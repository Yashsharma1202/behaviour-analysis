"""
Per-stock Diwali + Christmas backtest -> TWO workbooks (Nifty 50, F&O non-Nifty 50).
Each stock gets its OWN best side (LONG/SHORT) and OWN best window (T-lead .. T+hold), chosen by
win rate then average signed return over its history (needs >=7 valid event years).
A walk-forward check (selection redone each year using only earlier years) is shown beside every pick.
One sheet per event per year (Diwali_2000 ... Christmas_2025) lists every stock's trade that year.
Additive only: writes new files, touches nothing else, pushes nothing.
"""
import json, warnings
import numpy as np, pandas as pd
warnings.filterwarnings("ignore")

ROOT = r"D:\behaviour analysis"
OUT = {"NIFTY 50": ROOT + r"\Nifty50_Diwali_Christmas_Report_v4_8x8.xlsx",
       "F&O": ROOT + r"\FO_NonNifty50_Diwali_Christmas_Report_v4_8x8.xlsx"}
LEADS = list(range(1, 9)); HOLDS = list(range(1, 9)); YEARS = list(range(2000, 2026))   # PROJECT 8x8 GRID
MAX_MOVE = 0.35; MIN_YEARS = 7

hol = pd.read_csv(f"{ROOT}/NSE_BSE_Historical_Trading_Holidays_2000_2026.csv", parse_dates=["Holiday_Date"])
px = pd.read_parquet(f"{ROOT}/Stock_Data_213/all_stocks_daily_ohlc.parquet", columns=["Symbol", "Date", "Adj Close", "Volume"]).dropna().sort_values(["Symbol", "Date"])
# v4 FIX 1 -- real market calendar. Early rows contain zero-volume PLACEHOLDER rows on market holidays (e.g. 16-Nov-2001);
# counting them as trading days shifted every window. A date is a market day only if >=30% of the stocks listed that day traded volume.
_act = px.groupby("Date").agg(n=("Symbol", "size"), pos=("Volume", lambda v: (v > 0).sum()))
CAL = _act.index[(_act.n >= 5) & (_act.pos / _act.n >= 0.30)].values.astype("datetime64[D]")
lots = pd.read_csv(f"{ROOT}/Stock_Data_213/lot_sizes.csv").set_index("Symbol")["LotSize"]
rep = pd.read_csv(f"{ROOT}/Stock_Data_213/download_report.csv").set_index("Symbol")
meta = {r["symbol"]: r for r in json.load(open(f"{ROOT}/dashboard_data/fo_stocks_211.json"))}
NIFTY50 = {k for k, v in meta.items() if str(v.get("is_nifty50", "")).upper() == "YES"}

# ---------------------------------------------------------------- event anchors
d_ = hol[hol.Holiday_Name.str.contains("Diwali", case=False)]
ANCH = {"Diwali": {}, "Christmas": {}}
# v4 FIX 2 -- the holiday CSV's Diwali date is off by 1-3 days in several years (2000, 2001, 2003, 2005, 2007 ...).
# Laxmi Pujan dates below; verified against the price data: 2000-04 show the volume dip on the day itself, 2005+ show the exchange closure/Muhurat row.
LAXMI_PUJAN = {2000: "10-26", 2001: "11-14", 2002: "11-04", 2003: "10-25", 2004: "11-12", 2005: "11-01", 2006: "10-21", 2007: "11-09",
               2008: "10-28", 2009: "10-17", 2010: "11-05", 2011: "10-26", 2012: "11-13", 2013: "11-03", 2014: "10-23", 2015: "11-11",
               2016: "10-30", 2017: "10-19", 2018: "11-07", 2019: "10-27", 2020: "11-14", 2021: "11-04", 2022: "10-24", 2023: "11-12",
               2024: "11-01", 2025: "10-21", 2026: "11-08"}
for y in YEARS + [2026]:
    ANCH["Diwali"][y] = pd.Timestamp(f"{y}-{LAXMI_PUJAN[y]}")
    ANCH["Christmas"][y] = pd.Timestamp(y, 12, 25)
EVENTS = ("Diwali", "Christmas")

# ---------------------------------------------------------------- return cube
def build(dates, adj, ev):
    R = np.full((len(YEARS), 8, 8), np.nan); IP = np.full(len(YEARS), -1); IQ = np.full(len(YEARS), -1)
    for yi, y in enumerate(YEARS):
        T = np.datetime64(ANCH[ev][y], "D")
        ipre = np.searchsorted(dates, T, "left") - 1; ipost = np.searchsorted(dates, T, "right")
        if ipre < 0 or ipost >= len(dates): continue
        if (T - dates[ipre]).astype(int) > 7 or (dates[ipost] - T).astype(int) > 7: continue
        IP[yi], IQ[yi] = ipre, ipost
        for li in range(8):
            e = ipre - li
            if e < 0: continue
            for hi in range(8):
                x = ipost + hi
                if x >= len(dates) or (dates[x] - dates[e]).astype(int) > 40: continue
                seg = adj[e:x + 1]
                if len(seg) < 2 or np.max(np.abs(np.diff(seg) / seg[:-1])) > MAX_MOVE: continue
                R[yi, li, hi] = (adj[x] / adj[e] - 1) * 100
    return R, IP, IQ

def select(C):
    """best (side, lead, hold) by win rate then avg signed return, needs >=MIN_YEARS valid years."""
    valid = ~np.isnan(C); n = valid.sum(0)
    with np.errstate(all="ignore"):
        wl = ((C > 0) & valid).sum(0) / n; wsh = ((C < 0) & valid).sum(0) / n; avg = np.nanmean(C, 0)
    sl = np.where(n >= MIN_YEARS, wl + 1e-6 * avg, -1); ss = np.where(n >= MIN_YEARS, wsh - 1e-6 * avg, -1)
    bl, bs = np.nanmax(sl), np.nanmax(ss)
    if max(bl, bs) < 0: return None
    if bl >= bs: li, hi = np.unravel_index(np.nanargmax(sl), sl.shape); return ("LONG", li, hi, wl[li, hi] * 100, avg[li, hi], int(n[li, hi]))
    li, hi = np.unravel_index(np.nanargmax(ss), ss.shape); return ("SHORT", li, hi, wsh[li, hi] * 100, -avg[li, hi], int(n[li, hi]))

data = {ev: {} for ev in EVENTS}
symbols = sorted(s for s in px.Symbol.unique() if s in lots.index)
for s in symbols:
    g = px[px.Symbol == s].drop_duplicates("Date").set_index("Date")["Adj Close"]
    g.index = g.index.values.astype("datetime64[D]")
    g = g[g.index.isin(CAL)]                                   # drop non-market-day (placeholder) rows
    cal_s = CAL[(CAL >= g.index.min()) & (CAL <= g.index.max())]
    g = g.reindex(cal_s).ffill()                               # stock-specific gaps on market days -> carry price
    dts = cal_s; a = g.values.astype(float)
    for ev in EVENTS:
        keep = ~np.isin(dts, [np.datetime64(ANCH[ev][y], "D") for y in YEARS]) if ev == "Diwali" else np.ones(len(dts), bool)
        dd, aa = dts[keep], a[keep]
        R, IP, IQ = build(dd, aa, ev)
        data[ev][s] = dict(R=R, IP=IP, IQ=IQ, d=dd, a=aa)
print("cubes built", len(symbols))

# event-wide basket window (equal-weight basket of all stocks; best |t-stat|) -- used as the placeholder for stocks with <7 valid years
EVENT_WINDOW = {}
for _ev in EVENTS:
    _st = np.stack([data[_ev][s]["R"] for s in symbols]); _B = np.nanmean(_st, axis=0)
    _m = np.nanmean(_B, 0); _sd = np.nanstd(_B, 0, ddof=1); _n = np.sum(~np.isnan(_B), 0); _t = _m / (_sd / np.sqrt(_n))
    _i, _j = np.unravel_index(np.nanargmax(np.abs(_t)), _t.shape)
    _w = _B[:, _i, _j]; _w = _w[~np.isnan(_w)]
    EVENT_WINDOW[_ev] = dict(window=f"T-{_i + 1} to T+{_j + 1}", side="LONG" if _m[_i, _j] > 0 else "SHORT", t=round(float(_t[_i, _j]), 2),
                             basket_wr=round(float((_w > 0).mean() * 100), 1), n=len(_w))
print("EVENT WINDOWS", EVENT_WINDOW)
json.dump(EVENT_WINDOW, open(ROOT + r"\Diwali_Christmas_event_windows_v4_8x8.json", "w"), indent=1)

# ---------------------------------------------------------------- per-stock analysis
rows = {ev: [] for ev in EVENTS}; trades = {ev: [] for ev in EVENTS}
for ev in EVENTS:
    for s in symbols:
        D = data[ev][s]; C = D["R"]
        nmax = int((~np.isnan(C)).sum(0).max())
        grp = "NIFTY 50" if s in NIFTY50 else "F&O"
        base = dict(Group=grp, Symbol=s, Name=meta.get(s, {}).get("name", ""), Lot=lots.get(s, np.nan),
                    Last_Close=rep.LatestClose.get(s, np.nan), Data_From=rep.FirstDate.get(s, ""))
        sel = select(C)
        if sel is None:
            rows[ev].append(dict(**base, Position="NO CALL", Conviction="DATA INSUFFICIENT", Window="", Valid_Years=nmax,
                                 Data_Note=f"DATA INSUFFICIENT: best window has only {nmax} valid yr(s), need {MIN_YEARS} -- no position"))
            continue
        side, li, hi, wrate, avg, n = sel; sg = 1 if side == "LONG" else -1
        v = sg * C[:, li, hi]; ok = ~np.isnan(v); vv = v[ok]
        l10 = v[-10:]; l10 = l10[~np.isnan(l10)]
        # walk-forward (selection uses only earlier years)
        oos = []
        for yi in range(len(YEARS)):
            if np.isnan(C[yi]).all(): continue
            p = select(C[:yi]) if yi >= MIN_YEARS else None
            if p is None: continue
            r = (1 if p[0] == "LONG" else -1) * C[yi, p[1], p[2]]
            if not np.isnan(r): oos.append(r)
        oos = np.array(oos); oos_wr = np.mean(oos > 0) * 100 if len(oos) else np.nan
        l10wr = np.mean(l10 > 0) * 100 if len(l10) else np.nan
        if n >= 10 and wrate >= 75 and avg > 0 and len(oos) >= 5 and oos_wr >= 60: conv = "STRONG"
        elif wrate >= 65 and avg > 0 and (len(oos) < 5 or oos_wr >= 50): conv = "MODERATE"
        else: conv = "NO TRADE"
        rows[ev].append(dict(**base, Position=side, Conviction=conv, Window=f"T-{LEADS[li]} to T+{HOLDS[hi]}",
                             Lead=LEADS[li], Hold=HOLDS[hi], Valid_Years=n, Win_Rate=round(wrate, 1), Avg_Signed_Ret=round(avg, 2),
                             Median_Signed_Ret=round(float(np.median(vv)), 2), Best_Year=round(float(vv.max()), 2), Worst_Year=round(float(vv.min()), 2),
                             Last10_WinRate=None if np.isnan(l10wr) else round(l10wr, 1),
                             WF_OOS_Trades=len(oos), WF_OOS_WinRate=None if np.isnan(oos_wr) else round(oos_wr, 1),
                             WF_OOS_AvgRet=None if not len(oos) else round(float(oos.mean()), 2), Data_Note=""))
        # per-year trade log at the chosen window
        for yi, y in enumerate(YEARS):
            if np.isnan(C[yi, li, hi]): continue
            e = D["IP"][yi] - li; x = D["IQ"][yi] + hi
            trades[ev].append(dict(Year=y, Group=grp, Symbol=s, Name=base["Name"], Position=side, Conviction=conv,
                                   Window=f"T-{LEADS[li]} to T+{HOLDS[hi]}", Entry_Date=pd.Timestamp(D["d"][e]).strftime("%d-%b-%Y (%a)"),
                                   Exit_Date=pd.Timestamp(D["d"][x]).strftime("%d-%b-%Y (%a)"),
                                   Entry_AdjClose=round(float(D["a"][e]), 2), Exit_AdjClose=round(float(D["a"][x]), 2),
                                   LONG_Ret_pct=round(float(C[yi, li, hi]), 2), Signed_Ret_pct=round(float(sg * C[yi, li, hi]), 2),
                                   Result="WIN" if sg * C[yi, li, hi] > 0 else "LOSS", Lot=lots.get(s, np.nan),
                                   Approx_PnL_per_Lot=round(float(sg * (D["a"][x] - D["a"][e]) * lots.get(s, np.nan)), 0)))
    print(ev, pd.DataFrame(rows[ev]).Conviction.value_counts().to_dict())

# ---------------------------------------------------------------- 2026 plan dates
closed = set(hol[hol.Is_Trading_Holiday == "YES"].Holiday_Date)
closed.add(pd.Timestamp("2026-11-10"))   # Diwali Balipratipada 2026: missing from the holiday CSV but closed per the repo's NSE_2026_HOLIDAYS lists / existing dashboard dates
def plan_dates(ev, lead, hold):
    T = ANCH[ev][2026]; cal = [d for d in pd.bdate_range(T - pd.Timedelta(days=40), T + pd.Timedelta(days=40)) if d not in closed and d != T]
    return [d for d in cal if d < T][-lead], [d for d in cal if d > T][hold - 1]

# ---------------------------------------------------------------- Excel
def make_book(grp):
    wbw = pd.ExcelWriter(OUT[grp], engine="xlsxwriter"); wb = wbw.book
    F = lambda **k: wb.add_format({"font_name": "Calibri", "font_size": 10, **k})
    f_title = F(bold=True, font_size=16, font_color="#FFFFFF", bg_color="#1F3864", valign="vcenter")
    f_h = F(bold=True, font_color="#FFFFFF", bg_color="#2F5597", border=1, align="center", valign="vcenter", text_wrap=True)
    f_cell = F(border=1, border_color="#D9D9D9"); f_num = F(border=1, border_color="#D9D9D9", num_format="0.00"); f_int = F(border=1, border_color="#D9D9D9", num_format="#,##0")
    f_b = F(bold=True, font_size=12, font_color="#1F3864"); f_sub = F(italic=True, font_color="#595959"); f_wrap = F(text_wrap=True, valign="top")
    green = F(bg_color="#C6EFCE", font_color="#006100"); red = F(bg_color="#FFC7CE", font_color="#9C0006"); amber = F(bg_color="#FFEB9C", font_color="#9C5700")
    bgreen = F(bg_color="#63BE7B", bold=True); grey = F(bg_color="#BFBFBF", font_color="#404040", bold=True); greyrow = F(font_color="#7F7F7F", italic=True)
    scale3 = lambda lo, mid, hi: {"type": "3_color_scale", "min_type": "num", "min_value": lo, "mid_type": "num", "mid_value": mid, "max_type": "num", "max_value": hi, "min_color": "#F8696B", "mid_color": "#FFFFFF", "max_color": "#63BE7B"}

    def write_df(ws, df, r0, widths=None):
        for c, name in enumerate(df.columns): ws.write(r0, c, name.replace("_", " "), f_h)
        ws.set_row(r0, 32)
        for r, row in enumerate(df.itertuples(index=False), start=r0 + 1):
            for c, val in enumerate(row):
                if val is None or (isinstance(val, float) and np.isnan(val)): ws.write_blank(r, c, None, f_cell)
                elif isinstance(val, (int, np.integer)): ws.write_number(r, c, int(val), f_int)
                elif isinstance(val, (float, np.floating)): ws.write_number(r, c, float(val), f_num)
                else: ws.write(r, c, val, f_cell)
        for c, name in enumerate(df.columns): ws.set_column(c, c, (widths or {}).get(name, max(10, min(22, len(name) + 2))))

    def tag_cf(ws, df, r0):
        cols = {n: i for i, n in enumerate(df.columns)}; last = r0 + len(df)
        if "Position" in cols:
            ws.conditional_format(r0 + 1, cols["Position"], last, cols["Position"], {"type": "cell", "criteria": "==", "value": '"LONG"', "format": green})
            ws.conditional_format(r0 + 1, cols["Position"], last, cols["Position"], {"type": "cell", "criteria": "==", "value": '"SHORT"', "format": red})
            ws.conditional_format(r0 + 1, cols["Position"], last, cols["Position"], {"type": "cell", "criteria": "==", "value": '"NO CALL"', "format": grey})
        if "Conviction" in cols:
            for val, fm in (("STRONG", bgreen), ("MODERATE", amber), ("DATA INSUFFICIENT", grey)):
                ws.conditional_format(r0 + 1, cols["Conviction"], last, cols["Conviction"], {"type": "cell", "criteria": "==", "value": f'"{val}"', "format": fm})
        if "Result" in cols:
            ws.conditional_format(r0 + 1, cols["Result"], last, cols["Result"], {"type": "cell", "criteria": "==", "value": '"WIN"', "format": green})
            ws.conditional_format(r0 + 1, cols["Result"], last, cols["Result"], {"type": "cell", "criteria": "==", "value": '"LOSS"', "format": red})
        return cols, last

    dfs = {}
    for ev in EVENTS:
        d = pd.DataFrame(rows[ev]); d = d[d.Group == grp].copy()
        d["_o"] = d.Conviction.map({"STRONG": 0, "MODERATE": 1, "NO TRADE": 2, "DATA INSUFFICIENT": 3})
        d = d.sort_values(["_o", "Win_Rate", "Avg_Signed_Ret"], ascending=[True, False, False]).drop(columns="_o").reset_index(drop=True)
        d.insert(0, "Rank", d.index + 1); dfs[ev] = d

    # ---- Summary
    ws = wb.add_worksheet("Summary"); ws.hide_gridlines(2)
    ws.merge_range("A1:L1", f"Diwali & Christmas  |  {grp}  |  per-stock best side + best window  |  2000-2025", f_title); ws.set_row(0, 32)
    ws.write("A2", "Review workbook -- NOT pushed to dashboard. Source: Stock_Data_213 (Adj Close) + NSE/BSE holiday calendar. Entry = close of T-lead, exit = close of T+hold; T = Laxmi Pujan/Muhurat date (Diwali) or 25-Dec.", f_sub)
    cnt = []
    for ev in EVENTS:
        d = dfs[ev]
        for conv in ("STRONG", "MODERATE", "NO TRADE", "DATA INSUFFICIENT"):
            s_ = d[d.Conviction == conv]
            cnt.append(dict(Event=ev, Conviction=conv, Stocks=len(s_), LONG=int((s_.Position == "LONG").sum()), SHORT=int((s_.Position == "SHORT").sum()),
                            Avg_InSample_WinRate=None if conv == "DATA INSUFFICIENT" or not len(s_) else round(float(s_.Win_Rate.mean()), 1),
                            Avg_WalkForward_WinRate=None if conv == "DATA INSUFFICIENT" or not len(s_) else round(float(s_.WF_OOS_WinRate.mean()), 1)))
    ws.write("A4", "Stock counts by conviction (each stock's own best side/window)", f_b)
    cd = pd.DataFrame(cnt); write_df(ws, cd, 4, {"Event": 12, "Conviction": 20}); tag_cf(ws, cd, 4)
    r = 4 + len(cd) + 2
    ws.write(r, 0, "How to read", f_b)
    for i, t in enumerate([
        "Each stock is scanned over lead 1-10 x hold 1-10 trading days and BOTH sides; the (side, window) with the highest win rate, ties broken by average signed return, is its pick. Needs >=7 valid event years in that window.",
        "Position-taking window: 'T-4 to T+3' = enter at close of the 4th trading day before the holiday, exit at close of the 3rd trading day after it.",
        "Per-stock picking from 200 combinations is optimistic in-sample. WF (walk-forward) columns redo the pick every year using ONLY earlier years -- if WF win rate is far below Win Rate, treat the pick as fragile.",
        "STRONG: >=10 yrs, win rate >=75%, positive avg, WF win rate >=60% (>=5 WF trades).  MODERATE: win rate >=65%, positive avg, WF >=50% (or too few WF trades).  NO TRADE: anything else.  DATA INSUFFICIENT: <7 valid event years -> NO CALL, listed at the bottom for reference.",
        "Windows with any daily move beyond +/-35% are skipped (split/bad-tick guard). Approx P&L per lot uses Adj Close prices x lot size and ignores costs.",
        "Year sheets (Diwali_2000 ... Christmas_2025) show each stock's trade that year at its chosen window. Those windows were picked with all years, so the year sheets are in-sample history, not a forecast."]):
        ws.merge_range(r + 1 + i, 0, r + 1 + i, 11, "- " + t, f_wrap); ws.set_row(r + 1 + i, 30)

    # ---- Positions sheets
    SW = {"Name": 30, "Symbol": 14, "Rank": 6, "Position": 10, "Conviction": 18, "Window": 13, "Data_Note": 62, "Group": 10}
    cols_order = ["Rank", "Symbol", "Name", "Position", "Window", "Conviction", "Win_Rate", "Avg_Signed_Ret", "Valid_Years", "Last10_WinRate",
                  "WF_OOS_Trades", "WF_OOS_WinRate", "WF_OOS_AvgRet", "Median_Signed_Ret", "Best_Year", "Worst_Year", "Lot", "Last_Close", "Data_From", "Data_Note"]
    for ev in EVENTS:
        d = dfs[ev][cols_order]
        ws = wb.add_worksheet(f"{ev}_Positions")
        ws.merge_range(0, 0, 0, 12, f"{ev} | {grp} | {len(d)} stocks -- each stock's own best side + window ({int((d.Conviction=='DATA INSUFFICIENT').sum())} data-insufficient at bottom, no call)", f_title); ws.set_row(0, 30)
        write_df(ws, d, 2, SW); ws.freeze_panes(3, 3); ws.autofilter(2, 0, 2 + len(d), len(d.columns) - 1)
        cols, last = tag_cf(ws, d, 2)
        for k in ("Win_Rate", "Last10_WinRate", "WF_OOS_WinRate"): ws.conditional_format(3, cols[k], last, cols[k], scale3(30, 55, 85))
        for k in ("Avg_Signed_Ret", "WF_OOS_AvgRet"): ws.conditional_format(3, cols[k], last, cols[k], {"type": "data_bar", "bar_color": "#5B9BD5"})
        ws.conditional_format(3, 0, last, len(d.columns) - 1, {"type": "formula", "criteria": '=$F4="DATA INSUFFICIENT"', "format": greyrow})

    # ---- 2026 plan
    ws = wb.add_worksheet("Plan_2026")
    ws.merge_range(0, 0, 0, 14, f"2026 trade plan | {grp} -- STRONG + MODERATE only (review before any dashboard change)", f_title); ws.set_row(0, 30)
    r = 2
    for ev in EVENTS:
        d = dfs[ev]; d = d[d.Conviction.isin(["STRONG", "MODERATE"])].copy()
        ent, ext = zip(*[plan_dates(ev, int(a), int(b)) for a, b in zip(d.Lead, d.Hold)]) if len(d) else ([], [])
        d["Entry_Date"] = [x.strftime("%d-%b-%Y (%a)") for x in ent]; d["Exit_Date"] = [x.strftime("%d-%b-%Y (%a)") for x in ext]
        d["Futures_Action"] = np.where(d.Position == "LONG", "BUY FUTURES", "SELL FUTURES"); d["Option_Play"] = np.where(d.Position == "LONG", "1% ITM CALL", "1% ITM PUT")
        d["Contract_Value"] = (d.Lot * d.Last_Close).round(0); d["Margin_20pct"] = (d.Contract_Value * 0.2).round(0)
        d = d[["Rank", "Symbol", "Name", "Position", "Conviction", "Window", "Entry_Date", "Exit_Date", "Win_Rate", "Avg_Signed_Ret", "WF_OOS_WinRate", "Valid_Years", "Futures_Action", "Option_Play", "Lot", "Last_Close", "Contract_Value", "Margin_20pct"]]
        ws.write(r, 0, f"{ev} 2026  (T = {ANCH[ev][2026].strftime('%d-%b-%Y (%a)')})  --  {len(d)} stocks", f_b)
        write_df(ws, d, r + 1, {**SW, "Entry_Date": 18, "Exit_Date": 18, "Futures_Action": 15, "Option_Play": 13}); tag_cf(ws, d, r + 1)
        r += len(d) + 4

    # ---- Year-wise matrices
    for ev in EVENTS:
        t = pd.DataFrame(trades[ev]); t = t[t.Group == grp]
        mat = t.pivot(index="Symbol", columns="Year", values="Signed_Ret_pct").reindex(columns=YEARS)
        mat.columns = [str(c) for c in mat.columns]
        order = dfs[ev].set_index("Symbol")
        mat.insert(0, "Window", order.Window.reindex(mat.index)); mat.insert(0, "Position", order.Position.reindex(mat.index))
        mat = mat.reset_index(); mat = mat.sort_values("Symbol")
        ws = wb.add_worksheet(f"{ev}_YearMatrix")
        ws.merge_range(0, 0, 0, 12, f"{ev} | {grp} -- SIGNED % return by year at each stock's own side/window (positive = the position won)", f_title); ws.set_row(0, 30)
        write_df(ws, mat, 2, {"Symbol": 14, "Position": 10, "Window": 13}); ws.freeze_panes(3, 3)
        for c in range(3, len(mat.columns)): ws.set_column(c, c, 7.5)
        ws.conditional_format(3, 3, 2 + len(mat), len(mat.columns) - 1, scale3(-8, 0, 8)); tag_cf(ws, mat, 2)

    # ---- one sheet per event per year
    ycols = ["Rank", "Symbol", "Name", "Position", "Conviction", "Window", "Entry_Date", "Exit_Date", "Entry_AdjClose", "Exit_AdjClose", "LONG_Ret_pct", "Signed_Ret_pct", "Result", "Lot", "Approx_PnL_per_Lot"]
    for ev in EVENTS:
        t_all = pd.DataFrame(trades[ev]); t_all = t_all[t_all.Group == grp]
        for y in YEARS:
            t = t_all[t_all.Year == y].sort_values("Signed_Ret_pct", ascending=False).reset_index(drop=True)
            ws = wb.add_worksheet(f"{ev}_{y}")
            T = ANCH[ev][y]
            ws.merge_range(0, 0, 0, 12, f"{ev} {y}  |  {grp}  |  T = {T.strftime('%d-%b-%Y (%a)')}", f_title); ws.set_row(0, 28)
            if t.empty:
                ws.write(2, 0, "No stock had a valid trade this year (no price data / windows skipped).", f_sub); continue
            t.insert(0, "Rank", t.index + 1); t = t[ycols]
            wins = int((t.Result == "WIN").sum())
            ws.write(1, 0, f"Trades {len(t)}  |  Wins {wins}  |  Win rate {wins / len(t) * 100:.1f}%  |  Avg signed return {t.Signed_Ret_pct.mean():.2f}%  |  LONG {int((t.Position=='LONG').sum())} / SHORT {int((t.Position=='SHORT').sum())}  |  Windows chosen using all years (in-sample)", f_b)
            write_df(ws, t, 2, {**SW, "Entry_Date": 18, "Exit_Date": 18}); ws.freeze_panes(3, 3); ws.autofilter(2, 0, 2 + len(t), len(t.columns) - 1)
            cols, last = tag_cf(ws, t, 2)
            ws.conditional_format(3, cols["Signed_Ret_pct"], last, cols["Signed_Ret_pct"], scale3(-6, 0, 6))

    # ---- anchors
    ws = wb.add_worksheet("Event_Dates")
    ad = pd.DataFrame([dict(Year=y, Diwali_T=ANCH["Diwali"][y].strftime("%d-%b-%Y (%a)"), Christmas_T=ANCH["Christmas"][y].strftime("%d-%b-%Y (%a)")) for y in YEARS + [2026]])
    ws.merge_range(0, 0, 0, 3, "Event anchor dates (T) used", f_title); write_df(ws, ad, 2, {"Year": 7, "Diwali_T": 20, "Christmas_T": 20})
    wbw.close(); print("written", OUT[grp])

for g in OUT: make_book(g)
