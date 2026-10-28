"""
Diwali + Christmas holiday backtest on the 213-stock F&O universe (Stock_Data_213).
Writes ONE review workbook only -- Diwali_Christmas_213_Backtest_Review.xlsx.
Does NOT touch dashboard_data/ or any existing file (additive only, nothing pushed).

Convention: T = holiday date. Trade day T-k = k-th trading day before T; T+k = k-th trading
day after. Entry at CLOSE of T-lead, exit at CLOSE of T+hold. Diwali T = Laxmi Pujan / Muhurat
date (the Muhurat row, if any, is dropped from the price series); Christmas T = 25-Dec.
"""
import json, sys
import warnings
import numpy as np, pandas as pd
warnings.filterwarnings('ignore')

ROOT = r"D:\behaviour analysis"
OUT = ROOT + r"\Diwali_Christmas_213_Backtest_Review_8x8.xlsx"
LEADS = list(range(1, 9)); HOLDS = list(range(1, 9))          # PROJECT 8x8 GRID
YEARS = list(range(2000, 2026))
MAX_DAILY_MOVE = 0.20       # skip window if any daily move beyond +-20% (split/bad-tick artefact)
MIN_WF_PRIOR = 6            # min prior years before walk-forward picks a window

# ---------------------------------------------------------------- inputs
hol = pd.read_csv(f"{ROOT}/NSE_BSE_Historical_Trading_Holidays_2000_2026.csv", parse_dates=["Holiday_Date"])
px = pd.read_parquet(f"{ROOT}/Stock_Data_213/all_stocks_daily_ohlc.parquet", columns=["Symbol", "Date", "Adj Close"])
px = px.dropna().sort_values(["Symbol", "Date"])
lots = pd.read_csv(f"{ROOT}/Stock_Data_213/lot_sizes.csv").set_index("Symbol")["LotSize"]
rep = pd.read_csv(f"{ROOT}/Stock_Data_213/download_report.csv").set_index("Symbol")
meta = {r["symbol"]: r for r in json.load(open(f"{ROOT}/dashboard_data/fo_stocks_211.json"))}

def anchors():
    d = hol[hol.Holiday_Name.str.contains("Diwali", case=False)]
    out = {}
    for y in YEARS + [2026]:
        g = d[d.Year == y]
        lp = g[g.Is_Trading_Holiday == "SPECIAL_SESSION"]
        pick = lp.iloc[0] if len(lp) else g.iloc[0]
        out[y] = (pick.Holiday_Date, "Laxmi Pujan/Muhurat" if len(lp) else "Balipratipada (no Muhurat row)")
    diw = out
    xm = {y: (pd.Timestamp(y, 12, 25), "25-Dec") for y in YEARS + [2026]}
    return {"Diwali": diw, "Christmas": xm}
ANCH = anchors()

# ---------------------------------------------------------------- returns cube
def stock_cube(dates, adj, event):
    """R[year_idx, lead_idx, hold_idx] long % return, NaN when unusable."""
    R = np.full((len(YEARS), len(LEADS), len(HOLDS)), np.nan)
    lr = np.abs(np.diff(adj) / adj[:-1])
    for yi, y in enumerate(YEARS):
        T = ANCH[event][y][0]
        ipre = np.searchsorted(dates, np.datetime64(T), "left") - 1
        ipost = np.searchsorted(dates, np.datetime64(T), "right")
        # drop a row dated exactly on T (Muhurat) -> handled because ipost uses 'right'
        if ipre < 0 or ipost >= len(dates):
            continue
        if (np.datetime64(T) - dates[ipre]).astype("timedelta64[D]").astype(int) > 7: continue
        if (dates[ipost] - np.datetime64(T)).astype("timedelta64[D]").astype(int) > 7: continue
        for li, l in enumerate(LEADS):
            e = ipre - (l - 1)
            if e < 0: continue
            for hi, h in enumerate(HOLDS):
                x = ipost + (h - 1)
                if x >= len(dates): continue
                if (dates[x] - dates[e]).astype("timedelta64[D]").astype(int) > 40: continue
                seg = adj[e:x + 1]
                if len(seg) < 2: continue
                # window path includes the T row itself if Muhurat row present: exclude it
                if np.max(np.abs(np.diff(seg) / seg[:-1])) > MAX_DAILY_MOVE: continue
                R[yi, li, hi] = (adj[x] / adj[e] - 1) * 100
    return R

# Muhurat-day rows are removed from the series so T is purely a gap
cubes = {"Diwali": {}, "Christmas": {}}
symbols = [s for s in sorted(px.Symbol.unique())]
for s in symbols:
    g = px[px.Symbol == s]
    dts = g.Date.values.astype("datetime64[D]"); a = g["Adj Close"].values.astype(float)
    for ev in cubes:
        T_set = {np.datetime64(ANCH[ev][y][0], "D") for y in YEARS}
        keep = ~np.isin(dts, list(T_set)) if ev == "Diwali" else np.ones(len(dts), bool)
        cubes[ev][s] = stock_cube(dts[keep], a[keep], ev)
print("cubes built", len(symbols))

# ---------------------------------------------------------------- helpers
def wr(v): v = v[~np.isnan(v)]; return (np.mean(v > 0) * 100) if len(v) else np.nan

def basket_grid(ev):
    """equal-weight basket (all stocks with data) year x lead x hold return."""
    stack = np.stack([cubes[ev][s] for s in symbols])           # S,Y,L,H
    return np.nanmean(stack, axis=0), np.sum(~np.isnan(stack).all(axis=(2, 3)), axis=0)

def choose_window(B):
    """best event window by t-stat of basket yearly mean return (sign gives direction)."""
    best = None; mean = np.nanmean(B, axis=0); sd = np.nanstd(B, axis=0, ddof=1); n = np.sum(~np.isnan(B), axis=0)
    t = mean / (sd / np.sqrt(n))
    li, hi = np.unravel_index(np.nanargmax(np.abs(t)), t.shape)
    return li, hi, t

summary_rows, stock_tables, year_tables, basket_tables, grids, plans = [], {}, {}, {}, {}, {}
for ev in ("Diwali", "Christmas"):
    B, nB = basket_grid(ev)
    li, hi, tgrid = choose_window(B)
    lead, hold = LEADS[li], HOLDS[hi]
    bw = B[:, li, hi]
    direction = "LONG" if np.nanmean(bw) > 0 else "SHORT"
    sgn = 1 if direction == "LONG" else -1
    ev_wr = np.nanmean((sgn * bw[~np.isnan(bw)]) > 0) * 100
    # fixed benchmark windows from the earlier index study
    bench = {"Diwali": (4, 3), "Christmas": (1, 7)}[ev]
    bb = B[:, bench[0] - 1, bench[1] - 1]
    # walk-forward for basket window
    wf_hits = []
    for yi in range(MIN_WF_PRIOR, len(YEARS)):
        P = B[:yi]; m = np.nanmean(P, axis=0); a, b = np.unravel_index(np.nanargmax(np.abs(m)), m.shape)
        r = np.sign(m[a, b]) * B[yi, a, b]
        if not np.isnan(r): wf_hits.append(r)
    wf_hits = np.array(wf_hits)
    summary_rows.append(dict(
        Event=ev, T_2026=ANCH[ev][2026][0].strftime("%d-%b-%Y (%a)"), Best_Window=f"T-{lead} to T+{hold}", Direction=direction,
        Basket_WinRate_pct=round(ev_wr, 1), Basket_AvgRet_pct=round(float(np.nanmean(sgn * bw)), 2),
        Years=int(np.sum(~np.isnan(bw))), T_stat=round(float(tgrid[li, hi]), 2),
        Last10_WinRate_pct=round(np.nanmean((sgn * bw[-10:][~np.isnan(bw[-10:])]) > 0) * 100, 1),
        WalkForward_OOS_WR_pct=round(float(np.mean(wf_hits > 0) * 100), 1), WalkForward_OOS_n=len(wf_hits),
        Prior_Study_Window=f"T-{bench[0]} to T+{bench[1]}",
        Prior_Window_LONG_WR_pct=round(wr(bb), 1), Prior_Window_LONG_Avg_pct=round(float(np.nanmean(bb)), 2),
        Stocks_In_Universe=len(symbols)))
    grids[ev] = dict(avg=np.nanmean(B, axis=0), wr=np.array([[wr(B[:, i, j]) for j in range(8)] for i in range(8)]), best=(li, hi))
    basket_tables[ev] = pd.DataFrame(dict(Year=YEARS, T_date=[ANCH[ev][y][0].strftime("%d-%b-%Y") for y in YEARS],
                                          Anchor=[ANCH[ev][y][1] for y in YEARS], Stocks_With_Data=nB,
                                          Basket_LONG_Ret_pct=np.round(bw, 2), Signed_Ret_pct=np.round(sgn * bw, 2)))

    # ---------------- per stock
    rows, yrows = [], []
    for s in symbols:
        C = cubes[ev][s]; v = C[:, li, hi]; ok = ~np.isnan(v)
        n = int(ok.sum())
        if n == 0:
            continue
        lw = wr(v); sw = np.mean(v[ok] < 0) * 100; avg = float(np.nanmean(v))
        pdir = "LONG" if lw >= sw else "SHORT"
        sg = 1 if pdir == "LONG" else -1
        swr = max(lw, sw); savg = sg * avg
        l10 = v[-10:]; l10 = l10[~np.isnan(l10)]
        rwr = np.mean(sg * l10 > 0) * 100 if len(l10) else np.nan
        # per-stock walk-forward (own window tuned on prior years only) -> overfit check
        wf = []
        for yi in range(len(YEARS)):
            if np.isnan(C[yi]).all(): continue
            P = C[:yi]
            if np.sum(~np.isnan(P).all(axis=(1, 2))) < MIN_WF_PRIOR: continue
            with np.errstate(all="ignore"):
                m = np.nanmean(P, axis=0)
            if np.isnan(m).all(): continue
            a, b = np.unravel_index(np.nanargmax(np.abs(np.nan_to_num(m, nan=0))), m.shape)
            r = np.sign(m[a, b]) * C[yi, a, b]
            if not np.isnan(r): wf.append(r)
        wf = np.array(wf)
        if n >= 10 and swr >= 70 and savg > 0 and (np.isnan(rwr) or rwr >= 60): conv = "STRONG"
        elif n >= 8 and swr >= 60 and savg > 0: conv = "MODERATE"
        else: conv = "NO TRADE"
        worst = float(np.nanmin(sg * v)); best = float(np.nanmax(sg * v))
        sym_meta = meta.get(s, {})
        rows.append(dict(
            Symbol=s, Name=sym_meta.get("name", ""), Nifty50=sym_meta.get("is_nifty50", ""),
            Years_Of_Data=n, First_Event_Year=YEARS[int(np.argmax(ok))],
            LONG_WinRate=round(lw, 1), SHORT_WinRate=round(sw, 1), Avg_LONG_Ret=round(avg, 2),
            Median_LONG_Ret=round(float(np.nanmedian(v)), 2),
            Position=pdir, Position_WinRate=round(swr, 1), Position_AvgRet=round(savg, 2),
            Best_Year_Ret=round(best, 2), Worst_Year_Ret=round(worst, 2),
            Last10_WinRate=None if np.isnan(rwr) else round(rwr, 1),
            Conviction=conv, Matches_Event_Direction="YES" if pdir == direction else "no",
            WF_OOS_Trades=len(wf), WF_OOS_WinRate=None if len(wf) == 0 else round(float(np.mean(wf > 0) * 100), 1),
            WF_OOS_AvgRet=None if len(wf) == 0 else round(float(wf.mean()), 2),
            Lot=lots.get(s, np.nan), Last_Close=rep.LatestClose.get(s, np.nan), Data_From=rep.FirstDate.get(s, "")))
        yr = dict(Symbol=s); yr.update({str(y): (None if np.isnan(v[i]) else round(float(v[i]), 2)) for i, y in enumerate(YEARS)})
        yrows.append(yr)
    df = pd.DataFrame(rows)
    order = {"STRONG": 0, "MODERATE": 1, "NO TRADE": 2}
    df = df.sort_values(by=["Conviction", "Position_WinRate", "Position_AvgRet"], key=lambda c: c.map(order) if c.name == "Conviction" else -c).reset_index(drop=True)
    df.insert(0, "Rank", df.index + 1)
    stock_tables[ev] = df; year_tables[ev] = pd.DataFrame(yrows)

    # ---------------- 2026 trade plan
    hol_closed = set(hol[(hol.Is_Trading_Holiday == "YES")].Holiday_Date)
    T = ANCH[ev][2026][0]
    cal = pd.bdate_range(T - pd.Timedelta(days=40), T + pd.Timedelta(days=40))
    cal = [d for d in cal if d not in hol_closed and d != T]
    before = [d for d in cal if d < T]; after = [d for d in cal if d > T]
    entry, exit_ = before[-lead], after[hold - 1]
    plan = df[df.Conviction != "NO TRADE"].copy()
    plan["Entry_Date"] = entry.strftime("%d-%b-%Y (%a)"); plan["Exit_Date"] = exit_.strftime("%d-%b-%Y (%a)")
    plan["Window"] = f"T-{lead} to T+{hold}"
    plan["Futures_Action"] = np.where(plan.Position == "LONG", "BUY FUTURES", "SELL FUTURES")
    plan["Option_Play"] = np.where(plan.Position == "LONG", "1% ITM CALL", "1% ITM PUT")
    plan["Contract_Value"] = (plan.Lot * plan.Last_Close).round(0)
    plan["Margin_20pct"] = (plan.Contract_Value * 0.2).round(0)
    plans[ev] = plan[["Rank", "Symbol", "Name", "Position", "Conviction", "Position_WinRate", "Position_AvgRet",
                      "Years_Of_Data", "Last10_WinRate", "Window", "Entry_Date", "Exit_Date", "Futures_Action",
                      "Option_Play", "Lot", "Last_Close", "Contract_Value", "Margin_20pct"]]
    print(ev, direction, f"T-{lead}/T+{hold}", df.Conviction.value_counts().to_dict(), df.groupby("Conviction").Position.value_counts().to_dict())

summary = pd.DataFrame(summary_rows)

# ---------------------------------------------------------------- Excel
wbw = pd.ExcelWriter(OUT, engine="xlsxwriter")
wb = wbw.book
F = lambda **k: wb.add_format({"font_name": "Calibri", "font_size": 10, **k})
f_title = F(bold=True, font_size=18, font_color="#FFFFFF", bg_color="#1F3864", valign="vcenter")
f_sub = F(italic=True, font_color="#595959")
f_h = F(bold=True, font_color="#FFFFFF", bg_color="#2F5597", border=1, align="center", valign="vcenter", text_wrap=True)
f_cell = F(border=1, border_color="#D9D9D9"); f_num = F(border=1, border_color="#D9D9D9", num_format="0.00")
f_pct = F(border=1, border_color="#D9D9D9", num_format="0.0"); f_int = F(border=1, border_color="#D9D9D9", num_format="#,##0")
f_wrap = F(text_wrap=True, valign="top"); f_b = F(bold=True, font_size=12, font_color="#1F3864")
green = F(bg_color="#C6EFCE", font_color="#006100"); red = F(bg_color="#FFC7CE", font_color="#9C0006"); amber = F(bg_color="#FFEB9C", font_color="#9C5700")
bold_green = F(bg_color="#63BE7B", bold=True); bold_red = F(bg_color="#F8696B", bold=True)

def write_df(ws, df, r0, widths=None, fmt_cols=None):
    for c, name in enumerate(df.columns):
        ws.write(r0, c, name.replace("_", " "), f_h)
    ws.set_row(r0, 32)
    for r, row in enumerate(df.itertuples(index=False), start=r0 + 1):
        for c, val in enumerate(row):
            if val is None or (isinstance(val, float) and np.isnan(val)): ws.write_blank(r, c, None, f_cell)
            elif isinstance(val, (int, np.integer)): ws.write_number(r, c, int(val), f_int)
            elif isinstance(val, (float, np.floating)): ws.write_number(r, c, float(val), f_num)
            else: ws.write(r, c, val, f_cell)
    for c, name in enumerate(df.columns):
        w = widths.get(name, 12) if widths else max(10, min(28, len(name) + 2))
        ws.set_column(c, c, w)

# --- Summary
ws = wb.add_worksheet("Summary"); wbw.sheets["Summary"] = ws; ws.hide_gridlines(2)
ws.merge_range("A1:N1", "Diwali & Christmas Holiday Backtest  |  213-Stock F&O Universe  |  2000-2025", f_title); ws.set_row(0, 34)
ws.write("A2", "Review workbook -- NOT pushed to dashboard. Source: Stock_Data_213 (Adj Close) + NSE_BSE_Historical_Trading_Holidays_2000_2026. Entry = close of T-lead, exit = close of T+hold.", f_sub)
ws.write("A4", "1) Event-level verdict (equal-weight basket of all stocks with data; window chosen by highest |t-stat|)", f_b)
write_df(ws, summary, 4)
ws.set_column(0, 0, 12); ws.set_column(1, 1, 20); ws.set_column(2, 20, 15)
ws.conditional_format(5, 3, 6, 3, {"type": "cell", "criteria": "==", "value": '"LONG"', "format": green})
ws.conditional_format(5, 3, 6, 3, {"type": "cell", "criteria": "==", "value": '"SHORT"', "format": red})
ws.write("A10", "2) Stock-level position counts (common event window)", f_b)
r = 10
cnt = []
for ev in ("Diwali", "Christmas"):
    d = stock_tables[ev]
    for conv in ("STRONG", "MODERATE", "NO TRADE"):
        s_ = d[d.Conviction == conv]
        cnt.append(dict(Event=ev, Conviction=conv, Stocks=len(s_), LONG=int((s_.Position == "LONG").sum()) if conv != "NO TRADE" else None,
                        SHORT=int((s_.Position == "SHORT").sum()) if conv != "NO TRADE" else None,
                        Avg_Position_WinRate=round(float(s_.Position_WinRate.mean()), 1) if len(s_) else None))
write_df(ws, pd.DataFrame(cnt), r)
ws.write(r + 9, 0, "3) How to read", f_b)
notes = [
    "Position = the side (LONG/SHORT) with the higher historical win rate for that stock at the event window; NO TRADE means no side clears the bar.",
    "STRONG: >=10 yrs data, position win rate >=70%, positive avg signed return, last-10-yr win rate >=60%.  MODERATE: >=8 yrs, win rate >=60%, positive avg.",
    "WF OOS columns = per-stock walk-forward (window+direction re-tuned each year using ONLY earlier years). Compare with in-sample win rate: a big drop means the stock's pattern is fragile.",
    "Windows scanned: lead 1-10 x hold 1-10 trading days (100 combos). Event window is chosen on the whole sample, so in-sample rates are optimistic -- see the walk-forward line.",
    "Any window with a daily move beyond +/-20% is skipped (split / bad-tick guard). Adj Close used. Missing years = stock not yet listed / no data near the event.",
]
for i, t in enumerate(notes): ws.merge_range(r + 10 + i, 0, r + 10 + i, 13, "- " + t, f_wrap); ws.set_row(r + 10 + i, 28)

# --- stock sheets
SW = {"Name": 30, "Symbol": 14, "Rank": 6, "Position": 10, "Conviction": 12}
for ev in ("Diwali", "Christmas"):
    d = stock_tables[ev]
    ws = wb.add_worksheet(f"{ev}_Positions"); wbw.sheets[f"{ev}_Positions"] = ws
    sm = summary[summary.Event == ev].iloc[0]
    ws.merge_range(0, 0, 0, 12, f"{ev} -- stock-wise LONG / SHORT positions  |  window {sm.Best_Window}  |  event bias {sm.Direction}", f_title); ws.set_row(0, 30)
    write_df(ws, d, 2, SW); ws.freeze_panes(3, 3); ws.autofilter(2, 0, 2 + len(d), len(d.columns) - 1)
    cols = {n: i for i, n in enumerate(d.columns)}
    last = 2 + len(d)
    ws.conditional_format(3, cols["Position"], last, cols["Position"], {"type": "cell", "criteria": "==", "value": '"LONG"', "format": green})
    ws.conditional_format(3, cols["Position"], last, cols["Position"], {"type": "cell", "criteria": "==", "value": '"SHORT"', "format": red})
    ws.conditional_format(3, cols["Conviction"], last, cols["Conviction"], {"type": "cell", "criteria": "==", "value": '"STRONG"', "format": bold_green})
    ws.conditional_format(3, cols["Conviction"], last, cols["Conviction"], {"type": "cell", "criteria": "==", "value": '"MODERATE"', "format": amber})
    for k in ("LONG_WinRate", "SHORT_WinRate", "Position_WinRate", "WF_OOS_WinRate", "Last10_WinRate"):
        ws.conditional_format(3, cols[k], last, cols[k], {"type": "3_color_scale", "min_type": "num", "min_value": 30, "mid_type": "num", "mid_value": 55, "max_type": "num", "max_value": 85, "min_color": "#F8696B", "mid_color": "#FFFFFF", "max_color": "#63BE7B"})
    for k in ("Avg_LONG_Ret", "Position_AvgRet", "WF_OOS_AvgRet"):
        ws.conditional_format(3, cols[k], last, cols[k], {"type": "data_bar", "bar_color": "#5B9BD5"})

# --- year matrices
for ev in ("Diwali", "Christmas"):
    d = year_tables[ev]; sm = summary[summary.Event == ev].iloc[0]
    ws = wb.add_worksheet(f"{ev}_YearWise"); wbw.sheets[f"{ev}_YearWise"] = ws
    ws.merge_range(0, 0, 0, 12, f"{ev} -- LONG % return by year at window {sm.Best_Window} (green = LONG would have won)", f_title); ws.set_row(0, 30)
    write_df(ws, d, 2, {"Symbol": 14}); ws.freeze_panes(3, 1)
    for c in range(1, len(d.columns)): ws.set_column(c, c, 7.5)
    ws.conditional_format(3, 1, 2 + len(d), len(d.columns) - 1, {"type": "3_color_scale", "min_type": "num", "min_value": -8, "mid_type": "num", "mid_value": 0, "max_type": "num", "max_value": 8, "min_color": "#F8696B", "mid_color": "#FFFFFF", "max_color": "#63BE7B"})

# --- basket + heatmap
ws = wb.add_worksheet("Basket_And_Windows"); wbw.sheets["Basket_And_Windows"] = ws; ws.hide_gridlines(2)
ws.merge_range(0, 0, 0, 22, "Event basket by year + window heat-maps (rows = lead days, cols = hold days)", f_title); ws.set_row(0, 30)
col = 0
for ev in ("Diwali", "Christmas"):
    ws.write(2, col, f"{ev} basket (equal-weight) at best window", f_b)
    write_df(ws, basket_tables[ev], 3, {"Year": 7, "T_date": 12, "Anchor": 26, "Stocks_With_Data": 10, "Basket_LONG_Ret_pct": 12, "Signed_Ret_pct": 12})
    ws.conditional_format(4, 4, 3 + 26, 5, {"type": "3_color_scale", "min_color": "#F8696B", "mid_color": "#FFFFFF", "max_color": "#63BE7B", "mid_type": "num", "mid_value": 0})
    # heatmaps below
    for gi, (gname, key, lo, mid, hi) in enumerate((("LONG avg return % (basket)", "avg", -2, 0, 2), ("LONG win rate % (basket)", "wr", 30, 50, 70))):
        r0 = 33 + gi * 14
        ws.write(r0, col, f"{ev}: {gname}", f_b)
        ws.write(r0 + 1, col, "lead \\ hold", f_h)
        for j, h in enumerate(HOLDS): ws.write(r0 + 1, col + 1 + j, f"T+{h}", f_h)
        for i, l in enumerate(LEADS):
            ws.write(r0 + 2 + i, col, f"T-{l}", f_h)
            for j in range(8):
                x = float(grids[ev][key][i, j])
                if np.isnan(x): ws.write_blank(r0 + 2 + i, col + 1 + j, None, f_cell)
                else: ws.write_number(r0 + 2 + i, col + 1 + j, x, f_num)
        ws.conditional_format(r0 + 2, col + 1, r0 + 11, col + 10, {"type": "3_color_scale", "min_type": "num", "min_value": lo, "mid_type": "num", "mid_value": mid, "max_type": "num", "max_value": hi, "min_color": "#F8696B", "mid_color": "#FFFFFF", "max_color": "#63BE7B"})
    col += 14

# --- trade plan
ws = wb.add_worksheet("Trade_Plan_2026"); wbw.sheets["Trade_Plan_2026"] = ws
ws.merge_range(0, 0, 0, 14, "2026 trade plan -- STRONG + MODERATE stocks only (review before any dashboard change)", f_title); ws.set_row(0, 30)
r = 2
for ev in ("Diwali", "Christmas"):
    p = plans[ev]
    ws.write(r, 0, f"{ev} 2026  (T = {ANCH[ev][2026][0].strftime('%d-%b-%Y')})  --  {len(p)} stocks", f_b)
    write_df(ws, p, r + 1, SW)
    cols = {n: i for i, n in enumerate(p.columns)}
    ws.conditional_format(r + 2, cols["Position"], r + 1 + len(p), cols["Position"], {"type": "cell", "criteria": "==", "value": '"LONG"', "format": green})
    ws.conditional_format(r + 2, cols["Position"], r + 1 + len(p), cols["Position"], {"type": "cell", "criteria": "==", "value": '"SHORT"', "format": red})
    ws.conditional_format(r + 2, cols["Conviction"], r + 1 + len(p), cols["Conviction"], {"type": "cell", "criteria": "==", "value": '"STRONG"', "format": bold_green})
    r += len(p) + 4
for c, w in {10: 18, 11: 18, 12: 14, 13: 13}.items(): ws.set_column(c, c, w)

# --- method / data
ws = wb.add_worksheet("Anchors_And_Coverage"); wbw.sheets["Anchors_And_Coverage"] = ws
ws.merge_range(0, 0, 0, 8, "Event anchor dates used (from holiday CSV) + data coverage", f_title); ws.set_row(0, 30)
arows = [dict(Year=y, Diwali_T=ANCH["Diwali"][y][0].strftime("%d-%b-%Y (%a)"), Diwali_Anchor=ANCH["Diwali"][y][1],
              Christmas_T=ANCH["Christmas"][y][0].strftime("%d-%b-%Y (%a)")) for y in YEARS + [2026]]
write_df(ws, pd.DataFrame(arows), 2, {"Year": 7, "Diwali_T": 20, "Diwali_Anchor": 34, "Christmas_T": 20})
cov = rep.reset_index()[["Symbol", "Status", "Rows", "FirstDate", "LastDate"]]
write_df(ws, cov, 2 + len(arows) + 3, {"Symbol": 14, "Status": 10, "Rows": 9, "FirstDate": 12, "LastDate": 12})

wbw.close()
print("written", OUT)
