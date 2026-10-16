"""
All 17 NSE/BSE holiday events x 213 F&O stocks x 2000-2026: per-stock best side + best window backtest
(same method as the corrected Diwali/Christmas v4 build) -> per-stock trade logs for the dashboard
"Stock Deep Dive -> Holiday Performance" sub-tab, plus a review workbook.

Outputs (all NEW files, nothing existing is touched):
  dashboard_data/holiday_trades/<SYMBOL>.json      one file per stock (lazy-loaded by the dashboard)
  dashboard_data/holiday_trades_index.json         holiday meta, date audit, exclusions, event-wide default windows
  All_Holidays_26Y_Report.xlsx                     review workbook (positions, date audit)

Method
* T = holiday date. Entry = close of T-lead, exit = close of T+hold, lead/hold in 1..10 trading days (real market calendar).
* Per stock: pick the (side, lead, hold) with the highest win rate over 2000-2025, ties -> higher avg signed return; needs >=7 valid years.
  A walk-forward check (re-pick each year using only earlier years) is stored beside every pick. 2026 completed events are
  appended as extra out-of-sample trades (never used for the pick).
* Stocks with <7 valid years get NO pick: they are shown at the holiday's event-wide window and flagged DATA INSUFFICIENT.
* Event dates: project record dates (nifty_futures_holiday_behaviour.json); Diwali = verified Laxmi Pujan table; Muharram 2018 = 20-Sep
  (market closed that day); Good Friday years that are missing are filled from the Easter calendar. A holiday-year is used ONLY if the
  market was actually closed that day (weekend / exchange closure / Muhurat low-volume day); years where the market traded normally
  are EXCLUDED and listed in the index + workbook.
"""
import json, os, warnings
import numpy as np, pandas as pd
from dateutil.easter import easter
warnings.filterwarnings("ignore")

ROOT = r"D:\behaviour analysis"
OUT_DIR = ROOT + r"\dashboard_data\holiday_trades"
OUT_IDX = ROOT + r"\dashboard_data\holiday_trades_index.json"
OUT_XL = ROOT + r"\All_Holidays_26Y_Report.xlsx"
YEARS_SEL = list(range(2000, 2026)); YEARS_ALL = YEARS_SEL + [2026]; NSEL = len(YEARS_SEL)
MAX_MOVE = 0.35; MIN_YEARS = 7; MIN_WF_PRIOR = 7

# ---------------------------------------------------------------- data
px = pd.read_parquet(f"{ROOT}/Stock_Data_213/all_stocks_daily_ohlc.parquet", columns=["Symbol", "Date", "Adj Close", "Volume"]).dropna().sort_values(["Symbol", "Date"])
lots = pd.read_csv(f"{ROOT}/Stock_Data_213/lot_sizes.csv").set_index("Symbol")["LotSize"]
meta = {r["symbol"]: r for r in json.load(open(f"{ROOT}/dashboard_data/fo_stocks_211.json", encoding="utf-8"))}
NIFTY50 = {k for k, v in meta.items() if str(v.get("is_nifty50", "")).upper() == "YES"}
hol_csv = pd.read_csv(f"{ROOT}/NSE_BSE_Historical_Trading_Holidays_2000_2026.csv", parse_dates=["Holiday_Date"])
rec = json.load(open(f"{ROOT}/dashboard_data/nifty_futures_holiday_behaviour.json", encoding="utf-8"))["holidays"]

_act = px.groupby("Date").agg(n=("Symbol", "size"), pos=("Volume", lambda v: (v > 0).sum()))
CAL = _act.index[(_act.n >= 5) & (_act.pos / _act.n >= 0.30)].values.astype("datetime64[D]")     # real market days
CAL_IDX = pd.DatetimeIndex(CAL); LAST = CAL_IDX.max()
_pv = px.pivot(index="Date", columns="Symbol", values="Volume")
REL = (_pv / _pv.rolling(20, min_periods=5).median().shift(1)).median(axis=1)                   # market-wide relative volume

# ---------------------------------------------------------------- event dates
LAXMI_PUJAN = {2000: "10-26", 2001: "11-14", 2002: "11-04", 2003: "10-25", 2004: "11-12", 2005: "11-01", 2006: "10-21", 2007: "11-09",
               2008: "10-28", 2009: "10-17", 2010: "11-05", 2011: "10-26", 2012: "11-13", 2013: "11-03", 2014: "10-23", 2015: "11-11",
               2016: "10-30", 2017: "10-19", 2018: "11-07", 2019: "10-27", 2020: "11-14", 2021: "11-04", 2022: "10-24", 2023: "11-12",
               2024: "11-01", 2025: "10-21", 2026: "11-08"}
CSV_PAT = {"republic": "Republic", "mahashivratri": "Mahashivratri", "holi": "Holi", "ramnavami": "Ram Navami", "mahavir": "Mahavir",
           "goodfriday": "Good Friday", "ambedkar": "Ambedkar", "maharashtra": "Maharashtra Day", "bakriid": "Bakri", "muharram": "Muharram",
           "independence": "Independence", "ganesh": "Ganesh", "gandhi": "Gandhi", "dussehra": "Dussehra", "diwali": "Diwali",
           "gurunanak": "Gurunanak", "christmas": "Christmas"}
HOL = []        # [{id,name,dates:{year:Timestamp},src:{year:str}}]
for h in rec:
    dates, src = {}, {}
    for t in h["trades"]:
        T = pd.Timestamp(t["holiday_date"]); dates[T.year] = T; src[T.year] = "project record"
    for y in YEARS_ALL:                                   # fill 2026 (and any gaps) from the holiday CSV
        if y not in dates and h["id"] not in ("diwali",):
            g = hol_csv[(hol_csv.Year == y) & hol_csv.Holiday_Name.str.contains(CSV_PAT[h["id"]], case=False)]
            if y == 2026 and len(g): dates[y] = g.iloc[0].Holiday_Date; src[y] = "holiday CSV (2026)"
    if h["id"] == "goodfriday":
        for y in YEARS_ALL:
            gf = pd.Timestamp(easter(y)) - pd.Timedelta(days=2)
            if y not in dates: dates[y] = gf; src[y] = "Easter calendar"
            assert dates[y] == gf, ("Good Friday mismatch", y, dates[y], gf)
    if h["id"] == "diwali":
        for y in YEARS_ALL: dates[y] = pd.Timestamp(f"{y}-{LAXMI_PUJAN[y]}"); src[y] = "verified Laxmi Pujan table"
    if h["id"] == "muharram" and 2018 in dates:
        dates[2018] = pd.Timestamp("2018-09-20"); src[2018] = "corrected: market closed 20-Sep-2018 (record had 21-Sep, a trading day)"
    HOL.append(dict(id=h["id"], name=h["name"], dates=dates, src=src))

def day_status(T):
    if T > LAST: return "future"
    if T.weekday() >= 5: return "weekend"
    if T not in CAL_IDX: return "closed"
    return "muhurat" if REL.get(T, 1) < 0.6 else "TRADED"

AUDIT, VALID = [], {}
for H in HOL:
    VALID[H["id"]] = {}
    for y in YEARS_ALL:
        if y not in H["dates"]: continue
        T = H["dates"][y]; st = day_status(T)
        use = st in ("weekend", "closed", "muhurat") or (st == "future" and False)
        if st == "TRADED": reason = "EXCLUDED: market traded normally on this date (not an exchange holiday that year)"
        elif st == "future": reason = "future event - not backtested"
        else: reason = "used (" + st + ")"
        AUDIT.append(dict(Holiday=H["name"], Year=y, Date=T.strftime("%Y-%m-%d"), Weekday=T.strftime("%a"), Market=st, Source=H["src"][y], Status=reason))
        if use: VALID[H["id"]][y] = T

# ---------------------------------------------------------------- cube
def build(dates, adj, Tmap):
    """R[year,lead,hold] long % return, NaN when unusable. Vectorised over (lead,hold)."""
    n = len(dates); R = np.full((len(YEARS_ALL), 10, 10), np.nan)
    mv = np.zeros(n); mv[1:] = np.abs(adj[1:] / adj[:-1] - 1) > MAX_MOVE; bad = np.cumsum(mv)
    ar = np.arange(10)
    for yi, y in enumerate(YEARS_ALL):
        if y not in Tmap: continue
        T = np.datetime64(Tmap[y], "D")
        ipre = np.searchsorted(dates, T, "left") - 1; ipost = np.searchsorted(dates, T, "right")
        if ipre < 0 or ipost >= n: continue
        if (T - dates[ipre]).astype(int) > 7 or (dates[ipost] - T).astype(int) > 7: continue
        E, X = np.meshgrid(ipre - ar, ipost + ar, indexing="ij")
        ok = (E >= 0) & (X < n)
        Ec, Xc = np.clip(E, 0, n - 1), np.clip(X, 0, n - 1)
        ok &= ((dates[Xc] - dates[Ec]).astype(int) <= 40) & ((bad[Xc] - bad[Ec]) == 0) & (Xc > Ec)
        R[yi] = np.where(ok, (adj[Xc] / adj[Ec] - 1) * 100, np.nan)
    return R

def select(C):
    valid = ~np.isnan(C); nn = valid.sum(0)
    with np.errstate(all="ignore"):
        wl = ((C > 0) & valid).sum(0) / nn; ws = ((C < 0) & valid).sum(0) / nn; avg = np.nanmean(C, 0)
    sl = np.where(nn >= MIN_YEARS, wl + 1e-6 * avg, -1); ss = np.where(nn >= MIN_YEARS, ws - 1e-6 * avg, -1)
    bl, bs = np.nanmax(sl), np.nanmax(ss)
    if max(bl, bs) < 0: return None
    if bl >= bs:
        i, j = np.unravel_index(np.nanargmax(sl), sl.shape); return ("LONG", i, j, wl[i, j] * 100, avg[i, j], int(nn[i, j]))
    i, j = np.unravel_index(np.nanargmax(ss), ss.shape); return ("SHORT", i, j, ws[i, j] * 100, -avg[i, j], int(nn[i, j]))

symbols = sorted(s for s in px.Symbol.unique() if s in lots.index)
SER = {}
for s in symbols:
    g = px[px.Symbol == s].drop_duplicates("Date").set_index("Date")["Adj Close"]; g.index = g.index.values.astype("datetime64[D]")
    g = g[g.index.isin(CAL)]; cs = CAL[(CAL >= g.index.min()) & (CAL <= g.index.max())]
    SER[s] = (cs, g.reindex(cs).ffill().values.astype(float))

per_stock = {s: [] for s in symbols}
index_h = []; rows_x = []
for H in HOL:
    hid = H["id"]; Tmap = VALID[hid]
    cubes = {}
    for s in symbols:
        dts, a = SER[s]
        dts_h, a_h = dts, a
        if any(np.datetime64(T, "D") in set(dts) for T in Tmap.values()):          # drop Muhurat-type rows that sit exactly on T
            keep = ~np.isin(dts, [np.datetime64(T, "D") for T in Tmap.values()]); dts_h, a_h = dts[keep], a[keep]
        cubes[s] = (build(dts_h, a_h, Tmap), dts_h, a_h)
    # event-wide default window (equal-weight basket, best |t-stat|) -> placeholder for <7-year stocks
    B = np.nanmean(np.stack([cubes[s][0][:NSEL] for s in symbols]), axis=0)
    m = np.nanmean(B, 0); sd = np.nanstd(B, 0, ddof=1); nn = np.sum(~np.isnan(B), 0)
    with np.errstate(all="ignore"): tt = m / (sd / np.sqrt(nn))
    tt = np.where(nn >= MIN_YEARS, tt, np.nan)
    di, dj = np.unravel_index(np.nanargmax(np.abs(tt)), tt.shape); dside = "LONG" if m[di, dj] > 0 else "SHORT"
    bw = B[:, di, dj]; bw = bw[~np.isnan(bw)]
    index_h.append(dict(id=hid, name=H["name"], default_side=dside, default_window=f"T-{di + 1} to T+{dj + 1}", basket_win_rate=round(float(((np.sign(m[di, dj]) * bw) > 0).mean() * 100), 1),
                        basket_years=int(len(bw)), years_used=sorted(y for y in Tmap if y <= 2025), years_2026=bool(2026 in Tmap),
                        years_excluded=[a["Year"] for a in AUDIT if a["Holiday"] == H["name"] and a["Status"].startswith("EXCLUDED")]))
    for s in symbols:
        C_all, dts_h, a_h = cubes[s]; C = C_all[:NSEL]
        sel = select(C); flag = "OK"
        if sel is None:
            side, li, hi = dside, di, dj; flag = "DATA INSUFFICIENT"
            v = C[:, li, hi]; ok = ~np.isnan(v); sg = 1 if side == "LONG" else -1
            wr = float(np.mean(sg * v[ok] > 0) * 100) if ok.any() else 0.0; avg = float(np.mean(sg * v[ok])) if ok.any() else 0.0; n = int(ok.sum())
            conv = "DATA INSUFFICIENT"; wf_wr = None; wf_n = 0
        else:
            side, li, hi, wr, avg, n = sel; sg = 1 if side == "LONG" else -1
            oos = []
            for yi in range(NSEL):
                if np.isnan(C[yi]).all() or yi < MIN_WF_PRIOR: continue
                p = select(C[:yi])
                if p is None: continue
                r = (1 if p[0] == "LONG" else -1) * C[yi, p[1], p[2]]
                if not np.isnan(r): oos.append(r)
            oos = np.array(oos); wf_n = len(oos); wf_wr = round(float(np.mean(oos > 0) * 100), 1) if wf_n else None
            if n >= 10 and wr >= 75 and avg > 0 and wf_n >= 5 and wf_wr >= 60: conv = "STRONG"
            elif wr >= 65 and avg > 0 and (wf_n < 5 or wf_wr >= 50): conv = "MODERATE"
            else: conv = "NO TRADE"
            wr, avg = float(wr), float(avg)
        # trade log (2000-2026) at the chosen window
        trades = []
        dts_np = dts_h
        for yi, y in enumerate(YEARS_ALL):
            r = C_all[yi, li, hi]
            if np.isnan(r): continue
            T = np.datetime64(Tmap[y], "D"); ip = np.searchsorted(dts_np, T, "left") - 1; iq = np.searchsorted(dts_np, T, "right")
            e, x = ip - li, iq + hi
            trades.append([y, str(T), str(dts_np[e]), str(dts_np[x]), round(float(a_h[e]), 2), round(float(a_h[x]), 2), round(float(r), 2), round(float(sg * r), 2)])
        per_stock[s].append(dict(id=hid, name=H["name"], pos=side, win=f"T-{li + 1} to T+{hi + 1}", lead=int(li + 1), hold=int(hi + 1), conv=conv, flag=flag,
                                 wr=round(wr, 1), avg=round(avg, 2), yrs=n, wf=wf_wr, wft=wf_n, t=trades))
        rows_x.append(dict(Holiday=H["name"], Symbol=s, Group="NIFTY 50" if s in NIFTY50 else "F&O", Position="NO CALL" if flag != "OK" else side,
                           Window=f"T-{li + 1} to T+{hi + 1}", Rating=conv, Win_Rate=round(wr, 1), Avg_Signed_Ret=round(avg, 2), Valid_Years=n, WF_Win_Rate=wf_wr, WF_Trades=wf_n,
                           Note="event-wide placeholder window" if flag != "OK" else ""))
    c = pd.Series([r["conv"] for s in symbols for r in per_stock[s] if r["id"] == hid]).value_counts().to_dict()
    print(f"{H['name'][:28]:28s} years used {len([y for y in Tmap if y <= 2025]):2d} (+2026:{2026 in Tmap}) default {dside} {index_h[-1]['default_window']:10s} {c}")

# ---------------------------------------------------------------- write dashboard files
os.makedirs(OUT_DIR, exist_ok=True)
for s in symbols:
    doc = dict(s=s, n=meta.get(s, {}).get("name", s), lot=None if pd.isna(lots[s]) else int(lots[s]), nifty50=s in NIFTY50, h=per_stock[s])
    with open(os.path.join(OUT_DIR, s.replace("&", "_AND_") + ".json"), "w", encoding="utf-8", newline="\n") as f:
        json.dump(doc, f, separators=(",", ":"), ensure_ascii=False, allow_nan=False); f.write("\n")
idx = dict(generated="2026-10-16", years="2000-2025 for the pick; 2026 completed events appended as out-of-sample", min_years=MIN_YEARS,
           holidays=index_h, symbols=[dict(s=s, n=meta.get(s, {}).get("name", s), nifty50=s in NIFTY50, file=s.replace("&", "_AND_") + ".json") for s in symbols],
           date_audit=AUDIT)
with open(OUT_IDX, "w", encoding="utf-8", newline="\n") as f: json.dump(idx, f, separators=(",", ":"), ensure_ascii=False, allow_nan=False); f.write("\n")
print("stock files:", len(symbols), "| total size MB:", round(sum(os.path.getsize(os.path.join(OUT_DIR, x)) for x in os.listdir(OUT_DIR)) / 1e6, 1))

# ---------------------------------------------------------------- review workbook
X = pd.DataFrame(rows_x); A = pd.DataFrame(AUDIT)
with pd.ExcelWriter(OUT_XL, engine="xlsxwriter") as w:
    wb = w.book
    hd = wb.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#2F5597", "border": 1, "text_wrap": True, "valign": "vcenter", "align": "center"})
    green = wb.add_format({"bg_color": "#C6EFCE", "font_color": "#006100"}); red = wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})
    amber = wb.add_format({"bg_color": "#FFEB9C", "font_color": "#9C5700"}); grey = wb.add_format({"bg_color": "#BFBFBF", "font_color": "#404040", "bold": True})
    strong = wb.add_format({"bg_color": "#63BE7B", "bold": True})
    def sheet(df, name, widths=None):
        df.to_excel(w, sheet_name=name, index=False, startrow=0); ws = w.sheets[name]
        for c, col in enumerate(df.columns): ws.write(0, c, col.replace("_", " "), hd); ws.set_column(c, c, (widths or {}).get(col, max(10, min(30, len(col) + 3))))
        ws.freeze_panes(1, 0); ws.autofilter(0, 0, len(df), len(df.columns) - 1); return ws
    summ = pd.DataFrame([dict(Holiday=h["name"], Default_Side=h["default_side"], Default_Window=h["default_window"], Basket_Win_Rate=h["basket_win_rate"], Years_Used=len(h["years_used"]),
                              Years_Excluded=", ".join(map(str, h["years_excluded"])) or "-", Includes_2026=h["years_2026"],
                              STRONG=int(((X.Holiday == h["name"]) & (X.Rating == "STRONG")).sum()), MODERATE=int(((X.Holiday == h["name"]) & (X.Rating == "MODERATE")).sum()),
                              NO_TRADE=int(((X.Holiday == h["name"]) & (X.Rating == "NO TRADE")).sum()), DATA_INSUFFICIENT=int(((X.Holiday == h["name"]) & (X.Rating == "DATA INSUFFICIENT")).sum())) for h in index_h])
    sheet(summ, "Summary", {"Holiday": 34, "Years_Excluded": 38})
    ws = sheet(X, "All_Positions", {"Holiday": 34, "Symbol": 14, "Note": 30}); n = len(X); cols = list(X.columns)
    for val, fm in (("LONG", green), ("SHORT", red), ("NO CALL", grey)):
        ws.conditional_format(1, cols.index("Position"), n, cols.index("Position"), {"type": "cell", "criteria": "==", "value": f'"{val}"', "format": fm})
    for val, fm in (("STRONG", strong), ("MODERATE", amber), ("DATA INSUFFICIENT", grey)):
        ws.conditional_format(1, cols.index("Rating"), n, cols.index("Rating"), {"type": "cell", "criteria": "==", "value": f'"{val}"', "format": fm})
    sheet(A, "Date_Audit", {"Holiday": 34, "Source": 44, "Status": 70})
    for h in index_h:
        sub = X[X.Holiday == h["name"]].drop(columns=["Holiday"]).copy()
        sub["_o"] = sub.Rating.map({"STRONG": 0, "MODERATE": 1, "NO TRADE": 2, "DATA INSUFFICIENT": 3}); sub = sub.sort_values(["_o", "Win_Rate"], ascending=[True, False]).drop(columns="_o")
        nm = h["id"][:28]; ws = sheet(sub, nm, {"Symbol": 14}); cc = list(sub.columns)
        for val, fm in (("LONG", green), ("SHORT", red), ("NO CALL", grey)):
            ws.conditional_format(1, cc.index("Position"), len(sub), cc.index("Position"), {"type": "cell", "criteria": "==", "value": f'"{val}"', "format": fm})
        for val, fm in (("STRONG", strong), ("MODERATE", amber), ("DATA INSUFFICIENT", grey)):
            ws.conditional_format(1, cc.index("Rating"), len(sub), cc.index("Rating"), {"type": "cell", "criteria": "==", "value": f'"{val}"', "format": fm})
print("workbook:", OUT_XL)
