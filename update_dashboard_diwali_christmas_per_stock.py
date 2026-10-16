"""
Replace the Diwali + Christmas 2026 stock lists in dashboard_data/event_dashboard_data.json with the
per-stock best-side / best-window results from Nifty50_Diwali_Christmas_Report_v3.xlsx (Nifty 50 only,
same scope as the dashboard's existing holiday tables).  Gandhi, Dussehra, quarters, RBI untouched.

* Old Diwali/Christmas entries are first saved to dashboard_data/diwali_christmas_pre_per_stock_backup.json (additive).
* Only STRONG + MODERATE stocks become positions.  SHORT additionally needs >=15 valid yrs and walk-forward >=60%.
* NO TRADE / SHORT-rule / data-insufficient stocks are listed in side lists, never as positions.
Run build_dashboard.py afterwards to regenerate the .js bundle and sync docs/.
"""
import json, copy, sys
import numpy as np, pandas as pd

ROOT = r"D:\behaviour analysis"
XL = ROOT + r"\Nifty50_Diwali_Christmas_Report_v3.xlsx"
JS = ROOT + r"\dashboard_data\event_dashboard_data.json"
BK = ROOT + r"\dashboard_data\diwali_christmas_pre_per_stock_backup.json"
HOL = pd.read_csv(ROOT + r"\NSE_BSE_Historical_Trading_Holidays_2000_2026.csv", parse_dates=["Holiday_Date"])
CLOSED = set(HOL[HOL.Is_Trading_Holiday == "YES"].Holiday_Date) | {pd.Timestamp("2026-11-10")}
T = {"diwali_2026": pd.Timestamp("2026-11-08"), "christmas_2026": pd.Timestamp("2026-12-25")}
EV = {"diwali_2026": "Diwali", "christmas_2026": "Christmas"}
MIN_SHORT_YEARS, MIN_SHORT_WF = 15, 60.0

DEFAULT_SIDE = "LONG"      # event-wide verdict from the 213-stock basket study (both events LONG)
DEFAULT_WIN = {"diwali_2026": ("T-2 to T+2",), "christmas_2026": ("T-1 to T+7",)}
YEARS = list(range(2000, 2026))
_px = pd.read_parquet(ROOT + r"\Stock_Data_213\all_stocks_daily_ohlc.parquet", columns=["Symbol", "Date", "Adj Close"]).dropna()

def anchor(ev_id, y):
    if ev_id == "christmas_2026": return pd.Timestamp(y, 12, 25)
    g = HOL[HOL.Holiday_Name.str.contains("Diwali", case=False) & (HOL.Year == y)]; lp = g[g.Is_Trading_Holiday == "SPECIAL_SESSION"]
    return (lp.iloc[0] if len(lp) else g.iloc[0]).Holiday_Date

def default_stats(sym, ev_id):
    """LONG stats of one stock at the event-wide default window (same rules as the backtest: T row dropped, 7-day gap, 40-day span, 35% daily-move guard)."""
    lead, hold = [int(v) for v in DEFAULT_WIN[ev_id][0].replace("T-", "").replace("T+", "").split(" to ")]
    g = _px[_px.Symbol == sym].sort_values("Date"); d = g.Date.values.astype("datetime64[D]"); a = g["Adj Close"].values.astype(float)
    if ev_id == "diwali_2026":
        keep = ~np.isin(d, [np.datetime64(anchor(ev_id, y), "D") for y in YEARS]); d, a = d[keep], a[keep]
    out = {}
    for y in YEARS:
        Tn = np.datetime64(anchor(ev_id, y), "D"); ip = np.searchsorted(d, Tn, "left") - 1; iq = np.searchsorted(d, Tn, "right")
        if ip < 0 or iq >= len(d) or (Tn - d[ip]).astype(int) > 7 or (d[iq] - Tn).astype(int) > 7: continue
        e, x = ip - (lead - 1), iq + (hold - 1)
        if e < 0 or x >= len(d) or (d[x] - d[e]).astype(int) > 40: continue
        seg = a[e:x + 1]
        if len(seg) < 2 or np.max(np.abs(np.diff(seg) / seg[:-1])) > 0.35: continue
        out[y] = (a[x] / a[e] - 1) * 100
    v = np.array(list(out.values())); l4 = np.array([out[y] for y in (2022, 2023, 2024, 2025) if y in out])
    return dict(n=len(v), wr=float((v > 0).mean() * 100) if len(v) else 0.0, avg=float(v.mean()) if len(v) else 0.0,
                l4wr=round(float((l4 > 0).mean() * 100), 1) if len(l4) else None, l4av=round(float(l4.mean()), 2) if len(l4) else None)

def wd(d): return d.strftime("%d-%b-%Y (%A)")
def window_dates(ev_id, lead, hold):
    cal = [d for d in pd.bdate_range(T[ev_id] - pd.Timedelta(days=40), T[ev_id] + pd.Timedelta(days=40)) if d not in CLOSED and d != T[ev_id]]
    return [d for d in cal if d < T[ev_id]][-lead], [d for d in cal if d > T[ev_id]][hold - 1]

data = json.load(open(JS, encoding="utf-8"))
old = {h["id"]: h for h in data["holidays"]}
DRY = "--dry-run" in sys.argv
DRY_DIR = r"C:\Users\PC2551\AppData\Local\Temp\claude\d--behaviour-analysis\eeb2be21-45f7-49ef-9851-c810459970a8\scratchpad"
import os
if not DRY and not os.path.exists(BK):      # never overwrite the original pre-change backup on a re-run
    json.dump({k: old[k] for k in T}, open(BK, "w", encoding="utf-8"), indent=1)   # backup (new file)

# lookup tables from the existing dashboard (names, sectors)
sector, name = {}, {}
for sec in ("holidays", "quarters"):
    for blk in data[sec]:
        for s in blk.get("stocks", []):
            if s.get("sector"): sector.setdefault(s["symbol"], s["sector"])
            if s.get("name"): name.setdefault(s["symbol"], s["name"])

report = {}
for ev_id, ev in EV.items():
    P = pd.read_excel(XL, sheet_name=f"{ev}_Positions", header=2)
    M = pd.read_excel(XL, sheet_name=f"{ev}_YearMatrix", header=2).set_index("Symbol")
    assert len(P) == 50, (ev, len(P))
    stocks, no_trade, insufficient, short_rule = [], [], [], []
    P.columns = [c.replace(" ", "_") for c in P.columns]
    for _, r in P.iterrows():
        sym = r.Symbol
        flag, flag_note = None, None
        if r.Conviction == "DATA INSUFFICIENT":
            insufficient.append({"symbol": sym, "reason": f"Only {int(r.Valid_Years)} valid {ev} event year(s) in price history -- minimum 7 required; shown with the event-wide side/window as a placeholder, NOT a backtested call."})
            flag = "DATA INSUFFICIENT"
            flag_note = f"DATA INSUFFICIENT: only {int(r.Valid_Years)} valid {ev} year(s) (<7). Event-wide default {DEFAULT_SIDE} {DEFAULT_WIN[ev_id][0]} shown as a placeholder -- not a per-stock backtested call."
        elif r.Conviction == "NO TRADE":
            no_trade.append({"symbol": sym, "reason": f"{r.Position} best win rate {r.Win_Rate}% over {int(r.Valid_Years)} yrs but walk-forward {r.WF_OOS_WinRate}% -- below the bar; shown as LOW CONFIDENCE."})
            flag = "LOW CONFIDENCE"
            flag_note = f"LOW CONFIDENCE (NO TRADE grade): own best {r.Position} {r.Window} scored {r.Win_Rate}% in-sample but only {r.WF_OOS_WinRate}% walk-forward."
        elif r.Position == "SHORT" and not (r.Valid_Years >= MIN_SHORT_YEARS and r.WF_OOS_WinRate >= MIN_SHORT_WF):
            short_rule.append({"symbol": sym, "reason": f"SHORT needs >={MIN_SHORT_YEARS} valid yrs and walk-forward >={MIN_SHORT_WF:.0f}%; has {int(r.Valid_Years)} yrs / WF {r.WF_OOS_WinRate}% -- shown as LOW CONFIDENCE."})
            flag = "LOW CONFIDENCE"
            flag_note = f"LOW CONFIDENCE (SHORT rule not met): needs >={MIN_SHORT_YEARS} yrs and walk-forward >={MIN_SHORT_WF:.0f}%; has {int(r.Valid_Years)} yrs / WF {r.WF_OOS_WinRate}%."
        else:
            assert r.Conviction in ("STRONG", "MODERATE") and r.Position in ("LONG", "SHORT"), (sym, r.Conviction, r.Position)
        if flag == "DATA INSUFFICIENT":
            st = default_stats(sym, ev_id)
            pos, win_s, n_valid, wr_v, avg_v, l4wr_v, l4av_v = DEFAULT_SIDE, DEFAULT_WIN[ev_id][0], st["n"], st["wr"], st["avg"], st["l4wr"], st["l4av"]
            wf_wr, wf_tr, conv_v = None, 0, "DATA INSUFFICIENT"
        else:
            pos, win_s, n_valid, wr_v, avg_v = r.Position, r.Window, int(r.Valid_Years), float(r.Win_Rate), float(r.Avg_Signed_Ret)
            wf_wr = None if pd.isna(r.WF_OOS_WinRate) else float(r.WF_OOS_WinRate); wf_tr = int(r.WF_OOS_Trades); conv_v = r.Conviction
        lead, hold = [int(x) for x in win_s.replace("T-", "").replace("T+", "").split(" to ")]
        e, x = window_dates(ev_id, lead, hold)
        assert e < T[ev_id] < x
        if flag == "DATA INSUFFICIENT":
            l4wr, l4av = st["l4wr"], st["l4av"]
        else:
            y = M.loc[sym, [c for c in M.columns if str(c) in ("2022", "2023", "2024", "2025")]].astype(float).dropna()
            l4wr = round(float((y > 0).mean() * 100), 1) if len(y) else None
            l4av = round(float(y.mean()), 2) if len(y) else None
        lot = int(r.Lot); close = float(r.Last_Close)
        wr, n = wr_v, n_valid
        verb, opt = ("BUY", "CALL") if pos == "LONG" else ("SELL", "PUT")
        basis = flag_note or f"Per-stock best side + window over {n} valid yrs ({win_s}): {pos} {wr:.1f}% win; walk-forward OOS {wf_wr}% ({wf_tr} trades); {conv_v}"
        stocks.append(dict(
            symbol=sym, name=name.get(sym) or r.Name, direction=pos, window=win_s, lead=lead, hold=hold,
            entry_date=wd(e), exit_date=wd(x), total_years=n, full_19y_wr=round(wr, 1), full_19y_avg_ret=round(avg_v, 2),
            last_4y_wr=l4wr, last_4y_avg_ret=l4av, lot=lot, margin=round(lot * close * 0.2, 0),
            fut_action=f"{verb} FUTURES ({pos} {lot} qty)", opt_action=f"BUY {x.strftime('%b').upper()} {opt} OPTION",
            sector=sector[sym], direction_basis=basis,
            long_win_rate_26y=round(wr if pos == "LONG" else 100 - wr, 1), short_win_rate_26y=round(wr if pos == "SHORT" else 100 - wr, 1), n_years=n,
            conviction=conv_v, wf_oos_win_rate=wf_wr, wf_oos_trades=wf_tr, confidence_flag=flag or "OK", confidence_note=flag_note or ""))
    rank = {"STRONG": 0, "MODERATE": 1, "NO TRADE": 2, "DATA INSUFFICIENT": 3}
    stocks.sort(key=lambda s: (rank[s["conviction"]], -s["full_19y_wr"], -s["full_19y_avg_ret"]))
    core = [s for s in stocks if s["confidence_flag"] == "OK"]; flagged = [s for s in stocks if s["confidence_flag"] != "OK"]
    nL = sum(s["direction"] == "LONG" for s in core); nS = len(core) - nL; nStrong = sum(s["conviction"] == "STRONG" for s in core)
    h = copy.deepcopy(old[ev_id])
    h["stocks"] = stocks; h["stocks_count"] = len(stocks)
    h["bias"] = (f"{ev} -- per-stock best side & window (walk-forward checked): {len(core)} positions ({nL} LONG / {nS} SHORT, {nStrong} STRONG) "
                 f"+ {len(flagged)} flagged LOW CONFIDENCE / DATA INSUFFICIENT shown at the bottom")
    h["top_pick"] = " & ".join(f"{s['symbol']} ({s['full_19y_wr']}% WR)" for s in core[:2])
    h["data_pending_symbols"] = insufficient
    h["no_trade_symbols"] = no_trade + short_rule
    data["holidays"] = [h if x["id"] == ev_id else x for x in data["holidays"]]
    report[ev] = dict(listed=len(stocks), positions=len(core), flagged=[(s["symbol"], s["direction"], s["window"], s["confidence_flag"]) for s in flagged], LONG=nL, SHORT=nS, STRONG=nStrong, no_trade=len(no_trade), short_rule=[s["symbol"] for s in short_rule], insufficient=[s["symbol"] for s in insufficient])

for k in ("diwali_2026", "christmas_2026"):
    nh = [x for x in data["holidays"] if x["id"] == k][0]; assert set(nh) >= set(old[k]) and len(nh["stocks"]) > 0
out_text = json.dumps(data, indent=2, ensure_ascii=False, allow_nan=False) + "\n"
if DRY:
    open(DRY_DIR + r"\event_dashboard_data.dry.json", "w", encoding="utf-8", newline="\n").write(out_text)
    print("DRY RUN -> nothing in the repo was written")
else:
    open(JS, "w", encoding="utf-8", newline="\n").write(out_text)
print(json.dumps(report, indent=1))
