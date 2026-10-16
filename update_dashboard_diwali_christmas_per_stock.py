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

def wd(d): return d.strftime("%d-%b-%Y (%A)")
def window_dates(ev_id, lead, hold):
    cal = [d for d in pd.bdate_range(T[ev_id] - pd.Timedelta(days=40), T[ev_id] + pd.Timedelta(days=40)) if d not in CLOSED and d != T[ev_id]]
    return [d for d in cal if d < T[ev_id]][-lead], [d for d in cal if d > T[ev_id]][hold - 1]

data = json.load(open(JS, encoding="utf-8"))
old = {h["id"]: h for h in data["holidays"]}
DRY = "--dry-run" in sys.argv
DRY_DIR = r"C:\Users\PC2551\AppData\Local\Temp\claude\d--behaviour-analysis\eeb2be21-45f7-49ef-9851-c810459970a8\scratchpad"
if not DRY:
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
        if r.Conviction == "DATA INSUFFICIENT":
            insufficient.append({"symbol": sym, "reason": f"Only {int(r.Valid_Years)} valid {ev} event year(s) in price history -- minimum 7 required; no position given."}); continue
        if r.Conviction == "NO TRADE":
            no_trade.append({"symbol": sym, "reason": f"{r.Position} best win rate {r.Win_Rate}% over {int(r.Valid_Years)} yrs but walk-forward {r.WF_OOS_WinRate}% -- below the bar; no position."}); continue
        if r.Position == "SHORT" and not (r.Valid_Years >= MIN_SHORT_YEARS and r.WF_OOS_WinRate >= MIN_SHORT_WF):
            short_rule.append({"symbol": sym, "reason": f"SHORT needs >={MIN_SHORT_YEARS} valid yrs and walk-forward >={MIN_SHORT_WF:.0f}%; has {int(r.Valid_Years)} yrs / WF {r.WF_OOS_WinRate}% -- no position."}); continue
        assert r.Conviction in ("STRONG", "MODERATE") and r.Position in ("LONG", "SHORT"), (sym, r.Conviction, r.Position)
        lead, hold = [int(x) for x in r.Window.replace("T-", "").replace("T+", "").split(" to ")]
        e, x = window_dates(ev_id, lead, hold)
        assert e < T[ev_id] < x
        y = M.loc[sym, [c for c in M.columns if str(c) in ("2022", "2023", "2024", "2025")]].astype(float).dropna()
        l4wr = round(float((y > 0).mean() * 100), 1) if len(y) else None
        l4av = round(float(y.mean()), 2) if len(y) else None
        lot = int(r.Lot); close = float(r.Last_Close)
        wr = float(r.Win_Rate); n = int(r.Valid_Years)
        verb, opt = ("BUY", "CALL") if r.Position == "LONG" else ("SELL", "PUT")
        stocks.append(dict(
            symbol=sym, name=name.get(sym) or r.Name, direction=r.Position, window=r.Window, lead=lead, hold=hold,
            entry_date=wd(e), exit_date=wd(x), total_years=n, full_19y_wr=round(wr, 1), full_19y_avg_ret=round(float(r.Avg_Signed_Ret), 2),
            last_4y_wr=l4wr, last_4y_avg_ret=l4av, lot=lot, margin=round(lot * close * 0.2, 0),
            fut_action=f"{verb} FUTURES ({r.Position} {lot} qty)", opt_action=f"BUY {x.strftime('%b').upper()} {opt} OPTION",
            sector=sector[sym],
            direction_basis=f"Per-stock best side + window over {n} valid yrs ({r.Window}): {r.Position} {wr:.1f}% win; walk-forward OOS {r.WF_OOS_WinRate}% ({int(r.WF_OOS_Trades)} trades); {r.Conviction}",
            long_win_rate_26y=round(wr if r.Position == "LONG" else 100 - wr, 1), short_win_rate_26y=round(wr if r.Position == "SHORT" else 100 - wr, 1), n_years=n,
            conviction=r.Conviction, wf_oos_win_rate=None if pd.isna(r.WF_OOS_WinRate) else float(r.WF_OOS_WinRate), wf_oos_trades=int(r.WF_OOS_Trades)))
    stocks.sort(key=lambda s: (s["conviction"] != "STRONG", -s["full_19y_wr"], -s["full_19y_avg_ret"]))
    nL = sum(s["direction"] == "LONG" for s in stocks); nS = len(stocks) - nL; nStrong = sum(s["conviction"] == "STRONG" for s in stocks)
    h = copy.deepcopy(old[ev_id])
    h["stocks"] = stocks; h["stocks_count"] = len(stocks)
    h["bias"] = f"{ev} -- per-stock best side & window (>=7 valid yrs, walk-forward checked): {nL} LONG / {nS} SHORT ({nStrong} STRONG)"
    h["top_pick"] = " & ".join(f"{s['symbol']} ({s['full_19y_wr']}% WR)" for s in stocks[:2])
    h["data_pending_symbols"] = insufficient
    h["no_trade_symbols"] = no_trade + short_rule
    data["holidays"] = [h if x["id"] == ev_id else x for x in data["holidays"]]
    report[ev] = dict(positions=len(stocks), LONG=nL, SHORT=nS, STRONG=nStrong, no_trade=len(no_trade), short_rule=[s["symbol"] for s in short_rule], insufficient=[s["symbol"] for s in insufficient])

for k in ("diwali_2026", "christmas_2026"):
    nh = [x for x in data["holidays"] if x["id"] == k][0]; assert set(nh) >= set(old[k]) and len(nh["stocks"]) > 0
out_text = json.dumps(data, indent=2, ensure_ascii=False, allow_nan=False) + "\n"
if DRY:
    open(DRY_DIR + r"\event_dashboard_data.dry.json", "w", encoding="utf-8", newline="\n").write(out_text)
    print("DRY RUN -> nothing in the repo was written")
else:
    open(JS, "w", encoding="utf-8", newline="\n").write(out_text)
print(json.dumps(report, indent=1))
