"""
Direct NSE check of Sep-quarter (FY27_Q3 table) result dates for every stock in the dashboard's Nifty table.

  python check_nifty_result_dates.py            # DRY RUN: report newly announced / changed dates, change nothing
  python check_nifty_result_dates.py --apply    # also update dashboard_data/event_dashboard_data.json (FY27_Q3 only)

Why this exists: the bulk NSE event-calendar feed used by fetch_nifty50_result_dates.py misses some filings (e.g. BSE) and drops
results once filed, so this asks NSE's per-symbol board-meetings API for each stock instead.

--apply touches ONLY stocks that are newly announced or whose date changed:
  * announced / result_date / entry+exit dates (NSE trading calendar) / counts are updated;
  * newly announced stocks get the best REACHABLE window (entry strictly after today) from the same walk-forward grid the
    project uses (fix_q3_position_window.py), on the stock's own prior quarterly results;
  * already-announced stocks whose date merely changed keep their side/window (only the dates move) and are flagged if the entry has passed.
Every other stock and every other section of the file is left untouched.  Run build_dashboard.py afterwards.
"""
import argparse, copy, http.cookiejar, json, re, sys, time, urllib.parse, urllib.request
from datetime import datetime, timedelta
import pandas as pd

sys.stdout.reconfigure(encoding="utf-8")
BASE = r"D:\behaviour analysis"
JS = BASE + r"\dashboard_data\event_dashboard_data.json"
HOL = {"2026-01-26", "2026-03-03", "2026-03-26", "2026-03-31", "2026-04-03", "2026-04-14", "2026-05-01", "2026-05-28", "2026-06-26", "2026-09-14",
       "2026-10-02", "2026-10-20", "2026-11-08", "2026-11-10", "2026-11-24", "2026-12-25", "2027-01-26"}   # same calendar fix_q3_position_window.py uses

ap = argparse.ArgumentParser(); ap.add_argument("--apply", action="store_true"); ap.add_argument("--today", default=None)
args = ap.parse_args()
TODAY = pd.Timestamp(args.today) if args.today else pd.Timestamp.now().normalize()

class NSE:
    def __init__(s):
        s.op = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar())); s.ok = False
        s.h = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36", "Accept": "*/*", "Accept-Language": "en-US,en;q=0.9"}
    def prime(s, sym):
        try:
            with s.op.open(urllib.request.Request("https://www.nseindia.com/get-quotes/equity?symbol=" + urllib.parse.quote(sym), headers=s.h), timeout=12) as r: r.read(); s.ok = True
        except Exception: pass
    def get(s, url, sym):
        if not s.ok: s.prime(sym)
        h = dict(s.h); h["Referer"] = "https://www.nseindia.com/get-quotes/equity?symbol=" + urllib.parse.quote(sym)
        err = None
        for _ in range(3):
            try:
                with s.op.open(urllib.request.Request(url, headers=h), timeout=15) as r: return json.loads(r.read().decode("utf-8", "ignore"))
            except Exception as e: err = e; s.prime(sym); time.sleep(0.6)
        return {"_error": str(err)}

def shift(dt, n):
    st = 1 if n >= 0 else -1; c = abs(n)
    while c:
        dt += timedelta(days=st)
        if dt.weekday() < 5 and dt.isoformat() not in HOL: c -= 1
    return dt

d = json.load(open(JS, encoding="utf-8"))
q = next(k for k in d["quarters"] if k["q_code"] == "FY27_Q3"); stocks = {s["symbol"]: s for s in q["stocks"]}
nse = NSE(); nse.prime("RELIANCE")
found, errs = {}, []
for sym in sorted(stocks):
    bm = nse.get("https://www.nseindia.com/api/corporate-board-meetings?index=equities&symbol=" + urllib.parse.quote(sym), sym)
    if not isinstance(bm, list): errs.append(sym); continue
    cand = []
    for b in bm:
        dt = pd.to_datetime(b.get("bm_date"), format="%d-%b-%Y", errors="coerce")
        txt = ((b.get("bm_purpose") or "") + " " + (b.get("bm_desc") or "")).lower()
        if pd.isna(dt) or dt < pd.Timestamp("2026-09-25"): continue
        if "financial result" in txt or "unaudited" in txt or "audited" in txt: cand.append((dt, b.get("bm_timestamp", ""), b.get("attachment", "")))
    if cand:
        dt, ts, att = min(cand); found[sym] = (dt, ts, att)
    time.sleep(0.25)

new, changed = [], []
for sym, (dt, ts, att) in found.items():
    s = stocks[sym]
    if not s.get("announced"): new.append(sym)
    elif s.get("result_date") != dt.strftime("%d-%b-%Y"): changed.append(sym)
pending = [s for s in sorted(stocks) if not stocks[s].get("announced") and s not in found]
print(f"as of {TODAY:%d-%b-%Y}: NSE lookups ok for {len(stocks) - len(errs)}/{len(stocks)} stocks" + (f" (errors: {errs})" if errs else ""))
print(f"dashboard announced: {sum(1 for s in stocks.values() if s.get('announced'))} of {len(stocks)}")
print("NEWLY ANNOUNCED:" + ("" if new else " none"))
for sym in new: print(f"  {sym:11s} result {found[sym][0]:%d-%b-%Y (%a)}   filed {found[sym][1]}   {found[sym][2]}")
print("DATE CHANGED:" + ("" if changed else " none"))
for sym in changed: print(f"  {sym:11s} {stocks[sym].get('result_date')} -> {found[sym][0]:%d-%b-%Y (%a)}   filed {found[sym][1]}")
print(f"STILL NOT ANNOUNCED ({len(pending)}): {', '.join(pending)}")
if not (new or changed): print("\nNothing to update."); sys.exit(0)
if not args.apply: print("\nDRY RUN - nothing changed. Re-run with --apply to update the dashboard data."); sys.exit(0)

# ----------------------------------------------------------------------------------------------- apply
fut = pd.read_parquet(BASE + r"\scraped_parquet\fo_futures_near_month_continuous.parquet", columns=["Date", "Instrument", "Close"])
fut["Date"] = pd.to_datetime(fut["Date"]).dt.date
hist = {}
for qq in d["quarters"]:
    if qq["q_code"] == "FY27_Q3": continue
    for s in qq["stocks"]:
        rd = s.get("result_date")
        if rd and rd != "Yet to come":
            try: hist.setdefault(s["symbol"], set()).add(datetime.strptime(rd, "%d-%b-%Y").date())
            except ValueError: pass

def ret(a, hd, bo, so):
    ib = -1
    for i, r in enumerate(a):
        if r[0] <= hd: ib = i
        else: break
    if ib < 0 or (hd - a[ib][0]).days > 15: return None
    en, ex = ib - bo, ib + so
    if en < 1 or ex >= len(a): return None
    g = 1.0
    for k in range(en, ex + 1):
        p0, p1 = a[k - 1][1], a[k][1]
        if not p0: return None
        dr = (p1 - p0) / p0
        if abs(dr) <= 0.20: g *= (1 + dr)
    return (g - 1) * 100

def best_reachable(sym, res):
    a = fut[fut.Instrument == sym].sort_values("Date")[["Date", "Close"]].values.tolist(); prior = sorted(hist.get(sym, []))
    if not a or len(prior) < 3: return None
    rows = []
    for side in ("LONG", "SHORT"):
        for bo in range(1, 9):
            if shift(res.date(), -bo) <= TODAY.date(): continue          # entry must be strictly after today
            for so in range(1, 9):
                r = [ret(a, h, bo, so) for h in prior]; r = [x for x in r if x is not None]; r = [x if side == "LONG" else -x for x in r]
                if len(r) >= 3: rows.append((sum(1 for x in r if x > 0) / len(r) * 100, sum(r) / len(r), side, bo, so, len(r)))
    return sorted(rows, reverse=True)[0] if rows else None

pre = copy.deepcopy(d); log = []
for sym in new + changed:
    s = stocks[sym]; res = found[sym][0]; is_new = sym in new
    bo, so, side, note = int(s.get("lead") or 2), int(s.get("hold") or 4), s["direction"], None
    if is_new:
        b = best_reachable(sym, res)
        if b:
            side, bo, so = b[2], b[3], b[4]
            note = f"Result date announced {found[sym][1]} (meeting {res:%d-%b-%Y}). Best REACHABLE window (entry after {TODAY:%d-%b-%Y}) from the walk-forward grid on {b[5]} prior quarters: {side} T-{bo} to T+{so}, win {b[0]:.1f}%, avg {b[1]:+.2f}%."
        else:
            note = f"Result date announced {found[sym][1]}; no prior quarterly history to pick a window - generic window kept (placeholder, not a backtested call)."
    en, ex = shift(res.date(), -bo), shift(res.date(), so)
    s.update(direction=side, window=f"T-{bo} to T+{so}", lead=bo, hold=so, entry_lead_days=bo, exit_hold_days=so, taking_window_raw=f"T-{bo} to T+{so}", pos_window=f"T-{bo} to T+{so}",
             result_date=res.strftime("%d-%b-%Y"), result_declaration_date=res.strftime("%d-%b-%Y"), entry_date=f"{en:%d-%b-%Y} ({en:%a})", exit_date=f"{ex:%d-%b-%Y} ({ex:%a})",
             announced=True, result_status="Announced", source=f"NSE board-meetings API (filed {found[sym][1]})")
    s["taking_window_full"] = f"Entry: {s['entry_date']} ({s['window']}) \u2794 Exit: {s['exit_date']}"
    if note: s["window_fix_note"] = note
    log.append((sym, "NEW" if is_new else "DATE CHANGED", s["direction"], s["window"], s["result_date"], s["entry_date"], s["exit_date"], "ENTRY ALREADY PASSED" if en <= TODAY.date() else ""))
q["announced_count"] = sum(1 for x in q["stocks"] if x.get("announced")); q["pending_count"] = len(q["stocks"]) - q["announced_count"]
open(JS, "w", encoding="utf-8", newline="\n").write(json.dumps(d, indent=2, ensure_ascii=False) + "\n")
print("\nAPPLIED:")
for r in log: print("  ", " | ".join(str(x) for x in r))
qp = lambda x: {t["symbol"]: t for t in next(k for k in x["quarters"] if k["q_code"] == "FY27_Q3")["stocks"]}
a, b2 = qp(pre), qp(d)
print("stocks changed in file:", [k for k in b2 if a[k] != b2[k]], "| rest of file identical:",
      all(pre[k] == d[k] for k in pre if k != "quarters") and all(x == y for x, y in zip([k for k in pre["quarters"] if k["q_code"] != "FY27_Q3"], [k for k in d["quarters"] if k["q_code"] != "FY27_Q3"])))
print(f"announced/pending now: {q['announced_count']}/{q['pending_count']}.  Next: python build_dashboard.py && python build_dashboard.py --check")
