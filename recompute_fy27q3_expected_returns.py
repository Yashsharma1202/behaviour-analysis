"""
recompute_fy27q3_expected_returns.py
===============================================================================
expected_ret for FY27_Q3 stocks was found stale for INFY (computed once,
never refreshed when that stock's window changed away from the generic
T-2/T+4). This recomputes expected_ret / exp_pnl_1lot for EVERY FY27_Q3
stock from its own prior-quarters result dates, replayed through its
CURRENT window and direction against real futures price data -- same
methodology already verified for INFY, extended to all 51.

Does NOT touch: direction, window, lead, hold, entry_date, exit_date,
entry_px, outcome, actual_pnl. Window/direction are left exactly as-is
per explicit instruction -- only the return/P&L forecast fields change.
"""
import json
from datetime import datetime
import pandas as pd

PATH = "dashboard_data/event_dashboard_data.json"
MAX_GAP_DAYS = 15


def parse_date(s):
    if not s or s in ("Yet to come", "Awaiting NSE Announcement"):
        return None
    core = s.split(" (")[0].strip()
    if len(core) == 10 and core[4] == "-" and core[7] == "-":
        return datetime.strptime(core, "%Y-%m-%d").date()
    return datetime.strptime(core, "%d-%b-%Y").date()


def ret_for_window(a, hd, bo, so):
    ib = -1
    for i, row in enumerate(a):
        if row[0] <= hd:
            ib = i
        else:
            break
    if ib < 0 or (hd - a[ib][0]).days > MAX_GAP_DAYS:
        return None
    en, ex = ib - bo, ib + so
    if en < 1 or ex >= len(a):
        return None
    g = 1.0
    for k in range(en, ex + 1):
        p0, p1 = a[k - 1][1], a[k][1]
        if not p0:
            return None
        dr = (p1 - p0) / p0
        if abs(dr) <= 0.20:
            g *= (1 + dr)
    return (g - 1) * 100


def parse_window(window_str, lead_fallback, hold_fallback):
    try:
        lead = int(window_str.split("T-")[1].split(" ")[0])
        hold = int(window_str.split("T+")[1])
        return lead, hold
    except Exception:
        return lead_fallback, hold_fallback


def main():
    with open(PATH, encoding="utf-8") as f:
        data = json.load(f)

    fut = pd.read_parquet(
        "scraped_parquet/fo_futures_near_month_continuous.parquet",
        columns=["Date", "Instrument", "Close"],
    )
    fut["Date"] = pd.to_datetime(fut["Date"]).dt.date
    fut_by_instr = {
        sym: g.sort_values("Date")[["Date", "Close"]].values.tolist()
        for sym, g in fut.groupby("Instrument")
    }

    q3 = next(q for q in data["quarters"] if q["q_code"] == "FY27_Q3")

    hist_dates = {}
    for q in data["quarters"]:
        if q["q_code"] == "FY27_Q3":
            continue
        for s in q["stocks"]:
            rd = parse_date(s.get("result_date"))
            if rd:
                hist_dates.setdefault(s["symbol"], set()).add(rd)

    updated, no_data, insufficient = [], [], []

    for s in q3["stocks"]:
        sym = s["symbol"]
        a = fut_by_instr.get(sym)
        dates = sorted(hist_dates.get(sym, []))
        if not a or not dates:
            no_data.append(sym)
            continue

        lead, hold = parse_window(s.get("window", ""), s.get("lead", 2), s.get("hold", 4))

        rets = []
        for hd in dates:
            v = ret_for_window(a, hd, lead, hold)
            if v is not None:
                rets.append(v)

        if len(rets) < 3:
            insufficient.append((sym, len(rets)))
            continue

        avg_raw = sum(rets) / len(rets)
        direction = (s.get("direction") or "LONG").upper()
        signed_avg = avg_raw if direction == "LONG" else -avg_raw
        win_rate = (
            sum(1 for v in rets if v > 0) / len(rets) * 100
            if direction == "LONG"
            else sum(1 for v in rets if v < 0) / len(rets) * 100
        )

        old_ret = s.get("expected_ret")
        old_pnl = s.get("exp_pnl_1lot")
        new_ret = round(signed_avg, 2)
        new_pnl = round(signed_avg / 100 * s["entry_px"] * s["lot_size"], 2)

        s["expected_ret"] = new_ret
        s["expected_return"] = new_ret
        s["exp_pnl_1lot"] = new_pnl
        s["avg_17q_wr"] = round(win_rate, 2)

        updated.append((sym, direction, len(rets), old_ret, new_ret, old_pnl, new_pnl, round(win_rate, 1)))

    with open(PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
        f.write("\n")

    print(f"Updated {len(updated)} stocks (window/direction untouched):")
    for u in updated:
        print(f"  {u[0]:<12} dir={u[1]:<5} n={u[2]:<3} ret {u[3]:>7} -> {u[4]:>7}  pnl {u[5]:>10} -> {u[6]:>10}  winrate -> {u[7]}%")
    print(f"\nNo futures data or no prior results found: {no_data}")
    print(f"Insufficient history (<3 valid quarters): {insufficient}")


if __name__ == "__main__":
    main()
