"""
apply_backtest_windows_unannounced.py
===============================================================================
Applies the backtest-optimal window (found via per-stock T-n/T+m grid search,
direction held fixed -- not re-selected) to the 24 still-unannounced FY27_Q3
stocks that were sitting on the generic T-2/T+4 placeholder (or an earlier
unverified non-uniform value). BSE excluded (no futures history).

Updates: window, lead, hold, expected_ret, expected_return, exp_pnl_1lot,
avg_17q_wr, plus the mirrored fields (entry_lead_days, exit_hold_days,
taking_window_raw, pos_window) and window_fix_note.
Does NOT touch: direction, result_date, entry_date, exit_date (still "Yet to
come" for all of these -- no real date to compute from yet), entry_px, outcome.
"""
import json
from datetime import datetime
import pandas as pd

PATH = "dashboard_data/event_dashboard_data.json"

SYMBOLS = ['ETERNAL','BEL','COALINDIA','ITC','POWERGRID','ONGC','TATASTEEL','SBIN','NTPC',
           'TATACONSUM','SHRIRAMFIN','BHARTIARTL','INDIGO','MAXHEALTH','ADANIENT','GRASIM',
           'SBILIFE','TMPV','SUNPHARMA','TITAN','LT','RELIANCE','APOLLOHOSP','EICHERMOT']


def parse_date(s):
    if not s or s in ("Yet to come", "Awaiting NSE Announcement"):
        return None
    core = s.split(" (")[0].strip()
    if len(core) == 10 and core[4] == "-" and core[7] == "-":
        return datetime.strptime(core, "%Y-%m-%d").date()
    return datetime.strptime(core, "%d-%b-%Y").date()


def ret_for_window(a, hd, bo, so, MAX_GAP_DAYS=15):
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


def best_window_fixed_side(a, dates, side, min_n=3):
    best, best_score = None, (-9e9, -9e9)
    for bo in range(1, 9):
        for so in range(1, 9):
            rets = []
            for hd in dates:
                v = ret_for_window(a, hd, bo, so)
                if v is not None:
                    rets.append(v if side == "LONG" else -v)
            if len(rets) >= min_n:
                win = sum(1 for v in rets if v > 0) / len(rets) * 100
                avg = sum(rets) / len(rets)
                if (win, avg) > best_score:
                    best_score = (win, avg)
                    best = (bo, so)
    return best, best_score


def main():
    with open(PATH, encoding="utf-8") as f:
        data = json.load(f)
    q3 = next(q for q in data["quarters"] if q["q_code"] == "FY27_Q3")

    hist = {}
    for q in data["quarters"]:
        if q["q_code"] == "FY27_Q3":
            continue
        for s in q["stocks"]:
            rd = parse_date(s.get("result_date"))
            if rd:
                hist.setdefault(s["symbol"], set()).add(rd)

    fut = pd.read_parquet(
        "scraped_parquet/fo_futures_near_month_continuous.parquet",
        columns=["Date", "Instrument", "Close"],
    )
    fut["Date"] = pd.to_datetime(fut["Date"]).dt.date
    fut_by = {
        sym: g.sort_values("Date")[["Date", "Close"]].values.tolist()
        for sym, g in fut.groupby("Instrument")
    }

    by_sym = {s["symbol"]: s for s in q3["stocks"]}
    updated = []

    for sym in SYMBOLS:
        s = by_sym[sym]
        direction = s["direction"]
        a = fut_by.get(sym)
        dates = sorted(hist.get(sym, []))
        if not a or len(dates) < 3:
            continue
        best, score = best_window_fixed_side(a, dates, direction)
        if best is None:
            continue
        bo, so = best
        win_rate, avg_ret = score
        new_window = f"T-{bo} to T+{so}"
        old_window = s.get("window")

        s["window"] = new_window
        s["lead"] = bo
        s["hold"] = so
        s["entry_lead_days"] = bo
        s["exit_hold_days"] = so
        s["taking_window_raw"] = new_window
        s["pos_window"] = new_window
        s["taking_window_full"] = f"Entry: T-{bo} trading days ({new_window}) ➔ Exit: T+{so} trading days (result not yet announced)"
        s["window_fix_note"] = (
            "Window computed via per-stock T-n/T+m backtest grid search on 09-Oct-2026 "
            "(direction held fixed, not re-selected) -- replaces the generic T-2/T+4 "
            "placeholder that was never individually verified for this still-unannounced stock."
        )
        s["expected_ret"] = round(avg_ret, 2)
        s["expected_return"] = round(avg_ret, 2)
        s["avg_17q_wr"] = round(win_rate, 2)
        s["exp_pnl_1lot"] = round(avg_ret / 100 * s["entry_px"] * s["lot_size"], 2)

        updated.append((sym, old_window, new_window, round(win_rate, 1), round(avg_ret, 2)))

    with open(PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
        f.write("\n")

    print(f"Updated {len(updated)} stocks:")
    for u in updated:
        print(" ", u)


if __name__ == "__main__":
    main()
