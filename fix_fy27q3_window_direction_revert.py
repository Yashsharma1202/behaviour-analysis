"""
fix_fy27q3_window_direction_revert.py
===============================================================================
During this session's big f821bdf8 commit, a per-stock walk-forward window
optimizer silently ran against FY27_Q3 (same bug class already caught and
excluded for TCS: see daily_refresh_and_publish.py's comment on
fix_q3_position_window.py). It changed 21 stocks away from the project's
uniform T-2/T+4 window, and for 10 of those it also flipped `direction` to
disagree with `best_strategy` (the real backtested signal, which was left
untouched). TCS itself was caught and manually reverted (commit 26a08bf4);
these other 20 were not (TRENT is excluded here -- its non-uniform window
pre-dates this bug and is a separate, legitimate exception).

Fix: revert window to T-2/T+4 (recomputed via the live NSE trading calendar,
not copied from any stale snapshot), restore `direction` to match
`best_strategy` where they disagree, and correct the stale `window_fix_note`.
All 20 stocks are still outcome=PENDING (no actual_pnl), so nothing already
realized/reported is touched.
"""
import json
from datetime import datetime
import nse_trading_calendar as ntc

PATH = "dashboard_data/event_dashboard_data.json"

AFFECTED = [
    "WIPRO", "HINDALCO", "BAJAJFINSV", "CIPLA", "HDFCLIFE", "ICICIBANK",
    "HDFCBANK", "ADANIPORTS", "NESTLEIND", "INFY", "JSWSTEEL", "TECHM",
    "M&M", "DRREDDY", "ASIANPAINT", "ULTRACEMCO", "HINDUNILVR", "MARUTI",
    "BAJAJ-AUTO", "BAJFINANCE",
]


def parse_result_date(s):
    return datetime.strptime(s.strip(), "%d-%b-%Y").date()


def fmt(d):
    return d.strftime("%d-%b-%Y (%a)")


def best_strategy_direction(best_strategy):
    bs = (best_strategy or "").upper()
    if "SHORT" in bs:
        return "SHORT"
    if "LONG" in bs:
        return "LONG"
    return None


def main():
    with open(PATH, encoding="utf-8") as f:
        data = json.load(f)

    holidays = ntc.load_trading_holidays()

    q3 = None
    for q in data["quarters"]:
        if q["q_code"] == "FY27_Q3":
            q3 = q
            break
    assert q3 is not None

    fixed_window = []
    fixed_direction = []

    for s in q3["stocks"]:
        if s["symbol"] not in AFFECTED:
            continue

        result_date = parse_result_date(s["result_date"])
        entry_date = ntc.add_trading_days(result_date, -2, holidays)
        exit_date = ntc.add_trading_days(result_date, 4, holidays)

        old_window = s.get("window")
        s["window"] = "T-2 to T+4"
        s["lead"] = 2
        s["hold"] = 4
        s["entry_lead_days"] = 2
        s["exit_hold_days"] = 4
        s["taking_window_raw"] = "T-2 to T+4"
        s["pos_window"] = "T-2 to T+4"
        s["entry_date"] = fmt(entry_date)
        s["exit_date"] = fmt(exit_date)
        if "entry_date_sample" in s:
            s["entry_date_sample"] = fmt(entry_date)
        if "q_entry_date" in s:
            s["q_entry_date"] = entry_date.strftime("%Y-%m-%d (%a)")
        if "exit_date_sample" in s:
            s["exit_date_sample"] = fmt(exit_date)
        if "q_exit_date" in s:
            s["q_exit_date"] = exit_date.strftime("%Y-%m-%d (%a)")
        s["taking_window_full"] = (
            f"Entry: {fmt(entry_date)} (T-2 to T+4) ➔ Exit: {fmt(exit_date)}"
        )
        s["window_fix_note"] = (
            "Reverted to uniform T-2/T+4 window on 08-Oct-2026: a per-stock "
            "walk-forward optimizer had silently overridden this quarter's "
            "window (same bug class caught and excluded for TCS)."
        )
        fixed_window.append((s["symbol"], old_window))

        bs_dir = best_strategy_direction(s.get("best_strategy"))
        if bs_dir and s.get("direction") != bs_dir:
            fixed_direction.append((s["symbol"], s.get("direction"), bs_dir))
            s["direction"] = bs_dir

    with open(PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
        f.write("\n")

    print(f"Window reverted to T-2/T+4 for {len(fixed_window)} stocks:")
    for sym, old in fixed_window:
        print(f"  {sym}: {old} -> T-2 to T+4")
    print(f"\nDirection restored to match best_strategy for {len(fixed_direction)} stocks:")
    for sym, old, new in fixed_direction:
        print(f"  {sym}: {old} -> {new}")


if __name__ == "__main__":
    main()
