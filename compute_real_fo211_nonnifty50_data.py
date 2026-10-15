"""
compute_real_fo211_nonnifty50_data.py
===============================================================================
fo_stocks_211.json's 161 non-Nifty-50 entries were entirely placeholder data:
spot_ltp hardcoded to 1000.0 for every stock, best_strategy identical generic
text for every stock, win rates suspiciously sequential (71, 72, 73, 74...).

Replaces all of that with real computed values per stock:
  - spot_ltp: latest real close from processed/price_cache/<SYM>.csv
  - direction + window: joint grid search (LONG/SHORT x T-1..T8/T+1..T+8)
    against each stock's own historical result dates (from
    <SYM>/financial_results.csv) and real futures price data, scored by
    win-rate then avg return -- same methodology used for the Nifty-50
    quarterly backtest today, just picking direction freshly since these
    161 stocks never had one.
  - q_win_rate / h_win_rate: the real win rate from that backtest
  - h_est_pnl: recomputed from the real expected return, price, and lot size
  - best_strategy: accurate per-stock text, not generic placeholder

Does NOT touch any Nifty-50 stock (those stay exactly as fixed earlier
today) and does NOT fabricate a live current-quarter result date -- entry/
exit "sample" dates are left as a generic marker so the dashboard's existing
"Yet to come" fallback renders instead of a fake specific date.
"""
import json
import os
from datetime import datetime
import pandas as pd

MAX_GAP_DAYS = 15


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


def best_direction_and_window(a, dates, min_n=3):
    best, best_score, best_side = None, (-9e9, -9e9), None
    for side in ("LONG", "SHORT"):
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
                        best_side = side
    return best_side, best, best_score


def main():
    with open("dashboard_data/fo_stocks_211.json", encoding="utf-8") as f:
        data = json.load(f)

    fut = pd.read_parquet(
        "scraped_parquet/fo_futures_near_month_continuous.parquet",
        columns=["Date", "Instrument", "Close"],
    )
    fut["Date"] = pd.to_datetime(fut["Date"]).dt.date
    fut_by = {
        sym: g.sort_values("Date")[["Date", "Close"]].values.tolist()
        for sym, g in fut.groupby("Instrument")
    }

    updated, skipped = [], []

    for s in data:
        if s.get("is_nifty50") == "YES":
            continue  # never touch Nifty-50 entries

        sym = s["symbol"]

        price_path = f"processed/price_cache/{sym}.csv"
        if not os.path.exists(price_path):
            skipped.append((sym, "no price_cache"))
            continue
        pdf = pd.read_csv(price_path)
        if pdf.empty:
            skipped.append((sym, "empty price_cache"))
            continue
        real_spot = float(pdf["adj"].iloc[-1])

        fr_path = f"{sym}/financial_results.csv"
        if not os.path.exists(fr_path):
            skipped.append((sym, "no financial_results.csv"))
            continue
        try:
            fr = pd.read_csv(fr_path, dtype=str).fillna("")
        except Exception:
            skipped.append((sym, "unreadable financial_results.csv"))
            continue
        if "broadCastDate" not in fr.columns:
            skipped.append((sym, "no broadCastDate column"))
            continue
        dates = []
        for d in fr["broadCastDate"]:
            try:
                dates.append(datetime.strptime(d, "%d-%b-%Y").date())
            except ValueError:
                continue
        dates = sorted(set(dates))

        a = fut_by.get(sym)
        if not a:
            skipped.append((sym, "no futures price data"))
            continue

        side, window, score = best_direction_and_window(a, dates, min_n=3)
        if side is None:
            skipped.append((sym, f"insufficient history (n={len(dates)})"))
            continue

        bo, so = window
        win_rate, avg_ret = score
        lot_size = s.get("lot_size", 1)
        margin = round(real_spot * lot_size * 0.20, 2)
        exp_pnl = round(avg_ret / 100 * real_spot * lot_size, 2)

        s["spot_ltp"] = round(real_spot, 2)
        s["margin_20pct"] = margin
        s["taking_window_raw"] = f"T-{bo} to T+{so}"
        s["entry_lead_days"] = bo
        s["exit_hold_days"] = so
        s["direction"] = side
        s["q_win_rate"] = round(win_rate, 1)
        s["h_win_rate"] = round(win_rate, 1)
        s["h_est_pnl"] = exp_pnl
        s["expected_ret"] = round(avg_ret, 2)
        s["best_strategy"] = f"Pre-Quarterly Run-up ({side})"
        s["n_hist_quarters"] = len(dates)
        # No live current-quarter date available for these 161 -- leave as
        # a generic marker so the dashboard's "Yet to come" fallback renders
        # instead of a fabricated specific date.
        s["entry_date_sample"] = "Yet to come"
        s["exit_date_sample"] = "Yet to come"
        s["announced"] = False
        s["data_source"] = (
            f"Real backtest: {len(dates)} historical quarters, "
            f"{bo}/{so} T-n/T+m grid search, direction chosen by win-rate "
            f"then avg return (same methodology as the Nifty-50 quarterly model)"
        )

        updated.append((sym, side, f"T-{bo}/T+{so}", round(win_rate, 1), round(avg_ret, 2), round(real_spot, 2)))

    with open("dashboard_data/fo_stocks_211.json", "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
        f.write("\n")

    print(f"Updated {len(updated)} non-Nifty-50 stocks with real computed data")
    print(f"Skipped {len(skipped)}: {skipped}")
    print()
    for u in updated[:20]:
        print(" ", u)
    print(f"  ... and {max(0, len(updated)-20)} more")


if __name__ == "__main__":
    main()
