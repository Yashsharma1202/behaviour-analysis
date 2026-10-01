"""
FIX THE DUSSEHRA POSITION-TAKING WINDOW \u2014 walk-forward 8x8 grid search.

Replaces the in-sample "best direction over all years" method used in today's
earlier F&O-211 Dussehra reports with the SAME rigorous, point-in-time method
already used for the Nifty index options model (optimize_8x8 in
backtest_nifty_index_options_1pct_itm_quarterly.py):

  For each Dussehra year, search entry offset bo in [1,8] trading days before
  the anchor x exit offset so in [1,8] trading days after, for LONG and SHORT
  separately (128 combos) -- using ONLY years strictly BEFORE the one being
  evaluated (no lookahead). The best (win_rate, avg_return) combo from that
  prior-years-only search is applied FORWARD to the year being evaluated,
  producing a genuine out-of-sample (OOS) trade return. Needs >=3 prior years
  before a combo is even considered (same confidence gate used everywhere
  else in this project).

Performance note: returns for every (date, bo, so) pair are precomputed once
per stock into a matrix, then every grid-search / walk-forward step is a
cheap lookup+average instead of re-scanning the price array -- otherwise this
is ~500M+ redundant operations across 211 stocks and does not finish.

ADDITIVE: reads existing price/date data; writes a new intermediate pickle +
summary; does not touch the live dashboard (a separate step applies the
result to event_dashboard_data.json).
"""
import json, pickle, time
from datetime import date
import pandas as pd

BASE = r'D:\behaviour analysis'
t0 = time.time()

# DATA QUALITY FIX: the cash-equity file (nifty50_all_stocks_daily_2000_2026.parquet)
# is missing 98.3% of all Fridays across its entire 26yr history (1,371 of
# 1,395 Fridays, 2000-2026) -- not a handful of gaps, the file is effectively
# a 4-day trading week for its whole life. Every "T-n trading days" countback
# and every return compounded on it has silently skipped real Friday price
# action project-wide. The futures file has no such defect (only ~6.5% of
# Fridays missing, consistent with genuine NSE holidays landing on a Friday).
# Switching ALL stocks to the futures source -- shorter history (2019-2026,
# ~7yr vs up to 26yr) but a complete, trustworthy trading calendar.
fut = pd.read_parquet(BASE + r'\scraped_parquet\fo_futures_near_month_continuous.parquet', columns=['Date', 'Instrument', 'Close'])
fut['Date'] = pd.to_datetime(fut['Date']).dt.date
fut_closes = {s: g.sort_values('Date')[['Date', 'Close']].values.tolist() for s, g in fut.groupby('Instrument')}

beh = json.load(open(BASE + r'\dashboard_data\nifty_futures_holiday_behaviour.json', encoding='utf-8'))
dh = [h for h in beh['holidays'] if 'Dussehra' in h['name']][0]
dussehra_dates = sorted(date.fromisoformat(t['holiday_date']) for t in dh['trades'])

fo = json.load(open(BASE + r'\dashboard_data\fo_stocks_211.json', encoding='utf-8'))
meta = {f['symbol']: f for f in fo}

BO_RANGE = range(1, 9)
SO_RANGE = range(1, 9)


def closes_for(sym):
    if sym in fut_closes:
        return fut_closes[sym], 'futures-7yr', [d for d in dussehra_dates if d.year >= 2019]
    return None, None, []


def build_matrix(a, dates_to_use):
    """ret_mat[date_idx][(bo,so)] = split-robust signed %% raw (LONG-convention) return."""
    # index of nearest trading day <= hd, found via a single forward pointer sweep (dates sorted)
    ib_list = []
    ptr = -1
    for hd in dates_to_use:
        while ptr + 1 < len(a) and a[ptr + 1][0] <= hd:
            ptr += 1
        ib_list.append(ptr)

    mat = []
    for ib in ib_list:
        row = {}
        if ib < 0:
            mat.append(row)
            continue
        for bo in BO_RANGE:
            en = ib - bo
            if en < 1:
                continue
            # incremental compounding reused across so values for this bo
            g = 1.0
            valid = True
            k = en
            # compound from en..ib first (shared prefix for all so)
            while k <= ib:
                p0, p1 = a[k - 1][1], a[k][1]
                if not p0:
                    valid = False
                    break
                dr = (p1 - p0) / p0
                if abs(dr) <= 0.20:
                    g *= (1 + dr)
                k += 1
            if not valid:
                continue
            g_at_anchor = g
            for so in SO_RANGE:
                ex = ib + so
                if ex >= len(a):
                    continue
                g2 = g_at_anchor
                ok = True
                for k2 in range(ib + 1, ex + 1):
                    p0, p1 = a[k2 - 1][1], a[k2][1]
                    if not p0:
                        ok = False
                        break
                    dr = (p1 - p0) / p0
                    if abs(dr) <= 0.20:
                        g2 *= (1 + dr)
                if ok:
                    row[(bo, so)] = (g2 - 1) * 100
        mat.append(row)
    return mat


def best_from_prior(mat, prior_idx, side):
    best, best_score = (2, 5), (-9e9, -9e9)
    for bo in BO_RANGE:
        for so in SO_RANGE:
            rets = []
            for i in prior_idx:
                v = mat[i].get((bo, so))
                if v is not None:
                    rets.append(v if side == 'LONG' else -v)
            if len(rets) >= 3:
                win = sum(1 for r in rets if r > 0) / len(rets) * 100
                avg = sum(rets) / len(rets)
                if (win, avg) > best_score:
                    best_score = (win, avg)
                    best = (bo, so)
    return best, best_score


results = {}
total = len(meta)
for i, sym in enumerate(meta, 1):
    a, source, dates_to_use = closes_for(sym)
    if a is None:
        results[sym] = dict(source=None, oos_n=0, low_confidence=True)
        continue

    mat = build_matrix(a, dates_to_use)

    oos_trades = []
    for yi in range(len(dates_to_use)):
        prior_idx = list(range(yi))
        if len(prior_idx) < 3:
            continue
        (lbo, lso), lsc = best_from_prior(mat, prior_idx, 'LONG')
        (sbo, sso), ssc = best_from_prior(mat, prior_idx, 'SHORT')
        if lsc == (-9e9, -9e9) and ssc == (-9e9, -9e9):
            continue
        if ssc > lsc:
            side, bo, so = 'SHORT', sbo, sso
        else:
            side, bo, so = 'LONG', lbo, lso
        raw = mat[yi].get((bo, so))
        if raw is None:
            continue
        signed = raw if side == 'LONG' else -raw
        oos_trades.append(dict(year=dates_to_use[yi].year, side=side, bo=bo, so=so, ret=round(signed, 2)))

    n = len(oos_trades)

    all_idx = list(range(len(dates_to_use)))
    (lbo, lso), lsc = best_from_prior(mat, all_idx, 'LONG')
    (sbo, sso), ssc = best_from_prior(mat, all_idx, 'SHORT')
    if lsc == (-9e9, -9e9) and ssc == (-9e9, -9e9):
        up_side, up_bo, up_so = 'LONG', 2, 5
    elif ssc > lsc:
        up_side, up_bo, up_so = 'SHORT', sbo, sso
    else:
        up_side, up_bo, up_so = 'LONG', lbo, lso

    if n >= 1:
        wins = sum(1 for t in oos_trades if t['ret'] > 0)
        wr = round(100 * wins / n, 1)
        avg = round(sum(t['ret'] for t in oos_trades) / n, 2)
    else:
        wr = avg = None

    results[sym] = dict(source=source, oos_n=n, oos_wr=wr, oos_avg=avg, oos_trades=oos_trades,
                         upcoming_side=up_side, upcoming_bo=up_bo, upcoming_so=up_so,
                         low_confidence=(n < 3))
    if i % 25 == 0 or i == total:
        print(f'  {i}/{total} stocks processed... ({time.time()-t0:.1f}s elapsed)')

pickle.dump(results, open(BASE + r'\scraped_parquet\_tmp_dussehra_walkforward.pkl', 'wb'))

qualified = {s: r for s, r in results.items() if r['oos_n'] >= 1}
zero = {s: r for s, r in results.items() if r['oos_n'] == 0}
print(f"\nTotal: {len(results)} | OOS-qualified (>=1 walk-forward trade): {len(qualified)} | zero: {len(zero)}")
low_conf = sum(1 for r in qualified.values() if r['low_confidence'])
print(f"Low-confidence (n<3 OOS trades): {low_conf}")
print(f"Total time: {time.time()-t0:.1f}s")

ranked = sorted(qualified.items(), key=lambda kv: -kv[1]['oos_avg'])
print("\nTop 15 by OOS avg return (genuinely walk-forward, not in-sample):")
for sym, r in ranked[:15]:
    print(f"  {sym:12} OOS_WR={r['oos_wr']:5.1f}% OOS_avg={r['oos_avg']:+6.2f}% n={r['oos_n']:2} "
          f"upcoming={r['upcoming_side']:5} T-{r['upcoming_bo']}/T+{r['upcoming_so']} src={r['source']}")
