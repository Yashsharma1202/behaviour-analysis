"""
APPLY the walk-forward 8x8 Dussehra results to the LIVE dashboard.
Replaces dussehra_2026's direction/window/win-rate/avg-return/entry-exit for
every stock with a genuine out-of-sample (OOS) walk-forward result. Stocks
with zero OOS trades (BSE, JIOFIN -- insufficient history even for the
3-prior-year minimum) are left on their existing fallback values, clearly
noted, since walk-forward literally cannot be computed for them.

ADDITIVE: updates dussehra_2026 stock entries in place (direction/window/
stats/dates change -- this IS the live position, not a new stock), nothing
is deleted or removed from the file.
"""
import json, pickle
from datetime import date, timedelta

BASE = r'D:\behaviour analysis'

results = pickle.load(open(BASE + r'\scraped_parquet\_tmp_dussehra_walkforward.pkl', 'rb'))

NSE_2026_HOLIDAYS = {
    '2026-01-26', '2026-03-03', '2026-03-04', '2026-03-26', '2026-03-31', '2026-04-03',
    '2026-04-14', '2026-05-01', '2026-05-28', '2026-06-26', '2026-09-14', '2026-10-02',
    '2026-10-20', '2026-11-24', '2026-12-25',
}
def is_trading_day(d):
    return d.weekday() < 5 and d.isoformat() not in NSE_2026_HOLIDAYS
def shift_trading_days(anchor, n):
    d = anchor
    step = 1 if n >= 0 else -1
    cnt = abs(n)
    while cnt > 0:
        d += timedelta(days=step)
        if is_trading_day(d): cnt -= 1
    return d

ANCHOR = date(2026, 10, 20)

path = BASE + r'\dashboard_data\event_dashboard_data.json'
d = json.load(open(path, encoding='utf-8'))
dus = next(h for h in d['holidays'] if h['id'] == 'dussehra_2026')

updated, skipped = 0, []
for s in dus['stocks']:
    sym = s['symbol']
    r = results.get(sym)
    if not r or r['oos_n'] < 1:
        skipped.append(sym)
        s['direction_basis'] = (s.get('direction_basis', '') + ' | Walk-forward 8x8 check: insufficient history for even 1 out-of-sample trade (needs >=4yr) -- kept on prior fallback value, NOT walk-forward validated.').strip(' |')
        continue

    bo, so = r['upcoming_bo'], r['upcoming_so']
    en = shift_trading_days(ANCHOR, -bo)
    ex = shift_trading_days(ANCHOR, so)

    old_dir, old_wr = s.get('direction'), s.get('full_19y_wr')
    s['direction'] = r['upcoming_side']
    s['window'] = f"T-{bo} to T+{so}"
    s['lead'] = bo
    s['hold'] = so
    s['entry_date'] = en.strftime('%d-%b-%Y (%a)')
    s['exit_date'] = ex.strftime('%d-%b-%Y (%a)')
    s['full_19y_wr'] = r['oos_wr']
    s['full_19y_avg_ret'] = r['oos_avg']
    s['last_4y_wr'] = r['oos_wr']
    s['last_4y_avg_ret'] = r['oos_avg']
    s['oos_trade_count'] = r['oos_n']
    s['low_confidence'] = r['low_confidence']
    s['methodology'] = (f"Walk-forward 8x8 grid search (point-in-time, no lookahead) -- "
                         f"same method as the Nifty index options model. {r['oos_n']} genuine "
                         f"out-of-sample trades. Window and direction were chosen using ONLY "
                         f"years strictly before each trade, never future data.")
    s.pop('data_pending_symbols', None)
    updated += 1
    changed = (old_dir != r['upcoming_side'])
    print(f"{sym:12} {old_dir or '?':5} (WR={old_wr}) -> {r['upcoming_side']:5} (WR={r['oos_wr']}, OOS avg={r['oos_avg']:+.2f}%, n={r['oos_n']}){'  [DIRECTION FLIPPED]' if changed else ''}")

dirs = {}
for s in dus['stocks']:
    dirs[s['direction']] = dirs.get(s['direction'], 0) + 1
dus['bias'] = f"Dussehra -- walk-forward 8x8 validated (point-in-time, no lookahead) -- {dirs}"

json.dump(d, open(path, 'w', encoding='utf-8'), indent=2)
print(f"\nUpdated {updated} stocks with walk-forward results. Skipped (insufficient history): {skipped}")
print("Direction split now:", dirs)
print("Saved:", path)
