"""
Compute the holiday DIRECTION backtest on the full Nifty futures index and write
dashboard_data/holiday_direction_backtest.json for the dashboard to display.

Source of truth = dashboard_data/nifty_futures_holiday_behaviour.json (26 index
trades per holiday, each holiday's optimal window). LONG return = index move over
the window; SHORT = mirror. Reported for 26yr / last-15yr / out-of-sample(2013-25).
ADDITIVE: reads existing data, writes one new JSON.
"""
import json, os, statistics as st
from datetime import date

BASE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(BASE, 'dashboard_data', 'nifty_futures_holiday_behaviour.json')
OUT = os.path.join(BASE, 'dashboard_data', 'holiday_direction_backtest.json')

beh = json.load(open(SRC, encoding='utf-8'))

def rows(name):
    h = [x for x in beh['holidays'] if name.lower() in x['name'].lower()][0]
    data = []
    for t in h['trades']:
        y = date.fromisoformat(t['holiday_date']).year
        raw = (t['exit_price'] - t['entry_price']) / t['entry_price'] * 100
        data.append((y, round(raw, 2)))
    return h['optimal_window'], h['direction'], sorted(data)

def stats(data, side, years):
    r = [(x if side == 'LONG' else -x) for y, x in data if y in years]
    if not r:
        return None
    return {'n': len(r), 'win_rate': round(100 * sum(1 for v in r if v > 0) / len(r), 1),
            'avg': round(st.mean(r), 2), 'total': round(sum(r), 1)}

HOL = [('Dussehra / Vijayadashami', 'dussehra_2026'), ('Diwali', 'diwali_2026'), ('Christmas', 'christmas_2026')]
out = {'generated_at': date.today().isoformat(),
       'basis': 'Full Nifty futures index; each holiday’s empirically-optimal window; split-free (index level).',
       'holidays': {}}
for name, hid in HOL:
    win, idxdir, data = rows(name)
    yrs = [y for y, _ in data]
    ally = set(yrs); l15 = set(y for y in yrs if y >= max(yrs) - 14); oos = set(y for y in yrs if y >= 2013)
    verdict = 'LONG' if idxdir.upper() == 'LONG' else 'SHORT'
    out['holidays'][hid] = {
        'name': name, 'window': win, 'verdict': verdict, 'span': '%d-%d' % (min(yrs), max(yrs)),
        'LONG': {'all': stats(data, 'LONG', ally), 'last15': stats(data, 'LONG', l15), 'oos': stats(data, 'LONG', oos)},
        'SHORT': {'all': stats(data, 'SHORT', ally), 'last15': stats(data, 'SHORT', l15), 'oos': stats(data, 'SHORT', oos)},
        'year_returns_long': [{'year': y, 'ret': r} for y, r in data],
    }

json.dump(out, open(OUT, 'w', encoding='utf-8'), indent=2)
print('Saved', OUT)
for hid, h in out['holidays'].items():
    v = h['verdict']; s = h[v]
    print('  %-14s verdict %-5s  win %s (26y) %s (15y) %s (oos)  | total %s%% (26y)'
          % (hid, v, s['all']['win_rate'], s['last15']['win_rate'], s['oos']['win_rate'], s['all']['total']))
