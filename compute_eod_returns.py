"""
END-OF-DAY "RETURN WE GET TILL NOW" for active event positions.
================================================================
NOT live/intraday. Reads the daily EOD closes from
  scraped_parquet/nifty50_all_stocks_daily_2000_2026.parquet
and, for every position whose entry has passed, computes:

  return_till_now = (latest_close - entry_close) / entry_close * (+1 LONG / -1 SHORT)

  entry_close  = the stock's CLOSE on its entry date (nearest trading day <= entry)
  latest_close = the stock's most recent CLOSE available, but not past the exit date

Writes dashboard_data/eod_returns.json (keyed by holiday/event -> symbol) so the
dashboard can show a "Return Till Now" column that updates whenever you refresh
AFTER the daily EOD parquet has been updated. If the EOD file has not yet been
updated past a position's entry date, that position is reported as AWAITING_DATA.

ADDITIVE: reads existing data, writes one new JSON. Nothing else touched.
"""
import json, os, re
from datetime import date, timedelta
import pandas as pd

BASE = os.path.dirname(os.path.abspath(__file__))
PARQUET = os.path.join(BASE, 'scraped_parquet', 'nifty50_all_stocks_daily_2000_2026.parquet')
EV = os.path.join(BASE, 'dashboard_data', 'event_dashboard_data.json')
OUT = os.path.join(BASE, 'dashboard_data', 'eod_returns.json')

_MONS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
def pdmy(s):
    m = re.search(r'(\d{1,2})-([A-Za-z]{3})-(\d{4})', str(s))
    return date(int(m.group(3)), _MONS.index(m.group(2)) + 1, int(m.group(1))) if m else None

TODAY = date.today()

# ---- load EOD closes into {symbol: sorted [(date, close)]} ----
df = pd.read_parquet(PARQUET, columns=['DATE', 'SYMBOL', 'CLOSE'])
df['DATE'] = pd.to_datetime(df['DATE']).dt.date
closes = {}
for sym, g in df.groupby('SYMBOL'):
    closes[sym] = g.sort_values('DATE')[['DATE', 'CLOSE']].values.tolist()
DATA_MAX = max(df['DATE'])
DATA_SYMS = set(closes)

def close_on_or_before(sym, d):
    arr = closes.get(sym)
    if not arr: return None, None
    best = None
    for dt, c in arr:
        if dt <= d: best = (dt, float(c))
        else: break
    return best if best else (None, None)

def latest_close_upto(sym, d):
    arr = closes.get(sym)
    if not arr: return None, None
    best = None
    for dt, c in arr:
        if dt <= d: best = (dt, float(c))
        else: break
    return best if best else (None, None)

def compute_position(sym, direction, entry_date, exit_date):
    dirsign = -1 if str(direction).upper() == 'SHORT' else 1
    if sym not in DATA_SYMS:
        return {'status': 'NO_SYMBOL_DATA', 'symbol': sym}
    if entry_date > DATA_MAX:
        return {'status': 'AWAITING_DATA', 'symbol': sym,
                'note': 'entry %s is after last EOD data %s' % (entry_date, DATA_MAX)}
    en_dt, en_px = close_on_or_before(sym, entry_date)
    asof_cap = min(TODAY, exit_date, DATA_MAX)
    ax_dt, ax_px = latest_close_upto(sym, asof_cap)
    if en_px is None or ax_px is None or en_px == 0:
        return {'status': 'NO_PRICE', 'symbol': sym}
    ret = (ax_px - en_px) / en_px * dirsign * 100
    closed = exit_date <= DATA_MAX and exit_date <= TODAY
    return {
        'status': 'CLOSED' if closed else 'OPEN',
        'symbol': sym, 'direction': str(direction).upper(),
        'entry_date_used': str(en_dt), 'entry_close': round(en_px, 2),
        'asof_date': str(ax_dt), 'asof_close': round(ax_px, 2),
        'return_till_now_pct': round(ret, 2),
    }

def main():
    ev = json.load(open(EV, encoding='utf-8'))
    out = {'generated_at': TODAY.isoformat(), 'eod_data_through': str(DATA_MAX),
           'source': 'nifty50_all_stocks_daily_2000_2026.parquet', 'events': {}}

    # Holidays
    for h in ev.get('holidays', []):
        block = {}
        for s in h.get('stocks', []):
            en, xt = pdmy(s.get('entry_date')), pdmy(s.get('exit_date'))
            if not en or not xt: continue
            if en > TODAY:  # not entered yet -> skip (upcoming)
                continue
            block[s['symbol']] = compute_position(s['symbol'], s.get('direction'), en, xt)
        if block:
            out['events'][h['id']] = block

    json.dump(out, open(OUT, 'w', encoding='utf-8'), indent=2, default=str)

    # ---- console report ----
    print('EOD RETURN ENGINE  |  today=%s  |  EOD data through=%s' % (TODAY, DATA_MAX))
    for eid, block in out['events'].items():
        computed = [v for v in block.values() if v.get('status') in ('OPEN', 'CLOSED')]
        awaiting = [v for v in block.values() if v.get('status') == 'AWAITING_DATA']
        nosym = [v for v in block.values() if v.get('status') in ('NO_SYMBOL_DATA', 'NO_PRICE')]
        print('\n== %s ==  computed:%d  awaiting-data:%d  no-symbol:%d'
              % (eid, len(computed), len(awaiting), len(nosym)))
        for v in sorted(computed, key=lambda x: -x['return_till_now_pct'])[:60]:
            print('   %-11s %-5s entry %s @%.2f -> %s @%.2f  ret_till_now %+6.2f%%  [%s]' % (
                v['symbol'], v['direction'], v['entry_date_used'], v['entry_close'],
                v['asof_date'], v['asof_close'], v['return_till_now_pct'], v['status']))
        if awaiting:
            print('   AWAITING EOD update for: ' + ', '.join(v['symbol'] for v in awaiting))
    print('\nWrote', OUT)


if __name__ == '__main__':
    main()
