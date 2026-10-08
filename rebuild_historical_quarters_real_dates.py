"""
REBUILD THE 17 HISTORICAL QUARTERS (FY23_Q2 .. FY27_Q2) ON REAL NSE RESULT
DATES instead of the synthetic placeholder dates discovered this session
(every stock "reporting" on a fixed day-of-month, exactly 3 months apart,
for 17 straight quarters -- not how real companies report).

Real dates come from data_center/nse_real_quarterly_result_dates.json,
itself extracted from two authentic local NSE archives (see
extract_real_result_dates.py for full provenance). For the 717/850
(symbol, quarter) slots with a real date:

  1. result_date is replaced with the real NSE date.
  2. entry_date / exit_date are recomputed by shifting that real date by the
     stock's EXISTING window (lead/hold -- not changed, e.g. T-2 to T+4)
     along the real NSE trading-day calendar.
  3. entry_px / exit_px are refetched from real Yahoo Finance adjusted OHLC
     (Open=entry, Close=exit) at the NEW dates.
  4. actual_ret / actual_pnl / outcome are recomputed from the new prices.

For the 133 slots with no real date yet, the stock is left untouched and
explicitly tagged date_source='synthetic_placeholder_unverified' so this
is visible in the data rather than silently mixed in with verified rows.

Quarter-level aggregates (win_rate_avg, total_pnl, bias, ...) are
recomputed from the rebuilt per-stock data afterward.

ADDITIVE: reads event_dashboard_data.json + the data_center extraction +
fresh Yahoo Finance prices; updates price/date-dependent fields in place.
"""
import json
import datetime
import pandas as pd
import yfinance as yf

BASE = r'D:\behaviour analysis'
d = json.load(open(BASE + r'\dashboard_data\event_dashboard_data.json', encoding='utf-8'))
real = json.load(open(BASE + r'\data_center\nse_real_quarterly_result_dates.json', encoding='utf-8'))

TARGET_QS = {q['q_code'] for q in real['quarters']}

# ---- NSE trading-day calendar (2022-2027), reused from scratch/fix_all_dates_across_all_quarters.py ----
nse_holidays = set([
    '2022-01-26', '2022-03-01', '2022-03-18', '2022-04-14', '2022-04-15', '2022-05-03', '2022-08-09', '2022-08-15', '2022-08-31', '2022-10-05', '2022-10-24', '2022-10-26', '2022-11-08',
    '2023-01-26', '2023-03-07', '2023-03-30', '2023-04-04', '2023-04-07', '2023-04-14', '2023-05-01', '2023-06-29', '2023-08-15', '2023-09-19', '2023-10-02', '2023-10-24', '2023-11-14', '2023-11-27', '2023-12-25',
    '2024-01-22', '2024-01-26', '2024-03-08', '2024-03-25', '2024-03-29', '2024-04-11', '2024-04-17', '2024-05-01', '2024-05-20', '2024-06-17', '2024-07-17', '2024-08-15', '2024-10-02', '2024-11-01', '2024-11-15', '2024-12-25',
    '2025-01-26', '2025-02-26', '2025-03-14', '2025-03-31', '2025-04-10', '2025-04-14', '2025-04-18', '2025-05-01', '2025-08-15', '2025-08-27', '2025-10-02', '2025-10-21', '2025-10-22', '2025-11-05', '2025-12-25',
    '2026-01-26', '2026-03-03', '2026-03-20', '2026-04-03', '2026-04-14', '2026-05-01', '2026-05-28', '2026-08-15', '2026-10-02', '2026-10-20', '2026-11-09', '2026-12-25'
])
all_dates = pd.date_range(start='2022-01-01', end='2027-12-31', freq='B')
td_list = sorted(d.strftime('%Y-%m-%d') for d in all_dates if d.strftime('%Y-%m-%d') not in nse_holidays)
td_set = set(td_list)
_DOW = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']


def shift_trading_days(base_date, offset):
    ds = base_date.strftime('%Y-%m-%d')
    dt = base_date
    while ds not in td_set:
        dt -= datetime.timedelta(days=1)
        ds = dt.strftime('%Y-%m-%d')
    idx = td_list.index(ds)
    target_idx = max(0, min(len(td_list) - 1, idx + offset))
    return datetime.datetime.strptime(td_list[target_idx], '%Y-%m-%d').date()


def fmt_dmy(dt):
    return f"{dt.day:02d}-{dt.strftime('%b')}-{dt.year} ({_DOW[dt.weekday()]})"


def parse_window(win):
    lead, hold = 2, 4
    if win and 'T-' in win and 'to T+' in win:
        try:
            lead = int(win.split('T-')[1].split(' ')[0])
            hold = int(win.split('T+')[1].split(' ')[0])
        except Exception:
            pass
    return lead, hold


# ---- 1. Apply real dates + recompute entry/exit across all touched stocks ----
touched = []  # (q_code, stock_dict)
untouched_synthetic = 0
for q in d['quarters']:
    if q['q_code'] not in TARGET_QS:
        continue
    for s in q['stocks']:
        if s.get('is_nifty50') != 'YES':
            continue
        rd = real['result_dates'].get(s['symbol'], {}).get(q['q_code'])
        if not rd:
            s['date_source'] = 'synthetic_placeholder_unverified'
            untouched_synthetic += 1
            continue
        result_dt = datetime.datetime.strptime(rd, '%d-%b-%Y').date()
        lead, hold = parse_window(s.get('window'))
        en_dt = shift_trading_days(result_dt, -lead)
        ex_dt = shift_trading_days(result_dt, hold)

        s['result_date'] = rd
        s['entry_date'] = fmt_dmy(en_dt)
        s['exit_date'] = fmt_dmy(ex_dt)
        s['date_source'] = 'NSE_verified'
        touched.append((q['q_code'], s))

print(f"Real NSE dates applied to {len(touched)} stock-quarter records; "
      f"{untouched_synthetic} left on the old synthetic placeholder (flagged).")

# ---- 2. Refetch real prices for every touched stock's full date range ----
syms = sorted({s['symbol'] for _, s in touched})
print(f"\nFetching real price history for {len(syms)} symbols...")
price_data = {}
failed = []
TICKER_ALIASES = {'TMPV': 'TMPV.NS'}
for sym in syms:
    ticker = TICKER_ALIASES.get(sym, sym + '.NS')
    try:
        h = yf.Ticker(ticker).history(start='2022-06-01', end='2026-10-05', auto_adjust=True)
        if h.empty:
            failed.append(sym)
            continue
        h.index = h.index.date
        price_data[sym] = h
    except Exception as e:
        failed.append(sym)
        print(f"  FAILED {sym}: {e}")
print(f"Fetched {len(price_data)} of {len(syms)} symbols. Failed: {failed}")


def nearest_trading_price(hist, target_date, field, direction='forward'):
    dates = hist.index
    if target_date in dates:
        return hist.loc[target_date, field]
    candidates = sorted(dates)
    if direction == 'forward':
        for dt in candidates:
            if dt >= target_date:
                return hist.loc[dt, field]
    else:
        for dt in reversed(candidates):
            if dt <= target_date:
                return hist.loc[dt, field]
    return None


# ---- 3. Recompute price-dependent fields for touched stocks ----
updated = 0
skipped = 0
flips = []
for q_code, s in touched:
    hist = price_data.get(s['symbol'])
    if hist is None:
        skipped += 1
        continue
    en_date = datetime.datetime.strptime(s['entry_date'].split(' (')[0], '%d-%b-%Y').date()
    ex_date = datetime.datetime.strptime(s['exit_date'].split(' (')[0], '%d-%b-%Y').date()
    en_px = nearest_trading_price(hist, en_date, 'Open', 'forward')
    ex_px = nearest_trading_price(hist, ex_date, 'Close', 'backward')
    if en_px is None or ex_px is None or en_px == 0:
        skipped += 1
        continue

    direction = s.get('direction', 'LONG')
    sign = 1 if direction == 'LONG' else -1
    ret_pct = round((ex_px - en_px) / en_px * 100 * sign, 2)
    lot = s.get('lot') or s.get('lot_size') or 1
    pnl = round(ret_pct / 100 * en_px * lot, 2)
    outcome = 'WIN' if ret_pct > 0 else 'LOSS'

    old_outcome = s.get('outcome')
    s['entry_px'] = round(float(en_px), 2)
    s['exit_px'] = round(float(ex_px), 2)
    s['entry_price'] = s['entry_px']
    s['exit_price'] = s['exit_px']
    s['actual_ret'] = ret_pct
    s['actual_return'] = ret_pct
    s['return_pct'] = ret_pct
    s['actual_pnl'] = pnl
    s['pnl'] = pnl
    s['outcome'] = outcome
    s['data_source'] = ('Rebuilt on REAL NSE result date (from data_center/nse_real_quarterly_result_dates.json) '
                         '+ real Yahoo Finance adjusted OHLC (2026-10-05), Open=entry/Close=exit; window/direction unchanged.')
    updated += 1
    if old_outcome != outcome:
        flips.append((q_code, s['symbol'], old_outcome, outcome))

print(f"\nUpdated {updated} price-dependent records ({skipped} skipped -- no price data at the new dates).")
print(f"Outcome FLIPPED for {len(flips)} records.")

# ---- 4. Recompute quarter-level aggregates ----
print("\nRecomputed quarter-level aggregates:")
for q in d['quarters']:
    if q['q_code'] not in TARGET_QS:
        continue
    n50_stocks = [s for s in q['stocks'] if s.get('is_nifty50') == 'YES']
    rets = [s['actual_ret'] for s in n50_stocks if s.get('actual_ret') is not None]
    pnls = [s['actual_pnl'] for s in n50_stocks if s.get('actual_pnl') is not None]
    wins = sum(1 for s in n50_stocks if s.get('outcome') == 'WIN')
    losses = sum(1 for s in n50_stocks if s.get('outcome') == 'LOSS')
    n = wins + losses
    old_wr = q.get('win_rate_avg')
    q['win_rate_avg'] = round(100 * wins / n, 1) if n else q.get('win_rate_avg')
    q['avg_win_rate'] = q['win_rate_avg']
    q['total_wins'] = wins
    q['total_losses'] = losses
    q['total_pnl'] = round(sum(pnls), 2) if pnls else q.get('total_pnl')
    n_long = sum(1 for s in n50_stocks if s.get('direction') == 'LONG')
    n_short = sum(1 for s in n50_stocks if s.get('direction') == 'SHORT')
    q['bias'] = f"{n_short} SHORT / {n_long} LONG"
    print(f"  {q['q_code']:<10} win_rate: {old_wr} -> {q['win_rate_avg']}   wins/losses={wins}/{losses} (of {n} scored; {len(n50_stocks)-n} still on unverified dates)")

json.dump(d, open(BASE + r'\dashboard_data\event_dashboard_data.json', 'w', encoding='utf-8'), indent=2)
print("\nSaved:", BASE + r'\dashboard_data\event_dashboard_data.json')
