"""
REBUILD THE LAST 12 QUARTERS' TRADE OUTCOMES FROM REAL PRICES.

Per explicit user decision: the entry/exit prices stored for at least
FY24_Q3/FY24_Q4 don't match real market data from any checked source
(NSE cash, NSE futures, or even a fresh Yahoo Finance pull, accounting for
adjusted-price drift). Rather than trying to reproduce the original (now
unreproducible) adjustment snapshot, this rebuilds ALL 12 trailing quarters'
price-dependent outcome fields fresh, on today's basis.

KEPT AS-IS (not touched): entry_date, exit_date, result_date, window,
lead, hold, direction -- these are the strategy's own parameters/dates,
already independently date-validated earlier this session, not something
to re-derive from price data.

REBUILT (price-dependent only): entry_px, exit_px, actual_ret, actual_pnl,
pnl, outcome, entry_price, exit_price, return_pct -- using real Yahoo
Finance adjusted OHLC for the stock's own already-trusted entry/exit date,
Open=entry (09:20 proxy), Close=exit (15:15 proxy), same convention used
throughout this project.

ADDITIVE: reads event_dashboard_data.json + fetches fresh prices; updates
price-dependent fields in place for the 12 trailing quarters.
"""
import json
import yfinance as yf
import pandas as pd
from datetime import datetime

BASE = r'D:\behaviour analysis'
d = json.load(open(BASE + r'\dashboard_data\event_dashboard_data.json', encoding='utf-8'))

LAST12 = ['FY27_Q2', 'FY27_Q1', 'FY26_Q4', 'FY26_Q3', 'FY26_Q2', 'FY26_Q1',
          'FY25_Q4', 'FY25_Q3', 'FY25_Q2', 'FY25_Q1', 'FY24_Q4', 'FY24_Q3']

syms = set()
for q in d['quarters']:
    if q['q_code'] in LAST12:
        for s in q['stocks']:
            syms.add(s['symbol'])
syms = sorted(syms)
print(f"Fetching real price history for {len(syms)} symbols...")

price_data = {}
failed = []
for sym in syms:
    ticker = sym + '.NS'
    try:
        h = yf.Ticker(ticker).history(start='2023-09-01', end='2026-09-10', auto_adjust=True)
        if h.empty:
            failed.append(sym)
            continue
        h.index = h.index.date
        price_data[sym] = h
    except Exception as e:
        failed.append(sym)
        print(f"  FAILED {sym}: {e}")

print(f"Fetched {len(price_data)} of {len(syms)} symbols. Failed: {failed}")


def parse_dmy(s):
    return datetime.strptime(s.split(' (')[0], '%d-%b-%Y').date()


def nearest_trading_price(hist, target_date, field, direction='forward'):
    """Find target_date in the index; if not a trading day, step forward
    (entry) or backward (exit) to the nearest real trading day, mirroring
    how a real order would fill on the next/prior session."""
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


updated = 0
skipped = 0
details = []
for q in d['quarters']:
    if q['q_code'] not in LAST12:
        continue
    for s in q['stocks']:
        sym = s['symbol']
        hist = price_data.get(sym)
        if hist is None or not s.get('entry_date') or not s.get('exit_date'):
            skipped += 1
            continue
        try:
            en_date = parse_dmy(s['entry_date'])
            ex_date = parse_dmy(s['exit_date'])
        except Exception:
            skipped += 1
            continue

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

        old_pnl = s.get('actual_pnl')
        old_ret = s.get('actual_ret')
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
        s['data_source'] = 'Rebuilt from real Yahoo Finance adjusted OHLC (2026-10-05), Open=entry/Close=exit on the already-validated entry/exit date; window/direction unchanged.'
        updated += 1
        if old_outcome != outcome:
            details.append((q['q_code'], sym, old_outcome, outcome, old_ret, ret_pct))

print(f"\nUpdated {updated} records across 12 quarters ({skipped} skipped - no price data / bad dates).")
print(f"Outcome FLIPPED (WIN<->LOSS) for {len(details)} records:")
for det in details[:40]:
    print(' ', det)

# ---- Recompute quarter-level aggregates from the rebuilt per-stock data ----
for q in d['quarters']:
    if q['q_code'] not in LAST12:
        continue
    rets = [s['actual_ret'] for s in q['stocks'] if s.get('actual_ret') is not None]
    pnls = [s['actual_pnl'] for s in q['stocks'] if s.get('actual_pnl') is not None]
    wins = sum(1 for s in q['stocks'] if s.get('outcome') == 'WIN')
    losses = sum(1 for s in q['stocks'] if s.get('outcome') == 'LOSS')
    n = wins + losses
    q['win_rate_avg'] = round(100 * wins / n, 1) if n else None
    q['avg_win_rate'] = q['win_rate_avg']
    q['total_wins'] = wins
    q['total_losses'] = losses
    q['total_pnl'] = round(sum(pnls), 2)
    n_long = sum(1 for s in q['stocks'] if s.get('direction') == 'LONG')
    n_short = sum(1 for s in q['stocks'] if s.get('direction') == 'SHORT')
    q['bias'] = f"{n_short} SHORT / {n_long} LONG"

print("\nRecomputed quarter-level aggregates:")
for q in d['quarters']:
    if q['q_code'] in LAST12:
        print(f"  {q['q_code']:<10} win_rate_avg={q['win_rate_avg']}  total_pnl={q['total_pnl']}  wins/losses={q['total_wins']}/{q['total_losses']}")

json.dump(d, open(BASE + r'\dashboard_data\event_dashboard_data.json', 'w', encoding='utf-8'), indent=2)
print("\nSaved:", BASE + r'\dashboard_data\event_dashboard_data.json')
