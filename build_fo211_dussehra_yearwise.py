"""
ALL 211 F&O STOCKS — DUSSEHRA, YEAR-WISE, RANKED.
For every F&O stock (its own existing T-n/T+m window from holidays_dataset.json),
computes the split-robust year-by-year return using its best available price
history: 26yr cash data (55 stocks) where it exists, else near-month F&O futures
(up to 7yr, the remaining 156 stocks). Direction = best of LONG/SHORT by win rate
(same method as the Nifty-50 Dussehra report). Output: a Stock x Year performance
grid, ranked by average return, with full transparency on data source/sample size
per stock (no fabricated years).

ADDITIVE: reads existing data files; writes new report files only.
"""
import json, statistics as st
from datetime import date
import pandas as pd

BASE = r'D:\behaviour analysis'

# ---- load both price sources ----
cash = pd.read_parquet(BASE + r'\scraped_parquet\nifty50_all_stocks_daily_2000_2026.parquet', columns=['DATE', 'SYMBOL', 'CLOSE'])
cash['DATE'] = pd.to_datetime(cash['DATE']).dt.date
cash_closes = {s: g.sort_values('DATE')[['DATE', 'CLOSE']].values.tolist() for s, g in cash.groupby('SYMBOL')}

fut = pd.read_parquet(BASE + r'\scraped_parquet\fo_futures_near_month_continuous.parquet', columns=['Date', 'Instrument', 'Close'])
fut['Date'] = pd.to_datetime(fut['Date']).dt.date
fut_closes = {s: g.sort_values('Date')[['Date', 'Close']].values.tolist() for s, g in fut.groupby('Instrument')}

# ---- Dussehra historical anchor dates (26 years, 2000-2025) ----
beh = json.load(open(BASE + r'\dashboard_data\nifty_futures_holiday_behaviour.json', encoding='utf-8'))
dh = [h for h in beh['holidays'] if 'Dussehra' in h['name']][0]
dussehra_dates = sorted(date.fromisoformat(t['holiday_date']) for t in dh['trades'])

# ---- windows for all 211 stocks ----
# For the 50 Nifty-50 names, reuse the ALREADY-VERIFIED live dashboard window
# (event_dashboard_data.json, just fixed/pushed) so this report doesn't show
# different numbers than the live site for the same stocks. For the other 161
# F&O-only names (no live window exists), use the F&O-211 dataset's window.
hds = json.load(open(BASE + r'\dashboard_data\holidays_dataset.json', encoding='utf-8'))
du = [h for h in hds if h.get('id') == 'dussehra'][0]
windows = {s['symbol']: (s['entry_lead_days'], s['exit_hold_days'], s.get('lot_size')) for s in du['top_stocks']}

live = json.load(open(BASE + r'\dashboard_data\event_dashboard_data.json', encoding='utf-8'))
live_dussehra = [h for h in live['holidays'] if h['id'] == 'dussehra_2026'][0]
window_source = {}
for s in live_dussehra['stocks']:
    windows[s['symbol']] = (s['lead'], s['hold'], s.get('lot'))
    window_source[s['symbol']] = 'live-dashboard'
for sym in windows:
    window_source.setdefault(sym, 'fo211-dataset')

fo = json.load(open(BASE + r'\dashboard_data\fo_stocks_211.json', encoding='utf-8'))
meta = {f['symbol']: f for f in fo}


def raw_ret(closes, sym, hd, lead, hold):
    a = closes.get(sym)
    if not a: return None
    ib = -1
    for i, (d, _) in enumerate(a):
        if d <= hd: ib = i
        else: break
    if ib < 0: return None
    en, ex = ib - lead, ib + hold
    if en < 1 or ex >= len(a): return None
    g = 1.0
    for k in range(en, ex + 1):
        p0, p1 = a[k - 1][1], a[k][1]
        if not p0: return None
        dr = (p1 - p0) / p0
        if abs(dr) > 0.20: continue
        g *= (1 + dr)
    return (g - 1) * 100


results = {}   # symbol -> dict(source, n, year_rets{year:raw_move}, pick, wr, avg)
for sym, (lead, hold, lot) in windows.items():
    if sym in cash_closes:
        source = 'cash-26yr'
        closes = cash_closes
        dates_to_use = dussehra_dates
    else:
        source = 'futures-7yr'
        closes = fut_closes
        dates_to_use = [d for d in dussehra_dates if d.year >= 2019]

    year_rets = {}
    for hd in dates_to_use:
        v = raw_ret(closes, sym, hd, lead, hold)
        if v is not None:
            year_rets[hd.year] = v
    n = len(year_rets)
    wsrc = window_source.get(sym, 'fo211-dataset')
    if n < 1:
        # genuinely zero price history for any Dussehra date -> nothing computable
        results[sym] = dict(source=source, window_source=wsrc, n=n, lead=lead, hold=hold, lot=lot, year_rets=year_rets, pick=None)
        continue
    vals = list(year_rets.values())
    long_wins = sum(1 for v in vals if v > 0)
    long_wr = round(100 * long_wins / n, 1)
    long_avg = round(st.mean(vals), 2)
    short_wr = round(100 - long_wr, 1)
    short_avg = round(-long_avg, 2)
    # with n<3 a "higher win rate" pick is mostly a coin flip; still show it (transparently
    # flagged via 'n' and a low_confidence marker), never silently drop real data.
    pick = 'LONG' if long_wr >= short_wr else 'SHORT'
    results[sym] = dict(source=source, window_source=wsrc, n=n, lead=lead, hold=hold, lot=lot, year_rets=year_rets,
                        pick=pick, long_wr=long_wr, long_avg=long_avg, short_wr=short_wr, short_avg=short_avg,
                        pick_wr=(long_wr if pick == 'LONG' else short_wr),
                        pick_avg=(long_avg if pick == 'LONG' else short_avg),
                        low_confidence=(n < 3))

# JIOFIN and BEL: the live dashboard already used the F&O-211 4yr FALLBACK
# (insufficient cash-26yr data: n=3/n=0) rather than the generic futures-based
# recompute this script would otherwise produce for them. Force these 2 to
# match the already-verified live dashboard exactly, so this report never
# shows different numbers than the live site for a Nifty-50 stock.
LIVE_OVERRIDE = {
    'JIOFIN': dict(pick='SHORT', pick_wr=100.0, pick_avg=None, source='fo211-4yr-fallback (live dashboard)'),
    'BEL':    dict(pick='LONG', pick_wr=75.0, pick_avg=None, source='fo211-4yr-fallback (live dashboard)'),
}
for sym, ov in LIVE_OVERRIDE.items():
    if sym in results:
        results[sym].update(ov)
        results[sym]['n'] = 4  # F&O-211 fallback sample size
        results[sym]['override_note'] = 'matches live dashboard (insufficient 26yr cash data)'

qualified = {s: r for s, r in results.items() if r['pick']}
insufficient = {s: r for s, r in results.items() if not r['pick']}
low_conf = sum(1 for r in qualified.values() if r.get('low_confidence'))
print(f"Total F&O stocks: {len(results)} | qualified (>=1yr): {len(qualified)} ({low_conf} low-confidence, n<3) | zero-data: {len(insufficient)}")
sources = {}
for r in qualified.values(): sources[r['source']] = sources.get(r['source'], 0) + 1
print("Source breakdown (qualified):", sources)

import pickle
pickle.dump(dict(results=results, meta=meta), open(BASE + r'\scraped_parquet\_tmp_fo211_dussehra.pkl', 'wb'))
print("Saved intermediate results.")

ranked = sorted(qualified.items(), key=lambda kv: -(kv[1]['pick_avg'] if kv[1]['pick_avg'] is not None else -999))
print("\nTop 10 by avg return:")
for sym, r in ranked[:10]:
    avg_str = f"{r['pick_avg']:+6.2f}%" if r['pick_avg'] is not None else "  n/a "
    print(f"  {sym:12} {r['pick']:5} wr={r['pick_wr']:5.1f}% avg={avg_str} n={r['n']:2} src={r['source']}")
