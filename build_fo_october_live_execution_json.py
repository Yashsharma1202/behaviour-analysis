"""
FO OCTOBER SEASONALITY \u2014 LIVE EXECUTION JSON (October 2026).
=================================================================
Turns the validated, split-robust October seasonality backtest into a concrete
LIVE execution file for October 2026: each stock's own historically-best entry
trading-day-of-October is mapped to its REAL October-2026 calendar date using
the project's official NSE-2026 holiday calendar (same one used everywhere
else in the dashboard/live_pnl_server, for consistency \u2014 NOT the illustrative
2025 reference dates from the Excel report, which would be WRONG for 2026
because the holiday-day positions differ between years).

Exit for every position = the October-2026 month-end close (30-Oct-2026, the
20th and last NSE trading day of the month).

Source: intermediate split-robust results from fo_october_seasonality.py
(dashboard_data or scraped_parquet _tmp file), dashboard_data/fo_stocks_211.json
for sector/lot metadata. ADDITIVE: writes a new JSON only.
"""
import json, os
from datetime import date, timedelta
import pandas as pd

BASE = r'D:\behaviour analysis'
OUT = os.path.join(BASE, 'dashboard_data', 'fo_october_seasonality_live_2026.json')

# ---- SAME official NSE-2026 holiday calendar used across the whole project
# (event_trading_dashboard.html, live_pnl_server.py) \u2014 single source of truth ----
NSE_2026_HOLIDAYS = {
    # '2026-03-04' removed (not a real NSE holiday -- leftover from Holi's date
    # once being wrongly typed as 04-Mar; real Holi holiday is 03-Mar, kept
    # below). '2026-11-08'/'2026-11-10' (Diwali) added -- were missing entirely.
    '2026-01-26', '2026-03-03', '2026-03-26', '2026-03-31', '2026-04-03',
    '2026-04-14', '2026-05-01', '2026-05-28', '2026-06-26', '2026-09-14', '2026-10-02',
    '2026-10-20', '2026-11-08', '2026-11-10', '2026-11-24', '2026-12-25',
}


def is_trading_day(d):
    return d.weekday() < 5 and d.isoformat() not in NSE_2026_HOLIDAYS


def october_2026_trading_days():
    d = date(2026, 10, 1)
    out = []
    while d.month == 10:
        if is_trading_day(d):
            out.append(d)
        d += timedelta(days=1)
    return out


CAL_2026 = october_2026_trading_days()  # index 0 = trading-day 1
N_DAYS_2026 = len(CAL_2026)
EXIT_DATE_2026 = CAL_2026[-1]

# ---- validation of the calendar before using it for anything ----
assert N_DAYS_2026 >= 15, 'October 2026 trading-day count looks wrong: %d' % N_DAYS_2026
assert date(2026, 10, 2) not in CAL_2026, 'Gandhi Jayanti (2-Oct) must be excluded'
assert date(2026, 10, 20) not in CAL_2026, 'Dussehra (20-Oct) must be excluded'
assert all(CAL_2026[i] < CAL_2026[i + 1] for i in range(len(CAL_2026) - 1)), 'calendar must be strictly increasing'
assert EXIT_DATE_2026 == date(2026, 10, 30), 'month-end exit date mismatch: got %s' % EXIT_DATE_2026


def day_idx_to_2026_date(day_idx):
    """Map a historical trading-day-index (1-based) to its real Oct-2026 date.
    If a stock's best day-index exceeds this year's trading-day count (can
    happen since some years have up to 22 days), clamp to the last trading day
    so every position still has a valid, in-month entry date."""
    idx = min(int(day_idx), N_DAYS_2026) - 1
    return CAL_2026[idx]


# ---- load the split-robust backtest results ----
res = pd.read_parquet(os.path.join(BASE, 'scraped_parquet', '_tmp_oct_seasonality_results.parquet'))

# ---- basket-level (all stocks) recommendation ----
basket_yr = res.groupby(['Year', 'DayIdx'])['Ret'].mean().reset_index()
pivot = basket_yr.pivot(index='DayIdx', columns='Year', values='Ret')
stats = pivot.agg(['mean', 'median', 'std'], axis=1)
stats['win_rate'] = (pivot > 0).mean(axis=1) * 100
ex2019 = pivot.drop(columns=2019, errors='ignore')
stats['mean_ex2019'] = ex2019.mean(axis=1)
stats['win_rate_ex2019'] = (ex2019 > 0).mean(axis=1) * 100
stats['sharpe_like'] = stats['mean'] / stats['std']
best_day = int(stats['sharpe_like'].idxmax())
bstat = stats.loc[best_day]
basket_entry_date = day_idx_to_2026_date(best_day)

basket_reco = {
    'day_index': best_day,
    'entry_date': basket_entry_date.isoformat(),
    'entry_weekday': basket_entry_date.strftime('%A'),
    'exit_date': EXIT_DATE_2026.isoformat(),
    'exit_weekday': EXIT_DATE_2026.strftime('%A'),
    'direction': 'LONG',
    'entry_time': '09:20 AM',
    'exit_time': '03:15 PM',
    'historical_mean_return_pct': round(float(bstat['mean']), 2),
    'historical_mean_ex2019_pct': round(float(bstat['mean_ex2019']), 2),
    'historical_median_return_pct': round(float(bstat['median']), 2),
    'historical_win_rate_pct': round(float(bstat['win_rate']), 1),
    'historical_win_rate_ex2019_pct': round(float(bstat['win_rate_ex2019']), 1),
    'sharpe_like': round(float(bstat['sharpe_like']), 2),
    'years_backtested': int(pivot.loc[best_day].notna().sum()),
}

# ---- per-stock recommendations ----
fo = json.load(open(os.path.join(BASE, 'dashboard_data', 'fo_stocks_211.json'), encoding='utf-8'))
meta = {f['symbol']: f for f in fo}
ALL_FO_STOCKS = set(meta.keys())
STOCKS_WITH_ANY_DATA = set(res['Instrument'].unique())
zero_data_stocks = sorted(ALL_FO_STOCKS - STOCKS_WITH_ANY_DATA)

# fo_stocks_211.json has NO sector field at all. Sector classification in this
# project only exists for the Nifty-50 subset (event_dashboard_data.json quarters
# + rbi_policy_behaviour.json stock_analysis). Merge that in for the ~45 overlap
# stocks; the remaining non-Nifty-50 F&O names genuinely have no sector data
# anywhere in the project and are left as "" (not fabricated) with a coverage
# note so this is documented, not silently missing.
SECTOR = {}
_ev = json.load(open(os.path.join(BASE, 'dashboard_data', 'event_dashboard_data.json'), encoding='utf-8'))
for _q in _ev.get('quarters', []):
    for _s in _q.get('stocks', []):
        if _s.get('sector'):
            SECTOR[_s['symbol']] = _s['sector']
_rbi = json.load(open(os.path.join(BASE, 'dashboard_data', 'rbi_policy_behaviour.json'), encoding='utf-8'))
for _s in _rbi.get('stock_analysis', []):
    sym = _s.get('symbol')
    if sym and sym not in SECTOR and (_s.get('core_sector') or _s.get('sector')):
        SECTOR[sym] = _s.get('core_sector') or _s.get('sector')
SECTOR_FROM_PROJECT_DATA = set(SECTOR.keys())

# For F&O names outside the Nifty-50 (no sector anywhere in this project's data),
# hand-classify using standard NSE sector groupings, reusing the project's
# existing category labels where they fit and adding a few new ones only where
# none of the existing labels apply (Industrials & Capital Goods, Chemicals &
# Materials, Realty). Marked separately (sector_source) so it's never confused
# with the validated project-sourced sectors above.
MANUAL_SECTOR_MAP = {
    # Banking & Financials (incl. NBFC / insurance / AMC / exchanges, matching
    # how the project already buckets BAJFINANCE/SBILIFE/KOTAKBANK etc.)
    'ABCAPITAL': 'Banking & Financials', 'AUBANK': 'Banking & Financials',
    'BANDHANBNK': 'Banking & Financials', 'BANKBARODA': 'Banking & Financials',
    'CANBK': 'Banking & Financials', 'CHOLAFIN': 'Banking & Financials',
    'FEDERALBNK': 'Banking & Financials', 'HDFCAMC': 'Banking & Financials',
    'ICICIGI': 'Banking & Financials', 'ICICIPRULI': 'Banking & Financials',
    'IDFCFIRSTB': 'Banking & Financials', 'INDUSINDBK': 'Banking & Financials',
    'LICHSGFIN': 'Banking & Financials', 'MANAPPURAM': 'Banking & Financials',
    'MFSL': 'Banking & Financials', 'MUTHOOTFIN': 'Banking & Financials',
    'PFC': 'Banking & Financials', 'PNB': 'Banking & Financials',
    'RBLBANK': 'Banking & Financials', 'RECLTD': 'Banking & Financials',
    'SBICARD': 'Banking & Financials', 'MCX': 'Banking & Financials',
    'IEX': 'Banking & Financials',
    # Automobile & Auto Components
    'ASHOKLEY': 'Automobile & Auto Components', 'BHARATFORG': 'Automobile & Auto Components',
    'BOSCHLTD': 'Automobile & Auto Components', 'HEROMOTOCO': 'Automobile & Auto Components',
    'MOTHERSON': 'Automobile & Auto Components', 'TATAMOTORS': 'Automobile & Auto Components',
    'TVSMOTOR': 'Automobile & Auto Components',
    # Healthcare & Pharma
    'ALKEM': 'Healthcare & Pharma', 'AUROPHARMA': 'Healthcare & Pharma',
    'BIOCON': 'Healthcare & Pharma', 'DIVISLAB': 'Healthcare & Pharma',
    'GLENMARK': 'Healthcare & Pharma', 'LAURUSLABS': 'Healthcare & Pharma',
    'LUPIN': 'Healthcare & Pharma', 'TORNTPHARM': 'Healthcare & Pharma',
    'ZYDUSLIFE': 'Healthcare & Pharma',
    # Oil, Gas & Energy
    'BPCL': 'Oil, Gas & Energy', 'GAIL': 'Oil, Gas & Energy',
    'HINDPETRO': 'Oil, Gas & Energy', 'IOC': 'Oil, Gas & Energy',
    'PETRONET': 'Oil, Gas & Energy', 'TATAPOWER': 'Oil, Gas & Energy',
    # Metals & Mining
    'JINDALSTEL': 'Metals & Mining', 'NATIONALUM': 'Metals & Mining',
    'NMDC': 'Metals & Mining', 'SAIL': 'Metals & Mining', 'VEDL': 'Metals & Mining',
    # Cement & Construction
    'AMBUJACEM': 'Cement & Construction', 'SHREECEM': 'Cement & Construction',
    # FMCG & Consumption
    'BRITANNIA': 'FMCG & Consumption', 'COLPAL': 'FMCG & Consumption',
    'DABUR': 'FMCG & Consumption', 'GODREJCP': 'FMCG & Consumption',
    'MARICO': 'FMCG & Consumption', 'PIDILITIND': 'FMCG & Consumption',
    # Consumer Durables & Retail (incl. QSR / hospitality — consumer discretionary)
    'CROMPTON': 'Consumer Durables & Retail', 'DIXON': 'Consumer Durables & Retail',
    'HAVELLS': 'Consumer Durables & Retail', 'PAGEIND': 'Consumer Durables & Retail',
    'VOLTAS': 'Consumer Durables & Retail', 'JUBLFOOD': 'Consumer Durables & Retail',
    'INDHOTEL': 'Consumer Durables & Retail',
    # Information Technology
    'COFORGE': 'Information Technology', 'MPHASIS': 'Information Technology',
    'OFSS': 'Information Technology', 'NAUKRI': 'Information Technology',
    'PERSISTENT': 'Information Technology',
    # Telecom & Logistics
    'CONCOR': 'Telecom & Logistics', 'INDUSTOWER': 'Telecom & Logistics',
    'IDEA': 'Telecom & Logistics',
    # Industrials & Capital Goods (NEW bucket — no existing label fit)
    'ABB': 'Industrials & Capital Goods', 'BHEL': 'Industrials & Capital Goods',
    'CUMMINSIND': 'Industrials & Capital Goods', 'SIEMENS': 'Industrials & Capital Goods',
    'POLYCAB': 'Industrials & Capital Goods', 'ASTRAL': 'Industrials & Capital Goods',
    # Chemicals & Materials (NEW bucket)
    'SRF': 'Chemicals & Materials', 'UPL': 'Chemicals & Materials', 'PIIND': 'Chemicals & Materials',
    # Realty (NEW bucket)
    'DLF': 'Realty', 'GODREJPROP': 'Realty', 'OBEROIRLTY': 'Realty',
    # Other / Diversified (matches how the project already classifies BEL, a similar PSU)
    'HAL': 'Other / Diversified',
}
for _sym, _sec in MANUAL_SECTOR_MAP.items():
    SECTOR.setdefault(_sym, _sec)

stock_rows = []
skipped_low_data = []
for sym, g in res.groupby('Instrument'):
    piv = g.pivot_table(index='DayIdx', columns='Year', values='Ret', aggfunc='mean')
    nyears = piv.shape[1]
    if nyears < 4:
        skipped_low_data.append(sym)
        continue
    mean_, std_ = piv.mean(axis=1), piv.std(axis=1)
    wr = (piv > 0).mean(axis=1) * 100
    sharpe = mean_ / std_
    bd = int(sharpe.idxmax())
    entry_dt = day_idx_to_2026_date(bd)
    m = meta.get(sym, {})
    edge = 'positive' if sharpe[bd] > 0.15 else ('weak' if sharpe[bd] > 0 else 'negative')
    stock_rows.append({
        'symbol': sym,
        'name': m.get('name', sym),
        'sector': SECTOR.get(sym, ''),
        'sector_source': ('project_data' if sym in SECTOR_FROM_PROJECT_DATA
                          else ('manual_classification' if sym in SECTOR else 'unclassified')),
        'lot_size': m.get('lot_size'),
        'is_nifty50': m.get('is_nifty50', 'NO'),
        'direction': 'LONG',
        'day_index': bd,
        'entry_date': entry_dt.isoformat(),
        'entry_weekday': entry_dt.strftime('%A'),
        'exit_date': EXIT_DATE_2026.isoformat(),
        'exit_weekday': EXIT_DATE_2026.strftime('%A'),
        'entry_time': '09:20 AM',
        'exit_time': '03:15 PM',
        'historical_mean_return_pct': round(float(mean_[bd]), 2),
        'historical_median_return_pct': round(float(piv.loc[bd].median()), 2),
        'historical_win_rate_pct': round(float(wr[bd]), 1),
        'sharpe_like': round(float(sharpe[bd]), 2),
        'years_backtested': nyears,
        'edge_quality': edge,
    })

# sort strongest edge first
stock_rows.sort(key=lambda r: -r['sharpe_like'])

out = {
    'generated_at': date.today().isoformat(),
    'event': 'October F&O Seasonality \u2014 Live Execution',
    'month': 'October', 'year': 2026,
    'source': 'Split-robust backtest on fo_futures_near_month_continuous.parquet (211 F&O stocks, near-month futures, 2019-2025)',
    'methodology': (
        'For each stock, find the historically best trading-day-of-October to enter a LONG '
        '(risk-adjusted / Sharpe-like ranked, not raw highest-mean \u2014 see caveat), held to the '
        'October month-end close. Corporate-action (split/bonus) days excluded from return '
        'calculation. Entry dates below are mapped onto the REAL October-2026 NSE trading '
        'calendar (2-Oct Gandhi Jayanti and 20-Oct Dussehra excluded), not a prior year\u2019s dates.'
    ),
    'trading_calendar_2026': {
        'total_trading_days': N_DAYS_2026,
        'holidays_excluded': ['2026-10-02 (Gandhi Jayanti, Friday)', '2026-10-20 (Dussehra, Tuesday)'],
        'first_trading_day': CAL_2026[0].isoformat(),
        'month_end_exit_date': EXIT_DATE_2026.isoformat(),
    },
    'caveats': [
        'Backtest sample is only 7 Octobers (2019-2025); per-stock results need >=4 years and are still a small sample.',
        'The 2019 outlier (corporate-tax-cut rally) inflates raw mean returns for early-month entries; ranking here uses risk-adjusted (Sharpe-like) score, with the ex-2019 mean/win-rate also reported at the basket level for transparency.',
        'Prices are near-month continuous FUTURES closes, not exact intraday fill prices; entry/exit times (09:20/15:15) are the project convention, not guaranteed fills.',
        'This is Phase 1 (October only). Not yet extended to other months.',
    ],
    'basket_recommendation': basket_reco,
    'stocks': stock_rows,
    'stocks_skipped_insufficient_history': skipped_low_data,
    'stocks_with_zero_october_data': zero_data_stocks,
    'coverage_summary': {
        'total_fo_stocks': len(ALL_FO_STOCKS),
        'stocks_with_live_entry': len(stock_rows),
        'stocks_skipped_insufficient_history_1to3yr': len(skipped_low_data),
        'stocks_with_zero_history': len(zero_data_stocks),
        'stocks_with_sector_classified': sum(1 for s in stock_rows if s['sector']),
        'stocks_sector_from_project_data': sum(1 for s in stock_rows if s['sector_source'] == 'project_data'),
        'stocks_sector_manually_classified': sum(1 for s in stock_rows if s['sector_source'] == 'manual_classification'),
        'sector_note': ('Sector for Nifty-50 names (sector_source="project_data") comes from the project’s own '
                       'validated data. Sector for the remaining F&O names (sector_source="manual_classification") '
                       'was hand-classified using standard NSE sector groupings, since no sector data exists for '
                       'them anywhere in this project — check sector_source per stock before relying on it for '
                       'risk/diversification decisions.'),
    },
}

# ---- final validation before writing ----
errs = []
if not (1 <= basket_reco['day_index'] <= N_DAYS_2026):
    errs.append('basket day_index out of range')
if date.fromisoformat(basket_reco['entry_date']) > date.fromisoformat(basket_reco['exit_date']):
    errs.append('basket entry_date is after exit_date')
seen = set()
for s in stock_rows:
    if s['symbol'] in seen:
        errs.append('duplicate symbol: ' + s['symbol'])
    seen.add(s['symbol'])
    ed, xd = date.fromisoformat(s['entry_date']), date.fromisoformat(s['exit_date'])
    if ed > xd:
        errs.append('%s: entry_date after exit_date' % s['symbol'])
    if not is_trading_day(ed):
        errs.append('%s: entry_date %s is not a valid Oct-2026 trading day' % (s['symbol'], ed))
    if xd != EXIT_DATE_2026:
        errs.append('%s: exit_date does not match month-end' % s['symbol'])
if len(stock_rows) + len(skipped_low_data) != res['Instrument'].nunique():
    errs.append('stock count mismatch: %d + %d != %d' % (len(stock_rows), len(skipped_low_data), res['Instrument'].nunique()))
if len(stock_rows) + len(skipped_low_data) + len(zero_data_stocks) != len(ALL_FO_STOCKS):
    errs.append('full universe accounting mismatch: %d + %d + %d != %d'
                % (len(stock_rows), len(skipped_low_data), len(zero_data_stocks), len(ALL_FO_STOCKS)))
unclassified = [s['symbol'] for s in stock_rows if not s['sector']]
if unclassified:
    errs.append('stocks still missing sector after manual classification: ' + ', '.join(unclassified))
if errs:
    raise SystemExit('VALIDATION FAILED:\n' + '\n'.join(errs))

os.makedirs(os.path.dirname(OUT), exist_ok=True)
with open(OUT, 'w', encoding='utf-8') as f:
    json.dump(out, f, indent=2, ensure_ascii=False)

# ---- round-trip check: re-read and re-verify ----
reloaded = json.load(open(OUT, encoding='utf-8'))
assert len(reloaded['stocks']) == len(stock_rows)
assert reloaded['basket_recommendation']['entry_date'] == basket_reco['entry_date']

print('VALIDATION PASSED. Saved:', OUT)
print('October 2026: %d trading days, month-end exit %s' % (N_DAYS_2026, EXIT_DATE_2026))
print('Basket recommendation: Day-Index %d -> entry %s (%s), exit %s (%s)'
      % (basket_reco['day_index'], basket_reco['entry_date'], basket_reco['entry_weekday'],
         basket_reco['exit_date'], basket_reco['exit_weekday']))
print('Stocks with live entry dates: %d | skipped (1-3yr history): %d | zero history: %d | total F&O universe: %d'
      % (len(stock_rows), len(skipped_low_data), len(zero_data_stocks), len(ALL_FO_STOCKS)))
print('Top 5 stocks by edge:')
for s in stock_rows[:5]:
    print('  %-11s day%-3d entry %s (%s)  mean %+5.2f%%  wr %5.1f%%  sharpe %+.2f'
          % (s['symbol'], s['day_index'], s['entry_date'], s['entry_weekday'],
             s['historical_mean_return_pct'], s['historical_win_rate_pct'], s['sharpe_like']))
