"""
OCTOBER SEASONALITY \u2014 F&O Stocks, Phase 1 (single month, split-robust).
========================================================================
Concept: find the best trading-day-of-October to ENTER a position, held to the
October month-end close, using the real F&O universe (211 stocks, near-month
continuous futures). Trading days are taken exactly as they appear in the data
(NSE weekends + holidays are already absent from the source \u2014 verified, no
synthetic calendar needed).

CRITICAL FIX (found + applied): raw closes are NOT split/bonus adjusted. A
corporate action shows as a fake single-day move of 50-90% (e.g. DRREDDY 5:1
split on 28-Oct-2024 crashed the close from Rs 6512 -> Rs 1310; RELIANCE also
had one on the same day). All returns here are computed by compounding DAILY
% changes from entry to month-end and SKIPPING any single day where
|daily change| > 20% (treated as a split/bonus, not a real price move) \u2014
same method already validated earlier in this project.

OUTPUT: fo_october_seasonality.py prints validated results; a separate builder
writes the Excel (kept separate so this file can be re-run cheaply for checks).
"""
import pandas as pd, numpy as np

BASE = r'D:\behaviour analysis'
SPLIT_CAP = 0.20  # skip any single-day |chg| > 20% when compounding (corporate action)
MIN_YEARS_STOCK = 4  # per-stock results need at least this many Octobers to report

d = pd.read_parquet(BASE + r'\scraped_parquet\fo_futures_near_month_continuous.parquet')
d['Date'] = pd.to_datetime(d['Date'])
cov = pd.read_csv(BASE + r'\scraped_parquet\fo_futures_coverage.csv')
STOCK_SYMS = set(cov[cov['CATEGORY'] == 'Stock']['SYMBOL'])

d = d[d['Instrument'].isin(STOCK_SYMS)].sort_values(['Instrument', 'Date']).reset_index(drop=True)
d['PrevClose'] = d.groupby('Instrument')['Close'].shift(1)
d['DailyChg'] = (d['Close'] - d['PrevClose']) / d['PrevClose']
d['SplitDay'] = d['DailyChg'].abs() > SPLIT_CAP

oct = d[d['Date'].dt.month == 10].copy()
oct['Year'] = oct['Date'].dt.year
oct = oct.sort_values(['Instrument', 'Year', 'Date'])
oct['DayIdx'] = oct.groupby(['Instrument', 'Year']).cumcount() + 1
last_idx = oct.groupby(['Instrument', 'Year'])['DayIdx'].transform('max')
oct['HoldDays'] = last_idx - oct['DayIdx']


def split_robust_ret(instr, year, entry_date, exit_date):
    """Compound daily % changes from entry_date (exclusive of that day's own
    change) to exit_date, skipping any split day. Returns (ret%, had_split)."""
    seg = d[(d['Instrument'] == instr) & (d['Date'] > entry_date) & (d['Date'] <= exit_date)]
    if seg.empty:
        return 0.0, False
    g = 1.0; had_split = False
    for chg, is_split in zip(seg['DailyChg'], seg['SplitDay']):
        if pd.isna(chg):
            continue
        if is_split:
            had_split = True
            continue
        g *= (1 + chg)
    return (g - 1) * 100, had_split


# ---- build split-robust return for every (Instrument, Year, DayIdx), keeping
# real entry/exit DATES and CLOSE PRICES for full auditability ----
monthend = oct.groupby(['Instrument', 'Year']).agg(
    MonthEndDate=('Date', 'max')).reset_index()
oct = oct.merge(monthend, on=['Instrument', 'Year'])
exit_close = oct[oct['Date'] == oct['MonthEndDate']][['Instrument', 'Year', 'Close']].rename(columns={'Close': 'ExitClose'})
oct = oct.merge(exit_close, on=['Instrument', 'Year'])

rows = []
for (instr, yr), g in oct.groupby(['Instrument', 'Year']):
    me_date = g['MonthEndDate'].iloc[0]
    exit_px = g['ExitClose'].iloc[0]
    for _, r in g.iterrows():
        ret, had_split = split_robust_ret(instr, yr, r['Date'], me_date)
        rows.append((instr, yr, r['DayIdx'], r['HoldDays'], r['Date'], r['Close'],
                    me_date, exit_px, ret, had_split))
res = pd.DataFrame(rows, columns=['Instrument', 'Year', 'DayIdx', 'HoldDays',
                                   'EntryDate', 'EntryClose', 'ExitDate', 'ExitClose',
                                   'Ret', 'HadSplit'])
res = res[res['DayIdx'] <= 19]

print('=== VALIDATION: DRREDDY Oct 2024, split-robust (should be small, not -80%) ===')
dr = res[(res['Instrument'] == 'DRREDDY') & (res['Year'] == 2024)]
print(dr[['DayIdx', 'EntryDate', 'EntryClose', 'ExitDate', 'ExitClose', 'Ret', 'HadSplit']].round(2).to_string(index=False))

n_split_windows = res['HadSplit'].sum()
print('\nStock-year-day windows containing a split/bonus (raw price ratio would be WRONG, Ret column is fixed): %d of %d' % (n_split_windows, len(res)))

# ---- BASKET LEVEL: average across stocks WITHIN each year, then across years ----
basket_yr = res.groupby(['Year', 'DayIdx'])['Ret'].mean().reset_index()
pivot = basket_yr.pivot(index='DayIdx', columns='Year', values='Ret')
avg_hold = res.groupby('DayIdx')['HoldDays'].mean()

stats = pivot.agg(['mean', 'median', 'std'], axis=1)
stats['win_rate'] = (pivot > 0).mean(axis=1) * 100
ex2019 = pivot.drop(columns=2019, errors='ignore')
stats['mean_ex2019'] = ex2019.mean(axis=1)
stats['win_rate_ex2019'] = (ex2019 > 0).mean(axis=1) * 100
stats['avg_hold_days'] = avg_hold.reindex(stats.index)
stats['sharpe_like'] = stats['mean'] / stats['std']
stats = stats.sort_values('sharpe_like', ascending=False)

print('\n=== BASKET (split-robust) \u2014 ranked by risk-adjusted (Sharpe-like) score ===')
cols = ['mean', 'mean_ex2019', 'median', 'std', 'sharpe_like', 'win_rate', 'win_rate_ex2019', 'avg_hold_days']
print(stats[cols].round(2).head(12).to_string())

print('\n=== Full Day-Index x Year matrix (split-robust basket avg %) ===')
print(pivot.round(2).to_string())

# save for the Excel builder
res.to_parquet(BASE + r'\scraped_parquet\_tmp_oct_seasonality_results.parquet')
print('\nSaved intermediate results for Excel builder.')
