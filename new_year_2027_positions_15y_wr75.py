"""1-Jan-2027 position book: TOP-50 stocks, each with ALL 15 New Years (2012-2026) valid at its window and 15-yr net win rate >= 75%.
Best window/direction per stock by 15-yr t-stat (net of 0.10% cost). New outputs only."""
s = open('D:/behaviour analysis/new_year_2027_positions.py', encoding='utf-8').read()
inject = '''exec(src, globals())
_d = D[D.Year >= 2012]; _o = []
for _dr, _col in (('LONG', 'Net_L'), ('SHORT', 'Net_S')):
    _g = _d.groupby(['Symbol', 'N', 'H'])[_col].agg(Yrs='count', Avg='mean', Sd='std', Win=lambda v: (v > 0).mean()).reset_index(); _g['Dir'] = _dr; _o.append(_g)
_c = pd.concat(_o); _c = _c[(_c.Yrs == 15) & (_c.Win >= 0.75) & (_c.Avg > 0) & (_c.Sd > 0)]; _c['t'] = _c.Avg / _c.Sd * np.sqrt(_c.Yrs)
_c = _c.sort_values('t', ascending=False).drop_duplicates('Symbol').head(50).reset_index(drop=True)
LV = pd.DataFrame({'Rank': range(1, len(_c) + 1), 'Symbol': _c.Symbol, 'Dir': _c.Dir, 'Enter_N_days_before_T': _c.N.astype(int), 'Hold_days': _c.H.astype(int), 'Years_of_data': _c.Yrs.astype(int),
                   'Avg_Net_pct': 100 * _c.Avg, 'Win_pct': 100 * _c.Win, 'Per_hold_day_pct': 100 * _c.Avg / _c.H, 't': _c.t})
print('LV', len(LV), LV.Win_pct.min())'''
assert 'exec(src, globals())' in s
s = s.replace('exec(src, globals())', inject, 1)
s = s.replace('New_Year_1Jan2027_Position_Book', 'New_Year_1Jan2027_Position_Book_Top50_15yr_WR75')
s = s.replace('The backtest of this rule averaged +1.83% a year (2012-26, 14 of 15 years positive).', 'Live list = stocks with all 15 New Years valid and net win rate >= 75% over 2012-26 at the chosen window (in-sample, so optimistic; the walk-forward book returned +1.83% a year, 14 of 15 years positive).')
s = s.replace('backtest of this rule: +1.83%/yr, 14 of 15 years positive', 'all 15 yrs of data, 15-yr net win rate >= 75% (in-sample)')
s = s.replace('ranked on ALL events 2006-2026 (LV)', 'ranked on 2012-2026 (LV)')
exec(compile(s, 'wr75', 'exec'), globals())
