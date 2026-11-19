"""Live 1-Jan-2027 list: stocks with ALL 15 New Years (2012-2026) valid at the window, net win rate >= 75%, best window per stock, top 50 by t-stat (15-yr, net of 0.10% cost)."""
import sys
src = open('D:/behaviour analysis/new_year_top50_best_ls.py', encoding='utf-8').read().split('OUT = ')[0]
exec(src, globals())
d = D[D.Year >= 2012]; out = []
for dr, col in (('LONG', 'Net_L'), ('SHORT', 'Net_S')):
    g = d.groupby(['Symbol', 'N', 'H'])[col].agg(Yrs='count', Avg='mean', Sd='std', Win=lambda v: (v > 0).mean()).reset_index(); g['Dir'] = dr; out.append(g)
c = pd.concat(out); c = c[(c.Yrs == 15) & (c.Win >= 0.75) & (c.Avg > 0) & (c.Sd > 0)]; c['t'] = c.Avg / c.Sd * np.sqrt(c.Yrs)
c = c.sort_values('t', ascending=False).drop_duplicates('Symbol')
print('eligible stocks', len(c), c.Dir.value_counts().to_dict()); print(c.head(60).round(3).to_string())
c.to_csv('New_Year_15y_wr75_candidates.csv', index=False)
