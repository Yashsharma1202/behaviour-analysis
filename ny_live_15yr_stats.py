"""Per-stock 15-yr (2012-2026) NET win rate at the live 1-Jan-2027 window, plus walk-forward OOS stats."""
import sys
BOOK = sys.argv[1] if len(sys.argv) > 1 else 'New_Year_1Jan2027_Position_Book.xlsx'; CSV = sys.argv[2] if len(sys.argv) > 2 else 'New_Year_live_15yr_stats.csv'
src = open('D:/behaviour analysis/new_year_top50_best_ls.py', encoding='utf-8').read().split('OUT = ')[0]
exec(src, globals())
pb = pd.read_excel(BOOK, sheet_name='Position_Book', header=3).dropna(subset=['Symbol'])
out = []
for _, r in pb.iterrows():
    s, n, h = r.Symbol, int(r['Enter N before T']), int(r['Hold days'])
    d = D[(D.Symbol == s) & (D.N == n) & (D.H == h) & (D.Year >= 2012)]
    col = 'Net_L' if r.Direction == 'LONG' else 'Net_S'
    t = TR[TR.Symbol == s]
    out.append(dict(Symbol=s, N15=len(d), WR15=100 * (d[col] > 0).mean() if len(d) else None, AVG15=100 * d[col].mean() if len(d) else None,
                    OOS_n=len(t), OOS_wr=100 * (t.Net_pct > 0).mean() if len(t) else None))
o = pd.DataFrame(out); o.to_csv(CSV, index=False)
print(o.describe().T[['count','mean','min','max']]); print(o.head(8))
