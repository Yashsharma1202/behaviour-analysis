"""
Long-form Diwali strategy report in the format of SENSEX_Straddle_Final_Strategy_Report.pdf.
Reads portfolio_cache/diwali_report_8x8.pkl (built by diwali_report_data.py). Charts are hand-drawn inline SVG.
Output: Diwali_Final_Strategy_Report.pdf (+ .html). Additive only.
"""
import sys, pickle, json, base64
sys.stdout.reconfigure(encoding='utf-8')
import numpy as np, pandas as pd
ROOT = 'D:/behaviour analysis/'
R = pickle.load(open(ROOT + 'portfolio_cache/diwali_report_8x8.pkl', 'rb'))
M, MN, MG, POS, BOOKS, YEARS, TD = R['M'], R['MN'], R['MARGIN'], R['POS'], R['BOOKS'], R['YEARS'], R['T_DATE']
NOTIONAL, LASTC, LOTS = R['notional'], R['last_close'], R['lots']
yt = R['year_tbl']; fh = R['fh']
NAVY, BLUE = '#1f3a5f', '#2b6cb0'
repick = {}
exp3 = json.load(open(ROOT + 'Diwali_2026_3_LowHistory_Stocks_Positions_8x8.json', encoding='utf-8'))['stocks']
assert all(p['direction'] == 'LONG' for p in POS), 'report text assumes an all-LONG book'
LS = R['LS']; ALL50 = R['ALL50']
EX_SYM = sorted([p for p in POS if p['conv'] == 'STRONG'], key=lambda p: -p['wr'])[0]['symbol']

# ---------------------------------------------------------------- formatting
mn = lambda v: ('−' if v < 0 else '')
inr = lambda v: ('−' if v < 0 else '') + '₹' + f'{abs(v):,.0f}'
sinr = lambda v: ('+' if v >= 0 else '−') + '₹' + f'{abs(v):,.0f}'
lac = lambda v: ('−' if v < 0 else '') + '₹' + f'{abs(v) / 1e5:,.2f} L'
pc = lambda v, d=1: ('−' if v < 0 else '') + f'{abs(v):.{d}f}%'
spc = lambda v, d=2: ('+' if v >= 0 else '−') + f'{abs(v):.{d}f}%'
n2 = lambda v, d=2: '—' if v is None or (isinstance(v, float) and np.isnan(v)) else f'{v:.{d}f}'
dfm = lambda ts: pd.Timestamp(ts).strftime('%d-%b-%Y')
dsh = lambda ts: pd.Timestamp(ts).strftime('%d-%b')

def table(head, rows, cls='', hl=None, colw=None, small=False):
    hl = hl or []
    th = ''.join(f'<th>{h}</th>' for h in head)
    body = ''
    for r in rows:
        body += '<tr>' + ''.join(f'<td class="{"hl" if i in hl else ""}">{c}</td>' for i, c in enumerate(r)) + '</tr>'
    return f'<table class="lg {cls}"><tr>{th}</tr>{body}</table>'

# ---------------------------------------------------------------- SVG charts
def svg_lines(series, w=700, h=250, ylabel='', title='', xlabels=None, hline0=True, ylim=None):
    pl, pr, pt, pb = 52, 12, 24, 28
    xs_all = [x for s in series for x in s['x']]; ys_all = [y for s in series for y in s['y']]
    x0, x1 = min(xs_all), max(xs_all); y0, y1 = ylim or (min(ys_all + [0]), max(ys_all + [0]))
    pad = (y1 - y0) * 0.06 or 1; y0 -= pad if not ylim else 0; y1 += pad if not ylim else 0
    X = lambda x: pl + (x - x0) / (x1 - x0 or 1) * (w - pl - pr); Y = lambda y: pt + (y1 - y) / (y1 - y0 or 1) * (h - pt - pb)
    o = f'<svg viewBox="0 0 {w} {h}" width="100%" xmlns="http://www.w3.org/2000/svg" font-family="Arial" font-size="10">'
    if title: o += f'<text x="{w/2}" y="13" text-anchor="middle" font-size="11" fill="#333">{title}</text>'
    step = max(1, int(round((y1 - y0) / 5 / (10 ** np.floor(np.log10(max((y1 - y0) / 5, 1e-9))))))) * 10 ** np.floor(np.log10(max((y1 - y0) / 5, 1e-9)))
    t = np.ceil(y0 / step) * step
    while t <= y1:
        o += f'<line x1="{pl}" x2="{w-pr}" y1="{Y(t):.1f}" y2="{Y(t):.1f}" stroke="#e6e9ef"/><text x="{pl-6}" y="{Y(t)+3:.1f}" text-anchor="end" fill="#555">{t:g}</text>'
        t += step
    if hline0 and y0 < 0 < y1: o += f'<line x1="{pl}" x2="{w-pr}" y1="{Y(0):.1f}" y2="{Y(0):.1f}" stroke="#999"/>'
    for xv, lab in (xlabels or []):
        o += f'<line x1="{X(xv):.1f}" x2="{X(xv):.1f}" y1="{pt}" y2="{h-pb}" stroke="#f0f2f6"/><text x="{X(xv):.1f}" y="{h-pb+13}" text-anchor="middle" fill="#555">{lab}</text>'
    for s in series:
        pts = ' '.join(f'{X(a):.1f},{Y(b):.1f}' for a, b in zip(s['x'], s['y']))
        o += f'<polyline fill="none" stroke="{s["c"]}" stroke-width="{s.get("w",1.8)}" {"stroke-dasharray=%s" % chr(34)+s["dash"]+chr(34) if s.get("dash") else ""} points="{pts}"/>'
    lx = pl + 8
    for s in series:
        o += f'<line x1="{lx}" x2="{lx+16}" y1="{pt+8}" y2="{pt+8}" stroke="{s["c"]}" stroke-width="2.4"/><text x="{lx+20}" y="{pt+11}" fill="#333">{s["n"]}</text>'; lx += 26 + 6.2 * len(s['n'])
    o += f'<text x="12" y="{(pt+h-pb)/2}" transform="rotate(-90 12 {(pt+h-pb)/2})" text-anchor="middle" fill="#555">{ylabel}</text></svg>'
    return o

def svg_bars(vals, labels, w=700, h=200, title='', ylabel=''):
    pl, pr, pt, pb = 52, 12, 22, 26
    y1 = max(max(vals), 0) * 1.1 or 1; y0 = min(min(vals), 0) * 1.15
    X = lambda i: pl + (i + 0.5) / len(vals) * (w - pl - pr); Y = lambda y: pt + (y1 - y) / (y1 - y0) * (h - pt - pb)
    bw = (w - pl - pr) / len(vals) * 0.7
    o = f'<svg viewBox="0 0 {w} {h}" width="100%" xmlns="http://www.w3.org/2000/svg" font-family="Arial" font-size="10"><text x="{w/2}" y="13" text-anchor="middle" font-size="11" fill="#333">{title}</text>'
    for t in np.arange(np.ceil(y0 / 10) * 10, y1 + 1e-9, 10):
        o += f'<line x1="{pl}" x2="{w-pr}" y1="{Y(t):.1f}" y2="{Y(t):.1f}" stroke="#eef0f4"/><text x="{pl-6}" y="{Y(t)+3:.1f}" text-anchor="end" fill="#555">{t:.0f}</text>'
    o += f'<line x1="{pl}" x2="{w-pr}" y1="{Y(0):.1f}" y2="{Y(0):.1f}" stroke="#888"/>'
    for i, v in enumerate(vals):
        c = '#2ca02c' if v >= 0 else '#c0392b'
        o += f'<rect x="{X(i)-bw/2:.1f}" y="{min(Y(v),Y(0)):.1f}" width="{bw:.1f}" height="{abs(Y(v)-Y(0)):.1f}" fill="{c}"/>'
        if i % 2 == 0 or len(vals) < 14: o += f'<text x="{X(i):.1f}" y="{h-pb+12}" text-anchor="middle" fill="#555" font-size="9">{labels[i]}</text>'
    o += f'<text x="12" y="{(pt+h-pb)/2}" transform="rotate(-90 12 {(pt+h-pb)/2})" text-anchor="middle" fill="#555">{ylabel}</text></svg>'
    return o

def color_scale(v, lo, mid, hi):
    def mix(a, b, t): return tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))
    red, wht, grn = (215, 48, 39), (255, 255, 255), (26, 152, 80)
    c = mix(red, wht, (v - lo) / (mid - lo)) if v < mid else mix(wht, grn, min(1, (v - mid) / (hi - mid)))
    return '#%02x%02x%02x' % c if v >= lo else '#d73027'

def svg_heat(grid, lo, mid, hi, fmt, title, w=330, h=300):
    NN = len(grid); pl, pt, cw, ch = 34, 26, (w - 40) / NN, (h - 50) / NN
    o = f'<svg viewBox="0 0 {w} {h}" width="100%" xmlns="http://www.w3.org/2000/svg" font-family="Arial" font-size="9"><text x="{w/2}" y="13" text-anchor="middle" font-size="10.5" fill="#333">{title}</text>'
    for i in range(NN):
        o += f'<text x="{pl-5}" y="{pt+i*ch+ch/2+3:.1f}" text-anchor="end" fill="#555">T-{i+1}</text><text x="{pl+i*cw+cw/2:.1f}" y="{pt+NN*ch+12}" text-anchor="middle" fill="#555">T+{i+1}</text>'
        for j in range(NN):
            v = grid[i][j]
            o += f'<rect x="{pl+j*cw:.1f}" y="{pt+i*ch:.1f}" width="{cw:.1f}" height="{ch:.1f}" fill="{color_scale(v,lo,mid,hi)}" stroke="#fff"/><text x="{pl+j*cw+cw/2:.1f}" y="{pt+i*ch+ch/2+3:.1f}" text-anchor="middle" fill="#222" font-size="8">{fmt(v)}</text>'
    return o + '</svg>'

# ---------------------------------------------------------------- derived tables
A, B, FIX, WF35, WF50, T35, T50, TR35, TF35, IDX = [M[k] for k in ['A', 'B', 'FIX', 'WF35', 'WF50', 'TEST35', 'TEST50', 'TRAIN35', 'TESTFIX35', 'IDX']]
AN, BN = MN['A'], MN['B']
nA, nB = len(R['SYM35']), len(R['SYM9'])
strong_names = ', '.join(R['SYM9'])
cyc = A['cyc']

def episodes(c, margin, k=5):
    dd = c.dd.values; cum = c.cum.values; yrs = list(c.index); out = []; i = 0; n = len(dd)
    while i < n:
        if dd[i] < 0:
            j = i
            while j < n and dd[j] < 0: j += 1
            seg = dd[i:j]; ti = i + int(np.argmin(seg)); pk = yrs[i - 1] if i > 0 else 'start'
            out.append((pk, yrs[ti], yrs[j] if j < n else 'ongoing', seg.min(), seg.min() / margin * 100, ti - i + 1, (j - i + 1) if j < n else 'ongoing')); i = j
        else: i += 1
    return sorted(out, key=lambda e: e[3])[:k]
EPS = episodes(cyc, MG['A'])

# ledger helpers (raw close prices)
px = pd.read_parquet(ROOT + 'Stock_Data_213/all_stocks_daily_ohlc.parquet', columns=['Symbol', 'Date', 'Close']).dropna()
ns = {}; head = open(ROOT + 'build_diwali_christmas_perstock_reports_v4.py', encoding='utf-8').read().split('# ---------------------------------------------------------------- per-stock analysis')[0].replace('json.dump(EVENT_WINDOW', '(lambda *a, **k: None)(EVENT_WINDOW').replace(r'open(ROOT + r"\Diwali_Christmas_event_windows_v4.json", "w")', 'open(__import__("os").devnull, "w")')  # DEVNULLFIX: never truncate the tracked windows file
exec(compile(head, 'v4', 'exec'), ns); data = ns['data']['Diwali']
rawc = {s: g.drop_duplicates('Date').set_index('Date').Close for s, g in px.groupby('Symbol')}
def ledger(year):
    rows = []; yi = YEARS.index(year)
    for p in POS:
        D = data[p['symbol']]; ip, iq = D['IP'][yi], D['IQ'][yi]
        if ip < 0: continue
        e, x = ip - (p['lead'] - 1), iq + (p['hold'] - 1)
        if e < 0 or x >= len(D['d']): continue
        ed, xd = pd.Timestamp(D['d'][e]), pd.Timestamp(D['d'][x]); r = D['R'][yi, p['lead'] - 1, p['hold'] - 1]
        if np.isnan(r): continue
        rs = rawc[p['symbol']]; pe_, px_ = float(rs.get(ed, np.nan)), float(rs.get(xd, np.nan))
        rows.append(dict(symbol=p['symbol'], window=f"T-{p['lead']} to T+{p['hold']}", entry=ed, entry_px=pe_, exit=xd, exit_px=px_, ret=float(r), pnl=float(r) / 100 * NOTIONAL[p['symbol']]))
    return pd.DataFrame(rows).sort_values('pnl', ascending=False)
L25, L24 = ledger(2025), ledger(2024)
def ledger_html(L):
    return ''.join(f"<tr><td>{r.symbol}</td><td class=g><b>LONG</b></td><td>{r.window}</td><td>{dsh(r.entry)}</td><td class=R>{n2(r.entry_px,1)}</td><td>{dsh(r.exit)}</td><td class=R>{n2(r.exit_px,1)}</td><td class='{'g' if r.ret>=0 else 'r'}'><b>{spc(r.ret)}</b></td><td class='{'g' if r.ret>=0 else 'r'}'><b>{'WIN' if r.ret>0 else 'LOSS'}</b></td><td class='R {'g' if r.pnl>=0 else 'r'}'>{sinr(r.pnl)}</td></tr>" for r in L.itertuples())

# worked example (TITAN 2025)
def worked(sym, year):
    p = [q for q in POS if q['symbol'] == sym][0]; yi = YEARS.index(year); D = data[sym]; ip, iq = D['IP'][yi], D['IQ'][yi]
    e, x = ip - (p['lead'] - 1), iq + (p['hold'] - 1); rs = rawc[sym]; rows = []; base = float(rs[pd.Timestamp(D['d'][e])])
    idx = list(range(e, ip + 1)) + [None] + list(range(iq, x + 1))
    for k in idx:
        if k is None:
            rows.append(f"<tr style='background:#fff7e0'><td>{dfm(TD[year])}</td><td>Laxmi Pujan / Muhurat (T)</td><td colspan=3>holiday session, not counted as a trading day</td></tr>"); continue
        d = pd.Timestamp(D['d'][k]); c = float(rs[d]); lab = f"T-{ip-k+1}" if k <= ip else f"T+{k-iq+1}"
        act = 'BUY at close (entry)' if k == e else ('SELL at close (exit)' if k == x else '')
        rows.append(f"<tr><td>{dfm(d)}</td><td>{lab}</td><td class=R>{c:,.1f}</td><td class='{'g' if c>=base else 'r'}'>{spc((c/base-1)*100)}</td><td><b>{act}</b></td></tr>")
    return p, '<table class=lg><tr><th>Date</th><th>Day</th><th>Close ₹</th><th>vs entry</th><th>Action</th></tr>' + ''.join(rows) + '</table>'
wp, WORK = worked(EX_SYM, 2025)

# year-by-year (Final A)
yr_rows = ''
for y in YEARS:
    a = cyc.loc[y]; ix = BOOKS['IDX'].set_index('year').ret.get(y, np.nan); f_ = yt.set_index('year').fix_ret[y]
    yr_rows += f"<tr><td>{y}</td><td>{TD[y]:%d-%b-%Y}</td><td>{int(a.n)}</td><td>{int(a.w)}</td><td>{a.w/a.n*100:.0f}%</td><td class={'g' if a.pnl>=0 else 'r'}>{sinr(a.pnl)}</td><td class={'g' if a.pnl>=0 else 'r'}>{spc(a.pnl/MG['A']*100)}</td><td>{spc(a.ret)}</td><td>{spc(f_)}</td><td>{'—' if np.isnan(ix) else spc(ix)}</td></tr>"

# timing / holding buckets, sectors, october vs november
dA = BOOKS['A'].copy()
def bucket(df, key, bins):
    rows = []
    for lab, f in bins:
        s = df[f(df[key])]
        if len(s): rows.append([lab, len(s), f"{(s.ret>0).mean()*100:.1f}%", spc(s.ret.mean()), sinr(s.pnl.sum()), f"{s.groupby('year').pnl.sum().gt(0).mean()*100:.0f}%"])
    return rows
t_rows = bucket(dA, 'lead', [('Early entry (T-6 to T-8)', lambda x: x >= 6), ('Mid entry (T-3 to T-5)', lambda x: (x >= 3) & (x <= 5)), ('Late entry (T-1 to T-2)', lambda x: x <= 2)])
h_rows = bucket(dA, 'hold', [('Short hold (T+1 to T+3)', lambda x: x <= 3), ('Medium hold (T+4 to T+6)', lambda x: (x >= 4) & (x <= 6)), ('Long hold (T+7 to T+8)', lambda x: x >= 7)])
dA['Tm'] = dA.year.map(lambda y: TD[y].month)
m_rows = []
for lab, mm in (('Diwali in October', 10), ('Diwali in November', 11)):
    s = dA[dA.Tm == mm]; yrs_ = sorted(s.year.unique()); fx = yt[yt.year.isin(yrs_)]
    m_rows.append([lab, len(yrs_), f"{(s.ret>0).mean()*100:.1f}%", spc(s.ret.mean()), pc(s.groupby('year').pnl.sum().gt(0).mean()*100,0), spc(fx.fix_ret.mean())])
sec_rows = []
for sec, syms in sorted(R['sector_tbl'].items(), key=lambda kv: -len(kv[1])):
    s = dA[dA.symbol.isin(syms)]; pp = [p for p in POS if p['symbol'] in syms]
    sec_rows.append([sec or 'Other', len(syms), ', '.join(syms[:5]) + ('…' if len(syms) > 5 else ''), f"{(s.ret>0).mean()*100:.1f}%", spc(s.ret.mean()), f"{np.nanmean([p['wf'] for p in pp if p['wf'] is not None]):.1f}%"])
# worst basket years
wy = yt.sort_values('fix_ret').head(6)
wy_rows = [[int(r.year), f"{r['T']:%d-%b-%Y}", spc(r.fix_ret), f"{r.fix_wr:.0f}%", spc(cyc.loc[int(r.year)].ret), sinr(cyc.loc[int(r.year)].pnl)] for _, r in wy.iterrows()]

# conviction table
def conv_row(c):
    pp = [p for p in POS if p['conv'] == c]; s = dA[dA.symbol.isin([p['symbol'] for p in pp])]
    return [c, len(pp), f"{np.mean([p['wr'] for p in pp]):.1f}%", f"{np.nanmean([p['wf'] for p in pp if p['wf'] is not None]):.1f}%", f"{(s.ret>0).mean()*100:.1f}%", spc(s.ret.mean())]
conv_rows = [conv_row('STRONG'), conv_row('MODERATE')]

# ---------------------------------------------------------------- charts
cumx = lambda m, k: dict(x=list(m['cyc'].index), y=list(m['cyc'].cum.values / MG[k] * 100))
ch_cum = svg_lines([dict(n='Final A (35)', c='#1f3a5f', **cumx(A, 'A')), dict(n='Final B (9 STRONG)', c='#c0392b', **cumx(B, 'B')), dict(n='Fixed T-3/T+3 (50)', c='#2ca02c', **cumx(FIX, 'FIX')),
                    dict(n='Walk-forward (35)', c='#ff7f0e', dash='5,3', **cumx(WF35, 'WF35'))], ylabel='Cumulative return (% of margin)', title='Cumulative return on own margin: one Diwali cycle per year',
                   xlabels=[(y, str(y)) for y in range(2000, 2026, 5)] + [(2025, '2025')], h=270)
ch_dd = svg_lines([dict(n='Final A', c='#1f3a5f', x=list(A['cyc'].index), y=list(A['cyc'].dd.values / MG['A'] * 100)), dict(n='Final B', c='#c0392b', x=list(B['cyc'].index), y=list(B['cyc'].dd.values / MG['B'] * 100)),
                   dict(n='Walk-forward (35)', c='#ff7f0e', x=list(WF35['cyc'].index), y=list(WF35['cyc'].dd.values / MG['WF35'] * 100))], ylabel='Drawdown (% of margin)', h=170, xlabels=[(y, str(y)) for y in range(2000, 2026, 5)])
ch_bar = svg_bars(list(cyc.pnl.values / MG['A'] * 100), [str(y)[2:] for y in cyc.index], title='Final A: return per Diwali cycle (% of margin)', ylabel='% of margin')
offs = list(range(-8, 0)) + list(range(1, 9)); xs = list(range(16)); allp = np.array([v for v in R['paths'].values()])
ch_path = svg_lines([dict(n='Average of all years', c='#1f3a5f', x=xs, y=list(allp.mean(0) - allp.mean(0)[0] * 0)), dict(n='2024', c='#c0392b', x=xs, y=list(R['paths'][2024])), dict(n='2025', c='#2ca02c', x=xs, y=list(R['paths'][2025]))],
                    ylabel='% from T-8 close', title=f'Average path of the {len(POS)} positions around Diwali (T-8 to T+8, indexed to T-8 close)', h=250,
                    xlabels=[(0, 'T-8'), (3, 'T-5'), (7, 'T-1'), (8, 'T+1'), (11, 'T+4'), (15, 'T+8')])
hm1 = svg_heat([[R['HM_ret'][i][j] for j in range(8)] for i in range(8)], -0.5, 0.8, 3.2, lambda v: f'{v:.1f}', 'Avg LONG return % (Nifty 50 basket)')
hm2 = svg_heat([[R['HM_wr'][i][j] for j in range(8)] for i in range(8)], 40, 60, 90, lambda v: f'{v:.0f}', 'Basket win rate % (years the basket was up)')

# metrics table (headline)
def mrow(label, key, f, cols): return [label] + [f(c) for c in cols]
cols = [IDX, FIX, WF50, A, B]; keys = ['IDX', 'FIX', 'WF50', 'A', 'B']
def tab_metrics(cs, ks, hlcols):
    L = [('Capital (margin)', lambda c, k: inr(MG[k])), ('Cycles (years)', lambda c, k: str(c['cycles'])), ('Trades', lambda c, k: f"{c['trades']:,}"),
         ('Avg annual return (simple)', lambda c, k: pc(c['annual'])), ('CAGR', lambda c, k: pc(c['cagr'])), ('Max drawdown (% of capital)', lambda c, k: pc(c['maxdd_pct'])),
         ('Max drawdown (₹)', lambda c, k: inr(c['maxdd'])), ('Calmar (CAGR ÷ DD %)', lambda c, k: n2(c['calmar'])), ('Sharpe / Sortino (per cycle)', lambda c, k: f"{n2(c['sharpe'])} / {n2(c['sortino'])}"),
         ('Recovery factor (net ÷ max DD)', lambda c, k: n2(c['recovery'], 1)), ('Longest drawdown (cycles)', lambda c, k: str(c['longest_dd'])), ('Win trades', lambda c, k: pc(c['wr'])),
         ('Positive cycles', lambda c, k: pc(c['pos_cycles'], 0)), ('Profit factor (trades)', lambda c, k: n2(c['pf'])), ('Avg win / avg loss (return)', lambda c, k: f"{spc(c['avg_win'])} / {spc(c['avg_loss'])}"),
         ('Avg trade return', lambda c, k: spc(c['avg_ret'])), ('Payoff (avg win ÷ avg loss)', lambda c, k: n2(c['payoff'])), ('Best cycle / worst cycle', lambda c, k: f"{lac(c['best_cycle'])} / {lac(c['worst_cycle'])}"),
         ('Best / worst single trade', lambda c, k: f"{inr(c['best_trade'])} / {inr(c['worst_trade'])}"), ('Max losing cycles in a row', lambda c, k: str(c['max_lose_cycles'])),
         ('Avg trades per cycle', lambda c, k: n2(c['per_cycle'], 1)), ('Net P&L (gross)', lambda c, k: lac(c['net'])), ('Net P&L after 0.15% costs', lambda c, k: lac(MN[k]['net'])), ('Total cost paid', lambda c, k: inr(MN[k]['cost']))]
    return [[lab] + [f(c, k) for c, k in zip(cs, ks)] for lab, f in L]

# ---------------------------------------------------------------- positions 2026
T26 = pd.Timestamp('2026-11-08'); closed = set(ns['hol'][ns['hol'].Is_Trading_Holiday == 'YES'].Holiday_Date); closed.add(pd.Timestamp('2026-11-10'))
cal26 = [d for d in pd.bdate_range(T26 - pd.Timedelta(days=40), T26 + pd.Timedelta(days=40)) if d not in closed]
pre26 = [d for d in cal26 if d < T26]; post26 = [d for d in cal26 if d > T26]
plan = sorted(POS, key=lambda p: (pre26[-p['lead']], p['symbol']))
ENT = [pre26[-p['lead']] for p in POS]; EXT = [post26[p['hold'] - 1] for p in POS]
ENT0, ENT1, EXT0, EXT1 = min(ENT), max(ENT), min(EXT), max(EXT)
plan_rows = [[p['symbol'], 'LONG', f"T-{p['lead']} to T+{p['hold']}", pre26[-p['lead']].strftime('%d-%b (%a)'), post26[p['hold'] - 1].strftime('%d-%b (%a)'), f"{p['wr']:.1f}%", '—' if p['wf'] is None else f"{p['wf']:.1f}%",
              p['conv'] + (' · re-picked' if p['repick'] else ''), f"{LOTS[p['symbol']]:,.0f}", inr(0.2 * NOTIONAL[p['symbol']])] for p in plan]
top_rows = [[i, p['symbol'], p['name'], f"T-{p['lead']} to T+{p['hold']}", f"{p['wr']:.1f}%", spc(p['avg']), f"{p['wf']:.1f}% ({p['wf_n']})", p['conv']] for i, p in enumerate(sorted([q for q in POS if q['conv'] == 'STRONG'], key=lambda q: -q['wr']), 1)]
rp_rows = [[r['symbol'], f"{r['original']['window']} · {r['original']['entry_date'][:6]}", f"{r['repick']['window']} · {r['repick']['entry_date'][:6]}", f"{r['original']['win_rate']}% → {r['repick']['win_rate']}%",
            f"{r['original']['conviction']} → {r['repick']['conviction']}"] for r in repick.values() if r.get('repick')]
exp_rows = [[e['symbol'], e['direction'], e['window'], e['entry_date'][:11], e['exit_date'][:11], e['valid_years'], f"{e['win_rate']:.0f}%", spc(e['avg_signed_ret'])] for e in exp3]

# ---------------------------------------------------------------- narrative numbers
oos = dict(wf=WF35, t35=T35, t50=T50, tr=TR35, tf=TF35)
A_ann, A_dd = A['annual'], A['maxdd_pct']; best_c = A['best_cycle_y']; worst_c = A['worst_cycle_y']
fix_up = int((yt.fix_ret > 0).sum())
lng = 'LONG'
CSS = """
@page{size:A4;margin:16mm 14mm 16mm 14mm}
body{font-family:Arial,Helvetica,sans-serif;color:#111;font-size:9.6pt;margin:0;line-height:1.38}
h1{font-size:21pt;color:#1f3a5f;margin:0 0 6px;line-height:1.15} .cover h1{font-size:25pt;margin-top:120px}
h2{font-size:15.5pt;color:#1f3a5f;margin:0 0 7px;page-break-after:avoid} h3{font-size:11.6pt;color:#2b6cb0;margin:13px 0 5px;page-break-after:avoid}
.sec{page-break-before:always} .cover{page-break-after:always}
.cs{font-size:12.4pt;color:#444;margin:6px 0 30px;line-height:1.35}
table{border-collapse:collapse;width:100%;margin:5px 0 3px} td,th{border:1px solid #c9d3e0;padding:3.6px 6px;vertical-align:top;font-size:8.7pt;text-align:left}
th{background:#1f3a5f;color:#fff;font-weight:700} table.lg tr:nth-child(even) td{background:#eef3f8} table.cv td{font-size:9.2pt;padding:5px 8px} table.cv tr:nth-child(even) td{background:#eef3f8}
td.hl{background:#e3f1e4 !important} .dense td,.dense th{padding:2.4px 5px;font-size:8pt} .R{text-align:right} .g{color:#146c2e} .r{color:#b71c1c}
.note{font-size:8pt;color:#555;font-style:italic;margin:3px 0 8px} .box{border:1px solid #2b6cb0;background:#e8f4ec;padding:8px 11px;margin:9px 0;font-size:9.4pt}
.warn{border:1px solid #d9a441;background:#fff7e0;padding:8px 11px;margin:9px 0;font-size:9.2pt} ul{margin:4px 0 6px 18px;padding:0} li{margin:3px 0}
.two{display:flex;gap:10px} .two>div{flex:1} svg{display:block}
"""
cover_rows = [('Builds on', 'The per-stock Diwali scan (build_diwali_christmas_perstock_reports_v4) re-run on the project 8×8 grid; supersedes the earlier 10×10 Diwali reports'),
              ('Data', f"Daily prices of {R['n_stocks']} NSE F&O stocks (Nifty 50 = 50), 26 Diwali events 2000–2025, Laxmi Pujan / Muhurat dates; Nifty futures history for the index check"),
              ('Independent check', f"Per-year walk-forward, a 2000–2018 train / 2019–2025 test split, and the Nifty futures index (T-5 to T+4, 26 events) — none used to pick the stock windows"),
              ('Final strategy', f"{nA} LONG futures positions on the project 8x8 grid, each at its own window (T-N to T+M, closes), entries {ENT0:%d-%b-%Y} to {ENT1:%d-%b-%Y}, exits {EXT0:%d-%b} to {EXT1:%d-%b-%Y}; Final B = the {nB} STRONG names only"),
              ('Headline', f"In-sample {pc(A_ann)} a year on margin, max drawdown {pc(A_dd)}, {pc(A['pos_cycles'],0)} positive cycles. Out-of-sample: {pc(T35['annual'])} / {pc(T35['maxdd_pct'])} (train/test split) and {pc(WF35['annual'])} / {pc(WF35['maxdd_pct'])} (walk-forward)"),
              ('Costs', f"0.15% of contract value per round trip, applied in the ‘net’ rows; lot sizes are today's lots, values at today's prices"),
              ('Prepared for', 'Management review'), ('Date', '27 October 2026')]

HMr = np.array(R['HM_ret']); HMw = np.array(R['HM_wr'])
GRIDTXT = (f"On the 8×8 grid the basket's average return is positive in {int((HMr>0).sum())} of 64 windows and the up-year share is at least 60% in {int((HMw>=60).sum())} of 64. "
           f"Entering 5–8 days before and leaving 3–8 days after averages {HMr[4:8,2:8].mean():+.2f}%, against {HMr[0:2,:].mean():+.2f}% for entering only 1–2 days before; the profitable zone is wide, so the result does not hinge on one exact window.")
ENTRYTXT = f"early entries (T-6 to T-8) average {dA[dA.lead>=6].ret.mean():+.2f}% per trade against {dA[dA.lead<=2].ret.mean():+.2f}% for late entries (T-1 to T-2); see table 4.3."
NT_ROWS = [[r['symbol'], r['analysis']['direction'], f"T-{r['analysis']['lead']} to T+{r['analysis']['hold']}", f"{r['analysis']['wr']:.1f}%", '—' if r['analysis']['wf'] is None else f"{r['analysis']['wf']:.1f}%", 'walk-forward / conviction below the bar'] for r in ALL50 if r['status'] == 'NO TRADE']
html = f"""<!doctype html><html><head><meta charset=utf-8><title>Diwali final strategy report (8x8 grid)</title><style>{CSS}</style></head><body>

<div class=cover>
<h1>Nifty 50 Diwali Drift:<br>From a 26-Year Scan to One Final Position Book</h1>
<div class=cs>Per-stock best side and window around Laxmi Pujan, what drives the result, how much of it survives out-of-sample, and the final 2026 position book with full performance</div>
<table class=cv><tr><td style="background:#dfe8f3" colspan=2>&nbsp;</td></tr>{''.join(f'<tr><td style="width:22%">{a}</td><td>{b}</td></tr>' for a, b in cover_rows)}</table>
<p class=note style="margin-top:18px">All figures are from a historical backtest on daily data, with entries and exits at the close of the stated trading day. Past results do not guarantee future returns. Windows were chosen on the same history they are scored on; read the out-of-sample rows, not only the in-sample ones.</p>
</div>

<div class=sec>
<h2>1. Executive summary</h2>
<p>This document follows the Diwali work end to end: a scan of every F&amp;O stock over 128 side/window combinations (the project 8×8 grid), a look at <b>what drives</b> the result, an honest test of how much of it survives out-of-sample, and the final position book for Diwali 2026 (Laxmi Pujan, Sunday 8 November).</p>
{table(['Step','What was done','What came out'],[
 ['1. Scan', f"{R['n_stocks']} stocks × 8 entry leads × 8 holds × 2 sides (project 8×8 grid), 26 Diwali events (2000–2025), per-stock best by win rate", f"Nifty 50: {LS['NIFTY']['long']} of {LS['NIFTY']['tested']} stocks have LONG as the better side; avg best LONG win rate {LS['NIFTY']['avg_long']:.1f}% vs best SHORT {LS['NIFTY']['avg_short']:.1f}%"],
 ['2. What drives it', 'Basket by year, entry/hold windows, Oct vs Nov Diwali, sectors, worst years', f"Equal-weight basket up in {fix_up} of 26 years; the gain builds from about T-8 and runs to T+3..T+8; common-window down years: {', '.join(str(int(y)) for y in yt[yt.fix_ret<0].year)}; Final A was roughly flat in 2024 ({sinr(cyc.loc[2024].pnl)})"],
 ['3. Choose and check', 'Conviction rules, walk-forward, 2000–18 / 2019–25 split, Nifty futures index', f"{nA} tradeable LONG positions ({nB} STRONG). In-sample {pc(A['wr'])} win rate; walk-forward {pc(WF35['wr'])}; train/test split {pc(T35['wr'])}"],
 ['4. Final book', '8×8 window check (no entry date has passed), 2026 entry/exit calendar', f"No re-pick needed (the earliest 8×8 entry is T-8 = 28-Oct); entries {ENT0:%d-%b} to {ENT1:%d-%b}, exits {EXT0:%d-%b} to {EXT1:%d-%b}"]], hl=[])}
<h3>Headline results</h3>
{table(['Metric','Index futures T-5/T+4','Fixed T-3/T+3, all 50','Walk-forward, all 50 (pure OOS)','Final A (35, in-sample)','Final B (9 STRONG)'], tab_metrics(cols, keys, [4])[:19], hl=[4], small=True)}
<p class=note>Each book is one portfolio on its own margin (20% of contract value, today's prices and lots; the index column is one Nifty futures lot on ₹1.25 L). One Diwali cycle a year, so ‘annual return’ is the average cycle return. Final A and Final B windows were chosen on the full 2000–2025 history, so their rows are in-sample; the walk-forward column is not.</p>
<div class=box><b>Recommendation: use Final A as the position book, but plan on the out-of-sample numbers, not the in-sample ones.</b> In-sample Final A shows {pc(A_ann)} a year with a max drawdown of {pc(A_dd)}. Re-choosing every window on 2000–2018 only and scoring 2019–2025 gives {pc(T35['annual'])} a year and {pc(T35['maxdd_pct'])} drawdown on the same 35 stocks (win rate {pc(T35['wr'])}); a strict year-by-year walk-forward is weaker still ({pc(WF35['annual'])}, {pc(WF35['wr'])} win rate). Final B ({nB} STRONG names) looks best in-sample ({pc(B['annual'])} / {pc(B['maxdd_pct'])}) but the STRONG label itself uses a walk-forward filter, so part of that edge is selection, not proof.</div>
</div>

<div class=sec>
<h2>2. Data and how the backtest works</h2>
{table(['Item','Setting'],[
 ['Prices','Daily Adj Close (returns, so splits and dividends do not distort) from Stock_Data_213; raw Close shown in the example ledgers'],
 ['Event date T','Laxmi Pujan / Muhurat date, 2000–2025 (hand-corrected: the holiday CSV was off by 1–3 days in several early years); 2026 = Sunday 08-Nov-2026'],
 ['Trading calendar','A day counts only if at least 30% of listed stocks traded volume that day (removes placeholder rows on holidays); the Muhurat session itself is not counted'],
 ['Entry / exit','Entry at the close of trading day T-N; exit at the close of trading day T+M (daily data, no intraday timing). T-1 is the last session before Laxmi Pujan, T+1 the first after it'],
 ['Window grid','PROJECT 8×8 GRID: N = 1..8 trading days before, M = 1..8 days after, LONG and SHORT (128 combinations per stock)'],
 ['Selection','Highest win rate over the stock\’s valid years; ties broken by average signed return; needs at least 7 valid event years'],
 ['Guards','A window is dropped for a year if any daily move exceeds ±35%, or the window spans more than 40 calendar days, or no session falls within 7 days of T'],
 ['Position size','One futures lot per stock at today\’s lot size; contract value = lot × latest close; same ₹ weight for every year (so years are comparable)'],
 ['Margin / capital',f'20% of contract value summed over the book (Final A {lac(MG["A"])}, Final B {lac(MG["B"])}); fixed, no compounding'],
 ['Costs','0.15% of contract value per round trip, shown in the ‘net’ rows (assumption; not taken from broker data)'],
 ['Universe','The 50 current Nifty stocks with data; 3 (ETERNAL, JIOFIN, MAXHEALTH) have too little history for a backtested call']], cls='dense')}
<h3>2.1 The trade mechanics</h3>
<ul><li><b>Buy before:</b> buy the stock future at the close of the N-th trading day before Laxmi Pujan.</li>
<li><b>Hold through:</b> carry the position over the Muhurat session and the closed days.</li>
<li><b>Sell after:</b> sell at the close of the M-th trading day after it. Return = (P<sub>exit</sub> / P<sub>entry</sub> − 1) × direction; all current positions are LONG.</li>
<li>There is no stop-loss inside the backtest. The −3.5% spot stop in the playbook is a risk control suggested for live trading and is not in the numbers.</li></ul>
<h3>2.2 One trade, day by day (illustration)</h3>
<p>{wp['symbol']} in Diwali 2025, window T-{wp['lead']} to T+{wp['hold']}:</p>
{WORK}
<p class=note>Return {spc(float(ledger(2025).set_index('symbol').loc[EX_SYM].ret))} on the stock; at one lot ({LOTS[EX_SYM]:,.0f} shares, contract value {inr(NOTIONAL[EX_SYM])}) that is {sinr(float(ledger(2025).set_index('symbol').loc[EX_SYM].pnl))} before costs.</p>
</div>

<div class=sec>
<h2>3. Step 1: the per-stock scan</h2>
{table(['Rule','Values tested','Count'],[['Stocks','F&O stocks with price data',str(R['n_stocks'])],['Entry lead N','1 to 8 trading days before T','8'],['Hold M','1 to 8 trading days after T','8'],['Side','LONG / SHORT','2'],['Stock-windows scored',f"{R['n_stocks']} × 8 × 8 × 2",f"{R['n_stocks']*128:,}"],['Years per window','2000 to 2025, minimum 7 valid','≤ 26'],['Out-of-sample check','Pick redone each year from earlier years only (walk-forward)','19 picks / stock']], cls='')}
<h3>3.1 What the scan produced for the Nifty 50</h3>
{table(['Conviction','Stocks','Avg in-sample win rate','Avg walk-forward win rate','Final-window win rate (26 yrs)','Avg trade return'], conv_rows)}
<p class=note>STRONG = at least 10 years, win rate ≥ 75%, positive average, walk-forward win rate ≥ 60%. MODERATE = win rate ≥ 65%, positive average, walk-forward ≥ 50% (or too few walk-forward trades). The other {sum(r['status']=='NO TRADE' for r in ALL50)} Nifty stocks are NO TRADE and {sum(r['status']=='DATA INSUFFICIENT' for r in ALL50)} are data-insufficient; they are not in the book.</p>
<h3>3.2 The simplest version: one window for everyone (T-3 to T+3), year by year</h3>
{table(['Year','Laxmi Pujan','Nifty-50 basket return','Stocks up','Year','Laxmi Pujan','Nifty-50 basket return','Stocks up'],
   [[int(a.year), f"{a['T']:%d-%b-%Y}", spc(a.fix_ret), f"{a.fix_wr:.0f}%", int(b.year), f"{b['T']:%d-%b-%Y}", spc(b.fix_ret), f"{b.fix_wr:.0f}%"] for (_, a), (_, b) in zip(yt.iloc[:13].iterrows(), yt.iloc[13:].iterrows())], cls='dense')}
<p class=note>Equal-weight average of the Nifty 50 stocks that had data that year, LONG, entry at the close of T-3, exit at the close of T+3.</p>
<h3>3.3 What this step showed</h3>
<ul><li>The basket was up in <b>{fix_up} of 26 years</b> with one common window. The down years were {', '.join(str(int(y)) for y in yt[yt.fix_ret<0].year)}.</li>
<li>Individual stocks differ a lot: a stock’s own best window wins far more often in-sample than the common window ({pc(A['wr'])} vs {pc(FIX['wr'])} of trades) — but part of that is the search itself, which is why step 3 exists.</li>
<li>All trades in one year share the same market backdrop, so the {A['trades']} trades of Final A are really <b>26 observations</b> of the market, not {A['trades']}.</li></ul>
</div>

<div class=sec>
<h2>4. Step 2: what drives the result</h2>
<h3>4.1 LONG versus SHORT</h3>
<p>For each stock the best LONG window was compared with its best SHORT window on the 8×8 grid (≥ 7 valid years). <b>Nifty 50: LONG is better for {LS['NIFTY']['long']} of {LS['NIFTY']['tested']} stocks</b>, SHORT for {LS['NIFTY']['short']}, with {LS['NIFTY']['tie']} exact tie (ties default to LONG in the scan). Across the {LS['FO']['tested']} tested non-Nifty F&amp;O stocks LONG is better for {LS['FO']['long']}, SHORT for {LS['FO']['short']}, and {LS['FO']['tie']} are ties (see Diwali_Positions_8x8_Nifty_FO.xlsx). Average best-window win rate on the Nifty: LONG {LS['NIFTY']['avg_long']:.1f}% vs SHORT {LS['NIFTY']['avg_short']:.1f}%. Part of the LONG bias is general upward drift in prices; part is Diwali.</p>
<h3>4.2 Which windows work: the 8 × 8 grid for the Nifty-50 basket (all years)</h3>
<div class=two><div>{hm1}</div><div>{hm2}</div></div>
<p class=note>Rows = entry N days before T, columns = exit M days after T. Left: average LONG return of the equal-weight basket. Right: share of the 26 years in which the basket was up.</p>
<p>{GRIDTXT}</p>
<h3>4.3 Entry timing and holding length (Final A trades, 2000–2025)</h3>
{table(['Entry timing','Trades','Win rate','Avg return','Net P&L','Cycles positive'], t_rows)}
{table(['Holding length','Trades','Win rate','Avg return','Net P&L','Cycles positive'], h_rows)}
<h3>4.4 Diwali in October versus November</h3>
{table(['Diwali date','Years','Final A win rate','Final A avg return','Cycles positive','Fixed-window basket avg'], m_rows)}
<p class=note>Diwali moves between mid-October and mid-November, so the window sits in different parts of the market calendar.</p>
<h3>4.5 By sector (Final A)</h3>
{table(['Sector','Stocks','Names','Win rate','Avg return','Avg walk-forward win rate'], sec_rows, cls='dense')}
<h3>4.6 The six weakest years for the common window</h3>
{table(['Year','Laxmi Pujan','Fixed-window basket','Stocks up','Final A avg trade','Final A P&L'], wy_rows)}
<h3>4.7 Average path around Diwali</h3>
{ch_path}
<p class=note>Equal-weight average of the Final A names, indexed to the T-8 close. The gap between T-1 and T+1 is the Muhurat/holiday.</p>
<h3>4.8 Conclusions that went into the final design</h3>
<ul><li><b>Direction:</b> LONG only; no SHORT passed the evidence bar.</li><li><b>Entry:</b> {ENTRYTXT}</li><li><b>Per-stock windows</b> beat one common window in-sample, but only a part of the gain is real — hence the conviction filter.</li><li><b>One event a year:</b> size the book as a single yearly bet; a bad cycle hits every position together.</li></ul>
</div>

<div class=sec>
<h2>5. Step 3: choosing the windows and checking them</h2>
<h3>5.1 How the book was chosen (protection against curve-fitting)</h3>
<ul><li><b>Minimum history:</b> a stock needs ≥ 7 valid Diwali years in the chosen window; three Nifty stocks failed this and are left out of the main book.</li>
<li><b>Walk-forward:</b> for each year from 2007, the best window is re-picked using only earlier years and traded in that year. This is shown next to every in-sample pick.</li>
<li><b>Conviction gate:</b> only STRONG and MODERATE stocks are in the book. NO TRADE stocks (best in-sample win rate but walk-forward below the bar) are excluded.</li>
<li><b>Grid:</b> the project's 8×8 grid (N, M = 1..8) is used for every stock, so the earliest 2026 entry is T-8 = 28-Oct and no re-pick for passed entries is needed.</li>
<li><b>Honest limit:</b> 128 combinations are scanned per stock, so the best in-sample win rate is optimistic by construction.</li></ul>
<h3>5.2 The STRONG positions</h3>
{table(['#','Stock','Company','Window','Win rate (26y)','Avg return','Walk-forward (trades)','Conviction'], top_rows)}
<h3>5.3 Nifty stocks excluded by the conviction gate</h3>
{table(['Stock','Best side','Window','Win rate (26y)','Walk-forward','Why not in the book'], NT_ROWS, cls='dense')}
<p class=note>NO TRADE = the best in-sample window fails the conviction gate (win rate ≥ 65%, positive average, walk-forward ≥ 50%). Data-insufficient stocks (fewer than 7 valid Diwali years) are in section 7.2.</p>
<h3>5.4 Approaches compared on the same 26 years</h3>
{table(['Approach','Role','Win rate','Avg trade return','Cycles','Positive cycles','Avg annual','Max DD'], [
 ['Index futures T-5 to T+4','Independent index check',pc(IDX['wr']),spc(IDX['avg_ret']),IDX['cycles'],pc(IDX['pos_cycles'],0),'—','—'],
 ['Fixed T-3 to T+3, all 50','No stock selection',pc(FIX['wr']),spc(FIX['avg_ret']),FIX['cycles'],pc(FIX['pos_cycles'],0),pc(FIX['annual']),pc(FIX['maxdd_pct'])],
 ['Per-stock best, all 35 (Final A)','In-sample',pc(A['wr']),spc(A['avg_ret']),A['cycles'],pc(A['pos_cycles'],0),pc(A['annual']),pc(A['maxdd_pct'])],
 ['Same 35, picked on 2000–18, scored 2019–25','Out-of-sample',pc(T35['wr']),spc(T35['avg_ret']),T35['cycles'],pc(T35['pos_cycles'],0),pc(T35['annual']),pc(T35['maxdd_pct'])],
 ['All 50, picked on 2000–18, scored 2019–25','Out-of-sample, no stock filter',pc(T50['wr']),spc(T50['avg_ret']),T50['cycles'],pc(T50['pos_cycles'],0),pc(T50['annual']),pc(T50['maxdd_pct'])],
 ['Walk-forward each year, same 35','Out-of-sample (2007–25)',pc(WF35['wr']),spc(WF35['avg_ret']),WF35['cycles'],pc(WF35['pos_cycles'],0),pc(WF35['annual']),pc(WF35['maxdd_pct'])],
 ['Walk-forward each year, all 50','Out-of-sample, no stock filter',pc(WF50['wr']),spc(WF50['avg_ret']),WF50['cycles'],pc(WF50['pos_cycles'],0),pc(WF50['annual']),pc(WF50['maxdd_pct'])]], cls='dense')}
<p class=note>‘Same 35’ rows carry a mild look-ahead: which 35 stocks are in the book was decided with knowledge of the full history. The ‘all 50’ rows have no such filter. The 2019–25 window is only 7 cycles.</p>
</div>

<div class=sec>
<h2>6. Final Strategy A: the 35-stock Diwali book</h2>
<div class=box><b>Buy the stock future of each of the {nA} names at the close of its entry day (T-N) and sell it at the close of T+M.</b> All positions LONG, one lot each, entries from {ENT0:%d-%b-%Y} to {ENT1:%d-%b-%Y}, exits from {EXT0:%d-%b} to {EXT1:%d-%b-%Y}. Laxmi Pujan / Muhurat is Sunday 08-Nov-2026; 10-Nov (Balipratipada) is a market holiday. Planned margin ≈ {lac(MG['A'])} (20% of contract value).</div>
<h3>6.1 Daily checklist</h3>
{table(['When','Check / action','If not met'],[['Before 28-Oct','Confirm the Laxmi Pujan date (08-Nov) and the NSE holiday list (10-Nov closed)','Recompute T-N / T+M dates'],['Entry day close','Buy the future of every stock whose T-N date is today (table 6.3)','Skip that stock; do not chase next day'],['Each day','Check margin (20% of contract value) and the cumulative book P&L','—'],['Exit day close','Sell the future of every stock whose T+M date is today','Exit by the next close at the latest'],['Risk','Optional −3.5% spot stop per position (not part of the backtest)','—']], cls='')}
<h3>6.2 Why each rule is there</h3>
{table(['Rule','Reason (from the analysis above)'],[['LONG only',f"LONG beats SHORT in {LS['NIFTY']['long']} of {LS['NIFTY']['tested']} Nifty stocks; ties are not signals"],['Per-stock window','Each stock\’s own best window beat the common window in-sample; the conviction gate removes the fragile ones'],['Walk-forward gate','Stocks whose in-sample win rate collapses out-of-sample are NO TRADE'],['35 names, one lot each','Diversifies stock-specific noise; the market-wide Diwali factor is still one bet'],['8×8 grid','Entry no earlier than T-8 (28-Oct-2026), so no entry date has passed']], cls='')}
<h3>6.3 The 2026 position calendar</h3>
{table(['Stock','Side','Window','Entry (close)','Exit (close)','Win rate (26y)','Walk-forward','Conviction','Lot','Margin 20%'], plan_rows, cls='dense')}
</div>

<div class=sec>
<h2>6.4 Real example: Diwali 2025 (Laxmi Pujan 21-Oct-2025)</h2>
<p>All {len(L25)} positions at their windows, ranked by P&amp;L (return is price return on the stock; P&amp;L = return × today's one-lot contract value). Cycle result: <b>{int((L25.ret>0).sum())} wins / {int((L25.ret<=0).sum())} losses, {sinr(L25.pnl.sum())}</b>.</p>
<table class="lg dense"><tr><th>Symbol</th><th>Dir</th><th>Window</th><th>Entry</th><th>Entry ₹</th><th>Exit</th><th>Exit ₹</th><th>Return</th><th>Result</th><th>P&amp;L / lot</th></tr>{ledger_html(L25)}</table>
</div>

<div class=sec>
<h2>6.5 Real example: Diwali 2024, the bad cycle (Laxmi Pujan 01-Nov-2024)</h2>
<p>The same book in the one recent year it did not work well: <b>{int((L24.ret>0).sum())} wins / {int((L24.ret<=0).sum())} losses, {sinr(L24.pnl.sum())}</b>, an average trade of {spc(L24.ret.mean())}. A weak year is not protected: the book gave back almost all of its edge. Stock-level detail:</p>
<table class="lg dense"><tr><th>Symbol</th><th>Dir</th><th>Window</th><th>Entry</th><th>Entry ₹</th><th>Exit</th><th>Exit ₹</th><th>Return</th><th>Result</th><th>P&amp;L / lot</th></tr>{ledger_html(L24)}</table>
</div>

<div class=sec>
<h2>6.6 Performance</h2>
{table(['Metric','Final A (35)','Final A after 0.15% costs','Final B (9 STRONG)','Final B after costs'],
  [[lab, a_, an_, b_, bn_] for (lab,a_,an_,b_,bn_) in [
   ('Capital (margin)', inr(MG['A']), inr(MG['A']), inr(MG['B']), inr(MG['B'])),
   ('Cycles / trades', f"{A['cycles']} / {A['trades']}", f"{AN['cycles']} / {AN['trades']}", f"{B['cycles']} / {B['trades']}", f"{BN['cycles']} / {BN['trades']}"),
   ('Net P&L', lac(A['net']), lac(AN['net']), lac(B['net']), lac(BN['net'])),
   ('Avg annual return (simple)', pc(A['annual']), pc(AN['annual']), pc(B['annual']), pc(BN['annual'])),
   ('CAGR', pc(A['cagr']), pc(AN['cagr']), pc(B['cagr']), pc(BN['cagr'])),
   ('Max drawdown (% of capital)', pc(A['maxdd_pct']), pc(AN['maxdd_pct']), pc(B['maxdd_pct']), pc(BN['maxdd_pct'])),
   ('Max drawdown (₹)', inr(A['maxdd']), inr(AN['maxdd']), inr(B['maxdd']), inr(BN['maxdd'])),
   ('Calmar (CAGR ÷ DD %)', n2(A['calmar']), n2(AN['calmar']), n2(B['calmar']), n2(BN['calmar'])),
   ('Sharpe / Sortino (per cycle)', f"{n2(A['sharpe'])} / {n2(A['sortino'])}", f"{n2(AN['sharpe'])} / {n2(AN['sortino'])}", f"{n2(B['sharpe'])} / {n2(B['sortino'])}", f"{n2(BN['sharpe'])} / {n2(BN['sortino'])}"),
   ('Recovery factor (net ÷ max DD)', n2(A['recovery'],1), n2(AN['recovery'],1), n2(B['recovery'],1), n2(BN['recovery'],1)),
   ('Longest drawdown (cycles)', A['longest_dd'], AN['longest_dd'], B['longest_dd'], BN['longest_dd']),
   ('Win trades', pc(A['wr']), pc(AN['wr']), pc(B['wr']), pc(BN['wr'])),
   ('Positive cycles', pc(A['pos_cycles'],0), pc(AN['pos_cycles'],0), pc(B['pos_cycles'],0), pc(BN['pos_cycles'],0)),
   ('Profit factor (trades)', n2(A['pf']), n2(AN['pf']), n2(B['pf']), n2(BN['pf'])),
   ('Avg win / avg loss (return)', f"{spc(A['avg_win'])} / {spc(A['avg_loss'])}", '', f"{spc(B['avg_win'])} / {spc(B['avg_loss'])}", ''),
   ('Payoff (avg win ÷ avg loss)', n2(A['payoff']), n2(AN['payoff']), n2(B['payoff']), n2(BN['payoff'])),
   ('Best / worst cycle', f"{lac(A['best_cycle'])} ({A['best_cycle_y']}) / {lac(A['worst_cycle'])} ({A['worst_cycle_y']})", '', f"{lac(B['best_cycle'])} ({B['best_cycle_y']}) / {lac(B['worst_cycle'])} ({B['worst_cycle_y']})", ''),
   ('Max losing cycles in a row', A['max_lose_cycles'], '', B['max_lose_cycles'], ''),
   ('Total cost paid', '—', inr(AN['cost']), '—', inr(BN['cost']))]], hl=[1])}
<p class=note>Calmar here = CAGR ÷ max drawdown %. One cycle per year, so Sharpe/Sortino are per-cycle ratios on very few observations. These are in-sample numbers; see 6.10 for out-of-sample.</p>
<h3>6.7 Cumulative return and drawdown</h3>
{ch_cum}
{ch_dd}
<p class=note>Each line is the cumulative P&amp;L of that book as a % of its own margin. The walk-forward line starts in 2007 because it needs 7 prior years.</p>
</div>

<div class=sec>
<h2>6.8 Year by year (Final A, 1 lot per stock)</h2>
<table class="lg dense"><tr><th>Year</th><th>Laxmi Pujan</th><th>Trades</th><th>Wins</th><th>Win rate</th><th>Net P&amp;L</th><th>Return on margin</th><th>Avg trade</th><th>Fixed-window basket</th><th>Index fut T-5/T+4</th></tr>{yr_rows}</table>
<p class=note>P&amp;L = return × today's contract value for one lot per stock (not the prices actually traded that year); early years have fewer than 35 trades where a stock was not yet listed.</p>
{ch_bar}
<h3>6.9 Drawdowns at cycle level, Final A ({len(EPS)} occurred)</h3>
{table(['#','Peak','Trough','Recovered','Depth (₹)','% of capital','Fall (cycles)','Total length (cycles)'], [[i+1, e[0], e[1], e[2], inr(e[3]), pc(e[4]), e[5], e[6]] for i, e in enumerate(EPS)])}
<h3>6.10 Out-of-sample checks</h3>
{table(['Period','Role','Win rate','Avg trade','Annual return on margin','Max drawdown','Calmar','Cycles'], [
 ['Same 35 windows, 2000–2018','Used to choose (in-sample)',pc(TR35['wr']),spc(TR35['avg_ret']),pc(TR35['annual']),pc(TR35['maxdd_pct']),n2(TR35['calmar']),TR35['cycles']],
 ['Same 35 windows, 2019–2025','In-sample (windows chosen on all years)',pc(TF35['wr']),spc(TF35['avg_ret']),pc(TF35['annual']),pc(TF35['maxdd_pct']),n2(TF35['calmar']),TF35['cycles']],
 ['Windows re-chosen on 2000–2018, scored 2019–2025 (35)','Out-of-sample',pc(T35['wr']),spc(T35['avg_ret']),pc(T35['annual']),pc(T35['maxdd_pct']),n2(T35['calmar']),T35['cycles']],
 ['Same, all 50 stocks','Out-of-sample, no stock filter',pc(T50['wr']),spc(T50['avg_ret']),pc(T50['annual']),pc(T50['maxdd_pct']),n2(T50['calmar']),T50['cycles']],
 ['Walk-forward every year 2007–2025 (35)','Out-of-sample',pc(WF35['wr']),spc(WF35['avg_ret']),pc(WF35['annual']),pc(WF35['maxdd_pct']),n2(WF35['calmar']),WF35['cycles']],
 ['Nifty futures index, 2000–2025','Other instrument, T-5 to T+4',pc(IDX['wr']),spc(IDX['avg_ret']),'—','—','—',IDX['cycles']]], cls='dense')}
<p class=note>Reading: the honest out-of-sample expectation is the range between the walk-forward row ({pc(WF35['annual'])} a year) and the 2019–25 split ({pc(T35['annual'])} a year), with drawdowns of roughly {pc(T35['maxdd_pct'],0)} to {pc(WF35['maxdd_pct'],0)}, not the in-sample {pc(A['annual'])} / {pc(A['maxdd_pct'])}.</p>
</div>

<div class=sec>
<h2>7. Final Strategy B: the 9 STRONG names</h2>
<div class=warn>Same rules as Final A, restricted to the {nB} STRONG stocks: {strong_names}. Margin {lac(MG['B'])}.</div>
{table(['Metric','A (35)','B (9)','A after costs','B after costs'], [[lab, a_, b_, an_, bn_] for (lab,a_,b_,an_,bn_) in [
 ('Capital (margin)', inr(MG['A']), inr(MG['B']), inr(MG['A']), inr(MG['B'])), ('Trades', A['trades'], B['trades'], AN['trades'], BN['trades']), ('Net P&L', lac(A['net']), lac(B['net']), lac(AN['net']), lac(BN['net'])),
 ('Avg annual return (simple)', pc(A['annual']), pc(B['annual']), pc(AN['annual']), pc(BN['annual'])), ('Max drawdown (% of capital)', pc(A['maxdd_pct']), pc(B['maxdd_pct']), pc(AN['maxdd_pct']), pc(BN['maxdd_pct'])),
 ('Calmar', n2(A['calmar']), n2(B['calmar']), n2(AN['calmar']), n2(BN['calmar'])), ('Win trades', pc(A['wr']), pc(B['wr']), pc(AN['wr']), pc(BN['wr'])), ('Positive cycles', pc(A['pos_cycles'],0), pc(B['pos_cycles'],0), pc(AN['pos_cycles'],0), pc(BN['pos_cycles'],0)),
 ('Profit factor', n2(A['pf']), n2(B['pf']), n2(AN['pf']), n2(BN['pf'])), ('Worst cycle', lac(A['worst_cycle']), lac(B['worst_cycle']), lac(AN['worst_cycle']), lac(BN['worst_cycle']))]])}
<h3>7.1 Which one</h3>
{table(['Priority','Choose','Why'],[['Broadest, steadiest, matches the position list','A',f"{nA} names spread stock-specific noise; out-of-sample {pc(T35['annual'])} / {pc(T35['maxdd_pct'])} on the 2019–25 split"],['Less capital, higher in-sample return','B',f"{pc(B['annual'])} in-sample on {lac(MG['B'])}; but STRONG already embeds a walk-forward filter, so part of this is selection"],['Compare with a single common window','A or B',f"Fixed T-3/T+3 on all 50 gave {pc(FIX['annual'])} a year with {pc(FIX['maxdd_pct'])} drawdown"]])}
<h3>7.2 The three low-history stocks (experimental, not in A or B)</h3>
{table(['Stock','Side','Window','Entry','Exit','Years of data','Win rate','Avg'], exp_rows)}
<p class=note>MAXHEALTH, ETERNAL and JIOFIN have 3–6 Diwali years; a 100% win rate on 3–6 points is expected by chance among 100 windows. Treat as indicative; JIOFIN (the only SHORT) is the weakest.</p>
</div>

<div class=sec>
<h2>8. Assumptions and limitations</h2>
<ul><li><b>Few independent observations:</b> 26 Diwali events. The {A['trades']} trades of Final A all share each year's market backdrop; effective sample size is nearer 26 than {A['trades']}.</li>
<li><b>Chosen by search:</b> 128 combinations per stock, picked on the same years they are scored on. Even with walk-forward and a conviction gate, a searched rule usually does worse live than in the backtest — the out-of-sample rows are the better guide.</li>
<li><b>Look-ahead in the stock list:</b> which 35 stocks are in the book was decided using full-history evidence; the ‘all 50’ out-of-sample rows have no such filter.</li>
<li><b>Daily closes only:</b> entries and exits are at closes, not 09:20 / 15:15 intraday fills; real slippage, especially on stocks that move into the close, is not modelled.</li>
<li><b>P&amp;L scale:</b> ₹ P&amp;L uses today’s lot sizes and prices for every year, so it is a consistent weight, not what a trader would have earned at that year's prices. This is different from the earlier Diwali_Holiday_Performance_Report, which used raw prices of 2022–2025; its ₹ figures will not match.</li>
<li><b>Costs and margin:</b> 0.15% round trip and 20% margin are assumptions; actual SPAN + exposure and brokerage vary.</li>
<li><b>Universe bias:</b> only stocks that are in the Nifty 50 / F&amp;O list today; stocks that dropped out earlier are not tested.</li>
<li><b>Calendar:</b> the Laxmi Pujan dates 2000–2025 were hand-corrected and verified against price data; the 2026 holiday list assumes 10-Nov is closed. Re-check before trading.</li>
<li><b>Grid change:</b> this report uses the project's 8×8 grid. The earlier Diwali_Final_Strategy_Report and Diwali_Holiday_Performance_Report used a 10×10 scan and are superseded.</li>
<li><b>Gap risk:</b> a long future can lose more than any stop; there is no stop in the backtest.</li></ul>
<h2>9. Files</h2>
{table(['What','Where (D:\\behaviour analysis\\)'],[['Position list (book + status of all 50)','Diwali_2026_Nifty50_Position_List_8x8.json'],['3 low-history stocks','Diwali_2026_3_LowHistory_Stocks_Positions_8x8.json'],['Master file (positions + past data)','Diwali_Master_8x8.json'],['Backtest data (Nifty, 26 events)','Diwali_Past_Events_Backtest_Data_8x8.json'],['All F&O stocks data','Diwali_FO_All_Stocks_Data_8x8.json'],['Positions workbook (Nifty + F&O)','Diwali_Positions_8x8_Nifty_FO.xlsx'],['Per-stock scan workbooks','Nifty50_Diwali_Christmas_Report_v4.xlsx, FO_NonNifty50_Diwali_Christmas_Report_v4.xlsx'],['Code','build_diwali_christmas_perstock_reports_v4.py (scan), diwali_report_data.py (data layer), build_diwali_strategy_report_8x8.py (this report), build_diwali_8x8_lists.py']], cls='')}
</div></body></html>"""
for _a, _b in [('(35)', f'({nA})'), ('(35,', f'({nA},'), ('all 35', f'all {nA}'), ('same 35', f'same {nA}'), ('Same 35', f'Same {nA}'), ('the 35-stock', f'the {nA}-stock'), ('which 35 stocks', f'which {nA} stocks'),
               ('35 names', f'{nA} names'), ('fewer than 35 trades', f'fewer than {nA} trades'), ('(9)', f'({nB})'), ('9 STRONG', f'{nB} STRONG'), ('9 names', f'{nB} names')]:
    html = html.replace(_a, _b)
open(ROOT + 'Diwali_Final_Strategy_Report_8x8.html', 'w', encoding='utf-8').write(html)
from playwright.sync_api import sync_playwright
hdr = '<div style="font-family:Arial;font-size:8px;color:#666;width:100%;padding:0 14mm;display:flex;justify-content:space-between"><span>Nifty 50 Diwali Drift - final strategy report, project 8x8 grid (backtest 2000-2025, 26 events)</span><span>Page <span class="pageNumber"></span></span></div>'
with sync_playwright() as p:
    b = p.chromium.launch(); pg = b.new_page(); pg.set_content(html)
    pg.pdf(path=ROOT + 'Diwali_Final_Strategy_Report_8x8.pdf', format='A4', print_background=True, display_header_footer=True, header_template=hdr, footer_template='<span></span>', margin=dict(top='16mm', bottom='14mm', left='14mm', right='14mm'))
    b.close()
import re
print('pages', len(re.findall(rb'/Type\s*/Page[^s]', open(ROOT + 'Diwali_Final_Strategy_Report_8x8.pdf', 'rb').read())))
