"""
Data layer for the long-form Diwali strategy report (SENSEX-report format).
Builds every table used by build_diwali_strategy_report.py and pickles it to portfolio_cache/diwali_report.pkl.
Additive only: reads existing data, writes one new pickle.
"""
import sys, json, pickle
sys.stdout.reconfigure(encoding='utf-8')
ROOT = 'D:/behaviour analysis/'
src = open(ROOT + 'build_diwali_christmas_perstock_reports_v4.py', encoding='utf-8').read()
head = src.split('# ---------------------------------------------------------------- per-stock analysis')[0].replace('json.dump(EVENT_WINDOW', '(lambda *a, **k: None)(EVENT_WINDOW').replace(r'open(ROOT + r"\Diwali_Christmas_event_windows_v4.json", "w")', 'open(__import__("os").devnull, "w")')  # DEVNULLFIX: never truncate the tracked windows file
ns = {}; exec(compile(head, 'v4', 'exec'), ns)
np, pd = ns['np'], ns['pd']
data = ns['data']['Diwali']
for _s in data: data[_s]['R'] = data[_s]['R'][:, :8, :8]          # PROJECT 8x8 GRID (lead 1..8 x hold 1..8)
YEARS = ns['YEARS']; select = ns['select']; MIN_YEARS = ns['MIN_YEARS']
lots, meta, NIFTY50, ANCH, rep = ns['lots'], ns['meta'], ns['NIFTY50'], ns['ANCH'], ns['rep']
COST = 0.15
N50 = sorted(s for s in NIFTY50 if s in data)
last_close = {s: float(rep.LatestClose[s]) for s in data}
notional = {s: float(lots[s]) * last_close[s] for s in data}              # 1-lot contract value at today's price
margin_of = lambda syms: 0.2 * sum(notional[s] for s in syms)
T_DATE = {y: pd.Timestamp(ANCH['Diwali'][y]) for y in YEARS + [2026]}

# ---------------- positions on the 8x8 grid: same per-stock rule + conviction gate as the v4 scan, applied to the 8x8 cube
dash = json.load(open(ROOT + 'dashboard_data/event_dashboard_data.json', encoding='utf-8'))
sector = {s['symbol']: s.get('sector', '') for s in [x for x in dash['holidays'] if x['id'] == 'diwali_2026'][0]['stocks']}
closed = set(ns['hol'][ns['hol'].Is_Trading_Holiday == 'YES'].Holiday_Date); closed.add(pd.Timestamp('2026-11-10'))
_T26 = pd.Timestamp('2026-11-08'); _c26 = [d for d in pd.bdate_range(_T26 - pd.Timedelta(days=40), _T26 + pd.Timedelta(days=40)) if d not in closed]
pre26 = [d for d in _c26 if d < _T26][-8:]; post26 = [d for d in _c26 if d > _T26][:8]
def analyse(s):
    C = data[s]['R']; sel = select(C)
    if sel is None: return None
    side, li, hi, wr, avg, n = sel; sg = 1 if side == 'LONG' else -1
    oos = []
    for yi in range(len(YEARS)):
        if np.isnan(C[yi]).all(): continue
        p = select(C[:yi]) if yi >= MIN_YEARS else None
        if p is None: continue
        r = (1 if p[0] == 'LONG' else -1) * C[yi, p[1], p[2]]
        if not np.isnan(r): oos.append(float(r))
    oos = np.array(oos); owr = float(np.mean(oos > 0) * 100) if len(oos) else None
    if n >= 10 and wr >= 75 and avg > 0 and len(oos) >= 5 and owr >= 60: conv = 'STRONG'
    elif wr >= 65 and avg > 0 and (len(oos) < 5 or owr >= 50): conv = 'MODERATE'
    else: conv = 'NO TRADE'
    v = sg * C[:, li, hi]; v = v[~np.isnan(v)]; l10 = v[-10:]
    return dict(direction=side, lead=int(li + 1), hold=int(hi + 1), wr=round(float(wr), 1), avg=round(float(avg), 2), n=int(n), wf=None if owr is None else round(owr, 1), wf_n=int(len(oos)), conv=conv,
                median=round(float(np.median(v)), 2), best=round(float(v.max()), 2), worst=round(float(v.min()), 2), last10=round(float(np.mean(l10 > 0) * 100), 1))
ALL50 = []; POS = []
for s in N50 + [x for x in sorted(NIFTY50) if x not in N50]:
    r = analyse(s) if s in data else None
    row = dict(symbol=s, name=meta.get(s, {}).get('name', ''), sector=sector.get(s, ''), status='DATA INSUFFICIENT' if r is None else ('TRADE' if r['conv'] != 'NO TRADE' else 'NO TRADE'), analysis=r)
    ALL50.append(row)
    if r and r['conv'] != 'NO TRADE':
        POS.append(dict(symbol=s, name=row['name'], direction=r['direction'], lead=r['lead'], hold=r['hold'], wr=r['wr'], avg=r['avg'], wf=r['wf'], wf_n=r['wf_n'], conv=r['conv'], repick=False, sector=row['sector']))
SYM35 = [p['symbol'] for p in POS]; SYM9 = [p['symbol'] for p in POS if p['conv'] == 'STRONG']
print('8x8 book:', len(POS), 'positions;', len(SYM9), 'STRONG;', sum(p['direction'] == 'SHORT' for p in POS), 'SHORT;', sum(r['status'] == 'NO TRADE' for r in ALL50), 'NO TRADE;', sum(r['status'] == 'DATA INSUFFICIENT' for r in ALL50), 'insufficient')
# long vs short check on 8x8 (Nifty / non-Nifty)
def best_wr(C, sg):
    v = ~np.isnan(C); n = v.sum(0)
    with np.errstate(all='ignore'): w = ((sg * C > 0) & v).sum(0) / n * 100
    w = np.where(n >= MIN_YEARS, w, -1); return None if w.max() < 0 else float(w.max())
LS = {}
for grp, test in (('NIFTY', lambda s: s in NIFTY50), ('FO', lambda s: s not in NIFTY50)):
    L_, S_, T_, bl_, bs_ = 0, 0, 0, [], []
    for s in data:
        if not test(s): continue
        a, b = best_wr(data[s]['R'], 1), best_wr(data[s]['R'], -1)
        if a is None and b is None: continue
        a, b = a if a is not None else -1, b if b is not None else -1
        bl_.append(a); bs_.append(b); L_ += a > b + 0.05; S_ += b > a + 0.05; T_ += abs(a - b) <= 0.05
    LS[grp] = dict(tested=len(bl_), long=int(L_), short=int(S_), tie=int(T_), avg_long=float(np.mean(bl_)), avg_short=float(np.mean(bs_)))
print('LS', LS)
# experimental positions for the 3 low-history stocks (best window over their FULL available history, 8x8)
EXP = []
for s in ['MAXHEALTH', 'ETERNAL', 'JIOFIN']:
    if s not in data: continue
    C = data[s]['R']; v = ~np.isnan(C); n = v.sum(0); mx = int(n.max())
    with np.errstate(all='ignore'): wl = ((C > 0) & v).sum(0) / n; ws_ = ((C < 0) & v).sum(0) / n; av = np.nanmean(C, 0)
    ok_ = n >= mx; sl = np.where(ok_, wl + 1e-6 * av, -1); ss = np.where(ok_, ws_ - 1e-6 * av, -1)
    if sl.max() >= ss.max(): side = 'LONG'; i, j = np.unravel_index(np.argmax(sl), sl.shape); wr = wl[i, j]; a_ = av[i, j]
    else: side = 'SHORT'; i, j = np.unravel_index(np.argmax(ss), ss.shape); wr = ws_[i, j]; a_ = -av[i, j]
    lead, hold = int(i) + 1, int(j) + 1
    EXP.append(dict(symbol=s, name=meta.get(s, {}).get('name', ''), direction=side, window=f'T-{lead} to T+{hold}', lead=lead, hold=hold, entry_date=pre26[-lead].strftime('%d-%b-%Y (%A)'), exit_date=post26[hold - 1].strftime('%d-%b-%Y (%A)'),
                    valid_years=mx, win_rate=round(float(wr) * 100, 1), avg_signed_ret=round(float(a_), 2), conviction='EXPERIMENTAL - NOT BACKTESTED TO BAR',
                    reason=f'Only {mx} Diwali years of data (min 7 required); best of {side} windows over those years.'))

def trades_from(pairs, name):
    rows = []
    for s, lead, hold, sg in pairs:
        C = data[s]['R']
        for yi, y in enumerate(YEARS):
            r = C[yi, lead - 1, hold - 1]
            if np.isnan(r): continue
            rows.append(dict(book=name, year=y, symbol=s, lead=lead, hold=hold, ret=sg * r, pnl=sg * r / 100 * notional[s], cost=COST / 100 * notional[s]))
    return pd.DataFrame(rows)

BOOKS = {}
BOOKS['A'] = trades_from([(p['symbol'], p['lead'], p['hold'], 1 if p['direction'] == 'LONG' else -1) for p in POS], 'A')
BOOKS['B'] = BOOKS['A'][BOOKS['A'].symbol.isin(SYM9)].copy()
BOOKS['FIX'] = trades_from([(s, 3, 3, 1) for s in N50], 'FIX')
def wf_book(syms, first=MIN_YEARS, last=len(YEARS), name='WF'):
    rows = []
    for s in syms:
        C = data[s]['R']
        for yi in range(first, last):
            p = select(C[:yi])
            if p is None: continue
            side, li, hi = p[0], p[1], p[2]; sg = 1 if side == 'LONG' else -1; r = C[yi, li, hi]
            if np.isnan(r): continue
            rows.append(dict(book=name, year=YEARS[yi], symbol=s, lead=li + 1, hold=hi + 1, ret=sg * r, pnl=sg * r / 100 * notional[s], cost=COST / 100 * notional[s]))
    return pd.DataFrame(rows)
BOOKS['WF35'] = wf_book(SYM35, name='WF35'); BOOKS['WF50'] = wf_book(N50, name='WF50')
# train (<=2018) / test (2019-2025): rule chosen on 2000-2018 only
BOOKS['TEST35'] = wf_book(SYM35, 0, 0, 'TEST35'); BOOKS['TEST50'] = wf_book(N50, 0, 0, 'TEST50')
def split_book(syms, name):
    rows = []; cut = YEARS.index(2018) + 1
    for s in syms:
        C = data[s]['R']; p = select(C[:cut])
        if p is None: continue
        sg = 1 if p[0] == 'LONG' else -1
        for yi in range(cut, len(YEARS)):
            r = C[yi, p[1], p[2]]
            if np.isnan(r): continue
            rows.append(dict(book=name, year=YEARS[yi], symbol=s, lead=p[1] + 1, hold=p[2] + 1, ret=sg * r, pnl=sg * r / 100 * notional[s], cost=COST / 100 * notional[s]))
    return pd.DataFrame(rows)
BOOKS['TEST35'] = split_book(SYM35, 'TEST35'); BOOKS['TEST50'] = split_book(N50, 'TEST50')
# training-period (<=2018) in-sample score of the final windows, same stocks
TRAIN35 = BOOKS['A'][BOOKS['A'].year <= 2018]; TEST_FIXEDWIN35 = BOOKS['A'][BOOKS['A'].year >= 2019]
BOOKS['TRAIN35'] = TRAIN35.copy(); BOOKS['TESTFIX35'] = TEST_FIXEDWIN35.copy()
# Nifty futures index
fj = json.load(open(ROOT + 'dashboard_data/nifty_futures_holiday_behaviour.json', encoding='utf-8'))
fh = [x for x in fj['holidays'] if x['id'] == 'diwali'][0]
BOOKS['IDX'] = pd.DataFrame([dict(book='IDX', year=t['year'], symbol='NIFTYFUT', lead=5, hold=4, ret=t['ret_pct'], pnl=t['pnl_rs'], cost=0.0) for t in fh['trades']])
MARGIN = {'A': margin_of(SYM35), 'B': margin_of(SYM9), 'FIX': margin_of(N50), 'WF35': margin_of(SYM35), 'WF50': margin_of(N50), 'TEST35': margin_of(SYM35), 'TEST50': margin_of(N50), 'TRAIN35': margin_of(SYM35), 'TESTFIX35': margin_of(SYM35), 'IDX': float(fj['margin_per_lot'])}

def metrics(df, margin, net=False):
    d = df.copy(); d['p'] = d.pnl - (d.cost if net else 0)
    W, L = d[d.p > 0], d[d.p <= 0]
    cyc = d.groupby('year').agg(pnl=('p', 'sum'), n=('p', 'size'), w=('p', lambda x: (x > 0).sum()), ret=('ret', 'mean'))
    cyc['r'] = cyc.pnl / margin; cum = cyc.pnl.cumsum(); peak = np.maximum.accumulate(np.r_[0, cum.values])[1:]; dd = cum.values - peak
    ddp = dd.min() / margin * 100
    streak = 0; best = 0
    for v in dd:
        streak = streak + 1 if v < 0 else 0; best = max(best, streak)
    ls = 0; bl = 0
    for v in cyc.pnl:
        ls = ls + 1 if v < 0 else 0; bl = max(bl, ls)
    cagr = (np.prod(1 + cyc.r.values) ** (1 / len(cyc)) - 1) * 100 if (1 + cyc.r.values).min() > 0 else np.nan
    down = cyc.r[cyc.r < 0]
    return dict(trades=len(d), wins=len(W), losses=len(L), wr=len(W) / len(d) * 100, avg_ret=d.ret.mean(), avg_win=W.ret.mean(), avg_loss=L.ret.mean() if len(L) else np.nan,
                pf=W.p.sum() / abs(L.p.sum()) if len(L) and L.p.sum() != 0 else np.nan, payoff=(W.p.mean() / abs(L.p.mean())) if len(L) else np.nan,
                best_trade=d.p.max(), worst_trade=d.p.min(), cycles=len(cyc), pos_cycles=(cyc.pnl > 0).mean() * 100, net=d.p.sum(), avg_cycle=cyc.pnl.mean(),
                best_cycle=cyc.pnl.max(), best_cycle_y=int(cyc.pnl.idxmax()), worst_cycle=cyc.pnl.min(), worst_cycle_y=int(cyc.pnl.idxmin()),
                annual=cyc.r.mean() * 100, cagr=cagr, maxdd=dd.min(), maxdd_pct=ddp, calmar=(cagr / abs(ddp)) if ddp < 0 and not np.isnan(cagr) else np.nan,
                sharpe=cyc.r.mean() / cyc.r.std(ddof=1) if len(cyc) > 2 else np.nan, sortino=cyc.r.mean() / np.sqrt((down ** 2).sum() / len(cyc)) if len(down) else np.nan,
                recovery=d.p.sum() / abs(dd.min()) if dd.min() < 0 else np.nan, longest_dd=best, max_lose_cycles=bl, per_cycle=cyc.n.mean(), cost=d.cost.sum(), losing_cycles=int((cyc.pnl < 0).sum()),
                cyc=cyc.assign(cum=cum.values, dd=dd))
M = {k: metrics(v, MARGIN[k]) for k, v in BOOKS.items()}
MN = {k: metrics(v, MARGIN[k], net=True) for k, v in BOOKS.items()}
for k in ['A', 'B', 'FIX', 'WF35', 'WF50', 'TEST35', 'TEST50', 'TRAIN35', 'TESTFIX35', 'IDX']:
    m = M[k]; print(f"{k:7} n={m['trades']:4} WR={m['wr']:.1f} avg={m['avg_ret']:+.2f} PF={m['pf']:.2f} cycles={m['cycles']} poscyc={m['pos_cycles']:.0f}% ann={m['annual']:.1f}% DD={m['maxdd_pct']:.1f}% calmar={m['calmar']:.2f} net={m['net']:,.0f} margin={MARGIN[k]:,.0f}")

# ---------------- extra tables
# basket year table (Nifty 50 equal weight, fixed T-3/T+3 LONG) and per-year cube heatmaps
N50R = np.stack([data[s]['R'] for s in N50])                      # stock x year x lead x hold
BASK = np.nanmean(N50R, axis=0)                                   # year x lead x hold  (mean LONG return)
WRG = np.nanmean((N50R > 0).astype(float) * ~np.isnan(N50R), axis=(0, 1)) / np.maximum(np.nanmean(~np.isnan(N50R), axis=(0, 1)), 1e-9)
HM_ret = np.nanmean(BASK, axis=0); HM_wr = np.array([[np.mean(BASK[~np.isnan(BASK[:, i, j]), i, j] > 0) * 100 for j in range(8)] for i in range(8)])
# average path around Diwali (basket of the 35 final LONG names), offsets T-10 .. T+10 -> index 0..19 ; base = close at T-10
paths = {}
for yi, y in enumerate(YEARS):
    cur = []
    for s in SYM35:
        D = data[s]; ip, iq = D['IP'][yi], D['IQ'][yi]
        if ip < 0 or ip - 7 < 0 or iq + 7 >= len(D['a']): continue
        seg = np.r_[D['a'][ip - 7:ip + 1], D['a'][iq:iq + 8]]; cur.append(seg / seg[0] * 100 - 100)
    if cur: paths[y] = np.mean(cur, axis=0)
# conviction / sector / side tables
def sub(bk, syms): return metrics(BOOKS[bk][BOOKS[bk].symbol.isin(syms)], margin_of(syms)) if len(syms) else None
conv_tbl = {c: [p['symbol'] for p in POS if p['conv'] == c] for c in ['STRONG', 'MODERATE']}
sector_tbl = {}
for p in POS: sector_tbl.setdefault(p['sector'] or 'Other', []).append(p['symbol'])
# search scale
n_pairs = sum(1 for s in data)
# per-year basket of the fixed window, year table
year_tbl = pd.DataFrame({'year': YEARS, 'T': [T_DATE[y] for y in YEARS], 'fix_ret': [np.nanmean(N50R[:, i, 2, 2]) for i in range(len(YEARS))],
                         'fix_wr': [np.nanmean(N50R[:, i, 2, 2] > 0) * 100 for i in range(len(YEARS))], 'n': [int((~np.isnan(N50R[:, i, 2, 2])).sum()) for i in range(len(YEARS))]})
# nifty-universe summary counts from earlier check
res = dict(ALL50=ALL50, LS=LS, EXP=EXP, pre26=pre26, post26=post26, BOOKS=BOOKS, M=M, MN=MN, MARGIN=MARGIN, POS=POS, SYM35=SYM35, SYM9=SYM9, N50=N50, HM_ret=HM_ret, HM_wr=HM_wr, paths=paths, year_tbl=year_tbl, T_DATE=T_DATE,
           notional=notional, last_close=last_close, lots={s: float(lots[s]) for s in data}, conv_tbl=conv_tbl, sector_tbl=sector_tbl, fh=fh, YEARS=YEARS, n_stocks=len(data),
           TRAIN35=TRAIN35, TEST_FIXEDWIN35=TEST_FIXEDWIN35)
import os; os.makedirs(ROOT + 'portfolio_cache', exist_ok=True)
pickle.dump(res, open(ROOT + 'portfolio_cache/diwali_report_8x8.pkl', 'wb'))
print('saved; fixed-window basket years up:', int((year_tbl.fix_ret > 0).sum()), 'of', len(year_tbl))
print('train-35 (<=2018, final windows):', f"WR {(TRAIN35.ret>0).mean()*100:.1f}% avg {TRAIN35.ret.mean():.2f}", '| test-35 same windows 2019+:', f"WR {(TEST_FIXEDWIN35.ret>0).mean()*100:.1f}% avg {TEST_FIXEDWIN35.ret.mean():.2f}")
print(M['A']['cyc'][['pnl', 'n', 'w']].tail(8))
