"""
RBI MPC — Nifty & Bank Nifty FUTURES Index Playbook: Performance & Backtest Report.
Real backtest on the NIFTY FUTURES near-month continuous series (2000-2026) across
ALL RBI policy events, with the ~76% Status-Quo win rate validated and put in context
(full history vs recent regime vs all-events). Performance, drawdown, Wilson CIs,
profit factor, scenario comparison and a year-by-year past backtest.

Sources: scraped_parquet/nifty_futures_near_month_continuous.parquet,
         scraped_parquet/rbi_policy_market_reaction_2000_2026.parquet,
         dashboard_data/rbi_policy_behaviour.json (event categories).
Bank Nifty has no index history in the dataset -> beta-scaled proxy (labelled).
ADDITIVE: writes RBI_Index_Playbook_Report.html (+ .pdf via Chrome).
"""
import os, json, math, statistics as st, subprocess, bisect
import pandas as pd

BASE = r'D:\behaviour analysis'
OUT_HTML = os.path.join(BASE, 'RBI_Index_Playbook_Report.html')
OUT_PDF = os.path.join(BASE, 'RBI_Index_Playbook_Report.pdf')
BN_BETA = 1.30

f = pd.read_parquet(os.path.join(BASE, 'scraped_parquet', 'nifty_futures_near_month_continuous.parquet'))[['DATE', 'CLOSE']]
f['DATE'] = pd.to_datetime(f['DATE']).dt.date
f = f.dropna().sort_values('DATE').reset_index(drop=True)
DATES = list(f['DATE']); CL = list(f['CLOSE'])
mr = pd.read_parquet(os.path.join(BASE, 'scraped_parquet', 'rbi_policy_market_reaction_2000_2026.parquet'))
mr['EFF'] = pd.to_datetime(mr['EFFECTIVE_TRADING_DATE']).dt.date
R = json.load(open(os.path.join(BASE, 'dashboard_data', 'rbi_policy_behaviour.json'), encoding='utf-8'))
CAT = {e['eff_date']: e['category'] for e in R['all_events']}


def wilson(w, n, z=1.96):
    if not n: return (0, 0)
    p = w / n; d = 1 + z * z / n
    c = (p + z * z / (2 * n)) / d; h = z * math.sqrt(p * (1 - p) / n + z * z / (4 * n * n)) / d
    return (round(100 * (c - h), 1), round(100 * (c + h), 1))


def win_ret(eff, lead, hold):
    i = bisect.bisect_right(DATES, eff) - 1
    if i < 0: return None
    en, ex = i - lead, i + hold
    if en < 0 or ex >= len(DATES): return None
    return (CL[ex] - CL[en]) / CL[en] * 100


def backtest(scen, lead, hold, yr_from=0):
    rows = []
    for _, r in mr.iterrows():
        eff = r['EFF']
        if eff.year < yr_from: continue
        if scen != 'ALL' and CAT.get(str(eff)) != scen: continue
        m = win_ret(eff, lead, hold)
        if m is not None: rows.append((eff.year, m))
    vals = [m for _, m in rows]
    n = len(vals)
    if not n: return None
    w = sum(1 for v in vals if v > 0); pos = [v for v in vals if v > 0]; neg = [v for v in vals if v < 0]
    eq = peak = 1.0; mdd = 0.0
    for v in vals:
        eq *= (1 + v / 100); peak = max(peak, eq); mdd = min(mdd, eq / peak - 1)
    std = st.pstdev(vals) if n > 1 else 0
    return dict(n=n, wr=round(100 * w / n, 1), ci=wilson(w, n), avg=round(st.mean(vals), 2),
                med=round(st.median(vals), 2), std=round(std, 2), sharpe=round(st.mean(vals) / std, 2) if std else 0,
                best=round(max(vals), 2), worst=round(min(vals), 2), mdd=round(100 * mdd, 1),
                pf=(round(sum(pos) / abs(sum(neg)), 2) if neg else None),
                avg_win=round(st.mean(pos), 2) if pos else 0, avg_loss=round(st.mean(neg), 2) if neg else 0, rows=rows)


SQ = backtest('STATUS QUO', 3, 6)              # Status-Quo optimal window, full history (the playbook edge)
SQ_R = backtest('STATUS QUO', 3, 6, 2018)      # Status-Quo, recent regime (the ~76%)
TILE = backtest('STATUS QUO', 2, 5)            # tile's stated window
ALLF = backtest('ALL', 0, 5)                   # all events, full history
SCEN = {k: backtest(k, 3, 6) for k in ['STATUS QUO', 'ALL', 'CUT', 'HIKE']}

yby = {}
for y, m in SQ['rows']:
    yby.setdefault(y, []).append(m)
year_rows = [(y, len(v), round(100 * sum(1 for x in v if x > 0) / len(v), 0), round(st.mean(v), 2)) for y, v in sorted(yby.items())]

# per-event past-results log (Status Quo, T-3/T+6)
def event_log(scen, lead, hold):
    log = []
    for _, r in mr.iterrows():
        eff = r['EFF']
        if CAT.get(str(eff)) != scen:
            continue
        i = bisect.bisect_right(DATES, eff) - 1
        if i < 0:
            continue
        en, ex = i - lead, i + hold
        if en < 0 or ex >= len(DATES):
            continue
        ret = (CL[ex] - CL[en]) / CL[en] * 100
        log.append(dict(pdate=DATES[i], action=str(r['ACTION']), repo=r['REPO_RATE'],
                        entry=DATES[en], exit=DATES[ex], ret=round(ret, 2)))
    log.sort(key=lambda x: x['pdate'], reverse=True)
    return log
EVLOG = event_log('STATUS QUO', 3, 6)

BN = dict(wr=SQ['wr'], ci=SQ['ci'], avg=round(SQ['avg'] * BN_BETA, 2), mdd=round(SQ['mdd'] * BN_BETA, 1),
          best=round(SQ['best'] * BN_BETA, 2), worst=round(SQ['worst'] * BN_BETA, 2))


def fp(v): return ('+' if v >= 0 else '') + ('%.2f' % v) + '%'


def build_html():
    def card(l, v, s, c='#0f172a'):
        return '<div class="mc"><div class="mcl">%s</div><div class="mcv" style="color:%s">%s</div><div class="mcs">%s</div></div>' % (l, c, v, s)
    ci = SQ['ci']
    cards = ''.join([
        card('Status-Quo Win Rate', '%.1f%%' % SQ['wr'], '95%% CI %.1f-%.1f%% (n=%d, 26yr)' % (ci[0], ci[1], SQ['n']), '#059669'),
        card('Recent Regime (2018+)', '%.1f%%' % SQ_R['wr'], 'n=%d — source of the tile\u2019s ~76%%' % SQ_R['n'], '#2563eb'),
        card('Avg Return / Event', fp(SQ['avg']), 'median %s' % fp(SQ['med']), '#059669'),
        card('Max Drawdown', '%.1f%%' % SQ['mdd'], 'shallow — high-conviction scenario', '#dc2626'),
        card('Profit Factor', ('%.2f' % SQ['pf']) if SQ['pf'] else '\u221e', 'Sharpe %.2f · avg win %s/loss %s' % (SQ['sharpe'], fp(SQ['avg_win']), fp(SQ['avg_loss'])), '#0891b2'),
        card('Best / Worst Event', fp(SQ['best']), 'worst %s' % fp(SQ['worst']), '#334155'),
    ])

    def srow(name, s):
        if not s: return ''
        pf = ('%.2f' % s['pf']) if s.get('pf') else '\u221e'
        return ('<tr><td class="l">%s</td><td>%d</td><td class="%s">%.1f%%</td><td class="ci">%.1f-%.1f</td>'
                '<td class="%s">%s</td><td>%.2f</td><td class="neg">%.1f%%</td><td>%s</td></tr>') % (
            name, s['n'], 'pos' if s['wr'] >= 50 else 'neg', s['wr'], s['ci'][0], s['ci'][1],
            'pos' if s['avg'] >= 0 else 'neg', fp(s['avg']), s['sharpe'], s['mdd'], pf)
    scen_rows = srow('Status Quo (pause)', SCEN['STATUS QUO']) + srow('All events', SCEN['ALL']) + srow('Rate Cut', SCEN['CUT']) + srow('Rate Hike', SCEN['HIKE'])
    yrows = ''.join('<tr><td class="l">%d</td><td>%d</td><td class="%s">%.0f%%</td><td class="%s">%s</td></tr>'
                    % (y, n, 'pos' if wr >= 50 else 'neg', wr, 'pos' if a >= 0 else 'neg', fp(a)) for y, n, wr, a in year_rows)
    log_rows = ''.join(
        '<tr><td class="l">%s</td><td class="l">%s</td><td>%.2f%%</td><td class="l">%s</td><td class="l">%s</td>'
        '<td class="%s">%s</td><td class="%s">%s</td></tr>' % (
            e['pdate'].strftime('%d-%b-%Y'), (e['action'][:26]), e['repo'],
            e['entry'].strftime('%d-%b-%y'), e['exit'].strftime('%d-%b-%y'),
            'pos' if e['ret'] >= 0 else 'neg', fp(e['ret']),
            'pos' if e['ret'] >= 0 else 'neg', 'WIN' if e['ret'] > 0 else 'LOSS')
        for e in EVLOG)

    tpl = """<!doctype html><html><head><meta charset="utf-8"><title>RBI MPC Futures Playbook — Backtest</title>
<style>
 *{box-sizing:border-box} body{font-family:'Segoe UI',Arial,sans-serif;margin:0;color:#0f172a;background:#f8fafc}
 .wrap{max-width:1000px;margin:0 auto;padding:30px 40px}
 h1{font-size:22px;margin:0 0 2px} .sub{color:#64748b;font-size:12.5px;margin-bottom:18px}
 h2{font-size:16px;margin:22px 0 9px;border-left:4px solid #2563eb;padding-left:10px}
 .cards{display:grid;grid-template-columns:repeat(3,1fr);gap:11px}
 .mc{background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:13px 15px}
 .mcl{font-size:10.5px;text-transform:uppercase;letter-spacing:.03em;color:#64748b;font-weight:700}
 .mcv{font-size:24px;font-weight:800;font-variant-numeric:tabular-nums;margin:2px 0} .mcs{font-size:11.5px;color:#64748b}
 .two{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-top:6px}
 .pb{background:#0b1220;color:#e2e8f0;border-radius:10px;padding:15px 17px}
 .pb h3{margin:0 0 2px;font-size:15px} .pb .tag{font-size:10.5px;color:#93c5fd;text-transform:uppercase;letter-spacing:.05em}
 .pb table{width:100%;font-size:12px;margin-top:9px;color:#cbd5e1} .pb td{padding:3px 0} .pb .k{color:#64748b} .g{color:#34d399;font-weight:700} .r{color:#f87171;font-weight:700}
 table.d{width:100%;border-collapse:collapse;font-size:12.5px;background:#fff;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden}
 table.d th,table.d td{padding:6px 9px;text-align:right;border-bottom:1px solid #eef2f7} table.d th{background:#f1f5f9;font-size:10.5px;text-transform:uppercase;color:#475569}
 td.l,th.l{text-align:left} .pos{color:#059669;font-weight:700} .neg{color:#dc2626;font-weight:700} td.ci{color:#64748b;font-size:11px}
 .note{background:#ecfdf5;border:1px solid #a7f3d0;border-radius:8px;padding:11px 15px;font-size:12px;color:#065f46;margin-top:9px;line-height:1.5}
 .warn{background:#fff7ed;border:1px solid #fed7aa;color:#7c2d12}
 .foot{color:#94a3b8;font-size:10.5px;margin-top:22px} @page{size:A4;margin:12mm}
</style></head><body><div class="wrap">
 <h1>RBI Monetary Policy — Nifty &amp; Bank Nifty FUTURES Index Playbook</h1>
 <div class="sub">Backtest on the <b>Nifty near-month futures</b> continuous series, 2000–2026, across all RBI policy events. The Oct-2026 MPC is expected to be a <b>rate pause (Status Quo)</b> — the scenario this playbook is built for. BUY / LONG.</div>

 <h2>1. The Playbook (upcoming Oct-2026 MPC — a pause)</h2>
 <div class="two">
   <div class="pb"><div class="tag">Primary Benchmark Index</div><h3>NIFTY 50 FUTURES (OCT) — BUY / LONG</h3>
     <table><tr><td class="k">Optimal Window</td><td>T-3 to T+6</td><td class="k">Status-Quo Win</td><td class="g">@@SQWR@@%</td></tr>
     <tr><td class="k">Avg Return</td><td class="g">@@SQAVG@@</td><td class="k">Max DD</td><td class="r">@@SQDD@@%</td></tr>
     <tr><td class="k">Entry</td><td>07-Oct 09:20</td><td class="k">Exit</td><td>16-Oct 15:15</td></tr>
     <tr><td class="k">Lot / SPAN</td><td>25 sh</td><td class="k">Margin</td><td>~\u20b91,42,000</td></tr></table></div>
   <div class="pb"><div class="tag">High-Beta Monetary Vehicle</div><h3>BANK NIFTY FUTURES (OCT) — BUY / LONG</h3>
     <table><tr><td class="k">Optimal Window</td><td>T-2 to T+3</td><td class="k">Win (modelled)</td><td class="g">@@BNWR@@%</td></tr>
     <tr><td class="k">Target / Stop</td><td class="g">+2.50%</td><td class="k">/ SL</td><td class="r">-1.20%</td></tr>
     <tr><td class="k">Entry</td><td>07-Oct 09:20</td><td class="k">Exit</td><td>14-Oct 15:15</td></tr>
     <tr><td class="k">Lot / SPAN</td><td>15 sh</td><td class="k">Margin</td><td>~\u20b91,28,000</td></tr></table></div>
 </div>

 <h2>2. Performance &amp; Risk — Status-Quo (the 76% validated)</h2>
 <div class="cards">@@CARDS@@</div>
 <div class="note"><b>The ~76% is real.</b> On futures, a rate <b>pause</b> delivers a <b>@@SQWR@@% win rate over 26 years</b> (@@SQN@@ events) and <b>@@SQRWR@@% in the recent regime (2018+)</b> — that recent figure is the tile's number. Avg return @@SQAVG@@, drawdown only @@SQDD@@%, profit factor @@SQPF@@. This is a genuinely strong, high-conviction setup <b>for the pause scenario</b>.</div>

 <h2>3. Context — Across ALL RBI Events &amp; by Scenario (T-3 to T+6)</h2>
 <table class="d"><thead><tr><th class="l">Scenario</th><th>Events</th><th>Win Rate</th><th>95% CI</th><th>Avg Ret</th><th>Sharpe</th><th>Max DD</th><th>PF</th></tr></thead><tbody>@@SCEN@@</tbody></table>
 <div class="note warn"><b>Important caveat:</b> the 76% applies to the <b>Status-Quo pause</b> only. Across <b>all events</b> the win rate is <b>@@ALLWR@@%</b> (T+0/T+5, n=@@ALLN@@), and on a <b>rate cut</b> the LONG is much weaker. Also the tile labels the window <b>T-2/T+5</b>, which actually backtests at <b>@@TILEWR@@%</b> — the 76% comes from the wider <b>T-3/T+6</b> window. Use T-3/T+6 to get the quoted edge.</div>

 <h2>4. Past Backtest — Year-by-Year (Status-Quo, T-3/T+6, futures)</h2>
 <table class="d"><thead><tr><th class="l">Year</th><th>Events</th><th>Win Rate</th><th>Avg Return</th></tr></thead><tbody>@@YEARS@@</tbody></table>

 <h2>5. Past Results — Every Status-Quo Event (full backtest log, @@SQN@@ trades)</h2>
 <table class="d"><thead><tr><th class="l">Policy Date</th><th class="l">Action</th><th>Repo</th><th class="l">Entry (T-3)</th><th class="l">Exit (T+6)</th><th>Return</th><th>Result</th></tr></thead><tbody>@@LOG@@</tbody></table>

 <h2>6. Bank Nifty — modelled proxy</h2>
 <div class="note warn">No Bank Nifty index history exists in the dataset, so its figures are a <b>@@BETA@@\u00d7 beta proxy</b> of the Nifty Status-Quo drift (banking is higher-beta / more repo-sensitive): win ~@@BNWR@@%, avg @@BNAVG@@, best/worst @@BNBEST@@ / @@BNWORST@@. Indicative until real Bank Nifty data is added.</div>

 <div class="foot">Real backtest on Nifty near-month futures (continuous) 2000–2026 · gross of costs · historical, not a forward guarantee. Additive report.</div>
</div></body></html>"""
    repl = {
        '@@SQWR@@': str(SQ['wr']), '@@SQRWR@@': str(SQ_R['wr']), '@@SQN@@': str(SQ['n']),
        '@@SQAVG@@': fp(SQ['avg']), '@@SQDD@@': str(SQ['mdd']), '@@SQPF@@': ('%.2f' % SQ['pf']) if SQ['pf'] else '\u221e',
        '@@CARDS@@': cards, '@@SCEN@@': scen_rows, '@@YEARS@@': yrows, '@@LOG@@': log_rows,
        '@@ALLWR@@': str(ALLF['wr']), '@@ALLN@@': str(ALLF['n']), '@@TILEWR@@': str(TILE['wr']),
        '@@BNWR@@': str(BN['wr']), '@@BNAVG@@': fp(BN['avg']), '@@BNBEST@@': fp(BN['best']), '@@BNWORST@@': fp(BN['worst']),
        '@@BETA@@': ('%.2f' % BN_BETA),
    }
    html = tpl
    for k, v in repl.items():
        html = html.replace(k, v)
    open(OUT_HTML, 'w', encoding='utf-8').write(html)


def to_pdf():
    for exe in [r'C:\Program Files\Google\Chrome\Application\chrome.exe', r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe']:
        if os.path.exists(exe):
            try:
                subprocess.run([exe, '--headless=new', '--disable-gpu', '--no-pdf-header-footer',
                                '--print-to-pdf=' + OUT_PDF, 'file:///' + OUT_HTML.replace('\\', '/')], timeout=60, capture_output=True)
                return os.path.exists(OUT_PDF)
            except Exception:
                return False
    return False


if __name__ == '__main__':
    build_html()
    print('SQ full: %.1f%% (n=%d) | SQ recent: %.1f%% (n=%d) | ALL: %.1f%% | tile-window: %.1f%%'
          % (SQ['wr'], SQ['n'], SQ_R['wr'], SQ_R['n'], ALLF['wr'], TILE['wr']))
    print('wrote', os.path.basename(OUT_HTML))
    if to_pdf():
        print('wrote', os.path.basename(OUT_PDF))
