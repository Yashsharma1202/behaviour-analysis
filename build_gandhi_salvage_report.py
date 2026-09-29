"""
GANDHI JAYANTI (2-Oct-2026) — MISSED POSITIONS & NEW-ENTRY SALVAGE REPORT
========================================================================
For every Gandhi Jayanti stock whose entry window opened BEFORE today (28-Sep),
decide whether a NEW ENTRY is still possible (enter next session 28-Sep = T-4,
hold to the original exit) or whether it should be AVOIDed.

NEW ENTRY possible  <=>  exit still ahead (>= today)  AND  the historical setup
                         is a winner (avg return > 0).
Otherwise           ->   AVOID.

OUTPUT
  Gandhi_MissedPositions_NewEntry_Report.html  (+ .pdf if Chrome is available)
ADDITIVE: reads dashboard_data/event_dashboard_data.json only.
"""
import os, json, re, subprocess
from datetime import date, timedelta

BASE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(BASE, 'dashboard_data', 'event_dashboard_data.json')
OUT_HTML = os.path.join(BASE, 'Gandhi_MissedPositions_NewEntry_Report.html')
OUT_PDF = os.path.join(BASE, 'Gandhi_MissedPositions_NewEntry_Report.pdf')

MONS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
DAYS = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun']
HOLS = {'2026-10-20','2026-10-02','2026-11-24','2026-12-25','2026-09-14','2026-01-26',
        '2026-03-03','2026-03-04','2026-03-26','2026-03-31','2026-04-03','2026-04-14',
        '2026-05-01','2026-05-28','2026-06-26'}
TODAY = date(2026, 9, 28)
ANCHOR = date(2026, 10, 2)      # Gandhi Jayanti
SALVAGE = TODAY                 # earliest catch-up session (Mon, trading day)


def trad(d): return d.weekday() < 5 and d.isoformat() not in HOLS
def pdmy(s):
    m = re.search(r'(\d{1,2})-([A-Za-z]{3})-(\d{4})', str(s))
    return date(int(m.group(3)), MONS.index(m.group(2)) + 1, int(m.group(1))) if m else None
def fmt(d): return '%02d-%s-%d (%s)' % (d.day, MONS[d.month - 1], d.year, DAYS[d.weekday()])
def sess(a, b):
    if b < a: return 0
    n = 0; d = a
    while d <= b:
        if trad(d): n += 1
        d += timedelta(days=1)
    return n
def t_off(d, anc):
    if d == anc: return 0
    step = 1 if d > anc else -1; c = 0; x = anc
    while x != d:
        x += timedelta(days=step)
        if trad(x): c += 1
    return c * step


def compute():
    dd = json.load(open(SRC, encoding='utf-8'))
    g = [h for h in dd['holidays'] if h['id'] == 'gandhi_2026'][0]
    missed = []
    for s in g['stocks']:
        en, xt = pdmy(s['entry_date']), pdmy(s['exit_date'])
        if not en or not xt or en >= TODAY:
            continue
        total = sess(en, xt); remain = sess(SALVAGE, xt)
        frac = remain / total if total else 0
        avg = s.get('full_19y_avg_ret') or 0
        salv = avg * frac
        catchable = xt >= TODAY
        possible = catchable and avg > 0
        missed.append({
            'sym': s['symbol'], 'dir': s['direction'], 'win': s['window'],
            'entry': en, 'exit': xt, 'wr': s.get('full_19y_wr'), 'avg': avg,
            'remain': remain, 'total': total, 'frac': frac, 'salv': salv,
            'possible': possible, 'texit': t_off(xt, ANCHOR),
        })
    missed.sort(key=lambda r: (not r['possible'], -r['salv']))
    return g, missed


def build_html(g, missed):
    n_poss = sum(1 for r in missed if r['possible'])
    n_avoid = len(missed) - n_poss

    def action_cell(r):
        if not r['possible']:
            reason = 'losing setup (avg %+.2f%%)' % r['avg'] if r['avg'] <= 0 else 'window closed'
            return '<td class="act avoid">🔴 AVOID<div class="sub">%s</div></td>' % reason
        side = 'BUY' if r['dir'] == 'LONG' else 'SELL'
        return ('<td class="act new"><b>%s @ 28-Sep</b> (T-4)<div class="sub">hold to %s (T%+d) · est <b>%+.2f%%</b></div></td>'
                % (side, r['exit'].strftime('%d-%b'), r['texit'], r['salv']))

    rows = ''
    for r in missed:
        dcls = 'lng' if r['dir'] == 'LONG' else 'sht'
        cls = 'avoidrow' if not r['possible'] else ('strong' if r['salv'] >= 1.5 else '')
        rows += ('<tr class="%s"><td class="l"><b>%s</b></td><td class="%s">%s</td>'
                 '<td>%s</td><td>%s</td><td>%s (T%+d)</td><td>%.1f%%</td>'
                 '<td class="%s">%+.2f%%</td><td>%d/%d <span class="mut">(%d%%)</span></td>%s</tr>') % (
            cls, r['sym'], dcls, r['dir'], r['win'], r['entry'].strftime('%d-%b'),
            r['exit'].strftime('%d-%b'), r['texit'], (r['wr'] or 0),
            'pos' if r['avg'] >= 0 else 'neg', r['avg'],
            r['remain'], r['total'], round(r['frac'] * 100), action_cell(r))

    tpl = """<!doctype html><html><head><meta charset="utf-8">
<title>Gandhi Jayanti Missed Positions — New Entry vs Avoid</title>
<style>
 *{box-sizing:border-box} body{font-family:'Segoe UI',Arial,sans-serif;margin:0;color:#0f172a;background:#fff}
 .wrap{max-width:1080px;margin:0 auto;padding:26px 34px}
 h1{font-size:22px;margin:0 0 2px} .sub{color:#64748b;font-size:12.5px}
 .banner{background:#0b1220;color:#e2e8f0;border-radius:10px;padding:14px 18px;margin:16px 0}
 .banner b{color:#fbbf24}
 .kpis{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin:14px 0}
 .k{border:1px solid #e2e8f0;border-radius:9px;padding:12px 14px}
 .kl{font-size:10.5px;text-transform:uppercase;letter-spacing:.04em;color:#64748b;font-weight:700}
 .kv{font-size:24px;font-weight:800}
 table{width:100%;border-collapse:collapse;font-size:12px;margin-top:6px}
 th,td{padding:6px 8px;text-align:right;border-bottom:1px solid #eef2f7}
 th{background:#f1f5f9;font-size:10px;text-transform:uppercase;letter-spacing:.03em;color:#475569;position:sticky;top:0}
 td.l,th.l{text-align:left} .mut{color:#94a3b8} .pos{color:#059669;font-weight:700} .neg{color:#dc2626;font-weight:700}
 .lng{color:#059669;font-weight:700} .sht{color:#dc2626;font-weight:700}
 td.act{text-align:left;font-size:11.5px} td.act.new{background:#ecfdf5} td.act.avoid{background:#fef2f2;color:#b91c1c;font-weight:700}
 td.act .sub{font-size:10.5px;color:#475569;font-weight:400}
 tr.strong td{background:#f0fdf4} tr.avoidrow td{opacity:.85}
 .note{background:#fff7ed;border:1px solid #fed7aa;border-radius:8px;padding:11px 15px;font-size:11.5px;color:#7c2d12;margin-top:14px;line-height:1.55}
 .foot{color:#94a3b8;font-size:10.5px;margin-top:18px}
 @page{size:A4 landscape;margin:12mm}
</style></head><body><div class="wrap">
 <h1>Mahatma Gandhi Jayanti (2-Oct-2026) — Missed Positions: New Entry vs Avoid</h1>
 <div class="sub">Nifty-50 holiday drift playbook &bull; as of <b>28-Sep-2026</b> &bull; anchor 02-Oct-2026 (Fri) &bull; catch-up session = 28-Sep (<b>T-4</b>)</div>

 <div class="banner">Positions whose entry opened before today were <b>missed</b>. Today is a trading day at <b>T-4</b>, so a <b>new entry</b> (enter 28-Sep, hold to the original exit) still captures the event-day &amp; post-event drift. A stock is marked <b>AVOID</b> only if its historical setup is a net loser (or the window has closed).</div>

 <div class="kpis">
  <div class="k"><div class="kl">Missed positions</div><div class="kv">@@NMISS@@</div></div>
  <div class="k"><div class="kl" style="color:#059669">New entry possible</div><div class="kv" style="color:#059669">@@NPOSS@@</div></div>
  <div class="k"><div class="kl" style="color:#dc2626">Avoid</div><div class="kv" style="color:#dc2626">@@NAVOID@@</div></div>
 </div>

 <table><thead><tr>
   <th class="l">Stock</th><th class="l">Dir</th><th>Orig Window</th><th>Missed Entry</th>
   <th>Exit</th><th>Hist WR</th><th>Hist Avg</th><th>Runway Left</th><th class="l">Recommended Action</th>
 </tr></thead><tbody>@@ROWS@@</tbody></table>

 <div class="note"><b>How to read it:</b> <b>Runway Left</b> = trading sessions from 28-Sep to the exit vs the full window. <b>Est %</b> in the action = historical avg return &times; runway fraction — a <b>conservative linear</b> figure. Holiday drift usually concentrates around T-1&ndash;T+2, which a T-4 entry today still captures, so short-hold names (T+1/T+2 exits) likely retain <b>more</b> than the linear estimate. Directions: LONG=BUY, SHORT=SELL the near-month future / 1% ITM option. Sizing is your call — a late entry has less runway, so scale accordingly.</div>
 <div class="foot">Additive analysis &bull; source: dashboard_data/event_dashboard_data.json &bull; historical stats are 2007&ndash;2025 window behaviour, not a forward guarantee.</div>
</div></body></html>"""
    repl = {'@@NMISS@@': str(len(missed)), '@@NPOSS@@': str(n_poss),
            '@@NAVOID@@': str(n_avoid), '@@ROWS@@': rows}
    for k, v in repl.items():
        tpl = tpl.replace(k, v)
    with open(OUT_HTML, 'w', encoding='utf-8') as f:
        f.write(tpl)


def to_pdf():
    for exe in [r'C:\Program Files\Google\Chrome\Application\chrome.exe',
                r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe']:
        if os.path.exists(exe):
            try:
                subprocess.run([exe, '--headless=new', '--disable-gpu', '--no-pdf-header-footer',
                                '--print-to-pdf=' + OUT_PDF, 'file:///' + OUT_HTML.replace('\\', '/')],
                               timeout=60, capture_output=True)
                return os.path.exists(OUT_PDF)
            except Exception as e:
                print('  chrome pdf failed:', e); return False
    print('  Chrome not found — HTML only.')
    return False


if __name__ == '__main__':
    g, missed = compute()
    build_html(g, missed)
    n_poss = sum(1 for r in missed if r['possible'])
    print('Missed: %d | New entry possible: %d | Avoid: %d' % (len(missed), n_poss, len(missed) - n_poss))
    print('wrote', os.path.basename(OUT_HTML))
    if to_pdf():
        print('wrote', os.path.basename(OUT_PDF))
