"""
Q3 FY 2026-27 (Oct-Dec 2026 Results) \u2014 ANNOUNCED POSITIONS Excel.
Every stock with a live, NSE-confirmed result date for the current quarter
(not "Yet to come"), with its position-taking window, entry/exit date+price
basis, lot, margin and historical win rate \u2014 ready to act on.

ADDITIVE: reads event_dashboard_data.json; writes a new report.
"""
import json, os, subprocess
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

BASE = r'D:\behaviour analysis'
OUT_XLSX = os.environ.get('Q3ANN_OUT_XLSX', BASE + r'\Q3_FY2026-27_Announced_Positions.xlsx')
OUT_HTML = BASE + r'\Q3_FY2026-27_Announced_Positions.html'
OUT_PDF = BASE + r'\Q3_FY2026-27_Announced_Positions.pdf'

d = json.load(open(BASE + r'\dashboard_data\event_dashboard_data.json', encoding='utf-8'))
q3 = next(q for q in d['quarters'] if q['q_code'] == 'FY27_Q3')
announced = [s for s in q3['stocks'] if s.get('announced')]
pending_n = len(q3['stocks']) - len(announced)

# ---- Last 12 completed quarters (excludes the current, still-pending FY27_Q3) ----
LAST12_QCODES = ['FY27_Q2', 'FY27_Q1', 'FY26_Q4', 'FY26_Q3', 'FY26_Q2', 'FY26_Q1',
                  'FY25_Q4', 'FY25_Q3', 'FY25_Q2', 'FY25_Q1', 'FY24_Q4', 'FY24_Q3']
qname_by_code = {q['q_code']: q.get('q_name', q['q_code']) for q in d['quarters']}
by_qcode_stock = {q['q_code']: {s['symbol']: s for s in q['stocks']} for q in d['quarters'] if q['q_code'] in LAST12_QCODES}

trades_12q = []  # one row per (symbol, quarter), oldest to newest
for s in announced:
    sym = s['symbol']
    for qc in reversed(LAST12_QCODES):  # oldest -> newest for a clean running-PnL read
        rec = by_qcode_stock[qc].get(sym)
        if not rec:
            continue
        trades_12q.append(dict(symbol=sym, qcode=qc, qname=qname_by_code[qc], direction=rec['direction'],
                                entry_date=rec['entry_date'], entry_px=rec['entry_px'], exit_date=rec['exit_date'],
                                exit_px=rec['exit_px'], ret_pct=rec['actual_ret'], pnl=rec['actual_pnl'],
                                outcome=rec['outcome'], lot=rec['lot'], margin=rec['margin']))

def parse_dmy(s):
    return datetime.strptime(s.split(' (')[0], '%d-%b-%Y')

announced.sort(key=lambda s: parse_dmy(s['entry_date']))

NAVY = PatternFill('solid', fgColor='1A202C'); WHITE = Font(bold=True, color='FFFFFF', size=9)
TITLE = Font(bold=True, color='1A202C', size=13); SUB = Font(color='4A5568', size=10, italic=True)
GF = PatternFill('solid', fgColor='E6F4EA'); RF = PatternFill('solid', fgColor='FCE8E6')
LBL = PatternFill('solid', fgColor='F1F5F9')
GOOD = Font(color='137333', bold=True); BAD = Font(color='C5221F', bold=True); BOLD = Font(bold=True)
THIN = Border(*[Side(style='thin', color='D9E2EC')] * 4)
CEN = Alignment(horizontal='center'); RIGHT = Alignment(horizontal='right'); WRAP = Alignment(wrap_text=True, vertical='top')

wb = Workbook()

# ============ Sheet 1: Position Schedule ============
ps = wb.active; ps.title = 'Announced_Positions'
ps.merge_cells('A1:M1')
ps.cell(1, 1, 'Q3 FY 2026-27 (Oct-Dec 2026 Results) \u2014 ANNOUNCED POSITIONS').font = TITLE
ps.merge_cells('A2:M2')
ps.cell(2, 1, f'{len(announced)} of {len(q3["stocks"])} stocks have an NSE-confirmed result date so far ({pending_n} still "Yet to come"). Sorted by entry date. Entry 09:20 AM, Exit 03:15 PM.').font = SUB
hdr = ['#', 'Symbol', 'Company', 'Sector', 'Direction', 'Window', 'Result Date', 'Entry Date', 'Entry Price (\u20b9)',
       'Exit Date', 'Lot', 'Margin (\u20b9)', 'Hist. Win Rate %']
for j, h in enumerate(hdr, 1):
    c = ps.cell(4, j, h); c.fill = NAVY; c.font = WHITE; c.border = THIN; c.alignment = CEN
row = 5
for i, s in enumerate(announced, 1):
    vals = [i, s['symbol'], s['name'], s.get('sector', ''), s['direction'], s['window'], s['result_date'],
            s['entry_date'], s.get('entry_px', s.get('spot_ltp')), s['exit_date'], s['lot'], s['margin'], s.get('avg_17q_wr')]
    fill = GF if s['direction'] == 'LONG' else RF
    for j, v in enumerate(vals, 1):
        c = ps.cell(row, j, v); c.border = THIN; c.fill = fill
        if j in (1, 5, 6, 7): c.alignment = CEN
        if j in (9, 11, 12, 13): c.alignment = RIGHT
    row += 1
for j, w in enumerate([5, 12, 26, 20, 10, 13, 13, 18, 13, 18, 7, 13, 13], 1):
    ps.column_dimensions[get_column_letter(j)].width = w
ps.freeze_panes = 'A5'

# ============ Sheet 2: Overview ============
ov = wb.create_sheet('Overview', 0)
ov.merge_cells('A1:B1')
ov.cell(1, 1, 'Q3 FY 2026-27 \u2014 ANNOUNCED POSITIONS \u2014 OVERVIEW').font = TITLE
n_long = sum(1 for s in announced if s['direction'] == 'LONG'); n_short = len(announced) - n_long
total_margin = round(sum(s['margin'] for s in announced), 2)
avg_wr = round(sum(s.get('avg_17q_wr', 0) for s in announced) / len(announced), 1) if announced else 0
notes = [
    ('Universe', f'All {len(q3["stocks"])} stocks tracked for Q3 FY 2026-27 (Oct-Dec 2026 results). {len(announced)} have a real NSE board-meeting intimation filed so far; {pending_n} are still "Yet to come" and excluded from this report \u2014 re-run once more results are announced.'),
    ('Window & method', 'Each stock\u2019s own T-2 to T+4 window around its announced result date, direction from its historical quarterly win rate (Hist. Win Rate % column, prior-quarters track record).'),
    ('Direction split', f'{n_long} LONG / {n_short} SHORT.'),
    ('Margin', f'Total \u20b9{total_margin:,.0f} to deploy across all {len(announced)} positions (1 lot each), 20% SPAN proxy per stock.'),
    ('Average historical win rate', f'{avg_wr}% across these {len(announced)} stocks\u2019 own prior-quarter track record \u2014 this is NOT this quarter\u2019s outcome (still pending), it\u2019s what picked the direction.'),
    ('Source', 'Result dates pulled live from NSE\u2019s corporate board-meeting feed (verified against NSE\u2019s own API). Refresh the dashboard or re-run this report as more results get announced through the quarter.'),
    ('One sheet per quarter (FY24_Q3 \u2026 FY27_Q2)', 'Each of the 12 tabs shows all 15 stocks\u2019 real, already-settled trade for that single quarter \u2014 Entry/Exit Date+Price, Return %, "DD Booked This Qtr" (that quarter\u2019s own loss, if any), and "Running DD to Date" (the cumulative drawdown from the running equity-curve peak at that point, walking chronologically from FY24_Q3).'),
    ('Last_12Q_Performance sheet', 'One row per stock, aggregating all 12 quarters: Win Rate, Avg Return, CAGR, Max Profit % (best single quarter), Max Loss % (worst single quarter), Max DD % (worst peak-to-trough on the running equity curve), Sharpe-like, Total PnL.'),
]
r0 = 3
for lbl, txt in notes:
    a = ov.cell(r0, 1, lbl); a.font = BOLD; a.fill = LBL; a.border = THIN; a.alignment = WRAP
    b = ov.cell(r0, 2, txt); b.alignment = WRAP; b.border = THIN
    ov.row_dimensions[r0].height = 60; r0 += 1
ov.column_dimensions['A'].width = 22; ov.column_dimensions['B'].width = 105

# ============ Sheet 3: Last 12 Quarters Trade Log ============
import statistics as st

# ---- Precompute each stock's chronological equity curve + running DD, so
# every per-quarter sheet can show "DD booked this quarter" AND "running DD
# to date" without recomputing per sheet. ----
equity_by_sym = {}   # sym -> {qcode: dict(ret, dd_this_qtr, running_dd_after, eq_after)}
perf_rows = []
for s in announced:
    sym = s['symbol']
    trs = sorted([t for t in trades_12q if t['symbol'] == sym], key=lambda x: LAST12_QCODES.index(x['qcode']), reverse=True)  # oldest->newest
    rets = [t['ret_pct'] for t in trs]
    n = len(rets)
    if n == 0:
        continue
    wins = sum(1 for v in rets if v > 0)
    wr = round(100 * wins / n, 1)
    avg = round(sum(rets) / n, 2)
    max_profit = round(max(rets), 2)
    max_loss = round(min(rets), 2)
    cum = 1.0
    for v in rets: cum *= (1 + v / 100)
    cagr = round((cum ** (4 / n) - 1) * 100, 2) if cum > 0 else None  # 4 quarters/year
    eq = [100.0]
    peak = 100.0
    qmap = {}
    for t, v in zip(trs, rets):
        new_eq = eq[-1] * (1 + v / 100)
        peak = max(peak, new_eq)
        running_dd = (new_eq - peak) / peak * 100
        qmap[t['qcode']] = dict(ret=v, running_dd=round(running_dd, 2))
        eq.append(new_eq)
    equity_by_sym[sym] = qmap
    maxdd = min(q['running_dd'] for q in qmap.values())
    sharpe = round(st.mean(rets) / st.pstdev(rets), 2) if n > 1 and st.pstdev(rets) > 0 else None
    total_pnl = round(sum(t['pnl'] for t in trs), 2)
    perf_rows.append((sym, n, wr, avg, cagr, max_profit, max_loss, round(maxdd, 2), sharpe, total_pnl, s['direction']))

# ============ Sheet 3..14: one sheet PER QUARTER (all 15 stocks) ============
for qc in LAST12_QCODES:
    safe_name = qc.replace('FY', 'FY')  # sheet name, Excel-safe (<=31 chars, already is)
    qs = wb.create_sheet(safe_name)
    qs.merge_cells('A1:L1')
    qs.cell(1, 1, f'{qname_by_code[qc]} — REAL TRADE RESULT, EACH ANNOUNCED-THIS-QUARTER STOCK').font = TITLE
    qs.merge_cells('A2:L2')
    qs.cell(2, 1, 'Already-settled quarterly-earnings trade. "DD Booked This Qtr" = this quarter’s own return if negative. "Running DD to Date" = cumulative drawdown from the running equity-curve peak after this quarter, walking chronologically from FY24_Q3.').font = SUB
    hdr = ['Symbol', 'Direction', 'Entry Date', 'Entry Price (₹)', 'Exit Date', 'Exit Price (₹)', 'Lot', 'Margin (₹)',
           'Return %', 'DD Booked This Qtr %', 'Running DD to Date %', 'Outcome']
    for j, h in enumerate(hdr, 1):
        c = qs.cell(4, j, h); c.fill = NAVY; c.font = WHITE; c.border = THIN; c.alignment = CEN
    row = 5
    q_trades = sorted([t for t in trades_12q if t['qcode'] == qc], key=lambda x: x['symbol'])
    for t in q_trades:
        sym = t['symbol']
        dd_info = equity_by_sym.get(sym, {}).get(qc, {})
        dd_this_qtr = round(min(0.0, t['ret_pct']), 2)
        running_dd = dd_info.get('running_dd', '—')
        vals = [sym, t['direction'], t['entry_date'], round(t['entry_px'], 2), t['exit_date'], round(t['exit_px'], 2),
                t['lot'], round(t['margin'], 2), t['ret_pct'], dd_this_qtr, running_dd, t['outcome']]
        fill = GF if t['outcome'] == 'WIN' else RF
        for j, v in enumerate(vals, 1):
            c = qs.cell(row, j, v); c.border = THIN; c.fill = fill
            if j in (1, 2, 12): c.alignment = CEN
            if j in (4, 6, 8, 9, 10, 11): c.alignment = RIGHT
            if j == 9: c.font = GOOD if t['ret_pct'] >= 0 else BAD
            if j in (10, 11) and isinstance(v, (int, float)) and v < 0: c.font = BAD
        row += 1
    for j, w in enumerate([12, 10, 20, 13, 20, 13, 7, 13, 10, 16, 17, 9], 1):
        qs.column_dimensions[get_column_letter(j)].width = w
    qs.freeze_panes = 'A5'

# ============ Sheet: Last 12 Quarters Performance Measures (summary, all quarters combined) ============
pm12 = wb.create_sheet('Last_12Q_Performance')
pm12.merge_cells('A1:K1')
pm12.cell(1, 1, 'LAST 12 QUARTERS — PERFORMANCE MEASURES PER STOCK (across all 12, see the per-quarter sheets for the trade-by-trade detail)').font = TITLE
pm12.merge_cells('A2:K2')
pm12.cell(2, 1, 'CAGR and Max DD computed on each stock’s own quarterly-trade equity curve (one trade/quarter, starting at 100, walked chronologically oldest to newest).').font = SUB
hdr = ['Symbol', 'Quarters', 'Win Rate %', 'Avg Return %', 'CAGR %', 'Max Profit % (best qtr)', 'Max Loss % (worst qtr)',
       'Max DD %', 'Sharpe-like', 'Total PnL (₹)', 'This Qtr Direction']
for j, h in enumerate(hdr, 1):
    c = pm12.cell(4, j, h); c.fill = NAVY; c.font = WHITE; c.border = THIN; c.alignment = CEN
row = 5
for r in perf_rows:
    sym, n, wr, avg, cagr, max_profit, max_loss, maxdd, sharpe, total_pnl, this_dir = r
    vals = [sym, n, wr, avg, (cagr if cagr is not None else '—'), max_profit, max_loss, maxdd,
            (sharpe if sharpe is not None else '—'), total_pnl, this_dir]
    fill = GF if total_pnl >= 0 else RF
    for j, v in enumerate(vals, 1):
        c = pm12.cell(row, j, v); c.border = THIN; c.fill = fill
        if j in (2, 11): c.alignment = CEN
        if j in (3, 4, 5, 6, 7, 8, 9, 10): c.alignment = RIGHT
        if j == 4: c.font = GOOD if avg >= 0 else BAD
        if j == 6: c.font = GOOD
        if j == 7: c.font = BAD
        if j == 8: c.font = BAD if maxdd < 0 else GOOD
        if j == 10: c.font = GOOD if total_pnl >= 0 else BAD
    row += 1
for j, w in enumerate([12, 10, 11, 13, 10, 15, 15, 10, 11, 14, 16], 1):
    pm12.column_dimensions[get_column_letter(j)].width = w
pm12.freeze_panes = 'A5'

wb.save(OUT_XLSX)
print('Saved Excel:', OUT_XLSX)
print(f"{len(announced)} announced ({n_long} LONG / {n_short} SHORT) | Margin Rs{total_margin:,.0f} | {pending_n} still pending")

# ============ PDF ============
def esc(s): return str(s).replace('&', '&amp;').replace('<', '&lt;')
def pdf_row(i, s):
    return (f'<tr><td>{i}</td><td><b>{esc(s["symbol"])}</b></td><td>{esc(s["name"])[:22]}</td>'
            f'<td class="c">{s["direction"]}</td><td class="c">{esc(s["window"])}</td>'
            f'<td class="c">{s["result_date"]}</td><td class="c">{esc(s["entry_date"])}</td>'
            f'<td class="c">{esc(s["exit_date"])}</td><td class="r">{s.get("avg_17q_wr","-")}%</td></tr>')
rows_all = ''.join(pdf_row(i + 1, s) for i, s in enumerate(announced))

html = f"""<!doctype html><html><head><meta charset="utf-8">
<title>Q3 FY27 Announced Positions</title>
<style>
 *{{box-sizing:border-box}} body{{font-family:'Segoe UI',Arial,sans-serif;margin:0;color:#0f172a;background:#fff}}
 .wrap{{max-width:1000px;margin:0 auto;padding:28px 36px}}
 h1{{font-size:21px;margin:0 0 3px}} .sub{{color:#64748b;font-size:12px;margin-bottom:16px}}
 h2{{font-size:15px;margin:20px 0 8px;border-left:4px solid #2563eb;padding-left:9px}}
 table{{width:100%;border-collapse:collapse;font-size:11px}}
 th,td{{padding:5px 7px;text-align:right;border-bottom:1px solid #eef2f7}} th{{background:#f1f5f9;font-size:9.5px;text-transform:uppercase;color:#475569}}
 td.c,th.c{{text-align:center}}
 .cards{{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin:10px 0}}
 .mc{{background:#fff;border:1px solid #e2e8f0;border-radius:8px;padding:10px 12px}}
 .mcl{{font-size:9.5px;text-transform:uppercase;color:#64748b;font-weight:700}} .mcv{{font-size:19px;font-weight:800}}
 @page{{size:A4;margin:12mm}}
</style></head><body><div class="wrap">
 <h1>Q3 FY 2026-27 &mdash; Announced Positions</h1>
 <div class="sub">{len(announced)} of {len(q3['stocks'])} stocks have a confirmed NSE result date &bull; sorted by entry date &bull; entry 09:20 AM, exit 03:15 PM</div>
 <div class="cards">
  <div class="mc"><div class="mcl">Announced</div><div class="mcv">{len(announced)}</div></div>
  <div class="mc"><div class="mcl">Still Pending</div><div class="mcv">{pending_n}</div></div>
  <div class="mc"><div class="mcl">LONG / SHORT</div><div class="mcv">{n_long} / {n_short}</div></div>
  <div class="mc"><div class="mcl">Margin Needed</div><div class="mcv">\u20b9{total_margin:,.0f}</div></div>
 </div>
 <h2>Position Schedule</h2>
 <table><thead><tr><th>#</th><th class="c">Symbol</th><th class="c">Company</th><th class="c">Dir</th><th class="c">Window</th><th class="c">Result</th><th class="c">Entry</th><th class="c">Exit</th><th>Hist. WR</th></tr></thead>
 <tbody>{rows_all}</tbody></table>
</div></body></html>"""
open(OUT_HTML, 'w', encoding='utf-8').write(html)
print('Saved HTML:', OUT_HTML)

for exe in [r'C:\Program Files\Google\Chrome\Application\chrome.exe', r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe']:
    if os.path.exists(exe):
        try:
            subprocess.run([exe, '--headless=new', '--disable-gpu', '--no-pdf-header-footer',
                            '--print-to-pdf=' + OUT_PDF, 'file:///' + OUT_HTML.replace('\\', '/')], timeout=60, capture_output=True)
            print('Saved PDF:', OUT_PDF, 'exists:', os.path.exists(OUT_PDF))
        except Exception as e:
            print('PDF render failed:', e)
        break
