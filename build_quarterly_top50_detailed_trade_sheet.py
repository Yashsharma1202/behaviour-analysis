"""
TOP 50 F&O STOCKS \u2014 LAST QUARTER RESULTS (FY27_Q2, Jul-Sep 2026), DETAILED
TRADE SHEET. Uses the already-computed, fully-announced quarterly earnings-
drift trades (walk-forward 8x8 grid, trained on prior quarters, sourced from
an independent Yahoo Finance price fetch -- not the corrupted cash-equity
parquet file used elsewhere in this project).

ADDITIVE: reads event_dashboard_data.json; writes a new report.
"""
import json, os, subprocess, statistics as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

BASE = r'D:\behaviour analysis'
OUT_XLSX = os.environ.get('QTOP50_OUT_XLSX', BASE + r'\Quarterly_Top50_FO_Detailed_Trade_Sheet_FY27Q2.xlsx')
OUT_HTML = BASE + r'\Quarterly_Top50_FO_Detailed_Trade_Sheet_FY27Q2.html'
OUT_PDF = BASE + r'\Quarterly_Top50_FO_Detailed_Trade_Sheet_FY27Q2.pdf'

d = json.load(open(BASE + r'\dashboard_data\event_dashboard_data.json', encoding='utf-8'))
q2 = next(q for q in d['quarters'] if q['q_code'] == 'FY27_Q2')
stocks = q2['stocks']

ranked = sorted(stocks, key=lambda s: -s['actual_pnl'])

NAVY = PatternFill('solid', fgColor='1A202C'); WHITE = Font(bold=True, color='FFFFFF', size=9)
TITLE = Font(bold=True, color='1A202C', size=13); SUB = Font(color='4A5568', size=10, italic=True)
GF = PatternFill('solid', fgColor='E6F4EA'); RF = PatternFill('solid', fgColor='FCE8E6')
AM = PatternFill('solid', fgColor='FEF3C7'); LBL = PatternFill('solid', fgColor='F1F5F9')
GOOD = Font(color='137333', bold=True); BAD = Font(color='C5221F', bold=True); BOLD = Font(bold=True)
THIN = Border(*[Side(style='thin', color='D9E2EC')] * 4)
CEN = Alignment(horizontal='center'); RIGHT = Alignment(horizontal='right'); WRAP = Alignment(wrap_text=True, vertical='top')

wb = Workbook()

# ============ Sheet 1: Detailed Trade Log ============
tl = wb.active; tl.title = 'Trade_Log_FY27_Q2'
tl.merge_cells('A1:P1')
tl.cell(1, 1, 'TOP 50 F&O STOCKS \u2014 Q2 FY 2026-27 (Jul-Sep 2026) RESULTS \u2014 DETAILED TRADE LOG').font = TITLE
tl.merge_cells('A2:P2')
tl.cell(2, 1, 'Quarterly earnings-drift trades, walk-forward 8x8 grid search (trained on prior-quarter results, independent Yahoo Finance price source). Ranked by PnL, all 50 fully announced/closed.').font = SUB
hdr = ['Rank', 'Symbol', 'Company', 'Sector', 'Direction', 'Window', 'Result Date', 'Entry Date', 'Entry Price (\u20b9)',
       'Exit Date', 'Exit Price (\u20b9)', 'Lot', 'Margin (\u20b9)', 'Return %', 'PnL (\u20b9)', 'Outcome', 'Hist. WR % (prior qtrs)']
for j, h in enumerate(hdr, 1):
    c = tl.cell(4, j, h); c.fill = NAVY; c.font = WHITE; c.border = THIN; c.alignment = CEN
row = 5
for i, s in enumerate(ranked, 1):
    vals = [i, s['symbol'], s['name'], s.get('sector', ''), s['direction'], s['window'], s['result_date'],
            s['entry_date'], round(s['entry_px'], 2), s['exit_date'], round(s['exit_px'], 2), s['lot'],
            round(s['margin'], 2), s['actual_ret'], round(s['actual_pnl'], 2), s['outcome'], s['avg_17q_wr']]
    fill = GF if s['outcome'] == 'WIN' else RF
    for j, v in enumerate(vals, 1):
        c = tl.cell(row, j, v); c.border = THIN; c.fill = fill
        if j in (1, 5, 6, 16): c.alignment = CEN
        if j in (9, 11, 13, 14, 15, 17): c.alignment = RIGHT
        if j == 14: c.font = GOOD if v >= 0 else BAD
        if j == 15: c.font = GOOD if v >= 0 else BAD
    row += 1
for j, w in enumerate([6, 12, 24, 20, 10, 13, 13, 15, 13, 15, 13, 7, 13, 10, 13, 9, 14], 1):
    tl.column_dimensions[get_column_letter(j)].width = w
tl.freeze_panes = 'A5'

# ============ Sheet 2: Overview ============
ov = wb.create_sheet('Overview', 0)
ov.merge_cells('A1:B1')
ov.cell(1, 1, 'TOP 50 F&O STOCKS \u2014 Q2 FY 2026-27 RESULTS \u2014 OVERVIEW').font = TITLE
wins = sum(1 for s in stocks if s['outcome'] == 'WIN')
win_rate = round(100 * wins / len(stocks), 1)
total_pnl = round(sum(s['actual_pnl'] for s in stocks), 2)
total_margin = round(sum(s['margin'] for s in stocks), 2)
avg_ret = round(sum(s['actual_ret'] for s in stocks) / len(stocks), 2)
best = max(stocks, key=lambda s: s['actual_pnl'])
worst = min(stocks, key=lambda s: s['actual_pnl'])
n_long = sum(1 for s in stocks if s['direction'] == 'LONG'); n_short = len(stocks) - n_long
notes = [
    ('Universe', f'Top 50 F&O stocks tracked for quarterly-earnings-drift trading, Q2 FY 2026-27 (Jul-Sep 2026 results, announced Jul-Oct 2026) \u2014 all 50 fully announced and closed.'),
    ('Method', 'Each stock\u2019s own T-n/T+m window + direction chosen via a walk-forward 8x8 grid search trained on that stock\u2019s prior-quarter results (see Hist. WR % column \u2014 the win rate across those prior quarters, not this one).'),
    ('Data source', 'Prices sourced from an independent Yahoo Finance fetch (build_twelve_quarters_all_cases.py) \u2014 NOT the local cash-equity parquet file found to be missing 98.3% of Fridays across its 26yr history; this dataset does not carry that defect.'),
    ('This quarter\u2019s result', f'{wins}/50 wins ({win_rate}%). Average return {avg_ret:+.2f}%. Total PnL \u20b9{total_pnl:,.0f} on \u20b9{total_margin:,.0f} margin deployed (1 lot each) \u2014 return on margin {(total_pnl/total_margin*100 if total_margin else 0):+.2f}%.'),
    ('Direction split', f'{n_long} LONG / {n_short} SHORT.'),
    ('Best / Worst', f'Best: {best["symbol"]} ({best["actual_pnl"]:+,.0f} \u20b9, {best["direction"]}). Worst: {worst["symbol"]} ({worst["actual_pnl"]:+,.0f} \u20b9, {worst["direction"]}).'),
    ('Margin', '20% SPAN proxy on entry price \u00d7 lot size \u2014 a planning estimate, not an exact historical broker margin.'),
    ('Hist. WR % column', 'This is each stock\u2019s own win rate across PRIOR quarters (used to pick this quarter\u2019s direction) \u2014 not this quarter\u2019s own outcome. Compare it to the Outcome column to see where the historical edge held up and where it didn\u2019t.'),
]
r0 = 3
for lbl, txt in notes:
    a = ov.cell(r0, 1, lbl); a.font = BOLD; a.fill = AM; a.border = THIN; a.alignment = WRAP
    b = ov.cell(r0, 2, txt); b.alignment = WRAP; b.border = THIN
    ov.row_dimensions[r0].height = 60; r0 += 1
ov.column_dimensions['A'].width = 22; ov.column_dimensions['B'].width = 108

wb.save(OUT_XLSX)
print('Saved Excel:', OUT_XLSX)
print(f"Win rate {win_rate}% ({wins}/50) | Total PnL Rs{total_pnl:,.0f} | Margin Rs{total_margin:,.0f} | {n_long} LONG / {n_short} SHORT")

# ============ PDF ============
def esc(s): return str(s).replace('&', '&amp;').replace('<', '&lt;')
def pdf_row(rank, s):
    win = s['outcome'] == 'WIN'
    rowstyle = '' if win else ' style="background:#fef2f2;"'
    return (f'<tr{rowstyle}><td>{rank}</td><td><b>{esc(s["symbol"])}</b></td><td>{esc(s["name"])[:22]}</td>'
            f'<td class="c">{s["direction"]}</td><td class="c">{esc(s["window"])}</td>'
            f'<td class="r {"pos" if s["actual_ret"]>=0 else "neg"}">{s["actual_ret"]:+.2f}%</td>'
            f'<td class="r {"pos" if s["actual_pnl"]>=0 else "neg"}">{s["actual_pnl"]:+,.0f}</td>'
            f'<td class="c">{s["outcome"]}</td><td class="r">{s["avg_17q_wr"]:.0f}%</td></tr>')
rows_all = ''.join(pdf_row(i + 1, s) for i, s in enumerate(ranked))

html = f"""<!doctype html><html><head><meta charset="utf-8">
<title>Top 50 F&amp;O Q2 FY27 Results</title>
<style>
 *{{box-sizing:border-box}} body{{font-family:'Segoe UI',Arial,sans-serif;margin:0;color:#0f172a;background:#fff}}
 .wrap{{max-width:1050px;margin:0 auto;padding:28px 36px}}
 h1{{font-size:22px;margin:0 0 3px}} .sub{{color:#64748b;font-size:12px;margin-bottom:16px}}
 h2{{font-size:15px;margin:20px 0 8px;border-left:4px solid #2563eb;padding-left:9px}}
 table{{width:100%;border-collapse:collapse;font-size:11px}}
 th,td{{padding:5px 7px;text-align:right;border-bottom:1px solid #eef2f7}} th{{background:#f1f5f9;font-size:9.5px;text-transform:uppercase;color:#475569}}
 td.c,th.c{{text-align:center}} .pos{{color:#059669;font-weight:700}} .neg{{color:#dc2626;font-weight:700}}
 .cards{{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin:10px 0}}
 .mc{{background:#fff;border:1px solid #e2e8f0;border-radius:8px;padding:10px 12px}}
 .mcl{{font-size:9.5px;text-transform:uppercase;color:#64748b;font-weight:700}} .mcv{{font-size:19px;font-weight:800}}
 .note{{background:#fff7ed;border:1px solid #fed7aa;border-radius:8px;padding:11px 14px;font-size:11px;color:#7c2d12;margin-top:10px;line-height:1.5}}
 @page{{size:A4;margin:12mm}}
</style></head><body><div class="wrap">
 <h1>Top 50 F&amp;O Stocks &mdash; Q2 FY 2026-27 Results, Detailed Trade Log</h1>
 <div class="sub">Jul-Sep 2026 quarterly earnings &bull; walk-forward 8x8 grid, trained on prior quarters &bull; all 50 fully announced and closed</div>
 <div class="cards">
  <div class="mc"><div class="mcl">Win Rate</div><div class="mcv pos">{win_rate}%</div></div>
  <div class="mc"><div class="mcl">Total PnL</div><div class="mcv {"pos" if total_pnl>=0 else "neg"}">\u20b9{total_pnl:,.0f}</div></div>
  <div class="mc"><div class="mcl">Margin Deployed</div><div class="mcv">\u20b9{total_margin:,.0f}</div></div>
  <div class="mc"><div class="mcl">LONG / SHORT</div><div class="mcv">{n_long} / {n_short}</div></div>
 </div>
 <h2>All 50 \u2014 Ranked by PnL</h2>
 <table><thead><tr><th>Rank</th><th class="c">Symbol</th><th class="c">Company</th><th class="c">Dir</th><th class="c">Window</th><th>Return</th><th>PnL (\u20b9)</th><th class="c">Outcome</th><th>Hist. WR</th></tr></thead>
 <tbody>{rows_all}</tbody></table>
 <div class="note">Hist. WR% is each stock's own win rate across PRIOR quarters (what picked this quarter's direction), not this quarter's own result \u2014 compare it to Outcome to see where it held. Data sourced from an independent Yahoo Finance fetch, not the local cash-equity file found missing 98.3% of Fridays across its history.</div>
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
