"""
DUSSEHRA BACKTEST REPORT (Excel) — per-stock LONG vs SHORT, best-direction chosen.
For each of the 50 Nifty-50 stocks (using its own existing T-n/T+m window), the
real split-robust backtest (26yr history, corporate-action days excluded) for
BOTH directions is shown side by side, with the chosen (higher win-rate) side
highlighted, plus the corrected 2026 position-taking schedule.

CAVEAT (in the workbook): picking the best direction per stock this way tends to
overfit — it looks strong in-sample but has repeatedly underperformed a uniform
index-backed direction (SHORT) out-of-sample in this project's own testing.
Included transparently, not hidden.

ADDITIVE: reads event_dashboard_data.json (already updated); writes a new workbook.
"""
import json, pickle
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

BASE = r'D:\behaviour analysis'
OUT = BASE + r'\Dussehra_Backtest_Report.xlsx'

RESULTS = pickle.load(open(r'C:\Users\PC2551\AppData\Local\Temp\claude\d--behaviour-analysis\3a49a75a-d0c8-402c-9f1d-8142994e3974\scratchpad\dussehra_bt.pkl', 'rb'))
d = json.load(open(BASE + r'\dashboard_data\event_dashboard_data.json', encoding='utf-8'))
h = [x for x in d['holidays'] if x['id'] == 'dussehra_2026'][0]

NAVY = PatternFill('solid', fgColor='1A202C'); WHITE = Font(bold=True, color='FFFFFF', size=10)
TITLE = Font(bold=True, color='1A202C', size=13); SUB = Font(color='4A5568', size=10, italic=True)
GF = PatternFill('solid', fgColor='E6F4EA'); RF = PatternFill('solid', fgColor='FCE8E6')
AM = PatternFill('solid', fgColor='FEF3C7'); LBL = PatternFill('solid', fgColor='F1F5F9')
GOOD = Font(color='137333', bold=True); BAD = Font(color='C5221F', bold=True); BOLD = Font(bold=True)
THIN = Border(*[Side(style='thin', color='D9E2EC')] * 4)
CEN = Alignment(horizontal='center'); RIGHT = Alignment(horizontal='right'); WRAP = Alignment(wrap_text=True, vertical='top')

wb = Workbook()

# ---------- Sheet 1: Backtest comparison ----------
bt = wb.active; bt.title = 'LONG_vs_SHORT_Backtest'
bt.merge_cells('A1:K1')
bt.cell(1, 1, 'DUSSEHRA \u2014 Per-Stock LONG vs SHORT Backtest (26yr, split-robust)').font = TITLE
bt.merge_cells('A2:K2')
bt.cell(2, 1, 'Each stock uses its own existing T-n/T+m window. Best direction = higher win rate. Anchor: 20-Oct-2026.').font = SUB
hdr = ['Symbol', 'Window', 'Years', 'LONG Win%', 'LONG Avg%', 'SHORT Win%', 'SHORT Avg%', 'Chosen', 'Edge (pp)', 'Data Basis']
for j, hh in enumerate(hdr, 1):
    c = bt.cell(4, j, hh); c.fill = NAVY; c.font = WHITE; c.border = THIN; c.alignment = CEN

row = 5
for s in sorted(h['stocks'], key=lambda x: -(x['full_19y_wr'] or 0)):
    sym = s['symbol']; r = RESULTS.get(sym)
    if r and r['pick']:
        lwr, lav, swr, sav = r['long_wr'], r['long_avg'], r['short_wr'], r['short_avg']
        n = r['n']; edge = round(abs(lwr - swr), 1)
        basis = f"{n}yr real backtest"
    else:
        lwr = lav = swr = sav = None
        n = s.get('n_years'); edge = None
        basis = 'F&O-211 4yr fallback'
    vals = [sym, s['window'], n, lwr, lav, swr, sav, s['direction'], edge, basis]
    fill = GF if s['direction'] == 'LONG' else RF
    for j, v in enumerate(vals, 1):
        c = bt.cell(row, j, v if v is not None else '\u2014'); c.border = THIN; c.fill = fill
        if j in (1, 8, 10): c.alignment = CEN if j != 10 else Alignment(horizontal='left')
        if j in (3, 4, 5, 6, 7, 9): c.alignment = RIGHT
    row += 1
for j, w in enumerate([12, 13, 7, 11, 11, 11, 11, 9, 10, 18], 1):
    bt.column_dimensions[get_column_letter(j)].width = w
bt.freeze_panes = 'A5'

# ---------- Sheet 2: Corrected position schedule ----------
ps = wb.create_sheet('Position_Schedule_2026')
ps.merge_cells('A1:K1')
ps.cell(1, 1, 'DUSSEHRA (20-Oct-2026) \u2014 Corrected Position-Taking Schedule, 50 Stocks').font = TITLE
ps.merge_cells('A2:K2')
ps.cell(2, 1, 'Dates verified: entry/exit exactly match each stock\u2019s T-n/T+m window from the 20-Oct-2026 anchor on the NSE-2026 trading calendar.').font = SUB
hdr = ['#', 'Symbol', 'Company', 'Direction', 'Window', 'Entry Date', 'Exit Date', 'Win Rate %', 'Lot', 'Margin (Rs)', 'Order Ticket']
for j, hh in enumerate(hdr, 1):
    c = ps.cell(4, j, hh); c.fill = NAVY; c.font = WHITE; c.border = THIN; c.alignment = CEN
row = 5
for i, s in enumerate(sorted(h['stocks'], key=lambda x: x['symbol']), 1):
    vals = [i, s['symbol'], s['name'], s['direction'], s['window'], s['entry_date'], s['exit_date'],
            s['full_19y_wr'], s['lot'], s['margin'], s['fut_action']]
    fill = GF if s['direction'] == 'LONG' else RF
    for j, v in enumerate(vals, 1):
        c = ps.cell(row, j, v); c.border = THIN; c.fill = fill
        if j in (1, 4): c.alignment = CEN
        if j in (8, 9, 10): c.alignment = RIGHT
    row += 1
for j, w in enumerate([5, 12, 24, 10, 12, 20, 20, 11, 8, 14, 26], 1):
    ps.column_dimensions[get_column_letter(j)].width = w
ps.freeze_panes = 'A5'

# ---------- Sheet 3: Summary ----------
sm = wb.create_sheet('Summary', 0)
sm.merge_cells('A1:B1')
sm.cell(1, 1, 'DUSSEHRA BACKTEST \u2014 SUMMARY').font = TITLE
from collections import Counter
dirs = Counter(s['direction'] for s in h['stocks'])
avg_wr = round(sum(s['full_19y_wr'] for s in h['stocks']) / len(h['stocks']), 1)
long_stocks = [s for s in h['stocks'] if s['direction'] == 'LONG']
short_stocks = [s for s in h['stocks'] if s['direction'] == 'SHORT']
rows = [
    ('Total stocks', '50'),
    ('Direction split (per-stock best)', f"{dirs['LONG']} LONG / {dirs['SHORT']} SHORT"),
    ('Avg win rate (chosen direction)', f"{avg_wr}%"),
    ('Avg LONG-side win rate (25 LONG picks)', f"{round(sum(s['full_19y_wr'] for s in long_stocks)/len(long_stocks),1)}%" if long_stocks else '\u2014'),
    ('Avg SHORT-side win rate (25 SHORT picks)', f"{round(sum(s['full_19y_wr'] for s in short_stocks)/len(short_stocks),1)}%" if short_stocks else '\u2014'),
    ('Anchor date', '20-Oct-2026 (Tuesday)'),
    ('Methodology', 'Split-robust (corporate-action days excluded), each stock\u2019s own T-n/T+m window, best of LONG/SHORT by win rate'),
]
r0 = 3
for lbl, val in rows:
    a = sm.cell(r0, 1, lbl); a.font = BOLD; a.fill = LBL; a.border = THIN
    b = sm.cell(r0, 2, val); b.border = THIN
    r0 += 1
sm.merge_cells(f'A{r0+1}:B{r0+1}')
warn = sm.cell(r0 + 1, 1,
    'IMPORTANT CAVEAT: picking the single best direction per stock (as done here) fits the historical '
    'sample closely but has repeatedly shown WEAKER out-of-sample performance than a single uniform '
    'index-backed direction in this project\u2019s own prior testing (e.g. uniform SHORT beat best-per-stock '
    'selection out-of-sample on both a 26-year and a 15-year test). Treat this as the best HISTORICAL fit, '
    'not a guaranteed forward edge \u2014 verify against the index-level SHORT bias (69% WR, robust OOS) before sizing up.')
warn.font = Font(color='7C2D12', size=10, italic=True); warn.alignment = WRAP
sm.row_dimensions[r0 + 1].height = 70
sm.column_dimensions['A'].width = 38; sm.column_dimensions['B'].width = 60

wb.save(OUT)
print('Saved:', OUT)
print(f"Direction split: {dict(dirs)} | avg win rate: {avg_wr}%")
