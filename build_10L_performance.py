"""
Rs 10 LAKH / QUARTER PERFORMANCE — 1% ITM OPTIONS vs FUTURES (12 quarters).
Deploy a FRESH Rs 10,00,000 book each quarter (additive, NOT compounded).
Return% = quarter net P&L / capital deployed (premium for options, SPAN margin for
futures); Rs on 10L = Return% x 10,00,000. Full performance measures per vehicle.

FIXES applied vs the raw calc:
 - no bogus 'compounded' number (10L/qtr is fresh capital, so P&L is additive);
 - a prominent Assumptions & Reality-Check sheet (1% synthetic model, 3-yr in-sample
   bull window where every quarter won -> optimistic; not a forward guarantee).
Sources: the two 12-quarter masters. ADDITIVE: writes a new workbook.
"""
import openpyxl, statistics as st, math
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

BASE = r'D:\behaviour analysis'
OUT = BASE + r'\GitHub_Dashboard_10L_Quarter_Performance.xlsx'
CAP = 1_000_000  # Rs 10 lakh
ORDER = ['FY24_Q3', 'FY24_Q4', 'FY25_Q1', 'FY25_Q2', 'FY25_Q3', 'FY25_Q4',
         'FY26_Q1', 'FY26_Q2', 'FY26_Q3', 'FY26_Q4', 'FY27_Q1', 'FY27_Q2']


def load(path, pnl, cap):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True); out = {}
    for sn in wb.sheetnames:
        if sn == 'Dashboard_Summary':
            continue
        rows = list(wb[sn].iter_rows(values_only=True)); hdr = rows[3]
        idx = {str(h).strip(): j for j, h in enumerate(hdr) if h}
        def C(n):
            for k, j in idx.items():
                if k.lower().startswith(n.lower()):
                    return j
        js, jp, jc = C('Symbol'), C(pnl), C(cap)
        tp = tc = 0.0
        for r in rows[4:]:
            if r and r[js] is not None:
                try: tp += float(r[jp])
                except Exception: pass
                try: tc += float(r[jc])
                except Exception: pass
        out[sn] = (tp, tc)
    return out


OPT = load(BASE + r'\GitHub_Dashboard_12_Quarters_1PCT_ITM_Options_Master.xlsx', 'Option P&L', 'Premium Paid')
FUT = load(BASE + r'\GitHub_Dashboard_12_Quarters_Master.xlsx', 'Realised P&L', '20% SPAN Margin')


def measures(rets):
    n = len(rets); wins = sum(1 for r in rets if r > 0)
    pnl = [r / 100 * CAP for r in rets]
    cum = 0; peak = 0; mdd = 0
    for x in pnl:
        cum += x; peak = max(peak, cum); mdd = min(mdd, cum - peak)
    std = st.pstdev(rets) if n > 1 else 0
    down = st.pstdev([min(0, r) for r in rets]) if n > 1 else 0
    pos = [r for r in rets if r > 0]; neg = [r for r in rets if r < 0]
    return dict(total_pnl=sum(pnl), overall_ret=sum(pnl) / (CAP * n) * 100,
                wr=100 * wins / n, wins=wins, losses=n - wins,
                avg=st.mean(rets), med=st.median(rets), best=max(rets), worst=min(rets),
                std=std, sharpe=(st.mean(rets) / std * math.sqrt(4)) if std else 0,
                sortino=(st.mean(rets) / down * math.sqrt(4)) if down else None,
                pf=(sum(pos) / abs(sum(neg))) if neg else None, mdd=mdd)


opt_rets = [OPT[q][0] / OPT[q][1] * 100 if OPT[q][1] else 0 for q in ORDER]
fut_rets = [FUT[q][0] / FUT[q][1] * 100 if FUT[q][1] else 0 for q in ORDER]
OM = measures(opt_rets); FM = measures(fut_rets)

# ---- styles ----
NAVY = PatternFill('solid', fgColor='1A202C'); WHITE = Font(bold=True, color='FFFFFF', size=10)
TITLE = Font(bold=True, color='1A202C', size=13); SUB = Font(color='4A5568', size=10, italic=True)
GREENF = PatternFill('solid', fgColor='E6F4EA'); REDF = PatternFill('solid', fgColor='FCE8E6')
AMBERF = PatternFill('solid', fgColor='FEF3C7'); LBL = PatternFill('solid', fgColor='F1F5F9')
GOOD = Font(color='137333', bold=True); BAD = Font(color='C5221F', bold=True); BOLD = Font(bold=True)
THIN = Border(*[Side(style='thin', color='CBD5E0')] * 4)
CEN = Alignment(horizontal='center'); RIGHT = Alignment(horizontal='right'); WRAP = Alignment(wrap_text=True, vertical='top')

wb = openpyxl.Workbook()


def perf_sheet(name, rets, data, m, cap_label, title):
    ws = wb.create_sheet(name)
    ws.merge_cells('A1:F1'); ws.cell(1, 1, title).font = TITLE
    ws.merge_cells('A2:F2'); ws.cell(2, 1, 'Fresh Rs 10,00,000 deployed each quarter (additive). Return %% = quarter P&L / %s.' % cap_label).font = SUB
    hdr = ['Quarter', 'Quarter Net P&L (Rs)', cap_label + ' (Rs)', 'Return %', 'Rs P&L on 10L', 'Cumulative Rs (10L/qtr)']
    for j, h in enumerate(hdr, 1):
        c = ws.cell(4, j, h); c.fill = NAVY; c.font = WHITE; c.border = THIN; c.alignment = CEN
    cum = 0
    for i, q in enumerate(ORDER, 5):
        p, capd = data[q]; r = rets[i - 5]; rs = r / 100 * CAP; cum += rs
        vals = [q, round(p), round(capd), round(r, 1), round(rs), round(cum)]
        fill = GREENF if r >= 0 else REDF
        for j, v in enumerate(vals, 1):
            c = ws.cell(i, j, v); c.border = THIN; c.fill = fill
            if j >= 2: c.alignment = RIGHT
            if j in (4, 5) and r < 0: c.font = BAD
            if j in (4, 5) and r >= 0: c.font = GOOD
    r0 = 5 + len(ORDER) + 1
    ws.cell(r0, 1, 'PERFORMANCE MEASURES (12 quarters)').font = BOLD
    rows = [
        ('Total P&L on 10L/quarter', 'Rs {:+,.0f}'.format(m['total_pnl'])),
        ('Overall return (on Rs {:,.0f} total deployed)'.format(CAP * 12), '{:+.1f}%'.format(m['overall_ret'])),
        ('Win rate (quarters)', '{:.1f}%  ({}W / {}L)'.format(m['wr'], m['wins'], m['losses'])),
        ('Avg / Median return per quarter', '{:+.1f}% / {:+.1f}%'.format(m['avg'], m['med'])),
        ('Best / Worst quarter', '{:+.1f}% / {:+.1f}%'.format(m['best'], m['worst'])),
        ('Std dev (quarterly)', '{:.1f}%'.format(m['std'])),
        ('Sharpe (annualised)', '{:.2f}'.format(m['sharpe'])),
        ('Sortino (annualised)', ('{:.2f}'.format(m['sortino']) if m['sortino'] is not None else 'n/a (no losing quarter)')),
        ('Profit factor', ('{:.2f}'.format(m['pf']) if m['pf'] is not None else 'n/a (no losing quarter)')),
        ('Max drawdown (cumulative Rs)', 'Rs {:,.0f}'.format(m['mdd'])),
    ]
    for k, (lbl, val) in enumerate(rows):
        rr = r0 + 1 + k
        a = ws.cell(rr, 1, lbl); a.fill = LBL; a.font = BOLD; a.border = THIN
        b = ws.cell(rr, 2, val); b.border = THIN
        ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=3)
    for j, w in enumerate([34, 20, 20, 11, 16, 22], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    return ws


perf_sheet('Options_10L', opt_rets, OPT, OM, 'Premium Deployed',
           '1% ITM OPTIONS — Rs 10 Lakh / Quarter Performance')
perf_sheet('Futures_10L', fut_rets, FUT, FM, 'SPAN Margin',
           'FUTURES — Rs 10 Lakh / Quarter Performance (for comparison)')

# summary compare
cs = wb.active; cs.title = 'Summary_Compare'
cs.merge_cells('A1:D1'); cs.cell(1, 1, 'Rs 10 Lakh / Quarter — Options vs Futures (12 quarters)').font = TITLE
cs.merge_cells('A2:D2'); cs.cell(2, 1, 'Additive: fresh Rs 10L each quarter. See "Assumptions" sheet before relying on these.').font = SUB
hdr = ['Measure', '1% ITM Options', 'Futures', 'Note']
for j, h in enumerate(hdr, 1):
    c = cs.cell(4, j, h); c.fill = NAVY; c.font = WHITE; c.border = THIN; c.alignment = CEN
comp = [
    ('Total P&L (12 x 10L)', 'Rs {:+,.0f}'.format(OM['total_pnl']), 'Rs {:+,.0f}'.format(FM['total_pnl']), 'on Rs 1.2 Cr deployed'),
    ('Overall return', '{:+.1f}%'.format(OM['overall_ret']), '{:+.1f}%'.format(FM['overall_ret']), ''),
    ('Avg return / quarter', '{:+.1f}%'.format(OM['avg']), '{:+.1f}%'.format(FM['avg']), 'options amplify'),
    ('Win rate (quarters)', '{:.0f}%'.format(OM['wr']), '{:.0f}%'.format(FM['wr']), '3-yr sample: all quarters won'),
    ('Best / Worst quarter', '{:+.0f}% / {:+.0f}%'.format(OM['best'], OM['worst']), '{:+.0f}% / {:+.0f}%'.format(FM['best'], FM['worst']), ''),
    ('Sharpe (annualised)', '{:.2f}'.format(OM['sharpe']), '{:.2f}'.format(FM['sharpe']), 'in-sample, optimistic'),
    ('Max drawdown (Rs)', 'Rs {:,.0f}'.format(OM['mdd']), 'Rs {:,.0f}'.format(FM['mdd']), 'Rs 0 = no losing qtr in window'),
]
for i, row in enumerate(comp, 5):
    for j, v in enumerate(row, 1):
        c = cs.cell(i, j, v); c.border = THIN
        if j == 1: c.fill = LBL; c.font = BOLD
        if j in (2, 3): c.alignment = CEN
        if j == 4: c.font = Font(color='92400E', size=9, italic=True)
for j, w in enumerate([26, 18, 18, 30], 1):
    cs.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w

# assumptions / reality-check
asf = wb.create_sheet('Assumptions_RealityCheck')
asf.merge_cells('A1:B1'); asf.cell(1, 1, 'ASSUMPTIONS & REALITY CHECK — read before relying on these numbers').font = Font(bold=True, color='7C2D12', size=12)
notes = [
    ('Capital model', 'Fresh Rs 10,00,000 each quarter (additive). Return % = quarter net P&L / capital deployed (premium for options, 20% SPAN margin for futures), scaled to Rs 10L. NOT compounded (a fresh book each quarter cannot compound).'),
    ('Option pricing = MODEL, not real', 'Option P&L uses the synthetic 1% ITM convention (1% intrinsic + 1% time-value entry / 0.5% exit, 0.15% cost) - the same model on your dashboard. These are NOT real option-chain fills. Real near-month 1% ITM options around earnings carry 3-5% time value (high IV then crush), which would LOWER these returns and create losing quarters.'),
    ('Why 100% quarterly win rate', 'The backtest window is only 12 quarters (FY24_Q3 to FY27_Q2, ~2023-2026), a period in which the underlying futures strategy won EVERY quarter. So both options and futures show 100% quarterly win rate and zero drawdown here. That is a small, in-sample, bull-market artifact - not proof the strategy never loses a quarter.'),
    ('Book-level netting', 'Return is on the whole 50-stock basket. At the stock level most options LOSE (only ~30% win); the few big winners dominate the net. A concentrated book (few names) would look very different and far riskier.'),
    ('Sharpe / drawdown caveats', 'Sharpe is in-sample on 12 correlated quarters and is optimistic. Max drawdown shows Rs 0 only because no quarter lost in this window - it is NOT a real risk estimate. Expect real losing quarters and real drawdowns going forward.'),
    ('Bottom line', 'Directionally: options out-earn futures over this window by amplifying the winners. But treat the absolute figures (e.g. +80%/quarter) as an OPTIMISTIC model ceiling, not a forward expectation. For a realistic view, re-run the options at 3% time value.'),
]
r = 3
for lbl, txt in notes:
    a = asf.cell(r, 1, lbl); a.font = BOLD; a.fill = AMBERF; a.border = THIN; a.alignment = WRAP
    b = asf.cell(r, 2, txt); b.alignment = WRAP; b.border = THIN
    asf.row_dimensions[r].height = 62
    r += 1
asf.column_dimensions['A'].width = 24; asf.column_dimensions['B'].width = 105

wb.move_sheet('Summary_Compare', -(len(wb.sheetnames) - 1))
wb.save(OUT)
print('Saved:', OUT)
print('OPTIONS: total Rs {:+,.0f} on 10L/qtr, avg {:+.1f}%/qtr, win {:.0f}%, maxDD Rs {:,.0f}'.format(OM['total_pnl'], OM['avg'], OM['wr'], OM['mdd']))
print('FUTURES: total Rs {:+,.0f} on 10L/qtr, avg {:+.1f}%/qtr, win {:.0f}%'.format(FM['total_pnl'], FM['avg'], FM['wr']))
