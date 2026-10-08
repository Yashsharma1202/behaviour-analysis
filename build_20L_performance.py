"""
Rs 20 LAKH / QUARTER PERFORMANCE — 1% ITM OPTIONS, 12 quarters, with a REAL drawdown.
Each quarter gets its own fresh Rs 20,00,000 (individual funds). Premium is scaled so
each quarter deploys Rs 20L; each stock option is then a real trade. Drawdown is measured
at the TRADE level (order every stock's option by exit date and run the equity curve) —
so intra-quarter losing streaks show up, unlike the (wrong) quarter-level DD of Rs 0.

Source: GitHub_Dashboard_12_Quarters_1PCT_ITM_Options_Master.xlsx.
ADDITIVE: writes a new workbook.
"""
import openpyxl, statistics as st, math, re
from datetime import date
from collections import defaultdict
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

BASE = r'D:\behaviour analysis'
OUT = BASE + r'\GitHub_Dashboard_20L_Quarter_Options_Performance.xlsx'
CAP = 2_000_000
ORDER = ['FY24_Q3', 'FY24_Q4', 'FY25_Q1', 'FY25_Q2', 'FY25_Q3', 'FY25_Q4',
         'FY26_Q1', 'FY26_Q2', 'FY26_Q3', 'FY26_Q4', 'FY27_Q1', 'FY27_Q2']
MONS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']


def pdate(s):
    s = str(s)
    m = re.search(r'(\d{4})-(\d{2})-(\d{2})', s)
    if m: return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    m = re.search(r'(\d{1,2})-([A-Za-z]{3})-(\d{4})', s)
    if m: return date(int(m.group(3)), MONS.index(m.group(2)) + 1, int(m.group(1)))
    return None


wb0 = openpyxl.load_workbook(BASE + r'\GitHub_Dashboard_12_Quarters_1PCT_ITM_Options_Master.xlsx', read_only=True, data_only=True)
trades = []; premq = defaultdict(float)
for sn in wb0.sheetnames:
    if sn == 'Dashboard_Summary': continue
    rows = list(wb0[sn].iter_rows(values_only=True)); hdr = rows[3]
    idx = {str(h).strip(): j for j, h in enumerate(hdr) if h}
    def C(n):
        for k, j in idx.items():
            if k.lower().startswith(n.lower()): return j
    jS, jx, jp, jpr = C('Symbol'), C('Exit Date'), C('Option P&L'), C('Premium Paid')
    for r in rows[4:]:
        if not r or r[jS] is None: continue
        try: pnl = float(r[jp]); prem = float(r[jpr])
        except Exception: continue
        trades.append({'q': sn, 'sym': str(r[jS]), 'xd': pdate(r[jx]), 'pnl': pnl, 'prem': prem})
        premq[sn] += prem
for t in trades:
    t['spnl'] = t['pnl'] * (CAP / premq[t['q']]) if premq[t['q']] else 0

# continuous trade-level equity (ordered by exit date)
trades.sort(key=lambda t: (t['xd'] or date(1900, 1, 1)))
cum = peak = mdd = 0.0; mdd_date = None
for t in trades:
    cum += t['spnl']; peak = max(peak, cum)
    if cum - peak < mdd: mdd = cum - peak; mdd_date = t['xd']
trade_win = 100 * sum(1 for t in trades if t['spnl'] > 0) / len(trades)

# per-quarter: return, Rs, intra-quarter DD
byq = defaultdict(list)
for t in trades: byq[t['q']].append(t)
Q = {}
for q in ORDER:
    ts = sorted(byq[q], key=lambda t: (t['xd'] or date(1900, 1, 1)))
    c = pk = md = 0.0
    for t in ts:
        c += t['spnl']; pk = max(pk, c); md = min(md, c - pk)
    Q[q] = dict(ret=c / CAP * 100, pnl=c, dd=md, prem=premq[q])

rets = [Q[q]['ret'] for q in ORDER]; pnls = [Q[q]['pnl'] for q in ORDER]
wins = sum(1 for r in rets if r > 0); std = st.pstdev(rets)
neg = [r for r in rets if r < 0]; pos = [r for r in rets if r > 0]
M = dict(total=sum(pnls), overall=sum(pnls) / (CAP * 12) * 100, wr=100 * wins / 12,
         avg=st.mean(rets), med=st.median(rets), best=max(rets), worst=min(rets), std=std,
         sharpe=st.mean(rets) / std * math.sqrt(4) if std else 0,
         pf=(sum(pos) / abs(sum(neg))) if neg else None,
         mdd=mdd, mdd_pct=mdd / CAP * 100, mdd_date=mdd_date,
         worst_intraq=min(Q[q]['dd'] for q in ORDER), trade_win=trade_win)

# ---- styles ----
NAVY = PatternFill('solid', fgColor='1A202C'); WHITE = Font(bold=True, color='FFFFFF', size=10)
TITLE = Font(bold=True, color='1A202C', size=13); SUB = Font(color='4A5568', size=10, italic=True)
GREENF = PatternFill('solid', fgColor='E6F4EA'); REDF = PatternFill('solid', fgColor='FCE8E6')
AMBERF = PatternFill('solid', fgColor='FEF3C7'); LBL = PatternFill('solid', fgColor='F1F5F9')
GOOD = Font(color='137333', bold=True); BAD = Font(color='C5221F', bold=True); BOLD = Font(bold=True)
THIN = Border(*[Side(style='thin', color='CBD5E0')] * 4)
CEN = Alignment(horizontal='center'); RIGHT = Alignment(horizontal='right'); WRAP = Alignment(wrap_text=True, vertical='top')

wb = openpyxl.Workbook()
ws = wb.active; ws.title = 'Options_20L'
ws.merge_cells('A1:G1'); ws.cell(1, 1, '1% ITM OPTIONS — Rs 20 Lakh / Quarter Performance (12 quarters, real trade-level drawdown)').font = TITLE
ws.merge_cells('A2:G2'); ws.cell(2, 1, 'Individual Rs 20,00,000 each quarter. Return %% = quarter P&L / premium deployed. Intra-quarter DD = worst dip as stock options resolve.').font = SUB
hdr = ['Quarter', 'Premium Deployed (Rs)', 'Return %', 'Rs P&L on 20L', 'Intra-Qtr Max DD (Rs)', 'Intra-Qtr DD %', 'Cumulative Rs']
for j, h in enumerate(hdr, 1):
    c = ws.cell(4, j, h); c.fill = NAVY; c.font = WHITE; c.border = THIN; c.alignment = CEN
cum = 0
for i, q in enumerate(ORDER, 5):
    d = Q[q]; cum += d['pnl']
    vals = [q, round(CAP), round(d['ret'], 1), round(d['pnl']), round(d['dd']), round(d['dd'] / CAP * 100, 1), round(cum)]
    fill = GREENF if d['ret'] >= 0 else REDF
    for j, v in enumerate(vals, 1):
        c = ws.cell(i, j, v); c.border = THIN; c.fill = fill
        if j >= 2: c.alignment = RIGHT
        if j == 3: c.font = GOOD if d['ret'] >= 0 else BAD
        if j in (5, 6): c.font = BAD
r0 = 5 + len(ORDER) + 1
ws.cell(r0, 1, 'PERFORMANCE MEASURES (12 quarters, Rs 20L/quarter)').font = BOLD
rows = [
    ('Total P&L (12 x Rs 20L)', 'Rs {:+,.0f}'.format(M['total'])),
    ('Overall return (on Rs 2.4 Cr deployed)', '{:+.1f}%'.format(M['overall'])),
    ('Quarter win rate', '{:.0f}%  ({}W / {}L)'.format(M['wr'], wins, 12 - wins)),
    ('Trade-level win rate (per option)', '{:.1f}%  (the ~{:.0f}% losers cluster into the drawdown)'.format(M['trade_win'], 100 - M['trade_win'])),
    ('Avg / Median return per quarter', '{:+.1f}% / {:+.1f}%'.format(M['avg'], M['med'])),
    ('Best / Worst quarter', '{:+.1f}% / {:+.1f}%'.format(M['best'], M['worst'])),
    ('Std dev (quarterly)', '{:.1f}%'.format(M['std'])),
    ('Sharpe (annualised)', '{:.2f}'.format(M['sharpe'])),
    ('Profit factor (quarterly)', ('{:.2f}'.format(M['pf']) if M['pf'] is not None else 'n/a (no losing quarter)')),
    ('>> MAX DRAWDOWN (real, trade-level)', 'Rs {:,.0f}   ({:.1f}% of a 20L book)  in {}'.format(M['mdd'], M['mdd_pct'], M['mdd_date'])),
    ('Worst intra-quarter drawdown', '{:.1f}%'.format(M['worst_intraq'] / CAP * 100)),
]
for k, (lbl, val) in enumerate(rows):
    rr = r0 + 1 + k
    a = ws.cell(rr, 1, lbl); a.fill = (AMBERF if lbl.startswith('>>') else LBL); a.font = BOLD; a.border = THIN
    b = ws.cell(rr, 2, val); b.border = THIN
    ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=4)
for j, w in enumerate([36, 22, 11, 16, 20, 15, 18], 1):
    ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
ws.freeze_panes = 'A5'

# assumptions
asf = wb.create_sheet('Assumptions_RealityCheck')
asf.merge_cells('A1:B1'); asf.cell(1, 1, 'ASSUMPTIONS & REALITY CHECK — read first').font = Font(bold=True, color='7C2D12', size=12)
notes = [
    ('Capital', 'Individual Rs 20,00,000 each quarter (fresh, additive). Return % = quarter net P&L / premium deployed, scaled so each quarter deploys Rs 20L.'),
    ('Drawdown is now REAL', 'Measured at the TRADE level: every stock option is ordered by its exit date and the running equity is tracked. Losing options (the majority) create real dips - max drawdown here is {:.1f}% (Rs {:,.0f}). The earlier "Rs 0 drawdown" was wrong (it only looked at quarter totals).'.format(M['mdd_pct'], M['mdd'])),
    ('Quarter win 100% vs trade win {:.0f}%'.format(M['trade_win']), 'Every one of the 12 quarters (2023-2026) netted positive, so quarter win rate is 100% - a small in-sample, bull-market window. At the single-option level ~{:.0f}% win and ~{:.0f}% lose; when the losers cluster (by exit date) the book dips -{:.0f}% intra-quarter. That is the real risk, hidden by the quarter totals.'.format(M['trade_win'], 100 - M['trade_win'], abs(M['mdd_pct']))),
    ('Option pricing = MODEL', 'Uses the synthetic 1% ITM convention (1% intrinsic + 1% time-value entry / 0.5% exit, 0.15% cost), same as your dashboard - NOT real option fills. Real near-month premiums (3-5% time value around earnings) would lower returns and deepen drawdowns.'),
    ('Bottom line', 'On this window options turned Rs 20L/quarter into Rs {:+,.0f} total with a real -{:.0f}% peak dip. Treat the absolute returns as an OPTIMISTIC model ceiling; the -{:.0f}% drawdown and ~{:.0f}% trade win rate are the honest risk picture. Re-run at 3% time value for a conservative view.'.format(M['total'], abs(M['mdd_pct']), abs(M['mdd_pct']), M['trade_win'])),
]
r = 3
for lbl, txt in notes:
    a = asf.cell(r, 1, lbl); a.font = BOLD; a.fill = AMBERF; a.border = THIN; a.alignment = WRAP
    b = asf.cell(r, 2, txt); b.alignment = WRAP; b.border = THIN
    asf.row_dimensions[r].height = 70; r += 1
asf.column_dimensions['A'].width = 26; asf.column_dimensions['B'].width = 108

wb.save(OUT)
print('Saved:', OUT)
print('Rs 20L/qtr OPTIONS: total Rs {:+,.0f} | avg {:+.1f}%/qtr | REAL maxDD {:.1f}% (Rs {:,.0f}) | trade-win {:.0f}%'
      .format(M['total'], M['avg'], M['mdd_pct'], M['mdd'], M['trade_win']))
