"""
Two-page portfolio performance report (PDF) + detailed Excel with an audit trail.
 - reads November_Portfolio_Performance.xlsx (all numbers come from it; nothing is typed in by hand)
 - every number printed in the tables is logged with its source cell -> sheet 'Report_Numbers'
 - the log is re-checked against the workbook by an independent lookup, and every printed value is searched for in the PDF text
 - verify_portfolio.run_checks() re-simulates strategies from raw prices (53 checks) -> sheet 'Verification'
Outputs: November_Portfolio_2Page_Report.pdf, November_Portfolio_Report_Data.xlsx (new files)
"""
import sys, re
sys.path.insert(0, 'D:/behaviour analysis'); sys.stdout.reconfigure(encoding='utf-8')
import numpy as np, pandas as pd, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.graphics.shapes import Drawing, String, Line, Rect
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.charts.lineplots import LinePlot
import verify_portfolio as vp

ROOT = 'D:/behaviour analysis/'
PERF = ROOT + 'November_Portfolio_Performance.xlsx'
OUT_PDF = ROOT + 'November_Portfolio_2Page_Report.pdf'; OUT_XL = ROOT + 'November_Portfolio_Report_Data.xlsx'
SIZES = (1, 3, 5, 10)

A = pd.read_excel(PERF, sheet_name='All_Measures')
EQ = pd.read_excel(PERF, sheet_name='Equity_Daily'); YR = pd.read_excel(PERF, sheet_name='Yearly_Returns')
UT = pd.read_excel(PERF, sheet_name='Utilisation_By_Month').iloc[:12]; SCL = pd.read_excel(PERF, sheet_name='Scaled_Positions'); CST = pd.read_excel(PERF, sheet_name='Cost_Sensitivity')
nm = lambda pol, n, sy: f'{pol}-Top{n} ({sy}-25)'

# ---------------- number log ----------------
LOG = []
def fmt_num(v, d, plus=False):
    if v is None or (isinstance(v, float) and np.isnan(v)): return '-'
    return f'{v:+.{d}f}' if plus else f'{v:.{d}f}'
def mget(col, label, nth=0, d=2, plus=False, sec=''):
    vals = A[A.Measure == label][col].tolist(); v = float(vals[min(nth, len(vals) - 1)])
    s = fmt_num(v, d, plus); LOG.append(dict(Section=sec, Item=f'{col} | {label}', Displayed=s, Raw=v, SourceSheet='All_Measures', SourceRow=label, SourceCol=col, Nth=nth, Decimals=d, Plus=plus)); return s
def mraw(col, label, nth=0, sec=''):
    vals = A[A.Measure == label][col].tolist(); v = float(vals[min(nth, len(vals) - 1)])
    LOG.append(dict(Section=sec, Item=f'{col} | {label}', Displayed=f'{v:.4f}', Raw=v, SourceSheet='All_Measures', SourceRow=label, SourceCol=col, Nth=nth, Decimals=4, Plus=False)); return v
def other(sheet, df, rowcol, rowval, col, d=2, plus=False, sec='', extra=None):
    q = df[df[rowcol] == rowval]
    if extra:
        for k, v in extra.items(): q = q[q[k] == v]
    v = float(q[col].iloc[0]); s = fmt_num(v, d, plus)
    LOG.append(dict(Section=sec, Item=f'{rowval} | {col}', Displayed=s, Raw=v, SourceSheet=sheet, SourceRow=f'{rowcol}={rowval}' + (f' & {extra}' if extra else ''), SourceCol=col, Nth=0, Decimals=d, Plus=plus)); return s

# ---------------- styles ----------------
ss = getSampleStyleSheet(); NAVY = colors.HexColor('#1F3864')
def PS(name, size, **kw): return ParagraphStyle(name, parent=ss['BodyText'], fontSize=size, leading=size + 1.5, **kw)
TITLE = PS('t', 13.5, fontName='Helvetica-Bold', textColor=NAVY); SUB = PS('s', 6.8, textColor=colors.HexColor('#555555'))
H2 = PS('h2', 8.6, fontName='Helvetica-Bold', textColor=NAVY, spaceBefore=3, spaceAfter=1.5)
BODY = PS('b', 7.7); SMALL = PS('sm', 6.7, textColor=colors.HexColor('#444444'))
def cell_styles(sz): return PS(f'c{sz}', sz), PS(f'ch{sz}', sz, fontName='Helvetica-Bold', textColor=colors.white), PS(f'cb{sz}', sz, fontName='Helvetica-Bold')
PG = landscape(A4); LM = 9 * mm; W = PG[0] - 2 * LM
P = lambda t, st=BODY: Paragraph(t, st)

BUMP, PADB = 0.8, 0.3
def tbl(rows, widths, fs=6.4, pad=0.9, extra=None, boldcol0=True):
    fs = fs + BUMP; pad = pad + PADB
    C, CH, CB = cell_styles(fs)
    data = [[Paragraph(str(c), CH if i == 0 else (CB if (boldcol0 and j == 0) else C)) for j, c in enumerate(r)] for i, r in enumerate(rows)]
    t = Table(data, colWidths=widths, repeatRows=1)
    sty = [('BACKGROUND', (0, 0), (-1, 0), NAVY), ('GRID', (0, 0), (-1, -1), 0.25, colors.HexColor('#B7B7B7')), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
           ('LEFTPADDING', (0, 0), (-1, -1), 2), ('RIGHTPADDING', (0, 0), (-1, -1), 2), ('TOPPADDING', (0, 0), (-1, -1), pad), ('BOTTOMPADDING', (0, 0), (-1, -1), pad)]
    for i in range(2, len(rows), 2): sty.append(('BACKGROUND', (0, i), (-1, i), colors.HexColor('#F4F6FA')))
    if extra: sty += extra
    t.setStyle(TableStyle(sty)); return t

def footer(cv, doc):
    cv.saveState(); cv.setFont('Helvetica', 6.2); cv.setFillColor(colors.HexColor('#666666'))
    cv.drawString(LM, 5.5 * mm, 'November Top-N strategy - portfolio performance. Historical simulation on adjusted data, test years 2013-2025; not investment advice. Detail: November_Portfolio_Report_Data.xlsx')
    cv.drawRightString(PG[0] - LM, 5.5 * mm, f'Page {doc.page} of 2'); cv.restoreState()

def line_chart(title, series, w, h):
    d = Drawing(w, h); lp = LinePlot(); lp.x = 30; lp.y = 20; lp.width = w - 42; lp.height = h - 34
    lp.data = [s[1] for s in series]
    for i, s in enumerate(series):
        lp.lines[i].strokeColor = colors.HexColor(s[2]); lp.lines[i].strokeWidth = 1.1
        if s[3]: lp.lines[i].strokeDashArray = [3, 2]
    lp.xValueAxis.valueMin = 2013; lp.xValueAxis.valueMax = 2026; lp.xValueAxis.valueSteps = [2013, 2015, 2017, 2019, 2021, 2023, 2025]; lp.xValueAxis.labels.fontSize = 6
    lp.yValueAxis.labels.fontSize = 6; lp.yValueAxis.valueMin = 0; lp.xValueAxis.labelTextFormat = '%d'; lp.yValueAxis.labelTextFormat = '%.0f'
    d.add(lp); d.add(String(w / 2, h - 8, title, fontSize=7.2, fontName='Helvetica-Bold', textAnchor='middle'))
    x = 32
    for s in series:
        d.add(Line(x, 8, x + 12, 8, strokeColor=colors.HexColor(s[2]), strokeWidth=1.6, strokeDashArray=[3, 2] if s[3] else None)); d.add(String(x + 15, 5.5, s[0], fontSize=6)); x += 26 + 3.6 * len(s[0])
    return d

def bar_chart(title, cats, vals, color, w, h, fmt='%.1f', ymax=None):
    d = Drawing(w, h); c = VerticalBarChart(); c.x = 28; c.y = 18; c.width = w - 40; c.height = h - 32
    c.data = [vals]; c.categoryAxis.categoryNames = cats; c.categoryAxis.labels.fontSize = 6; c.valueAxis.labels.fontSize = 6; c.valueAxis.valueMin = 0
    if ymax: c.valueAxis.valueMax = ymax
    c.bars[0].fillColor = colors.HexColor(color); c.bars.strokeColor = None; c.barLabelFormat = fmt; c.barLabels.fontSize = 5.6; c.barLabels.nudge = 4; c.groupSpacing = 6
    d.add(c); d.add(String(w / 2, h - 8, title, fontSize=7.2, fontName='Helvetica-Bold', textAnchor='middle')); return d

a1, a3, a5, a10 = [nm('ALL', n, 2013) for n in SIZES]
h3 = nm('HINDSIGHT', 3, 2013); e3 = nm('EXPANDING', 3, 2018); a3b = nm('ALL', 3, 2018)
LAYOUT = {}
class Doc(SimpleDocTemplate):
    def afterFlowable(self, flowable):
        LAYOUT[self.page] = (self.frame._y - self.frame._y1) / mm      # free height (mm) left in the frame after this flowable
doc = Doc(OUT_PDF, pagesize=PG, leftMargin=LM, rightMargin=LM, topMargin=7 * mm, bottomMargin=9 * mm, title='November Top-N Strategy - Portfolio Performance (2 pages)', author='Behaviour Analysis project')
E = []

# ======================= PAGE 1 =======================
E += [P('November Top-N Strategy - Portfolio Performance: CAGR, Fund Utilisation, Sharpe and more', TITLE),
      P('Rs 10,00,000 account trading the walk-forward Top 1 / 3 / 5 / 10 stocks of the 30 November dates, test years 2013-2025 &nbsp;|&nbsp; stocks and rules are always chosen from EARLIER years only &nbsp;|&nbsp; Nifty 50 = benchmark &nbsp;|&nbsp; prepared 23 Oct 2026', SUB), Spacer(1, 2)]
sec = 'Plain words'
cagr0 = mget(a3, 'CAGR % (idle cash 0%)', d=1, sec=sec); cagr1 = mget(a1, 'CAGR % (idle cash 0%)', d=1, sec=sec); cagr10 = mget(a10, 'CAGR % (idle cash 0%)', d=1, sec=sec)
cagrrf = mget(a3, 'CAGR % (idle cash earns 6.5%)', d=1, sec=sec); nifc = mget(a3, 'Nifty 50 CAGR % (same window)', d=1, sec=sec)
util = mget(a3, 'Average fund utilised % (of equity, all days)', d=1, sec=sec); dim = mget(a3, 'Days in market %', d=0, sec=sec); rod = mget(a3, 'Return on deployed capital % per year', d=0, sec=sec)
sh0 = mget(a3, 'Sharpe ratio (annualised, rf 6.5%)', d=2, sec=sec); shrf = mget(a3, 'Sharpe ratio (annualised)', d=2, sec=sec); nsh = mget(a3, 'Nifty Sharpe (same window)', d=2, sec=sec)
exc = mraw(a3, 'CAGR % (idle cash earns 6.5%)', sec=sec) - 6.5; nexc = mraw(a3, 'Nifty 50 CAGR % (same window)', sec=sec) - 6.5
mdd3 = mget(a3, 'Maximum drawdown %', d=1, sec=sec); mdd1 = mget(a1, 'Maximum drawdown %', d=1, sec=sec); nmdd = mget(a3, 'Nifty max drawdown % (same window)', d=0, sec=sec)
ch = mget(h3, 'CAGR % (idle cash 0%)', d=1, sec=sec); ce = mget(e3, 'CAGR % (idle cash 0%)', d=1, sec=sec); ca = mget(a3b, 'CAGR % (idle cash 0%)', d=1, sec=sec)
b5 = mget(a3, 'Bootstrap CAGR 5th percentile %', d=1, sec=sec); b95 = mget(a3, 'Bootstrap CAGR 95th percentile %', d=1, sec=sec)
words = [['Question', 'Answer'],
         ['CAGR?', f'<b>{cagr0}% a year</b> for the all-dates Top 3 account (Top 1 {cagr1}%, Top 10 {cagr10}%) if idle cash earns nothing, <b>{cagrrf}%</b> if idle cash sits in a 6.5% liquid fund. Nifty 50 over the same years: <b>{nifc}%</b>. The strategy does <u>not</u> beat buy-and-hold.'],
         ['Fund utilised?', f'Only <b>{util}% of the account</b> is deployed on average (positions are open on {dim}% of days, almost all between late October and mid-December). The trades earn about <b>{rod}% a year on the capital while it is working</b>; the rest of the account sits idle.'],
         ['Sharpe?', f'<b>{sh0}</b> with idle cash at 0% (the account earns less than the 6.5% risk-free rate). With idle cash earning 6.5%: <b>{shrf}</b>, level with the Nifty ({nsh}) - but only because volatility is tiny: the account is just {exc:.1f} points a year above the risk-free rate, the Nifty {nexc:.1f}.'],
         ['Drawdown?', f'Max drawdown {mdd3}% for Top 3 ({mdd1}% for Top 1) - little capital is at risk at any time; Nifty 50 maximum drawdown in the same years: {nmdd}%.'],
         ['Best dates?', f'Choosing dates with hindsight lifts Top 3 CAGR to {ch}% (biased). Choosing them from earlier years only (2018-2025) gives {ce}% vs {ca}% for trading all dates: no real gain. Bootstrap range of Top 3 CAGR: {b5}% to {b95}% - small but reliably positive; 13 test years only.']]
E += [tbl(words, [19 * mm, W - 19 * mm], fs=6.7, pad=1.1)]
E += [P('<b>How it is run:</b> each year, for each of the 30 dates (a "sleeve"), the Top-N stocks and their "buy N days before / sell M days after" rule are picked from earlier years; equity is split equally across the sleeves in use and each sleeve across its N stocks; buy at the close of T-n, sell at the close of T+m (T = first trading day on/after the date); '
         'cost 0.15% per position; daily mark-to-market; risk-free 6.5%. Policies: <b>ALL</b> = all 30 dates; <b>EXPAND</b> = dates qualify from their record in earlier years (honest, 2018-2025); <b>HINDSIGHT*</b> = dates chosen using the whole period (biased, reference only).', SMALL), Spacer(1, 1.5)]

E += [P('1. Headline measures', H2)]
heads = ['Measure']; colsel = []
for n in SIZES:
    for pol, sy, lab in (('ALL', 2013, 'ALL'), ('EXPANDING', 2018, 'EXPAND'), ('HINDSIGHT', 2013, 'HIND.*')):
        colsel.append(nm(pol, n, sy)); heads.append(f'{lab} Top {n}<br/>{sy}-25')
ROWS = [('CAGR % (idle cash 0%)', 'CAGR % (idle cash 0%)', 0, 2, False), ('CAGR % (idle cash 6.5%)', 'CAGR % (idle cash earns 6.5%)', 0, 2, False), ('Nifty 50 CAGR %, same window', 'Nifty 50 CAGR % (same window)', 0, 2, False),
        ('Average fund utilised %', 'Average fund utilised % (of equity, all days)', 0, 2, False), ('Peak fund utilised %', 'Peak fund utilised %', 0, 0, False), ('Days in market %', 'Days in market %', 0, 0, False),
        ('Return on deployed capital %/yr', 'Return on deployed capital % per year', 0, 1, False), ('Sharpe (idle cash 0%)', 'Sharpe ratio (annualised, rf 6.5%)', 0, 2, False), ('Sharpe (idle cash 6.5%)', 'Sharpe ratio (annualised)', 0, 2, False),
        ('Max drawdown % (idle 0%)', 'Maximum drawdown %', 0, 1, False), ('Annual volatility % (idle 0%)', 'Annualised volatility %', 0, 2, False), ('Beta vs Nifty', 'Beta (daily)', 0, 3, False),
        ('Trade win %', 'Trade win %', 0, 1, False), ('Profit factor', 'Profit factor', 0, 2, False), ('Worst year %', 'Worst year %', 0, 2, False)]
rows = [heads]
for lab, key, nth, d, plus in ROWS: rows.append([lab] + [mget(c, key, nth, d, plus, sec='Headline table') for c in colsel])
hl = []
for j, c in enumerate(colsel, 1):
    if 'HINDSIGHT' in c:
        hl.append(('BACKGROUND', (j, 0), (j, 0), colors.HexColor('#B71C1C')))
        for i in range(1, len(rows)): hl.append(('BACKGROUND', (j, i), (j, i), colors.HexColor('#F8E1E1')))
E += [tbl(rows, [37 * mm] + [(W - 37 * mm) / 12] * 12, fs=6.1, pad=0.55, extra=hl),
      P('* HINDSIGHT (red) picks the dates using the whole 2013-2025 results - biased, shown for reference only. EXPAND covers 2018-2025. Sharpe with idle cash at 0% is negative because the account earns less than the 6.5% risk-free rate.', SMALL), Spacer(1, 1.5)]

# charts
def series_for(colname_, label, color, dashed=False):
    s = EQ[['Date', colname_]].dropna().iloc[::6]; x = pd.to_datetime(s['Date']); xf = x.dt.year + (x.dt.dayofyear - 1) / 365.25
    return (label, list(zip(xf.tolist(), (s[colname_] / 100000).tolist())), color, dashed)
cw = (W - 4 * mm) / 2
LOG.append(dict(Section='Chart data', Item='Equity_Daily columns used by the growth chart', Displayed='see sheet', Raw=float('nan'), SourceSheet='Equity_Daily', SourceRow='all dates', SourceCol='Nifty, ALL-Top3, ALL-Top10, HINDSIGHT-Top3', Nth=0, Decimals=0, Plus=False))
c1 = line_chart('Account value, Rs lakh (idle cash 0%): Nifty 50 vs ALL Top 3 / Top 10 vs HINDSIGHT Top 3 (biased)',
                [series_for('Nifty 50 (rebased to Rs 10,00,000)', 'Nifty 50', '#7F7F7F'), series_for(a3 + ' equity', 'ALL Top 3', '#2E75B6'), series_for(a10 + ' equity', 'ALL Top 10', '#2E9E9E'), series_for(h3 + ' equity', 'HINDSIGHT Top 3*', '#B71C1C', True)], cw, 46 * mm)
u3 = UT[a3 + ' utilised %'] if (a3 + ' utilised %') in UT.columns else UT[a3]
for m_, v_ in zip(UT['Calendar month'], u3): LOG.append(dict(Section='Utilisation chart', Item=f'{a3} | {m_}', Displayed=f'{v_:.1f}', Raw=float(v_), SourceSheet='Utilisation_By_Month', SourceRow=m_, SourceCol=a3, Nth=0, Decimals=1, Plus=False))
c2 = bar_chart('Average % of the account deployed, by calendar month (ALL 30 dates, Top 3)', UT['Calendar month'].tolist(), [round(float(v), 1) for v in u3], '#2E75B6', cw, 46 * mm, ymax=45)
E += [Table([[c1, c2]], colWidths=[cw + 2 * mm, cw + 2 * mm], style=TableStyle([('LEFTPADDING', (0, 0), (-1, -1), 0), ('RIGHTPADDING', (0, 0), (-1, -1), 0), ('TOPPADDING', (0, 0), (-1, -1), 0), ('BOTTOMPADDING', (0, 0), (-1, -1), 0)]))]

# ======================= PAGE 2 =======================
E += [PageBreak()]
LW = 136 * mm; RW = W - LW - 5 * mm
# --- left column
left = [P('2. Risk and market sensitivity (ALL 30 dates, 2013-2025)', H2)]
rows = [['Measure', 'Top 1', 'Top 3', 'Top 5', 'Top 10', 'Nifty 50']]
def rr(lab, key, d, nifkey=None, nth=0):
    cells = [mget(nm('ALL', n, 2013), key, nth, d, sec='Risk table') for n in SIZES]
    return [lab] + cells + [mget(a3, nifkey, 0, d, sec='Risk table') if nifkey else '-']
rows += [rr('Annual volatility % (idle 0%)', 'Annualised volatility %', 2, 'Nifty volatility % (same window)'), rr('Max drawdown % (idle 0%)', 'Maximum drawdown %', 1, 'Nifty max drawdown % (same window)'),
         rr('Longest time below a peak (days)', 'Longest time below a peak (days)', 0), rr('Sortino (idle cash 6.5%)', 'Sortino ratio', 2, None, 1), rr('Calmar (idle cash 6.5%)', 'Calmar ratio', 2),
         rr('Daily VaR 95% (invested days) %', 'Daily VaR 95% (days invested) %', 2), rr('Daily CVaR 95% (invested days) %', 'Daily CVaR 95% (days invested) %', 2), rr('Worst day %', 'Worst day %', 2), rr('Worst month %', 'Worst month %', 2),
         rr('Beta vs Nifty', 'Beta (daily)', 3), rr('Alpha % / yr', 'Alpha % per year (daily regression)', 2), rr('Correlation with Nifty', 'Correlation', 3), rr('Up-market capture', 'Up-market capture', 2),
         rr('Down-market capture', 'Down-market capture', 2), rr('Information ratio', 'Information ratio', 2), rr('Positive years %', 'Positive years %', 0)]
left += [tbl(rows, [41 * mm] + [(LW - 41 * mm) / 5] * 5, fs=6.1, pad=0.55)]
left += [P('3. Calendar-year returns of the account (idle cash 0%), %', H2)]
rows = [['Year', 'Nifty 50', 'ALL Top 1', 'ALL Top 3', 'ALL Top 5', 'ALL Top 10']]; ex = []
for i, y in enumerate(YR['Year'].tolist(), 1):
    r = [str(int(y))]
    for j, c in enumerate(['Nifty 50 %', a1, a3, a5, a10], 1):
        v = float(YR.loc[YR.Year == y, c].iloc[0]); r.append(fmt_num(v, 2, True))
        LOG.append(dict(Section='Yearly returns', Item=f'{int(y)} | {c}', Displayed=fmt_num(v, 2, True), Raw=v, SourceSheet='Yearly_Returns', SourceRow=f'Year={int(y)}', SourceCol=c, Nth=0, Decimals=2, Plus=True))
        if j >= 2: ex.append(('TEXTCOLOR', (j, i), (j, i), colors.HexColor('#B71C1C' if v < 0 else '#1B5E20')))
    rows.append(r)
left += [tbl(rows, [14 * mm] + [(LW - 14 * mm) / 5] * 5, fs=6.0, pad=0.4, extra=ex)]

# --- right column
right = [P('4. Does picking the "best" dates help? Like-for-like 2018-2025', H2)]
rows = [['Portfolio', 'ALL dates: CAGR % / Sharpe* / worst yr %', 'EXPANDING (honest)', 'HINDSIGHT (biased)', 'EXPAND - ALL, CAGR pts']]
for n in SIZES:
    cA, cE, cH = nm('ALL', n, 2018), nm('EXPANDING', n, 2018), nm('HINDSIGHT', n, 2018)
    def comp(c): return f"{mget(c, 'CAGR % (idle cash 0%)', 0, 2, sec='Date selection')} / {mget(c, 'Sharpe ratio (annualised)', 0, 2, sec='Date selection')} / {mget(c, 'Worst year %', 0, 2, sec='Date selection')}"
    dif = mraw(cE, 'CAGR % (idle cash 0%)', sec='Date selection') - mraw(cA, 'CAGR % (idle cash 0%)', sec='Date selection')
    rows.append([f'Top {n}', comp(cA), comp(cE), comp(cH), fmt_num(dif, 2, True)])
right += [tbl(rows, [13 * mm] + [(RW - 13 * mm - 17 * mm) / 3] * 3 + [17 * mm], fs=5.8, pad=0.5),
          P('*Sharpe with idle cash earning 6.5%. The honest date filter does not clearly beat trading all dates; only hindsight improves it.', SMALL)]
right += [P('5. Fund utilisation and scaled positions (ALL 30 dates, 2013-2025)', H2)]
rows = [['Measure', 'Top 1', 'Top 3', 'Top 5', 'Top 10']]
for lab, key, d in (('Average fund utilised % (all days)', 'Average fund utilised % (of equity, all days)', 2), ('Peak fund utilised %', 'Peak fund utilised %', 0), ('Days in market %', 'Days in market %', 1),
                    ('Return on deployed capital %/yr', 'Return on deployed capital % per year', 1), ('Futures scenario: return on margin %/yr (20% margin)', 'Return on margin % per year - futures scenario at 20% margin', 0)):
    rows.append([lab] + [mget(nm('ALL', n, 2013), key, 0, d, sec='Utilisation table') for n in SIZES])
sc_row = ['Scaled to ~98% peak: CAGR % (idle 0% / 6.5%)']
for n in SIZES:
    a_ = other('Scaled_Positions', SCL, 'Strategy', f'ALL-Top{n}', 'CAGR % (idle cash 0%)', 2, sec='Scaled', extra={'Window': '2013-2025'}); b_ = other('Scaled_Positions', SCL, 'Strategy', f'ALL-Top{n}', 'CAGR % (idle cash 6.5%)', 2, sec='Scaled', extra={'Window': '2013-2025'})
    sc_row.append(f'{a_} / {b_}')
rows.append(sc_row)
mm_row = ['Scaled: max drawdown %']
for n in SIZES: mm_row.append(other('Scaled_Positions', SCL, 'Strategy', f'ALL-Top{n}', 'Max drawdown % (idle cash 0%)', 1, sec='Scaled', extra={'Window': '2013-2025'}))
rows.append(mm_row)
right += [tbl(rows, [50 * mm] + [(RW - 50 * mm) / 4] * 4, fs=5.9, pad=0.5)]
right += [P('6. Cost sensitivity: CAGR % with idle cash at 6.5% (ALL dates, 2013-2025)', H2)]
rows = [['Cost per position', 'Top 1', 'Top 3', 'Top 5', 'Top 10']]
for cst_ in (0.15, 0.30, 0.50):
    row = [f'{cst_:.2f}%' + (' (base)' if cst_ == 0.15 else '')]
    for n in SIZES: row.append(other('Cost_Sensitivity', CST, 'Strategy', f'ALL-Top{n}', 'CAGR_IdleCashRf_Pct', 2, sec='Cost sensitivity', extra={'Window': '2013-2025', 'CostPerPositionPct': cst_}))
    rows.append(row)
right += [tbl(rows, [50 * mm] + [(RW - 50 * mm) / 4] * 4, fs=5.9, pad=0.5)]
right += [P('7. What the measures mean', H2)]
defs = [['Measure', 'Meaning'],
        ['CAGR', 'Compound annual growth: (ending/starting capital)^(1/years) - 1; idle cash earning 0% or 6.5%.'],
        ['Fund utilised %', 'Capital in open positions / equity, averaged over ALL days (idle days = 0). Peak = busiest day.'],
        ['Return on deployed capital', 'Total P&amp;L / average rupees deployed, per year: what the money earned while working.'],
        ['Sharpe / Sortino / Calmar', 'Daily excess return over 6.5%, annualised, / volatility, / downside volatility, / CAGR over |max drawdown|.'],
        ['VaR / CVaR 95%', 'Daily loss beaten on 5% of invested days / average loss on those days.'],
        ['Beta, alpha, capture, IR', 'Daily regression vs Nifty 50; up/down capture = average return on Nifty up/down days relative to Nifty; IR = excess return per unit of tracking error.'],
        ['Profit factor', 'Total gains / total losses in rupees (all trades).'],
        ['Bootstrap CAGR', 'Yearly returns resampled 5,000 times: 5th-95th percentile of CAGR.']]
right += [tbl(defs, [28 * mm, RW - 28 * mm], fs=5.8, pad=0.4)]
right += [P('8. Limits', H2),
          P('13 test years (8 for the honest date filter); the 30 dates overlap in time. Survivorship bias: stock list = today\'s F&amp;O universe. Cash-settled P&amp;L at the close, flat 0.15% cost, no slippage, margin interest or tax; SHORT trades assumed possible at the same cost (they need futures); the futures return-on-margin line is a rough scenario only. '
            'HINDSIGHT rows use future information and are not achievable. Every figure in this report is traceable in <b>November_Portfolio_Report_Data.xlsx</b> (sheets Report_Numbers, Verification, All_Measures).', SMALL)]
E += [Table([[left, right]], colWidths=[LW + 2.5 * mm, RW + 2.5 * mm], style=TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP'), ('LEFTPADDING', (0, 0), (-1, -1), 0), ('RIGHTPADDING', (0, 0), (0, 0), 5 * mm), ('RIGHTPADDING', (1, 0), (1, 0), 0), ('TOPPADDING', (0, 0), (-1, -1), 0), ('BOTTOMPADDING', (0, 0), (-1, -1), 0)]))]
doc.build(E, onFirstPage=footer, onLaterPages=footer)

# ======================= checks =======================
import pypdf
rd = pypdf.PdfReader(OUT_PDF); npages = len(rd.pages); text = ' '.join(p.extract_text() for p in rd.pages); flat = re.sub(r'\s+', ' ', text)
checks = []
for pg_, free in sorted(LAYOUT.items()):
    checks.append(dict(Check=f'Layout: page {pg_} free space left (mm), must be >= 0', Detail='frame height remaining after the last element', Workbook=free, Independent=free, AbsDiff=0.0, Tolerance=0.0, Result='PASS' if free >= 0 else 'FAIL'))
print('free space per page (mm):', {k: round(v, 1) for k, v in LAYOUT.items()})
checks.append(dict(Check='Report has exactly 2 pages', Detail='PDF page count', Workbook=2, Independent=npages, AbsDiff=abs(npages - 2), Tolerance=0, Result='PASS' if npages == 2 else 'FAIL'))
# re-read every logged number from the workbook by an independent route (openpyxl cell lookup) and compare
wbk = openpyxl.load_workbook(PERF, data_only=True)
def lookup(r):
    ws = wbk[r['SourceSheet']]; hdr = [c.value for c in ws[1]]
    if r['SourceSheet'] == 'All_Measures':
        ci = hdr.index(r['SourceCol']) + 1; hits = [i for i in range(2, ws.max_row + 1) if ws.cell(i, 1).value == r['SourceRow']]; return ws.cell(hits[min(r['Nth'], len(hits) - 1)], ci).value
    if r['SourceSheet'] == 'Utilisation_By_Month':
        ci = hdr.index(r['SourceCol']) + 1
        for i in range(2, ws.max_row + 1):
            if ws.cell(i, 1).value == r['SourceRow']: return ws.cell(i, ci).value
    if r['SourceSheet'] == 'Yearly_Returns':
        ci = hdr.index(r['SourceCol']) + 1; yr = int(r['SourceRow'].split('=')[1])
        for i in range(2, ws.max_row + 1):
            if ws.cell(i, 1).value == yr: return ws.cell(i, ci).value
    if r['SourceSheet'] in ('Scaled_Positions', 'Cost_Sensitivity'):
        ci = hdr.index(r['SourceCol']) + 1; strat = r['SourceRow'].split('=')[1].split(' & ')[0]; win = '2013-2025'
        want_cost = None
        m = re.search(r"'CostPerPositionPct': ([0-9.]+)", r['SourceRow'])
        if m: want_cost = float(m.group(1))
        for i in range(2, ws.max_row + 1):
            if ws.cell(i, hdr.index('Strategy') + 1).value == strat and ws.cell(i, hdr.index('Window') + 1).value == win and (want_cost is None or abs(ws.cell(i, hdr.index('CostPerPositionPct') + 1).value - want_cost) < 1e-9): return ws.cell(i, ci).value
    return None
bad = 0; checked = 0; notfound = 0
for r in LOG:
    if r['SourceSheet'] in ('Equity_Daily',) or (isinstance(r['Raw'], float) and np.isnan(r['Raw'])): continue
    v = lookup(r); checked += 1
    if v is None: notfound += 1; continue
    if abs(float(v) - float(r['Raw'])) > 1e-9: bad += 1
checks.append(dict(Check='Every printed number re-read from the workbook (independent lookup)', Detail=f'{checked} numbers; lookups not found: {notfound}', Workbook=checked, Independent=checked - bad - notfound, AbsDiff=bad + notfound, Tolerance=0, Result='PASS' if (bad + notfound) == 0 else 'FAIL'))
# displayed string must appear in the PDF text at least as often as it was logged
from collections import Counter
cnt = Counter(r['Displayed'] for r in LOG if r['Displayed'] not in ('see sheet',) and not (isinstance(r['Raw'], float) and np.isnan(r['Raw'])) and r['Decimals'] != 4)
miss = 0
for s_, k in cnt.items():
    pat = r'(?<![0-9.])' + re.escape(s_) + r'(?![0-9])'
    if len(re.findall(pat, flat)) < k: miss += 1
checks.append(dict(Check='Every printed value is present in the PDF text', Detail=f'{len(cnt)} distinct values searched', Workbook=len(cnt), Independent=len(cnt) - miss, AbsDiff=miss, Tolerance=0, Result='PASS' if miss == 0 else 'FAIL'))
V = vp.run_checks()
V = pd.concat([pd.DataFrame(checks), V], ignore_index=True)
print('pages', npages, '| logged numbers', len(LOG), '| checks:', V.Result.value_counts().to_dict())

# ======================= Excel: performance workbook + audit sheets =======================
wb = openpyxl.load_workbook(PERF)
hdr = PatternFill('solid', fgColor='1F3864'); hf = Font(bold=True, color='FFFFFF')
def put(ws, df, widths):
    for j, c in enumerate(df.columns, 1):
        x = ws.cell(1, j, c); x.fill = hdr; x.font = hf; x.alignment = Alignment(wrap_text=True, vertical='center')
        ws.column_dimensions[get_column_letter(j)].width = widths[j - 1]
    for i, row in enumerate(df.itertuples(index=False), 2):
        for j, v in enumerate(row, 1):
            if isinstance(v, (np.floating,)): v = None if np.isnan(v) else float(v)
            elif isinstance(v, (np.integer,)): v = int(v)
            ws.cell(i, j, v)
    ws.freeze_panes = 'A2'; ws.auto_filter.ref = ws.dimensions
RN = pd.DataFrame(LOG)[['Section', 'Item', 'Displayed', 'Raw', 'SourceSheet', 'SourceRow', 'SourceCol']]
RN.columns = ['Report section', 'Item', 'Printed in report', 'Exact value', 'Source sheet', 'Source row', 'Source column']
ws = wb.create_sheet('Report_Numbers', 1); put(ws, RN, [20, 60, 18, 18, 22, 52, 38])
VS = V.copy(); ws2 = wb.create_sheet('Verification', 2); put(ws2, VS, [52, 56, 18, 18, 14, 12, 10])
for i in range(2, ws2.max_row + 1):
    c = ws2.cell(i, 7); c.font = Font(bold=True, color='1B5E20' if c.value == 'PASS' else 'B71C1C')
    for j in (3, 4, 5): ws2.cell(i, j).number_format = '0.000000'
note = wb['ReadMe']; nxt = note.max_row + 2
for t_ in ['', 'ADDED FOR THE 2-PAGE REPORT', 'Report_Numbers = every number printed in November_Portfolio_2Page_Report.pdf with its source sheet / row / column. Verification = independent checks (separate re-simulation from raw prices, trade-list integrity, picks vs earlier validated workbooks, PDF page count and text).',
           'Rebuild everything: python run_portfolio_pipeline.py ; python build_portfolio_excel.py ; python build_two_page_report.py']:
    note.cell(nxt, 1, t_); nxt += 1
wb.save(OUT_XL)
print('saved', OUT_PDF, OUT_XL)
print(V[['Check', 'Detail', 'Result']].groupby('Result').size().to_dict())
fails = V[V.Result != 'PASS']
if len(fails): print(fails.to_string())
