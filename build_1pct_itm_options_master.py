"""
BUILD: 12-QUARTER MASTER WORKBOOK — 1% ITM OPTIONS VERSION
==========================================================
Mirror of  GitHub_Dashboard_12_Quarters_Master.xlsx  (the FUTURES master), but the
strategy is expressed as the project's 1% ITM OPTION leg instead of futures:
  LONG  bias  ->  BUY 1% ITM CALL   (strike = spot_entry * 0.99)
  SHORT bias  ->  BUY 1% ITM PUT    (strike = spot_entry * 1.01)

It reuses the SAME real per-trade entry/exit spot prices, dates, windows, lots and
win-rates already audited in the futures master (600 trades = 50 stocks x 12 quarters)
and re-prices each as an option using the project's synthetic convention (identical to
backtest_nifty_index_options_1pct_itm_quarterly.py):
  entry premium = intrinsic(1% of spot) + 2.0% * spot_entry
  exit  premium = intrinsic_at_exit    + 0.5% * spot_exit
  option return = (exit_prem/entry_prem - 1) - 0.15% round-trip cost

ADDITIVE: reads the existing futures master; writes a NEW workbook. Nothing else touched.
"""
import openpyxl, re, argparse
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

SRC = r'D:\behaviour analysis\GitHub_Dashboard_12_Quarters_Master.xlsx'
OUT = r'D:\behaviour analysis\GitHub_Dashboard_12_Quarters_1PCT_ITM_Options_Master.xlsx'

# ---- synthetic 1% ITM option convention (extrinsic % of spot is configurable) ----
_ap = argparse.ArgumentParser()
_ap.add_argument('--entry-extrinsic', type=float, default=0.01, help='time value at entry, fraction of spot (default 0.01)')
_ap.add_argument('--exit-extrinsic', type=float, default=0.005, help='time value at exit, fraction of spot (default 0.005)')
_args = _ap.parse_args()
EXTRINSIC = _args.entry_extrinsic       # time value at entry (default 1.0% of spot)
EXTRINSIC_EXIT = _args.exit_extrinsic   # time value at exit  (default 0.5% of spot)
COST = 0.0015                           # 0.15% round-trip cost on the premium
_BREAKEVEN = (EXTRINSIC - EXTRINSIC_EXIT) / (1 + EXTRINSIC_EXIT) * 100  # ~ favourable move needed

def num(v):
    if v is None: return None
    if isinstance(v, (int, float)): return float(v)
    try: return float(str(v).replace(',', '').strip())
    except Exception: return None

def option_leg(side, spot_en, spot_ex):
    """Return dict with strike, entry/exit premium, return %, per-share P&L (net of cost)."""
    call = (side == 'LONG')
    strike = spot_en * (0.99 if call else 1.01)
    intr_en = max(0.0, (spot_en - strike) if call else (strike - spot_en))   # = 1% of spot_en
    intr_ex = max(0.0, (spot_ex - strike) if call else (strike - spot_ex))
    p_en = intr_en + spot_en * EXTRINSIC
    p_ex = intr_ex + spot_ex * EXTRINSIC_EXIT
    ret = ((p_ex - p_en) / p_en - COST) if p_en > 0 else 0.0
    return dict(opt_type=('CALL' if call else 'PUT'), strike=strike,
                p_en=p_en, p_ex=p_ex, ret=ret, pnl_share=(p_ex - p_en) - COST * p_en)

# ---------- read the futures master ----------
src = openpyxl.load_workbook(SRC, read_only=True, data_only=True)
sumrows = list(src['Dashboard_Summary'].iter_rows(values_only=True))
# map q_code -> (cycle, period) from the futures summary
meta = {}
for r in sumrows[4:16]:
    if r and r[0] and str(r[0]).startswith('FY'):
        meta[str(r[0]).strip()] = (str(r[1] or ''), str(r[2] or ''))

QCODES = [s for s in src.sheetnames if s != 'Dashboard_Summary']

def read_quarter(qcode):
    rows = list(src[qcode].iter_rows(values_only=True))
    hdr = rows[3]
    idx = {str(h).strip(): j for j, h in enumerate(hdr) if h is not None}
    def col(*names):
        for n in names:
            for k, j in idx.items():
                if k.lower().startswith(n.lower()): return j
        return None
    ci = dict(sym=col('Symbol'), name=col('Company'), bias=col('Futures Bias', 'Bias'),
              win=col('Position Window', 'Window'), ed=col('Entry Date'), ep=col('Entry Price'),
              rd=col('Result Date'), xd=col('Exit Date'), xp=col('Exit Price'),
              wr=col('Win Rate'), lot=col('Lot Size'))
    out = []
    for r in rows[4:]:
        if not r or r[ci['sym']] is None: continue
        ep, xp = num(r[ci['ep']]), num(r[ci['xp']])
        if ep is None or xp is None: continue
        out.append(dict(
            sym=str(r[ci['sym']]).strip(), name=str(r[ci['name']] or '').strip(),
            bias=str(r[ci['bias']] or 'LONG').strip().upper(),
            win=str(r[ci['win']] or '').strip(),
            ed=str(r[ci['ed']] or ''), rd=str(r[ci['rd']] or ''), xd=str(r[ci['xd']] or ''),
            ep=ep, xp=xp, wr=num(r[ci['wr']]) or 0.0, lot=int(num(r[ci['lot']]) or 0)))
    return out

# ---------- styles ----------
NAVY = PatternFill('solid', fgColor='1A202C'); WHITE = Font(bold=True, color='FFFFFF', size=10)
TITLE = Font(bold=True, color='1A202C', size=13); SUB = Font(color='4A5568', size=10, italic=True)
WINF = PatternFill('solid', fgColor='E6F4EA'); LOSSF = PatternFill('solid', fgColor='FCE8E6')
GOOD = Font(color='137333', bold=True); BAD = Font(color='C5221F', bold=True)
THIN = Border(*[Side(style='thin', color='CBD5E0')] * 4)
CEN = Alignment(horizontal='center'); RIGHT = Alignment(horizontal='right')

HEADERS = ['#', 'Symbol', 'Company', 'Option Bias', 'Position Window', 'Entry Date (09:20 AM)',
           'Spot Entry (Rs)', 'Result Date (T)', 'Exit Date (03:15 PM)', 'Spot Exit (Rs)',
           'Underlying Move %', 'Strike (1% ITM)', 'Entry Premium (Rs)', 'Exit Premium (Rs)',
           'Stock Win Rate (%)', 'Lot Size', 'Option Return %', 'Option P&L (1 Lot) (Rs)',
           'Premium Paid (Rs)', 'Outcome', 'Option Order Ticket']
WIDTHS = [5, 13, 30, 13, 15, 22, 14, 17, 22, 14, 15, 14, 16, 16, 15, 10, 15, 20, 17, 10, 24]

out = openpyxl.Workbook()
summary = out.active; summary.title = 'Dashboard_Summary'

qsummary = []  # (qcode, cycle, period, wr, w, l, pnl, prem, avg_ret)
for qcode in QCODES:
    trades = read_quarter(qcode)
    ws = out.create_sheet(qcode)
    cycle, period = meta.get(qcode, ('', ''))
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(HEADERS))
    ws.cell(1, 1, f'{qcode}  |  {cycle}  |  {period}  |  1% ITM OPTIONS').font = TITLE
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(HEADERS))

    tot_pnl = tot_prem = 0.0; wins = 0; rets = []
    for j, h in enumerate(HEADERS, 1):
        c = ws.cell(4, j, h); c.fill = NAVY; c.font = WHITE; c.border = THIN; c.alignment = CEN
    rownum = 5
    for i, t in enumerate(trades, 1):
        leg = option_leg(t['bias'], t['ep'], t['xp'])
        pnl = leg['pnl_share'] * t['lot']; prem = leg['p_en'] * t['lot']
        umove = (t['xp'] - t['ep']) / t['ep'] * (1 if t['bias'] == 'LONG' else -1) * 100
        win = leg['ret'] > 0
        wins += win; tot_pnl += pnl; tot_prem += prem; rets.append(leg['ret'] * 100)
        ticket = 'BUY 1% ITM CALL' if t['bias'] == 'LONG' else 'BUY 1% ITM PUT'
        vals = [i, t['sym'], t['name'], ('BUY CALL' if t['bias'] == 'LONG' else 'BUY PUT'),
                t['win'], t['ed'], round(t['ep'], 2), t['rd'], t['xd'], round(t['xp'], 2),
                round(umove, 2), round(leg['strike'], 2), round(leg['p_en'], 2), round(leg['p_ex'], 2),
                round(t['wr'] * 100, 1) if t['wr'] <= 1 else round(t['wr'], 1), t['lot'],
                round(leg['ret'] * 100, 2), round(pnl, 0), round(prem, 0),
                ('WIN' if win else 'LOSS'), ticket]
        fill = WINF if win else LOSSF
        for j, v in enumerate(vals, 1):
            c = ws.cell(rownum, j, v); c.border = THIN; c.fill = fill
            if j in (1, 4, 5, 11, 15, 17, 20): c.alignment = CEN
            if j in (7, 10, 12, 13, 14, 18, 19): c.alignment = RIGHT
            if j == 17: c.font = GOOD if win else BAD
        rownum += 1
    n = len(trades); wr = wins / n if n else 0
    ws.cell(2, 1, (f'Option Realised Win Rate: {wr*100:.1f}%  |  {wins}W / {n-wins}L  |  '
                   f'Net Option PnL: Rs. {tot_pnl:,.0f}  |  Premium Deployed: Rs. {tot_prem:,.0f}  |  '
                   f'Avg Option Return: {sum(rets)/n:+.1f}%')).font = SUB
    for j, w in enumerate(WIDTHS, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    ws.freeze_panes = 'A5'
    qsummary.append((qcode, cycle, period, wr, wins, n - wins, tot_pnl, tot_prem, sum(rets)/n if n else 0))

# ---------- Dashboard_Summary ----------
summary.merge_cells('A1:I1'); summary.merge_cells('A2:I2')
summary.cell(1, 1, 'SMC GLOBAL - NIFTY 50 EVENT INTELLIGENCE | 12 QUARTERS — 1% ITM OPTIONS SUMMARY').font = TITLE
summary.cell(2, 1, (f'Same 12-quarter universe as the futures master, priced as the 1% ITM option leg  |  '
                    f'Time value: {EXTRINSIC*100:.1f}% entry / {EXTRINSIC_EXIT*100:.1f}% exit  |  '
                    f'Break-even favourable move ~ +{_BREAKEVEN:.2f}%  |  {COST*100:.2f}% cost')).font = SUB
shead = ['Quarter Code', 'Earnings Event Cycle', 'Reporting Period', 'Status', 'Option Win Rate',
         'Wins / Losses', 'Net Option PnL (Rs)', 'Premium Deployed (Rs)', 'Avg Option Ret %']
for j, h in enumerate(shead, 1):
    c = summary.cell(4, j, h); c.fill = NAVY; c.font = WHITE; c.border = THIN; c.alignment = CEN
tw = tl = 0; tpnl = tprem = 0.0; allret = []
for i, (qc, cy, pe, wr, w, l, pnl, prem, avg) in enumerate(qsummary, 5):
    tw += w; tl += l; tpnl += pnl; tprem += prem; allret.append(avg)
    vals = [qc, cy, pe, 'COMPLETED', f'{wr*100:.1f}%', f'{w}W / {l}L',
            round(pnl, 0), round(prem, 0), f'{avg:+.1f}%']
    fill = WINF if pnl > 0 else LOSSF
    for j, v in enumerate(vals, 1):
        c = summary.cell(i, j, v); c.border = THIN; c.fill = fill
        if j in (4, 5, 6, 9): c.alignment = CEN
        if j in (7, 8): c.alignment = RIGHT
row = 5 + len(qsummary)
tot_n = tw + tl
vals = ['CONSOLIDATED TOTAL', '12 Quarters (600 option trades)', f'{QCODES[-1]} to {QCODES[0]}',
        '12 QUARTERS', f'{tw/tot_n*100:.1f}%', f'{tw}W / {tl}L', round(tpnl, 0), round(tprem, 0),
        f'{sum(allret)/len(allret):+.1f}%']
for j, v in enumerate(vals, 1):
    c = summary.cell(row, j, v); c.fill = NAVY; c.font = WHITE; c.border = THIN
    if j in (4, 5, 6, 9): c.alignment = CEN
    if j in (7, 8): c.alignment = RIGHT
for j, w in enumerate([16, 26, 18, 14, 15, 16, 20, 22, 16], 1):
    summary.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
summary.freeze_panes = 'A5'

out.save(OUT)
print('Saved:', OUT)
print(f'12 quarters | 600 option trades | overall win {tw/tot_n*100:.1f}% ({tw}W/{tl}L) | '
      f'net option PnL Rs {tpnl:,.0f} on premium Rs {tprem:,.0f}')
print('Per-quarter option win rate & PnL:')
for qc, cy, pe, wr, w, l, pnl, prem, avg in qsummary:
    print(f'  {qc:8} win {wr*100:5.1f}%  {w:2}W/{l:2}L  netPnL Rs {pnl:>12,.0f}  avgRet {avg:+6.1f}%')
