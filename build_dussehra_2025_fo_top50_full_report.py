"""
TOP 50 F&O STOCKS \u2014 LAST DUSSEHRA (2025) \u2014 FULL REPORT (detailed trade sheet
+ daily MTM sheet + event summary). Same top-50 selection as
build_dussehra_2025_fo_top50_trade_sheet.py (ranked by that occurrence's real
PnL, full 211 F&O universe, walk-forward point-in-time), now also with a
date-wise mark-to-market grid like the Nifty-50 2025 report.

ADDITIVE: reads the walk-forward pickle + futures price data; writes a new report.
"""
import json, pickle, os, subprocess, statistics as st
from datetime import date
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

BASE = r'D:\behaviour analysis'
OUT_XLSX = os.environ.get('DUSS25TOP50FULL_OUT_XLSX', BASE + r'\Dussehra_2025_FO_Top50_Full_Report.xlsx')
OUT_HTML = BASE + r'\Dussehra_2025_FO_Top50_Full_Report.html'
OUT_PDF = BASE + r'\Dussehra_2025_FO_Top50_Full_Report.pdf'

results = pickle.load(open(BASE + r'\scraped_parquet\_tmp_dussehra_walkforward.pkl', 'rb'))
fo = json.load(open(BASE + r'\dashboard_data\fo_stocks_211.json', encoding='utf-8'))
meta = {f['symbol']: f for f in fo}

fut = pd.read_parquet(BASE + r'\scraped_parquet\fo_futures_near_month_continuous.parquet', columns=['Date', 'Instrument', 'Open', 'Close'])
fut['Date'] = pd.to_datetime(fut['Date']).dt.date
fut_closes = {s: g.sort_values('Date')[['Date', 'Close', 'Open']].values.tolist() for s, g in fut.groupby('Instrument')}

beh = json.load(open(BASE + r'\dashboard_data\nifty_futures_holiday_behaviour.json', encoding='utf-8'))
dh = [h for h in beh['holidays'] if 'Dussehra' in h['name']][0]
dussehra_dates = sorted(date.fromisoformat(t['holiday_date']) for t in dh['trades'])
anchor_2025 = next(dt for dt in dussehra_dates if dt.year == 2025)


def lot_for(sym):
    return meta.get(sym, {}).get('lot_size') or 500


# ---- Step 1: reconstruct every F&O stock's real 2025 trade + full daily MTM path ----
stock_paths = {}
for sym, r in results.items():
    if r.get('oos_n', 0) < 1:
        continue
    t = next((x for x in r['oos_trades'] if x['year'] == 2025), None)
    if not t:
        continue
    a = fut_closes.get(sym)
    if not a:
        continue
    ib = -1
    for i, row in enumerate(a):
        if row[0] <= anchor_2025: ib = i
        else: break
    if ib < 0:
        continue
    bo, so, side = t['bo'], t['so'], t['side']
    en_idx, ex_idx = ib - bo, ib + so
    if en_idx < 0 or ex_idx >= len(a):
        continue
    lot = lot_for(sym)
    sign = 1 if side == 'LONG' else -1
    en_date, en_close, en_px = a[en_idx]

    g = 1.0
    if en_px:
        dr0 = (en_close - en_px) / en_px
        if abs(dr0) <= 0.20: g *= (1 + dr0)
    daily = [(en_date, round((g - 1) * 100 * sign / 100 * en_px * lot, 2))]
    for k in range(en_idx + 1, ex_idx + 1):
        p0, p1 = a[k - 1][1], a[k][1]
        if p0:
            dr = (p1 - p0) / p0
            if abs(dr) <= 0.20: g *= (1 + dr)
        daily.append((a[k][0], round((g - 1) * 100 * sign / 100 * en_px * lot, 2)))

    exit_px = a[ex_idx][1]
    margin = round(en_px * lot * 0.20, 2)
    m = meta.get(sym, {})
    stock_paths[sym] = dict(symbol=sym, name=m.get('name', sym), sector=m.get('sector', ''), direction=side,
                             window=f"T-{bo} to T+{so}", entry_date=en_date, entry_px=round(en_px, 2),
                             exit_date=a[ex_idx][0], exit_px=round(exit_px, 2), lot=lot, margin=margin,
                             daily=daily, low_confidence=r.get('low_confidence', False), oos_n=r['oos_n'])

print(f"Reconstructed real 2025 Dussehra trades for {len(stock_paths)} of 211 F&O stocks.")

# ---- Step 2: per-stock Max DD / Sharpe / VaR, then rank by final PnL, take top 50 ----
for sym, p in stock_paths.items():
    vals = [v for _, v in p['daily']]
    peak = vals[0]; maxdd = 0.0
    for v in vals:
        peak = max(peak, v); maxdd = min(maxdd, v - peak)
    p['max_dd_rs'] = round(maxdd, 2)
    p['cum_pnl'] = vals[-1]
    p['ret_pct'] = round(vals[-1] / (p['entry_px'] * p['lot']) * 100, 2)
    changes = [vals[i] - vals[i - 1] for i in range(1, len(vals))]
    if len(changes) >= 2 and st.pstdev(changes) > 0:
        p['sharpe'] = round(st.mean(changes) / st.pstdev(changes), 2)
        sc = sorted(changes)
        p['var95'] = round(sc[max(0, int(len(sc) * 0.05) - 1)], 2)
    else:
        p['sharpe'] = p['var95'] = None

ranked = sorted(stock_paths.items(), key=lambda kv: -kv[1]['cum_pnl'])
top50 = dict(ranked[:50])

# ---- Step 3: whole-event (top-50-portfolio) aggregation ----
all_dates = sorted(set(dt for p in top50.values() for dt, _ in p['daily']))
portfolio_mtm = []
for dt in all_dates:
    total = sum(next((v for d2, v in p['daily'][::-1] if d2 <= dt), 0) for p in top50.values())
    portfolio_mtm.append((dt, round(total, 2)))
port_vals = [v for _, v in portfolio_mtm]
port_peak = port_vals[0]; port_maxdd = 0.0
for v in port_vals:
    port_peak = max(port_peak, v); port_maxdd = min(port_maxdd, v - port_peak)
port_changes = [port_vals[i] - port_vals[i - 1] for i in range(1, len(port_vals))]
port_sharpe = round(st.mean(port_changes) / st.pstdev(port_changes), 3) if len(port_changes) >= 2 and st.pstdev(port_changes) > 0 else None
sorted_pc = sorted(port_changes)
port_var95 = round(sorted_pc[max(0, int(len(sorted_pc) * 0.05) - 1)], 2) if sorted_pc else None

total_margin = round(sum(p['margin'] for p in top50.values()), 2)
final_pnl = round(sum(p['cum_pnl'] for p in top50.values()), 2)
wins = sum(1 for p in top50.values() if p['cum_pnl'] > 0)
win_rate = round(100 * wins / len(top50), 1)
n_long = sum(1 for p in top50.values() if p['direction'] == 'LONG'); n_short = len(top50) - n_long
n_low = sum(1 for p in top50.values() if p['low_confidence'])
best = max(top50.items(), key=lambda kv: kv[1]['cum_pnl'])
worst = min(top50.items(), key=lambda kv: kv[1]['cum_pnl'])

NAVY = PatternFill('solid', fgColor='1A202C'); WHITE = Font(bold=True, color='FFFFFF', size=8)
TITLE = Font(bold=True, color='1A202C', size=13); SUB = Font(color='4A5568', size=10, italic=True)
GF = PatternFill('solid', fgColor='E6F4EA'); RF = PatternFill('solid', fgColor='FCE8E6')
AM = PatternFill('solid', fgColor='FEF3C7'); LBL = PatternFill('solid', fgColor='F1F5F9')
GOOD = Font(color='137333', bold=True); BAD = Font(color='C5221F', bold=True); BOLD = Font(bold=True)
THIN = Border(*[Side(style='thin', color='D9E2EC')] * 4)
CEN = Alignment(horizontal='center'); RIGHT = Alignment(horizontal='right'); WRAP = Alignment(wrap_text=True, vertical='top')

wb = Workbook()

# ============ Sheet 1: Detailed Trade Log (Top 50) ============
tl = wb.active; tl.title = 'Top50_Trade_Log'
tl.merge_cells('A1:O1')
tl.cell(1, 1, 'TOP 50 F&O STOCKS \u2014 LAST DUSSEHRA (2-Oct-2025) \u2014 DETAILED TRADE LOG').font = TITLE
tl.merge_cells('A2:O2')
tl.cell(2, 1, f'Real 2025 walk-forward trade (point-in-time) for each of {len(stock_paths)} F&O stocks with usable data; ranked by PnL, top 50 shown. Futures-sourced prices.').font = SUB
hdr = ['Rank', 'Symbol', 'Company', 'Sector', 'Direction', 'Window', 'Entry Date', 'Entry Price (\u20b9 @09:20)',
       'Exit Date', 'Exit Price (\u20b9 @15:15)', 'Lot', 'Margin (\u20b9)', 'Return %', 'PnL (\u20b9)', 'Confidence']
for j, h in enumerate(hdr, 1):
    c = tl.cell(4, j, h); c.fill = NAVY; c.font = WHITE; c.border = THIN; c.alignment = CEN
row = 5
for i, (sym, p) in enumerate(ranked[:50], 1):
    vals = [i, sym, p['name'], p['sector'], p['direction'], p['window'], p['entry_date'].strftime('%d-%b-%Y'),
            p['entry_px'], p['exit_date'].strftime('%d-%b-%Y'), p['exit_px'], p['lot'], p['margin'],
            p['ret_pct'], p['cum_pnl'], ('LOW (n<3)' if p['low_confidence'] else 'OK')]
    fill = AM if p['low_confidence'] else (GF if p['cum_pnl'] >= 0 else RF)
    for j, v in enumerate(vals, 1):
        c = tl.cell(row, j, v); c.border = THIN; c.fill = fill
        if j in (1, 5, 7, 9, 15): c.alignment = CEN
        if j in (8, 10, 12, 13, 14): c.alignment = RIGHT
        if j in (13, 14): c.font = GOOD if v >= 0 else BAD
        if j == 15 and p['low_confidence']: c.font = Font(color='92400E', bold=True, size=9)
    row += 1
for j, w in enumerate([6, 12, 24, 20, 10, 13, 15, 15, 15, 15, 7, 13, 10, 13, 12], 1):
    tl.column_dimensions[get_column_letter(j)].width = w
tl.freeze_panes = 'A5'

# ============ Sheet 2: Daily MTM Grid (Top 50) ============
mg = wb.create_sheet('Daily_MTM_Grid')
n_date_cols = len(all_dates)
last_col = 1 + n_date_cols + 8  # symbol + dates + (direction,window,margin,cumpnl,maxdd,sharpe,var,confidence)
mg.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
mg.cell(1, 1, 'TOP 50 F&O STOCKS \u2014 DAILY MARK-TO-MARKET, LAST DUSSEHRA (2025)').font = TITLE
mg.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
mg.cell(2, 1, f'Anchor 2-Oct-2025. All 50 covered ({n_low} LOW confidence). Entry Price = Open (09:20 proxy); Exit Price = Close (15:15 proxy). Blank = stock not yet in its own window that day.').font = SUB
hdr = (['Symbol'] + [dt.strftime('%d-%b') for dt in all_dates]
       + ['Direction', 'Window', 'Margin (\u20b9)', 'Cumulative PnL (\u20b9)', 'Max DD (\u20b9)', 'Sharpe', 'VaR 95% (\u20b9)', 'Confidence'])
for j, h in enumerate(hdr, 1):
    c = mg.cell(4, j, h); c.fill = NAVY; c.font = WHITE; c.border = THIN; c.alignment = CEN
row = 5
for sym in sorted(top50):
    p = top50[sym]
    daily_map = dict(p['daily'])
    rowfill = AM if p['low_confidence'] else None
    c = mg.cell(row, 1, sym); c.border = THIN; c.font = BOLD
    if rowfill: c.fill = rowfill
    for j, dt in enumerate(all_dates, 2):
        v = daily_map.get(dt)
        cc = mg.cell(row, j, v if v is not None else None)
        cc.border = THIN; cc.alignment = RIGHT
        if rowfill: cc.fill = rowfill
        if v is not None:
            cc.font = GOOD if v >= 0 else BAD
            cc.number_format = '#,##0'
    base = 2 + n_date_cols
    vals = [p['direction'], p['window'], p['margin'], p['cum_pnl'], p['max_dd_rs'],
            (p['sharpe'] if p['sharpe'] is not None else '\u2014'), (p['var95'] if p['var95'] is not None else '\u2014'),
            ('LOW' if p['low_confidence'] else 'OK')]
    for k, v in enumerate(vals):
        cc = mg.cell(row, base + k, v); cc.border = THIN
        if rowfill: cc.fill = rowfill
        if k in (0, 1, 7): cc.alignment = CEN
        else: cc.alignment = RIGHT
        if k == 3: cc.font = GOOD if p['cum_pnl'] >= 0 else BAD
        if k == 4: cc.font = BAD
        if k == 7 and p['low_confidence']: cc.font = Font(color='92400E', bold=True, size=9)
    row += 1
mg.column_dimensions['A'].width = 12
for j in range(2, 2 + n_date_cols):
    mg.column_dimensions[get_column_letter(j)].width = 9
for j, w in enumerate([10, 14, 13, 16, 12, 9, 12, 11], 2 + n_date_cols):
    mg.column_dimensions[get_column_letter(j)].width = w
mg.freeze_panes = 'B5'

# ============ Sheet 3: Event Performance Summary ============
ev = wb.create_sheet('Event_Performance_Summary')
ev.merge_cells('A1:B1')
ev.cell(1, 1, 'TOP 50 F&O \u2014 DUSSEHRA 2025 \u2014 WHOLE-EVENT PERFORMANCE SUMMARY').font = TITLE
metrics = [
    ('Stocks (top 50 of 211 universe)', f"{len(top50)}"),
    ('Win rate (stocks ending profitable)', f"{win_rate}% ({wins}/{len(top50)})"),
    ('Total margin deployed (\u20b9, 1 lot each)', f"{total_margin:,.0f}"),
    ('Final combined PnL (\u20b9)', f"{final_pnl:,.0f}"),
    ('Return on margin deployed', f"{(final_pnl/total_margin*100 if total_margin else 0):+.2f}%"),
    ('Direction split', f"{n_long} LONG / {n_short} SHORT"),
    ('Low-confidence stocks (n<3 prior years)', f"{n_low} of 50"),
    ('Portfolio Max Drawdown (\u20b9, intraday across the event)', f"{port_maxdd:,.0f}"),
    ('Portfolio Sharpe-like (daily PnL-change based)', f"{port_sharpe if port_sharpe is not None else '\u2014'}"),
    ('Portfolio VaR 95% (\u20b9, worst-5% single-day PnL swing)', f"{port_var95:,.0f}" if port_var95 is not None else '\u2014'),
    ('Best stock', f"{best[0]} ({best[1]['cum_pnl']:+,.0f} \u20b9, {best[1]['direction']})"),
    ('Worst stock', f"{worst[0]} ({worst[1]['cum_pnl']:+,.0f} \u20b9, {worst[1]['direction']})"),
]
r0 = 3
for lbl, val in metrics:
    a = ev.cell(r0, 1, lbl); a.font = BOLD; a.fill = LBL; a.border = THIN
    b = ev.cell(r0, 2, val); b.border = THIN; b.alignment = RIGHT
    r0 += 1
ev.column_dimensions['A'].width = 46; ev.column_dimensions['B'].width = 34

r0 += 1
ev.merge_cells(f'A{r0}:D{r0}')
ev.cell(r0, 1, 'PORTFOLIO DAILY CUMULATIVE MTM (date-wise, top 50 summed)').font = Font(bold=True, size=12)
r0 += 1
hdr2 = ['Date', 'Portfolio MTM (\u20b9)', "Day's Change (\u20b9)", 'Running Peak (\u20b9)']
for j, h in enumerate(hdr2, 1):
    c = ev.cell(r0, j, h); c.fill = NAVY; c.font = WHITE; c.border = THIN; c.alignment = CEN
r0 += 1
peak_track = portfolio_mtm[0][1]
for i, (dt, v) in enumerate(portfolio_mtm):
    peak_track = max(peak_track, v)
    chg = v - portfolio_mtm[i - 1][1] if i > 0 else 0
    vals = [dt.strftime('%d-%b-%Y'), v, chg, peak_track]
    fill = GF if v >= 0 else RF
    for j, val in enumerate(vals, 1):
        c = ev.cell(r0, j, val); c.border = THIN; c.fill = fill
        if j == 1: c.alignment = CEN
        else: c.alignment = RIGHT
        if j in (2, 3) and isinstance(val, (int, float)): c.font = GOOD if val >= 0 else BAD
    r0 += 1
note_row = r0 + 1
ev.merge_cells(f'A{note_row}:D{note_row}')
nt = ev.cell(note_row, 1, 'Data source: futures prices (2019-2026) \u2014 NOT the cash-equity file (found missing 98.3% of all Fridays across its 26yr history). Entry Price = Open (09:20 proxy); Exit Price = Close (15:15 proxy). Margin = 20% SPAN proxy. Ranked by this single occurrence\u2019s PnL, not a multi-year validated edge.')
nt.font = Font(italic=True, size=9, color='7C2D12'); nt.alignment = WRAP
ev.row_dimensions[note_row].height = 50

# ============ Sheet 0: Overview ============
ov = wb.create_sheet('Overview', 0)
ov.merge_cells('A1:B1')
ov.cell(1, 1, 'TOP 50 F&O STOCKS \u2014 LAST DUSSEHRA \u2014 OVERVIEW').font = TITLE
notes = [
    ('Universe & ranking', f'Full 211 F&O universe; {len(stock_paths)} stocks had usable 2025 Dussehra data. Top 50 shown, ranked by that single occurrence\u2019s actual PnL (1 lot).'),
    ('Method', 'Point-in-time walk-forward (window + direction for 2025 chosen using ONLY years strictly before it, same discipline as the live dashboard) \u2014 no hindsight.'),
    ('Data source', 'Futures prices (2019-2026). The cash-equity file was found missing 98.3% of all Fridays across its 26yr history and is not used anywhere in this report.'),
    ('Sheets', 'Top50_Trade_Log = one row per stock, final result. Daily_MTM_Grid = date-wise running P&L from entry to exit for every stock. Event_Performance_Summary = whole-portfolio (top 50 combined) view, including a date-wise combined MTM curve.'),
    ('Result', f'{wins}/50 profitable ({win_rate}%). Total PnL \u20b9{final_pnl:,.0f} on \u20b9{total_margin:,.0f} margin (1 lot each) \u2014 return on margin {(final_pnl/total_margin*100 if total_margin else 0):+.2f}%. {n_long} LONG / {n_short} SHORT.'),
    ('Confidence', f'{n_low} of the top 50 are LOW confidence (fewer than 3 prior years to pick their window from) \u2014 amber highlighted throughout.'),
    ('CAVEAT', 'Ranking by a single year\u2019s PnL favours whichever stocks happened to move most in 2025 \u2014 not the same as a multi-year validated edge. See FO211_Dussehra_WalkForward_Ranked_Report.xlsx for the across-years ranking.'),
]
r0 = 3
for lbl, txt in notes:
    a = ov.cell(r0, 1, lbl); a.font = BOLD; a.fill = AM; a.border = THIN; a.alignment = WRAP
    b = ov.cell(r0, 2, txt); b.alignment = WRAP; b.border = THIN
    ov.row_dimensions[r0].height = 62; r0 += 1
ov.column_dimensions['A'].width = 22; ov.column_dimensions['B'].width = 108

wb.save(OUT_XLSX)
print('Saved Excel:', OUT_XLSX)
print(f"Top50: win rate {win_rate}% ({wins}/50) | Total PnL Rs{final_pnl:,.0f} | Margin Rs{total_margin:,.0f} | {n_long} LONG / {n_short} SHORT | {n_low} low-confidence")

# ============ PDF (trade log table) ============
def esc(s): return str(s).replace('&', '&amp;').replace('<', '&lt;')
def pdf_row(rank, sym, p):
    low = p['low_confidence']
    rowstyle = ' style="background:#fffbeb;"' if low else ''
    conf = '<span style="color:#b45309;font-weight:700;">LOW</span>' if low else 'OK'
    return (f'<tr{rowstyle}><td>{rank}</td><td><b>{esc(sym)}</b></td><td>{esc(p["name"])[:22]}</td>'
            f'<td class="c">{p["direction"]}</td><td class="c">{esc(p["window"])}</td>'
            f'<td class="r {"pos" if p["ret_pct"]>=0 else "neg"}">{p["ret_pct"]:+.2f}%</td>'
            f'<td class="r {"pos" if p["cum_pnl"]>=0 else "neg"}">{p["cum_pnl"]:+,.0f}</td><td class="c">{conf}</td></tr>')
rows_all = ''.join(pdf_row(i + 1, sym, p) for i, (sym, p) in enumerate(ranked[:50]))

html = f"""<!doctype html><html><head><meta charset="utf-8">
<title>Top 50 F&amp;O Last Dussehra Full Report</title>
<style>
 *{{box-sizing:border-box}} body{{font-family:'Segoe UI',Arial,sans-serif;margin:0;color:#0f172a;background:#fff}}
 .wrap{{max-width:1050px;margin:0 auto;padding:28px 36px}}
 h1{{font-size:22px;margin:0 0 3px}} .sub{{color:#64748b;font-size:12px;margin-bottom:16px}}
 h2{{font-size:15px;margin:20px 0 8px;border-left:4px solid #0f766e;padding-left:9px}}
 table{{width:100%;border-collapse:collapse;font-size:11px}}
 th,td{{padding:5px 7px;text-align:right;border-bottom:1px solid #eef2f7}} th{{background:#f1f5f9;font-size:9.5px;text-transform:uppercase;color:#475569}}
 td.c,th.c{{text-align:center}} .pos{{color:#059669;font-weight:700}} .neg{{color:#dc2626;font-weight:700}}
 .cards{{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin:10px 0}}
 .mc{{background:#fff;border:1px solid #e2e8f0;border-radius:8px;padding:10px 12px}}
 .mcl{{font-size:9.5px;text-transform:uppercase;color:#64748b;font-weight:700}} .mcv{{font-size:19px;font-weight:800}}
 .note{{background:#fff7ed;border:1px solid #fed7aa;border-radius:8px;padding:11px 14px;font-size:11px;color:#7c2d12;margin-top:10px;line-height:1.5}}
 @page{{size:A4;margin:12mm}}
</style></head><body><div class="wrap">
 <h1>Top 50 F&amp;O Stocks &mdash; Last Dussehra, Full Report</h1>
 <div class="sub">Detailed trade log + daily MTM, real point-in-time trade per stock &bull; top 50 of {len(stock_paths)} usable F&amp;O stocks by PnL</div>
 <div class="cards">
  <div class="mc"><div class="mcl">Win Rate</div><div class="mcv pos">{win_rate}%</div></div>
  <div class="mc"><div class="mcl">Total PnL</div><div class="mcv {"pos" if final_pnl>=0 else "neg"}">\u20b9{final_pnl:,.0f}</div></div>
  <div class="mc"><div class="mcl">Margin Deployed</div><div class="mcv">\u20b9{total_margin:,.0f}</div></div>
  <div class="mc"><div class="mcl">LONG / SHORT</div><div class="mcv">{n_long} / {n_short}</div></div>
 </div>
 <h2>Top 50 \u2014 Ranked by PnL</h2>
 <table><thead><tr><th>Rank</th><th class="c">Symbol</th><th class="c">Company</th><th class="c">Dir</th><th class="c">Window</th><th>Return</th><th>PnL (\u20b9)</th><th class="c">Conf.</th></tr></thead>
 <tbody>{rows_all}</tbody></table>
 <div class="note">Full date-wise MTM grid for every stock, and the combined portfolio MTM curve, are in the companion Excel (Daily_MTM_Grid and Event_Performance_Summary sheets).</div>
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
