"""
TOP 50 F&O STOCKS — DUSSEHRA PERFORMANCE REPORT (Excel + PDF).
Ranks the 195 F&O stocks with usable Dussehra backtest data by average return
(same metric/method as the all-211 report) and takes the top 50. For each,
computes full trade-level performance measures on its yearly Dussehra-window
equity curve: CAGR, Max Drawdown, Sharpe-like ratio (mean/stdev of yearly
returns), best/worst year, plus the upcoming 20-Oct-2026 position (entry/exit
dates, direction, lot) using the same live-dashboard-consistent method as the
all-211 report's Position_Schedule_2026 sheet.

ADDITIVE: reads the existing intermediate pickle + live dashboard data; writes
new report files only.
"""
import json, pickle, os, subprocess, statistics as st
from datetime import date, timedelta, datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

BASE = r'D:\behaviour analysis'
OUT_XLSX = os.environ.get('FO211_TOP50_OUT_XLSX', BASE + r'\FO211_Top50_Dussehra_Performance_Report.xlsx')
OUT_HTML = BASE + r'\FO211_Top50_Dussehra_Performance_Report.html'
OUT_PDF = BASE + r'\FO211_Top50_Dussehra_Performance_Report.pdf'

data = pickle.load(open(BASE + r'\scraped_parquet\_tmp_fo211_dussehra.pkl', 'rb'))
results, meta = data['results'], data['meta']
qualified = {s: r for s, r in results.items() if r['pick']}
ranked_all = sorted(qualified.items(), key=lambda kv: -(kv[1]['pick_avg'] if kv[1]['pick_avg'] is not None else -999))
top50 = ranked_all[:50]
n_low_conf_50 = sum(1 for _, r in top50 if r.get('low_confidence'))
ALL_YEARS = sorted(set(y for _, r in top50 for y in r['year_rets']))

# ---- upcoming Dussehra 2026 position dates (identical method as the all-211 report) ----
NSE_2026_HOLIDAYS = {
    # '2026-03-04' removed (not a real NSE holiday -- leftover from Holi's date
    # once being wrongly typed as 04-Mar; real Holi holiday is 03-Mar, kept
    # below). '2026-11-08'/'2026-11-10' (Diwali) added -- were missing entirely.
    '2026-01-26', '2026-03-03', '2026-03-26', '2026-03-31', '2026-04-03',
    '2026-04-14', '2026-05-01', '2026-05-28', '2026-06-26', '2026-09-14', '2026-10-02',
    '2026-10-20', '2026-11-08', '2026-11-10', '2026-11-24', '2026-12-25',
}
def is_trading_day(d):
    return d.weekday() < 5 and d.isoformat() not in NSE_2026_HOLIDAYS
def shift_trading_days(anchor, n):
    d = anchor
    step = 1 if n >= 0 else -1
    cnt = abs(n)
    while cnt > 0:
        d += timedelta(days=step)
        if is_trading_day(d): cnt -= 1
    return d
DUSSEHRA_ANCHOR_2026 = date(2026, 10, 20)

_live = json.load(open(BASE + r'\dashboard_data\event_dashboard_data.json', encoding='utf-8'))
_live_dussehra = [h for h in _live['holidays'] if h['id'] == 'dussehra_2026'][0]
LIVE_POSITIONS = {s['symbol']: s for s in _live_dussehra['stocks']}

def position_for(sym, r):
    if sym in LIVE_POSITIONS:
        s = LIVE_POSITIONS[sym]
        return dict(direction=s['direction'], entry=s['entry_date'], exit=s['exit_date'], lot=s.get('lot'))
    en = shift_trading_days(DUSSEHRA_ANCHOR_2026, -r['lead'])
    ex = shift_trading_days(DUSSEHRA_ANCHOR_2026, r['hold'])
    return dict(direction=r['pick'], entry=en.strftime('%d-%b-%Y'), exit=ex.strftime('%d-%b-%Y'), lot=r.get('lot'))

# ---- performance measures on the yearly Dussehra-trade equity curve ----
def perf_measures(r):
    sign = 1 if r['pick'] == 'LONG' else -1
    years_sorted = sorted(r['year_rets'].keys())
    signed = [r['year_rets'][y] * sign for y in years_sorted]
    n = len(signed)
    cum = 1.0
    for v in signed: cum *= (1 + v / 100)
    cagr = (cum ** (1 / n) - 1) * 100 if n > 0 and cum > 0 else None
    eq = [100.0]
    for v in signed: eq.append(eq[-1] * (1 + v / 100))
    peak = eq[0]; maxdd = 0.0
    for v in eq[1:]:
        peak = max(peak, v)
        dd = (v - peak) / peak * 100
        maxdd = min(maxdd, dd)
    sharpe = (st.mean(signed) / st.pstdev(signed)) if n > 1 and st.pstdev(signed) > 0 else None
    best_i = signed.index(max(signed)); worst_i = signed.index(min(signed))
    return dict(cagr=cagr, maxdd=maxdd, sharpe=sharpe, best_year=years_sorted[best_i], best_ret=signed[best_i],
                worst_year=years_sorted[worst_i], worst_ret=signed[worst_i], n=n)

NAVY = PatternFill('solid', fgColor='1A202C'); WHITE = Font(bold=True, color='FFFFFF', size=9)
TITLE = Font(bold=True, color='1A202C', size=13); SUB = Font(color='4A5568', size=10, italic=True)
GF = PatternFill('solid', fgColor='E6F4EA'); RF = PatternFill('solid', fgColor='FCE8E6')
AM = PatternFill('solid', fgColor='FEF3C7'); LBL = PatternFill('solid', fgColor='F1F5F9')
GOOD = Font(color='137333', bold=True); BAD = Font(color='C5221F', bold=True); BOLD = Font(bold=True)
THIN = Border(*[Side(style='thin', color='D9E2EC')] * 4)
CEN = Alignment(horizontal='center'); RIGHT = Alignment(horizontal='right'); WRAP = Alignment(wrap_text=True, vertical='top')

wb = Workbook()

# ============ Sheet 1: Performance Measures (Top 50) ============
pm = wb.active; pm.title = 'Top50_Performance_Measures'
pm.merge_cells('A1:N1')
pm.cell(1, 1, 'TOP 50 F&O STOCKS \u2014 DUSSEHRA PERFORMANCE MEASURES (ranked by avg return)').font = TITLE
pm.merge_cells('A2:N2')
pm.cell(2, 1, 'Ranking/method identical to the all-211 report. CAGR & Max DD computed on each stock\u2019s own yearly Dussehra-trade equity curve (not calendar-year price return). Sharpe-like = mean / population-stdev of yearly returns.').font = SUB
hdr = ['Rank', 'Symbol', 'Company', 'Sector', 'Direction', 'Win Rate %', 'Avg Return %', 'CAGR %', 'Max DD %', 'Sharpe-like', 'Best Yr', 'Worst Yr', 'Years', 'Confidence']
for j, hh in enumerate(hdr, 1):
    c = pm.cell(4, j, hh); c.fill = NAVY; c.font = WHITE; c.border = THIN; c.alignment = CEN
row = 5
perf_cache = {}
for i, (sym, r) in enumerate(top50, 1):
    m = meta.get(sym, {})
    p = perf_measures(r); perf_cache[sym] = p
    low = r.get('low_confidence')
    vals = [i, sym, m.get('name', sym), m.get('sector', ''), r['pick'], r['pick_wr'], round(r['pick_avg'], 2),
            (round(p['cagr'], 2) if p['cagr'] is not None else '\u2014'), round(p['maxdd'], 2),
            (round(p['sharpe'], 2) if p['sharpe'] is not None else '\u2014'),
            f"{p['best_year']} ({p['best_ret']:+.1f}%)", f"{p['worst_year']} ({p['worst_ret']:+.1f}%)",
            r['n'], ('LOW (n<3)' if low else 'OK')]
    fill = AM if low else (GF if r['pick_avg'] >= 0 else RF)
    for j, v in enumerate(vals, 1):
        c = pm.cell(row, j, v); c.border = THIN; c.fill = fill
        if j in (1, 5, 13, 14): c.alignment = CEN
        if j in (6, 7, 8, 9, 10): c.alignment = RIGHT
        if j == 7: c.font = GOOD if r['pick_avg'] >= 0 else BAD
        if j == 9: c.font = BAD if p['maxdd'] < 0 else GOOD
    row += 1
for j, w in enumerate([6, 12, 26, 20, 10, 11, 12, 9, 10, 11, 15, 15, 7, 12], 1):
    pm.column_dimensions[get_column_letter(j)].width = w
pm.freeze_panes = 'A5'

# ============ Sheet 2: Upcoming Position Schedule (Top 50) ============
ps = wb.create_sheet('Position_Schedule_2026')
ps.merge_cells('A1:K1')
ps.cell(1, 1, 'DUSSEHRA 20-Oct-2026 \u2014 Position Schedule, Top 50 F&O Stocks by Avg Return').font = TITLE
ps.merge_cells('A2:K2')
ps.cell(2, 1, f'Entry/exit on the NSE-2026 trading calendar; Nifty-50 names use the exact live-dashboard dates. Sorted by entry date. {n_low_conf_50} of these 50 are LOW confidence (n<3 yrs, amber) \u2014 check before sizing.').font = SUB
hdr = ['Symbol', 'Company', 'Direction', 'Window', 'Entry Date', 'Exit Date', 'Lot Size', 'Win Rate %', 'Avg Return %', 'Years', 'Confidence']
for j, hh in enumerate(hdr, 1):
    c = ps.cell(4, j, hh); c.fill = NAVY; c.font = WHITE; c.border = THIN; c.alignment = CEN
sched = []
for sym, r in top50:
    pos = position_for(sym, r)
    m = meta.get(sym, {})
    sort_key = datetime.strptime(pos['entry'].split(' (')[0].strip(), '%d-%b-%Y')
    sched.append((sort_key, sym, m.get('name', sym), pos['direction'], f"T-{r['lead']} to T+{r['hold']}",
                  pos['entry'], pos['exit'], pos['lot'], r['pick_wr'], round(r['pick_avg'], 2), r['n'], r.get('low_confidence')))
sched.sort(key=lambda t: (t[0], t[1]))
row = 5
for _, sym, name, direc, window, entry, exit_, lot, wr, avg, n, low in sched:
    vals = [sym, name, direc, window, entry, exit_, lot, wr, avg, n, ('LOW (n<3)' if low else 'OK')]
    fill = AM if low else (GF if direc == 'LONG' else RF)
    for j, v in enumerate(vals, 1):
        c = ps.cell(row, j, v); c.border = THIN; c.fill = fill
        if j in (1, 3, 4, 5, 6, 10, 11): c.alignment = CEN
        if j in (7, 8, 9): c.alignment = RIGHT
        if j == 11 and low: c.font = Font(color='92400E', bold=True, size=9)
    row += 1
for j, w in enumerate([13, 28, 10, 15, 16, 16, 9, 11, 12, 7, 12], 1):
    ps.column_dimensions[get_column_letter(j)].width = w
ps.freeze_panes = 'A5'

# ============ Sheet 3: Year-by-Year Detail (Top 50) ============
yw = wb.create_sheet('Year_by_Year_Detail')
yw.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6 + len(ALL_YEARS))
yw.cell(1, 1, 'TOP 50 \u2014 STOCK x YEAR, Dussehra return %% using each stock\u2019s chosen direction').font = TITLE
yw.merge_cells(start_row=2, start_column=1, end_row=2, end_column=6 + len(ALL_YEARS))
yw.cell(2, 1, 'Blank = no data that year (varies by listing date/data source). Amber row = LOW confidence (n<3 years) \u2014 this is where a high win rate on very few trades becomes visible.').font = SUB
hdr = ['Rank', 'Symbol', 'Direction', 'Avg %', 'Yrs', 'Confidence'] + [str(y) for y in ALL_YEARS]
for j, h in enumerate(hdr, 1):
    c = yw.cell(4, j, h); c.fill = NAVY; c.font = WHITE; c.border = THIN; c.alignment = CEN
row = 5
for i, (sym, r) in enumerate(top50, 1):
    sign = 1 if r['pick'] == 'LONG' else -1
    avg = r['pick_avg']; low = r.get('low_confidence')
    vals = [i, sym, r['pick'], (round(avg, 2) if avg is not None else '\u2014'), r['n'], ('LOW (n<3)' if low else 'OK')]
    for y in ALL_YEARS:
        raw = r['year_rets'].get(y)
        vals.append(round(raw * sign, 2) if raw is not None else '')
    rowfill = AM if low else None
    for j, v in enumerate(vals, 1):
        c = yw.cell(row, j, v); c.border = THIN
        if rowfill: c.fill = rowfill
        if j in (1, 3, 5, 6): c.alignment = CEN
        elif j >= 4:
            c.alignment = RIGHT
            if isinstance(v, (int, float)) and j not in (5,):
                c.font = GOOD if v >= 0 else BAD
    row += 1
for j, w in enumerate([6, 12, 9, 8, 6, 12] + [8] * len(ALL_YEARS), 1):
    yw.column_dimensions[get_column_letter(j)].width = w
yw.freeze_panes = 'F5'

# ============ Sheet 3: Overview ============
ov = wb.create_sheet('Overview', 0)
ov.merge_cells('A1:B1')
ov.cell(1, 1, 'TOP 50 F&O STOCKS \u2014 DUSSEHRA PERFORMANCE REPORT').font = TITLE
avg_cagr = round(st.mean(p['cagr'] for p in perf_cache.values() if p['cagr'] is not None), 2)
avg_dd = round(st.mean(p['maxdd'] for p in perf_cache.values()), 2)
n_long_50 = sum(1 for _, r in top50 if r['pick'] == 'LONG'); n_short_50 = 50 - n_long_50
notes = [
    ('Universe & ranking', 'Top 50 of the 195 F&O stocks with usable Dussehra backtest data, ranked by average return (best direction per stock) \u2014 identical metric/method as the companion all-211-stock report.'),
    ('CONFIDENCE \u2014 READ FIRST', f'{n_low_conf_50} of these 50 stocks ({n_low_conf_50*2}%) are LOW confidence \u2014 fewer than 3 years of real backtested data (n<3), highlighted AMBER in every sheet. Ranking purely by average return pulls in thin-sample outliers (e.g. a 100% win rate on just 1 trade) alongside genuinely robust 7-year track records \u2014 always check the Years/Confidence column before acting on any row here, especially near the top.'),
    ('CAGR', 'Compounded annual growth rate of the stock\u2019s own sequence of yearly Dussehra-trade returns (not a calendar-year price CAGR) \u2014 i.e. the return if the Dussehra trade alone were compounded year over year.'),
    ('Max Drawdown', 'Peak-to-trough decline on the cumulative equity curve built purely from the yearly Dussehra trades (one trade per year) \u2014 a trade-level DD, not an intra-trade/intraday DD.'),
    ('Sharpe-like ratio', 'Mean yearly return / population-stdev of yearly returns. Not annualized to daily vol \u2014 a simple risk-adjusted-return measure for comparing stocks on this one event.'),
    ('Detail', f'Direction split: {n_long_50} LONG / {n_short_50} SHORT. Full year-by-year return for every one of the 50 is in the Year_by_Year_Detail sheet \u2014 this is where a high win rate built on very few trades becomes visible at a glance (amber rows).'),
    ('Upcoming positions', 'See Position_Schedule_2026 for the 20-Oct-2026 entry/exit dates, direction, lot size and Confidence flag for all 50, sorted by entry date.'),
    ('Portfolio snapshot', f'Average CAGR across these 50: {avg_cagr:+.2f}% | Average Max DD: {avg_dd:.2f}%.'),
    ('CAVEAT', 'Best-direction-per-stock selection fits history closely but has repeatedly underperformed a single uniform index-backed direction out-of-sample in this project\u2019s own testing \u2014 treat as historical fit, not guaranteed forward edge. This applies with extra force to the LOW-confidence (amber) rows.'),
]
r0 = 3
for lbl, txt in notes:
    a = ov.cell(r0, 1, lbl); a.font = BOLD; a.fill = AM; a.border = THIN; a.alignment = WRAP
    b = ov.cell(r0, 2, txt); b.alignment = WRAP; b.border = THIN
    ov.row_dimensions[r0].height = 60; r0 += 1
ov.column_dimensions['A'].width = 22; ov.column_dimensions['B'].width = 105

wb.save(OUT_XLSX)
print('Saved Excel:', OUT_XLSX)

# ============ PDF ============
def esc(s): return str(s).replace('&', '&amp;').replace('<', '&lt;')
def pdf_row(rank, sym, r, p):
    low = r.get('low_confidence')
    rowstyle = ' style="background:#fffbeb;"' if low else ''
    conf = '<span style="color:#b45309;font-weight:700;">LOW</span>' if low else 'OK'
    cagr = f"{p['cagr']:+.1f}%" if p['cagr'] is not None else '\u2014'
    sharpe = f"{p['sharpe']:.2f}" if p['sharpe'] is not None else '\u2014'
    return (f'<tr{rowstyle}><td>{rank}</td><td><b>{esc(sym)}</b></td><td>{esc(meta.get(sym,{}).get("name",sym))[:22]}</td>'
            f'<td class="c">{r["pick"]}</td><td class="r">{r["pick_wr"]:.1f}%</td>'
            f'<td class="r {"pos" if r["pick_avg"]>=0 else "neg"}">{r["pick_avg"]:+.2f}%</td>'
            f'<td class="r">{cagr}</td><td class="r neg">{p["maxdd"]:.1f}%</td><td class="r">{sharpe}</td>'
            f'<td class="c">{r["n"]}</td><td class="c">{conf}</td></tr>')
rows_all = ''.join(pdf_row(i + 1, s, r, perf_cache[s]) for i, (s, r) in enumerate(top50))

html = f"""<!doctype html><html><head><meta charset="utf-8">
<title>Top 50 F&amp;O Stocks \u2014 Dussehra Performance</title>
<style>
 *{{box-sizing:border-box}} body{{font-family:'Segoe UI',Arial,sans-serif;margin:0;color:#0f172a;background:#fff}}
 .wrap{{max-width:1050px;margin:0 auto;padding:26px 34px}}
 h1{{font-size:21px;margin:0 0 2px}} .sub{{color:#64748b;font-size:12px;margin-bottom:14px}}
 h2{{font-size:15px;margin:20px 0 8px;border-left:4px solid #2563eb;padding-left:9px}}
 table{{width:100%;border-collapse:collapse;font-size:10.8px}}
 th,td{{padding:5px 6px;text-align:right;border-bottom:1px solid #eef2f7}} th{{background:#f1f5f9;font-size:9px;text-transform:uppercase;color:#475569}}
 td.c,th.c{{text-align:center}} .pos{{color:#059669;font-weight:700}} .neg{{color:#dc2626;font-weight:700}}
 .note{{background:#fff7ed;border:1px solid #fed7aa;border-radius:8px;padding:11px 14px;font-size:11px;color:#7c2d12;margin-top:8px;line-height:1.5}}
 .cards{{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin:10px 0}}
 .mc{{background:#fff;border:1px solid #e2e8f0;border-radius:8px;padding:10px 12px}}
 .mcl{{font-size:9.5px;text-transform:uppercase;color:#64748b;font-weight:700}} .mcv{{font-size:19px;font-weight:800}}
 .foot{{color:#94a3b8;font-size:9.5px;margin-top:18px}} @page{{size:A4;margin:12mm}}
</style></head><body><div class="wrap">
 <h1>Top 50 F&amp;O Stocks &mdash; Dussehra Performance Measures</h1>
 <div class="sub">Ranked by average return (best direction per stock, split-robust backtest) &bull; full year-by-year detail for every stock in the companion all-211 Excel.</div>
 <div class="cards">
  <div class="mc"><div class="mcl">Stocks ranked</div><div class="mcv">50</div></div>
  <div class="mc"><div class="mcl">Low-confidence (n&lt;3yr)</div><div class="mcv" style="color:#b45309">{n_low_conf_50}</div></div>
  <div class="mc"><div class="mcl">Avg CAGR</div><div class="mcv pos">{avg_cagr:+.1f}%</div></div>
  <div class="mc"><div class="mcl">Avg Max DD</div><div class="mcv neg">{avg_dd:.1f}%</div></div>
 </div>
 <h2>Top 50 \u2014 Performance Measures</h2>
 <table><thead><tr><th>Rank</th><th class="c">Symbol</th><th class="c">Company</th><th class="c">Dir</th><th>Win%</th><th>Avg Ret</th><th>CAGR</th><th>Max DD</th><th>Sharpe</th><th class="c">Yrs</th><th class="c">Conf.</th></tr></thead>
 <tbody>{rows_all}</tbody></table>
 <div class="note"><b>Definitions:</b> CAGR and Max DD are computed on each stock\u2019s own yearly Dussehra-trade equity curve (one trade/year, not a calendar-year price return). Sharpe-like = mean / stdev of yearly returns (not annualized). Rows marked <b style="color:#b45309;">LOW</b> confidence have under 3 years of data. <b>Caveat:</b> best-direction-per-stock selection fits history closely but has repeatedly underperformed a single uniform index-backed direction out-of-sample in this project's prior testing.</div>
 <div class="foot">Entry/exit schedule for all 50 (20-Oct-2026 anchor) is in the companion Excel, Position_Schedule_2026 sheet.</div>
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
