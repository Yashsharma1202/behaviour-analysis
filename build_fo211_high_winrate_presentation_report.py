"""
HIGH-CONVICTION DUSSEHRA PLAYBOOK — presentation report.
Filters the full 211-stock F&O Dussehra backtest universe down to stocks with
Win Rate >= 75% AND Confidence = OK (n >= 3 years, not low-confidence) — a
clean, presentable shortlist. Includes full performance measures (CAGR, Max
DD, Sharpe-like) and the upcoming 20-Oct-2026 position schedule for exactly
these stocks.

ADDITIVE: reads the existing intermediate pickle + live dashboard data; writes
new report files only.
"""
import json, pickle, os, subprocess, statistics as st
from datetime import date, timedelta, datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

BASE = r'D:\behaviour analysis'
OUT_XLSX = os.environ.get('FO211_HWR_OUT_XLSX', BASE + r'\FO211_Dussehra_High_WinRate_75plus_Report.xlsx')
OUT_HTML = BASE + r'\FO211_Dussehra_High_WinRate_75plus_Report.html'
OUT_PDF = BASE + r'\FO211_Dussehra_High_WinRate_75plus_Report.pdf'

WR_MIN = 75.0

data = pickle.load(open(BASE + r'\scraped_parquet\_tmp_fo211_dussehra.pkl', 'rb'))
results, meta = data['results'], data['meta']
qualified = {s: r for s, r in results.items() if r['pick']}
sel = {s: r for s, r in qualified.items() if r['pick_wr'] >= WR_MIN and not r.get('low_confidence')}
ranked = sorted(sel.items(), key=lambda kv: (-kv[1]['pick_wr'], -(kv[1]['pick_avg'] if kv[1]['pick_avg'] is not None else -999)))

# ---- upcoming Dussehra 2026 position dates (same method as the other reports) ----
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
TITLE = Font(bold=True, color='1A202C', size=14); SUB = Font(color='4A5568', size=10, italic=True)
GF = PatternFill('solid', fgColor='E6F4EA'); RF = PatternFill('solid', fgColor='FCE8E6')
LBL = PatternFill('solid', fgColor='F1F5F9')
GOOD = Font(color='137333', bold=True); BAD = Font(color='C5221F', bold=True); BOLD = Font(bold=True)
THIN = Border(*[Side(style='thin', color='D9E2EC')] * 4)
CEN = Alignment(horizontal='center'); RIGHT = Alignment(horizontal='right'); WRAP = Alignment(wrap_text=True, vertical='top')

wb = Workbook()

# ============ Sheet 1: High Win-Rate Shortlist ============
sh = wb.active; sh.title = 'High_WinRate_Shortlist'
sh.merge_cells('A1:N1')
sh.cell(1, 1, f'DUSSEHRA HIGH-CONVICTION SHORTLIST \u2014 Win Rate \u2265 {WR_MIN:.0f}% AND Confidence = OK ({len(ranked)} stocks)').font = TITLE
sh.merge_cells('A2:N2')
sh.cell(2, 1, 'Filtered from all 211 F&O stocks: win rate on the chosen (best) direction \u2265 75%, and at least 3 years of real backtested data (no 1-2 year "lucky streak" picks). Sorted by win rate, then avg return.').font = SUB
hdr = ['Rank', 'Symbol', 'Company', 'Sector', 'Direction', 'Win Rate %', 'Avg Return %', 'CAGR %', 'Max DD %', 'Sharpe-like', 'Years', 'Data Basis', 'Best Yr', 'Worst Yr']
for j, hh in enumerate(hdr, 1):
    c = sh.cell(4, j, hh); c.fill = NAVY; c.font = WHITE; c.border = THIN; c.alignment = CEN
row = 5
perf_cache = {}
for i, (sym, r) in enumerate(ranked, 1):
    m = meta.get(sym, {})
    p = perf_measures(r); perf_cache[sym] = p
    avg = r['pick_avg']
    vals = [i, sym, m.get('name', sym), m.get('sector', ''), r['pick'], r['pick_wr'],
            (round(avg, 2) if avg is not None else '\u2014'),
            (round(p['cagr'], 2) if p['cagr'] is not None else '\u2014'), round(p['maxdd'], 2),
            (round(p['sharpe'], 2) if p['sharpe'] is not None else '\u2014'), r['n'], r['source'],
            f"{p['best_year']} ({p['best_ret']:+.1f}%)", f"{p['worst_year']} ({p['worst_ret']:+.1f}%)"]
    fill = GF if (avg is None or avg >= 0) else RF
    for j, v in enumerate(vals, 1):
        c = sh.cell(row, j, v); c.border = THIN; c.fill = fill
        if j in (1, 5, 11): c.alignment = CEN
        if j in (6, 7, 8, 9, 10): c.alignment = RIGHT
        if j == 7 and isinstance(v, (int, float)): c.font = GOOD if v >= 0 else BAD
        if j == 9: c.font = BAD if p['maxdd'] < 0 else GOOD
    row += 1
for j, w in enumerate([6, 12, 24, 20, 10, 11, 12, 9, 10, 11, 7, 30, 15, 15], 1):
    sh.column_dimensions[get_column_letter(j)].width = w
sh.freeze_panes = 'A5'

# ============ Sheet 2: Position Schedule 2026 ============
ps = wb.create_sheet('Position_Schedule_2026')
ps.merge_cells('A1:J1')
ps.cell(1, 1, f'DUSSEHRA 20-Oct-2026 \u2014 Position Schedule, {len(ranked)} High-Conviction Stocks').font = TITLE
ps.merge_cells('A2:J2')
ps.cell(2, 1, 'Entry/exit on the NSE-2026 trading calendar; Nifty-50 names use the exact live-dashboard dates. Sorted by entry date.').font = SUB
hdr = ['Symbol', 'Company', 'Direction', 'Window', 'Entry Date', 'Exit Date', 'Lot Size', 'Win Rate %', 'Avg Return %', 'Years']
for j, hh in enumerate(hdr, 1):
    c = ps.cell(4, j, hh); c.fill = NAVY; c.font = WHITE; c.border = THIN; c.alignment = CEN
sched = []
for sym, r in ranked:
    pos = position_for(sym, r)
    m = meta.get(sym, {})
    sort_key = datetime.strptime(pos['entry'].split(' (')[0].strip(), '%d-%b-%Y')
    avg = r['pick_avg']
    sched.append((sort_key, sym, m.get('name', sym), pos['direction'], f"T-{r['lead']} to T+{r['hold']}",
                  pos['entry'], pos['exit'], pos['lot'], r['pick_wr'], (round(avg, 2) if avg is not None else '\u2014'), r['n']))
sched.sort(key=lambda t: (t[0], t[1]))
row = 5
for _, sym, name, direc, window, entry, exit_, lot, wr, avg, n in sched:
    vals = [sym, name, direc, window, entry, exit_, lot, wr, avg, n]
    fill = GF if direc == 'LONG' else RF
    for j, v in enumerate(vals, 1):
        c = ps.cell(row, j, v); c.border = THIN; c.fill = fill
        if j in (1, 3, 4, 5, 6, 10): c.alignment = CEN
        if j in (7, 8, 9): c.alignment = RIGHT
    row += 1
for j, w in enumerate([13, 28, 10, 15, 16, 16, 9, 11, 12, 7], 1):
    ps.column_dimensions[get_column_letter(j)].width = w
ps.freeze_panes = 'A5'

# ============ Sheet 3: Overview ============
ov = wb.create_sheet('Overview', 0)
ov.merge_cells('A1:B1')
ov.cell(1, 1, 'DUSSEHRA HIGH-CONVICTION SHORTLIST \u2014 OVERVIEW').font = TITLE
avg_cagr = round(st.mean(p['cagr'] for p in perf_cache.values() if p['cagr'] is not None), 2)
avg_dd = round(st.mean(p['maxdd'] for p in perf_cache.values()), 2)
avg_wr = round(st.mean(r['pick_wr'] for _, r in ranked), 1)
n_long = sum(1 for _, r in ranked if r['pick'] == 'LONG'); n_short = len(ranked) - n_long
n_7yr_only = sum(1 for _, r in ranked if r['n'] < 7 and r['source'] != 'cash-26yr')
notes = [
    ('Filter criteria', f'Win Rate \u2265 {WR_MIN:.0f}% on the chosen (best) direction, AND Confidence = OK (\u2265 3 years of real backtested data). {len(ranked)} of 211 F&O stocks pass both filters.'),
    ('Result', f'{n_long} LONG / {n_short} SHORT. Average win rate {avg_wr}%, average CAGR {avg_cagr:+.2f}%, average Max DD {avg_dd:.2f}%.'),
    ('IMPORTANT \u2014 read before presenting', f'"Confidence = OK" here means \u2265 3 years, not necessarily a long track record \u2014 {n_7yr_only} of these {len(ranked)} stocks have only 4-7 years of F&O futures history (2019-2025), not the full 26-year window. A 100% or 85.7% win rate on 5-7 trades is a strong signal, not a certainty \u2014 state the sample size (Years column) whenever you quote a win rate.'),
    ('Deepest history', 'ADANIPORTS is the only stock here with the full 18-26yr cash-equity backtest; all others use the 2019-2025 F&O futures series.'),
    ('CAGR', 'Compounded annual growth from the stock\u2019s own sequence of yearly Dussehra-trade returns (one trade/year), not a calendar-year price CAGR.'),
    ('Max Drawdown', 'Peak-to-trough decline on the cumulative equity curve of the yearly Dussehra trades \u2014 a trade-level DD, not intraday.'),
    ('Upcoming positions', 'See Position_Schedule_2026 for the 20-Oct-2026 entry/exit dates, direction and lot size for all stocks on this shortlist, sorted by entry date.'),
    ('CAVEAT', 'Best-direction-per-stock selection fits history closely but has repeatedly underperformed a single uniform index-backed direction out-of-sample in this project\u2019s own testing \u2014 treat as historical fit, not a guaranteed forward edge, especially on the shorter (4-7yr) samples.'),
]
r0 = 3
for lbl, txt in notes:
    a = ov.cell(r0, 1, lbl); a.font = BOLD; a.fill = LBL; a.border = THIN; a.alignment = WRAP
    b = ov.cell(r0, 2, txt); b.alignment = WRAP; b.border = THIN
    ov.row_dimensions[r0].height = 62; r0 += 1
ov.column_dimensions['A'].width = 26; ov.column_dimensions['B'].width = 105

wb.save(OUT_XLSX)
print('Saved Excel:', OUT_XLSX)
print(f'{len(ranked)} stocks qualify (WR>=75, confidence OK): {n_long} LONG / {n_short} SHORT')

# ============ PDF (presentation style) ============
def esc(s): return str(s).replace('&', '&amp;').replace('<', '&lt;')
def pdf_row(rank, sym, r, p):
    avg = r['pick_avg']
    avg_str = f"{avg:+.2f}%" if avg is not None else '\u2014'
    cagr = f"{p['cagr']:+.1f}%" if p['cagr'] is not None else '\u2014'
    return (f'<tr><td>{rank}</td><td><b>{esc(sym)}</b></td><td>{esc(meta.get(sym,{}).get("name",sym))[:26]}</td>'
            f'<td class="c">{r["pick"]}</td><td class="r">{r["pick_wr"]:.1f}%</td>'
            f'<td class="r {"pos" if (avg or 0)>=0 else "neg"}">{avg_str}</td>'
            f'<td class="r">{cagr}</td><td class="r neg">{p["maxdd"]:.1f}%</td><td class="c">{r["n"]}</td></tr>')
rows_all = ''.join(pdf_row(i + 1, s, r, perf_cache[s]) for i, (s, r) in enumerate(ranked))

html = f"""<!doctype html><html><head><meta charset="utf-8">
<title>Dussehra High-Conviction Shortlist</title>
<style>
 *{{box-sizing:border-box}} body{{font-family:'Segoe UI',Arial,sans-serif;margin:0;color:#0f172a;background:#fff}}
 .wrap{{max-width:1050px;margin:0 auto;padding:30px 36px}}
 h1{{font-size:23px;margin:0 0 3px}} .sub{{color:#64748b;font-size:12.5px;margin-bottom:16px}}
 h2{{font-size:15px;margin:22px 0 8px;border-left:4px solid #137333;padding-left:9px}}
 table{{width:100%;border-collapse:collapse;font-size:11px}}
 th,td{{padding:6px 7px;text-align:right;border-bottom:1px solid #eef2f7}} th{{background:#f1f5f9;font-size:9.5px;text-transform:uppercase;color:#475569}}
 td.c,th.c{{text-align:center}} .pos{{color:#059669;font-weight:700}} .neg{{color:#dc2626;font-weight:700}}
 .note{{background:#fff7ed;border:1px solid #fed7aa;border-radius:8px;padding:12px 15px;font-size:11px;color:#7c2d12;margin-top:10px;line-height:1.55}}
 .cards{{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin:10px 0 4px}}
 .mc{{background:#fff;border:1px solid #e2e8f0;border-radius:8px;padding:11px 13px}}
 .mcl{{font-size:9.5px;text-transform:uppercase;color:#64748b;font-weight:700}} .mcv{{font-size:20px;font-weight:800}}
 .foot{{color:#94a3b8;font-size:9.5px;margin-top:18px}} @page{{size:A4;margin:12mm}}
</style></head><body><div class="wrap">
 <h1>Dussehra High-Conviction Shortlist</h1>
 <div class="sub">F&amp;O-wide screen &bull; Win Rate \u2265 75% AND \u2265 3 years of real backtested data &bull; {len(ranked)} of 211 F&amp;O stocks qualify</div>
 <div class="cards">
  <div class="mc"><div class="mcl">Stocks</div><div class="mcv">{len(ranked)}</div></div>
  <div class="mc"><div class="mcl">Avg Win Rate</div><div class="mcv pos">{avg_wr}%</div></div>
  <div class="mc"><div class="mcl">Avg CAGR</div><div class="mcv pos">{avg_cagr:+.1f}%</div></div>
  <div class="mc"><div class="mcl">LONG / SHORT</div><div class="mcv">{n_long} / {n_short}</div></div>
 </div>
 <h2>Shortlist \u2014 Ranked by Win Rate</h2>
 <table><thead><tr><th>Rank</th><th class="c">Symbol</th><th class="c">Company</th><th class="c">Dir</th><th>Win%</th><th>Avg Ret</th><th>CAGR</th><th>Max DD</th><th class="c">Yrs</th></tr></thead>
 <tbody>{rows_all}</tbody></table>
 <div class="note"><b>Read before presenting:</b> {n_7yr_only} of these {len(ranked)} stocks have only 4-7 years of F&amp;O futures history (2019-2025) \u2014 a high win rate on 5-7 trades is a strong signal, not a certainty. Only ADANIPORTS here has the full multi-decade cash-equity backtest. <b>Caveat:</b> best-direction-per-stock selection fits history closely but has repeatedly underperformed a single uniform index-backed direction out-of-sample in this project's own testing.</div>
 <div class="foot">Upcoming 20-Oct-2026 entry/exit schedule for every stock here is in the companion Excel, Position_Schedule_2026 sheet.</div>
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
