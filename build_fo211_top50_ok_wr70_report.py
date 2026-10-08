"""
TOP 50 F&O STOCKS — OK CONFIDENCE + WIN RATE >= 70% — DUSSEHRA.
Filters to Confidence = OK (n>=3 years, no thin-sample picks) AND Win Rate >=
70%, then takes the Top 50 of that filtered set ranked by average return.
Full performance measures (CAGR, Max DD, Sharpe-like) + year-by-year detail +
upcoming 20-Oct-2026 position schedule, all with confidence clearly shown.

ADDITIVE: reads the existing intermediate pickle + live dashboard data; writes
new report files only.
"""
import json, pickle, os, subprocess, statistics as st
from datetime import date, timedelta, datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

BASE = r'D:\behaviour analysis'
OUT_XLSX = os.environ.get('FO211_T50OK_OUT_XLSX', BASE + r'\FO211_Dussehra_Top50_OKConfidence_WR70plus_Report.xlsx')
OUT_HTML = BASE + r'\FO211_Dussehra_Top50_OKConfidence_WR70plus_Report.html'
OUT_PDF = BASE + r'\FO211_Dussehra_Top50_OKConfidence_WR70plus_Report.pdf'

WR_MIN = 70.0

data = pickle.load(open(BASE + r'\scraped_parquet\_tmp_fo211_dussehra.pkl', 'rb'))
results, meta = data['results'], data['meta']
qualified = {s: r for s, r in results.items() if r['pick']}
filtered = {s: r for s, r in qualified.items() if r['pick_wr'] >= WR_MIN and not r.get('low_confidence')}
ranked_all = sorted(filtered.items(), key=lambda kv: -(kv[1]['pick_avg'] if kv[1]['pick_avg'] is not None else -999))
top50 = ranked_all[:50]
ALL_YEARS = sorted(set(y for _, r in top50 for y in r['year_rets']))

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
                worst_year=years_sorted[worst_i], worst_ret=signed[worst_i], n=n, signed=dict(zip(years_sorted, signed)))

NAVY = PatternFill('solid', fgColor='1A202C'); WHITE = Font(bold=True, color='FFFFFF', size=9)
TITLE = Font(bold=True, color='1A202C', size=13); SUB = Font(color='4A5568', size=10, italic=True)
GF = PatternFill('solid', fgColor='E6F4EA'); RF = PatternFill('solid', fgColor='FCE8E6')
LBL = PatternFill('solid', fgColor='F1F5F9')
GOOD = Font(color='137333', bold=True); BAD = Font(color='C5221F', bold=True); BOLD = Font(bold=True)
THIN = Border(*[Side(style='thin', color='D9E2EC')] * 4)
CEN = Alignment(horizontal='center'); RIGHT = Alignment(horizontal='right'); WRAP = Alignment(wrap_text=True, vertical='top')

wb = Workbook()

# ============ Sheet 1: Performance Measures ============
pm = wb.active; pm.title = 'Top50_Performance_Measures'
pm.merge_cells('A1:N1')
pm.cell(1, 1, f'TOP 50 \u2014 CONFIDENCE = OK AND WIN RATE \u2265 {WR_MIN:.0f}% ({len(filtered)} qualify, top 50 by avg return shown)').font = TITLE
pm.merge_cells('A2:N2')
pm.cell(2, 1, 'Every row here has >=3 years of real data AND a win rate of at least 70% \u2014 no thin-sample or sub-70% names. Ranked by average return.').font = SUB
hdr = ['Rank', 'Symbol', 'Company', 'Sector', 'Direction', 'Win Rate %', 'Avg Return %', 'CAGR %', 'Max DD %', 'Sharpe-like', 'Years', 'Data Basis', 'Best Yr', 'Worst Yr']
for j, hh in enumerate(hdr, 1):
    c = pm.cell(4, j, hh); c.fill = NAVY; c.font = WHITE; c.border = THIN; c.alignment = CEN
row = 5
perf_cache = {}
for i, (sym, r) in enumerate(top50, 1):
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
        c = pm.cell(row, j, v); c.border = THIN; c.fill = fill
        if j in (1, 5, 11): c.alignment = CEN
        if j in (6, 7, 8, 9, 10): c.alignment = RIGHT
        if j == 7 and isinstance(v, (int, float)): c.font = GOOD if v >= 0 else BAD
        if j == 9: c.font = BAD if p['maxdd'] < 0 else GOOD
        if j == 6: c.font = BOLD
    row += 1
for j, w in enumerate([6, 12, 26, 20, 10, 11, 12, 9, 10, 11, 7, 30, 15, 15], 1):
    pm.column_dimensions[get_column_letter(j)].width = w
pm.freeze_panes = 'A5'

# ============ Sheet 2: Position Schedule 2026 ============
ps = wb.create_sheet('Position_Schedule_2026')
ps.merge_cells('A1:J1')
ps.cell(1, 1, f'DUSSEHRA 20-Oct-2026 \u2014 Position Schedule, Top {len(top50)} (OK Confidence + WR\u2265{WR_MIN:.0f}%)').font = TITLE
ps.merge_cells('A2:J2')
ps.cell(2, 1, 'Entry/exit on the NSE-2026 trading calendar; Nifty-50 names use the exact live-dashboard dates. Sorted by entry date.').font = SUB
hdr = ['Symbol', 'Company', 'Direction', 'Window', 'Entry Date', 'Exit Date', 'Lot Size', 'Win Rate %', 'Avg Return %', 'Years']
for j, hh in enumerate(hdr, 1):
    c = ps.cell(4, j, hh); c.fill = NAVY; c.font = WHITE; c.border = THIN; c.alignment = CEN
sched = []
for sym, r in top50:
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

# ============ Sheet 3: Year-by-Year Detail ============
yw = wb.create_sheet('Year_by_Year_Detail')
yw.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5 + len(ALL_YEARS))
yw.cell(1, 1, f'TOP {len(top50)} \u2014 STOCK x YEAR, Dussehra return %% using each stock\u2019s chosen direction').font = TITLE
hdr = ['Rank', 'Symbol', 'Direction', 'Avg %', 'Yrs'] + [str(y) for y in ALL_YEARS]
for j, h in enumerate(hdr, 1):
    c = yw.cell(4, j, h); c.fill = NAVY; c.font = WHITE; c.border = THIN; c.alignment = CEN
row = 5
for i, (sym, r) in enumerate(top50, 1):
    sign = 1 if r['pick'] == 'LONG' else -1
    avg = r['pick_avg']
    vals = [i, sym, r['pick'], (round(avg, 2) if avg is not None else '\u2014'), r['n']]
    for y in ALL_YEARS:
        raw = r['year_rets'].get(y)
        vals.append(round(raw * sign, 2) if raw is not None else '')
    for j, v in enumerate(vals, 1):
        c = yw.cell(row, j, v); c.border = THIN
        if j in (1, 3, 5): c.alignment = CEN
        elif j >= 4:
            c.alignment = RIGHT
            if isinstance(v, (int, float)) and j != 5:
                c.font = GOOD if v >= 0 else BAD
    row += 1
for j, w in enumerate([6, 12, 9, 8, 6] + [8] * len(ALL_YEARS), 1):
    yw.column_dimensions[get_column_letter(j)].width = w
yw.freeze_panes = 'E5'

# ============ Sheet 4: Overview ============
ov = wb.create_sheet('Overview', 0)
ov.merge_cells('A1:B1')
ov.cell(1, 1, 'TOP 50 \u2014 CONFIDENCE OK + WIN RATE 70%+ \u2014 OVERVIEW').font = TITLE
avg_cagr = round(st.mean(p['cagr'] for p in perf_cache.values() if p['cagr'] is not None), 2)
avg_dd = round(st.mean(p['maxdd'] for p in perf_cache.values()), 2)
avg_wr = round(st.mean(r['pick_wr'] for _, r in top50), 1)
n_long = sum(1 for _, r in top50 if r['pick'] == 'LONG'); n_short = len(top50) - n_long
notes = [
    ('Filter criteria', f'Confidence = OK (\u2265 3 years of real data, no thin-sample picks) AND Win Rate \u2265 {WR_MIN:.0f}%. {len(filtered)} of 211 F&O stocks pass both; the top {len(top50)} by average return are shown here.'),
    ('Why this view', 'Every prior ranked/percentage report mixed in thin-sample (n<3) outliers alongside robust multi-year records. This report removes that noise entirely \u2014 everything here is both reasonably proven (\u22653 years) and a genuinely strong performer (\u226570% win rate).'),
    ('Result', f'{n_long} LONG / {n_short} SHORT. Average win rate {avg_wr}%, average CAGR {avg_cagr:+.2f}%, average Max DD {avg_dd:.2f}%.'),
    ('Data depth caveat', 'Most of these use 2019-2025 F&O futures history (4-7 years), not the full 26-year window \u2014 "OK confidence" means \u22653 years, not necessarily a long track record. Check the Years column per stock.'),
    ('CAGR / Max DD / Sharpe', 'Computed on each stock\u2019s own yearly Dussehra-trade equity curve (one trade/year, split-robust) \u2014 not a calendar-year price return.'),
    ('Upcoming positions', 'See Position_Schedule_2026 for the 20-Oct-2026 entry/exit dates, direction and lot size for all 50, sorted by entry date.'),
    ('CAVEAT', 'Best-direction-per-stock selection fits history closely but has repeatedly underperformed a single uniform index-backed direction out-of-sample in this project\u2019s own testing \u2014 treat as historical fit, not a guaranteed forward edge.'),
]
r0 = 3
for lbl, txt in notes:
    a = ov.cell(r0, 1, lbl); a.font = BOLD; a.fill = LBL; a.border = THIN; a.alignment = WRAP
    b = ov.cell(r0, 2, txt); b.alignment = WRAP; b.border = THIN
    ov.row_dimensions[r0].height = 62; r0 += 1
ov.column_dimensions['A'].width = 24; ov.column_dimensions['B'].width = 105

# ============ Sheet 5: Methodology (logic behind every measure) ============
mt = wb.create_sheet('Methodology')
mt.merge_cells('A1:C1')
mt.cell(1, 1, 'METHODOLOGY — HOW EACH PERFORMANCE MEASURE IS CALCULATED').font = TITLE
mt_hdr = ['Step / Measure', 'Logic', 'Formula']
for j, h in enumerate(mt_hdr, 1):
    c = mt.cell(3, j, h); c.fill = NAVY; c.font = WHITE; c.border = THIN; c.alignment = CEN
mt_rows = [
    ('1. The event & window',
     'Dussehra anchor date each year (20-Oct-2026 for the upcoming trade). Each stock has its own entry/exit window, expressed as T-n (n trading days BEFORE the anchor) to T+m (m trading days AFTER). Entry at 09:20 AM, exit at 03:15 PM on the NSE-2026 trading-day calendar (weekends + 15 listed holidays excluded).',
     'Entry date = anchor shifted back n trading days. Exit date = anchor shifted forward m trading days.'),
    ('2. One trade per stock per year',
     'For every historical Dussehra date (up to 26 years, 2000-2025), the stock’s close-to-close price path from entry day to exit day is used — this produces exactly one return number per stock per year it has price data for.',
     'year_return = cumulative %% change in price from entry-day close to exit-day close.'),
    ('3. Split-robust compounding',
     'Each day inside the window, the daily %% price change is computed and compounded day-by-day (not a simple start-to-end %%). Any single day where the price jumps by more than 20%% is treated as a stock split/bonus/demerger artifact, not a real market move, and is SKIPPED from the compounding (confirmed necessary — e.g. DRREDDY’s 5:1 split, RELIANCE’s split, 135+ such days project-wide would otherwise fake huge returns).',
     'cumulative_gain = Π(1 + daily_%%_change) for each day in window, skipping any |daily_%%_change| > 20%%. year_return = (cumulative_gain − 1) × 100.'),
    ('4. Direction: LONG vs SHORT',
     'Both directions are backtested on the SAME raw yearly price moves. LONG wins when the raw move is positive; SHORT wins when it is negative (mirror image). The direction with the HIGHER win rate across all available years is chosen per stock — this is a best-of-both-directions, per-stock selection, not a fixed index-wide rule.',
     'LONG win rate = %% of years with raw_move > 0. SHORT win rate = 100 − LONG win rate. Chosen direction = whichever win rate is higher.'),
    ('5. Win Rate %',
     'Of all the years with usable data for that stock, the %% where the CHOSEN direction was profitable.',
     'Win Rate = (winning years ÷ total years with data) × 100.'),
    ('6. Avg Return %',
     'The simple average of the stock’s own yearly returns, signed for its chosen direction (a SHORT pick flips the sign of the raw price move).',
     'Avg Return = mean(signed yearly returns).'),
    ('7. CAGR % (this report’s addition)',
     'Not a price CAGR — it is the compounded annual growth if ONLY the Dussehra trade were taken every year and the proceeds reinvested into the next year’s Dussehra trade. Rewards consistency; a stock with 3 wins of +5%% compounds to a higher CAGR than one with 2 wins of +8%% and 1 loss of -6%%.',
     'cumulative = Π(1 + yearly_return/100) across all years. CAGR = (cumulative^(1/n) − 1) × 100, where n = years of data.'),
    ('8. Max Drawdown % (this report’s addition)',
     'Peak-to-trough decline on the cumulative equity curve built purely from the yearly Dussehra trades (starting at 100, compounding one trade per year in chronological order) — a TRADE-LEVEL drawdown across the strategy’s history, not an intraday or intra-trade drawdown.',
     'equity[0] = 100; equity[t] = equity[t-1] × (1 + yearly_return[t]/100). Max DD = min over all t of (equity[t] − running_peak) ÷ running_peak × 100.'),
    ('9. Sharpe-like ratio (this report’s addition)',
     'A simple risk-adjusted-return measure comparing the stock’s average yearly return to how much that return varies year to year — NOT annualized to daily volatility like a textbook Sharpe ratio, so only compare it stock-to-stock within this report, not to a standard market Sharpe number.',
     'Sharpe-like = mean(yearly returns) ÷ population_stdev(yearly returns). Left blank where fewer than 2 years of data exist (stdev undefined).'),
    ('10. Confidence flag',
     'Stocks with fewer than 3 years of usable data are flagged LOW — a win rate built on 1-2 trades is not a statistically meaningful sample, even if it reads 100%%. This report ONLY includes OK-confidence stocks (≥ 3 years) by design — every row has already passed this filter.',
     'Confidence = OK if years_of_data ≥ 3, else LOW. (This report excludes all LOW rows.)'),
    ('11. Data source per stock',
     'Preference order: 26-year cash-equity daily close (2000-2025) where the stock has it (mostly Nifty-50 names); otherwise near-month F&O futures continuous series (2019-2025, up to 7 years) for F&O-only names. Shown in the Data Basis column — a 26yr figure and a 5yr figure are not equally reliable even at the same win rate.',
     'source = ‘cash-26yr’ if available, else ‘futures-7yr’.'),
    ('12. This report’s filter & rank',
     f'From the full 211-stock F&amp;O universe: keep only Confidence = OK AND Win Rate ≥ {WR_MIN:.0f}%% ({len(filtered)} stocks pass), then rank the survivors by Avg Return % and take the top {len(top50)}.',
     f'Shown here = Top {len(top50)} of {len(filtered)} stocks with (confidence=OK AND win_rate≥70), sorted by avg_return descending.'),
]
row = 4
for lbl, logic, formula in mt_rows:
    a = mt.cell(row, 1, lbl); a.font = BOLD; a.fill = LBL; a.border = THIN; a.alignment = WRAP
    b = mt.cell(row, 2, logic); b.alignment = WRAP; b.border = THIN
    c = mt.cell(row, 3, formula); c.alignment = WRAP; c.border = THIN; c.font = Font(italic=True, size=9, color='1E3A8A')
    mt.row_dimensions[row].height = 85
    row += 1
mt.merge_cells(f'A{row+1}:C{row+1}')
warn = mt.cell(row + 1, 1,
    'CAVEAT (applies to every measure above): choosing the single best-performing direction per stock, as this entire methodology does, fits the historical sample closely — but in this project’s own repeated out-of-sample testing, a single uniform index-backed direction has outperformed this per-stock approach going forward. All figures here are the best HISTORICAL fit, not a guaranteed forward edge.')
warn.font = Font(color='7C2D12', size=10, italic=True); warn.alignment = WRAP
mt.row_dimensions[row + 1].height = 55
mt.column_dimensions['A'].width = 26; mt.column_dimensions['B'].width = 70; mt.column_dimensions['C'].width = 45
mt.freeze_panes = 'A4'

wb.save(OUT_XLSX)
print('Saved Excel:', OUT_XLSX)
print(f'{len(filtered)} total qualify (OK confidence, WR>={WR_MIN:.0f}%); top {len(top50)} shown: {n_long} LONG / {n_short} SHORT')

# ============ PDF ============
def esc(s): return str(s).replace('&', '&amp;').replace('<', '&lt;')
def pdf_row(rank, sym, r, p):
    avg = r['pick_avg']; avg_str = f"{avg:+.2f}%" if avg is not None else '\u2014'
    cagr = f"{p['cagr']:+.1f}%" if p['cagr'] is not None else '\u2014'
    return (f'<tr><td>{rank}</td><td><b>{esc(sym)}</b></td><td>{esc(meta.get(sym,{}).get("name",sym))[:26]}</td>'
            f'<td class="c">{r["pick"]}</td><td class="r"><b>{r["pick_wr"]:.1f}%</b></td>'
            f'<td class="r {"pos" if (avg or 0)>=0 else "neg"}">{avg_str}</td>'
            f'<td class="r">{cagr}</td><td class="r neg">{p["maxdd"]:.1f}%</td><td class="c">{r["n"]}</td></tr>')
rows_all = ''.join(pdf_row(i + 1, s, r, perf_cache[s]) for i, (s, r) in enumerate(top50))

html = f"""<!doctype html><html><head><meta charset="utf-8">
<title>Top 50 OK-Confidence WR70+</title>
<style>
 *{{box-sizing:border-box}} body{{font-family:'Segoe UI',Arial,sans-serif;margin:0;color:#0f172a;background:#fff}}
 .wrap{{max-width:1050px;margin:0 auto;padding:30px 36px}}
 h1{{font-size:22px;margin:0 0 3px}} .sub{{color:#64748b;font-size:12.5px;margin-bottom:16px}}
 h2{{font-size:15px;margin:22px 0 8px;border-left:4px solid #0f766e;padding-left:9px}}
 table{{width:100%;border-collapse:collapse;font-size:11px}}
 th,td{{padding:6px 7px;text-align:right;border-bottom:1px solid #eef2f7}} th{{background:#f1f5f9;font-size:9.5px;text-transform:uppercase;color:#475569}}
 td.c,th.c{{text-align:center}} .pos{{color:#059669;font-weight:700}} .neg{{color:#dc2626;font-weight:700}}
 .note{{background:#f0fdfa;border:1px solid #99f6e4;border-radius:8px;padding:12px 15px;font-size:11px;color:#134e4a;margin-top:10px;line-height:1.55}}
 .cards{{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin:10px 0 4px}}
 .mc{{background:#fff;border:1px solid #e2e8f0;border-radius:8px;padding:11px 13px}}
 .mcl{{font-size:9.5px;text-transform:uppercase;color:#64748b;font-weight:700}} .mcv{{font-size:20px;font-weight:800}}
 .foot{{color:#94a3b8;font-size:9.5px;margin-top:18px}} @page{{size:A4;margin:12mm}}
</style></head><body><div class="wrap">
 <h1>Top {len(top50)} &mdash; Confidence OK &amp; Win Rate \u2265 {WR_MIN:.0f}%</h1>
 <div class="sub">F&amp;O-wide screen &bull; \u2265 3 years of real data AND win rate \u2265 70% &bull; {len(filtered)} of 211 F&amp;O stocks qualify, top {len(top50)} by avg return shown</div>
 <div class="cards">
  <div class="mc"><div class="mcl">Stocks shown</div><div class="mcv">{len(top50)}</div></div>
  <div class="mc"><div class="mcl">Avg Win Rate</div><div class="mcv pos">{avg_wr}%</div></div>
  <div class="mc"><div class="mcl">Avg CAGR</div><div class="mcv pos">{avg_cagr:+.1f}%</div></div>
  <div class="mc"><div class="mcl">LONG / SHORT</div><div class="mcv">{n_long} / {n_short}</div></div>
 </div>
 <h2>Top {len(top50)} \u2014 Ranked by Avg Return</h2>
 <table><thead><tr><th>Rank</th><th class="c">Symbol</th><th class="c">Company</th><th class="c">Dir</th><th>Win%</th><th>Avg Ret</th><th>CAGR</th><th>Max DD</th><th class="c">Yrs</th></tr></thead>
 <tbody>{rows_all}</tbody></table>
 <div class="note"><b>This is the clean list:</b> no thin-sample (n&lt;3) picks, no sub-70% win rates \u2014 every stock here is both reasonably proven and a strong historical performer. Most still rely on 2019-2025 F&amp;O futures history (4-7 years), not a full multi-decade record. <b>Caveat:</b> best-direction-per-stock selection has repeatedly underperformed a uniform index-backed direction out-of-sample in this project's prior testing.</div>
 <h2>Methodology \u2014 How Each Measure Is Calculated</h2>
 <table><thead><tr><th style="text-align:left">Measure</th><th style="text-align:left">Logic</th><th style="text-align:left">Formula</th></tr></thead>
 <tbody>
 <tr><td style="text-align:left"><b>Window &amp; event</b></td><td style="text-align:left">Each stock has its own T-n/T+m trading-day window around the 20-Oct-2026 Dussehra anchor, on the NSE-2026 calendar.</td><td style="text-align:left;font-style:italic;color:#1e3a8a">Entry = anchor \u2212 n trading days; Exit = anchor + m trading days.</td></tr>
 <tr><td style="text-align:left"><b>Split-robust return</b></td><td style="text-align:left">Daily %% changes inside the window are compounded; any single-day move &gt;20%% (stock split/bonus artifact) is skipped.</td><td style="text-align:left;font-style:italic;color:#1e3a8a">year_return = (\u03a0(1+daily_%%), skip |daily_%%|&gt;20%%) \u2212 1, \u00d7100.</td></tr>
 <tr><td style="text-align:left"><b>Direction (LONG/SHORT)</b></td><td style="text-align:left">Both directions backtested on the same yearly moves; the one with the higher win rate across all years is chosen per stock.</td><td style="text-align:left;font-style:italic;color:#1e3a8a">Chosen = direction with max(LONG win%%, SHORT win%%).</td></tr>
 <tr><td style="text-align:left"><b>Win Rate %%</b></td><td style="text-align:left">Share of years the chosen direction was profitable.</td><td style="text-align:left;font-style:italic;color:#1e3a8a">winning years \u00f7 total years with data \u00d7 100.</td></tr>
 <tr><td style="text-align:left"><b>Avg Return %%</b></td><td style="text-align:left">Simple average of the signed yearly returns.</td><td style="text-align:left;font-style:italic;color:#1e3a8a">mean(signed yearly returns).</td></tr>
 <tr><td style="text-align:left"><b>CAGR %%</b></td><td style="text-align:left">Compounded annual growth if only this one trade/year were reinvested year over year \u2014 rewards consistency over lucky one-offs.</td><td style="text-align:left;font-style:italic;color:#1e3a8a">(\u03a0(1+r/100))^(1/n) \u2212 1, \u00d7100.</td></tr>
 <tr><td style="text-align:left"><b>Max Drawdown %%</b></td><td style="text-align:left">Peak-to-trough decline on the cumulative equity curve of the yearly trades (trade-level, not intraday).</td><td style="text-align:left;font-style:italic;color:#1e3a8a">min((equity_t \u2212 running_peak) \u00f7 running_peak) \u00d7100.</td></tr>
 <tr><td style="text-align:left"><b>Sharpe-like</b></td><td style="text-align:left">Risk-adjusted return vs. year-to-year variability; NOT annualized \u2014 only compare within this report.</td><td style="text-align:left;font-style:italic;color:#1e3a8a">mean(yearly returns) \u00f7 population_stdev(yearly returns).</td></tr>
 <tr><td style="text-align:left"><b>Confidence &amp; this report\u2019s filter</b></td><td style="text-align:left">OK = \u22653 years of data (LOW excluded entirely here). This report keeps only Confidence=OK AND Win Rate\u2265{WR_MIN:.0f}%%, ranks by avg return, shows the top {len(top50)}.</td><td style="text-align:left;font-style:italic;color:#1e3a8a">{len(filtered)} stocks pass the filter \u2192 top {len(top50)} by avg_return shown.</td></tr>
 </tbody></table>
 <div class="foot">Upcoming 20-Oct-2026 entry/exit schedule for every stock here is in the companion Excel, Position_Schedule_2026 sheet. Full formula detail is in the Excel's Methodology sheet.</div>
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
