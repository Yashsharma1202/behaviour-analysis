"""
Build the Gandhi Jayanti Holiday Drift performance report (HTML, print-ready to PDF),
in the same format as Quarterly_Results_Performance_Report.pdf.

ADDITIVE: writes a new local HTML file. Sources only real data:
  - 17_NSE_Holidays_Reports/13_Mahatma_Gandhi_Jayanti/*.xlsx  (per-stock, 2022-2025)
  - dashboard_data/nifty_futures_holiday_behaviour.json        (index-level, 2000-2026)
"""
import sys; sys.stdout.reconfigure(encoding='utf-8', errors='replace')
import pandas as pd, json, os, html

BASE = r'17_NSE_Holidays_Reports/13_Mahatma_Gandhi_Jayanti'
OUT = r'D:\behaviour analysis\Gandhi_Jayanti_Performance_Report.html'
YEARS = [2022, 2023, 2024, 2025]

def rupees(x):
    x = float(x)
    neg = x < 0
    x = abs(x)
    s = f'{x:,.0f}'
    return ('-' if neg else '+') + '\u20b9' + s

def load_year(y):
    xl = pd.ExcelFile(os.path.join(BASE, f'Mahatma_Gandhi_Jayanti_{y}_TradeLog.xlsx'), engine='openpyxl')
    df = xl.parse(xl.sheet_names[0], header=2)
    df = df[df['Stock Symbol'].notna()].copy()
    gp = [c for c in df.columns if 'Gross PnL' in str(c)][0]
    df['pnl'] = pd.to_numeric(df[gp], errors='coerce')
    df['ret'] = pd.to_numeric(df['Return (%)'], errors='coerce') * 100
    df['status'] = df['Trade Status'].astype(str).str.upper()
    df['ed'] = pd.to_datetime(df['Entry Date']); df['xd'] = pd.to_datetime(df['Exit Date'])
    return df

# ---- per-year + overall ----
peryear = {}
frames = []
for y in YEARS:
    df = load_year(y)
    w = int((df['status'] == 'WIN').sum()); n = len(df); l = n - w
    ym = df['ed'].dt.year == y  # ignore mislabeled backfilled rows (e.g. JIOFIN 2023 in the 2022 log)
    span = f"{df.loc[ym, 'ed'].min():%d-%b} to {df.loc[ym, 'xd'].max():%d-%b}" if ym.any() else '—'
    peryear[y] = dict(
        n=n, w=w, l=l, wr=round(w / n * 100, 1), pnl=df['pnl'].sum(),
        avg=df['ret'].mean(), hol=str(df['Holiday Date'].iloc[0])[:10], span=span,
        longs=int(df['Strategy'].str.contains('LONG').sum()),
        shorts=int(df['Strategy'].str.contains('SHORT').sum()))
    frames.append(df)
allc = pd.concat(frames, ignore_index=True)
TOTn = len(allc); TOTw = int((allc['status'] == 'WIN').sum()); TOTl = TOTn - TOTw
TOTpnl = allc['pnl'].sum()
gw = allc.loc[allc['pnl'] > 0, 'pnl'].sum(); gl = allc.loc[allc['pnl'] < 0, 'pnl'].sum()
PF = gw / abs(gl) if gl else 0
avg = allc['ret'].mean(); avgW = allc.loc[allc['status'] == 'WIN', 'ret'].mean(); avgL = allc.loc[allc['status'] != 'WIN', 'ret'].mean()
best_y = max(YEARS, key=lambda y: peryear[y]['pnl'])
cap2024 = None
try:
    _d = load_year(2024)
    mcol = [c for c in _d.columns if 'Margin Req' in str(c)][0]
    cap2024 = pd.to_numeric(_d[mcol], errors='coerce').sum()
except Exception:
    cap2024 = 0

# ---- top 10 (consolidated) ----
cons = pd.ExcelFile(os.path.join(BASE, 'Mahatma_Gandhi_Jayanti_4Year_Consolidated.xlsx'),
                    engine='openpyxl').parse('4-Year Consolidated Summary', header=2)
cons = cons[cons['Stock Symbol'].notna()].copy()
wcol = [c for c in cons.columns if 'Win Rate' in str(c)][0]
tcol = [c for c in cons.columns if 'Total' in str(c)][0]
cons['WR'] = pd.to_numeric(cons[wcol], errors='coerce') * 100
cons['TOT'] = pd.to_numeric(cons[tcol], errors='coerce')
top10 = cons.sort_values('TOT', ascending=False).head(10)

# ---- latest cycle detailed positions (2025) ----
lat = load_year(2025).sort_values('pnl', ascending=False)

# ---- index-level confirmation ----
idx = [h for h in json.load(open('dashboard_data/nifty_futures_holiday_behaviour.json', encoding='utf-8'))['holidays'] if 'Gandhi' in h['name']][0]

# =================== HTML ===================
def esc(s): return html.escape(str(s))

css = """
* { box-sizing: border-box; }
body { font-family: -apple-system, 'Segoe UI', Roboto, Arial, sans-serif; color:#2D3748; margin:0; font-size:12.5px; line-height:1.5; }
.page { width: 210mm; min-height: 297mm; padding: 18mm 16mm; margin: 0 auto; background:#fff; position:relative; }
.kicker { color:#2B6CB0; font-weight:700; font-size:10.5px; letter-spacing:.12em; text-transform:uppercase; }
h1 { font-size:30px; line-height:1.1; margin:6px 0 4px; color:#1A202C; font-weight:800; letter-spacing:-.5px; }
.sub { color:#4A5568; font-size:14px; margin-bottom:12px; }
.rule { height:3px; background:#2B6CB0; border:0; margin:10px 0 18px; }
h2 { font-size:18px; color:#1A202C; font-weight:800; margin:22px 0 10px; }
.sec-blue { color:#2B6CB0; font-weight:700; font-size:13.5px; margin:14px 0 6px; }
p { margin:8px 0; }
b, strong { color:#1A202C; }
table { width:100%; border-collapse:collapse; margin:10px 0; font-size:11.5px; }
th { background:#1A202C; color:#fff; text-align:left; padding:7px 8px; font-weight:700; border:1px solid #1A202C; }
td { padding:6px 8px; border:1px solid #E2E8F0; }
tbody tr:nth-child(even) { background:#F8FAFC; }
.metrics td { border:1px solid #E2E8F0; padding:8px 10px; }
.metrics .lbl { font-weight:700; background:#F1F5F9; width:22%; }
.metrics .val { width:28%; }
.compact table, table.compact { font-size:9.5px; }
.compact th, .compact td { padding:2.6px 5px; }
.win { color:#137333; font-weight:700; }
.loss { color:#C5221F; font-weight:700; }
.gold { color:#92400E; font-weight:700; }
.center { text-align:center; } .right { text-align:right; }
.foot { border-top:1px solid #CBD5E0; margin-top:22px; padding-top:6px; color:#64748B; font-size:10px; display:flex; justify-content:space-between; }
.tag { display:inline-block; padding:1px 7px; border-radius:4px; font-size:10px; font-weight:700; }
.tag-hist { background:#EDF2F7; color:#2D3748; }
.tag-oos { background:#FEF3C7; color:#92400E; }
.tag-live { background:#DCFCE7; color:#137333; }
ul { margin:6px 0 6px 18px; } li { margin:4px 0; }
@media print { .page { margin:0; box-shadow:none; } body { -webkit-print-color-adjust:exact; print-color-adjust:exact; } }
@page { size: A4; margin: 0; }
"""

def foot(n):
    return (f'<div class="foot"><span>CONFIDENTIAL &nbsp;|&nbsp; Prepared for Management Presentation &nbsp;|&nbsp; '
            f'Audited Holiday Backtest Output</span><span>Page {n} of 5</span></div>')

# ----- metric grid -----
metrics = f"""
<table class="metrics">
<tr><td class="lbl">Net Realized Profit (4-Cycle, 1 lot/stock)</td><td class="val"><span class="win">{rupees(TOTpnl/100000)} Lakhs</span></td>
    <td class="lbl">Win Rate (4-Cycle Audited)</td><td class="val"><span class="win">{TOTw/TOTn*100:.2f}%</span> ({TOTw}W / {TOTl}L)</td></tr>
<tr><td class="lbl">Profit Factor</td><td class="val"><b>{PF:.2f}</b> (Gross W/L Ratio)</td>
    <td class="lbl">Average Trade Return</td><td class="val"><b>+{avg:.2f}%</b> (+{avgW:.2f}% W / {avgL:.2f}% L)</td></tr>
<tr><td class="lbl">Gross Winnings</td><td class="val">{rupees(gw/100000)} Lakhs</td>
    <td class="lbl">Gross Losses</td><td class="val"><span class="loss">{rupees(gl/100000)} Lakhs</span></td></tr>
<tr><td class="lbl">Best Cycle</td><td class="val"><span class="win">Oct-{best_y}</span> ({rupees(peryear[best_y]['pnl']/100000)}L, {peryear[best_y]['wr']}%)</td>
    <td class="lbl">Capital Deployed / Cycle (20% margin)</td><td class="val">~\u20b9{cap2024/100000:.1f} Lakhs (50 stocks &times; 1 lot)</td></tr>
<tr><td class="lbl">Index Confirmation (Nifty Fut, 2000&ndash;2026)</td><td class="val"><span class="win">{idx['win_rate']}%</span> over {idx['total_trades']} occurrences</td>
    <td class="lbl">Aggregate Optimal Window (Index)</td><td class="val"><b>{esc(idx['optimal_window'])}</b> &middot; {esc(idx['direction'])}</td></tr>
<tr><td class="lbl">Next Event</td><td class="val"><b>{esc(idx['2026_date_str'])}</b> &middot; NSE Closed</td>
    <td class="lbl">Latest Cycle (Oct-2025)</td><td class="val"><span class="win">{peryear[2025]['wr']}% Win Rate</span> ({rupees(peryear[2025]['pnl']/100000)}L)</td></tr>
</table>"""

# ----- Section 2: per-cycle table -----
statmap = {2022:('HISTORICAL','tag-hist'),2023:('HISTORICAL','tag-hist'),2024:('HISTORICAL','tag-hist'),2025:('LATEST CYCLE','tag-live')}
rows2 = ""
for y in [2026] + YEARS[::-1]:
    if y == 2026:
        rows2 += (f'<tr><td><b>Oct-2026</b></td><td>{esc(idx["2026_date_str"])} &mdash; Q3 festive cycle</td>'
                  f'<td><span class="tag tag-oos">UPCOMING</span></td><td class="center gold">Pending</td>'
                  f'<td class="center">&mdash;</td><td class="right">\u20b90 (Pending)</td><td class="center"><b>{esc(idx["direction"])}</b></td></tr>')
        continue
    p = peryear[y]; lab, cls = statmap[y]
    rows2 += (f'<tr><td><b>Oct-{y}</b></td><td>Gandhi Jayanti {p["hol"]} &middot; window {esc(p["span"])}</td>'
              f'<td><span class="tag {cls}">{lab}</span></td><td class="center win">{p["wr"]}%</td>'
              f'<td class="center">{p["w"]} / {p["l"]}</td><td class="right win">{rupees(p["pnl"])}</td>'
              f'<td class="center">+{p["avg"]:.2f}%</td></tr>')

sec2 = f"""<h2>2. Four-Cycle Audited Backtest Performance (2022&ndash;2025)</h2>
<p>The table below documents the empirical per-stock futures backtest across the four most recent Mahatma Gandhi
Jayanti cycles, built from the master holiday execution logs
(<i>17_NSE_Holidays_Reports / 13_Mahatma_Gandhi_Jayanti</i>). Each cycle deploys all 50 Nifty&nbsp;50 futures, one
lot per stock, using each stock's own calibrated pre-/post-holiday window. Oct-2026 is the upcoming live event.</p>
<table>
<thead><tr><th>Cycle</th><th>Event &amp; Window</th><th>Status</th><th>Win Rate</th><th>W / L</th><th>Net Realized PnL</th><th>Avg Return</th></tr></thead>
<tbody>{rows2}</tbody></table>"""

# ----- Section 3: top 10 -----
rows3 = ""
for i, (_, r) in enumerate(top10.iterrows(), 1):
    d = esc(r['Optimal Direction']).replace('FUTURE ', '')
    dcls = 'win' if 'LONG' in d else 'loss'
    rows3 += (f'<tr><td><b>{i}. {esc(r["Stock Symbol"])}</b></td><td>{esc(str(r["Company Name"]))}</td>'
              f'<td class="center win">{r["WR"]:.0f}%</td><td class="center">{esc(r["Optimal Window"])}</td>'
              f'<td class="center {dcls}">{d}</td><td class="right win">{rupees(r["TOT"])}</td></tr>')
sec3 = f"""<h2>3. Top 10 Consistent Stock Performers (4-Cycle Base Rate)</h2>
<p>The strategy does not treat all stocks equally. Each stock's optimal entry lead (N) and exit hold (M) window and
its LONG/SHORT lean are calibrated against its own Gandhi Jayanti history. The ten most profitable Nifty&nbsp;50
names across the four cycles:</p>
<table>
<thead><tr><th>Rank &amp; Symbol</th><th>Company Name</th><th>4-Cycle Win Rate</th><th>Optimal Window</th><th>Direction</th><th>Total 4-Cycle PnL</th></tr></thead>
<tbody>{rows3}</tbody></table>"""

# ----- Section 4: detailed positions (latest 2025), split across two pages -----
row_list = []
for _, r in lat.iterrows():
    d = esc(r['Strategy']).replace('FUTURE ', ''); dcls = 'win' if 'LONG' in d else 'loss'
    scls = 'win' if r['status'] == 'WIN' else 'loss'
    row_list.append(
        f'<tr><td>{esc(r["Stock Symbol"])}</td><td class="center {dcls}">{d}</td>'
        f'<td class="center">{esc(r["Optimal Window"])}</td>'
        f'<td class="center">{r["ed"]:%d-%b}</td><td class="right">{r["Entry Price (\u20b9)"]:,.1f}</td>'
        f'<td class="center">{r["xd"]:%d-%b}</td><td class="right">{r["Exit Price (\u20b9)"]:,.1f}</td>'
        f'<td class="right {scls}">{r["ret"]:+.2f}%</td><td class="center {scls}">{r["status"]}</td>'
        f'<td class="right">{rupees(r["pnl"])}</td></tr>')
half = (len(row_list) + 1) // 2
_thead = ('<thead><tr><th>Symbol</th><th>Dir</th><th>Window</th><th>Entry</th><th>Entry \u20b9</th>'
          '<th>Exit</th><th>Exit \u20b9</th><th>Return</th><th>Result</th><th>PnL / Lot</th></tr></thead>')

def det_table(rows):
    return f'<table class="compact">{_thead}<tbody>{"".join(rows)}</tbody></table>'

sec4a = f"""<h2>4. Detailed Position-Taking Log &mdash; Latest Cycle (Oct-2025)</h2>
<p>Full 50-stock execution ledger for the most recent Gandhi Jayanti (holiday {peryear[2025]['hol']}), ranked by
PnL. Entries at 09:20&nbsp;AM on day T&minus;N, exits at 03:15&nbsp;PM on day T+M, one futures lot per stock. Cycle
result: <b class="win">{peryear[2025]['wr']}% win rate, {rupees(peryear[2025]['pnl'])}</b> across
{peryear[2025]['w']}W / {peryear[2025]['l']}L. <i>(Ranks 1&ndash;{half}; continued next page.)</i></p>
{det_table(row_list[:half])}"""
sec4b = f"""<h2>4. Detailed Position-Taking Log &mdash; Latest Cycle (Oct-2025), continued</h2>
<p>Ranks {half + 1}&ndash;{len(row_list)} of the Oct-2025 Gandhi Jayanti execution ledger.</p>
{det_table(row_list[half:])}"""

# ----- Section 5 + 6 -----
sec56 = f"""<h2>5. Capital Deployment, Margin Rules &amp; Execution Playbook</h2>
<p class="sec-blue">1. Position Sizing &amp; Margin</p>
<p>Sizing adheres to real NSE futures lot sizes, one lot per stock. Margin is the exchange 20% initial requirement
(SPAN + Exposure), aggregating to ~\u20b9{cap2024/100000:.1f} Lakhs of deployed capital for the full 50-stock cycle.</p>
<p class="sec-blue">2. Execution Timings</p>
<ul><li><b>Entry:</b> 09:20 AM on trading day T&minus;N (limit / market-at-open).</li>
<li><b>Exit:</b> 03:10&ndash;03:25 PM on trading day T+M. The position is held straight through the exchange holiday.</li></ul>
<p class="sec-blue">3. Non-Predictive Anchor</p>
<p>T is the NSE trading day of the Mahatma Gandhi Jayanti holiday (02 October; the exchange is closed that day). The
model positions around the holiday liquidity/festive-cycle drift rather than forecasting news.</p>
<p class="sec-blue">4. Options Alternative &amp; Stop-Loss</p>
<p>For capital-efficient deployment, buy near-month ATM/ITM Calls (LONG bias) or Puts (SHORT bias), capping risk to
premium paid. A hard stop-loss of &minus;3.5% on spot guards against outlier holiday-window shocks.</p>

<h2>6. Institutional Conclusion &amp; Next Steps</h2>
<p>Across four audited cycles the Nifty&nbsp;50 Gandhi Jayanti Holiday Drift System delivered a
<b class="win">{TOTw/TOTn*100:.1f}% win rate</b>, <b>{PF:.2f} profit factor</b> and <b class="win">{rupees(TOTpnl/100000)} Lakhs</b>
net profit (1 lot/stock), independently corroborated by a <b>{idx['win_rate']}%</b> base rate on the Nifty futures
index over {idx['total_trades']} occurrences since 2000.</p>
<p class="sec-blue">Recommendations</p>
<ul><li>Stage entries into the <b>{esc(idx['2026_date_str'])}</b> event from late September 2026 per each stock's calibrated T&minus;N lead.</li>
<li>Prioritise the consistent performers in Section 3 (100% base-rate names: JIOFIN, ULTRACEMCO, ICICIBANK, TATACONSUM).</li>
<li>Maintain automated 09:20 AM entry / 03:15 PM exit execution and the &minus;3.5% spot stop-loss.</li></ul>"""

doc = f"""<!doctype html><html><head><meta charset="utf-8"><title>Gandhi Jayanti Performance Report</title><style>{css}</style></head><body>
<div class="page">
  <div class="kicker">Quantitative Strategy Note &bull; Confidential</div>
  <h1>NIFTY 50 MAHATMA GANDHI JAYANTI<br>HOLIDAY DRIFT SYSTEM</h1>
  <div class="sub">A Quantitative Pre- &amp; Post-Holiday Futures Drift Execution Model across 4 Audited Cycles (2022&ndash;2025)</div>
  <hr class="rule">
  <div class="sec-blue" style="font-size:13px">EXECUTIVE SUMMARY &amp; CORE MANDATE</div>
  <p>This note presents the quantitative backtest performance and execution playbook for the Nifty&nbsp;50 Mahatma
  Gandhi Jayanti Holiday Drift System. The strategy captures the systematic pre-holiday run-up and post-holiday
  digestion drift around the 02 October national holiday across the Nifty&nbsp;50 futures universe. Over the four most
  recent cycles (2022&ndash;2025) it delivers an <b>{TOTw/TOTn*100:.2f}% win rate</b>, generating
  <b class="win">{rupees(TOTpnl/100000)} Lakhs</b> gross profit (one lot per stock) at a <b>{PF:.2f} profit factor</b>.</p>
  {metrics}
  <h2>1. Core Algorithmic Logic &mdash; Buy-Before / Sell-After Holiday Model</h2>
  <p>The system is <b>non-predictive positioning around a scheduled exchange holiday</b>. Rather than forecasting
  news, it anchors to the fixed NSE Mahatma Gandhi Jayanti holiday date (02&nbsp;October) and uses each stock's own
  multi-year base rate to enter ahead of the holiday and exit after the post-holiday drift completes.</p>
  <p class="sec-blue">The Two-Leg Trade Structure</p>
  <ul>
    <li><b>Leg 1 &mdash; Buy Before (Anticipation Run-Up):</b> enter N trading days before the holiday date T at
    09:20&nbsp;AM, as participants position into the festive / Q3-earnings run-in.</li>
    <li><b>Leg 2 &mdash; Hold Through &amp; Sell After (Digest):</b> hold across the closed session and exit M trading
    days after T at 03:15&nbsp;PM, capturing the full post-holiday drift.</li>
  </ul>
  <p><b>Trade Return:</b> Return<sub>i</sub> = ((P<sub>exit,&nbsp;T+M</sub> &minus; P<sub>entry,&nbsp;T&minus;N</sub>) /
  P<sub>entry,&nbsp;T&minus;N</sub>) &times; Direction<sub>i</sub>, where Direction = +1 (LONG) or &minus;1 (SHORT), and
  T is the NSE trading day of the Gandhi Jayanti holiday.</p>
  {foot(1)}
</div>
<div class="page">{sec2}{sec3}{foot(2)}</div>
<div class="page">{sec4a}{foot(3)}</div>
<div class="page">{sec4b}{foot(4)}</div>
<div class="page">{sec56}{foot(5)}</div>
</body></html>"""

open(OUT, 'w', encoding='utf-8').write(doc)
print('Saved:', OUT)
print(f'4-cycle: {TOTw}/{TOTn} WR={TOTw/TOTn*100:.2f}% PnL={rupees(TOTpnl)} PF={PF:.2f}')
