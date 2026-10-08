"""
DUSSEHRA 2025 - POINT-IN-TIME (walk-forward) DETAILED RESULTS, ALL QUALIFIED STOCKS.
Each stock's 2025 window+direction was chosen using ONLY years before 2025
(scraped_parquet/_tmp_dussehra_walkforward.pkl). P&L rebuilt from real near-month
futures closes, same method/lot overrides as build_dussehra_2025_fo_top50_full_report.py.
ADDITIVE: writes a new report only.
"""
import json, pickle, statistics as st
from datetime import date
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

BASE = r'D:\behaviour analysis'
OUT = BASE + r'\Dussehra_2025_PIT_Detailed_Report.xlsx'

results = pickle.load(open(BASE + r'\scraped_parquet\_tmp_dussehra_walkforward.pkl', 'rb'))
meta = {f['symbol']: f for f in json.load(open(BASE + r'\dashboard_data\fo_stocks_211.json', encoding='utf-8'))}

fut = pd.read_parquet(BASE + r'\scraped_parquet\fo_futures_near_month_continuous.parquet', columns=['Date', 'Instrument', 'Open', 'Close'])
fut['Date'] = pd.to_datetime(fut['Date']).dt.date
fut_closes = {s: g.sort_values('Date')[['Date', 'Close', 'Open']].values.tolist() for s, g in fut.groupby('Instrument')}

beh = json.load(open(BASE + r'\dashboard_data\nifty_futures_holiday_behaviour.json', encoding='utf-8'))
dh = [h for h in beh['holidays'] if 'Dussehra' in h['name']][0]
anchor_2025 = next(date.fromisoformat(t['holiday_date']) for t in dh['trades'] if t['holiday_date'].startswith('2025'))

# Point-in-time lot overrides (NSE FAOP70554, eff. 29-Oct-2025), as in the 2025 top-50 report
LOT_OVERRIDE = {'TVSMOTOR': 350, 'LAURUSLABS': 1700, 'SIEMENS': 125, 'RECLTD': 1275}
def lot_for(sym):
    return LOT_OVERRIDE.get(sym) or meta.get(sym, {}).get('lot_size') or 500

rows, paths = [], {}
for sym, r in results.items():
    t = next((x for x in r.get('oos_trades', []) if x['year'] == 2025), None)
    if not t or sym not in fut_closes:
        continue
    a = fut_closes[sym]
    ib = max((i for i, row in enumerate(a) if row[0] <= anchor_2025), default=-1)
    if ib < 0:
        continue
    en_idx, ex_idx = ib - t['bo'], ib + t['so']
    if en_idx < 0 or ex_idx >= len(a):
        continue
    sign = 1 if t['side'] == 'LONG' else -1
    lot = lot_for(sym)
    en_date, en_close, en_px = a[en_idx]
    exit_date, exit_close, _ = a[ex_idx]
    # compounding daily path, same as the 2025 report (moves over 20% in a day ignored as data errors)
    g, daily = 1.0, []
    if en_px and abs((en_close - en_px) / en_px) <= 0.20:
        g *= (1 + (en_close - en_px) / en_px)
    daily.append((en_date, (g - 1) * sign * en_px * lot))
    for k in range(en_idx + 1, ex_idx + 1):
        p0, p1 = a[k - 1][1], a[k][1]
        if p0 and abs((p1 - p0) / p0) <= 0.20:
            g *= (1 + (p1 - p0) / p0)
        daily.append((a[k][0], (g - 1) * sign * en_px * lot))
    vals = [v for _, v in daily]
    peak, maxdd = vals[0], 0.0
    for v in vals:
        peak = max(peak, v); maxdd = min(maxdd, v - peak)
    chg = [vals[i] - vals[i - 1] for i in range(1, len(vals))]
    sharpe = round(st.mean(chg) / st.pstdev(chg), 2) if len(chg) >= 2 and st.pstdev(chg) > 0 else None
    pnl = vals[-1]
    raw_ret = (exit_close - en_px) / en_px * 100 * sign if en_px else None
    m = meta.get(sym, {})
    rows.append({
        'Symbol': sym, 'Name': m.get('name', sym), 'Sector': m.get('sector', ''),
        'Direction': t['side'], 'Window': f"T-{t['bo']} to T+{t['so']}",
        'Entry Date': en_date, 'Entry Px (fut close)': round(en_px, 2),
        'Exit Date': exit_date, 'Exit Px (fut close)': round(exit_close, 2),
        'Raw Move %': round((exit_close - en_px) / en_px * 100, 2) if en_px else None,
        'Direction-Adj Return %': round(raw_ret, 2) if raw_ret is not None else None,
        'Pickle OOS Ret % (cross-check)': t['ret'],
        'Lot Size': lot, 'Notional per Lot': round(en_px * lot, 0),
        'Indicative Margin (20%)': round(en_px * lot * 0.20, 0),
        'P&L Rs': round(pnl, 0), 'Max DD Rs': round(maxdd, 0), 'Daily Sharpe': sharpe,
        'Prior-yr OOS n': r.get('oos_n'), 'Low Confidence': 'YES' if r.get('low_confidence') else 'NO',
    })
    paths[sym] = daily

df = pd.DataFrame(rows).sort_values('P&L Rs', ascending=False).reset_index(drop=True)
df.insert(0, 'PnL Rank', range(1, len(df) + 1))

# Daily cumulative P&L grid (stocks as columns)
all_dates = sorted({d for p in paths.values() for d, _ in p})
grid = pd.DataFrame(index=all_dates)
for sym, p in paths.items():
    s = pd.Series(dict(p)).reindex(all_dates).ffill().fillna(0.0)
    grid[sym] = s.round(0)
grid['TOTAL (all stocks)'] = grid.sum(axis=1).round(0)
grid.index.name = 'Date'

wins = df[df['P&L Rs'] > 0]
port = grid['TOTAL (all stocks)']
summary = [
    ('Qualified stocks with a 2025 point-in-time trade', len(df)),
    ('Long / Short split', f"{(df.Direction == 'LONG').sum()} / {(df.Direction == 'SHORT').sum()}"),
    ('Winning trades (P&L > 0)', f"{len(wins)} ({len(wins) / len(df) * 100:.1f}%)"),
    ('Total P&L, 1 lot each (Rs)', round(df['P&L Rs'].sum(), 0)),
    ('Average P&L per trade (Rs)', round(df['P&L Rs'].mean(), 0)),
    ('Total indicative margin, 1 lot each (Rs)', round(df['Indicative Margin (20%)'].sum(), 0)),
    ('Portfolio max DD (Rs, all stocks combined)', round(min(0, (port - port.cummax()).min()), 0)),
    ('Lot-size overrides applied', ', '.join(f'{k}={v}' for k, v in LOT_OVERRIDE.items())),
]

wb = Workbook()
ws = wb.active; ws.title = 'Overview'
ws.append(['DUSSEHRA 2025 - POINT-IN-TIME (WALK-FORWARD) DETAILED RESULTS'])
ws['A1'].font = Font(bold=True, size=13)
ws.append([])
for k, v in summary:
    ws.append([k, v])
ws.append([])
for line in [
    "Method: each stock's 2025 window and direction were chosen using ONLY years before 2025 (walk-forward 8x8 grid).",
    'P&L: real near-month futures closes, compounded daily, 1 lot, direction-adjusted. Moves over 20% in a day are ignored as data errors.',
    "Cross-check column compares the rebuilt Direction-Adj Return with the pickle's stored OOS return.",
    'Lot sizes: local fo_stocks_211.json with 2025 point-in-time overrides. Not verified against the NSE circular.',
    'Margin is indicative (20% of notional), not exchange SPAN.',
]:
    ws.append([line])
ws.column_dimensions['A'].width = 48; ws.column_dimensions['B'].width = 60


def write_table(wsx, frame, title):
    wsx.append([title]); wsx['A1'].font = Font(bold=True, size=12); wsx.append([])
    wsx.append(list(frame.columns))
    hr = wsx.max_row
    for c in wsx[hr]:
        c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='1F4E78')
        c.alignment = Alignment(horizontal='center', wrap_text=True)
    for r in frame.itertuples(index=False):
        wsx.append([None if pd.isna(v) else (v.date() if hasattr(v, 'date') and not isinstance(v, str) else v) for v in r])
    wsx.freeze_panes = wsx.cell(row=hr + 1, column=2)
    for i, col in enumerate(frame.columns, 1):
        wsx.column_dimensions[wsx.cell(row=hr, column=i).column_letter].width = max(12, min(28, len(str(col)) + 2))


write_table(wb.create_sheet('PIT_2025_Trades'), df, 'ALL QUALIFIED STOCKS - 2025 POINT-IN-TIME TRADES (ranked by P&L)')
write_table(wb.create_sheet('Daily_Cum_PnL'), grid.reset_index(), 'DAILY CUMULATIVE P&L (Rs, 1 lot each, stocks as columns)')
wb.save(OUT)

print(df[['PnL Rank', 'Symbol', 'Direction', 'Window', 'Entry Date', 'Exit Date', 'Direction-Adj Return %', 'Pickle OOS Ret % (cross-check)', 'P&L Rs']].head(10).to_string(index=False))
print(len(df), 'stocks')
for k, v in summary:
    print(f'{k}: {v}')
mis = df[(df['Direction-Adj Return %'] - df['Pickle OOS Ret % (cross-check)']).abs() > 1]
print('cross-check mismatches (>1pt):', len(mis))
