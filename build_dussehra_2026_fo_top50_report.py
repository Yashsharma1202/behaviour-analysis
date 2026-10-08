"""
TOP 50 F&O STOCKS - UPCOMING DUSSEHRA (20-Oct-2026) - trade sheet.
Source: FO211_Dussehra_WalkForward_Ranked_Report.xlsx (point-in-time walk-forward,
ranked by OOS avg return). Lot size from Stock_Data_213/lot_sizes.csv; margin is an
indicative 20% of lot notional at latest downloaded close (NOT exchange SPAN).
ADDITIVE: writes a new report only.
"""
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

BASE = r'D:\behaviour analysis'
SRC = BASE + r'\FO211_Dussehra_WalkForward_Ranked_Report.xlsx'
OUT = BASE + r'\Dussehra_2026_FO_Top50_WinRate_LotFixed.xlsx'

rank = pd.read_excel(SRC, sheet_name='WalkForward_Ranked', header=3)
rank = rank[pd.to_numeric(rank['Rank'], errors='coerce').notna()].copy()
rank = rank.sort_values(['OOS Win Rate %', 'OOS Avg Return %', 'OOS Trades (n)'], ascending=False).head(50)
rank['Rank'] = range(1, len(rank) + 1)

sched = pd.read_excel(SRC, sheet_name='Position_Schedule_2026', header=3)
sched = sched[['Symbol', 'Entry Date', 'Exit Date']].drop_duplicates('Symbol')

lots = pd.read_csv(BASE + r'\Stock_Data_213\lot_sizes.csv')[['Symbol', 'LotSize']]
px = pd.read_parquet(BASE + r'\Stock_Data_213\all_stocks_daily_ohlc.parquet', columns=['Symbol', 'Date', 'Close'])
last = px.sort_values('Date').groupby('Symbol').tail(1).rename(columns={'Date': 'LastDate', 'Close': 'LastClose'})

df = rank.merge(sched, on='Symbol', how='left').merge(lots, on='Symbol', how='left').merge(last, on='Symbol', how='left')
df['NotionalPerLot'] = (df.LotSize * df.LastClose).round(0)
df['IndicativeMargin_20pct'] = (0.20 * df.NotionalPerLot).round(0)

cols = ['Rank', 'Symbol', 'Upcoming Direction', 'Window', 'Entry Date', 'Exit Date',
        'OOS Win Rate %', 'OOS Avg Return %', 'OOS Trades (n)', 'Confidence',
        'LotSize', 'LastClose', 'NotionalPerLot', 'IndicativeMargin_20pct']
df = df[cols]

wb = Workbook(); ws = wb.active; ws.title = 'Top50_Trade_Sheet'
ws.append(['TOP 50 F&O STOCKS - DUSSEHRA 20-Oct-2026 (walk-forward ranked, point-in-time)'])
ws['A1'].font = Font(bold=True, size=13)
ws.append(['Ranked by out-of-sample avg return. Lot sizes UNVERIFIED vs NSE; margin is indicative (20% of notional), not SPAN.'])
ws.append([])
ws.append(cols)
hdr_row = ws.max_row
for c in ws[hdr_row]:
    c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='1F4E78')
    c.alignment = Alignment(horizontal='center', wrap_text=True)
for r in df.itertuples(index=False):
    ws.append([None if pd.isna(v) else v for v in r])
for col, w in zip('ABCDEFGHIJKLMN', [7, 13, 12, 12, 16, 16, 12, 14, 12, 12, 10, 12, 16, 20]):
    ws.column_dimensions[col].width = w
for row in ws.iter_rows(min_row=hdr_row + 1):
    row[11].number_format = '#,##0.00'; row[12].number_format = '#,##0'; row[13].number_format = '#,##0'
ws.freeze_panes = ws.cell(row=hdr_row + 1, column=1)

notes = wb.create_sheet('Notes')
for line in [
    'Dussehra 20-Oct-2026 - Top 50 F&O stocks',
    '',
    'Selection: FO211_Dussehra_WalkForward_Ranked_Report.xlsx (walk-forward 8x8 grid, only prior years used).',
    'Ranking: OOS win rate %, then OOS avg return %, then trade count. Confidence LOW = fewer than 3 OOS trades.',
    'LotSize: NSE fo_mktlots.csv, OCT-26 column (Stock_Data_213/fo_mktlots_NSE.csv).',
    'LastClose: latest Yahoo Finance daily Close. NotionalPerLot = LotSize x LastClose.',
    'IndicativeMargin_20pct = 20% of NotionalPerLot. Not exchange SPAN+exposure.',
]:
    notes.append([line])
notes.column_dimensions['A'].width = 120
wb.save(OUT)
print(df[['Rank', 'Symbol', 'Upcoming Direction', 'Entry Date', 'Exit Date', 'OOS Win Rate %', 'LotSize', 'IndicativeMargin_20pct']].to_string(index=False))
print('missing lot:', df.LotSize.isna().sum(), 'missing sched:', df['Entry Date'].isna().sum())
