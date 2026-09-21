import pandas as pd
import pathlib
import sys
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

sys.stdout.reconfigure(encoding='utf-8')

# Base directory containing OI data for each stock
BASE_DIR = pathlib.Path('D:/behaviour analysis/OI_DATA')
OUTPUT_PATH = pathlib.Path('D:/behaviour analysis/Future_Futures_Master.xlsx')

if not BASE_DIR.exists():
    print(f'❗ Base directory not found: {BASE_DIR}')
    sys.exit(1)

# Helper to get the latest parquet file inside a folder
def latest_parquet(folder: pathlib.Path):
    files = sorted(folder.glob('*.parquet'))
    return files[-1] if files else None

# Helper to read a parquet and return the close price at market close (15:29:59)
def get_close_price(df: pd.DataFrame):
    close_row = df[df['Time'] == '15:29:59']
    if close_row.empty:
        close_row = df.sort_values('Time').tail(1)
    return float(close_row.iloc[0]['Close'])

# Prepare workbook and header styling
wb = Workbook()
ws = wb.active
ws.title = 'Futures_Master'

headers = [
    'Stock',
    'Spot Price (₹)',
    'Futures Expiry Date',
    'Futures Close (₹)',
    'Daily Change (%)',
    'Data Date'  # parquet file date used for futures data
]
ws.append(headers)
header_font = Font(bold=True, color='FFFFFFFF')
header_fill = PatternFill(start_color='FF004488', end_color='FF004488', fill_type='solid')
for col in range(1, len(headers) + 1):
    cell = ws.cell(row=1, column=col)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal='center')

today_str = datetime.now().strftime('%Y-%m-%d')

processed = 0
for stock_dir in sorted([d for d in BASE_DIR.iterdir() if d.is_dir()]):
    stock = stock_dir.name
    # Spot data – always present
    spot_dir = stock_dir / 'spot_parquet'
    spot_file = latest_parquet(spot_dir)
    if not spot_file:
        print(f'⚠️ No spot data for {stock}, skipping')
        continue
    df_spot = pd.read_parquet(spot_file)
    spot_price = get_close_price(df_spot)

    # Futures data – may exist under futures_parquet; if not fall back to spot as proxy
    futures_dir = stock_dir / 'futures_parquet'
    futures_file = latest_parquet(futures_dir) if futures_dir.exists() else None
    if futures_file:
        df_fut = pd.read_parquet(futures_file)
        # Determine expiry date >= today (choose earliest qualifying)
        expiry_dates = sorted(df_fut['ExpiryDate'].unique())
        chosen_expiry = None
        for exp in expiry_dates:
            if exp >= today_str:
                chosen_expiry = exp
                break
        if not chosen_expiry:
            chosen_expiry = expiry_dates[0] if expiry_dates else ''
        # Futures close price for the chosen expiry (use Close column, ignoring Type)
        # If multiple rows, pick the one with Time == 15:29:59
        fut_price = get_close_price(df_fut)
        # Compute daily change if we have yesterday's file
        daily_change = ''
        prev_date = (datetime.strptime(futures_file.stem, '%Y-%m-%d') - timedelta(days=1)).strftime('%Y-%m-%d')
        prev_file = futures_dir / f'{prev_date}.parquet'
        if prev_file.exists():
            df_prev = pd.read_parquet(prev_file)
            prev_close = get_close_price(df_prev)
            daily_change = ((fut_price - prev_close) / prev_close) * 100
            daily_change = round(daily_change, 2)
    else:
        # No futures parquet – use spot price as a proxy for futures close
        fut_price = spot_price
        chosen_expiry = ''
        daily_change = ''

    ws.append([
        stock,
        spot_price,
        chosen_expiry,
        fut_price,
        daily_change,
        futures_file.stem if futures_file else spot_file.stem
    ])
    processed += 1

# Auto‑size columns
for col in ws.columns:
    max_len = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
    ws.column_dimensions[col[0].column_letter].width = max_len + 2

wb.save(OUTPUT_PATH)
print(f'✅ Futures‑only master Excel created at: {OUTPUT_PATH}')
print(f'Rows written (stocks processed): {processed}')
