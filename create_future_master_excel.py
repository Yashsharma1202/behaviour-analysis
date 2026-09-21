import pandas as pd
import pathlib
import sys
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = pathlib.Path('D:/behaviour analysis/OI_DATA')
OUTPUT_PATH = pathlib.Path('D:/behaviour analysis/Future_Master.xlsx')

if not BASE_DIR.exists():
    print(f'❗ Base directory not found: {BASE_DIR}')
    sys.exit(1)

# Helper to get latest parquet file in a folder
def latest_parquet(folder: pathlib.Path):
    files = sorted(folder.glob('*.parquet'))
    return files[-1] if files else None

# Prepare workbook
wb = Workbook()
ws = wb.active
ws.title = 'Future_Master'

# Header columns (similar to earlier master sheet but simplified)
headers = [
    'Stock',
    'Sector',               # placeholder – could be filled later from a static mapping if available
    'Spot Price (₹)',
    'ATM Strike (₹)',
    'Expiry Date',
    'Call Premium (₹)',
    'Put Premium (₹)',
    'Synthetic Long P&L (₹)',
    'Synthetic Short P&L (₹)',
    'Data Date'            # the date of the parquet used
]
ws.append(headers)
# Apply simple styling to header
header_font = Font(bold=True, color='FFFFFFFF')
header_fill = PatternFill(start_color='FF333399', end_color='FF333399', fill_type='solid')
for col in range(1, len(headers)+1):
    cell = ws.cell(row=1, column=col)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal='center')

# Iterate each stock folder
for stock_dir in sorted([d for d in BASE_DIR.iterdir() if d.is_dir()]):
    stock = stock_dir.name
    opt_dir = stock_dir / 'options_parquet'
    spot_dir = stock_dir / 'spot_parquet'
    opt_file = latest_parquet(opt_dir)
    spot_file = latest_parquet(spot_dir)
    if not opt_file or not spot_file:
        # skip if missing data
        continue
    # Load dataframes
    df_opt = pd.read_parquet(opt_file)
    df_spot = pd.read_parquet(spot_file)
    # Spot close price (15:29:59 if exists)
    spot_close_row = df_spot[df_spot['Time'] == '15:29:59']
    if spot_close_row.empty:
        spot_close_row = df_spot.sort_values('Time').tail(1)
    spot_price = float(spot_close_row.iloc[0]['Close'])
    # Determine ATM strike (nearest 50 Rs)
    atm_strike = round(spot_price / 50) * 50
    # Determine closest expiry that is >= today (use the minimum expiry date in the file)
    today_str = datetime.now().strftime('%Y-%m-%d')
    # Filter expiries that are >= today
    expiries = sorted(df_opt['ExpiryDate'].unique())
    chosen_expiry = None
    for exp in expiries:
        if exp >= today_str:
            chosen_expiry = exp
            break
    if not chosen_expiry:
        # fallback to the earliest expiry
        chosen_expiry = expiries[0]
    # Get Call and Put rows for ATM strike and chosen expiry
    call_row = df_opt[(df_opt['Strike'] == atm_strike) & (df_opt['Type'] == 'CE') & (df_opt['ExpiryDate'] == chosen_expiry)]
    put_row = df_opt[(df_opt['Strike'] == atm_strike) & (df_opt['Type'] == 'PE') & (df_opt['ExpiryDate'] == chosen_expiry)]
    # Use market close price if available
    call_price = None
    put_price = None
    if not call_row.empty:
        cr = call_row[call_row['Time'] == '15:29:59']
        if cr.empty:
            cr = call_row.sort_values('Time').tail(1)
        call_price = float(cr.iloc[0]['Close'])
    if not put_row.empty:
        pr = put_row[put_row['Time'] == '15:29:59']
        if pr.empty:
            pr = put_row.sort_values('Time').tail(1)
        put_price = float(pr.iloc[0]['Close'])
    # If either premium missing, skip this stock (cannot compute synthetic)
    if call_price is None or put_price is None:
        continue
    # Synthetic P&L calculations (simple – difference between call and put)
    synthetic_long = call_price - put_price
    synthetic_short = put_price - call_price
    # Append row
    ws.append([
        stock,
        '',  # Sector placeholder
        spot_price,
        atm_strike,
        chosen_expiry,
        call_price,
        put_price,
        synthetic_long,
        synthetic_short,
        opt_file.stem  # date string of the parquet used
    ])

# Adjust column widths for readability
for col in ws.columns:
    max_length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
    adjusted_width = (max_length + 2)
    ws.column_dimensions[col[0].column_letter].width = adjusted_width

# Save workbook
wb.save(OUTPUT_PATH)
print('✅ Future master Excel created at:', OUTPUT_PATH)
print('Rows written (stocks processed):', ws.max_row - 1)
