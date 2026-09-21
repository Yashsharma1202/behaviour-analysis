import pandas as pd
import pathlib
import sys
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

sys.stdout.reconfigure(encoding='utf-8')

# ------------------- Configuration -------------------
# Root folder where each stock has futures_parquet and spot_parquet subfolders
FUTURES_ROOT = pathlib.Path('D:/behaviour analysis/OI_DATA')
# Output master workbook path
OUTPUT_PATH = pathlib.Path('D:/behaviour analysis/Futures_Only_Master.xlsx')

# ------------------- Helper functions -------------------

def get_latest_parquet(folder: pathlib.Path):
    """Return the most recent .parquet file in a folder, or None."""
    files = sorted(folder.glob('*.parquet'))
    return files[-1] if files else None

def get_close_price(df: pd.DataFrame) -> float:
    """Return closing price at 15:29:59 if present; otherwise last row's Close."""
    if 'Time' in df.columns:
        row = df[df['Time'] == '15:29:59']
        if not row.empty:
            return float(row.iloc[0]['Close'])
    # fallback to last row
    return float(df.iloc[-1]['Close'])

def load_futures_price(stock: str, date_str: str):
    """Load futures close price for a given stock/date.
    Returns (price, expiry_date) – expiry may be empty string if not present.
    """
    stock_dir = FUTURES_ROOT / stock / 'futures_parquet'
    fpath = stock_dir / f"{date_str}.parquet"
    if fpath.exists():
        df = pd.read_parquet(fpath)
        price = get_close_price(df)
        expiry = ''
        if 'ExpiryDate' in df.columns:
            expiry = sorted(df['ExpiryDate'].unique())[0] if df['ExpiryDate'].nunique() else ''
        return price, expiry
    return None, ''

def load_spot_price(stock: str, date_str: str):
    """Fallback: load spot close price if futures file missing."""
    stock_dir = FUTURES_ROOT / stock / 'spot_parquet'
    fpath = stock_dir / f"{date_str}.parquet"
    if fpath.exists():
        df = pd.read_parquet(fpath)
        return get_close_price(df)
    return None

# ------------------- Build master dataset -------------------
records = []
# Iterate over all stock directories
for stock_dir in FUTURES_ROOT.iterdir():
    if not stock_dir.is_dir():
        continue
    stock = stock_dir.name
    futures_folder = stock_dir / 'futures_parquet'
    if not futures_folder.is_dir():
        continue
    # Process each parquet file (assumed naming: YYYY-MM-DD.parquet)
    for parquet_path in sorted(futures_folder.glob('*.parquet')):
        date_str = parquet_path.stem  # filename without .parquet
        try:
            df = pd.read_parquet(parquet_path)
        except Exception:
            continue
        price = get_close_price(df)
        expiry = ''
        if 'ExpiryDate' in df.columns:
            expiry = sorted(df['ExpiryDate'].unique())[0] if df['ExpiryDate'].nunique() else ''
        # Spot fallback if needed (optional, not required for futures‑only view)
        spot_price = load_spot_price(stock, date_str)
        records.append([
            stock,
            date_str,
            price,
            spot_price if spot_price is not None else '',
            expiry,
            '',  # Position (will be filled later)
            '',  # Entry Date
            '',  # Exit Date
            '',  # Expected Return (%)
            '',  # Return We Get (%)
            '',  # Booked P&L
            ''   # Quarterly Result Date
        ])

# Create DataFrame
columns = [
    'Symbol',
    'Date',
    'Futures Close Price (₹)',
    'Spot Close Price (₹)',
    'Futures Expiry Date',
    'Position',            # LONG / SHORT (blank for now)
    'Entry Date',
    'Exit Date',
    'Expected Return (%)',
    'Return We Get (%)',
    'Booked P&L (₹)',
    'Quarterly Result Date'
]
master_df = pd.DataFrame(records, columns=columns)

# ------------------- Write to Excel with styling -------------------
wb = Workbook()
ws = wb.active
ws.title = 'Futures_Only_Master'
ws.append(columns)
# Header style – dark teal background, white bold text
header_font = Font(bold=True, color='FFFFFFFF')
header_fill = PatternFill(start_color='FF006688', end_color='FF006688', fill_type='solid')
for col_idx in range(1, len(columns) + 1):
    cell = ws.cell(row=1, column=col_idx)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal='center')
# Data rows
for row in master_df.itertuples(index=False, name=None):
    ws.append(row)
# Auto‑size columns
for col in ws.columns:
    max_len = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
    ws.column_dimensions[col[0].column_letter].width = max_len + 3

# Safe save (handles possible permission errors)
def safe_save(wb_obj, path):
    try:
        wb_obj.save(path)
        return True
    except PermissionError:
        for i in range(2, 6):
            versioned = path.with_name(f"{path.stem}_v{i}{path.suffix}")
            try:
                wb_obj.save(versioned)
                print(f'⚠️ Permission error on {path.name}; saved as {versioned.name}')
                return True
            except PermissionError:
                continue
        print(f'❌ Could not save {path.name} after multiple attempts.')
        return False

if safe_save(wb, OUTPUT_PATH):
    print(f'✅ Futures‑only master Excel created at: {OUTPUT_PATH}')
    print(f'Rows written: {len(master_df)}')
else:
    print('❌ Failed to write master workbook.')
