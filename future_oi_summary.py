import pandas as pd
import pathlib
import sys
from datetime import datetime

sys.stdout.reconfigure(encoding='utf-8')

# Base directory (the duplicate OI_DATA under the workspace)
BASE_DIR = pathlib.Path('D:/behaviour analysis/OI_DATA')

if not BASE_DIR.exists():
    print(f'❗ Directory not found: {BASE_DIR}')
    sys.exit(1)

# Prepare a list to collect summary rows
summary_rows = []

# Helper to get latest parquet file in a folder
def latest_parquet(folder: pathlib.Path) -> pathlib.Path | None:
    files = sorted(folder.glob('*.parquet'))
    return files[-1] if files else None

# Iterate over each stock folder
for stock_dir in sorted([d for d in BASE_DIR.iterdir() if d.is_dir()]):
    stock = stock_dir.name
    opt_dir = stock_dir / 'options_parquet'
    spot_dir = stock_dir / 'spot_parquet'

    opt_file = latest_parquet(opt_dir) if opt_dir.exists() else None
    spot_file = latest_parquet(spot_dir) if spot_dir.exists() else None

    # Default placeholders
    opt_date = ''
    spot_date = ''
    opt_rows = ''
    spot_rows = ''
    expiry_dates = ''
    sample_option = ''
    sample_spot = ''

    if opt_file:
        opt_date = opt_file.stem
        df_opt = pd.read_parquet(opt_file)
        opt_rows = df_opt.shape[0]
        # Unique expiries (just show a few)
        uniq_exp = sorted(df_opt['ExpiryDate'].unique())[:5]
        expiry_dates = ', '.join(uniq_exp)
        # Grab a sample row (first CE at market close if exists)
        sample = df_opt[(df_opt['Time'] == '15:29:59') & (df_opt['Type'] == 'CE')]
        if not sample.empty:
            sample_option = f"CE {sample.iloc[0]['Strike']} @ {sample.iloc[0]['Close']}"
        else:
            # fallback to any row
            sample_option = f"{df_opt.iloc[0]['Type']} {df_opt.iloc[0]['Strike']} @ {df_opt.iloc[0]['Close']}"

    if spot_file:
        spot_date = spot_file.stem
        df_spot = pd.read_parquet(spot_file)
        spot_rows = df_spot.shape[0]
        # Spot close price at market close
        spot_close = df_spot[df_spot['Time'] == '15:29:59']
        if not spot_close.empty:
            sample_spot = f"{spot_close.iloc[0]['Close']:.2f}"
        else:
            sample_spot = f"{df_spot.iloc[-1]['Close']:.2f}"

    summary_rows.append({
        'Stock': stock,
        'Options_Date': opt_date,
        'Options_Rows': opt_rows,
        'Spot_Date': spot_date,
        'Spot_Rows': spot_rows,
        'Expiry_Dates_Sample': expiry_dates,
        'Sample_Option': sample_option,
        'Spot_Close': sample_spot,
    })

# Build DataFrame
summary_df = pd.DataFrame(summary_rows)

# Write to Excel – one sheet with the summary
output_path = pathlib.Path('D:/behaviour analysis/Future_OI_Summary.xlsx')
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    summary_df.to_excel(writer, index=False, sheet_name='Summary')
    # Also add a timestamp sheet
    ts_df = pd.DataFrame({
        'Generated_On': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')]
    })
    ts_df.to_excel(writer, index=False, sheet_name='Metadata')

print('✅ Summary Excel created at:', output_path)
print('Rows written:', len(summary_df))
print('Columns:', list(summary_df.columns))
