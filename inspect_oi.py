import pandas as pd
import sys
import pathlib

sys.stdout.reconfigure(encoding='utf-8')

base = pathlib.Path('E:/OI_DATA')

print("=" * 70)
print("E:\\OI_DATA — COMPLETE DIRECTORY AUDIT")
print("=" * 70)

all_stocks = sorted([d.name for d in base.iterdir() if d.is_dir()])
print(f"\nTotal Stock Folders: {len(all_stocks)}")
print(f"Stocks: {', '.join(all_stocks)}")
print()

# For each stock, show sub-folder counts and date ranges
print(f"{'Stock':<14} {'Options Files':>14} {'Spot Files':>12} {'First Date':>12} {'Last Date':>12}")
print("-" * 70)

first_date_global = None
last_date_global  = None

for sym in all_stocks:
    sym_dir = base / sym
    opt_files  = sorted((sym_dir / 'options_parquet').glob('*.parquet')) if (sym_dir / 'options_parquet').exists() else []
    spot_files = sorted((sym_dir / 'spot_parquet').glob('*.parquet'))    if (sym_dir / 'spot_parquet').exists()   else []

    first_d = opt_files[0].stem  if opt_files  else 'N/A'
    last_d  = opt_files[-1].stem if opt_files  else 'N/A'

    if opt_files:
        if first_date_global is None or first_d < first_date_global:
            first_date_global = first_d
        if last_date_global is None or last_d > last_date_global:
            last_date_global = last_d

    print(f"{sym:<14} {len(opt_files):>14,} {len(spot_files):>12,} {first_d:>12} {last_d:>12}")

print()
print(f"Global Date Range: {first_date_global}  →  {last_date_global}")
print()

# Deep-inspect the latest HDFCBANK options parquet
print("=" * 70)
print("SAMPLE FILE SCHEMA — HDFCBANK options_parquet (latest date)")
print("=" * 70)

sample_file = sorted((base / 'HDFCBANK' / 'options_parquet').glob('*.parquet'))[-1]
print(f"\nFile: {sample_file.name}  ({sample_file.stat().st_size / 1024:.1f} KB)")

df = pd.read_parquet(sample_file)
print(f"Shape: {df.shape[0]:,} rows  x  {df.shape[1]} columns")
print(f"\nColumns & dtypes:")
for col, dt in df.dtypes.items():
    print(f"  {col:<20} {str(dt)}")

print(f"\nUnique ExpiryDates in this file: {sorted(df['ExpiryDate'].unique())}")
print(f"Unique Types: {sorted(df['Type'].unique())}")
print(f"Strike range: {df['Strike'].min():.0f}  to  {df['Strike'].max():.0f}")
print(f"Time range:   {df['Time'].min()}  to  {df['Time'].max()}")

print(f"\n--- Sample rows at 15:29:59 (market close) ---")
close_df = df[df['Time'] == '15:29:59']
sample_close = close_df.groupby(['ExpiryDate', 'Type']).first().reset_index()
print(sample_close[['Ticker', 'Time', 'ExpiryDate', 'Type', 'Strike', 'Close', 'OpenInterest']].head(10).to_string(index=False))

# Spot parquet sample
print()
print("=" * 70)
print("SAMPLE FILE SCHEMA — HDFCBANK spot_parquet (latest date)")
print("=" * 70)
spot_file = sorted((base / 'HDFCBANK' / 'spot_parquet').glob('*.parquet'))[-1]
print(f"\nFile: {spot_file.name}  ({spot_file.stat().st_size / 1024:.1f} KB)")
df_s = pd.read_parquet(spot_file)
print(f"Shape: {df_s.shape[0]:,} rows  x  {df_s.shape[1]} columns")
print(f"Columns: {df_s.columns.tolist()}")
close_s = df_s[df_s['Time'] == '15:29:59']
if close_s.empty:
    close_s = df_s.sort_values('Time').tail(1)
print(f"\nSpot Close at 15:29:59: {float(close_s.iloc[0]['Close']):.2f}")
