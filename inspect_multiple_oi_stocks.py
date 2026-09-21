import os, pathlib, sys, pandas as pd

sys.stdout.reconfigure(errors='replace')

base_dir = pathlib.Path(r'D:\behaviour analysis\OI_DATA')

stocks_to_check = ['RELIANCE', 'INFY', 'TCS', 'ADANIENT', 'HDFCBANK']

print("==========================================================================================")
print("INSPECTING MULTIPLE MAJOR STOCKS IN D:\\behaviour analysis\\OI_DATA")
print("==========================================================================================")

for sym in stocks_to_check:
    stock_dir = base_dir / sym
    print(f"\n==================== STOCK: {sym} ====================")
    
    for category in ['spot_parquet', 'futures_parquet', 'options_parquet']:
        folder_path = stock_dir / category
        if folder_path.exists():
            files = sorted([f for f in os.listdir(folder_path) if f.endswith('.parquet')])
            print(f"  📁 Category: {category:17s} -> Found {len(files):,}.parquet files")
            if files:
                sample_file = folder_path / files[0]
                try:
                    df = pd.read_parquet(sample_file)
                    print(f"     Sample ({sample_file.name}): {len(df):,} rows x {len(df.columns)} cols | Date Range: {df['Date'].min()} | Ticker Sample: {df['Ticker'].iloc[0]}")
                except Exception as e:
                    print(f"     Error reading {sample_file.name}: {e}")

print("\n==========================================================================================")
