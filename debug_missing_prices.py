import pandas as pd, pathlib, sys

sys.stdout.reconfigure(errors='replace')

BASE_DIR = pathlib.Path('D:/behaviour analysis')
OI_ROOT = BASE_DIR / 'OI_DATA'
OPTIONS_MASTER_PATH = BASE_DIR / 'Nifty50_12_Quarters_Options_OI_Master_v4.xlsx'

# List stock folders in OI_ROOT
oi_stocks = sorted([d.name for d in OI_ROOT.iterdir() if d.is_dir()])
print(f"Total stock folders in OI_DATA: {len(oi_stocks)}")
print("Sample stock folders:", oi_stocks[:15])

# Check trades with 0 price
df_trades = pd.read_excel(OPTIONS_MASTER_PATH, sheet_name='All_12Q_Options_Trades', engine='openpyxl')

missing_stocks = set()
for idx, row in df_trades.iterrows():
    sym = str(row['Symbol']).strip()
    entry_dt = pd.to_datetime(row['Entry Date']).strftime('%Y-%m-%d')
    exit_dt = pd.to_datetime(row['Exit Date']).strftime('%Y-%m-%d')
    
    stock_dir = OI_ROOT / sym
    if not stock_dir.exists():
        # Try matching ignoring case or special chars
        match = next((s for s in oi_stocks if s.upper() == sym.upper() or s.upper().replace('&','').replace('-','') == sym.upper().replace('&','').replace('-','')), None)
        print(f"Stock dir missing: '{sym}' -> Matched in OI_DATA: '{match}'")
        missing_stocks.add(sym)
    else:
        # Check if entry_dt parquet exists
        fut_file = stock_dir / 'futures_parquet' / f"{entry_dt}.parquet"
        spot_file = stock_dir / 'spot_parquet' / f"{entry_dt}.parquet"
        if not fut_file.exists() and not spot_file.exists():
            # Check available files around that date
            all_futs = sorted([f.stem for f in (stock_dir / 'futures_parquet').glob('*.parquet')]) if (stock_dir / 'futures_parquet').exists() else []
            all_spots = sorted([f.stem for f in (stock_dir / 'spot_parquet').glob('*.parquet')]) if (stock_dir / 'spot_parquet').exists() else []
            print(f"No parquet on {entry_dt} for {sym}. Available futs count: {len(all_futs)}, spots count: {len(all_spots)}")

print(f"\nTotal missing stock folders: {len(missing_stocks)}")
