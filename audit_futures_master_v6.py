import pandas as pd, pathlib, sys

sys.stdout.reconfigure(errors='replace')

BASE_DIR = pathlib.Path('D:/behaviour analysis')
OI_DATA_DIR = BASE_DIR / 'OI_DATA'
V6_PATH = BASE_DIR / 'Nifty50_12_Quarters_Futures_OI_Master_v6.xlsx'

print("==========================================", flush=True)
print("Auditing Nifty50_12_Quarters_Futures_OI_Master_v6.xlsx against raw OI_DATA futures parquets...", flush=True)
print("==========================================", flush=True)

df_trades = pd.read_excel(V6_PATH, sheet_name='All_12Q_Futures_Trades', engine='openpyxl')
print(f"Loaded {len(df_trades)} trades from v6 workbook.", flush=True)

# Cache futures parquets
futures_cache = {}

def get_actual_futures_price(symbol, date_str):
    if (symbol, date_str) in futures_cache:
        return futures_cache[(symbol, date_str)]
    
    sym_dir = OI_DATA_DIR / symbol / 'futures_parquet'
    if not sym_dir.exists():
        futures_cache[(symbol, date_str)] = None
        return None
    
    # Try exact date file first
    exact_file = sym_dir / f"{date_str}.parquet"
    if exact_file.exists():
        try:
            df = pd.read_parquet(exact_file)
            if 'close' in df.columns and not df['close'].empty:
                val = float(df['close'].iloc[-1])
                futures_cache[(symbol, date_str)] = val
                return val
        except Exception:
            pass
            
    # Try nearest date file
    all_files = list(sym_dir.glob('*.parquet'))
    if not all_files:
        futures_cache[(symbol, date_str)] = None
        return None
        
    target_dt = pd.to_datetime(date_str)
    file_dts = []
    for f in all_files:
        try:
            f_dt = pd.to_datetime(f.stem)
            file_dts.append((abs((f_dt - target_dt).days), f))
        except:
            pass
            
    file_dts.sort(key=lambda x: x[0])
    if file_dts and file_dts[0][0] <= 5: # Within 5 days
        best_file = file_dts[0][1]
        try:
            df = pd.read_parquet(best_file)
            if 'close' in df.columns and not df['close'].empty:
                val = float(df['close'].iloc[-1])
                futures_cache[(symbol, date_str)] = val
                return val
        except Exception:
            pass

    futures_cache[(symbol, date_str)] = None
    return None

# Audit each trade
exact_matches = 0
diffs = 0
missing_parquet = 0

for idx, r in df_trades.iterrows():
    sym = str(r['Symbol']).strip()
    entry_dt = str(r['Entry Date'])[:10]
    exit_dt = str(r['Exit Date'])[:10]
    
    v6_entry = float(r['Entry Futures Price (₹)'])
    v6_exit = float(r['Exit Futures Price (₹)'])
    
    raw_entry = get_actual_futures_price(sym, entry_dt)
    raw_exit = get_actual_futures_price(sym, exit_dt)
    
    if raw_entry is None or raw_exit is None:
        missing_parquet += 1
    else:
        if abs(raw_entry - v6_entry) < 1.0 and abs(raw_exit - v6_exit) < 1.0:
            exact_matches += 1
        else:
            diffs += 1
            if diffs <= 10:
                print(f"Diff Trade #{r['Trade No']} {sym}: Entry {entry_dt} (v6: {v6_entry}, raw: {raw_entry}), Exit {exit_dt} (v6: {v6_exit}, raw: {raw_exit})", flush=True)

print("\n--- Audit Summary ---", flush=True)
print(f"Exact Matches with Raw Parquet: {exact_matches}", flush=True)
print(f"Differences from Raw Parquet: {diffs}", flush=True)
print(f"Missing Raw Parquet Files: {missing_parquet}", flush=True)
