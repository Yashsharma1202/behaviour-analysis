import pandas as pd, pathlib, sys, time

sys.stdout.reconfigure(errors='replace')

t0 = time.time()
BASE_DIR = pathlib.Path('D:/behaviour analysis')
OI_DATA_DIR = BASE_DIR / 'OI_DATA'
V6_PATH = BASE_DIR / 'Nifty50_12_Quarters_Futures_OI_Master_v6.xlsx'

print("==========================================", flush=True)
print("Auditing Nifty50_12_Quarters_Futures_OI_Master_v6.xlsx against raw OI_DATA futures parquets...", flush=True)
print("==========================================", flush=True)

df_trades = pd.read_excel(V6_PATH, sheet_name='All_12Q_Futures_Trades', engine='openpyxl')
print(f"Loaded {len(df_trades)} trades from v6 workbook.", flush=True)

# Pre-scan directory files
dir_cache = {}
for sym_dir in OI_DATA_DIR.glob('*/futures_parquet'):
    sym = sym_dir.parent.name
    files = {f.stem: f for f in sym_dir.glob('*.parquet')}
    dir_cache[sym] = files

print(f"Pre-scanned futures parquets for {len(dir_cache)} symbols in {time.time() - t0:.2f}s", flush=True)

futures_cache = {}

def get_actual_futures_price(symbol, date_str):
    if (symbol, date_str) in futures_cache:
        return futures_cache[(symbol, date_str)]
    
    files = dir_cache.get(symbol, {})
    if not files:
        futures_cache[(symbol, date_str)] = None
        return None
    
    if date_str in files:
        try:
            df = pd.read_parquet(files[date_str])
            if 'close' in df.columns and not df['close'].empty:
                val = float(df['close'].iloc[-1])
                futures_cache[(symbol, date_str)] = val
                return val
        except Exception:
            pass
            
    # Nearest date
    target_dt = pd.to_datetime(date_str)
    file_dts = []
    for d_str, f_path in files.items():
        try:
            f_dt = pd.to_datetime(d_str)
            file_dts.append((abs((f_dt - target_dt).days), f_path))
        except:
            pass
            
    file_dts.sort(key=lambda x: x[0])
    if file_dts and file_dts[0][0] <= 5:
        try:
            df = pd.read_parquet(file_dts[0][1])
            if 'close' in df.columns and not df['close'].empty:
                val = float(df['close'].iloc[-1])
                futures_cache[(symbol, date_str)] = val
                return val
        except Exception:
            pass

    futures_cache[(symbol, date_str)] = None
    return None

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
                print(f"Diff Trade #{r['Trade No']} {sym}: Entry {entry_dt} (v6: {v6_entry:.2f}, raw: {raw_entry:.2f}), Exit {exit_dt} (v6: {v6_exit:.2f}, raw: {raw_exit:.2f})", flush=True)

print("\n--- Audit Summary ---", flush=True)
print(f"Exact Matches with Raw Parquet: {exact_matches}", flush=True)
print(f"Differences from Raw Parquet: {diffs}", flush=True)
print(f"Missing Raw Parquet Files: {missing_parquet}", flush=True)
print(f"Audit completed in {time.time() - t0:.2f}s", flush=True)
