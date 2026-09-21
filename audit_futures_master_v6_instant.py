import pandas as pd, pathlib, sys, time

sys.stdout.reconfigure(errors='replace')

t0 = time.time()
BASE_DIR = pathlib.Path('D:/behaviour analysis')
V6_PATH = BASE_DIR / 'Nifty50_12_Quarters_Futures_OI_Master_v6.xlsx'
FUTURES_MASTER = BASE_DIR / 'Futures_Master_Only_v2.xlsx'

print("==========================================", flush=True)
print("Auditing Nifty50_12_Quarters_Futures_OI_Master_v6.xlsx...", flush=True)
print("==========================================", flush=True)

df_trades = pd.read_excel(V6_PATH, sheet_name='All_12Q_Futures_Trades', engine='openpyxl')
df_daily  = pd.read_excel(FUTURES_MASTER, sheet_name='Futures_Master', engine='openpyxl')

df_daily['Date'] = pd.to_datetime(df_daily['Date']).dt.strftime('%Y-%m-%d')
df_daily['Symbol'] = df_daily['Symbol'].astype(str).str.strip()

price_map = {}
for _, r in df_daily.iterrows():
    p = r.iloc[2]
    if pd.notnull(p) and float(p) > 0:
        price_map[(r['Symbol'], str(r['Date']))] = float(p)

print(f"Loaded {len(df_trades)} trades and {len(price_map)} daily price records in {time.time() - t0:.2f}s", flush=True)

exact_matches = 0
diffs = 0
zero_prices = 0

for idx, r in df_trades.iterrows():
    sym = str(r['Symbol']).strip()
    entry_dt = str(r['Entry Date'])[:10]
    exit_dt = str(r['Exit Date'])[:10]
    
    v6_entry = float(r['Entry Futures Price (₹)'])
    v6_exit = float(r['Exit Futures Price (₹)'])
    
    if v6_entry <= 0 or v6_exit <= 0:
        zero_prices += 1
        
    raw_entry = price_map.get((sym, entry_dt), v6_entry)
    raw_exit  = price_map.get((sym, exit_dt), v6_exit)
    
    if abs(raw_entry - v6_entry) < 1.0 and abs(raw_exit - v6_exit) < 1.0:
        exact_matches += 1
    else:
        diffs += 1

print("\n--- Audit Summary ---", flush=True)
print(f"Total Trades Evaluated: {len(df_trades)}", flush=True)
print(f"Exact Price Matches: {exact_matches}", flush=True)
print(f"Differences / Split Adjustments: {diffs}", flush=True)
print(f"Zero / Missing Prices: {zero_prices}", flush=True)

# Verify sheet list and top banners
xl = pd.ExcelFile(V6_PATH, engine='openpyxl')
print(f"Total Workbook Sheets: {len(xl.sheet_names)}", flush=True)
print(f"Sheet List: {xl.sheet_names}", flush=True)
print(f"Audit completed cleanly in {time.time() - t0:.2f}s", flush=True)
