import pandas as pd, pathlib, sys

sys.stdout.reconfigure(errors='replace')

BASE_DIR = pathlib.Path('D:/behaviour analysis')
CONSOLIDATED_PATH = BASE_DIR / 'Nifty50_12_Quarters_Consolidated.xlsx'
FUTURES_DAILY = BASE_DIR / 'Futures_Master_Only_v2.xlsx'

# Load daily price lookup
df_daily = pd.read_excel(FUTURES_DAILY, sheet_name='Futures_Master', engine='openpyxl')
df_daily['Date'] = pd.to_datetime(df_daily['Date']).dt.strftime('%Y-%m-%d')
df_daily['Symbol'] = df_daily['Symbol'].astype(str).str.strip()

price_lookup_map = {}
for _, r in df_daily.iterrows():
    p = r.iloc[2]
    if pd.notnull(p) and float(p) > 0:
        price_lookup_map[(r['Symbol'], str(r['Date']))] = float(p)

df_cons = pd.read_excel(CONSOLIDATED_PATH, sheet_name='All_12_Quarters_Master_Trades', engine='openpyxl')

print(f"Loaded {len(df_cons)} trades from Consolidated Master.")

# Test evaluation on Futures
wins = 0
losses = 0
tot_pnl = 0.0

for idx, r in df_cons.iterrows():
    sym = str(r['Symbol']).strip()
    entry_dt = pd.to_datetime(r['Entry Date']).strftime('%Y-%m-%d')
    exit_dt = pd.to_datetime(r['Exit Date']).strftime('%Y-%m-%d')
    strat = str(r['Strategy']).strip().upper()
    direction = 1 if 'LONG' in strat else -1
    
    eq_entry = float(r.get('Entry Price (₹)', 0.0) or r.get('Entry Price (?)', 1000.0))
    eq_exit  = float(r.get('Exit Price (₹)', 0.0) or r.get('Exit Price (?)', 1000.0))
    if eq_entry <= 0: eq_entry = 1000.0
    if eq_exit <= 0: eq_exit = eq_entry
    
    entry_p = price_lookup_map.get((sym, entry_dt), eq_entry)
    exit_p  = price_lookup_map.get((sym, exit_dt), eq_exit)
    if entry_p <= 0: entry_p = eq_entry
    if exit_p <= 0: exit_p = eq_exit
    
    lot_size = 500
    if sym in ['RELIANCE', 'TCS', 'INFY', 'HDFCBANK', 'ICICIBANK', 'AXISBANK']: lot_size = 550
    elif sym in ['MARUTI', 'ULTRACEMCO', 'BAJAJ-AUTO']: lot_size = 100
    elif sym in ['COALINDIA', 'ITC', 'BEL']: lot_size = 1500
    
    pnl = direction * (exit_p - entry_p) * lot_size
    tot_pnl += pnl
    if pnl > 0: wins += 1
    else: losses += 1

print(f"Total Wins: {wins} ({wins/len(df_cons):.2%}) | Total Losses: {losses}")
print(f"Total Net Futures PnL: ₹{tot_pnl:,.2f}")
