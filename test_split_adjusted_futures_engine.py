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

ret_col = next(c for c in df_cons.columns if 'Realised Return' in c)
entry_price_col = next(c for c in df_cons.columns if 'Entry Price' in c)
exit_price_col = next(c for c in df_cons.columns if 'Exit Price' in c)

df_cons[ret_col] = pd.to_numeric(df_cons[ret_col], errors='coerce')

# Lot Size Map
lot_size_map = {
    'RELIANCE': 250, 'TCS': 175, 'INFY': 400, 'HDFCBANK': 550, 'ICICIBANK': 700,
    'AXISBANK': 625, 'SBIN': 1500, 'BHARTIARTL': 950, 'ITC': 1600, 'KOTAKBANK': 400,
    'LTIM': 150, 'LT': 300, 'HINDUNILVR': 300, 'BAJFINANCE': 125, 'MARUTI': 100,
    'ASIANPAINT': 200, 'HCLTECH': 350, 'TITAN': 175, 'SUNPHARMA': 350, 'TATAMOTORS': 1425,
    'ULTRACEMCO': 100, 'POWERGRID': 1800, 'NTPC': 1500, 'COALINDIA': 2100, 'TATASTEEL': 5500,
    'JIOFIN': 2000, 'JSWSTEEL': 675, 'M&M': 350, 'ADANIENT': 300, 'ADANIPORTS': 800,
    'GRASIM': 475, 'BAJAJFINSV': 500, 'BAJAJ-AUTO': 125, 'NESTLEIND': 250, 'APOLLOHOSP': 125,
    'WIPRO': 1500, 'EICHERMOT': 175, 'DIVISLAB': 150, 'DRREDDY': 125, 'CIPLA': 650,
    'BPCL': 1800, 'TATACONSUM': 900, 'BRITANNIA': 200, 'HEROMOTOCO': 150, 'INDUSINDBK': 500,
    'HDFCLIFE': 1100, 'SBILIFE': 750, 'BEL': 1500, 'SHRIRAMFIN': 300, 'TECHM': 600
}

def get_lot(sym):
    return lot_size_map.get(sym, 500)

wins = 0
losses = 0
tot_pnl = 0.0

for idx, r in df_cons.iterrows():
    sym = str(r['Symbol']).strip()
    entry_dt = pd.to_datetime(r['Entry Date']).strftime('%Y-%m-%d')
    exit_dt = pd.to_datetime(r['Exit Date']).strftime('%Y-%m-%d')
    strat = str(r['Strategy']).strip().upper()
    direction = 1 if 'LONG' in strat else -1
    
    eq_entry = float(r[entry_price_col])
    eq_exit  = float(r[exit_price_col])
    ret_pct  = float(r[ret_col])
    
    # Split-adjusted Futures Price calculation matching actual market contract behavior:
    fut_entry = price_lookup_map.get((sym, entry_dt), eq_entry)
    # Check for unadjusted split in parquet lookup: if fut_entry is > 1.8x or < 0.5x eq_entry, adjust to eq_entry scale!
    if fut_entry > 1.8 * eq_entry or fut_entry < 0.5 * eq_entry:
        fut_entry = eq_entry
        
    fut_exit = fut_entry * (1.0 + direction * ret_pct)
    
    lot_size = get_lot(sym)
    pnl = direction * (fut_exit - fut_entry) * lot_size
    
    tot_pnl += pnl
    if ret_pct > 0:
        wins += 1
    else:
        losses += 1

print("==========================================")
print(f"Split-Adjusted Consolidated Futures Engine Results:")
print(f"Total Trades: {len(df_cons)}")
print(f"Wins: {wins} ({wins/len(df_cons):.1%}) | Losses: {losses}")
print(f"Total 12Q Futures Net P&L: ₹{tot_pnl:,.2f}")
print("==========================================")
