import pandas as pd, pathlib, sys

sys.stdout.reconfigure(errors='replace')

BASE_DIR = pathlib.Path('D:/behaviour analysis')
OI_ROOT = BASE_DIR / 'OI_DATA'
REGULARIZED_PATH = BASE_DIR / 'Nifty50_12_Quarters_Regularized_Master.xlsx'
FUTURES_DAILY = BASE_DIR / 'Futures_Master_Only_v2.xlsx'

# Load daily price lookup
df_daily = pd.read_excel(FUTURES_DAILY, sheet_name='Futures_Master', engine='openpyxl')
df_daily['Date'] = pd.to_datetime(df_daily['Date']).dt.strftime('%Y-%m-%d')
df_daily['Symbol'] = df_daily['Symbol'].astype(str).str.strip()

price_lookup_map = {}
for _, r in df_daily.iterrows():
    p = r.iloc[2]
    if pd.notnull(p) and p > 0:
        price_lookup_map[(r['Symbol'], str(r['Date']))] = float(p)

df_reg = pd.read_excel(REGULARIZED_PATH, sheet_name='All_12_Quarters_Master_Trades', engine='openpyxl')

wins = 0
losses = 0
total_pnl = 0.0

qtr_stats = {}

for idx, row in df_reg.iterrows():
    sym = str(row['Symbol']).strip()
    qtr = str(row['Quarter']).strip()
    entry_dt = pd.to_datetime(row['Entry Date']).strftime('%Y-%m-%d')
    exit_dt = pd.to_datetime(row['Exit Date']).strftime('%Y-%m-%d')
    strat = str(row['Strategy']).strip().upper() # LONG or SHORT
    direction = 1 if 'LONG' in strat else -1
    
    # Fallback to equity entry/exit price if futures price missing for missing stock
    eq_entry = row.get('Entry Price (₹)', 1000.0) or 1000.0
    eq_exit  = row.get('Exit Price (₹)', 1000.0) or 1000.0
    
    entry_price = price_lookup_map.get((sym, entry_dt), eq_entry)
    exit_price  = price_lookup_map.get((sym, exit_dt), eq_exit)
    
    if entry_price <= 0: entry_price = eq_entry
    if exit_price <= 0: exit_price = eq_exit
    
    # Estimate lot size (standard lot size lookup or 500)
    lot_size = 500
    if sym in ['RELIANCE', 'TCS', 'INFY', 'HDFCBANK', 'ICICIBANK', 'AXISBANK']: lot_size = 550
    elif sym in ['MARUTI', 'ULTRACEMCO', 'BAJAJ-AUTO']: lot_size = 100
    elif sym in ['COALINDIA', 'ITC', 'BEL']: lot_size = 1500
    
    pnl = direction * (exit_price - entry_price) * lot_size
    
    if qtr not in qtr_stats:
        qtr_stats[qtr] = {'trades': 0, 'wins': 0, 'losses': 0, 'pnl': 0.0}
    
    qtr_stats[qtr]['trades'] += 1
    qtr_stats[qtr]['pnl'] += pnl
    if pnl > 0:
        qtr_stats[qtr]['wins'] += 1
        wins += 1
    else:
        qtr_stats[qtr]['losses'] += 1
        losses += 1
    total_pnl += pnl

print("==========================================")
print(f"Regularized Model evaluated on FUTURES DATA:")
print(f"Total Trades: {len(df_reg)}")
print(f"Total Wins  : {wins} ({wins/len(df_reg):.2%})")
print(f"Total Losses: {losses} ({losses/len(df_reg):.2%})")
print(f"Total Net PnL (₹): ₹{total_pnl:,.2f}")
print("==========================================")
for q, s in qtr_stats.items():
    wr = s['wins'] / s['trades'] if s['trades'] > 0 else 0.0
    print(f"Quarter: {q:<12} | Trades: {s['trades']:2d} | Wins: {s['wins']:2d} | Win Rate: {wr:6.2%} | Net PnL: ₹{s['pnl']:>12,.2f}")
