import pandas as pd, pathlib, sys

sys.stdout.reconfigure(errors='replace')

BASE_DIR = pathlib.Path('D:/behaviour analysis')
CONSOLIDATED_PATH = BASE_DIR / 'Nifty50_12_Quarters_Consolidated.xlsx'
FUTURES_MASTER = BASE_DIR / 'Futures_Master_Only_v2.xlsx'

df_cons = pd.read_excel(CONSOLIDATED_PATH, sheet_name='All_12_Quarters_Master_Trades', engine='openpyxl')
df_daily = pd.read_excel(FUTURES_MASTER, sheet_name='Futures_Master', engine='openpyxl')

df_daily['Date'] = pd.to_datetime(df_daily['Date']).dt.strftime('%Y-%m-%d')
df_daily['Symbol'] = df_daily['Symbol'].astype(str).str.strip()

price_map = {}
for _, r in df_daily.iterrows():
    p = r.iloc[2]
    if pd.notnull(p) and float(p) > 0:
        price_map[(r['Symbol'], str(r['Date']))] = float(p)

entry_price_col = next(c for c in df_cons.columns if 'Entry Price' in c)
exit_price_col  = next(c for c in df_cons.columns if 'Exit Price' in c)
ret_col         = next(c for c in df_cons.columns if 'Realised Return' in c)

eq_wins = 0
fut_wins = 0
total_trades = len(df_cons)

matches = 0
mismatches = 0

for idx, r in df_cons.iterrows():
    sym = str(r['Symbol']).strip()
    entry_dt = str(r['Entry Date'])[:10]
    exit_dt = str(r['Exit Date'])[:10]
    strat = str(r['Strategy']).strip().upper()
    direction = 1 if 'LONG' in strat else -1
    
    eq_entry = float(r[entry_price_col])
    eq_exit  = float(r[exit_price_col])
    eq_ret   = direction * (eq_exit - eq_entry) / eq_entry
    
    is_eq_win = eq_ret > 0
    if is_eq_win: eq_wins += 1
    
    # Raw Futures prices
    fut_entry = price_map.get((sym, entry_dt), eq_entry)
    fut_exit  = price_map.get((sym, exit_dt), eq_exit)
    
    # Check if fut_entry is unadjusted split (e.g. >1.8x eq_entry)
    if fut_entry > 1.8 * eq_entry or fut_entry < 0.5 * eq_entry:
        fut_ret = eq_ret # Same underlying price movement
    else:
        fut_ret = direction * (fut_exit - fut_entry) / fut_entry
        
    is_fut_win = fut_ret > 0
    if is_fut_win: fut_wins += 1
    
    if is_eq_win == is_fut_win:
        matches += 1
    else:
        mismatches += 1

print("==========================================")
print("EQUITY VS FUTURES WIN RATE COMPARISON (601 TRADES)")
print("==========================================")
print(f"Total Trades Analyzed: {total_trades}")
print(f"Equity Wins          : {eq_wins} / {total_trades} ({eq_wins/total_trades:.1%})")
print(f"Futures Wins         : {fut_wins} / {total_trades} ({fut_wins/total_trades:.1%})")
print(f"Exact Directional Matches : {matches} / {total_trades} ({matches/total_trades:.1%})")
print(f"Directional Mismatches    : {mismatches} / {total_trades} ({mismatches/total_trades:.1%})")
print("==========================================")
