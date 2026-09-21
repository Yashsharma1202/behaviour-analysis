import pandas as pd, sys

sys.stdout.reconfigure(errors='replace')

path = r'D:/behaviour analysis/Nifty50_12_Quarters_Consolidated.xlsx'
xl = pd.ExcelFile(path, engine='openpyxl')

df_all = xl.parse('All_12_Quarters_Master_Trades')

print("Shape:", df_all.shape)

ret_col = next(c for c in df_all.columns if 'Realised Return' in c)
pnl_col = next(c for c in df_all.columns if 'Realised P' in c or 'Profit' in c or 'P&L' in c)
entry_price_col = next(c for c in df_all.columns if 'Entry Price' in c)
exit_price_col = next(c for c in df_all.columns if 'Exit Price' in c)

df_all[ret_col] = pd.to_numeric(df_all[ret_col], errors='coerce')
df_all[pnl_col] = pd.to_numeric(df_all[pnl_col], errors='coerce')

wins_eq = len(df_all[df_all[ret_col] > 0])
losses_eq = len(df_all[df_all[ret_col] <= 0])
win_rate_eq = wins_eq / len(df_all)

print(f"\nEquity Wins: {wins_eq} ({win_rate_eq:.1%}) | Losses: {losses_eq}")

# Now let's check how Futures prices from OI_DATA map to these exact trades
df_daily = pd.read_excel(r'D:/behaviour analysis/Futures_Master_Only_v2.xlsx', sheet_name='Futures_Master', engine='openpyxl')
df_daily['Date'] = pd.to_datetime(df_daily['Date']).dt.strftime('%Y-%m-%d')
df_daily['Symbol'] = df_daily['Symbol'].astype(str).str.strip()

price_map = {}
for _, r in df_daily.iterrows():
    p = r.iloc[2]
    if pd.notnull(p) and float(p) > 0:
        price_map[(r['Symbol'], str(r['Date']))] = float(p)

diff_count = 0
fut_wins = 0
fut_losses = 0

for idx, r in df_all.iterrows():
    sym = str(r['Symbol']).strip()
    entry_dt = pd.to_datetime(r['Entry Date']).strftime('%Y-%m-%d')
    exit_dt = pd.to_datetime(r['Exit Date']).strftime('%Y-%m-%d')
    strat = str(r['Strategy']).strip().upper()
    direction = 1 if 'LONG' in strat else -1
    
    eq_entry = float(r[entry_price_col])
    eq_exit  = float(r[exit_price_col])
    eq_ret   = float(r[ret_col])
    
    fut_entry = price_map.get((sym, entry_dt), eq_entry)
    fut_exit  = price_map.get((sym, exit_dt), eq_exit)
    
    fut_ret = direction * (fut_exit - fut_entry) / fut_entry if fut_entry > 0 else eq_ret
    
    if fut_ret > 0: fut_wins += 1
    else: fut_losses += 1
    
    # Check if equity win != futures win
    if (eq_ret > 0 and fut_ret <= 0) or (eq_ret <= 0 and fut_ret > 0):
        diff_count += 1
        if diff_count <= 10:
            print(f"Mismatch Trade #{r['Trade No']} {sym} ({strat}): Entry {entry_dt}, Exit {exit_dt} | Eq Ret: {eq_ret:.2%}, Fut Ret: {fut_ret:.2%} | Eq Prices: ({eq_entry:.2f}, {eq_exit:.2f}), Fut Prices: ({fut_entry:.2f}, {fut_exit:.2f})")

print(f"\nFutures Wins: {fut_wins} ({fut_wins/len(df_all):.1%}) | Futures Losses: {fut_losses}")
print(f"Total Mismatches between Equity & Futures Direction: {diff_count}")
