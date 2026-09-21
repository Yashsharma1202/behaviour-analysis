import pandas as pd, pathlib, sys

sys.stdout.reconfigure(errors='replace')

BASE_DIR = pathlib.Path('D:/behaviour analysis')
V7_PATH = BASE_DIR / 'Nifty50_12_Quarters_Futures_OI_Master_v7.xlsx'

xl = pd.ExcelFile(V7_PATH, engine='openpyxl')
df_trades = xl.parse('All_12Q_Futures_Trades')
df_lead   = xl.parse('Stock_1Lot_Leaderboard')

print("==========================================")
print("AUDITING LEADERBOARD & SHEET CONSISTENCY")
print("==========================================")

pnl_col = next(c for c in df_trades.columns if 'P&L' in c or 'Profit' in c)
sym_col = next(c for c in df_trades.columns if 'Symbol' in c)

# Check Leaderboard PnL sums vs All Trades PnL sums per stock
lead_mismatches = 0
for idx, r in df_lead.iterrows():
    sym = str(r['Symbol']).strip()
    lead_pnl = round(float(r['Total 12Q Futures P&L (₹)']), 2)
    
    trades_stock = df_trades[df_trades[sym_col] == sym]
    sum_stock_pnl = round(trades_stock[pnl_col].astype(float).sum(), 2)
    
    if abs(lead_pnl - sum_stock_pnl) > 1.0:
        lead_mismatches += 1
        print(f"Leaderboard PnL Mismatch for {sym}: Leaderboard={lead_pnl}, Trades Sum={sum_stock_pnl}")

print(f"Leaderboard Stock PnL Audit Mismatches: {lead_mismatches}")
