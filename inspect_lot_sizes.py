import pandas as pd, sys

sys.stdout.reconfigure(errors='replace')

path = r'D:/behaviour analysis/Nifty50_12_Quarters_Futures_OI_Master_v13.xlsx'
xl = pd.ExcelFile(path, engine='openpyxl')
df_trades = xl.parse('All_12Q_Futures_Trades')

# Extract unique symbol lot size mapping
df_symbols = df_trades.drop_duplicates(subset=['Symbol'])[['Symbol', 'Company Name', 'Sector', 'Lot Size (Qty)', 'Entry Futures Price (₹)']].copy()

df_symbols['Contract Value (₹)'] = df_symbols['Entry Futures Price (₹)'] * df_symbols['Lot Size (Qty)']
df_symbols['20% Initial Margin (₹)'] = df_symbols['Contract Value (₹)'] * 0.20

df_symbols = df_symbols.sort_values(by='Symbol').reset_index(drop=True)

print("==========================================================================================================")
print("EXACT NSE LOT SIZES, CONTRACT VALUES & 20% MARGIN REQUIRED FOR ALL 50 NIFTY STOCKS (1 LOT MODEL)")
print("==========================================================================================================")
print(df_symbols.head(15).to_string())
print("\nAverage Lot Size across 50 Stocks:", int(df_symbols['Lot Size (Qty)'].mean()))
print("Average Contract Value per Lot   : ₹", f"{df_symbols['Contract Value (₹)'].mean():,.2f}")
print("Average 20% Margin per Lot       : ₹", f"{df_symbols['20% Initial Margin (₹)'].mean():,.2f}")
print("==========================================================================================================")
