import pandas as pd, sys

sys.stdout.reconfigure(errors='replace')

SOURCE_FUT = r'D:/behaviour analysis/Nifty50_12_Quarters_Futures_OI_Master_v13.xlsx'
df = pd.read_excel(SOURCE_FUT, sheet_name='All_12Q_Futures_Trades', engine='openpyxl')

unique_lots = df[['Symbol', 'Company Name', 'Sector', 'Lot Size (Qty)']].drop_duplicates(subset=['Symbol']).reset_index(drop=True)

print("==========================================================================================")
print("OFFICIAL NSE DERIVATIVE LOT SIZES PULLED FROM MASTER WORKBOOK")
print("==========================================================================================")
print(unique_lots.head(25).to_string())
print("==========================================================================================")
