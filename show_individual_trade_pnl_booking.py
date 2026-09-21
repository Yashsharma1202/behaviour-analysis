import openpyxl, pandas as pd, sys

sys.stdout.reconfigure(errors='replace')

master_path = r'D:/behaviour analysis/Nifty50_12_Quarters_Options_1PCT_ITM_Point_In_Time_Master.xlsx'
xl_master = pd.ExcelFile(master_path, engine='openpyxl')

df_q3 = xl_master.parse('Q3 2023-24', header=6)

print("==========================================================================================")
print("INDIVIDUAL TRADE-BY-TRADE P&L BOOKING SAMPLE (Q3 2023-24)")
print("==========================================================================================")
sample_cols = [
    'Trade #', 'Symbol', 'Option Strategy', 'Entry Date', 'Exit Date',
    'Entry Spot Price (₹)', '1% ITM Strike Price (₹)', 'Option Entry Premium (₹)',
    'Option Exit Premium (₹)', 'Lot Size (Qty)', 'Capital Outlay per Trade (₹)', 'Booked Option P&L (₹)'
]

df_sample = df_q3[sample_cols].head(12)
print(df_sample.to_string())
print("==========================================================================================")
