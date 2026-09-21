import openpyxl
import pandas as pd
import sys

sys.stdout.reconfigure(errors='replace')

master_path = r'D:/behaviour analysis/Nifty50_12_Quarters_Options_1PCT_ITM_40PCT_SL_Master.xlsx'
xl_master = pd.ExcelFile(master_path, engine='openpyxl')

df_q3 = xl_master.parse('Q3 2023-24', header=6)

df_q3['Entry Date DT'] = pd.to_datetime(df_q3['Entry Date'])
df_q3['Exit Date DT']  = pd.to_datetime(df_q3['Exit Date'])

peak_date = pd.to_datetime('2023-10-31')

active_on_peak = df_q3[(df_q3['Entry Date DT'] <= peak_date) & (peak_date <= df_q3['Exit Date DT'])].copy().reset_index(drop=True)

print("==========================================================================================")
print("EXACT BREAKUP OF PEAK CONCURRENT CAPITAL (₹4,18,871.44) ON 2023-10-31 IN Q3 2023-24")
print("==========================================================================================")
print(f"Total Open Concurrent Trades on 2023-10-31: {len(active_on_peak)}")
print(f"Combined Capital Invested                     : ₹{active_on_peak['Capital Outlay per Trade (₹)'].sum():,.2f}")

show_cols = [
    'Trade #', 'Symbol', 'Company Name', 'Option Strategy', 'Entry Date', 'Exit Date',
    'Entry Spot Price (₹)', '1% ITM Strike Price (₹)', 'Option Entry Premium (₹)',
    'Lot Size (Qty)', 'Capital Outlay per Trade (₹)'
]

print("\nDetailed Stock-by-Stock Capital Breakup:")
print(active_on_peak[show_cols].to_string())

print("==========================================================================================")
