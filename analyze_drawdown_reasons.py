import openpyxl
import pandas as pd
import numpy as np
import pathlib
import sys

sys.stdout.reconfigure(errors='replace')

master_path = r'D:/behaviour analysis/Nifty50_12_Quarters_Options_1PCT_ITM_Detailed_Master.xlsx'
xl_master = pd.ExcelFile(master_path, engine='openpyxl')

quarters_to_analyze = ['Q2 2025-26', 'Q1 2026-27', 'Q4 2024-25', 'Q3 2025-26']

print("==========================================================================================")
print("EMPIRICAL ROOT CAUSE ANALYSIS OF STRATEGY DRAWDOWNS (DD)")
print("==========================================================================================")

for qtr in quarters_to_analyze:
    df_q = xl_master.parse(qtr, header=6)
    
    df_q['Cum PnL'] = df_q['Booked Option P&L (₹)'].cumsum()
    df_q['Running Max'] = np.maximum.accumulate(df_q['Cum PnL'])
    df_q['Drawdown'] = df_q['Cum PnL'] - df_q['Running Max']
    
    trough_idx = df_q['Drawdown'].idxmin()
    max_dd_val = abs(df_q.loc[trough_idx, 'Drawdown'])
    
    # Find peak index before trough
    peak_idx = df_q.loc[:trough_idx, 'Running Max'].idxmax()
    
    dd_trades = df_q.loc[peak_idx:trough_idx].copy()
    
    losing_trades_in_dd = dd_trades[dd_trades['Booked Option P&L (₹)'] < 0]
    total_dd_losses = losing_trades_in_dd['Booked Option P&L (₹)'].sum()
    
    top_loss_stocks = losing_trades_in_dd.sort_values(by='Booked Option P&L (₹)').head(5)
    
    print(f"\n📌 --- {qtr.upper()} DRAWDOWN ANALYSIS ---")
    print(f"Max Rupee Drawdown  : -₹{max_dd_val:,.2f}")
    print(f"Drawdown Window     : Trade #{peak_idx + 1} ({df_q.loc[peak_idx, 'Symbol']}) to Trade #{trough_idx + 1} ({df_q.loc[trough_idx, 'Symbol']})")
    print(f"Trades in DD Window : {len(dd_trades)} Trades ({len(losing_trades_in_dd)} Losses, {len(dd_trades) - len(losing_trades_in_dd)} Wins)")
    print(f"Win Rate in DD Window: {(len(dd_trades) - len(losing_trades_in_dd))/len(dd_trades):.1%}")
    print(f"Total Losses in DD  : -₹{abs(total_dd_losses):,.2f}")
    
    print("\nTop 5 Loss Contributor Stocks in Drawdown Window:")
    cols_show = ['Trade #', 'Symbol', 'Option Strategy', 'Entry Spot Price (₹)', 'Exit Spot Price (₹)', 'Option Entry Premium (₹)', 'Option Exit Premium (₹)', 'Booked Option P&L (₹)']
    print(top_loss_stocks[cols_show].to_string())
    print("-" * 90)

print("==========================================================================================")
