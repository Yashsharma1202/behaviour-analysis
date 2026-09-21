import openpyxl
import pandas as pd
import numpy as np
import pathlib
import sys

sys.stdout.reconfigure(errors='replace')

master_path = r'D:/behaviour analysis/Nifty50_12_Quarters_Options_New_Strike_No_SL_Master.xlsx'
xl_master = pd.ExcelFile(master_path, engine='openpyxl')

quarters_order = [
    'Q3 2023-24', 'Q4 2023-24', 'Q1 2024-25', 'Q2 2024-25',
    'Q3 2024-25', 'Q4 2024-25', 'Q1 2025-26', 'Q2 2025-26',
    'Q3 2025-26', 'Q4 2025-26', 'Q1 2026-27', 'Q2 2026-27'
]

all_trades = []
for qtr in quarters_order:
    df_q = xl_master.parse(qtr, header=6)
    df_q['Quarter'] = qtr
    all_trades.append(df_q)

df_all = pd.concat(all_trades, ignore_index=True)

# 1. Baseline Win Rate
baseline_wins = len(df_all[df_all['Booked Option P&L (₹)'] > 0])
baseline_win_rate = baseline_wins / len(df_all)

# 2. Target Lock +30%
df_all['Target_30_PnL'] = df_all.apply(
    lambda r: r['Option Entry Premium (₹)'] * 0.30 * r['Lot Size (Qty)']
    if (r['Option Exit Premium (₹)'] - r['Option Entry Premium (₹)']) / r['Option Entry Premium (₹)'] >= 0.30
    else r['Booked Option P&L (₹)'], axis=1
)
tp30_wins = len(df_all[df_all['Target_30_PnL'] > 0])
tp30_win_rate = tp30_wins / len(df_all)

# 3. Target Lock +20%
df_all['Target_20_PnL'] = df_all.apply(
    lambda r: r['Option Entry Premium (₹)'] * 0.20 * r['Lot Size (Qty)']
    if (r['Option Exit Premium (₹)'] - r['Option Entry Premium (₹)']) / r['Option Entry Premium (₹)'] >= 0.20
    else r['Booked Option P&L (₹)'], axis=1
)
tp20_wins = len(df_all[df_all['Target_20_PnL'] > 0])
tp20_win_rate = tp20_wins / len(df_all)

# 4. Trend Filter Simulation (Filtering out counter-trend trades)
# Suppose we filter out trades where Option Strategy is counter to quarterly sector trend
# Let's filter high-conviction trades with >1% spot move
high_conv_trades = df_all[abs((df_all['Exit Spot Price (₹)'] - df_all['Entry Spot Price (₹)']) / df_all['Entry Spot Price (₹)']) >= 0.015]
hc_wins = len(high_conv_trades[high_conv_trades['Booked Option P&L (₹)'] > 0])
hc_win_rate = hc_wins / len(high_conv_trades)

print("========================================================================================================================")
print("QUANTITATIVE WIN RATE BOOSTER SIMULATION Across 600 TRADES")
print("========================================================================================================================")
print(f"1. Baseline Win Rate (All 600 Trades)         : {baseline_win_rate:.2%} ({baseline_wins}/600 Wins)")
print(f"2. Profit Target Lock at +30% Gain             : {tp30_win_rate:.2%} ({tp30_wins}/600 Wins)")
print(f"3. Profit Target Lock at +20% Gain             : {tp20_win_rate:.2%} ({tp20_wins}/600 Wins)")
print(f"4. High-Conviction Trend Move Filter (>=1.5%) : {hc_win_rate:.2%} ({hc_wins}/{len(high_conv_trades)} Wins)")
print("========================================================================================================================")
