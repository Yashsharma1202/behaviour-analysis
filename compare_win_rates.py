import pandas as pd, sys

sys.stdout.reconfigure(errors='replace')

fut_path = r'D:/behaviour analysis/Nifty50_12_Quarters_Point_In_Time_Futures_Master.xlsx'
opt_path = r'D:/behaviour analysis/Nifty50_12_Quarters_Options_1PCT_ITM_Detailed_Master.xlsx'

xl_fut = pd.ExcelFile(fut_path, engine='openpyxl')
xl_opt = pd.ExcelFile(opt_path, engine='openpyxl')

df_fut_sum = xl_fut.parse('Exec_12Q_Combined_Summary', header=6)
df_opt_sum = xl_opt.parse('Exec_12Q_Combined_Summary', header=6)

print("==========================================================================================")
print("SIDE-BY-SIDE WIN RATE COMPARISON: FUTURES MODEL VS. OPTIONS BUYING MODEL")
print("==========================================================================================")

quarters_order = [
    'Q3 2023-24', 'Q4 2023-24', 'Q1 2024-25', 'Q2 2024-25',
    'Q3 2024-25', 'Q4 2024-25', 'Q1 2025-26', 'Q2 2025-26',
    'Q3 2025-26', 'Q4 2025-26', 'Q1 2026-27', 'Q2 2026-27'
]

comp_data = []

for qtr in quarters_order:
    q_fut = xl_fut.parse(qtr, header=6)
    q_opt = xl_opt.parse(qtr, header=6)
    
    pnl_col = next(c for c in q_fut.columns if 'P&L' in str(c) or 'PnL' in str(c))
    fut_wins = len(q_fut[q_fut[pnl_col] > 0])
    opt_wins = len(q_opt[q_opt['Booked Option P&L (₹)'] > 0])
    
    fut_wr = fut_wins / len(q_fut)
    opt_wr = opt_wins / len(q_opt)
    
    diff_wr = opt_wr - fut_wr
    
    comp_data.append({
        'Quarter': qtr,
        'Futures Wins': f"{fut_wins}/50",
        'Futures Win Rate (%)': f"{fut_wr:.2%}",
        'Options Wins': f"{opt_wins}/50",
        'Options Win Rate (%)': f"{opt_wr:.2%}",
        'Difference (%)': f"{diff_wr:+.2%}"
    })

df_comp = pd.DataFrame(comp_data)
print(df_comp.to_string(index=False))

print("==========================================================================================")
