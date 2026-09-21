import openpyxl
import pandas as pd
import numpy as np
import pathlib
import sys

sys.stdout.reconfigure(errors='replace')

master_path = r'D:/behaviour analysis/Nifty50_12_Quarters_Options_1PCT_ITM_Detailed_Master.xlsx'
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

# Define Stop Loss & Profit Target Scenarios
def simulate_scenario(stop_loss_pct=None, target_pct=None):
    sim_summaries = []
    
    for qtr in quarters_order:
        q_tr = df_all[df_all['Quarter'] == qtr].copy().reset_index(drop=True)
        
        sim_pnls = []
        sim_outlays = []
        
        for idx, row in q_tr.iterrows():
            p_entry  = float(row['Option Entry Premium (₹)'])
            p_exit   = float(row['Option Exit Premium (₹)'])
            lot_size = int(row['Lot Size (Qty)'])
            outlay   = p_entry * lot_size
            
            raw_ret  = (p_exit - p_entry) / p_entry if p_entry > 0 else 0.0
            
            # Apply Stop Loss
            if stop_loss_pct is not None and raw_ret < -stop_loss_pct:
                effective_ret = -stop_loss_pct
            # Apply Target
            elif target_pct is not None and raw_ret > target_pct:
                effective_ret = target_pct
            else:
                effective_ret = raw_ret
                
            sim_pnl = round(effective_ret * outlay, 2)
            sim_pnls.append(sim_pnl)
            sim_outlays.append(outlay)
            
        q_tr['Sim_PnL'] = sim_pnls
        q_tr['Opt_Outlay'] = sim_outlays
        
        tot_pnl = round(sum(sim_pnls), 2)
        wins = len([p for p in sim_pnls if p > 0])
        win_rate = wins / len(sim_pnls)
        
        # Peak Concurrent Capital
        q_tr['Entry Date DT'] = pd.to_datetime(q_tr['Entry Date'])
        q_tr['Exit Date DT']  = pd.to_datetime(q_tr['Exit Date'])
        min_date = q_tr['Entry Date DT'].min()
        max_date = q_tr['Exit Date DT'].max()
        date_range = pd.date_range(min_date, max_date)
        
        peak_capital = 0.0
        for d in date_range:
            active_trades = q_tr[(q_tr['Entry Date DT'] <= d) & (d <= q_tr['Exit Date DT'])]
            opt_tot = active_trades['Opt_Outlay'].sum()
            if opt_tot > peak_capital:
                peak_capital = opt_tot
                
        # Drawdown math
        q_tr['Cum PnL'] = q_tr['Sim_PnL'].cumsum()
        q_tr['Running Max'] = np.maximum.accumulate(q_tr['Cum PnL'])
        q_tr['Drawdown'] = q_tr['Cum PnL'] - q_tr['Running Max']
        
        trough_idx = q_tr['Drawdown'].idxmin()
        max_dd_val = abs(q_tr.loc[trough_idx, 'Drawdown'])
        peak_cum_pnl = q_tr.loc[:trough_idx, 'Cum PnL'].max()
        
        dd_pct_init = -abs(max_dd_val / peak_capital) if peak_capital > 0 else 0.0
        dd_pct_peak = -abs(max_dd_val / (peak_capital + max(0.0, peak_cum_pnl))) if peak_capital > 0 else 0.0
        
        sim_summaries.append({
            'Quarter': qtr,
            'Net P&L (₹)': tot_pnl,
            'Peak Capital (₹)': round(peak_capital, 2),
            'Max DD (₹)': round(max_dd_val, 2),
            'DD % Init': dd_pct_init,
            'DD % Peak': dd_pct_peak,
            'Win Rate': win_rate
        })
        
    df_res = pd.DataFrame(sim_summaries)
    return {
        'Total PnL': df_res['Net P&L (₹)'].sum(),
        'Avg Capital': df_res['Peak Capital (₹)'].mean(),
        'Avg Max DD (₹)': df_res['Max DD (₹)'].mean(),
        'Worst Max DD (₹)': df_res['Max DD (₹)'].max(),
        'Avg DD % Init': df_res['DD % Init'].mean(),
        'Worst DD % Init': df_res['DD % Init'].min(),
        'Avg DD % Peak': df_res['DD % Peak'].mean(),
        'Worst DD % Peak': df_res['DD % Peak'].min(),
        'Global Win Rate': df_res['Win Rate'].mean()
    }

print("========================================================================================================================")
print("TESTING DRAWDOWN REDUCTION STRATEGIES (STOP LOSS & PROFIT TARGET RULES)")
print("========================================================================================================================")

baseline = simulate_scenario(stop_loss_pct=None, target_pct=None)
sl40     = simulate_scenario(stop_loss_pct=0.40, target_pct=None)
sl50     = simulate_scenario(stop_loss_pct=0.50, target_pct=None)
sl40_tp75 = simulate_scenario(stop_loss_pct=0.40, target_pct=0.75)

scenarios = [
    ('Baseline (No Stop Loss)', baseline),
    ('Strict -40% Premium Stop Loss', sl40),
    ('Moderate -50% Premium Stop Loss', sl50),
    ('-40% SL + +75% Target Lock', sl40_tp75)
]

for name, sc in scenarios:
    print(f"\n🔹 SCENARIO: {name.upper()}")
    print(f"  12-Quarter Total Net Profit : ₹{sc['Total PnL']:,.2f}")
    print(f"  Avg Quarterly Peak Capital : ₹{sc['Avg Capital']:,.2f}")
    print(f"  Average Max Rupee Drawdown  : -₹{sc['Avg Max DD (₹)']:,.2f}")
    print(f"  Worst Single Quarter Rupee DD: -₹{sc['Worst Max DD (₹)']:,.2f}")
    print(f"  Average DD % on Initial Cap : {sc['Avg DD % Init']:.2%}")
    print(f"  Worst DD % on Initial Cap   : {sc['Worst DD % Init']:.2%}")
    print(f"  Average DD % on Peak Port   : {sc['Avg DD % Peak']:.2%}")
    print(f"  Worst DD % on Peak Port     : {sc['Worst DD % Peak']:.2%}")

print("========================================================================================================================")
