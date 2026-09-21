import openpyxl
import pandas as pd
import numpy as np
import pathlib
import sys
import time

sys.stdout.reconfigure(errors='replace')

t0 = time.time()

# ------------------- Configuration & Paths -------------------
BASE_DIR = pathlib.Path('D:/behaviour analysis')
SOURCE_OPT_MASTER = BASE_DIR / 'Nifty50_12_Quarters_Options_New_Strike_No_SL_Master.xlsx'
OUTPUT_OPT_EXCEL  = BASE_DIR / 'Stop_Loss_Optimization_Report.xlsx'

quarters_order = [
    'Q3 2023-24', 'Q4 2023-24', 'Q1 2024-25', 'Q2 2024-25',
    'Q3 2024-25', 'Q4 2024-25', 'Q1 2025-26', 'Q2 2025-26',
    'Q3 2025-26', 'Q4 2025-26', 'Q1 2026-27', 'Q2 2026-27'
]

print("Loading Trade Data for Fast Stop Loss Optimization Engine...", flush=True)
xl_opt = pd.ExcelFile(SOURCE_OPT_MASTER, engine='openpyxl')

quarter_data = {}
for qtr in quarters_order:
    df_q = xl_opt.parse(qtr, header=6)
    df_q['Entry Date DT'] = pd.to_datetime(df_q['Entry Date'])
    df_q['Exit Date DT']  = pd.to_datetime(df_q['Exit Date'])
    t_no_col = next((c for c in df_q.columns if 'Trade' in str(c)), df_q.columns[0])
    df_q = df_q.sort_values(by=['Entry Date DT', t_no_col]).reset_index(drop=True)
    
    # Pre-calculate Peak Concurrent Capital per Quarter
    min_date = df_q['Entry Date DT'].min()
    max_date = df_q['Exit Date DT'].max()
    date_range = pd.date_range(min_date, max_date)
    
    peak_cap = 0.0
    for d in date_range:
        active_trades = df_q[(df_q['Entry Date DT'] <= d) & (d <= df_q['Exit Date DT'])]
        tot_outlay = active_trades['Capital Outlay per Trade (₹)'].sum()
        if tot_outlay > peak_cap:
            peak_cap = tot_outlay
            
    quarter_data[qtr] = {
        'df': df_q,
        'peak_capital': peak_cap
    }

def evaluate_sl_level(sl_pct):
    q_results = []
    
    for qtr in quarters_order:
        q_info = quarter_data[qtr]
        q_tr   = q_info['df']
        peak_capital = q_info['peak_capital']
        
        p_entries = q_tr['Option Entry Premium (₹)'].values
        p_exits   = q_tr['Option Exit Premium (₹)'].values
        lot_sizes = q_tr['Lot Size (Qty)'].values
        outlays   = q_tr['Capital Outlay per Trade (₹)'].values
        
        raw_rets = (p_exits - p_entries) / p_entries
        raw_rets = np.nan_to_num(raw_rets, nan=0.0)
        
        if sl_pct is not None:
            eff_rets = np.where(raw_rets < -sl_pct, -sl_pct, raw_rets)
            sl_hits  = np.sum(raw_rets < -sl_pct)
        else:
            eff_rets = raw_rets
            sl_hits  = 0
            
        trade_pnls = np.round(eff_rets * outlays, 2)
        tot_pnl    = np.round(np.sum(trade_pnls), 2)
        wins       = np.sum(trade_pnls > 0)
        win_rate   = wins / len(trade_pnls) if len(trade_pnls) > 0 else 0.0
        
        # Fast Drawdown Calculation
        cum_pnl = np.cumsum(trade_pnls)
        running_max = np.maximum.accumulate(cum_pnl)
        drawdown = cum_pnl - running_max
        
        max_dd_val = abs(np.min(drawdown))
        dd_pct_init = -abs(max_dd_val / peak_capital) if peak_capital > 0 else 0.0
        
        q_results.append({
            'Quarter': qtr,
            'Net P&L': tot_pnl,
            'Peak Capital': peak_capital,
            'Max DD': max_dd_val,
            'DD % Init': dd_pct_init,
            'SL Hits': sl_hits,
            'Win Rate': win_rate
        })
        
    df_res = pd.DataFrame(q_results)
    
    tot_pnl_12q  = df_res['Net P&L'].sum()
    avg_cap_12q  = df_res['Peak Capital'].mean()
    avg_dd_val   = df_res['Max DD'].mean()
    worst_dd_val = df_res['Max DD'].max()
    avg_dd_pct   = df_res['DD % Init'].mean()
    worst_dd_pct = df_res['DD % Init'].min()
    tot_sl_hits  = df_res['SL Hits'].sum()
    calmar_ratio = (tot_pnl_12q / worst_dd_val) if worst_dd_val > 0 else 0.0
    
    return {
        'SL Level (%)': f"-{sl_pct:.0%}" if sl_pct is not None else "No SL",
        'SL Value': sl_pct if sl_pct is not None else 1.0,
        '12-Qtr Net Profit (₹)': tot_pnl_12q,
        'Avg Quarterly Capital (₹)': avg_cap_12q,
        'Avg Max Drawdown (₹)': avg_dd_val,
        'Worst Quarter DD (₹)': worst_dd_val,
        'Avg DD % (Initial Cap)': avg_dd_pct,
        'Worst Single Quarter DD %': worst_dd_pct,
        'Total SL Hits': tot_sl_hits,
        'Calmar Ratio (Profit / Worst DD)': calmar_ratio
    }

print("Running Fast Optimization Sweep from -5% to -90%...", flush=True)

sweep_results = []
# Test No SL
sweep_results.append(evaluate_sl_level(None))

# Sweep -5% to -90% in 5% increments
for sl_val in np.arange(0.05, 0.95, 0.05):
    res = evaluate_sl_level(round(sl_val, 2))
    sweep_results.append(res)

df_sweep = pd.DataFrame(sweep_results)

best_profit_row = df_sweep.loc[df_sweep['12-Qtr Net Profit (₹)'].idxmax()]
best_calmar_row = df_sweep.loc[df_sweep['Calmar Ratio (Profit / Worst DD)'].idxmax()]
best_dd_row     = df_sweep.loc[df_sweep['Avg DD % (Initial Cap)'].idxmax()]

print("========================================================================================================================")
print("STOP LOSS OPTIMIZATION RESULTS SUMMARY (ALL 600 TRADES)")
print("========================================================================================================================")
print(f"🥇 MAXIMUM PROFIT SWEET SPOT       : {best_profit_row['SL Level (%)']} SL -> Profit: ₹{best_profit_row['12-Qtr Net Profit (₹)']:,.2f} | Avg DD: {best_profit_row['Avg DD % (Initial Cap)']:.2%}")
print(f"🛡️ MAXIMUM RISK-ADJUSTED (CALMAR) : {best_calmar_row['SL Level (%)']} SL -> Calmar: {best_calmar_row['Calmar Ratio (Profit / Worst DD)']:.2f} | Profit: ₹{best_calmar_row['12-Qtr Net Profit (₹)']:,.2f} | Avg DD: {best_calmar_row['Avg DD % (Initial Cap)']:.2%}")
print(f"🔒 MINIMUM DRAWDOWN SWEET SPOT    : {best_dd_row['SL Level (%)']} SL -> Avg DD: {best_dd_row['Avg DD % (Initial Cap)']:.2%} | Profit: ₹{best_dd_row['12-Qtr Net Profit (₹)']:,.2f}")
print("========================================================================================================================")

print("\nFull Sweep Results Table:")
print(df_sweep[['SL Level (%)', '12-Qtr Net Profit (₹)', 'Avg Max Drawdown (₹)', 'Avg DD % (Initial Cap)', 'Worst Single Quarter DD %', 'Total SL Hits', 'Calmar Ratio (Profit / Worst DD)']].to_string(index=False))

df_sweep.to_excel(OUTPUT_OPT_EXCEL, index=False)
print(f"\n🎉 FAST OPTIMIZATION SWEEP COMPLETED IN {time.time() - t0:.1f}s!", flush=True)
