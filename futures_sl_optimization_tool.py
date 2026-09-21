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
SOURCE_FUT_MASTER = BASE_DIR / 'Nifty50_12_Quarters_Point_In_Time_Futures_Master.xlsx'
OUTPUT_FUT_EXCEL  = BASE_DIR / 'Futures_Stop_Loss_Optimization_Report.xlsx'

quarters_order = [
    'Q3 2023-24', 'Q4 2023-24', 'Q1 2024-25', 'Q2 2024-25',
    'Q3 2024-25', 'Q4 2024-25', 'Q1 2025-26', 'Q2 2025-26',
    'Q3 2025-26', 'Q4 2025-26', 'Q1 2026-27', 'Q2 2026-27'
]

print("Loading Futures Trade Data for Fast Optimization Engine...", flush=True)
xl_fut = pd.ExcelFile(SOURCE_FUT_MASTER, engine='openpyxl')

quarter_data = {}
for qtr in quarters_order:
    df_q = xl_fut.parse(qtr, header=6)
    df_q['Entry Date DT'] = pd.to_datetime(df_q['Entry Date'])
    df_q['Exit Date DT']  = pd.to_datetime(df_q['Exit Date'])
    t_no_col = next((c for c in df_q.columns if 'Trade' in str(c)), df_q.columns[0])
    df_q = df_q.sort_values(by=['Entry Date DT', t_no_col]).reset_index(drop=True)
    
    # Pre-calculate Peak Margin Capital (PIT Lot Margin ~ 20% of contract value)
    min_date = df_q['Entry Date DT'].min()
    max_date = df_q['Exit Date DT'].max()
    date_range = pd.date_range(min_date, max_date)
    
    peak_cap = 0.0
    for d in date_range:
        active_trades = df_q[(df_q['Entry Date DT'] <= d) & (d <= df_q['Exit Date DT'])]
        # 20% margin outlay per futures contract
        tot_margin = (active_trades['Entry Futures Price (₹)'] * active_trades['Lot Size (Qty)'] * 0.20).sum()
        if tot_margin > peak_cap:
            peak_cap = tot_margin
            
    quarter_data[qtr] = {
        'df': df_q,
        'peak_capital': peak_cap
    }

def evaluate_futures_sl(sl_pct):
    q_results = []
    
    for qtr in quarters_order:
        q_info = quarter_data[qtr]
        q_tr   = q_info['df']
        peak_capital = q_info['peak_capital']
        
        s_entries = q_tr['Entry Futures Price (₹)'].values
        s_exits   = q_tr['Exit Futures Price (₹)'].values
        lot_sizes = q_tr['Lot Size (Qty)'].values
        strats    = q_tr['Strategy'].astype(str).str.upper().values
        
        is_long = np.vectorize(lambda s: 'LONG' in s or 'BUY' in s)(strats)
        
        raw_rets = np.where(is_long, (s_exits - s_entries)/s_entries, (s_entries - s_exits)/s_entries)
        raw_rets = np.nan_to_num(raw_rets, nan=0.0)
        
        if sl_pct is not None:
            eff_rets = np.where(raw_rets < -sl_pct, -sl_pct, raw_rets)
            sl_hits  = np.sum(raw_rets < -sl_pct)
        else:
            eff_rets = raw_rets
            sl_hits  = 0
            
        contract_vals = s_entries * lot_sizes
        trade_pnls    = np.round(eff_rets * contract_vals, 2)
        tot_pnl       = np.round(np.sum(trade_pnls), 2)
        wins          = np.sum(trade_pnls > 0)
        win_rate      = wins / len(trade_pnls) if len(trade_pnls) > 0 else 0.0
        
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
    global_wr    = df_res['Win Rate'].mean()
    
    return {
        'SL Level (% Spot Move)': f"-{sl_pct:.2%}" if sl_pct is not None else "No SL",
        'SL Value': sl_pct if sl_pct is not None else 1.0,
        '12-Qtr Net Profit (₹)': tot_pnl_12q,
        'Avg Quarterly Peak Margin (₹)': avg_cap_12q,
        'Avg Max Drawdown (₹)': avg_dd_val,
        'Worst Quarter DD (₹)': worst_dd_val,
        'Avg DD % (Initial Cap)': avg_dd_pct,
        'Worst Single Quarter DD %': worst_dd_pct,
        'Total SL Hits': tot_sl_hits,
        'Win Rate (%)': global_wr,
        'Calmar Ratio (Profit / Worst DD)': calmar_ratio
    }

print("Running Futures Optimization Sweep from -0.25% to -5.00% Spot Move...", flush=True)

sweep_results = []
# Test No SL
sweep_results.append(evaluate_futures_sl(None))

# Sweep -0.25% to -5.00% in 0.25% increments
for sl_val in np.arange(0.0025, 0.0525, 0.0025):
    res = evaluate_futures_sl(round(sl_val, 4))
    sweep_results.append(res)

df_sweep = pd.DataFrame(sweep_results)

best_profit_row = df_sweep.loc[df_sweep['12-Qtr Net Profit (₹)'].idxmax()]
best_calmar_row = df_sweep.loc[df_sweep['Calmar Ratio (Profit / Worst DD)'].idxmax()]
best_dd_row     = df_sweep.loc[df_sweep['Avg DD % (Initial Cap)'].idxmax()]

print("========================================================================================================================")
print("FUTURES STOP LOSS OPTIMIZATION RESULTS SUMMARY (ALL 600 TRADES)")
print("========================================================================================================================")
print(f"🥇 MAXIMUM PROFIT SWEET SPOT       : {best_profit_row['SL Level (% Spot Move)']} SL -> Profit: ₹{best_profit_row['12-Qtr Net Profit (₹)']:,.2f} | Avg DD: {best_profit_row['Avg DD % (Initial Cap)']:.2%}")
print(f"🛡️ MAXIMUM RISK-ADJUSTED (CALMAR) : {best_calmar_row['SL Level (% Spot Move)']} SL -> Calmar: {best_calmar_row['Calmar Ratio (Profit / Worst DD)']:.2f} | Profit: ₹{best_calmar_row['12-Qtr Net Profit (₹)']:,.2f} | Avg DD: {best_calmar_row['Avg DD % (Initial Cap)']:.2%}")
print(f"🔒 MINIMUM DRAWDOWN SWEET SPOT    : {best_dd_row['SL Level (% Spot Move)']} SL -> Avg DD: {best_dd_row['Avg DD % (Initial Cap)']:.2%} | Profit: ₹{best_dd_row['12-Qtr Net Profit (₹)']:,.2f}")
print("========================================================================================================================")

print("\nFull Sweep Results Table:")
print(df_sweep[['SL Level (% Spot Move)', '12-Qtr Net Profit (₹)', 'Avg Max Drawdown (₹)', 'Avg DD % (Initial Cap)', 'Worst Single Quarter DD %', 'Total SL Hits', 'Win Rate (%)', 'Calmar Ratio (Profit / Worst DD)']].to_string(index=False))

df_sweep.to_excel(OUTPUT_FUT_EXCEL, index=False)
print(f"\n🎉 FUTURES OPTIMIZATION SWEEP COMPLETED IN {time.time() - t0:.1f}s!", flush=True)
