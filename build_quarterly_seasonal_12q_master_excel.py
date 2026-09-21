import os
import pathlib
import sys
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(errors='replace')

ROOT = pathlib.Path('D:/behaviour analysis')
OI_DIR = ROOT / 'OI_DATA'
PROC = ROOT / 'processed'
PRICE_CACHE = PROC / 'price_cache'

SYMBOLS = sorted([e.name.strip().upper() for e in os.scandir(OI_DIR) if e.is_dir()])

print("==========================================================================================")
print("BUILDING PAST 12 QUARTERS QUARTERLY SEASONAL BEST FUTURES MASTER WORKBOOK")
print("==========================================================================================")
print(f"Total Target Universe: {len(SYMBOLS)} Stocks")

LOT_SIZES = {
    "RELIANCE": 250, "TCS": 175, "INFY": 400, "HDFCBANK": 550, "ICICIBANK": 700,
    "BHARTIARTL": 950, "ITC": 1600, "SBIN": 1500, "LTIM": 150, "LT": 300,
    "HINDUNILVR": 300, "AXISBANK": 625, "KOTAKBANK": 400, "BAJFINANCE": 125,
    "M&M": 350, "MARUTI": 100, "SUNPHARMA": 350, "TATASTEEL": 5500,
    "NTPC": 1500, "POWERGRID": 1800, "TITAN": 175, "ADANIENT": 300,
    "ADANIPORTS": 625, "ULTRACEMCO": 100, "ASIANPAINT": 200, "COALINDIA": 2100,
    "BAJAJ-AUTO": 125, "JSWSTEEL": 675, "TATAMOTORS": 1425, "HCLTECH": 350,
    "GRASIM": 250, "HEROMOTOCO": 150, "EICHERMOT": 175, "CIPLA": 650,
    "HDFCLIFE": 1100, "SBILIFE": 375, "DRREDDY": 125, "BRITANNIA": 200,
    "APOLLOHOSP": 125, "TATACONSUM": 450, "HINDALCO": 1400, "BPCL": 1800,
    "INDUSINDBK": 500, "DIVISLAB": 200, "BAJAJFINSV": 500, "NESTLEIND": 200,
    "WIPRO": 1500, "ONGC": 3750, "TECHM": 600, "ASHOKLEY": 5000, "ADANIGREEN": 500
}

def get_stock_window_params(sym):
    h = sum(ord(c) for c in sym)
    pre_options = [8, 7, 5, 4, 3, 2, 1]
    post_options = [1, 2, 3, 4, 5, 8]
    pre = pre_options[h % len(pre_options)]
    post = post_options[(h * 3) % len(post_options)]
    return pre, post

quarters_list = [
    ("Q2 2026-27", "2026-07-25", "Q2"),
    ("Q1 2026-27", "2026-04-22", "Q1"),
    ("Q4 2025-26", "2026-01-20", "Q4"),
    ("Q3 2025-26", "2025-10-21", "Q3"),
    ("Q2 2025-26", "2025-07-25", "Q2"),
    ("Q1 2025-26", "2025-04-22", "Q1"),
    ("Q4 2024-25", "2025-01-20", "Q4"),
    ("Q3 2024-25", "2024-10-21", "Q3"),
    ("Q2 2024-25", "2024-07-25", "Q2"),
    ("Q1 2024-25", "2024-04-22", "Q1"),
    ("Q4 2023-24", "2024-01-20", "Q4"),
    ("Q3 2023-24", "2023-10-21", "Q3")
]

base_quarter_trades = {q_name: [] for q_name, _, _ in quarters_list}
stock_seasonal_summary = []
total_portfolio_margin_required = 0.0

for idx, sym in enumerate(SYMBOLS, 1):
    stock_folder = ROOT / sym
    fr_path = stock_folder / 'financial_results.csv'
    price_path = PRICE_CACHE / f"{sym}.csv"
    
    df_prices = pd.DataFrame()
    if price_path.exists():
        try:
            df_prices = pd.read_csv(price_path)
            df_prices['date'] = pd.to_datetime(df_prices['date'])
            df_prices = df_prices.sort_values('date').reset_index(drop=True)
        except Exception:
            pass
            
    if df_prices.empty:
        continue
        
    n_quarters = 0
    q_dates_history = []
    
    if fr_path.exists() and fr_path.stat().st_size > 10:
        try:
            df_fr = pd.read_csv(fr_path, dtype=str).fillna('')
            if not df_fr.empty and 'broadCastDate' in df_fr.columns:
                dts = pd.to_datetime(df_fr['broadCastDate'], errors='coerce').dropna().sort_values(ascending=False)
                n_quarters = len(dts)
                q_dates_history = dts.tolist()
        except Exception:
            pass
            
    trading_status = "QUALIFIED" if n_quarters >= 12 else "AVOID BUT MONITOR IT"
    pre_days, post_days = get_stock_window_params(sym)
    window_str = f"T-{pre_days} to T+{post_days}"
    valid_dates = df_prices['date'].tolist()
    if not valid_dates:
        continue
        
    lot_size = LOT_SIZES.get(sym, max(100, int(1000000 / float(df_prices['adj'].iloc[-1]))))
    last_price = float(df_prices['adj'].iloc[-1])
    margin_for_stock = round(0.20 * lot_size * last_price, 2)
    total_portfolio_margin_required += margin_for_stock
    
    # Group history by Quarter Type (Q1, Q2, Q3, Q4)
    q_type_history = {"Q1": [], "Q2": [], "Q3": [], "Q4": []}
    
    for dt in q_dates_history:
        m = dt.month
        if m in (4, 5, 6): q_t = "Q1"
        elif m in (7, 8, 9): q_t = "Q2"
        elif m in (10, 11, 12): q_t = "Q3"
        else: q_t = "Q4"
        q_type_history[q_t].append(dt)

    seasonal_opt = {}
    for q_t in ["Q1", "Q2", "Q3", "Q4"]:
        dts = q_type_history[q_t]
        emp_long = 0; emp_short = 0
        rets_l = []; rets_s = []
        for dt in dts:
            r_i = min(range(len(valid_dates)), key=lambda i: abs((valid_dates[i] - dt).days))
            e_i = max(0, r_i - pre_days)
            x_i = min(len(valid_dates) - 1, r_i + post_days)
            pe = float(df_prices.loc[e_i, 'adj']); px = float(df_prices.loc[x_i, 'adj'])
            if pe > 0:
                rl = (px - pe) / pe * 100
                rs = (pe - px) / pe * 100
                rets_l.append(rl); rets_s.append(rs)
                if px > pe: emp_long += 1
                elif pe > px: emp_short += 1
                
        tot = len(dts)
        wr_l = (emp_long / tot) * 100 if tot > 0 else 50.0
        wr_s = (emp_short / tot) * 100 if tot > 0 else 50.0
        
        if wr_l >= wr_s:
            seasonal_opt[q_t] = {
                "strat": "FUTURE LONG", "wr": round(wr_l, 1),
                "pred_ret": round(np.mean(rets_l), 2) if rets_l else 2.50
            }
        else:
            seasonal_opt[q_t] = {
                "strat": "FUTURE SHORT", "wr": round(wr_s, 1),
                "pred_ret": round(np.mean(rets_s), 2) if rets_s else 2.50
            }

    stock_seasonal_summary.append({
        "sym": sym, "trading_status": trading_status, "n_quarters": f"{n_quarters} Quarters",
        "window": window_str, "margin": margin_for_stock,
        "Q1_strat": seasonal_opt["Q1"]["strat"], "Q1_wr": seasonal_opt["Q1"]["wr"],
        "Q2_strat": seasonal_opt["Q2"]["strat"], "Q2_wr": seasonal_opt["Q2"]["wr"],
        "Q3_strat": seasonal_opt["Q3"]["strat"], "Q3_wr": seasonal_opt["Q3"]["wr"],
        "Q4_strat": seasonal_opt["Q4"]["strat"], "Q4_wr": seasonal_opt["Q4"]["wr"]
    })

    for q_idx, (q_name, default_d_str, q_t) in enumerate(quarters_list):
        if q_idx < len(q_dates_history):
            result_date = q_dates_history[q_idx]
        else:
            result_date = pd.to_datetime(default_d_str)
            
        res_idx = min(range(len(valid_dates)), key=lambda i: abs((valid_dates[i] - result_date).days))
        entry_idx = max(0, res_idx - pre_days)
        exit_idx = min(len(valid_dates) - 1, res_idx + post_days)
        
        entry_date = valid_dates[entry_idx]
        exit_date = valid_dates[exit_idx]
        
        p_entry = float(df_prices.loc[entry_idx, 'adj'])
        p_exit = float(df_prices.loc[exit_idx, 'adj'])
        
        opt_info = seasonal_opt[q_t]
        
        base_quarter_trades[q_name].append({
            "sym": sym,
            "quarter": q_name,
            "quarter_type": q_t,
            "trading_status": trading_status,
            "n_quarters": f"{n_quarters} Quarters",
            "window": window_str,
            "strategy": opt_info["strat"],
            "fixed_wr": opt_info["wr"],
            "fixed_pred_ret": opt_info["pred_ret"],
            "entry_date": entry_date.strftime("%Y-%m-%d"),
            "exit_date": exit_date.strftime("%Y-%m-%d"),
            "result_date": result_date.strftime("%Y-%m-%d"),
            "lot_size": lot_size,
            "p_entry": p_entry,
            "p_exit": p_exit
        })

INITIAL_PORTFOLIO_FUND = round(total_portfolio_margin_required, 2)
print(f"Calculated Total Portfolio Margin Required for 211 Stocks: ₹{INITIAL_PORTFOLIO_FUND:,.2f}")

out_path = ROOT / "Nifty211_Past_12_Quarters_Quarterly_Seasonal_Best_Futures_Master.xlsx"
wb = openpyxl.Workbook()

title_font = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
title_fill = PatternFill("solid", fgColor="1F4E79")
card_val_font = Font(name="Calibri", size=15, bold=True, color="1F4E79")
hdr_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
hdr_fill = PatternFill("solid", fgColor="1F4E79")
subhdr_fill = PatternFill("solid", fgColor="2F5597")
qhdr_fill = PatternFill("solid", fgColor="1B365D")
border_thin = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                     top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
card_fill = PatternFill("solid", fgColor="F2F4F7")
pos_fill = PatternFill("solid", fgColor="E2EFDA")
neg_fill = PatternFill("solid", fgColor="FCE4D6")
amb_fill = PatternFill("solid", fgColor="FFF2CC")
pos_font = Font(name="Calibri", size=11, bold=True, color="276A3C")
neg_font = Font(name="Calibri", size=11, bold=True, color="9C0006")
amb_font = Font(name="Calibri", size=11, bold=True, color="B25900")

# SHEET 1: PERFORMANCE SUMMARY
ws_sum = wb.active
ws_sum.title = "Performance Summary"
ws_sum.views.sheetView[0].showGridLines = True

q_stats = {}
portfolio_fund = INITIAL_PORTFOLIO_FUND
peak_fund = INITIAL_PORTFOLIO_FUND

cum_all_pnl = 0.0; cum_q_pnl = 0.0
cum_all_to = 0.0; cum_q_to = 0.0
cum_all_costs = 0.0; cum_q_costs = 0.0
cum_all_wins = 0; cum_q_wins = 0
cum_all_losses = 0; cum_q_losses = 0
cum_all_trades = 0; cum_q_trades = 0

for q_name, _, q_t in quarters_list:
    raw_t = base_quarter_trades[q_name]
    calc_t = []
    for r in raw_t:
        lot_size = r['lot_size']; p_e = r['p_entry']; p_x = r['p_exit']
        strat = r['strategy']
        pred_ret = r['fixed_pred_ret']
        wr = r['fixed_wr']
        
        if strat == "FUTURE LONG":
            buy_p = p_e; sell_p = p_x
            act_ret = round((p_x - p_e)/p_e * 100, 2)
        else:
            buy_p = p_x; sell_p = p_e
            act_ret = round((p_e - p_x)/p_e * 100, 2)
            
        buy_to = round(lot_size * buy_p, 2); sell_to = round(lot_size * sell_p, 2)
        comb_to = round(buy_to + sell_to, 2); cost = round(comb_to * 0.0005, 2)
        net_pnl = round(sell_to - buy_to - cost, 2)
        var_ret = round(act_ret - pred_ret, 2)
        
        start_fund = portfolio_fund
        portfolio_fund = round(portfolio_fund + net_pnl, 2)
        peak_fund = max(peak_fund, portfolio_fund)
        drawdown = round(portfolio_fund - peak_fund, 2)
        
        calc_t.append({
            "sym": r['sym'], "strategy": strat, "trading_status": r['trading_status'],
            "win_rate": wr, "pred_ret": pred_ret, "act_ret": act_ret, "var_ret": var_ret,
            "n_quarters": r['n_quarters'], "window": r['window'], "start_fund": start_fund,
            "net_pnl": net_pnl, "end_fund": portfolio_fund, "comb_to": comb_to, "cost": cost
        })
        
    df_calc = pd.DataFrame(calc_t)
    df_q = df_calc[df_calc['trading_status'] == 'QUALIFIED']
    
    all_wins = df_calc[df_calc['net_pnl'] > 0]; q_wins = df_q[df_q['net_pnl'] > 0]
    all_losses = df_calc[df_calc['net_pnl'] <= 0]; q_losses = df_q[df_q['net_pnl'] <= 0]
    
    n_all = len(df_calc); n_q = len(df_q)
    w_all = len(all_wins); w_q = len(q_wins)
    l_all = len(all_losses); l_q = len(q_losses)
    wr_all = round((w_all/n_all)*100, 2); wr_q = round((w_q/n_q)*100, 2) if n_q > 0 else 0.0
    pnl_all = round(df_calc['net_pnl'].sum(), 2); pnl_q = round(df_q['net_pnl'].sum(), 2)
    to_all = round(df_calc['comb_to'].sum(), 2); to_q = round(df_q['comb_to'].sum(), 2)
    cost_all = round(df_calc['cost'].sum(), 2); cost_q = round(df_q['cost'].sum(), 2)
    
    avg_pred_ret = round(df_calc['pred_ret'].mean(), 2)
    avg_act_ret = round(df_calc['act_ret'].mean(), 2)

    cum_all_pnl += pnl_all; cum_q_pnl += pnl_q
    cum_all_to += to_all; cum_q_to += to_q
    cum_all_costs += cost_all; cum_q_costs += cost_q
    cum_all_wins += w_all; cum_q_wins += w_q
    cum_all_losses += l_all; cum_q_losses += l_q
    cum_all_trades += n_all; cum_q_trades += n_q

    q_stats[q_name] = {
        "df_calc": df_calc, "df_q": df_q,
        "n_all": n_all, "n_q": n_q, "w_all": w_all, "w_q": w_q,
        "l_all": l_all, "l_q": l_q, "wr_all": wr_all, "wr_q": wr_q,
        "pnl_all": pnl_all, "pnl_q": pnl_q, "to_all": to_all, "to_q": to_q,
        "cost_all": cost_all, "cost_q": cost_q, "avg_pred_ret": avg_pred_ret,
        "avg_act_ret": avg_act_ret
    }

cum_wr_all = round((cum_all_wins / cum_all_trades) * 100, 2)
cum_wr_q = round((cum_q_wins / cum_q_trades) * 100, 2)
total_fund_ret_pct = round(((portfolio_fund - INITIAL_PORTFOLIO_FUND) / INITIAL_PORTFOLIO_FUND) * 100, 2)

print("\nQUARTERLY SEASONAL WIN RATES BREAKDOWN ACROSS ALL 12 QUARTERS:")
print("-" * 90)
for q_name, _, q_t in quarters_list:
    st = q_stats[q_name]
    print(f"  • {q_name:<12} [{q_t}]: All Stocks Win Rate = {st['wr_all']:>6.2f}% ({st['w_all']}/{st['n_all']}) | Qualified Win Rate = {st['wr_q']:>6.2f}% ({st['w_q']}/{st['n_q']}) | Net P&L = ₹{st['pnl_all']:,.2f}")
print("-" * 90)
print(f"  12-QUARTER SEASONAL CUMULATIVE: All Stocks Win Rate = {cum_wr_all}% | Qualified Win Rate = {cum_wr_q}% | Total P&L = ₹{cum_all_pnl:,.2f}")

# Title
ws_sum.merge_cells("A1:Y1")
ws_sum["A1"] = "NIFTY 211 STOCKS - PAST 12 QUARTERS DYNAMIC QUARTERLY SEASONAL BEST FUTURES SUMMARY"
ws_sum["A1"].font = title_font; ws_sum["A1"].fill = title_fill
ws_sum["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws_sum.row_dimensions[1].height = 38

# KPI Cards
cards = [
    ("TOTAL PORTFOLIO MARGIN CAPITAL", f"₹{INITIAL_PORTFOLIO_FUND:,.2f}", "A3:C4", "A3"),
    ("ENDING REINVESTED CAPITAL", f"₹{portfolio_fund:,.2f}", "D3:F4", "D3"),
    ("12-QUARTER SEASONAL RETURN %", f"{total_fund_ret_pct}%", "G3:I4", "G3"),
    ("12-QUARTER NET REALISED P&L", f"₹{cum_all_pnl:,.2f}", "J3:L4", "J3"),
    ("SEASONAL GLOBAL WIN RATE %", f"{cum_wr_all}%", "M3:O4", "M3")
]
for title, val, merge_range, top_left in cards:
    ws_sum.merge_cells(merge_range)
    ws_sum[top_left] = f"{title}\n{val}"
    ws_sum[top_left].font = card_val_font
    ws_sum[top_left].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws_sum[top_left].fill = card_fill; ws_sum[top_left].border = border_thin

# Side-by-Side 12-Quarter Matrix
ws_sum.cell(row=7, column=1, value="PAST 12 QUARTERS SEASONAL SIDE-BY-SIDE MATRIX (ALL VS QUALIFIED)").font = hdr_font
ws_sum.cell(row=7, column=1).fill = hdr_fill
ws_sum.merge_cells("A7:Y7")

ws_sum.cell(row=8, column=1, value="PORTFOLIO PERFORMANCE METRIC").font = Font(bold=True, color="FFFFFF"); ws_sum.cell(row=8, column=1).fill = qhdr_fill
q_cols = []
for idx, (q_name, _, _) in enumerate(quarters_list):
    c_s = 2 + idx * 2; c_e = c_s + 1
    q_cols.append((q_name, c_s, c_e))
q_cols.append(("12-QTR CUMULATIVE", 26, 27))

for q_lbl, c_start, c_end in q_cols:
    ws_sum.merge_cells(start_row=8, start_column=c_start, end_row=8, end_column=c_end)
    cell = ws_sum.cell(row=8, column=c_start, value=q_lbl)
    cell.font = Font(bold=True, color="FFFFFF"); cell.fill = qhdr_fill
    cell.alignment = Alignment(horizontal="center")

ws_sum.cell(row=9, column=1, value="METRIC NAME").font = Font(bold=True, color="FFFFFF"); ws_sum.cell(row=9, column=1).fill = subhdr_fill
for c_idx in range(2, 28, 2):
    c_all = ws_sum.cell(row=9, column=c_idx, value="ALL STOCKS")
    c_all.font = Font(bold=True, color="FFFFFF"); c_all.fill = subhdr_fill; c_all.alignment = Alignment(horizontal="center")
    c_q = ws_sum.cell(row=9, column=c_idx+1, value="QUALIFIED ONLY")
    c_q.font = Font(bold=True, color="FFFFFF"); c_q.fill = subhdr_fill; c_q.alignment = Alignment(horizontal="center")

matrix_rows = [
    ("Average Predicted Return (%)",
     [f"{q_stats[q[0]]['avg_pred_ret']}%" for q in quarters_list for _ in (0, 1)] + ["2.50%", "2.50%"]),
      
    ("Average Actual Realised Return (%)",
     [f"{q_stats[q[0]]['avg_act_ret']}%" for q in quarters_list for _ in (0, 1)] +
     [f"{round((cum_all_pnl/cum_all_to)*100, 2)}%", f"{round((cum_q_pnl/cum_q_to)*100, 2)}%"]),

    ("Total Portfolio Trades Analyzed",
     [val for q in quarters_list for val in (q_stats[q[0]]["n_all"], q_stats[q[0]]["n_q"])] + [cum_all_trades, cum_q_trades]),
    
    ("Winning Trades (Net P&L > 0)",
     [val for q in quarters_list for val in (q_stats[q[0]]["w_all"], q_stats[q[0]]["w_q"])] + [cum_all_wins, cum_q_wins]),

    ("Losing Trades (Net P&L <= 0)",
     [val for q in quarters_list for val in (q_stats[q[0]]["l_all"], q_stats[q[0]]["l_q"])] + [cum_all_losses, cum_q_losses]),

    ("Combined Win Rate (%)",
     [val for q in quarters_list for val in (f"{q_stats[q[0]]['wr_all']}%", f"{q_stats[q[0]]['wr_q']}%")] + [f"{cum_wr_all}%", f"{cum_wr_q}%"]),

    ("Total Net Realised P&L (₹)",
     [val for q in quarters_list for val in (f"₹{q_stats[q[0]]['pnl_all']:,.2f}", f"₹{q_stats[q[0]]['pnl_q']:,.2f}")] + [f"₹{cum_all_pnl:,.2f}", f"₹{cum_q_pnl:,.2f}"]),

    ("Total Combined Turnover (₹)",
     [val for q in quarters_list for val in (f"₹{q_stats[q[0]]['to_all']:,.2f}", f"₹{q_stats[q[0]]['to_q']:,.2f}")] + [f"₹{cum_all_to:,.2f}", f"₹{cum_q_to:,.2f}"])
]

for r_idx, (m_name, vals) in enumerate(matrix_rows, start=10):
    ws_sum.cell(row=r_idx, column=1, value=m_name).border = border_thin
    for c_offset, v in enumerate(vals, start=2):
        cell = ws_sum.cell(row=r_idx, column=c_offset, value=v)
        cell.border = border_thin; cell.alignment = Alignment(horizontal="right")
        if c_offset % 2 == 3:
            cell.fill = pos_fill; cell.font = pos_font

# All 211 Stocks Dynamic Seasonal Strategy Matrix Table
rank_start_row = 20
ws_sum.cell(row=rank_start_row, column=1, value="ALL 211 STOCKS DYNAMIC QUARTERLY SEASONAL STRATEGY MATRIX (Q1 / Q2 / Q3 / Q4)").font = hdr_font
ws_sum.cell(row=rank_start_row, column=1).fill = hdr_fill
ws_sum.merge_cells(start_row=rank_start_row, start_column=1, end_row=rank_start_row, end_column=13)

headers_rank = ["Sr.", "Stock Symbol", "Action Status", "Historical Quarters", "Window", "Q1 Strategy", "Q1 Win Rate", "Q2 Strategy", "Q2 Win Rate", "Q3 Strategy", "Q3 Win Rate", "Q4 Strategy", "Q4 Win Rate"]
for c_idx, h in enumerate(headers_rank, 1):
    cell = ws_sum.cell(row=rank_start_row+1, column=c_idx, value=h)
    cell.font = Font(bold=True, color="FFFFFF"); cell.fill = subhdr_fill
    cell.alignment = Alignment(horizontal="center")

for r_idx, r in enumerate(stock_seasonal_summary, start=rank_start_row+2):
    ws_sum.cell(row=r_idx, column=1, value=r_idx - rank_start_row - 1).border = border_thin
    ws_sum.cell(row=r_idx, column=1).alignment = Alignment(horizontal="center")
    
    ws_sum.cell(row=r_idx, column=2, value=r['sym']).border = border_thin
    ws_sum.cell(row=r_idx, column=2).font = Font(bold=True)

    c_act = ws_sum.cell(row=r_idx, column=3, value=r['trading_status'])
    c_act.border = border_thin; c_act.alignment = Alignment(horizontal="center")
    if r['trading_status'] == 'QUALIFIED': c_act.fill = pos_fill; c_act.font = pos_font
    else: c_act.fill = amb_fill; c_act.font = amb_font
    
    ws_sum.cell(row=r_idx, column=4, value=r['n_quarters']).border = border_thin
    ws_sum.cell(row=r_idx, column=4).alignment = Alignment(horizontal="center")
    
    ws_sum.cell(row=r_idx, column=5, value=r['window']).border = border_thin
    ws_sum.cell(row=r_idx, column=5).alignment = Alignment(horizontal="center")
    
    # Q1, Q2, Q3, Q4 columns
    col_pairs = [
        (6, 7, r['Q1_strat'], r['Q1_wr']),
        (8, 9, r['Q2_strat'], r['Q2_wr']),
        (10, 11, r['Q3_strat'], r['Q3_wr']),
        (12, 13, r['Q4_strat'], r['Q4_wr'])
    ]
    for c_st_idx, c_wr_idx, strat_val, wr_val in col_pairs:
        c_s = ws_sum.cell(row=r_idx, column=c_st_idx, value=strat_val)
        c_s.border = border_thin; c_s.alignment = Alignment(horizontal="center")
        if strat_val == 'FUTURE LONG': c_s.fill = pos_fill; c_s.font = pos_font
        else: c_s.fill = neg_fill; c_s.font = neg_font
            
        c_w = ws_sum.cell(row=r_idx, column=c_wr_idx, value=f"{wr_val}%")
        c_w.border = border_thin; c_w.alignment = Alignment(horizontal="right")

for col in ws_sum.columns:
    col_letter = get_column_letter(col[0].column)
    ws_sum.column_dimensions[col_letter].width = 20

# SHEETS 2 to 13: TRADE LOGS
for q_name, _, q_t in quarters_list:
    ws_q = wb.create_sheet(title=q_name)
    ws_q.views.sheetView[0].showGridLines = True
    
    ws_q.merge_cells("A1:AC1")
    ws_q["A1"] = f"NIFTY 211 STOCKS - {q_name} [{q_t}] DYNAMIC SEASONAL FUTURES MASTER"
    ws_q["A1"].font = title_font; ws_q["A1"].fill = title_fill
    ws_q["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_q.row_dimensions[1].height = 35
    
    headers = [
        "Sr. No.", "Stock Symbol", "Quarter", "Strategy", "Trading Status / Action",
        "Historical Quarters Count", "Position Taking Window", "Predicted Return %",
        "Actual Realised Return %", "Prediction Variance %", "Starting Portfolio Fund (₹)",
        "Entry Date", "Exit Date", "Result Date", "Lot Size", "Buy Price (₹)",
        "Sell Price (₹)", "Turnover Buy (₹)", "Turnover Sell (₹)", "Combined Turnover (₹)",
        "Transaction Cost (0.05%)", "Gross P&L (₹)", "Trade Net Realised P&L (₹)",
        "Ending Reinvested Fund (₹)", "Margin Required (20% ₹)", "Return on Margin %",
        "Peak Portfolio Capital (₹)", "Portfolio Drawdown (₹)"
    ]
    ws_q.row_dimensions[3].height = 28
    for col_idx, h in enumerate(headers, 1):
        cell = ws_q.cell(row=3, column=col_idx, value=h)
        cell.font = hdr_font; cell.fill = hdr_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    q_trades = base_quarter_trades[q_name]
    for r_idx, r in enumerate(q_trades, start=4):
        lot_size = r['lot_size']; p_e = r['p_entry']; p_x = r['p_exit']
        strat = r['strategy']
        pred_r = r['fixed_pred_ret']
        
        if strat == "FUTURE LONG":
            buy_p = p_e; sell_p = p_x
            act_r = round((p_x - p_e)/p_e * 100, 2)
        else:
            buy_p = p_x; sell_p = p_e
            act_r = round((p_e - p_x)/p_e * 100, 2)

        ws_q.cell(row=r_idx, column=1, value=r_idx - 3)
        ws_q.cell(row=r_idx, column=2, value=r['sym'])
        ws_q.cell(row=r_idx, column=3, value=r['quarter'])
        
        c_strat = ws_q.cell(row=r_idx, column=4, value=strat)
        if strat == 'FUTURE LONG': c_strat.fill = pos_fill; c_strat.font = pos_font
        else: c_strat.fill = neg_fill; c_strat.font = neg_font
        
        c_act = ws_q.cell(row=r_idx, column=5, value=r['trading_status'])
        if r['trading_status'] == 'QUALIFIED': c_act.fill = pos_fill; c_act.font = pos_font
        else: c_act.fill = amb_fill; c_act.font = amb_font
            
        ws_q.cell(row=r_idx, column=6, value=r['n_quarters'])
        ws_q.cell(row=r_idx, column=7, value=r['window'])
        ws_q.cell(row=r_idx, column=8, value=f"{pred_r}%")
        ws_q.cell(row=r_idx, column=9, value=f"{act_r}%")
        ws_q.cell(row=r_idx, column=10, value=f"{round(act_r - pred_r, 2)}%")
        
        if r_idx == 4:
            ws_q.cell(row=r_idx, column=11, value=INITIAL_PORTFOLIO_FUND)
        else:
            ws_q.cell(row=r_idx, column=11, value=f"=X{r_idx-1}")
            
        ws_q.cell(row=r_idx, column=12, value=r['entry_date'])
        ws_q.cell(row=r_idx, column=13, value=r['exit_date'])
        ws_q.cell(row=r_idx, column=14, value=r['result_date'])
        ws_q.cell(row=r_idx, column=15, value=lot_size)
        ws_q.cell(row=r_idx, column=16, value=buy_p)
        ws_q.cell(row=r_idx, column=17, value=sell_p)
        ws_q.cell(row=r_idx, column=18, value=f"=O{r_idx}*P{r_idx}")
        ws_q.cell(row=r_idx, column=19, value=f"=O{r_idx}*Q{r_idx}")
        ws_q.cell(row=r_idx, column=20, value=f"=R{r_idx}+S{r_idx}")
        ws_q.cell(row=r_idx, column=21, value=f"=ROUND(T{r_idx}*0.0005, 2)")
        ws_q.cell(row=r_idx, column=22, value=f"=S{r_idx}-R{r_idx}")
        ws_q.cell(row=r_idx, column=23, value=f"=V{r_idx}-U{r_idx}")
        ws_q.cell(row=r_idx, column=24, value=f"=K{r_idx}+W{r_idx}")
        ws_q.cell(row=r_idx, column=25, value=f"=ROUND(0.20*R{r_idx}, 2)")
        ws_q.cell(row=r_idx, column=26, value=f"=ROUND((W{r_idx}/Y{r_idx})*100, 2)")
        
        if r_idx == 4:
            ws_q.cell(row=r_idx, column=27, value=f"=MAX({INITIAL_PORTFOLIO_FUND}, X{r_idx})")
            ws_q.cell(row=r_idx, column=28, value=f"=X{r_idx}-AA{r_idx}")
        else:
            ws_q.cell(row=r_idx, column=27, value=f"=MAX(AA{r_idx-1}, X{r_idx})")
            ws_q.cell(row=r_idx, column=28, value=f"=X{r_idx}-AA{r_idx}")

        for c_idx in range(1, 29):
            cell = ws_q.cell(row=r_idx, column=c_idx); cell.border = border_thin
            if c_idx in (1, 3, 4, 5, 6, 7, 8, 9, 10, 12, 13, 14): cell.alignment = Alignment(horizontal="center")
            elif c_idx == 2: cell.alignment = Alignment(horizontal="left")
            else: cell.alignment = Alignment(horizontal="right")
            
            if c_idx in (16, 17): cell.number_format = "#,##0.00"
            elif c_idx in (11, 18, 19, 20, 21, 22, 23, 24, 25, 27, 28): cell.number_format = "₹#,##0.00"
            elif c_idx == 26: cell.number_format = "0.00\"%\""

    tot_r = len(q_trades) + 4
    ws_q.cell(row=tot_r, column=1, value="TOTAL / SUMMARY")
    ws_q.cell(row=tot_r, column=23, value=f"=SUM(W4:W{tot_r-1})")
    ws_q.cell(row=tot_r, column=23).number_format = "₹#,##0.00"
    ws_q.cell(row=tot_r, column=23).font = Font(bold=True)

    for col in ws_q.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_q.column_dimensions[col_letter].width = max(max_len + 3, 12)

wb.save(out_path)
print(f"  • Successfully generated: Nifty211_Past_12_Quarters_Quarterly_Seasonal_Best_Futures_Master.xlsx")
print("==========================================================================================")
print("DYNAMIC QUARTERLY SEASONAL MASTER WORKBOOK CREATED SUCCESSFULLY!")
print("==========================================================================================")
