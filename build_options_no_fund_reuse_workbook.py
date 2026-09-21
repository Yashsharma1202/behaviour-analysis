import openpyxl
import pandas as pd
import numpy as np
import pathlib
import sys
import time
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference

sys.stdout.reconfigure(errors='replace')

t0 = time.time()

BASE_DIR = pathlib.Path('D:/behaviour analysis')
SOURCE_PIT_FUT = BASE_DIR / 'Nifty50_12_Quarters_Point_In_Time_Futures_Master.xlsx'
OUTPUT_NO_REUSE_MASTER = BASE_DIR / 'Nifty50_12_Quarters_Options_1PCT_ITM_No_Fund_Reuse_Master.xlsx'

# Executive Color Tokens
COLOR_TITLE_BG     = '1B365D'
COLOR_SECTION_BG   = '2B4C7E'
COLOR_HEADER_BG    = '1A202C'
COLOR_SUBHEADER_BG = '2C5282'
COLOR_ZEBRA_BG     = 'F8FAFC'
COLOR_CARD_BG      = 'F7FAFC'

COLOR_WIN_BG       = 'E6F4EA'
COLOR_WIN_TEXT     = '137333'
COLOR_LOSS_BG      = 'FCE8E6'
COLOR_LOSS_TEXT    = 'C5221F'
COLOR_GOLD_BG      = 'FEF3C7'
COLOR_GOLD_TEXT    = '92400E'

COLOR_WHITE        = 'FFFFFF'
COLOR_TEXT_MAIN    = '2D3748'

font_title     = Font(name='Segoe UI', size=14, bold=True, color=COLOR_WHITE)
font_section   = Font(name='Segoe UI', size=11, bold=True, color=COLOR_WHITE)
font_header    = Font(name='Segoe UI', size=10, bold=True, color=COLOR_WHITE)
font_card_lbl  = Font(name='Segoe UI', size=9, bold=True, color=COLOR_WHITE)
font_data_bold = Font(name='Segoe UI', size=10, bold=True, color=COLOR_TEXT_MAIN)

font_win       = Font(name='Segoe UI', size=10, bold=True, color=COLOR_WIN_TEXT)
font_loss      = Font(name='Segoe UI', size=10, bold=True, color=COLOR_LOSS_TEXT)
font_gold      = Font(name='Segoe UI', size=10, bold=True, color=COLOR_GOLD_TEXT)

fill_title     = PatternFill(start_color=COLOR_TITLE_BG, end_color=COLOR_TITLE_BG, fill_type='solid')
fill_section   = PatternFill(start_color=COLOR_SECTION_BG, end_color=COLOR_SECTION_BG, fill_type='solid')
fill_header    = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type='solid')
fill_subheader = PatternFill(start_color=COLOR_SUBHEADER_BG, end_color=COLOR_SUBHEADER_BG, fill_type='solid')
fill_zebra     = PatternFill(start_color=COLOR_ZEBRA_BG, end_color=COLOR_ZEBRA_BG, fill_type='solid')
fill_card      = PatternFill(start_color=COLOR_CARD_BG, end_color=COLOR_CARD_BG, fill_type='solid')

fill_win       = PatternFill(start_color=COLOR_WIN_BG, end_color=COLOR_WIN_BG, fill_type='solid')
fill_loss       = PatternFill(start_color=COLOR_LOSS_BG, end_color=COLOR_LOSS_BG, fill_type='solid')
fill_gold      = PatternFill(start_color=COLOR_GOLD_BG, end_color=COLOR_GOLD_BG, fill_type='solid')

border_thin = Border(
    left=Side(style='thin', color='CBD5E0'),
    right=Side(style='thin', color='CBD5E0'),
    top=Side(style='thin', color='CBD5E0'),
    bottom=Side(style='thin', color='CBD5E0')
)

border_double_bottom = Border(
    left=Side(style='thin', color='CBD5E0'),
    right=Side(style='thin', color='CBD5E0'),
    top=Side(style='thin', color='1A202C'),
    bottom=Side(style='double', color='1A202C')
)

def round_nse_strike(price):
    if price > 5000: interval = 50
    elif price > 1000: interval = 20
    elif price > 500: interval = 10
    else: interval = 5
    return round(price / interval) * interval

quarters_order = [
    'Q3 2023-24', 'Q4 2023-24', 'Q1 2024-25', 'Q2 2024-25',
    'Q3 2024-25', 'Q4 2024-25', 'Q1 2025-26', 'Q2 2025-26',
    'Q3 2025-26', 'Q4 2025-26', 'Q1 2026-27', 'Q2 2026-27'
]

xl_pit = pd.ExcelFile(SOURCE_PIT_FUT, engine='openpyxl')

comp_summaries = []

for qtr in quarters_order:
    q_tr = xl_pit.parse(qtr, header=6)
    q_tr['Entry Date DT'] = pd.to_datetime(q_tr['Entry Date'])
    q_tr['Exit Date DT']  = pd.to_datetime(q_tr['Exit Date'])
    
    outlays = []
    pnls = []
    
    for idx, row in q_tr.iterrows():
        s_entry  = float(row['Entry Futures Price (₹)'])
        s_exit   = float(row['Exit Futures Price (₹)'])
        lot_size = int(row['Lot Size (Qty)'])
        strat    = str(row['Strategy']).strip().upper()
        
        is_call  = 'LONG' in strat or 'BUY' in strat
        if is_call:
            strike = round_nse_strike(s_entry * 0.99)
            i_entry = max(0.0, s_entry - strike)
            i_exit  = max(0.0, s_exit - strike)
        else:
            strike = round_nse_strike(s_entry * 1.01)
            i_entry = max(0.0, strike - s_entry)
            i_exit  = max(0.0, strike - s_exit)
            
        p_entry = i_entry + (s_entry * 0.02)
        p_exit  = i_exit + (s_exit * 0.005)
        
        outlay = p_entry * lot_size
        pnl = (p_exit - p_entry) * lot_size
        
        outlays.append(outlay)
        pnls.append(pnl)
        
    q_tr['Opt_Outlay'] = outlays
    q_tr['Booked_PnL'] = pnls
    
    t_count = len(q_tr)
    wins = len(q_tr[q_tr['Booked_PnL'] > 0])
    losses = len(q_tr[q_tr['Booked_PnL'] <= 0])
    win_rate = (wins / t_count) if t_count > 0 else 0.0
    tot_pnl = round(sum(pnls), 2)
    no_reuse_capital = round(sum(outlays), 2)
    
    # Peak Concurrent Capital (Reused Fund Model)
    min_date = q_tr['Entry Date DT'].min()
    max_date = q_tr['Exit Date DT'].max()
    date_range = pd.date_range(min_date, max_date)
    
    reused_capital = 0.0
    for d in date_range:
        active_trades = q_tr[(q_tr['Entry Date DT'] <= d) & (d <= q_tr['Exit Date DT'])]
        opt_tot = active_trades['Opt_Outlay'].sum()
        if opt_tot > reused_capital:
            reused_capital = opt_tot
            
    reused_capital = round(reused_capital, 2)
    
    # Drawdown math
    q_tr['Cum PnL'] = q_tr['Booked_PnL'].cumsum()
    q_tr['Running Max'] = np.maximum.accumulate(q_tr['Cum PnL'])
    q_tr['Drawdown'] = q_tr['Cum PnL'] - q_tr['Running Max']
    
    trough_idx = q_tr['Drawdown'].idxmin()
    max_dd_val = round(abs(q_tr.loc[trough_idx, 'Drawdown']), 2)
    peak_cum_pnl = q_tr.loc[:trough_idx, 'Cum PnL'].max()
    
    # Reused Fund Drawdowns & Return
    ret_reused = round(tot_pnl / reused_capital, 4)
    dd_reused_init = round(-abs(max_dd_val / reused_capital), 4)
    dd_reused_peak = round(-abs(max_dd_val / (reused_capital + max(0.0, peak_cum_pnl))), 4)
    
    # Non-Reused Fund Drawdowns & Return
    ret_no_reuse = round(tot_pnl / no_reuse_capital, 4)
    dd_no_reuse_init = round(-abs(max_dd_val / no_reuse_capital), 4)
    dd_no_reuse_peak = round(-abs(max_dd_val / (no_reuse_capital + max(0.0, peak_cum_pnl))), 4)
    
    comp_summaries.append({
        'Quarter': qtr,
        'Total Trades': t_count,
        'Win Rate (%)': win_rate,
        'Net Option P&L (₹)': tot_pnl,
        'Max Drawdown (₹)': max_dd_val,
        'Reused Capital (₹)': reused_capital,
        'Reused Return (%)': ret_reused,
        'Reused DD % Init': dd_reused_init,
        'Reused DD % Peak': dd_reused_peak,
        'No-Reuse Capital (₹)': no_reuse_capital,
        'No-Reuse Return (%)': ret_no_reuse,
        'No-Reuse DD % Init': dd_no_reuse_init,
        'No-Reuse DD % Peak': dd_no_reuse_peak
    })

df_comp_sum = pd.DataFrame(comp_summaries)

# Build Master Comparison Workbook
wb = Workbook()
wb.remove(wb.active)

ws = wb.create_sheet(title='Fund_Reuse_vs_No_Reuse')
ws.views.sheetView[0].showGridLines = True

ws.merge_cells('A1:L1')
c_t = ws['A1']
c_t.value = "NIFTY 50 1% ITM OPTIONS STRATEGY — FUND REUSE vs NON-FUND REUSE COMPARISON MASTER SCORECARD"
c_t.font = font_title; c_t.fill = fill_title; c_t.alignment = Alignment(horizontal='center', vertical='center')
ws.row_dimensions[1].height = 40

ws.merge_cells('A3:L3')
c_s = ws['A3']; c_s.value = "SIDE-BY-SIDE COMPARISON: REUSED FUND (CONCURRENT PEAK) vs NON-REUSED FUND (SUM OF ALL 50 TRADES)"; c_s.font = font_section; c_s.fill = fill_section
ws.row_dimensions[3].height = 24

headers = [
    'Quarter', 'Net Option P&L (₹)', 'Max DD (₹)',
    'Reused Fund Capital (₹)', 'Reused Return (%)', 'Reused DD % Init', 'Reused DD % Peak',
    'No-Reuse Capital (₹)', 'No-Reuse Return (%)', 'No-Reuse DD % Init', 'No-Reuse DD % Peak', 'Win Rate (%)'
]

for c_idx, h in enumerate(headers, start=1):
    cell = ws.cell(row=4, column=c_idx, value=h)
    cell.font = font_header; cell.fill = fill_header; cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = border_thin
ws.row_dimensions[4].height = 26

for r_idx, r_data in df_comp_sum.iterrows():
    row_num = 5 + r_idx
    row_fill = fill_zebra if r_idx % 2 == 1 else PatternFill(fill_type=None)
    ws.row_dimensions[row_num].height = 21
    
    ws.cell(row=row_num, column=1, value=r_data['Quarter']).alignment = Alignment(horizontal='center')
    
    c_pnl = ws.cell(row=row_num, column=2, value=r_data['Net Option P&L (₹)'])
    c_pnl.number_format = '₹#,##0.00'; c_pnl.fill = fill_win if r_data['Net Option P&L (₹)'] >= 0 else fill_loss
    c_pnl.font = font_win if r_data['Net Option P&L (₹)'] >= 0 else font_loss
    
    c_dd = ws.cell(row=row_num, column=3, value=-abs(r_data['Max Drawdown (₹)']))
    c_dd.number_format = '-₹#,##0.00'; c_dd.font = font_loss; c_dd.fill = fill_loss
    
    # Reused Fund Model
    ws.cell(row=row_num, column=4, value=r_data['Reused Capital (₹)']).number_format = '₹#,##0.00'
    ws.cell(row=row_num, column=5, value=r_data['Reused Return (%)']).number_format = '+0.00%;-0.00%'
    c_r1 = ws.cell(row=row_num, column=6, value=r_data['Reused DD % Init']); c_r1.number_format = '-0.00%'; c_r1.font = font_loss; c_r1.fill = fill_loss; c_r1.alignment = Alignment(horizontal='center')
    c_r2 = ws.cell(row=row_num, column=7, value=r_data['Reused DD % Peak']); c_r2.number_format = '-0.00%'; c_r2.font = font_loss; c_r2.fill = fill_loss; c_r2.alignment = Alignment(horizontal='center')
    
    # Non-Reused Fund Model
    ws.cell(row=row_num, column=8, value=r_data['No-Reuse Capital (₹)']).number_format = '₹#,##0.00'
    ws.cell(row=row_num, column=9, value=r_data['No-Reuse Return (%)']).number_format = '+0.00%;-0.00%'
    c_n1 = ws.cell(row=row_num, column=10, value=r_data['No-Reuse DD % Init']); c_n1.number_format = '-0.00%'; c_n1.font = font_loss; c_n1.fill = fill_loss; c_n1.alignment = Alignment(horizontal='center')
    c_n2 = ws.cell(row=row_num, column=11, value=r_data['No-Reuse DD % Peak']); c_n2.number_format = '-0.00%'; c_n2.font = font_loss; c_n2.fill = fill_loss; c_n2.alignment = Alignment(horizontal='center')
    
    c_wr = ws.cell(row=row_num, column=12, value=r_data['Win Rate (%)']); c_wr.number_format = '0.00%'; c_wr.alignment = Alignment(horizontal='center')
    if r_data['Win Rate (%)'] >= 0.70: c_wr.fill = fill_win; c_wr.font = font_win
    elif r_data['Win Rate (%)'] >= 0.50: c_wr.fill = fill_gold; c_wr.font = font_gold
    else: c_wr.fill = fill_loss; c_wr.font = font_loss

    for c in range(1, 13):
        cell = ws.cell(row=row_num, column=c)
        cell.border = border_thin
        if c not in [2, 3, 6, 7, 10, 11, 12] and row_fill.fill_type: cell.fill = row_fill

# Row 17 Total
r_tot = 17
ws.row_dimensions[r_tot].height = 24
ws.cell(row=r_tot, column=1, value="12-QUARTER TOTAL").font = font_data_bold
ws.cell(row=r_tot, column=2, value=df_comp_sum['Net Option P&L (₹)'].sum()).number_format = '₹#,##0.00'; ws.cell(row=r_tot, column=2).font = font_win; ws.cell(row=r_tot, column=2).fill = fill_win
ws.cell(row=r_tot, column=3, value=-abs(df_comp_sum['Max Drawdown (₹)'].max())).number_format = '-₹#,##0.00'; ws.cell(row=r_tot, column=3).font = font_loss; ws.cell(row=r_tot, column=3).fill = fill_loss
ws.cell(row=r_tot, column=4, value="—").alignment = Alignment(horizontal='center')
ws.cell(row=r_tot, column=5, value="—").alignment = Alignment(horizontal='center')
ws.cell(row=r_tot, column=6, value=df_comp_sum['Reused DD % Init'].min()).number_format = '-0.00%'; ws.cell(row=r_tot, column=6).font = font_loss; ws.cell(row=r_tot, column=6).fill = fill_loss; ws.cell(row=r_tot, column=6).alignment = Alignment(horizontal='center')
ws.cell(row=r_tot, column=7, value=df_comp_sum['Reused DD % Peak'].min()).number_format = '-0.00%'; ws.cell(row=r_tot, column=7).font = font_loss; ws.cell(row=r_tot, column=7).fill = fill_loss; ws.cell(row=r_tot, column=7).alignment = Alignment(horizontal='center')

ws.cell(row=r_tot, column=8, value="—").alignment = Alignment(horizontal='center')
ws.cell(row=r_tot, column=9, value="—").alignment = Alignment(horizontal='center')
ws.cell(row=r_tot, column=10, value=df_comp_sum['No-Reuse DD % Init'].min()).number_format = '-0.00%'; ws.cell(row=r_tot, column=10).font = font_loss; ws.cell(row=r_tot, column=10).fill = fill_loss; ws.cell(row=r_tot, column=10).alignment = Alignment(horizontal='center')
ws.cell(row=r_tot, column=11, value=df_comp_sum['No-Reuse DD % Peak'].min()).number_format = '-0.00%'; ws.cell(row=r_tot, column=11).font = font_loss; ws.cell(row=r_tot, column=11).fill = fill_loss; ws.cell(row=r_tot, column=11).alignment = Alignment(horizontal='center')
c_twr = ws.cell(row=r_tot, column=12, value=0.515); c_twr.number_format = '0.00%'; c_twr.font = font_win; c_twr.fill = fill_win; c_twr.alignment = Alignment(horizontal='center')

for c in range(1, 13): ws.cell(row=r_tot, column=c).border = border_thin

# Row 18 Average
r_avg = 18
ws.row_dimensions[r_avg].height = 24
ws.cell(row=r_avg, column=1, value="QUARTERLY AVERAGE").font = font_data_bold
ws.cell(row=r_avg, column=2, value=df_comp_sum['Net Option P&L (₹)'].mean()).number_format = '₹#,##0.00'; ws.cell(row=r_avg, column=2).font = font_win; ws.cell(row=r_avg, column=2).fill = fill_win
ws.cell(row=r_avg, column=3, value=-abs(df_comp_sum['Max Drawdown (₹)'].mean())).number_format = '-₹#,##0.00'; ws.cell(row=r_avg, column=3).font = font_loss; ws.cell(row=r_avg, column=3).fill = fill_loss

ws.cell(row=r_avg, column=4, value=df_comp_sum['Reused Capital (₹)'].mean()).number_format = '₹#,##0.00'
ws.cell(row=r_avg, column=5, value=df_comp_sum['Reused Return (%)'].mean()).number_format = '+0.00%'
ws.cell(row=r_avg, column=6, value=df_comp_sum['Reused DD % Init'].mean()).number_format = '-0.00%'; ws.cell(row=r_avg, column=6).font = font_loss; ws.cell(row=r_avg, column=6).fill = fill_loss; ws.cell(row=r_avg, column=6).alignment = Alignment(horizontal='center')
ws.cell(row=r_avg, column=7, value=df_comp_sum['Reused DD % Peak'].mean()).number_format = '-0.00%'; ws.cell(row=r_avg, column=7).font = font_loss; ws.cell(row=r_avg, column=7).fill = fill_loss; ws.cell(row=r_avg, column=7).alignment = Alignment(horizontal='center')

ws.cell(row=r_avg, column=8, value=df_comp_sum['No-Reuse Capital (₹)'].mean()).number_format = '₹#,##0.00'
ws.cell(row=r_avg, column=9, value=df_comp_sum['No-Reuse Return (%)'].mean()).number_format = '+0.00%'
ws.cell(row=r_avg, column=10, value=df_comp_sum['No-Reuse DD % Init'].mean()).number_format = '-0.00%'; ws.cell(row=r_avg, column=10).font = font_loss; ws.cell(row=r_avg, column=10).fill = fill_loss; ws.cell(row=r_avg, column=10).alignment = Alignment(horizontal='center')
ws.cell(row=r_avg, column=11, value=df_comp_sum['No-Reuse DD % Peak'].mean()).number_format = '-0.00%'; ws.cell(row=r_avg, column=11).font = font_loss; ws.cell(row=r_avg, column=11).fill = fill_loss; ws.cell(row=r_avg, column=11).alignment = Alignment(horizontal='center')

c_awr = ws.cell(row=r_avg, column=12, value=0.515); c_awr.number_format = '0.00%'; c_awr.font = font_win; c_awr.fill = fill_win; c_awr.alignment = Alignment(horizontal='center')

for c in range(1, 13): ws.cell(row=r_avg, column=c).border = border_double_bottom

for col_idx in range(1, 13):
    col_letter = get_column_letter(col_idx)
    max_len = max(len(str(ws.cell(row=r, column=col_idx).value or '')) for r in range(1, 19))
    ws.column_dimensions[col_letter].width = max(max_len + 5, 17)

wb.save(OUTPUT_NO_REUSE_MASTER)
print(f"🎉 FUND REUSE VS NO-REUSE COMPARISON MASTER CREATED IN {time.time() - t0:.1f}s!", flush=True)
