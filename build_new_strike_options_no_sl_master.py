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

# ------------------- Configuration & Paths -------------------
BASE_DIR = pathlib.Path('D:/behaviour analysis')
SOURCE_PIT_FUT = BASE_DIR / 'Nifty50_12_Quarters_Point_In_Time_Futures_Master.xlsx'
OUTPUT_NEW_STRIKE_MASTER  = BASE_DIR / 'Nifty50_12_Quarters_Options_New_Strike_No_SL_Master.xlsx'
OUTPUT_NEW_STRIKE_SUMMARY = BASE_DIR / 'Nifty50_12_Quarters_Options_New_Strike_No_SL_Summary.xlsx'

# Executive Color Palette (Royal Blue / Dark Slate Header)
COLOR_TITLE_BG     = '1E3A8A'  # Royal Blue Title Banner
COLOR_SECTION_BG   = '1E293B'  # Slate Navy Section Header
COLOR_HEADER_BG    = '0F172A'  # Deep Slate Table Header
COLOR_SUBHEADER_BG = '3B82F6'  # Accent Blue Subheader
COLOR_ZEBRA_BG     = 'F8FAFC'  # Light Gray Zebra
COLOR_CARD_BG      = 'F1F5F9'  # Cool Gray Card Fill

COLOR_WIN_BG       = 'DCFCE7'  # Soft Green Fill
COLOR_WIN_TEXT     = '15803D'  # Dark Green Text
COLOR_LOSS_BG      = 'FEE2E2'  # Soft Red Fill
COLOR_LOSS_TEXT    = 'B91C1C'  # Dark Red Text
COLOR_GOLD_BG      = 'FEF3C7'  # Soft Gold Fill
COLOR_GOLD_TEXT    = 'B45309'  # Dark Gold Text

COLOR_WHITE        = 'FFFFFF'
COLOR_TEXT_MAIN    = '334155'  # Dark Slate Text

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
    top=Side(style='thin', color='1E293B'),
    bottom=Side(style='double', color='1E293B')
)

def get_nse_strike_interval(price):
    if price > 5000: return 50
    elif price > 1000: return 20
    elif price > 500: return 10
    else: return 5

def round_nse_strike(price):
    interval = get_nse_strike_interval(price)
    return round(price / interval) * interval

quarters_order = [
    'Q3 2023-24', 'Q4 2023-24', 'Q1 2024-25', 'Q2 2024-25',
    'Q3 2024-25', 'Q4 2024-25', 'Q1 2025-26', 'Q2 2025-26',
    'Q3 2025-26', 'Q4 2025-26', 'Q1 2026-27', 'Q2 2026-27'
]

print("Loading Source Data for New Strike Selection Model (No SL)...", flush=True)
xl_pit = pd.ExcelFile(SOURCE_PIT_FUT, engine='openpyxl')

new_strike_trades = []

for qtr in quarters_order:
    q_tr = xl_pit.parse(qtr, header=6)
    
    t_no_col = next(c for c in q_tr.columns if 'Trade' in str(c))
    q_tr['Entry Date DT'] = pd.to_datetime(q_tr['Entry Date'])
    q_tr = q_tr.sort_values(by=['Entry Date DT', t_no_col]).reset_index(drop=True)
    
    for idx, row in q_tr.iterrows():
        sym      = str(row['Symbol']).strip()
        lot_size = int(row['Lot Size (Qty)'])
        s_entry  = float(row['Entry Futures Price (₹)'])
        s_exit   = float(row['Exit Futures Price (₹)'])
        strat    = str(row['Strategy']).strip().upper()
        
        is_call  = 'LONG' in strat or 'BUY' in strat
        opt_type = '1 STRIKE ITM CALL (CE)' if is_call else '1 STRIKE OTM PUT (PE)'
        
        interval = get_nse_strike_interval(s_entry)
        k_atm = round_nse_strike(s_entry)
        
        # USER NEW STRIKE RULE:
        # CALL: 1 Strike BELOW ATM (In-The-Money Call: K_CE = K_ATM - Interval)
        # PUT : 1 Strike BELOW ATM (Out-of-The-Money Put: K_PE = K_ATM - Interval)
        strike = k_atm - interval
        
        if is_call:
            intrinsic_entry = max(0.0, s_entry - strike)
            intrinsic_exit  = max(0.0, s_exit - strike)
        else:
            intrinsic_entry = max(0.0, strike - s_entry)
            intrinsic_exit  = max(0.0, strike - s_exit)
            
        extrinsic_entry = s_entry * 0.02
        p_entry = intrinsic_entry + extrinsic_entry
        
        extrinsic_exit = s_exit * 0.005
        p_exit = intrinsic_exit + extrinsic_exit
        
        opt_pnl = round((p_exit - p_entry) * lot_size, 2)
        capital_outlay = round(p_entry * lot_size, 2)
        opt_return = round((p_exit - p_entry) / p_entry, 4) if p_entry > 0 else 0.0
        
        trade_rec = {
            'Quarter': qtr,
            'FY': str(row.get('FY', '')),
            'Trade No': idx + 1,
            'Symbol': sym,
            'Company Name': str(row.get('Company Name', '')),
            'Sector': str(row.get('Sector', '')),
            'Option Strategy': opt_type,
            'Position Taking Window': str(row.get('Position Taking Window', '')),
            'Entry Date': str(row['Entry Date']),
            'Exit Date': str(row['Exit Date']),
            'Entry Spot Price (₹)': s_entry,
            'Exit Spot Price (₹)': s_exit,
            'ATM Strike Price (₹)': k_atm,
            'Selected Strike Price (₹)': strike,
            'Option Entry Premium (₹)': round(p_entry, 2),
            'Option Exit Premium (₹)': round(p_exit, 2),
            'Lot Size (Qty)': lot_size,
            'Capital Outlay per Trade (₹)': capital_outlay,
            'Booked Option P&L (₹)': opt_pnl,
            'Option Return (%)': opt_return,
            'Assigned Slot': str(row.get('Assigned Slot', f"Slot {(idx%15)+1}")),
            'Re-entry Type': str(row.get('Re-entry Type', 'First Entry')),
            'Quarterly Result Date': str(row.get('Quarterly Result Date', ''))
        }
        new_strike_trades.append(trade_rec)

df_all_new_trades = pd.DataFrame(new_strike_trades)

# ------------------- 2. Calculate Day-by-Day Peak Capital & Dual Drawdowns (No SL) -------------------
new_summaries = []

for qtr in quarters_order:
    q_tr = df_all_new_trades[df_all_new_trades['Quarter'] == qtr].copy()
    q_tr['Entry Date DT'] = pd.to_datetime(q_tr['Entry Date'])
    q_tr['Exit Date DT']  = pd.to_datetime(q_tr['Exit Date'])
    q_tr = q_tr.sort_values(by=['Entry Date DT', 'Trade No']).reset_index(drop=True)
    
    t_count = len(q_tr)
    wins = len(q_tr[q_tr['Booked Option P&L (₹)'] > 0])
    losses = len(q_tr[q_tr['Booked Option P&L (₹)'] <= 0])
    win_rate = (wins / t_count) if t_count > 0 else 0.0
    tot_opt_pnl = round(q_tr['Booked Option P&L (₹)'].sum(), 2)
    
    # Peak Concurrent Capital
    min_date = q_tr['Entry Date DT'].min()
    max_date = q_tr['Exit Date DT'].max()
    date_range = pd.date_range(min_date, max_date)
    
    peak_opt_capital = 0.0
    peak_opt_date    = None
    peak_opt_open_trades = 0
    
    for d in date_range:
        active_trades = q_tr[(q_tr['Entry Date DT'] <= d) & (d <= q_tr['Exit Date DT'])]
        opt_tot = active_trades['Capital Outlay per Trade (₹)'].sum()
        if opt_tot > peak_opt_capital:
            peak_opt_capital = opt_tot
            peak_opt_date = d.strftime('%Y-%m-%d')
            peak_opt_open_trades = len(active_trades)
            
    final_val_quarter = round(peak_opt_capital + tot_opt_pnl, 2)
    opt_booked_return = (tot_opt_pnl / peak_opt_capital) if peak_opt_capital > 0 else 0.0
    
    # Chronological Cumulative & Peak Drawdown
    q_tr['Cum PnL'] = q_tr['Booked Option P&L (₹)'].cumsum()
    q_tr['Running Max'] = np.maximum.accumulate(q_tr['Cum PnL'])
    q_tr['Drawdown'] = q_tr['Cum PnL'] - q_tr['Running Max']
    
    trough_idx = q_tr['Drawdown'].idxmin()
    max_dd_val = abs(q_tr.loc[trough_idx, 'Drawdown'])
    peak_cum_pnl = q_tr.loc[:trough_idx, 'Cum PnL'].max()
    
    dd_pct_initial = -abs(max_dd_val / peak_opt_capital) if peak_opt_capital > 0 else 0.0
    peak_portfolio_val = peak_opt_capital + max(0.0, peak_cum_pnl)
    dd_pct_peak = -abs(max_dd_val / peak_portfolio_val) if peak_portfolio_val > 0 else 0.0
    
    new_summaries.append({
        'Quarter': qtr,
        'FY': q_tr.iloc[0]['FY'] if 'FY' in q_tr.columns else '',
        'Total Trades': t_count,
        'Wins': wins,
        'Losses': losses,
        'Win Rate (%)': win_rate,
        'Peak Concurrent Capital (₹)': round(peak_opt_capital, 2),
        'Peak Capital Date': peak_opt_date,
        'Peak Open Trades': peak_opt_open_trades,
        'Net Option P&L (₹)': tot_opt_pnl,
        'Final Quarter Value (₹)': final_val_quarter,
        'Booked Return (%)': opt_booked_return,
        'Peak Cum P&L (₹)': round(peak_cum_pnl, 2),
        'Trough Cum P&L (₹)': round(q_tr.loc[trough_idx, 'Cum PnL'], 2),
        'Max Drawdown (₹)': round(max_dd_val, 2),
        'DD % on Initial Capital': round(dd_pct_initial, 4),
        'True DD % on Peak Portfolio': round(dd_pct_peak, 4),
        'Trough Stock': q_tr.loc[trough_idx, 'Symbol']
    })

df_new_summary = pd.DataFrame(new_summaries)

# ------------------- 3. Build Master 13-Sheet Excel Workbook (New Strike, No SL) -------------------
print("Building Master 13-Sheet Workbook for New Strike Selection Model (No SL)...", flush=True)
wb = Workbook()
wb.remove(wb.active)

def style_table_headers(ws, start_row, headers, fill_style=fill_header):
    for c_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=start_row, column=c_idx, value=h)
        cell.font = font_header; cell.fill = fill_style
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border_thin

def autofit_columns(ws, max_cols=25):
    ws.views.sheetView[0].showGridLines = True
    for col_idx in range(1, max_cols + 1):
        col_letter = get_column_letter(col_idx)
        max_len = 0
        for row in range(1, min(ws.max_row + 1, 100)):
            val = str(ws.cell(row=row, column=col_idx).value or '')
            if len(val) > max_len: max_len = len(val)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 13)

# --- SHEET 1: Exec_12Q_Combined_Summary ---
ws_exec = wb.create_sheet(title='Exec_12Q_Combined_Summary')

ws_exec.merge_cells('A1:I1')
c_t = ws_exec['A1']
c_t.value = "NIFTY 50 OPTIONS STRATEGY — NEW STRIKE SELECTION MODEL (NO STOP-LOSS) EXECUTIVE SCORECARD"
c_t.font = font_title; c_t.fill = fill_title; c_t.alignment = Alignment(horizontal='center', vertical='center')
ws_exec.row_dimensions[1].height = 42

global_tot_pnl = round(df_all_new_trades['Booked Option P&L (₹)'].sum(), 2)
global_win_rate = len(df_all_new_trades[df_all_new_trades['Booked Option P&L (₹)'] > 0]) / len(df_all_new_trades)

ws_exec.merge_cells('A3:B3')
c_h1 = ws_exec['A3']; c_h1.value = "TOTAL TRADES ANALYZED"; c_h1.font = font_card_lbl; c_h1.fill = fill_subheader; c_h1.alignment = Alignment(horizontal='center', vertical='center')
ws_exec.merge_cells('A4:B4')
c_v1 = ws_exec['A4']; c_v1.value = len(df_all_new_trades); c_v1.font = Font(name='Segoe UI', size=16, bold=True, color=COLOR_TEXT_MAIN); c_v1.fill = fill_card; c_v1.alignment = Alignment(horizontal='center', vertical='center')

ws_exec.merge_cells('C3:E3')
c_h2 = ws_exec['C3']; c_h2.value = "GLOBAL WIN RATE"; c_h2.font = font_card_lbl; c_h2.fill = fill_subheader; c_h2.alignment = Alignment(horizontal='center', vertical='center')
ws_exec.merge_cells('C4:E4')
c_v2 = ws_exec['C4']; c_v2.value = global_win_rate; c_v2.font = Font(name='Segoe UI', size=16, bold=True, color=COLOR_WIN_TEXT); c_v2.fill = fill_win; c_v2.number_format = '0.0%'; c_v2.alignment = Alignment(horizontal='center', vertical='center')

ws_exec.merge_cells('F3:I3')
c_h3 = ws_exec['F3']; c_h3.value = "TOTAL 12-QUARTER NET OPTION PROFIT (NO SL)"; c_h3.font = font_card_lbl; c_h3.fill = fill_subheader; c_h3.alignment = Alignment(horizontal='center', vertical='center')
ws_exec.merge_cells('F4:I4')
c_v3 = ws_exec['F4']; c_v3.value = global_tot_pnl; c_v3.font = Font(name='Segoe UI', size=16, bold=True, color=COLOR_WIN_TEXT); c_v3.fill = fill_win; c_v3.number_format = '₹#,##0.00'; c_v3.alignment = Alignment(horizontal='center', vertical='center')

ws_exec.row_dimensions[3].height = 22
ws_exec.row_dimensions[4].height = 30

for r in [3, 4]:
    for c in range(1, 10): ws_exec.cell(row=r, column=c).border = border_thin

ws_exec.merge_cells('A6:I6')
c_sec = ws_exec['A6']
c_sec.value = "CONSOLIDATED 12-QUARTER PERFORMANCE SCORECARD (CALL: 1 STRIKE ITM | PUT: 1 STRIKE OTM | NO SL)"
c_sec.font = font_section; c_sec.fill = fill_section; c_sec.alignment = Alignment(horizontal='left', vertical='center')
ws_exec.row_dimensions[6].height = 26

headers_summary = ['Quarter', 'Peak Concurrent Capital (₹)', 'Peak Date (Open Trades)', 'Net Option Realised P&L (₹)', 'Return on Peak Capital (%)', 'Max Realised Drawdown (₹)', 'DD % on Initial Capital', 'True DD % on Peak Portfolio', 'Win Rate (%)']
style_table_headers(ws_exec, 7, headers_summary, fill_header)
ws_exec.row_dimensions[7].height = 28

for r_idx, row_data in df_new_summary.iterrows():
    row_num = 8 + r_idx
    row_fill = fill_zebra if r_idx % 2 == 1 else PatternFill(fill_type=None)
    ws_exec.row_dimensions[row_num].height = 22
    
    ws_exec.cell(row=row_num, column=1, value=row_data['Quarter']).alignment = Alignment(horizontal='center')
    ws_exec.cell(row=row_num, column=2, value=row_data['Peak Concurrent Capital (₹)']).number_format = '₹#,##0.00'
    ws_exec.cell(row=row_num, column=3, value=f"{row_data['Peak Capital Date']} ({row_data['Peak Open Trades']} Trades)").alignment = Alignment(horizontal='center')
    
    c_pnl = ws_exec.cell(row=row_num, column=4, value=row_data['Net Option P&L (₹)'])
    c_pnl.number_format = '₹#,##0.00'; c_pnl.fill = fill_win if row_data['Net Option P&L (₹)'] >= 0 else fill_loss
    c_pnl.font = font_win if row_data['Net Option P&L (₹)'] >= 0 else font_loss
    
    c_ret = ws_exec.cell(row=row_num, column=5, value=row_data['Booked Return (%)'])
    c_ret.number_format = '+0.00%;-0.00%'
    
    c_dd = ws_exec.cell(row=row_num, column=6, value=-abs(row_data['Max Drawdown (₹)']))
    c_dd.number_format = '-₹#,##0.00'; c_dd.font = font_loss; c_dd.fill = fill_loss
    
    c_ddi = ws_exec.cell(row=row_num, column=7, value=row_data['DD % on Initial Capital'])
    c_ddi.number_format = '-0.00%'; c_ddi.font = font_loss; c_ddi.fill = fill_loss; c_ddi.alignment = Alignment(horizontal='center')
    
    c_ddp = ws_exec.cell(row=row_num, column=8, value=row_data['True DD % on Peak Portfolio'])
    c_ddp.number_format = '-0.00%'; c_ddp.font = font_loss; c_ddp.fill = fill_loss; c_ddp.alignment = Alignment(horizontal='center')
    
    c_wr = ws_exec.cell(row=row_num, column=9, value=row_data['Win Rate (%)'])
    c_wr.number_format = '0.00%'; c_wr.alignment = Alignment(horizontal='center')
    if row_data['Win Rate (%)'] >= 0.70: c_wr.fill = fill_win; c_wr.font = font_win
    elif row_data['Win Rate (%)'] >= 0.50: c_wr.fill = fill_gold; c_wr.font = font_gold
    else: c_wr.fill = fill_loss; c_wr.font = font_loss

    for c in range(1, 10):
        cell = ws_exec.cell(row=row_num, column=c)
        cell.border = border_thin
        if c not in [4, 6, 7, 8, 9] and row_fill.fill_type: cell.fill = row_fill

# Row 20: 12-QUARTER TOTAL ROW
r_tot = 20
ws_exec.row_dimensions[r_tot].height = 25
ws_exec.cell(row=r_tot, column=1, value="12-QUARTER TOTAL").font = font_data_bold
ws_exec.cell(row=r_tot, column=2, value="—").alignment = Alignment(horizontal='center')
ws_exec.cell(row=r_tot, column=3, value="—").alignment = Alignment(horizontal='center')
ws_exec.cell(row=r_tot, column=4, value=global_tot_pnl).number_format = '₹#,##0.00'
ws_exec.cell(row=r_tot, column=4).font = font_win; ws_exec.cell(row=r_tot, column=4).fill = fill_win
ws_exec.cell(row=r_tot, column=5, value="—").alignment = Alignment(horizontal='center')
ws_exec.cell(row=r_tot, column=6, value=-abs(df_new_summary['Max Drawdown (₹)'].max())).number_format = '-₹#,##0.00'
ws_exec.cell(row=r_tot, column=6).font = font_loss; ws_exec.cell(row=r_tot, column=6).fill = fill_loss
ws_exec.cell(row=r_tot, column=7, value=df_new_summary['DD % on Initial Capital'].min()).number_format = '-0.00%'
ws_exec.cell(row=r_tot, column=7).font = font_loss; ws_exec.cell(row=r_tot, column=7).fill = fill_loss; ws_exec.cell(row=r_tot, column=7).alignment = Alignment(horizontal='center')
ws_exec.cell(row=r_tot, column=8, value=df_new_summary['True DD % on Peak Portfolio'].min()).number_format = '-0.00%'
ws_exec.cell(row=r_tot, column=8).font = font_loss; ws_exec.cell(row=r_tot, column=8).fill = fill_loss; ws_exec.cell(row=r_tot, column=8).alignment = Alignment(horizontal='center')
c_tot_wr = ws_exec.cell(row=r_tot, column=9, value=global_win_rate)
c_tot_wr.number_format = '0.00%'; c_tot_wr.font = font_win; c_tot_wr.fill = fill_win; c_tot_wr.alignment = Alignment(horizontal='center')

for c in range(1, 10): ws_exec.cell(row=r_tot, column=c).border = border_thin

# Row 21: QUARTERLY AVERAGE ROW
r_avg = 21
ws_exec.row_dimensions[r_avg].height = 25
ws_exec.cell(row=r_avg, column=1, value="QUARTERLY AVERAGE").font = font_data_bold
ws_exec.cell(row=r_avg, column=2, value=df_new_summary['Peak Concurrent Capital (₹)'].mean()).number_format = '₹#,##0.00'
ws_exec.cell(row=r_avg, column=3, value="Avg ~17 Trades").alignment = Alignment(horizontal='center')
ws_exec.cell(row=r_avg, column=4, value=df_new_summary['Net Option P&L (₹)'].mean()).number_format = '₹#,##0.00'
ws_exec.cell(row=r_avg, column=4).font = font_win; ws_exec.cell(row=r_avg, column=4).fill = fill_win
ws_exec.cell(row=r_avg, column=5, value=df_new_summary['Booked Return (%)'].mean()).number_format = '+0.00%'
ws_exec.cell(row=r_avg, column=6, value=-abs(df_new_summary['Max Drawdown (₹)'].mean())).number_format = '-₹#,##0.00'
ws_exec.cell(row=r_avg, column=6).font = font_loss; ws_exec.cell(row=r_avg, column=6).fill = fill_loss
ws_exec.cell(row=r_avg, column=7, value=df_new_summary['DD % on Initial Capital'].mean()).number_format = '-0.00%'
ws_exec.cell(row=r_avg, column=7).font = font_loss; ws_exec.cell(row=r_avg, column=7).fill = fill_loss; ws_exec.cell(row=r_avg, column=7).alignment = Alignment(horizontal='center')
ws_exec.cell(row=r_avg, column=8, value=df_new_summary['True DD % on Peak Portfolio'].mean()).number_format = '-0.00%'
ws_exec.cell(row=r_avg, column=8).font = font_loss; ws_exec.cell(row=r_avg, column=8).fill = fill_loss; ws_exec.cell(row=r_avg, column=8).alignment = Alignment(horizontal='center')
c_avg_wr = ws_exec.cell(row=r_avg, column=9, value=global_win_rate)
c_avg_wr.number_format = '0.00%'; c_avg_wr.font = font_win; c_avg_wr.fill = fill_win; c_avg_wr.alignment = Alignment(horizontal='center')

for c in range(1, 10): ws_exec.cell(row=r_avg, column=c).border = border_double_bottom

autofit_columns(ws_exec, max_cols=9)

# Chart
chart_opt = LineChart()
chart_opt.title = "Master 12-Quarter Net Option Buying P&L (New Strike Selection - No SL) (₹)"
chart_opt.style = 13
chart_opt.y_axis.title = "Net Option P&L (₹)"
chart_opt.x_axis.title = "Quarter"
chart_opt.width = 16
chart_opt.height = 11

data_o = Reference(ws_exec, min_col=4, min_row=7, max_row=19)
cats_o = Reference(ws_exec, min_col=1, min_row=8, max_row=19)
chart_opt.add_data(data_o, titles_from_data=True)
chart_opt.set_categories(cats_o)
ws_exec.add_chart(chart_opt, "K3")

# --- SHEETS 2-13: INDIVIDUAL 12 QUARTER SHEETS ---
trade_headers = [
    'Trade #', 'Symbol', 'Company Name', 'Sector', 'Option Strategy',
    'Position Taking Window', 'Entry Date', 'Exit Date',
    'Entry Spot Price (₹)', 'Exit Spot Price (₹)', 'ATM Strike Price (₹)', 'Selected Strike Price (₹)',
    'Option Entry Premium (₹)', 'Option Exit Premium (₹)', 'Lot Size (Qty)',
    'Capital Outlay per Trade (₹)', 'Booked Option P&L (₹)', 'Cumulative P&L (₹)', 'Drawdown (₹)',
    'Assigned Slot', 'Re-entry Type', 'Quarterly Result Date'
]

for qtr in quarters_order:
    ws_q = wb.create_sheet(title=qtr)
    
    q_tr = df_all_new_trades[df_all_new_trades['Quarter'] == qtr].copy()
    q_tr['Entry Date DT'] = pd.to_datetime(q_tr['Entry Date'])
    q_tr = q_tr.sort_values(by=['Entry Date DT', 'Trade No']).reset_index(drop=True)
    
    q_sum = df_new_summary[df_new_summary['Quarter'] == qtr].iloc[0]
    
    q_tr['Cumulative P&L (₹)'] = q_tr['Booked Option P&L (₹)'].cumsum()
    q_tr['Running Max'] = np.maximum.accumulate(q_tr['Cumulative P&L (₹)'])
    q_tr['Drawdown (₹)'] = q_tr['Cumulative P&L (₹)'] - q_tr['Running Max']
    
    ws_q.merge_cells('A1:V1')
    c_t = ws_q['A1']
    c_t.value = f"NIFTY 50 OPTIONS BUYING MODEL (CALL: 1 STRIKE ITM | PUT: 1 STRIKE OTM | NO SL) — {qtr.upper()}"
    c_t.font = font_title; c_t.fill = fill_title; c_t.alignment = Alignment(horizontal='center', vertical='center')
    ws_q.row_dimensions[1].height = 40
    
    card_lbls = ['Peak Concurrent Capital (₹)', 'Peak Date (Open Trades)', 'Net Option P&L (₹)', 'Return on Peak Capital (%)', 'Win Rate (%)', 'Total Trades', 'Winning Trades', 'Losing Trades', 'Max Drawdown (₹)', 'DD % Initial Capital', 'True DD % Peak Portfolio']
    ws_q.merge_cells('A2:K2')
    c_s = ws_q['A2']; c_s.value = f"{qtr.upper()} — PERFORMANCE METRICS & CAPITAL OVERVIEW (NO SL)"; c_s.font = font_section; c_s.fill = fill_section
    ws_q.row_dimensions[2].height = 24
    
    style_table_headers(ws_q, 3, card_lbls, fill_subheader)
    ws_q.row_dimensions[3].height = 22
    
    ws_q.cell(row=4, column=1, value=q_sum['Peak Concurrent Capital (₹)']).number_format = '₹#,##0.00'
    ws_q.cell(row=4, column=2, value=f"{q_sum['Peak Capital Date']} ({q_sum['Peak Open Trades']} Trades)").alignment = Alignment(horizontal='center')
    c_pnl = ws_q.cell(row=4, column=3, value=q_sum['Net Option P&L (₹)'])
    c_pnl.number_format = '₹#,##0.00'; c_pnl.font = font_win if q_sum['Net Option P&L (₹)'] >= 0 else font_loss; c_pnl.fill = fill_win if q_sum['Net Option P&L (₹)'] >= 0 else fill_loss
    
    ws_q.cell(row=4, column=4, value=q_sum['Booked Return (%)']).number_format = '0.00%'
    c_wr = ws_q.cell(row=4, column=5, value=q_sum['Win Rate (%)'])
    c_wr.number_format = '0.00%'; c_wr.font = font_win; c_wr.fill = fill_win; c_wr.alignment = Alignment(horizontal='center')
    
    ws_q.cell(row=4, column=6, value=int(q_sum['Total Trades'])).alignment = Alignment(horizontal='center')
    ws_q.cell(row=4, column=7, value=int(q_sum['Wins'])).alignment = Alignment(horizontal='center')
    ws_q.cell(row=4, column=8, value=int(q_sum['Losses'])).alignment = Alignment(horizontal='center')
    
    c_dd1 = ws_q.cell(row=4, column=9, value=-abs(q_sum['Max Drawdown (₹)']))
    c_dd1.number_format = '-₹#,##0.00'; c_dd1.font = font_loss; c_dd1.fill = fill_loss
    
    c_ddi = ws_q.cell(row=4, column=10, value=q_sum['DD % on Initial Capital'])
    c_ddi.number_format = '-0.00%'; c_ddi.font = font_loss; c_ddi.fill = fill_loss; c_ddi.alignment = Alignment(horizontal='center')
    
    c_dd2 = ws_q.cell(row=4, column=11, value=q_sum['True DD % on Peak Portfolio'])
    c_dd2.number_format = '-0.00%'; c_dd2.font = font_loss; c_dd2.fill = fill_loss; c_dd2.alignment = Alignment(horizontal='center')
    
    ws_q.row_dimensions[4].height = 26
    
    for c in range(1, 12): ws_q.cell(row=4, column=c).border = border_thin
    
    ws_q.merge_cells('A6:V6')
    c_s2 = ws_q['A6']; c_s2.value = f"{qtr.upper()} — COMPLETE DETAILED 50-TRADE OPTIONS JOURNAL (NO SL | ENTRY DATE CHRONOLOGICAL ORDER)"; c_s2.font = font_section; c_s2.fill = fill_section
    ws_q.row_dimensions[6].height = 24
    
    style_table_headers(ws_q, 7, trade_headers, fill_header)
    ws_q.row_dimensions[7].height = 26
    
    for r_idx, t_row in q_tr.iterrows():
        row_num = 8 + r_idx
        row_fill = fill_zebra if r_idx % 2 == 1 else PatternFill(fill_type=None)
        ws_q.row_dimensions[row_num].height = 21
        
        ws_q.cell(row=row_num, column=1, value=r_idx + 1).alignment = Alignment(horizontal='center')
        ws_q.cell(row=row_num, column=2, value=t_row['Symbol']).alignment = Alignment(horizontal='center')
        ws_q.cell(row=row_num, column=3, value=t_row['Company Name'])
        ws_q.cell(row=row_num, column=4, value=t_row['Sector'])
        
        c_st = ws_q.cell(row=row_num, column=5, value=t_row['Option Strategy'])
        c_st.alignment = Alignment(horizontal='center')
        if 'CALL' in t_row['Option Strategy']: c_st.fill = fill_win; c_st.font = font_win
        else: c_st.fill = fill_loss; c_st.font = font_loss
        
        ws_q.cell(row=row_num, column=6, value=t_row['Position Taking Window']).alignment = Alignment(horizontal='center')
        ws_q.cell(row=row_num, column=7, value=t_row['Entry Date']).alignment = Alignment(horizontal='center')
        ws_q.cell(row=row_num, column=8, value=t_row['Exit Date']).alignment = Alignment(horizontal='center')
        
        ws_q.cell(row=row_num, column=9, value=t_row['Entry Spot Price (₹)']).number_format = '₹#,##0.00'
        ws_q.cell(row=row_num, column=10, value=t_row['Exit Spot Price (₹)']).number_format = '₹#,##0.00'
        ws_q.cell(row=row_num, column=11, value=t_row['ATM Strike Price (₹)']).number_format = '₹#,##0.00'
        ws_q.cell(row=row_num, column=12, value=t_row['Selected Strike Price (₹)']).number_format = '₹#,##0.00'
        
        ws_q.cell(row=row_num, column=13, value=t_row['Option Entry Premium (₹)']).number_format = '₹#,##0.00'
        ws_q.cell(row=row_num, column=14, value=t_row['Option Exit Premium (₹)']).number_format = '₹#,##0.00'
        ws_q.cell(row=row_num, column=15, value=int(t_row['Lot Size (Qty)'])).alignment = Alignment(horizontal='center')
        ws_q.cell(row=row_num, column=16, value=t_row['Capital Outlay per Trade (₹)']).number_format = '₹#,##0.00'
        
        c_op = ws_q.cell(row=row_num, column=17, value=t_row['Booked Option P&L (₹)'])
        c_op.number_format = '₹#,##0.00'
        c_op.font = font_win if t_row['Booked Option P&L (₹)'] >= 0 else font_loss
        c_op.fill = fill_win if t_row['Booked Option P&L (₹)'] >= 0 else fill_loss
        
        c_cp = ws_q.cell(row=row_num, column=18, value=t_row['Cumulative P&L (₹)'])
        c_cp.number_format = '₹#,##0.00'; c_cp.font = font_data_bold
        
        c_dd = ws_q.cell(row=row_num, column=19, value=t_row['Drawdown (₹)'])
        c_dd.number_format = '₹#,##0.00'
        if t_row['Drawdown (₹)'] < 0: c_dd.font = font_loss; c_dd.fill = fill_loss
        
        ws_q.cell(row=row_num, column=20, value=t_row['Assigned Slot']).alignment = Alignment(horizontal='center')
        ws_q.cell(row=row_num, column=21, value=t_row.get('Re-entry Type', 'First Entry')).alignment = Alignment(horizontal='center')
        ws_q.cell(row=row_num, column=22, value=t_row.get('Quarterly Result Date', '')).alignment = Alignment(horizontal='center')
        
        for c in range(1, 23):
            cell = ws_q.cell(row=row_num, column=c)
            cell.border = border_thin
            if c not in [5, 17, 19] and row_fill.fill_type: cell.fill = row_fill

    autofit_columns(ws_q, max_cols=22)
    ws_q.freeze_panes = 'A8'

    chart_q = LineChart()
    chart_q.title = f"{qtr} — Option Buying (New Strike Selection - No SL) Cumulative P&L Curve (Max DD: -₹{q_sum['Max Drawdown (₹)']:,.2f})"
    chart_q.style = 13
    chart_q.y_axis.title = "Rupees (₹)"
    chart_q.x_axis.title = "Trade #"
    chart_q.width = 16
    chart_q.height = 10

    max_r = 7 + len(q_tr)
    data_q = Reference(ws_q, min_col=18, min_row=7, max_col=19, max_row=max_r)
    cats_q = Reference(ws_q, min_col=1, min_row=8, max_row=max_r)
    chart_q.add_data(data_q, titles_from_data=True)
    chart_q.set_categories(cats_q)
    ws_q.add_chart(chart_q, "X6")

wb.save(OUTPUT_NEW_STRIKE_MASTER)

# ------------------- 4. Build Standalone Summary Workbook (New Strike, No SL) -------------------
print("Writing Standalone Summary Excel Workbook (New Strike, No SL)...", flush=True)
wb_s = openpyxl.Workbook()
wb_s.remove(wb_s.active)

ws_s = wb_s.create_sheet(title='Options_New_Strike_No_SL_Sum')
ws_s.views.sheetView[0].showGridLines = True

ws_s.merge_cells('A1:I1')
c_t = ws_s['A1']
c_t.value = "NIFTY 50 OPTIONS STRATEGY — NEW STRIKE SELECTION MODEL (NO STOP-LOSS) SUMMARY DASHBOARD"
c_t.font = font_title; c_t.fill = fill_title; c_t.alignment = Alignment(horizontal='center', vertical='center')
ws_s.row_dimensions[1].height = 40

ws_s.merge_cells('A3:B3')
c_h1 = ws_s['A3']; c_h1.value = "TOTAL TRADES ANALYZED"; c_h1.font = font_card_lbl; c_h1.fill = fill_subheader; c_h1.alignment = Alignment(horizontal='center', vertical='center')
ws_s.merge_cells('A4:B4')
c_v1 = ws_s['A4']; c_v1.value = len(df_all_new_trades); c_v1.font = Font(name='Segoe UI', size=16, bold=True, color=COLOR_TEXT_MAIN); c_v1.fill = fill_card; c_v1.alignment = Alignment(horizontal='center', vertical='center')

ws_s.merge_cells('C3:E3')
c_h2 = ws_s['C3']; c_h2.value = "GLOBAL WIN RATE"; c_h2.font = font_card_lbl; c_h2.fill = fill_subheader; c_h2.alignment = Alignment(horizontal='center', vertical='center')
ws_s.merge_cells('C4:E4')
c_v2 = ws_s['C4']; c_v2.value = global_win_rate; c_v2.font = Font(name='Segoe UI', size=16, bold=True, color=COLOR_WIN_TEXT); c_v2.fill = fill_win; c_v2.number_format = '0.0%'; c_v2.alignment = Alignment(horizontal='center', vertical='center')

ws_s.merge_cells('F3:I3')
c_h3 = ws_s['F3']; c_h3.value = "TOTAL 12-QUARTER NET OPTION PROFIT (NO SL)"; c_h3.font = font_card_lbl; c_h3.fill = fill_subheader; c_h3.alignment = Alignment(horizontal='center', vertical='center')
ws_s.merge_cells('F4:I4')
c_v3 = ws_s['F4']; c_v3.value = global_tot_pnl; c_v3.font = Font(name='Segoe UI', size=16, bold=True, color=COLOR_WIN_TEXT); c_v3.fill = fill_win; c_v3.number_format = '₹#,##0.00'; c_v3.alignment = Alignment(horizontal='center', vertical='center')

ws_s.row_dimensions[3].height = 20
ws_s.row_dimensions[4].height = 28
for r in [3, 4]:
    for c in range(1, 10): ws_s.cell(row=r, column=c).border = border_thin

ws_s.merge_cells('A6:I6')
c_s = ws_s['A6']
c_s.value = "CONSOLIDATED 12-QUARTER PERFORMANCE SCORECARD (CALL: 1 STRIKE ITM | PUT: 1 STRIKE OTM | NO SL)"
c_s.font = font_section; c_s.fill = fill_section; c_s.alignment = Alignment(horizontal='left', vertical='center')
ws_s.row_dimensions[6].height = 24

for c_idx, h in enumerate(headers_summary, start=1):
    cell = ws_s.cell(row=7, column=c_idx, value=h)
    cell.font = font_header; cell.fill = fill_header; cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = border_thin
ws_s.row_dimensions[7].height = 26

for r_idx, row_data in df_new_summary.iterrows():
    row_num = 8 + r_idx
    row_fill = fill_zebra if r_idx % 2 == 1 else PatternFill(fill_type=None)
    ws_s.row_dimensions[row_num].height = 21
    
    ws_s.cell(row=row_num, column=1, value=row_data['Quarter']).alignment = Alignment(horizontal='center')
    ws_s.cell(row=row_num, column=2, value=row_data['Peak Concurrent Capital (₹)']).number_format = '₹#,##0.00'
    ws_s.cell(row=row_num, column=3, value=f"{row_data['Peak Capital Date']} ({row_data['Peak Open Trades']} Trades)").alignment = Alignment(horizontal='center')
    
    c_pnl = ws_s.cell(row=row_num, column=4, value=row_data['Net Option P&L (₹)'])
    c_pnl.number_format = '₹#,##0.00'; c_pnl.fill = fill_win if row_data['Net Option P&L (₹)'] >= 0 else fill_loss
    c_pnl.font = font_win if row_data['Net Option P&L (₹)'] >= 0 else font_loss
    
    ws_s.cell(row=row_num, column=5, value=row_data['Booked Return (%)']).number_format = '+0.00%;-0.00%'
    
    c_dd = ws_s.cell(row=row_num, column=6, value=-abs(row_data['Max Drawdown (₹)']))
    c_dd.number_format = '-₹#,##0.00'; c_dd.font = font_loss; c_dd.fill = fill_loss
    
    c_ddi = ws_s.cell(row=row_num, column=7, value=row_data['DD % on Initial Capital'])
    c_ddi.number_format = '-0.00%'; c_ddi.font = font_loss; c_ddi.fill = fill_loss; c_ddi.alignment = Alignment(horizontal='center')
    
    c_ddp = ws_s.cell(row=row_num, column=8, value=row_data['True DD % on Peak Portfolio'])
    c_ddp.number_format = '-0.00%'; c_ddp.font = font_loss; c_ddp.fill = fill_loss; c_ddp.alignment = Alignment(horizontal='center')
    
    c_wr = ws_s.cell(row=row_num, column=9, value=row_data['Win Rate (%)'])
    c_wr.number_format = '0.00%'; c_wr.alignment = Alignment(horizontal='center')
    if row_data['Win Rate (%)'] >= 0.70: c_wr.fill = fill_win; c_wr.font = font_win
    elif row_data['Win Rate (%)'] >= 0.50: c_wr.fill = fill_gold; c_wr.font = font_gold
    else: c_wr.fill = fill_loss; c_wr.font = font_loss

    for c in range(1, 10):
        cell = ws_s.cell(row=row_num, column=c)
        cell.border = border_thin
        if c not in [4, 6, 7, 8, 9] and row_fill.fill_type: cell.fill = row_fill

# Row 20 Total
r_tot = 20
ws_s.row_dimensions[r_tot].height = 24
ws_s.cell(row=r_tot, column=1, value="12-QUARTER TOTAL").font = font_data_bold
ws_s.cell(row=r_tot, column=2, value="—").alignment = Alignment(horizontal='center')
ws_s.cell(row=r_tot, column=3, value="—").alignment = Alignment(horizontal='center')
ws_s.cell(row=r_tot, column=4, value=global_tot_pnl).number_format = '₹#,##0.00'
ws_s.cell(row=r_tot, column=4).font = font_win; ws_s.cell(row=r_tot, column=4).fill = fill_win
ws_s.cell(row=r_tot, column=5, value="—").alignment = Alignment(horizontal='center')
ws_s.cell(row=r_tot, column=6, value=-abs(df_new_summary['Max Drawdown (₹)'].max())).number_format = '-₹#,##0.00'
ws_s.cell(row=r_tot, column=6).font = font_loss; ws_s.cell(row=r_tot, column=6).fill = fill_loss
ws_s.cell(row=r_tot, column=7, value=df_new_summary['DD % on Initial Capital'].min()).number_format = '-0.00%'
ws_s.cell(row=r_tot, column=7).font = font_loss; ws_s.cell(row=r_tot, column=7).fill = fill_loss; ws_s.cell(row=r_tot, column=7).alignment = Alignment(horizontal='center')
ws_s.cell(row=r_tot, column=8, value=df_new_summary['True DD % on Peak Portfolio'].min()).number_format = '-0.00%'
ws_s.cell(row=r_tot, column=8).font = font_loss; ws_s.cell(row=r_tot, column=8).fill = fill_loss; ws_s.cell(row=r_tot, column=8).alignment = Alignment(horizontal='center')
c_twr = ws_s.cell(row=r_tot, column=9, value=global_win_rate)
c_twr.number_format = '0.00%'; c_twr.font = font_win; c_twr.fill = fill_win; c_twr.alignment = Alignment(horizontal='center')
for c in range(1, 10): ws_s.cell(row=r_tot, column=c).border = border_thin

# Row 21 Average
r_avg = 21
ws_s.row_dimensions[r_avg].height = 24
ws_s.cell(row=r_avg, column=1, value="QUARTERLY AVERAGE").font = font_data_bold
ws_s.cell(row=r_avg, column=2, value=df_new_summary['Peak Concurrent Capital (₹)'].mean()).number_format = '₹#,##0.00'
ws_s.cell(row=r_avg, column=3, value="Avg ~17 Trades").alignment = Alignment(horizontal='center')
ws_s.cell(row=r_avg, column=4, value=df_new_summary['Net Option P&L (₹)'].mean()).number_format = '₹#,##0.00'
ws_s.cell(row=r_avg, column=4).font = font_win; ws_s.cell(row=r_avg, column=4).fill = fill_win
ws_s.cell(row=r_avg, column=5, value=df_new_summary['Booked Return (%)'].mean()).number_format = '+0.00%'
ws_s.cell(row=r_avg, column=6, value=-abs(df_new_summary['Max Drawdown (₹)'].mean())).number_format = '-₹#,##0.00'
ws_s.cell(row=r_avg, column=6).font = font_loss; ws_s.cell(row=r_avg, column=6).fill = fill_loss
ws_s.cell(row=r_avg, column=7, value=df_new_summary['DD % on Initial Capital'].mean()).number_format = '-0.00%'
ws_s.cell(row=r_avg, column=7).font = font_loss; ws_s.cell(row=r_avg, column=7).fill = fill_loss; ws_s.cell(row=r_avg, column=7).alignment = Alignment(horizontal='center')
ws_s.cell(row=r_avg, column=8, value=df_new_summary['True DD % on Peak Portfolio'].mean()).number_format = '-0.00%'
ws_s.cell(row=r_avg, column=8).font = font_loss; ws_s.cell(row=r_avg, column=8).fill = fill_loss; ws_s.cell(row=r_avg, column=8).alignment = Alignment(horizontal='center')
c_awr = ws_s.cell(row=r_avg, column=9, value=global_win_rate)
c_awr.number_format = '0.00%'; c_awr.font = font_win; c_awr.fill = fill_win; c_awr.alignment = Alignment(horizontal='center')
for c in range(1, 10): ws_s.cell(row=r_avg, column=c).border = border_double_bottom

# Autofit
for col_idx in range(1, 10):
    col_letter = get_column_letter(col_idx)
    max_len = max(len(str(ws_s.cell(row=r, column=col_idx).value or '')) for r in range(1, 22))
    ws_s.column_dimensions[col_letter].width = max(max_len + 5, 18)

# Chart
chart = LineChart()
chart.title = "Consolidated 12-Quarter Net Option Buying P&L (New Strike Selection - No SL) (₹)"
chart.style = 13
chart.y_axis.title = "Profit (₹)"
chart.x_axis.title = "Quarter"
chart.width = 16
chart.height = 11

data = Reference(ws_s, min_col=4, min_row=7, max_row=19)
cats = Reference(ws_s, min_col=1, min_row=8, max_row=19)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
ws_s.add_chart(chart, "K3")

wb_s.save(OUTPUT_NEW_STRIKE_SUMMARY)
print(f"\n🎉 NEW STRIKE NO SL OPTIONS MASTER & SUMMARY CREATED IN {time.time() - t0:.1f}s!", flush=True)
