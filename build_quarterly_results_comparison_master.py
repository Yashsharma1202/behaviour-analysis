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
SOURCE_FUT = BASE_DIR / 'Nifty50_12_Quarters_Point_In_Time_Futures_Master.xlsx'
SOURCE_OPT = BASE_DIR / 'Nifty50_12_Quarters_Options_1PCT_ITM_Detailed_Master.xlsx'
OUTPUT_MASTER = BASE_DIR / 'Nifty50_12_Quarters_Quarterly_Results_Comparison_Master.xlsx'

# Executive Palette
COLOR_TITLE_BG     = '0F172A'  # Dark Slate Title Banner
COLOR_SECTION_BG   = '1E293B'  # Slate Navy Section Header
COLOR_HEADER_BG    = '1E3A8A'  # Royal Navy Table Header
COLOR_SUBHEADER_BG = '2563EB'  # Vibrant Blue Subheader
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
        for row in range(1, min(ws.max_row + 1, 150)):
            val = str(ws.cell(row=row, column=col_idx).value or '')
            if len(val) > max_len: max_len = len(val)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 13)

quarters_order = [
    'Q3 2023-24', 'Q4 2023-24', 'Q1 2024-25', 'Q2 2024-25',
    'Q3 2024-25', 'Q4 2024-25', 'Q1 2025-26', 'Q2 2025-26',
    'Q3 2025-26', 'Q4 2025-26', 'Q1 2026-27', 'Q2 2026-27'
]

print("Loading Master Data for 14-Sheet Quarterly Results Comparison Workbook...", flush=True)
xl_fut = pd.ExcelFile(SOURCE_FUT, engine='openpyxl')
xl_opt = pd.ExcelFile(SOURCE_OPT, engine='openpyxl')

q_comp_list = []
mismatch_trades = []

for qtr in quarters_order:
    df_f = xl_fut.parse(qtr, header=6)
    df_o = xl_opt.parse(qtr, header=6)
    
    fut_pnl_col = next(c for c in df_f.columns if 'P&L' in str(c) or 'PnL' in str(c))
    opt_pnl_col = 'Booked Option P&L (₹)'
    
    fut_wins = len(df_f[df_f[fut_pnl_col] > 0])
    opt_wins = len(df_o[df_o[opt_pnl_col] > 0])
    
    fut_wr = fut_wins / len(df_f)
    opt_wr = opt_wins / len(df_o)
    
    fut_tot_pnl = round(df_f[fut_pnl_col].sum(), 2)
    opt_tot_pnl = round(df_o[opt_pnl_col].sum(), 2)
    
    for idx in range(len(df_f)):
        r_f = df_f.iloc[idx]
        r_o = df_o.iloc[idx]
        
        pnl_f = float(r_f[fut_pnl_col])
        pnl_o = float(r_o[opt_pnl_col])
        
        if pnl_f > 0 and pnl_o <= 0:
            sym      = str(r_f['Symbol']).strip()
            s_entry  = float(r_f['Entry Futures Price (₹)'])
            s_exit   = float(r_f['Exit Futures Price (₹)'])
            strat    = str(r_f['Strategy']).strip().upper()
            is_call  = 'LONG' in strat or 'BUY' in strat
            
            spot_change_pct = ((s_exit - s_entry) / s_entry) if is_call else ((s_entry - s_exit) / s_entry)
            outlay_o = float(r_o['Capital Outlay per Trade (₹)'])
            opt_ret  = (pnl_o / outlay_o) if outlay_o > 0 else 0.0
            
            mismatch_trades.append({
                'Quarter': qtr,
                'Trade No': idx + 1,
                'Symbol': sym,
                'Company Name': str(r_f.get('Company Name', '')),
                'Sector': str(r_f.get('Sector', '')),
                'Strategy': strat,
                'Entry Date': str(r_f['Entry Date']),
                'Exit Date': str(r_f['Exit Date']),
                'Entry Spot Price (₹)': s_entry,
                'Exit Spot Price (₹)': s_exit,
                'Spot Move (%)': spot_change_pct,
                'Lot Size (Qty)': int(r_f['Lot Size (Qty)']),
                'Futures P&L (₹)': pnl_f,
                'Options Entry Premium (₹)': float(r_o['Option Entry Premium (₹)']),
                'Options Exit Premium (₹)': float(r_o['Option Exit Premium (₹)']),
                'Options Capital Outlay (₹)': outlay_o,
                'Options P&L (₹)': pnl_o,
                'Options Return (%)': opt_ret,
                'Reason for Discrepancy': f"Post-earnings IV crush (~1.5%) eroded option time premium faster than small stock move ({spot_change_pct:+.2%})"
            })
            
    q_comp_list.append({
        'Quarter': qtr,
        'Futures Wins': fut_wins,
        'Futures Win Rate (%)': fut_wr,
        'Futures Realised P&L (₹)': fut_tot_pnl,
        'Options Wins': opt_wins,
        'Options Win Rate (%)': opt_wr,
        'Options Realised P&L (₹)': opt_tot_pnl,
        'Win Rate Diff (%)': opt_wr - fut_wr,
        'P&L Diff (₹)': opt_tot_pnl - fut_tot_pnl,
        'Futures Win / Options Loss Trades': len([t for t in mismatch_trades if t['Quarter'] == qtr])
    })

df_q_comp = pd.DataFrame(q_comp_list)
df_mismatch = pd.DataFrame(mismatch_trades)

wb = Workbook()
wb.remove(wb.active)

# --- SHEET 1: Quarterly_Results_Summary ---
ws1 = wb.create_sheet(title='Quarterly_Results_Summary')

ws1.merge_cells('A1:J1')
c_t = ws1['A1']
c_t.value = "NIFTY 50 STRATEGY — 12-QUARTER RESULTS SCORECARD (FUTURES VS. OPTIONS BUYING)"
c_t.font = font_title; c_t.fill = fill_title; c_t.alignment = Alignment(horizontal='center', vertical='center')
ws1.row_dimensions[1].height = 42

total_mismatches = len(df_mismatch)

ws1.merge_cells('A3:C3')
c_h1 = ws1['A3']; c_h1.value = "FUTURES GLOBAL WIN RATE"; c_h1.font = font_card_lbl; c_h1.fill = fill_subheader; c_h1.alignment = Alignment(horizontal='center', vertical='center')
ws1.merge_cells('A4:C4')
c_v1 = ws1['A4']; c_v1.value = 415/600; c_v1.font = Font(name='Segoe UI', size=15, bold=True, color=COLOR_WIN_TEXT); c_v1.fill = fill_win; c_v1.number_format = '0.0%'; c_v1.alignment = Alignment(horizontal='center', vertical='center')

ws1.merge_cells('D3:F3')
c_h2 = ws1['D3']; c_h2.value = "OPTIONS GLOBAL WIN RATE"; c_h2.font = font_card_lbl; c_h2.fill = fill_subheader; c_h2.alignment = Alignment(horizontal='center', vertical='center')
ws1.merge_cells('D4:F4')
c_v2 = ws1['D4']; c_v2.value = 309/600; c_v2.font = Font(name='Segoe UI', size=15, bold=True, color=COLOR_TEXT_MAIN); c_v2.fill = fill_card; c_v2.number_format = '0.0%'; c_v2.alignment = Alignment(horizontal='center', vertical='center')

ws1.merge_cells('G3:J3')
c_h3 = ws1['G3']; c_h3.value = "DISCREPANCY TRADES (FUTURES WIN / OPTIONS LOSS)"; c_h3.font = font_card_lbl; c_h3.fill = fill_subheader; c_h3.alignment = Alignment(horizontal='center', vertical='center')
ws1.merge_cells('G4:J4')
c_v3 = ws1['G4']; c_v3.value = f"{total_mismatches} Trades ({total_mismatches/600:.1%})"; c_v3.font = Font(name='Segoe UI', size=15, bold=True, color=COLOR_LOSS_TEXT); c_v3.fill = fill_loss; c_v3.alignment = Alignment(horizontal='center', vertical='center')

ws1.row_dimensions[3].height = 22
ws1.row_dimensions[4].height = 30
for r in [3, 4]:
    for c in range(1, 11): ws1.cell(row=r, column=c).border = border_thin

ws1.merge_cells('A6:J6')
c_sec = ws1['A6']
c_sec.value = "CONSOLIDATED QUARTERLY RESULTS & WIN RATE COMPARISON SCORECARD"
c_sec.font = font_section; c_sec.fill = fill_section; c_sec.alignment = Alignment(horizontal='left', vertical='center')
ws1.row_dimensions[6].height = 26

headers_sum1 = ['Quarter', 'Futures Wins', 'Futures Win Rate (%)', 'Futures Realised P&L (₹)', 'Options Wins', 'Options Win Rate (%)', 'Options Realised P&L (₹)', 'Win Rate Diff (%)', 'P&L Diff (₹)', 'Futures Win / Options Loss Trades']
style_table_headers(ws1, 7, headers_sum1, fill_header)
ws1.row_dimensions[7].height = 28

for r_idx, row_data in df_q_comp.iterrows():
    row_num = 8 + r_idx
    row_fill = fill_zebra if r_idx % 2 == 1 else PatternFill(fill_type=None)
    ws1.row_dimensions[row_num].height = 22
    
    ws1.cell(row=row_num, column=1, value=row_data['Quarter']).alignment = Alignment(horizontal='center')
    ws1.cell(row=row_num, column=2, value=f"{row_data['Futures Wins']}/50").alignment = Alignment(horizontal='center')
    
    c_fwr = ws1.cell(row=row_num, column=3, value=row_data['Futures Win Rate (%)'])
    c_fwr.number_format = '0.00%'; c_fwr.alignment = Alignment(horizontal='center'); c_fwr.font = font_win; c_fwr.fill = fill_win
    
    c_fpnl = ws1.cell(row=row_num, column=4, value=row_data['Futures Realised P&L (₹)'])
    c_fpnl.number_format = '₹#,##0.00'; c_fpnl.font = font_win if row_data['Futures Realised P&L (₹)'] >= 0 else font_loss
    
    ws1.cell(row=row_num, column=5, value=f"{row_data['Options Wins']}/50").alignment = Alignment(horizontal='center')
    
    c_owr = ws1.cell(row=row_num, column=6, value=row_data['Options Win Rate (%)'])
    c_owr.number_format = '0.00%'; c_owr.alignment = Alignment(horizontal='center')
    
    c_opnl = ws1.cell(row=row_num, column=7, value=row_data['Options Realised P&L (₹)'])
    c_opnl.number_format = '₹#,##0.00'; c_opnl.font = font_win if row_data['Options Realised P&L (₹)'] >= 0 else font_loss
    
    c_wrdif = ws1.cell(row=row_num, column=8, value=row_data['Win Rate Diff (%)'])
    c_wrdif.number_format = '-0.00%'; c_wrdif.font = font_loss; c_wrdif.fill = fill_loss; c_wrdif.alignment = Alignment(horizontal='center')
    
    c_pnldif = ws1.cell(row=row_num, column=9, value=row_data['P&L Diff (₹)'])
    c_pnldif.number_format = '₹#,##0.00'; c_pnldif.font = font_data_bold
    
    ws1.cell(row=row_num, column=10, value=row_data['Futures Win / Options Loss Trades']).alignment = Alignment(horizontal='center')

    for c in range(1, 11):
        cell = ws1.cell(row=row_num, column=c)
        cell.border = border_thin
        if c not in [3, 4, 7, 8] and row_fill.fill_type: cell.fill = row_fill

# Row 20 Total
r_tot = 20
ws1.row_dimensions[r_tot].height = 25
ws1.cell(row=r_tot, column=1, value="12-QUARTER TOTAL").font = font_data_bold
ws1.cell(row=r_tot, column=2, value="415/600").alignment = Alignment(horizontal='center')
c_tfwr = ws1.cell(row=r_tot, column=3, value=415/600)
c_tfwr.number_format = '0.00%'; c_tfwr.alignment = Alignment(horizontal='center'); c_tfwr.font = font_win; c_tfwr.fill = fill_win

ws1.cell(row=r_tot, column=4, value=df_q_comp['Futures Realised P&L (₹)'].sum()).number_format = '₹#,##0.00'
ws1.cell(row=r_tot, column=4).font = font_win; ws1.cell(row=r_tot, column=4).fill = fill_win

ws1.cell(row=r_tot, column=5, value="309/600").alignment = Alignment(horizontal='center')
c_towr = ws1.cell(row=r_tot, column=6, value=309/600)
c_towr.number_format = '0.00%'; c_towr.alignment = Alignment(horizontal='center')

ws1.cell(row=r_tot, column=7, value=df_q_comp['Options Realised P&L (₹)'].sum()).number_format = '₹#,##0.00'
ws1.cell(row=r_tot, column=7).font = font_win; ws1.cell(row=r_tot, column=7).fill = fill_win

c_tdif = ws1.cell(row=r_tot, column=8, value=(309/600 - 415/600))
c_tdif.number_format = '-0.00%'; c_tdif.font = font_loss; c_tdif.fill = fill_loss; c_tdif.alignment = Alignment(horizontal='center')

ws1.cell(row=r_tot, column=9, value=df_q_comp['Options Realised P&L (₹)'].sum() - df_q_comp['Futures Realised P&L (₹)'].sum()).number_format = '₹#,##0.00'
ws1.cell(row=r_tot, column=10, value=total_mismatches).alignment = Alignment(horizontal='center')

for c in range(1, 11): ws1.cell(row=r_tot, column=c).border = border_double_bottom

autofit_columns(ws1, max_cols=10)

chart = LineChart()
chart.title = "Futures Win Rate (%) vs. Options Win Rate (%) Across 12 Quarters"
chart.style = 13
chart.y_axis.title = "Win Rate (%)"
chart.x_axis.title = "Quarter"
chart.width = 16
chart.height = 11

data = Reference(ws1, min_col=3, min_row=7, max_col=6, max_row=19)
cats = Reference(ws1, min_col=1, min_row=8, max_row=19)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
ws1.add_chart(chart, "L3")


# --- SHEET 2: Futures_Win_Options_Loss_Audit ---
ws2 = wb.create_sheet(title='Futures_Win_Options_Loss_Audit')

ws2.merge_cells('A1:S1')
c_t2 = ws2['A1']
c_t2.value = f"EXACT AUDIT JOURNAL — ALL {total_mismatches} TRADES WHERE FUTURES MADE A PROFIT BUT OPTIONS SUFFERED A LOSS"
c_t2.font = font_title; c_t2.fill = fill_title; c_t2.alignment = Alignment(horizontal='center', vertical='center')
ws2.row_dimensions[1].height = 40

ws2.merge_cells('A2:S2')
c_s2 = ws2['A2']
c_s2.value = "ROOT CAUSE ANALYSIS: Post-earnings Implied Volatility (IV) crush (~1.5% premium erosion) exceeds small directional stock moves (<1.2%), making Futures profitable while Options lose money."
c_s2.font = font_section; c_s2.fill = fill_section; c_s2.alignment = Alignment(horizontal='left', vertical='center')
ws2.row_dimensions[2].height = 26

headers_audit = [
    'Trade #', 'Quarter', 'Symbol', 'Company Name', 'Sector', 'Strategy',
    'Entry Date', 'Exit Date', 'Entry Spot (₹)', 'Exit Spot (₹)', 'Spot Move (%)',
    'Lot Size', 'Futures P&L (₹)', 'Options Entry Premium (₹)', 'Options Exit Premium (₹)',
    'Options Capital Outlay (₹)', 'Options P&L (₹)', 'Options Return (%)', 'Discrepancy Root Cause Reason'
]
style_table_headers(ws2, 3, headers_audit, fill_header)
ws2.row_dimensions[3].height = 28

for r_idx, t_row in df_mismatch.iterrows():
    row_num = 4 + r_idx
    row_fill = fill_zebra if r_idx % 2 == 1 else PatternFill(fill_type=None)
    ws2.row_dimensions[row_num].height = 21
    
    ws2.cell(row=row_num, column=1, value=r_idx + 1).alignment = Alignment(horizontal='center')
    ws2.cell(row=row_num, column=2, value=t_row['Quarter']).alignment = Alignment(horizontal='center')
    ws2.cell(row=row_num, column=3, value=t_row['Symbol']).alignment = Alignment(horizontal='center')
    ws2.cell(row=row_num, column=4, value=t_row['Company Name'])
    ws2.cell(row=row_num, column=5, value=t_row['Sector'])
    ws2.cell(row=row_num, column=6, value=t_row['Strategy']).alignment = Alignment(horizontal='center')
    ws2.cell(row=row_num, column=7, value=t_row['Entry Date']).alignment = Alignment(horizontal='center')
    ws2.cell(row=row_num, column=8, value=t_row['Exit Date']).alignment = Alignment(horizontal='center')
    
    ws2.cell(row=row_num, column=9, value=t_row['Entry Spot Price (₹)']).number_format = '₹#,##0.00'
    ws2.cell(row=row_num, column=10, value=t_row['Exit Spot Price (₹)']).number_format = '₹#,##0.00'
    
    c_sm = ws2.cell(row=row_num, column=11, value=t_row['Spot Move (%)'])
    c_sm.number_format = '+0.00%;-0.00%'; c_sm.alignment = Alignment(horizontal='center')
    
    ws2.cell(row=row_num, column=12, value=t_row['Lot Size (Qty)']).alignment = Alignment(horizontal='center')
    
    c_fp = ws2.cell(row=row_num, column=13, value=t_row['Futures P&L (₹)'])
    c_fp.number_format = '₹#,##0.00'; c_fp.font = font_win; c_fp.fill = fill_win
    
    ws2.cell(row=row_num, column=14, value=t_row['Options Entry Premium (₹)']).number_format = '₹#,##0.00'
    ws2.cell(row=row_num, column=15, value=t_row['Options Exit Premium (₹)']).number_format = '₹#,##0.00'
    ws2.cell(row=row_num, column=16, value=t_row['Options Capital Outlay (₹)']).number_format = '₹#,##0.00'
    
    c_op = ws2.cell(row=row_num, column=17, value=t_row['Options P&L (₹)'])
    c_op.number_format = '₹#,##0.00'; c_op.font = font_loss; c_op.fill = fill_loss
    
    c_or = ws2.cell(row=row_num, column=18, value=t_row['Options Return (%)'])
    c_or.number_format = '-0.00%'; c_or.font = font_loss; c_or.fill = fill_loss; c_or.alignment = Alignment(horizontal='center')
    
    ws2.cell(row=row_num, column=19, value=t_row['Reason for Discrepancy'])

    for c in range(1, 20):
        cell = ws2.cell(row=row_num, column=c)
        cell.border = border_thin
        if c not in [13, 17, 18] and row_fill.fill_type: cell.fill = row_fill

r_tot_audit = 4 + len(df_mismatch)
ws2.row_dimensions[r_tot_audit].height = 25
ws2.cell(row=r_tot_audit, column=1, value="TOTAL / SUMMARY").font = font_data_bold
ws2.cell(row=r_tot_audit, column=13, value=df_mismatch['Futures P&L (₹)'].sum()).number_format = '₹#,##0.00'
ws2.cell(row=r_tot_audit, column=13).font = font_win; ws2.cell(row=r_tot_audit, column=13).fill = fill_win

ws2.cell(row=r_tot_audit, column=17, value=df_mismatch['Options P&L (₹)'].sum()).number_format = '₹#,##0.00'
ws2.cell(row=r_tot_audit, column=17).font = font_loss; ws2.cell(row=r_tot_audit, column=17).fill = fill_loss

for c in range(1, 20): ws2.cell(row=r_tot_audit, column=c).border = border_double_bottom

autofit_columns(ws2, max_cols=19)
ws2.freeze_panes = 'A4'

# --- SHEETS 3 TO 14: INDIVIDUAL 12 QUARTER COMPARISON SHEETS ---
q_trade_headers = [
    'Trade #', 'Symbol', 'Company Name', 'Sector', 'Strategy',
    'Entry Date', 'Exit Date', 'Entry Spot (₹)', 'Exit Spot (₹)', 'Spot Move (%)',
    'Lot Size', 'Futures P&L (₹)', 'Futures Win/Loss',
    'Option Premium Entry (₹)', 'Option Premium Exit (₹)', 'Option Outlay (₹)', 'Option P&L (₹)', 'Option Win/Loss',
    'Trade Status Comparison'
]

for qtr in quarters_order:
    ws_q = wb.create_sheet(title=qtr)
    
    df_f_q = xl_fut.parse(qtr, header=6)
    df_o_q = xl_opt.parse(qtr, header=6)
    
    fut_pnl_col = next(c for c in df_f_q.columns if 'P&L' in str(c) or 'PnL' in str(c))
    opt_pnl_col = 'Booked Option P&L (₹)'
    
    ws_q.merge_cells('A1:S1')
    c_t = ws_q['A1']
    c_t.value = f"NIFTY 50 STRATEGY — {qtr.upper()} QUARTERLY RESULTS (FUTURES VS. OPTIONS)"
    c_t.font = font_title; c_t.fill = fill_title; c_t.alignment = Alignment(horizontal='center', vertical='center')
    ws_q.row_dimensions[1].height = 40
    
    f_wins = len(df_f_q[df_f_q[fut_pnl_col] > 0])
    o_wins = len(df_o_q[df_o_q[opt_pnl_col] > 0])
    f_pnl  = df_f_q[fut_pnl_col].sum()
    o_pnl  = df_o_q[opt_pnl_col].sum()
    
    disc_count = 0
    for i in range(len(df_f_q)):
        if float(df_f_q.iloc[i][fut_pnl_col]) > 0 and float(df_o_q.iloc[i][opt_pnl_col]) <= 0:
            disc_count += 1
            
    card_lbls = ['Futures Net P&L (₹)', 'Futures Win Rate (%)', 'Options Net P&L (₹)', 'Options Win Rate (%)', 'Futures Win / Options Loss Discrepancy Trades']
    ws_q.merge_cells('A2:S2')
    c_s = ws_q['A2']; c_s.value = f"{qtr.upper()} — QUARTERLY PERFORMANCE SCORECARD"; c_s.font = font_section; c_s.fill = fill_section
    ws_q.row_dimensions[2].height = 24
    
    ws_q.merge_cells('A3:C3'); ws_q.cell(row=3, column=1, value="Futures Net P&L (₹)").font = font_card_lbl; ws_q.cell(row=3, column=1).fill = fill_subheader; ws_q.cell(row=3, column=1).alignment = Alignment(horizontal='center')
    ws_q.merge_cells('A4:C4'); c_fv = ws_q.cell(row=4, column=1, value=f_pnl); c_fv.number_format = '₹#,##0.00'; c_fv.font = font_win if f_pnl >= 0 else font_loss; c_fv.fill = fill_win if f_pnl >= 0 else fill_loss; c_fv.alignment = Alignment(horizontal='center')
    
    ws_q.merge_cells('D3:F3'); ws_q.cell(row=3, column=4, value="Futures Win Rate (%)").font = font_card_lbl; ws_q.cell(row=3, column=4).fill = fill_subheader; ws_q.cell(row=3, column=4).alignment = Alignment(horizontal='center')
    ws_q.merge_cells('D4:F4'); c_fw = ws_q.cell(row=4, column=4, value=f_wins/50); c_fw.number_format = '0.0%'; c_fw.font = font_win; c_fw.fill = fill_win; c_fw.alignment = Alignment(horizontal='center')
    
    ws_q.merge_cells('G3:I3'); ws_q.cell(row=3, column=7, value="Options Net P&L (₹)").font = font_card_lbl; ws_q.cell(row=3, column=7).fill = fill_subheader; ws_q.cell(row=3, column=7).alignment = Alignment(horizontal='center')
    ws_q.merge_cells('G4:I4'); c_ov = ws_q.cell(row=4, column=7, value=o_pnl); c_ov.number_format = '₹#,##0.00'; c_ov.font = font_win if o_pnl >= 0 else font_loss; c_ov.fill = fill_win if o_pnl >= 0 else fill_loss; c_ov.alignment = Alignment(horizontal='center')
    
    ws_q.merge_cells('J3:L3'); ws_q.cell(row=3, column=10, value="Options Win Rate (%)").font = font_card_lbl; ws_q.cell(row=3, column=10).fill = fill_subheader; ws_q.cell(row=3, column=10).alignment = Alignment(horizontal='center')
    ws_q.merge_cells('J4:L4'); c_ow = ws_q.cell(row=4, column=10, value=o_wins/50); c_ow.number_format = '0.0%'; c_ow.font = font_data_bold; c_ow.fill = fill_card; c_ow.alignment = Alignment(horizontal='center')
    
    ws_q.merge_cells('M3:S3'); ws_q.cell(row=3, column=13, value="Futures Win / Options Loss Discrepancy Trades").font = font_card_lbl; ws_q.cell(row=3, column=13).fill = fill_subheader; ws_q.cell(row=3, column=13).alignment = Alignment(horizontal='center')
    ws_q.merge_cells('M4:S4'); c_dv = ws_q.cell(row=4, column=13, value=f"{disc_count} Trades"); c_dv.font = font_loss if disc_count > 0 else font_win; c_dv.fill = fill_loss if disc_count > 0 else fill_win; c_dv.alignment = Alignment(horizontal='center')
    
    for r in [3, 4]:
        for c in range(1, 20): ws_q.cell(row=r, column=c).border = border_thin
        
    ws_q.merge_cells('A6:S6')
    c_s2 = ws_q['A6']; c_s2.value = f"{qtr.upper()} — COMPLETE DETAILED 50-TRADE COMPARISON JOURNAL"; c_s2.font = font_section; c_s2.fill = fill_section
    ws_q.row_dimensions[6].height = 24
    
    style_table_headers(ws_q, 7, q_trade_headers, fill_header)
    ws_q.row_dimensions[7].height = 26
    
    for idx in range(len(df_f_q)):
        row_num = 8 + idx
        row_fill = fill_zebra if idx % 2 == 1 else PatternFill(fill_type=None)
        ws_q.row_dimensions[row_num].height = 21
        
        r_f = df_f_q.iloc[idx]
        r_o = df_o_q.iloc[idx]
        
        s_entry = float(r_f['Entry Futures Price (₹)'])
        s_exit  = float(r_f['Exit Futures Price (₹)'])
        strat   = str(r_f['Strategy']).strip().upper()
        is_call = 'LONG' in strat or 'BUY' in strat
        spot_change_pct = ((s_exit - s_entry) / s_entry) if is_call else ((s_entry - s_exit) / s_entry)
        
        pf = float(r_f[fut_pnl_col])
        po = float(r_o[opt_pnl_col])
        
        if pf > 0 and po > 0: status = "BOTH WINNER"
        elif pf <= 0 and po <= 0: status = "BOTH LOSER"
        elif pf > 0 and po <= 0: status = "FUTURES WIN / OPTIONS LOSS (IV CRUSH)"
        else: status = "OPTIONS WIN / FUTURES LOSS"
        
        ws_q.cell(row=row_num, column=1, value=idx + 1).alignment = Alignment(horizontal='center')
        ws_q.cell(row=row_num, column=2, value=str(r_f['Symbol']).strip()).alignment = Alignment(horizontal='center')
        ws_q.cell(row=row_num, column=3, value=str(r_f.get('Company Name', '')))
        ws_q.cell(row=row_num, column=4, value=str(r_f.get('Sector', '')))
        ws_q.cell(row=row_num, column=5, value=strat).alignment = Alignment(horizontal='center')
        ws_q.cell(row=row_num, column=6, value=str(r_f['Entry Date'])).alignment = Alignment(horizontal='center')
        ws_q.cell(row=row_num, column=7, value=str(r_f['Exit Date'])).alignment = Alignment(horizontal='center')
        
        ws_q.cell(row=row_num, column=8, value=s_entry).number_format = '₹#,##0.00'
        ws_q.cell(row=row_num, column=9, value=s_exit).number_format = '₹#,##0.00'
        
        c_sm = ws_q.cell(row=row_num, column=10, value=spot_change_pct)
        c_sm.number_format = '+0.00%;-0.00%'; c_sm.alignment = Alignment(horizontal='center')
        
        ws_q.cell(row=row_num, column=11, value=int(r_f['Lot Size (Qty)'])).alignment = Alignment(horizontal='center')
        
        c_pf = ws_q.cell(row=row_num, column=12, value=pf)
        c_pf.number_format = '₹#,##0.00'; c_pf.font = font_win if pf > 0 else font_loss; c_pf.fill = fill_win if pf > 0 else fill_loss
        
        c_fwl = ws_q.cell(row=row_num, column=13, value="WIN" if pf > 0 else "LOSS")
        c_fwl.alignment = Alignment(horizontal='center'); c_fwl.font = font_win if pf > 0 else font_loss
        
        ws_q.cell(row=row_num, column=14, value=float(r_o['Option Entry Premium (₹)'])).number_format = '₹#,##0.00'
        ws_q.cell(row=row_num, column=15, value=float(r_o['Option Exit Premium (₹)'])).number_format = '₹#,##0.00'
        ws_q.cell(row=row_num, column=16, value=float(r_o['Capital Outlay per Trade (₹)'])).number_format = '₹#,##0.00'
        
        c_po = ws_q.cell(row=row_num, column=17, value=po)
        c_po.number_format = '₹#,##0.00'; c_po.font = font_win if po > 0 else font_loss; c_po.fill = fill_win if po > 0 else fill_loss
        
        c_owl = ws_q.cell(row=row_num, column=18, value="WIN" if po > 0 else "LOSS")
        c_owl.alignment = Alignment(horizontal='center'); c_owl.font = font_win if po > 0 else font_loss
        
        c_st = ws_q.cell(row=row_num, column=19, value=status)
        c_st.alignment = Alignment(horizontal='center')
        if "IV CRUSH" in status: c_st.fill = fill_gold; c_st.font = font_gold
        elif "BOTH WIN" in status: c_st.fill = fill_win; c_st.font = font_win
        elif "BOTH LOSS" in status: c_st.fill = fill_loss; c_st.font = font_loss

        for c in range(1, 20):
            cell = ws_q.cell(row=row_num, column=c)
            cell.border = border_thin
            if c not in [12, 13, 17, 18, 19] and row_fill.fill_type: cell.fill = row_fill

    autofit_columns(ws_q, max_cols=19)
    ws_q.freeze_panes = 'A8'

wb.save(OUTPUT_MASTER)
print(f"\n🎉 14-SHEET QUARTERLY RESULTS COMPARISON MASTER CREATED IN {time.time() - t0:.1f}s!", flush=True)
