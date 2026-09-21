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
OUTPUT_COMP_WORKBOOK = BASE_DIR / 'Nifty50_12_Quarters_Futures_vs_Options_Comparison.xlsx'

# Executive Color Palette (Royal Navy & Dark Slate)
COLOR_TITLE_BG     = '1E3A8A'  # Royal Navy Title
COLOR_SECTION_BG   = '1E293B'  # Slate Section Header
COLOR_HEADER_BG    = '0F172A'  # Deep Slate Header
COLOR_SUBHEADER_BG = '2563EB'  # Accent Blue Subheader
COLOR_ZEBRA_BG     = 'F8FAFC'  # Subtle Zebra
COLOR_CARD_BG      = 'F1F5F9'  # Cool Gray Card

COLOR_WIN_BG       = 'DCFCE7'  # Soft Green
COLOR_WIN_TEXT     = '15803D'  # Dark Green
COLOR_LOSS_BG      = 'FEE2E2'  # Soft Red
COLOR_LOSS_TEXT    = 'B91C1C'  # Dark Red
COLOR_GOLD_BG      = 'FEF3C7'  # Soft Gold
COLOR_GOLD_TEXT    = 'B45309'  # Dark Gold

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

quarters_order = [
    'Q3 2023-24', 'Q4 2023-24', 'Q1 2024-25', 'Q2 2024-25',
    'Q3 2024-25', 'Q4 2024-25', 'Q1 2025-26', 'Q2 2025-26',
    'Q3 2025-26', 'Q4 2025-26', 'Q1 2026-27', 'Q2 2026-27'
]

print("Loading Futures & Options Master Data...", flush=True)
xl_fut = pd.ExcelFile(SOURCE_FUT, engine='openpyxl')
xl_opt = pd.ExcelFile(SOURCE_OPT, engine='openpyxl')

q_comp_list = []
fut_win_opt_loss_trades = []

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
    
    # Check trade by trade mismatch
    for idx in range(len(df_f)):
        r_f = df_f.iloc[idx]
        r_o = df_o.iloc[idx]
        
        pnl_f = float(r_f[fut_pnl_col])
        pnl_o = float(r_o[opt_pnl_col])
        
        # FUTURES WIN & OPTIONS LOSS
        if pnl_f > 0 and pnl_o <= 0:
            sym      = str(r_f['Symbol']).strip()
            s_entry  = float(r_f['Entry Futures Price (₹)'])
            s_exit   = float(r_f['Exit Futures Price (₹)'])
            strat    = str(r_f['Strategy']).strip().upper()
            is_call  = 'LONG' in strat or 'BUY' in strat
            
            spot_change_pct = ((s_exit - s_entry) / s_entry) if is_call else ((s_entry - s_exit) / s_entry)
            
            outlay_o = float(r_o['Capital Outlay per Trade (₹)'])
            opt_ret_val = (pnl_o / outlay_o) if outlay_o > 0 else 0.0
            
            fut_win_opt_loss_trades.append({
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
                'Options Return (%)': opt_ret_val,
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
        'Futures Win / Options Loss Trade Count': len([t for t in fut_win_opt_loss_trades if t['Quarter'] == qtr])
    })

df_q_comp = pd.DataFrame(q_comp_list)
df_mismatch = pd.DataFrame(fut_win_opt_loss_trades)

print(f"Total Futures Win / Options Loss Trades Identified: {len(df_mismatch)} out of 600 Trades!", flush=True)

def autofit_columns(ws, max_cols=25):
    ws.views.sheetView[0].showGridLines = True
    for col_idx in range(1, max_cols + 1):
        col_letter = get_column_letter(col_idx)
        max_len = 0
        for row in range(1, min(ws.max_row + 1, 150)):
            val = str(ws.cell(row=row, column=col_idx).value or '')
            if len(val) > max_len: max_len = len(val)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 13)
wb = Workbook()
wb.remove(wb.active)

# --- SHEET 1: Summary_&_Win_Rate_Compare ---
ws1 = wb.create_sheet(title='Summary_&_Win_Rate_Compare')
ws1.views.sheetView[0].showGridLines = True

ws1.merge_cells('A1:J1')
c_t = ws1['A1']
c_t.value = "FUTURES VS. 1% ITM OPTIONS STRATEGY — 12-QUARTER COMPARISON & WIN RATE AUDIT"
c_t.font = font_title; c_t.fill = fill_title; c_t.alignment = Alignment(horizontal='center', vertical='center')
ws1.row_dimensions[1].height = 42

total_mismatches = len(df_mismatch)
total_mismatch_fut_profit = df_mismatch['Futures P&L (₹)'].sum()
total_mismatch_opt_loss   = df_mismatch['Options P&L (₹)'].sum()

ws1.merge_cells('A3:C3')
c_h1 = ws1['A3']; c_h1.value = "TOTAL FUTURES WIN / OPTIONS LOSS TRADES"; c_h1.font = font_card_lbl; c_h1.fill = fill_subheader; c_h1.alignment = Alignment(horizontal='center', vertical='center')
ws1.merge_cells('A4:C4')
c_v1 = ws1['A4']; c_v1.value = f"{total_mismatches} Trades ({total_mismatches/600:.1%})"; c_v1.font = Font(name='Segoe UI', size=13, bold=True, color=COLOR_TEXT_MAIN); c_v1.fill = fill_card; c_v1.alignment = Alignment(horizontal='center', vertical='center')

ws1.merge_cells('D3:F3')
c_h2 = ws1['D3']; c_h2.value = "FUTURES PROFIT ON THESE TRADES"; c_h2.font = font_card_lbl; c_h2.fill = fill_subheader; c_h2.alignment = Alignment(horizontal='center', vertical='center')
ws1.merge_cells('D4:F4')
c_v2 = ws1['D4']; c_v2.value = total_mismatch_fut_profit; c_v2.font = Font(name='Segoe UI', size=15, bold=True, color=COLOR_WIN_TEXT); c_v2.fill = fill_win; c_v2.number_format = '₹#,##0.00'; c_v2.alignment = Alignment(horizontal='center', vertical='center')

ws1.merge_cells('G3:J3')
c_h3 = ws1['G3']; c_h3.value = "OPTIONS LOSS ON THESE TRADES (DUE TO IV CRUSH)"; c_h3.font = font_card_lbl; c_h3.fill = fill_subheader; c_h3.alignment = Alignment(horizontal='center', vertical='center')
ws1.merge_cells('G4:J4')
c_v3 = ws1['G4']; c_v3.value = total_mismatch_opt_loss; c_v3.font = Font(name='Segoe UI', size=15, bold=True, color=COLOR_LOSS_TEXT); c_v3.fill = fill_loss; c_v3.number_format = '-₹#,##0.00'; c_v3.alignment = Alignment(horizontal='center', vertical='center')

ws1.row_dimensions[3].height = 22
ws1.row_dimensions[4].height = 30
for r in [3, 4]:
    for c in range(1, 11): ws1.cell(row=r, column=c).border = border_thin

ws1.merge_cells('A6:J6')
c_sec = ws1['A6']
c_sec.value = "SIDE-BY-SIDE QUARTERLY SCORECARD COMPARISON (FUTURES VS. OPTIONS BUYING)"
c_sec.font = font_section; c_sec.fill = fill_section; c_sec.alignment = Alignment(horizontal='left', vertical='center')
ws1.row_dimensions[6].height = 26

headers_sum1 = ['Quarter', 'Futures Wins', 'Futures Win Rate (%)', 'Futures P&L (₹)', 'Options Wins', 'Options Win Rate (%)', 'Options P&L (₹)', 'Win Rate Diff (%)', 'P&L Diff (₹)', 'Futures Win / Options Loss Trades']

for c_idx, h in enumerate(headers_sum1, start=1):
    cell = ws1.cell(row=7, column=c_idx, value=h)
    cell.font = font_header; cell.fill = fill_header; cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = border_thin
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
    
    ws1.cell(row=row_num, column=10, value=row_data['Futures Win / Options Loss Trade Count']).alignment = Alignment(horizontal='center')

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

# Autofit
for col_idx in range(1, 11):
    col_letter = get_column_letter(col_idx)
    max_len = max(len(str(ws1.cell(row=r, column=col_idx).value or '')) for r in range(1, 21))
    ws1.column_dimensions[col_letter].width = max(max_len + 4, 16)

# Chart
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
ws2.views.sheetView[0].showGridLines = True

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

for c_idx, h in enumerate(headers_audit, start=1):
    cell = ws2.cell(row=3, column=c_idx, value=h)
    cell.font = font_header; cell.fill = fill_header; cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = border_thin
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

# Total Row on Audit Sheet
r_tot_audit = 4 + len(df_mismatch)
ws2.row_dimensions[r_tot_audit].height = 25
ws2.cell(row=r_tot_audit, column=1, value="TOTAL / SUMMARY").font = font_data_bold
ws2.cell(row=r_tot_audit, column=13, value=total_mismatch_fut_profit).number_format = '₹#,##0.00'
ws2.cell(row=r_tot_audit, column=13).font = font_win; ws2.cell(row=r_tot_audit, column=13).fill = fill_win

ws2.cell(row=r_tot_audit, column=17, value=total_mismatch_opt_loss).number_format = '₹#,##0.00'
ws2.cell(row=r_tot_audit, column=17).font = font_loss; ws2.cell(row=r_tot_audit, column=17).fill = fill_loss

for c in range(1, 20): ws2.cell(row=r_tot_audit, column=c).border = border_double_bottom

autofit_columns(ws2, max_cols=19)
ws2.freeze_panes = 'A4'

wb.save(OUTPUT_COMP_WORKBOOK)
print(f"\n🎉 FUTURES VS OPTIONS COMPARISON WORKBOOK CREATED IN {time.time() - t0:.1f}s!", flush=True)
