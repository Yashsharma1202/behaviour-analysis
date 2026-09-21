import pandas as pd
import pathlib
import sys
import time
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference

sys.stdout.reconfigure(errors='replace')

t0 = time.time()

BASE_DIR = pathlib.Path('D:/behaviour analysis')
OUTPUT_EXCEL = BASE_DIR / 'Nifty50_12_Quarters_True_Peak_Drawdown_Summary.xlsx'

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

# ------------------- Data Table with True Peak Drawdown % -------------------
table_data = [
    {'Quarter': 'Q3 2023-24', 'Trades': 50, 'Margin Deployed': 117749.07, 'Net Futures P&L': 630821.00, 'Peak Cum PnL': 803616.50, 'Trough Cum PnL': 630821.00, 'Max DD ₹': -172795.50, 'True Peak DD %': -0.1875, 'Win Rate': 0.9200, 'Trough Stock': 'TMPV'},
    {'Quarter': 'Q4 2023-24', 'Trades': 50, 'Margin Deployed': 133482.27, 'Net Futures P&L': 1662111.75, 'Peak Cum PnL': 1834907.25, 'Trough Cum PnL': 1662111.75, 'Max DD ₹': -172795.50, 'True Peak DD %': -0.0878, 'Win Rate': 0.9400, 'Trough Stock': 'TMPV'},
    {'Quarter': 'Q1 2024-25', 'Trades': 50, 'Margin Deployed': 145795.34, 'Net Futures P&L': 368367.25, 'Peak Cum PnL': 541162.75, 'Trough Cum PnL': 368367.25, 'Max DD ₹': -172795.50, 'True Peak DD %': -0.2515, 'Win Rate': 0.6400, 'Trough Stock': 'TMPV'},
    {'Quarter': 'Q2 2024-25', 'Trades': 50, 'Margin Deployed': 158146.39, 'Net Futures P&L': 637875.75, 'Peak Cum PnL': 495426.50, 'Trough Cum PnL': 294232.50, 'Max DD ₹': -201194.00, 'True Peak DD %': -0.3078, 'Win Rate': 0.6800, 'Trough Stock': 'TMPV'},
    {'Quarter': 'Q3 2024-25', 'Trades': 50, 'Margin Deployed': 159872.57, 'Net Futures P&L': 578677.50, 'Peak Cum PnL': 420331.00, 'Trough Cum PnL': 333532.00, 'Max DD ₹': -86799.00, 'True Peak DD %': -0.1496, 'Win Rate': 0.6200, 'Trough Stock': 'TRENT'},
    {'Quarter': 'Q4 2024-25', 'Trades': 50, 'Margin Deployed': 147019.84, 'Net Futures P&L': 264661.25, 'Peak Cum PnL': 70982.50, 'Trough Cum PnL': -48316.75, 'Max DD ₹': -119299.25, 'True Peak DD %': -0.5472, 'Win Rate': 0.5200, 'Trough Stock': 'LT'},
    {'Quarter': 'Q1 2025-26', 'Trades': 50, 'Margin Deployed': 151114.54, 'Net Futures P&L': 770148.75, 'Peak Cum PnL': 62078.50, 'Trough Cum PnL': -142424.50, 'Max DD ₹': -204503.00, 'True Peak DD %': -0.9592, 'Win Rate': 0.6200, 'Trough Stock': 'TECHM'},
    {'Quarter': 'Q2 2025-26', 'Trades': 50, 'Margin Deployed': 155685.80, 'Net Futures P&L': -127162.50, 'Peak Cum PnL': 12007.00, 'Trough Cum PnL': -204031.50, 'Max DD ₹': -216038.50, 'True Peak DD %': -1.2883, 'Win Rate': 0.5200, 'Trough Stock': 'ULTRACEMCO'},
    {'Quarter': 'Q3 2025-26', 'Trades': 50, 'Margin Deployed': 163725.49, 'Net Futures P&L': 680108.50, 'Peak Cum PnL': 318329.75, 'Trough Cum PnL': 221078.25, 'Max DD ₹': -97251.50, 'True Peak DD %': -0.2017, 'Win Rate': 0.6600, 'Trough Stock': 'HINDALCO'},
    {'Quarter': 'Q4 2025-26', 'Trades': 50, 'Margin Deployed': 162284.00, 'Net Futures P&L': 1152472.75, 'Peak Cum PnL': 793176.75, 'Trough Cum PnL': 687343.25, 'Max DD ₹': -105833.50, 'True Peak DD %': -0.1108, 'Win Rate': 0.7200, 'Trough Stock': 'POWERGRID'},
    {'Quarter': 'Q1 2026-27', 'Trades': 50, 'Margin Deployed': 160520.17, 'Net Futures P&L': 516113.25, 'Peak Cum PnL': 49138.25, 'Trough Cum PnL': -53174.75, 'Max DD ₹': -102313.00, 'True Peak DD %': -0.4880, 'Win Rate': 0.6000, 'Trough Stock': 'TRENT'},
    {'Quarter': 'Q2 2026-27', 'Trades': 50, 'Margin Deployed': 3523469.33, 'Net Futures P&L': 1443141.00, 'Peak Cum PnL': 1349085.50, 'Trough Cum PnL': 1286570.50, 'Max DD ₹': -62515.00, 'True Peak DD %': -0.0128, 'Win Rate': 0.8600, 'Trough Stock': 'MAXHEALTH'}
]

df_table = pd.DataFrame(table_data)

print("Building Standalone True Peak Drawdown Summary Excel...", flush=True)
wb = Workbook()
wb.remove(wb.active)

ws = wb.create_sheet(title='True_Peak_Drawdown_Summary')
ws.views.sheetView[0].showGridLines = True

# Title Banner
ws.merge_cells('A1:I1')
c_title = ws['A1']
c_title.value = "NIFTY 50 FUTURES STRATEGY — 12-QUARTER TRUE PEAK DRAWDOWN SUMMARY SCORECARD"
c_title.font = font_title; c_title.fill = fill_title; c_title.alignment = Alignment(horizontal='center', vertical='center')
ws.row_dimensions[1].height = 40

# Hero Summary Cards
ws.merge_cells('A3:C3')
c_h1 = ws['A3']; c_h1.value = "TOTAL TRADES ANALYZED"; c_h1.font = font_card_lbl; c_h1.fill = fill_subheader; c_h1.alignment = Alignment(horizontal='center', vertical='center')
ws.merge_cells('A4:C4')
c_v1 = ws['A4']; c_v1.value = 600; c_v1.font = Font(name='Segoe UI', size=15, bold=True, color=COLOR_TEXT_MAIN); c_v1.fill = fill_card; c_v1.alignment = Alignment(horizontal='center', vertical='center')

ws.merge_cells('D3:F3')
c_h2 = ws['D3']; c_h2.value = "GLOBAL WIN RATE"; c_h2.font = font_card_lbl; c_h2.fill = fill_subheader; c_h2.alignment = Alignment(horizontal='center', vertical='center')
ws.merge_cells('D4:F4')
c_v2 = ws['D4']; c_v2.value = 0.6917; c_v2.font = Font(name='Segoe UI', size=15, bold=True, color=COLOR_WIN_TEXT); c_v2.fill = fill_win; c_v2.number_format = '0.0%'; c_v2.alignment = Alignment(horizontal='center', vertical='center')

ws.merge_cells('G3:I3')
c_h3 = ws['G3']; c_h3.value = "TOTAL 12-QUARTER NET FUTURES PROFIT"; c_h3.font = font_card_lbl; c_h3.fill = fill_subheader; c_h3.alignment = Alignment(horizontal='center', vertical='center')
ws.merge_cells('G4:I4')
c_v3 = ws['G4']; c_v3.value = df_table['Net Futures P&L'].sum(); c_v3.font = Font(name='Segoe UI', size=15, bold=True, color=COLOR_WIN_TEXT); c_v3.fill = fill_win; c_v3.number_format = '₹#,##0.00'; c_v3.alignment = Alignment(horizontal='center', vertical='center')

ws.row_dimensions[3].height = 20
ws.row_dimensions[4].height = 28

for r in [3, 4]:
    for c in range(1, 10): ws.cell(row=r, column=c).border = border_thin

# Section Banner
ws.merge_cells('A6:I6')
c_sec = ws['A6']
c_sec.value = "CONSOLIDATED 12-QUARTER TRUE PEAK-TO-TROUGH DRAWDOWN SCORECARD"
c_sec.font = font_section; c_sec.fill = fill_section; c_sec.alignment = Alignment(horizontal='left', vertical='center')
ws.row_dimensions[6].height = 24

# Headers
headers = ['Quarter', 'Peak Cum P&L (₹)', 'Trough Cum P&L (₹)', 'Net Futures P&L (₹)', 'Max Drawdown (₹)', 'True Peak DD (%)', 'Win Rate (%)', 'Trough Stock']
for c_idx, h in enumerate(headers, start=1):
    cell = ws.cell(row=7, column=c_idx, value=h)
    cell.font = font_header; cell.fill = fill_header; cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = border_thin
ws.row_dimensions[7].height = 26

# Data Rows
for r_idx, r_data in df_table.iterrows():
    row_num = 8 + r_idx
    row_fill = fill_zebra if r_idx % 2 == 1 else PatternFill(fill_type=None)
    ws.row_dimensions[row_num].height = 21
    
    ws.cell(row=row_num, column=1, value=r_data['Quarter']).alignment = Alignment(horizontal='center')
    ws.cell(row=row_num, column=2, value=r_data['Peak Cum PnL']).number_format = '₹#,##0.00'
    ws.cell(row=row_num, column=3, value=r_data['Trough Cum PnL']).number_format = '₹#,##0.00'
    
    c_pnl = ws.cell(row=row_num, column=4, value=r_data['Net Futures P&L'])
    c_pnl.number_format = '₹#,##0.00'; c_pnl.font = font_win if r_data['Net Futures P&L'] >= 0 else font_loss
    c_pnl.fill = fill_win if r_data['Net Futures P&L'] >= 0 else fill_loss
    
    c_dd = ws.cell(row=row_num, column=5, value=r_data['Max DD ₹'])
    c_dd.number_format = '-₹#,##0.00'; c_dd.font = font_loss; c_dd.fill = fill_loss
    
    c_ddp = ws.cell(row=row_num, column=6, value=r_data['True Peak DD %'])
    c_ddp.number_format = '-0.00%'; c_ddp.font = font_loss; c_ddp.fill = fill_loss; c_ddp.alignment = Alignment(horizontal='center')
    
    c_wr = ws.cell(row=row_num, column=7, value=r_data['Win Rate'])
    c_wr.number_format = '0.00%'; c_wr.alignment = Alignment(horizontal='center')
    if r_data['Win Rate'] >= 0.70: c_wr.fill = fill_win; c_wr.font = font_win
    elif r_data['Win Rate'] >= 0.50: c_wr.fill = fill_gold; c_wr.font = font_gold
    else: c_wr.fill = fill_loss; c_wr.font = font_loss
    
    ws.cell(row=row_num, column=8, value=r_data['Trough Stock']).alignment = Alignment(horizontal='center')

    for c in range(1, 9):
        cell = ws.cell(row=row_num, column=c)
        cell.border = border_thin
        if c not in [4, 5, 6, 7] and row_fill.fill_type: cell.fill = row_fill

# Row 20: 12-QUARTER TOTAL ROW
r_tot = 20
ws.row_dimensions[r_tot].height = 24
ws.cell(row=r_tot, column=1, value="12-QUARTER TOTAL").font = font_data_bold
ws.cell(row=r_tot, column=2, value="—").alignment = Alignment(horizontal='center')
ws.cell(row=r_tot, column=3, value="—").alignment = Alignment(horizontal='center')
ws.cell(row=r_tot, column=4, value=df_table['Net Futures P&L'].sum()).number_format = '₹#,##0.00'
ws.cell(row=r_tot, column=4).font = font_win; ws.cell(row=r_tot, column=4).fill = fill_win

ws.cell(row=r_tot, column=5, value=df_table['Max DD ₹'].min()).number_format = '-₹#,##0.00'
ws.cell(row=r_tot, column=5).font = font_loss; ws.cell(row=r_tot, column=5).fill = fill_loss

ws.cell(row=r_tot, column=6, value=df_table['True Peak DD %'].min()).number_format = '-0.00%'
ws.cell(row=r_tot, column=6).font = font_loss; ws.cell(row=r_tot, column=6).fill = fill_loss; ws.cell(row=r_tot, column=6).alignment = Alignment(horizontal='center')

ws.cell(row=r_tot, column=7, value=0.6917).number_format = '0.00%'
ws.cell(row=r_tot, column=7).font = font_win; ws.cell(row=r_tot, column=7).fill = fill_win; ws.cell(row=r_tot, column=7).alignment = Alignment(horizontal='center')

for c in range(1, 9): ws.cell(row=r_tot, column=c).border = border_thin

# Row 21: QUARTERLY AVERAGE ROW
r_avg = 21
ws.row_dimensions[r_avg].height = 24
ws.cell(row=r_avg, column=1, value="QUARTERLY AVERAGE").font = font_data_bold
ws.cell(row=r_avg, column=2, value=df_table['Peak Cum PnL'].mean()).number_format = '₹#,##0.00'
ws.cell(row=r_avg, column=3, value=df_table['Trough Cum PnL'].mean()).number_format = '₹#,##0.00'
ws.cell(row=r_avg, column=4, value=df_table['Net Futures P&L'].mean()).number_format = '₹#,##0.00'
ws.cell(row=r_avg, column=4).font = font_win; ws.cell(row=r_avg, column=4).fill = fill_win

ws.cell(row=r_avg, column=5, value=df_table['Max DD ₹'].mean()).number_format = '-₹#,##0.00'
ws.cell(row=r_avg, column=5).font = font_loss; ws.cell(row=r_avg, column=5).fill = fill_loss

ws.cell(row=r_avg, column=6, value=df_table['True Peak DD %'].mean()).number_format = '-0.00%'
ws.cell(row=r_avg, column=6).font = font_loss; ws.cell(row=r_avg, column=6).fill = fill_loss; ws.cell(row=r_avg, column=6).alignment = Alignment(horizontal='center')

ws.cell(row=r_avg, column=7, value=0.6917).number_format = '0.00%'
ws.cell(row=r_avg, column=7).font = font_win; ws.cell(row=r_avg, column=7).fill = fill_win; ws.cell(row=r_avg, column=7).alignment = Alignment(horizontal='center')

for c in range(1, 9): ws.cell(row=r_avg, column=c).border = border_double_bottom

# Autofit Columns
for col_idx in range(1, 9):
    col_letter = get_column_letter(col_idx)
    max_len = max(len(str(ws.cell(row=r, column=col_idx).value or '')) for r in range(1, 22))
    ws.column_dimensions[col_letter].width = max(max_len + 5, 18)

# Chart
chart = LineChart()
chart.title = "Consolidated 12-Quarter Net Futures Realised Profit Trend (₹)"
chart.style = 13
chart.y_axis.title = "Profit (₹)"
chart.x_axis.title = "Quarter"
chart.width = 16
chart.height = 11

data = Reference(ws, min_col=4, min_row=7, max_row=19)
cats = Reference(ws, min_col=1, min_row=8, max_row=19)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
ws.add_chart(chart, "J3")

wb.save(OUTPUT_EXCEL)
print(f"\n🎉 STANDALONE TRUE PEAK DRAWDOWN SUMMARY EXCEL CREATED IN {time.time() - t0:.1f}s!", flush=True)
