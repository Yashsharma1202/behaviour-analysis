import pandas as pd
import pathlib
import sys
import time
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference, Series

sys.stdout.reconfigure(errors='replace')

t0 = time.time()

# ------------------- Configuration & Paths -------------------
BASE_DIR = pathlib.Path('D:/behaviour analysis')
V13_PATH = BASE_DIR / 'Nifty50_12_Quarters_Futures_OI_Master_v13.xlsx'
OUTPUT_CHARTS_ONLY = BASE_DIR / 'Nifty50_12_Quarters_Drawdown_Charts_Only.xlsx'

# ------------------- Color Palette Tokens -------------------
COLOR_TITLE_BG     = '1B365D'  # Deep Imperial Navy
COLOR_SECTION_BG   = '2B4C7E'  # Royal Slate Accent
COLOR_HEADER_BG    = '1A202C'  # Rich Charcoal Navy
COLOR_WHITE        = 'FFFFFF'
COLOR_TEXT_MAIN    = '2D3748'

font_title     = Font(name='Segoe UI', size=14, bold=True, color=COLOR_WHITE)
font_section   = Font(name='Segoe UI', size=11, bold=True, color=COLOR_WHITE)
font_header    = Font(name='Segoe UI', size=10, bold=True, color=COLOR_WHITE)

fill_title     = PatternFill(start_color=COLOR_TITLE_BG, end_color=COLOR_TITLE_BG, fill_type='solid')
fill_section   = PatternFill(start_color=COLOR_SECTION_BG, end_color=COLOR_SECTION_BG, fill_type='solid')
fill_header    = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type='solid')

# Load trades from v13 Master
xl_v13 = pd.ExcelFile(V13_PATH, engine='openpyxl')
df_all_trades = xl_v13.parse('All_12Q_Futures_Trades')
df_qtr_summary = xl_v13.parse('Exec_12Q_Combined_Summary')

trade_no_col = next(c for c in df_all_trades.columns if 'Trade' in c)
pnl_col = next(c for c in df_all_trades.columns if 'P&L' in c or 'Profit' in c)

quarters_order = [
    'Q3 2023-24', 'Q4 2023-24', 'Q1 2024-25', 'Q2 2024-25',
    'Q3 2024-25', 'Q4 2024-25', 'Q1 2025-26', 'Q2 2025-26',
    'Q3 2025-26', 'Q4 2025-26', 'Q1 2026-27', 'Q2 2026-27'
]

qtr_date_ranges = {
    'Q3 2023-24': 'Oct 2023 - Dec 2023',
    'Q4 2023-24': 'Jan 2024 - Mar 2024',
    'Q1 2024-25': 'Apr 2024 - Jun 2024',
    'Q2 2024-25': 'Jul 2024 - Sep 2024',
    'Q3 2024-25': 'Oct 2024 - Dec 2024',
    'Q4 2024-25': 'Jan 2025 - Mar 2025',
    'Q1 2025-26': 'Apr 2025 - Jun 2025',
    'Q2 2025-26': 'Jul 2025 - Sep 2025',
    'Q3 2025-26': 'Oct 2025 - Dec 2025',
    'Q4 2025-26': 'Jan 2026 - Mar 2026',
    'Q1 2026-27': 'Apr 2026 - Jun 2026',
    'Q2 2026-27': 'Jul 2026 - Sep 2026'
}

print("Building CHARTS-ONLY Master Excel Workbook...", flush=True)
wb = Workbook()
wb.remove(wb.active)

# ------------------- Hidden Data Sheet (Chart_Data) -------------------
ws_data = wb.create_sheet(title='Chart_Data')

# Prepare Master Summary Data in Chart_Data
ws_data.cell(row=1, column=1, value='Quarter')
ws_data.cell(row=1, column=2, value='Net Futures P&L (₹)')

q_summaries_list = []
for r_idx, qtr in enumerate(quarters_order):
    q_trades = df_all_trades[df_all_trades['Quarter'] == qtr].copy()
    q_trades['Entry Date DT'] = pd.to_datetime(q_trades['Entry Date'])
    q_trades = q_trades.sort_values(by=['Entry Date DT', trade_no_col]).reset_index(drop=True)
    
    q_trades['Cum PnL'] = q_trades[pnl_col].cumsum()
    q_trades['Running Max'] = np.maximum.accumulate(q_trades['Cum PnL'])
    q_trades['Drawdown'] = q_trades['Cum PnL'] - q_trades['Running Max']
    
    max_dd_val = abs(q_trades['Drawdown'].min())
    net_pnl = q_trades[pnl_col].sum()
    
    q_summaries_list.append({
        'Quarter': qtr,
        'Net PnL': net_pnl,
        'Max DD': max_dd_val,
        'Trades DF': q_trades
    })
    
    ws_data.cell(row=2 + r_idx, column=1, value=qtr)
    ws_data.cell(row=2 + r_idx, column=2, value=net_pnl)

# Write Per-Quarter Trade Series into Chart_Data
q_data_col_map = {}

curr_col = 4
for q_info in q_summaries_list:
    qtr = q_info['Quarter']
    df_q = q_info['Trades DF']
    
    ws_data.cell(row=1, column=curr_col, value=f"{qtr}_Trade#")
    ws_data.cell(row=1, column=curr_col+1, value="Cumulative Realised P&L (₹)")
    ws_data.cell(row=1, column=curr_col+2, value="Drawdown (₹)")
    
    q_data_col_map[qtr] = {
        'trade_col': curr_col,
        'pnl_col': curr_col + 1,
        'dd_col': curr_col + 2,
        'count': len(df_q)
    }
    
    for idx, r in df_q.iterrows():
        row_num = 2 + idx
        ws_data.cell(row=row_num, column=curr_col, value=idx + 1)
        ws_data.cell(row=row_num, column=curr_col+1, value=r['Cum PnL'])
        ws_data.cell(row=row_num, column=curr_col+2, value=r['Drawdown'])
        
    curr_col += 4

# Hide Chart_Data Sheet from view
ws_data.sheet_state = 'hidden'

# ------------------- 1. SHEET 1: Master_Executive_Dashboard (Grid of All Charts) -------------------
ws_dash = wb.create_sheet(title='Master_Executive_Dashboard')
ws_dash.views.sheetView[0].showGridLines = True

# Title Banner
ws_dash.merge_cells('A1:R1')
c_title = ws_dash['A1']
c_title.value = "NIFTY 50 STRATEGY — 12-QUARTER DRAWDOWN & EQUITY CURVE DASHBOARD"
c_title.font = font_title
c_title.fill = fill_title
c_title.alignment = Alignment(horizontal='center', vertical='center')
ws_dash.row_dimensions[1].height = 40

# Add Master 12-Quarter Cumulative Chart at Top
chart_master = LineChart()
chart_master.title = "Master 12-Quarter Cumulative Net Futures P&L (₹)"
chart_master.style = 13
chart_master.y_axis.title = "Net P&L (₹)"
chart_master.x_axis.title = "Quarter"
chart_master.width = 30
chart_master.height = 12

data_m = Reference(ws_data, min_col=2, min_row=1, max_row=13)
cats_m = Reference(ws_data, min_col=1, min_row=2, max_row=13)
chart_master.add_data(data_m, titles_from_data=True)
chart_master.set_categories(cats_m)
ws_dash.add_chart(chart_master, "A3")

# Layout 12 Individual Quarter Charts in a clean 3x4 Grid starting at Row 17
grid_positions = [
    "A17", "G17", "M17",
    "A30", "G30", "M30",
    "A43", "G43", "M43",
    "A56", "G56", "M56"
]

for q_idx, q_info in enumerate(q_summaries_list):
    qtr = q_info['Quarter']
    max_dd = q_info['Max DD']
    cmap = q_data_col_map[qtr]
    cell_anchor = grid_positions[q_idx]
    
    chart_q = LineChart()
    chart_q.title = f"{qtr} — Drawdown Curve (Max DD: ₹{max_dd:,.2f})"
    chart_q.style = 13
    chart_q.y_axis.title = "Rupees (₹)"
    chart_q.x_axis.title = "Trade #"
    chart_q.width = 14
    chart_q.height = 8.5
    
    max_r = 1 + cmap['count']
    data_q = Reference(ws_data, min_col=cmap['pnl_col'], min_row=1, max_col=cmap['dd_col'], max_row=max_r)
    cats_q = Reference(ws_data, min_col=cmap['trade_col'], min_row=2, max_row=max_r)
    chart_q.add_data(data_q, titles_from_data=True)
    chart_q.set_categories(cats_q)
    
    ws_dash.add_chart(chart_q, cell_anchor)

# ------------------- 2. SHEETS 2-13: Individual Quarter Sheets (Charts ONLY) -------------------
for q_info in q_summaries_list:
    qtr = q_info['Quarter']
    max_dd = q_info['Max DD']
    net_pnl = q_info['Net PnL']
    cmap = q_data_col_map[qtr]
    dt_range = qtr_date_ranges.get(qtr, '')
    
    ws_q = wb.create_sheet(title=qtr)
    ws_q.views.sheetView[0].showGridLines = True
    
    # Title Banner
    ws_q.merge_cells('A1:P1')
    c_t = ws_q['A1']
    c_t.value = f"NIFTY 50 STRATEGY — {qtr.upper()} ({dt_range.upper()}) | NET P&L: ₹{net_pnl:,.2f} | MAX DD: ₹{max_dd:,.2f}"
    c_t.font = font_title
    c_t.fill = fill_title
    c_t.alignment = Alignment(horizontal='center', vertical='center')
    ws_q.row_dimensions[1].height = 38
    
    # Large High-Resolution Chart filling the entire page
    chart_large = LineChart()
    chart_large.title = f"{qtr} — Cumulative Realised Futures P&L & Drawdown Curve (Max DD: ₹{max_dd:,.2f})"
    chart_large.style = 13
    chart_large.y_axis.title = "Rupees (₹)"
    chart_large.x_axis.title = "Trade #"
    chart_large.width = 28
    chart_large.height = 14
    
    max_r = 1 + cmap['count']
    data_q = Reference(ws_data, min_col=cmap['pnl_col'], min_row=1, max_col=cmap['dd_col'], max_row=max_r)
    cats_q = Reference(ws_data, min_col=cmap['trade_col'], min_row=2, max_row=max_r)
    chart_large.add_data(data_q, titles_from_data=True)
    chart_large.set_categories(cats_q)
    
    ws_q.add_chart(chart_large, "A3")

# Save Workbook
print(f"Saving CHARTS-ONLY workbook to {OUTPUT_CHARTS_ONLY.name}...", flush=True)
wb.save(OUTPUT_CHARTS_ONLY)

print(f"\n🎉 CHARTS-ONLY EXCEL WORKBOOK CREATED IN {time.time() - t0:.1f}s!", flush=True)
