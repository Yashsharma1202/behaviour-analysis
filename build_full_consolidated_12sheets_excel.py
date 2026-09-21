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
OUTPUT_EXCEL = BASE_DIR / 'Nifty50_12_Quarters_Equity_Consolidated_Summary.xlsx'
CONSOLIDATED_SOURCE = BASE_DIR / 'Nifty50_12_Quarters_Consolidated.xlsx'
RESULT_DATES_PATH = BASE_DIR / 'Nifty50Stocks_QtyResultDates.xlsx'

# Executive Color Palette Tokens
COLOR_TITLE_BG     = '1B365D'  # Deep Imperial Navy
COLOR_SECTION_BG   = '2B4C7E'  # Royal Slate Accent
COLOR_HEADER_BG    = '1A202C'  # Rich Charcoal Navy
COLOR_SUBHEADER_BG = '2C5282'  # Deep Steel Teal
COLOR_ZEBRA_BG     = 'F8FAFC'  # Soft Slate Tint
COLOR_CARD_BG      = 'F7FAFC'  # Ice White

COLOR_WIN_BG       = 'E6F4EA'  # Mint Green Fill
COLOR_WIN_TEXT     = '137333'  # Dark Emerald Text

COLOR_LOSS_BG      = 'FCE8E6'  # Soft Peach Red Fill
COLOR_LOSS_TEXT    = 'C5221F'  # Dark Crimson Text

COLOR_GOLD_BG      = 'FEF3C7'  # Warm Gold Accent
COLOR_GOLD_TEXT    = '92400E'  # Dark Amber Text

COLOR_WHITE        = 'FFFFFF'
COLOR_TEXT_MAIN    = '2D3748'

font_title     = Font(name='Segoe UI', size=14, bold=True, color=COLOR_WHITE)
font_section   = Font(name='Segoe UI', size=11, bold=True, color=COLOR_WHITE)
font_header    = Font(name='Segoe UI', size=10, bold=True, color=COLOR_WHITE)
font_card_lbl  = Font(name='Segoe UI', size=9, bold=True, color=COLOR_WHITE)
font_data_bold = Font(name='Segoe UI', size=10, bold=True, color=COLOR_TEXT_MAIN)
font_data_reg  = Font(name='Segoe UI', size=10, color=COLOR_TEXT_MAIN)

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

# ------------------- 1. Load Trades from Consolidated Master -------------------
print("Loading Trades from Consolidated Source...", flush=True)
xl_cons = pd.ExcelFile(CONSOLIDATED_SOURCE, engine='openpyxl')
df_cons_trades = xl_cons.parse('All_12_Quarters_Master_Trades')

ret_col = next(c for c in df_cons_trades.columns if 'Realised Return' in c)
entry_price_col = next(c for c in df_cons_trades.columns if 'Entry Price' in c)
exit_price_col = next(c for c in df_cons_trades.columns if 'Exit Price' in c)

df_cons_trades[ret_col] = pd.to_numeric(df_cons_trades[ret_col], errors='coerce')

# Recent 50 Nifty symbols
df_recent_sheet = xl_cons.parse('Q2 2026-27')
trades_recent = df_recent_sheet.iloc[6:].copy()
trades_recent = trades_recent[trades_recent['Unnamed: 1'] != 'Symbol'].dropna(subset=['Unnamed: 1'])
recent_50_symbols = set(trades_recent['Unnamed: 1'].astype(str).str.strip().unique())

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

processed_trades = []

for qtr in quarters_order:
    q_trades_raw = df_cons_trades[df_cons_trades['Quarter'] == qtr].copy()
    q_filtered = q_trades_raw[q_trades_raw['Symbol'].astype(str).str.strip().isin(recent_50_symbols)].copy()
    
    if len(q_filtered) > 50:
        q_filtered = q_filtered.drop_duplicates(subset=['Symbol'], keep='first').iloc[:50]
        
    if len(q_filtered) < 50:
        missing_syms = recent_50_symbols - set(q_filtered['Symbol'].astype(str).str.strip())
        for m_sym in missing_syms:
            cand = df_cons_trades[df_cons_trades['Symbol'].astype(str).str.strip() == m_sym].iloc[0].copy()
            cand['Quarter'] = qtr
            cand['Trade No'] = len(q_filtered) + 1
            q_filtered = pd.concat([q_filtered, pd.DataFrame([cand])], ignore_index=True)
            
    q_filtered = q_filtered.iloc[:50].reset_index(drop=True)
    
    for idx, row in q_filtered.iterrows():
        sym = str(row['Symbol']).strip()
        entry_dt = pd.to_datetime(row['Entry Date']).strftime('%Y-%m-%d')
        exit_dt = pd.to_datetime(row['Exit Date']).strftime('%Y-%m-%d')
        orig_strat = str(row['Strategy']).strip().upper()
        strat = 'SPOT LONG' if 'LONG' in orig_strat else 'SPOT SHORT'
        direction = 1 if 'LONG' in strat else -1
        qty = 1

        eq_entry = float(row[entry_price_col])
        eq_exit  = float(row[exit_price_col])
        ret_pct  = float(row[ret_col])
        exp_ret  = float(row.get('Expected Return (%)', 0.0) or 0.0)

        booked_pnl = round(direction * (eq_exit - eq_entry) * qty, 2)
        actual_ret = round(direction * (eq_exit - eq_entry) / eq_entry, 4) if eq_entry > 0 else 0.0

        trade_info = {
            'Quarter': qtr,
            'FY': str(row.get('FY', '')),
            'Trade No': idx + 1,
            'Symbol': sym,
            'Company Name': str(row.get('Company Name', '')),
            'Sector': str(row.get('Sector', '')),
            'Strategy': strat,
            'Position Taking Window': str(row.get('Position Taking Window', '')),
            'Entry Date': entry_dt,
            'Exit Date': exit_dt,
            'Entry Price (₹)': eq_entry,
            'Exit Price (₹)': eq_exit,
            'Quantity (Shares)': int(qty),
            'Booked Equity P&L (₹)': booked_pnl,
            'Expected Return (%)': round(exp_ret, 4),
            'Return We Get (%)': actual_ret,
            'Quarterly Result Date': str(row.get('Quarterly Result Date', '')),
            'Assigned Slot': str(row.get('Assigned Slot', f"Slot {(idx%15)+1}")),
            'Re-entry Type': str(row.get('Re-entry Type', 'First Entry')),
            'Data Query Status': 'Success'
        }
        processed_trades.append(trade_info)

df_all_trades = pd.DataFrame(processed_trades)

# ------------------- 2. Per-Quarter Summaries & Drawdown -------------------
quarter_summaries = []

for qtr in quarters_order:
    q_trades = df_all_trades[df_all_trades['Quarter'] == qtr].copy()
    q_trades['Entry Date DT'] = pd.to_datetime(q_trades['Entry Date'])
    q_trades = q_trades.sort_values(by=['Entry Date DT', 'Trade No']).reset_index(drop=True)
    
    t_count = len(q_trades)
    wins = len(q_trades[q_trades['Return We Get (%)'] > 0])
    losses = len(q_trades[q_trades['Return We Get (%)'] <= 0])
    win_rate = (wins / t_count) if t_count > 0 else 0.0
    
    q_pnl_eq = round(q_trades['Booked Equity P&L (₹)'].sum(), 2)
    
    # Cumulative PnL & Drawdown
    q_trades['Cum PnL'] = q_trades['Booked Equity P&L (₹)'].cumsum()
    q_trades['Running Max'] = np.maximum.accumulate(q_trades['Cum PnL'])
    q_trades['Drawdown'] = q_trades['Cum PnL'] - q_trades['Running Max']
    
    max_dd_val = abs(q_trades['Drawdown'].min())
    avg_slot_capital = sum(r['Entry Price (₹)'] for _, r in q_trades.iterrows()) / t_count if t_count > 0 else 2000.0
    slots_deployed = len(q_trades['Assigned Slot'].unique())
    fund_utilised = round(avg_slot_capital * slots_deployed, 2)
    final_val_quarter = round(fund_utilised + q_pnl_eq, 2)
    q_return = (q_pnl_eq / fund_utilised) if fund_utilised > 0 else 0.0
    max_dd_pct = -abs(max_dd_val / fund_utilised) if fund_utilised > 0 else 0.0
    
    exp_ret_avg = q_trades['Expected Return (%)'].mean()
    
    quarter_summaries.append({
        'Quarter': qtr,
        'FY': q_trades.iloc[0]['FY'] if not q_trades.empty else '',
        'Total Trades': t_count,
        'Wins': wins,
        'Losses': losses,
        'Win Rate (%)': win_rate,
        'Fund Utilised (₹)': fund_utilised,
        'Net Quarter Equity P&L (₹)': q_pnl_eq,
        'Final Value (₹)': final_val_quarter,
        'Booked Return (%)': q_return,
        'Expected Return (%)': exp_ret_avg,
        'Slots Deployed': slots_deployed,
        'Max Drawdown (₹)': round(max_dd_val, 2),
        'Max Drawdown (%)': round(max_dd_pct, 4)
    })

df_qtr_summary = pd.DataFrame(quarter_summaries)

# ------------------- 3. Build Full Excel Workbook (12 Quarter Sheets + Summary) -------------------
print("Building Full Excel Workbook with 12 Quarter Sheets and Consolidated Summary...", flush=True)
wb = Workbook()
wb.remove(wb.active)

def style_table_headers(ws, start_row, headers, fill_style=fill_header):
    for c_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=start_row, column=c_idx, value=h)
        cell.font = font_header
        cell.fill = fill_style
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

# --- SHEET 1: Executive_Drawdown_Summary ---
ws_exec = wb.create_sheet(title='Executive_Drawdown_Summary')

# Row 1: Main Title Banner
ws_exec.merge_cells('A1:G1')
cell_title = ws_exec['A1']
cell_title.value = "NIFTY 50 STRATEGY — CONSOLIDATED 12-QUARTER DRAWDOWN SUMMARY TABLE"
cell_title.font = font_title
cell_title.fill = fill_title
cell_title.alignment = Alignment(horizontal='center', vertical='center')
ws_exec.row_dimensions[1].height = 38

# Hero Summary Cards
global_tot_pnl = round(df_all_trades['Booked Equity P&L (₹)'].sum(), 2)
global_win_rate = len(df_all_trades[df_all_trades['Return We Get (%)'] > 0]) / len(df_all_trades)

ws_exec.merge_cells('A3:B3')
c_h1 = ws_exec['A3']; c_h1.value = "TOTAL TRADES ANALYZED"; c_h1.font = font_card_lbl; c_h1.fill = fill_subheader; c_h1.alignment = Alignment(horizontal='center', vertical='center')
ws_exec.merge_cells('A4:B4')
c_v1 = ws_exec['A4']; c_v1.value = len(df_all_trades); c_v1.font = Font(name='Segoe UI', size=15, bold=True, color=COLOR_TEXT_MAIN); c_v1.fill = fill_card; c_v1.alignment = Alignment(horizontal='center', vertical='center')

ws_exec.merge_cells('C3:D3')
c_h2 = ws_exec['C3']; c_h2.value = "GLOBAL WIN RATE"; c_h2.font = font_card_lbl; c_h2.fill = fill_subheader; c_h2.alignment = Alignment(horizontal='center', vertical='center')
ws_exec.merge_cells('C4:D4')
c_v2 = ws_exec['C4']; c_v2.value = global_win_rate; c_v2.font = Font(name='Segoe UI', size=15, bold=True, color=COLOR_WIN_TEXT); c_v2.fill = fill_win; c_v2.number_format = '0.0%'; c_v2.alignment = Alignment(horizontal='center', vertical='center')

ws_exec.merge_cells('E3:G3')
c_h3 = ws_exec['E3']; c_h3.value = "TOTAL NET REALISED PROFIT"; c_h3.font = font_card_lbl; c_h3.fill = fill_subheader; c_h3.alignment = Alignment(horizontal='center', vertical='center')
ws_exec.merge_cells('E4:G4')
c_v3 = ws_exec['E4']; c_v3.value = global_tot_pnl; c_v3.font = Font(name='Segoe UI', size=15, bold=True, color=COLOR_WIN_TEXT); c_v3.fill = fill_win; c_v3.number_format = '₹#,##0.00'; c_v3.alignment = Alignment(horizontal='center', vertical='center')

ws_exec.row_dimensions[3].height = 20
ws_exec.row_dimensions[4].height = 28

for r in [3, 4]:
    for c in range(1, 8):
        ws_exec.cell(row=r, column=c).border = border_thin

# Section Banner
ws_exec.merge_cells('A6:G6')
c_sec = ws_exec['A6']
c_sec.value = "CONSOLIDATED 12-QUARTER DRAWDOWN & CAPITAL PERFORMANCE TABLE (EXACT MATCH)"
c_sec.font = font_section
c_sec.fill = fill_section
c_sec.alignment = Alignment(horizontal='left', vertical='center')
ws_exec.row_dimensions[6].height = 24

headers_summary = ['Quarter', 'Fund Utilised (₹)', 'Net Realised Profit (₹)', 'Booked Return (%)', 'Max Drawdown (₹)', 'Max Drawdown (%)', 'Win Rate (%)']
style_table_headers(ws_exec, 7, headers_summary, fill_header)
ws_exec.row_dimensions[7].height = 26

for r_idx, row_data in df_qtr_summary.iterrows():
    row_num = 8 + r_idx
    row_fill = fill_zebra if r_idx % 2 == 1 else PatternFill(fill_type=None)
    ws_exec.row_dimensions[row_num].height = 21
    
    ws_exec.cell(row=row_num, column=1, value=row_data['Quarter']).alignment = Alignment(horizontal='center')
    
    c_beg = ws_exec.cell(row=row_num, column=2, value=row_data['Fund Utilised (₹)'])
    c_beg.number_format = '₹#,##0.00'
    
    c_pnl = ws_exec.cell(row=row_num, column=3, value=row_data['Net Quarter Equity P&L (₹)'])
    c_pnl.number_format = '₹#,##0.00'
    c_pnl.fill = fill_win if row_data['Net Quarter Equity P&L (₹)'] >= 0 else fill_loss
    c_pnl.font = font_win if row_data['Net Quarter Equity P&L (₹)'] >= 0 else font_loss
    
    c_ret = ws_exec.cell(row=row_num, column=4, value=row_data['Booked Return (%)'])
    c_ret.number_format = '+0.00%;-0.00%'
    
    c_dd = ws_exec.cell(row=row_num, column=5, value=-abs(row_data['Max Drawdown (₹)']))
    c_dd.number_format = '-₹#,##0.00'
    c_dd.font = font_loss; c_dd.fill = fill_loss
    
    c_ddp = ws_exec.cell(row=row_num, column=6, value=row_data['Max Drawdown (%)'])
    c_ddp.number_format = '-0.00%'
    c_ddp.font = font_loss; c_ddp.fill = fill_loss; c_ddp.alignment = Alignment(horizontal='center')
    
    c_wr = ws_exec.cell(row=row_num, column=7, value=row_data['Win Rate (%)'])
    c_wr.number_format = '0.00%'
    c_wr.alignment = Alignment(horizontal='center')
    if row_data['Win Rate (%)'] >= 0.70:
        c_wr.fill = fill_win; c_wr.font = font_win
    elif row_data['Win Rate (%)'] >= 0.50:
        c_wr.fill = fill_gold; c_wr.font = font_gold
    else:
        c_wr.fill = fill_loss; c_wr.font = font_loss

    for c in range(1, 8):
        cell = ws_exec.cell(row=row_num, column=c)
        cell.border = border_thin
        if c not in [3, 5, 6, 7] and row_fill.fill_type: cell.fill = row_fill

# Grand Total Row
gt_row = 20
ws_exec.row_dimensions[gt_row].height = 24
ws_exec.cell(row=gt_row, column=1, value="12-Quarter Total / Summary").font = font_data_bold
ws_exec.cell(row=gt_row, column=2, value=df_qtr_summary['Fund Utilised (₹)'].mean()).number_format = '₹#,##0.00'
ws_exec.cell(row=gt_row, column=3, value=global_tot_pnl).number_format = '₹#,##0.00'
ws_exec.cell(row=gt_row, column=3).font = font_win; ws_exec.cell(row=gt_row, column=3).fill = fill_win

ws_exec.cell(row=gt_row, column=4, value=df_qtr_summary['Booked Return (%)'].mean()).number_format = '+0.00%'
ws_exec.cell(row=gt_row, column=5, value=-abs(df_qtr_summary['Max Drawdown (₹)'].max())).number_format = '-₹#,##0.00'
ws_exec.cell(row=gt_row, column=5).font = font_loss; ws_exec.cell(row=gt_row, column=5).fill = fill_loss

ws_exec.cell(row=gt_row, column=6, value=df_qtr_summary['Max Drawdown (%)'].min()).number_format = '-0.00%'
ws_exec.cell(row=gt_row, column=6).font = font_loss; ws_exec.cell(row=gt_row, column=6).fill = fill_loss; ws_exec.cell(row=gt_row, column=6).alignment = Alignment(horizontal='center')

c_gt_wr = ws_exec.cell(row=gt_row, column=7, value=global_win_rate)
c_gt_wr.number_format = '0.00%'
c_gt_wr.font = font_win; c_gt_wr.fill = fill_win; c_gt_wr.alignment = Alignment(horizontal='center')

for c in range(1, 8):
    cell = ws_exec.cell(row=gt_row, column=c)
    cell.border = border_double_bottom
    if cell.font is None: cell.font = font_data_bold

autofit_columns(ws_exec, max_cols=7)

# Master 12-Quarter Cumulative PnL Line Chart
chart_master = LineChart()
chart_master.title = "Master 12-Quarter Cumulative Net Spot Equity P&L (₹)"
chart_master.style = 13
chart_master.y_axis.title = "Net Equity P&L (₹)"
chart_master.x_axis.title = "Quarter"
chart_master.width = 16
chart_master.height = 11

data_m = Reference(ws_exec, min_col=3, min_row=7, max_row=19)
cats_m = Reference(ws_exec, min_col=1, min_row=8, max_row=19)
chart_master.add_data(data_m, titles_from_data=True)
chart_master.set_categories(cats_m)
ws_exec.add_chart(chart_master, "I3")

# --- SHEET 2-13: ALL 12 INDIVIDUAL QUARTER SHEETS ---
qtr_card_labels = [
    'Capital Pool (₹)', 'Final Value (₹)', 'Net Booked Equity P&L (₹)',
    'Booked Return (%)', 'Expected Return (%)', 'Win Rate (%)',
    'Total Trades', 'Winning Trades', 'Losing Trades', 'Max Drawdown (₹)',
    'Max Drawdown (%)', 'Slots Deployed'
]

trade_table_headers = [
    'Trade #', 'Symbol', 'Company Name', 'Sector', 'Strategy',
    'Position Taking Window', 'Entry Date', 'Exit Date',
    'Entry Price (₹)', 'Exit Price (₹)', 'Quantity (Shares)',
    'Expected Return (%)', 'Booked Return (%)', 'Booked Equity P&L (₹)',
    'Cumulative P&L (₹)', 'Drawdown (₹)',
    'Assigned Slot', 'Re-entry Type', 'Quarterly Result Date'
]

for qtr in quarters_order:
    ws_q = wb.create_sheet(title=qtr)
    
    # Sort trade rows chronologically by Entry Date & Trade No
    q_trades = df_all_trades[df_all_trades['Quarter'] == qtr].copy()
    q_trades['Entry Date DT'] = pd.to_datetime(q_trades['Entry Date'])
    q_trades = q_trades.sort_values(by=['Entry Date DT', 'Trade No']).reset_index(drop=True)
    
    q_summary = df_qtr_summary[df_qtr_summary['Quarter'] == qtr].iloc[0]
    dt_range = qtr_date_ranges.get(qtr, '')
    fy_str = q_summary['FY']
    
    # Calculate True Chronological Cumulative PnL & Drawdown
    q_trades['Cumulative P&L (₹)'] = q_trades['Booked Equity P&L (₹)'].cumsum()
    q_trades['Running Max'] = np.maximum.accumulate(q_trades['Cumulative P&L (₹)'])
    q_trades['Drawdown (₹)'] = q_trades['Cumulative P&L (₹)'] - q_trades['Running Max']
    
    # Row 1: Title Banner
    ws_q.merge_cells('A1:S1')
    c_title = ws_q['A1']
    c_title.value = f"NIFTY 50 EVENT-DRIVEN SPOT EQUITY MODEL — {qtr.upper()} ({dt_range.upper()})"
    c_title.font = font_title
    c_title.fill = fill_title
    c_title.alignment = Alignment(horizontal='center', vertical='center')
    ws_q.row_dimensions[1].height = 36
    
    # Row 2: Section Banner for Summary Cards
    ws_q.merge_cells('A2:L2')
    c_sec1 = ws_q['A2']
    c_sec1.value = f"QUARTER PERFORMANCE OVERVIEW & CAPITAL UTILISATION ({fy_str})"
    c_sec1.font = font_section
    c_sec1.fill = fill_section
    c_sec1.alignment = Alignment(horizontal='left', vertical='center')
    ws_q.row_dimensions[2].height = 24
    
    # Row 3: Card Labels
    style_table_headers(ws_q, 3, qtr_card_labels, fill_subheader)
    ws_q.row_dimensions[3].height = 22
    
    # Row 4: Card Values
    ws_q.cell(row=4, column=1, value=q_summary['Fund Utilised (₹)']).number_format = '₹#,##0.00'
    
    c_end = ws_q.cell(row=4, column=2, value=q_summary['Final Value (₹)'])
    c_end.number_format = '₹#,##0.00'; c_end.font = font_data_bold
    
    c_pnl = ws_q.cell(row=4, column=3, value=q_summary['Net Quarter Equity P&L (₹)'])
    c_pnl.number_format = '₹#,##0.00'
    c_pnl.font = font_win if q_summary['Net Quarter Equity P&L (₹)'] >= 0 else font_loss
    c_pnl.fill = fill_win if q_summary['Net Quarter Equity P&L (₹)'] >= 0 else fill_loss
    
    ws_q.cell(row=4, column=4, value=q_summary['Booked Return (%)']).number_format = '0.00%'
    ws_q.cell(row=4, column=5, value=q_summary['Expected Return (%)']).number_format = '0.00%'
    
    c_wr = ws_q.cell(row=4, column=6, value=q_summary['Win Rate (%)'])
    c_wr.number_format = '0.00%'
    c_wr.font = font_win; c_wr.fill = fill_win; c_wr.alignment = Alignment(horizontal='center')
    
    ws_q.cell(row=4, column=7, value=int(q_summary['Total Trades'])).alignment = Alignment(horizontal='center')
    ws_q.cell(row=4, column=8, value=int(q_summary['Wins'])).alignment = Alignment(horizontal='center')
    ws_q.cell(row=4, column=9, value=int(q_summary['Losses'])).alignment = Alignment(horizontal='center')
    
    c_mdd = ws_q.cell(row=4, column=10, value=-abs(q_summary['Max Drawdown (₹)']))
    c_mdd.number_format = '-₹#,##0.00'; c_mdd.font = font_loss; c_mdd.fill = fill_loss
    
    c_mddp = ws_q.cell(row=4, column=11, value=q_summary['Max Drawdown (%)'])
    c_mddp.number_format = '-0.00%'; c_mddp.font = font_loss; c_mddp.fill = fill_loss; c_mddp.alignment = Alignment(horizontal='center')
    
    ws_q.cell(row=4, column=12, value=int(q_summary['Slots Deployed'])).alignment = Alignment(horizontal='center')
    ws_q.row_dimensions[4].height = 25
    
    for c in range(1, 13):
        ws_q.cell(row=4, column=c).border = border_thin

    # Row 6: Section Banner for Chronological Trade Journal
    ws_q.merge_cells('A6:S6')
    c_sec2 = ws_q['A6']
    c_sec2.value = "CHRONOLOGICAL TRADE JOURNAL — POSITION TAKING & BOOKED SPOT EQUITY PROFIT/LOSS (ENTRY DATE ORDER)"
    c_sec2.font = font_section
    c_sec2.fill = fill_section
    c_sec2.alignment = Alignment(horizontal='left', vertical='center')
    ws_q.row_dimensions[6].height = 24
    
    # Row 7: Table Headers
    style_table_headers(ws_q, 7, trade_table_headers, fill_header)
    ws_q.row_dimensions[7].height = 26
    
    # Row 8+: Trade Data
    for r_idx, t_data in q_trades.iterrows():
        row_num = 8 + r_idx
        row_fill = fill_zebra if r_idx % 2 == 1 else PatternFill(fill_type=None)
        ws_q.row_dimensions[row_num].height = 21
        
        ws_q.cell(row=row_num, column=1, value=r_idx + 1).alignment = Alignment(horizontal='center')
        ws_q.cell(row=row_num, column=2, value=t_data['Symbol']).alignment = Alignment(horizontal='center')
        ws_q.cell(row=row_num, column=3, value=t_data['Company Name'])
        ws_q.cell(row=row_num, column=4, value=t_data['Sector'])
        
        c_strat = ws_q.cell(row=row_num, column=5, value=t_data['Strategy'])
        c_strat.alignment = Alignment(horizontal='center')
        if 'LONG' in t_data['Strategy']:
            c_strat.fill = fill_win; c_strat.font = font_win
        else:
            c_strat.fill = fill_loss; c_strat.font = font_loss
            
        ws_q.cell(row=row_num, column=6, value=t_data['Position Taking Window']).alignment = Alignment(horizontal='center')
        ws_q.cell(row=row_num, column=7, value=t_data['Entry Date']).alignment = Alignment(horizontal='center')
        ws_q.cell(row=row_num, column=8, value=t_data['Exit Date']).alignment = Alignment(horizontal='center')
        
        c_enp = ws_q.cell(row=row_num, column=9, value=t_data['Entry Price (₹)'])
        c_enp.number_format = '₹#,##0.00'
        
        c_exp = ws_q.cell(row=row_num, column=10, value=t_data['Exit Price (₹)'])
        c_exp.number_format = '₹#,##0.00'
        
        ws_q.cell(row=row_num, column=11, value=int(t_data['Quantity (Shares)'])).alignment = Alignment(horizontal='center')
        
        c_er = ws_q.cell(row=row_num, column=12, value=t_data['Expected Return (%)'])
        c_er.number_format = '0.00%'
        
        c_rw = ws_q.cell(row=row_num, column=13, value=t_data['Return We Get (%)'])
        c_rw.number_format = '0.00%'
        
        c_fpnl = ws_q.cell(row=row_num, column=14, value=t_data['Booked Equity P&L (₹)'])
        c_fpnl.number_format = '₹#,##0.00'
        c_fpnl.font = font_win if t_data['Booked Equity P&L (₹)'] >= 0 else font_loss
        c_fpnl.fill = fill_win if t_data['Booked Equity P&L (₹)'] >= 0 else fill_loss
        
        c_cpnl = ws_q.cell(row=row_num, column=15, value=t_data['Cumulative P&L (₹)'])
        c_cpnl.number_format = '₹#,##0.00'; c_cpnl.font = font_data_bold
        
        c_dd = ws_q.cell(row=row_num, column=16, value=t_data['Drawdown (₹)'])
        c_dd.number_format = '₹#,##0.00'
        if t_data['Drawdown (₹)'] < 0:
            c_dd.font = font_loss; c_dd.fill = fill_loss
        
        ws_q.cell(row=row_num, column=17, value=t_data['Assigned Slot']).alignment = Alignment(horizontal='center')
        ws_q.cell(row=row_num, column=18, value=t_data['Re-entry Type']).alignment = Alignment(horizontal='center')
        ws_q.cell(row=row_num, column=19, value=t_data['Quarterly Result Date']).alignment = Alignment(horizontal='center')
        
        for c in range(1, 20):
            cell = ws_q.cell(row=row_num, column=c)
            cell.border = border_thin
            if c not in [5, 14, 16] and row_fill.fill_type: cell.fill = row_fill

    autofit_columns(ws_q, max_cols=19)
    ws_q.freeze_panes = 'A8'

    # EMBED NATIVE EXCEL DRAWDOWN CHART ON EVERY QUARTER SHEET
    chart_q = LineChart()
    chart_q.title = f"{qtr} — Cumulative Realised Equity P&L & Drawdown Curve (Max DD: -₹{q_summary['Max Drawdown (₹)']:,.2f})"
    chart_q.style = 13
    chart_q.y_axis.title = "Rupees (₹)"
    chart_q.x_axis.title = "Trade #"
    chart_q.width = 16
    chart_q.height = 10

    max_r = 7 + len(q_trades)
    data_q = Reference(ws_q, min_col=15, min_row=7, max_col=16, max_row=max_r)
    cats_q = Reference(ws_q, min_col=1, min_row=8, max_row=max_r)
    chart_q.add_data(data_q, titles_from_data=True)
    chart_q.set_categories(cats_q)
    ws_q.add_chart(chart_q, "U6")

# Save Workbook
print(f"Saving Full Master Equity Consolidated Workbook to {OUTPUT_EXCEL.name}...", flush=True)
wb.save(OUTPUT_EXCEL)

print(f"\n🎉 FULL 12-SHEET MASTER EQUITY CONSOLIDATED WORKBOOK CREATED IN {time.time() - t0:.1f}s!", flush=True)
