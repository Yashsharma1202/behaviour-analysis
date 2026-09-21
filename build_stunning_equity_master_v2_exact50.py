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
CONSOLIDATED_PATH = BASE_DIR / 'Nifty50_12_Quarters_Consolidated.xlsx'
RESULT_DATES_PATH = BASE_DIR / 'Nifty50Stocks_QtyResultDates.xlsx'

OUTPUT_EQUITY_MASTER = BASE_DIR / 'Nifty50_12_Quarters_Equity_Master_v2.xlsx'

# ------------------- Executive Color Palette Tokens -------------------
COLOR_TITLE_BG     = '1B365D'  # Deep Imperial Navy
COLOR_SECTION_BG   = '2B4C7E'  # Royal Slate Accent
COLOR_HEADER_BG    = '1A202C'  # Rich Charcoal Navy
COLOR_SUBHEADER_BG = '2C5282'  # Deep Steel Teal
COLOR_ZEBRA_BG     = 'F8FAFC'  # Ultra-soft Slate Tint
COLOR_CARD_BG      = 'F7FAFC'  # Clean Ice White

COLOR_WIN_BG       = 'E6F4EA'  # Mint Green Fill
COLOR_WIN_TEXT     = '137333'  # Dark Emerald Text

COLOR_LOSS_BG      = 'FCE8E6'  # Soft Peach Red Fill
COLOR_LOSS_TEXT    = 'C5221F'  # Dark Crimson Text

COLOR_GOLD_BG      = 'FEF3C7'  # Warm Gold Accent
COLOR_GOLD_TEXT    = '92400E'  # Dark Amber Text

COLOR_WHITE        = 'FFFFFF'
COLOR_TEXT_MAIN    = '2D3748'

# Fonts
font_title     = Font(name='Segoe UI', size=14, bold=True, color=COLOR_WHITE)
font_section   = Font(name='Segoe UI', size=11, bold=True, color=COLOR_WHITE)
font_header    = Font(name='Segoe UI', size=10, bold=True, color=COLOR_WHITE)
font_card_lbl  = Font(name='Segoe UI', size=9, bold=True, color=COLOR_WHITE)
font_data_bold = Font(name='Segoe UI', size=10, bold=True, color=COLOR_TEXT_MAIN)
font_data_reg  = Font(name='Segoe UI', size=10, color=COLOR_TEXT_MAIN)

font_win       = Font(name='Segoe UI', size=10, bold=True, color=COLOR_WIN_TEXT)
font_loss      = Font(name='Segoe UI', size=10, bold=True, color=COLOR_LOSS_TEXT)
font_gold      = Font(name='Segoe UI', size=10, bold=True, color=COLOR_GOLD_TEXT)

# Fills
fill_title     = PatternFill(start_color=COLOR_TITLE_BG, end_color=COLOR_TITLE_BG, fill_type='solid')
fill_section   = PatternFill(start_color=COLOR_SECTION_BG, end_color=COLOR_SECTION_BG, fill_type='solid')
fill_header    = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type='solid')
fill_subheader = PatternFill(start_color=COLOR_SUBHEADER_BG, end_color=COLOR_SUBHEADER_BG, fill_type='solid')
fill_zebra     = PatternFill(start_color=COLOR_ZEBRA_BG, end_color=COLOR_ZEBRA_BG, fill_type='solid')
fill_card      = PatternFill(start_color=COLOR_CARD_BG, end_color=COLOR_CARD_BG, fill_type='solid')

fill_win       = PatternFill(start_color=COLOR_WIN_BG, end_color=COLOR_WIN_BG, fill_type='solid')
fill_loss       = PatternFill(start_color=COLOR_LOSS_BG, end_color=COLOR_LOSS_BG, fill_type='solid')
fill_gold      = PatternFill(start_color=COLOR_GOLD_BG, end_color=COLOR_GOLD_BG, fill_type='solid')

# Borders
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

# ------------------- 1. Load Result Dates -------------------
print("Loading Quarterly Result Dates...", flush=True)
result_date_map = {}

if RESULT_DATES_PATH.exists():
    try:
        xl_rd = pd.ExcelFile(RESULT_DATES_PATH, engine='openpyxl')
        for sheet in xl_rd.sheet_names:
            df_rd = xl_rd.parse(sheet)
            sym_col = next((c for c in df_rd.columns if 'SYM' in str(c).upper()), None)
            date_col = next((c for c in df_rd.columns if 'DATE' in str(c).upper() or 'REPORTING' in str(c).upper()), None)
            if sym_col and date_col:
                for _, r in df_rd.iterrows():
                    sym = str(r[sym_col]).strip()
                    dt_val = r[date_col]
                    if pd.notnull(dt_val):
                        dt_str = pd.to_datetime(dt_val).strftime('%Y-%m-%d')
                        result_date_map[(sym, sheet)] = dt_str
    except Exception as e:
        print(f"Error loading result dates: {e}", flush=True)

print(f"Loaded {len(result_date_map)} quarterly result date mappings.", flush=True)

# ------------------- 2. Load Recent 50 Nifty Stock Symbols -------------------
xl_cons = pd.ExcelFile(CONSOLIDATED_PATH, engine='openpyxl')
df_recent_sheet = xl_cons.parse('Q2 2026-27')
trades_recent = df_recent_sheet.iloc[6:].copy()
trades_recent = trades_recent[trades_recent['Unnamed: 1'] != 'Symbol'].dropna(subset=['Unnamed: 1'])
recent_50_symbols = set(trades_recent['Unnamed: 1'].astype(str).str.strip().unique())

print(f"Official Recent Nifty 50 Count: {len(recent_50_symbols)} stocks.")

# ------------------- 3. Load & Filter Equity Trades to EXACTLY 50 per Quarter -------------------
print("Loading Trades from Consolidated Master...", flush=True)
df_cons_trades = xl_cons.parse('All_12_Quarters_Master_Trades')

ret_col = next(c for c in df_cons_trades.columns if 'Realised Return' in c)
entry_price_col = next(c for c in df_cons_trades.columns if 'Entry Price' in c)
exit_price_col = next(c for c in df_cons_trades.columns if 'Exit Price' in c)

df_cons_trades[ret_col] = pd.to_numeric(df_cons_trades[ret_col], errors='coerce')

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
    
    # Filter trades matching recent 50 symbols
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
        qty = 1  # 1 Share per stock in Spot Equity

        eq_entry = float(row[entry_price_col])
        eq_exit  = float(row[exit_price_col])
        ret_pct  = float(row[ret_col])
        exp_ret  = float(row.get('Expected Return (%)', 0.0) or 0.0)

        booked_pnl = round(direction * (eq_exit - eq_entry) * qty, 2)
        actual_ret = round(direction * (eq_exit - eq_entry) / eq_entry, 4) if eq_entry > 0 else 0.0
        
        result_dt = result_date_map.get((sym, qtr), '')

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
            'Quarterly Result Date': result_dt,
            'Assigned Slot': str(row.get('Assigned Slot', f"Slot {(idx%15)+1}")),
            'Re-entry Type': str(row.get('Re-entry Type', 'First Entry')),
            'Data Query Status': 'Success'
        }
        processed_trades.append(trade_info)

df_all_trades = pd.DataFrame(processed_trades)
print(f"Total Processed Equity Trades Across 12 Quarters: {len(df_all_trades)} (EXACTLY 50 PER QUARTER!)", flush=True)

# ------------------- 4. Per-Quarter Fund & Drawdown Summaries -------------------
quarter_summaries = []

for qtr in quarters_order:
    q_trades = df_all_trades[df_all_trades['Quarter'] == qtr].copy()
    q_trades['Entry Date DT'] = pd.to_datetime(q_trades['Entry Date'])
    q_trades['Exit Date DT']  = pd.to_datetime(q_trades['Exit Date'])
    q_trades = q_trades.sort_values(by=['Entry Date DT', 'Trade No']).reset_index(drop=True)
    
    t_count = len(q_trades)
    wins = len(q_trades[q_trades['Return We Get (%)'] > 0])
    losses = len(q_trades[q_trades['Return We Get (%)'] <= 0])
    longs = len(q_trades[q_trades['Strategy'] == 'SPOT LONG'])
    shorts = len(q_trades[q_trades['Strategy'] == 'SPOT SHORT'])
    win_rate = (wins / t_count) if t_count > 0 else 0.0
    
    q_pnl_eq = round(q_trades['Booked Equity P&L (₹)'].sum(), 2)
    
    # Calculate True Chronological Cumulative PnL & Drawdown (Entry Date Order)
    q_trades['Cum PnL'] = q_trades['Booked Equity P&L (₹)'].cumsum()
    q_trades['Running Max'] = np.maximum.accumulate(q_trades['Cum PnL'])
    q_trades['Drawdown'] = q_trades['Cum PnL'] - q_trades['Running Max']
    
    max_dd_val = abs(q_trades['Drawdown'].min()) if not q_trades.empty else 0.0
    
    avg_slot_capital = sum(r['Entry Price (₹)'] for _, r in q_trades.iterrows()) / t_count if t_count > 0 else 2000.0
    slots_deployed = len(q_trades['Assigned Slot'].unique())
    fund_utilised = round(avg_slot_capital * slots_deployed, 2)
    final_val_quarter = round(fund_utilised + q_pnl_eq, 2)
    q_return = (q_pnl_eq / fund_utilised) if fund_utilised > 0 else 0.0
    max_dd_pct = -abs(max_dd_val / fund_utilised) if fund_utilised > 0 else 0.0
    
    exp_ret_avg = q_trades['Expected Return (%)'].mean()
    re_entries = len(q_trades[q_trades['Re-entry Type'] != 'First Entry'])
    
    # Calculate Concurrent Trades
    min_date = q_trades['Entry Date DT'].min()
    max_date = q_trades['Exit Date DT'].max()
    all_b_dates = pd.date_range(min_date, max_date, freq='B')
    
    daily_conc = []
    for dt in all_b_dates:
        open_c = q_trades[(q_trades['Entry Date DT'] <= dt) & (q_trades['Exit Date DT'] >= dt)]
        daily_conc.append(len(open_c))
        
    max_concurrent = max(daily_conc) if daily_conc else 0
    avg_concurrent = round(float(np.mean(daily_conc)), 1) if daily_conc else 0.0
    
    quarter_summaries.append({
        'Quarter': qtr,
        'FY': q_trades.iloc[0]['FY'] if not q_trades.empty else '',
        'Total Trades': t_count,
        'Wins': wins,
        'Losses': losses,
        'Long Trades': longs,
        'Short Trades': shorts,
        'Win Rate (%)': win_rate,
        'Fund Utilised (₹)': fund_utilised,
        'Net Quarter Equity P&L (₹)': q_pnl_eq,
        'Final Value (₹)': final_val_quarter,
        'Booked Return (%)': q_return,
        'Expected Return (%)': exp_ret_avg,
        'Slots Deployed': slots_deployed,
        'Re-entries': re_entries,
        'Max Drawdown (₹)': round(max_dd_val, 2),
        'Max Drawdown (%)': round(max_dd_pct, 4),
        'Max Concurrent Scripts Open': max_concurrent,
        'Avg Concurrent Scripts Open': avg_concurrent
    })

df_qtr_summary = pd.DataFrame(quarter_summaries)

# ------------------- 5. Stock Leaderboard (1 Share Equity) -------------------
stock_leaderboard = []
for sym, group in df_all_trades.groupby('Symbol'):
    c_name = group.iloc[0]['Company Name']
    sector = group.iloc[0]['Sector']
    
    t_total = len(group)
    wins = len(group[group['Return We Get (%)'] > 0])
    losses = len(group[group['Return We Get (%)'] <= 0])
    long_wins = len(group[(group['Strategy'] == 'SPOT LONG') & (group['Return We Get (%)'] > 0)])
    short_wins = len(group[(group['Strategy'] == 'SPOT SHORT') & (group['Return We Get (%)'] > 0)])
    win_rate = (wins / t_total) if t_total > 0 else 0.0
    
    total_pnl = round(group['Booked Equity P&L (₹)'].sum(), 2)
    avg_ret = group['Return We Get (%)'].mean()
    
    prof_qtrs = 0
    for qtr in quarters_order:
        q_p = group[group['Quarter'] == qtr]['Booked Equity P&L (₹)'].sum()
        if q_p > 0: prof_qtrs += 1
        
    stock_leaderboard.append({
        'Symbol': sym,
        'Company Name': c_name,
        'Sector': sector,
        'Shares per Trade': 1,
        'Quarters Traded': len(group['Quarter'].unique()),
        'Profitable Quarters': prof_qtrs,
        'Wins': wins,
        'Losses': losses,
        'Long Wins': long_wins,
        'Short Wins': short_wins,
        'Win Rate (%)': win_rate,
        'Avg Return per Trade (%)': avg_ret,
        'Total 12Q Equity P&L (₹)': total_pnl
    })

df_leaderboard = pd.DataFrame(stock_leaderboard)
df_leaderboard = df_leaderboard.sort_values(by='Total 12Q Equity P&L (₹)', ascending=False).reset_index(drop=True)
df_leaderboard.insert(0, 'Rank', range(1, len(df_leaderboard) + 1))

# ------------------- 6. Losing Trades Audit -------------------
df_losing = df_all_trades[df_all_trades['Return We Get (%)'] <= 0].copy().reset_index(drop=True)

# ------------------- 7. Build Perfect Equity Master Excel Workbook -------------------
print("Building Master Equity Workbook (v2)...", flush=True)
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

# --- SHEET 1: Exec_12Q_Combined_Summary ---
ws_exec = wb.create_sheet(title='Exec_12Q_Combined_Summary')

# Row 1: Main Title Banner
ws_exec.merge_cells('A1:K1')
cell_title = ws_exec['A1']
cell_title.value = "NIFTY 50 EVENT-DRIVEN SPOT EQUITY STRATEGY — 12-QUARTERS EXECUTIVE SCORECARD"
cell_title.font = font_title
cell_title.fill = fill_title
cell_title.alignment = Alignment(horizontal='center', vertical='center')
ws_exec.row_dimensions[1].height = 38

# Row 3-4: Big Hero Summary Cards
global_tot_pnl = round(df_all_trades['Booked Equity P&L (₹)'].sum(), 2)
global_win_rate = len(df_all_trades[df_all_trades['Return We Get (%)'] > 0]) / len(df_all_trades)

ws_exec.merge_cells('A3:C3')
c_h1 = ws_exec['A3']
c_h1.value = "TOTAL TRADES ANALYZED"
c_h1.font = font_card_lbl
c_h1.fill = fill_subheader
c_h1.alignment = Alignment(horizontal='center', vertical='center')

ws_exec.merge_cells('A4:C4')
c_v1 = ws_exec['A4']
c_v1.value = len(df_all_trades)
c_v1.font = Font(name='Segoe UI', size=16, bold=True, color=COLOR_TEXT_MAIN)
c_v1.fill = fill_card
c_v1.alignment = Alignment(horizontal='center', vertical='center')

ws_exec.merge_cells('D3:F3')
c_h2 = ws_exec['D3']
c_h2.value = "GLOBAL WIN RATE"
c_h2.font = font_card_lbl
c_h2.fill = fill_subheader
c_h2.alignment = Alignment(horizontal='center', vertical='center')

ws_exec.merge_cells('D4:F4')
c_v2 = ws_exec['D4']
c_v2.value = global_win_rate
c_v2.font = Font(name='Segoe UI', size=16, bold=True, color=COLOR_WIN_TEXT)
c_v2.fill = fill_win
c_v2.number_format = '0.0%'
c_v2.alignment = Alignment(horizontal='center', vertical='center')

ws_exec.merge_cells('G3:K3')
c_h3 = ws_exec['G3']
c_h3.value = "TOTAL 12-QUARTER NET SPOT EQUITY PROFIT"
c_h3.font = font_card_lbl
c_h3.fill = fill_subheader
c_h3.alignment = Alignment(horizontal='center', vertical='center')

ws_exec.merge_cells('G4:K4')
c_v3 = ws_exec['G4']
c_v3.value = global_tot_pnl
c_v3.font = Font(name='Segoe UI', size=16, bold=True, color=COLOR_WIN_TEXT)
c_v3.fill = fill_win
c_v3.number_format = '₹#,##0.00'
c_v3.alignment = Alignment(horizontal='center', vertical='center')

ws_exec.row_dimensions[3].height = 20
ws_exec.row_dimensions[4].height = 30

for r in [3, 4]:
    for c in range(1, 12):
        ws_exec.cell(row=r, column=c).border = border_thin

# Row 6: Section Banner
ws_exec.merge_cells('A6:K6')
c_sec = ws_exec['A6']
c_sec.value = "QUARTER-BY-QUARTER PERFORMANCE COMPARISON & CHRONOLOGICAL DRAWDOWN (1-SHARE EQUITY MODEL)"
c_sec.font = font_section
c_sec.fill = fill_section
c_sec.alignment = Alignment(horizontal='left', vertical='center')
ws_exec.row_dimensions[6].height = 24

headers_summary = [
    'Quarter', 'Total Trades', 'Wins', 'Losses', 'Win Rate (%)',
    'Capital Pool Utilised (₹)', 'Net Quarter Equity P&L (₹)', 'Final Quarter Value (₹)',
    'Booked Return (%)', 'Max Drawdown (₹)', 'Max Drawdown (%)'
]
style_table_headers(ws_exec, 7, headers_summary, fill_header)
ws_exec.row_dimensions[7].height = 26

for r_idx, row_data in df_qtr_summary.iterrows():
    row_num = 8 + r_idx
    row_fill = fill_zebra if r_idx % 2 == 1 else PatternFill(fill_type=None)
    ws_exec.row_dimensions[row_num].height = 21
    
    ws_exec.cell(row=row_num, column=1, value=row_data['Quarter']).alignment = Alignment(horizontal='center')
    ws_exec.cell(row=row_num, column=2, value=int(row_data['Total Trades'])).alignment = Alignment(horizontal='center')
    ws_exec.cell(row=row_num, column=3, value=int(row_data['Wins'])).alignment = Alignment(horizontal='center')
    ws_exec.cell(row=row_num, column=4, value=int(row_data['Losses'])).alignment = Alignment(horizontal='center')
    
    cell_wr = ws_exec.cell(row=row_num, column=5, value=row_data['Win Rate (%)'])
    cell_wr.number_format = '0.00%'
    cell_wr.alignment = Alignment(horizontal='center')
    if row_data['Win Rate (%)'] >= 0.70:
        cell_wr.fill = fill_win; cell_wr.font = font_win
    elif row_data['Win Rate (%)'] >= 0.50:
        cell_wr.fill = fill_gold; cell_wr.font = font_gold
    else:
        cell_wr.fill = fill_loss; cell_wr.font = font_loss
    
    cell_beg = ws_exec.cell(row=row_num, column=6, value=row_data['Fund Utilised (₹)'])
    cell_beg.number_format = '₹#,##0.00'
    
    cell_pnl = ws_exec.cell(row=row_num, column=7, value=row_data['Net Quarter Equity P&L (₹)'])
    cell_pnl.number_format = '₹#,##0.00'
    cell_pnl.fill = fill_win if row_data['Net Quarter Equity P&L (₹)'] >= 0 else fill_loss
    cell_pnl.font = font_win if row_data['Net Quarter Equity P&L (₹)'] >= 0 else font_loss
    
    cell_end = ws_exec.cell(row=row_num, column=8, value=row_data['Final Value (₹)'])
    cell_end.number_format = '₹#,##0.00'
    cell_end.font = font_data_bold
    
    cell_ret = ws_exec.cell(row=row_num, column=9, value=row_data['Booked Return (%)'])
    cell_ret.number_format = '0.00%'
    
    cell_dd = ws_exec.cell(row=row_num, column=10, value=-abs(row_data['Max Drawdown (₹)']))
    cell_dd.number_format = '-₹#,##0.00'
    cell_dd.font = font_loss; cell_dd.fill = fill_loss
    
    cell_ddp = ws_exec.cell(row=row_num, column=11, value=row_data['Max Drawdown (%)'])
    cell_ddp.number_format = '-0.00%'
    cell_ddp.font = font_loss; cell_ddp.fill = fill_loss; cell_ddp.alignment = Alignment(horizontal='center')

    for c in range(1, 12):
        cell = ws_exec.cell(row=row_num, column=c)
        cell.border = border_thin
        if c not in [5, 7, 10, 11] and row_fill.fill_type: cell.fill = row_fill

# Grand Total Row
gt_row = 20
ws_exec.row_dimensions[gt_row].height = 24
ws_exec.cell(row=gt_row, column=1, value="12-Quarter Total / Summary").font = font_data_bold
ws_exec.cell(row=gt_row, column=2, value=len(df_all_trades)).alignment = Alignment(horizontal='center')
ws_exec.cell(row=gt_row, column=3, value=len(df_all_trades[df_all_trades['Return We Get (%)'] > 0])).alignment = Alignment(horizontal='center')
ws_exec.cell(row=gt_row, column=4, value=len(df_all_trades[df_all_trades['Return We Get (%)'] <= 0])).alignment = Alignment(horizontal='center')

c_gt_wr = ws_exec.cell(row=gt_row, column=5, value=global_win_rate)
c_gt_wr.number_format = '0.00%'
c_gt_wr.font = font_win; c_gt_wr.fill = fill_win; c_gt_wr.alignment = Alignment(horizontal='center')

c_gt_pnl = ws_exec.cell(row=gt_row, column=7, value=global_tot_pnl)
c_gt_pnl.number_format = '₹#,##0.00'
c_gt_pnl.font = font_win; c_gt_pnl.fill = fill_win

for c in range(1, 12):
    cell = ws_exec.cell(row=gt_row, column=c)
    cell.border = border_double_bottom
    if cell.font is None: cell.font = font_data_bold

autofit_columns(ws_exec, max_cols=11)

# Master 12-Quarter Cumulative PnL Line Chart
chart_master = LineChart()
chart_master.title = "Master 12-Quarter Cumulative Net Spot Equity P&L (₹)"
chart_master.style = 13
chart_master.y_axis.title = "Net Equity P&L (₹)"
chart_master.x_axis.title = "Quarter"
chart_master.width = 16
chart_master.height = 10

data_m = Reference(ws_exec, min_col=7, min_row=7, max_row=19)
cats_m = Reference(ws_exec, min_col=1, min_row=8, max_row=19)
chart_master.add_data(data_m, titles_from_data=True)
chart_master.set_categories(cats_m)
ws_exec.add_chart(chart_master, "M3")


# --- SHEET 2: Concurrent_Trades_Analysis ---
ws_conc = wb.create_sheet(title='Concurrent_Trades_Analysis')

# Title Banner
ws_conc.merge_cells('A1:H1')
c_ctitle = ws_conc['A1']
c_ctitle.value = "EQUITY CAPITAL EFFICIENCY & CONCURRENT OPEN SCRIPT ANALYSIS (EXACTLY 50 STOCKS PER QUARTER)"
c_ctitle.font = font_title
c_ctitle.fill = fill_title
c_ctitle.alignment = Alignment(horizontal='center', vertical='center')
ws_conc.row_dimensions[1].height = 36

headers_conc = [
    'Quarter', 'Total Trades', 'Slots Deployed', 'Max Concurrent Open Scripts',
    'Avg Concurrent Open Scripts', 'Total Net Equity P&L (₹)',
    'Max Realised Drawdown (₹)', 'Max Realised DD (%)'
]
style_table_headers(ws_conc, 3, headers_conc, fill_header)
ws_conc.row_dimensions[3].height = 26

for r_idx, r_data in df_qtr_summary.iterrows():
    row_num = 4 + r_idx
    row_fill = fill_zebra if r_idx % 2 == 1 else PatternFill(fill_type=None)
    ws_conc.row_dimensions[row_num].height = 21
    
    ws_conc.cell(row=row_num, column=1, value=r_data['Quarter']).alignment = Alignment(horizontal='center')
    ws_conc.cell(row=row_num, column=2, value=int(r_data['Total Trades'])).alignment = Alignment(horizontal='center')
    ws_conc.cell(row=row_num, column=3, value=int(r_data['Slots Deployed'])).alignment = Alignment(horizontal='center')
    
    c_mc = ws_conc.cell(row=row_num, column=4, value=int(r_data['Max Concurrent Scripts Open']))
    c_mc.alignment = Alignment(horizontal='center'); c_mc.font = font_data_bold; c_mc.fill = fill_gold
    
    ws_conc.cell(row=row_num, column=5, value=float(r_data['Avg Concurrent Scripts Open'])).alignment = Alignment(horizontal='center')
    
    c_pnl = ws_conc.cell(row=row_num, column=6, value=r_data['Net Quarter Equity P&L (₹)'])
    c_pnl.number_format = '₹#,##0.00'; c_pnl.font = font_win; c_pnl.fill = fill_win
    
    c_dd = ws_conc.cell(row=row_num, column=7, value=-abs(r_data['Max Drawdown (₹)']))
    c_dd.number_format = '-₹#,##0.00'; c_dd.font = font_loss; c_dd.fill = fill_loss
    
    c_ddp = ws_conc.cell(row=row_num, column=8, value=r_data['Max Drawdown (%)'])
    c_ddp.number_format = '-0.00%'; c_ddp.alignment = Alignment(horizontal='center')

    for c in range(1, 9):
        cell = ws_conc.cell(row=row_num, column=c)
        cell.border = border_thin
        if c not in [4, 6, 7] and row_fill.fill_type: cell.fill = row_fill

autofit_columns(ws_conc, max_cols=8)

# --- SHEET 3-14: 12 SEPARATE QUARTER SHEETS WITH DRAWDOWN CHARTS ---
qtr_card_labels = [
    'Capital Pool (₹)', 'Final Value (₹)', 'Net Booked Equity P&L (₹)',
    'Booked Return (%)', 'Expected Return (%)', 'Win Rate (%)',
    'Total Trades', 'Winning Trades', 'Losing Trades', 'Max Drawdown (₹)',
    'Max Drawdown (%)', 'Slots Deployed', 'Max Open Scripts'
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
    ws_q.merge_cells('A2:M2')
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
    c_mos = ws_q.cell(row=4, column=13, value=int(q_summary['Max Concurrent Scripts Open']))
    c_mos.alignment = Alignment(horizontal='center'); c_mos.font = font_data_bold; c_mos.fill = fill_gold
    ws_q.row_dimensions[4].height = 25
    
    for c in range(1, 14):
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


# --- SHEET 15: Stock_1Share_Leaderboard ---
ws_lead = wb.create_sheet(title='Stock_1Share_Leaderboard')
headers_lead = [
    'Rank', 'Symbol', 'Company Name', 'Sector', 'Shares per Trade',
    'Quarters Traded', 'Profitable Quarters', 'Wins', 'Losses',
    'Long Wins', 'Short Wins', 'Win Rate (%)', 'Avg Return per Trade (%)',
    'Total 12Q Equity P&L (₹)'
]
style_table_headers(ws_lead, 1, headers_lead, fill_header)
ws_lead.row_dimensions[1].height = 26

for r_idx, r_data in df_leaderboard.iterrows():
    row_num = 2 + r_idx
    row_fill = fill_zebra if r_idx % 2 == 1 else PatternFill(fill_type=None)
    ws_lead.row_dimensions[row_num].height = 21
    
    c_rank = ws_lead.cell(row=row_num, column=1, value=int(r_data['Rank']))
    c_rank.alignment = Alignment(horizontal='center')
    if r_data['Rank'] <= 3:
        c_rank.fill = fill_gold; c_rank.font = font_gold
        
    ws_lead.cell(row=row_num, column=2, value=r_data['Symbol']).alignment = Alignment(horizontal='center')
    ws_lead.cell(row=row_num, column=3, value=r_data['Company Name'])
    ws_lead.cell(row=row_num, column=4, value=r_data['Sector'])
    ws_lead.cell(row=row_num, column=5, value=int(r_data['Shares per Trade'])).alignment = Alignment(horizontal='center')
    ws_lead.cell(row=row_num, column=6, value=int(r_data['Quarters Traded'])).alignment = Alignment(horizontal='center')
    ws_lead.cell(row=row_num, column=7, value=int(r_data['Profitable Quarters'])).alignment = Alignment(horizontal='center')
    ws_lead.cell(row=row_num, column=8, value=int(r_data['Wins'])).alignment = Alignment(horizontal='center')
    ws_lead.cell(row=row_num, column=9, value=int(r_data['Losses'])).alignment = Alignment(horizontal='center')
    ws_lead.cell(row=row_num, column=10, value=int(r_data['Long Wins'])).alignment = Alignment(horizontal='center')
    ws_lead.cell(row=row_num, column=11, value=int(r_data['Short Wins'])).alignment = Alignment(horizontal='center')
    
    cell_wr = ws_lead.cell(row=row_num, column=12, value=r_data['Win Rate (%)'])
    cell_wr.number_format = '0.00%'
    cell_wr.alignment = Alignment(horizontal='center')
    
    cell_ar = ws_lead.cell(row=row_num, column=13, value=r_data['Avg Return per Trade (%)'])
    cell_ar.number_format = '0.00%'
    
    cell_tot = ws_lead.cell(row=row_num, column=14, value=r_data['Total 12Q Equity P&L (₹)'])
    cell_tot.number_format = '₹#,##0.00'
    cell_tot.font = font_win if r_data['Total 12Q Equity P&L (₹)'] >= 0 else font_loss
    cell_tot.fill = fill_win if r_data['Total 12Q Equity P&L (₹)'] >= 0 else fill_loss

    for c in range(1, 15):
        cell = ws_lead.cell(row=row_num, column=c)
        cell.border = border_thin
        if c not in [1, 14] and row_fill.fill_type: cell.fill = row_fill

autofit_columns(ws_lead, max_cols=14)

# --- SHEET 16: All_12Q_Equity_Trades ---
ws_all = wb.create_sheet(title='All_12Q_Equity_Trades')
all_headers = ['Quarter', 'FY'] + trade_table_headers
style_table_headers(ws_all, 1, all_headers, fill_header)
ws_all.row_dimensions[1].height = 26

for r_idx, t_data in df_all_trades.iterrows():
    row_num = 2 + r_idx
    row_fill = fill_zebra if r_idx % 2 == 1 else PatternFill(fill_type=None)
    ws_all.row_dimensions[row_num].height = 21
    
    ws_all.cell(row=row_num, column=1, value=t_data['Quarter']).alignment = Alignment(horizontal='center')
    ws_all.cell(row=row_num, column=2, value=t_data['FY']).alignment = Alignment(horizontal='center')
    ws_all.cell(row=row_num, column=3, value=int(t_data['Trade No'])).alignment = Alignment(horizontal='center')
    ws_all.cell(row=row_num, column=4, value=t_data['Symbol']).alignment = Alignment(horizontal='center')
    ws_all.cell(row=row_num, column=5, value=t_data['Company Name'])
    ws_all.cell(row=row_num, column=6, value=t_data['Sector'])
    
    c_strat = ws_all.cell(row=row_num, column=7, value=t_data['Strategy'])
    c_strat.alignment = Alignment(horizontal='center')
    if 'LONG' in t_data['Strategy']:
        c_strat.fill = fill_win; c_strat.font = font_win
    else:
        c_strat.fill = fill_loss; c_strat.font = font_loss
        
    ws_all.cell(row=row_num, column=8, value=t_data['Position Taking Window']).alignment = Alignment(horizontal='center')
    ws_all.cell(row=row_num, column=9, value=t_data['Entry Date']).alignment = Alignment(horizontal='center')
    ws_all.cell(row=row_num, column=10, value=t_data['Exit Date']).alignment = Alignment(horizontal='center')
    
    c_enp = ws_all.cell(row=row_num, column=11, value=t_data['Entry Price (₹)'])
    c_enp.number_format = '₹#,##0.00'
    
    c_exp = ws_all.cell(row=row_num, column=12, value=t_data['Exit Price (₹)'])
    c_exp.number_format = '₹#,##0.00'
    
    ws_all.cell(row=row_num, column=13, value=int(t_data['Quantity (Shares)'])).alignment = Alignment(horizontal='center')
    
    c_er = ws_all.cell(row=row_num, column=14, value=t_data['Expected Return (%)'])
    c_er.number_format = '0.00%'
    
    c_rw = ws_all.cell(row=row_num, column=15, value=t_data['Return We Get (%)'])
    c_rw.number_format = '0.00%'
    
    c_fpnl = ws_all.cell(row=row_num, column=16, value=t_data['Booked Equity P&L (₹)'])
    c_fpnl.number_format = '₹#,##0.00'
    c_fpnl.font = font_win if t_data['Booked Equity P&L (₹)'] >= 0 else font_loss
    c_fpnl.fill = fill_win if t_data['Booked Equity P&L (₹)'] >= 0 else fill_loss
    
    ws_all.cell(row=row_num, column=17, value=t_data['Assigned Slot']).alignment = Alignment(horizontal='center')
    ws_all.cell(row=row_num, column=18, value=t_data['Re-entry Type']).alignment = Alignment(horizontal='center')
    ws_all.cell(row=row_num, column=19, value=t_data['Quarterly Result Date']).alignment = Alignment(horizontal='center')

    for c in range(1, 20):
        cell = ws_all.cell(row=row_num, column=c)
        cell.border = border_thin
        if c not in [7, 16] and row_fill.fill_type: cell.fill = row_fill

autofit_columns(ws_all, max_cols=19)

# --- SHEET 17: Losing_Trades_Audit ---
ws_loss = wb.create_sheet(title='Losing_Trades_Audit')
loss_headers = [
    'Quarter', 'Symbol', 'Company Name', 'Sector', 'Strategy',
    'Entry Date', 'Exit Date', 'Entry Price (₹)', 'Exit Price (₹)',
    'Quantity (Shares)', 'Booked Equity Loss (₹)', 'Return We Get (%)'
]
style_table_headers(ws_loss, 1, loss_headers, fill_header)
ws_loss.row_dimensions[1].height = 26

for r_idx, t_data in df_losing.iterrows():
    row_num = 2 + r_idx
    row_fill = fill_zebra if r_idx % 2 == 1 else PatternFill(fill_type=None)
    ws_loss.row_dimensions[row_num].height = 21
    
    ws_loss.cell(row=row_num, column=1, value=t_data['Quarter']).alignment = Alignment(horizontal='center')
    ws_loss.cell(row=row_num, column=2, value=t_data['Symbol']).alignment = Alignment(horizontal='center')
    ws_loss.cell(row=row_num, column=3, value=t_data['Company Name'])
    ws_loss.cell(row=row_num, column=4, value=t_data['Sector'])
    
    c_strat = ws_loss.cell(row=row_num, column=5, value=t_data['Strategy'])
    c_strat.alignment = Alignment(horizontal='center')
    c_strat.fill = fill_loss; c_strat.font = font_loss
    
    ws_loss.cell(row=row_num, column=6, value=t_data['Entry Date']).alignment = Alignment(horizontal='center')
    ws_loss.cell(row=row_num, column=7, value=t_data['Exit Date']).alignment = Alignment(horizontal='center')
    
    c_enp = ws_loss.cell(row=row_num, column=8, value=t_data['Entry Price (₹)'])
    c_enp.number_format = '₹#,##0.00'
    
    c_exp = ws_loss.cell(row=row_num, column=9, value=t_data['Exit Price (₹)'])
    c_exp.number_format = '₹#,##0.00'
    
    ws_loss.cell(row=row_num, column=10, value=int(t_data['Quantity (Shares)'])).alignment = Alignment(horizontal='center')
    
    c_pnl = ws_loss.cell(row=row_num, column=11, value=t_data['Booked Equity P&L (₹)'])
    c_pnl.number_format = '₹#,##0.00'
    c_pnl.font = font_loss; c_pnl.fill = fill_loss
    
    c_rw = ws_loss.cell(row=row_num, column=12, value=t_data['Return We Get (%)'])
    c_rw.number_format = '0.00%'

    for c in range(1, 13):
        cell = ws_loss.cell(row=row_num, column=c)
        cell.border = border_thin
        if c not in [5, 11] and row_fill.fill_type: cell.fill = row_fill

autofit_columns(ws_loss, max_cols=12)

# Save Workbook
print(f"Saving Master Equity Workbook to {OUTPUT_EQUITY_MASTER.name}...", flush=True)
wb.save(OUTPUT_EQUITY_MASTER)

print(f"\n🎉 PERFECT EQUITY MASTER WORKBOOK V2 CREATED IN {time.time() - t0:.1f}s!", flush=True)
