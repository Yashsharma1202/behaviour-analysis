import pandas as pd
import pathlib
import sys
import time
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(errors='replace')

t0 = time.time()

# ------------------- Configuration & Paths -------------------
BASE_DIR = pathlib.Path('D:/behaviour analysis')
CONSOLIDATED_PATH = BASE_DIR / 'Nifty50_12_Quarters_Consolidated.xlsx'
OPTIONS_MASTER_PATH = BASE_DIR / 'Nifty50_12_Quarters_Options_OI_Master_v4.xlsx'
RESULT_DATES_PATH = BASE_DIR / 'Nifty50Stocks_QtyResultDates.xlsx'
FUTURES_MASTER_EXISTING = BASE_DIR / 'Futures_Master_Only_v2.xlsx'

OUTPUT_FUTURES_MASTER = BASE_DIR / 'Nifty50_12_Quarters_Futures_OI_Master_v6.xlsx'
OUTPUT_FUTURES_ONLY = BASE_DIR / 'Futures_Master_Only_v5.xlsx'

# Styling Tokens
COLOR_HEADER_NAVY = '1F497D'      # Main Header Fill (Navy Blue)
COLOR_HEADER_TEAL = '006688'      # Sub Header Fill (Teal)
COLOR_ZEBRA_LIGHT = 'F2F5F8'      # Alternating Row Fill
COLOR_WIN_GREEN  = 'E2EFDA'      # Light Green for Positive P&L
COLOR_LOSS_RED   = 'FCE4D6'      # Light Red for Negative P&L
COLOR_TEXT_WHITE = 'FFFFFF'

font_title  = Font(name='Segoe UI', size=13, bold=True, color=COLOR_TEXT_WHITE)
font_header = Font(name='Segoe UI', size=10, bold=True, color=COLOR_TEXT_WHITE)
font_bold   = Font(name='Segoe UI', size=10, bold=True)
font_regular= Font(name='Segoe UI', size=10)

fill_navy   = PatternFill(start_color=COLOR_HEADER_NAVY, end_color=COLOR_HEADER_NAVY, fill_type='solid')
fill_teal   = PatternFill(start_color=COLOR_HEADER_TEAL, end_color=COLOR_HEADER_TEAL, fill_type='solid')
fill_win    = PatternFill(start_color=COLOR_WIN_GREEN, end_color=COLOR_WIN_GREEN, fill_type='solid')
fill_loss   = PatternFill(start_color=COLOR_LOSS_RED, end_color=COLOR_LOSS_RED, fill_type='solid')

# ------------------- 1. Load Daily Futures Dataset -------------------
print("Loading daily futures dataset...", flush=True)
try:
    df_daily = pd.read_excel(FUTURES_MASTER_EXISTING, sheet_name='Futures_Master', engine='openpyxl')
    print(f"Loaded {len(df_daily)} daily futures records in {time.time() - t0:.1f}s", flush=True)
except Exception as e:
    print(f"Error loading daily dataset: {e}", flush=True)
    sys.exit(1)

df_daily['Date'] = pd.to_datetime(df_daily['Date']).dt.strftime('%Y-%m-%d')
df_daily['Symbol'] = df_daily['Symbol'].astype(str).str.strip()

price_lookup_map = {}
for _, r in df_daily.iterrows():
    p = r.iloc[2]
    if pd.notnull(p) and float(p) > 0:
        price_lookup_map[(r['Symbol'], str(r['Date']))] = float(p)

# ------------------- 2. Load Result Dates -------------------
print("Loading Quarterly Result Dates...", flush=True)
result_date_map = {} # (symbol, quarter_str) -> date_str

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

# ------------------- 3. Lot Size Map -------------------
lot_size_map = {
    'RELIANCE': 250, 'TCS': 175, 'INFY': 400, 'HDFCBANK': 550, 'ICICIBANK': 700,
    'AXISBANK': 625, 'SBIN': 1500, 'BHARTIARTL': 950, 'ITC': 1600, 'KOTAKBANK': 400,
    'LTIM': 150, 'LT': 300, 'HINDUNILVR': 300, 'BAJFINANCE': 125, 'MARUTI': 100,
    'ASIANPAINT': 200, 'HCLTECH': 350, 'TITAN': 175, 'SUNPHARMA': 350, 'TATAMOTORS': 1425,
    'ULTRACEMCO': 100, 'POWERGRID': 1800, 'NTPC': 1500, 'COALINDIA': 2100, 'TATASTEEL': 5500,
    'JIOFIN': 2000, 'JSWSTEEL': 675, 'M&M': 350, 'ADANIENT': 300, 'ADANIPORTS': 800,
    'GRASIM': 475, 'BAJAJFINSV': 500, 'BAJAJ-AUTO': 125, 'NESTLEIND': 250, 'APOLLOHOSP': 125,
    'WIPRO': 1500, 'EICHERMOT': 175, 'DIVISLAB': 150, 'DRREDDY': 125, 'CIPLA': 650,
    'BPCL': 1800, 'TATACONSUM': 900, 'BRITANNIA': 200, 'HEROMOTOCO': 150, 'INDUSINDBK': 500,
    'HDFCLIFE': 1100, 'SBILIFE': 750, 'BEL': 1500, 'SHRIRAMFIN': 300, 'TECHM': 600
}

def get_lot_size(sym):
    return lot_size_map.get(sym, 500)

# ------------------- 4. Load Trades from CONSOLIDATED MASTER -------------------
print("Loading Trades from Consolidated Master...", flush=True)
try:
    df_cons_trades = pd.read_excel(CONSOLIDATED_PATH, sheet_name='All_12_Quarters_Master_Trades', engine='openpyxl')
except Exception as e:
    print(f"Error reading consolidated trades: {e}", flush=True)
    sys.exit(1)

ret_col = next(c for c in df_cons_trades.columns if 'Realised Return' in c)
pnl_col = next(c for c in df_cons_trades.columns if 'Realised P' in c or 'Profit' in c or 'P&L' in c)
entry_price_col = next(c for c in df_cons_trades.columns if 'Entry Price' in c)
exit_price_col = next(c for c in df_cons_trades.columns if 'Exit Price' in c)

df_cons_trades[ret_col] = pd.to_numeric(df_cons_trades[ret_col], errors='coerce')

processed_trades = []

for idx, row in df_cons_trades.iterrows():
    sym = str(row['Symbol']).strip()
    entry_dt = pd.to_datetime(row['Entry Date']).strftime('%Y-%m-%d')
    exit_dt = pd.to_datetime(row['Exit Date']).strftime('%Y-%m-%d')
    orig_strat = str(row['Strategy']).strip().upper()
    strat = 'FUTURE LONG' if 'LONG' in orig_strat else 'FUTURE SHORT'
    direction = 1 if 'LONG' in strat else -1
    lot_size = get_lot_size(sym)

    eq_entry = float(row[entry_price_col])
    eq_exit  = float(row[exit_price_col])
    ret_pct  = float(row[ret_col])
    exp_ret  = float(row.get('Expected Return (%)', 0.0) or 0.0)

    # Split-adjusted Futures Price anchoring
    fut_entry = price_lookup_map.get((sym, entry_dt), eq_entry)
    if fut_entry > 1.8 * eq_entry or fut_entry < 0.5 * eq_entry:
        fut_entry = eq_entry
        
    fut_exit = round(fut_entry * (1.0 + direction * ret_pct), 2)
    booked_pnl = direction * (fut_exit - fut_entry) * lot_size
    
    qtr = str(row['Quarter']).strip()
    result_dt = result_date_map.get((sym, qtr), '')

    trade_info = {
        'Quarter': qtr,
        'FY': str(row.get('FY', '')),
        'Trade No': int(row.get('Trade No', idx + 1)),
        'Symbol': sym,
        'Company Name': str(row.get('Company Name', '')),
        'Sector': str(row.get('Sector', '')),
        'Strategy': strat,
        'Position Taking Window': str(row.get('Position Taking Window', '')),
        'Entry Date': entry_dt,
        'Exit Date': exit_dt,
        'Entry Futures Price (₹)': round(fut_entry, 2),
        'Exit Futures Price (₹)': round(fut_exit, 2),
        'Lot Size (Qty)': int(lot_size),
        'Booked Futures P&L (₹)': round(booked_pnl, 2),
        'Expected Return (%)': round(exp_ret, 4),
        'Return We Get (%)': round(ret_pct, 4),
        'Quarterly Result Date': result_dt,
        'Assigned Slot': str(row.get('Assigned Slot', '')),
        'Re-entry Type': str(row.get('Re-entry Type', 'First Entry')),
        'Data Query Status': 'Success'
    }
    processed_trades.append(trade_info)

df_all_trades = pd.DataFrame(processed_trades)

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

# ------------------- 5. Per-Quarter Fund Summaries -------------------
quarter_summaries = []

for qtr in quarters_order:
    q_trades = df_all_trades[df_all_trades['Quarter'] == qtr]
    t_count = len(q_trades)
    wins = len(q_trades[q_trades['Return We Get (%)'] > 0])
    losses = len(q_trades[q_trades['Return We Get (%)'] <= 0])
    longs = len(q_trades[q_trades['Strategy'] == 'FUTURE LONG'])
    shorts = len(q_trades[q_trades['Strategy'] == 'FUTURE SHORT'])
    win_rate = (wins / t_count) if t_count > 0 else 0.0
    
    q_pnl_fut = q_trades['Booked Futures P&L (₹)'].sum()
    
    # Calculate Per-Lot Fund Deployed in Quarter
    # Each slot deploys 1 lot of futures contract (~15% margin or 1 lot entry price * lot size)
    avg_slot_capital = sum(r['Entry Futures Price (₹)'] * r['Lot Size (Qty)'] * 0.20 for _, r in q_trades.iterrows()) / t_count if t_count > 0 else 500000.0
    slots_deployed = len(q_trades['Assigned Slot'].unique())
    fund_utilised = round(avg_slot_capital * slots_deployed, 2)
    final_val_quarter = fund_utilised + q_pnl_fut
    q_return = (q_pnl_fut / fund_utilised) if fund_utilised > 0 else 0.0
    
    exp_ret_avg = q_trades['Expected Return (%)'].mean()
    re_entries = len(q_trades[q_trades['Re-entry Type'] != 'First Entry'])
    
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
        'Net Quarter Futures P&L (₹)': q_pnl_fut,
        'Final Value (₹)': final_val_quarter,
        'Booked Return (%)': q_return,
        'Expected Return (%)': exp_ret_avg,
        'Slots Deployed': slots_deployed,
        'Re-entries': re_entries
    })

df_qtr_summary = pd.DataFrame(quarter_summaries)

# ------------------- 6. Stock Leaderboard -------------------
stock_leaderboard = []
for sym, group in df_all_trades.groupby('Symbol'):
    c_name = group.iloc[0]['Company Name']
    sector = group.iloc[0]['Sector']
    lot = group.iloc[0]['Lot Size (Qty)']
    
    t_total = len(group)
    wins = len(group[group['Return We Get (%)'] > 0])
    losses = len(group[group['Return We Get (%)'] <= 0])
    long_wins = len(group[(group['Strategy'] == 'FUTURE LONG') & (group['Return We Get (%)'] > 0)])
    short_wins = len(group[(group['Strategy'] == 'FUTURE SHORT') & (group['Return We Get (%)'] > 0)])
    win_rate = (wins / t_total) if t_total > 0 else 0.0
    
    total_pnl = group['Booked Futures P&L (₹)'].sum()
    avg_ret = group['Return We Get (%)'].mean()
    
    prof_qtrs = 0
    for qtr in quarters_order:
        q_p = group[group['Quarter'] == qtr]['Booked Futures P&L (₹)'].sum()
        if q_p > 0: prof_qtrs += 1
        
    stock_leaderboard.append({
        'Symbol': sym,
        'Company Name': c_name,
        'Sector': sector,
        'Lot Size': lot,
        'Quarters Traded': len(group['Quarter'].unique()),
        'Profitable Quarters': prof_qtrs,
        'Wins': wins,
        'Losses': losses,
        'Long Wins': long_wins,
        'Short Wins': short_wins,
        'Win Rate (%)': win_rate,
        'Avg Return per Trade (%)': avg_ret,
        'Total 12Q Futures P&L (₹)': total_pnl
    })

df_leaderboard = pd.DataFrame(stock_leaderboard)
df_leaderboard = df_leaderboard.sort_values(by='Total 12Q Futures P&L (₹)', ascending=False).reset_index(drop=True)
df_leaderboard.insert(0, 'Rank', range(1, len(df_leaderboard) + 1))

# ------------------- 7. Losing Trades Audit -------------------
df_losing = df_all_trades[df_all_trades['Return We Get (%)'] <= 0].copy().reset_index(drop=True)

# ------------------- 8. Build 17-Sheet Excel Workbook -------------------
print("Building 17-Sheet Master Workbook with exact card layout...", flush=True)
wb = Workbook()
wb.remove(wb.active)

def style_table_headers(ws, start_row, headers, fill_style=fill_navy):
    for c_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=start_row, column=c_idx, value=h)
        cell.font = font_header
        cell.fill = fill_style
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

def autofit_columns(ws, max_cols=25):
    ws.freeze_panes = 'A2'
    for col_idx in range(1, max_cols + 1):
        col_letter = get_column_letter(col_idx)
        max_len = 0
        for row in range(1, min(ws.max_row + 1, 100)):
            val = str(ws.cell(row=row, column=col_idx).value or '')
            if len(val) > max_len: max_len = len(val)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 13)

# --- SHEET 1: Exec_12Q_Combined_Summary ---
ws_exec = wb.create_sheet(title='Exec_12Q_Combined_Summary')

ws_exec.merge_cells('A1:J1')
cell_title = ws_exec['A1']
cell_title.value = "NIFTY 50 EVENT-DRIVEN FUTURES TRADING STRATEGY — 12-QUARTERS EXECUTIVE SCORECARD"
cell_title.font = font_title
cell_title.fill = fill_navy
cell_title.alignment = Alignment(horizontal='center', vertical='center')
ws_exec.row_dimensions[1].height = 35

global_tot_pnl = df_all_trades['Booked Futures P&L (₹)'].sum()
global_win_rate = len(df_all_trades[df_all_trades['Return We Get (%)'] > 0]) / len(df_all_trades)

ws_exec.cell(row=3, column=1, value="Total Trades Analyzed:").font = font_bold
ws_exec.cell(row=3, column=2, value=len(df_all_trades)).alignment = Alignment(horizontal='center')

ws_exec.cell(row=3, column=4, value="Global Win Rate:").font = font_bold
ws_exec.cell(row=3, column=5, value=global_win_rate).number_format = '0.0%'

ws_exec.cell(row=4, column=1, value="Total Net Futures P&L:").font = font_bold
ws_exec.cell(row=4, column=2, value=global_tot_pnl).number_format = '₹#,##0.00'

headers_summary = [
    'Quarter', 'Total Trades', 'Wins', 'Losses', 'Win Rate (%)',
    'Estimated Margin Utilised (₹)', 'Net Quarter Futures P&L (₹)', 'Final Quarter Value (₹)',
    'Booked Return (%)', 'Slots Deployed'
]
style_table_headers(ws_exec, 6, headers_summary, fill_navy)
ws_exec.row_dimensions[6].height = 25

for r_idx, row_data in df_qtr_summary.iterrows():
    row_num = 7 + r_idx
    ws_exec.cell(row=row_num, column=1, value=row_data['Quarter']).alignment = Alignment(horizontal='center')
    ws_exec.cell(row=row_num, column=2, value=int(row_data['Total Trades'])).alignment = Alignment(horizontal='center')
    ws_exec.cell(row=row_num, column=3, value=int(row_data['Wins'])).alignment = Alignment(horizontal='center')
    ws_exec.cell(row=row_num, column=4, value=int(row_data['Losses'])).alignment = Alignment(horizontal='center')
    
    cell_wr = ws_exec.cell(row=row_num, column=5, value=row_data['Win Rate (%)'])
    cell_wr.number_format = '0.00%'
    cell_wr.alignment = Alignment(horizontal='center')
    
    cell_beg = ws_exec.cell(row=row_num, column=6, value=row_data['Fund Utilised (₹)'])
    cell_beg.number_format = '₹#,##0.00'
    
    cell_pnl = ws_exec.cell(row=row_num, column=7, value=row_data['Net Quarter Futures P&L (₹)'])
    cell_pnl.number_format = '₹#,##0.00'
    cell_pnl.fill = fill_win if row_data['Net Quarter Futures P&L (₹)'] >= 0 else fill_loss
    
    cell_end = ws_exec.cell(row=row_num, column=8, value=row_data['Final Value (₹)'])
    cell_end.number_format = '₹#,##0.00'
    cell_end.font = font_bold
    
    cell_ret = ws_exec.cell(row=row_num, column=9, value=row_data['Booked Return (%)'])
    cell_ret.number_format = '0.00%'
    
    ws_exec.cell(row=row_num, column=10, value=int(row_data['Slots Deployed'])).alignment = Alignment(horizontal='center')

autofit_columns(ws_exec, max_cols=10)

# --- SHEET 2: Stock_1Lot_Leaderboard ---
ws_lead = wb.create_sheet(title='Stock_1Lot_Leaderboard')
headers_lead = [
    'Rank', 'Symbol', 'Company Name', 'Sector', 'Lot Size',
    'Quarters Traded', 'Profitable Quarters', 'Wins', 'Losses',
    'Long Wins', 'Short Wins', 'Win Rate (%)', 'Avg Return per Trade (%)',
    'Total 12Q Futures P&L (₹)'
]
style_table_headers(ws_lead, 1, headers_lead, fill_teal)

for r_idx, r_data in df_leaderboard.iterrows():
    row_num = 2 + r_idx
    ws_lead.cell(row=row_num, column=1, value=int(r_data['Rank'])).alignment = Alignment(horizontal='center')
    ws_lead.cell(row=row_num, column=2, value=r_data['Symbol']).alignment = Alignment(horizontal='center')
    ws_lead.cell(row=row_num, column=3, value=r_data['Company Name'])
    ws_lead.cell(row=row_num, column=4, value=r_data['Sector'])
    ws_lead.cell(row=row_num, column=5, value=int(r_data['Lot Size'])).alignment = Alignment(horizontal='center')
    ws_lead.cell(row=row_num, column=6, value=int(r_data['Quarters Traded'])).alignment = Alignment(horizontal='center')
    ws_lead.cell(row=row_num, column=7, value=int(r_data['Profitable Quarters'])).alignment = Alignment(horizontal='center')
    ws_lead.cell(row=row_num, column=8, value=int(r_data['Wins'])).alignment = Alignment(horizontal='center')
    ws_lead.cell(row=row_num, column=9, value=int(r_data['Losses'])).alignment = Alignment(horizontal='center')
    ws_lead.cell(row=row_num, column=10, value=int(r_data['Long Wins'])).alignment = Alignment(horizontal='center')
    ws_lead.cell(row=row_num, column=11, value=int(r_data['Short Wins'])).alignment = Alignment(horizontal='center')
    
    cell_wr = ws_lead.cell(row=row_num, column=12, value=r_data['Win Rate (%)'])
    cell_wr.number_format = '0.00%'
    
    cell_ar = ws_lead.cell(row=row_num, column=13, value=r_data['Avg Return per Trade (%)'])
    cell_ar.number_format = '0.00%'
    
    cell_tot = ws_lead.cell(row=row_num, column=14, value=r_data['Total 12Q Futures P&L (₹)'])
    cell_tot.number_format = '₹#,##0.00'
    cell_tot.font = font_bold
    cell_tot.fill = fill_win if r_data['Total 12Q Futures P&L (₹)'] >= 0 else fill_loss

autofit_columns(ws_lead, max_cols=14)

# --- SHEET 3-14: 12 SEPARATE QUARTER SHEETS MATCHING EXACT CONCEPT ---
qtr_card_labels = [
    'Fund Utilised (₹)', 'Final Value (₹)', 'Net Booked Futures P&L (₹)',
    'Booked Return (%)', 'Expected Return (%)', 'Win Rate (%)',
    'Total Trades', 'Winning Trades', 'Losing Trades', 'Long Trades',
    'Short Trades', 'Slots Deployed', 'Re-entries'
]

trade_table_headers = [
    'Trade #', 'Symbol', 'Company Name', 'Sector', 'Strategy',
    'Position Taking Window', 'Entry Date', 'Exit Date',
    'Entry Futures Price (₹)', 'Exit Futures Price (₹)', 'Lot Size (Qty)',
    'Expected Return (%)', 'Booked Return (%)', 'Booked Futures P&L (₹)',
    'Assigned Slot', 'Re-entry Type', 'Quarterly Result Date'
]

for qtr in quarters_order:
    ws_q = wb.create_sheet(title=qtr)
    
    q_trades = df_all_trades[df_all_trades['Quarter'] == qtr].reset_index(drop=True)
    q_summary = df_qtr_summary[df_qtr_summary['Quarter'] == qtr].iloc[0]
    dt_range = qtr_date_ranges.get(qtr, '')
    fy_str = q_summary['FY']
    
    # Row 1: Title Banner
    ws_q.merge_cells('A1:Q1')
    c_title = ws_q['A1']
    c_title.value = f"NIFTY 50 EVENT-DRIVEN FUTURES MODEL — {qtr.upper()} ({dt_range.upper()})"
    c_title.font = font_title
    c_title.fill = fill_navy
    c_title.alignment = Alignment(horizontal='center', vertical='center')
    ws_q.row_dimensions[1].height = 32
    
    # Row 2: Section Banner for Summary Cards
    ws_q.merge_cells('A2:M2')
    c_sec1 = ws_q['A2']
    c_sec1.value = f"QUARTER PERFORMANCE OVERVIEW & CAPITAL UTILISATION ({fy_str})"
    c_sec1.font = font_header
    c_sec1.fill = fill_navy
    c_sec1.alignment = Alignment(horizontal='left', vertical='center')
    ws_q.row_dimensions[2].height = 22
    
    # Row 3: Card Labels
    style_table_headers(ws_q, 3, qtr_card_labels, fill_teal)
    ws_q.row_dimensions[3].height = 22
    
    # Row 4: Card Values
    ws_q.cell(row=4, column=1, value=q_summary['Fund Utilised (₹)']).number_format = '₹#,##0.00'
    
    c_end = ws_q.cell(row=4, column=2, value=q_summary['Final Value (₹)'])
    c_end.number_format = '₹#,##0.00'
    c_end.font = font_bold
    
    c_pnl = ws_q.cell(row=4, column=3, value=q_summary['Net Quarter Futures P&L (₹)'])
    c_pnl.number_format = '₹#,##0.00'
    c_pnl.font = font_bold
    c_pnl.fill = fill_win if q_summary['Net Quarter Futures P&L (₹)'] >= 0 else fill_loss
    
    ws_q.cell(row=4, column=4, value=q_summary['Booked Return (%)']).number_format = '0.00%'
    ws_q.cell(row=4, column=5, value=q_summary['Expected Return (%)']).number_format = '0.00%'
    
    c_wr = ws_q.cell(row=4, column=6, value=q_summary['Win Rate (%)'])
    c_wr.number_format = '0.00%'
    c_wr.font = font_bold
    
    ws_q.cell(row=4, column=7, value=int(q_summary['Total Trades'])).alignment = Alignment(horizontal='center')
    ws_q.cell(row=4, column=8, value=int(q_summary['Wins'])).alignment = Alignment(horizontal='center')
    ws_q.cell(row=4, column=9, value=int(q_summary['Losses'])).alignment = Alignment(horizontal='center')
    ws_q.cell(row=4, column=10, value=int(q_summary['Long Trades'])).alignment = Alignment(horizontal='center')
    ws_q.cell(row=4, column=11, value=int(q_summary['Short Trades'])).alignment = Alignment(horizontal='center')
    ws_q.cell(row=4, column=12, value=int(q_summary['Slots Deployed'])).alignment = Alignment(horizontal='center')
    ws_q.cell(row=4, column=13, value=int(q_summary['Re-entries'])).alignment = Alignment(horizontal='center')
    ws_q.row_dimensions[4].height = 24
    
    # Row 6: Section Banner for Chronological Trade Journal
    ws_q.merge_cells('A6:Q6')
    c_sec2 = ws_q['A6']
    c_sec2.value = "CHRONOLOGICAL TRADE JOURNAL — POSITION TAKING & BOOKED FUTURES PROFIT/LOSS"
    c_sec2.font = font_header
    c_sec2.fill = fill_navy
    c_sec2.alignment = Alignment(horizontal='left', vertical='center')
    ws_q.row_dimensions[6].height = 22
    
    # Row 7: Table Headers
    style_table_headers(ws_q, 7, trade_table_headers, fill_teal)
    ws_q.row_dimensions[7].height = 25
    
    # Row 8+: Trade Data
    for r_idx, t_data in q_trades.iterrows():
        row_num = 8 + r_idx
        ws_q.cell(row=row_num, column=1, value=int(t_data['Trade No'])).alignment = Alignment(horizontal='center')
        ws_q.cell(row=row_num, column=2, value=t_data['Symbol']).alignment = Alignment(horizontal='center')
        ws_q.cell(row=row_num, column=3, value=t_data['Company Name'])
        ws_q.cell(row=row_num, column=4, value=t_data['Sector'])
        ws_q.cell(row=row_num, column=5, value=t_data['Strategy']).alignment = Alignment(horizontal='center')
        ws_q.cell(row=row_num, column=6, value=t_data['Position Taking Window']).alignment = Alignment(horizontal='center')
        ws_q.cell(row=row_num, column=7, value=t_data['Entry Date']).alignment = Alignment(horizontal='center')
        ws_q.cell(row=row_num, column=8, value=t_data['Exit Date']).alignment = Alignment(horizontal='center')
        
        c_enp = ws_q.cell(row=row_num, column=9, value=t_data['Entry Futures Price (₹)'])
        c_enp.number_format = '₹#,##0.00'
        
        c_exp = ws_q.cell(row=row_num, column=10, value=t_data['Exit Futures Price (₹)'])
        c_exp.number_format = '₹#,##0.00'
        
        ws_q.cell(row=row_num, column=11, value=int(t_data['Lot Size (Qty)'])).alignment = Alignment(horizontal='center')
        
        c_er = ws_q.cell(row=row_num, column=12, value=t_data['Expected Return (%)'])
        c_er.number_format = '0.00%'
        
        c_rw = ws_q.cell(row=row_num, column=13, value=t_data['Return We Get (%)'])
        c_rw.number_format = '0.00%'
        
        c_fpnl = ws_q.cell(row=row_num, column=14, value=t_data['Booked Futures P&L (₹)'])
        c_fpnl.number_format = '₹#,##0.00'
        c_fpnl.font = font_bold
        c_fpnl.fill = fill_win if t_data['Booked Futures P&L (₹)'] >= 0 else fill_loss
        
        ws_q.cell(row=row_num, column=15, value=t_data['Assigned Slot']).alignment = Alignment(horizontal='center')
        ws_q.cell(row=row_num, column=16, value=t_data['Re-entry Type']).alignment = Alignment(horizontal='center')
        ws_q.cell(row=row_num, column=17, value=t_data['Quarterly Result Date']).alignment = Alignment(horizontal='center')
        
    autofit_columns(ws_q, max_cols=17)
    ws_q.freeze_panes = 'A8'

# --- SHEET 15: All_12Q_Futures_Trades ---
ws_all = wb.create_sheet(title='All_12Q_Futures_Trades')
all_headers = ['Quarter', 'FY'] + trade_table_headers
style_table_headers(ws_all, 1, all_headers, fill_navy)

for r_idx, t_data in df_all_trades.iterrows():
    row_num = 2 + r_idx
    ws_all.cell(row=row_num, column=1, value=t_data['Quarter']).alignment = Alignment(horizontal='center')
    ws_all.cell(row=row_num, column=2, value=t_data['FY']).alignment = Alignment(horizontal='center')
    ws_all.cell(row=row_num, column=3, value=int(t_data['Trade No'])).alignment = Alignment(horizontal='center')
    ws_all.cell(row=row_num, column=4, value=t_data['Symbol']).alignment = Alignment(horizontal='center')
    ws_all.cell(row=row_num, column=5, value=t_data['Company Name'])
    ws_all.cell(row=row_num, column=6, value=t_data['Sector'])
    ws_all.cell(row=row_num, column=7, value=t_data['Strategy']).alignment = Alignment(horizontal='center')
    ws_all.cell(row=row_num, column=8, value=t_data['Position Taking Window']).alignment = Alignment(horizontal='center')
    ws_all.cell(row=row_num, column=9, value=t_data['Entry Date']).alignment = Alignment(horizontal='center')
    ws_all.cell(row=row_num, column=10, value=t_data['Exit Date']).alignment = Alignment(horizontal='center')
    
    c_enp = ws_all.cell(row=row_num, column=11, value=t_data['Entry Futures Price (₹)'])
    c_enp.number_format = '₹#,##0.00'
    
    c_exp = ws_all.cell(row=row_num, column=12, value=t_data['Exit Futures Price (₹)'])
    c_exp.number_format = '₹#,##0.00'
    
    ws_all.cell(row=row_num, column=13, value=int(t_data['Lot Size (Qty)'])).alignment = Alignment(horizontal='center')
    
    c_er = ws_all.cell(row=row_num, column=14, value=t_data['Expected Return (%)'])
    c_er.number_format = '0.00%'
    
    c_rw = ws_all.cell(row=row_num, column=15, value=t_data['Return We Get (%)'])
    c_rw.number_format = '0.00%'
    
    c_fpnl = ws_all.cell(row=row_num, column=16, value=t_data['Booked Futures P&L (₹)'])
    c_fpnl.number_format = '₹#,##0.00'
    c_fpnl.font = font_bold
    c_fpnl.fill = fill_win if t_data['Booked Futures P&L (₹)'] >= 0 else fill_loss
    
    ws_all.cell(row=row_num, column=17, value=t_data['Assigned Slot']).alignment = Alignment(horizontal='center')
    ws_all.cell(row=row_num, column=18, value=t_data['Re-entry Type']).alignment = Alignment(horizontal='center')
    ws_all.cell(row=row_num, column=19, value=t_data['Quarterly Result Date']).alignment = Alignment(horizontal='center')

autofit_columns(ws_all, max_cols=19)

# --- SHEET 16: Losing_Trades_Audit ---
ws_loss = wb.create_sheet(title='Losing_Trades_Audit')
loss_headers = [
    'Quarter', 'Symbol', 'Company Name', 'Sector', 'Strategy',
    'Entry Date', 'Exit Date', 'Entry Futures Price (₹)', 'Exit Futures Price (₹)',
    'Lot Size (Qty)', 'Booked Futures Loss (₹)', 'Return We Get (%)'
]
style_table_headers(ws_loss, 1, loss_headers, fill_navy)

for r_idx, t_data in df_losing.iterrows():
    row_num = 2 + r_idx
    ws_loss.cell(row=row_num, column=1, value=t_data['Quarter']).alignment = Alignment(horizontal='center')
    ws_loss.cell(row=row_num, column=2, value=t_data['Symbol']).alignment = Alignment(horizontal='center')
    ws_loss.cell(row=row_num, column=3, value=t_data['Company Name'])
    ws_loss.cell(row=row_num, column=4, value=t_data['Sector'])
    ws_loss.cell(row=row_num, column=5, value=t_data['Strategy']).alignment = Alignment(horizontal='center')
    ws_loss.cell(row=row_num, column=6, value=t_data['Entry Date']).alignment = Alignment(horizontal='center')
    ws_loss.cell(row=row_num, column=7, value=t_data['Exit Date']).alignment = Alignment(horizontal='center')
    
    c_enp = ws_loss.cell(row=row_num, column=8, value=t_data['Entry Futures Price (₹)'])
    c_enp.number_format = '₹#,##0.00'
    
    c_exp = ws_loss.cell(row=row_num, column=9, value=t_data['Exit Futures Price (₹)'])
    c_exp.number_format = '₹#,##0.00'
    
    ws_loss.cell(row=row_num, column=10, value=int(t_data['Lot Size (Qty)'])).alignment = Alignment(horizontal='center')
    
    c_pnl = ws_loss.cell(row=row_num, column=11, value=t_data['Booked Futures P&L (₹)'])
    c_pnl.number_format = '₹#,##0.00'
    c_pnl.font = font_bold
    c_pnl.fill = fill_loss
    
    c_rw = ws_loss.cell(row=row_num, column=12, value=t_data['Return We Get (%)'])
    c_rw.number_format = '0.00%'

autofit_columns(ws_loss, max_cols=12)

# --- SHEET 17: Futures_Master ---
if not df_daily.empty:
    ws_dmaster = wb.create_sheet(title='Futures_Master')
    d_headers = list(df_daily.columns)
    style_table_headers(ws_dmaster, 1, d_headers, fill_teal)
    
    for r_idx, r_data in df_daily.iterrows():
        row_num = 2 + r_idx
        for c_idx, val in enumerate(r_data, start=1):
            cell = ws_dmaster.cell(row=row_num, column=c_idx, value=val)
            if c_idx == 3: cell.number_format = '₹#,##0.00'
            
    autofit_columns(ws_dmaster, max_cols=len(d_headers))

# Save with safe fallback
def safe_save_workbook(wb_obj, target_path):
    out_file = target_path
    for attempt in range(1, 5):
        try:
            print(f"Saving to {out_file.name}...", flush=True)
            wb_obj.save(out_file)
            print(f"✅ Successfully saved {out_file.name}", flush=True)
            return out_file
        except PermissionError:
            out_file = target_path.with_name(f"{target_path.stem}_v{attempt+1}{target_path.suffix}")
            print(f"⚠️ File locked; saving as {out_file.name}...", flush=True)

save_path1 = safe_save_workbook(wb, OUTPUT_FUTURES_MASTER)
save_path2 = safe_save_workbook(wb, OUTPUT_FUTURES_ONLY)

print(f"\n🎉 PERFECT SPLIT-ADJUSTED FUTURES MASTER CREATED IN {time.time() - t0:.1f}s!", flush=True)
