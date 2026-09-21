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
PORTFOLIO_MASTER = ROOT / 'Nifty211_All_Stocks_Portfolio_Master.xlsx'
OUTPUT_FILE = ROOT / 'Nifty211_RBI_Policy_Combined_Master.xlsx'

# 1. Load Portfolio Master Strategy & Window Mapping
print("="*90)
print(f"Loading stock strategy mapping from: {PORTFOLIO_MASTER.name} ...")
df_pm = pd.read_excel(PORTFOLIO_MASTER, sheet_name='ALL QUARTERS COMBINED', skiprows=6)
df_pm['Stock Symbol'] = df_pm['Stock Symbol'].astype(str).str.strip().str.upper()

STOCK_MAP = {}
for _, row in df_pm.iterrows():
    sym = row['Stock Symbol']
    strat = str(row['Strategy']).strip().upper()  # 'FUTURE LONG' or 'FUTURE SHORT'
    window_raw = str(row['Position Window']).strip()  # e.g. 'T-5 to T+8'
    status = str(row['Action Status']).strip()  # 'QUALIFIED' or 'AVOID BUT MONITOR IT'
    hist_q_str = str(row['Total History Quarters']).strip()
    sector = str(row.get('Industry Sector', 'Equity')).strip()
    
    try:
        parts = window_raw.replace('T-', '').replace('T+', '').split(' to ')
        pre_days = int(parts[0])
        post_days = int(parts[1])
    except Exception:
        pre_days, post_days = 3, 3
        
    STOCK_MAP[sym] = {
        "strategy": "LONG" if "LONG" in strat else "SHORT",
        "raw_strategy": strat,
        "window_str": window_raw,
        "pre_days": pre_days,
        "post_days": post_days,
        "status": status,
        "hist_q_str": hist_q_str,
        "sector": sector
    }

SYMBOLS = sorted(list(STOCK_MAP.keys()))

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

# 12 Official RBI Monetary Policy Committee (MPC) Announcement Events
RBI_POLICY_EVENTS = [
    {"code": "RBI_Oct_2023", "label": "RBI MPC Policy Oct 2023", "date_str": "2023-10-06", "period": "FY 2023-24 Q3 Policy"},
    {"code": "RBI_Dec_2023", "label": "RBI MPC Policy Dec 2023", "date_str": "2023-12-08", "period": "FY 2023-24 Q3 Policy"},
    {"code": "RBI_Feb_2024", "label": "RBI MPC Policy Feb 2024", "date_str": "2024-02-08", "period": "FY 2023-24 Q4 Policy"},
    {"code": "RBI_Apr_2024", "label": "RBI MPC Policy Apr 2024", "date_str": "2024-04-05", "period": "FY 2024-25 Q1 Policy"},
    {"code": "RBI_Jun_2024", "label": "RBI MPC Policy Jun 2024", "date_str": "2024-06-07", "period": "FY 2024-25 Q1 Policy"},
    {"code": "RBI_Aug_2024", "label": "RBI MPC Policy Aug 2024", "date_str": "2024-08-08", "period": "FY 2024-25 Q2 Policy"},
    {"code": "RBI_Oct_2024", "label": "RBI MPC Policy Oct 2024", "date_str": "2024-10-09", "period": "FY 2024-25 Q3 Policy"},
    {"code": "RBI_Dec_2024", "label": "RBI MPC Policy Dec 2024", "date_str": "2024-12-06", "period": "FY 2024-25 Q3 Policy"},
    {"code": "RBI_Feb_2025", "label": "RBI MPC Policy Feb 2025", "date_str": "2025-02-07", "period": "FY 2024-25 Q4 Policy"},
    {"code": "RBI_Apr_2025", "label": "RBI MPC Policy Apr 2025", "date_str": "2025-04-09", "period": "FY 2025-26 Q1 Policy"},
    {"code": "RBI_Jun_2025", "label": "RBI MPC Policy Jun 2025", "date_str": "2025-06-06", "period": "FY 2025-26 Q1 Policy"},
    {"code": "RBI_Aug_2025", "label": "RBI MPC Policy Aug 2025", "date_str": "2025-08-08", "period": "FY 2025-26 Q2 Policy"},
]

print("Calculating Consolidated Stock Performance Across All 12 RBI Policy Events ...")

stock_combined_stats = []

for sym in SYMBOLS:
    stock_info = STOCK_MAP[sym]
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
        
    valid_dates = df_prices['date'].tolist()
    min_date = valid_dates[0]
    max_date = valid_dates[-1]
    last_price = float(df_prices['adj'].iloc[-1])
    lot_size = LOT_SIZES.get(sym, max(100, int(1000000 / last_price)))
    
    stock_status = stock_info['status']
    pre_days = stock_info['pre_days']
    post_days = stock_info['post_days']
    window_str = stock_info['window_str']
    trade_dir = stock_info['strategy']
    raw_strat = stock_info['raw_strategy']
    sector = stock_info['sector']
    
    exec_trades = 0
    wins = 0
    losses = 0
    total_pnl = 0.0
    margin_cap = 0.0
    returns = []
    
    for p_info in RBI_POLICY_EVENTS:
        p_dt = pd.to_datetime(p_info['date_str'])
        
        if p_dt < min_date or p_dt > max_date:
            continue

        res_idx = min(range(len(valid_dates)), key=lambda i: abs((valid_dates[i] - p_dt).days))
        entry_idx = max(0, res_idx - pre_days)
        exit_idx = min(len(valid_dates) - 1, res_idx + post_days)
        
        p_entry = float(df_prices.loc[entry_idx, 'adj'])
        p_exit = float(df_prices.loc[exit_idx, 'adj'])
        
        pos_val = round(lot_size * p_entry, 2)
        m_req = round(pos_val * 0.20, 2)
        margin_cap = max(margin_cap, m_req)
        
        if trade_dir == "LONG":
            gross_pnl = (p_exit - p_entry) * lot_size
        else:
            gross_pnl = (p_entry - p_exit) * lot_size
            
        comb_turnover = (p_entry + p_exit) * lot_size
        cost = comb_turnover * 0.0005
        net_pnl = round(gross_pnl - cost, 2)
        ret_pct = (net_pnl / pos_val * 100) if pos_val > 0 else 0.0
        
        exec_trades += 1
        total_pnl += net_pnl
        returns.append(ret_pct)
        if net_pnl > 0:
            wins += 1
        else:
            losses += 1
            
    win_rate = (wins / exec_trades * 100) if exec_trades > 0 else 0.0
    avg_ret = np.mean(returns) if returns else 0.0
    win_ratio_str = f"{wins} Wins / {exec_trades} Events" if exec_trades > 0 else "N/A"
    
    stock_combined_stats.append({
        "symbol": sym, "sector": sector, "status": stock_status, "n_quarters": stock_info['hist_q_str'],
        "strategy": raw_strat, "window": window_str, "exec_events": exec_trades,
        "wins": wins, "losses": losses, "win_rate": round(win_rate, 2), "win_ratio_str": win_ratio_str,
        "avg_return": round(avg_ret, 2), "margin_req": round(margin_cap, 2), "total_pnl": round(total_pnl, 2)
    })

# Sort stocks by Total Net Realised P&L descending
stock_combined_stats = sorted(stock_combined_stats, key=lambda x: x['total_pnl'], reverse=True)

print("="*90)
print(f"Total Stocks Processed: {len(stock_combined_stats)}")
print("Top 5 Performing Stocks in RBI Policy Events:")
for s_i, s in enumerate(stock_combined_stats[:5], 1):
    print(f"  Rank {s_i:<2}: {s['symbol']:<12} | Status: {s['status']:<20} | Strategy: {s['strategy']:<12} | Win Rate: {s['win_rate']}% | P&L: ₹{s['total_pnl']:,.2f}")
print("="*90)

# Build Excel Workbook
print(f"\nBuilding Excel Workbook: {OUTPUT_FILE.name} ...")
wb = openpyxl.Workbook()

# Styling tokens
f_title = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
f_subtitle = Font(name="Calibri", size=10, italic=True, color="E0E0E0")
f_sec_hdr = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
f_tbl_hdr = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
f_card_val = Font(name="Calibri", size=14, bold=True, color="1F4E79")
f_data = Font(name="Calibri", size=10, color="000000")
f_bold = Font(name="Calibri", size=10, bold=True, color="000000")

fill_navy = PatternFill("solid", fgColor="1F4E79")
fill_subnavy = PatternFill("solid", fgColor="2F5597")
fill_darknavy = PatternFill("solid", fgColor="1B365D")
fill_card = PatternFill("solid", fgColor="F2F4F7")
fill_qual = PatternFill("solid", fgColor="E2EFDA")
fill_mon = PatternFill("solid", fgColor="FFF2CC")
fill_tot = PatternFill("solid", fgColor="D9E1F2")

f_qual = Font(name="Calibri", size=10, bold=True, color="276A3C")
f_mon = Font(name="Calibri", size=10, bold=True, color="B25900")
f_loss = Font(name="Calibri", size=10, bold=True, color="9C0006")
f_win = Font(name="Calibri", size=10, bold=True, color="276A3C")

thin_side = Side(style='thin', color='D9D9D9')
border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

# ----------------------------------------------------
# SHEET 1: RBI POLICIES COMBINED MASTER
# ----------------------------------------------------
ws_comb = wb.active
ws_comb.title = "RBI POLICIES COMBINED MASTER"
ws_comb.views.sheetView[0].showGridLines = True

ws_comb.merge_cells("A1:N1")
ws_comb["A1"] = "RBI MONETARY POLICY COMMITTEE (MPC) — COMBINED PERFORMANCE MASTER"
ws_comb["A1"].font = f_title
ws_comb["A1"].fill = fill_navy
ws_comb["A1"].alignment = Alignment(horizontal="center", vertical="center")

ws_comb.merge_cells("A2:N2")
ws_comb["A2"] = "Consolidated Performance Audit of 211 Nifty Stocks Across All 12 RBI Policy Events (Oct 2023 – Aug 2025)"
ws_comb["A2"].font = f_subtitle
ws_comb["A2"].fill = fill_navy
ws_comb["A2"].alignment = Alignment(horizontal="center", vertical="center")

ws_comb.row_dimensions[1].height = 28
ws_comb.row_dimensions[2].height = 18

tot_stocks = len(stock_combined_stats)
tot_qual_stocks = sum(1 for s in stock_combined_stats if s['status'] == 'QUALIFIED')
tot_qual_pnl = sum(s['total_pnl'] for s in stock_combined_stats if s['status'] == 'QUALIFIED')
tot_all_pnl = sum(s['total_pnl'] for s in stock_combined_stats)

qual_wins_sum = sum(s['wins'] for s in stock_combined_stats if s['status'] == 'QUALIFIED')
qual_events_sum = sum(s['exec_events'] for s in stock_combined_stats if s['status'] == 'QUALIFIED')
overall_qual_wr = (qual_wins_sum / qual_events_sum * 100) if qual_events_sum > 0 else 0.0

cards_data = [
    ("TOTAL STOCKS EVALUATED", f"{tot_stocks} Stocks", "A4:C5"),
    ("QUALIFIED STOCKS (≥12Q)", f"{tot_qual_stocks} ({tot_qual_stocks/tot_stocks*100:.1f}%)", "D4:F5"),
    ("QUALIFIED NET P&L", f"₹{tot_qual_pnl:,.2f}", "G4:J5"),
    ("QUALIFIED WIN RATE", f"{overall_qual_wr:.2f}%", "K4:N5")
]

for title, val, rng in cards_data:
    ws_comb.merge_cells(rng)
    top_cell = ws_comb[rng.split(":")[0]]
    top_cell.value = f"{title}\n{val}"
    top_cell.font = f_card_val
    top_cell.fill = fill_card
    top_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    start_col, start_row = rng.split(":")[0][0], int(rng.split(":")[0][1:])
    end_col, end_row = rng.split(":")[1][0], int(rng.split(":")[1][1:])
    for r in range(start_row, end_row+1):
        for c in range(ord(start_col)-64, ord(end_col)-64+1):
            ws_comb.cell(row=r, column=c).border = border_all

ws_comb.row_dimensions[4].height = 20
ws_comb.row_dimensions[5].height = 28

ws_comb.merge_cells("A7:N7")
ws_comb["A7"] = "📊 ALL 211 STOCKS RANKED BY TOTAL REALISED NET P&L ACROSS ALL 12 RBI POLICY EVENTS"
ws_comb["A7"].font = f_sec_hdr
ws_comb["A7"].fill = fill_subnavy
ws_comb["A7"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

headers_comb = [
    "Rank", "Stock Symbol", "Industry Sector", "Stock Status", "Total History", 
    "Strategy", "Position Window", "RBI Events Executed", "Win Rate %", 
    "Win Ratio", "Losing Events", "Average Return %", "Margin Capital (₹)", "Net Realised P&L (₹)"
]

ws_comb.row_dimensions[8].height = 24
for col_idx, h in enumerate(headers_comb, 1):
    cell = ws_comb.cell(row=8, column=col_idx, value=h)
    cell.font = f_tbl_hdr
    cell.fill = fill_darknavy
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = border_all

row_idx = 9
for rank_i, s in enumerate(stock_combined_stats, 1):
    ws_comb.row_dimensions[row_idx].height = 20
    vals = [
        rank_i, s['symbol'], s['sector'], s['status'], s['n_quarters'],
        s['strategy'], s['window'], s['exec_events'], s['win_rate'],
        s['win_ratio_str'], s['losses'], s['avg_return'], s['margin_req'], s['total_pnl']
    ]
    
    for col_i, v in enumerate(vals, 1):
        cell = ws_comb.cell(row=row_idx, column=col_i, value=v)
        cell.font = f_data
        cell.border = border_all
        
        if s['status'] == 'QUALIFIED':
            cell.fill = fill_qual if s['total_pnl'] > 0 else fill_card
        else:
            cell.fill = fill_mon
            
        if col_i in [1, 2, 3, 5, 6, 7, 8, 10, 11]:
            cell.alignment = Alignment(horizontal="center", vertical="center")
        elif col_i == 4:
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.font = f_qual if v == 'QUALIFIED' else f_mon
        elif col_i in [9, 12]:
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = "0.00\"%\""
            cell.font = f_win if v >= 50 else f_loss
        elif col_i in [13, 14]:
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = "₹#,##0.00;[Red]-₹#,##0.00"
            if col_i == 14:
                cell.font = f_win if v > 0 else f_loss
                
    row_idx += 1

# Subtotal Row: Qualified Stocks Only
ws_comb.row_dimensions[row_idx].height = 24
qual_stocks_list = [s for s in stock_combined_stats if s['status'] == 'QUALIFIED']
qual_pnl_tot = sum(s['total_pnl'] for s in qual_stocks_list)
qual_margin_tot = sum(s['margin_req'] for s in qual_stocks_list)
qual_avg_ret = np.mean([s['avg_return'] for s in qual_stocks_list]) if qual_stocks_list else 0.0

sub_vals = [
    "SUBTOTAL", "QUALIFIED STOCKS ONLY", "-", f"{len(qual_stocks_list)} Qualified", "-",
    "-", "-", sum(s['exec_events'] for s in qual_stocks_list), round(overall_qual_wr, 2),
    f"{qual_wins_sum} Wins / {qual_events_sum}", sum(s['losses'] for s in qual_stocks_list),
    round(qual_avg_ret, 2), round(qual_margin_tot, 2), round(qual_pnl_tot, 2)
]

for col_i, v in enumerate(sub_vals, 1):
    cell = ws_comb.cell(row=row_idx, column=col_i, value=v)
    cell.font = f_bold
    cell.fill = PatternFill("solid", fgColor="E2EFDA")
    cell.border = border_all
    if col_i in [1, 2, 4, 10]:
        cell.alignment = Alignment(horizontal="center", vertical="center")
    elif col_i in [8, 11]:
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.number_format = "#,##0"
    elif col_i in [9, 12]:
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.number_format = "0.00\"%\""
    elif col_i in [13, 14]:
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.number_format = "₹#,##0.00;[Red]-₹#,##0.00"
        cell.font = f_win if col_i == 14 and v > 0 else f_bold

row_idx += 1

# Total Row: All 211 Stocks
ws_comb.row_dimensions[row_idx].height = 24
all_events_sum = sum(s['exec_events'] for s in stock_combined_stats)
all_wins_sum = sum(s['wins'] for s in stock_combined_stats)
all_losses_sum = sum(s['losses'] for s in stock_combined_stats)
overall_all_wr = (all_wins_sum / all_events_sum * 100) if all_events_sum > 0 else 0.0
all_margin_tot = sum(s['margin_req'] for s in stock_combined_stats)
all_avg_ret = np.mean([s['avg_return'] for s in stock_combined_stats]) if stock_combined_stats else 0.0

tot_vals = [
    "TOTAL", "ALL 211 STOCKS COMBINED", "-", f"{len(stock_combined_stats)} Stocks", "-",
    "-", "-", all_events_sum, round(overall_all_wr, 2),
    f"{all_wins_sum} Wins / {all_events_sum}", all_losses_sum,
    round(all_avg_ret, 2), round(all_margin_tot, 2), round(tot_all_pnl, 2)
]

for col_i, v in enumerate(tot_vals, 1):
    cell = ws_comb.cell(row=row_idx, column=col_i, value=v)
    cell.font = f_bold
    cell.fill = fill_tot
    cell.border = border_all
    if col_i in [1, 2, 4, 10]:
        cell.alignment = Alignment(horizontal="center", vertical="center")
    elif col_i in [8, 11]:
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.number_format = "#,##0"
    elif col_i in [9, 12]:
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.number_format = "0.00\"%\""
    elif col_i in [13, 14]:
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.number_format = "₹#,##0.00;[Red]-₹#,##0.00"
        cell.font = f_win if col_i == 14 and v > 0 else f_bold

for col in ws_comb.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws_comb.column_dimensions[col_letter].width = max(max_len + 3, 13)

ws_comb.column_dimensions['A'].width = 10
ws_comb.column_dimensions['B'].width = 16
ws_comb.column_dimensions['C'].width = 24
ws_comb.column_dimensions['D'].width = 24
ws_comb.column_dimensions['E'].width = 16
ws_comb.column_dimensions['F'].width = 14
ws_comb.column_dimensions['G'].width = 14
ws_comb.column_dimensions['H'].width = 18
ws_comb.column_dimensions['I'].width = 14
ws_comb.column_dimensions['J'].width = 20
ws_comb.column_dimensions['K'].width = 14
ws_comb.column_dimensions['L'].width = 16
ws_comb.column_dimensions['M'].width = 20
ws_comb.column_dimensions['N'].width = 22

# Save Combined Master Workbook
saved_path = None
for candidate in [
    ROOT / 'Nifty211_RBI_Policy_Combined_Master.xlsx',
    ROOT / f'Nifty211_RBI_Policy_Combined_Master_{pd.Timestamp.now().strftime("%H%M%S")}.xlsx'
]:
    try:
        wb.save(candidate)
        saved_path = candidate
        break
    except PermissionError:
        continue

print(f"\n✅ SUCCESSFULLY BUILT AND SAVED COMBINED RBI POLICY MASTER EXCEL:")
print(f"   Location: {saved_path}")
print(f"   Sheets Generated ({len(wb.sheetnames)} total): {', '.join(wb.sheetnames)}")
print("="*90)
