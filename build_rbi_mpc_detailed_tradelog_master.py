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
OUTPUT_FILE = ROOT / 'Nifty211_RBI_Policy_Detailed_Trade_Log_Master.xlsx'

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
        "hist_q_str": hist_q_str
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
    {"code": "RBI_Oct_2023", "label": "RBI MPC Policy Oct 2023", "date_str": "2023-10-06", "period": "FY 2023-24 Q3 Policy", "ann": "Repo Rate Decision & Stance"},
    {"code": "RBI_Dec_2023", "label": "RBI MPC Policy Dec 2023", "date_str": "2023-12-08", "period": "FY 2023-24 Q3 Policy", "ann": "Repo Rate Decision & Stance"},
    {"code": "RBI_Feb_2024", "label": "RBI MPC Policy Feb 2024", "date_str": "2024-02-08", "period": "FY 2023-24 Q4 Policy", "ann": "Repo Rate Decision & Stance"},
    {"code": "RBI_Apr_2024", "label": "RBI MPC Policy Apr 2024", "date_str": "2024-04-05", "period": "FY 2024-25 Q1 Policy", "ann": "Repo Rate Decision & Stance"},
    {"code": "RBI_Jun_2024", "label": "RBI MPC Policy Jun 2024", "date_str": "2024-06-07", "period": "FY 2024-25 Q1 Policy", "ann": "Repo Rate Decision & Stance"},
    {"code": "RBI_Aug_2024", "label": "RBI MPC Policy Aug 2024", "date_str": "2024-08-08", "period": "FY 2024-25 Q2 Policy", "ann": "Repo Rate Decision & Stance"},
    {"code": "RBI_Oct_2024", "label": "RBI MPC Policy Oct 2024", "date_str": "2024-10-09", "period": "FY 2024-25 Q3 Policy", "ann": "Repo Rate Decision & Stance"},
    {"code": "RBI_Dec_2024", "label": "RBI MPC Policy Dec 2024", "date_str": "2024-12-06", "period": "FY 2024-25 Q3 Policy", "ann": "Repo Rate Decision & Stance"},
    {"code": "RBI_Feb_2025", "label": "RBI MPC Policy Feb 2025", "date_str": "2025-02-07", "period": "FY 2024-25 Q4 Policy", "ann": "Repo Rate Cut Announcement"},
    {"code": "RBI_Apr_2025", "label": "RBI MPC Policy Apr 2025", "date_str": "2025-04-09", "period": "FY 2025-26 Q1 Policy", "ann": "Repo Rate Decision & Stance"},
    {"code": "RBI_Jun_2025", "label": "RBI MPC Policy Jun 2025", "date_str": "2025-06-06", "period": "FY 2025-26 Q1 Policy", "ann": "Repo Rate Decision & Stance"},
    {"code": "RBI_Aug_2025", "label": "RBI MPC Policy Aug 2025", "date_str": "2025-08-08", "period": "FY 2025-26 Q2 Policy", "ann": "Repo Rate Decision & Stance"},
]

print("Target Universe         : 211 Nifty F&O / Broad Market Stocks")
print("Target Policy Events    : 12 RBI Monetary Policy Announcements (Oct 2023 – Aug 2025)")
print("-" * 90)

rbi_trades = {p['label']: [] for p in RBI_POLICY_EVENTS}

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
    
    for p_info in RBI_POLICY_EVENTS:
        p_label = p_info['label']
        p_dt = pd.to_datetime(p_info['date_str'])
        
        # If RBI Policy Date is beyond available price history for this stock, mark NA
        if p_dt < min_date or p_dt > max_date:
            rbi_trades[p_label].append({
                "symbol": sym, "sector": "Equity", "status": stock_status, "n_quarters": stock_info['hist_q_str'],
                "strategy": raw_strat, "window": window_str,
                "broadcast_date": "N/A", "entry_date": "N/A", "entry_price": "N/A",
                "exit_date": "N/A", "exit_price": "N/A", "lot_size": lot_size,
                "pos_value": "N/A", "margin_req": "N/A", "net_pnl": "N/A",
                "ret_pct": "N/A", "win_loss": "N/A", "is_valid": False
            })
            continue

        res_idx = min(range(len(valid_dates)), key=lambda i: abs((valid_dates[i] - p_dt).days))
        entry_idx = max(0, res_idx - pre_days)
        exit_idx = min(len(valid_dates) - 1, res_idx + post_days)
        
        entry_dt = valid_dates[entry_idx]
        exit_dt = valid_dates[exit_idx]
        
        p_entry = float(df_prices.loc[entry_idx, 'adj'])
        p_exit = float(df_prices.loc[exit_idx, 'adj'])
        
        pos_val = round(lot_size * p_entry, 2)
        margin_req = round(pos_val * 0.20, 2)
        
        if trade_dir == "LONG":
            gross_pnl = (p_exit - p_entry) * lot_size
        else:
            gross_pnl = (p_entry - p_exit) * lot_size
            
        comb_turnover = (p_entry + p_exit) * lot_size
        cost = comb_turnover * 0.0005
        net_pnl = round(gross_pnl - cost, 2)
        ret_pct = round((net_pnl / pos_val) * 100, 2) if pos_val > 0 else 0.0
        win_loss = "WIN" if net_pnl > 0 else "LOSS"
        
        rbi_trades[p_label].append({
            "symbol": sym, "sector": "Equity", "status": stock_status, "n_quarters": stock_info['hist_q_str'],
            "strategy": raw_strat, "window": window_str, "broadcast_date": p_dt.strftime("%Y-%m-%d"),
            "entry_date": entry_dt.strftime("%Y-%m-%d"), "entry_price": p_entry,
            "exit_date": exit_dt.strftime("%Y-%m-%d"), "exit_price": p_exit,
            "lot_size": lot_size, "pos_value": pos_val, "margin_req": margin_req,
            "net_pnl": net_pnl, "ret_pct": ret_pct, "win_loss": win_loss,
            "is_valid": True
        })

print("\nVERIFICATION OF EXECUTED vs N/A TRADES ACROSS ALL 12 RBI POLICY EVENTS:")
print(f"{'RBI Policy Event':<24} | {'Total Stocks':<12} | {'Executed Trades':<16} | {'N/A Trades':<12} | {'QUALIFIED Exec':<15} | {'Net P&L (₹)':<15}")
print("-" * 105)

rbi_summary_stats = []
for p in RBI_POLICY_EVENTS:
    p_label = p['label']
    trades = rbi_trades[p_label]
    n_total = len(trades)
    n_exec = sum(1 for t in trades if t['is_valid'])
    n_na = n_total - n_exec
    n_qual_exec = sum(1 for t in trades if t['is_valid'] and t['status'] == 'QUALIFIED')
    
    valid_trades = [t for t in trades if t['is_valid']]
    net_pnl = sum(t['net_pnl'] for t in valid_trades)
    
    print(f"{p_label:<24} | {n_total:<12} | {n_exec:<16} | {n_na:<12} | {n_qual_exec:<15} | {net_pnl:<15,.2f}")
    rbi_summary_stats.append({
        "info": p, "n_total": n_total, "n_exec": n_exec, "n_na": n_na, "n_qual_exec": n_qual_exec, "trades": trades
    })

print("="*105)

# Build Master Excel Workbook for RBI Policy Events
print("\nGenerating Master Workbook: Nifty211_RBI_Policy_Detailed_Trade_Log_Master.xlsx ...")
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
fill_na = PatternFill("solid", fgColor="F2F2F2")

f_qual = Font(name="Calibri", size=10, bold=True, color="276A3C")
f_mon = Font(name="Calibri", size=10, bold=True, color="B25900")
f_loss = Font(name="Calibri", size=10, bold=True, color="9C0006")
f_win = Font(name="Calibri", size=10, bold=True, color="276A3C")
f_na = Font(name="Calibri", size=10, italic=True, color="7F7F7F")

thin_side = Side(style='thin', color='D9D9D9')
border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

# ----------------------------------------------------
# SHEET 1: MASTER SUMMARY
# ----------------------------------------------------
ws_sum = wb.active
ws_sum.title = "MASTER SUMMARY"
ws_sum.views.sheetView[0].showGridLines = True

ws_sum.merge_cells("A1:K1")
ws_sum["A1"] = "RBI MONETARY POLICY COMMITTEE (MPC) DETAILED TRADE LOG & PERFORMANCE MASTER"
ws_sum["A1"].font = f_title
ws_sum["A1"].fill = fill_navy
ws_sum["A1"].alignment = Alignment(horizontal="center", vertical="center")

ws_sum.merge_cells("A2:K2")
ws_sum["A2"] = "Comprehensive Audit of RBI Interest Rate Decisions Across 211 Nifty Stocks (Strict N/A for Out-of-Range Dates)"
ws_sum["A2"].font = f_subtitle
ws_sum["A2"].fill = fill_navy
ws_sum["A2"].alignment = Alignment(horizontal="center", vertical="center")

ws_sum.row_dimensions[1].height = 28
ws_sum.row_dimensions[2].height = 18

tot_all_stocks_slots = sum(qs['n_total'] for qs in rbi_summary_stats)
tot_exec_trades = sum(qs['n_exec'] for qs in rbi_summary_stats)
tot_na_trades = sum(qs['n_na'] for qs in rbi_summary_stats)
tot_qual_exec_trades = sum(qs['n_qual_exec'] for qs in rbi_summary_stats)

all_qual_exec_trades = [t for qs in rbi_summary_stats for t in qs['trades'] if t['is_valid'] and t['status'] == 'QUALIFIED']
qual_wins = sum(1 for t in all_qual_exec_trades if t['win_loss'] == 'WIN')
qual_wr = (qual_wins / len(all_qual_exec_trades) * 100) if len(all_qual_exec_trades) > 0 else 0.0
qual_net_pnl = sum(t['net_pnl'] for t in all_qual_exec_trades)

all_exec_trades_list = [t for qs in rbi_summary_stats for t in qs['trades'] if t['is_valid']]
all_wins = sum(1 for t in all_exec_trades_list if t['win_loss'] == 'WIN')
all_wr = (all_wins / len(all_exec_trades_list) * 100) if len(all_exec_trades_list) > 0 else 0.0
all_net_pnl = sum(t['net_pnl'] for t in all_exec_trades_list)

cards_data = [
    ("TOTAL RBI TRADES EXECUTED", f"{tot_exec_trades:,} / {tot_all_stocks_slots:,}\n({tot_na_trades:,} N/A Out of Range)", "A4:C5"),
    ("QUALIFIED EXECUTED TRADES", f"{tot_qual_exec_trades:,} ({tot_qual_exec_trades/tot_exec_trades*100:.1f}%)", "D4:F5"),
    ("QUALIFIED NET P&L", f"₹{qual_net_pnl:,.2f}", "G4:I5"),
    ("QUALIFIED WIN RATE", f"{qual_wr:.2f}%", "J4:K5")
]

for title, val, rng in cards_data:
    ws_sum.merge_cells(rng)
    top_cell = ws_sum[rng.split(":")[0]]
    top_cell.value = f"{title}\n{val}"
    top_cell.font = f_card_val
    top_cell.fill = fill_card
    top_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    start_col, start_row = rng.split(":")[0][0], int(rng.split(":")[0][1:])
    end_col, end_row = rng.split(":")[1][0], int(rng.split(":")[1][1:])
    for r in range(start_row, end_row+1):
        for c in range(ord(start_col)-64, ord(end_col)-64+1):
            ws_sum.cell(row=r, column=c).border = border_all

ws_sum.row_dimensions[4].height = 20
ws_sum.row_dimensions[5].height = 28

ws_sum.merge_cells("A7:K7")
ws_sum["A7"] = "🏛️ RBI MONETARY POLICY BREAKDOWN BY POLICY EVENT (211 UNIVERSE COVERAGE)"
ws_sum["A7"].font = f_sec_hdr
ws_sum["A7"].fill = fill_subnavy
ws_sum["A7"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

headers_sum = [
    "RBI Policy Event", "Policy Period", "Policy Type / Stance", "Total Stocks", 
    "Executed Trades", "N/A Out of Range", "Qual. Win Rate %", 
    "Qual. Net P&L (₹)", "Overall Net P&L (₹)", "Total Margin Required (₹)", "Coverage Status"
]

ws_sum.row_dimensions[8].height = 24
for col_idx, h in enumerate(headers_sum, 1):
    cell = ws_sum.cell(row=8, column=col_idx, value=h)
    cell.font = f_tbl_hdr
    cell.fill = fill_darknavy
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = border_all

row_idx = 9
for qs in rbi_summary_stats:
    q_info = qs['info']
    trades = qs['trades']
    
    q_qual_exec = [t for t in trades if t['is_valid'] and t['status'] == 'QUALIFIED']
    q_qual_wins = sum(1 for t in q_qual_exec if t['win_loss'] == 'WIN')
    q_qual_wr = (q_qual_wins / len(q_qual_exec) * 100) if len(q_qual_exec) > 0 else 0.0
    q_qual_pnl = sum(t['net_pnl'] for t in q_qual_exec)
    
    q_all_exec = [t for t in trades if t['is_valid']]
    q_all_pnl = sum(t['net_pnl'] for t in q_all_exec)
    q_tot_margin = sum(t['margin_req'] for t in q_all_exec)
    
    if qs['n_na'] == 0:
        status_str = "100% Date Coverage"
    else:
        status_str = f"{qs['n_exec']} Executed · {qs['n_na']} N/A"
    
    vals = [
        q_info['label'], q_info['period'], q_info['ann'], qs['n_total'],
        qs['n_exec'], qs['n_na'], round(q_qual_wr, 2) if len(q_qual_exec) > 0 else "N/A",
        round(q_qual_pnl, 2), round(q_all_pnl, 2), round(q_tot_margin, 2), status_str
    ]
    
    ws_sum.row_dimensions[row_idx].height = 20
    for col_i, v in enumerate(vals, 1):
        cell = ws_sum.cell(row=row_idx, column=col_i, value=v)
        cell.font = f_data
        cell.border = border_all
        
        if col_i in [1, 2, 3, 11]:
            cell.alignment = Alignment(horizontal="center", vertical="center")
        elif col_i in [4, 5, 6]:
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = "#,##0"
        elif col_i == 7:
            cell.alignment = Alignment(horizontal="right", vertical="center")
            if isinstance(v, (int, float)):
                cell.number_format = "0.00\"%\""
                cell.font = f_win if v >= 50 else f_loss
            else:
                cell.font = f_na
        elif col_i in [8, 9, 10]:
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = "₹#,##0.00;[Red]-₹#,##0.00"
            if col_i in [8, 9]:
                cell.font = f_win if v > 0 else (f_loss if v < 0 else f_data)
        if col_i == 11:
            cell.font = f_qual if qs['n_na'] == 0 else f_mon
            
    row_idx += 1

# Total Summary Row
ws_sum.row_dimensions[row_idx].height = 24
tot_vals = [
    "TOTAL (12 RBI POLICIES)", "Oct 2023 - Aug 2025", "12 Rate Announcements", tot_all_stocks_slots,
    tot_exec_trades, tot_na_trades, round(qual_wr, 2),
    round(qual_net_pnl, 2), round(all_net_pnl, 2), "-", f"{tot_exec_trades} Executed Trades"
]

for col_i, v in enumerate(tot_vals, 1):
    cell = ws_sum.cell(row=row_idx, column=col_i, value=v)
    cell.font = f_bold
    cell.fill = fill_tot
    cell.border = border_all
    if col_i in [1, 2, 3, 11]:
        cell.alignment = Alignment(horizontal="center", vertical="center")
    elif col_i in [4, 5, 6]:
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.number_format = "#,##0"
    elif col_i == 7:
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.number_format = "0.00\"%\""
    elif col_i in [8, 9]:
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.number_format = "₹#,##0.00;[Red]-₹#,##0.00"

# Dynamic RBI Policy Fund Allocation Table
row_idx += 3
ws_sum.merge_cells(f"A{row_idx}:K{row_idx}")
ws_sum[f"A{row_idx}"] = "💰 DYNAMIC RBI POLICY EVENT FUND ALLOCATION & PERFORMANCE RETURN %"
ws_sum[f"A{row_idx}"].font = f_sec_hdr
ws_sum[f"A{row_idx}"].fill = fill_subnavy
ws_sum[f"A{row_idx}"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

row_idx += 1
fund_headers = [
    "RBI Policy Event", "Policy Period", "Total Stocks", "Executed Trades", 
    "Qual. Win Rate %", "New Policy Fund Required (Margin ₹)", "Qual. Net P&L (₹)", 
    "Qual. Return on Required Fund %", "Overall Net P&L (₹)", "Overall Return on Required Fund %", "Policy-End Fund Action"
]
ws_sum.row_dimensions[row_idx].height = 24

for col_i, h in enumerate(fund_headers, 1):
    cell = ws_sum.cell(row=row_idx, column=col_i, value=h)
    cell.font = f_tbl_hdr
    cell.fill = fill_darknavy
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = border_all

row_idx += 1
tot_qual_pnl_all = 0.0
tot_all_pnl_all = 0.0
tot_margin_all = 0.0
tot_qual_margin_all = 0.0

for qs in rbi_summary_stats:
    q_info = qs['info']
    trades = qs['trades']
    q_qual_exec = [t for t in trades if t['is_valid'] and t['status'] == 'QUALIFIED']
    q_all_exec = [t for t in trades if t['is_valid']]
    
    q_qual_pnl = sum(t['net_pnl'] for t in q_qual_exec)
    q_all_pnl = sum(t['net_pnl'] for t in q_all_exec)
    
    q_qual_margin = sum(t['margin_req'] for t in q_qual_exec)
    q_total_margin = sum(t['margin_req'] for t in q_all_exec)
    
    tot_qual_pnl_all += q_qual_pnl
    tot_all_pnl_all += q_all_pnl
    tot_margin_all += q_total_margin
    tot_qual_margin_all += q_qual_margin
    
    q_qual_wins = sum(1 for t in q_qual_exec if t['win_loss'] == 'WIN')
    q_qual_wr = (q_qual_wins / len(q_qual_exec) * 100) if len(q_qual_exec) > 0 else 0.0
    
    ret_qual_fund = (q_qual_pnl / q_qual_margin * 100) if q_qual_margin > 0 else 0.0
    ret_overall_fund = (q_all_pnl / q_total_margin * 100) if q_total_margin > 0 else 0.0
    
    if len(q_all_exec) == 0:
        action_str = "No Date Coverage (N/A)"
    else:
        action_str = "Profit Booked & Capital Released" if q_all_pnl >= 0 else "Loss Booked & Capital Released"
    
    row_vals = [
        q_info['label'], q_info['period'], qs['n_total'], qs['n_exec'],
        round(q_qual_wr, 2) if len(q_qual_exec) > 0 else "N/A",
        round(q_total_margin, 2), round(q_qual_pnl, 2), 
        round(ret_qual_fund, 2) if q_qual_margin > 0 else "N/A",
        round(q_all_pnl, 2), round(ret_overall_fund, 2) if q_total_margin > 0 else "N/A", action_str
    ]
    
    ws_sum.row_dimensions[row_idx].height = 20
    for col_i, v in enumerate(row_vals, 1):
        cell = ws_sum.cell(row=row_idx, column=col_i, value=v)
        cell.font = f_data
        cell.border = border_all
        if col_i in [1, 2, 11]:
            cell.alignment = Alignment(horizontal="center", vertical="center")
        elif col_i in [3, 4]:
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = "#,##0"
        elif col_i in [5, 8, 10]:
            cell.alignment = Alignment(horizontal="right", vertical="center")
            if isinstance(v, (int, float)):
                cell.number_format = "0.00\"%\""
                cell.font = f_win if v > 0 else (f_loss if v < 0 else f_data)
            else:
                cell.font = f_na
        elif col_i in [6, 7, 9]:
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = "₹#,##0.00;[Red]-₹#,##0.00"
            if col_i in [7, 9]:
                cell.font = f_win if v > 0 else (f_loss if v < 0 else f_data)
            
    row_idx += 1

ws_sum.row_dimensions[row_idx].height = 24
avg_quarter_margin = tot_margin_all / len(rbi_summary_stats)
overall_qual_ret_fund = (tot_qual_pnl_all / tot_qual_margin_all * 100) if tot_qual_margin_all > 0 else 0.0
overall_all_ret_fund = (tot_all_pnl_all / tot_margin_all * 100) if tot_margin_all > 0 else 0.0

tot_fund_vals = [
    "TOTAL (12 RBI POLICIES)", "Oct 2023 - Aug 2025", tot_all_stocks_slots, tot_exec_trades,
    round(qual_wr, 2), round(avg_quarter_margin, 2), round(tot_qual_pnl_all, 2),
    round(overall_qual_ret_fund, 2), round(tot_all_pnl_all, 2), round(overall_all_ret_fund, 2), "Fresh Fund Allocated Per RBI Policy"
]

for col_i, v in enumerate(tot_fund_vals, 1):
    cell = ws_sum.cell(row=row_idx, column=col_i, value=v)
    cell.font = f_bold
    cell.fill = fill_tot
    cell.border = border_all
    if col_i in [1, 2, 11]:
        cell.alignment = Alignment(horizontal="center", vertical="center")
    elif col_i in [3, 4]:
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.number_format = "#,##0"
    elif col_i in [5, 8, 10]:
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.number_format = "0.00\"%\""
        cell.font = f_win if v > 0 else f_loss
    elif col_i in [6, 7, 9]:
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.number_format = "₹#,##0.00;[Red]-₹#,##0.00"
        if col_i in [7, 9]:
            cell.font = f_win if v > 0 else f_loss

for col in ws_sum.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws_sum.column_dimensions[col_letter].width = max(max_len + 4, 15)

ws_sum.column_dimensions['A'].width = 24
ws_sum.column_dimensions['B'].width = 22
ws_sum.column_dimensions['C'].width = 24
ws_sum.column_dimensions['D'].width = 28

# ----------------------------------------------------
# SHEETS 2 to 13: INDIVIDUAL RBI POLICY SHEETS
# ----------------------------------------------------
for qs in rbi_summary_stats:
    q_info = qs['info']
    q_code = q_info['code']
    q_label = q_info['label']
    trades = qs['trades']
    
    ws = wb.create_sheet(title=q_code)
    ws.views.sheetView[0].showGridLines = True
    
    ws.merge_cells("A1:Q1")
    ws["A1"] = f"{q_label.upper()} · MONETARY POLICY DETAILED TRADE LOG"
    ws["A1"].font = f_title
    ws["A1"].fill = fill_navy
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    
    ws.merge_cells("A2:Q2")
    ws["A2"] = f"🏛️ RBI Policy Date: {q_info['date_str']}  |  Policy Focus: {q_info['ann']}  |  Execution Engine: 20% Futures Margin + 0.05% Friction"
    ws["A2"].font = f_subtitle
    ws["A2"].fill = fill_navy
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    
    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 18
    
    ws.merge_cells("A3:Q3")
    ws["A3"] = "ℹ️ QUALIFIED STOCKS: ≥12 Quarters History  |  AVOID BUT MONITOR IT: <12 Quarters History  |  N/A: Out of Price Data Range"
    ws["A3"].font = Font(name="Calibri", size=9, bold=True, color="1F4E79")
    ws["A3"].fill = PatternFill("solid", fgColor="E6EEF8")
    ws["A3"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[3].height = 18
    
    n_tot = len(trades)
    n_exec = sum(1 for t in trades if t['is_valid'])
    n_na = n_tot - n_exec
    
    q_qual_exec = [t for t in trades if t['is_valid'] and t['status'] == 'QUALIFIED']
    n_qual_exec = len(q_qual_exec)
    qual_wins = sum(1 for t in q_qual_exec if t['win_loss'] == 'WIN')
    qual_wr = (qual_wins / n_qual_exec * 100) if n_qual_exec > 0 else 0.0
    qual_pnl = sum(t['net_pnl'] for t in q_qual_exec)

    all_exec = [t for t in trades if t['is_valid']]
    all_wins = sum(1 for t in all_exec if t['win_loss'] == 'WIN')
    all_wr = (all_wins / n_exec * 100) if n_exec > 0 else 0.0
    all_pnl = sum(t['net_pnl'] for t in all_exec)
    tot_margin = sum(t['margin_req'] for t in all_exec)
    ret_on_q_fund = (all_pnl / tot_margin * 100) if tot_margin > 0 else 0.0
    
    kpi_items = [
        ("TOTAL STOCKS", f"{n_tot}", "A5:B6"),
        ("EXECUTED TRADES", f"{n_exec} ({n_exec/n_tot*100:.1f}%)", "C5:D6"),
        ("N/A OUT OF RANGE", f"{n_na} ({n_na/n_tot*100:.1f}%)", "E5:F6"),
        ("QUAL. WIN RATE", f"{qual_wr:.2f}%" if n_qual_exec > 0 else "N/A", "G5:H6"),
        ("QUAL. NET P&L", f"₹{qual_pnl:,.2f}", "I5:K6"),
        ("NEW POLICY FUND (MARGIN)", f"₹{tot_margin:,.2f}", "L5:N6"),
        ("RETURN ON POLICY FUND %", f"{ret_on_q_fund:.2f}%" if tot_margin > 0 else "N/A", "O5:Q6")
    ]
    
    for title, val, rng in kpi_items:
        ws.merge_cells(rng)
        top_cell = ws[rng.split(":")[0]]
        top_cell.value = f"{title}\n{val}"
        top_cell.font = f_card_val
        top_cell.fill = fill_card
        top_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        start_col, start_row = rng.split(":")[0][0], int(rng.split(":")[0][1:])
        end_col, end_row = rng.split(":")[1][0], int(rng.split(":")[1][1:])
        for r in range(start_row, end_row+1):
            for c in range(openpyxl.utils.column_index_from_string(start_col), openpyxl.utils.column_index_from_string(end_col)+1):
                ws.cell(row=r, column=c).border = border_all

    ws.row_dimensions[5].height = 18
    ws.row_dimensions[6].height = 24
    
    ws.merge_cells("A8:Q8")
    ws["A8"] = f"DETAILED TRADE LOG — {n_exec} EXECUTED TRADES ({n_qual_exec} QUALIFIED) · {n_na} N/A OUT OF RANGE STOCKS"
    ws["A8"].font = f_sec_hdr
    ws["A8"].fill = fill_subnavy
    ws["A8"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    
    headers_trade = [
        "Trade #", "Stock Symbol", "Sector", "Stock Status", "Strategy", "Window",
        "RBI Policy Date", "Entry Date", "Entry Price (₹)", "Exit Date", "Exit Price (₹)",
        "Lot Size", "Position Value (₹)", "Margin Req. (20%)", "Net P&L (₹)", "Return %", "Win / Loss"
    ]
    
    ws.row_dimensions[9].height = 24
    for col_i, h in enumerate(headers_trade, 1):
        cell = ws.cell(row=9, column=col_i, value=h)
        cell.font = f_tbl_hdr
        cell.fill = fill_darknavy
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_all

    # Sort trades: Executed Qualified first, then Executed Monitor, then N/A trades
    trades_sorted = sorted(trades, key=lambda x: (
        0 if (x['is_valid'] and x['status'] == 'QUALIFIED') else (1 if x['is_valid'] else 2),
        -x['net_pnl'] if isinstance(x['net_pnl'], (int, float)) else 0
    ))
    
    r_idx = 10
    for t_i, t in enumerate(trades_sorted, 1):
        ws.row_dimensions[r_idx].height = 20
        row_vals = [
            t_i, t['symbol'], t['sector'], t['status'], t['strategy'], t['window'],
            t['broadcast_date'], t['entry_date'], t['entry_price'], t['exit_date'], t['exit_price'],
            t['lot_size'], t['pos_value'], t['margin_req'], t['net_pnl'], t['ret_pct'], t['win_loss']
        ]
        
        for c_i, v in enumerate(row_vals, 1):
            cell = ws.cell(row=r_idx, column=c_i, value=v)
            cell.border = border_all
            
            if not t['is_valid']:
                cell.fill = fill_na
                cell.font = f_na
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.font = f_data
                if t['status'] == 'QUALIFIED':
                    cell.fill = fill_qual if t['win_loss'] == 'WIN' else fill_card
                else:
                    cell.fill = fill_mon
                    
                if c_i in [1, 2, 3, 5, 6, 7, 8, 10]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif c_i == 4:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.font = f_qual if v == 'QUALIFIED' else f_mon
                elif c_i in [9, 11, 13, 14, 15]:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    cell.number_format = "₹#,##0.00;[Red]-₹#,##0.00"
                    if c_i == 15:
                        cell.font = f_win if v > 0 else f_loss
                elif c_i == 12:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    cell.number_format = "#,##0"
                elif c_i == 16:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    cell.number_format = "0.00\"%\""
                    cell.font = f_win if v > 0 else f_loss
                elif c_i == 17:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.font = f_win if v == 'WIN' else f_loss
                
        r_idx += 1
        
    ws.row_dimensions[r_idx].height = 22
    q_valid_trades = [t for t in trades_sorted if t['is_valid'] and t['status'] == 'QUALIFIED']
    q_pnl_sum = sum(t['net_pnl'] for t in q_valid_trades)
    q_margin_sum = sum(t['margin_req'] for t in q_valid_trades)
    q_pos_sum = sum(t['pos_value'] for t in q_valid_trades)
    
    sub_q_vals = [
        "SUBTOTAL", "QUALIFIED EXECUTED ONLY", "-", f"{len(q_valid_trades)} Trades", "-", "-",
        "-", "-", "-", "-", "-",
        "-", q_pos_sum, q_margin_sum, q_pnl_sum, (q_pnl_sum/q_pos_sum*100) if q_pos_sum>0 else 0.0, f"{qual_wr:.1f}% WR" if len(q_valid_trades)>0 else "N/A"
    ]
    for c_i, v in enumerate(sub_q_vals, 1):
        cell = ws.cell(row=r_idx, column=c_i, value=v)
        cell.font = f_bold
        cell.fill = PatternFill("solid", fgColor="E2EFDA")
        cell.border = border_all
        if c_i in [1, 2, 4, 17]:
            cell.alignment = Alignment(horizontal="center", vertical="center")
        elif c_i in [13, 14, 15]:
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = "₹#,##0.00;[Red]-₹#,##0.00"
            cell.font = f_win if c_i == 15 and isinstance(v, (int, float)) and v > 0 else f_bold
        elif c_i == 16:
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = "0.00\"%\""
            
    r_idx += 1

    ws.row_dimensions[r_idx].height = 24
    all_valid_trades = [t for t in trades_sorted if t['is_valid']]
    all_pnl_sum = sum(t['net_pnl'] for t in all_valid_trades)
    all_margin_sum = sum(t['margin_req'] for t in all_valid_trades)
    all_pos_sum = sum(t['pos_value'] for t in all_valid_trades)
    
    tot_all_row_vals = [
        "TOTAL", "ALL EXECUTED TRADES", "-", f"{len(all_valid_trades)} Trades", "-", "-",
        "-", "-", "-", "-", "-",
        "-", all_pos_sum, all_margin_sum, all_pnl_sum, (all_pnl_sum/all_pos_sum*100) if all_pos_sum>0 else 0.0, f"{all_wr:.1f}% WR" if len(all_valid_trades)>0 else "N/A"
    ]
    for c_i, v in enumerate(tot_all_row_vals, 1):
        cell = ws.cell(row=r_idx, column=c_i, value=v)
        cell.font = f_bold
        cell.fill = fill_tot
        cell.border = border_all
        if c_i in [1, 2, 4, 17]:
            cell.alignment = Alignment(horizontal="center", vertical="center")
        elif c_i in [13, 14, 15]:
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = "₹#,##0.00;[Red]-₹#,##0.00"
            cell.font = f_win if c_i == 15 and isinstance(v, (int, float)) and v > 0 else f_bold
        elif c_i == 16:
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = "0.00\"%\""

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 13)
        
    ws.column_dimensions['A'].width = 10
    ws.column_dimensions['B'].width = 16
    ws.column_dimensions['C'].width = 14
    ws.column_dimensions['D'].width = 24
    ws.column_dimensions['E'].width = 12
    ws.column_dimensions['F'].width = 14
    ws.column_dimensions['G'].width = 16
    ws.column_dimensions['H'].width = 14
    ws.column_dimensions['I'].width = 16
    ws.column_dimensions['J'].width = 14
    ws.column_dimensions['K'].width = 16
    ws.column_dimensions['L'].width = 12
    ws.column_dimensions['M'].width = 20
    ws.column_dimensions['N'].width = 20
    ws.column_dimensions['O'].width = 18
    ws.column_dimensions['P'].width = 14
    ws.column_dimensions['Q'].width = 14

# Save RBI Policy Master Excel
saved_path = None
for candidate in [
    ROOT / 'Nifty211_RBI_Policy_Detailed_Trade_Log_Master.xlsx',
    ROOT / f'Nifty211_RBI_Policy_Detailed_Trade_Log_Master_{pd.Timestamp.now().strftime("%H%M%S")}.xlsx'
]:
    try:
        wb.save(candidate)
        saved_path = candidate
        break
    except PermissionError:
        continue

print(f"\n✅ SUCCESSFULLY BUILT AND SAVED RBI POLICY MASTER EXCEL:")
print(f"   Location: {saved_path}")
print(f"   Sheets Generated ({len(wb.sheetnames)} total): {', '.join(wb.sheetnames)}")
print("="*90)
