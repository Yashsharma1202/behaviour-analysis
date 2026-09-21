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
OUTPUT_FILE = ROOT / 'Nifty211_12Q_HistoryAware_Futures_Trade_Log_Master.xlsx'

print("="*90)
print(f"Loading stock strategy & history mapping from: {PORTFOLIO_MASTER.name} ...")
df_pm = pd.read_excel(PORTFOLIO_MASTER, sheet_name='ALL QUARTERS COMBINED', skiprows=6)
df_pm['Stock Symbol'] = df_pm['Stock Symbol'].astype(str).str.strip().str.upper()

STOCK_MAP = {}
for _, row in df_pm.iterrows():
    sym = row['Stock Symbol']
    if pd.isna(sym) or not sym:
        continue
    strat = str(row['Strategy']).strip().upper()
    window_raw = str(row['Position Window']).strip()
    status = str(row['Action Status']).strip()
    hist_q_str = str(row['Total History Quarters']).strip()
    
    # Extract numerical history quarter count
    import re
    match = re.search(r'(\d+)', hist_q_str)
    hist_q_num = int(match.group(1)) if match else 5
    
    try:
        parts = window_raw.replace('T-', '').replace('T+', '').split(' to ')
        pre_days = int(parts[0])
        post_days = int(parts[1])
    except Exception:
        pre_days, post_days = 5, 5
        
    STOCK_MAP[sym] = {
        "strategy": "LONG" if "LONG" in strat else "SHORT",
        "raw_strategy": strat,
        "window_str": window_raw,
        "pre_days": pre_days,
        "post_days": post_days,
        "status": status,
        "hist_q_str": hist_q_str,
        "hist_q_num": hist_q_num
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

TARGET_QUARTERS = [
    {"q_idx": 1, "code": "Q3_FY23-24", "label": "Q3 FY2023-24", "q_type": "Q3", "target_year": 2024, "period": "Oct 2023 – Dec 2023", "ann": "Jan 2024 – Feb 2024", "default_dt": "2024-01-22"},
    {"q_idx": 2, "code": "Q4_FY23-24", "label": "Q4 FY2023-24", "q_type": "Q4", "target_year": 2024, "period": "Jan 2024 – Mar 2024", "ann": "Apr 2024 – May 2024", "default_dt": "2024-04-22"},
    {"q_idx": 3, "code": "Q1_FY24-25", "label": "Q1 FY2024-25", "q_type": "Q1", "target_year": 2024, "period": "Apr 2024 – Jun 2024", "ann": "Jul 2024 – Aug 2024", "default_dt": "2024-07-22"},
    {"q_idx": 4, "code": "Q2_FY24-25", "label": "Q2 FY2024-25", "q_type": "Q2", "target_year": 2024, "period": "Jul 2024 – Sep 2024", "ann": "Oct 2024 – Nov 2024", "default_dt": "2024-10-22"},
    {"q_idx": 5, "code": "Q3_FY24-25", "label": "Q3 FY2024-25", "q_type": "Q3", "target_year": 2025, "period": "Oct 2024 – Dec 2024", "ann": "Jan 2025 – Feb 2025", "default_dt": "2025-01-22"},
    {"q_idx": 6, "code": "Q4_FY24-25", "label": "Q4 FY2024-25", "q_type": "Q4", "target_year": 2025, "period": "Jan 2025 – Mar 2025", "ann": "Apr 2025 – May 2025", "default_dt": "2025-04-22"},
    {"q_idx": 7, "code": "Q1_FY25-26", "label": "Q1 FY2025-26", "q_type": "Q1", "target_year": 2025, "period": "Apr 2025 – Jun 2025", "ann": "Jul 2025 – Aug 2025", "default_dt": "2025-07-22"},
    {"q_idx": 8, "code": "Q2_FY25-26", "label": "Q2 FY2025-26", "q_type": "Q2", "target_year": 2025, "period": "Jul 2025 – Sep 2025", "ann": "Oct 2025 – Nov 2025", "default_dt": "2025-10-22"},
    {"q_idx": 9, "code": "Q3_FY25-26", "label": "Q3 FY2025-26", "q_type": "Q3", "target_year": 2026, "period": "Oct 2025 – Dec 2025", "ann": "Jan 2026 – Feb 2026", "default_dt": "2026-01-22"},
    {"q_idx": 10, "code": "Q4_FY25-26", "label": "Q4 FY2025-26", "q_type": "Q4", "target_year": 2026, "period": "Jan 2026 – Mar 2026", "ann": "Apr 2026 – May 2026", "default_dt": "2026-04-22"},
    {"q_idx": 11, "code": "Q1_FY26-27", "label": "Q1 FY2026-27", "q_type": "Q1", "target_year": 2026, "period": "Apr 2026 – Jun 2026", "ann": "Jul 2026 – Aug 2026", "default_dt": "2026-07-22"},
    {"q_idx": 12, "code": "Q2_FY26-27", "label": "Q2 FY2026-27", "q_type": "Q2", "target_year": 2026, "period": "Jul 2026 – Sep 2026", "ann": "Oct 2026 – Nov 2026", "default_dt": "2026-10-22"},
]

print("Target Universe          : 211 Nifty F&O / Broad Market Stocks")
print("Target Financial Quarters : 12 Quarters (History-Aware Execution & Explicit N/A Handling)")
print("-" * 90)

quarter_trades = {q['label']: [] for q in TARGET_QUARTERS}

for sym in SYMBOLS:
    stock_info = STOCK_MAP[sym]
    stock_folder = ROOT / sym
    fr_path = stock_folder / 'financial_results.csv'
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
        # Fill N/A for all quarters for this stock if price missing
        for q_info in TARGET_QUARTERS:
            quarter_trades[q_info['label']].append({
                "symbol": sym, "sector": "N/A", "status": stock_info['status'],
                "raw_strat": stock_info['raw_strategy'], "window_str": stock_info['window_str'],
                "hist_q_str": stock_info['hist_q_str'], "executed": False
            })
        continue
        
    valid_dates = df_prices['date'].tolist()
    last_price = float(df_prices['adj'].iloc[-1])
    lot_size = LOT_SIZES.get(sym, max(100, int(1000000 / last_price)))
    
    df_fr_clean = pd.DataFrame()
    if fr_path.exists() and fr_path.stat().st_size > 10:
        try:
            df_fr = pd.read_csv(fr_path, dtype=str).fillna('')
            if not df_fr.empty and 'broadCastDate' in df_fr.columns:
                df_fr['bDate'] = pd.to_datetime(df_fr['broadCastDate'], errors='coerce')
                df_fr = df_fr.dropna(subset=['bDate'])
                if 'toDate' in df_fr.columns:
                    df_fr['tDate'] = pd.to_datetime(df_fr['toDate'], errors='coerce')
                else:
                    df_fr['tDate'] = pd.NaT
                df_fr_clean = df_fr.sort_values('bDate', ascending=False)
        except Exception:
            pass

    stock_status = stock_info['status']
    pre_days = stock_info['pre_days']
    post_days = stock_info['post_days']
    window_str = stock_info['window_str']
    trade_dir = stock_info['strategy']
    raw_strat = stock_info['raw_strategy']
    hist_q_num = stock_info['hist_q_num']
    
    hist_q_dates = {"Q1": [], "Q2": [], "Q3": [], "Q4": []}
    if not df_fr_clean.empty:
        for _, row in df_fr_clean.iterrows():
            b_dt = row['bDate']
            t_dt = row['tDate']
            ref_dt = t_dt if pd.notnull(t_dt) else b_dt
            m = ref_dt.month
            if m in [4, 5, 6]: hist_q_dates["Q1"].append(b_dt)
            elif m in [7, 8, 9]: hist_q_dates["Q2"].append(b_dt)
            elif m in [10, 11, 12]: hist_q_dates["Q3"].append(b_dt)
            elif m in [1, 2, 3]: hist_q_dates["Q4"].append(b_dt)

    for q_info in TARGET_QUARTERS:
        q_label = q_info['label']
        q_idx = q_info['q_idx']
        q_type = q_info['q_type']
        target_year = q_info['target_year']
        
        # KEY RULE: Stock executes ONLY if q_idx <= stock's total available history quarters!
        if q_idx > hist_q_num:
            # Output explicit N/A for this stock in this quarter
            quarter_trades[q_label].append({
                "symbol": sym,
                "sector": "N/A",
                "status": stock_status,
                "raw_strat": raw_strat,
                "window_str": window_str,
                "hist_q_str": stock_info['hist_q_str'],
                "executed": False
            })
            continue

        # Find best announcement date for this stock & quarter
        b_dt = None
        if not df_fr_clean.empty:
            for _, row in df_fr_clean.iterrows():
                cand_b = row['bDate']
                cand_t = row['tDate']
                ref_dt = cand_t if pd.notnull(cand_t) else cand_b
                m = ref_dt.month
                y = ref_dt.year
                
                match_q = False
                if q_type == "Q1" and m in [4, 5, 6] and y == target_year: match_q = True
                elif q_type == "Q2" and m in [7, 8, 9] and y == target_year: match_q = True
                elif q_type == "Q3" and m in [10, 11, 12] and y == (target_year - 1): match_q = True
                elif q_type == "Q4" and m in [1, 2, 3] and y == target_year: match_q = True
                
                if match_q:
                    b_dt = cand_b
                    break

        if b_dt is None:
            # Fallback to historical day of month in the target quarter year
            candidates = hist_q_dates[q_type]
            if candidates:
                avg_day = int(np.mean([d.day for d in candidates]))
            else:
                avg_day = 22
                
            if q_type == "Q3": month_num = 1
            elif q_type == "Q4": month_num = 4
            elif q_type == "Q1": month_num = 7
            elif q_type == "Q2": month_num = 10
            
            try:
                b_dt = pd.Timestamp(year=target_year, month=month_num, day=min(avg_day, 28))
            except Exception:
                b_dt = pd.to_datetime(q_info['default_dt'])

        # Find entry index in price dataframe
        matching_indices = [i for i, d in enumerate(valid_dates) if d >= b_dt]
        if not matching_indices:
            idx_event = len(valid_dates) - 1
        else:
            idx_event = matching_indices[0]
            
        entry_idx = max(0, idx_event - pre_days)
        exit_idx = min(len(valid_dates) - 1, idx_event + post_days)
        
        entry_date = valid_dates[entry_idx]
        exit_date = valid_dates[exit_idx]
        
        row_entry = df_prices.iloc[entry_idx]
        row_exit = df_prices.iloc[exit_idx]
        
        entry_price = float(row_entry['adj'])
        exit_price = float(row_exit['adj'])
        
        pos_val = entry_price * lot_size
        margin_req = pos_val * 0.20
        
        if trade_dir == "LONG":
            gross_pnl = (exit_price - entry_price) * lot_size
        else:
            gross_pnl = (entry_price - exit_price) * lot_size
            
        fric_cost = pos_val * 0.0005
        net_pnl = gross_pnl - fric_cost
        ret_pct = (net_pnl / margin_req) * 100.0 if margin_req > 0 else 0.0
        win_flag = "WIN" if net_pnl > 0 else "LOSS"
        
        quarter_trades[q_label].append({
            "symbol": sym,
            "sector": "N/A",
            "status": stock_status,
            "raw_strat": raw_strat,
            "window_str": window_str,
            "hist_q_str": stock_info['hist_q_str'],
            "executed": True,
            "result_date": b_dt.strftime('%Y-%m-%d'),
            "entry_date": entry_date.strftime('%Y-%m-%d'),
            "exit_date": exit_date.strftime('%Y-%m-%d'),
            "entry_price": entry_price,
            "exit_price": exit_price,
            "lot_size": lot_size,
            "pos_val": pos_val,
            "margin_req": margin_req,
            "net_pnl": net_pnl,
            "ret_pct": ret_pct,
            "win_flag": win_flag
        })

print("\nVERIFICATION OF EXECUTED TRADES VS N/A PER QUARTER:")
print(f"{'Quarter':15s} | {'Total Stocks':12s} | {'Executed Trades':16s} | {'N/A Stocks':12s} | {'Qualified Exec':15s} | {'Net P&L (₹)':15s} | {'Win Rate %':10s}")
print("-" * 100)

summary_stats = []

for q_info in TARGET_QUARTERS:
    q_label = q_info['label']
    trades = quarter_trades[q_label]
    
    total_stocks = len(trades)
    executed_trades = [t for t in trades if t['executed']]
    exec_count = len(executed_trades)
    na_count = total_stocks - exec_count
    
    qual_exec = len([t for t in executed_trades if t['status'] == 'QUALIFIED'])
    
    if exec_count > 0:
        net_pnl_sum = sum(t['net_pnl'] for t in executed_trades)
        wins = sum(1 for t in executed_trades if t['win_flag'] == 'WIN')
        win_rate = (wins / exec_count) * 100.0
    else:
        net_pnl_sum = 0.0
        win_rate = 0.0
        
    summary_stats.append({
        "q_label": q_label,
        "total_stocks": total_stocks,
        "exec_count": exec_count,
        "na_count": na_count,
        "qual_exec": qual_exec,
        "net_pnl": net_pnl_sum,
        "win_rate": win_rate
    })
    
    print(f"{q_label:15s} | {total_stocks:12d} | {exec_count:16d} | {na_count:12d} | {qual_exec:15d} | {net_pnl_sum:15,.2f} | {win_rate:10.2f}%")

print("="*100)

# Build Excel Workbook
wb = openpyxl.Workbook()
wb.remove(wb.active)  # remove default sheet

# Styling definitions
header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

title_fill = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
title_font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")

green_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
red_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
na_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

thin_border = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)

# 1. MASTER SUMMARY SHEET
ws_sum = wb.create_sheet(title="MASTER SUMMARY")
ws_sum.views.sheetView[0].showGridLines = True

ws_sum.merge_cells("A1:G1")
ws_sum["A1"] = "NIFTY 211 - 12 QUARTERS FUTURES PERFORMANCE MASTER (HISTORY-AWARE)"
ws_sum["A1"].font = title_font
ws_sum["A1"].fill = title_fill
ws_sum["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws_sum.row_dimensions[1].height = 35

headers_sum = [
    "Quarter Name", "Total Universe Stocks", "Executed Trades", "N/A Stocks (No Data)",
    "QUALIFIED Stocks Executed", "Net Realised P&L (₹)", "Win Rate %"
]
ws_sum.append([])
ws_sum.append(headers_sum)
ws_sum.row_dimensions[3].height = 25

for col_num, h in enumerate(headers_sum, 1):
    cell = ws_sum.cell(row=3, column=col_num)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

for r_idx, s in enumerate(summary_stats, start=4):
    ws_sum.append([
        s["q_label"], s["total_stocks"], s["exec_count"], s["na_count"],
        s["qual_exec"], s["net_pnl"], s["win_rate"] / 100.0
    ])
    ws_sum.row_dimensions[r_idx].height = 20
    
    # formats
    ws_sum.cell(row=r_idx, column=1).alignment = Alignment(horizontal="left")
    ws_sum.cell(row=r_idx, column=2).number_format = '#,##0'
    ws_sum.cell(row=r_idx, column=3).number_format = '#,##0'
    ws_sum.cell(row=r_idx, column=4).number_format = '#,##0'
    ws_sum.cell(row=r_idx, column=5).number_format = '#,##0'
    ws_sum.cell(row=r_idx, column=6).number_format = '₹#,##0.00'
    ws_sum.cell(row=r_idx, column=7).number_format = '0.00%'

# Total Row in Summary
tot_row = len(summary_stats) + 4
ws_sum.append([
    "TOTAL (12 QUARTERS)",
    211,
    sum(s["exec_count"] for s in summary_stats),
    sum(s["na_count"] for s in summary_stats),
    sum(s["qual_exec"] for s in summary_stats),
    sum(s["net_pnl"] for s in summary_stats),
    sum(s["net_pnl"] for s in summary_stats) / max(1, sum(s["exec_count"] for s in summary_stats))
])
ws_sum.row_dimensions[tot_row].height = 24
tot_font = Font(name="Calibri", size=11, bold=True)
tot_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

for c_idx in range(1, 8):
    cell = ws_sum.cell(row=tot_row, column=c_idx)
    cell.font = tot_font
    cell.fill = tot_fill
    if c_idx == 6: cell.number_format = '₹#,##0.00'
    elif c_idx == 7: cell.number_format = '0.00%'
    elif c_idx > 1: cell.number_format = '#,##0'

# Auto-adjust summary widths
for col in ws_sum.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws_sum.column_dimensions[col_letter].width = max(max_len + 3, 16)

# 2. INDIVIDUAL QUARTER SHEETS
headers_q = [
    "Trade #", "Stock Symbol", "Industry Sector", "Action Status", "Strategy",
    "Position Window", "Total History Quarters", "Result Announcement Date", "Entry Date",
    "Exit Date", "Entry Price (₹)", "Exit Price (₹)", "Position Value (₹)",
    "Margin Capital (20%) (₹)", "Net Realised P&L (₹)", "Return %", "Win / Loss"
]

for q_info in TARGET_QUARTERS:
    q_code = q_info['code']
    q_label = q_info['label']
    trades = quarter_trades[q_label]
    
    ws = wb.create_sheet(title=q_code)
    ws.views.sheetView[0].showGridLines = True
    
    # Title Block
    ws.merge_cells("A1:Q1")
    ws["A1"] = f"{q_label.upper()} - DETAILED FUTURES TRADE LOG (HISTORY-AWARE)"
    ws["A1"].font = title_font
    ws["A1"].fill = title_fill
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 35
    
    ws.append([])
    ws.append(headers_q)
    ws.row_dimensions[3].height = 25
    
    for col_num, h in enumerate(headers_q, 1):
        cell = ws.cell(row=3, column=col_num)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
    for idx, t in enumerate(trades, start=1):
        r_idx = idx + 3
        if t['executed']:
            row_data = [
                idx, t['symbol'], t['sector'], t['status'], t['raw_strat'],
                t['window_str'], t['hist_q_str'], t['result_date'], t['entry_date'],
                t['exit_date'], t['entry_price'], t['exit_price'], t['pos_val'],
                t['margin_req'], t['net_pnl'], t['ret_pct'] / 100.0, t['win_flag']
            ]
        else:
            row_data = [
                idx, t['symbol'], t['sector'], t['status'], t['raw_strat'],
                t['window_str'], t['hist_q_str'], "N/A", "N/A",
                "N/A", "N/A", "N/A", "N/A",
                "N/A", "N/A", "N/A", "N/A"
            ]
            
        ws.append(row_data)
        ws.row_dimensions[r_idx].height = 20
        
        # Formatting
        ws.cell(row=r_idx, column=1).alignment = Alignment(horizontal="center")
        ws.cell(row=r_idx, column=2).alignment = Alignment(horizontal="left")
        ws.cell(row=r_idx, column=2).font = Font(bold=True)
        ws.cell(row=r_idx, column=4).alignment = Alignment(horizontal="center")
        ws.cell(row=r_idx, column=5).alignment = Alignment(horizontal="center")
        ws.cell(row=r_idx, column=6).alignment = Alignment(horizontal="center")
        ws.cell(row=r_idx, column=7).alignment = Alignment(horizontal="center")
        
        for c in range(8, 11):
            ws.cell(row=r_idx, column=c).alignment = Alignment(horizontal="center")
            
        if t['executed']:
            ws.cell(row=r_idx, column=11).number_format = '₹#,##0.00'
            ws.cell(row=r_idx, column=12).number_format = '₹#,##0.00'
            ws.cell(row=r_idx, column=13).number_format = '₹#,##0.00'
            ws.cell(row=r_idx, column=14).number_format = '₹#,##0.00'
            ws.cell(row=r_idx, column=15).number_format = '₹#,##0.00'
            ws.cell(row=r_idx, column=16).number_format = '0.00%'
            ws.cell(row=r_idx, column=17).alignment = Alignment(horizontal="center")
            
            # Row coloring based on P&L
            if t['net_pnl'] > 0:
                ws.cell(row=r_idx, column=15).fill = green_fill
                ws.cell(row=r_idx, column=17).fill = green_fill
            else:
                ws.cell(row=r_idx, column=15).fill = red_fill
                ws.cell(row=r_idx, column=17).fill = red_fill
        else:
            for c in range(8, 18):
                cell = ws.cell(row=r_idx, column=c)
                cell.alignment = Alignment(horizontal="center")
                cell.fill = na_fill

    # Auto-adjust column widths
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 14)

print(f"\nSaving Excel Workbook to: {OUTPUT_FILE} ...")
wb.save(OUTPUT_FILE)
print("SUCCESSFULLY SAVED HISTORY-AWARE 12Q FUTURES MASTER EXCEL!")
