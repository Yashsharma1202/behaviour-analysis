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
OUTPUT_FILE = ROOT / 'Nifty50_Only_12Q_and_8Q_Futures_Master.xlsx'

NIFTY50_SYMBOLS = [
    'RELIANCE', 'TCS', 'INFY', 'HDFCBANK', 'ICICIBANK', 'BHARTIARTL', 'ITC', 'SBIN', 'LTM', 'LT',
    'HINDUNILVR', 'AXISBANK', 'KOTAKBANK', 'BAJFINANCE', 'M&M', 'MARUTI', 'SUNPHARMA', 'TATASTEEL',
    'NTPC', 'POWERGRID', 'TITAN', 'ADANIENT', 'ADANIPORTS', 'ULTRACEMCO', 'ASIANPAINT', 'COALINDIA',
    'BAJAJ-AUTO', 'JSWSTEEL', 'TATAMOTORS', 'HCLTECH', 'GRASIM', 'HEROMOTOCO', 'EICHERMOT', 'CIPLA',
    'HDFCLIFE', 'SBILIFE', 'DRREDDY', 'BRITANNIA', 'APOLLOHOSP', 'TATACONSUM', 'HINDALCO', 'BPCL',
    'INDUSINDBK', 'DIVISLAB', 'BAJAJFINSV', 'NESTLEIND', 'WIPRO', 'ONGC', 'TECHM', 'SHRIRAMFIN'
]

def clean_pct(val):
    if pd.isna(val): return 0.0
    val_str = str(val).replace('%', '').replace('₹', '').strip()
    try:
        v = float(val_str)
        return v / 100.0 if v > 1.0 else v
    except Exception:
        return 0.0

print("="*90)
print(f"Loading NIFTY 50 stock strategies & predicted benchmarks from: {PORTFOLIO_MASTER.name} ...")
df_pm = pd.read_excel(PORTFOLIO_MASTER, sheet_name='ALL QUARTERS COMBINED', skiprows=6)
df_pm['Stock Symbol'] = df_pm['Stock Symbol'].astype(str).str.strip().str.upper()

# Filter ONLY Nifty 50 stocks
df_pm_nifty50 = df_pm[df_pm['Stock Symbol'].isin(NIFTY50_SYMBOLS)].copy()

STOCK_MAP = {}
for _, row in df_pm_nifty50.iterrows():
    sym = row['Stock Symbol']
    strat = str(row['Strategy']).strip().upper()
    window_raw = str(row['Position Window']).strip()
    status = str(row['Action Status']).strip()
    hist_q_str = str(row['Total History Quarters']).strip()
    sector = str(row['Industry Sector']).strip() if 'Industry Sector' in row and pd.notna(row['Industry Sector']) else "N/A"
    
    pred_win_rate = clean_pct(row.get('Win Rate %', 0))
    pred_avg_ret = clean_pct(row.get('Average Return %', 0))
    
    import re
    match = re.search(r'(\d+)', hist_q_str)
    hist_q_num = int(match.group(1)) if match else 70
    
    try:
        parts = window_raw.replace('T-', '').replace('T+', '').split(' to ')
        pre_days = int(parts[0])
        post_days = int(parts[1])
    except Exception:
        pre_days, post_days = 5, 5
        
    STOCK_MAP[sym] = {
        "sector": sector,
        "strategy": "LONG" if "LONG" in strat else "SHORT",
        "raw_strategy": strat,
        "window_str": window_raw,
        "pre_days": pre_days,
        "post_days": post_days,
        "status": status,
        "hist_q_str": hist_q_str,
        "hist_q_num": hist_q_num,
        "pred_win_rate": pred_win_rate,
        "pred_avg_ret": pred_avg_ret
    }

SYMBOLS = sorted(list(STOCK_MAP.keys()))
print(f"Total Nifty 50 stocks filtered: {len(SYMBOLS)}")

LOT_SIZES = {
    "RELIANCE": 250, "TCS": 175, "INFY": 400, "HDFCBANK": 550, "ICICIBANK": 700,
    "BHARTIARTL": 950, "ITC": 1600, "SBIN": 1500, "LTM": 150, "LT": 300,
    "HINDUNILVR": 300, "AXISBANK": 625, "KOTAKBANK": 400, "BAJFINANCE": 125,
    "M&M": 350, "MARUTI": 100, "SUNPHARMA": 350, "TATASTEEL": 5500,
    "NTPC": 1500, "POWERGRID": 1800, "TITAN": 175, "ADANIENT": 300,
    "ADANIPORTS": 625, "ULTRACEMCO": 100, "ASIANPAINT": 200, "COALINDIA": 2100,
    "BAJAJ-AUTO": 125, "JSWSTEEL": 675, "TATAMOTORS": 1425, "HCLTECH": 350,
    "GRASIM": 250, "HEROMOTOCO": 150, "EICHERMOT": 175, "CIPLA": 650,
    "HDFCLIFE": 1100, "SBILIFE": 375, "DRREDDY": 125, "BRITANNIA": 200,
    "APOLLOHOSP": 125, "TATACONSUM": 450, "HINDALCO": 1400, "BPCL": 1800,
    "INDUSINDBK": 500, "DIVISLAB": 200, "BAJAJFINSV": 500, "NESTLEIND": 200,
    "WIPRO": 1500, "ONGC": 3750, "TECHM": 600, "SHRIRAMFIN": 300
}

TARGET_QUARTERS = [
    {"q_idx": 1, "code": "Q3_FY23-24", "label": "Q3 FY2023-24", "q_type": "Q3", "target_year": 2024, "default_dt": "2024-01-22"},
    {"q_idx": 2, "code": "Q4_FY23-24", "label": "Q4 FY2023-24", "q_type": "Q4", "target_year": 2024, "default_dt": "2024-04-22"},
    {"q_idx": 3, "code": "Q1_FY24-25", "label": "Q1 FY2024-25", "q_type": "Q1", "target_year": 2024, "default_dt": "2024-07-22"},
    {"q_idx": 4, "code": "Q2_FY24-25", "label": "Q2 FY2024-25", "q_type": "Q2", "target_year": 2024, "default_dt": "2024-10-22"},
    {"q_idx": 5, "code": "Q3_FY24-25", "label": "Q3 FY2024-25", "q_type": "Q3", "target_year": 2025, "default_dt": "2025-01-22"},
    {"q_idx": 6, "code": "Q4_FY24-25", "label": "Q4 FY2024-25", "q_type": "Q4", "target_year": 2025, "default_dt": "2025-04-22"},
    {"q_idx": 7, "code": "Q1_FY25-26", "label": "Q1 FY2025-26", "q_type": "Q1", "target_year": 2025, "default_dt": "2025-07-22"},
    {"q_idx": 8, "code": "Q2_FY25-26", "label": "Q2 FY2025-26", "q_type": "Q2", "target_year": 2025, "default_dt": "2025-10-22"},
    {"q_idx": 9, "code": "Q3_FY25-26", "label": "Q3 FY2025-26", "q_type": "Q3", "target_year": 2026, "default_dt": "2026-01-22"},
    {"q_idx": 10, "code": "Q4_FY25-26", "label": "Q4 FY2025-26", "q_type": "Q4", "target_year": 2026, "default_dt": "2026-04-22"},
    {"q_idx": 11, "code": "Q1_FY26-27", "label": "Q1 FY2026-27", "q_type": "Q1", "target_year": 2026, "default_dt": "2026-07-22"},
    {"q_idx": 12, "code": "Q2_FY26-27", "label": "Q2 FY2026-27", "q_type": "Q2", "target_year": 2026, "default_dt": "2026-10-22"},
]

quarter_trades = {q['label']: [] for q in TARGET_QUARTERS}
stock_aggregated_trades_12q = {sym: [] for sym in SYMBOLS}
stock_aggregated_trades_8q = {sym: [] for sym in SYMBOLS}

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
        for q_info in TARGET_QUARTERS:
            quarter_trades[q_info['label']].append({
                "symbol": sym, "sector": stock_info['sector'], "status": stock_info['status'],
                "raw_strat": stock_info['raw_strategy'], "window_str": stock_info['window_str'],
                "hist_q_str": stock_info['hist_q_str'], "pred_win_rate": stock_info['pred_win_rate'],
                "pred_avg_ret": stock_info['pred_avg_ret'], "executed": False
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
    sector = stock_info['sector']
    pred_win_rate = stock_info['pred_win_rate']
    pred_avg_ret = stock_info['pred_avg_ret']
    
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
        
        if q_idx > hist_q_num:
            t_record = {
                "symbol": sym, "sector": sector, "status": stock_status,
                "raw_strat": raw_strat, "window_str": window_str,
                "hist_q_str": stock_info['hist_q_str'], "pred_win_rate": pred_win_rate,
                "pred_avg_ret": pred_avg_ret, "executed": False
            }
            quarter_trades[q_label].append(t_record)
            continue

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
        act_ret_pct = (net_pnl / margin_req) * 100.0 if margin_req > 0 else 0.0
        win_flag = "WIN" if net_pnl > 0 else "LOSS"
        
        t_record = {
            "symbol": sym, "sector": sector, "status": stock_status,
            "raw_strat": raw_strat, "window_str": window_str,
            "hist_q_str": stock_info['hist_q_str'], "pred_win_rate": pred_win_rate,
            "pred_avg_ret": pred_avg_ret, "executed": True,
            "result_date": b_dt.strftime('%Y-%m-%d'), "entry_date": entry_date.strftime('%Y-%m-%d'),
            "exit_date": exit_date.strftime('%Y-%m-%d'), "entry_price": entry_price,
            "exit_price": exit_price, "lot_size": lot_size, "pos_val": pos_val,
            "margin_req": margin_req, "net_pnl": net_pnl, "act_ret_pct": act_ret_pct, "win_flag": win_flag
        }
        
        quarter_trades[q_label].append(t_record)
        stock_aggregated_trades_12q[sym].append(t_record)
        if q_idx <= 8:
            stock_aggregated_trades_8q[sym].append(t_record)

# Function to compute rankings DataFrame
def build_rankings(aggregated_trades_map):
    rank_list = []
    for sym in SYMBOLS:
        trades = aggregated_trades_map[sym]
        executed = [t for t in trades if t.get('executed', False)]
        tot_trades = len(executed)
        
        if tot_trades > 0:
            wins = sum(1 for t in executed if t['win_flag'] == 'WIN')
            losses = tot_trades - wins
            win_rate = (wins / tot_trades) * 100.0
            total_pnl = sum(t['net_pnl'] for t in executed)
            avg_act_ret = np.mean([t['act_ret_pct'] for t in executed])
            avg_margin = np.mean([t['margin_req'] for t in executed])
        else:
            wins, losses, win_rate, total_pnl, avg_act_ret, avg_margin = 0, 0, 0.0, 0.0, 0.0, 0.0
            
        s_info = STOCK_MAP[sym]
        rank_list.append({
            "symbol": sym,
            "sector": s_info['sector'],
            "status": s_info['status'],
            "strategy": s_info['raw_strategy'],
            "window_str": s_info['window_str'],
            "hist_q_str": s_info['hist_q_str'],
            "pred_win_rate": s_info['pred_win_rate'],
            "pred_avg_ret": s_info['pred_avg_ret'],
            "tot_trades": tot_trades,
            "wins": wins,
            "losses": losses,
            "win_rate": win_rate,
            "avg_act_ret": avg_act_ret,
            "avg_margin": avg_margin,
            "total_pnl": total_pnl
        })
    return sorted(rank_list, key=lambda x: (x['win_rate'], x['total_pnl']), reverse=True)

rank_list_12q = build_rankings(stock_aggregated_trades_12q)
rank_list_8q = build_rankings(stock_aggregated_trades_8q)

SYM_RANK_MAP_12Q = {item['symbol']: r_idx for r_idx, item in enumerate(rank_list_12q, start=1)}
for r_idx, item in enumerate(rank_list_12q, start=1): item['rank'] = r_idx

SYM_RANK_MAP_8Q = {item['symbol']: r_idx for r_idx, item in enumerate(rank_list_8q, start=1)}
for r_idx, item in enumerate(rank_list_8q, start=1): item['rank'] = r_idx

# Qualified Stocks Filtering
qual_list_12q = [dict(item) for item in rank_list_12q if item['status'] == 'QUALIFIED']
for r_idx, item in enumerate(qual_list_12q, start=1): item['qual_rank'] = r_idx

print("\nNIFTY 50 TOP 15 STOCKS BY WIN RATE & P&L (12 QUARTERS):")
print(f"{'Rank':5s} | {'Symbol':12s} | {'Status':20s} | {'Strategy':14s} | {'Trades':7s} | {'Wins':5s} | {'Win Rate %':10s} | {'Total P&L (₹)':16s}")
print("-" * 105)
for item in rank_list_12q[:15]:
    print(f"{item['rank']:5d} | {item['symbol']:12s} | {item['status']:20s} | {item['strategy']:14s} | {item['tot_trades']:7d} | {item['wins']:5d} | {item['win_rate']:10.2f}% | {item['total_pnl']:16,.2f}")

# Create Excel Workbook
wb = openpyxl.Workbook()
wb.remove(wb.active)

# Styles
header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

title_fill = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
title_font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")

kpi_header_fill = PatternFill(start_color="203764", end_color="203764", fill_type="solid")

green_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
red_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
na_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

# Helper to write summary sheet
def write_summary_sheet(sheet_title, title_text, rank_data_list, total_q_num):
    ws = wb.create_sheet(title=sheet_title)
    ws.views.sheetView[0].showGridLines = True
    
    ws.merge_cells("A1:P1")
    ws["A1"] = title_text
    ws["A1"].font = title_font
    ws["A1"].fill = title_fill
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 35
    
    tot_pnl = sum(item['total_pnl'] for item in rank_data_list)
    tot_trades = sum(item['tot_trades'] for item in rank_data_list)
    tot_wins = sum(item['wins'] for item in rank_data_list)
    overall_wr = (tot_wins / max(1, tot_trades)) * 100.0
    
    qual_items = [item for item in rank_data_list if item['status'] == 'QUALIFIED']
    qual_pnl = sum(item['total_pnl'] for item in qual_items)
    qual_trades = sum(item['tot_trades'] for item in qual_items)
    qual_wins = sum(item['wins'] for item in qual_items)
    qual_wr = (qual_wins / max(1, qual_trades)) * 100.0
    
    top_stk = rank_data_list[0]['symbol']
    top_pnl = rank_data_list[0]['total_pnl']
    top_wr = rank_data_list[0]['win_rate']
    
    ws.merge_cells("A3:P3")
    ws["A3"] = f"NIFTY 50 EXECUTIVE PORTFOLIO SUMMARY ({total_q_num} FINANCIAL QUARTERS)"
    ws["A3"].font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    ws["A3"].fill = kpi_header_fill
    ws["A3"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[3].height = 24
    
    ws.cell(row=4, column=1, value="Total Portfolio Net P&L (₹):").font = Font(bold=True)
    ws.cell(row=4, column=2, value=tot_pnl).number_format = '₹#,##0.00'
    ws.cell(row=4, column=2).font = Font(bold=True, color="006100")
    
    ws.cell(row=4, column=4, value="Overall Portfolio Win Rate %:").font = Font(bold=True)
    ws.cell(row=4, column=5, value=overall_wr / 100.0).number_format = '0.00%'
    ws.cell(row=4, column=5).font = Font(bold=True)
    
    ws.cell(row=4, column=7, value="Total Executed Trades:").font = Font(bold=True)
    ws.cell(row=4, column=8, value=tot_trades).number_format = '#,##0'
    
    ws.cell(row=5, column=1, value="QUALIFIED Stocks Net P&L (₹):").font = Font(bold=True)
    ws.cell(row=5, column=2, value=qual_pnl).number_format = '₹#,##0.00'
    ws.cell(row=5, column=2).font = Font(bold=True, color="006100")
    
    ws.cell(row=5, column=4, value="QUALIFIED Stocks Win Rate %:").font = Font(bold=True)
    ws.cell(row=5, column=5, value=qual_wr / 100.0).number_format = '0.00%'
    ws.cell(row=5, column=5).font = Font(bold=True)
    
    ws.cell(row=5, column=7, value="Top Performing Stock:").font = Font(bold=True)
    ws.cell(row=5, column=8, value=f"{top_stk} ({top_wr:.1f}% WR)").font = Font(bold=True, color="002060")
    
    headers = [
        "Rank", "Stock Symbol", "Industry Sector", "Action Status", "Strategy",
        "Position Window", "Total History Quarters", "Pred Win Rate %", "Pred Avg Return %",
        "Total Trades Executed", "Winning Trades", "Losing Trades", "Actual Win Rate %",
        "Actual Avg Return %", "Margin Capital (20%) (₹)", "Total Net Realised P&L (₹)"
    ]
    
    ws.append([])
    ws.append(headers)
    ws.row_dimensions[7].height = 25
    
    for col_num, h in enumerate(headers, 1):
        cell = ws.cell(row=7, column=col_num)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
    for idx, item in enumerate(rank_data_list, start=1):
        r_idx = idx + 7
        ws.append([
            idx, item['symbol'], item['sector'], item['status'], item['strategy'],
            item['window_str'], item['hist_q_str'], item['pred_win_rate'], item['pred_avg_ret'],
            item['tot_trades'], item['wins'], item['losses'], item['win_rate'] / 100.0,
            item['avg_act_ret'] / 100.0, item['avg_margin'], item['total_pnl']
        ])
        ws.row_dimensions[r_idx].height = 20
        
        ws.cell(row=r_idx, column=1).alignment = Alignment(horizontal="center")
        ws.cell(row=r_idx, column=2).alignment = Alignment(horizontal="left")
        ws.cell(row=r_idx, column=2).font = Font(bold=True)
        ws.cell(row=r_idx, column=4).alignment = Alignment(horizontal="center")
        ws.cell(row=r_idx, column=5).alignment = Alignment(horizontal="center")
        ws.cell(row=r_idx, column=6).alignment = Alignment(horizontal="center")
        ws.cell(row=r_idx, column=7).alignment = Alignment(horizontal="center")
        
        ws.cell(row=r_idx, column=8).number_format = '0.00%'
        ws.cell(row=r_idx, column=9).number_format = '0.00%'
        ws.cell(row=r_idx, column=10).number_format = '#,##0'
        ws.cell(row=r_idx, column=11).number_format = '#,##0'
        ws.cell(row=r_idx, column=12).number_format = '#,##0'
        ws.cell(row=r_idx, column=13).number_format = '0.00%'
        ws.cell(row=r_idx, column=14).number_format = '0.00%'
        ws.cell(row=r_idx, column=15).number_format = '₹#,##0.00'
        ws.cell(row=r_idx, column=16).number_format = '₹#,##0.00'
        
        if item['total_pnl'] > 0:
            ws.cell(row=r_idx, column=16).fill = green_fill
        elif item['total_pnl'] < 0:
            ws.cell(row=r_idx, column=16).fill = red_fill
            
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 15)

# 1. Write NIFTY 50 12Q Summary
write_summary_sheet("NIFTY 50 RANKINGS (12Q)", "NIFTY 50 STOCKS PERFORMANCE SUMMARY & WIN RATE RANKINGS (12 QUARTERS)", rank_list_12q, 12)

# 2. Write NIFTY 50 8Q Summary
write_summary_sheet("NIFTY 50 RANKINGS (8Q)", "NIFTY 50 STOCKS PERFORMANCE SUMMARY & WIN RATE RANKINGS (LAST 8 QUARTERS)", rank_list_8q, 8)

# 3. Write NIFTY 50 QUALIFIED RANKINGS Sheet
ws_q = wb.create_sheet(title="NIFTY 50 QUALIFIED RANKINGS")
ws_q.views.sheetView[0].showGridLines = True

ws_q.merge_cells("A1:P1")
ws_q["A1"] = "NIFTY 50 - QUALIFIED STOCKS PERFORMANCE SUMMARY & RANKINGS"
ws_q["A1"].font = title_font
ws_q["A1"].fill = PatternFill(start_color="004B50", end_color="004B50", fill_type="solid")
ws_q["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws_q.row_dimensions[1].height = 35

qual_tot_pnl = sum(item['total_pnl'] for item in qual_list_12q)
qual_tot_trades = sum(item['tot_trades'] for item in qual_list_12q)
qual_tot_wins = sum(item['wins'] for item in qual_list_12q)
qual_tot_wr = (qual_tot_wins / max(1, qual_tot_trades)) * 100.0

ws_q.merge_cells("A3:P3")
ws_q["A3"] = "NIFTY 50 QUALIFIED STOCKS PERFORMANCE HIGHLIGHTS (12 QUARTERS)"
ws_q["A3"].font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
ws_q["A3"].fill = PatternFill(start_color="006666", end_color="006666", fill_type="solid")
ws_q["A3"].alignment = Alignment(horizontal="center", vertical="center")
ws_q.row_dimensions[3].height = 24

ws_q.cell(row=4, column=1, value="NIFTY 50 QUALIFIED Stocks Count:").font = Font(bold=True)
ws_q.cell(row=4, column=2, value=len(qual_list_12q)).number_format = '#,##0'

ws_q.cell(row=4, column=4, value="Total Executed Trades (12Q):").font = Font(bold=True)
ws_q.cell(row=4, column=5, value=qual_tot_trades).number_format = '#,##0'

ws_q.cell(row=4, column=7, value="Total Winning Trades:").font = Font(bold=True)
ws_q.cell(row=4, column=8, value=qual_tot_wins).number_format = '#,##0'

ws_q.cell(row=4, column=10, value="QUALIFIED Win Rate %:").font = Font(bold=True)
ws_q.cell(row=4, column=11, value=qual_tot_wr / 100.0).number_format = '0.00%'
ws_q.cell(row=4, column=11).font = Font(bold=True)

ws_q.cell(row=5, column=1, value="QUALIFIED Net P&L (₹):").font = Font(bold=True)
ws_q.cell(row=5, column=2, value=qual_tot_pnl).number_format = '₹#,##0.00'
ws_q.cell(row=5, column=2).font = Font(bold=True, color="006100")

headers_qual_rank = [
    "Qual Rank", "Overall Rank", "Stock Symbol", "Industry Sector", "Strategy",
    "Position Window", "Total History Quarters", "Pred Win Rate %", "Pred Avg Return %",
    "Total Trades Executed", "Winning Trades", "Losing Trades", "Actual Win Rate %",
    "Actual Avg Return %", "Margin Capital (20%) (₹)", "Net Realised P&L (₹)"
]

ws_q.append([])
ws_q.append(headers_qual_rank)
ws_q.row_dimensions[7].height = 25

for col_num, h in enumerate(headers_qual_rank, 1):
    cell = ws_q.cell(row=7, column=col_num)
    cell.font = header_font
    cell.fill = PatternFill(start_color="004B50", end_color="004B50", fill_type="solid")
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

for idx, item in enumerate(qual_list_12q, start=1):
    r_idx = idx + 7
    ws_q.append([
        item['qual_rank'], item['rank'], item['symbol'], item['sector'], item['strategy'],
        item['window_str'], item['hist_q_str'], item['pred_win_rate'], item['pred_avg_ret'],
        item['tot_trades'], item['wins'], item['losses'], item['win_rate'] / 100.0,
        item['avg_act_ret'] / 100.0, item['avg_margin'], item['total_pnl']
    ])
    ws_q.row_dimensions[r_idx].height = 20
    
    ws_q.cell(row=r_idx, column=1).alignment = Alignment(horizontal="center")
    ws_q.cell(row=r_idx, column=1).font = Font(bold=True)
    ws_q.cell(row=r_idx, column=2).alignment = Alignment(horizontal="center")
    ws_q.cell(row=r_idx, column=3).alignment = Alignment(horizontal="left")
    ws_q.cell(row=r_idx, column=3).font = Font(bold=True)
    ws_q.cell(row=r_idx, column=5).alignment = Alignment(horizontal="center")
    ws_q.cell(row=r_idx, column=6).alignment = Alignment(horizontal="center")
    ws_q.cell(row=r_idx, column=7).alignment = Alignment(horizontal="center")
    
    ws_q.cell(row=r_idx, column=8).number_format = '0.00%'
    ws_q.cell(row=r_idx, column=9).number_format = '0.00%'
    ws_q.cell(row=r_idx, column=10).number_format = '#,##0'
    ws_q.cell(row=r_idx, column=11).number_format = '#,##0'
    ws_q.cell(row=r_idx, column=12).number_format = '#,##0'
    ws_q.cell(row=r_idx, column=13).number_format = '0.00%'
    ws_q.cell(row=r_idx, column=14).number_format = '0.00%'
    ws_q.cell(row=r_idx, column=15).number_format = '₹#,##0.00'
    ws_q.cell(row=r_idx, column=16).number_format = '₹#,##0.00'
    
    if item['total_pnl'] > 0:
        ws_q.cell(row=r_idx, column=16).fill = green_fill
    elif item['total_pnl'] < 0:
        ws_q.cell(row=r_idx, column=16).fill = red_fill

for col in ws_q.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws_q.column_dimensions[col_letter].width = max(max_len + 3, 15)

# 4. INDIVIDUAL 12 QUARTER SHEETS (SORTED BY NIFTY 50 RANK #1 TO #50)
headers_q = [
    "Rank", "Stock Symbol", "Industry Sector", "Action Status", "Strategy",
    "Position Window", "Total History Quarters", "Pred Win Rate %", "Pred Avg Return %",
    "Result Announcement Date", "Entry Date", "Exit Date", "Entry Price (₹)", "Exit Price (₹)",
    "Position Value (₹)", "Margin Capital (20%) (₹)", "Net Realised P&L (₹)",
    "Actual Realised Return %", "Win / Loss"
]

for q_info in TARGET_QUARTERS:
    q_code = q_info['code']
    q_label = q_info['label']
    trades = quarter_trades[q_label]
    
    trades_sorted = sorted(trades, key=lambda x: SYM_RANK_MAP_12Q.get(x['symbol'], 999))
    
    ws = wb.create_sheet(title=q_code)
    ws.views.sheetView[0].showGridLines = True
    
    ws.merge_cells("A1:S1")
    ws["A1"] = f"{q_label.upper()} - NIFTY 50 FUTURES TRADE LOG (RANKED #1 TO #50)"
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
        
    for idx, t in enumerate(trades_sorted, start=1):
        r_idx = idx + 3
        rank_num = SYM_RANK_MAP_12Q.get(t['symbol'], idx)
        
        if t['executed']:
            row_data = [
                rank_num, t['symbol'], t['sector'], t['status'], t['raw_strat'],
                t['window_str'], t['hist_q_str'], t['pred_win_rate'], t['pred_avg_ret'],
                t['result_date'], t['entry_date'], t['exit_date'], t['entry_price'],
                t['exit_price'], t['pos_val'], t['margin_req'], t['net_pnl'],
                t['act_ret_pct'] / 100.0, t['win_flag']
            ]
        else:
            row_data = [
                rank_num, t['symbol'], t['sector'], t['status'], t['raw_strat'],
                t['window_str'], t['hist_q_str'], t['pred_win_rate'], t['pred_avg_ret'],
                "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A"
            ]
            
        ws.append(row_data)
        ws.row_dimensions[r_idx].height = 20
        
        ws.cell(row=r_idx, column=1).alignment = Alignment(horizontal="center")
        ws.cell(row=r_idx, column=2).alignment = Alignment(horizontal="left")
        ws.cell(row=r_idx, column=2).font = Font(bold=True)
        ws.cell(row=r_idx, column=4).alignment = Alignment(horizontal="center")
        ws.cell(row=r_idx, column=5).alignment = Alignment(horizontal="center")
        ws.cell(row=r_idx, column=6).alignment = Alignment(horizontal="center")
        ws.cell(row=r_idx, column=7).alignment = Alignment(horizontal="center")
        
        ws.cell(row=r_idx, column=8).number_format = '0.00%'
        ws.cell(row=r_idx, column=9).number_format = '0.00%'
        
        for c in range(10, 13):
            ws.cell(row=r_idx, column=c).alignment = Alignment(horizontal="center")
            
        if t['executed']:
            ws.cell(row=r_idx, column=13).number_format = '₹#,##0.00'
            ws.cell(row=r_idx, column=14).number_format = '₹#,##0.00'
            ws.cell(row=r_idx, column=15).number_format = '₹#,##0.00'
            ws.cell(row=r_idx, column=16).number_format = '₹#,##0.00'
            ws.cell(row=r_idx, column=17).number_format = '₹#,##0.00'
            ws.cell(row=r_idx, column=18).number_format = '0.00%'
            ws.cell(row=r_idx, column=19).alignment = Alignment(horizontal="center")
            
            if t['net_pnl'] > 0:
                ws.cell(row=r_idx, column=17).fill = green_fill
                ws.cell(row=r_idx, column=19).fill = green_fill
            else:
                ws.cell(row=r_idx, column=17).fill = red_fill
                ws.cell(row=r_idx, column=19).fill = red_fill
        else:
            for c in range(10, 20):
                cell = ws.cell(row=r_idx, column=c)
                cell.alignment = Alignment(horizontal="center")
                cell.fill = na_fill

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 14)

print(f"\nSaving NIFTY 50 ONLY Master Workbook to: {OUTPUT_FILE} ...")
wb.save(OUTPUT_FILE)
print("SUCCESSFULLY SAVED NIFTY 50 ONLY FUTURES MASTER EXCEL!")
