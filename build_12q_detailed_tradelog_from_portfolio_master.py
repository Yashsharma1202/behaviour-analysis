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
OUTPUT_FILE = ROOT / 'Nifty211_12Q_Detailed_Trade_Log_Master.xlsx'

# 1. Load Portfolio Master Strategy & Window Mapping
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
    
    # Parse window pre/post days
    # Format: T-x to T+y
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

TARGET_QUARTERS = [
    {"code": "Q3_FY23-24", "label": "Q3 FY2023-24", "q_type": "Q3", "period": "Oct 2023 – Dec 2023", "ann": "Jan 2024 – Feb 2024"},
    {"code": "Q4_FY23-24", "label": "Q4 FY2023-24", "q_type": "Q4", "period": "Jan 2024 – Mar 2024", "ann": "Apr 2024 – May 2024"},
    {"code": "Q1_FY24-25", "label": "Q1 FY2024-25", "q_type": "Q1", "period": "Apr 2024 – Jun 2024", "ann": "Jul 2024 – Aug 2024"},
    {"code": "Q2_FY24-25", "label": "Q2 FY2024-25", "q_type": "Q2", "period": "Jul 2024 – Sep 2024", "ann": "Oct 2024 – Nov 2024"},
    {"code": "Q3_FY24-25", "label": "Q3 FY2024-25", "q_type": "Q3", "period": "Oct 2024 – Dec 2024", "ann": "Jan 2025 – Feb 2025"},
    {"code": "Q4_FY24-25", "label": "Q4 FY2024-25", "q_type": "Q4", "period": "Jan 2025 – Mar 2025", "ann": "Apr 2025 – May 2025"},
    {"code": "Q1_FY25-26", "label": "Q1 FY2025-26", "q_type": "Q1", "period": "Apr 2025 – Jun 2025", "ann": "Jul 2025 – Aug 2025"},
    {"code": "Q2_FY25-26", "label": "Q2 FY2025-26", "q_type": "Q2", "period": "Jul 2025 – Sep 2025", "ann": "Oct 2025 – Nov 2025"},
    {"code": "Q3_FY25-26", "label": "Q3 FY2025-26", "q_type": "Q3", "period": "Oct 2025 – Dec 2025", "ann": "Jan 2026 – Feb 2026"},
    {"code": "Q4_FY25-26", "label": "Q4 FY2025-26", "q_type": "Q4", "period": "Jan 2026 – Mar 2026", "ann": "Apr 2026 – May 2026"},
    {"code": "Q1_FY26-27", "label": "Q1 FY2026-27", "q_type": "Q1", "period": "Apr 2026 – Jun 2026", "ann": "Jul 2026 – Aug 2026"},
    {"code": "Q2_FY26-27", "label": "Q2 FY2026-27", "q_type": "Q2", "period": "Jul 2026 – Sep 2026", "ann": "Oct 2026 – Nov 2026"},
]

print("="*90)
print("  DATA VERIFICATION — 211 STOCKS MAPPED FROM PORTFOLIO MASTER EXCEL")
print("="*90)

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
    trade_dir = stock_info['strategy']  # 'LONG' or 'SHORT'
    raw_strat = stock_info['raw_strategy']
    
    for q_info in TARGET_QUARTERS:
        q_label = q_info['label']
        
        b_dt = None
        if not df_fr_clean.empty:
            for _, row in df_fr_clean.iterrows():
                r_bdt = row['bDate']
                r_tdt = row['tDate']
                ref_dt = r_tdt if pd.notnull(r_tdt) else r_bdt
                m = ref_dt.month
                y = ref_dt.year
                
                if m in [1, 2, 3] and q_label == f"Q4 FY{y-1}-{str(y)[-2:]}": b_dt = r_bdt; break
                elif m in [4, 5, 6] and q_label == f"Q1 FY{y}-{str(y+1)[-2:]}": b_dt = r_bdt; break
                elif m in [7, 8, 9] and q_label == f"Q2 FY{y}-{str(y+1)[-2:]}": b_dt = r_bdt; break
                elif m in [10, 11, 12] and q_label == f"Q3 FY{y}-{str(y+1)[-2:]}": b_dt = r_bdt; break
                
        if b_dt is None:
            quarter_trades[q_label].append({
                "symbol": sym, "sector": "Equity", "status": stock_status, "n_quarters": stock_info['hist_q_str'],
                "strategy": raw_strat, "window": window_str,
                "broadcast_date": "N/A", "entry_date": "N/A", "entry_price": "N/A",
                "exit_date": "N/A", "exit_price": "N/A", "lot_size": lot_size,
                "pos_value": "N/A", "margin_req": "N/A", "net_pnl": "N/A",
                "ret_pct": "N/A", "win_loss": "N/A", "is_valid": False
            })
            continue

        res_idx = min(range(len(valid_dates)), key=lambda i: abs((valid_dates[i] - b_dt).days))
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
        else: # SHORT
            gross_pnl = (p_entry - p_exit) * lot_size
            
        comb_turnover = (p_entry + p_exit) * lot_size
        cost = comb_turnover * 0.0005
        net_pnl = round(gross_pnl - cost, 2)
        ret_pct = round((net_pnl / pos_val) * 100, 2) if pos_val > 0 else 0.0
        win_loss = "WIN" if net_pnl > 0 else "LOSS"
        
        quarter_trades[q_label].append({
            "symbol": sym, "sector": "Equity", "status": stock_status, "n_quarters": stock_info['hist_q_str'],
            "strategy": raw_strat, "window": window_str, "broadcast_date": b_dt.strftime("%Y-%m-%d"),
            "entry_date": entry_dt.strftime("%Y-%m-%d"), "entry_price": p_entry,
            "exit_date": exit_dt.strftime("%Y-%m-%d"), "exit_price": p_exit,
            "lot_size": lot_size, "pos_value": pos_val, "margin_req": margin_req,
            "net_pnl": net_pnl, "ret_pct": ret_pct, "win_loss": win_loss,
            "is_valid": True
        })

print("\nPERFORMANCE VERIFICATION (MAPPED FROM PORTFOLIO MASTER EXCEL):")
print(f"{'Quarter':<15} | {'Total Stocks':<12} | {'Executed Trades':<16} | {'N/A Trades':<12} | {'QUALIFIED Exec':<15} | {'Net P&L (₹)':<15} | {'Win Rate %':<12}")
print("-" * 105)

for q in TARGET_QUARTERS:
    q_label = q['label']
    trades = quarter_trades[q_label]
    n_total = len(trades)
    n_exec = sum(1 for t in trades if t['is_valid'])
    n_na = n_total - n_exec
    n_qual_exec = sum(1 for t in trades if t['is_valid'] and t['status'] == 'QUALIFIED')
    
    valid_trades = [t for t in trades if t['is_valid']]
    net_pnl = sum(t['net_pnl'] for t in valid_trades)
    wins = sum(1 for t in valid_trades if t['win_loss'] == 'WIN')
    wr = (wins / len(valid_trades) * 100) if len(valid_trades) > 0 else 0.0
    
    print(f"{q_label:<15} | {n_total:<12} | {n_exec:<16} | {n_na:<12} | {n_qual_exec:<15} | {net_pnl:<15,.2f} | {wr:<12.2f}%")

print("="*105)
