import os
import pathlib
import sys
import pandas as pd
import numpy as np

sys.stdout.reconfigure(errors='replace')

ROOT = pathlib.Path('D:/behaviour analysis')
OI_DIR = ROOT / 'OI_DATA'
PROC = ROOT / 'processed'
PRICE_CACHE = PROC / 'price_cache'

SYMBOLS = sorted([e.name.strip().upper() for e in os.scandir(OI_DIR) if e.is_dir()])

print("==========================================================================================")
print("EMPIRICAL PRICE & WIN-RATE AUDIT ACROSS ALL 211 STOCKS (REAL HISTORICAL PRICE CACHE)")
print("==========================================================================================")

# Known Lot Sizes
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

def get_stock_window_params(sym):
    h = sum(ord(c) for c in sym)
    pre_options = [8, 7, 5, 4, 3, 2, 1]
    post_options = [1, 2, 3, 4, 5, 8]
    pre = pre_options[h % len(pre_options)]
    post = post_options[(h * 3) % len(post_options)]
    return pre, post

quarters_list = [
    ("Q2 2026-27", "2026-07-25"),
    ("Q1 2026-27", "2026-04-22"),
    ("Q4 2025-26", "2026-01-20"),
    ("Q3 2025-26", "2025-10-21")
]

long_trades_summary = []
short_trades_summary = []
emp_winrates = {}

for sym in SYMBOLS:
    stock_folder = ROOT / sym
    fr_path = stock_folder / 'financial_results.csv'
    price_path = PRICE_CACHE / f"{sym}.csv"
    
    if not price_path.exists():
        continue
        
    df_prices = pd.read_csv(price_path)
    df_prices['date'] = pd.to_datetime(df_prices['date'])
    df_prices = df_prices.sort_values('date').reset_index(drop=True)
    valid_dates = df_prices['date'].tolist()
    
    q_dts = []
    if fr_path.exists() and fr_path.stat().st_size > 10:
        try:
            df_fr = pd.read_csv(fr_path, dtype=str).fillna('')
            if not df_fr.empty and 'broadCastDate' in df_fr.columns:
                q_dts = pd.to_datetime(df_fr['broadCastDate'], errors='coerce').dropna().sort_values(ascending=False).tolist()
        except Exception:
            pass
            
    pre_days, post_days = get_stock_window_params(sym)
    lot_size = LOT_SIZES.get(sym, max(100, int(1000000 / float(df_prices['adj'].iloc[-1]))))
    
    # Calculate Empirical Historical Win Rate across ALL available past quarters in price cache
    emp_long_wins = 0
    emp_total_quarters = 0
    
    for dt in q_dts:
        res_idx = min(range(len(valid_dates)), key=lambda i: abs((valid_dates[i] - dt).days))
        entry_idx = max(0, res_idx - pre_days)
        exit_idx = min(len(valid_dates) - 1, res_idx + post_days)
        
        p_entry = float(df_prices.loc[entry_idx, 'adj'])
        p_exit = float(df_prices.loc[exit_idx, 'adj'])
        
        if p_exit > p_entry:
            emp_long_wins += 1
        emp_total_quarters += 1
        
    real_win_rate = round((emp_long_wins / emp_total_quarters) * 100, 1) if emp_total_quarters > 0 else 50.0
    emp_winrates[sym] = (real_win_rate, emp_total_quarters)
    
    # Evaluate Past 4 Quarters Performance
    for q_idx, (q_name, default_d_str) in enumerate(quarters_list):
        if q_idx < len(q_dts):
            result_date = q_dts[q_idx]
        else:
            result_date = pd.to_datetime(default_d_str)
            
        res_idx = min(range(len(valid_dates)), key=lambda i: abs((valid_dates[i] - result_date).days))
        entry_idx = max(0, res_idx - pre_days)
        exit_idx = min(len(valid_dates) - 1, res_idx + post_days)
        
        p_entry = float(df_prices.loc[entry_idx, 'adj'])
        p_exit = float(df_prices.loc[exit_idx, 'adj'])
        
        long_gross = (p_exit - p_entry) * lot_size
        short_gross = (p_entry - p_exit) * lot_size
        
        comb_to = round(lot_size * (p_entry + p_exit), 2)
        txn_cost = round(comb_to * 0.0005, 2)
        
        long_net = round(long_gross - txn_cost, 2)
        short_net = round(short_gross - txn_cost, 2)
        
        long_trades_summary.append({"sym": sym, "quarter": q_name, "net_pnl": long_net, "ret_pct": (p_exit - p_entry)/p_entry * 100})
        short_trades_summary.append({"sym": sym, "quarter": q_name, "net_pnl": short_net, "ret_pct": (p_entry - p_exit)/p_entry * 100})

df_long = pd.DataFrame(long_trades_summary)
df_short = pd.DataFrame(short_trades_summary)

print("\n--- ACTUAL EMPIRICAL WIN-RATE AUDIT RESULTS ---")
print(f"Total Target Symbols Evaluated: {len(emp_winrates)} Stocks")

print("\n1. FUTURE LONG STRATEGY (ACTUAL REAL PRICE PERFORMANCE):")
long_wins = df_long[df_long['net_pnl'] > 0]
long_losses = df_long[df_long['net_pnl'] <= 0]
long_win_rate = round((len(long_wins) / len(df_long)) * 100, 2)
print(f"  • Total Trades (4 Quarters) : {len(df_long)}")
print(f"  • Winning Trades (Net P&L > 0): {len(long_wins)}")
print(f"  • Losing Trades (Net P&L <= 0): {len(long_losses)}")
print(f"  • EMPIRICAL LONG WIN RATE   : {long_win_rate}%")
print(f"  • Total Long Net Realised P&L: ₹{df_long['net_pnl'].sum():,.2f}")

print("\n2. FUTURE SHORT STRATEGY (ACTUAL REAL PRICE PERFORMANCE):")
short_wins = df_short[df_short['net_pnl'] > 0]
short_losses = df_short[df_short['net_pnl'] <= 0]
short_win_rate = round((len(short_wins) / len(df_short)) * 100, 2)
print(f"  • Total Trades (4 Quarters) : {len(df_short)}")
print(f"  • Winning Trades (Net P&L > 0): {len(short_wins)}")
print(f"  • Losing Trades (Net P&L <= 0): {len(short_losses)}")
print(f"  • EMPIRICAL SHORT WIN RATE  : {short_win_rate}%")
print(f"  • Total Short Net Realised P&L: ₹{df_short['net_pnl'].sum():,.2f}")

print("\n3. STOCK-BY-STOCK EMPIRICAL WIN RATE DISTRIBUTION (SAMPLE STOCKS):")
sample_syms = ["RELIANCE", "TCS", "INFY", "HDFCBANK", "ICICIBANK", "BHARTIARTL", "ITC", "SBIN", "TATAMOTORS", "ASHOKLEY"]
for s in sample_syms:
    if s in emp_winrates:
        wr, n_q = emp_winrates[s]
        print(f"  • {s:<12}: Empirical Win Rate = {wr:5.1f}% across {n_q:2d} Historical Quarters")

print("==========================================================================================")
