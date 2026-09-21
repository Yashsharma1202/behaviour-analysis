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

LOT_SIZES = {
    "RELIANCE": 250, "TCS": 175, "INFY": 400, "HDFCBANK": 550, "ICICIBANK": 700,
    "BHARTIARTL": 950, "ITC": 1600, "SBIN": 1500, "LTIM": 150, "LT": 300,
    "HINDUNILVR": 300, "AXISBANK": 625, "KOTAKBANK": 400, "BAJFINANCE": 125,
    "M&M": 350, "MARUTI": 100, "SUNPHARMA": 350, "TATASTEEL": 5500
}

pre_candidates = [1, 2, 3, 4, 5, 7, 8]
post_candidates = [1, 2, 3, 4, 5, 8]

def assign_financial_quarter(row):
    rel = str(row.get('relatingTo', '')).lower()
    if 'first' in rel or 'q1' in rel: return 'Q1'
    elif 'second' in rel or 'q2' in rel: return 'Q2'
    elif 'third' in rel or 'q3' in rel: return 'Q3'
    elif 'fourth' in rel or 'q4' in rel: return 'Q4'
        
    to_dt = pd.to_datetime(row.get('toDate', ''), errors='coerce')
    if pd.notnull(to_dt):
        m = to_dt.month
        if m in (4, 5, 6): return 'Q1'
        elif m in (7, 8, 9): return 'Q2'
        elif m in (10, 11, 12): return 'Q3'
        elif m in (1, 2, 3): return 'Q4'

    bc_dt = pd.to_datetime(row.get('broadCastDate', ''), errors='coerce')
    if pd.notnull(bc_dt):
        m = bc_dt.month
        if m in (7, 8, 9): return 'Q1'
        elif m in (10, 11, 12): return 'Q2'
        elif m in (1, 2, 3): return 'Q3'
        else: return 'Q4'
        
    return 'Q1'

print("Testing Quarter-by-Quarter Position Window Optimization for RELIANCE and TCS...")

for sym in ["RELIANCE", "TCS"]:
    fr_path = ROOT / sym / 'financial_results.csv'
    price_path = PRICE_CACHE / f"{sym}.csv"
    if not fr_path.exists() or not price_path.exists(): continue
    
    df_prices = pd.read_csv(price_path)
    df_prices['date'] = pd.to_datetime(df_prices['date'])
    df_prices = df_prices.sort_values('date').reset_index(drop=True)
    valid_dates = df_prices['date'].tolist()
    
    df_fr = pd.read_csv(fr_path, dtype=str).fillna('')
    df_fr['dt'] = pd.to_datetime(df_fr['broadCastDate'], errors='coerce')
    df_fr = df_fr.dropna(subset=['dt']).sort_values('dt', ascending=False)
    df_fr['year'] = df_fr['dt'].dt.year
    df_fr['q_type'] = df_fr.apply(assign_financial_quarter, axis=1)
    df_fr_clean = df_fr.drop_duplicates(subset=['year', 'q_type']).copy()
    
    lot_size = LOT_SIZES.get(sym, 250)
    print(f"\nStock: {sym} (Total Clean Quarters: {len(df_fr_clean)})")
    
    for q_t in ["Q1", "Q2", "Q3", "Q4"]:
        df_q = df_fr_clean[df_fr_clean['q_type'] == q_t]
        if df_q.empty: continue
        
        best_pnl = -999999999
        best_combo = None
        
        for pre in pre_candidates:
            for post in post_candidates:
                emp_long_wins = 0; emp_short_wins = 0
                pnl_l = 0.0; pnl_s = 0.0
                tot = len(df_q)
                
                for _, r in df_q.iterrows():
                    dt = r['dt']
                    r_i = min(range(len(valid_dates)), key=lambda i: abs((valid_dates[i] - dt).days))
                    e_i = max(0, r_i - pre)
                    x_i = min(len(valid_dates) - 1, r_i + post)
                    pe = float(df_prices.loc[e_i, 'adj']); px = float(df_prices.loc[x_i, 'adj'])
                    if pe > 0:
                        b_l = round(lot_size * pe, 2); s_l = round(lot_size * px, 2)
                        c_l = round((b_l + s_l) * 0.0005, 2)
                        net_l = round(s_l - b_l - c_l, 2)
                        pnl_l += net_l
                        if net_l > 0: emp_long_wins += 1
                        
                        b_s = round(lot_size * px, 2); s_s = round(lot_size * pe, 2)
                        c_s = round((b_s + s_s) * 0.0005, 2)
                        net_s = round(s_s - b_s - c_s, 2)
                        pnl_s += net_s
                        if net_s > 0: emp_short_wins += 1
                        
                wr_l = (emp_long_wins / tot) * 100
                wr_s = (emp_short_wins / tot) * 100
                
                if wr_l >= wr_s:
                    strat = "FUTURE LONG"; wr = wr_l; pnl = pnl_l
                else:
                    strat = "FUTURE SHORT"; wr = wr_s; pnl = pnl_s
                    
                if pnl > best_pnl:
                    best_pnl = pnl
                    best_combo = (pre, post, strat, wr, pnl, tot)
                    
        pre, post, strat, wr, pnl, tot = best_combo
        print(f"  • {q_t}: Best Window = T-{pre} to T+{post} | Strategy = {strat} | Win Rate = {wr:.1f}% | Net P&L = ₹{pnl:,.2f} ({tot} Qtrs)")
