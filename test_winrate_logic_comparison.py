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

# 12 Target Quarters
TARGET_QUARTERS = [
    {"label": "Q3 FY2023-24", "q_type": "Q3", "target_year": 2024, "default_dt": "2024-01-22"},
    {"label": "Q4 FY2023-24", "q_type": "Q4", "target_year": 2024, "default_dt": "2024-04-22"},
    {"label": "Q1 FY2024-25", "q_type": "Q1", "target_year": 2024, "default_dt": "2024-07-22"},
    {"label": "Q2 FY2024-25", "q_type": "Q2", "target_year": 2024, "default_dt": "2024-10-22"},
    {"label": "Q3 FY2024-25", "q_type": "Q3", "target_year": 2025, "default_dt": "2025-01-22"},
    {"label": "Q4 FY2024-25", "q_type": "Q4", "target_year": 2025, "default_dt": "2025-04-22"},
    {"label": "Q1 FY2025-26", "q_type": "Q1", "target_year": 2025, "default_dt": "2025-07-22"},
    {"label": "Q2 FY2025-26", "q_type": "Q2", "target_year": 2025, "default_dt": "2025-10-22"},
    {"label": "Q3 FY2025-26", "q_type": "Q3", "target_year": 2026, "default_dt": "2026-01-22"},
    {"label": "Q4 FY2025-26", "q_type": "Q4", "target_year": 2026, "default_dt": "2026-04-22"},
    {"label": "Q1 FY2026-27", "q_type": "Q1", "target_year": 2026, "default_dt": "2026-07-22"},
    {"label": "Q2 FY2026-27", "q_type": "Q2", "target_year": 2026, "default_dt": "2026-10-22"},
]

def get_stock_window_params(sym):
    h = sum(ord(c) for c in sym)
    pre_options = [8, 7, 5, 4, 3, 2, 1]
    post_options = [1, 2, 3, 4, 5, 8]
    pre = pre_options[h % len(pre_options)]
    post = post_options[(h * 3) % len(post_options)]
    return pre, post

# We will test 3 strategy selection methods:
# Method A: Current Method (Strategy = LONG if avg_long_ret >= avg_short_ret else SHORT)
# Method B: Win-Rate Based Strategy (Strategy = LONG if hist_wr_long >= hist_wr_short else SHORT)
# Method C: Event-Optimal Mode (Strategy = LONG if long_gross >= short_gross else SHORT per trade)

results_A = []
results_B = []
results_C = []

for sym in SYMBOLS:
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
            
    if df_prices.empty: continue
        
    valid_dates = df_prices['date'].tolist()
    pre_days, post_days = get_stock_window_params(sym)
    
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
            
    # Calculate historical win rates and returns
    hist_rets_l, hist_rets_s = [], []
    wins_l, wins_s = 0, 0
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
            
            r_i = min(range(len(valid_dates)), key=lambda i: abs((valid_dates[i] - b_dt).days))
            e_i = max(0, r_i - pre_days)
            x_i = min(len(valid_dates) - 1, r_i + post_days)
            pe = float(df_prices.loc[e_i, 'adj'])
            px = float(df_prices.loc[x_i, 'adj'])
            if pe > 0:
                rl = (px - pe) / pe * 100
                rs = (pe - px) / pe * 100
                hist_rets_l.append(rl)
                hist_rets_s.append(rs)
                if px > pe: wins_l += 1
                elif pe > px: wins_s += 1
                
    total_h = len(hist_rets_l)
    wr_l = (wins_l / total_h * 100) if total_h > 0 else 50.0
    wr_s = (wins_s / total_h * 100) if total_h > 0 else 50.0
    avg_l = np.mean(hist_rets_l) if hist_rets_l else 0.0
    avg_s = np.mean(hist_rets_s) if hist_rets_s else 0.0
    
    strat_A = "LONG" if avg_l >= avg_s else "SHORT"
    strat_B = "LONG" if wr_l >= wr_s else "SHORT"
    
    # Process 12 Target Quarters
    for q_info in TARGET_QUARTERS:
        q_label = q_info['label']
        q_type = q_info['q_type']
        target_year = q_info['target_year']
        
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
            q_past = hist_q_dates.get(q_type, [])
            if q_past:
                avg_month = int(np.round(np.mean([d.month for d in q_past])))
                avg_day = int(np.round(np.mean([d.day for d in q_past])))
                try: b_dt = pd.Timestamp(year=target_year, month=avg_month, day=min(avg_day, 28))
                except Exception: b_dt = pd.to_datetime(q_info['default_dt'])
            else:
                b_dt = pd.to_datetime(q_info['default_dt'])
                
        res_idx = min(range(len(valid_dates)), key=lambda i: abs((valid_dates[i] - b_dt).days))
        entry_idx = max(0, res_idx - pre_days)
        exit_idx = min(len(valid_dates) - 1, res_idx + post_days)
        
        p_entry = float(df_prices.loc[entry_idx, 'adj'])
        p_exit = float(df_prices.loc[exit_idx, 'adj'])
        
        # Method A
        pnl_A = (p_exit - p_entry) if strat_A == "LONG" else (p_entry - p_exit)
        results_A.append(1 if pnl_A > 0 else 0)
        
        # Method B (Win-rate strategy)
        pnl_B = (p_exit - p_entry) if strat_B == "LONG" else (p_entry - p_exit)
        results_B.append(1 if pnl_B > 0 else 0)
        
        # Method C (Optimal Mode per Trade)
        pnl_long = (p_exit - p_entry)
        pnl_short = (p_entry - p_exit)
        pnl_C = max(pnl_long, pnl_short)
        results_C.append(1 if pnl_C > 0 else 0)

print(f"Total Trades Audited: {len(results_A)}")
print(f"Method A (Average Return Direction) Win Rate: {np.mean(results_A)*100:.2f}%")
print(f"Method B (Win-Rate Direction)       Win Rate: {np.mean(results_B)*100:.2f}%")
print(f"Method C (Event Optimal Direction)  Win Rate: {np.mean(results_C)*100:.2f}%")
