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

TARGET_QUARTERS = [
    {"code": "Q3_FY23-24", "label": "Q3 FY2023-24", "period": "Oct 2023 – Dec 2023", "ann": "Jan 2024 – Feb 2024"},
    {"code": "Q4_FY23-24", "label": "Q4 FY2023-24", "period": "Jan 2024 – Mar 2024", "ann": "Apr 2024 – May 2024"},
    {"code": "Q1_FY24-25", "label": "Q1 FY2024-25", "period": "Apr 2024 – Jun 2024", "ann": "Jul 2024 – Aug 2024"},
    {"code": "Q2_FY24-25", "label": "Q2 FY2024-25", "period": "Jul 2024 – Sep 2024", "ann": "Oct 2024 – Nov 2024"},
    {"code": "Q3_FY24-25", "label": "Q3 FY2024-25", "period": "Oct 2024 – Dec 2024", "ann": "Jan 2025 – Feb 2025"},
    {"code": "Q4_FY24-25", "label": "Q4 FY2024-25", "period": "Jan 2025 – Mar 2025", "ann": "Apr 2025 – May 2025"},
    {"code": "Q1_FY25-26", "label": "Q1 FY2025-26", "period": "Apr 2025 – Jun 2025", "ann": "Jul 2025 – Aug 2025"},
    {"code": "Q2_FY25-26", "label": "Q2 FY2025-26", "period": "Jul 2025 – Sep 2025", "ann": "Oct 2025 – Nov 2025"},
    {"code": "Q3_FY25-26", "label": "Q3 FY2025-26", "period": "Oct 2025 – Dec 2025", "ann": "Jan 2026 – Feb 2026"},
    {"code": "Q4_FY25-26", "label": "Q4 FY2025-26", "period": "Jan 2026 – Mar 2026", "ann": "Apr 2026 – May 2026"},
    {"code": "Q1_FY26-27", "label": "Q1 FY2026-27", "period": "Apr 2026 – Jun 2026", "ann": "Jul 2026 – Aug 2026"},
    {"code": "Q2_FY26-27", "label": "Q2 FY2026-27", "period": "Jul 2026 – Sep 2026", "ann": "Oct 2026 – Nov 2026"},
]

def get_stock_window_params(sym):
    h = sum(ord(c) for c in sym)
    pre_options = [8, 7, 5, 4, 3, 2, 1]
    post_options = [1, 2, 3, 4, 5, 8]
    pre = pre_options[h % len(pre_options)]
    post = post_options[(h * 3) % len(post_options)]
    return pre, post

quarter_counts = {q['label']: {'total_rows': 0, 'executed': 0, 'na_rows': 0, 'qual_executed': 0, 'qual_wins': 0} for q in TARGET_QUARTERS}

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
        except Exception: pass
            
    if df_prices.empty: continue
    valid_dates = df_prices['date'].tolist()
    pre_days, post_days = get_stock_window_params(sym)
    
    df_fr_clean = pd.DataFrame()
    n_quarters = 0
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
                n_quarters = len(df_fr_clean)
        except Exception: pass
            
    status = "QUALIFIED" if n_quarters >= 12 else "AVOID BUT MONITOR IT"
    
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
                
        quarter_counts[q_label]['total_rows'] += 1
        
        if b_dt is not None:
            res_idx = min(range(len(valid_dates)), key=lambda i: abs((valid_dates[i] - b_dt).days))
            entry_idx = max(0, res_idx - pre_days)
            exit_idx = min(len(valid_dates) - 1, res_idx + post_days)
            p_entry = float(df_prices.loc[entry_idx, 'adj'])
            p_exit = float(df_prices.loc[exit_idx, 'adj'])
            
            pnl = (p_exit - p_entry) if p_exit >= p_entry else (p_entry - p_exit)
            cost = (p_entry + p_exit) * 0.0005
            net_pnl = pnl - cost
            is_win = net_pnl > 0
            
            quarter_counts[q_label]['executed'] += 1
            if status == 'QUALIFIED':
                quarter_counts[q_label]['qual_executed'] += 1
                if is_win:
                    quarter_counts[q_label]['qual_wins'] += 1
        else:
            quarter_counts[q_label]['na_rows'] += 1

print("\n--- NA LOGIC VERIFICATION ACROSS ALL 12 QUARTERS ---")
print(f"{'Quarter':<15} | {'Total Rows':<10} | {'Executed':<10} | {'N/A Rows':<10} | {'Qual WR %':<12}")
print("-" * 75)

for q in TARGET_QUARTERS:
    ql = q['label']
    c = quarter_counts[ql]
    wr = (c['qual_wins'] / c['qual_executed'] * 100) if c['qual_executed'] > 0 else 0.0
    print(f"{ql:<15} | {c['total_rows']:<10} | {c['executed']:<10} | {c['na_rows']:<10} | {wr:<12.2f}%")
