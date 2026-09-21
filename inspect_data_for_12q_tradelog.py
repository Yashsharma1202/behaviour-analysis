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
print(f"Total Universe Symbols: {len(SYMBOLS)}")

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

# Collect all result dates per stock
all_results = []

for sym in SYMBOLS:
    stock_folder = ROOT / sym
    fr_path = stock_folder / 'financial_results.csv'
    price_path = PRICE_CACHE / f"{sym}.csv"
    
    if not fr_path.exists() or fr_path.stat().st_size <= 10:
        continue
        
    try:
        df_fr = pd.read_csv(fr_path, dtype=str).fillna('')
        if df_fr.empty or 'broadCastDate' not in df_fr.columns:
            continue
            
        df_fr['bDate'] = pd.to_datetime(df_fr['broadCastDate'], errors='coerce')
        df_fr = df_fr.dropna(subset=['bDate'])
        
        # also parse toDate if present
        if 'toDate' in df_fr.columns:
            df_fr['tDate'] = pd.to_datetime(df_fr['toDate'], errors='coerce')
        else:
            df_fr['tDate'] = pd.NaT
            
        for _, row in df_fr.iterrows():
            b_dt = row['bDate']
            t_dt = row['tDate']
            
            # Determine Financial Quarter from broadcast date or toDate
            # Standard Indian FY:
            # Q1: Apr-Jun result (announced Jul-Aug) -> toDate in Jun
            # Q2: Jul-Sep result (announced Oct-Nov) -> toDate in Sep
            # Q3: Oct-Dec result (announced Jan-Feb) -> toDate in Dec
            # Q4: Jan-Mar result (announced Apr-May) -> toDate in Mar
            
            ref_dt = t_dt if pd.notnull(t_dt) else b_dt
            m = ref_dt.month
            y = ref_dt.year
            
            if m in [1, 2, 3]:
                q_name = f"Q4 FY{y-1}-{str(y)[-2:]}"
                q_order = y * 10 + 4
                q_period = f"Jan {y} - Mar {y}"
                results_ann = f"Apr {y} - May {y}"
            elif m in [4, 5, 6]:
                q_name = f"Q1 FY{y}-{str(y+1)[-2:]}"
                q_order = y * 10 + 1
                q_period = f"Apr {y} - Jun {y}"
                results_ann = f"Jul {y} - Aug {y}"
            elif m in [7, 8, 9]:
                q_name = f"Q2 FY{y}-{str(y+1)[-2:]}"
                q_order = y * 10 + 2
                q_period = f"Jul {y} - Sep {y}"
                results_ann = f"Oct {y} - Nov {y}"
            else:
                q_name = f"Q3 FY{y}-{str(y+1)[-2:]}"
                q_order = y * 10 + 3
                q_period = f"Oct {y} - Dec {y}"
                results_ann = f"Jan {y+1} - Feb {y+1}"
                
            all_results.append({
                'symbol': sym,
                'broadcast_date': b_dt,
                'to_date': t_dt,
                'quarter': q_name,
                'q_order': q_order,
                'q_period': q_period,
                'results_ann': results_ann
            })
    except Exception as e:
        pass

df_res = pd.DataFrame(all_results)
print(f"Total raw broadcast records found: {len(df_res)}")

# Deduplicate stock + quarter
df_res_dedup = df_res.sort_values('broadcast_date', ascending=False).drop_duplicates(subset=['symbol', 'quarter'])
print(f"Total deduped stock-quarter records: {len(df_res_dedup)}")

q_summary = df_res_dedup.groupby(['q_order', 'quarter', 'q_period', 'results_ann']).agg(
    stock_count=('symbol', 'nunique'),
    min_date=('broadcast_date', 'min'),
    max_date=('broadcast_date', 'max')
).reset_index().sort_values('q_order', ascending=False)

print("\n--- QUARTER BREAKDOWN ---")
print(q_summary.to_string(index=False))
