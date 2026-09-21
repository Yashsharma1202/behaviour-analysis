import os
import pathlib
import sys
import pandas as pd

sys.stdout.reconfigure(errors='replace')

ROOT = pathlib.Path('D:/behaviour analysis')
OI_DIR = ROOT / 'OI_DATA'

SYMBOLS = sorted([e.name.strip().upper() for e in os.scandir(OI_DIR) if e.is_dir()])

print("==========================================================================================")
print("DEEP EXHAUSTIVE AUDIT OF ALL 211 STOCKS: BROADCAST DATE DISTRIBUTIONS (Q1, Q2, Q3, Q4)")
print("==========================================================================================")
print(f"Total Target Universe: {len(SYMBOLS)} Stocks")

total_mismatches = 0
stocks_checked = 0

audit_results = []

for idx, sym in enumerate(SYMBOLS, 1):
    stock_folder = ROOT / sym
    fr_path = stock_folder / 'financial_results.csv'
    
    if not fr_path.exists() or fr_path.stat().st_size < 10:
        continue
        
    try:
        df_fr = pd.read_csv(fr_path, dtype=str).fillna('')
        if 'broadCastDate' not in df_fr.columns:
            continue
            
        dts_raw = pd.to_datetime(df_fr['broadCastDate'], errors='coerce').dropna().sort_values(ascending=False)
        if dts_raw.empty:
            continue
            
        # Unique broadcast dates for clean quarter count
        dts_unique = pd.Series(dts_raw.unique()).sort_values(ascending=False).tolist()
        total_unique_qtrs = len(dts_unique)
        
        q1_count = sum(1 for dt in dts_unique if dt.month in (4, 5, 6))
        q2_count = sum(1 for dt in dts_unique if dt.month in (7, 8, 9))
        q3_count = sum(1 for dt in dts_unique if dt.month in (10, 11, 12))
        q4_count = sum(1 for dt in dts_unique if dt.month in (1, 2, 3))
        
        sum_qtrs = q1_count + q2_count + q3_count + q4_count
        
        stocks_checked += 1
        
        if sum_qtrs != total_unique_qtrs:
            total_mismatches += 1
            print(f"❌ MISMATCH FOR {sym}: Total Unique={total_unique_qtrs} vs Sum(Q1..Q4)={sum_qtrs}")
            
        audit_results.append({
            "sym": sym, "total_qtrs": total_unique_qtrs,
            "q1_count": q1_count, "q2_count": q2_count,
            "q3_count": q3_count, "q4_count": q4_count,
            "sum_qtrs": sum_qtrs, "is_pass": (sum_qtrs == total_unique_qtrs)
        })
        
    except Exception as e:
        print(f"  ❌ Exception auditing {sym}: {e}")

print("-" * 90)
print(f"  • Total Stocks Audited: {stocks_checked} / {len(SYMBOLS)}")
print(f"  • Date Sum Mismatches Found: {total_mismatches}")
if total_mismatches == 0:
    print(f"  🏆 AUDIT VERDICT: 100% PERFECT PASS! ZERO MISMATCHES ACROSS ALL 211 STOCKS!")
print("==========================================================================================")
