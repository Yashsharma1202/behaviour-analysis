import pathlib
import sys
import pandas as pd

sys.stdout.reconfigure(errors='replace')

ROOT = pathlib.Path('D:/behaviour analysis')

test_syms = ["ATHERENERG", "COCHINSHIP", "HAL", "RVNL", "KAYNES", "COALINDIA"]

print("==========================================================================================")
print("FINANCIAL QUARTER DEDUPLICATION & UNIQUE PERIOD GROUPING AUDIT")
print("==========================================================================================")

for sym in test_syms:
    fr_path = ROOT / sym / 'financial_results.csv'
    print(f"\nAUDITING STOCK: {sym}")
    print("-" * 90)
    if not fr_path.exists():
        continue
    df_fr = pd.read_csv(fr_path, dtype=str).fillna('')
    if 'broadCastDate' not in df_fr.columns:
        continue
        
    df_fr['dt'] = pd.to_datetime(df_fr['broadCastDate'], errors='coerce')
    df_fr = df_fr.dropna(subset=['dt']).sort_values('dt', ascending=False)
    
    # Extract Year and Month
    df_fr['year'] = df_fr['dt'].dt.year
    df_fr['month'] = df_fr['dt'].dt.month
    
    # Assign Quarter Type: Q1 (4,5,6), Q2 (7,8,9), Q3 (10,11,12), Q4 (1,2,3)
    def assign_q(m):
        if m in (4, 5, 6): return "Q1"
        elif m in (7, 8, 9): return "Q2"
        elif m in (10, 11, 12): return "Q3"
        else: return "Q4"
        
    df_fr['q_type'] = df_fr['month'].apply(assign_q)
    
    # For deduplication, group by (year, q_type) so multiple filings in the same quarter count as 1 unique financial quarter!
    unique_q_periods = df_fr[['year', 'q_type']].drop_duplicates()
    
    total_unique_periods = len(unique_q_periods)
    q1_unique = len(unique_q_periods[unique_q_periods['q_type'] == 'Q1'])
    q2_unique = len(unique_q_periods[unique_q_periods['q_type'] == 'Q2'])
    q3_unique = len(unique_q_periods[unique_q_periods['q_type'] == 'Q3'])
    q4_unique = len(unique_q_periods[unique_q_periods['q_type'] == 'Q4'])
    
    print(f"  • Raw Filings Count: {len(df_fr)}")
    print(f"  • Deduplicated Unique Financial Quarters Count: {total_unique_periods}")
    print(f"  • Breakdown: Q1={q1_unique} | Q2={q2_unique} | Q3={q3_unique} | Q4={q4_unique}")
    print(f"  • Detailed Unique Financial Quarter Periods:")
    for _, row in unique_q_periods.iterrows():
        print(f"     - Year {row['year']} {row['q_type']}")

print("==========================================================================================")
