import pathlib
import sys
import pandas as pd

sys.stdout.reconfigure(errors='replace')

ROOT = pathlib.Path('D:/behaviour analysis')

test_syms = ["COCHINSHIP", "ATHERENERG", "HAL", "RVNL", "KAYNES"]

for sym in test_syms:
    fr_path = ROOT / sym / 'financial_results.csv'
    print(f"\n==========================================================================================")
    print(f"EXACT BROADCAST DATES IN financial_results.csv FOR {sym}")
    print(f"==========================================================================================")
    if fr_path.exists():
        df_fr = pd.read_csv(fr_path, dtype=str).fillna('')
        if 'broadCastDate' in df_fr.columns:
            dts = pd.to_datetime(df_fr['broadCastDate'], errors='coerce').dropna().sort_values(ascending=False)
            # Deduplicate dates
            dts_unique = pd.Series(dts.unique()).sort_values(ascending=False)
            print(f"Total Unique Broadcast Dates: {len(dts_unique)}")
            for d in dts_unique:
                m = d.month
                if m in (4, 5, 6): q_type = "Q1 (Apr/May/Jun)"
                elif m in (7, 8, 9): q_type = "Q2 (Jul/Aug/Sep)"
                elif m in (10, 11, 12): q_type = "Q3 (Oct/Nov/Dec)"
                else: q_type = "Q4 (Jan/Feb/Mar)"
                print(f"  • Date: {d.strftime('%Y-%m-%d')} | Month: {d.month:02d} | Quarter Type: {q_type}")

print("==========================================================================================")
