import os
import pathlib
import sys
import pandas as pd

sys.stdout.reconfigure(errors='replace')

ROOT = pathlib.Path('D:/behaviour analysis')
OI_DIR = ROOT / 'OI_DATA'

SYMBOLS = sorted([e.name.strip().upper() for e in os.scandir(OI_DIR) if e.is_dir()])

print("==========================================================================================")
print("INSPECTING FINANCIAL QUARTER MAPPING ACROSS 211 STOCKS")
print("==========================================================================================")

sample_counts = {"Q1 (Apr-Jun)": 0, "Q2 (Jul-Sep)": 0, "Q3 (Oct-Dec)": 0, "Q4 (Jan-Mar)": 0, "Other/Annual": 0}

for sym in SYMBOLS[:10]:
    fr_path = ROOT / sym / 'financial_results.csv'
    if not fr_path.exists():
        continue
    df = pd.read_csv(fr_path, dtype=str).fillna('')
    print(f"\nStock: {sym} (Total rows: {len(df)})")
    if 'relatingTo' in df.columns and 'broadCastDate' in df.columns:
        for idx, r in df.head(6).iterrows():
            print(f"  Broadcast: {r.get('broadCastDate', '')} | relatingTo: {r.get('relatingTo', '')} | toDate: {r.get('toDate', '')} | period: {r.get('period', '')}")
