import pandas as pd, pathlib, sys

sys.stdout.reconfigure(errors='replace')

BASE_DIR = pathlib.Path('D:/behaviour analysis')
CONSOLIDATED_PATH = BASE_DIR / 'Nifty50_12_Quarters_Consolidated.xlsx'

xl = pd.ExcelFile(CONSOLIDATED_PATH, engine='openpyxl')

quarters_order = [
    'Q3 2023-24', 'Q4 2023-24', 'Q1 2024-25', 'Q2 2024-25',
    'Q3 2024-25', 'Q4 2024-25', 'Q1 2025-26', 'Q2 2025-26',
    'Q3 2025-26', 'Q4 2025-26', 'Q1 2026-27', 'Q2 2026-27'
]

print("==========================================")
print("EXACT TRADE COUNT & SYMBOL AUDIT PER QUARTER")
print("==========================================")

# Standard 50 Nifty symbols from recent quarter (e.g. Q2 2026-27)
df_recent = xl.parse('Q2 2026-27')
trades_recent = df_recent.iloc[6:].copy()
trades_recent = trades_recent[trades_recent['Unnamed: 1'] != 'Symbol'].dropna(subset=['Unnamed: 1'])
recent_50_symbols = list(trades_recent['Unnamed: 1'].astype(str).str.strip().unique())

print(f"Recent Quarter (Q2 2026-27) Symbol Count: {len(recent_50_symbols)}")
print("50 Stock Symbols:", sorted(recent_50_symbols))

for q in quarters_order:
    sheet_name = next(s for s in xl.sheet_names if q in s)
    df_q = xl.parse(sheet_name)
    trades = df_q.iloc[6:].copy()
    trades = trades[trades['Unnamed: 1'] != 'Symbol'].dropna(subset=['Unnamed: 1'])
    
    syms = list(trades['Unnamed: 1'].astype(str).str.strip())
    counts = len(syms)
    
    missing_from_recent = set(recent_50_symbols) - set(syms)
    extra_than_recent = set(syms) - set(recent_50_symbols)
    
    print(f"\nQuarter {q}: {counts} trades")
    if missing_from_recent:
        print(f"  ❌ Missing stock(s) vs Recent 50: {missing_from_recent}")
    if extra_than_recent:
        print(f"  ➕ Extra stock(s) vs Recent 50: {extra_than_recent}")

print("==========================================")
