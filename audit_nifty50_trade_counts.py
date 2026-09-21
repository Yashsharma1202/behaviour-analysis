import pandas as pd, pathlib, sys

sys.stdout.reconfigure(errors='replace')

BASE_DIR = pathlib.Path('D:/behaviour analysis')
CONSOLIDATED_PATH = BASE_DIR / 'Nifty50_12_Quarters_Consolidated.xlsx'

xl = pd.ExcelFile(CONSOLIDATED_PATH, engine='openpyxl')

print("==========================================")
print("AUDITING TRADE COUNTS & SYMBOLS PER QUARTER")
print("==========================================")

quarters_order = [
    'Q3 2023-24', 'Q4 2023-24', 'Q1 2024-25', 'Q2 2024-25',
    'Q3 2024-25', 'Q4 2024-25', 'Q1 2025-26', 'Q2 2025-26',
    'Q3 2025-26', 'Q4 2025-26', 'Q1 2026-27', 'Q2 2026-27'
]

# Get standard 50 Nifty symbols from a 50-trade quarter like Q3 2023-24
df_q3 = xl.parse('Q3 2023-24')
t_q3 = df_q3.iloc[6:].copy().dropna(subset=[df_q3.columns[1]])
nifty_50_symbols = list(t_q3['Unnamed: 1'].astype(str).str.strip().unique())

print(f"Standard Nifty 50 Symbols Count: {len(nifty_50_symbols)}")
print("Symbols:", nifty_50_symbols[:10], "... [total 50]")

for q in quarters_order:
    sheet_name = next(s for s in xl.sheet_names if q in s)
    df_q = xl.parse(sheet_name)
    trades = df_q.iloc[6:].copy().dropna(subset=[df_q.columns[1]])
    
    syms = list(trades['Unnamed: 1'].astype(str).str.strip())
    unique_syms = set(syms)
    
    counts = len(trades)
    print(f"\nQuarter {q}: {counts} trades, {len(unique_syms)} unique symbols")
    
    if counts > 50:
        # Find duplicate symbols
        dups = [s for s in syms if syms.count(s) > 1]
        print(f"  ⚠️ Duplicate symbols in {q}: {set(dups)}")
    elif counts < 50:
        # Find missing symbol
        missing = set(nifty_50_symbols) - set(syms)
        print(f"  ⚠️ Missing symbol in {q}: {missing}")

print("==========================================")
