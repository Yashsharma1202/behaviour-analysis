import pandas as pd, sys

sys.stdout.reconfigure(errors='replace')

path = r'D:/behaviour analysis/Nifty50_12_Quarters_Consolidated.xlsx'
xl = pd.ExcelFile(path, engine='openpyxl')

quarters = [
    "Q3 2023-24", "Q4 2023-24", "Q1 2024-25", "Q2 2024-25",
    "Q3 2024-25", "Q4 2024-25", "Q1 2025-26", "Q2 2025-26",
    "Q3 2025-26", "Q4 2025-26", "Q1 2026-27", "Q2 2026-27"
]

print("==========================================")
print("Quarter Summary Metrics from Consolidated Master:")
print("==========================================")

for q in quarters:
    # find sheet matching q
    sheet_name = next(s for s in xl.sheet_names if q in s)
    df_q = xl.parse(sheet_name)
    
    row_hdr = df_q.iloc[2].values
    row_val = df_q.iloc[3].values
    
    print(f"\nQuarter: {q}")
    for h, v in zip(row_hdr[:13], row_val[:13]):
        if pd.notnull(h):
            print(f"  {str(h):<25}: {v}")
