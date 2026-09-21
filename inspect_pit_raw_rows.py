import pandas as pd, sys

sys.stdout.reconfigure(errors='replace')

BASE_DIR = pathlib.Path('D:/behaviour analysis') if 'pathlib' in locals() else __import__('pathlib').Path('D:/behaviour analysis')
FUT_PIT = BASE_DIR / 'Nifty50_12_Quarters_Point_In_Time_Futures_Master.xlsx'

xl_pit = pd.ExcelFile(FUT_PIT, engine='openpyxl')
df_q3 = xl_pit.parse('Q3 2023-24', header=None)

for r in range(0, 10):
    row_vals = df_q3.iloc[r].tolist()
    print(f"Row {r}: {row_vals[:6]}")
