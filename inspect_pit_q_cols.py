import pandas as pd, sys

sys.stdout.reconfigure(errors='replace')

BASE_DIR = pathlib.Path('D:/behaviour analysis') if 'pathlib' in locals() else __import__('pathlib').Path('D:/behaviour analysis')
FUT_PIT = BASE_DIR / 'Nifty50_12_Quarters_Point_In_Time_Futures_Master.xlsx'

xl_pit = pd.ExcelFile(FUT_PIT, engine='openpyxl')
df_q3 = xl_pit.parse('Q3 2023-24')

print("Q3 2023-24 Columns:")
print(df_q3.columns.tolist())
print("\nRow 1-8 sample:")
print(df_q3.head(8).to_string())
