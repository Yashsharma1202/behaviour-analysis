import pandas as pd, sys

sys.stdout.reconfigure(errors='replace')

BASE_DIR = pathlib.Path('D:/behaviour analysis') if 'pathlib' in locals() else __import__('pathlib').Path('D:/behaviour analysis')
EQ_V2   = BASE_DIR / 'Nifty50_12_Quarters_Equity_Master_v2.xlsx'

xl_eq = pd.ExcelFile(EQ_V2, engine='openpyxl')
df_eq = xl_eq.parse('All_12Q_Equity_Trades')

print("Equity Master Columns:")
print(df_eq.columns.tolist())
