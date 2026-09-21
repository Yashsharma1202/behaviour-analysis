import pandas as pd
import sys

excel_path = r'D:/behaviour analysis/Nifty50_12_Quarters_Options_OI_Master_v4.xlsx'

try:
    # Load the first sheet (assuming the master data is in the first sheet)
    df = pd.read_excel(excel_path, sheet_name=0)
except Exception as e:
    print(f'❌ Failed to read options master Excel: {e}')
    sys.exit(1)



print(f"\nRows total: {len(df)}")
