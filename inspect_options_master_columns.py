import pandas as pd, sys

path = r'D:/behaviour analysis/Nifty50_12_Quarters_Options_OI_Master_v4.xlsx'

try:
    df = pd.read_excel(path, sheet_name=0, engine='openpyxl')
except Exception as e:
    print(f'❌ Failed to read options master: {e}')
    sys.exit(1)

print('--- Column list (index, repr) ---')
for i, col in enumerate(df.columns):
    # Use utf-8 safe representation
    try:
        safe = str(col)
    except Exception:
        safe = repr(col)
    print(i, safe)

print('\nTotal columns:', len(df.columns))
