import pandas as pd, sys

path = r'D:/behaviour analysis/Futures_Master_Only.xlsx'

try:
    df = pd.read_excel(path, engine='openpyxl')
except Exception as e:
    print(f'❌ Failed to read futures master: {e}')
    sys.exit(1)

print('--- Futures Master Columns (repr) ---')
for i, col in enumerate(df.columns):
    print(i, repr(col))
print('\nTotal rows:', len(df))
