import pandas as pd, sys

# Prevent Windows console cp1252 encoding errors
sys.stdout.reconfigure(errors='replace')

path = r'D:/behaviour analysis/Futures_Master_Only.xlsx'

try:
    df = pd.read_excel(path, engine='openpyxl')
except Exception as e:
    print(f'Failed to read futures master: {e}')
    sys.exit(1)

print('--- Futures Master Columns ---')
for i, col in enumerate(df.columns):
    print(f"{i}: {col}")
print('\nTotal rows:', len(df))
print('\nSample head:')
print(df.head(3))
