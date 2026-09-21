import pandas as pd, sys

path = r'D:/behaviour analysis/Nifty50_12_Quarters_Options_OI_Master_v4.xlsx'

try:
    # Read the first sheet (assuming the master data is there)
    df = pd.read_excel(path, sheet_name=0, engine='openpyxl')
except Exception as e:
    print(f'❌ Failed to read Excel: {e}')
    sys.exit(1)

# Print column names
print('--- Columns in Options Master ---')
for col in df.columns:
    print(col)

print('\n--- First 5 rows (raw) ---')
try:
    rows_str = df.head(5).to_string(index=False)
    print(rows_str.encode('utf-8', errors='ignore').decode('utf-8'))
except Exception as e:
    print(f'⚠️ Could not display rows: {e}')

print(f"\nTotal rows: {len(df)}")
