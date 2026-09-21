import pandas as pd, sys

path = r'D:/behaviour analysis/Nifty50_12_Quarters_Options_OI_Master_v4.xlsx'

# Try reading with first row as header (default) and also with second row as header to see actual column names
for header_row in [0, 1]:
    try:
        df = pd.read_excel(path, sheet_name=0, header=header_row, engine='openpyxl')
        print(f'--- Reading with header={header_row} ---')
        print('Columns:')
        for i, col in enumerate(df.columns):
            print(i, repr(col))
        print('Rows loaded:', len(df))
        # Show first few rows of data
        if not df.empty:
            print(df.head(3).to_string(index=False))
    except Exception as e:
        print(f'Error reading with header={header_row}:', e)
        sys.exit(1)
