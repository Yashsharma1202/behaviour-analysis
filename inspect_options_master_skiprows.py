import pandas as pd, sys

path = r'D:/behaviour analysis/Nifty50_12_Quarters_Options_OI_Master_v4.xlsx'

for skip in range(2, 6):  # try skiprows 2,3,4,5
    try:
        df = pd.read_excel(path, sheet_name=0, skiprows=skip, engine='openpyxl')
        print(f'--- skiprows={skip} ---')
        print('Columns (first 10):', df.columns.tolist()[:10])
        print('Rows loaded:', len(df))
        if not df.empty:
            print(df.head(2).to_string(index=False))
        else:
            print('Dataframe empty')
        print('\n')
    except Exception as e:
        print(f'Error with skiprows={skip}:', e)
        sys.exit(1)
