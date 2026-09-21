import pandas as pd, pathlib, sys

sys.stdout.reconfigure(errors='replace')

workspace = pathlib.Path('d:/behaviour analysis')
files = sorted(workspace.glob('*_Combined_Best_Capital_Utilisation.xlsx'))

print(f"Found {len(files)} quarter files:")
for f in files:
    print(f"\n--- File: {f.name} ---")
    try:
        xl = pd.ExcelFile(f, engine='openpyxl')
        for sheet in xl.sheet_names:
            df = xl.parse(sheet, nrows=2)
            print(f"  Sheet: {sheet}, columns: {list(df.columns)}")
    except Exception as e:
        print(f"  Error reading {f.name}: {e}")
