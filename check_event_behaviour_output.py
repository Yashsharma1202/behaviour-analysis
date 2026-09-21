import pandas as pd, pathlib, sys

sys.stdout.reconfigure(errors='replace')

p_summary = pathlib.Path('D:/behaviour analysis/processed/event_behaviour_summary.csv')

print("==========================================================================================")
print("VERIFYING EVENT BEHAVIOUR SUMMARY OUTPUT")
print("==========================================================================================")

if p_summary.exists():
    df = pd.read_csv(p_summary)
    print(f"✅ event_behaviour_summary.csv exists! Total rows: {len(df)}")
    print(f"Columns: {list(df.columns)}")
    if not df.empty:
        print("\nSample Rows:")
        print(df.head(10))
else:
    print("❌ event_behaviour_summary.csv does not exist yet.")

print("==========================================================================================")
