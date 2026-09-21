import pandas as pd, sys, glob

sys.stdout.reconfigure(errors='replace')

files = glob.glob('d:/behaviour analysis/options_parquet/*.parquet')
print(f"Total Parquet Files Found: {len(files)}")

if files:
    sample_file = files[-10]
    print(f"\nInspecting Sample Parquet File: {sample_file}")
    df_sample = pd.read_parquet(sample_file)
    print("Columns:")
    print(df_sample.columns.tolist())
    print("\nShape:", df_sample.shape)
    print("\nHead 10 rows:")
    print(df_sample.head(10).to_string())
