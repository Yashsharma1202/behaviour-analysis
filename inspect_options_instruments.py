import pandas as pd, sys, glob

sys.stdout.reconfigure(errors='replace')

files = glob.glob('d:/behaviour analysis/options_parquet/*.parquet')
sample_file = files[-100]

print(f"Reading file: {sample_file}")
df = pd.read_parquet(sample_file)

print("\nUnique Instruments found:")
print(df['Instrument'].unique().tolist()[:30])

print("\nUnique Tickers sample:")
print(df['Ticker'].unique().tolist()[:20])
