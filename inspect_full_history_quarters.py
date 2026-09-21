import os
import pathlib
import sys
import pandas as pd

sys.stdout.reconfigure(errors='replace')

ROOT = pathlib.Path('D:/behaviour analysis')
OI_DIR = ROOT / 'OI_DATA'

SYMBOLS = sorted([e.name.strip().upper() for e in os.scandir(OI_DIR) if e.is_dir()])

total_q_count = 0
stock_q_counts = []

for sym in SYMBOLS:
    fr_path = ROOT / sym / 'financial_results.csv'
    n_q = 0
    if fr_path.exists() and fr_path.stat().st_size > 10:
        try:
            df_fr = pd.read_csv(fr_path, dtype=str).fillna('')
            if not df_fr.empty and 'broadCastDate' in df_fr.columns:
                dts = pd.to_datetime(df_fr['broadCastDate'], errors='coerce').dropna()
                n_q = len(dts)
        except Exception:
            pass
    total_q_count += n_q
    stock_q_counts.append((sym, n_q))

print(f"Total Target Universe: {len(SYMBOLS)} Stocks")
print(f"Total Combined Historical Quarters Analyzed Across All 211 Stocks: {total_q_count} Quarters / Trades")

df_counts = pd.DataFrame(stock_q_counts, columns=['Symbol', 'HistoricalQuarters'])
print("\nTop 10 Stocks with Highest Historical Quarters:")
print(df_counts.sort_values('HistoricalQuarters', ascending=False).head(10).to_string(index=False))

print("\nQualified Stocks (>=12 Qtrs) Count:", len(df_counts[df_counts['HistoricalQuarters'] >= 12]))
print("Avoid but Monitor (<12 Qtrs) Count:", len(df_counts[df_counts['HistoricalQuarters'] < 12]))
