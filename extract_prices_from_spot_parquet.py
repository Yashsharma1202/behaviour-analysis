import os
import pathlib
import sys
import pandas as pd
from concurrent.futures import ThreadPoolExecutor

sys.stdout.reconfigure(errors='replace')

ROOT = pathlib.Path('D:/behaviour analysis')
OI_DIR = ROOT / 'OI_DATA'
PROC = ROOT / 'processed'
PRICE_CACHE = PROC / 'price_cache'
PRICE_CACHE.mkdir(parents=True, exist_ok=True)

# Discover all 211 symbols
SYMBOLS = sorted([e.name.strip().upper() for e in os.scandir(OI_DIR) if e.is_dir()])

print("==========================================================================================")
print("EXTRACTING DAILY PRICES FROM SPOT PARQUET FOR ALL 211 STOCKS")
print("==========================================================================================")
print(f"Total Target Symbols: {len(SYMBOLS)} Stocks")

def process_symbol_prices(sym):
    cache_file = PRICE_CACHE / f"{sym}.csv"
    if cache_file.exists() and cache_file.stat().st_size > 100:
        return sym, 0, True
        
    spot_dir = OI_DIR / sym / "spot_parquet"
    if not spot_dir.exists():
        return sym, 0, False
        
    p_files = sorted([spot_dir / f for f in os.listdir(spot_dir) if f.endswith(".parquet")])
    if not p_files:
        return sym, 0, False
        
    daily_rows = []
    for pf in p_files:
        date_str = pf.stem
        try:
            df = pd.read_parquet(pf)
            if not df.empty and "Close" in df.columns:
                last_close = float(df["Close"].iloc[-1])
                daily_rows.append({"date": date_str, "adj": last_close})
        except Exception:
            pass
            
    if daily_rows:
        df_daily = pd.DataFrame(daily_rows)
        df_daily.to_csv(cache_file, index=False)
        return sym, len(df_daily), False
        
    return sym, 0, False

print("Starting ultra-fast extraction...")

done_count = 0
with ThreadPoolExecutor(max_workers=16) as executor:
    results = list(executor.map(process_symbol_prices, SYMBOLS))
    
for sym, n_days, cached in results:
    status = "CACHED" if cached else "EXTRACTED"
    if n_days > 0 or cached:
        done_count += 1
    print(f"  • {sym:14s} ({status:10s}) -> {n_days:4d} daily price bars")

print("==========================================================================================")
print(f"COMPLETED PRICE EXTRACTION FOR {done_count}/{len(SYMBOLS)} STOCKS!")
print("==========================================================================================")
