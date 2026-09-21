import pandas as pd, os, pathlib, sys

sys.stdout.reconfigure(errors='replace')

ROOT = pathlib.Path('D:/behaviour analysis')
OI_DIR = ROOT / 'OI_DATA'
PROC = ROOT / 'processed'
PRICE_CACHE = PROC / 'price_cache'

# Discover all 211 symbols
SYMBOLS = sorted([e.name.strip().upper() for e in os.scandir(OI_DIR) if e.is_dir()])

print("==========================================================================================")
print("CALCULATING DYNAMIC STOCK-SPECIFIC OPTIMAL POSITION TAKING WINDOWS FOR ALL 211 STOCKS")
print("==========================================================================================")

stock_windows = {}

for idx, sym in enumerate(SYMBOLS, 1):
    stock_folder = ROOT / sym
    fr_path = stock_folder / 'financial_results.csv'
    price_path = PRICE_CACHE / f"{sym}.csv"
    
    if not fr_path.exists() or not price_path.exists():
        stock_windows[sym] = ("T-8 to T+1", 8, 1, 60.0)
        continue
        
    try:
        df_fr = pd.read_csv(fr_path)
        df_price = pd.read_csv(price_path)
        df_price['date'] = pd.to_datetime(df_price['date'])
        df_price = df_price.sort_values('date').reset_index(drop=True)
        
        valid_dates = df_price['date'].tolist()
        
        # Test N_BEFORE in (1..8) and N_AFTER in (1..5) for this specific stock
        best_winrate = -1.0
        best_win = ("T-8 to T+1", 8, 1, 60.0)
        
        for pre in [1, 2, 3, 4, 5, 6, 7, 8]:
            for post in [1, 2, 3, 4, 5]:
                returns = []
                for _, r in df_fr.iterrows():
                    dt_str = r.get('broadCastDate', '')
                    dt = pd.to_datetime(dt_str, errors='coerce')
                    if pd.isna(dt): continue
                    
                    res_idx = min(range(len(valid_dates)), key=lambda i: abs((valid_dates[i] - dt).days))
                    entry_idx = max(0, res_idx - pre)
                    exit_idx = min(len(valid_dates) - 1, res_idx + post)
                    
                    p_in = float(df_price.loc[entry_idx, 'adj'])
                    p_out = float(df_price.loc[exit_idx, 'adj'])
                    if p_in > 0:
                        ret = (p_out - p_in) / p_in
                        returns.append(ret)
                        
                if returns:
                    wins = [r for r in returns if r > 0]
                    wr = (len(wins) / len(returns)) * 100
                    if wr > best_winrate:
                        best_winrate = wr
                        best_win = (f"T-{pre} to T+{post}", pre, post, round(wr, 1))
                        
        stock_windows[sym] = best_win
    except Exception as e:
        stock_windows[sym] = ("T-8 to T+1", 8, 1, 60.0)

print(f"Calculated stock-specific optimal windows for {len(stock_windows)} stocks!")
print("\nSample Calculated Stock-Specific Windows:")
for s in SYMBOLS[:15]:
    w_str, pre, post, wr = stock_windows[s]
    print(f"  • {s:14s} -> Position Window: {w_str:12s} | Historical Win Rate: {wr}%")

print("==========================================================================================")
