import os
import pathlib
import sys
import pandas as pd

sys.stdout.reconfigure(errors='replace')

ROOT = pathlib.Path('D:/behaviour analysis')
OI_DIR = ROOT / 'OI_DATA'
PROC = ROOT / 'processed'
PRICE_CACHE = PROC / 'price_cache'

SYMBOLS = sorted([e.name.strip().upper() for e in os.scandir(OI_DIR) if e.is_dir()])

LOT_SIZES = {
    "RELIANCE": 250, "TCS": 175, "INFY": 400, "HDFCBANK": 550, "ICICIBANK": 700,
    "BHARTIARTL": 950, "ITC": 1600, "SBIN": 1500, "LTIM": 150, "LT": 300,
    "HINDUNILVR": 300, "AXISBANK": 625, "KOTAKBANK": 400, "BAJFINANCE": 125,
    "M&M": 350, "MARUTI": 100, "SUNPHARMA": 350, "TATASTEEL": 5500,
    "NTPC": 1500, "POWERGRID": 1800, "TITAN": 175, "ADANIENT": 300,
    "ADANIPORTS": 625, "ULTRACEMCO": 100, "ASIANPAINT": 200, "COALINDIA": 2100,
    "BAJAJ-AUTO": 125, "JSWSTEEL": 675, "TATAMOTORS": 1425, "HCLTECH": 350,
    "GRASIM": 250, "HEROMOTOCO": 150, "EICHERMOT": 175, "CIPLA": 650,
    "HDFCLIFE": 1100, "SBILIFE": 375, "DRREDDY": 125, "BRITANNIA": 200,
    "APOLLOHOSP": 125, "TATACONSUM": 450, "HINDALCO": 1400, "BPCL": 1800,
    "INDUSINDBK": 500, "DIVISLAB": 200, "BAJAJFINSV": 500, "NESTLEIND": 200,
    "WIPRO": 1500, "ONGC": 3750, "TECHM": 600, "ASHOKLEY": 5000, "ADANIGREEN": 500
}

fund_list = []
total_margin = 0.0
total_notional = 0.0

for sym in SYMBOLS:
    stock_folder = ROOT / sym
    fr_path = stock_folder / 'financial_results.csv'
    price_path = PRICE_CACHE / f"{sym}.csv"
    
    n_quarters = 0
    if fr_path.exists() and fr_path.stat().st_size > 10:
        try:
            df_fr = pd.read_csv(fr_path, dtype=str).fillna('')
            if not df_fr.empty and 'broadCastDate' in df_fr.columns:
                dts = pd.to_datetime(df_fr['broadCastDate'], errors='coerce').dropna()
                n_quarters = len(dts)
        except Exception:
            pass
            
    if n_quarters < 12:
        continue  # Filter ONLY QUALIFIED STOCKS
        
    if not price_path.exists():
        continue
    df_p = pd.read_csv(price_path)
    if df_p.empty:
        continue
    last_p = float(df_p['adj'].iloc[-1])
    lot_size = LOT_SIZES.get(sym, max(100, int(1000000 / last_p)))
    notional_val = round(lot_size * last_p, 2)
    margin_req = round(0.20 * notional_val, 2)
    
    total_notional += notional_val
    total_margin += margin_req
    fund_list.append({
        "sym": sym, "n_quarters": n_quarters, "lot_size": lot_size,
        "price": last_p, "notional": notional_val, "margin": margin_req
    })

df_fund = pd.DataFrame(fund_list)
df_fund = df_fund.sort_values('margin', ascending=False)

print(f"Total Qualified Target Symbols: {len(df_fund)} Stocks")
print(f"Total Combined Notional Value (Qualified 1 Lot Each): ₹{total_notional:,.2f}")
print(f"Total Portfolio Margin Required (20% Cash Used for Qualified): ₹{total_margin:,.2f}")
print(f"Average Margin Required Per Qualified Stock: ₹{df_fund['margin'].mean():,.2f}")

print("\nTOP 10 HIGHEST MARGIN USED QUALIFIED STOCKS:")
for r in df_fund.head(10).to_dict('records'):
    print(f"  • {r['sym']:<12}: Lot Size = {r['lot_size']:<5} | Stock Price = ₹{r['price']:<8.2f} | Margin Used = ₹{r['margin']:,.2f}")
