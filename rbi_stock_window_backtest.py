"""
RBI Monetary Policy — Individual Nifty 50 Stocks Window Backtest Engine (Fast Vectorized)
Computes T-3 to T+5 pre/post window performance for all 50 Nifty stocks across 162 RBI events (2000-2026).
"""

import pandas as pd
import numpy as np
import json
import os
import sys

# ── Paths ──────────────────────────────────────────────────────────────────────
BASE = r"D:\behaviour analysis"
PARQUET_DIR = os.path.join(BASE, "scraped_parquet")
STOCKS_FILE = os.path.join(PARQUET_DIR, "nifty50_all_stocks_daily_2000_2026.parquet")
RBI_REACTION_FILE = os.path.join(PARQUET_DIR, "rbi_policy_market_reaction_2000_2026.parquet")
JSON_PATH = os.path.join(BASE, "dashboard_data", "rbi_policy_behaviour.json")
DOCS_JSON_PATH = os.path.join(BASE, "docs", "dashboard_data", "rbi_policy_behaviour.json")
EXCEL_PATH = os.path.join(BASE, "RBI_Policy_Behaviour_Master_2000_2026.xlsx")

# ── Sector Mapping for 50 Nifty Stocks ───────────────────────────────────────
SECTORS = {
    'ADANIENT': 'Metals & Mining / Diversified',
    'ADANIPORTS': 'Infrastructure / Logistics',
    'APOLLOHOSP': 'Healthcare / Hospitals',
    'ASIANPAINT': 'Consumer Durables / Paints',
    'AXISBANK': 'Banking & Financials',
    'BAJAJ-AUTO': 'Automobile',
    'BAJAJFINSV': 'Financial Services',
    'BAJFINANCE': 'Financial Services / NBFC',
    'BHARTIARTL': 'Telecommunication',
    'BPCL': 'Oil & Gas',
    'BRITANNIA': 'FMCG',
    'CIPLA': 'Pharmaceuticals',
    'COALINDIA': 'Metals & Mining / Coal',
    'DIVISLAB': 'Pharmaceuticals',
    'DRREDDY': 'Pharmaceuticals',
    'EICHERMOT': 'Automobile',
    'GRASIM': 'Materials / Cement',
    'HCLTECH': 'Information Technology',
    'HDFCBANK': 'Banking & Financials',
    'HDFCLIFE': 'Insurance / Financials',
    'HEROMOTOCO': 'Automobile',
    'HINDALCO': 'Metals & Mining',
    'HINDUNILVR': 'FMCG',
    'ICICIBANK': 'Banking & Financials',
    'INDUSINDBK': 'Banking & Financials',
    'INFY': 'Information Technology',
    'ITC': 'FMCG',
    'JSWSTEEL': 'Metals & Mining',
    'KOTAKBANK': 'Banking & Financials',
    'LT': 'Construction & Engineering',
    'LTIM': 'Information Technology',
    'M&M': 'Automobile',
    'MARUTI': 'Automobile',
    'NESTLEIND': 'FMCG',
    'NTPC': 'Power & Utilities',
    'ONGC': 'Oil & Gas',
    'POWERGRID': 'Power & Utilities',
    'RELIANCE': 'Energy & Telecommunications',
    'SBILIFE': 'Insurance / Financials',
    'SBIN': 'Banking & Financials',
    'SHRIRAMFIN': 'Financial Services / NBFC',
    'SUNPHARMA': 'Pharmaceuticals',
    'TATACONSUM': 'FMCG',
    'TATAMOTORS': 'Automobile',
    'TATASTEEL': 'Metals & Mining',
    'TCS': 'Information Technology',
    'TECHM': 'Information Technology',
    'TITAN': 'Consumer Durables / Retail',
    'ULTRACEMCO': 'Materials / Cement',
    'WIPRO': 'Information Technology'
}

WINDOWS = [
    (-3, +1), (-3, +2), (-3, +3), (-3, +5),
    (-2, +1), (-2, +2), (-2, +3), (-2, +5),
    (-1, +1), (-1, +2), (-1, +3), (-1, +5),
    ( 0, +1), ( 0, +2), ( 0, +3), ( 0, +5),
    (+1, +2), (+1, +3), (+1, +5),
]

def categorize(action):
    a = str(action).upper()
    if 'CUT' in a: return 'CUT'
    elif 'HIKE' in a: return 'HIKE'
    elif 'STATUS QUO' in a or 'NO CHANGE' in a or 'UNCHANGED' in a: return 'STATUS QUO'
    else: return 'OTHER'

print("Loading RBI events and Stock daily data...", flush=True)
rbi_reaction = pd.read_parquet(RBI_REACTION_FILE)
rbi_reaction['POLICY_DATE'] = pd.to_datetime(rbi_reaction['POLICY_DATE'])
rbi_reaction['EFFECTIVE_TRADING_DATE'] = pd.to_datetime(rbi_reaction['EFFECTIVE_TRADING_DATE'])
rbi_reaction['CATEGORY'] = rbi_reaction['ACTION'].apply(categorize)

stocks_df = pd.read_parquet(STOCKS_FILE)
stocks_df['DATE'] = pd.to_datetime(stocks_df['DATE'])
stocks_df = stocks_df.sort_values(['SYMBOL', 'DATE']).reset_index(drop=True)

symbols = sorted(stocks_df['SYMBOL'].unique())
print(f"Loaded {len(symbols)} unique stock symbols.", flush=True)

stock_analysis_results = []

for sym in symbols:
    sub = stocks_df[stocks_df['SYMBOL'] == sym].drop_duplicates(subset=['DATE']).sort_values('DATE').reset_index(drop=True)
    if len(sub) == 0:
        continue
        
    t_days = sub['DATE'].dt.strftime('%Y-%m-%d').values
    opens = sub['OPEN'].values
    closes = sub['CLOSE'].values
    n_days = len(t_days)
    
    date_to_idx = {d: i for i, d in enumerate(t_days)}
    
    sym_trades = []
    events_traded_set = set()
    
    for _, row in rbi_reaction.iterrows():
        eff_str = row['EFFECTIVE_TRADING_DATE'].strftime('%Y-%m-%d')
        cat = row['CATEGORY']
        
        # Find eff_str or nearest next trading day
        if eff_str in date_to_idx:
            idx0 = date_to_idx[eff_str]
        else:
            # find next available day in stock data
            future_days = [d for d in t_days if d >= eff_str]
            if not future_days:
                continue
            idx0 = date_to_idx[future_days[0]]
            
        events_traded_set.add(row['POLICY_DATE'].strftime('%Y-%m-%d'))
        
        for (entry_off, exit_off) in WINDOWS:
            e_idx = idx0 + entry_off
            x_idx = idx0 + exit_off
            
            if 0 <= e_idx < n_days and 0 <= x_idx < n_days:
                px_e = opens[e_idx] if entry_off < 0 else closes[e_idx]
                px_x = closes[x_idx]
                
                if px_e > 0 and not np.isnan(px_e) and not np.isnan(px_x):
                    ret = ((px_x - px_e) / px_e) * 100.0
                    sym_trades.append({
                        'CATEGORY': cat,
                        'WINDOW': f"T{entry_off:+d} to T{exit_off:+d}",
                        'RETURN_PCT': ret
                    })
                    
    if not sym_trades:
        continue
        
    tdf = pd.DataFrame(sym_trades)
    
    opt_dict = {}
    for cat in ['ALL', 'CUT', 'HIKE', 'STATUS QUO']:
        if cat == 'ALL':
            cat_df = tdf
        else:
            cat_df = tdf[tdf['CATEGORY'] == cat]
            
        if len(cat_df) == 0:
            opt_dict[cat] = {'window': 'N/A', 'win_rate': 0.0, 'avg_return': 0.0, 'count': 0, 'score': 0.0}
            continue
            
        w_stats = []
        for w_name, group in cat_df.groupby('WINDOW'):
            rets = group['RETURN_PCT'].values
            n_tr = len(rets)
            if n_tr < 3:
                continue
            wins = np.sum(rets > 0)
            wr = round(float(wins / n_tr * 100.0), 1)
            avg_r = round(float(np.mean(rets)), 2)
            med_r = round(float(np.median(rets)), 2)
            score = round(float(wr * avg_r / 10.0), 2)
            w_stats.append({
                'window': w_name,
                'win_rate': wr,
                'avg_return': avg_r,
                'median_return': med_r,
                'count': int(n_tr),
                'score': score
            })
            
        if not w_stats:
            opt_dict[cat] = {'window': 'N/A', 'win_rate': 0.0, 'avg_return': 0.0, 'count': 0, 'score': 0.0}
        else:
            # Sort by wr >= 55%, wr desc, avg_return desc
            w_stats.sort(key=lambda x: (x['win_rate'] >= 55.0, x['win_rate'], x['avg_return']), reverse=True)
            best = w_stats[0]
            opt_dict[cat] = best

    stock_analysis_results.append({
        'symbol': sym,
        'sector': SECTORS.get(sym, 'Other'),
        'total_events': len(events_traded_set),
        'optimal_all': opt_dict['ALL'],
        'optimal_cut': opt_dict['CUT'],
        'optimal_hike': opt_dict['HIKE'],
        'optimal_status_quo': opt_dict['STATUS QUO']
    })

print(f"Fast analysis completed for {len(stock_analysis_results)} stocks!", flush=True)

# ── Update JSON File ──────────────────────────────────────────────────────────
with open(JSON_PATH, 'r', encoding='utf-8') as f:
    rbi_json_data = json.load(f)

rbi_json_data['stock_analysis'] = stock_analysis_results

with open(JSON_PATH, 'w', encoding='utf-8') as f:
    json.dump(rbi_json_data, f, indent=2)

os.makedirs(os.path.dirname(DOCS_JSON_PATH), exist_ok=True)
with open(DOCS_JSON_PATH, 'w', encoding='utf-8') as f:
    json.dump(rbi_json_data, f, indent=2)

print(f"Updated JSON at {JSON_PATH} and {DOCS_JSON_PATH}", flush=True)

# ── Update Excel File ─────────────────────────────────────────────────────────
try:
    rows = []
    for s in stock_analysis_results:
        rows.append({
            'Symbol': s['symbol'],
            'Sector': s['sector'],
            'Events Traded': s['total_events'],
            'Best Window (All)': s['optimal_all']['window'],
            'Win Rate % (All)': s['optimal_all']['win_rate'],
            'Avg Ret % (All)': s['optimal_all']['avg_return'],
            'Best Window (Cut)': s['optimal_cut']['window'],
            'Win Rate % (Cut)': s['optimal_cut']['win_rate'],
            'Avg Ret % (Cut)': s['optimal_cut']['avg_return'],
            'Best Window (Hike)': s['optimal_hike']['window'],
            'Win Rate % (Hike)': s['optimal_hike']['win_rate'],
            'Avg Ret % (Hike)': s['optimal_hike']['avg_return'],
            'Best Window (Status Quo)': s['optimal_status_quo']['window'],
            'Win Rate % (Status Quo)': s['optimal_status_quo']['win_rate'],
            'Avg Ret % (Status Quo)': s['optimal_status_quo']['avg_return'],
        })
    stock_sum_df = pd.DataFrame(rows)
    with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        stock_sum_df.to_excel(writer, sheet_name='Nifty50_Stocks_RBI_Summary', index=False)
    print("Excel updated successfully with sheet 'Nifty50_Stocks_RBI_Summary'", flush=True)
except Exception as e:
    print(f"Excel update note: {e}", flush=True)

print("Done! RBI Stock Backtest complete.", flush=True)
