import pathlib, os, json, pandas as pd, numpy as np, sys
sys.stdout.reconfigure(encoding='utf-8')

ROOT = pathlib.Path('D:/behaviour analysis')
PROC = ROOT / 'processed'
PRICE_CACHE = PROC / 'price_cache'

NIFTY50 = [
    'ULTRACEMCO', 'RELIANCE', 'TCS', 'INFY', 'HDFCBANK', 'ICICIBANK', 'BHARTIARTL', 'ITC', 'SBIN', 'LT',
    'AXISBANK', 'KOTAKBANK', 'BAJFINANCE', 'MARUTI', 'SUNPHARMA', 'TATASTEEL', 'NTPC', 'POWERGRID', 'TITAN', 'ADANIENT',
    'ADANIPORTS', 'ASIANPAINT', 'COALINDIA', 'BAJAJ-AUTO', 'JSWSTEEL', 'TATAMOTORS', 'HCLTECH', 'GRASIM', 'HEROMOTOCO', 'EICHERMOT',
    'CIPLA', 'HDFCLIFE', 'SBILIFE', 'DRREDDY', 'BRITANNIA', 'APOLLOHOSP', 'TATACONSUM', 'HINDALCO', 'BPCL', 'INDUSINDBK',
    'DIVISLAB', 'BAJAJFINSV', 'NESTLEIND', 'WIPRO', 'ONGC', 'TECHM', 'ASHOKLEY', 'ADANIGREEN', 'SHRIRAMFIN', 'HINDUNILVR'
]

LOT_SIZES = {
    'RELIANCE': 250, 'TCS': 175, 'INFY': 400, 'HDFCBANK': 550, 'ICICIBANK': 700,
    'BHARTIARTL': 950, 'ITC': 1600, 'SBIN': 1500, 'LTIM': 150, 'LT': 300,
    'HINDUNILVR': 300, 'AXISBANK': 625, 'KOTAKBANK': 400, 'BAJFINANCE': 125,
    'M&M': 350, 'MARUTI': 100, 'SUNPHARMA': 350, 'TATASTEEL': 5500,
    'NTPC': 1500, 'POWERGRID': 1800, 'TITAN': 175, 'ADANIENT': 300,
    'ADANIPORTS': 625, 'ULTRACEMCO': 100, 'ASIANPAINT': 200, 'COALINDIA': 2100,
    'BAJAJ-AUTO': 125, 'JSWSTEEL': 675, 'TATAMOTORS': 1425, 'HCLTECH': 350,
    'GRASIM': 250, 'HEROMOTOCO': 150, 'EICHERMOT': 175, 'CIPLA': 650,
    'HDFCLIFE': 1100, 'SBILIFE': 375, 'DRREDDY': 125, 'BRITANNIA': 200,
    'APOLLOHOSP': 125, 'TATACONSUM': 450, 'HINDALCO': 1400, 'BPCL': 1800,
    'INDUSINDBK': 500, 'DIVISLAB': 200, 'BAJAJFINSV': 500, 'NESTLEIND': 200,
    'WIPRO': 1500, 'ONGC': 3750, 'TECHM': 600, 'ASHOKLEY': 5000, 'ADANIGREEN': 500,
    'SHRIRAMFIN': 300, 'TRENT': 100
}

pre_candidates = [1, 2, 3, 4, 5, 7, 8]
post_candidates = [1, 2, 3, 4, 5, 8]

q3_results = []

for sym in NIFTY50:
    fr_path = ROOT / sym / 'financial_results.csv'
    price_path = PRICE_CACHE / f'{sym}.csv'
    
    if not price_path.exists(): continue
    
    df_p = pd.read_csv(price_path)
    df_p['date'] = pd.to_datetime(df_p['date'])
    df_p = df_p.sort_values('date').reset_index(drop=True)
    valid_dates = df_p['date'].tolist()
    if not valid_dates: continue
    
    last_p = float(df_p['adj'].iloc[-1])
    lot_size = LOT_SIZES.get(sym, max(100, int(1000000 / last_p)))
    margin_req = round(0.20 * lot_size * last_p, 2)
    
    clean_dts = []
    if fr_path.exists() and fr_path.stat().st_size > 10:
        try:
            df_fr = pd.read_csv(fr_path)
            raw_dts = pd.to_datetime(df_fr['broadCastDate'], errors='coerce').dropna().sort_values(ascending=False).tolist()
            for d in raw_dts:
                if not clean_dts or (clean_dts[-1] - d).days > 15:
                    clean_dts.append(d)
        except Exception:
            pass
            
    oct_nov_dts = [d for d in clean_dts if d.month in (10, 11)]
    if len(oct_nov_dts) == 0: oct_nov_dts = clean_dts[:12]
    tot_count = len(oct_nov_dts)
    
    best_wr = -1.0; best_win_str = 'T-1 to T+1'; best_strat = 'FUTURE LONG'; best_avg_ret = 0.0
    best_pre = 1; best_post = 1; best_wins = 0
    
    for pre in pre_candidates:
        for post in post_candidates:
            emp_l = 0; emp_s = 0
            rets_l = []; rets_s = []
            for dt in oct_nov_dts:
                r_i = min(range(len(valid_dates)), key=lambda i: abs((valid_dates[i] - dt).days))
                e_i = max(0, r_i - pre)
                x_i = min(len(valid_dates) - 1, r_i + post)
                pe = float(df_p.loc[e_i, 'adj']); px = float(df_p.loc[x_i, 'adj'])
                if pe > 0:
                    rl = (px - pe) / pe * 100
                    rs = (pe - px) / pe * 100
                    rets_l.append(rl); rets_s.append(rs)
                    if px > pe: emp_l += 1
                    elif pe > px: emp_s += 1
            
            wr_l = (emp_l / tot_count) * 100 if tot_count > 0 else 50.0
            wr_s = (emp_s / tot_count) * 100 if tot_count > 0 else 50.0
            
            avg_l = np.mean(rets_l) if rets_l else 0.0
            avg_s = np.mean(rets_s) if rets_s else 0.0
            
            if wr_l >= wr_s:
                if wr_l > best_wr or (abs(wr_l - best_wr) < 1e-5 and avg_l > best_avg_ret):
                    best_wr = wr_l; best_win_str = f'T-{pre} to T+{post}'; best_strat = 'FUTURE LONG'; best_avg_ret = avg_l; best_pre = pre; best_post = post; best_wins = emp_l
            else:
                if wr_s > best_wr or (abs(wr_s - best_wr) < 1e-5 and avg_s > best_avg_ret):
                    best_wr = wr_s; best_win_str = f'T-{pre} to T+{post}'; best_strat = 'FUTURE SHORT'; best_avg_ret = avg_s; best_pre = pre; best_post = post; best_wins = emp_s

    stage = 2 if sym == 'ULTRACEMCO' else 1
    date_str = '19-Oct-2026 (Monday)' if sym == 'ULTRACEMCO' else 'Unannounced (Pending NSE Filing)'
    
    past_trades = []
    for q_i, q_dt in enumerate(oct_nov_dts[:4], 1):
        r_i = min(range(len(valid_dates)), key=lambda i: abs((valid_dates[i] - q_dt).days))
        e_i = max(0, r_i - best_pre)
        x_i = min(len(valid_dates) - 1, r_i + best_post)
        pe = float(df_p.loc[e_i, 'adj']); px = float(df_p.loc[x_i, 'adj'])
        
        r_val = (px - pe) / pe * 100 if best_strat == 'FUTURE LONG' else (pe - px) / pe * 100
        pnl_val = round(lot_size * (px - pe) if best_strat == 'FUTURE LONG' else lot_size * (pe - px), 2)
        
        past_trades.append({
            'q': f"Q3 {q_dt.year}",
            't': q_dt.strftime('%Y-%m-%d'),
            'entry': valid_dates[e_i].strftime('%Y-%m-%d'),
            'exit': valid_dates[x_i].strftime('%Y-%m-%d'),
            'pEntry': round(pe, 2),
            'pExit': round(px, 2),
            'ret': round(r_val, 2),
            'pnl': pnl_val,
            'win': pnl_val > 0,
            'side': best_strat.replace('FUTURE ', '')
        })

    q3_results.append({
        'sym': sym,
        'name': sym + ' Ltd.',
        'ltp': round(last_p, 2),
        'lot': lot_size,
        'margin': margin_req,
        'stage': stage,
        'dateStr': date_str,
        'winStr': best_win_str,
        'pre': best_pre,
        'post': best_post,
        'strat': best_strat.replace('FUTURE ', ''),
        'opt_strat': best_strat,
        'wr': f"{round(best_wr, 2)}%",
        'raw_wr': round(best_wr, 2),
        'wins': best_wins,
        'tot_quarters': tot_count,
        'avgRet': f"+{round(best_avg_ret, 2)}%" if best_avg_ret >= 0 else f"{round(best_avg_ret, 2)}%",
        'entry': "16-Oct-2026 (Friday)" if sym == 'ULTRACEMCO' else "Pending NSE Filing",
        'result': "19-Oct-2026 (Monday)" if sym == 'ULTRACEMCO' else "Pending NSE Filing",
        'exit': "20-Oct-2026 (Tuesday)" if sym == 'ULTRACEMCO' else "Pending NSE Filing",
        'trades': past_trades
    })

print(f"Generated Q3 Windows for {len(q3_results)} Nifty 50 stocks.")

# Write Excel
excel_path = ROOT / "Nifty50_Q3_Position_Execution_Windows_Master.xlsx"
df_excel = pd.DataFrame([
    {
        'Symbol': r['sym'],
        'Company Name': r['name'],
        'Filing Stage': 'Stage 2 (Confirmed)' if r['stage'] == 2 else 'Stage 1 (Pending)',
        'Result Date': r['dateStr'],
        'Optimal Window': r['winStr'],
        'Optimal Strategy': r['opt_strat'],
        'Q3 Win Rate': r['wr'],
        'Expected Gain': r['avgRet'],
        'LTP (₹)': r['ltp'],
        'Lot Size': r['lot'],
        '20% Futures Margin (₹)': r['margin']
    }
    for r in q3_results
])
df_excel.to_excel(excel_path, index=False)
print("Saved Excel:", excel_path.name)

# Write JSON
json_path = ROOT / "Nifty50_Q3_Position_Execution_Windows_Master.json"
with open(json_path, 'w', encoding='utf-8') as f:
    json.dump(q3_results, f, indent=2)
print("Saved JSON:", json_path.name)
