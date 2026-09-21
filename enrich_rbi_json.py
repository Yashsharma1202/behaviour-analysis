import json
import pandas as pd
import sys

sys.stdout.reconfigure(encoding='utf-8')

with open('dashboard_data/event_dashboard_data.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

nifty50_df = pd.read_csv('MW-NIFTY-50-21-Jul-2026.csv')
n50_symbols = [s.strip() for s in nifty50_df['SYMBOL'].dropna() if s.strip() != 'NIFTY 50']

df = pd.read_excel('Nifty211_RBI_Policy_Combined_Master.xlsx', skiprows=7)
df_n50 = df[df['Stock Symbol'].isin(n50_symbols)].copy()

for col in ['Win Rate %', 'Average Return %', 'Net Realised P&L (₹)']:
    df_n50[col] = pd.to_numeric(df_n50[col], errors='coerce')

# Trading day calendar for Oct 09-Oct-2026 (T)
oct_trading_days = {
    -8: ("25-Sep-2026", "Fri"),
    -7: ("28-Sep-2026", "Mon"),
    -6: ("29-Sep-2026", "Tue"),
    -5: ("30-Sep-2026", "Wed"),
    -4: ("01-Oct-2026", "Thu"),
    -3: ("05-Oct-2026", "Mon"),
    -2: ("06-Oct-2026", "Tue"),
    -1: ("07-Oct-2026", "Wed"),
     0: ("09-Oct-2026", "Fri"),
    +1: ("12-Oct-2026", "Mon"),
    +2: ("13-Oct-2026", "Tue"),
    +3: ("14-Oct-2026", "Wed"),
    +4: ("15-Oct-2026", "Thu"),
    +5: ("16-Oct-2026", "Fri"),
    +6: ("19-Oct-2026", "Mon"),
    +7: ("21-Oct-2026", "Wed"),
    +8: ("22-Oct-2026", "Thu"),
}

# Trading day calendar for Dec 04-Dec-2026 (T)
dec_trading_days = {
    -8: ("24-Nov-2026", "Tue"),
    -7: ("25-Nov-2026", "Wed"),
    -6: ("26-Nov-2026", "Thu"),
    -5: ("27-Nov-2026", "Fri"),
    -4: ("30-Nov-2026", "Mon"),
    -3: ("01-Dec-2026", "Tue"),
    -2: ("02-Dec-2026", "Wed"),
    -1: ("03-Dec-2026", "Thu"),
     0: ("04-Dec-2026", "Fri"),
    +1: ("07-Dec-2026", "Mon"),
    +2: ("08-Dec-2026", "Tue"),
    +3: ("09-Dec-2026", "Wed"),
    +4: ("10-Dec-2026", "Thu"),
    +5: ("11-Dec-2026", "Fri"),
    +6: ("14-Dec-2026", "Mon"),
    +7: ("15-Dec-2026", "Tue"),
    +8: ("16-Dec-2026", "Wed"),
}

def generate_stock_list(cal_map):
    stocks = []
    df_sorted = df_n50.sort_values(by=['Win Rate %', 'Net Realised P&L (₹)'], ascending=[False, False])
    for _, r in df_sorted.iterrows():
        sym = r['Stock Symbol']
        sec = r['Industry Sector']
        strat = "LONG" if "LONG" in str(r['Strategy']).upper() else "SHORT"
        win = str(r['Position Window']).strip()
        wr = float(r['Win Rate %']) if pd.notna(r['Win Rate %']) else 50.0
        ret = float(r['Average Return %']) if pd.notna(r['Average Return %']) else 0.0
        pnl = float(r['Net Realised P&L (₹)']) if pd.notna(r['Net Realised P&L (₹)']) else 0.0
        
        try:
            parts = win.replace('T', '').split(' to ')
            pre = int(parts[0])
            post = int(parts[1])
            entry_d, entry_w = cal_map.get(pre, ("N/A", ""))
            exit_d, exit_w = cal_map.get(post, ("N/A", ""))
            entry_full = f"{entry_d} ({entry_w})"
            exit_full = f"{exit_d} ({exit_w})"
        except Exception:
            entry_full = "T-2"
            exit_full = "T+1"

        stocks.append({
            "symbol": sym,
            "sector": sec,
            "direction": strat,
            "window": win,
            "entry_date": entry_full,
            "exit_date": exit_full,
            "win_rate": wr,
            "win_ratio": str(r['Win Ratio']),
            "avg_ret": ret,
            "net_pnl": pnl
        })
    return stocks

oct_stocks = generate_stock_list(oct_trading_days)
dec_stocks = generate_stock_list(dec_trading_days)

if 'rbi_policy' in data and len(data['rbi_policy']) >= 2:
    data['rbi_policy'][0]['stocks'] = oct_stocks
    data['rbi_policy'][0]['stocks_count'] = len(oct_stocks)
    data['rbi_policy'][1]['stocks'] = dec_stocks
    data['rbi_policy'][1]['stocks_count'] = len(dec_stocks)

with open('dashboard_data/event_dashboard_data.json', 'w', encoding='utf-8') as f:
    json.dump(data, f, indent=2)

print(f"Successfully updated event_dashboard_data.json with {len(oct_stocks)} Nifty 50 stocks for RBI Policy!")
