import pandas as pd
import sys

sys.stdout.reconfigure(encoding='utf-8')

nifty50_df = pd.read_csv('MW-NIFTY-50-21-Jul-2026.csv')
n50_symbols = [s.strip() for s in nifty50_df['SYMBOL'].dropna() if s.strip() != 'NIFTY 50']

df = pd.read_excel('Nifty211_RBI_Policy_Combined_Master.xlsx', skiprows=7)
df_n50 = df[df['Stock Symbol'].isin(n50_symbols)].copy()

for col in ['Win Rate %', 'Average Return %', 'Net Realised P&L (₹)']:
    df_n50[col] = pd.to_numeric(df_n50[col], errors='coerce')

# Map Upcoming Dates for Oct 2026 MPC
# Announcement Date: Friday, 09-Oct-2026 (T)
# Holiday on 02-Oct (Gandhi Jayanti) and 20-Oct (Dussehra)
trading_days = {
    -8: ("25-Sep-2026", "Fri"),
    -7: ("28-Sep-2026", "Mon"),
    -6: ("29-Sep-2026", "Tue"),
    -5: ("30-Sep-2026", "Wed"),
    -4: ("01-Oct-2026", "Thu"),
    # 02-Oct is holiday
    -3: ("05-Oct-2026", "Mon"),
    -2: ("06-Oct-2026", "Tue"),
    -1: ("07-Oct-2026", "Wed"), # MPC Meeting starts
     0: ("09-Oct-2026", "Fri"), # Policy Announcement at 10 AM
    +1: ("12-Oct-2026", "Mon"),
    +2: ("13-Oct-2026", "Tue"),
    +3: ("14-Oct-2026", "Wed"),
    +4: ("15-Oct-2026", "Thu"),
    +5: ("16-Oct-2026", "Fri"),
    +6: ("19-Oct-2026", "Mon"),
    # 20-Oct is Dussehra
    +7: ("21-Oct-2026", "Wed"),
    +8: ("22-Oct-2026", "Thu"),
}

# Sort by Sector, then Win Rate descending
df_n50 = df_n50.sort_values(by=['Industry Sector', 'Win Rate %'], ascending=[True, False])

print("Symbol | Sector | Strategy | Window | Entry (Oct 26) | Exit (Oct 26) | WinRate | WinRatio | AvgReturn | NetPnL")
print("-" * 120)

for _, r in df_n50.iterrows():
    sym = r['Stock Symbol']
    sec = r['Industry Sector']
    strat = r['Strategy']
    win = str(r['Position Window'])
    wr = r['Win Rate %']
    ratio = r['Win Ratio']
    ret = r['Average Return %']
    pnl = r['Net Realised P&L (₹)']
    
    # Parse window: e.g. T-8 to T+2
    try:
        parts = win.replace('T', '').split(' to ')
        pre = int(parts[0])
        post = int(parts[1])
        entry_date, entry_day = trading_days.get(pre, ("N/A", ""))
        exit_date, exit_day = trading_days.get(post, ("N/A", ""))
        entry_str = f"{entry_date} ({entry_day})"
        exit_str = f"{exit_date} ({exit_day})"
    except Exception:
        entry_str = "N/A"
        exit_str = "N/A"

    strat_icon = "🟢 LONG" if "LONG" in strat else "🔴 SHORT"
    print(f"{sym:12} | {sec:28} | {strat_icon:8} | {win:10} | {entry_str:18} | {exit_str:18} | {wr:5.1f}% | {ratio:18} | {ret:6.2f}% | ₹{pnl:10,.0f}")
