import pandas as pd
import sys

from rbi_mpc_calendar import refresh_from_rbi, get_next_meeting, get_most_recent_meeting
from nse_trading_calendar import load_trading_holidays, trading_day_window

sys.stdout.reconfigure(encoding='utf-8')

nifty50_df = pd.read_csv('MW-NIFTY-50-21-Jul-2026.csv')
n50_symbols = [s.strip() for s in nifty50_df['SYMBOL'].dropna() if s.strip() != 'NIFTY 50']

df = pd.read_excel('Nifty211_RBI_Policy_Combined_Master.xlsx', skiprows=7)
df_n50 = df[df['Stock Symbol'].isin(n50_symbols)].copy()

for col in ['Win Rate %', 'Average Return %', 'Net Realised P&L (₹)']:
    df_n50[col] = pd.to_numeric(df_n50[col], errors='coerce')

# RBI's decision date + the T-8..T+8 trading-day window around it: fetched
# live (RBI's own press-release feed for the date, NSE's own holiday-master
# API for which days are tradeable) instead of hand-typed. This used to say
# "Announcement Date: Friday, 09-Oct-2026" here -- RBI's actual decision date
# is 07-Oct-2026 (confirmed by RBI's resolution press release).
refresh_from_rbi(verbose=False)
today = pd.Timestamp.now().normalize()
_recent = get_most_recent_meeting(today)
meeting = _recent if _recent and (today - pd.to_datetime(_recent['date_str'])).days <= 5 else get_next_meeting(today)

holidays = load_trading_holidays()
_window = trading_day_window(meeting['date_str'], range(-8, 9), holidays)
trading_days = {off: (d.strftime('%d-%b-%Y'), d.strftime('%a')) for off, d in _window.items()}

status_note = '' if meeting['status'] == 'confirmed' else '  [SCHEDULED by RBI, not yet confirmed]'
print(f"RBI MPC Policy used for this window: {meeting['label']} | Decision Date (T): {meeting['date_str']}{status_note}")
print()

# Sort by Sector, then Win Rate descending
df_n50 = df_n50.sort_values(by=['Industry Sector', 'Win Rate %'], ascending=[True, False])

print(f"Symbol | Sector | Strategy | Window | Entry ({meeting['label'][-8:]}) | Exit ({meeting['label'][-8:]}) | WinRate | WinRatio | AvgReturn | NetPnL")
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
