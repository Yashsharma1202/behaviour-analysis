import json
import pandas as pd
import sys

from rbi_mpc_calendar import refresh_from_rbi, load_calendar
from nse_trading_calendar import load_trading_holidays, trading_day_window

sys.stdout.reconfigure(encoding='utf-8')

with open('dashboard_data/event_dashboard_data.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

nifty50_df = pd.read_csv('MW-NIFTY-50-21-Jul-2026.csv')
n50_symbols = [s.strip() for s in nifty50_df['SYMBOL'].dropna() if s.strip() != 'NIFTY 50']

df = pd.read_excel('Nifty211_RBI_Policy_Combined_Master.xlsx', skiprows=7)
df_n50 = df[df['Stock Symbol'].isin(n50_symbols)].copy()

for col in ['Win Rate %', 'Average Return %', 'Net Realised P&L (₹)']:
    df_n50[col] = pd.to_numeric(df_n50[col], errors='coerce')

# The two RBI meetings dashboard_data/event_dashboard_data.json tracks
# (data['rbi_policy'][0] and [1]) are picked live, chronologically, instead
# of being hand-typed as "the Oct 2026 one" and "the Dec 2026 one": the
# nearest meeting on/after today, and the one after that. Their T-8..T+8
# trading-day windows are computed from RBI's own decision date (not a hand-
# typed table -- that table used to say T = 09-Oct-2026; RBI's actual
# decision date is 07-Oct-2026) and NSE's live holiday calendar.
refresh_from_rbi(verbose=False)
holidays = load_trading_holidays()
today = pd.Timestamp.now().normalize()
upcoming_meetings = [e for e in load_calendar() if pd.to_datetime(e['date_str']) >= today][:2]
if len(upcoming_meetings) < 2:
    # Not enough future meetings cached (e.g. calendar needs a refresh) --
    # pad with the most recent past ones rather than crashing.
    past = [e for e in load_calendar() if pd.to_datetime(e['date_str']) < today]
    upcoming_meetings = (past[-(2 - len(upcoming_meetings)):] + upcoming_meetings) if past else upcoming_meetings

def _cal_map(meeting):
    window = trading_day_window(meeting['date_str'], range(-8, 9), holidays)
    return {off: (d.strftime('%d-%b-%Y'), d.strftime('%a')) for off, d in window.items()}

oct_trading_days = _cal_map(upcoming_meetings[0])
dec_trading_days = _cal_map(upcoming_meetings[1]) if len(upcoming_meetings) > 1 else oct_trading_days

print(f"RBI meeting #1 (event_dashboard_data.rbi_policy[0]): {upcoming_meetings[0]['label']} -> T = {upcoming_meetings[0]['date_str']}")
if len(upcoming_meetings) > 1:
    print(f"RBI meeting #2 (event_dashboard_data.rbi_policy[1]): {upcoming_meetings[1]['label']} -> T = {upcoming_meetings[1]['date_str']}")

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

import re as _re

def _parse_window(win_str):
    """'T-2 to T+1 (Entry: ..., Exit: ...)' -> (-2, 1). Only looks at the
    leading 'T<n> to T<m>' part, so a stale '(Entry: ..., Exit: ...)' suffix
    already baked into the string (as this one was) doesn't break the parse.
    Returns None if it doesn't match."""
    m = _re.search(r'T\s*([+-]?\d+)\s+to\s+T\s*([+-]?\d+)', str(win_str))
    return (int(m.group(1)), int(m.group(2))) if m else None

def _record_date_fields(meeting, cal_map, existing_window_str):
    """Recompute the record's own date-bearing fields (date span, decision_day,
    status, window's Entry/Exit) from the live calendar -- these used to be
    hand-typed directly into dashboard_data/event_dashboard_data.json itself
    (e.g. decision_day: '09-Oct-2026 (Friday 10:00 AM)' when RBI's actual
    decision date is 07-Oct-2026), independent of the per-stock `stocks` list
    enrich_rbi_json.py already fixed above."""
    decision_dt = pd.to_datetime(meeting['date_str'])
    meeting_start_d, _ = cal_map.get(-2, (decision_dt.strftime('%d-%b-%Y'), ''))
    fields = {
        "date": f"{meeting_start_d} to {decision_dt.strftime('%d-%b-%Y')}",
        "decision_day": f"{decision_dt.strftime('%d-%b-%Y')} ({decision_dt.strftime('%A')} 10:00 AM)",
    }
    days_away = (decision_dt - today).days
    if meeting['status'] != 'confirmed':
        fields["status"] = f"SCHEDULED FOR {decision_dt.strftime('%B').upper()} (RBI schedule, not yet confirmed)"
    elif days_away < 0:
        fields["status"] = f"DECIDED {decision_dt.strftime('%d-%b-%Y').upper()}"
    elif days_away == 0:
        fields["status"] = "DECISION TODAY"
    else:
        fields["status"] = f"UPCOMING IN {decision_dt.strftime('%B').upper()}"

    parsed = _parse_window(existing_window_str)
    if parsed:
        pre, post = parsed
        entry_d, _ = cal_map.get(pre, ("N/A", ""))
        exit_d, _ = cal_map.get(post, ("N/A", ""))
        fields["window"] = f"T{pre:+d} to T{post:+d} (Entry: {entry_d}, Exit: {exit_d})"
    return fields

if 'rbi_policy' in data and len(data['rbi_policy']) >= 2:
    data['rbi_policy'][0].update(_record_date_fields(upcoming_meetings[0], oct_trading_days, data['rbi_policy'][0].get('window')))
    data['rbi_policy'][0]['stocks'] = oct_stocks
    data['rbi_policy'][0]['stocks_count'] = len(oct_stocks)
    data['rbi_policy'][1].update(_record_date_fields(upcoming_meetings[1], dec_trading_days, data['rbi_policy'][1].get('window')))
    data['rbi_policy'][1]['stocks'] = dec_stocks
    data['rbi_policy'][1]['stocks_count'] = len(dec_stocks)

with open('dashboard_data/event_dashboard_data.json', 'w', encoding='utf-8') as f:
    json.dump(data, f, indent=2)

print(f"Successfully updated event_dashboard_data.json with {len(oct_stocks)} Nifty 50 stocks for RBI Policy!")
