"""
sync_q3_live_nse_results.py
===============================================================================
Updates the FY27_Q3 section of dashboard_data/event_dashboard_data.json from
a fresh, direct NSE fetch (fetch_nifty50_result_dates.py's output workbook),
instead of relying only on the external D:\\share_live\\action_cop scraper
(sync_live_corporate_feeds.py's source), which can lag by several days --
on 2026-10-07 it was 6 days stale and missing 7 Nifty 50 stocks that had
since filed their board-meeting intimation with NSE directly.

Run fetch_nifty50_result_dates.py first (regenerates
Nifty50_Upcoming_Quarterly_Result_Dates_2026.xlsx), then this script.

Entry/exit dates are computed from the announced result date using NSE's
live holiday calendar (nse_trading_calendar.py) -- not hand-typed -- so a
result date that lands a T+n exit on a market holiday (e.g. Dussehra,
20-Oct-2026) is handled correctly instead of landing on a closed day.

Also fixes announced_count / pending_count, which sync_live_corporate_feeds.py's
sync_q3_earnings() updates individual stock records but never recomputes.
===============================================================================
"""
import json
import pathlib
import re
import sys

import pandas as pd

from nse_trading_calendar import load_trading_holidays, add_trading_days

sys.stdout.reconfigure(encoding='utf-8')

ROOT = pathlib.Path(__file__).resolve().parent
DASH_JSON = ROOT / 'dashboard_data' / 'event_dashboard_data.json'
SRC_XLSX = ROOT / 'Nifty50_Upcoming_Quarterly_Result_Dates_2026.xlsx'
Q_CODE = 'FY27_Q3'  # hardcodes the live quarter -- update when it rolls over

if not SRC_XLSX.exists():
    raise SystemExit(f"{SRC_XLSX.name} not found -- run fetch_nifty50_result_dates.py first.")

df = pd.read_excel(SRC_XLSX)
holidays = load_trading_holidays()

announced = {}
for _, r in df.iterrows():
    sym = str(r['Nifty 50 Symbol']).strip().upper()
    status = str(r['NSE Announcement Status'])
    raw = str(r['Upcoming Result Date / Expected Window']).strip()
    if 'ANNOUNCED' not in status.upper():
        continue
    m = re.match(r'(\d{2}-[A-Za-z]{3}-\d{4})', raw)
    if not m:
        continue
    announced[sym] = pd.to_datetime(m.group(1), format='%d-%b-%Y')

print(f"Live NSE fetch: {len(announced)} Nifty 50 stocks with an announced Q2 FY27 result date.")

with open(DASH_JSON, 'r', encoding='utf-8') as f:
    data = json.load(f)

q = next((x for x in data['quarters'] if x['q_code'] == Q_CODE), None)
if q is None:
    raise SystemExit(f"No quarter with q_code == {Q_CODE!r} found in {DASH_JSON.name}.")

updated = 0
newly_announced = []
for s in q['stocks']:
    sym = s.get('symbol')
    if sym not in announced:
        continue
    res_dt = announced[sym]
    was_announced = bool(s.get('announced'))
    lead_days = int(s.get('lead', s.get('entry_lead_days', 2)) or 2)
    hold_days = int(s.get('hold', s.get('exit_hold_days', 4)) or 4)
    en_dt = add_trading_days(res_dt, -lead_days, holidays)
    ex_dt = add_trading_days(res_dt, hold_days, holidays)

    # result_date/result_declaration_date use the project-wide "%d-%b-%Y"
    # convention (no weekday), matching sync_live_corporate_feeds.py's
    # format_date_str() -- an earlier version of this script wrote ISO
    # (YYYY-MM-DD) here, which broke fix_q3_position_window.py downstream
    # (it parses result_date with strptime('%d-%b-%Y')).
    changed = (s.get('result_date') != res_dt.strftime('%d-%b-%Y')
               or not was_announced)
    s['result_date'] = res_dt.strftime('%d-%b-%Y')
    s['result_declaration_date'] = res_dt.strftime('%d-%b-%Y')
    s['entry_date'] = f"{en_dt.strftime('%d-%b-%Y')} ({en_dt.strftime('%a')})"
    s['exit_date'] = f"{ex_dt.strftime('%d-%b-%Y')} ({ex_dt.strftime('%a')})"
    s['announced'] = True
    s['result_status'] = 'Announced'
    s['source'] = 'NSE live fetch (fetch_nifty50_result_dates.py)'
    if changed:
        updated += 1
        if not was_announced:
            newly_announced.append(sym)

q['announced_count'] = sum(1 for s in q['stocks'] if s.get('announced'))
q['pending_count'] = len(q['stocks']) - q['announced_count']

if updated:
    with open(DASH_JSON, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)
    print(f"Updated {updated} stock record(s) ({len(newly_announced)} newly announced: {', '.join(newly_announced)}).")
    print(f"announced_count={q['announced_count']}  pending_count={q['pending_count']}")
    print("Run `python build_dashboard.py` next to regenerate .js bundles and sync docs/.")
else:
    print("Nothing to update -- dashboard already matches the live NSE fetch.")
