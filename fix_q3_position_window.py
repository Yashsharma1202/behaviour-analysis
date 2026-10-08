"""
FIX THE Q3 FY2026-27 POSITION-TAKING WINDOW for announced stocks.

Bug found: every announced stock showed the SAME "T-2 to T+4" window -- a
generic placeholder used for display while a result is "Yet to come", which
was never replaced with the stock's own real backtested-optimal window when
sync_live_corporate_feeds.py populated real entry/exit dates on announcement.
Confirmed by comparing to these same stocks' PRIOR quarter (FY27_Q2), where
their real windows varied widely (e.g. CIPLA T-6/T+4, HDFCLIFE T-7/T+2).

Fix: for each announced stock, run the same walk-forward 8x8 grid search
already verified for Dussehra -- trained on ALL of that stock's own prior
quarterly result dates (17 available per stock, back to FY23_Q2), scored on
clean futures OHLC data -- to find its real best (bo, so, side). Apply that
window to the ALREADY NSE-CONFIRMED Q3 result date via the standard NSE-2026
trading-calendar shift (no future price data needed, just calendar math).

ADDITIVE: reads event_dashboard_data.json + futures price data; updates the
15 announced FY27_Q3 stock records in place (direction/window/dates change --
this corrects the live position, nothing is deleted).
"""
import json
from datetime import date, datetime, timedelta
import pandas as pd

BASE = r'D:\behaviour analysis'

d = json.load(open(BASE + r'\dashboard_data\event_dashboard_data.json', encoding='utf-8'))
q3 = next(q for q in d['quarters'] if q['q_code'] == 'FY27_Q3')
announced = [s for s in q3['stocks'] if s.get('announced')]

# ---- gather each stock's own prior result dates (all quarters before Q3) ----
hist_dates = {}
for q in d['quarters']:
    if q['q_code'] == 'FY27_Q3':
        continue
    for s in q['stocks']:
        if s['symbol'] in {a['symbol'] for a in announced} and s.get('result_date'):
            try:
                dt = datetime.strptime(s['result_date'], '%d-%b-%Y').date()
            except ValueError:
                continue
            hist_dates.setdefault(s['symbol'], []).append(dt)
for sym in hist_dates:
    hist_dates[sym] = sorted(set(hist_dates[sym]))

# ---- futures price data (clean, no Friday-gap defect) ----
fut = pd.read_parquet(BASE + r'\scraped_parquet\fo_futures_near_month_continuous.parquet', columns=['Date', 'Instrument', 'Close'])
fut['Date'] = pd.to_datetime(fut['Date']).dt.date
fut_closes = {s: g.sort_values('Date')[['Date', 'Close']].values.tolist() for s, g in fut.groupby('Instrument')}

MAX_GAP_DAYS = 15


def ret_for_window(a, hd, bo, so):
    ib = -1
    for i, row in enumerate(a):
        if row[0] <= hd: ib = i
        else: break
    if ib < 0 or (hd - a[ib][0]).days > MAX_GAP_DAYS:
        return None
    en, ex = ib - bo, ib + so
    if en < 1 or ex >= len(a): return None
    g = 1.0
    for k in range(en, ex + 1):
        p0, p1 = a[k - 1][1], a[k][1]
        if not p0: return None
        dr = (p1 - p0) / p0
        if abs(dr) <= 0.20: g *= (1 + dr)
    return (g - 1) * 100


def best_window(a, prior_dates, side, min_n=3):
    best, best_score = (2, 4), (-9e9, -9e9)
    for bo in range(1, 9):
        for so in range(1, 9):
            rets = []
            for hd in prior_dates:
                v = ret_for_window(a, hd, bo, so)
                if v is not None:
                    rets.append(v if side == 'LONG' else -v)
            if len(rets) >= min_n:
                win = sum(1 for v in rets if v > 0) / len(rets) * 100
                avg = sum(rets) / len(rets)
                if (win, avg) > best_score:
                    best_score = (win, avg)
                    best = (bo, so)
    return best, best_score


NSE_2026_HOLIDAYS = {
    # '2026-03-04' removed (not a real NSE holiday -- leftover from Holi's date
    # once being wrongly typed as 04-Mar; real Holi holiday is 03-Mar, kept
    # below). '2026-11-08'/'2026-11-10' (Diwali) added -- were missing entirely.
    '2026-01-26', '2026-03-03', '2026-03-26', '2026-03-31', '2026-04-03',
    '2026-04-14', '2026-05-01', '2026-05-28', '2026-06-26', '2026-09-14', '2026-10-02',
    '2026-10-20', '2026-11-08', '2026-11-10', '2026-11-24', '2026-12-25',
}
NSE_2026_HOLIDAYS_NEXT = {
    '2027-01-26',
}
def is_trading_day(dt):
    return dt.weekday() < 5 and dt.isoformat() not in NSE_2026_HOLIDAYS and dt.isoformat() not in NSE_2026_HOLIDAYS_NEXT
def shift_trading_days(anchor, n):
    dt = anchor
    step = 1 if n >= 0 else -1
    cnt = abs(n)
    while cnt > 0:
        dt += timedelta(days=step)
        if is_trading_day(dt): cnt -= 1
    return dt


print(f"{'Symbol':<12}{'Old Window':<14}{'New Window':<14}{'Old Dir':<9}{'New Dir':<9}{'n':<4}{'New Entry':<14}{'New Exit'}")
updated = 0
for s in announced:
    sym = s['symbol']
    a = fut_closes.get(sym)
    prior = hist_dates.get(sym, [])
    if not a or len(prior) < 3:
        print(f"{sym:<12} SKIPPED - insufficient data (prior={len(prior)}, has_price={a is not None})")
        continue

    (lbo, lso), lsc = best_window(a, prior, 'LONG')
    (sbo, sso), ssc = best_window(a, prior, 'SHORT')
    if ssc > lsc:
        side, bo, so = 'SHORT', sbo, sso
    else:
        side, bo, so = 'LONG', lbo, lso

    result_dt = datetime.strptime(s['result_date'], '%d-%b-%Y').date()
    en_dt = shift_trading_days(result_dt, -bo)
    ex_dt = shift_trading_days(result_dt, so)

    old_window, old_dir = s['window'], s['direction']
    s['direction'] = side
    s['window'] = f"T-{bo} to T+{so}"
    s['lead'] = bo
    s['hold'] = so
    s['entry_lead_days'] = bo
    s['exit_hold_days'] = so
    s['taking_window_raw'] = s['window']
    s['entry_date'] = en_dt.strftime('%d-%b-%Y (%a)')
    s['exit_date'] = ex_dt.strftime('%d-%b-%Y (%a)')
    s['pos_window'] = s['window']
    s['taking_window_full'] = f"Entry: {s['entry_date']} ({s['window']}) \u2794 Exit: {s['exit_date']}"
    s['window_fix_note'] = f"Corrected from generic T-2/T+4 placeholder; real window from walk-forward grid search on {len(prior)} prior quarters."

    print(f"{sym:<12}{old_window:<14}{s['window']:<14}{old_dir:<9}{side:<9}{len(prior):<4}{s['entry_date']:<22}{s['exit_date']}")
    updated += 1

print(f"\nUpdated {updated} of {len(announced)} announced stocks with real per-stock windows.")

json.dump(d, open(BASE + r'\dashboard_data\event_dashboard_data.json', 'w', encoding='utf-8'), indent=2)
print("Saved:", BASE + r'\dashboard_data\event_dashboard_data.json')
