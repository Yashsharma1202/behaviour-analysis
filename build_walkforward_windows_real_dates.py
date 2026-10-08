"""
REPLACE THE UNIFORM T-2/T+4 WINDOW WITH A GENUINE, VERIFIED, WALK-FORWARD
PER-STOCK WINDOW -- the same rigorous out-of-sample methodology already
used and validated elsewhere in this project (Dussehra walk-forward model,
fix_dussehra_position_window_walkforward.py) -- now applied on top of the
REAL NSE result dates just verified in rebuild_historical_quarters_real_dates.py.

Context: the user showed an old, abandoned Excel pipeline
(Futures_Master_Only_v10.xlsx) with much higher win rates (94%/96%) than
this dashboard, built on per-stock custom windows (T-1/T+1, T-8/T+2, ...)
instead of a uniform T-2/T+4. That Excel's exact window-selection script
could not be found (likely lost), so its logic cannot be literally
reproduced or trusted. Per explicit instruction, this script instead
builds a LEGITIMATE equivalent: for each stock, at each quarter with >=3
PRIOR *verified* (real-date) quarters of history, grid-search all 128
(entry-offset 1-8 x exit-offset 1-8 x LONG/SHORT) combinations scored by
(win-rate, avg-return) using ONLY prior quarters -- no lookahead -- then
apply the winning combo forward to the real result date being evaluated.

Quarters with fewer than 3 prior verified quarters, or whose own result
date is not NSE-verified (see date_source field), are left on the
existing (already real-date-rebuilt) T-2/T+4 trade -- there isn't enough
trustworthy history yet to walk-forward optimize them.

ADDITIVE: reads event_dashboard_data.json + fresh Yahoo Finance prices;
updates window/direction/entry/exit/price/outcome fields in place only
for stocks that get a genuine walk-forward pick.
"""
import json
import datetime
import pandas as pd
import yfinance as yf

BASE = r'D:\behaviour analysis'
d = json.load(open(BASE + r'\dashboard_data\event_dashboard_data.json', encoding='utf-8'))

Q_ORDER = ['FY23_Q2', 'FY23_Q3', 'FY23_Q4', 'FY24_Q1', 'FY24_Q2', 'FY24_Q3', 'FY24_Q4',
           'FY25_Q1', 'FY25_Q2', 'FY25_Q3', 'FY25_Q4', 'FY26_Q1', 'FY26_Q2', 'FY26_Q3',
           'FY26_Q4', 'FY27_Q1', 'FY27_Q2']
Q_RANK = {q: i for i, q in enumerate(Q_ORDER)}

# ---- NSE trading-day calendar (2022-2027), same as rebuild_historical_quarters_real_dates.py ----
nse_holidays = set([
    '2022-01-26', '2022-03-01', '2022-03-18', '2022-04-14', '2022-04-15', '2022-05-03', '2022-08-09', '2022-08-15', '2022-08-31', '2022-10-05', '2022-10-24', '2022-10-26', '2022-11-08',
    '2023-01-26', '2023-03-07', '2023-03-30', '2023-04-04', '2023-04-07', '2023-04-14', '2023-05-01', '2023-06-29', '2023-08-15', '2023-09-19', '2023-10-02', '2023-10-24', '2023-11-14', '2023-11-27', '2023-12-25',
    '2024-01-22', '2024-01-26', '2024-03-08', '2024-03-25', '2024-03-29', '2024-04-11', '2024-04-17', '2024-05-01', '2024-05-20', '2024-06-17', '2024-07-17', '2024-08-15', '2024-10-02', '2024-11-01', '2024-11-15', '2024-12-25',
    '2025-01-26', '2025-02-26', '2025-03-14', '2025-03-31', '2025-04-10', '2025-04-14', '2025-04-18', '2025-05-01', '2025-08-15', '2025-08-27', '2025-10-02', '2025-10-21', '2025-10-22', '2025-11-05', '2025-12-25',
    '2026-01-26', '2026-03-03', '2026-03-20', '2026-04-03', '2026-04-14', '2026-05-01', '2026-05-28', '2026-08-15', '2026-10-02', '2026-10-20', '2026-11-09', '2026-12-25'
])
all_dates = pd.date_range(start='2020-01-01', end='2027-12-31', freq='B')
td_list = sorted(dt.strftime('%Y-%m-%d') for dt in all_dates if dt.strftime('%Y-%m-%d') not in nse_holidays)
td_set = set(td_list)
_DOW = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']


def shift_trading_days(base_date, offset):
    ds = base_date.strftime('%Y-%m-%d')
    dt = base_date
    while ds not in td_set:
        dt -= datetime.timedelta(days=1)
        ds = dt.strftime('%Y-%m-%d')
    idx = td_list.index(ds)
    target_idx = max(0, min(len(td_list) - 1, idx + offset))
    return datetime.datetime.strptime(td_list[target_idx], '%Y-%m-%d').date()


def fmt_dmy(dt):
    return f"{dt.day:02d}-{dt.strftime('%b')}-{dt.year} ({_DOW[dt.weekday()]})"


# ---- 1. Collect each stock's chronological list of VERIFIED (q_code, real_result_date) ----
stock_hist = {}  # symbol -> [(q_rank, q_code, result_date), ...] chronological
stock_refs = {}  # symbol -> {q_code: stock_dict}
for q in d['quarters']:
    if q['q_code'] not in Q_RANK:
        continue
    for s in q['stocks']:
        if s.get('is_nifty50') != 'YES':
            continue
        stock_refs.setdefault(s['symbol'], {})[q['q_code']] = s
        if s.get('date_source') == 'NSE_verified':
            try:
                rd = datetime.datetime.strptime(s['result_date'], '%d-%b-%Y').date()
            except Exception:
                continue
            stock_hist.setdefault(s['symbol'], []).append((Q_RANK[q['q_code']], q['q_code'], rd))

for sym in stock_hist:
    stock_hist[sym].sort(key=lambda x: x[0])

syms = sorted(stock_hist.keys())
print(f"{len(syms)} symbols have at least some NSE-verified quarters.")

# ---- 2. Fetch price history (Open for entry, Close for exit + grid-search matrix) ----
print("\nFetching price history for the walk-forward grid search...")
TICKER_ALIASES = {'TMPV': 'TMPV.NS'}
price_data = {}
failed = []
for sym in syms:
    ticker = TICKER_ALIASES.get(sym, sym + '.NS')
    try:
        h = yf.Ticker(ticker).history(start='2021-01-01', end='2026-10-05', auto_adjust=True)
        if h.empty:
            failed.append(sym)
            continue
        h.index = h.index.date
        price_data[sym] = h
    except Exception as e:
        failed.append(sym)
        print(f"  FAILED {sym}: {e}")
print(f"Fetched {len(price_data)} of {len(syms)} symbols. Failed: {failed}")

MAX_GAP_DAYS = 15
BO_RANGE = range(1, 9)
SO_RANGE = range(1, 9)


def build_close_array(hist):
    return sorted((dt, float(px)) for dt, px in hist['Close'].items())


def build_matrix(a, dates_to_use):
    ib_list = []
    ptr = -1
    for hd in dates_to_use:
        while ptr + 1 < len(a) and a[ptr + 1][0] <= hd:
            ptr += 1
        if ptr >= 0 and (hd - a[ptr][0]).days > MAX_GAP_DAYS:
            ib_list.append(-1)
        else:
            ib_list.append(ptr)
    mat = []
    for ib in ib_list:
        row = {}
        if ib < 0:
            mat.append(row)
            continue
        for bo in BO_RANGE:
            en = ib - bo
            if en < 1:
                continue
            g = 1.0
            valid = True
            for k in range(en, ib + 1):
                p0, p1 = a[k - 1][1], a[k][1]
                if not p0:
                    valid = False
                    break
                dr = (p1 - p0) / p0
                if abs(dr) <= 0.20:
                    g *= (1 + dr)
            if not valid:
                continue
            g_anchor = g
            for so in SO_RANGE:
                ex = ib + so
                if ex >= len(a):
                    continue
                g2 = g_anchor
                ok = True
                for k2 in range(ib + 1, ex + 1):
                    p0, p1 = a[k2 - 1][1], a[k2][1]
                    if not p0:
                        ok = False
                        break
                    dr = (p1 - p0) / p0
                    if abs(dr) <= 0.20:
                        g2 *= (1 + dr)
                if ok:
                    row[(bo, so)] = (g2 - 1) * 100
        mat.append(row)
    return mat


def best_from_prior(mat, prior_idx, side):
    best, best_score = None, (-9e9, -9e9)
    for bo in BO_RANGE:
        for so in SO_RANGE:
            rets = []
            for i in prior_idx:
                v = mat[i].get((bo, so))
                if v is not None:
                    rets.append(v if side == 'LONG' else -v)
            if len(rets) >= 3:
                win = sum(1 for r in rets if r > 0) / len(rets) * 100
                avg = sum(rets) / len(rets)
                if (win, avg) > best_score:
                    best_score = (win, avg)
                    best = (bo, so)
    return best, best_score


def nearest_trading_price(hist, target_date, field, direction='forward'):
    dates = hist.index
    if target_date in dates:
        return hist.loc[target_date, field]
    candidates = sorted(dates)
    if direction == 'forward':
        for dt in candidates:
            if dt >= target_date:
                return hist.loc[dt, field]
    else:
        for dt in reversed(candidates):
            if dt <= target_date:
                return hist.loc[dt, field]
    return None


# ---- 3. Walk-forward per stock ----
wf_applied = 0
wf_skipped_low_conf = 0
flips = []
for sym in syms:
    hist_df = price_data.get(sym)
    if hist_df is None:
        continue
    a = build_close_array(hist_df)
    events = stock_hist[sym]
    dates_to_use = [e[2] for e in events]
    mat = build_matrix(a, dates_to_use)

    for i, (q_rank, q_code, result_dt) in enumerate(events):
        prior_idx = list(range(i))
        if len(prior_idx) < 3:
            wf_skipped_low_conf += 1
            continue
        (lbo, lso), lsc = best_from_prior(mat, prior_idx, 'LONG')
        (sbo, sso), ssc = best_from_prior(mat, prior_idx, 'SHORT')
        if lsc == (-9e9, -9e9) and ssc == (-9e9, -9e9):
            wf_skipped_low_conf += 1
            continue
        if ssc > lsc:
            side, bo, so = 'SHORT', sbo, sso
        else:
            side, bo, so = 'LONG', lbo, lso

        s = stock_refs[sym][q_code]
        en_dt = shift_trading_days(result_dt, -bo)
        ex_dt = shift_trading_days(result_dt, so)
        en_px = nearest_trading_price(hist_df, en_dt, 'Open', 'forward')
        ex_px = nearest_trading_price(hist_df, ex_dt, 'Close', 'backward')
        if en_px is None or ex_px is None or en_px == 0:
            continue

        sign = 1 if side == 'LONG' else -1
        ret_pct = round((ex_px - en_px) / en_px * 100 * sign, 2)
        lot = s.get('lot') or s.get('lot_size') or 1
        pnl = round(ret_pct / 100 * en_px * lot, 2)
        outcome = 'WIN' if ret_pct > 0 else 'LOSS'
        old_outcome = s.get('outcome')

        s['window'] = f"T-{bo} to T+{so}"
        s['lead'] = bo
        s['hold'] = so
        s['entry_lead_days'] = bo
        s['exit_hold_days'] = so
        s['taking_window_raw'] = s['window']
        s['direction'] = side
        s['entry_date'] = fmt_dmy(en_dt)
        s['exit_date'] = fmt_dmy(ex_dt)
        s['entry_px'] = round(float(en_px), 2)
        s['exit_px'] = round(float(ex_px), 2)
        s['entry_price'] = s['entry_px']
        s['exit_price'] = s['exit_px']
        s['actual_ret'] = ret_pct
        s['actual_return'] = ret_pct
        s['return_pct'] = ret_pct
        s['actual_pnl'] = pnl
        s['pnl'] = pnl
        s['outcome'] = outcome
        s['window_fix_note'] = (f"Walk-forward optimized window (OOS, no lookahead) from {len(prior_idx)} prior "
                                 f"NSE-verified quarters; replaces the uniform T-2/T+4 window per explicit request "
                                 f"to verify & adopt a per-stock windowing strategy.")
        s['data_source'] = ('Walk-forward per-stock window + real NSE result date + real Yahoo Finance adjusted OHLC '
                             '(2026-10-05), Open=entry/Close=exit.')
        wf_applied += 1
        if old_outcome != outcome:
            flips.append((q_code, sym, old_outcome, outcome))

print(f"\nWalk-forward window applied to {wf_applied} stock-quarter records.")
print(f"Skipped (insufficient prior verified quarters, <3): {wf_skipped_low_conf} -- left on existing T-2/T+4 real-date trade.")
print(f"Outcome flipped vs the T-2/T+4 real-date version for {len(flips)} records.")

# ---- 4. Recompute quarter-level aggregates ----
print("\nRecomputed quarter-level aggregates (Nifty 50 scope):")
for q in d['quarters']:
    if q['q_code'] not in Q_RANK:
        continue
    n50_stocks = [s for s in q['stocks'] if s.get('is_nifty50') == 'YES']
    pnls = [s['actual_pnl'] for s in n50_stocks if s.get('actual_pnl') is not None]
    wins = sum(1 for s in n50_stocks if s.get('outcome') == 'WIN')
    losses = sum(1 for s in n50_stocks if s.get('outcome') == 'LOSS')
    n = wins + losses
    old_wr = q.get('win_rate_avg')
    q['win_rate_avg'] = round(100 * wins / n, 1) if n else q.get('win_rate_avg')
    q['avg_win_rate'] = q['win_rate_avg']
    q['total_wins'] = wins
    q['total_losses'] = losses
    q['total_pnl'] = round(sum(pnls), 2) if pnls else q.get('total_pnl')
    n_long = sum(1 for s in n50_stocks if s.get('direction') == 'LONG')
    n_short = sum(1 for s in n50_stocks if s.get('direction') == 'SHORT')
    q['bias'] = f"{n_short} SHORT / {n_long} LONG"
    print(f"  {q['q_code']:<10} win_rate: {old_wr} -> {q['win_rate_avg']}   wins/losses={wins}/{losses} (of {n} scored, {len(n50_stocks)-n} unscored)")

json.dump(d, open(BASE + r'\dashboard_data\event_dashboard_data.json', 'w', encoding='utf-8'), indent=2)
print("\nSaved:", BASE + r'\dashboard_data\event_dashboard_data.json')
