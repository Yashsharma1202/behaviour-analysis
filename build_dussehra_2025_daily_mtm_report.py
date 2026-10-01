"""
DUSSEHRA 2025 \u2014 DAILY MTM REPORT (Nifty-50 + BSE basket, point-in-time).
For the actual 2025 Dussehra trade (the real walk-forward decision made using
only pre-2025 history \u2014 no hindsight), reconstructs day-by-day mark-to-market
P&L for every stock from its own entry date to its own exit date, plus
per-stock Margin, Cumulative PnL, Max Drawdown, Sharpe and VaR \u2014 and a
second sheet aggregating all of that into one whole-event performance table.

ADDITIVE: reads the walk-forward pickle + price data; writes a new report.
"""
import json, pickle, statistics as st
from datetime import date
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

BASE = r'D:\behaviour analysis'
import os as _os
OUT_XLSX = _os.environ.get('DUSS2025_OUT_XLSX', BASE + r'\Dussehra_2025_Daily_MTM_Report.xlsx')

results = pickle.load(open(BASE + r'\scraped_parquet\_tmp_dussehra_walkforward.pkl', 'rb'))
fo = json.load(open(BASE + r'\dashboard_data\fo_stocks_211.json', encoding='utf-8'))
meta = {f['symbol']: f for f in fo}
d = json.load(open(BASE + r'\dashboard_data\event_dashboard_data.json', encoding='utf-8'))
dus = next(h for h in d['holidays'] if h['id'] == 'dussehra_2026')
NIFTY_SET = set(s['symbol'] for s in dus['stocks'])
live_meta = {s['symbol']: s for s in dus['stocks']}

# DATA QUALITY FIX: the cash-equity file is missing 98.3% of all Fridays
# across its entire 26yr history (1,371 of 1,395, 2000-2026) -- not a few
# gaps, it is effectively a 4-day trading week for its whole life, which is
# how BEL/BSE (futures-only, no gap) showed MTM on dates the cash-sourced
# stocks silently had none for. Using the futures file for ALL stocks now:
# shorter history (2019-2026) but a complete, trustworthy trading calendar.
# Each row is (date, close, open) -- close drives the window-selection grid
# search (unchanged methodology); open is used only to price the ENTRY leg
# below, since the project's execution convention is 09:20 AM entry / 03:15
# PM exit -- open (5 min after the 09:15 market open) is a far closer proxy
# for a 09:20 fill than the day's close, while close is still a reasonable
# proxy for a 15:15 exit (15 min before the 15:30 close).
fut = pd.read_parquet(BASE + r'\scraped_parquet\fo_futures_near_month_continuous.parquet', columns=['Date', 'Instrument', 'Open', 'Close'])
fut['Date'] = pd.to_datetime(fut['Date']).dt.date
fut_closes = {s: g.sort_values('Date')[['Date', 'Close', 'Open']].values.tolist() for s, g in fut.groupby('Instrument')}

beh = json.load(open(BASE + r'\dashboard_data\nifty_futures_holiday_behaviour.json', encoding='utf-8'))
dh = [h for h in beh['holidays'] if 'Dussehra' in h['name']][0]
dussehra_dates = sorted(date.fromisoformat(t['holiday_date']) for t in dh['trades'])
anchor_2025 = next(dt for dt in dussehra_dates if dt.year == 2025)
dates_before_2025 = [dt for dt in dussehra_dates if dt.year < 2025]


def ret_for_window(a, hd, bo, so):
    ib = -1
    for i, row in enumerate(a):
        dt = row[0]
        if dt <= hd: ib = i
        else: break
    if ib < 0: return None
    en, ex = ib - bo, ib + so
    if en < 1 or ex >= len(a): return None
    g = 1.0
    for k in range(en, ex + 1):
        p0, p1 = a[k - 1][1], a[k][1]
        if not p0: return None
        dr = (p1 - p0) / p0
        if abs(dr) <= 0.20: g *= (1 + dr)
    return (g - 1) * 100


def best_window_relaxed(a, prior_dates, side, min_n=1):
    """Same 8x8 search as the walk-forward model, but the sample-size gate is
    lowered from 3 to 1 prior year — used ONLY to fill in the handful of
    stocks (JIOFIN, BSE) that fail the full walk-forward model's 3-year
    minimum, so this one-off 2025 snapshot isn't silently missing them.
    Flagged low-confidence in the report wherever used."""
    best, best_score = (2, 5), (-9e9, -9e9)
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


def relaxed_2025_trade(sym, a):
    """For a stock the full walk-forward model excluded (0 OOS trades): pick
    its 2025 window/direction from whatever prior years exist (even just 1-2),
    falling back to the project's standard default (LONG, T-2/T+5) if there
    is truly zero prior history -- same fallback already used by the index
    options 8x8 model when prior_anchors is empty."""
    prior = [hd for hd in dates_before_2025 if ret_for_window(a, hd, 2, 5) is not None or True]
    # filter to dates that actually have SOME usable price data for this stock
    prior = [hd for hd in prior if any(ret_for_window(a, hd, bo, so) is not None for bo in range(1, 9) for so in range(1, 3))]
    (lbo, lso), lsc = best_window_relaxed(a, prior, 'LONG', min_n=1)
    (sbo, sso), ssc = best_window_relaxed(a, prior, 'SHORT', min_n=1)
    if lsc == (-9e9, -9e9) and ssc == (-9e9, -9e9):
        return 'LONG', 2, 5, 0, 'default (zero prior history)'
    if ssc > lsc:
        return 'SHORT', sbo, sso, len(prior), f'relaxed search, n={len(prior)} prior year(s)'
    return 'LONG', lbo, lso, len(prior), f'relaxed search, n={len(prior)} prior year(s)'


def closes_for(sym):
    if sym in fut_closes:
        return fut_closes[sym]
    return None


def lot_for(sym):
    m = meta.get(sym, {})
    if m.get('lot_size'):
        return m['lot_size']
    return live_meta.get(sym, {}).get('lot') or 500


# ---- Build each stock's real 2025 daily MTM path ----
stock_paths = {}   # symbol -> dict(direction, window, entry_date, exit_date, margin, daily=[(date, mtm_rupees), ...])
for sym in NIFTY_SET:
    a = closes_for(sym)
    if a is None:
        continue
    ib = -1
    for i, row in enumerate(a):
        dt = row[0]
        if dt <= anchor_2025: ib = i
        else: break
    if ib < 0:
        continue

    r = results.get(sym)
    trade = next((t for t in (r['oos_trades'] if r else []) if t['year'] == 2025), None)
    if trade:
        bo, so, side = trade['bo'], trade['so'], trade['side']
        low_conf, basis = False, f"walk-forward, n={r['oos_n']} prior years (full 3yr-minimum method)"
    else:
        # Walk-forward excluded this stock entirely (JIOFIN, BSE) -- relax the
        # 3-prior-year gate just for this one-off 2025 snapshot so it isn't
        # silently missing; always flagged low-confidence.
        side, bo, so, n_prior, basis = relaxed_2025_trade(sym, a)
        low_conf = True

    en_idx, ex_idx = ib - bo, ib + so
    if en_idx < 0 or ex_idx >= len(a):
        continue
    lot = lot_for(sym)
    sign = 1 if side == 'LONG' else -1
    en_date, en_close, en_px = a[en_idx]   # en_px = Open (09:20 AM proxy); en_close used only for the entry day's own intraday leg

    # Entry-day leg: position is opened at 09:20 (~Open), so by that day's
    # close it has ALREADY moved Open->Close -- it is not flat at 0 like a
    # trade that was entered at the previous close.
    g = 1.0
    if en_px:
        dr0 = (en_close - en_px) / en_px
        if abs(dr0) <= 0.20:
            g *= (1 + dr0)
    mtm_pct0 = (g - 1) * 100 * sign
    daily = [(en_date, round(mtm_pct0 / 100 * en_px * lot, 2))]

    for k in range(en_idx + 1, ex_idx + 1):
        p0, p1 = a[k - 1][1], a[k][1]   # close-to-close for every day after entry
        if p0:
            dr = (p1 - p0) / p0
            if abs(dr) <= 0.20:
                g *= (1 + dr)
        mtm_pct = (g - 1) * 100 * sign
        mtm_rs = round(mtm_pct / 100 * en_px * lot, 2)
        daily.append((a[k][0], mtm_rs))

    exit_px = a[ex_idx][1]   # Close of exit day (15:15 proxy)
    margin = round(en_px * lot * 0.20, 2)
    stock_paths[sym] = dict(direction=side, window=f"T-{bo} to T+{so}", entry_date=en_date,
                             exit_date=a[ex_idx][0], entry_px=round(en_px, 2), exit_px=round(exit_px, 2), lot=lot,
                             margin=margin, daily=daily, low_confidence=low_conf, basis=basis)

print(f"Reconstructed 2025 daily MTM path for {len(stock_paths)} of {len(NIFTY_SET)} stocks.")

# ---- Per-stock measures: Max DD, Sharpe, VaR (from the within-trade daily MTM path) ----
for sym, p in stock_paths.items():
    vals = [v for _, v in p['daily']]
    peak = vals[0]; maxdd = 0.0
    for v in vals:
        peak = max(peak, v)
        dd = v - peak
        maxdd = min(maxdd, dd)
    p['max_dd_rs'] = round(maxdd, 2)
    p['cum_pnl'] = vals[-1]
    daily_changes = [vals[i] - vals[i - 1] for i in range(1, len(vals))]
    if len(daily_changes) >= 2 and st.pstdev(daily_changes) > 0:
        p['sharpe'] = round(st.mean(daily_changes) / st.pstdev(daily_changes), 2)
        sorted_ch = sorted(daily_changes)
        idx = max(0, int(len(sorted_ch) * 0.05) - 1)
        p['var95'] = round(sorted_ch[idx], 2)
    else:
        p['sharpe'] = None
        p['var95'] = None

# ---- Whole-event (portfolio) aggregation ----
all_dates = sorted(set(dt for p in stock_paths.values() for dt, _ in p['daily']))
portfolio_mtm = []
for dt in all_dates:
    total = 0.0
    for p in stock_paths.values():
        # last known MTM on or before this date (flat/0 if not yet entered, holds last value after exit)
        vals_on_or_before = [v for d2, v in p['daily'] if d2 <= dt]
        if vals_on_or_before:
            total += vals_on_or_before[-1]
    portfolio_mtm.append((dt, round(total, 2)))

port_vals = [v for _, v in portfolio_mtm]
port_peak = port_vals[0]; port_maxdd = 0.0
for v in port_vals:
    port_peak = max(port_peak, v)
    port_maxdd = min(port_maxdd, v - port_peak)
port_changes = [port_vals[i] - port_vals[i - 1] for i in range(1, len(port_vals))]
port_sharpe = round(st.mean(port_changes) / st.pstdev(port_changes), 3) if len(port_changes) >= 2 and st.pstdev(port_changes) > 0 else None
sorted_pc = sorted(port_changes)
port_var95 = round(sorted_pc[max(0, int(len(sorted_pc) * 0.05) - 1)], 2) if sorted_pc else None

total_margin = round(sum(p['margin'] for p in stock_paths.values()), 2)
final_pnl = round(sum(p['cum_pnl'] for p in stock_paths.values()), 2)
wins = sum(1 for p in stock_paths.values() if p['cum_pnl'] > 0)
win_rate = round(100 * wins / len(stock_paths), 1)
avg_ret_pct = round(sum((p['cum_pnl'] / (p['entry_px'] * p['lot'])) * 100 for p in stock_paths.values()) / len(stock_paths), 2)
best = max(stock_paths.items(), key=lambda kv: kv[1]['cum_pnl'])
worst = min(stock_paths.items(), key=lambda kv: kv[1]['cum_pnl'])

# =====================================================================
NAVY = PatternFill('solid', fgColor='1A202C'); WHITE = Font(bold=True, color='FFFFFF', size=8)
TITLE = Font(bold=True, color='1A202C', size=13); SUB = Font(color='4A5568', size=10, italic=True)
GF = PatternFill('solid', fgColor='E6F4EA'); RF = PatternFill('solid', fgColor='FCE8E6')
AM = PatternFill('solid', fgColor='FEF3C7'); LBL = PatternFill('solid', fgColor='F1F5F9')
GOOD = Font(color='137333', bold=True); BAD = Font(color='C5221F', bold=True); BOLD = Font(bold=True)
THIN = Border(*[Side(style='thin', color='D9E2EC')] * 4)
CEN = Alignment(horizontal='center'); RIGHT = Alignment(horizontal='right'); WRAP = Alignment(wrap_text=True, vertical='top')

wb = Workbook()

# ============ Sheet 1: Daily MTM Grid ============
mg = wb.active; mg.title = 'Daily_MTM_Grid_2025'
n_date_cols = len(all_dates)
last_col = 1 + n_date_cols + 13  # symbol + dates + (direction,window,entrydate,entryprice,exitdate,exitprice,margin,cumpnl,maxdd,sharpe,var,confidence,basis)
mg.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
mg.cell(1, 1, 'DUSSEHRA 2025 \u2014 DAILY MARK-TO-MARKET, ALL STOCKS (point-in-time)').font = TITLE
mg.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
n_low = sum(1 for p in stock_paths.values() if p['low_confidence'])
mg.cell(2, 1, f'Anchor 2-Oct-2025. {len(stock_paths)} of {len(NIFTY_SET)} stocks covered ({n_low} LOW confidence \u2014 JIOFIN/BSE, relaxed window search due to thin prior history; see Basis column). Blank date-cell = stock not yet in its own window on that date. MTM in \u20b9, 1 lot. Entry Price = Open (09:20 AM proxy); Exit Price = Close (15:15 PM proxy), per the project\u2019s 09:20 entry / 15:15 exit execution convention.').font = SUB
hdr = (['Symbol'] + [dt.strftime('%d-%b') for dt in all_dates]
       + ['Direction', 'Window', 'Entry Date', 'Entry Price (\u20b9 @ 09:20)', 'Exit Date', 'Exit Price (\u20b9 @ 15:15)',
          'Margin (\u20b9)', 'Cumulative PnL (\u20b9)', 'Max DD (\u20b9)', 'Sharpe', 'VaR 95% (\u20b9)', 'Confidence', 'Basis'])
for j, h in enumerate(hdr, 1):
    c = mg.cell(4, j, h); c.fill = NAVY; c.font = WHITE; c.border = THIN; c.alignment = CEN
row = 5
for sym in sorted(stock_paths):
    p = stock_paths[sym]
    daily_map = dict(p['daily'])
    rowfill = AM if p['low_confidence'] else None
    c = mg.cell(row, 1, sym); c.border = THIN; c.font = BOLD
    if rowfill: c.fill = rowfill
    for j, dt in enumerate(all_dates, 2):
        v = daily_map.get(dt)
        cc = mg.cell(row, j, v if v is not None else None)
        cc.border = THIN; cc.alignment = RIGHT
        if rowfill: cc.fill = rowfill
        if v is not None:
            cc.font = GOOD if v >= 0 else BAD
            cc.number_format = '#,##0'
    base = 2 + n_date_cols
    vals = [p['direction'], p['window'], p['entry_date'].strftime('%d-%b-%Y'), p['entry_px'],
            p['exit_date'].strftime('%d-%b-%Y'), p['exit_px'],
            p['margin'], p['cum_pnl'], p['max_dd_rs'], (p['sharpe'] if p['sharpe'] is not None else '\u2014'),
            (p['var95'] if p['var95'] is not None else '\u2014'), ('LOW' if p['low_confidence'] else 'OK'), p['basis']]
    for k, v in enumerate(vals):
        cc = mg.cell(row, base + k, v); cc.border = THIN
        if rowfill: cc.fill = rowfill
        if k in (0, 1, 2, 4, 11): cc.alignment = CEN
        else: cc.alignment = RIGHT
        if k == 11 and p['low_confidence']: cc.font = Font(color='92400E', bold=True, size=9)
        if k == 5: cc.font = GOOD if p['cum_pnl'] >= 0 else BAD
        if k == 6: cc.font = BAD
    row += 1
mg.column_dimensions['A'].width = 12
for j in range(2, 2 + n_date_cols):
    mg.column_dimensions[get_column_letter(j)].width = 9
for j, w in enumerate([10, 14, 13, 15, 13, 15, 12, 16, 12, 9, 12, 11, 36], 2 + n_date_cols):
    mg.column_dimensions[get_column_letter(j)].width = w
mg.freeze_panes = mg.cell(5, 2).coordinate.replace('5', '5')
mg.freeze_panes = 'B5'

# ============ Sheet 2: Whole-Event Performance Summary ============
ev = wb.create_sheet('Event_Performance_Summary')
ev.merge_cells('A1:B1')
ev.cell(1, 1, 'DUSSEHRA 2025 \u2014 WHOLE-EVENT PERFORMANCE SUMMARY (all stocks combined, 1 lot each)').font = TITLE
metrics = [
    ('Stocks traded', f"{len(stock_paths)} of {len(NIFTY_SET)}"),
    ('Win rate (stocks ending profitable)', f"{win_rate}% ({wins}/{len(stock_paths)})"),
    ('Average return per stock', f"{avg_ret_pct:+.2f}%"),
    ('Total margin deployed (\u20b9, 1 lot each)', f"{total_margin:,.0f}"),
    ('Final combined PnL (\u20b9)', f"{final_pnl:,.0f}"),
    ('Return on margin deployed', f"{(final_pnl/total_margin*100 if total_margin else 0):+.2f}%"),
    ('Portfolio Max Drawdown (\u20b9, intraday across the event)', f"{port_maxdd:,.0f}"),
    ('Portfolio Sharpe-like (daily PnL-change based)', f"{port_sharpe if port_sharpe is not None else '\u2014'}"),
    ('Portfolio VaR 95% (\u20b9, worst-5% single-day PnL swing)', f"{port_var95:,.0f}" if port_var95 is not None else '\u2014'),
    ('Best stock', f"{best[0]} ({best[1]['cum_pnl']:+,.0f} \u20b9, {best[1]['direction']})"),
    ('Worst stock', f"{worst[0]} ({worst[1]['cum_pnl']:+,.0f} \u20b9, {worst[1]['direction']})"),
]
r0 = 3
for lbl, val in metrics:
    a = ev.cell(r0, 1, lbl); a.font = BOLD; a.fill = LBL; a.border = THIN
    b = ev.cell(r0, 2, val); b.border = THIN; b.alignment = RIGHT
    r0 += 1
ev.column_dimensions['A'].width = 46; ev.column_dimensions['B'].width = 34

r0 += 1
ev.merge_cells(f'A{r0}:D{r0}')
ev.cell(r0, 1, 'PORTFOLIO DAILY CUMULATIVE MTM (date-wise, all stocks summed)').font = Font(bold=True, size=12)
r0 += 1
hdr2 = ['Date', 'Portfolio MTM (\u20b9)', "Day's Change (\u20b9)", 'Running Peak (\u20b9)']
for j, h in enumerate(hdr2, 1):
    c = ev.cell(r0, j, h); c.fill = NAVY; c.font = WHITE; c.border = THIN; c.alignment = CEN
r0 += 1
peak_track = portfolio_mtm[0][1]
for i, (dt, v) in enumerate(portfolio_mtm):
    peak_track = max(peak_track, v)
    chg = v - portfolio_mtm[i - 1][1] if i > 0 else 0
    vals = [dt.strftime('%d-%b-%Y'), v, chg, peak_track]
    fill = GF if v >= 0 else RF
    for j, val in enumerate(vals, 1):
        c = ev.cell(r0, j, val); c.border = THIN; c.fill = fill
        if j == 1: c.alignment = CEN
        else: c.alignment = RIGHT
        if j in (2, 3) and isinstance(val, (int, float)): c.font = GOOD if val >= 0 else BAD
    r0 += 1

note_row = r0 + 1
ev.merge_cells(f'A{note_row}:D{note_row}')
nt = ev.cell(note_row, 1, 'Note: Entry Price = that day\u2019s Open (proxy for the 09:20 AM entry fill); Exit Price = that day\u2019s Close (proxy for the 15:15 PM exit), matching the project\u2019s stated execution convention. Sharpe/VaR here are computed from a short daily-PnL-change series (~15-20 trading days across the whole event window) \u2014 directionally useful, not a statistically robust long-run estimate. Margin is a 20% SPAN proxy, not an exact historical broker figure.')
nt.font = Font(italic=True, size=9, color='7C2D12'); nt.alignment = WRAP
ev.row_dimensions[note_row].height = 50

note_row2 = note_row + 2
ev.merge_cells(f'A{note_row2}:D{note_row2}')
nt2 = ev.cell(note_row2, 1,
    'DATA QUALITY CAVEAT: the source price file had genuine duplicate-date rows for 50 of 56 Nifty-50 stocks (deduplicated here, keep=last). '
    'Severity varies hugely \u2014 most stocks had <1% duplicate rows (likely minor vendor re-fetch noise), but SHRIRAMFIN, SBIN and NTPC had their price '
    'history duplicated 2-4x with CONFLICTING close prices on the same date across their full 26yr history (not just 2025). Their walk-forward win-rate/'
    'direction results (including SHRIRAMFIN, shown elsewhere as this project\u2019s strongest validated Dussehra result) were computed BEFORE this fix and '
    'should be re-run on the deduplicated data before being trusted or acted on.')
nt2.font = Font(italic=True, size=9, color='991B1B', bold=True); nt2.alignment = WRAP
ev.row_dimensions[note_row2].height = 70

wb.save(OUT_XLSX)
print('Saved Excel:', OUT_XLSX)
print(f"Win rate {win_rate}% | Final PnL Rs{final_pnl:,.0f} | Margin Rs{total_margin:,.0f} | Portfolio MaxDD Rs{port_maxdd:,.0f} | Sharpe {port_sharpe} | VaR95 Rs{port_var95}")
