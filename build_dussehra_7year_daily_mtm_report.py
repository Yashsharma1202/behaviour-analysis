"""
DUSSEHRA \u2014 LAST 7 YEARS (2019-2025), DAILY MTM REPORT (Nifty-50 + BSE basket).
Same structure as the single-year report, generalised: one Daily_MTM_Grid_YYYY
sheet per year (2019-2025, the full depth of the clean futures price data),
each built the same point-in-time way -- for year Y, the window/direction is
chosen using ONLY Dussehra dates strictly before Y (no lookahead), exactly
the walk-forward model's own method, just re-run year by year instead of
only for 2026's "upcoming" pick. Years with fewer than 3 prior years (2019-
2021) are flagged LOW confidence -- shown, not hidden, since this report is
meant to show what actually happened each year, not just the years with a
statistically strong sample.

Also adds one All_Years_Event_Summary sheet comparing all 7 years side by
side (win rate, PnL, margin, DD, Sharpe, VaR, best/worst stock per year).

ADDITIVE: reads the walk-forward pickle + price data; writes a new report.
"""
import json, pickle, os, statistics as st
from datetime import date
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

BASE = r'D:\behaviour analysis'
OUT_XLSX = os.environ.get('DUSS7Y_OUT_XLSX', BASE + r'\Dussehra_7Year_Daily_MTM_Report.xlsx')

results = pickle.load(open(BASE + r'\scraped_parquet\_tmp_dussehra_walkforward.pkl', 'rb'))
fo = json.load(open(BASE + r'\dashboard_data\fo_stocks_211.json', encoding='utf-8'))
meta = {f['symbol']: f for f in fo}
d = json.load(open(BASE + r'\dashboard_data\event_dashboard_data.json', encoding='utf-8'))
dus = next(h for h in d['holidays'] if h['id'] == 'dussehra_2026')
NIFTY_SET = set(s['symbol'] for s in dus['stocks'])
live_meta = {s['symbol']: s for s in dus['stocks']}

# Futures-only (the cash file is missing 98.3% of all Fridays across its
# whole 26yr history -- see the walk-forward model's own header note).
fut = pd.read_parquet(BASE + r'\scraped_parquet\fo_futures_near_month_continuous.parquet', columns=['Date', 'Instrument', 'Open', 'Close'])
fut['Date'] = pd.to_datetime(fut['Date']).dt.date
fut_closes = {s: g.sort_values('Date')[['Date', 'Close', 'Open']].values.tolist() for s, g in fut.groupby('Instrument')}

beh = json.load(open(BASE + r'\dashboard_data\nifty_futures_holiday_behaviour.json', encoding='utf-8'))
dh = [h for h in beh['holidays'] if 'Dussehra' in h['name']][0]
dussehra_dates = sorted(date.fromisoformat(t['holiday_date']) for t in dh['trades'])

YEARS = [y for y in range(2019, 2026)]  # last 7 years of clean futures-backed Dussehra data
anchor_by_year = {y: next(dt for dt in dussehra_dates if dt.year == y) for y in YEARS}


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


def pick_for_year(sym, a, target_year):
    """Point-in-time window/direction for target_year, using ONLY Dussehra
    dates strictly before it (same discipline as the full walk-forward
    model). Relaxed sample-size gate (min 1, not 3) so early years with thin
    prior history still get a result -- always flagged by n_prior count."""
    prior_all = [dt for dt in dussehra_dates if dt.year < target_year]
    prior = [hd for hd in prior_all if any(ret_for_window(a, hd, bo, so) is not None for bo in range(1, 9) for so in range(1, 3))]
    (lbo, lso), lsc = best_window_relaxed(a, prior, 'LONG', min_n=1)
    (sbo, sso), ssc = best_window_relaxed(a, prior, 'SHORT', min_n=1)
    if lsc == (-9e9, -9e9) and ssc == (-9e9, -9e9):
        return 'LONG', 2, 5, 0
    if ssc > lsc:
        return 'SHORT', sbo, sso, len(prior)
    return 'LONG', lbo, lso, len(prior)


def closes_for(sym):
    return fut_closes.get(sym)


def lot_for(sym):
    m = meta.get(sym, {})
    if m.get('lot_size'):
        return m['lot_size']
    return live_meta.get(sym, {}).get('lot') or 500


def build_year_paths(target_year):
    anchor = anchor_by_year[target_year]
    stock_paths = {}
    for sym in NIFTY_SET:
        a = closes_for(sym)
        if a is None:
            continue
        ib = -1
        for i, row in enumerate(a):
            dt = row[0]
            if dt <= anchor: ib = i
            else: break
        if ib < 0:
            continue

        side, bo, so, n_prior = pick_for_year(sym, a, target_year)
        low_conf = n_prior < 3
        basis = f"n={n_prior} prior year(s)" if n_prior > 0 else "default (zero prior history)"

        en_idx, ex_idx = ib - bo, ib + so
        if en_idx < 0 or ex_idx >= len(a):
            continue
        lot = lot_for(sym)
        sign = 1 if side == 'LONG' else -1
        en_date, en_close, en_px = a[en_idx]

        g = 1.0
        if en_px:
            dr0 = (en_close - en_px) / en_px
            if abs(dr0) <= 0.20:
                g *= (1 + dr0)
        mtm_pct0 = (g - 1) * 100 * sign
        daily = [(en_date, round(mtm_pct0 / 100 * en_px * lot, 2))]

        for k in range(en_idx + 1, ex_idx + 1):
            p0, p1 = a[k - 1][1], a[k][1]
            if p0:
                dr = (p1 - p0) / p0
                if abs(dr) <= 0.20:
                    g *= (1 + dr)
            mtm_pct = (g - 1) * 100 * sign
            daily.append((a[k][0], round(mtm_pct / 100 * en_px * lot, 2)))

        exit_px = a[ex_idx][1]
        margin = round(en_px * lot * 0.20, 2)
        vals = [v for _, v in daily]
        peak = vals[0]; maxdd = 0.0
        for v in vals:
            peak = max(peak, v); maxdd = min(maxdd, v - peak)
        daily_changes = [vals[i] - vals[i - 1] for i in range(1, len(vals))]
        if len(daily_changes) >= 2 and st.pstdev(daily_changes) > 0:
            sharpe = round(st.mean(daily_changes) / st.pstdev(daily_changes), 2)
            sorted_ch = sorted(daily_changes)
            var95 = round(sorted_ch[max(0, int(len(sorted_ch) * 0.05) - 1)], 2)
        else:
            sharpe = var95 = None

        stock_paths[sym] = dict(direction=side, window=f"T-{bo} to T+{so}", entry_date=en_date,
                                 exit_date=a[ex_idx][0], entry_px=round(en_px, 2), exit_px=round(exit_px, 2),
                                 lot=lot, margin=margin, daily=daily, low_confidence=low_conf, basis=basis,
                                 max_dd_rs=round(maxdd, 2), cum_pnl=vals[-1], sharpe=sharpe, var95=var95)
    return stock_paths


print("Building year-by-year Dussehra MTM paths (2019-2025)...")
year_paths = {y: build_year_paths(y) for y in YEARS}
for y in YEARS:
    print(f"  {y}: {len(year_paths[y])} of {len(NIFTY_SET)} stocks")

NAVY = PatternFill('solid', fgColor='1A202C'); WHITE = Font(bold=True, color='FFFFFF', size=8)
TITLE = Font(bold=True, color='1A202C', size=13); SUB = Font(color='4A5568', size=10, italic=True)
GF = PatternFill('solid', fgColor='E6F4EA'); RF = PatternFill('solid', fgColor='FCE8E6')
AM = PatternFill('solid', fgColor='FEF3C7'); LBL = PatternFill('solid', fgColor='F1F5F9')
GOOD = Font(color='137333', bold=True); BAD = Font(color='C5221F', bold=True); BOLD = Font(bold=True)
THIN = Border(*[Side(style='thin', color='D9E2EC')] * 4)
CEN = Alignment(horizontal='center'); RIGHT = Alignment(horizontal='right'); WRAP = Alignment(wrap_text=True, vertical='top')

wb = Workbook()
wb.remove(wb.active)

year_summaries = {}
for y in YEARS:
    stock_paths = year_paths[y]
    if not stock_paths:
        continue
    all_dates = sorted(set(dt for p in stock_paths.values() for dt, _ in p['daily']))
    portfolio_mtm = []
    for dt in all_dates:
        total = sum(next((v for d2, v in p['daily'][::-1] if d2 <= dt), 0) for p in stock_paths.values())
        portfolio_mtm.append((dt, round(total, 2)))
    port_vals = [v for _, v in portfolio_mtm]
    port_peak = port_vals[0]; port_maxdd = 0.0
    for v in port_vals:
        port_peak = max(port_peak, v); port_maxdd = min(port_maxdd, v - port_peak)

    total_margin = round(sum(p['margin'] for p in stock_paths.values()), 2)
    final_pnl = round(sum(p['cum_pnl'] for p in stock_paths.values()), 2)
    wins = sum(1 for p in stock_paths.values() if p['cum_pnl'] > 0)
    win_rate = round(100 * wins / len(stock_paths), 1)
    best = max(stock_paths.items(), key=lambda kv: kv[1]['cum_pnl'])
    worst = min(stock_paths.items(), key=lambda kv: kv[1]['cum_pnl'])
    n_low = sum(1 for p in stock_paths.values() if p['low_confidence'])
    year_summaries[y] = dict(n_stocks=len(stock_paths), win_rate=win_rate, wins=wins, final_pnl=final_pnl,
                              total_margin=total_margin, port_maxdd=port_maxdd, best=best, worst=worst, n_low=n_low)

    # ---- Daily_MTM_Grid_YYYY sheet ----
    mg = wb.create_sheet(f'Daily_MTM_Grid_{y}')
    n_date_cols = len(all_dates)
    last_col = 1 + n_date_cols + 13
    mg.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    mg.cell(1, 1, f'DUSSEHRA {y} \u2014 DAILY MARK-TO-MARKET, ALL STOCKS (point-in-time)').font = TITLE
    mg.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
    mg.cell(2, 1, f'Anchor {anchor_by_year[y].strftime("%d-%b-%Y")}. {len(stock_paths)} of {len(NIFTY_SET)} stocks covered ({n_low} LOW confidence \u2014 fewer than 3 prior years to choose the window from). Entry Price = Open (09:20 AM proxy); Exit Price = Close (15:15 PM proxy).').font = SUB
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
            if k == 7: cc.font = GOOD if p['cum_pnl'] >= 0 else BAD
            if k == 8: cc.font = BAD
        row += 1
    mg.column_dimensions['A'].width = 12
    for j in range(2, 2 + n_date_cols):
        mg.column_dimensions[get_column_letter(j)].width = 9
    for j, w in enumerate([10, 14, 13, 15, 13, 15, 12, 16, 12, 9, 12, 11, 36], 2 + n_date_cols):
        mg.column_dimensions[get_column_letter(j)].width = w
    mg.freeze_panes = 'B5'

# ============ All_Years_Event_Summary sheet ============
sm = wb.create_sheet('All_Years_Event_Summary', 0)
sm.merge_cells('A1:J1')
sm.cell(1, 1, 'DUSSEHRA \u2014 ALL 7 YEARS, WHOLE-EVENT COMPARISON (all stocks combined, 1 lot each)').font = TITLE
sm.merge_cells('A2:J2')
sm.cell(2, 1, '2019-2021 have few/no prior years to choose a window from (flagged via Low-Confidence Stocks) \u2014 shown for completeness, read with extra caution. 2022-2025 have a real 3+ prior-year walk-forward basis.').font = SUB
hdr = ['Year', 'Anchor Date', 'Stocks', 'Low-Conf Stocks', 'Win Rate %', 'Final PnL (\u20b9)', 'Margin Deployed (\u20b9)', 'Return on Margin %', 'Portfolio Max DD (\u20b9)', 'Best Stock', 'Worst Stock']
for j, h in enumerate(hdr, 1):
    c = sm.cell(4, j, h); c.fill = NAVY; c.font = WHITE; c.border = THIN; c.alignment = CEN
row = 5
for y in YEARS:
    s = year_summaries.get(y)
    if not s: continue
    rom = round(s['final_pnl'] / s['total_margin'] * 100, 2) if s['total_margin'] else 0
    vals = [y, anchor_by_year[y].strftime('%d-%b-%Y'), s['n_stocks'], s['n_low'], s['win_rate'],
            s['final_pnl'], s['total_margin'], rom, s['port_maxdd'],
            f"{s['best'][0]} ({s['best'][1]['cum_pnl']:+,.0f})", f"{s['worst'][0]} ({s['worst'][1]['cum_pnl']:+,.0f})"]
    fill = GF if s['final_pnl'] >= 0 else RF
    for j, v in enumerate(vals, 1):
        c = sm.cell(row, j, v); c.border = THIN; c.fill = fill
        if j in (1, 3, 4): c.alignment = CEN
        else: c.alignment = RIGHT if j not in (2, 10, 11) else CEN
        if j == 6: c.font = GOOD if s['final_pnl'] >= 0 else BAD
        if j == 9: c.font = BAD
    row += 1
# 7-year totals row
all_pnl = sum(s['final_pnl'] for s in year_summaries.values())
all_margin = sum(s['total_margin'] for s in year_summaries.values())
r_total = row
sm.cell(r_total, 1, '7-YR TOTAL').font = BOLD
sm.cell(r_total, 6, round(all_pnl, 2)).font = GOOD if all_pnl >= 0 else BAD
sm.cell(r_total, 7, round(all_margin, 2))
for j in (1, 6, 7):
    sm.cell(r_total, j).border = THIN
    sm.cell(r_total, j).alignment = RIGHT if j != 1 else CEN
for j, w in enumerate([8, 14, 8, 13, 11, 14, 16, 14, 16, 24, 24], 1):
    sm.column_dimensions[get_column_letter(j)].width = w
sm.freeze_panes = 'A5'

note_row = r_total + 2
sm.merge_cells(f'A{note_row}:D{note_row}')
nt = sm.cell(note_row, 1, 'Source: futures prices (2019-2026), chosen over the cash-equity file which is missing 98.3% of all Fridays across its full 26yr history. Window/direction for each year picked using ONLY Dussehra dates strictly before it (walk-forward discipline, no lookahead). Margin = 20% SPAN proxy. Sharpe/VaR per stock are on each Excel grid sheet.')
nt.font = Font(italic=True, size=9, color='7C2D12'); nt.alignment = WRAP
sm.row_dimensions[note_row].height = 40

wb.save(OUT_XLSX)
print('\nSaved Excel:', OUT_XLSX)
print('7-year total PnL: Rs', f'{all_pnl:,.0f}', '| 7-year total margin: Rs', f'{all_margin:,.0f}')
