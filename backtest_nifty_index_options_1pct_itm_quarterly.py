"""
NIFTY INDEX OPTIONS — 1% ITM QUARTERLY EARNINGS-SEASON BACKTEST (2019 -> 2026)
================================================================================
Applies the project's 8x8 matrix behaviour model (best LONG or SHORT by
win-rate/avg-return, point-in-time) to ONE NIFTY *index* option trade per quarter,
priced with REAL option quotes from D:/qunat_db/updated nifty 50.

- ADDITIVE ONLY: writes a new local workbook. Does NOT touch docs/ or dashboard_data/.
- NOT pushed to host.
- Shows EXPECTED (synthetic: intrinsic move + 2% extrinsic convention) vs
  ACTUAL (real premium from parquet) for every trade.
"""
import sys
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
import os, glob
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

OPT_DIR  = r'D:\qunat_db\updated nifty 50\options'
SPOT_DIR = r'D:\qunat_db\updated nifty 50\spot'
SPOT_DAILY = r'D:\qunat_db\updated nifty 50\spot\nifty_spot_daily_2000_2026.parquet'
OUT_XLSX = r'D:\behaviour analysis\Nifty_Index_Options_1PCT_ITM_Quarterly_Backtest_2019_2026.xlsx'

REPORT_LAST_N = 12     # report/summarize only the most recent N quarter-trades (last 12 quarters)
COST = 0.0015          # round-trip cost fraction on the option premium (matches project convention ~0.15%)
EXTRINSIC = 0.02       # 2.0% of spot extrinsic at entry (project synthetic-pricing convention)
EXTRINSIC_EXIT = 0.005 # 0.5% of spot extrinsic at exit (project convention)
ENTRY_TIME = '09:20'
EXIT_TIME  = '15:15'

# Earnings-season peak anchor (month, day) per fiscal quarter label
ANCHORS = {  # quarter of RESULTS -> approx peak announcement date
    'Q1': (7, 25),   # Apr-Jun results peak late Jul
    'Q2': (10, 25),  # Jul-Sep results peak late Oct
    'Q3': (1, 25),   # Oct-Dec results peak late Jan
    'Q4': (5, 8),    # Jan-Mar results peak early May
}

def round_nse_strike(p):
    return int(round(p / 50) * 50)

# ---------- Trading calendar & spot ----------
# NOTE: the spot DAILY master (nifty_spot_daily_2000_2026.parquet) has gaps (e.g. all of
# H1-2019 is missing). We therefore build the authoritative trading calendar from the
# actual option files, and read spot from the per-day spot parquets (complete for every
# option trading day).
print('Loading option-day calendar...', flush=True)
opt_days = sorted(os.path.basename(f)[:-8] for f in glob.glob(os.path.join(OPT_DIR, '*.parquet')))
opt_day_set = set(opt_days)
first_opt, last_opt = opt_days[0], opt_days[-1]
cal = opt_days  # authoritative trading calendar = real option trading days
print(f'  option data: {first_opt} .. {last_opt} ({len(opt_days)} days)')

_spot_df_cache = {}
def _spot_df(day):
    if day in _spot_df_cache:
        return _spot_df_cache[day]
    f = os.path.join(SPOT_DIR, day + '.parquet')
    df = None
    if os.path.exists(f):
        df = pd.read_parquet(f)
        df.columns = [c.lower() for c in df.columns]
        if 'time' in df.columns:
            df['time'] = df['time'].astype(str)
    _spot_df_cache[day] = df
    if len(_spot_df_cache) > 80:
        _spot_df_cache.pop(next(iter(_spot_df_cache)))
    return df

print('Precomputing daily spot closes from per-day files...', flush=True)
close_by_day = {}
for _i, _d in enumerate(opt_days):
    f = os.path.join(SPOT_DIR, _d + '.parquet')
    if not os.path.exists(f):
        continue
    try:
        _s = pd.read_parquet(f, columns=['Close'])
        close_by_day[_d] = float(_s['Close'].iloc[-1])
    except Exception:
        _s = pd.read_parquet(f)
        _s.columns = [c.lower() for c in _s.columns]
        if 'close' in _s.columns and len(_s):
            close_by_day[_d] = float(_s['close'].iloc[-1])
print(f'  daily spot closes loaded for {len(close_by_day)}/{len(opt_days)} days', flush=True)

def spot_at(day, when):
    """Intraday spot close at/just before HH:MM 'when' (falls back to daily close)."""
    df = _spot_df(day)
    if df is None or df.empty:
        return close_by_day.get(day)
    if 'time' in df.columns:
        r = df[df['time'].str[:5] <= when].tail(1)
        if len(r):
            return float(r['close'].iloc[0])
    return float(df['close'].iloc[-1])

def nearest_trading_on_or_after(target):
    for d in cal:
        if d >= target:
            return d
    return None

def idx_in_cal(day):
    lo, hi = 0, len(cal)
    import bisect
    i = bisect.bisect_left(cal, day)
    return i if i < len(cal) and cal[i] == day else None

def shift_trading_days(day, n):
    """n<0 => before, n>0 => after. day must be in cal (or nearest after)."""
    import bisect
    i = bisect.bisect_left(cal, day)
    if i >= len(cal):
        return None
    j = i + n
    if j < 0 or j >= len(cal):
        return None
    return cal[j]

# ---------- Build quarter anchors 2019..2026 ----------
def fiscal_label(q, anchor_year):
    # Indian FY = Apr..Mar. Q3(Oct-Dec) & Q4(Jan-Mar) results are announced in the
    # calendar year AFTER the FY start; Q1(Apr-Jun) & Q2(Jul-Sep) in the FY start year.
    fy0 = anchor_year - 1 if q in ('Q3', 'Q4') else anchor_year
    return f'{q} {fy0}-{str(fy0 + 1)[2:]}'

def build_quarters():
    qs = []
    for year in range(2019, 2027):
        for q, (mo, dy) in ANCHORS.items():
            target = f'{year:04d}-{mo:02d}-{dy:02d}'
            anchor = nearest_trading_on_or_after(target)
            if anchor is None:
                continue
            # keep only anchors within option-data span (with room for offsets)
            if anchor < first_opt or anchor > last_opt:
                continue
            qs.append({'label': fiscal_label(q, year), 'q': q, 'year': year, 'anchor': anchor})
    qs.sort(key=lambda x: x['anchor'])
    return qs

quarters = build_quarters()
print(f'Quarters in scope: {len(quarters)} ({quarters[0]["label"]} .. {quarters[-1]["label"]})')

# ---------- 8x8 matrix model on SPOT (point-in-time) ----------
def spot_close_on(day):
    return close_by_day.get(day)

def sim_spot_ret(anchor, side, bo, so):
    """Return signed % return on the underlying for entering bo days before anchor,
    exiting so days after (used only to pick direction+window like the project model)."""
    en = shift_trading_days(anchor, -bo)
    ex = shift_trading_days(anchor, so)
    if en is None or ex is None:
        return None
    pe, px = spot_close_on(en), spot_close_on(ex)
    if pe is None or px is None or pe == 0:
        return None
    r = (px / pe - 1.0)
    if side == 'SHORT':
        r = -r
    return (r - COST) * 100

def optimize_8x8(prior_anchors, side):
    best, best_score = (2, 5), (-9e9, -9e9)
    for bo in range(1, 9):
        for so in range(1, 9):
            rets = [sim_spot_ret(a, side, bo, so) for a in prior_anchors]
            rets = [r for r in rets if r is not None]
            if len(rets) >= 3:
                win = sum(1 for r in rets if r > 0) / len(rets) * 100
                avg = sum(rets) / len(rets)
                if (win, avg) > best_score:
                    best_score = (win, avg)
                    best = (bo, so)
    return best, best_score

# ---------- Real option pricing ----------
_opt_cache = {}
def load_opt_day(day):
    if day in _opt_cache:
        return _opt_cache[day]
    f = os.path.join(OPT_DIR, day + '.parquet')
    if not os.path.exists(f):
        _opt_cache[day] = None
        return None
    o = pd.read_parquet(f, columns=['Time','Close','Type','Strike','ExpiryDate','OpenInterest','Volume'])
    o['Time'] = o['Time'].astype(str)
    o['ExpiryDate'] = o['ExpiryDate'].astype(str)
    _opt_cache[day] = o
    if len(_opt_cache) > 40:  # keep memory bounded
        _opt_cache.pop(next(iter(_opt_cache)))
    return o

def pick_expiry(day, exit_day):
    """Nearest listed expiry on/after the exit day, within a sane near-month window
    (<= exit + 70 days). Rejects corrupt/absurd expiries (e.g. the junk 1970/2050
    strings in the Jan-2019 files) so such days drop cleanly instead of mispricing."""
    import datetime as dt
    o = load_opt_day(day)
    if o is None:
        return None
    try:
        xd = dt.date.fromisoformat(exit_day)
    except Exception:
        return None
    hi = xd + dt.timedelta(days=70)
    cand = []
    for e in sorted(o['ExpiryDate'].unique()):
        try:
            ed = dt.date.fromisoformat(e)
        except Exception:
            continue
        if xd <= ed <= hi:
            cand.append(e)
    return cand[0] if cand else None

def opt_premium(day, when, opt_type, strike, expiry):
    o = load_opt_day(day)
    if o is None:
        return None, None
    leg = o[(o['Type']==opt_type) & (o['ExpiryDate']==expiry)]
    if leg.empty:
        return None, None
    # nearest available strike to requested
    avail = leg['Strike'].unique()
    strike_use = min(avail, key=lambda s: abs(s - strike))
    leg = leg[leg['Strike']==strike_use]
    row = leg[leg['Time'].str[:5] <= when].tail(1)
    if row.empty:
        row = leg.sort_values('Time').head(1)
    if row.empty:
        return None, None
    return float(row['Close'].iloc[0]), float(strike_use)

# ---------- Run backtest ----------
trades = []
for i, qd in enumerate(quarters):
    anchor = qd['anchor']
    prior = [q['anchor'] for q in quarters[:i]]

    (lbo, lso), lsc = optimize_8x8(prior, 'LONG')
    (sbo, sso), ssc = optimize_8x8(prior, 'SHORT')
    # choose best side by (win-rate, avg-return); default LONG if no history
    if lsc == (-9e9, -9e9) and ssc == (-9e9, -9e9):
        side, bo, so, basis = 'LONG', 2, 5, 'default (no prior history)'
    elif ssc > lsc:
        side, bo, so, basis = 'SHORT', sbo, sso, f'8x8 best SHORT win={ssc[0]:.0f}% avg={ssc[1]:+.2f}%'
    else:
        side, bo, so, basis = 'LONG', lbo, lso, f'8x8 best LONG win={lsc[0]:.0f}% avg={lsc[1]:+.2f}%'

    opt_type = 'CE' if side == 'LONG' else 'PE'
    entry_day = shift_trading_days(anchor, -bo)
    exit_day  = shift_trading_days(anchor, so)
    if entry_day is None or exit_day is None or entry_day not in opt_day_set or exit_day not in opt_day_set:
        print(f"  DROP {qd['label']:10s}: entry/exit day missing option file (anchor={anchor}, side={side}, bo={bo}, so={so}, entry={entry_day}, exit={exit_day})", flush=True)
        continue

    sp_en = spot_at(entry_day, ENTRY_TIME)
    sp_ex = spot_at(exit_day, EXIT_TIME)
    if sp_en is None or sp_ex is None:
        print(f"  DROP {qd['label']:10s}: spot close missing", flush=True)
        continue

    # target strike: 1% ITM
    strike_req = round_nse_strike(sp_en * (0.99 if side=='LONG' else 1.01))
    expiry = pick_expiry(entry_day, exit_day)
    if expiry is None:
        print(f"  DROP {qd['label']:10s}: no expiry available on {entry_day}", flush=True)
        continue

    p_en, k_used = opt_premium(entry_day, ENTRY_TIME, opt_type, strike_req, expiry)
    p_ex, _      = opt_premium(exit_day,  EXIT_TIME,  opt_type, strike_req, expiry)
    if p_en is None or p_ex is None or p_en <= 0:
        print(f"  DROP {qd['label']:10s}: premium missing (type={opt_type}, strike~{strike_req}, expiry={expiry}, entry={entry_day}/{p_en}, exit={exit_day}/{p_ex})", flush=True)
        continue

    # EXPECTED (project synthetic convention): intrinsic move + 2% extrinsic at entry
    if side == 'LONG':
        intr_en = max(0.0, sp_en - k_used); intr_ex = max(0.0, sp_ex - k_used)
    else:
        intr_en = max(0.0, k_used - sp_en); intr_ex = max(0.0, k_used - sp_ex)
    exp_p_en = intr_en + sp_en * EXTRINSIC        # project convention: 2% extrinsic at entry
    exp_p_ex = intr_ex + sp_ex * EXTRINSIC_EXIT   # project convention: 0.5% extrinsic at exit
    exp_ret = ((exp_p_ex - exp_p_en) / exp_p_en - COST) * 100 if exp_p_en > 0 else float('nan')

    act_ret = ((p_ex - p_en) / p_en - COST) * 100

    trades.append({
        'Quarter': qd['label'], 'Anchor (T)': anchor, 'Side': side, 'Option': opt_type,
        'Window': f'T-{bo} to T+{so}', 'Basis': basis,
        'Entry Date': entry_day, 'Exit Date': exit_day,
        'Spot @Entry': round(sp_en, 1), 'Spot @Exit': round(sp_ex, 1),
        'Strike': int(k_used), 'Expiry': expiry,
        'Expected Entry Prem': round(exp_p_en, 2), 'Expected Exit Prem': round(exp_p_ex, 2),
        'Expected Ret %': round(exp_ret, 2),
        'Actual Entry Prem': round(p_en, 2), 'Actual Exit Prem': round(p_ex, 2),
        'Actual Ret %': round(act_ret, 2),
        'Exp-Act Gap %': round(exp_ret - act_ret, 2),
    })
    print(f"  {qd['label']:10s} {side:5s} {opt_type} {f'T-{bo}/T+{so}':10s} "
          f"K={int(k_used):6d} exp={exp_ret:+7.2f}%  act={act_ret:+7.2f}%", flush=True)

df = pd.DataFrame(trades)
# Point-in-time training used ALL prior quarters above; report only the last N quarters.
if REPORT_LAST_N and len(df) > REPORT_LAST_N:
    df = df.tail(REPORT_LAST_N).reset_index(drop=True)
    print(f"\nReporting the last {REPORT_LAST_N} quarters: {df['Quarter'].iloc[0]} .. {df['Quarter'].iloc[-1]}")

print('\n================= SUMMARY =================')
print(f'Trades: {len(df)}')
if len(df):
    print(f"Actual  : win-rate {(df['Actual Ret %']>0).mean()*100:5.1f}%   avg {df['Actual Ret %'].mean():+6.2f}%   total {df['Actual Ret %'].sum():+7.1f}%")
    print(f"Expected: win-rate {(df['Expected Ret %']>0).mean()*100:5.1f}%   avg {df['Expected Ret %'].mean():+6.2f}%   total {df['Expected Ret %'].sum():+7.1f}%")
    print(f"Mean Expected-vs-Actual gap: {df['Exp-Act Gap %'].mean():+.2f}%")

# ---------- Write Excel ----------
if len(df):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Trade Log'
    hdr_fill = PatternFill('solid', fgColor='1A202C'); hdr_font = Font(bold=True, color='FFFFFF', size=10)
    win_fill = PatternFill('solid', fgColor='E6F4EA'); loss_fill = PatternFill('solid', fgColor='FCE8E6')
    thin = Border(*[Side(style='thin', color='CBD5E0')]*4)
    cols = list(df.columns)
    ws.append(['NIFTY INDEX OPTIONS — 1% ITM QUARTERLY BACKTEST (2019–2026) | Real quotes | Expected vs Actual'])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(cols))
    ws['A1'].font = Font(bold=True, color='FFFFFF', size=12); ws['A1'].fill = PatternFill('solid', fgColor='1B365D')
    ws['A1'].alignment = Alignment(horizontal='center')
    ws.append(cols)
    for c in range(1, len(cols)+1):
        cell = ws.cell(row=2, column=c); cell.fill = hdr_fill; cell.font = hdr_font; cell.border = thin
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
    for _, r in df.iterrows():
        ws.append([r[c] for c in cols])
        rr = ws.max_row
        act = r['Actual Ret %']
        fill = win_fill if act > 0 else loss_fill
        for c in range(1, len(cols)+1):
            ws.cell(row=rr, column=c).border = thin
        ws.cell(row=rr, column=cols.index('Actual Ret %')+1).fill = fill
    # summary sheet
    ws2 = wb.create_sheet('Summary')
    summ = [
        ('Total Trades', len(df)),
        ('Actual Win Rate %', round((df['Actual Ret %']>0).mean()*100, 1)),
        ('Actual Avg Ret %', round(df['Actual Ret %'].mean(), 2)),
        ('Actual Total Ret %', round(df['Actual Ret %'].sum(), 1)),
        ('Expected Win Rate %', round((df['Expected Ret %']>0).mean()*100, 1)),
        ('Expected Avg Ret %', round(df['Expected Ret %'].mean(), 2)),
        ('Expected Total Ret %', round(df['Expected Ret %'].sum(), 1)),
        ('Mean Expected-Actual Gap %', round(df['Exp-Act Gap %'].mean(), 2)),
    ]
    ws2.append(['Metric', 'Value'])
    for k, v in summ: ws2.append([k, v])
    for cc in range(1, 3):
        ws2.cell(row=1, column=cc).fill = hdr_fill; ws2.cell(row=1, column=cc).font = hdr_font
    widths = [12,12,7,8,12,42,12,12,12,12,8,12,14,14,13,14,14,13,13]
    for i, w in enumerate(widths[:len(cols)], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = 'A3'
    wb.save(OUT_XLSX)
    print(f'\nSaved: {OUT_XLSX}')
else:
    print('No trades produced — check data span/anchors.')
