"""
Build an ADJUSTED copy of Stock_Data_213/daily_ohlc  ->  Stock_Data_213_Adjusted/daily_ohlc  (originals untouched).
Rules (every change is written to Adjustment_Log.csv):
  1. drop rows with non-positive prices or Close outside High/Low (BAD_ROW)
  2. a >35% close-to-close move in Adj Close is inspected:
       - price returns to the pre-jump level (within 35%) through a sharp >35% jump back, within 300 bars -> the stretch is a bad bar/segment:
            <=5 bars: DROP_SPIKE (rows removed);  longer & internally sane: RESCALE_SEGMENT (stretch scaled to join both sides);
            longer & not sane: DROP_SEGMENT
       - no return, and the day has a recorded Stock Split or zero volume -> RESCALE_HISTORY (earlier rows scaled to join: unadjusted split/demerger) -- also any move >60%
       - no return otherwise (35-60%) -> KEEP_PERSISTENT (treated as real, +-10 bars flagged WEAK)
       - from 2010 a traded (volume>0) move <60% with no split is kept as real (KEEP_REAL)
  3. Quality column: GOOD (2010+), OK (2003-2009), WEAK (before 2003, or any row touched by a fix, or in a 6+ identical-close run)
FINNIFTY / MIDCPNIFTY excluded (no usable data).
"""
import sys, glob, os, pathlib
import numpy as np, pandas as pd
sys.stdout.reconfigure(encoding='utf-8')
ROOT = pathlib.Path('D:/behaviour analysis')
SRC = ROOT / 'Stock_Data_213' / 'daily_ohlc'
DST = ROOT / 'Stock_Data_213_Adjusted' / 'daily_ohlc'
DST.mkdir(parents=True, exist_ok=True)
PX = ['Open', 'High', 'Low', 'Close', 'Adj Close']
JUMP, WINDOW, BIG = 0.35, 300, 0.60
MANUAL = {('TMPV', '2025-10-14')}   # Tata Motors demerger (known corporate action Yahoo did not adjust)
REAL_FROM = pd.Timestamp('2010-01-01')   # from here a traded (volume>0) big move is treated as real
SKIP = {'FINNIFTY', 'MIDCPNIFTY'}

def clean(sym, d):
    log = []
    def L(act, date, **kw): log.append(dict(Symbol=sym, Action=act, Date=pd.Timestamp(date).date(), **kw))
    d = d.sort_values('Date').drop_duplicates('Date').reset_index(drop=True)
    bad = (d[PX] <= 0).any(axis=1) | d[PX].isna().any(axis=1) | (d.Close > d.High * 1.001) | (d.Close < d.Low * 0.999)
    for x in d[bad].itertuples(): L('BAD_ROW', x.Date, Detail=f'close={x.Close:.4f}')
    d = d[~bad].reset_index(drop=True)
    d['flag'] = ''
    pos = 1
    for _ in range(5000):
        a = d['Adj Close'].values; n = len(d)
        mv = np.abs(a[pos:] / a[pos - 1:-1] - 1) if pos < n else np.array([])
        hit = np.where(mv > JUMP)[0]
        if len(hit) == 0: break
        i = pos + int(hit[0]); prev = a[i - 1]
        if d.Date[i] >= REAL_FROM and d.Volume[i] > 0 and d['Stock Splits'][i] == 0 and abs(a[i] / prev - 1) < BIG and (sym, str(d.Date[i].date())) not in MANUAL:
            L('KEEP_REAL', d.Date[i], Detail=f'move {a[i]/prev-1:+.0%} on volume {int(d.Volume[i])}, kept')
            d.loc[max(i-10,0):i+10, 'flag'] += 'K'; pos = i + 1; continue
        k_hi = min(i + 1 + WINDOW, n)
        seg = a[i + 1:k_hi]
        back = (np.abs(seg / prev - 1) < JUMP) & (np.abs(seg / a[i:k_hi - 1] - 1) > JUMP)   # returns THROUGH a sharp jump, not by drifting
        rng = back
        if rng.any():
            j = i + 1 + int(np.argmax(rng)); Ls = j - i
            if Ls <= 5:
                L('DROP_SPIKE', d.Date[i], Detail=f'{Ls} bar(s), move {a[i]/prev-1:+.0%}')
                d = d.drop(index=range(i, j)).reset_index(drop=True)
            else:
                inner = a[i + 1:j] / a[i:j - 1] - 1 if j - i > 1 else np.array([0])
                if np.abs(inner).max() < JUMP:
                    f = prev / a[i]
                    d.loc[i:j - 1, PX] = d.loc[i:j - 1, PX].values * f
                    d.loc[i:j - 1, 'flag'] = 'R'
                    L('RESCALE_SEGMENT', d.Date[i], Detail=f'{Ls} bars to {pd.Timestamp(d.Date[j-1]).date()}, factor {f:.4f}')
                    pos = i + 1
                else:
                    L('DROP_SEGMENT', d.Date[i], Detail=f'{Ls} bars')
                    d = d.drop(index=range(i, j)).reset_index(drop=True)
        else:
            m = a[i] / prev - 1
            if d['Stock Splits'][i] > 0 or d.Volume[i] == 0 or abs(m) > BIG or (sym, str(d.Date[i].date())) in MANUAL:
                f = a[i] / prev
                d.loc[:i - 1, PX] = d.loc[:i - 1, PX].values * f
                d.loc[max(i - 10, 0):i + 10, 'flag'] += 'H'
                L('RESCALE_HISTORY', d.Date[i], Detail=f'move {m:+.0%}; history before scaled x{f:.4f} (split/demerger/basis break)')
            else:
                d.loc[max(i - 10, 0):i + 10, 'flag'] += 'K'
                L('KEEP_PERSISTENT', d.Date[i], Detail=f'move {m:+.0%} kept as real (flagged)')
            pos = i + 1
    # stale runs
    run = d.Close.groupby((d.Close != d.Close.shift()).cumsum()).transform('size')
    stale = (run >= 6) & (d.Date >= '1996-01-01')
    d.loc[stale, 'flag'] += 'S'
    yr = d.Date.dt.year
    q = np.where(yr >= 2010, 'GOOD', np.where(yr >= 2003, 'OK', 'WEAK'))
    d['Quality'] = np.where(d.flag != '', 'WEAK', q)
    return d.drop(columns='flag'), log

def main():
    logs, summ = [], []
    for f in sorted(glob.glob(str(SRC / '*.csv'))):
        s = os.path.basename(f)[:-4]
        if s in SKIP: continue
        d = pd.read_csv(f, parse_dates=['Date'])
        n0 = len(d); out, lg = clean(s, d)
        cols = ['Symbol', 'YahooTicker', 'Date'] + PX + ['Volume', 'Dividends', 'Stock Splits', 'Quality']
        out[cols].to_csv(DST / f'{s}.csv', index=False)
        logs += lg
        summ.append(dict(Symbol=s, RowsBefore=n0, RowsAfter=len(out), Fixes=len(lg)))
    LG = pd.DataFrame(logs); LG.to_csv(ROOT / 'Stock_Data_213_Adjusted' / 'Adjustment_Log.csv', index=False)
    S = pd.DataFrame(summ); S.to_csv(ROOT / 'Stock_Data_213_Adjusted' / 'Adjustment_Summary.csv', index=False)
    print('symbols', len(S), '| rows before', S.RowsBefore.sum(), 'after', S.RowsAfter.sum(), '| logged actions', len(LG))
    print(LG.Action.value_counts().to_string())
    print('stocks touched:', LG.Symbol.nunique())
    print(LG[LG.Action == 'KEEP_PERSISTENT'].assign(Y=lambda x: pd.to_datetime(x.Date).dt.year).groupby('Y').size().to_string())

if __name__ == '__main__':
    main()
