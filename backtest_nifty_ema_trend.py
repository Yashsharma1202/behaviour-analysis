"""
BEST EMA TREND-FOLLOWING COMBINATION ON NIFTY 50 (2010-2026).
Tests a grid of fast/slow EMA pairs. Rule: LONG (in-market) when fast EMA >
slow EMA, flat (out) when fast EMA < slow EMA \u2014 signal taken the day AFTER
the crossover (no lookahead). Ranked by CAGR, Sharpe, Max Drawdown, number of
whipsaw trades, and vs buy-and-hold. The winning pair's CURRENT state (as of
the latest close) becomes the "Nifty Trend" regime filter for the market-
sentiment feature.

ADDITIVE: reads existing price data only; writes nothing (analysis script).
"""
import pandas as pd
import numpy as np

df = pd.read_parquet(r'D:\behaviour analysis\scraped_parquet\nifty_spot_daily.parquet')
df = df.sort_values('DATE').reset_index(drop=True)
close = df['CLOSE']
dates = df['DATE']

FAST = [5, 8, 9, 10, 12, 15, 20]
SLOW = [20, 21, 26, 34, 50, 100, 200]

def backtest(fast_n, slow_n):
    ema_f = close.ewm(span=fast_n, adjust=False).mean()
    ema_s = close.ewm(span=slow_n, adjust=False).mean()
    signal = (ema_f > ema_s).astype(int)          # 1 = long regime, 0 = flat
    signal = signal.shift(1).fillna(0)              # act next day, no lookahead
    daily_ret = close.pct_change().fillna(0)
    strat_ret = daily_ret * signal
    # skip split/data-glitch days (>20% single-day move) same convention as the rest of the project
    strat_ret = strat_ret.where(daily_ret.abs() <= 0.20, 0)

    n_years = (dates.iloc[-1] - dates.iloc[0]).days / 365.25
    cum = (1 + strat_ret).prod()
    cagr = (cum ** (1 / n_years) - 1) * 100
    ann_vol = strat_ret.std() * np.sqrt(252)
    ann_ret = strat_ret.mean() * 252
    sharpe = ann_ret / ann_vol if ann_vol > 0 else 0

    eq = (1 + strat_ret).cumprod()
    peak = eq.cummax()
    dd = (eq - peak) / peak
    max_dd = dd.min() * 100

    trades = int((signal.diff().abs() == 1).sum())
    days_in_market = int(signal.sum())
    pct_in_market = days_in_market / len(signal) * 100

    return dict(fast=fast_n, slow=slow_n, cagr=cagr, sharpe=sharpe, max_dd=max_dd,
                trades=trades, pct_in_market=pct_in_market, cum_x=cum)

# Buy & hold benchmark
bh_ret = close.pct_change().fillna(0)
bh_ret = bh_ret.where(bh_ret.abs() <= 0.20, 0)
n_years = (dates.iloc[-1] - dates.iloc[0]).days / 365.25
bh_cum = (1 + bh_ret).prod()
bh_cagr = (bh_cum ** (1 / n_years) - 1) * 100
bh_vol = bh_ret.std() * np.sqrt(252)
bh_sharpe = (bh_ret.mean() * 252) / bh_vol
bh_eq = (1 + bh_ret).cumprod(); bh_dd = (bh_eq - bh_eq.cummax()) / bh_eq.cummax()
print(f"BUY & HOLD BENCHMARK: CAGR={bh_cagr:.2f}% Sharpe={bh_sharpe:.2f} MaxDD={bh_dd.min()*100:.1f}%")
print("=" * 100)

results = []
for f in FAST:
    for s in SLOW:
        if f >= s:
            continue
        results.append(backtest(f, s))

results.sort(key=lambda r: -r['sharpe'])
print(f"{'Fast':<6}{'Slow':<6}{'CAGR%':<9}{'Sharpe':<9}{'MaxDD%':<9}{'Trades':<8}{'%InMkt':<8}")
for r in results[:15]:
    print(f"{r['fast']:<6}{r['slow']:<6}{r['cagr']:<9.2f}{r['sharpe']:<9.2f}{r['max_dd']:<9.1f}{r['trades']:<8}{r['pct_in_market']:<8.1f}")

best = results[0]
print()
print(f"BEST BY SHARPE: EMA({best['fast']}) / EMA({best['slow']})  CAGR={best['cagr']:.2f}% Sharpe={best['sharpe']:.2f} MaxDD={best['max_dd']:.1f}% Trades={best['trades']}")

# Current state of the best combination
ema_f = close.ewm(span=best['fast'], adjust=False).mean()
ema_s = close.ewm(span=best['slow'], adjust=False).mean()
cur_f, cur_s = ema_f.iloc[-1], ema_s.iloc[-1]
cur_close = close.iloc[-1]
cur_date = dates.iloc[-1]
trend = 'BULLISH (fast > slow)' if cur_f > cur_s else 'BEARISH (fast < slow)'
gap_pct = (cur_f - cur_s) / cur_s * 100
print()
print(f"CURRENT STATE as of {cur_date.date()}: Close={cur_close:.2f} EMA{best['fast']}={cur_f:.2f} EMA{best['slow']}={cur_s:.2f}")
print(f"NIFTY TREND REGIME: {trend}  (gap {gap_pct:+.2f}%)")

# How many days since the last crossover
signal = (ema_f > ema_s).astype(int)
last_val = signal.iloc[-1]
days_since_cross = 0
for v in signal.iloc[::-1]:
    if v != last_val:
        break
    days_since_cross += 1
print(f"Days in current regime: {days_since_cross}")
