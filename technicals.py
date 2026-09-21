"""
technicals.py
===============================================================================
Technical indicators computed from the OHLCV already on disk.

TWO PATHS, ONE DEFINITION
-------------------------
The dashboard wants indicators for ONE stock as a pandas Series. event_ml.py
wants them for EVERY stock on EVERY day as a GPU tensor. Writing RSI twice is
how the two silently drift apart, so each indicator is defined once here and
exposed through both:

    indicators(df)              -> dict of pandas Series   (dashboard, one stock)
    rsi_t(prices, n)  etc.      -> torch tensors           (event_ml, the panel)

test_technicals.py asserts the two agree to 1e-4. If you add an indicator, add
it to both and to that test.

WHY NOT pandas-ta / TA-Lib
--------------------------
pandas-ta is currently split across several competing forks with no clear
maintainer, and TA-Lib needs a C toolchain on Windows. This file is ~200 lines
of numpy for the dozen indicators actually used, and adds no dependency to a
project that deliberately runs on pandas + torch.

DATA
    processed/raw_price_cache/<SYM>.csv   OHLCV, split-adjusted, NOT dividend
                                          adjusted (results_dividend_behaviour.
                                          fetch_raw_prices)
    processed/price_cache/<SYM>.csv       adjusted close only — used for
                                          return-based work elsewhere

Wilder's smoothing (RSI, ATR, ADX) is an EWMA with alpha = 1/n, which is what
every charting package means by "14-period RSI". Using a simple mean instead
gives visibly different numbers and is the usual source of "my RSI doesn't match
TradingView".
===============================================================================
"""

from __future__ import annotations

import sys
from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parent
RAW_CACHE = ROOT / "processed" / "raw_price_cache"

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:                                 # noqa: BLE001
    pass


# --------------------------------------------------------------------------- #
# Loading
# --------------------------------------------------------------------------- #
def load_ohlcv(sym: str, fetch: bool = True) -> pd.DataFrame | None:
    """OHLCV for one symbol.

    Cache reads go through results_dividend_behaviour.fetch_raw_prices rather
    than being reimplemented here — it owns the cache format and knows to
    re-fetch files written before `volume` was added. Duplicating the read here
    is what silently left OBV and vol_z as NaN on the first pass.
    """
    import results_dividend_behaviour as RDB
    if not fetch and not (RAW_CACHE / f"{sym}.csv").exists():
        return None
    return RDB.fetch_raw_prices(sym)


# --------------------------------------------------------------------------- #
# Primitives
# --------------------------------------------------------------------------- #
def wilder(s: pd.Series, n: int) -> pd.Series:
    """Wilder's smoothing — the EWMA every charting package means by 'n-period'."""
    return s.ewm(alpha=1.0 / n, adjust=False, min_periods=n).mean()


def rsi(close: pd.Series, n: int = 14) -> pd.Series:
    d = close.diff()
    gain = wilder(d.clip(lower=0), n)
    loss = wilder((-d).clip(lower=0), n)
    rs = gain / loss.replace(0, np.nan)
    out = 100 - 100 / (1 + rs)
    return out.where(loss != 0, 100.0)            # no losses in the window -> 100


def macd(close: pd.Series, fast: int = 12, slow: int = 26, sig: int = 9):
    ef = close.ewm(span=fast, adjust=False).mean()
    es = close.ewm(span=slow, adjust=False).mean()
    line = ef - es
    signal = line.ewm(span=sig, adjust=False).mean()
    return line, signal, line - signal


def true_range(df: pd.DataFrame) -> pd.Series:
    pc = df["close"].shift(1)
    return pd.concat([df["high"] - df["low"],
                      (df["high"] - pc).abs(),
                      (df["low"] - pc).abs()], axis=1).max(axis=1)


def atr(df: pd.DataFrame, n: int = 14) -> pd.Series:
    return wilder(true_range(df), n)


def adx(df: pd.DataFrame, n: int = 14):
    """ADX plus the two directional indicators. Returns (adx, di_plus, di_minus)."""
    up = df["high"].diff()
    dn = -df["low"].diff()
    plus_dm = np.where((up > dn) & (up > 0), up, 0.0)
    minus_dm = np.where((dn > up) & (dn > 0), dn, 0.0)
    atr_n = wilder(true_range(df), n)
    di_p = 100 * wilder(pd.Series(plus_dm, index=df.index), n) / atr_n
    di_m = 100 * wilder(pd.Series(minus_dm, index=df.index), n) / atr_n
    dx = 100 * (di_p - di_m).abs() / (di_p + di_m).replace(0, np.nan)
    return wilder(dx, n), di_p, di_m


def bollinger(close: pd.Series, n: int = 20, k: float = 2.0):
    """Returns (%B, bandwidth). %B is where price sits inside the band: 0 = lower,
    1 = upper, outside the band goes beyond that range."""
    ma = close.rolling(n).mean()
    sd = close.rolling(n).std(ddof=0)
    upper, lower = ma + k * sd, ma - k * sd
    width = (upper - lower).replace(0, np.nan)
    return (close - lower) / width, (upper - lower) / ma


def obv(df: pd.DataFrame) -> pd.Series:
    if "volume" not in df.columns:
        return pd.Series(np.nan, index=df.index)
    sign = np.sign(df["close"].diff()).fillna(0.0)
    return (sign * df["volume"].fillna(0)).cumsum()


def slope_pct(s: pd.Series, n: int) -> pd.Series:
    """Change over n bars, normalised by the magnitude at the start of the window
    (OBV can be negative, so a plain pct_change is meaningless)."""
    prev = s.shift(n)
    return (s - prev) / prev.abs().replace(0, np.nan)


def zscore(s: pd.Series, n: int) -> pd.Series:
    m = s.rolling(n).mean()
    sd = s.rolling(n).std(ddof=0).replace(0, np.nan)
    return (s - m) / sd


def aroon(df: pd.DataFrame, n: int = 25):
    hi = df["high"].rolling(n + 1).apply(lambda w: float(np.argmax(w)), raw=True)
    lo = df["low"].rolling(n + 1).apply(lambda w: float(np.argmin(w)), raw=True)
    return 100 * hi / n, 100 * lo / n             # (aroon_up, aroon_down)


# --------------------------------------------------------------------------- #
# The full indicator set for one stock
# --------------------------------------------------------------------------- #
def indicators(df: pd.DataFrame) -> dict[str, pd.Series]:
    """Every indicator as a dated Series. Input needs at least a `close` column;
    high/low/volume unlock ADX/ATR/Aroon/OBV and are NaN-filled without them."""
    c = df["close"].astype(float)
    has_hl = {"high", "low"} <= set(df.columns)
    out: dict[str, pd.Series] = {}

    out["close"] = c
    out["rsi_14"] = rsi(c, 14)
    line, signal, hist = macd(c)
    out["macd"], out["macd_signal"], out["macd_hist"] = line, signal, hist

    for n in (20, 50, 200):
        out[f"sma_{n}"] = c.rolling(n).mean()
    out["above_ma50"] = (c > out["sma_50"]).astype(float)
    out["golden_cross"] = (out["sma_50"] > out["sma_200"]).astype(float)

    pctb, bw = bollinger(c, 20, 2.0)
    out["bb_pctb"], out["bb_width"] = pctb, bw

    for n in (5, 20, 60):
        out[f"mom_{n}"] = c.pct_change(n)
    r = c.pct_change()
    for n in (20, 60):
        out[f"vol_{n}"] = r.rolling(n).std(ddof=0)

    hi252 = c.rolling(252).max()
    lo252 = c.rolling(252).min()
    out["dist_52w_high"] = c / hi252 - 1.0
    out["dist_52w_low"] = c / lo252 - 1.0

    if has_hl:
        a, di_p, di_m = adx(df, 14)
        out["adx_14"], out["di_plus"], out["di_minus"] = a, di_p, di_m
        out["atr_pct"] = atr(df, 14) / c
        up, dn = aroon(df, 25)
        out["aroon_up"], out["aroon_down"] = up, dn
    else:
        for k in ("adx_14", "di_plus", "di_minus", "atr_pct",
                  "aroon_up", "aroon_down"):
            out[k] = pd.Series(np.nan, index=df.index)

    o = obv(df)
    out["obv"] = o
    out["obv_slope_20"] = slope_pct(o, 20)
    out["vol_z_20"] = (zscore(df["volume"].astype(float), 20)
                       if "volume" in df.columns
                       else pd.Series(np.nan, index=df.index))
    return out


def latest(sym: str, fetch: bool = True) -> dict | None:
    """The most recent value of every indicator — what the dashboard renders."""
    df = load_ohlcv(sym, fetch=fetch)
    if df is None or len(df) < 30:
        return None
    ind = indicators(df)
    out = {"symbol": sym, "date": str(df.index[-1].date()), "n_bars": int(len(df))}
    for k, s in ind.items():
        v = s.iloc[-1]
        out[k] = None if pd.isna(v) else round(float(v), 6)
    return out


# --------------------------------------------------------------------------- #
# torch path — same definitions over the (symbols x days) panel event_ml uses
# --------------------------------------------------------------------------- #
def _wilder_t(x, n: int):
    """Wilder smoothing along dim=1 of a (S, D) tensor. Sequential by
    construction — an EWMA is a recurrence — but it runs over all symbols at
    once, so it is S-way parallel and D iterations deep, not S*D."""
    import torch
    a = 1.0 / n
    out = torch.zeros_like(x)
    acc = x[:, :n].mean(dim=1)
    out[:, :n] = acc.unsqueeze(1)
    for t in range(n, x.shape[1]):
        acc = acc + a * (x[:, t] - acc)
        out[:, t] = acc
    return out


def rsi_t(close, n: int = 14):
    """RSI over a (symbols, days) tensor — matches rsi() to floating error."""
    import torch
    d = torch.zeros_like(close)
    d[:, 1:] = close[:, 1:] - close[:, :-1]
    gain = _wilder_t(d.clamp(min=0), n)
    loss = _wilder_t((-d).clamp(min=0), n)
    rs = gain / loss.clamp(min=1e-12)
    out = 100 - 100 / (1 + rs)
    return torch.where(loss <= 0, torch.full_like(out, 100.0), out)


def _ema_t(x, span: int):
    import torch
    a = 2.0 / (span + 1)
    out = torch.zeros_like(x)
    acc = x[:, 0]
    out[:, 0] = acc
    for t in range(1, x.shape[1]):
        acc = acc + a * (x[:, t] - acc)
        out[:, t] = acc
    return out


def macd_hist_t(close, fast: int = 12, slow: int = 26, sig: int = 9):
    line = _ema_t(close, fast) - _ema_t(close, slow)
    return line - _ema_t(line, sig)


def bb_pctb_t(close, n: int = 20, k: float = 2.0):
    """Bollinger %B over a (symbols, days) tensor."""
    import torch
    S, D = close.shape
    pad = torch.nn.functional.pad(close.unsqueeze(1), (n - 1, 0), mode="replicate")
    win = pad.squeeze(1).unfold(1, n, 1)          # (S, D, n)
    ma = win.mean(dim=2)
    sd = win.std(dim=2, unbiased=False)
    lower = ma - k * sd
    width = (2 * k * sd).clamp(min=1e-12)
    out = (close - lower) / width
    out[:, :n - 1] = float("nan")
    return out


def adx_t(high, low, close, n: int = 14):
    """ADX over (symbols, days) tensors — matches adx() to floating error."""
    import torch
    up = torch.zeros_like(high)
    dn = torch.zeros_like(low)
    up[:, 1:] = high[:, 1:] - high[:, :-1]
    dn[:, 1:] = low[:, :-1] - low[:, 1:]
    plus = torch.where((up > dn) & (up > 0), up, torch.zeros_like(up))
    minus = torch.where((dn > up) & (dn > 0), dn, torch.zeros_like(dn))

    pc = torch.zeros_like(close)
    pc[:, 1:] = close[:, :-1]
    pc[:, 0] = close[:, 0]
    tr = torch.maximum(high - low,
                       torch.maximum((high - pc).abs(), (low - pc).abs()))
    atr_n = _wilder_t(tr, n).clamp(min=1e-12)
    di_p = 100 * _wilder_t(plus, n) / atr_n
    di_m = 100 * _wilder_t(minus, n) / atr_n
    dx = 100 * (di_p - di_m).abs() / (di_p + di_m).clamp(min=1e-12)
    return _wilder_t(dx, n), di_p, di_m


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="print the latest indicators")
    ap.add_argument("symbols", nargs="*", default=["RELIANCE"])
    args = ap.parse_args()
    for sym in (args.symbols or ["RELIANCE"]):
        d = latest(sym)
        if not d:
            print(f"{sym}: no price data")
            continue
        print(f"\n{sym}  ({d['date']}, {d['n_bars']:,} bars)")
        for k in ("close", "rsi_14", "macd_hist", "adx_14", "atr_pct", "bb_pctb",
                  "mom_20", "mom_60", "dist_52w_high", "above_ma50",
                  "golden_cross", "obv_slope_20", "vol_z_20"):
            print(f"   {k:<16} {d.get(k)}")
