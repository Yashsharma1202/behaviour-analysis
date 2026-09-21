"""
test_technicals.py
===============================================================================
Guards the one thing that can quietly break technicals.py: the pandas path (the
dashboard, one stock) and the torch path (event_ml.py, the whole panel) drifting
apart. If RSI means two different things in two files, the ML model trains on
features the UI never shows, and nobody notices.

Also sanity-checks each indicator against its definition — RSI in [0,100],
%B ~ 0 at the lower band, ATR% positive, and MACD histogram = line - signal.

    python test_technicals.py
    python test_technicals.py RELIANCE TCS INFY

Exit code 0 if every check passes.
===============================================================================
"""

from __future__ import annotations

import sys

import numpy as np
import pandas as pd

import technicals as T

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:                                 # noqa: BLE001
    pass

TOL = 1e-4
fails: list[str] = []


def check(name: str, ok: bool, detail: str = ""):
    print(f"  {'PASS' if ok else 'FAIL'}  {name}{'  ' + detail if detail else ''}")
    if not ok:
        fails.append(name)


def max_rel_diff(a: np.ndarray, b: np.ndarray, warmup: int) -> float:
    """Largest relative gap after the warm-up bars, ignoring NaN on either side."""
    a, b = a[warmup:], b[warmup:]
    m = ~(np.isnan(a) | np.isnan(b))
    if not m.any():
        return float("nan")
    denom = np.maximum(np.abs(b[m]), 1.0)
    return float(np.max(np.abs(a[m] - b[m]) / denom))


def main():
    syms = sys.argv[1:] or ["RELIANCE", "TCS", "INFY"]
    frames = {}
    for s in syms:
        df = T.load_ohlcv(s)
        if df is not None and len(df) > 400:
            frames[s] = df
    if not frames:
        raise SystemExit("no price data — run technicals.py once to populate the cache")

    print("=" * 74)
    print(f"  TECHNICALS — pandas vs torch parity  ({', '.join(frames)})")
    print("=" * 74)

    try:
        import torch
    except ImportError:
        print("  torch not installed — parity checks skipped")
        torch = None

    # ---- definitional sanity, pandas path --------------------------------- #
    print("\n  definitions")
    for s, df in frames.items():
        ind = T.indicators(df)
        r = ind["rsi_14"].dropna()
        check(f"{s}: RSI within [0,100]", bool(((r >= 0) & (r <= 100)).all()),
              f"min={r.min():.1f} max={r.max():.1f}")

        line, sig, hist = ind["macd"], ind["macd_signal"], ind["macd_hist"]
        d = (hist - (line - sig)).abs().max()
        check(f"{s}: MACD hist == line - signal", d < 1e-9, f"maxdiff={d:.2e}")

        a = ind["atr_pct"].dropna()
        check(f"{s}: ATR% positive", bool((a > 0).all()))

        adxv = ind["adx_14"].dropna()
        check(f"{s}: ADX within [0,100]", bool(((adxv >= 0) & (adxv <= 100)).all()),
              f"last={adxv.iloc[-1]:.1f}")

        # %B must be ~0 at the lower band and ~1 at the upper band, by definition.
        c = df["close"].astype(float)
        ma, sd = c.rolling(20).mean(), c.rolling(20).std(ddof=0)
        pctb = ind["bb_pctb"]
        at_lo = ((c - (ma - 2 * sd)).abs() < 1e-9)
        ok = True if not at_lo.any() else bool((pctb[at_lo].abs() < 1e-6).all())
        check(f"{s}: %B == 0 on the lower band", ok)

        dh = ind["dist_52w_high"].dropna()
        check(f"{s}: dist_52w_high <= 0", bool((dh <= 1e-12).all()),
              f"max={dh.max():.2e}")

    # ---- pandas vs torch --------------------------------------------------- #
    if torch is not None:
        print("\n  pandas vs torch (relative, after warm-up)")
        close = torch.tensor(
            np.stack([frames[s]["close"].to_numpy(dtype=np.float64) [-1200:]
                      for s in frames]), dtype=torch.float64)
        high = torch.tensor(
            np.stack([frames[s]["high"].to_numpy(dtype=np.float64)[-1200:]
                      for s in frames]), dtype=torch.float64)
        low = torch.tensor(
            np.stack([frames[s]["low"].to_numpy(dtype=np.float64)[-1200:]
                      for s in frames]), dtype=torch.float64)

        rsi_t = T.rsi_t(close, 14).numpy()
        mh_t = T.macd_hist_t(close).numpy()
        bb_t = T.bb_pctb_t(close).numpy()
        adx_t, _dip, _dim = T.adx_t(high, low, close, 14)
        adx_t = adx_t.numpy()

        for i, s in enumerate(frames):
            c = frames[s]["close"].iloc[-1200:]
            sub = frames[s].iloc[-1200:]
            # 200 bars of warm-up: both paths seed their recurrences differently
            # at the very start, which is expected and irrelevant downstream.
            d = max_rel_diff(rsi_t[i], T.rsi(c, 14).to_numpy(), 200)
            check(f"{s}: RSI parity", d < TOL, f"maxrel={d:.2e}")
            d = max_rel_diff(mh_t[i], T.macd(c)[2].to_numpy(), 200)
            check(f"{s}: MACD-hist parity", d < TOL, f"maxrel={d:.2e}")
            d = max_rel_diff(bb_t[i], T.bollinger(c)[0].to_numpy(), 200)
            check(f"{s}: %B parity", d < TOL, f"maxrel={d:.2e}")
            d = max_rel_diff(adx_t[i], T.adx(sub, 14)[0].to_numpy(), 200)
            check(f"{s}: ADX parity", d < TOL, f"maxrel={d:.2e}")

    print("\n" + "=" * 74)
    if fails:
        print(f"  {len(fails)} CHECK(S) FAILED: {', '.join(fails)}")
        print("=" * 74)
        return 1
    print("  ALL CHECKS PASSED")
    print("=" * 74)
    return 0


if __name__ == "__main__":
    sys.exit(main())
