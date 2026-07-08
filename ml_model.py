"""
ml_model.py   (v3 — enriched with the full Screener fundamentals)
===============================================================================
Predict whether the "buy 6 days before results, hold 4 days" trade profits, and
test whether adding RICH FUNDAMENTALS (from fund_loader) helps a model that
previously used only price features.

FEATURE SETS COMPARED
---------------------
  PRICE      (8)  : mom_20, mom_60, vol_20, mkt_mom_20, rel_strength,
                    dist_52w_high, above_ma50, month
  +FUND     (+13) : last-known quarter YoY profit/sales growth & OPM;
                    latest annual sales-growth, profit-growth, OPM, div-payout;
                    ROCE, cash-conversion-cycle, debt/equity, FCF-positive,
                    CFO/OP ; plus is_annual result flag.
                    (All "as of" the entry day -> no look-ahead.)

THE HONEST HYPOTHESIS
---------------------
With only ~48 training rows, adding ~13 columns may HURT (curse of
dimensionality): the model gets more knobs to overfit but no more examples to
learn from. We compare test accuracy PRICE vs +FUND to see it happen -- the real
fix is more ROWS (multi-stock), not more columns.

Baseline to beat: "always predict WIN" (~62% because most trades win).

    python ml_model.py

Needs internet (Yahoo) + scikit-learn + the fundamental folders.
===============================================================================
"""

from __future__ import annotations

import json
import ssl
import sys
import time
import urllib.request
import warnings
from pathlib import Path

import numpy as np
import pandas as pd

import fund_loader

warnings.filterwarnings("ignore")
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:                              # noqa: BLE001
    pass

from sklearn.ensemble import RandomForestClassifier          # noqa: E402
from sklearn.linear_model import LogisticRegression          # noqa: E402
from sklearn.neural_network import MLPClassifier             # noqa: E402
from sklearn.preprocessing import StandardScaler             # noqa: E402

ROOT = Path(__file__).resolve().parent
PROC = ROOT / "processed"
STOCK, MARKET = "RELIANCE.NS", "^NSEI"
ENTER, HOLD, COST = 6, 4, 0.001

_HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
_SSL = ssl.create_default_context()
_SSL.check_hostname = False
_SSL.verify_mode = ssl.CERT_NONE

PRICE_FEATS = ["mom_20", "mom_60", "vol_20", "mkt_mom_20", "rel_strength",
               "dist_52w_high", "above_ma50", "month"]
FUND_FEATS = ["is_annual", "q_prof_yoy", "q_sales_yoy", "q_opm",
              "a_sales_growth", "a_profit_growth", "a_opm", "a_div_payout",
              "r_roce", "r_ccc", "b_debt_equity", "c_fcf_pos", "c_cfo_op"]


def fetch_adj(ticker: str) -> pd.Series:
    url = (f"https://query1.finance.yahoo.com/v8/finance/chart/{ticker}"
           f"?period1=1167609600&period2={int(time.time())}&interval=1d")
    req = urllib.request.Request(url, headers=_HEADERS)
    with urllib.request.urlopen(req, timeout=30, context=_SSL) as r:
        d = json.load(r)
    res = d["chart"]["result"][0]
    ind = res["indicators"]
    adj = ind.get("adjclose", [{}])[0].get("adjclose") or ind["quote"][0]["close"]
    return pd.Series(adj, index=pd.to_datetime(res["timestamp"], unit="s").normalize()).dropna().sort_index()


def _col(df, name):
    """Find a cleaned column by exact/contains match, else None."""
    for c in df.columns:
        if c.strip() == name:
            return c
    for c in df.columns:
        if name in c:
            return c
    return None


def asof(df, offset_days, d_in):
    """Latest row whose 'reported date' (period month-end + offset) <= entry day.
    Returns the row (Series) or None -> guarantees no look-ahead."""
    if df is None or df.empty:
        return None
    reported = df.index + pd.offsets.MonthEnd(0) + pd.Timedelta(days=offset_days)
    mask = reported <= d_in
    return df[mask].iloc[-1] if mask.any() else None


def fundamentals_features(fund, d_in, is_annual):
    """Build the fundamental feature dict using only data public before entry."""
    pnl, rat, bs, cf, q = (fund.get(k) for k in
                           ("pnl", "ratios", "balance_sheet", "cash_flow", "quarterly"))
    f = {k: 0.0 for k in FUND_FEATS}
    f["is_annual"] = float(is_annual)

    # last KNOWN quarter (reported ~45d after quarter end)
    qr = asof(q, 45, d_in)
    if qr is not None:
        f["q_prof_yoy"] = float(qr.get(_col(q, "YOY Profit Growth %") or "", 0) or 0)
        f["q_sales_yoy"] = float(qr.get(_col(q, "YOY Sales Growth %") or "", 0) or 0)
        f["q_opm"] = float(qr.get(_col(q, "OPM %") or "", 0) or 0)

    # latest annual (reported ~55d after FY end)
    pr = asof(pnl, 55, d_in)
    if pr is not None:
        f["a_sales_growth"] = float(pr.get(_col(pnl, "Sales Growth %") or "", 0) or 0)
        f["a_profit_growth"] = float(pr.get(_col(pnl, "Profit Growth %") or "", 0) or 0)
        f["a_opm"] = float(pr.get(_col(pnl, "OPM %") or "", 0) or 0)
        f["a_div_payout"] = float(pr.get(_col(pnl, "Dividend Payout %") or "", 0) or 0)
    rr = asof(rat, 55, d_in)
    if rr is not None:
        f["r_roce"] = float(rr.get(_col(rat, "ROCE %") or "", 0) or 0)
        f["r_ccc"] = float(rr.get(_col(rat, "Cash Conversion Cycle") or "", 0) or 0)
    br = asof(bs, 55, d_in)
    if br is not None:
        eq = float(br.get(_col(bs, "Equity Capital") or "", 0) or 0) + \
            float(br.get(_col(bs, "Reserves") or "", 0) or 0)
        bor = float(br.get(_col(bs, "Borrowings") or "", 0) or 0)
        f["b_debt_equity"] = bor / eq if eq else 0.0
    cr = asof(cf, 55, d_in)
    if cr is not None:
        f["c_fcf_pos"] = 1.0 if float(cr.get(_col(cf, "Free Cash Flow") or "", 0) or 0) > 0 else 0.0
        f["c_cfo_op"] = float(cr.get(_col(cf, "CFO/OP") or "", 0) or 0)
    return f


def build_dataset(px, mkt, events, fund):
    mkt = mkt.reindex(px.index).ffill()
    p, m = px.values, mkt.values
    rows = []
    for _, e in events.iterrows():
        ev = e["date"]
        pos = px.index.searchsorted(ev)
        i_in, i_out = pos - ENTER, pos - ENTER + HOLD
        if i_in - 60 < 0 or i_out >= len(px):
            continue
        d_in = px.index[i_in]
        hist_lo = max(0, i_in - 252)
        row = {
            "date": d_in,
            "mom_20": p[i_in] / p[i_in - 20] - 1,
            "mom_60": p[i_in] / p[i_in - 60] - 1,
            "vol_20": pd.Series(p[i_in - 20:i_in]).pct_change().std(),
            "mkt_mom_20": m[i_in] / m[i_in - 20] - 1,
            "rel_strength": (p[i_in] / p[i_in - 60] - 1) - (m[i_in] / m[i_in - 60] - 1),
            "dist_52w_high": p[i_in] / max(p[hist_lo:i_in + 1]) - 1,
            "above_ma50": 1.0 if p[i_in] > np.mean(p[i_in - 50:i_in]) else 0.0,
            "month": ev.month,
            "win": int((p[i_out] / p[i_in] - 1 - COST) > 0),
        }
        row.update(fundamentals_features(fund, d_in, e["is_annual"]))
        rows.append(row)
    return pd.DataFrame(rows).dropna().reset_index(drop=True)


def earnings_events():
    ev = pd.read_csv(PROC / "unified_events.csv", parse_dates=["date"])
    e = ev[ev["event_type"].str.startswith("RESULT_", na=False)].copy()
    e["date"] = e["date"].dt.normalize()
    e["is_annual"] = e["event_type"].str.contains("Annual").astype(int)
    return e.groupby("date")["is_annual"].max().reset_index().sort_values("date")


def run_models(data, feats, tag):
    X, y = data[feats].values, data["win"].values
    cut = int(len(data) * 0.70)
    Xtr, Xte, ytr, yte = X[:cut], X[cut:], y[:cut], y[cut:]
    sc = StandardScaler().fit(Xtr)
    Xtr_s, Xte_s = sc.transform(Xtr), sc.transform(Xte)
    models = {
        "LogReg": (LogisticRegression(C=0.5, max_iter=1000), Xtr_s, Xte_s),
        "Forest": (RandomForestClassifier(n_estimators=200, max_depth=3,
                   min_samples_leaf=3, random_state=0), Xtr, Xte),
        "NeuralNet": (MLPClassifier(hidden_layer_sizes=(64, 32), max_iter=3000,
                      random_state=0), Xtr_s, Xte_s),
    }
    out = {}
    for name, (mdl, a, b) in models.items():
        mdl.fit(a, ytr)
        out[name] = {"train": mdl.score(a, ytr) * 100, "test": mdl.score(b, yte) * 100}
    base = (yte == int(round(ytr.mean()))).mean() * 100
    return out, base, len(ytr), len(yte)


def main():
    print("Downloading prices + loading fundamentals ...")
    px, mkt = fetch_adj(STOCK), fetch_adj(MARKET)
    fund = fund_loader.load_stock("RELIANCE")
    events = earnings_events()
    data = build_dataset(px, mkt, events, fund).sort_values("date").reset_index(drop=True)
    data.to_csv(PROC / "ml_dataset.csv", index=False)

    price_res, base, ntr, nte = run_models(data, PRICE_FEATS, "price")
    full_res, _, _, _ = run_models(data, PRICE_FEATS + FUND_FEATS, "full")

    print(f"\nDataset: {len(data)} earnings events, train {ntr} / test {nte}")
    print(f"Baseline ('always WIN'): {base:.0f}% test accuracy  <-- beat this\n")

    print("=" * 72)
    print(f"  DOES ADDING FUNDAMENTALS HELP?  (test accuracy, {nte} recent events)")
    print("=" * 72)
    print(f"  {'model':10} {'PRICE only':>18} {'PRICE + FUNDAMENTALS':>22}")
    print(f"  {'':10} {'train':>8}{'test':>7}   {'train':>8}{'test':>7}   {'test Δ':>7}")
    print("-" * 72)
    for name in ("LogReg", "Forest", "NeuralNet"):
        pr, fu = price_res[name], full_res[name]
        delta = fu["test"] - pr["test"]
        print(f"  {name:10} {pr['train']:>7.0f}%{pr['test']:>6.0f}%   "
              f"{fu['train']:>7.0f}%{fu['test']:>6.0f}%   {delta:>+6.0f}pp")
    print("=" * 72)

    # feature importance on the full set
    X, y = data[PRICE_FEATS + FUND_FEATS].values, data["win"].values
    rf = RandomForestClassifier(n_estimators=400, max_depth=3, min_samples_leaf=3,
                                random_state=0).fit(X, y)
    imp = sorted(zip(PRICE_FEATS + FUND_FEATS, rf.feature_importances_),
                 key=lambda t: -t[1])[:10]
    print("\n  Top-10 feature importance (full model):")
    for f_, v in imp:
        kind = "fund" if f_ in FUND_FEATS else "price"
        print(f"    {f_:16} ({kind}) {'#'*int(v*70)} {v*100:.0f}%")

    # verdict
    best_price = max(price_res.values(), key=lambda r: r["test"])["test"]
    best_full = max(full_res.values(), key=lambda r: r["test"])["test"]
    print("\n  VERDICT:")
    print(f"    best PRICE-only test acc : {best_price:.0f}%")
    print(f"    best +FUNDAMENTALS       : {best_full:.0f}%   "
          f"(baseline {base:.0f}%)")
    if best_full <= best_price + 2:
        print("    -> Fundamentals did NOT improve out-of-sample accuracy. With ~48 train")
        print("       rows, 13 extra columns just add overfitting room, not signal.")
    else:
        print("    -> Fundamentals helped a little, but on ~20 test events treat with care.")
    print("    The features are GOOD (they'd matter across many stocks); the problem is")
    print("    rows. Same features on 3,400 stocks x ~50 quarters = the real experiment.")
    print(f"\n  Saved -> {PROC}\\ml_dataset.csv")


if __name__ == "__main__":
    main()
