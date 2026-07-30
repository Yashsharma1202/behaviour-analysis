"""
ml_predict.py
===============================================================================
POOLED, REGULARISED ML for event-trade outcomes.
-------------------------------------------------------------------------------
The earlier ml_model.py trained PER STOCK (~48 rows) and failed. This version
fixes the real problem: it POOLS every scheduled event across all 50 stocks into
one dataset (thousands of rows), engineers as-of-entry-day features (no
look-ahead), and trains REGULARISED models with a strict TIME-BASED split, so the
score is honest out-of-time — not curve-fitted.

TASK
    For each scheduled event (Results / Board meeting / Corporate action), define
    a standard trade — buy N days before, sell M days after — and predict whether
    it finishes positive (WIN=1 / LOSS=0).

MODELS (both regularised)
    • Logistic Regression  — L2 penalty (C small = strong regularisation), scaled
    • HistGradientBoosting — L2 leaf reg + shallow trees + early stopping

VALIDATION
    Time split: train on the older events, test on the most recent — never random.
    Baseline to beat: "always predict WIN" (= the WIN rate in the test period).

    python ml_predict.py
===============================================================================
"""
from __future__ import annotations

import json
import sys
import warnings
from pathlib import Path

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")
sys.stdout.reconfigure(encoding="utf-8")

import event_behaviour as EB
import fund_loader
from download_feeds import NIFTY50_FALLBACK

from sklearn.linear_model import LogisticRegression
from sklearn.ensemble import HistGradientBoostingClassifier
from sklearn.preprocessing import StandardScaler
from sklearn.impute import SimpleImputer
from sklearn.pipeline import make_pipeline
from sklearn.metrics import accuracy_score, roc_auc_score

# ── trade definition & knobs ────────────────────────────────────────────────
N_BEFORE = 5           # buy 5 trading days before the event
M_AFTER = 3            # sell 3 trading days after
EVENTS = ["RESULTS", "BOARD_MEETING", "CORPORATE_ACTION"]   # scheduled only
TEST_FROM = "2021-07-01"          # events on/after this date = out-of-time TEST set
MIN_HIST = 252                    # need 1y of prior prices for the 52w feature


def nifty() -> pd.Series | None:
    p = EB.PRICE_CACHE / "_NSEI.csv"
    if p.exists():
        return pd.read_csv(p, parse_dates=["date"]).set_index("date")["adj"].sort_index()
    return EB.fetch_prices_yahoo("^NSEI")


def features_for(px: pd.Series, nif: pd.Series, pos: int) -> dict | None:
    """As-of-entry-day features. Entry = close at pos-N_BEFORE; NOTHING after it."""
    i = pos - N_BEFORE
    if i - MIN_HIST < 0 or pos + M_AFTER >= len(px):
        return None
    w = px.iloc[:i + 1]                      # strictly up to & including entry day
    entry = float(w.iloc[-1])
    r = w.pct_change().dropna()
    mom20 = entry / float(w.iloc[-21]) - 1
    mom60 = entry / float(w.iloc[-61]) - 1
    vol20 = float(r.iloc[-20:].std())
    hi52 = float(w.iloc[-252:].max())
    ma50 = float(w.iloc[-50:].mean())
    # market relative strength on the entry date
    edate = w.index[-1]
    nw = nif.loc[:edate]
    nmom20 = (float(nw.iloc[-1]) / float(nw.iloc[-21]) - 1) if len(nw) > 21 else 0.0
    return {
        "mom20": mom20, "mom60": mom60, "vol20": vol20,
        "dist_52w_high": entry / hi52 - 1,
        "above_ma50": 1.0 if entry > ma50 else 0.0,
        "rel_strength": mom20 - nmom20,
        "month": edate.month,
    }


_FUND: dict = {}
FUND_LAG = pd.DateOffset(months=6)     # only use a fiscal year reported >=6 months before entry


def get_fund(sym: str):
    if sym not in _FUND:
        try:
            _FUND[sym] = fund_loader.load_stock(sym)
        except Exception:                              # noqa: BLE001
            _FUND[sym] = {}
    return _FUND[sym]


def fund_features(sym: str, edate: pd.Timestamp) -> dict:
    """Latest ANNUAL fundamentals known as-of the entry day (lagged 6m -> no lookahead)."""
    d = get_fund(sym)
    cut = edate - FUND_LAG
    out = {"f_roce": np.nan, "f_sales_g": np.nan, "f_opm": np.nan,
           "f_profit_g": np.nan, "f_de": np.nan, "f_payout": np.nan}

    def last(df):
        if df is None or df.empty:
            return None
        sub = df[df.index <= cut]
        return sub.iloc[-1] if len(sub) else None

    def num(row, col):
        if row is None or col not in row.index:
            return np.nan
        try:
            v = float(row[col]); return v if v == v else np.nan
        except (TypeError, ValueError):
            return np.nan

    p, r, b = last(d.get("pnl")), last(d.get("ratios")), last(d.get("balance_sheet"))
    out["f_roce"] = num(r, "ROCE %")
    out["f_sales_g"] = num(p, "Sales Growth %")
    out["f_opm"] = num(p, "OPM %")
    out["f_profit_g"] = num(p, "Profit Growth %")
    out["f_payout"] = num(p, "Dividend Payout %")
    if b is not None:
        bor = num(b, "Borrowings"); eq = num(b, "Equity Capital"); res = num(b, "Reserves")
        base = (eq if eq == eq else 0) + (res if res == res else 0)
        if bor == bor and base:
            out["f_de"] = bor / base
    return out


def build_dataset() -> pd.DataFrame:
    nif = nifty()
    if nif is None:
        raise SystemExit("no NIFTY price series")
    rows = []
    syms = sorted(set(NIFTY50_FALLBACK))
    for k, sym in enumerate(syms, 1):
        px = EB.fetch_prices_yahoo(sym)
        if px is None or len(px) < MIN_HIST + 30:
            continue
        ev = EB.load_events_from_feeds([sym])
        ev = ev[ev["event_type"].isin(EVENTS)]
        idx = px.index
        for _, e in ev.iterrows():
            pos = idx.searchsorted(pd.Timestamp(e["event_date"]))
            if pos <= 0 or pos >= len(idx):
                continue
            f = features_for(px, nif, pos)
            if f is None:
                continue
            entry = float(px.iloc[pos - N_BEFORE]); exit_ = float(px.iloc[pos + M_AFTER])
            f.update(fund_features(sym, idx[pos - N_BEFORE]))     # as-of entry-day fundamentals
            f.update({"symbol": sym, "date": idx[pos],
                      "etype": e["event_type"],
                      "win": int((exit_ / entry - 1) > 0)})
            rows.append(f)
        if k % 10 == 0:
            print(f"  built {k}/{len(syms)} stocks…")
    df = pd.DataFrame(rows)
    price_cols = ["mom20", "mom60", "vol20", "dist_52w_high", "above_ma50", "rel_strength", "month"]
    df = df.dropna(subset=price_cols)          # price features must be present; fundamentals may be NaN
    df = pd.concat([df, pd.get_dummies(df["etype"], prefix="ev")], axis=1)
    return df


def evaluate(name, Xtr, ytr, Xte, yte, base_acc):
    """Train both regularised models on a feature set, return metrics + the logistic pipe."""
    lr = make_pipeline(SimpleImputer(strategy="median"), StandardScaler(),
                       LogisticRegression(C=0.1, penalty="l2", class_weight="balanced",
                                          max_iter=2000)).fit(Xtr, ytr)
    pl = lr.predict(Xte); pp = lr.predict_proba(Xte)[:, 1]
    gb = HistGradientBoostingClassifier(       # HistGB handles NaN natively
        max_depth=3, learning_rate=0.05, l2_regularization=1.0, max_iter=400,
        early_stopping=True, validation_fraction=0.15, class_weight="balanced",
        random_state=0).fit(Xtr, ytr)
    pg = gb.predict(Xte); pgp = gb.predict_proba(Xte)[:, 1]
    la, ga = accuracy_score(yte, pl), accuracy_score(yte, pg)
    lauc, gauc = roc_auc_score(yte, pp), roc_auc_score(yte, pgp)
    print(f"\n  [{name}]")
    print(f"     Logistic (L2)          : acc {la*100:.1f}%  AUC {lauc:.3f}")
    print(f"     GradientBoosting (reg) : acc {ga*100:.1f}%  AUC {gauc:.3f}")
    best = max(la, ga)
    print(f"     -> vs baseline {base_acc*100:.1f}%: {'BEATS by %.1f pts'%((best-base_acc)*100) if best>base_acc+0.01 else 'no edge'}")
    m = {"name": name, "lr_acc": la, "lr_auc": lauc, "gb_acc": ga, "gb_auc": gauc, "best": best}
    return m, lr


def build_page(o: dict) -> None:
    """Write a self-contained ml_dashboard.html — its OWN page, separate from the host."""
    def bar(p):
        col = "#5ec98a" if p >= 55 else ("#fbbf5b" if p >= 50 else "#f87171")
        return (f'<div class="bar"><span style="width:{max(2,min(100,p)):.0f}%;background:{col}"></span></div>'
                f'<b style="color:{col}">{p:.0f}%</b>')
    rows = "".join(
        f'<tr><td class="rk">{i}</td><td class="sym">{r["symbol"]}</td>'
        f'<td class="mono">{r["event"].replace("_"," ").title()}</td>'
        f'<td class="mono">{r["date"]}</td><td class="pcell">{bar(r["prob"])}</td></tr>'
        for i, r in enumerate(o["preds"], 1))
    verdict = ("A real (modest) edge — the model beats the baseline."
               if o["edge"] else
               "No tradeable edge — the model does NOT beat the baseline. Short-term event "
               "direction is essentially unpredictable, even with fundamentals.")
    vclass = "good" if o["edge"] else "warn"
    sig = "".join(f'<tr><td class="mono">{s["f"]}</td><td class="num">{s["c"]:+.3f}</td></tr>'
                  for s in o["signals"])
    html = f"""<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>ML Experiment — Event-Trade Prediction</title><style>
:root{{--bg:#0c0f14;--panel:#141922;--ink:#e6e9ef;--mute:#98a2b3;--line:#232a35;--blue:#5b9bff;
--pos:#5ec98a;--neg:#f87171;--amb:#fbbf5b;--mono:'SFMono-Regular',Consolas,Menlo,monospace}}
*{{box-sizing:border-box}}body{{margin:0;background:var(--bg);color:var(--ink);
font-family:'Inter',-apple-system,'Segoe UI',Roboto,Arial,sans-serif;line-height:1.55}}
.wrap{{max-width:960px;margin:0 auto;padding:0 20px 80px}}
header{{padding:34px 0 18px;border-bottom:1px solid var(--line)}}
.kick{{font-family:var(--mono);font-size:.72rem;letter-spacing:.12em;text-transform:uppercase;color:var(--blue);font-weight:700}}
h1{{font-size:1.7rem;margin:.3em 0 .2em}}.sub{{color:var(--mute);margin:0;max-width:70ch}}
.sep{{display:inline-block;font-family:var(--mono);font-size:.7rem;color:var(--amb);border:1px solid var(--amb);
border-radius:20px;padding:.15em .7em;margin-top:10px}}
h2{{font-size:1.15rem;margin:30px 0 8px}}
.kpis{{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin:16px 0}}
@media(max-width:640px){{.kpis{{grid-template-columns:repeat(2,1fr)}}}}
.kpi{{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:14px;text-align:center}}
.kpi .v{{font-size:1.5rem;font-weight:800;color:var(--blue)}}.kpi .l{{font-size:.72rem;color:var(--mute);margin-top:2px}}
table{{width:100%;border-collapse:collapse;font-size:.9rem;background:var(--panel)}}
.tbl{{overflow:hidden;border:1px solid var(--line);border-radius:12px;margin:12px 0}}
th,td{{padding:9px 13px;border-bottom:1px solid var(--line);text-align:left}}
th{{font-size:.68rem;text-transform:uppercase;letter-spacing:.04em;color:var(--mute);background:#0f141c}}
tr:last-child td{{border-bottom:none}}.mono{{font-family:var(--mono);font-size:.84rem}}.num{{text-align:right;font-family:var(--mono)}}
.rk{{color:var(--mute);font-family:var(--mono);width:34px}}.sym{{font-weight:700}}
.pcell{{display:flex;align-items:center;gap:10px}}.bar{{flex:1;height:9px;background:#0f141c;border-radius:6px;overflow:hidden;max-width:260px}}
.bar span{{display:block;height:100%}}
.callout{{border-left:4px solid var(--amb);background:var(--panel);border-radius:0 10px 10px 0;padding:12px 16px;margin:14px 0}}
.callout.good{{border-left-color:var(--pos)}}.callout.warn{{border-left-color:var(--neg)}}
.mrow{{display:grid;grid-template-columns:1fr 1fr;gap:12px}}@media(max-width:640px){{.mrow{{grid-template-columns:1fr}}}}
.foot{{color:var(--mute);font-size:.8rem;margin-top:30px;font-family:var(--mono)}}
</style></head><body><div class="wrap">
<header><div class="kick">Nifty 50 · Machine-Learning Experiment</div>
<h1>Event-Trade Prediction</h1>
<p class="sub">Can a regularised model predict whether a standard event trade
(<b>{o['trade']}</b>) finishes positive? Pooled across all 50 stocks, time-validated, no look-ahead.</p>
<div class="sep">⚙ separate from the main dashboard (port 8090) — for review only</div></header>

<div class="kpis">
<div class="kpi"><div class="v">{o['n']:,}</div><div class="l">events (train {o['n_train']:,} / test {o['n_test']:,})</div></div>
<div class="kpi"><div class="v">{o['baseline']:.0f}%</div><div class="l">baseline (always WIN)</div></div>
<div class="kpi"><div class="v">{max(o['price']['best'],o['fund']['best'])*100:.1f}%</div><div class="l">best model accuracy</div></div>
<div class="kpi"><div class="v">{o['fund']['lr_auc']:.3f}</div><div class="l">best AUC (0.5 = random)</div></div>
</div>
<div class="callout {vclass}"><b>Verdict:</b> {verdict}</div>

<h2>Model comparison (out-of-time test, ≥ {o['test_from']})</h2>
<div class="tbl"><table><thead><tr><th>Feature set</th><th class="num">Logistic acc</th><th class="num">Logistic AUC</th>
<th class="num">GBM acc</th><th class="num">GBM AUC</th></tr></thead><tbody>
<tr><td>Baseline (always WIN)</td><td class="num">{o['baseline']:.1f}%</td><td class="num">—</td><td class="num">—</td><td class="num">—</td></tr>
<tr><td>Price only</td><td class="num">{o['price']['lr_acc']*100:.1f}%</td><td class="num">{o['price']['lr_auc']:.3f}</td><td class="num">{o['price']['gb_acc']*100:.1f}%</td><td class="num">{o['price']['gb_auc']:.3f}</td></tr>
<tr><td>Price + Fundamentals</td><td class="num">{o['fund']['lr_acc']*100:.1f}%</td><td class="num">{o['fund']['lr_auc']:.3f}</td><td class="num">{o['fund']['gb_acc']*100:.1f}%</td><td class="num">{o['fund']['gb_auc']:.3f}</td></tr>
</tbody></table></div>

<div class="mrow">
<div><h2>Strongest signals</h2><div class="tbl"><table><thead><tr><th>Feature</th><th class="num">Weight</th></tr></thead><tbody>{sig}</tbody></table></div></div>
<div><h2>How to read it</h2><div class="callout">Probabilities below are the model's estimate for each stock's
<b>most recent</b> event. With AUC≈0.5 they cluster near 50% — shown for transparency, not as trade signals.</div></div>
</div>

<h2>Per-stock prediction — latest event ({len(o['preds'])} stocks)</h2>
<div class="tbl"><table><thead><tr><th>#</th><th>Stock</th><th>Event</th><th>Date</th><th>Predicted WIN probability</th></tr></thead>
<tbody>{rows}</tbody></table></div>

<p class="foot">Standalone ML experiment · pooled + regularised + time-split · not wired into the live dashboard ·
for study, not investment advice.</p>
</div></body></html>"""
    Path("ml_dashboard.html").write_text(html, encoding="utf-8")


def main() -> None:
    print(f"Building pooled dataset — trade = buy {N_BEFORE}d before / sell {M_AFTER}d after, "
          f"events={EVENTS}\n")
    df = build_dataset()
    PRICE = ["mom20", "mom60", "vol20", "dist_52w_high", "above_ma50", "rel_strength", "month"] \
        + [c for c in df.columns if c.startswith("ev_")]
    FUND = ["f_roce", "f_sales_g", "f_opm", "f_profit_g", "f_de", "f_payout"]

    tr = df[df["date"] < TEST_FROM]
    te = df[df["date"] >= TEST_FROM]
    fund_cov = 100 * df[FUND].notna().any(axis=1).mean()
    print(f"\nDataset: {len(df):,} events  ({len(tr):,} train < {TEST_FROM}, {len(te):,} test ≥)")
    print(f"Overall WIN rate: {df['win'].mean()*100:.1f}%  |  rows with fundamentals: {fund_cov:.0f}%")

    base = te["win"].mean()
    base_acc = max(base, 1 - base)
    print("\n" + "=" * 62)
    print(f"  BASELINE (always predict majority class): {base_acc*100:.1f}%")
    print("=" * 62)

    mp, lr_price = evaluate("PRICE only", tr[PRICE].astype(float), tr["win"],
                            te[PRICE].astype(float), te["win"], base_acc)
    mf, _ = evaluate("PRICE + FUNDAMENTALS", tr[PRICE + FUND].astype(float), tr["win"],
                     te[PRICE + FUND].astype(float), te["win"], base_acc)

    print("\n" + "=" * 62)
    best = max(mp["best"], mf["best"])
    edge = best > base_acc + 0.01
    if edge:
        print(f"  RESULT: best model BEATS baseline by {(best-base_acc)*100:.1f} pts.")
    else:
        print(f"  RESULT: NO model beats the baseline ({best*100:.1f}% vs {base_acc*100:.1f}%).")
        print("          Even with fundamentals, short-term event direction is ~unpredictable here.")
    print("=" * 62)

    coef = pd.Series(lr_price.named_steps["logisticregression"].coef_[0],
                     index=PRICE).sort_values(key=abs, ascending=False)

    # ── per-stock prediction for the most RECENT event (model trained on all data) ──
    full = make_pipeline(SimpleImputer(strategy="median"), StandardScaler(),
                         LogisticRegression(C=0.1, penalty="l2", class_weight="balanced",
                                            max_iter=2000)).fit(df[PRICE + FUND].astype(float), df["win"])
    preds = []
    for sym, g in df.sort_values("date").groupby("symbol"):
        last = g.iloc[[-1]]
        prob = float(full.predict_proba(last[PRICE + FUND].astype(float))[0, 1])
        preds.append({"symbol": sym, "event": last["etype"].iloc[0],
                      "date": str(pd.Timestamp(last["date"].iloc[0]).date()),
                      "prob": round(prob * 100, 1)})
    preds.sort(key=lambda r: r["prob"], reverse=True)

    out = {
        "trade": f"buy {N_BEFORE}d before / sell {M_AFTER}d after",
        "events": EVENTS, "n": int(len(df)), "n_train": int(len(tr)), "n_test": int(len(te)),
        "win_rate": round(df["win"].mean() * 100, 1), "fund_cov": round(fund_cov, 0),
        "baseline": round(base_acc * 100, 1), "edge": bool(edge),
        "price": mp, "fund": mf,
        "signals": [{"f": k, "c": round(v, 3)} for k, v in coef.head(6).items()],
        "preds": preds,
        "test_from": TEST_FROM,
    }
    Path("ml_output.json").write_text(json.dumps(out, indent=2), encoding="utf-8")
    build_page(out)
    print(f"\nWrote ml_output.json + ml_dashboard.html  ({len(preds)} stock predictions)")
    print("Serve separately:  python ml_host.py   -> http://localhost:8095")


if __name__ == "__main__":
    main()
