"""
scoring.py
===============================================================================
Four 0-100 scores per stock, each returned WITH THE DRIVERS THAT PRODUCED IT.

    Quality     fundamentals — returns, margins, growth, leverage, cash quality
    Risk        governance red flags (100 = clean, 0 = alarming)
    Sentiment   tone of the company's own filings, from the FinBERT router
    Technical   trend / momentum / volatility regime

WHY PERCENTILE RANKS
--------------------
Every component is ranked WITHIN the universe rather than scored against fixed
thresholds. "ROCE above 20%" means something different for a bank and a cement
maker, and any absolute cut has to be re-tuned every time the universe changes.
A percentile says the only thing that is actually defensible from this data:
where the company sits among its peers, today.

The cost is that scores are RELATIVE. In a universe where everything is
mediocre, something still scores 100. The payload says so — `universe_n` is
returned alongside, and the UI must show it.

WHAT THE RISK SCORE IS BUILT FROM — AND WHAT IT IS NOT
------------------------------------------------------
The design called for six inputs. Three exist in this dataset; three do not:

  BUILT
    disclosure burial   share of filings released after hours / Friday evening,
                        from processed/disclosure_timing_announcements.csv
                        (88,548 filings, 50 symbols, permutation-tested by
                        disclosure_timing.py). This is the component nobody
                        else has.
    negative filings    share of material filings the router tags Negative
    governance events   auditor resignations, litigation, penalties, show-cause
                        notices, rating downgrades, as a share of all filings

  NOT BUILT — the data is not there
    promoter encumbrance  new data/20_Promoter_Encumbrance_Details.xlsx contains
                          one row: "No records found on NSE for Promoter
                          Encumbrance."
    insider trading (PIT) new data/16_Insider_Trading_PIT.xlsx likewise: "No
                          records found on NSE for Insider Trading."
    investor complaints   new data/18_Investor_Complaints.xlsx has 20 companies,
                          none of them in this universe.

  So Risk is a THREE-input score. Do not present it as a full governance audit.

    python scoring.py                # score the Nifty 50, print a table
    python scoring.py RELIANCE TCS   # a couple of stocks with their drivers
===============================================================================
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

import numpy as np
import pandas as pd

import score_feeds
import technicals as T

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:                                 # noqa: BLE001
    pass

ROOT = Path(__file__).resolve().parent
PROC = ROOT / "processed"
BURIAL = PROC / "disclosure_timing_announcements.csv"

# Filing categories that are governance events rather than business news.
GOV_EVENT = re.compile(
    r"resignation|resigned|cessation"
    r"|change in auditor|auditor"
    r"|litigation|court|tribunal|nclt|arbitrat"
    r"|penalt|fine|show cause|adjudicat"
    r"|search|seizure|raid|investigation"
    r"|credit rating[- ]?(?:revision|downgrade)"
    r"|default|insolvenc|delist|suspension",
    re.I,
)


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def pct_rank(values: dict[str, float], higher_is_better: bool = True) -> dict[str, float]:
    """Percentile rank in 0..100. Missing values stay missing rather than being
    silently imputed to the median — a stock with no cash-flow data should show
    a gap, not a fabricated average."""
    ok = {k: v for k, v in values.items()
          if v is not None and isinstance(v, (int, float)) and np.isfinite(v)}
    if not ok:
        return {k: None for k in values}
    s = pd.Series(ok)
    r = s.rank(pct=True, ascending=higher_is_better) * 100
    out = {k: None for k in values}
    for k, v in r.items():
        out[k] = round(float(v), 1)
    return out


def _blend(parts: list[tuple[float | None, float]]) -> float | None:
    """Weighted mean over the components that exist, re-normalised."""
    live = [(v, w) for v, w in parts if v is not None]
    if not live:
        return None
    tot = sum(w for _v, w in live)
    return round(sum(v * w for v, w in live) / tot, 1) if tot else None


def _f(v):
    try:
        v = float(v)
        return v if np.isfinite(v) else None
    except (TypeError, ValueError):
        return None


# --------------------------------------------------------------------------- #
# Raw component extraction
# --------------------------------------------------------------------------- #
def fundamental_raws(universe: list[str]) -> dict[str, dict]:
    """Latest annual fundamentals per symbol, as raw (un-ranked) numbers."""
    import fund_loader

    out: dict[str, dict] = {}
    for sym in universe:
        try:
            d = fund_loader.load_stock(sym)
        except Exception:                         # noqa: BLE001
            continue
        pnl, bs, rat, cf = (d.get("pnl"), d.get("balance_sheet"),
                            d.get("ratios"), d.get("cash_flow"))
        if pnl is None or pnl.empty:
            continue
        L = pnl.iloc[-1]

        de = None
        if bs is not None and not bs.empty:
            B = bs.iloc[-1]
            base = (_f(B.get("Equity Capital")) or 0) + (_f(B.get("Reserves")) or 0)
            bor = _f(B.get("Borrowings"))
            if bor is not None and base:
                de = bor / base

        # 3-year CAGR beats a single year's growth %, which is hostage to one
        # weak base year.
        def cagr(col):
            s = pd.to_numeric(pnl[col], errors="coerce").dropna() if col in pnl else None
            if s is None or len(s) < 4:
                return None
            a, b = float(s.iloc[-4]), float(s.iloc[-1])
            if a <= 0 or b <= 0:
                return None
            return ((b / a) ** (1 / 3) - 1) * 100

        cfo_op = None
        if cf is not None and not cf.empty and "CFO/OP" in cf.columns:
            s = pd.to_numeric(cf["CFO/OP"], errors="coerce").dropna()
            if len(s):
                cfo_op = float(s.tail(3).median())

        out[sym] = {
            "roce": _f(rat.iloc[-1].get("ROCE %")) if (rat is not None and not rat.empty) else None,
            "opm": _f(L.get("OPM %")),
            "sales_cagr3": cagr("Sales"),
            "profit_cagr3": cagr("Net Profit"),
            "de": de,
            "cfo_op": cfo_op,
        }
    return out


def burial_rates(universe: list[str]) -> dict[str, dict]:
    """Share of each stock's filings released after hours / Friday evening.

    Straight out of disclosure_timing.py's own output, so the definition of
    "buried" is the one that module permutation-tested, not a second opinion.
    """
    out = {s: {"burial": None, "n": 0} for s in universe}
    if not BURIAL.exists():
        return out
    try:
        df = pd.read_csv(BURIAL, usecols=["symbol", "is_buried"])
    except Exception:                             # noqa: BLE001
        return out
    g = df.groupby("symbol")["is_buried"].agg(["mean", "count"])
    for sym in universe:
        if sym in g.index:
            out[sym] = {"burial": float(g.loc[sym, "mean"]) * 100,
                        "n": int(g.loc[sym, "count"])}
    return out


def filing_risk(universe: list[str]) -> dict[str, dict]:
    """Negative-filing share and governance-event share, from the sentiment cache
    plus the announcement categories."""
    out = {}
    for sym in universe:
        df = score_feeds.sentiment_frame(sym)
        rec = {"neg_share": None, "gov_share": None, "n_material": 0}
        if not df.empty:
            mat = df[df["engine"] != "rule"]
            if len(mat) >= 20:
                rec["n_material"] = int(len(mat))
                rec["neg_share"] = float((mat["label"] == "Negative").mean()) * 100
            cats = df["category"].fillna("").astype(str)
            if len(df) >= 20:
                rec["gov_share"] = float(cats.str.contains(GOV_EVENT).mean()) * 100
        out[sym] = rec
    return out


# --------------------------------------------------------------------------- #
# The four scores
# --------------------------------------------------------------------------- #
QUALITY_WEIGHTS = [("roce", 0.25, True), ("opm", 0.15, True),
                   ("sales_cagr3", 0.15, True), ("profit_cagr3", 0.20, True),
                   ("de", 0.10, False), ("cfo_op", 0.15, True)]

RISK_WEIGHTS = [("burial", 0.40), ("neg_share", 0.35), ("gov_share", 0.25)]

DRIVER_LABEL = {
    "roce": "ROCE %", "opm": "Operating margin %", "sales_cagr3": "Sales CAGR 3y %",
    "profit_cagr3": "Profit CAGR 3y %", "de": "Debt / equity",
    "cfo_op": "Cash conversion (CFO/OP)", "burial": "Filings released after hours %",
    "neg_share": "Negative filings %", "gov_share": "Governance-event filings %",
}


def build_scores(universe: list[str]) -> dict:
    """Score the whole universe at once — percentile ranks need the peer set."""
    universe = sorted(set(universe))
    fund = fundamental_raws(universe)
    bur = burial_rates(universe)
    fil = filing_risk(universe)

    # ---- Quality ---------------------------------------------------------- #
    q_ranks = {}
    for key, _w, hib in QUALITY_WEIGHTS:
        q_ranks[key] = pct_rank({s: fund.get(s, {}).get(key) for s in universe},
                                higher_is_better=hib)

    # ---- Risk: rank each flag so that HIGH RANK = SAFE -------------------- #
    r_ranks = {
        "burial": pct_rank({s: bur[s]["burial"] for s in universe}, False),
        "neg_share": pct_rank({s: fil[s]["neg_share"] for s in universe}, False),
        "gov_share": pct_rank({s: fil[s]["gov_share"] for s in universe}, False),
    }

    # ---- Sentiment -------------------------------------------------------- #
    # Percentile-ranked like everything else, NOT mapped linearly from the EWMA.
    # Most filings are Neutral, so the raw 90-day EWMA sits within a whisker of
    # zero for every stock; (ewma+1)*50 put the entire Nifty 50 between 45 and
    # 59, which is a scale that cannot distinguish anything. The raw tone is
    # kept in `sentiment_detail` so the absolute number is still inspectable.
    sent_all = {s: score_feeds.sentiment_series(s) for s in universe}
    sent_rank = pct_rank({s: sent_all[s].get("ewma", {}).get("ewma_90")
                          for s in universe}, True)

    out = {}
    for sym in universe:
        quality = _blend([(q_ranks[k][sym], w) for k, w, _h in QUALITY_WEIGHTS])
        risk = _blend([(r_ranks[k][sym], w) for k, w in RISK_WEIGHTS])

        sent = sent_all[sym]
        sent_score = sent_rank[sym]

        tech = T.latest(sym, fetch=False)
        tech_score, tech_drivers = _technical(tech)

        drivers = []
        for k, _w, _h in QUALITY_WEIGHTS:
            raw = fund.get(sym, {}).get(k)
            if raw is not None:
                drivers.append({"group": "quality", "key": k,
                                "label": DRIVER_LABEL[k], "raw": round(raw, 2),
                                "pct": q_ranks[k][sym]})
        for k, _w in RISK_WEIGHTS:
            raw = (bur[sym]["burial"] if k == "burial" else fil[sym].get(k))
            if raw is not None:
                drivers.append({"group": "risk", "key": k,
                                "label": DRIVER_LABEL[k], "raw": round(raw, 2),
                                "pct": r_ranks[k][sym]})

        out[sym] = {
            "symbol": sym,
            "quality": quality, "risk": risk,
            "sentiment": sent_score, "technical": tech_score,
            "drivers": drivers + tech_drivers,
            "sentiment_detail": sent,
            "n_filings_buried_from": bur[sym]["n"],
        }
    return {"universe_n": len(universe), "scores": out}


def _technical(t: dict | None) -> tuple[float | None, list]:
    """Trend + momentum + volatility regime, each mapped to 0..100 on its own
    natural scale (RSI is already 0-100; ADX saturates around 50)."""
    if not t:
        return None, []
    parts, drivers = [], []

    rsi = t.get("rsi_14")
    if rsi is not None:
        parts.append((float(rsi), 0.30))
        drivers.append({"group": "technical", "key": "rsi_14",
                        "label": "RSI (14)", "raw": round(rsi, 1), "pct": None})

    # Trend: above the 50-day and a 50/200 golden cross, scaled by ADX so a
    # trend nobody is following counts for less than a strong one.
    above = t.get("above_ma50")
    gold = t.get("golden_cross")
    adxv = t.get("adx_14")
    if above is not None and gold is not None:
        strength = min(float(adxv or 0) / 40.0, 1.0)
        base = 50 + 25 * (float(above) * 2 - 1) + 25 * (float(gold) * 2 - 1)
        parts.append((50 + (base - 50) * (0.4 + 0.6 * strength), 0.45))
        drivers.append({"group": "technical", "key": "trend",
                        "label": "Trend (>50dma, 50/200 cross, ADX)",
                        "raw": f"{'above' if above else 'below'} 50dma, "
                               f"{'golden' if gold else 'death'} cross, "
                               f"ADX {adxv:.0f}" if adxv is not None else "n/a",
                        "pct": None})

    mh = t.get("macd_hist")
    cl = t.get("close")
    if mh is not None and cl:
        # Normalise the histogram by price so it is comparable across stocks.
        norm = np.clip(float(mh) / float(cl) * 2000, -50, 50)
        parts.append((50 + norm, 0.25))
        drivers.append({"group": "technical", "key": "macd_hist",
                        "label": "MACD histogram", "raw": round(float(mh), 2),
                        "pct": None})

    return _blend(parts), drivers


# --------------------------------------------------------------------------- #
def main():
    syms = [s.upper() for s in sys.argv[1:]]
    if not syms:
        from download_feeds import NIFTY50_FALLBACK
        syms = sorted(set(NIFTY50_FALLBACK))
    res = build_scores(syms)
    sc = res["scores"]

    print("=" * 78)
    print(f"  SCORES — percentile ranks within a {res['universe_n']}-stock universe")
    print("=" * 78)
    print(f"  {'symbol':<14}{'quality':>9}{'risk':>8}{'sentim':>9}{'tech':>8}")
    print("  " + "-" * 74)
    for sym in sorted(sc, key=lambda s: -(sc[s]["quality"] or -1)):
        d = sc[sym]
        f = lambda v: f"{v:>8.1f}" if v is not None else "       -"    # noqa: E731
        print(f"  {sym:<14}{f(d['quality'])} {f(d['risk'])}{f(d['sentiment'])}"
              f"{f(d['technical'])}")

    if len(syms) <= 3:
        for sym in syms:
            print(f"\n  {sym} drivers")
            for dr in sc[sym]["drivers"]:
                p = f"pct {dr['pct']:>5.1f}" if dr["pct"] is not None else "         "
                print(f"    {dr['group']:<10}{dr['label']:<38}{str(dr['raw']):>12}  {p}")


if __name__ == "__main__":
    main()
