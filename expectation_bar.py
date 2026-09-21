"""
expectation_bar.py  —  event-news attribution engine (deterministic v1)
===============================================================================
Implements the Expectation-Bar spec's RULES and JSON SCHEMA without an LLM:
causality (ingested_at < decision_ts), reactive-article exclusion, anticipatory
/ informational / reactive / noise classification, weighted bar scoring,
coverage->null, peer read-across, governance overrides, entity attribution.

It measures HOW HIGH THE BAR WAS SET before the event — never the outcome, never
a trade view. A high bar just means good news is already priced.

    python expectation_bar.py APOLLOHOSP 2026-08-12
    python expectation_bar.py --upcoming            # all near-term events

NOTE: Google News RSS gives headline + source only (no body), so classification
and entity attribution are headline-based. Swap `score_batch()` for a call to
your LLM prompt (same schema) once an API key is available — that's the only
change needed to upgrade this to the full engine.
===============================================================================
"""
from __future__ import annotations

import datetime as dt
import json
import re
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

from news_ingest import articles_for, names

ROOT = Path(__file__).resolve().parent

# ── sector map for peer read-across (compact; extend as needed) ───────────────
SECTORS = {
    "Banks": ["HDFCBANK", "ICICIBANK", "AXISBANK", "KOTAKBANK", "SBIN"],
    "NBFC/Fin": ["BAJFINANCE", "BAJAJFINSV", "SHRIRAMFIN", "JIOFIN"],
    "Insurance": ["HDFCLIFE", "SBILIFE"],
    "IT": ["TCS", "INFY", "HCLTECH", "TECHM", "WIPRO"],
    "Auto": ["MARUTI", "M&M", "BAJAJ-AUTO", "EICHERMOT", "TMPV"],
    "Metals": ["TATASTEEL", "JSWSTEEL", "HINDALCO", "COALINDIA"],
    "Energy": ["RELIANCE", "ONGC", "NTPC", "POWERGRID"],
    "Pharma/Health": ["SUNPHARMA", "CIPLA", "DRREDDY", "APOLLOHOSP", "MAXHEALTH"],
    "FMCG": ["HINDUNILVR", "ITC", "NESTLEIND", "TATACONSUM"],
    "Consumer": ["TITAN", "TRENT", "ASIANPAINT", "ETERNAL"],
    "Cement": ["ULTRACEMCO", "GRASIM"],
    "Infra/CapGoods": ["LT", "BEL"],
    "Ports/Diversified": ["ADANIENT", "ADANIPORTS"],
    "Telecom": ["BHARTIARTL"],
    "Aviation": ["INDIGO"],
}
SYM_SECTOR = {s: sec for sec, lst in SECTORS.items() for s in lst}

# ── lexicons (headline-based) ─────────────────────────────────────────────────
REACTIVE = re.compile(r"\b(surge|surges|surged|rally|rallies|rallied|jump|jumps|"
                      r"jumped|soar|soars|plunge|plunges|slump|slumps|tumble|"
                      r"tumbles|gain|gains|slip|slips|fall|falls|drop|drops|"
                      r"rise|rises|rose|hits?\s+\d|52-week|all-time high|"
                      r"top (nifty|sensex)|ahead of results|before results)\b", re.I)
ANTICIPATORY = re.compile(r"\b(expect|expected|estimate|estimates|preview|"
                          r"forecast|guidance|street sees|brokerage|analysts?|"
                          r"target price|to watch|ahead of|likely to|outlook|"
                          r"q1 preview|earnings preview)\b", re.I)
INFORMATIONAL = re.compile(r"\b(order|contract|capex|expansion|launch|acquire|"
                           r"acquisition|stake|price hike|approval|approved|"
                           r"partnership|deal|invest|plant|capacity|dividend|"
                           r"tariff|hike|appoint)\b", re.I)

POS = re.compile(r"\b(upgrade|upgrades|raised? target|hike target|beat|beats|"
                 r"strong|robust|outperform|buy rating|bullish|record|"
                 r"multibagger|top pick|accumulate|positive)\b", re.I)
NEG = re.compile(r"\b(downgrade|downgrades|cut target|lower target|miss|misses|"
                 r"weak|soft|underperform|sell rating|bearish|warning|warns|"
                 r"concern|pressure|decline|slowdown|negative)\b", re.I)

GOVERNANCE = {
    "auditor_resignation": re.compile(r"auditor.*(resign|quit)", re.I),
    "kmp_exit": re.compile(r"(cfo|ceo|md|chairman|director).*(resign|quit|step down|exit)", re.I),
    "promoter_pledge": re.compile(r"promoter.*(pledge|pledged)", re.I),
    "credit_downgrade": re.compile(r"(credit|rating).*(downgrade|cut)|icra|crisil.*downgrade", re.I),
    "default": re.compile(r"\b(default|insolvency|nclt|bankrupt)\b", re.I),
    "sebi_action": re.compile(r"sebi.*(action|order|probe|penalt|ban)", re.I),
    "delayed_filing": re.compile(r"(delay|postpone).*(result|filing|board meeting)", re.I),
}

CREDIBLE = {"economic times", "moneycontrol", "business standard", "livemint",
            "mint", "the hindu businessline", "businessline", "cnbc", "reuters",
            "bloomberg", "ndtv profit", "financial express", "business today",
            "zee business", "et now"}


def _src_weight(source: str) -> float:
    s = (source or "").lower()
    return 1.0 if any(k in s for k in CREDIBLE) else 0.5


def _parse(ts: str):
    if not ts:
        return None
    try:
        return dt.datetime.fromisoformat(ts.replace("Z", "+00:00"))
    except Exception:                                       # noqa: BLE001
        return None


def _clip12(text: str) -> str:
    w = text.split()
    return " ".join(w[:12]) + ("…" if len(w) > 12 else "")


def classify(title: str) -> str:
    if REACTIVE.search(title):
        return "REACTIVE"
    if ANTICIPATORY.search(title):
        return "ANTICIPATORY"
    if INFORMATIONAL.search(title):
        return "INFORMATIONAL"
    return "NOISE"


def _attribution(sym: str, name: str, title: str) -> str:
    t = title.lower()
    if sym.lower() in t or (name and name.split()[0].lower() in t):
        return "high"
    return "low"


def score_batch(symbol: str, decision_ts: dt.datetime) -> dict:
    """Deterministic implementation of the Expectation-Bar spec -> JSON dict."""
    if decision_ts.tzinfo is None:                         # treat as market close (15:30 IST)
        IST = dt.timezone(dt.timedelta(hours=5, minutes=30))
        decision_ts = decision_ts.replace(hour=15, minute=30, tzinfo=IST)
    nm = names()
    name = nm.get(symbol, symbol)
    arts = articles_for(symbol)

    counts = {"pre_event": 0, "post_event": 0, "reactive": 0,
              "informational": 0, "anticipatory": 0, "noise": 0}
    detail, gov = [], None
    bar_num, bar_den, credible_specific = 0.0, 0.0, 0

    for a in arts:
        ing = _parse(a["ingested_at"])
        pre = bool(ing and ing < decision_ts)              # missing -> post
        counts["pre_event" if pre else "post_event"] += 1
        cls = classify(a["title"])
        counts[cls.lower()] = counts.get(cls.lower(), 0) + 1
        reactive = cls == "REACTIVE"
        if reactive:
            counts["reactive"] += 1
        attr = _attribution(symbol, name, a["title"])
        contrib = None

        if pre and not reactive:                           # governance scan (pre-event only)
            for gtype, rx in GOVERNANCE.items():
                if rx.search(a["title"]):
                    gov = {"type": gtype, "headline": _clip12(a["title"]),
                           "source": a["source"]}

        if pre and not reactive and attr == "high":        # bar scoring pool
            recency = 1.0
            if ing:
                days = max(0.0, (decision_ts - ing).total_seconds() / 86400)
                recency = max(0.3, 1.0 - days / 30.0)       # fade over ~30d
            w = _src_weight(a["source"]) * recency
            s = (1 if POS.search(a["title"]) else 0) - (1 if NEG.search(a["title"]) else 0)
            if s != 0:
                bar_num += s * w
                bar_den += w
            credible_specific += 1
            contrib = round(s * w, 3)

        detail.append({
            "headline": _clip12(a["title"]), "source": a["source"] or None,
            "link": a.get("link") or None,
            "ingested_at": a["ingested_at"], "pre_event": pre,
            "classification": cls, "is_reactive": reactive,
            "attribution_confidence": attr, "bar_contrib": contrib,
        })

    no_coverage = credible_specific < 2
    bar = None if (no_coverage or bar_den == 0) else round(max(-1.0, min(1.0, bar_num / bar_den)), 3)

    # peer read-across: sector peers with pre-event results/guidance headlines
    peers = []
    sector = SYM_SECTOR.get(symbol)
    for p in SECTORS.get(sector, []):
        if p == symbol:
            continue
        for a in articles_for(p):
            ing = _parse(a["ingested_at"])
            if not (ing and ing < decision_ts):
                continue
            if re.search(r"\b(result|results|q1|q2|profit|guidance|earnings)\b", a["title"], re.I):
                direction = ("up" if POS.search(a["title"]) else
                             "down" if NEG.search(a["title"]) else "neutral")
                peers.append({"peer": p, "direction": direction,
                              "linked_to_subject": name.split()[0].lower() in a["title"].lower(),
                              "headline": _clip12(a["title"])})
                break

    return {
        "symbol": symbol, "decision_ts": decision_ts.isoformat(),
        "generated_at": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
        "engine": "deterministic-v1 (headline-based; swap in LLM for full spec)",
        "expectation_bar_raw": bar,
        "no_coverage": no_coverage,
        "counts": counts,
        "credible_specific_pre_event": credible_specific,
        "peer_signals": peers or None,
        "governance_flag": gov,
        "confidence": {
            "coverage": "none" if no_coverage else ("low" if credible_specific < 4 else "medium"),
            "attribution": "headline-only (no article body available via RSS)",
        },
        "articles": detail,
    }


def peer_readacross(symbol: str, decision_ts: dt.datetime) -> dict:
    """Full sector peer read-across for one event: every sector peer and, if it
    has a pre-event results/guidance headline, its direction + linkage. A peer
    beat/miss resets the bar even when the subject itself has no coverage."""
    if decision_ts.tzinfo is None:
        IST = dt.timezone(dt.timedelta(hours=5, minutes=30))
        decision_ts = decision_ts.replace(hour=15, minute=30, tzinfo=IST)
    nm = names()
    name = nm.get(symbol, symbol)
    first = (name.split()[0].lower() if name else symbol.lower())
    sector = SYM_SECTOR.get(symbol)
    rx = re.compile(r"\b(result|results|q1|q2|q3|q4|profit|guidance|earnings)\b", re.I)
    peers = []
    for p in SECTORS.get(sector, []):
        if p == symbol:
            continue
        sig = None
        for a in articles_for(p):
            ing = _parse(a["ingested_at"])
            if not (ing and ing < decision_ts) or not rx.search(a["title"]):
                continue
            d = ("up" if POS.search(a["title"]) else
                 "down" if NEG.search(a["title"]) else "neutral")
            sig = {"direction": d, "linked": first in a["title"].lower(),
                   "headline": _clip12(a["title"]), "source": a["source"] or None,
                   "link": a.get("link") or None}
            break
        peers.append({"peer": p, "name": nm.get(p, p), "signal": sig})
    up = sum(1 for x in peers if x["signal"] and x["signal"]["direction"] == "up")
    dn = sum(1 for x in peers if x["signal"] and x["signal"]["direction"] == "down")
    reported = sum(1 for x in peers if x["signal"])
    return {"symbol": symbol, "name": name, "sector": sector,
            "decision_ts": decision_ts.isoformat(),
            "reported": reported, "up": up, "down": dn,
            "net": ("up" if up > dn else "down" if dn > up else "mixed/none"),
            "peers": peers}


# ── convenience: find near-term events from the board-meeting feeds ────────────
def upcoming_events(days: int = 45) -> list[tuple]:
    import glob, os, pandas as pd, warnings
    warnings.filterwarnings("ignore")
    today = dt.datetime.now()
    out = []
    for p in glob.glob(str(ROOT / "*" / "board_meetings.csv")):
        sym = Path(p).parent.name
        try:
            df = pd.read_csv(p, dtype=str).fillna("")
        except Exception:                                  # noqa: BLE001
            continue
        if "bm_date" not in df.columns:
            continue
        df["d"] = pd.to_datetime(df["bm_date"], errors="coerce", format="mixed")
        for _, r in df[(df["d"] > today) &
                       (df["d"] <= today + pd.Timedelta(days=days))].iterrows():
            out.append((sym, r["d"].to_pydatetime()))
    return sorted(out, key=lambda x: x[1])


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    if "--upcoming" in sys.argv or not args:
        events = upcoming_events()
        print(f"{len(events)} upcoming event(s)\n")
        for sym, ts in events:
            print(json.dumps(score_batch(sym, ts), ensure_ascii=False, indent=2))
            print()
        return
    sym = args[0]
    ts = dt.datetime.fromisoformat(args[1]) if len(args) > 1 else dt.datetime.now()
    print(json.dumps(score_batch(sym, ts), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
