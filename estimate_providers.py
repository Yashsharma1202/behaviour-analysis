"""
estimate_providers.py
===============================================================================
Where a forward estimate can come from. Four providers behind one interface;
each may fail without taking the others down.

    own_history      next-quarter revenue/PAT extrapolated from quarterly/
    calendar         WHEN the next result is due, from board_meetings.csv
    web_search       live consensus/broker previews (opt-in, cached 24h)
    screener_scrape  a free web source (opt-in, DISABLED by default)

Every provider returns an Estimate or None. Never raise into the caller — an
estimate is a nice-to-have and must not break the page.

    Estimate(value, source, as_of, confidence, detail)

POINT-IN-TIME WARNING
---------------------
Only `own_history` and `calendar` may ever feed the ML model. The external
providers return TODAY's consensus, and there is no archive of what consensus
said in 2023 — using it on a 2023 event is look-ahead bias of the purest kind.
estimates.py enforces this via for_model=True; do not route around it.

WHY web_search AND screener_scrape SHIP INERT
---------------------------------------------
They are written, wired and tested, but return None unless explicitly enabled:
web_search needs a fetch callback the caller supplies (this process has no
outbound search binding), and screener_scrape is off because scraping a third
party's site is the caller's decision to make, not this module's. Turning
either on is one flag; neither is a stub.
===============================================================================
"""

from __future__ import annotations

import json
import re
import time
from dataclasses import dataclass, asdict
from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parent
CACHE = ROOT / "processed" / "estimates" / "cache"
WEB_TTL = 24 * 3600


@dataclass
class Estimate:
    value: float | None
    source: str
    as_of: str
    confidence: float          # 0..1 — how much the merge policy should trust it
    detail: dict

    def to_dict(self) -> dict:
        return asdict(self)


# --------------------------------------------------------------------------- #
# 1. own_history — always available, zero cost, the baseline
# --------------------------------------------------------------------------- #
def own_history(sym: str, metric: str = "Net Profit",
                asof: pd.Timestamp | None = None) -> Estimate | None:
    """Next quarter from the company's own history: seasonal naive plus trend.

    A quarter is compared to the SAME quarter a year earlier, not to the one
    before it — Indian quarterlies are strongly seasonal and a QoQ model reads
    festive-season swings as growth. So:

        estimate(Q+1) = Q(+1 - 4 quarters) * (1 + median YoY growth of last 4Q)

    The band is +/-1 sd of those YoY growth rates, which is an honest statement
    of how stable this company's growth has actually been — a volatile company
    gets a wide band and should be trusted less.

    This is the estimate-free proxy earnings_surprise.py already argues for,
    turned from a backward-looking surprise measure into a forward point.
    """
    import fund_loader
    try:
        q = fund_loader.load_stock(sym).get("quarterly")
    except Exception:                             # noqa: BLE001
        return None
    if q is None or q.empty or metric not in q.columns:
        return None

    s = pd.to_numeric(q[metric], errors="coerce").dropna()
    if asof is not None:                          # point-in-time: drop the future
        s = s[s.index <= asof]
    if len(s) < 8:
        return None                               # need 2 years to see a season

    yoy = (s / s.shift(4) - 1).dropna()
    if len(yoy) < 4:
        return None
    recent = yoy.tail(4)
    g = float(recent.median())
    sd = float(recent.std(ddof=0)) if len(recent) > 1 else abs(g)

    base = float(s.iloc[-4])                      # same quarter, a year ago
    if base <= 0:
        return None
    val = base * (1 + g)

    last_p = s.index[-1]
    nxt = (pd.Timestamp(last_p) + pd.DateOffset(months=3)).strftime("%b %Y")
    # Wide historical dispersion => low confidence, and it decays fast.
    conf = float(np.clip(1.0 / (1.0 + 6 * abs(sd)), 0.15, 0.75))
    return Estimate(
        value=round(val, 1), source="own_history",
        as_of=str(pd.Timestamp(last_p).date()), confidence=round(conf, 2),
        detail={"metric": metric, "for_quarter": nxt,
                "base_same_qtr_last_year": round(base, 1),
                "yoy_growth_median_4q": round(g * 100, 2),
                "yoy_growth_sd_4q": round(sd * 100, 2),
                "low": round(base * (1 + g - sd), 1),
                "high": round(base * (1 + g + sd), 1),
                "n_quarters": int(len(s))},
    )


# --------------------------------------------------------------------------- #
# 2. calendar — when, not what
# --------------------------------------------------------------------------- #
def calendar(sym: str, today: pd.Timestamp | None = None) -> Estimate | None:
    """Next result date from the board-meeting feed.

    Delegates to next_result.next_result_date rather than re-deriving it, so
    the Insight tab and processed/next_result.csv can never disagree.
    """
    try:
        import next_result as NR
        d = NR.next_result_date(sym, today or pd.Timestamp.now().normalize())
    except Exception:                             # noqa: BLE001
        return None
    if d is None or pd.isna(d):
        return None
    days = int((pd.Timestamp(d).normalize()
                - (today or pd.Timestamp.now().normalize())).days)
    return Estimate(value=None, source="calendar", as_of=str(pd.Timestamp(d).date()),
                    confidence=1.0,
                    detail={"next_result_date": str(pd.Timestamp(d).date()),
                            "days_away": days})


# --------------------------------------------------------------------------- #
# 3. web_search — live consensus (opt-in)
# --------------------------------------------------------------------------- #
def _cache_get(key: str):
    p = CACHE / f"{key}.json"
    if not p.exists():
        return None
    try:
        blob = json.loads(p.read_text(encoding="utf-8"))
    except Exception:                             # noqa: BLE001
        return None
    if time.time() - blob.get("_t", 0) > WEB_TTL:
        return None
    return blob.get("data")


def _cache_put(key: str, data) -> None:
    try:
        CACHE.mkdir(parents=True, exist_ok=True)
        (CACHE / f"{key}.json").write_text(
            json.dumps({"_t": time.time(), "data": data}), encoding="utf-8")
    except Exception:                             # noqa: BLE001
        pass


def web_search(sym: str, fetcher=None) -> Estimate | None:
    """Consensus estimates via a caller-supplied search callback.

    `fetcher(query: str) -> str | None` is whatever the caller wants it to be —
    a search API, an LLM tool call, a curl wrapper. This module deliberately
    does NOT hard-code one: the dashboard runs offline, and a provider that
    blocks the page on a third-party timeout is worse than no provider.

    Returns None when no fetcher is supplied, which is the default.
    """
    if fetcher is None:
        return None
    key = f"web_{sym}"
    hit = _cache_get(key)
    if hit is None:
        try:
            hit = fetcher(f"{sym} NSE next quarter consensus estimate "
                          f"revenue net profit analyst preview")
        except Exception:                         # noqa: BLE001
            return None
        if hit is None:
            return None
        _cache_put(key, hit)
    return Estimate(value=None, source="web_search",
                    as_of=str(pd.Timestamp.now().date()), confidence=0.5,
                    detail={"raw": str(hit)[:2000]})


# --------------------------------------------------------------------------- #
# 4. screener_scrape — OFF by default, and this is the only place it lives
# --------------------------------------------------------------------------- #
SCREENER_ENABLED = False


_RATIO_RE = re.compile(
    r"([A-Za-z][A-Za-z /%.\-]{2,30}?)\s*[\r\n]*"
    r"[^0-9\-₹]{0,40}?₹?\s*(-?[\d,]+(?:\.\d+)?)\s*(%|Cr\.?)?",
)


def screener_scrape(sym: str) -> Estimate | None:
    """Scrape Screener.in for the headline ratios it publishes.

    WHAT THIS CAN AND CANNOT DELIVER
    --------------------------------
    Screener publishes HISTORICAL fundamentals and current market ratios. It
    does NOT publish analyst consensus, so this provider cannot supply a
    forward estimate no matter how good the fetcher is — `value` stays None on
    purpose. What it does add over the local CSVs is live market data: price,
    market cap, current P/E, book value, dividend yield.

    For actual forward consensus you need a source that carries it (Trendlyne
    forecast pages, Moneycontrol estimates). That would be a sibling provider
    here, written the same way.

    Inert unless SCREENER_ENABLED (estimates.py --providers screener), because
    scraping a third party is the caller's decision, not this module's.
    Confining it here means the day their markup changes, exactly one thing
    breaks — and web_fetcher's adaptive selectors are meant to soften even that.

    NOT EXERCISED against the live site from this environment. Run
    `python estimates.py RELIANCE --providers screener` to validate.
    """
    if not SCREENER_ENABLED:
        return None
    import web_fetcher

    key = f"screener_{sym}"
    hit = _cache_get(key)
    ratios: dict = {}
    if isinstance(hit, dict):
        ratios = hit
    else:
        url = f"https://www.screener.in/company/{sym}/consolidated/"
        # Screener serves plain HTML; stealth is reserved for hosts that block.
        page = web_fetcher.fetch(url, stealth=False)
        if page is None or page.status >= 400 or not page.html:
            return None
        # Preferred path: Scrapling's adaptive selectors survive a re-skin.
        for li in page.css("#top-ratios li", adaptive=True):
            try:
                txt = " ".join(str(getattr(li, "text", li)).split())
            except Exception:                     # noqa: BLE001
                continue
            m = _RATIO_RE.match(txt)
            if m:
                name = m.group(1).strip().rstrip(":").title()
                try:
                    ratios[name] = float(m.group(2).replace(",", ""))
                except ValueError:
                    pass
        if not ratios:
            # Fallback: no Scrapling, or the selector found nothing. Pull the
            # same block out of raw HTML so the provider still returns data.
            blob = re.sub(r"<[^>]+>", " ", page.html)
            for name in ("Market Cap", "Current Price", "Stock P/E", "Book Value",
                         "Dividend Yield", "ROCE", "ROE", "Face Value"):
                m = re.search(re.escape(name) + r"\D{0,30}?(-?[\d,]+(?:\.\d+)?)", blob)
                if m:
                    try:
                        ratios[name] = float(m.group(1).replace(",", ""))
                    except ValueError:
                        pass
        if not ratios:
            return None
        _cache_put(key, ratios)

    return Estimate(
        value=None,                               # Screener has no consensus to give
        source="screener_scrape",
        as_of=str(pd.Timestamp.now().date()),
        confidence=0.0,                           # never outranks own_history
        detail={"live_ratios": ratios,
                "note": "market data only — Screener publishes no forward consensus"},
    )


# Providers usable as ML features — see the point-in-time warning above.
MODEL_SAFE = {"own_history", "calendar"}
