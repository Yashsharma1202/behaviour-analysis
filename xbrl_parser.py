"""
xbrl_parser.py
===============================================================================
Turns the `xbrl` LINKS in financial_results into actual NUMBERS.

THE PROBLEM
-----------
processed/financial_results_clean.csv tells you a result was filed and gives a
URL to an XBRL file -- but NOT the revenue or profit. Those live inside the XML.
This script downloads each XBRL file and extracts the headline figures:

        Revenue | Other Income | Total Income | PBT | Net Profit | Basic EPS

HOW XBRL WORKS (the one thing you must understand)
--------------------------------------------------
An XBRL file is XML where every number is a "fact" tagged with a CONCEPT and a
CONTEXT.  The context says WHICH period the number is for:

    <xbrli:context id="OneD"><xbrli:period>
        <xbrli:startDate>2024-10-01</xbrli:startDate>
        <xbrli:endDate>2024-12-31</xbrli:endDate>
    </xbrli:period></xbrli:context>

    <in-bse-fin:RevenueFromOperations contextRef="OneD">1282600000000</...>

The SAME concept appears several times (this quarter, prior quarter, year-to-
date, prior year...). To get "this quarter's revenue" we pick the fact whose
context period is ~one quarter long (~85-95 days) and ENDS on the filing's
period-end date.  That period-matching is the heart of this parser.

SCOPE
-----
Focuses on modern Ind-AS filings (2016+), which share the in-bse-fin schema.
Older Non-Ind-AS files use a different schema and are flagged & skipped.

Downloads are CACHED in processed/xbrl_cache/ so re-runs are instant and polite
to NSE's servers.

    python xbrl_parser.py            # parse all cached/available filings
    python xbrl_parser.py 20         # only the 20 most recent (quick test)
===============================================================================
"""

from __future__ import annotations

import ssl
import sys
import time
import urllib.request
import xml.etree.ElementTree as ET
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / "processed"
CACHE = OUT_DIR / "xbrl_cache"
CACHE.mkdir(parents=True, exist_ok=True)
SRC = OUT_DIR / "financial_results_clean.csv"

_HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
_SSL = ssl.create_default_context()
_SSL.check_hostname = False
_SSL.verify_mode = ssl.CERT_NONE

# XBRL concept local-names we want -> friendly column name.
# We match by LOCAL name (ignoring the namespace prefix) so the same code works
# across minor schema versions.
WANTED = {
    "RevenueFromOperations": "revenue",
    "OtherIncome": "other_income",
    "Income": "total_income",
    "ProfitBeforeTax": "profit_before_tax",
    "ProfitLossForPeriod": "net_profit",
    "BasicEarningsLossPerShareFromContinuingAndDiscontinuedOperations": "basic_eps",
}


# ---------------------------------------------------------------------------
def download(url: str) -> str | None:
    """Fetch an XBRL file, using a local cache. Returns the XML text or None."""
    fname = CACHE / url.rsplit("/", 1)[-1]
    if fname.exists():
        return fname.read_text(encoding="utf-8", errors="replace")
    try:
        req = urllib.request.Request(url, headers=_HEADERS)
        with urllib.request.urlopen(req, timeout=30, context=_SSL) as r:
            text = r.read().decode("utf-8", errors="replace")
        fname.write_text(text, encoding="utf-8")
        time.sleep(0.4)                       # be polite to NSE
        return text
    except Exception as e:                     # noqa: BLE001 - report and continue
        print(f"    ! download failed: {type(e).__name__}: {e}")
        return None


def local(tag: str) -> str:
    """Strip the '{namespace}' / 'prefix:' part of an XML tag -> bare local name."""
    if "}" in tag:
        tag = tag.split("}", 1)[1]
    if ":" in tag:
        tag = tag.split(":", 1)[1]
    return tag


def parse_contexts(root: ET.Element) -> dict[str, tuple[str, str, bool]]:
    """Map context id -> (startDate, endDate, has_dimension).

    `has_dimension` flags contexts carrying a segment/scenario member (e.g. a
    business segment). We want the plain company-total context, so we later skip
    dimensional ones."""
    ctx = {}
    for c in root.iter():
        if local(c.tag) != "context":
            continue
        cid = c.get("id")
        start = end = None
        dim = False
        for e in c.iter():
            ln = local(e.tag)
            if ln == "startDate":
                start = (e.text or "").strip()
            elif ln == "endDate":
                end = (e.text or "").strip()
            elif ln == "instant":
                start = end = (e.text or "").strip()
            elif ln == "explicitMember":
                dim = True
        if cid and end:
            ctx[cid] = (start or end, end, dim)
    return ctx


def period_context(contexts: dict, period_start: str, period_end: str) -> set[str]:
    """Pick the context id(s) matching THIS filing's own reporting period.

    Matches contexts whose end ~= period_end (toDate) AND start ~= period_start
    (fromDate). This works for quarters (~90d) AND annuals (~365d) alike, because
    we compare to the filing's declared dates instead of assuming a length.
    Skips dimensional (per-segment) contexts so we get the company total."""
    keep = set()
    try:
        ps, pe = pd.Timestamp(period_start), pd.Timestamp(period_end)
    except Exception:                          # noqa: BLE001
        return keep
    for cid, (start, end, dim) in contexts.items():
        if dim:
            continue                            # segment context, not the total
        try:
            s, e = pd.Timestamp(start), pd.Timestamp(end)
        except Exception:                      # noqa: BLE001
            continue
        if abs((e - pe).days) <= 3 and abs((s - ps).days) <= 6:
            keep.add(cid)
    return keep


def extract(xml_text: str, period_start: str, period_end: str) -> dict:
    """Pull the wanted headline figures for the filing's reporting period."""
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError as e:
        return {"_error": f"xml parse: {e}"}

    contexts = parse_contexts(root)
    q_ctx = period_context(contexts, period_start, period_end)
    if not q_ctx:
        return {"_error": "no matching period context"}

    out: dict[str, float] = {}
    for el in root.iter():
        ln = local(el.tag)
        col = WANTED.get(ln)
        if col is None or col in out:
            continue
        if el.get("contextRef") not in q_ctx:
            continue
        try:
            out[col] = float((el.text or "").replace(",", "").strip())
        except (TypeError, ValueError):
            continue
    return out


# ---------------------------------------------------------------------------
def main() -> None:
    if not SRC.exists():
        raise SystemExit("Run clean_events.py first (need financial_results_clean.csv)")

    df = pd.read_csv(SRC, dtype=str).fillna("")
    df = df[df["xbrl"].str.startswith("http")]                 # must have a link
    df = df.sort_values("broadCastDate", ascending=False)

    limit = int(sys.argv[1]) if len(sys.argv) > 1 else len(df)
    df = df.head(limit)
    print(f"Parsing {len(df)} XBRL filings (cache: {CACHE})\n")

    records, skipped = [], 0
    for _, row in df.iterrows():
        period_end = row["toDate"][:10] if row["toDate"] else ""
        period_start = row["fromDate"][:10] if row["fromDate"] else ""
        text = download(row["xbrl"])
        if text is None:
            skipped += 1
            continue
        # older Non-Ind-AS files use a different schema -> flag & skip
        if "in-bse-fin" not in text and "in-capmkt" not in text:
            skipped += 1
            continue
        figs = extract(text, period_start, period_end)
        if "_error" in figs:
            skipped += 1
            continue
        # money is reported in rupees -> convert to crore (1 crore = 10 million)
        rec = {
            "broadCastDate": row["broadCastDate"],
            "period_end": period_end,
            "relatingTo": row["relatingTo"],
            "consolidated": row["consolidated"],
            "audited": row["audited"],
            "revenue_cr": figs.get("revenue", float("nan")) / 1e7,
            "net_profit_cr": figs.get("net_profit", float("nan")) / 1e7,
            "pbt_cr": figs.get("profit_before_tax", float("nan")) / 1e7,
            "basic_eps": figs.get("basic_eps", float("nan")),
        }
        records.append(rec)

    if not records:
        raise SystemExit("No filings parsed (network blocked or all older schema).")

    fund = pd.DataFrame(records)
    fund["period_end"] = pd.to_datetime(fund["period_end"], errors="coerce")
    fund = fund.sort_values("period_end")
    fund.to_csv(OUT_DIR / "fundamentals.csv", index=False)

    print(f"\nParsed {len(fund)} filings, skipped {skipped}.")
    print(f"Saved -> {OUT_DIR}\\fundamentals.csv\n")

    # show the consolidated quarterly trend (the number analysts actually track)
    con = fund[fund["consolidated"] == "Consolidated"].dropna(subset=["net_profit_cr"])
    if not con.empty:
        print("Consolidated quarterly trend (Rs crore):")
        print(f"  {'period':>10s} {'quarter':>16s} {'revenue':>12s} "
              f"{'net_profit':>11s} {'EPS':>7s}")
        for _, r in con.tail(12).iterrows():
            print(f"  {r['period_end'].date()!s:>10s} {r['relatingTo']:>16s} "
                  f"{r['revenue_cr']:>12,.0f} {r['net_profit_cr']:>11,.0f} "
                  f"{r['basic_eps']:>7.2f}")


if __name__ == "__main__":
    main()
