"""
download_fundamentals.py
===============================================================================
Download Screener.in fundamentals for a stock and write the 5 statement CSVs in
the EXACT format of the existing folders (balance_sheet / cash_flow / pnl /
ratios / quarterly).

The public company page carries only the top-level rows; the indented sub-rows
(└─ Sales Growth %, └─ ROCE %, …) load from Screener's schedules API:
    /api/company/{warehouse_id}/schedules/?parent={row}&section={sec}&consolidated={0|1}
This script fetches both and re-assembles them into the wide "Metric, Mar YYYY…"
layout the rest of the project expects.

    python download_fundamentals.py RELIANCE            # validate vs existing file
    python download_fundamentals.py LTIM=540005 SBILIFE TATAMOTORS=TMCV
    python download_fundamentals.py --write LTIM=540005 # actually write the CSVs

Pass SYMBOL=slug to force the Screener page slug (e.g. LTIM=540005, TATAMOTORS=TMCV).
Without --write it prints a preview and (for existing symbols) a diff, writing
nothing.
===============================================================================
"""
from __future__ import annotations

import json
import re
import ssl
import sys
import time
import urllib.parse
import urllib.request
from pathlib import Path

import pandas as pd
from bs4 import BeautifulSoup

sys.stdout.reconfigure(encoding="utf-8")
ROOT = Path(__file__).resolve().parent

# section-id on the page  ->  (folder name, is-annual)
SECTIONS = [
    ("profit-loss",   "pnl",           True),
    ("balance-sheet", "balance_sheet", True),
    ("cash-flow",     "cash_flow",     True),
    ("ratios",        "ratios",        True),
    ("quarters",      "quarterly",     False),
]
_H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36"}
_SSL = ssl.create_default_context()
_SSL.check_hostname = False
_SSL.verify_mode = ssl.CERT_NONE


def _get(url: str, tries: int = 5) -> str:
    """GET with backoff — Screener rate-limits the schedules API under load."""
    for i in range(tries):
        try:
            req = urllib.request.Request(url, headers=_H)
            with urllib.request.urlopen(req, timeout=40, context=_SSL) as r:
                return r.read().decode("utf-8", "replace")
        except urllib.error.HTTPError as e:
            if e.code in (429, 503) and i < tries - 1:
                time.sleep(1.5 * (i + 1))            # back off and retry
                continue
            raise
        except Exception:                            # noqa: BLE001
            if i < tries - 1:
                time.sleep(1.0 * (i + 1))
                continue
            raise
    raise RuntimeError("unreachable")


def fetch_page(slug: str):
    """Return (soup, warehouse_id, consolidated_flag). Tries consolidated, then standalone."""
    for cons in (True, False):
        url = f"https://www.screener.in/company/{slug}/" + ("consolidated/" if cons else "")
        try:
            html = _get(url)
        except Exception:                                        # noqa: BLE001
            continue
        soup = BeautifulSoup(html, "html.parser")
        # the warehouse id used by the schedules API is NOT the URL slug (which may
        # be a BSE code like 540005). The Related-Party link /results/rpt/<id>/ and
        # any /api/company/<id>/ reference both carry the true warehouse id.
        m = (re.search(r"/results/rpt/(\d+)/", html) or
             re.search(r"/api/company/(\d+)/", html) or
             re.search(r'data-company-id["\']?\s*[:=]\s*["\']?(\d+)', html, re.I))
        sec = soup.find("section", id="profit-loss")
        # a variant is only usable if its P&L actually carries 'Mon YYYY' columns —
        # insurers/banks leave the CONSOLIDATED P&L empty and report on standalone.
        has_periods = bool(sec and sec.find("table") and any(
            re.match(r"[A-Za-z]{3}\s*\d{4}", th.get_text(strip=True))
            for th in sec.find("table").find_all("th")))
        if has_periods and m:
            return soup, int(m.group(1)), cons
    raise SystemExit(f"could not load a usable page with data for '{slug}'")


def _clean_label(td) -> str:
    txt = td.get_text(" ", strip=True)
    return re.sub(r"\s*\+\s*$", "", txt).strip()               # drop the "+" expand marker


def _val(x: str) -> str:
    """Match the existing files: strip thousands-commas, keep %/-/decimals."""
    return (x or "").replace(",", "").strip()


def schedule(wid: int, parent: str, section: str, cons: bool) -> dict:
    url = (f"https://www.screener.in/api/company/{wid}/schedules/"
           f"?parent={urllib.parse.quote(parent)}&section={section}"
           f"&consolidated={'true' if cons else 'false'}")
    try:
        return json.loads(_get(url))
    except Exception:                                            # noqa: BLE001
        return {}


def expand(wid, parent, section, cons, keep, depth, seen, out):
    """Recursively pull a row's schedule; some children (e.g. Material Cost %)
    have their own sub-schedule, so descend up to a few levels."""
    if depth > 3 or parent in seen:
        return
    seen.add(parent)
    sub = schedule(wid, parent, section, cons)
    time.sleep(0.2)
    for name, series in sub.items():
        out.append([f"└─ {name}"] + [_val(series.get(p, "")) for p in keep])
        expand(wid, name, section, cons, keep, depth + 1, seen, out)


def parse_section(soup, wid, sec_id, is_annual, cons) -> pd.DataFrame | None:
    sec = soup.find("section", id=sec_id)
    if not sec:
        return None
    table = sec.find("table")
    if not table:
        return None
    heads = [th.get_text(strip=True) for th in table.find("thead").find_all("th")]
    periods = heads[1:]                                          # first head is blank
    # annual files keep only the 'Mon YYYY' columns (drop TTM); quarterly keeps all
    keep = [p for p in periods if re.match(r"[A-Za-z]{3}\s*\d{4}", p)]
    colidx = [periods.index(p) for p in keep]

    rows = []
    for tr in table.find("tbody").find_all("tr"):
        tds = tr.find_all("td")
        if not tds:
            continue
        label = _clean_label(tds[0])
        if not label:
            continue
        vals = [_val(tds[1 + i].get_text(strip=True)) if 1 + i < len(tds) else "" for i in colidx]
        rows.append([label] + vals)
        # expandable? pull its schedule (recursively) and append the sub-rows
        btn = tr.find("button", onclick=re.compile("showSchedule"))
        if btn:
            m = re.search(r"showSchedule\('([^']+)',\s*'([^']+)'", btn["onclick"])
            if m:
                expand(wid, m.group(1), m.group(2), cons, keep, 1, set(), rows)
    df = pd.DataFrame(rows, columns=["Metric"] + keep)
    return df


def build(slug: str):
    soup, wid, cons = fetch_page(slug)
    name = soup.find("h1").get_text(strip=True) if soup.find("h1") else slug
    print(f"  page: {slug}  (warehouse id {wid}, "
          f"{'consolidated' if cons else 'standalone'})  — {name}")
    out = {}
    for sec_id, folder, is_annual in SECTIONS:
        df = parse_section(soup, wid, sec_id, is_annual, cons)
        if df is not None and len(df.columns) > 1:
            out[folder] = df
        time.sleep(0.3)
    return out, name


def main():
    args = sys.argv[1:]
    write = "--write" in args
    targets = [a for a in args if a != "--write"]
    if not targets:
        raise SystemExit("usage: python download_fundamentals.py [--write] SYMBOL[=slug] …")

    for tgt in targets:
        sym, _, slug = tgt.partition("=")
        slug = slug or sym
        print(f"\n=== {sym} ===")
        out, name = build(slug)
        for folder, df in out.items():
            path = ROOT / folder / f"{sym}.csv"
            exists = path.exists()
            note = ""
            if exists:                                           # validation diff
                old = pd.read_csv(path, dtype=str).fillna("")
                note = (f"  [existing has {len(old)} rows × {len(old.columns)-1} periods]")
            print(f"    {folder:<14} {len(df):>2} rows × {len(df.columns)-1} periods{note}")
            if write:
                df.to_csv(path, index=False)
        if write:
            print(f"    -> written into the 5 folders as {sym}.csv")
        else:
            # show a sample so the format can be eyeballed before writing
            if "pnl" in out:
                print("    pnl preview:")
                print(out["pnl"].head(4).to_string(index=False, max_colwidth=14)[:400])
    if not write:
        print("\n(dry run — nothing written. re-run with --write to save.)")


if __name__ == "__main__":
    main()
