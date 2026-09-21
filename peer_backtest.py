"""
peer_backtest.py  —  does peer read-across predict the actual move?
===============================================================================
For every stock that has ALREADY reported this quarter, the peer read-across
PREDICTION = average realised move of its sector peers that reported EARLIER
(strictly before the subject's result date — causally clean). The ACTUAL = the
subject's own realised move. We then compare predicted vs actual.

This uses real realised moves from the tracker — no news, no headline tone.

    python peer_backtest.py
===============================================================================
"""
from __future__ import annotations
import datetime as dt
import re
import statistics as st
import sys
from pathlib import Path

import openpyxl
sys.stdout.reconfigure(encoding="utf-8")
from expectation_bar import SECTORS, SYM_SECTOR

ROOT = Path(__file__).resolve().parent
XLSX = ROOT / "Nifty50_Q1Results_BehaviourEngine_UPDATED_06Aug_fixed.xlsx"


def _ev_date(val, D, E):
    if isinstance(val, dt.datetime):
        return val
    if isinstance(val, str) and val.startswith("=") and isinstance(D, dt.datetime) \
            and re.match(r"=D\d+-E\d+", val):
        return D - dt.timedelta(days=int(E))
    return D if isinstance(D, dt.datetime) else None


def load():
    """symbol -> (result_date, realised_move%) for reported stocks."""
    ws = openpyxl.load_workbook(XLSX).active
    d = {}
    for r in range(3, ws.max_row + 1):
        sym = (ws.cell(r, 2).value or "").strip()
        D = ws.cell(r, 4).value
        o = ws.cell(r, 15).value
        if sym and isinstance(D, dt.datetime) and isinstance(o, (int, float)):
            d[sym] = (D, float(o))
    return d


def _corr(pairs):
    if len(pairs) < 3:
        return None
    ps = [p for p, _ in pairs]; a = [x for _, x in pairs]
    mp, ma = st.mean(ps), st.mean(a)
    cov = sum((x - mp) * (y - ma) for x, y in pairs)
    dp = sum((x - mp) ** 2 for x in ps) ** .5
    da = sum((y - ma) ** 2 for y in a) ** .5
    return round(cov / (dp * da), 3) if dp and da else None


def compute() -> dict:
    d = load()
    rows = []
    for sym, (D, actual) in d.items():
        sector = SYM_SECTOR.get(sym)
        sect = [(p, d[p][1]) for p in SECTORS.get(sector, [])
                if p != sym and p in d and d[p][0] < D]              # earlier SAME-sector peers
        mkt = [(p, v[1]) for p, v in d.items() if p != sym and v[0] < D]  # any earlier reporter
        if sect:
            members, basis = sect, "sector"
        elif mkt:
            members, basis = mkt, "market"                          # fallback fills first-reporters
        else:
            members, basis = [], None                              # the very first reporter overall
        pred = round(st.mean(v for _, v in members), 2) if members else None
        sector_peers = [p for p in SECTORS.get(sector, []) if p != sym]
        used = {p for p, _ in sect}                              # sector peers that reported earlier
        rows.append({"symbol": sym, "sector": sector, "result": D.date().isoformat(),
                     "actual": round(actual, 2), "pred": pred, "basis": basis,
                     "n_peers": len(members),
                     "error": (round(actual - pred, 2) if pred is not None else None),
                     "sector_peers": sector_peers,
                     "used_peers": sorted(used),                 # the ones in the sector prediction
                     "members": ([{"peer": p, "ret": round(v, 2)} for p, v in sect]
                                 if basis == "sector" else [])})
    rows.sort(key=lambda x: x["result"])

    sect_pairs = [(r["pred"], r["actual"]) for r in rows if r["basis"] == "sector"]
    all_pairs = [(r["pred"], r["actual"]) for r in rows if r["pred"] is not None]
    mae = round(st.mean(abs(p - x) for p, x in all_pairs), 2) if all_pairs else None
    hit = (round(sum(1 for p, x in all_pairs if (p >= 0) == (x >= 0)) / len(all_pairs) * 100)
           if all_pairs else None)
    return {"rows": rows,
            "n": len(all_pairs), "n_sector": len(sect_pairs),
            "correlation": _corr(sect_pairs),        # clean sector-only signal
            "correlation_all": _corr(all_pairs),     # incl. market fallback
            "mae": mae, "direction_hit": hit,
            "note": "PREDICTED = mean realised move of sector peers that reported EARLIER; "
                    "if none, falls back to the market (all earlier reporters), tagged 'mkt'. "
                    "ACTUAL = subject's realised move. Real moves, one quarter."}


def main():
    d = compute()
    print(f"{'SYM':11}{'sector':16}{'result':>11}{'pred':>8}{'actual':>8}{'err':>8}{'peers':>6}")
    for r in d["rows"]:
        pr = "—" if r["pred"] is None else f"{r['pred']:+.2f}"
        er = "—" if r["error"] is None else f"{r['error']:+.2f}"
        print(f"{r['symbol']:11}{(r['sector'] or '—'):16}{r['result']:>11}{pr:>8}{r['actual']:>+8.2f}{er:>8}{r['n_peers']:>6}")
    print(f"\n  n filled = {d['n']} (sector-based {d['n_sector']}, rest market fallback)")
    print(f"  correlation sector-only = {d['correlation']}  |  incl. market fallback = {d['correlation_all']}")
    print(f"  mean abs error = {d['mae']}%   |   direction hit-rate = {d['direction_hit']}%")


if __name__ == "__main__":
    main()
