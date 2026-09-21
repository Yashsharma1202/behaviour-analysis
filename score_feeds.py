"""
score_feeds.py
===============================================================================
Batch-score every downloaded announcement feed with the FinBERT router, and
hold the results in a resumable on-disk cache.

WHY A CACHE AND NOT INLINE SCORING
----------------------------------
stock_server.build_events() used to call the lexicon on every announcement row
on every /api/events request. That is fine for a regex and fatal for a
transformer: RELIANCE alone is 3,311 rows, and the full downloaded universe is
~107k. So the model runs HERE, once, offline, and the dashboard does a dict
lookup.

The cache key is md5(desc + attchmntText), so re-running only scores rows that
are genuinely new — download more feeds, re-run, and it picks up the delta.

STORED PER ROW
    key         md5 of the scored text (the join key back to the feed CSV)
    date        sort_date  (the announcement date, for rollups)
    filed_at    exchdisstime / an_dt  — the EXCHANGE DISSEMINATION TIMESTAMP.
                Phase 5 needs this: a filing disclosed at 21:40 cannot be a
                feature for that day's close, and event_ml.py must know that.
    category    the `desc` column, kept so the gate can be audited
    label/score/confidence/engine   the router's verdict

RUN
    python score_feeds.py                    # every symbol with a feed on disk
    python score_feeds.py RELIANCE TCS       # just these
    python score_feeds.py --force            # rescore everything from scratch
    python score_feeds.py --nifty            # the Nifty-50 list only
===============================================================================
"""

from __future__ import annotations

import argparse
import hashlib
import sys
import time
from pathlib import Path

import pandas as pd

import sentiment_finbert as SF

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:                                 # noqa: BLE001
    pass

ROOT = Path(__file__).resolve().parent
CACHE_DIR = ROOT / "processed" / "sentiment"
CACHE_COLS = ["key", "date", "filed_at", "category",
              "label", "score", "confidence", "engine"]


# --------------------------------------------------------------------------- #
# Cache store  (imported by stock_server.py and, later, by event_ml.py)
# --------------------------------------------------------------------------- #
def text_key(desc: str, text: str) -> str:
    """Stable cache key for one announcement row."""
    blob = f"{desc or ''}\x00{text or ''}".encode("utf-8", "replace")
    return hashlib.md5(blob).hexdigest()


def cache_path(sym: str) -> Path:
    return CACHE_DIR / f"{sym}.csv"


def load_cache(sym: str) -> dict[str, dict]:
    """key -> {label, score, confidence, engine, ...}. Empty dict when absent,
    so callers can always fall back to the lexicon without special-casing."""
    path = cache_path(sym)
    if not path.exists():
        return {}
    try:
        df = pd.read_csv(path, dtype=str).fillna("")
    except Exception:                             # noqa: BLE001
        return {}
    out = {}
    for r in df.to_dict("records"):
        try:
            r["score"] = float(r.get("score") or 0.0)
            r["confidence"] = float(r.get("confidence") or 0.0)
        except (TypeError, ValueError):
            r["score"], r["confidence"] = 0.0, 0.0
        out[r["key"]] = r
    return out


def save_cache(sym: str, rows: list[dict]) -> None:
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    pd.DataFrame(rows, columns=CACHE_COLS).to_csv(cache_path(sym), index=False)


def cached_symbols() -> list[str]:
    return sorted(p.stem for p in CACHE_DIR.glob("*.csv")) if CACHE_DIR.exists() else []


# --------------------------------------------------------------------------- #
# Rollups — the features Phase 3 (scores) and Phase 5 (ML) consume
# --------------------------------------------------------------------------- #
def parse_dt(s: pd.Series) -> pd.Series:
    """Parse NSE feed dates, which come in TWO shapes in the same file:

        sort_date     2026-07-10 17:46:25     ISO
        an_dt         10-Jul-2026 17:46:25    day-first

    Passing dayfirst=True at both is what the rest of this repo does, and it
    silently corrupts the ISO form: pandas reads "2026-07-10" as
    year-DAY-month and returns 10 October 2026. On RELIANCE that pushed 14
    filings months into the future. Parse ISO first, day-first for the rest.
    """
    s = s.astype("string")
    out = pd.to_datetime(s, errors="coerce", format="ISO8601")
    miss = out.isna() & s.notna() & (s.str.strip() != "")
    if miss.any():
        out[miss] = pd.to_datetime(s[miss], errors="coerce",
                                   dayfirst=True, format="mixed")
    return out


def sentiment_frame(sym: str) -> pd.DataFrame:
    """Cached sentiment as a dated frame, oldest first. Empty if never scored."""
    cache = load_cache(sym)
    if not cache:
        return pd.DataFrame(columns=["date", "filed_at", "category", "label",
                                     "score", "confidence", "engine"])
    df = pd.DataFrame(list(cache.values()))
    df["date"] = parse_dt(df["date"])
    df["filed_at"] = parse_dt(df["filed_at"])
    df["score"] = pd.to_numeric(df["score"], errors="coerce").fillna(0.0)
    return df.dropna(subset=["date"]).sort_values("date").reset_index(drop=True)


def sentiment_series(sym: str, half_lives=(30, 90, 180)) -> dict:
    """Per-stock sentiment rollup: EWMA at several half-lives plus recent counts.

    Only rows the router actually judged (engine != 'rule') carry information —
    the boilerplate gate emits a structural Neutral, and averaging thousands of
    those in would drag every stock's EWMA to zero regardless of its news.
    """
    df = sentiment_frame(sym)
    out = {"n": int(len(df)), "n_material": 0, "last_date": None,
           "ewma": {}, "counts": {}, "sent_surprise": None}
    if df.empty:
        return out

    mat = df[df["engine"] != "rule"]
    out["n_material"] = int(len(mat))
    out["last_date"] = str(df["date"].iloc[-1].date())
    if mat.empty:
        return out

    # EWMA in CALENDAR time, not row count — filings are bursty, so a plain
    # .ewm(span=) would let a busy fortnight outweigh a quiet quarter.
    t = mat["date"].astype("int64") / 86_400_000_000_000      # ns -> days
    now = float(t.iloc[-1])
    s = mat["score"].to_numpy(dtype=float)
    for hl in half_lives:
        w = 0.5 ** ((now - t.to_numpy(dtype=float)) / hl)
        tot = float(w.sum())
        out["ewma"][f"ewma_{hl}"] = round(float((w * s).sum() / tot), 4) if tot else 0.0

    for win in (30, 90, 365):
        cut = mat["date"].iloc[-1] - pd.Timedelta(days=win)
        sub = mat[mat["date"] >= cut]
        out["counts"][f"d{win}"] = {
            "n": int(len(sub)),
            "pos": int((sub["label"] == "Positive").sum()),
            "neg": int((sub["label"] == "Negative").sum()),
        }

    # How far the latest material filing sits from the stock's own recent tone.
    base = out["ewma"].get("ewma_90", 0.0)
    out["sent_surprise"] = round(float(s[-1]) - float(base), 4)
    return out


# --------------------------------------------------------------------------- #
# Scoring one symbol
# --------------------------------------------------------------------------- #
def score_symbol(sym: str, force: bool = False) -> dict:
    """Score <sym>/announcements.csv, reusing whatever is already cached."""
    path = ROOT / sym / "announcements.csv"
    stat = {"symbol": sym, "rows": 0, "new": 0, "reused": 0, "gated": 0,
            "engines": {}, "seconds": 0.0}
    if not path.exists():
        stat["skipped"] = "no announcements.csv"
        return stat

    try:
        df = pd.read_csv(path, dtype=str).fillna("")
    except Exception as e:                        # noqa: BLE001
        stat["skipped"] = f"{type(e).__name__}: {e}"
        return stat
    if df.empty:
        return stat

    for c in ("desc", "attchmntText", "sort_date", "exchdisstime", "an_dt"):
        if c not in df.columns:
            df[c] = ""

    stat["rows"] = len(df)
    cache = {} if force else load_cache(sym)

    keys = [text_key(d, t) for d, t in zip(df["desc"], df["attchmntText"])]
    df["_key"] = keys

    # De-duplicate: the same boilerplate text repeats hundreds of times within a
    # single symbol, and there is no reason to score it more than once.
    todo: dict[str, tuple[str, str]] = {}
    for k, d, t in zip(keys, df["desc"], df["attchmntText"]):
        if k not in cache and k not in todo:
            todo[k] = (d, t)
    stat["reused"] = len(set(keys)) - len(todo)

    t0 = time.time()
    if todo:
        tk = list(todo)
        cats = [todo[k][0] for k in tk]
        txts = [f"{todo[k][0]} {todo[k][1]}".strip() for k in tk]
        res = SF.score_many(txts, cats)
        for k, cat, r in zip(tk, cats, res):
            cache[k] = {"key": k, "date": "", "filed_at": "", "category": cat,
                        "label": r.label, "score": r.score,
                        "confidence": r.confidence, "engine": r.engine}
        stat["new"] = len(todo)
    stat["seconds"] = round(time.time() - t0, 2)

    # Attach dates/timestamps from the feed (newest wins for repeated text).
    for k, sd, ed, ad in zip(keys, df["sort_date"], df["exchdisstime"], df["an_dt"]):
        rec = cache.get(k)
        if rec is not None:
            rec["date"] = sd or ad or rec.get("date", "")
            rec["filed_at"] = ed or ad or rec.get("filed_at", "")

    save_cache(sym, [{c: cache[k].get(c, "") for c in CACHE_COLS}
                     for k in cache])

    for rec in cache.values():
        eng = rec.get("engine", "?")
        stat["engines"][eng] = stat["engines"].get(eng, 0) + 1
    stat["gated"] = stat["engines"].get("rule", 0)
    return stat


# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("symbols", nargs="*", help="symbols (default: all downloaded)")
    ap.add_argument("--force", action="store_true", help="rescore, ignore the cache")
    ap.add_argument("--nifty", action="store_true", help="the Nifty-50 list only")
    args = ap.parse_args()

    if args.symbols:
        syms = [s.strip().upper() for s in args.symbols]
    elif args.nifty:
        from download_feeds import NIFTY50_FALLBACK
        syms = sorted(set(NIFTY50_FALLBACK))
    else:
        syms = sorted(p.parent.name for p in ROOT.glob("*/announcements.csv"))

    info = SF.model_info()
    print("=" * 78)
    print("  SENTIMENT SCORING — FinBERT router")
    print(f"  model  : {info['model']}")
    print(f"  device : {info.get('device') or 'n/a'}"
          f"{'' if info['available'] else '   (UNAVAILABLE — lexicon fallback)'}")
    print(f"  symbols: {len(syms)}")
    print("=" * 78)

    tot = {"rows": 0, "new": 0, "reused": 0}
    engines: dict[str, int] = {}
    t_all = time.time()

    for i, sym in enumerate(syms, 1):
        st = score_symbol(sym, force=args.force)
        if st.get("skipped"):
            print(f"[{i:>3}/{len(syms)}] {sym:<14} skipped ({st['skipped']})")
            continue
        for k in tot:
            tot[k] += st[k]
        for e, n in st["engines"].items():
            engines[e] = engines.get(e, 0) + n
        rate = f"{st['new'] / st['seconds']:,.0f}/s" if st["seconds"] > 0.05 else "-"
        print(f"[{i:>3}/{len(syms)}] {sym:<14} rows={st['rows']:>6,}  "
              f"new={st['new']:>6,}  reused={st['reused']:>6,}  "
              f"{st['seconds']:>6.2f}s  {rate:>9}")

    el = time.time() - t_all
    scored = sum(engines.values())
    print("=" * 78)
    print(f"  {tot['rows']:,} feed rows across {len(syms)} symbols in {el:.1f}s")
    print(f"  {tot['new']:,} newly scored, {tot['reused']:,} reused from cache")
    if scored:
        print("\n  engine breakdown (unique texts):")
        for e, n in sorted(engines.items(), key=lambda kv: -kv[1]):
            note = {"rule": "boilerplate gate — model never ran",
                    "numeric": "templated results filing — decided by arithmetic",
                    "finbert": "model verdict",
                    "override": "rule overruled the model",
                    "lexicon": "model unavailable, lexicon fallback"}.get(e, "")
            print(f"    {e:<10} {n:>7,}  {n / scored * 100:>5.1f}%   {note}")
    print(f"\n  cache -> {CACHE_DIR}")


if __name__ == "__main__":
    main()
