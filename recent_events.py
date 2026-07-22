"""
recent_events.py
===============================================================================
Show the most RECENT corporate events across the stocks whose NSE feeds you've
downloaded — results, board meetings, corporate actions and announcements —
newest first, plus any UPCOMING (future-dated) events flagged at the top.

    python recent_events.py             # all downloaded stocks, last 60 days
    python recent_events.py 30          # last 30 days
    python recent_events.py RELIANCE    # one stock (all recent)
    python recent_events.py RELIANCE 90 # one stock, last 90 days

Output: printed list + processed/recent_events.csv
===============================================================================
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:                     # noqa: BLE001
    pass

ROOT = Path(__file__).resolve().parent
PROC = ROOT / "processed"
PROC.mkdir(exist_ok=True)

DEFAULT_DAYS = 60

# feed file -> (label, date column, function building a short detail string)
FEEDS = {
    "financial_results": ("RESULTS", "broadCastDate",
                          lambda r: f"{r.get('relatingTo', '')} results "
                                    f"({r.get('consolidated', '')})".strip()),
    "board_meetings":    ("BOARD MEETING", "bm_date",
                          lambda r: (r.get("bm_purpose") or r.get("bm_desc") or "")[:80]),
    "corporate_actions": ("CORP ACTION", "exDate",
                          lambda r: (r.get("subject") or "")[:80]),
    "announcements":     ("ANNOUNCEMENT", "sort_date",
                          lambda r: (f"{r.get('desc', '')} — "
                                     f"{(r.get('attchmntText') or '')[:60]}").strip(" —")),
}


def feed_symbols() -> list[str]:
    out = []
    for p in ROOT.iterdir():
        if p.is_dir() and any((p / f"{k}.csv").exists() for k in FEEDS):
            out.append(p.name)
    return sorted(out)


def collect(symbols: list[str]) -> pd.DataFrame:
    rows = []
    for sym in symbols:
        for feed, (label, datecol, detail) in FEEDS.items():
            path = ROOT / sym / f"{feed}.csv"
            if not path.exists():
                continue
            df = pd.read_csv(path, dtype=str).fillna("")
            if datecol not in df.columns:
                continue
            dates = pd.to_datetime(df[datecol], errors="coerce",
                                   dayfirst=True, format="mixed")
            for (_, r), d in zip(df.iterrows(), dates):
                if pd.isna(d):
                    continue
                rows.append({"date": d.normalize(), "symbol": sym,
                             "event_type": label, "detail": detail(r)})
    if not rows:
        return pd.DataFrame(columns=["date", "symbol", "event_type", "detail"])
    return (pd.DataFrame(rows).drop_duplicates()
            .sort_values("date", ascending=False).reset_index(drop=True))


def show(df: pd.DataFrame, title: str, limit: int | None = None) -> None:
    if df.empty:
        return
    print(f"\n  {title}")
    print("  " + "-" * 96)
    for _, r in (df.head(limit) if limit else df).iterrows():
        print(f"  {r['date'].date()!s:>11}  {r['symbol']:<12} {r['event_type']:<14} "
              f"{str(r['detail'])[:60]}")


def main():
    args = sys.argv[1:]
    days = next((int(a) for a in args if a.isdigit()), DEFAULT_DAYS)
    syms = [a.upper() for a in args if not a.isdigit()] or feed_symbols()
    if not syms:
        raise SystemExit("No downloaded event feeds found. Run "
                         "'python download_feeds.py nifty50' first.")

    ev = collect(syms)
    if ev.empty:
        raise SystemExit("No events found in the downloaded feeds.")

    today = pd.Timestamp.today().normalize()
    cutoff = today - pd.Timedelta(days=days)
    upcoming = ev[ev["date"] > today]
    recent = ev[(ev["date"] <= today) & (ev["date"] >= cutoff)]

    ev.to_csv(PROC / "recent_events.csv", index=False)

    print("=" * 100)
    print(f"  RECENT EVENTS  ·  {len(syms)} stock(s)  ·  last {days} days "
          f"(as of {today.date()})")
    print("=" * 100)

    show(upcoming.sort_values("date"), f"UPCOMING / FUTURE-DATED  ({len(upcoming)})")
    show(recent, f"RECENT  ({len(recent)} in the last {days} days)")

    if recent.empty and upcoming.empty:
        last = ev.iloc[0]["date"].date()
        print(f"\n  No events in the last {days} days. Most recent on file: {last}.")
        show(ev, "LATEST ON FILE", limit=25)

    # per-type counts in the recent window
    if not recent.empty:
        print("\n  Recent breakdown by type:")
        for t, n in recent["event_type"].value_counts().items():
            print(f"     {n:>3}  {t}")

    print(f"\n  Full list saved -> {PROC}\\recent_events.csv")


if __name__ == "__main__":
    main()
