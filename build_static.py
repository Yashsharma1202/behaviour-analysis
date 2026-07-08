"""
build_static.py
===============================================================================
Bakes the live dashboard (stock_server.py) into a STATIC site under ./docs/ that
GitHub Pages can host — because Pages can't run the Python server or hit NSE.

It writes:
    docs/index.html                 (the same UI, data source swapped to files)
    docs/data/symbols.json          {symbols, count, ready}
    docs/data/fund/<SYM>.json       fundamentals payload for every stock
    docs/data/events/<SYM>.json     event feeds for stocks already downloaded

RUN
    python build_static.py
===============================================================================
"""

from __future__ import annotations

import json
import re
from pathlib import Path

import stock_server as S           # reuse discovery + payload builders

ROOT = Path(__file__).resolve().parent
DOCS = ROOT / "docs"
DATA = DOCS / "data"
FUND = DATA / "fund"
EVENTS = DATA / "events"
for d in (FUND, EVENTS):
    d.mkdir(parents=True, exist_ok=True)


def _write(path: Path, obj) -> None:
    path.write_text(json.dumps(obj, ensure_ascii=False), encoding="utf-8")


# --- the static index.html: take the live PAGE and swap the data source -------
STATIC_BLOCK = """/*__DATA_SOURCE__*/
const STATIC=true;
const SYMBOLS_URL="./data/symbols.json";
function STOCK_URL(s){return "./data/fund/"+encodeURIComponent(s)+".json";}
function EVENTS_URL(s){return "./data/events/"+encodeURIComponent(s)+".json";}
/*__END_DATA_SOURCE__*/"""


def build_index() -> None:
    page = S.PAGE
    page = re.sub(r"/\*__DATA_SOURCE__\*/.*?/\*__END_DATA_SOURCE__\*/",
                  STATIC_BLOCK, page, flags=re.DOTALL)
    # a small banner noting this is a snapshot (after the tab bar)
    note = ('<div class="hint" style="margin:0 0 10px">📦 Static snapshot for '
            'GitHub Pages — fundamentals for all stocks + events for '
            'pre-downloaded ones. Run <b>python stock_server.py</b> locally for '
            'live NSE downloads.</div>')
    page = page.replace('<div class="tabs" id="tabs">', note + '\n    <div class="tabs" id="tabs">')
    (DOCS / "index.html").write_text(page, encoding="utf-8")
    print("  wrote docs/index.html")


def main() -> None:
    symbols = S.SYMBOLS
    ready = S.downloaded_symbols()
    print(f"Baking static site: {len(symbols)} stocks, {len(ready)} with event feeds")

    _write(DATA / "symbols.json",
           {"symbols": symbols, "count": len(symbols), "ready": ready})

    # fundamentals for every stock
    fail = 0
    for i, sym in enumerate(symbols, 1):
        try:
            payload = S.build_payload(sym)
        except Exception as e:                       # noqa: BLE001
            fail += 1
            payload = {"symbol": sym, "have": [], "cards": {},
                       "annual": {"labels": [], "sales": [], "net_profit": []},
                       "quarterly": {"labels": [], "net_profit": [], "yoy": []},
                       "error": str(e)}
        _write(FUND / f"{sym}.json", payload)
        if i % 400 == 0 or i == len(symbols):
            print(f"  fundamentals {i}/{len(symbols)}")

    # events for already-downloaded stocks
    for sym in ready:
        try:
            _write(EVENTS / f"{sym}.json", S.build_events(sym))
        except Exception as e:                       # noqa: BLE001
            print(f"  ! events {sym}: {e}")
    print(f"  events baked for: {', '.join(ready)}")

    build_index()

    total = sum(1 for _ in DATA.rglob("*.json"))
    size = sum(f.stat().st_size for f in DOCS.rglob("*")) / 1_048_576
    print(f"\nDone. {total} json files, docs/ ~= {size:.1f} MB")
    print("Preview locally:  python -m http.server -d docs 9000  -> http://localhost:9000")
    if fail:
        print(f"({fail} stocks had no parseable fundamentals — still browsable)")


if __name__ == "__main__":
    main()
