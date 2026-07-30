"""
build_static_nifty.py
===============================================================================
Bakes the NIFTY-50 dashboard (nifty_dash.py / stock_server.py) into a STATIC
site under ./docs/nifty/ that GitHub Pages can host — Pages can't run the Python
server or hit NSE, so every dynamic endpoint is pre-rendered to a JSON file and
the frontend's fetch() calls are rewritten to read those files.

Purely additive: writes a NEW folder (docs/nifty/). It never touches the existing
full-universe snapshot at docs/index.html.

It writes:
    docs/nifty/index.html                  (the Nifty-50 UI, data source -> files)
    docs/nifty/data/symbols.json           {symbols, count, ready}
    docs/nifty/data/fund/<SYM>.json        fundamentals payload  (/api/stock)
    docs/nifty/data/events/<SYM>.json      event feeds           (/api/events)
    docs/nifty/data/behaviour/<SYM>.json   behaviour payload     (/api/behaviour)
    docs/nifty/data/rankings.json          rankings tab          (/api/rankings)
    docs/nifty/data/screener.json          screener tab          (/api/screener)

RUN
    python build_static_nifty.py
    python -m http.server -d docs/nifty 9001   # preview -> http://localhost:9001
===============================================================================
"""

from __future__ import annotations

import json
import re
import sys
import warnings
from pathlib import Path

warnings.filterwarnings("ignore")
sys.stdout.reconfigure(encoding="utf-8")

import stock_server as S                       # reuse discovery + payload builders
from download_feeds import NIFTY50_FALLBACK    # the Nifty 50 symbol list

# ── restrict the universe to the Nifty 50 (same as nifty_dash.py) ─────────────
S.SYMBOLS = sorted(set(NIFTY50_FALLBACK))

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "docs" / "nifty"
DATA = OUT / "data"
FUND = DATA / "fund"
EVENTS = DATA / "events"
BEHAV = DATA / "behaviour"
for d in (FUND, EVENTS, BEHAV):
    d.mkdir(parents=True, exist_ok=True)


def _write(path: Path, obj) -> None:
    path.write_text(json.dumps(obj, ensure_ascii=False), encoding="utf-8")


# --- the static index.html: take the live PAGE and swap every data source -----
# _V is a per-build cache-buster appended to every data URL, so browsers/CDN can
# never serve a stale symbols/rankings/… file after a rebuild.
STATIC_BLOCK = """/*__DATA_SOURCE__*/
const STATIC=true;
const _V="?v=__BUILD__";
const SYMBOLS_URL="./data/symbols.json"+_V;
function STOCK_URL(s){return "./data/fund/"+encodeURIComponent(s)+".json"+_V;}
function EVENTS_URL(s){return "./data/events/"+encodeURIComponent(s)+".json"+_V;}
/*__END_DATA_SOURCE__*/"""


def build_index() -> None:
    import time
    build = str(int(time.time()))          # per-build cache-buster version
    page = S.PAGE
    # 1) rebrand to the Nifty 50 (mirror nifty_dash.py)
    page = (page.replace("<title>NSE Stock Browser</title>", "<title>Nifty 50 Browser</title>")
                .replace("NSE Stock Browser", "Nifty 50 Browser")
                .replace("search any of 2,363 stocks", f"search any of {len(S.SYMBOLS)} stocks"))
    # 2) swap the symbols/stock/events data source to static files (+cache-buster)
    page = re.sub(r"/\*__DATA_SOURCE__\*/.*?/\*__END_DATA_SOURCE__\*/",
                  STATIC_BLOCK.replace("__BUILD__", build), page, flags=re.DOTALL)
    # 3) rewrite the remaining /api/ fetches to the pre-baked JSON files (+cache-buster)
    page = page.replace('fetch("/api/rankings")', 'fetch("./data/rankings.json"+_V)')
    page = page.replace('fetch("/api/screener")', 'fetch("./data/screener.json"+_V)')
    page = page.replace('fetch("/api/upcoming")', 'fetch("./data/upcoming.json"+_V)')
    page = page.replace('"/api/behaviour?sym="+encodeURIComponent(sym))',
                        '"./data/behaviour/"+encodeURIComponent(sym)+".json"+_V)')
    # 4) snapshot banner after the tab bar
    note = ('<div class="hint" style="margin:0 0 10px">📦 Static snapshot for '
            'GitHub Pages — the Nifty 50 with fundamentals, event feeds, behaviour, '
            'rankings, screener &amp; compare. Run <b>python nifty_dash.py</b> locally '
            'for live NSE downloads.</div>')
    page = page.replace('<div class="tabs" id="tabs">', note + '\n    <div class="tabs" id="tabs">')

    (OUT / "index.html").write_text(page, encoding="utf-8")
    # sanity: no live API path should survive in the baked page
    leftovers = re.findall(r'fetch\("/api/[^"]*', page)
    leftovers = [l for l in leftovers if "/api/fetch?sym=" not in l]  # fetch is STATIC-gated, ok
    print("  wrote docs/nifty/index.html"
          + (f"   ⚠ leftover api calls: {leftovers}" if leftovers else ""))


def main() -> None:
    symbols = S.SYMBOLS
    ready = S.downloaded_symbols()
    ready = [s for s in ready if s in set(symbols)]          # only Nifty ones
    print(f"Baking Nifty-50 static site: {len(symbols)} stocks, "
          f"{len(ready)} with event feeds")

    _write(DATA / "symbols.json",
           {"symbols": symbols, "count": len(symbols), "ready": ready})

    # fundamentals for every stock
    fail = 0
    for i, sym in enumerate(symbols, 1):
        try:
            payload = S.build_payload(sym)
        except Exception as e:                               # noqa: BLE001
            fail += 1
            payload = {"symbol": sym, "have": [], "cards": {},
                       "annual": {"labels": [], "sales": [], "net_profit": []},
                       "quarterly": {"labels": [], "net_profit": [], "yoy": []},
                       "error": str(e)}
        _write(FUND / f"{sym}.json", payload)

    # events for downloaded stocks
    for sym in ready:
        try:
            _write(EVENTS / f"{sym}.json", S.build_events(sym))
        except Exception as e:                               # noqa: BLE001
            print(f"  ! events {sym}: {e}")

    # behaviour payload for every stock (Behaviour + Compare tabs)
    bfail = 0
    for i, sym in enumerate(symbols, 1):
        try:
            _write(BEHAV / f"{sym}.json", S.build_behaviour(sym))
        except Exception as e:                               # noqa: BLE001
            bfail += 1
            _write(BEHAV / f"{sym}.json", {"symbol": sym, "events": {}, "error": str(e)})
        if i % 10 == 0 or i == len(symbols):
            print(f"  behaviour {i}/{len(symbols)}")

    # rankings + screener (whole-universe payloads)
    print("  building rankings…")
    _write(DATA / "rankings.json", S.build_rankings())
    print("  building screener…")
    _write(DATA / "screener.json", S.build_screener())
    print("  building upcoming…")
    _write(DATA / "upcoming.json", S.build_upcoming())

    build_index()

    total = sum(1 for _ in DATA.rglob("*.json"))
    size = sum(f.stat().st_size for f in OUT.rglob("*") if f.is_file()) / 1_048_576
    print(f"\nDone. {total} json files, docs/nifty/ ~= {size:.1f} MB")
    print("Preview:  python -m http.server -d docs/nifty 9001  -> http://localhost:9001")
    if fail:
        print(f"({fail} stocks had no parseable fundamentals — still browsable)")
    if bfail:
        print(f"({bfail} stocks had no behaviour payload — tab shows empty for them)")


if __name__ == "__main__":
    main()
