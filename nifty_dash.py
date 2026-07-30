"""
nifty_dash.py
===============================================================================
A SEPARATE dashboard for ONLY the Nifty 50 stocks — same UI/features as the full
stock_server.py dashboard (search, fundamentals, event tabs, 📈 Behaviour tab),
but on its own port so it never mixes with the 2,363-stock dashboard.

    python nifty_dash.py             # http://localhost:8090
    python nifty_dash.py 9000        # custom port
    python nifty_dash.py --no-open   # don't auto-open the browser

Shares the same underlying data (fundamentals folders + downloaded <SYMBOL>/
feeds) but only LISTS the Nifty 50 constituents.
===============================================================================
"""

from __future__ import annotations

import sys
import threading
import webbrowser
from urllib.parse import urlparse

import stock_server as S                      # reuse the whole dashboard engine
from download_feeds import NIFTY50_FALLBACK   # the Nifty 50 symbol list

_N = len(set(NIFTY50_FALLBACK))
# rebrand the page header/title so it's clearly the Nifty 50 dashboard, and fix
# the full-universe stock-count hint to the Nifty 50 count.
NIFTY_PAGE = (S.PAGE
              .replace("<title>NSE Stock Browser</title>", "<title>Nifty 50 Browser</title>")
              .replace("NSE Stock Browser", "Nifty 50 Browser")
              .replace("search any of 2,363 stocks", f"search any of {_N} stocks"))


class NiftyHandler(S.Handler):
    """Same handler as the full dashboard, but serves the rebranded page."""
    def do_GET(self):
        if urlparse(self.path).path in ("/", "/index.html"):
            self._send(200, NIFTY_PAGE, "text/html; charset=utf-8")
            return
        super().do_GET()


def main():
    port = 8090
    for a in sys.argv[1:]:
        if a.isdigit():
            port = int(a)

    # override the symbol universe -> only Nifty 50 (everything else reused as-is)
    S.SYMBOLS = sorted(set(NIFTY50_FALLBACK))

    # warm the expensive rankings/screener caches in the background so the
    # Rankings/Screener tabs return instantly instead of hanging on first click.
    def _warm():
        try:
            S.build_rankings(); S.build_screener()
            print("  caches warm — Rankings & Screener ready.")
        except Exception as e:                       # noqa: BLE001
            print(f"  cache warm failed: {e}")
    threading.Thread(target=_warm, daemon=True).start()

    lan = "--lan" in sys.argv
    host = "0.0.0.0" if lan else "127.0.0.1"
    ip = S.lan_ip() if lan else "localhost"
    url = f"http://{ip}:{port}"
    server = S.ThreadingHTTPServer((host, port), NiftyHandler)
    print("=" * 62)
    print(f"  NIFTY 50 Dashboard  —  {len(S.SYMBOLS)} stocks")
    print(f"  Open:  {url}" + ("   (reachable from other devices on your LAN)" if lan else ""))
    print("  (separate host from the 2,363-stock dashboard)")
    print("  Press Ctrl+C to stop.")
    print("=" * 62)

    if "--no-open" not in sys.argv:
        threading.Timer(0.6, lambda: webbrowser.open(f"http://localhost:{port}")).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")
        server.shutdown()


if __name__ == "__main__":
    main()
