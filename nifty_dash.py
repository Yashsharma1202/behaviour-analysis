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

import json
import sys
import threading
import webbrowser
from urllib.parse import urlparse, parse_qs

import stock_server as S                      # reuse the whole dashboard engine
import exp_bar_host as X                       # Peer / Forward analysis (promoted onto main host)
from download_feeds import NIFTY50_FALLBACK   # the Nifty 50 symbol list

_N = len(set(NIFTY50_FALLBACK))
# rebrand the page header/title so it's clearly the Nifty 50 dashboard, and fix
# the full-universe stock-count hint to the Nifty 50 count.
NIFTY_PAGE = (S.PAGE
              .replace("<title>NSE Stock Browser</title>", "<title>Nifty 50 Browser</title>")
              .replace("NSE Stock Browser", "Nifty 50 Browser")
              .replace("search any of 2,363 stocks", f"search any of {_N} stocks"))

# --- promote the Peer / Forward analysis onto the main host (purely additive) --
# Serve the whole expectation-bar app at /peer, with its API mounted at /pa/* so
# it never collides with the main dashboard's own /api/* routes. The main page is
# untouched apart from one floating link.
_BACK_TO_MAIN = ('<a href="/" style="display:inline-block;margin-bottom:12px;background:#1f2937;'
                 'color:#e6edf3;font-weight:700;padding:8px 14px;border-radius:8px;'
                 'text-decoration:none;border:1px solid #2b3a4f;font-family:-apple-system,Segoe UI,sans-serif">'
                 '← Main Dashboard</a>')
PEER_PAGE = (X.PAGE.replace("/api/", "/pa/")
             .replace('<div class="wrap">', '<div class="wrap">' + _BACK_TO_MAIN, 1))
_PEER_LINK = ('<a href="/peer" target="_blank" style="position:fixed;bottom:16px;right:16px;'
              'z-index:9999;background:#58a6ff;color:#0b1220;font-weight:700;padding:9px 15px;'
              'border-radius:9px;text-decoration:none;font-family:-apple-system,Segoe UI,sans-serif;'
              'box-shadow:0 4px 14px rgba(0,0,0,.4)">📊 Peer / Forward Analysis ↗</a>')
NIFTY_PAGE = NIFTY_PAGE.replace("</body>", _PEER_LINK + "</body>")


class NiftyHandler(S.Handler):
    """Same handler as the full dashboard, but serves the rebranded page."""
    def do_GET(self):
        p = urlparse(self.path)
        if p.path in ("/", "/index.html"):
            self._send(200, NIFTY_PAGE, "text/html; charset=utf-8")
            return
        if p.path in ("/peer", "/peer/", "/peer.html"):
            self._send(200, PEER_PAGE, "text/html; charset=utf-8")
            return
        if p.path.startswith("/pa/"):                       # Peer / Forward analysis API
            obj = X.api(p.path[len("/pa/"):], parse_qs(p.query))
            body = json.dumps(obj if obj is not None else {"error": "not found"},
                              ensure_ascii=False)
            self._send(200 if obj is not None else 404, body, "application/json")
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
