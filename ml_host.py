"""
ml_host.py
===============================================================================
The ML experiment on its OWN full dashboard — same look & feel as the main Nifty
50 dashboard (search, per-stock nav, all tabs) PLUS an extra "🤖 ML Signal" tab —
served on a SEPARATE port (8095). It reuses stock_server's engine at runtime but
NEVER modifies it, so the main dashboard on port 8090 is completely unaffected.

    python ml_predict.py     # first: train + write ml_output.json
    python ml_host.py        # then: serve the ML dashboard -> http://localhost:8095
    python ml_host.py --lan  # reachable from other devices on your LAN
===============================================================================
"""
from __future__ import annotations

import json
import sys
import threading
import webbrowser
from pathlib import Path
from urllib.parse import urlparse, parse_qs

import stock_server as S
from download_feeds import NIFTY50_FALLBACK
import ml_tab                                  # shared ML-tab injection + payload

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:                              # noqa: BLE001
    pass

ROOT = Path(__file__).resolve().parent
PORT = 8095

# restrict the engine to the Nifty 50 (same universe as the main dashboard)
S.SYMBOLS = sorted(set(NIFTY50_FALLBACK))

# build the ML dashboard page: rebrand + inject the shared ML tab (ml_tab.py)
ML_PAGE = ml_tab.inject(
    S.PAGE.replace("<title>NSE Stock Browser</title>", "<title>ML Experiment — Nifty 50</title>")
          .replace("NSE Stock Browser", "ML Experiment · Nifty 50")
          .replace("search any of 2,363 stocks", f"search any of {len(S.SYMBOLS)} stocks"))


class MLHandler(S.Handler):
    """Same handler as the main dashboard, plus the ML page & /api/ml route."""
    def do_GET(self):
        path = urlparse(self.path).path
        if path in ("/", "/index.html"):
            self._send(200, ML_PAGE, "text/html; charset=utf-8")
            return
        if path == "/api/ml":
            sym = parse_qs(urlparse(self.path).query).get("sym", [""])[0]
            self._json(ml_tab.payload(sym))
            return
        super().do_GET()


def main():
    port = PORT
    for a in sys.argv[1:]:
        if a.isdigit():
            port = int(a)
    lan = "--lan" in sys.argv
    host = "0.0.0.0" if lan else "127.0.0.1"
    ip = S.lan_ip() if lan else "localhost"
    srv = S.ThreadingHTTPServer((host, port), MLHandler)
    print("=" * 62)
    print(f"  ML EXPERIMENT dashboard  —  {len(S.SYMBOLS)} stocks  (separate host)")
    print(f"  Open:  http://{ip}:{port}" + ("   (reachable on your LAN)" if lan else ""))
    print("  Same UI as the main dashboard + an ML Signal tab.")
    print("  Main dashboard on port 8090 is UNAFFECTED.")
    print("  Ctrl+C to stop.")
    print("=" * 62)
    if "--no-open" not in sys.argv:
        threading.Timer(0.6, lambda: webbrowser.open(f"http://localhost:{port}")).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        srv.shutdown()


if __name__ == "__main__":
    main()
