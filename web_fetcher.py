"""
web_fetcher.py
===============================================================================
One place where outbound HTML fetching happens, with Scrapling used when it is
installed and a stdlib fallback when it is not.

WHY AN ADAPTER RATHER THAN CALLING SCRAPLING DIRECTLY
-----------------------------------------------------
Two reasons, both practical:

  * This project runs on pandas + torch + stdlib. Scrapling's stealth fetchers
    pull a patched Firefox (Camoufox) — hundreds of MB of browser binary. That
    is a fine trade for someone who needs it and a terrible default for someone
    who just wants the dashboard. So it is an OPTIONAL extra: import it if it
    is there, degrade to urllib if not, never hard-fail.
  * Scrapling's API has moved across versions (Fetcher / StealthyFetcher /
    PlayWrightFetcher, and `adaptive=` was `auto_match=` in older releases).
    Pinning every call site to one spelling means a version bump breaks the
    dashboard. Pinning ONE adapter means it breaks here, loudly, with a
    fallback still running.

WHAT SCRAPLING BUYS US
----------------------
  * NSE. stock_server._nse_prime hand-rolls a cookie-primed opener because NSE
    stalls plain requests. That works until it doesn't. StealthyFetcher solves
    the same problem with a real browser fingerprint.
  * Adaptive selectors. Scrapling can re-locate an element after a site
    restructures its markup, which is the failure mode that makes scrapers rot.

WHAT IT DOES NOT BUY US
-----------------------
  * Analyst consensus. Screener.in publishes HISTORICAL fundamentals, which
    this repo already has locally in quarterly/ and pnl/. No fetcher can scrape
    a number a site does not publish. Forward estimates need a source that
    actually carries them (Trendlyne forecasts, Moneycontrol) — see
    estimate_providers.py.

STATUS: the Scrapling code path below is written against the documented API but
has NOT been executed here — Scrapling is not installed in this environment and
no request was made to a live site. Treat `probe()` as the thing to run first.

    python web_fetcher.py            # report what backend is available
    python web_fetcher.py <url>      # fetch one URL and report which path ran
===============================================================================
"""

from __future__ import annotations

import ssl
import sys
import urllib.request
from dataclasses import dataclass

DEFAULT_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
              "(KHTML, like Gecko) Chrome/122.0 Safari/537.36")
TIMEOUT = 30

_BACKEND = {"checked": False, "scrapling": False, "detail": ""}


@dataclass
class Page:
    """A fetched page, whichever backend produced it."""
    url: str
    status: int
    html: str
    backend: str
    _obj: object = None                  # the Scrapling page, when there is one

    def css(self, selector: str, adaptive: bool = False) -> list:
        """Select elements. Falls back to a crude regex when Scrapling is absent
        — enough to keep a caller alive, not enough to be relied on."""
        if self._obj is not None:
            try:
                try:
                    return list(self._obj.css(selector, adaptive=adaptive))
                except TypeError:            # older Scrapling spelled it auto_match
                    return list(self._obj.css(selector, auto_match=adaptive))
            except Exception:                # noqa: BLE001
                return []
        return []

    def text_of(self, selector: str, adaptive: bool = False) -> str | None:
        els = self.css(selector, adaptive=adaptive)
        if not els:
            return None
        el = els[0]
        for attr in ("text", "get_all_text", "html_content"):
            v = getattr(el, attr, None)
            if v is None:
                continue
            return (v() if callable(v) else str(v)).strip()
        return str(el).strip()


def probe() -> dict:
    """What backend is actually available. Run this before trusting anything."""
    if _BACKEND["checked"]:
        return dict(_BACKEND)
    try:
        import scrapling                                    # noqa: F401
        _BACKEND["scrapling"] = True
        _BACKEND["detail"] = f"scrapling {getattr(scrapling, '__version__', '?')}"
    except Exception as e:                                  # noqa: BLE001
        _BACKEND["scrapling"] = False
        # ASCII only: this prints to a cp1252 console on Windows.
        _BACKEND["detail"] = f"scrapling unavailable ({type(e).__name__}) - using urllib"
    _BACKEND["checked"] = True
    return dict(_BACKEND)


def _urllib_get(url: str, headers: dict | None = None) -> Page:
    h = {"User-Agent": DEFAULT_UA, "Accept-Language": "en-US,en;q=0.9"}
    h.update(headers or {})
    ctx = ssl.create_default_context()
    req = urllib.request.Request(url, headers=h)
    with urllib.request.urlopen(req, timeout=TIMEOUT, context=ctx) as r:
        return Page(url=url, status=getattr(r, "status", 200),
                    html=r.read(2_000_000).decode("utf-8", "replace"),
                    backend="urllib")


def fetch(url: str, stealth: bool = False, headers: dict | None = None) -> Page | None:
    """Fetch one URL. Returns None rather than raising — a failed scrape must
    never take down the caller.

    stealth=True asks for a real browser fingerprint (Scrapling's
    StealthyFetcher). It is slow and needs the Camoufox binary, so use it only
    where a plain request is actually being blocked.
    """
    info = probe()
    if info["scrapling"]:
        try:
            if stealth:
                from scrapling.fetchers import StealthyFetcher
                r = StealthyFetcher.fetch(url, headless=True, network_idle=True)
            else:
                from scrapling.fetchers import Fetcher
                r = Fetcher.get(url, stealthy_headers=True)
            body = getattr(r, "html_content", None) or getattr(r, "body", "") or str(r)
            return Page(url=url, status=int(getattr(r, "status", 200)),
                        html=body if isinstance(body, str) else body.decode("utf-8", "replace"),
                        backend="scrapling-stealthy" if stealth else "scrapling",
                        _obj=r)
        except Exception as e:                              # noqa: BLE001
            print(f"  ! scrapling fetch failed ({type(e).__name__}: {e}) "
                  f"-> falling back to urllib", file=sys.stderr)
    try:
        return _urllib_get(url, headers)
    except Exception as e:                                  # noqa: BLE001
        print(f"  ! fetch failed for {url}: {type(e).__name__}: {e}", file=sys.stderr)
        return None


if __name__ == "__main__":
    info = probe()
    print(f"  backend: {info['detail']}")
    print(f"  scrapling available: {info['scrapling']}")
    if len(sys.argv) > 1:
        p = fetch(sys.argv[1])
        if p is None:
            print("  fetch returned None")
        else:
            print(f"  {p.status}  {len(p.html):,} bytes via {p.backend}")
    else:
        print("\n  pass a URL to test a live fetch, e.g.")
        print("    python web_fetcher.py https://www.screener.in/company/RELIANCE/")
        print("\n  install the optional backend with:")
        print("    pip install scrapling && scrapling install")
