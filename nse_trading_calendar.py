"""
nse_trading_calendar.py
===============================================================================
Single source of truth for "is this a trading day" across the project.

No holiday date is hand-typed here. The calendar is fetched live from NSE's
own public holiday-master API (the same one nseindia.com itself uses):

    https://www.nseindia.com/api/holiday-master?type=trading

...and cached to processed/nse_trading_holidays.json. Re-running
refresh_trading_holidays() merges in whatever NSE currently publishes --
it never deletes a previously cached year, so the cache only grows (NSE
usually publishes next calendar year's list around December; until then,
that year just isn't in the live response yet, and this module says so
rather than guessing).

Usage:
    from nse_trading_calendar import load_trading_holidays, trading_day_window, add_trading_days

    holidays = load_trading_holidays()                        # refreshes from NSE, falls back to cache
    d = add_trading_days('2026-10-07', -2, holidays)           # -> Timestamp('2026-10-05')  (T-2)
    window = trading_day_window('2026-10-07', range(-8, 9), holidays)

Run directly (`python nse_trading_calendar.py`) to refresh the cache and
print it.
===============================================================================
"""
import http.cookiejar
import json
import pathlib
import urllib.request

import pandas as pd

ROOT = pathlib.Path(__file__).resolve().parent
PROC = ROOT / 'processed'
PROC.mkdir(parents=True, exist_ok=True)
HOLIDAY_CACHE = PROC / 'nse_trading_holidays.json'

_HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': '*/*',
    'Referer': 'https://www.nseindia.com/',
}


def fetch_trading_holidays_live(segment='CM', timeout=10):
    """Hit NSE's live holiday-master API. Returns {'YYYY-MM-DD': description}."""
    cj = http.cookiejar.CookieJar()
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))
    opener.open(urllib.request.Request(
        'https://www.nseindia.com/get-quotes/equity?symbol=RELIANCE', headers=_HEADERS
    ), timeout=timeout).read()

    req = urllib.request.Request(
        'https://www.nseindia.com/api/holiday-master?type=trading', headers=_HEADERS
    )
    data = json.loads(opener.open(req, timeout=timeout).read().decode('utf-8', errors='ignore'))
    rows = data.get(segment) or data.get('CM') or []

    out = {}
    for r in rows:
        d = pd.to_datetime(r.get('tradingDate'), format='%d-%b-%Y', errors='coerce')
        if pd.notna(d):
            out[d.strftime('%Y-%m-%d')] = r.get('description', '')
    return out


def refresh_trading_holidays(verbose=True):
    """Fetch the live NSE calendar and merge it additively into the cache
    (live data wins on overlap; a year NSE hasn't published yet just stays
    whatever was cached before -- never deleted, never guessed)."""
    cached = {}
    if HOLIDAY_CACHE.exists():
        try:
            cached = json.loads(HOLIDAY_CACHE.read_text(encoding='utf-8'))
        except Exception:
            cached = {}
    try:
        live = fetch_trading_holidays_live()
        cached.update(live)
        HOLIDAY_CACHE.write_text(json.dumps(dict(sorted(cached.items())), indent=2), encoding='utf-8')
        if verbose:
            print(f"[nse_trading_calendar] Refreshed from NSE live API: {len(live)} holidays "
                  f"({min(live) if live else '-'} .. {max(live) if live else '-'}); {len(cached)} total cached.")
    except Exception as e:
        if verbose:
            print(f"[nse_trading_calendar] WARNING: live NSE fetch failed ({e}); "
                  f"using existing cache only ({len(cached)} dates).")
    return cached


def load_trading_holidays(auto_refresh=True):
    """Load the holiday set. Refreshes from NSE first unless auto_refresh=False
    (e.g. for an offline/no-network run), in which case it reads the cache only."""
    if auto_refresh:
        return refresh_trading_holidays(verbose=False)
    if HOLIDAY_CACHE.exists():
        try:
            return json.loads(HOLIDAY_CACHE.read_text(encoding='utf-8'))
        except Exception:
            return {}
    return {}


def is_trading_day(date, holidays):
    d = pd.Timestamp(date).normalize()
    if d.weekday() >= 5:
        return False
    return d.strftime('%Y-%m-%d') not in holidays


def add_trading_days(base_date, n, holidays):
    """Step n trading sessions forward (n>0) or back (n<0) from base_date,
    skipping weekends and NSE holidays. n == 0 returns base_date unchanged."""
    d = pd.Timestamp(base_date).normalize()
    if n == 0:
        return d
    step = 1 if n > 0 else -1
    remaining = abs(n)
    while remaining > 0:
        d += pd.Timedelta(days=step)
        if is_trading_day(d, holidays):
            remaining -= 1
    return d


def trading_day_window(center_date, offsets, holidays):
    """Return {offset: pd.Timestamp} for every offset, computed from the real
    NSE calendar -- not a hand-typed table."""
    center = pd.Timestamp(center_date).normalize()
    return {off: add_trading_days(center, off, holidays) for off in offsets}


if __name__ == '__main__':
    hols = refresh_trading_holidays()
    print(f"\n{len(hols)} cached trading holidays:")
    for d, desc in sorted(hols.items()):
        print(f"  {d}  {desc}")
