"""
refresh_price_cache.py
===============================================================================
processed/price_cache/<SYM>.csv and processed/raw_price_cache/<SYM>.csv were
last refreshed ~2026-07-10 (prefetch_prices.py only fetches a symbol if its
cache file does NOT already exist -- it never checks staleness, so a cache
that exists but is 3 months old is silently treated as current by every
downstream backtest).

This script re-fetches the FULL history from Yahoo (same source, same
adjustment methodology as the original caches -- not a different source
spliced on, which would risk a split/dividend discontinuity at the splice
point) for every symbol, and overwrites the cache file ONLY if the fresh
fetch succeeded and actually extends past what's currently on disk. A failed
fetch or a fetch that comes back no newer than the existing cache leaves the
file untouched -- this never deletes a cache file outright.

Usage:
    python refresh_price_cache.py              # all 211 symbols, both caches
    python refresh_price_cache.py --only-adj    # price_cache only (what the
                                                 # RBI/quarterly backtests read)
    python refresh_price_cache.py RELIANCE TCS
===============================================================================
"""
import argparse
import json
import os
import pathlib
import ssl
import sys
import time
import urllib.request
from urllib.parse import quote

import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')

ROOT = pathlib.Path(__file__).resolve().parent
OI_DIR = ROOT / 'OI_DATA'
ADJ_DIR = ROOT / 'processed' / 'price_cache'
RAW_DIR = ROOT / 'processed' / 'raw_price_cache'
ADJ_DIR.mkdir(parents=True, exist_ok=True)
RAW_DIR.mkdir(parents=True, exist_ok=True)

_HEADERS = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
_SSL = ssl.create_default_context()


def _fetch_json(symbol, period1):
    url = (f"https://query1.finance.yahoo.com/v8/finance/chart/{quote(symbol, safe='')}.NS"
           f"?period1={period1}&period2={int(time.time())}&interval=1d")
    req = urllib.request.Request(url, headers=_HEADERS)
    with urllib.request.urlopen(req, timeout=30, context=_SSL) as r:
        return json.load(r)


def fetch_adj(symbol):
    d = _fetch_json(symbol, 1167609600)
    res = d['chart']['result'][0]
    ind = res['indicators']
    adj = ind.get('adjclose', [{}])[0].get('adjclose') or ind['quote'][0]['close']
    s = pd.Series(adj, index=pd.to_datetime(res['timestamp'], unit='s').normalize())
    return s.dropna().sort_index().rename('adj').rename_axis('date')


def fetch_raw(symbol):
    d = _fetch_json(symbol, 1104537600)
    res = d['chart']['result'][0]
    q = res['indicators']['quote'][0]
    df = pd.DataFrame({
        'open': q.get('open'), 'high': q.get('high'),
        'low': q.get('low'), 'close': q.get('close'),
        'volume': q.get('volume'),
    }, index=pd.to_datetime(res['timestamp'], unit='s').normalize())
    df.index.name = 'date'
    return df.dropna(subset=['close']).sort_index()


def refresh_one(symbol, do_adj, do_raw):
    result = {'symbol': symbol, 'adj': 'skip', 'raw': 'skip'}

    if do_adj:
        adj_path = ADJ_DIR / f'{symbol}.csv'
        old_last = None
        if adj_path.exists():
            try:
                old_last = pd.read_csv(adj_path, parse_dates=['date'])['date'].max()
            except Exception:
                old_last = None
        try:
            fresh = fetch_adj(symbol)
            if len(fresh) and (old_last is None or fresh.index.max() > old_last):
                fresh.to_frame().to_csv(adj_path)
                result['adj'] = f"updated ({old_last.date() if old_last is not None else 'new'} -> {fresh.index.max().date()})"
            else:
                result['adj'] = 'already current'
        except Exception as e:
            result['adj'] = f'FAILED: {type(e).__name__}: {e}'
        time.sleep(0.25)

    if do_raw:
        raw_path = RAW_DIR / f'{symbol}.csv'
        old_last = None
        if raw_path.exists():
            try:
                old_last = pd.read_csv(raw_path, parse_dates=['date'])['date'].max()
            except Exception:
                old_last = None
        try:
            fresh = fetch_raw(symbol)
            if len(fresh) and (old_last is None or fresh.index.max() > old_last):
                fresh.to_csv(raw_path)
                result['raw'] = f"updated ({old_last.date() if old_last is not None else 'new'} -> {fresh.index.max().date()})"
            else:
                result['raw'] = 'already current'
        except Exception as e:
            result['raw'] = f'FAILED: {type(e).__name__}: {e}'
        time.sleep(0.25)

    return result


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('symbols', nargs='*')
    ap.add_argument('--only-adj', action='store_true')
    ap.add_argument('--only-raw', action='store_true')
    args = ap.parse_args()

    if args.symbols:
        syms = [s.upper() for s in args.symbols]
    else:
        syms = sorted(e.name.strip().upper() for e in os.scandir(OI_DIR) if e.is_dir())

    do_adj = not args.only_raw
    do_raw = not args.only_adj

    print(f"Refreshing {'adj' if do_adj and not do_raw else 'raw' if do_raw and not do_adj else 'adj+raw'} "
          f"price cache for {len(syms)} symbols from Yahoo (full re-fetch, overwrite only if it extends the data)...")

    updated_adj = updated_raw = failed = 0
    t0 = time.time()
    for i, sym in enumerate(syms, 1):
        r = refresh_one(sym, do_adj, do_raw)
        if 'updated' in r['adj']:
            updated_adj += 1
        if 'updated' in r['raw']:
            updated_raw += 1
        if 'FAILED' in r['adj'] or 'FAILED' in r['raw']:
            failed += 1
        if 'updated' in r['adj'] or 'updated' in r['raw'] or 'FAILED' in r['adj'] or 'FAILED' in r['raw']:
            print(f"  [{i}/{len(syms)}] {sym:<14} adj={r['adj']:<45} raw={r['raw']}")
        elif i % 50 == 0:
            print(f"  [{i}/{len(syms)}] ... ({time.time()-t0:.0f}s elapsed)")

    print(f"\nDone in {time.time()-t0:.0f}s. adj updated={updated_adj}, raw updated={updated_raw}, failed={failed}/{len(syms)}.")


if __name__ == '__main__':
    main()
