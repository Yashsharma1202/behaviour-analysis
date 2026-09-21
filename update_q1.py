"""
update_q1.py
===============================================================================
One command to refresh the Nifty-50 Q1-Result Behaviour tracker:

  1. reads the plan (entry/exit candles per stock) from  UPDATED.xlsx
  2. pulls the latest SETTLED daily closes from Yahoo (free, no API key)
  3. marks each position REALISED (exit candle reached) or OPEN (running % so far)
  4. writes  Nifty50_Q1Results_BehaviourEngine_UPDATED_06Aug_fixed.xlsx
  5. rebuilds the 2-page A4 PDF (via build_q1_report_pdf.py)

Meant to run just AFTER the market settles (~3:45-4:00 PM IST). It will NOT use
today's candle before 15:45 local time (it isn't final yet) — it falls back to
the last settled day, so the numbers are always reproducible.

    python update_q1.py

Windows Task Scheduler (runs every day at 3:45 PM): see the schtasks /Create
command printed by this script, or ask the assistant to register it for you.
===============================================================================
"""
from __future__ import annotations

import datetime as dt
import json
import re
import subprocess
import sys
import time
import urllib.request
from pathlib import Path

import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font

try:
    sys.stdout.reconfigure(encoding="utf-8")          # no-op if no console (Task Scheduler)
except Exception:                                     # noqa: BLE001
    pass

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "UPDATED.xlsx"
OUT = ROOT / "Nifty50_Q1Results_BehaviourEngine_UPDATED_06Aug_fixed.xlsx"
SETTLE_TIME = dt.time(15, 45)                     # today's candle trusted only after this

NOW = dt.datetime.now()
TODAY = pd.Timestamp(NOW.date())
INCLUDE_TODAY = NOW.time() >= SETTLE_TIME         # is today's bar final yet?


def yahoo(sym: str) -> pd.DataFrame | None:
    t = sym.replace("&", "%26") + ".NS"
    p1 = int(dt.datetime(2026, 1, 1).timestamp())  # Warm up period for EMA 50
    p2 = int((TODAY + pd.Timedelta(days=2)).timestamp())
    url = (f"https://query1.finance.yahoo.com/v8/finance/chart/{t}"
           f"?period1={p1}&period2={p2}&interval=1d")
    try:
        j = json.loads(urllib.request.urlopen(
            urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"}),
            timeout=25).read())
        res = j["chart"]["result"][0]
        idx = pd.to_datetime(res["timestamp"], unit="s").normalize()
        closes = res["indicators"]["quote"][0]["close"]
        lows = res["indicators"]["quote"][0]["low"]
        df = pd.DataFrame({"close": closes, "low": lows}, index=idx).dropna()
        if not INCLUDE_TODAY:                     # drop the still-live candle
            df = df[df.index < TODAY]
        return df
    except Exception as e:                        # noqa: BLE001
        print(f"  ! fetch {sym}: {e}")
        return None


def asof(df: pd.DataFrame, d, col="close") -> float | None:
    ss = df[df.index <= pd.Timestamp(d)]
    return round(float(ss[col].iloc[-1]), 2) if len(ss) else None


def ev(val, D, E, F):
    """Evaluate the sheet's date formulas the way Excel does (calendar days)."""
    if not isinstance(val, str) or not val.startswith("="):
        return val
    if D is None:
        return None
    if re.match(r"=D\d+-E\d+", val):
        return D - dt.timedelta(days=int(E))
    if re.match(r"=D\d+\+F\d+", val):
        return D + dt.timedelta(days=int(F))
    return None


def _fd(v):
    return v.strftime("%d-%b") if isinstance(v, dt.datetime) else "—"


def print_results(ws):
    """Print the full per-stock result table + summary to the terminal/log."""
    tty = sys.stdout.isatty()                          # colour only in a real terminal
    W, R, G, Y, DIM = (("\033[0m", "\033[31m", "\033[32m", "\033[33m", "\033[90m")
                       if tty else ("", "", "", "", ""))
    hdr = (f"{'#':>2}  {'SYMBOL':<11}{'RESULT':>7}{'ENTRY':>7}{'BUY':>10}"
           f"{'EXIT':>7}{'SELL':>10}{'RETURN':>11}  STATUS")
    print("\n" + "=" * len(hdr))
    print(hdr)
    print("-" * len(hdr))
    rr = []
    for r in range(3, ws.max_row + 1):
        D, E, F = ws.cell(r, 4).value, ws.cell(r, 5).value, ws.cell(r, 6).value
        sym = (ws.cell(r, 2).value or "").strip()
        en = ev(ws.cell(r, 11).value, D, E, F)
        ex = ev(ws.cell(r, 13).value, D, E, F)
        buy, sell, o = ws.cell(r, 12).value, ws.cell(r, 14).value, ws.cell(r, 15).value
        if isinstance(o, (int, float)):
            col = G if o >= 0 else R
            ret, status = f"{col}{o:+.2f}%{W}", f"{col}REALISED{W}"
            rr.append(o)
        elif isinstance(o, str) and "unrealised" in o:
            m = re.match(r"\s*([+-]?\d+(?:\.\d+)?)", o)
            v = float(m.group(1)) if m else 0.0
            col = G if v >= 0 else R
            ret, status = f"{col}{v:+.2f}%{W}", f"{Y}open{W}"
        elif isinstance(o, str) and "Filter Out" in o:
            ret, status = f"{DIM}Filter Out{W}", f"{DIM}Filter Out{W}"
        else:
            ret, status = f"{DIM}—{W}", f"{DIM}—{W}"
        b = f"{buy:,.2f}" if isinstance(buy, (int, float)) else "—"
        se = f"{sell:,.2f}" if isinstance(sell, (int, float)) else "—"
        print(f"{r-2:>2}  {sym:<11}{_fd(D):>7}{_fd(en):>7}{b:>10}{_fd(ex):>7}{se:>10}"
              f"{ret:>20}  {status}")
    if rr:
        wins = sum(1 for x in rr if x > 0)
        print("-" * len(hdr))
        print(f"  REALISED {len(rr)}: win-rate {wins}/{len(rr)} = {wins/len(rr)*100:.0f}%  |  "
              f"avg {sum(rr)/len(rr):+.2f}%  |  best {max(rr):+.2f}%  worst {min(rr):+.2f}%")
    print("=" * len(hdr))


def get_trading_date(s: pd.Series, result_date: dt.datetime, offset: int) -> dt.datetime | None:
    if s is None or len(s) == 0:
        return None
    D = pd.Timestamp(result_date).normalize()
    cal = s.index
    
    if D <= cal[-1]:
        pos = cal.searchsorted(D, side="right") - 1
        if pos < 0:
            pos = 0
        target_idx = pos + offset
        if 0 <= target_idx < len(cal):
            return cal[target_idx].to_pydatetime()
        elif target_idx < 0:
            return result_date + dt.timedelta(days=offset)
        else:
            rem_steps = target_idx - (len(cal) - 1)
            last_day = cal[-1]
            future_days = pd.bdate_range(start=last_day, periods=rem_steps + 1)
            return future_days[-1].to_pydatetime()
    else:
        b_days = pd.bdate_range(start=cal[-1], end=D)
        n_steps = len(b_days) - 1
        target_idx = n_steps + offset
        if target_idx >= 0:
            future_days = pd.bdate_range(start=cal[-1], periods=target_idx + 1)
            return future_days[-1].to_pydatetime()
        else:
            back_idx = len(cal) - 1 + target_idx
            if back_idx >= 0:
                return cal[back_idx].to_pydatetime()
            return D + dt.timedelta(days=offset)


def main():
    print(f"Update run {NOW:%Y-%m-%d %H:%M}  |  today's candle "
          f"{'INCLUDED' if INCLUDE_TODAY else 'excluded (before 15:45 — not settled)'}")
    wb = openpyxl.load_workbook(SRC)
    ws = wb.active

    prices = {}
    real, opn, fill = 0, 0, 0
    ctr = Alignment(horizontal="center")
    rgt = Alignment(horizontal="right")
    
    print("Pre-fetching prices and resolving trading day dates...")
    for r in range(3, ws.max_row + 1):
        sym = (ws.cell(r, 2).value or "").strip()
        if not sym:
            continue
            
        D, E, F = ws.cell(r, 4).value, ws.cell(r, 5).value, ws.cell(r, 6).value
        L = ws.cell(r, 12).value
        
        # Download prices
        if sym not in prices:
            prices[sym] = yahoo(sym)
            time.sleep(0.15)
            
        s = prices[sym]
        if s is None or not len(s):
            continue
            
        # Calculate entry and exit dates using trading day calendar
        if isinstance(D, dt.datetime):
            en = get_trading_date(s, D, -int(E))
            ex = get_trading_date(s, D, int(F))
            
            # Write static dates back to sheet
            ws.cell(r, 11).value = en
            ws.cell(r, 13).value = ex
        else:
            en, ex = None, None
            
        if en is None or ex is None:
            continue
            
        # Filter for active positions reaching into August whose entry has triggered
        if ex > dt.datetime(2026, 7, 31) and en <= NOW:
            last = s.index.max()
            Lc, Nc, Oc = ws.cell(r, 12), ws.cell(r, 14), ws.cell(r, 15)
            
            # Calculate EMA 50 Low Band
            s["ema50_low"] = s["low"].ewm(span=50, adjust=False).mean()
            buy_close = asof(s, en, "close")
            ema_low = asof(s, en, "ema50_low")
            
            if buy_close is not None and ema_low is not None and buy_close < ema_low:
                Lc.value = "-"; Lc.alignment = ctr
                Nc.value = "-"; Nc.alignment = ctr
                Oc.value = "Filter Out"; Oc.font = Font(italic=True, color="9F9F9F"); Oc.alignment = rgt
                continue
                
            buy = L if isinstance(L, (int, float)) else asof(s, en)
            if not isinstance(L, (int, float)):
                Lc.value = buy
                Lc.number_format = "0.00"
                fill += 1
            if buy is None:
                continue
                
            if ex <= last:                            # exit candle settled -> realise
                sell = asof(s, ex)
                ret = round((sell / buy - 1) * 100, 2)
                Nc.value = sell; Nc.number_format = "0.00"; Nc.font = Font()
                Oc.value = ret; Oc.number_format = "0.00"; Oc.font = Font(); Oc.alignment = rgt
                real += 1
            else:                                     # still open -> running %
                mark = asof(s, last)
                ret = round((mark / buy - 1) * 100, 2)
                Nc.value = "-"; Nc.alignment = ctr
                Oc.value = f"{ret:+.2f}% (unrealised)"; Oc.font = Font(italic=True); Oc.alignment = rgt
                opn += 1

    out = OUT
    try:
        wb.save(out)
    except PermissionError:
        out = OUT.with_name(OUT.stem + "_2.xlsx")
        wb.save(out)
        print(f"  (original open in Excel — saved to {out.name})")
        
    # Get last index max for display message
    last_dt_str = "N/A"
    if prices:
        first_valid_s = next((s for s in prices.values() if s is not None and len(s) > 0), None)
        if first_valid_s is not None:
            last_dt_str = str(first_valid_s.index.max().date())
            
    print(f"  last settled candle used: {last_dt_str}")
    print(f"  realised: {real} · open: {opn} · new buys filled: {fill}")
    print(f"  saved: {out.name}")

    print_results(ws)

    # rebuild the PDF from whatever file we just wrote
    try:
        subprocess.run([sys.executable, str(ROOT / "build_q1_report_pdf.py"), str(out)],
                       check=False)
    except Exception as e:                        # noqa: BLE001
        print(f"  ! PDF rebuild skipped: {e}")


if __name__ == "__main__":
    main()
