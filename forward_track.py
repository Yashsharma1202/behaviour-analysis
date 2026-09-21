"""
forward_track.py  —  lock in predictions BEFORE events, measure the move AFTER
===============================================================================
The only legitimate way to prove the peer/expected-move edge: snapshot each
upcoming event's prediction now (with a timestamp), then after it reports,
compute the ACTUAL move and compare. This accumulates a clean forward sample.

    python forward_track.py            # capture new upcoming + settle matured
    python forward_track.py capture    # only snapshot upcoming predictions
    python forward_track.py settle     # only fill actuals for matured events

Store: forward.db (SQLite). One row per (symbol, decision) — the FIRST snapshot
is kept (that's the honest pre-event prediction).
===============================================================================
"""
from __future__ import annotations

import datetime as dt
import json
import sqlite3
import sys
import time
import urllib.request
from pathlib import Path

import pandas as pd
sys.stdout.reconfigure(encoding="utf-8")

import expectation_bar as EB
import stock_server as S

ROOT = Path(__file__).resolve().parent
DB = ROOT / "forward.db"
TRACKER = ROOT / "Nifty50_Q1Results_BehaviourEngine_UPDATED_06Aug_fixed.xlsx"


def _conn():
    c = sqlite3.connect(DB)
    c.execute("""CREATE TABLE IF NOT EXISTS predictions(
        symbol TEXT, decision_ts TEXT, snapshot TEXT,
        subj_pred REAL, peer_pred REAL, bar REAL,
        buy_before INT, sell_after INT,
        entry_date TEXT, exit_date TEXT,
        actual REAL, status TEXT, settled TEXT,
        UNIQUE(symbol, decision_ts))""")
    return c


def _realised():
    import openpyxl
    out = {}
    try:
        ws = openpyxl.load_workbook(TRACKER).active
        for r in range(3, ws.max_row + 1):
            sym = (ws.cell(r, 2).value or "").strip()
            o = ws.cell(r, 15).value
            if sym and isinstance(o, (int, float)):
                out[sym] = float(o)
    except Exception:                                       # noqa: BLE001
        pass
    return out


def _subject(sym):
    """(subj_pred %, buy_before, sell_after) from the behaviour engine."""
    try:
        b = S.build_behaviour(sym)
        types = b.get("types", {}) if b.get("available") else {}
        et = ("RESULTS" if "RESULTS" in types else
              "BOARD_MEETING" if "BOARD_MEETING" in types else
              (next(iter(types)) if types else None))
        if not et:
            return None, None, None
        best = types[et]["best"]
        return best.get("avg_return_pct"), best.get("days_before"), best.get("days_after")
    except Exception:                                       # noqa: BLE001
        return None, None, None


def _peer_pred(sym, real):
    peers = [real[p] for p in EB.SECTORS.get(EB.SYM_SECTOR.get(sym), [])
             if p != sym and p in real]
    return round(sum(peers) / len(peers), 2) if peers else None


def capture():
    c = _conn()
    real = _realised()
    now = dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds")
    n = 0
    for sym, ts in EB.upcoming_events():
        dts = ts.isoformat()
        if c.execute("SELECT 1 FROM predictions WHERE symbol=? AND decision_ts=?",
                     (sym, dts)).fetchone():
            continue                                        # keep the first snapshot
        subj, bb, sa = _subject(sym)
        peer = _peer_pred(sym, real)
        try:
            bar = EB.score_batch(sym, ts).get("expectation_bar_raw")
        except Exception:                                   # noqa: BLE001
            bar = None
        c.execute("INSERT OR IGNORE INTO predictions "
                  "(symbol,decision_ts,snapshot,subj_pred,peer_pred,bar,buy_before,sell_after,status) "
                  "VALUES(?,?,?,?,?,?,?,?,?)",
                  (sym, dts, now, subj, peer, bar, bb, sa, "pending"))
        n += c.execute("SELECT changes()").fetchone()[0]
    c.commit(); c.close()
    print(f"capture: {n} new prediction snapshot(s) at {now}")
    return n


def _yahoo(sym, around: dt.datetime):
    t = sym.replace("&", "%26") + ".NS"
    p1 = int((around - dt.timedelta(days=40)).timestamp())
    p2 = int((around + dt.timedelta(days=40)).timestamp())
    url = (f"https://query1.finance.yahoo.com/v8/finance/chart/{t}"
           f"?period1={p1}&period2={p2}&interval=1d")
    j = json.loads(urllib.request.urlopen(
        urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"}), timeout=25).read())
    res = j["chart"]["result"][0]
    s = pd.Series(res["indicators"]["quote"][0]["close"],
                  index=pd.to_datetime(res["timestamp"], unit="s").normalize()).dropna()
    return s[s.index < pd.Timestamp(dt.date.today())]       # settled only


def settle():
    c = _conn()
    rows = c.execute("SELECT symbol,decision_ts,buy_before,sell_after,status FROM predictions "
                     "WHERE status!='realised'").fetchall()
    upd = 0
    for sym, dts, bb, sa, _st in rows:
        if bb is None or sa is None:
            continue
        D = pd.Timestamp(dt.datetime.fromisoformat(dts).date())
        try:
            s = _yahoo(sym, D.to_pydatetime())
        except Exception:                                   # noqa: BLE001
            continue
        if not len(s):
            continue
        cal = s.index
        pos = cal.searchsorted(D, side="right") - 1
        ei, xi = pos - int(bb), pos + int(sa)
        if ei < 0:
            continue                                        # not entered yet
        entry_dt = cal[ei]; buy = float(s.iloc[ei])
        last = cal[-1]
        if xi < len(cal):                                   # exit settled -> realised
            exit_dt = cal[xi]; sell = float(s.iloc[xi])
            actual = round((sell / buy - 1) * 100, 2); status = "realised"
        else:                                               # entered, running
            exit_dt = None; actual = round((float(s.iloc[-1]) / buy - 1) * 100, 2); status = "open"
        c.execute("UPDATE predictions SET entry_date=?,exit_date=?,actual=?,status=?,settled=? "
                  "WHERE symbol=? AND decision_ts=?",
                  (entry_dt.date().isoformat(), exit_dt.date().isoformat() if exit_dt is not None else None,
                   actual, status, last.date().isoformat(), sym, dts))
        upd += 1
        time.sleep(0.2)
    c.commit(); c.close()
    print(f"settle: updated {upd} row(s)")
    return upd


def rows():
    c = _conn()
    cols = ["symbol", "decision_ts", "snapshot", "subj_pred", "peer_pred", "bar",
            "buy_before", "sell_after", "entry_date", "exit_date", "actual", "status", "settled"]
    out = [dict(zip(cols, r)) for r in
           c.execute(f"SELECT {','.join(cols)} FROM predictions ORDER BY decision_ts").fetchall()]
    c.close()
    return out


def main():
    arg = sys.argv[1] if len(sys.argv) > 1 else "both"
    if arg in ("capture", "both"):
        capture()
    if arg in ("settle", "both"):
        settle()
    for r in rows():
        print(f"  {r['symbol']:11} decision {r['decision_ts'][:10]} | subj {r['subj_pred']} "
              f"peer {r['peer_pred']} bar {r['bar']} | actual {r['actual']} [{r['status']}]")


if __name__ == "__main__":
    main()
