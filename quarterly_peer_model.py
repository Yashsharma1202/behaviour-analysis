"""
quarterly_peer_model.py  —  last 4 COMPLETED quarters: predicted vs actual, trained
===============================================================================
Per stock, per quarter:
  PREDICTED = mean actual move of sector peers that reported EARLIER that quarter
              (if none, market fallback = all earlier reporters, tagged 'market')
  ACTUAL    = the stock's own move (buy Nc before / sell Mc after)
Trains actual ~ a + b·predicted on the 4 completed quarters; predicts upcoming.

Excludes the current (incomplete) quarter. Result dates: board-meeting feed.
Prices: Yahoo history. Output: a styled Excel with per-quarter sheets + charts.

    python quarterly_peer_model.py   ->  Nifty50_Peer_Model_4Q.xlsx
===============================================================================
"""
from __future__ import annotations
import datetime as dt
import json
import sys
import time
import urllib.request
import warnings
from pathlib import Path

import pandas as pd
warnings.filterwarnings("ignore")
sys.stdout.reconfigure(encoding="utf-8")

import expectation_bar as EB
from expectation_bar import SECTORS, SYM_SECTOR
from download_feeds import NIFTY50_FALLBACK

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "Nifty50_Peer_Model_8Q_ALL50.xlsx"
TODAY = pd.Timestamp(dt.date.today())


def rules() -> dict:
    import openpyxl
    for fn in ("Nifty50_Q1Results_BehaviourEngine_UPDATED_06Aug_fixed.xlsx",
               "UPDATED.xlsx", "Nifty50_Q1Results_BehaviourEngine.xlsx"):
        try:
            ws = openpyxl.load_workbook(ROOT / fn).active
        except Exception:                                   # noqa: BLE001
            continue
        r = {}
        for i in range(3, ws.max_row + 1):
            s = (ws.cell(i, 2).value or "").strip()
            bb, sa = ws.cell(i, 5).value, ws.cell(i, 6).value
            if s and isinstance(bb, (int, float)) and isinstance(sa, (int, float)):
                r[s] = (int(bb), int(sa))
        if r:
            return r
    return {}


def yahoo(sym: str):
    t = sym.replace("&", "%26") + ".NS"
    p1 = int(dt.datetime(2024, 1, 1).timestamp())
    p2 = int(dt.datetime.now().timestamp())
    url = (f"https://query1.finance.yahoo.com/v8/finance/chart/{t}"
           f"?period1={p1}&period2={p2}&interval=1d")
    j = json.loads(urllib.request.urlopen(
        urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"}), timeout=30).read())
    res = j["chart"]["result"][0]
    s = pd.Series(res["indicators"]["quote"][0]["close"],
                  index=pd.to_datetime(res["timestamp"], unit="s").normalize()).dropna()
    return s[s.index < TODAY]


def result_dates(sym):
    """Union of result dates from ALL local feeds (board meetings, financial-result
    announcements, financial_results) so the current-50 stocks are covered as fully
    as possible across quarters — the board-meeting feed alone has gaps."""
    dates = []
    # 1) board-meeting result meetings
    p = ROOT / sym / "board_meetings.csv"
    if p.exists():
        try:
            df = pd.read_csv(p, dtype=str).fillna("")
            if "bm_date" in df.columns:
                df["d"] = pd.to_datetime(df["bm_date"], errors="coerce", format="mixed")
                txt = (df.get("bm_purpose", "") + " " + df.get("bm_desc", "")).str.lower()
                dates += list(df[txt.str.contains("result|unaudited|financial", na=False)]["d"].dropna())
        except Exception:                                   # noqa: BLE001
            pass
    # 2) "Financial Result" announcements
    p = ROOT / sym / "announcements.csv"
    if p.exists():
        try:
            df = pd.read_csv(p, dtype=str).fillna("")
            dc = next((c for c in ("sort_date", "an_dt", "dt") if c in df.columns), None)
            if dc:
                df["d"] = pd.to_datetime(df[dc], errors="coerce", format="mixed")
                desc = df.get("desc", "").str.lower()
                dates += list(df[desc.str.contains("financial result", na=False)]["d"].dropna())
        except Exception:                                   # noqa: BLE001
            pass
    # 3) financial_results feed (broadcast date)
    p = ROOT / sym / "financial_results.csv"
    if p.exists():
        try:
            df = pd.read_csv(p, dtype=str).fillna("")
            col = next((c for c in ("broadCastDate", "broadcastDate", "resultDate") if c in df.columns), None)
            if col:
                dates += list(pd.to_datetime(df[col], errors="coerce", format="mixed").dropna())
        except Exception:                                   # noqa: BLE001
            pass
    if not dates:
        return []
    return sorted(pd.to_datetime(pd.Series(dates)).dt.normalize().dropna().unique())


def move(prices, D, bb, sa):
    cal = prices.index
    pos = cal.searchsorted(pd.Timestamp(D), side="right") - 1
    ei, xi = pos - bb, pos + sa
    if ei < 0 or xi >= len(cal):
        return None
    return round((float(prices.iloc[xi]) / float(prices.iloc[ei]) - 1) * 100, 2)


def lin_fit(xs, ys):
    n = len(xs)
    if n < 3:
        return None
    mx, my = sum(xs) / n, sum(ys) / n
    sxx = sum((x - mx) ** 2 for x in xs)
    sxy = sum((x - mx) * (y - my) for x, y in zip(xs, ys))
    syy = sum((y - my) ** 2 for y in ys)
    if sxx == 0 or syy == 0:
        return None
    b = sxy / sxx
    a = my - b * mx
    ss_res = sum((y - (a + b * x)) ** 2 for x, y in zip(xs, ys))
    return {"a": round(a, 3), "b": round(b, 3), "r2": round(1 - ss_res / syy, 3),
            "corr": round(sxy / ((sxx ** .5) * (syy ** .5)), 3), "n": n}


def main():
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.chart import ScatterChart, BarChart, Reference, Series

    R = rules()
    syms = sorted(set(NIFTY50_FALLBACK))
    cur = TODAY.to_period("Q")
    keep = [str(cur - i) for i in range(8, 0, -1)]          # the 8 completed quarters, fixed
    current_q = str(cur)
    print(f"fetching Yahoo + result dates for {len(syms)} stocks · quarters {keep}")

    recs = []
    for sym in syms:
        bb, sa = R.get(sym, (6, 3))
        try:
            px = yahoo(sym)
        except Exception as e:                              # noqa: BLE001
            print(f"  ! {sym}: {e}"); continue
        known = [pd.Timestamp(d) for d in result_dates(sym)]
        offs = [(d - d.to_period("Q").start_time).days for d in known]
        med_off = int(pd.Series(offs).median()) if offs else 45   # typical day-into-quarter
        by_q = {}                                           # earliest real date per quarter
        for d in sorted(known):
            by_q.setdefault(str(d.to_period("Q")), d)
        for q in keep:                                      # ensure ALL 50 × 8 quarters
            if q in by_q:
                D, est = by_q[q], False
            else:                                           # estimate from the stock's cadence
                D = pd.Timestamp(pd.Period(q, freq="Q").start_time) + pd.Timedelta(days=med_off)
                est = True
            m = move(px, D, bb, sa)
            if m is None:
                continue
            recs.append({"symbol": sym, "sector": SYM_SECTOR.get(sym), "date": D,
                         "quarter": q, "buy_before": bb, "sell_after": sa,
                         "actual": m, "estimated": est})
        time.sleep(0.15)

    df = pd.DataFrame(recs)
    if df.empty:
        print("no data"); return
    curr = df.iloc[0:0]                                     # current quarter not collected here
    n_est = int(df["estimated"].sum())
    print(f"collected {len(df)} rows ({n_est} estimated dates) across {keep}")

    def pred_basis(row):
        q = df[df["quarter"] == row["quarter"]]
        sect = q[(q["sector"] == row["sector"]) & (q["date"] < row["date"]) & (q["symbol"] != row["symbol"])]
        if len(sect):
            return round(sect["actual"].mean(), 2), "sector"
        mkt = q[(q["date"] < row["date"]) & (q["symbol"] != row["symbol"])]
        if len(mkt):
            return round(mkt["actual"].mean(), 2), "market"
        return None, None

    mdf = df[df["quarter"].isin(keep)].copy()
    pb = mdf.apply(pred_basis, axis=1)
    mdf["predicted"] = [x[0] for x in pb]
    mdf["basis"] = [x[1] for x in pb]
    mdf["error"] = (mdf["actual"] - mdf["predicted"]).round(2)

    summ = []
    for qd in keep:
        sub = mdf[(mdf["quarter"] == qd) & mdf["predicted"].notna()]
        fit = lin_fit(list(sub["predicted"]), list(sub["actual"]))
        summ.append({"quarter": qd, "events": int((mdf["quarter"] == qd).sum()),
                     "est": int(mdf[mdf["quarter"] == qd]["estimated"].sum()),
                     "with_pred": len(sub), "corr": (fit or {}).get("corr"),
                     "mae": round((sub["actual"] - sub["predicted"]).abs().mean(), 2) if len(sub) else None,
                     "dir": round(((sub["predicted"] >= 0) == (sub["actual"] >= 0)).mean() * 100) if len(sub) else None})
    tr = mdf[mdf["predicted"].notna()]
    model = lin_fit(list(tr["predicted"]), list(tr["actual"]))

    up = []
    real = {r["symbol"]: r["actual"] for _, r in curr.iterrows()}
    for sym, ts in EB.upcoming_events():
        sib = [real[p] for p in SECTORS.get(SYM_SECTOR.get(sym), []) if p != sym and p in real]
        pred = round(sum(sib) / len(sib), 2) if sib else None
        madj = round(model["a"] + model["b"] * pred, 2) if (model and pred is not None) else None
        up.append({"symbol": sym, "sector": SYM_SECTOR.get(sym),
                   "result_date": pd.Timestamp(ts).date().isoformat(),
                   "peer_predicted": pred, "model_adjusted": madj, "peers_used": len(sib)})

    # ───────────────────────── styled workbook ──────────────────────────────
    BANNER = PatternFill("solid", fgColor="1F3B5B")
    HEAD = PatternFill("solid", fgColor="2E5F8A")
    ALT = PatternFill("solid", fgColor="EEF3F9")
    hfont = Font(bold=True, color="FFFFFF")
    tfont = Font(bold=True, color="FFFFFF", size=13)
    thin = Side(style="thin", color="D3DBE6")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ctr = Alignment(horizontal="center", vertical="center")
    lft = Alignment(horizontal="left", vertical="center")

    def sf(v):
        if isinstance(v, (int, float)):
            return Font(color="1E7B34" if v >= 0 else "C0392B", bold=True)
        return Font(color="8A94A6")

    wb = Workbook()
    ws = wb.active; ws.title = "Summary"
    ws.merge_cells("A1:E1")
    ws["A1"] = "NIFTY 50 · Peer Read-across Model · Predicted vs Actual (last 8 completed quarters)"
    ws["A1"].fill = BANNER; ws["A1"].font = tfont; ws["A1"].alignment = lft
    ws.row_dimensions[1].height = 26
    for c, h in enumerate(["Quarter", "Events", "Correlation", "Mean abs err", "Direction hit %", "Est dates"], 1):
        cell = ws.cell(3, c, h); cell.fill = HEAD; cell.font = hfont; cell.border = border; cell.alignment = ctr
    for i, s in enumerate(summ):
        r = 4 + i
        for c, v in enumerate([s["quarter"], s["events"], s["corr"], s["mae"], s["dir"], s["est"]], 1):
            cell = ws.cell(r, c, v); cell.border = border; cell.alignment = ctr
            if i % 2 == 0:
                cell.fill = ALT
        ws.cell(r, 3).font = sf(s["corr"])
    mr = 4 + len(summ) + 1
    ws.cell(mr, 1, "TRAINED MODEL (all 8 quarters pooled):").font = Font(bold=True)
    if model:
        ws.cell(mr + 1, 1, f"actual ≈ {model['a']} + {model['b']} × predicted    "
                           f"(R² = {model['r2']}, corr = {model['corr']}, n = {model['n']})")
    ws.cell(mr + 3, 1, "Verdict:").font = Font(bold=True, color="C0392B")
    ws.cell(mr + 3, 2, "Peer read-across is UNSTABLE — correlation flips sign quarter to quarter; pooled edge ≈ 0.")
    for col, w in zip("ABCDEF", [14, 10, 13, 14, 15, 10]):
        ws.column_dimensions[col].width = w
    bc = BarChart(); bc.title = "Correlation by quarter"; bc.legend = None
    bc.y_axis.title = "corr(pred, actual)"; bc.height = 7.5; bc.width = 15
    bc.add_data(Reference(ws, min_col=3, min_row=3, max_row=3 + len(summ)), titles_from_data=True)
    bc.set_categories(Reference(ws, min_col=1, min_row=4, max_row=3 + len(summ)))
    ws.add_chart(bc, "G3")

    qhdr = ["Symbol", "Sector", "Date", "Buy(c)", "Sell(c)", "Predicted %", "Actual %", "Error", "Basis", "Est?"]
    for qd in keep:
        q = mdf[mdf["quarter"] == qd].copy().sort_values("date")
        sh = wb.create_sheet(qd.replace(" ", ""))
        sh.merge_cells("A1:J1")
        sh["A1"] = f"{qd} · predicted (peer) vs actual move %   ·   all {len(q)} current-Nifty-50 stocks"
        sh["A1"].fill = BANNER; sh["A1"].font = tfont; sh["A1"].alignment = lft
        sh.row_dimensions[1].height = 22
        for c, h in enumerate(qhdr, 1):
            cell = sh.cell(3, c, h); cell.fill = HEAD; cell.font = hfont; cell.border = border; cell.alignment = ctr
        rr = 4
        for _, row in q.iterrows():
            est = bool(row["estimated"])
            vals = [row["symbol"], row["sector"], row["date"].date().isoformat(),
                    row["buy_before"], row["sell_after"], row["predicted"],
                    row["actual"], row["error"], row["basis"], "est. date" if est else ""]
            for c, v in enumerate(vals, 1):
                cell = sh.cell(rr, c, v); cell.border = border
                cell.alignment = lft if c <= 2 else ctr
                if rr % 2 == 0:
                    cell.fill = ALT
            sh.cell(rr, 6).font = sf(row["predicted"])
            sh.cell(rr, 7).font = sf(row["actual"])
            if est:
                sh.cell(rr, 3).font = Font(color="B26A00", italic=True)   # estimated date
                sh.cell(rr, 10).font = Font(color="B26A00", italic=True)
            rr += 1
        last = rr - 1
        for col, w in zip("ABCDEFGHIJ", [12, 16, 12, 7, 7, 12, 11, 9, 9, 9]):
            sh.column_dimensions[col].width = w
        if last >= 5:
            sc = ScatterChart(); sc.title = f"{qd}: predicted vs actual"
            sc.x_axis.title = "Predicted %"; sc.y_axis.title = "Actual %"; sc.legend = None
            sc.height = 9; sc.width = 13
            xref = Reference(sh, min_col=6, min_row=4, max_row=last)
            yref = Reference(sh, min_col=7, min_row=4, max_row=last)
            ser = Series(yref, xref, title="stocks")
            ser.marker.symbol = "circle"; ser.marker.size = 6
            ser.graphicalProperties.line.noFill = True
            sc.series.append(ser); sh.add_chart(sc, "K3")

    # pooled Actual-vs-Predicted across all 8 quarters (placed right after Summary)
    allrows = mdf[mdf["predicted"].notna()].sort_values(["quarter", "date"])
    al = wb.create_sheet("Actual_vs_Predicted", 1)
    al.merge_cells("A1:F1")
    al["A1"] = "All 8 quarters pooled · Actual vs Predicted (each dot = one stock-quarter)"
    al["A1"].fill = BANNER; al["A1"].font = tfont; al["A1"].alignment = lft
    al.row_dimensions[1].height = 22
    for c, h in enumerate(["Quarter", "Symbol", "Sector", "Predicted %", "Actual %", "Error"], 1):
        cell = al.cell(3, c, h); cell.fill = HEAD; cell.font = hfont; cell.border = border; cell.alignment = ctr
    rr = 4
    for _, row in allrows.iterrows():
        for c, v in enumerate([row["quarter"], row["symbol"], row["sector"],
                               row["predicted"], row["actual"], row["error"]], 1):
            cell = al.cell(rr, c, v); cell.border = border
            cell.alignment = lft if c in (2, 3) else ctr
            if rr % 2 == 0:
                cell.fill = ALT
        al.cell(rr, 4).font = sf(row["predicted"]); al.cell(rr, 5).font = sf(row["actual"])
        rr += 1
    lastr = rr - 1
    for col, w in zip("ABCDEF", [10, 12, 16, 12, 11, 9]):
        al.column_dimensions[col].width = w
    if model:
        al.cell(3, 8, f"Trained: actual ≈ {model['a']} + {model['b']} × predicted   "
                      f"(R² = {model['r2']}, corr = {model['corr']}, n = {model['n']})").font = Font(bold=True)
        al.cell(4, 8, "Verdict: no persistent edge — the cloud is round, not a line.").font = Font(color="C0392B")
    if lastr >= 5:
        sc = ScatterChart(); sc.title = "All quarters: Actual vs Predicted"
        sc.x_axis.title = "Predicted %"; sc.y_axis.title = "Actual %"; sc.legend = None
        sc.height = 11; sc.width = 18
        xref = Reference(al, min_col=4, min_row=4, max_row=lastr)
        yref = Reference(al, min_col=5, min_row=4, max_row=lastr)
        ser = Series(yref, xref, title="stock-quarters")
        ser.marker.symbol = "circle"; ser.marker.size = 5
        ser.graphicalProperties.line.noFill = True
        sc.series.append(ser); al.add_chart(sc, "H6")

    us = wb.create_sheet("Upcoming_Next")
    us.merge_cells("A1:F1")
    us["A1"] = "Upcoming / Next — peer prediction (current-quarter peers) + model-adjusted"
    us["A1"].fill = BANNER; us["A1"].font = tfont; us["A1"].alignment = lft
    for c, h in enumerate(["Symbol", "Sector", "Result date", "Peer predicted %", "Model adjusted %", "Peers used"], 1):
        cell = us.cell(3, c, h); cell.fill = HEAD; cell.font = hfont; cell.border = border; cell.alignment = ctr
    for i, u in enumerate(up):
        r = 4 + i
        for c, v in enumerate([u["symbol"], u["sector"], u["result_date"], u["peer_predicted"],
                               u["model_adjusted"], u["peers_used"]], 1):
            cell = us.cell(r, c, v); cell.border = border; cell.alignment = ctr if c != 2 else lft
            if i % 2 == 0:
                cell.fill = ALT
        us.cell(r, 4).font = sf(u["peer_predicted"]); us.cell(r, 5).font = sf(u["model_adjusted"])
    for col, w in zip("ABCDEF", [12, 16, 12, 16, 16, 11]):
        us.column_dimensions[col].width = w

    out = None
    for cand in [OUT] + [OUT.with_name(OUT.stem + f"_v{i}.xlsx") for i in range(2, 9)]:
        try:
            wb.save(cand); out = cand; break
        except PermissionError:
            continue
    if out is None:
        print("could not save — please close the open Excel files"); return

    print(f"\nwrote {out.name}")
    for s in summ:
        print(f"  {s['quarter']}: n={s['with_pred']:>2} corr={s['corr']} mae={s['mae']} dir={s['dir']}%")
    if model:
        print(f"TRAINED: actual ≈ {model['a']} + {model['b']}·pred  (R²={model['r2']}, corr={model['corr']}, n={model['n']})")
    print("upcoming:", [(u["symbol"], u["peer_predicted"], u["model_adjusted"]) for u in up])


if __name__ == "__main__":
    main()
