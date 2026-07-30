"""
build_stock_profiles.py
===============================================================================
Generate per-stock profiles for every Nifty 50 stock — key fundamentals + the
best "buy-before / sell-after" rule & win-rate for each event type — GROUPED BY
SECTOR. Produces two things:

  1. Nifty50_Stock_Profiles.html   — standalone, light, sector-grouped
  2. injects a dark, sector-grouped APPENDIX into Nifty50_Presentation.html
     (at the <!--APPENDIX_CARDS--> marker)

    python build_stock_profiles.py
Reads live data from the running dashboard (localhost:8090).
===============================================================================
"""
from __future__ import annotations

import json
import re
import sys
import urllib.request
import warnings
from pathlib import Path

import pandas as pd

warnings.filterwarnings("ignore")
sys.stdout.reconfigure(encoding="utf-8")
ROOT = Path(__file__).resolve().parent
BASE = "http://127.0.0.1:8090"

EVENTS = [("RESULTS", "Results"), ("BOARD_MEETING", "Board mtg"),
          ("CORPORATE_ACTION", "Corp action"), ("ANNOUNCEMENT", "Announce")]

# Nifty 50 grouped into readable sectors (covers all 50)
SECTORS = [
    ("Banks", ["HDFCBANK", "ICICIBANK", "AXISBANK", "KOTAKBANK", "SBIN"]),
    ("Financial Services", ["BAJFINANCE", "BAJAJFINSV", "SHRIRAMFIN", "JIOFIN", "HDFCLIFE", "SBILIFE"]),
    ("IT & Technology", ["TCS", "INFY", "HCLTECH", "TECHM", "WIPRO"]),
    ("Automobiles", ["MARUTI", "M&M", "BAJAJ-AUTO", "EICHERMOT", "TMPV"]),
    ("Metals & Mining", ["TATASTEEL", "JSWSTEEL", "HINDALCO", "COALINDIA"]),
    ("Oil, Gas & Power", ["RELIANCE", "ONGC", "NTPC", "POWERGRID"]),
    ("Pharma & Healthcare", ["SUNPHARMA", "CIPLA", "DRREDDY", "APOLLOHOSP", "MAXHEALTH"]),
    ("FMCG", ["HINDUNILVR", "ITC", "NESTLEIND", "TATACONSUM"]),
    ("Consumer & Retail", ["TITAN", "TRENT", "ASIANPAINT", "ETERNAL"]),
    ("Cement & Materials", ["ULTRACEMCO", "GRASIM"]),
    ("Infra & Capital Goods", ["LT", "BEL"]),
    ("Diversified & Ports", ["ADANIENT", "ADANIPORTS"]),
    ("Telecom", ["BHARTIARTL"]),
    ("Aviation", ["INDIGO"]),
]


def get(url):
    return json.load(urllib.request.urlopen(url, timeout=90))


def inr(v):
    if v is None:
        return "—"
    n = int(round(v)); s = str(abs(n))
    if len(s) > 3:
        s = re.sub(r"(\d)(?=(\d\d)+$)", r"\1,", s[:-3]) + "," + s[-3:]
    return ("-" if n < 0 else "") + s


def wclass(w, dark=False):
    if w is None:
        return "na2" if dark else "na"
    base = "g" if w >= 60 else ("a" if w >= 50 else "r")
    return base + "2" if dark else base


def best_event(sym, best):
    bw, bet = -1, None
    for et, _ in EVENTS:
        r = best.get(et, {}).get(sym)
        if r and r.get("win") is not None and r["win"] > bw:
            bw, bet = r["win"], et
    return bet


def light_card(sym, s, best, name):
    star_et = best_event(sym, best)
    price = f"₹{inr(s.get('price'))}" if s.get("price") is not None else "—"
    pe = f"P/E {s['pe']:.1f}" if s.get("pe") is not None else "P/E —"
    fund = [("Sales", f"₹{inr(s.get('sales'))} cr"), ("Net profit", f"₹{inr(s.get('npat'))} cr"),
            ("ROCE", "—" if s.get("roce") is None else f"{s['roce']:.0f}%"),
            ("OPM", "—" if s.get("opm") is None else f"{s['opm']:.0f}%"),
            ("Debt/Eq", "—" if s.get("de") is None else f"{s['de']:.2f}"),
            ("Profit gr", "—" if s.get("profit_g") is None else f"{s['profit_g']:+.0f}%")]
    fh = "".join(f'<div class="fc"><span class="fl">{l}</span>'
                 f'<span class="fv {"neg" if str(v).startswith("-") else ""}">{v}</span></div>' for l, v in fund)
    rows = ""
    for et, lab in EVENTS:
        r = best.get(et, {}).get(sym)
        if r:
            rows += (f'<div class="er"><span class="el">{lab}{" ★" if et==star_et else ""}</span>'
                     f'<span class="erule">buy {r["db"]}d / sell {r["da"]}d</span>'
                     f'<span class="ew {wclass(r["win"])}">{r["win"]:.0f}%</span>'
                     f'<span class="en">n={r.get("events","")}</span></div>')
        else:
            rows += ('<div class="er"><span class="el">' + lab + '</span>'
                     '<span class="erule">—</span><span class="ew na">—</span><span class="en"></span></div>')
    return (f'<div class="pc"><div class="pch"><div class="sym">{sym}</div>'
            f'<div class="px">{price} · {pe}</div></div><div class="nm">{name}</div>'
            f'<div class="fund">{fh}</div>'
            f'<div class="evhd">Best event plays &nbsp;<span class="hint">win rate · ★ = strongest</span></div>'
            f'<div class="ev">{rows}</div></div>')


def dark_card(sym, s, best, name):
    star_et = best_event(sym, best)
    price = f"₹{inr(s.get('price'))}" if s.get("price") is not None else "—"
    pe = f"P/E {s['pe']:.1f}" if s.get("pe") is not None else "P/E —"
    fr = (f'<span>ROCE <b>{"—" if s.get("roce") is None else str(int(s["roce"]))+"%"}</b></span>'
          f'<span>OPM <b>{"—" if s.get("opm") is None else str(int(s["opm"]))+"%"}</b></span>'
          f'<span>D/E <b>{"—" if s.get("de") is None else f"{s["de"]:.2f}"}</b></span>'
          f'<span>Pft <b class="{"r2" if (s.get("profit_g") or 0)<0 else ""}">'
          f'{"—" if s.get("profit_g") is None else f"{s["profit_g"]:+.0f}%"}</b></span>')
    rows = ""
    for et, lab in EVENTS:
        r = best.get(et, {}).get(sym)
        w = f'{r["win"]:.0f}%' if r else "—"
        rows += (f'<div class="er"><span class="el">{lab}{" ★" if et==star_et else ""}</span>'
                 f'<span class="ew {wclass(r["win"] if r else None, dark=True)}">{w}</span></div>')
    return (f'<div class="ac"><div class="h"><span class="sym">{sym}</span><span class="px">{price} · {pe}</span></div>'
            f'<div class="nm">{name}</div><div class="fr">{fr}</div><div class="ev">{rows}</div></div>')


def main():
    sc = {s["symbol"]: s for s in get(f"{BASE}/api/screener")["stocks"]}
    rk = get(f"{BASE}/api/rankings")["ranks"]
    best = {et: {r["symbol"]: r for r in lst} for et, lst in rk.items()}
    try:
        nd = pd.read_excel(ROOT / "new data" / "08_Company_Directory.xlsx", dtype=str)
        names = {k: (v or "").replace(" Ltd.", "").replace(" Limited", "")
                 for k, v in zip(nd["symbol"], nd["company_name"])}
    except Exception:                                  # noqa: BLE001
        names = {}

    known = {s for _, syms in SECTORS for s in syms}
    extra = [s for s in sc if s not in known]
    sectors = SECTORS + ([("Other", extra)] if extra else [])

    # ---- 1) standalone light doc ----
    light_sec = ""
    for sec, syms in sectors:
        present = [s for s in syms if s in sc]
        if not present:
            continue
        cards = "".join(light_card(s, sc[s], best, names.get(s, "")) for s in present)
        light_sec += (f'<div class="sec">{sec}<span class="cnt">{len(present)} stocks</span></div>'
                      f'<div class="grid">{cards}</div>')
    css = """
:root{--bg:#f6f8fb;--panel:#fff;--ink:#111726;--mute:#5a6474;--faint:#8a94a6;--line:#e4e8ee;
 --blue:#2563eb;--g:#15803d;--a:#b45309;--r:#c2334d;--gb:#e5f3ea;--ab:#fbefd8;--rb:#fbe0e4;
 --mono:'SFMono-Regular',Consolas,Menlo,monospace;--shadow:0 1px 2px rgba(16,24,40,.05),0 6px 18px rgba(16,24,40,.06)}
*{box-sizing:border-box}html,body{margin:0;background:var(--bg);color:var(--ink);font-size:13px;line-height:1.5;
 font-family:'Inter',-apple-system,'Segoe UI',Roboto,Arial,sans-serif;-webkit-print-color-adjust:exact;print-color-adjust:exact}
.wrap{max-width:1080px;margin:0 auto;padding:0 22px 60px}
header{padding:40px 0 18px;border-bottom:2px solid var(--ink)}
.kick{font-family:var(--mono);font-size:.72rem;letter-spacing:.12em;text-transform:uppercase;color:var(--blue);font-weight:700}
h1{font-size:2rem;margin:.25em 0 .2em;letter-spacing:-.02em}.lede{color:var(--mute);margin:0;max-width:80ch}
.legend{display:flex;gap:16px;flex-wrap:wrap;margin:14px 0 4px;font-size:.78rem;color:var(--mute);font-family:var(--mono)}
.sw{width:11px;height:11px;border-radius:3px;display:inline-block;vertical-align:-1px;margin-right:5px}
.sec{font-size:1.2rem;font-weight:800;margin:26px 0 12px;padding-bottom:6px;border-bottom:2px solid var(--line);
 display:flex;align-items:baseline;gap:10px}
.sec .cnt{font-family:var(--mono);font-size:.75rem;color:var(--faint);font-weight:400}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.pc{background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:15px 17px;box-shadow:var(--shadow);break-inside:avoid}
.pch{display:flex;justify-content:space-between;align-items:baseline;gap:10px}
.sym{font-size:1.15rem;font-weight:800;letter-spacing:-.01em}
.px{font-family:var(--mono);font-size:.8rem;color:var(--mute);white-space:nowrap}
.nm{font-size:.82rem;color:var(--faint);margin:1px 0 10px}
.fund{display:grid;grid-template-columns:repeat(3,1fr);gap:7px 10px;padding:10px 0;border-top:1px solid var(--line);border-bottom:1px solid var(--line)}
.fc{display:flex;flex-direction:column}
.fl{font-size:.66rem;text-transform:uppercase;letter-spacing:.03em;color:var(--faint);font-family:var(--mono)}
.fv{font-weight:700;font-size:.92rem}.fv.neg{color:var(--r)}
.evhd{font-size:.72rem;font-weight:700;color:var(--mute);margin:10px 0 6px}
.evhd .hint{font-weight:400;color:var(--faint);font-family:var(--mono);font-size:.68rem}
.ev{display:flex;flex-direction:column;gap:4px}
.er{display:grid;grid-template-columns:74px 1fr auto auto;align-items:center;gap:8px;font-size:.82rem}
.el{color:var(--ink);font-weight:600}.erule{font-family:var(--mono);font-size:.76rem;color:var(--mute)}
.ew{font-family:var(--mono);font-weight:800;text-align:right;padding:.05em .45em;border-radius:6px}
.ew.g{color:var(--g);background:var(--gb)}.ew.a{color:var(--a);background:var(--ab)}
.ew.r{color:var(--r);background:var(--rb)}.ew.na{color:var(--faint)}
.en{font-family:var(--mono);font-size:.68rem;color:var(--faint);width:44px;text-align:right}
.foot{color:var(--mute);font-size:.8rem;margin-top:26px;font-family:var(--mono);border-top:1px solid var(--line);padding-top:12px}
@media print{@page{size:A4;margin:12mm}.wrap{max-width:100%;padding:0}.pc{box-shadow:none}.sec{break-after:avoid}header{padding-top:0}}
"""
    html = (f'<!DOCTYPE html><html lang="en"><head><meta charset="utf-8">'
            f'<meta name="viewport" content="width=device-width, initial-scale=1">'
            f'<title>Nifty 50 — Stock Profiles by Sector</title><style>{css}</style></head><body><div class="wrap">'
            f'<header><div class="kick">NSE India · Nifty 50 · Stock Profiles</div>'
            f'<h1>Every Nifty 50 stock, by sector</h1>'
            f'<p class="lede">All {len(sc)} constituents grouped into sectors — headline fundamentals plus the best '
            f'historical “buy-before / sell-after” rule &amp; win rate for each event type. ★ marks each stock’s '
            f'strongest event pattern.</p>'
            f'<div class="legend"><span><span class="sw" style="background:var(--g)"></span>win ≥ 60% (strong)</span>'
            f'<span><span class="sw" style="background:var(--a)"></span>50–59% (some edge)</span>'
            f'<span><span class="sw" style="background:var(--r)"></span>&lt; 50% (weak)</span>'
            f'<span>— = not applicable</span></div></header>{light_sec}'
            f'<p class="foot">Fundamentals: latest annual (Screener) · win rate = share of past events where the best '
            f'rule made money (in-sample) · read with the event count (n) · for study, not investment advice.</p>'
            f'</div></body></html>')
    (ROOT / "Nifty50_Stock_Profiles.html").write_text(html, encoding="utf-8")
    print(f"wrote Nifty50_Stock_Profiles.html  ({len(sc)} stocks, {len(sectors)} sectors)")

    # ---- 2) dark appendix injected into the presentation ----
    dark_sec = '<div class="ahead">All 50 stocks — by sector</div>'
    for sec, syms in sectors:
        present = [s for s in syms if s in sc]
        if not present:
            continue
        cards = "".join(dark_card(s, sc[s], best, names.get(s, "")) for s in present)
        dark_sec += (f'<div class="sec">{sec}<span class="cnt">{len(present)}</span></div>'
                     f'<div class="apxgrid">{cards}</div>')
    pres = ROOT / "Nifty50_Presentation.html"
    if pres.exists():
        p = pres.read_text(encoding="utf-8")
        payload = "<!--APPENDIX_CARDS-->" + dark_sec      # keep the marker so re-runs stay idempotent
        new = re.sub(r'(<div class="apx">).*?(</div>\s*</body>)',
                     lambda m: m.group(1) + payload + m.group(2), p, flags=re.DOTALL)
        if new != p:
            pres.write_text(new, encoding="utf-8")
            print("injected sector appendix into Nifty50_Presentation.html")
        else:
            print("! appendix container not found in presentation")


if __name__ == "__main__":
    main()
