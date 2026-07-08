"""
stock_server.py
===============================================================================
NSE STOCK BROWSER — web version, hosted on localhost.
-------------------------------------------------------------------------------
"One stock at a time, start from the first", served as a web page.

TABS per stock:
    Fundamentals      — snapshot cards + annual & quarterly charts (all stocks)
    Announcements     — NSE announcement feed        }
    Board Meetings    — board-meeting intimations     }  read from a per-stock
    Corporate Actions — dividends / bonus / splits …   }  <SYMBOL>/ folder
    Financial Results — quarterly / annual result rows}

The event tabs read <SYMBOL>/announcements.csv, board_meetings.csv,
corporate_actions.csv and financial_results.csv. Right now only the RELIANCE/
folder has these, so RELIANCE shows full detail and other stocks show a clear
"no event feed yet" note — the moment you drop those four CSVs into any
<SYMBOL>/ folder, its tabs light up automatically.

  * Pure Python standard library — NO Flask / pip install needed.
  * Fundamentals reuse fund_loader.load_stock(); fully OFFLINE.

RUN
    python stock_server.py              # http://localhost:8000
    python stock_server.py 8080         # choose a port
===============================================================================
"""

from __future__ import annotations

import json
import ssl
import sys
import threading
import time
import urllib.request
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse, parse_qs

import pandas as pd

import fund_loader

ROOT = Path(__file__).resolve().parent
STATEMENT_DIRS = ["pnl", "quarterly", "ratios", "balance_sheet", "cash_flow"]
MAX_EVENT_ROWS = 500          # cap rows sent per event feed (newest first)


# ---------------------------------------------------------------------------
# Symbol discovery
# ---------------------------------------------------------------------------
def discover_symbols() -> list[str]:
    syms: set[str] = set()
    for st in STATEMENT_DIRS:
        d = ROOT / st
        if d.exists():
            syms.update(p.stem for p in d.glob("*.csv"))
    # Drop purely-numeric BSE scrip codes (e.g. 500142) — those are old/delisted
    # entries with stale, incomplete data. Keep proper alphabetic NSE symbols.
    return sorted(s for s in syms if not s.isdigit())


# ---------------------------------------------------------------------------
# Fundamentals payload
# ---------------------------------------------------------------------------
def _col(df, name):
    if df is None:
        return None
    for c in df.columns:
        if c.strip() == name:
            return c
    for c in df.columns:
        if name.lower() in c.lower():
            return c
    return None


def _num(v):
    try:
        if v is None or pd.isna(v):
            return None
        return float(v)
    except (TypeError, ValueError):
        return None


def _latest(df, name):
    col = _col(df, name)
    if col is None:
        return None
    s = df[col].dropna()
    return _num(s.iloc[-1]) if not s.empty else None


def _series(df, name):
    col = _col(df, name)
    if df is None or col is None:
        return None
    return [_num(v) for v in df[col].tolist()]


def build_payload(sym: str) -> dict:
    data = fund_loader.load_stock(sym)
    pnl = data.get("pnl")
    q = data.get("quarterly")
    rat = data.get("ratios")
    bs = data.get("balance_sheet")

    eq, res, bor = (_latest(bs, "Equity Capital"),
                    _latest(bs, "Reserves"),
                    _latest(bs, "Borrowings"))
    de = None
    if None not in (eq, res, bor) and (eq + res):
        de = round(bor / (eq + res), 2)

    annual = {"labels": [], "sales": [], "net_profit": []}
    if pnl is not None and not pnl.empty:
        annual["labels"] = [d.strftime("%Y") for d in pnl.index]
        annual["sales"] = _series(pnl, "Sales") or []
        annual["net_profit"] = _series(pnl, "Net Profit") or []

    quarterly = {"labels": [], "net_profit": [], "yoy": []}
    if q is not None and not q.empty:
        sub = q.tail(12)
        quarterly["labels"] = [d.strftime("%b-%y") for d in sub.index]
        npc = _col(sub, "Net Profit")
        yoyc = _col(sub, "YOY Profit Growth %")
        quarterly["net_profit"] = [_num(v) for v in sub[npc].tolist()] if npc else []
        quarterly["yoy"] = [_num(v) for v in sub[yoyc].tolist()] if yoyc else []

    return {
        "symbol": sym,
        "have": [k for k in STATEMENT_DIRS if k in data],
        "cards": {
            "sales": _latest(pnl, "Sales"),
            "sales_g": _latest(pnl, "Sales Growth %"),
            "net_profit": _latest(pnl, "Net Profit"),
            "net_profit_g": _latest(pnl, "Profit Growth %"),
            "opm": _latest(pnl, "OPM %"),
            "roce": _latest(rat, "ROCE %"),
            "de": de,
        },
        "annual": annual,
        "quarterly": quarterly,
    }


# ---------------------------------------------------------------------------
# Event feeds payload  (announcements / board / corp-actions / results)
# ---------------------------------------------------------------------------
# feed key -> (nice label, date column for sorting, [(col, header, type)])
EVENT_FEEDS = {
    "announcements": {
        "label": "Announcements",
        "date": "sort_date",
        "columns": [
            ("sort_date", "Date", "date"),
            ("desc", "Category", "text"),
            ("attchmntText", "Details", "text"),
            ("attchmntFile", "File", "link"),
        ],
    },
    "board_meetings": {
        "label": "Board Meetings",
        "date": "bm_date",
        "columns": [
            ("bm_date", "Meeting date", "date"),
            ("bm_purpose", "Purpose", "text"),
            ("bm_desc", "Description", "text"),
            ("attachment", "File", "link"),
        ],
    },
    "corporate_actions": {
        "label": "Corporate Actions",
        "date": "exDate",
        "columns": [
            ("exDate", "Ex-date", "date"),
            ("recDate", "Record date", "date"),
            ("subject", "Subject", "text"),
            ("faceVal", "Face value", "text"),
        ],
    },
    "financial_results": {
        "label": "Financial Results",
        "date": "broadCastDate",
        "columns": [
            ("broadCastDate", "Broadcast", "date"),
            ("relatingTo", "Period", "text"),
            ("consolidated", "Basis", "text"),
            ("audited", "Audited", "text"),
            ("fromDate", "From", "date"),
            ("toDate", "To", "date"),
            ("xbrl", "XBRL", "link"),
        ],
    },
}


def _fmt_date(v: str) -> str:
    if not v:
        return ""
    ts = pd.to_datetime(v, errors="coerce", dayfirst=True, format="mixed")
    return ts.strftime("%d-%b-%Y") if pd.notna(ts) else str(v)[:20]


def build_events(sym: str) -> dict:
    """Read the four per-stock event CSVs from <SYMBOL>/ if present."""
    folder = ROOT / sym
    out = {"symbol": sym, "available": folder.exists(), "feeds": {}}

    for key, spec in EVENT_FEEDS.items():
        cols = spec["columns"]
        entry = {
            "label": spec["label"],
            "columns": [{"key": k, "label": lbl, "type": t} for k, lbl, t in cols],
            "rows": [], "total": 0, "shown": 0,
        }
        path = folder / f"{key}.csv"
        if path.exists():
            out["available"] = True
            try:
                df = pd.read_csv(path, dtype=str).fillna("")
            except Exception as e:                       # noqa: BLE001
                entry["error"] = f"{type(e).__name__}: {e}"
                out["feeds"][key] = entry
                continue
            dc = spec["date"]
            if dc in df.columns:
                df["_sort"] = pd.to_datetime(df[dc], errors="coerce", dayfirst=True)
                df = df.sort_values("_sort", ascending=False, na_position="last")
            entry["total"] = int(len(df))
            df = df.head(MAX_EVENT_ROWS)
            entry["shown"] = int(len(df))
            rows = []
            for _, r in df.iterrows():
                row = {}
                for k, _lbl, t in cols:
                    val = r[k] if k in df.columns else ""
                    row[k] = _fmt_date(val) if t == "date" else val
                rows.append(row)
            entry["rows"] = rows
        out["feeds"][key] = entry
    return out


def downloaded_symbols() -> list[str]:
    """Symbols whose <sym>/ folder already holds at least one event CSV."""
    files = [f"{k}.csv" for k in EVENT_FEEDS]
    ready = []
    for s in SYMBOLS:
        folder = ROOT / s
        if folder.exists() and any((folder / f).exists() for f in files):
            ready.append(s)
    return ready


# ---------------------------------------------------------------------------
# NSE downloader — fetch the four event feeds for ANY symbol on demand
# ---------------------------------------------------------------------------
_NSE_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/122.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.nseindia.com/companies-listing/corporate-filings-announcements",
}
_NSE_CTX = ssl.create_default_context()
_NSE_CTX.check_hostname = False
_NSE_CTX.verify_mode = ssl.CERT_NONE

NSE_API = {
    "announcements":     "https://www.nseindia.com/api/corporate-announcements?index=equities&symbol={s}",
    "board_meetings":    "https://www.nseindia.com/api/corporate-board-meetings?index=equities&symbol={s}",
    "corporate_actions": "https://www.nseindia.com/api/corporates-corporateActions?index=equities&symbol={s}",
}
NSE_RESULTS = "https://www.nseindia.com/api/corporates-financial-results?index=equities&symbol={s}&period={p}"


# ONE long-lived, cookie-primed NSE session, reused across every download so we
# don't re-hit NSE's homepage (and trip its bot-wall) on every request.
_NSE_SESSION = {"op": None}
_NSE_LOCK = threading.Lock()


def _nse_prime(sym: str):
    """Build a fresh cookie-primed opener (homepage + the stock's quote page)."""
    op = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(),
                                     urllib.request.HTTPSHandler(context=_NSE_CTX))
    for u in ("https://www.nseindia.com/",
              f"https://www.nseindia.com/get-quotes/equity?symbol={sym}"):
        try:
            op.open(urllib.request.Request(u, headers=_NSE_HEADERS), timeout=25).read()
        except Exception:                                # noqa: BLE001 - keep warming
            pass
        time.sleep(0.6)
    return op


def _nse_rows(sym: str, url: str) -> list:
    """GET one NSE json api, reusing the shared session; re-prime + back off on
    rejection (403 / empty). Raises the last error only if every attempt fails."""
    last = None
    for delay in (0, 2, 5):                              # exponential-ish backoff
        if delay:
            time.sleep(delay)
        with _NSE_LOCK:
            if _NSE_SESSION["op"] is None:
                _NSE_SESSION["op"] = _nse_prime(sym)
            op = _NSE_SESSION["op"]
        try:
            raw = op.open(urllib.request.Request(url, headers=_NSE_HEADERS),
                          timeout=25).read().decode("utf-8", "replace").strip()
            if raw and raw[0] in "[{":
                d = json.loads(raw)
                return d if isinstance(d, list) else d.get("data", [])
            last = ValueError("empty response from NSE")
        except Exception as e:                           # noqa: BLE001
            last = e
        with _NSE_LOCK:                                  # session went bad -> refresh
            _NSE_SESSION["op"] = None
    if last:
        raise last
    return []


def fetch_nse_feeds(sym: str) -> dict:
    """Download the four event feeds for `sym` -> <sym>/*.csv.

    Every feed is isolated: one feed failing (403/empty) never kills the others,
    and whatever succeeds is saved. Returns per-feed counts + any errors.
    """
    folder = ROOT / sym
    folder.mkdir(exist_ok=True)
    counts, errors = {}, {}

    for key, tpl in NSE_API.items():
        try:
            r = _nse_rows(sym, tpl.format(s=sym))
            if r:
                pd.DataFrame(r).to_csv(folder / f"{key}.csv", index=False)
            counts[key] = len(r)
        except Exception as e:                           # noqa: BLE001
            errors[key] = f"{type(e).__name__}: {e}"
            counts[key] = 0
        time.sleep(0.5)

    fr = []
    for period in ("Quarterly", "Annual"):
        try:
            fr += _nse_rows(sym, NSE_RESULTS.format(s=sym, p=period))
        except Exception as e:                           # noqa: BLE001
            errors[f"results_{period.lower()}"] = f"{type(e).__name__}: {e}"
        time.sleep(0.5)
    if fr:
        df = pd.DataFrame(fr)
        try:
            df = df.drop_duplicates()
        except Exception:                                # noqa: BLE001
            pass
        df.to_csv(folder / "financial_results.csv", index=False)
    counts["financial_results"] = len(fr)

    return {"counts": counts, "errors": errors, "any": any(counts.values())}


# ---------------------------------------------------------------------------
# The single-page front end (inline HTML + CSS + JS, canvas charts, no CDN)
# ---------------------------------------------------------------------------
PAGE = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>NSE Stock Browser</title>
<style>
  :root{
    --bg:#0f131b; --panel:#171d2b; --grid:#2a3346; --ink:#e6edf3; --mute:#8b98ad;
    --pos:#3fb950; --neg:#f85149; --blue:#58a6ff; --amber:#d29922;
    --mono:ui-monospace,"Cascadia Code",Consolas,monospace;
    --sans:-apple-system,"Segoe UI",system-ui,sans-serif;
  }
  *{box-sizing:border-box}
  body{margin:0;background:var(--bg);color:var(--ink);font-family:var(--sans)}
  .bar{display:flex;align-items:center;gap:10px;flex-wrap:wrap;
       padding:14px 20px;border-bottom:1px solid var(--grid);position:sticky;top:0;
       background:var(--bg);z-index:5}
  .bar h1{font-size:1.15rem;margin:0;font-weight:650;margin-right:auto}
  button{background:var(--panel);color:var(--ink);border:1px solid var(--grid);
    border-radius:7px;padding:7px 12px;font-family:var(--sans);font-weight:600;
    font-size:.9rem;cursor:pointer}
  button:hover{background:#243044}
  button:active{transform:translateY(1px)}
  input{background:var(--panel);color:var(--ink);border:1px solid var(--grid);
    border-radius:7px;padding:7px 12px;font-family:var(--mono);font-size:1rem;
    width:200px;text-transform:uppercase}
  input::placeholder{text-transform:none;color:var(--mute)}
  input:focus{outline:2px solid var(--blue)}
  .searchwrap{position:relative;display:inline-block}
  .ac{position:absolute;top:calc(100% + 4px);left:0;width:260px;max-height:340px;
    overflow-y:auto;background:var(--panel);border:1px solid var(--grid);border-radius:8px;
    box-shadow:0 10px 30px rgba(0,0,0,.5);z-index:30;display:none}
  .acitem{padding:8px 12px;font-family:var(--mono);font-size:.9rem;cursor:pointer;
    color:var(--mute);border-bottom:1px solid #1e2636;white-space:nowrap;
    display:flex;justify-content:space-between;align-items:center;gap:12px}
  .acitem:last-child{border-bottom:0}
  .acitem b{color:var(--blue);font-weight:700}
  .acitem.sel,.acitem:hover{background:#243044;color:var(--ink)}
  .rdy{color:var(--pos);font-size:.68rem;font-weight:700;white-space:nowrap}
  .notrdy{color:var(--mute);font-size:.68rem;white-space:nowrap;opacity:.7}
  #evstatus{margin-left:10px;font-family:var(--mono);font-size:.75rem}
  .acempty{padding:10px 12px;color:var(--mute);font-family:var(--mono);font-size:.85rem}
  .acmore{padding:7px 12px;color:var(--mute);font-family:var(--mono);font-size:.75rem;
    border-top:1px solid var(--grid)}
  .counter{font-family:var(--mono);color:var(--mute);font-size:.9rem;min-width:110px;
    text-align:right}
  main{max-width:1040px;margin:0 auto;padding:18px 20px 40px}
  .sym{font-size:2rem;font-weight:700;letter-spacing:-.01em}
  .have{font-family:var(--mono);color:var(--mute);font-size:.8rem;margin-left:10px}

  /* tabs */
  .tabs{display:flex;gap:4px;flex-wrap:wrap;margin:14px 0 16px;
        border-bottom:1px solid var(--grid)}
  .tab{background:transparent;border:0;border-bottom:2px solid transparent;
       color:var(--mute);padding:9px 14px;font-weight:600;font-size:.9rem;
       cursor:pointer;border-radius:0}
  .tab:hover{background:transparent;color:var(--ink)}
  .tab.active{color:var(--blue);border-bottom-color:var(--blue)}

  .cards{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin:0 0 18px}
  @media(max-width:720px){.cards{grid-template-columns:repeat(2,1fr)}}
  .card{background:var(--panel);border:1px solid var(--grid);border-radius:10px;
    padding:14px 16px}
  .card .t{color:var(--mute);font-size:.8rem}
  .card .v{font-size:1.5rem;font-weight:700;margin:2px 0}
  .card .s{color:var(--mute);font-size:.75rem;font-family:var(--mono)}
  .charts{display:grid;grid-template-columns:1fr 1fr;gap:12px}
  @media(max-width:720px){.charts{grid-template-columns:1fr}}
  .chart{background:var(--panel);border:1px solid var(--grid);border-radius:10px;padding:12px}
  .chart .ct{font-weight:650;font-size:.95rem}
  .chart .cs{color:var(--mute);font-size:.78rem;float:right;font-family:var(--mono)}
  canvas{width:100%;height:300px;display:block;margin-top:6px}

  /* event tables */
  .tblwrap{border:1px solid var(--grid);border-radius:10px;overflow:auto;max-height:62vh;
    background:var(--panel)}
  table{border-collapse:collapse;width:100%;font-size:.86rem}
  thead th{position:sticky;top:0;background:#1c2434;color:var(--mute);
    font-family:var(--mono);font-size:.72rem;letter-spacing:.04em;text-transform:uppercase;
    text-align:left;padding:9px 12px;border-bottom:1px solid var(--grid);white-space:nowrap}
  td{padding:9px 12px;border-bottom:1px solid var(--grid);vertical-align:top}
  tbody tr:hover{background:#1b2333}
  td:nth-child(1){white-space:nowrap;font-family:var(--mono);color:var(--ink)}
  .cat{font-family:var(--mono);font-size:.78rem;color:var(--amber)}
  .enote{color:var(--mute);font-size:.85rem;padding:10px 4px;font-family:var(--mono)}
  .banner{background:var(--panel);border:1px solid var(--grid);border-radius:10px;
    padding:22px;color:var(--mute);text-align:center;line-height:1.6}
  .banner b{color:var(--ink)}
  .dl{background:var(--blue);color:#0b1220;border:0;border-radius:8px;padding:10px 18px;
    font-weight:700;font-size:.95rem;cursor:pointer}
  .dl:hover{background:#7cb8ff}
  a{color:var(--blue)}
  .hint{color:var(--mute);font-size:.82rem;margin-top:16px;font-family:var(--mono)}
</style>
</head>
<body>
  <div class="bar">
    <h1>NSE Stock Browser</h1>
    <button id="first">◀◀ First</button>
    <button id="prev">◀ Prev</button>
    <span class="searchwrap">
      <input id="search" placeholder="🔍 search symbol…" autocomplete="off">
      <div id="ac" class="ac"></div>
    </span>
    <button id="next">Next ▶</button>
    <button id="last">Last ▶▶</button>
    <span class="counter" id="counter">—</span>
  </div>

  <main>
    <div><span class="sym" id="sym">…</span><span class="have" id="have"></span>
         <span id="evstatus"></span></div>

    <div class="tabs" id="tabs">
      <button class="tab active" data-tab="fundamentals">Fundamentals</button>
      <button class="tab" data-tab="announcements">Announcements</button>
      <button class="tab" data-tab="board_meetings">Board Meetings</button>
      <button class="tab" data-tab="corporate_actions">Corporate Actions</button>
      <button class="tab" data-tab="financial_results">Financial Results</button>
    </div>

    <!-- Fundamentals view -->
    <div id="fundview">
      <div class="cards" id="cards"></div>
      <div class="charts">
        <div class="chart">
          <span class="cs">Sales vs Net Profit</span><span class="ct">Annual trend (₹ cr)</span>
          <canvas id="c_annual"></canvas>
        </div>
        <div class="chart">
          <span class="cs">green = YoY up</span><span class="ct">Quarterly Net Profit (₹ cr)</span>
          <canvas id="c_quarter"></canvas>
        </div>
      </div>
    </div>

    <!-- Event view -->
    <div id="eventsview" style="display:none"></div>

    <div class="hint">🔍 search any of 2,363 stocks · ← / → arrow keys to step · click a tab for event details.</div>
  </main>

<script>
const COL={ink:"#e6edf3",mute:"#8b98ad",grid:"#2a3346",panel:"#171d2b",
           blue:"#58a6ff",amber:"#d29922",pos:"#3fb950",neg:"#f85149"};
let SYMS=[], idx=0, activeTab="fundamentals", stock=null, eventsCache={}, READY=new Set();

/*__DATA_SOURCE__*/           /* swapped by build_static.py for the GitHub Pages build */
const STATIC=false;
const SYMBOLS_URL="/api/symbols";
function STOCK_URL(s){return "/api/stock?sym="+encodeURIComponent(s);}
function EVENTS_URL(s){return "/api/events?sym="+encodeURIComponent(s);}
/*__END_DATA_SOURCE__*/

const $=id=>document.getElementById(id);
function fmt(v,suf="",nd=0){ return (v===null||v===undefined)?"—":
  Number(v).toLocaleString("en-IN",{minimumFractionDigits:nd,maximumFractionDigits:nd})+suf; }
function esc(s){ return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;"); }

/* ---------- canvas charts (HiDPI aware) ---------- */
function setup(cv){
  const r=cv.getBoundingClientRect(), dpr=window.devicePixelRatio||1;
  cv.width=r.width*dpr; cv.height=r.height*dpr;
  const ctx=cv.getContext("2d"); ctx.scale(dpr,dpr);
  return {ctx,w:r.width,h:r.height};
}
const M={l:56,r:18,t:14,b:30};
function niceBounds(vals,incZero){
  let v=vals.filter(x=>x!==null&&!isNaN(x)); if(incZero)v=v.concat([0]);
  if(!v.length)return null;
  let lo=Math.min(...v),hi=Math.max(...v),pad=((hi-lo)||1)*0.12;
  return {lo:lo-pad,hi:hi+pad};
}
function yaxis(ctx,w,h,lo,hi){
  ctx.font="11px Consolas"; ctx.fillStyle=COL.mute; ctx.strokeStyle=COL.grid;
  ctx.textAlign="right"; ctx.textBaseline="middle";
  for(let i=0;i<=5;i++){
    const val=lo+(hi-lo)*i/5, y=h-M.b-(val-lo)/((hi-lo)||1)*(h-M.b-M.t);
    ctx.beginPath();ctx.moveTo(M.l,y);ctx.lineTo(w-M.r,y);ctx.stroke();
    ctx.fillText(Math.round(val).toLocaleString("en-IN"),M.l-6,y);
  }
}
function xlab(ctx,w,h,labels){
  ctx.font="10px Consolas"; ctx.fillStyle=COL.mute; ctx.textAlign="center"; ctx.textBaseline="top";
  const n=labels.length, iw=w-M.l-M.r;
  labels.forEach((L,i)=>{const x=M.l+(n<2?iw/2:iw*i/(n-1)); ctx.fillText(L,x,h-M.b+6);});
}
function empty(ctx,w,h,msg){ ctx.fillStyle=COL.mute;ctx.font="13px sans-serif";
  ctx.textAlign="center";ctx.textBaseline="middle";ctx.fillText(msg||"no data",w/2,h/2); }

function lineChart(cv,labels,seriesList){
  const {ctx,w,h}=setup(cv); ctx.clearRect(0,0,w,h);
  const all=seriesList.flatMap(s=>s.data);
  const b=niceBounds(all,false); if(!b||!labels.length){empty(ctx,w,h,"no annual P&L");return;}
  yaxis(ctx,w,h,b.lo,b.hi); xlab(ctx,w,h,labels);
  const n=labels.length,iw=w-M.l-M.r;
  const X=i=>M.l+(n<2?iw/2:iw*i/(n-1));
  const Y=v=>h-M.b-(v-b.lo)/((b.hi-b.lo)||1)*(h-M.b-M.t);
  seriesList.forEach((s,si)=>{
    ctx.strokeStyle=s.color;ctx.lineWidth=2;ctx.beginPath();let started=false;
    s.data.forEach((v,i)=>{ if(v===null||isNaN(v))return;
      const x=X(i),y=Y(v); started?ctx.lineTo(x,y):ctx.moveTo(x,y); started=true; });
    ctx.stroke();
    s.data.forEach((v,i)=>{ if(v===null||isNaN(v))return;
      ctx.fillStyle=s.color;ctx.beginPath();ctx.arc(X(i),Y(v),3,0,7);ctx.fill(); });
    ctx.fillStyle=COL.ink;ctx.font="11px sans-serif";ctx.textAlign="left";ctx.textBaseline="middle";
    const ly=M.t+6+si*15;
    ctx.strokeStyle=s.color;ctx.lineWidth=3;ctx.beginPath();
    ctx.moveTo(w-M.r-118,ly);ctx.lineTo(w-M.r-98,ly);ctx.stroke();
    ctx.fillText(s.name,w-M.r-92,ly);
  });
}
function barChart(cv,labels,vals,colors){
  const {ctx,w,h}=setup(cv); ctx.clearRect(0,0,w,h);
  const b=niceBounds(vals,true); if(!b||!labels.length){empty(ctx,w,h,"no quarterly data");return;}
  yaxis(ctx,w,h,b.lo,b.hi); xlab(ctx,w,h,labels);
  const n=labels.length,iw=w-M.l-M.r,slot=iw/n,bw=slot*0.6;
  const Y=v=>h-M.b-(v-b.lo)/((b.hi-b.lo)||1)*(h-M.b-M.t);
  const yz=Y(0);
  ctx.strokeStyle=COL.mute;ctx.beginPath();ctx.moveTo(M.l,yz);ctx.lineTo(w-M.r,yz);ctx.stroke();
  vals.forEach((v,i)=>{ if(v===null||isNaN(v))return;
    const cx=M.l+slot*(i+0.5),y=Y(v);
    ctx.fillStyle=colors[i]||COL.blue;
    ctx.fillRect(cx-bw/2,Math.min(yz,y),bw,Math.abs(yz-y));
    ctx.fillStyle=COL.ink;ctx.font="10px Consolas";ctx.textAlign="center";
    ctx.textBaseline=v>=0?"bottom":"top";
    ctx.fillText(Math.round(v).toLocaleString("en-IN"),cx,v>=0?y-3:y+3);
  });
}

/* ---------- data ---------- */
async function loadStock(){
  const sym=SYMS[idx];
  $("counter").textContent=(idx+1).toLocaleString()+" / "+SYMS.length.toLocaleString();
  $("search").value=sym;
  stock=await (await fetch(STOCK_URL(sym))).json();
  $("sym").textContent=stock.symbol;
  $("have").textContent="statements: "+(stock.have.join(", ")||"none");
  updateEvStatus();
  showTab(activeTab);
}
function updateEvStatus(){
  const on=READY.has(SYMS[idx]);
  $("evstatus").innerHTML=on
    ? '<span class="rdy">✓ event feeds downloaded</span>'
    : '<span class="notrdy">○ event feeds not downloaded</span>';
}
async function ensureEvents(sym){
  if(!eventsCache[sym]){
    try{
      const r=await fetch(EVENTS_URL(sym));
      eventsCache[sym]= r.ok ? await r.json() : {symbol:sym,available:false,feeds:{}};
    }catch(e){ eventsCache[sym]={symbol:sym,available:false,feeds:{}}; }
  }
  return eventsCache[sym];
}

/* ---------- rendering ---------- */
function renderFundamentals(){
  const c=stock.cards;
  const npCol=(c.net_profit!==null&&c.net_profit<0)?COL.neg:COL.pos;
  $("cards").innerHTML=[
    card("Sales (latest FY)",fmt(c.sales," cr"),fmt(c.sales_g,"% growth",1),COL.ink),
    card("Net Profit",fmt(c.net_profit," cr"),fmt(c.net_profit_g,"% growth",1),npCol),
    card("Operating margin",fmt(c.opm,"%",1),"OPM",COL.blue),
    card("ROCE",fmt(c.roce,"%",1),"Debt/Equity "+fmt(c.de,"",2),COL.amber),
  ].join("");
  lineChart($("c_annual"),stock.annual.labels,[
    {name:"Sales",color:COL.blue,data:stock.annual.sales},
    {name:"Net Profit",color:COL.amber,data:stock.annual.net_profit},
  ]);
  const qc=stock.quarterly.net_profit.map((_,i)=>{
    const g=stock.quarterly.yoy[i]; return (g!==undefined&&g!==null&&g>=0)?COL.pos:COL.neg; });
  barChart($("c_quarter"),stock.quarterly.labels,stock.quarterly.net_profit,qc);
}
function card(t,v,s,col){
  return `<div class="card"><div class="t">${t}</div>`+
         `<div class="v" style="color:${col}">${v}</div><div class="s">${s}</div></div>`;
}

async function renderEvents(feed){
  const ev=$("eventsview");
  ev.innerHTML='<div class="enote">loading…</div>';
  const pay=await ensureEvents(SYMS[idx]);
  if(feed!==activeTab)return;                    // stale (user moved on)
  if(!pay.available){
    ev.innerHTML=`<div class="banner">No event feeds for <b>${esc(pay.symbol)}</b> yet.<br>`+
      `Announcements, board meetings, corporate actions and financial results can be pulled `+
      `live from NSE.<br><br>`+
      `<button class="dl" onclick="fetchNSE()">⬇ Download NSE feeds for ${esc(pay.symbol)}</button></div>`;
    return;
  }
  const F=pay.feeds[feed];
  if(!F||!F.rows.length){
    ev.innerHTML=`<div class="banner">No <b>${esc((F&&F.label)||feed)}</b> rows for <b>${esc(pay.symbol)}</b>.<br><br>`+
      `<button class="dl" onclick="fetchNSE()">⬇ Re-download NSE feeds for ${esc(pay.symbol)}</button></div>`;
    return;
  }
  let html='<div class="tblwrap"><table><thead><tr>'+
    F.columns.map(c=>`<th>${esc(c.label)}</th>`).join('')+'</tr></thead><tbody>';
  F.rows.forEach(r=>{
    html+='<tr>'+F.columns.map(c=>{
      const v=r[c.key]||'';
      if(c.type==='link') return v.startsWith('http')
        ? `<td><a href="${esc(v)}" target="_blank" rel="noopener">open ↗</a></td>` : `<td>${esc(v)}</td>`;
      if(c.key==='desc') return `<td class="cat">${esc(v)}</td>`;
      return `<td>${esc(v)}</td>`;
    }).join('')+'</tr>';
  });
  html+='</tbody></table></div>';
  if(F.total>F.shown) html+=`<div class="enote">Showing latest ${F.shown} of ${F.total.toLocaleString()} rows.</div>`;
  else html+=`<div class="enote">${F.total.toLocaleString()} row(s).</div>`;
  ev.innerHTML=html;
}

async function fetchNSE(){
  const sym=SYMS[idx];
  const ev=$("eventsview");
  if(STATIC){
    ev.innerHTML=`<div class="banner">This is a <b>static snapshot</b> — live NSE download isn't available here.<br>`+
      `To pull <b>${esc(sym)}</b>'s feeds from NSE, run the app locally:<br><br>`+
      `<span class="enote">python stock_server.py</span></div>`;
    return;
  }
  ev.innerHTML=`<div class="banner">Downloading NSE feeds for <b>${esc(sym)}</b>…<br>`+
    `<span class="enote">(announcements can be a few thousand rows — give it 5–15 s)</span></div>`;
  try{
    const j=await (await fetch("/api/fetch?sym="+encodeURIComponent(sym))).json();
    if(!j.ok){
      const reason=j.error || (j.errors && Object.values(j.errors)[0]) || "NSE refused the request";
      ev.innerHTML=`<div class="banner">Couldn't download <b>${esc(sym)}</b> from NSE.<br>`+
        `<span class="enote">${esc(reason)}</span><br><br>`+
        `NSE rate-limits rapid requests. Wait ~30–60 s, then:<br><br>`+
        `<button class="dl" onclick="fetchNSE()">Retry</button></div>`;
      return;
    }
    READY.add(sym);
    updateEvStatus();
    delete eventsCache[sym];
    renderEvents(activeTab);
  }catch(e){
    ev.innerHTML=`<div class="banner">Download error: ${esc(e)}<br><br>`+
      `NSE rate-limits rapid requests — wait ~30–60 s and retry.<br><br>`+
      `<button class="dl" onclick="fetchNSE()">Retry</button></div>`;
  }
}

function showTab(name){
  activeTab=name;
  document.querySelectorAll(".tab").forEach(t=>t.classList.toggle("active",t.dataset.tab===name));
  const isFund=name==="fundamentals";
  $("fundview").style.display=isFund?"":"none";
  $("eventsview").style.display=isFund?"none":"";
  if(isFund) renderFundamentals(); else renderEvents(name);
}

/* ---------- navigation ---------- */
function go(i){ idx=Math.max(0,Math.min(i,SYMS.length-1)); hideAC(); loadStock(); }

/* ---------- live search dropdown ---------- */
const AC=$("ac"), SEARCH=$("search");
let acItems=[], acSel=-1;
function hideAC(){ AC.style.display="none"; acSel=-1; }
function renderAC(){
  const q=SEARCH.value.trim().toUpperCase();
  if(!q){ hideAC(); return; }
  const starts=[], incl=[];
  for(const s of SYMS){
    if(s.startsWith(q)) starts.push(s);
    else if(s.includes(q)) incl.push(s);
  }
  const total=starts.length+incl.length;
  acItems=starts.concat(incl).slice(0,30);
  acSel=-1;
  if(!acItems.length){ AC.innerHTML='<div class="acempty">no match for “'+esc(q)+'”</div>'; AC.style.display="block"; return; }
  let html=acItems.map((s,i)=>{
    const k=s.indexOf(q);
    const hl=k<0?esc(s):esc(s.slice(0,k))+'<b>'+esc(s.slice(k,k+q.length))+'</b>'+esc(s.slice(k+q.length));
    const tag=READY.has(s)?'<span class="rdy">✓ data</span>':'';
    return `<div class="acitem" data-i="${i}"><span>${hl}</span>${tag}</div>`;
  }).join('');
  if(total>acItems.length) html+=`<div class="acmore">${total.toLocaleString()} matches — showing first ${acItems.length}</div>`;
  AC.innerHTML=html; AC.style.display="block";
}
function pickAC(i){
  if(i<0||i>=acItems.length)return;
  const j=SYMS.indexOf(acItems[i]);
  if(j>=0){ SEARCH.value=acItems[i]; go(j); SEARCH.blur(); }
}
function moveAC(d){
  if(AC.style.display==="none"||!acItems.length)return;
  acSel=(acSel+d+acItems.length)%acItems.length;
  [...AC.querySelectorAll(".acitem")].forEach((el,i)=>el.classList.toggle("sel",i===acSel));
  const el=AC.querySelector(".acitem.sel"); if(el)el.scrollIntoView({block:"nearest"});
}
SEARCH.addEventListener("input",renderAC);
SEARCH.addEventListener("focus",()=>{ if(SEARCH.value.trim())renderAC(); });
SEARCH.addEventListener("blur",()=>setTimeout(hideAC,150));   // allow click first
SEARCH.addEventListener("keydown",e=>{
  if(e.key==="ArrowDown"){ e.preventDefault(); moveAC(1); }
  else if(e.key==="ArrowUp"){ e.preventDefault(); moveAC(-1); }
  else if(e.key==="Enter"){
    e.preventDefault();
    if(acSel>=0) pickAC(acSel);
    else if(acItems.length) pickAC(0);            // best match
  }
  else if(e.key==="Escape"){ hideAC(); }
});
AC.addEventListener("mousedown",e=>{
  const it=e.target.closest(".acitem"); if(it) pickAC(+it.dataset.i);
});

$("first").onclick=()=>go(0);
$("prev").onclick=()=>go(idx-1);
$("next").onclick=()=>go(idx+1);
$("last").onclick=()=>go(SYMS.length-1);
document.querySelectorAll(".tab").forEach(t=>t.onclick=()=>showTab(t.dataset.tab));
document.addEventListener("keydown",e=>{
  if(document.activeElement.id==="search")return;
  if(e.key==="ArrowLeft")go(idx-1);
  else if(e.key==="ArrowRight")go(idx+1);
  else if(e.key==="Home")go(0);
  else if(e.key==="End")go(SYMS.length-1);
});
window.addEventListener("resize",()=>{ if(activeTab==="fundamentals"&&stock)renderFundamentals(); });

(async ()=>{
  const s=await (await fetch(SYMBOLS_URL)).json();
  SYMS=s.symbols;
  READY=new Set(s.ready||[]);
  go(0);                       // start from the FIRST stock
})();
</script>
</body>
</html>"""


# ---------------------------------------------------------------------------
# HTTP handler
# ---------------------------------------------------------------------------
SYMBOLS = discover_symbols()


class Handler(BaseHTTPRequestHandler):
    def _send(self, code, body, ctype):
        data = body.encode("utf-8") if isinstance(body, str) else body
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")      # always serve fresh
        self.end_headers()
        self.wfile.write(data)

    def _json(self, obj, code=200):
        self._send(code, json.dumps(obj), "application/json; charset=utf-8")

    def do_GET(self):
        parsed = urlparse(self.path)
        route = parsed.path
        qs = parse_qs(parsed.query)

        if route in ("/", "/index.html"):
            self._send(200, PAGE, "text/html; charset=utf-8")
        elif route == "/api/symbols":
            self._json({"symbols": SYMBOLS, "count": len(SYMBOLS),
                        "ready": downloaded_symbols()})
        elif route in ("/api/stock", "/api/events", "/api/fetch"):
            sym = (qs.get("sym", [""])[0]).strip().upper()
            if sym not in SYMBOLS:
                self._json({"error": f"unknown symbol '{sym}'"}, code=404)
                return
            try:
                if route == "/api/stock":
                    self._json(build_payload(sym))
                elif route == "/api/events":
                    self._json(build_events(sym))
                else:                                    # /api/fetch -> download NSE
                    res = fetch_nse_feeds(sym)
                    self._json({"ok": res["any"], "symbol": sym,
                                "counts": res["counts"], "errors": res["errors"]})
            except Exception as e:                       # noqa: BLE001
                self._json({"error": f"{type(e).__name__}: {e}"}, code=500)
        else:
            self._send(404, "not found", "text/plain; charset=utf-8")

    def log_message(self, *args):
        pass


def main():
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8000
    if not SYMBOLS:
        raise SystemExit("No stock fundamentals found in pnl/ quarterly/ ... folders.")

    url = f"http://localhost:{port}"
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    print("=" * 62)
    print(f"  NSE Stock Browser  —  serving {len(SYMBOLS):,} stocks")
    print(f"  Open:  {url}")
    print(f"  First: {SYMBOLS[0]}   Last: {SYMBOLS[-1]}")
    print("  Tabs: Fundamentals + Announcements / Board / Corp Actions / Results")
    print("  Press Ctrl+C to stop.")
    print("=" * 62)

    if "--no-open" not in sys.argv:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")
        server.shutdown()


if __name__ == "__main__":
    main()
