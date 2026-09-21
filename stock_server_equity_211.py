"""
stock_server_equity_211.py
===============================================================================
211 EQUITY STOCKS WEB APPLICATION HOST SERVER
-------------------------------------------------------------------------------
Serves all 211 Stock Symbols from D:\\behaviour analysis\\OI_DATA with:
  1. 1-Minute Intraday Equity (Spot Parquet) Price & Volume Charting
  2. Fundamentals Snapshot & Quarterly Financial Performance Cards
  3. Interactive Search & Stock Browser
  4. Localhost (8000/8080) & LAN (10.10.7.70) Binding
===============================================================================
"""

from __future__ import annotations

import json
import os
import pathlib
import re
import socket
import sys
import threading
import urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from functools import lru_cache

import pandas as pd

sys.stdout.reconfigure(errors='replace')

ROOT = Path(__file__).resolve().parent
OI_DATA_DIR = ROOT / "OI_DATA"

# ---------------------------------------------------------------------------
# 1. 211 Stock Symbol Discovery
# ---------------------------------------------------------------------------
def discover_all_211_symbols() -> list[str]:
    if OI_DATA_DIR.exists():
        symbols = sorted([e.name for e in os.scandir(OI_DATA_DIR) if e.is_dir()])
        if symbols:
            return symbols
    # Fallback to STATEMENT_DIRS if OI_DATA not available
    syms = set()
    for st in ["pnl", "quarterly", "ratios", "balance_sheet", "cash_flow"]:
        d = ROOT / st
        if d.exists():
            syms.update(p.stem for p in d.glob("*.csv"))
    return sorted(list(syms))

ALL_SYMBOLS = discover_all_211_symbols()

# ---------------------------------------------------------------------------
# 2. Equity Spot Parquet Data Loader & Caching
# ---------------------------------------------------------------------------
@lru_cache(maxsize=128)
def get_available_dates(symbol: str) -> list[str]:
    spot_dir = OI_DATA_DIR / symbol / "spot_parquet"
    if not spot_dir.exists():
        return []
    dates = [f[:-8] for f in os.listdir(spot_dir) if f.endswith(".parquet")]
    return sorted(dates)

def load_equity_spot_data(symbol: str, date_str: str | None = None) -> dict:
    dates = get_available_dates(symbol)
    if not dates:
        return {"symbol": symbol, "dates": [], "selected_date": None, "bars": []}
    
    selected_date = date_str if date_str in dates else dates[-1]
    file_path = OI_DATA_DIR / symbol / "spot_parquet" / f"{selected_date}.parquet"
    
    bars = []
    if file_path.exists():
        try:
            df = pd.read_parquet(file_path)
            for _, row in df.iterrows():
                bars.append({
                    "time": str(row.get("Time", "")),
                    "open": float(row.get("Open", 0.0)),
                    "high": float(row.get("High", 0.0)),
                    "low": float(row.get("Low", 0.0)),
                    "close": float(row.get("Close", 0.0)),
                    "volume": int(row.get("Volume", 0)),
                    "oi": int(row.get("OpenInterest", 0))
                })
        except Exception as e:
            print(f"Error loading parquet {file_path}: {e}")
            
    return {
        "symbol": symbol,
        "total_dates": len(dates),
        "dates": dates[-30:],  # Return last 30 dates for selector
        "selected_date": selected_date,
        "bars": bars
    }

# ---------------------------------------------------------------------------
# 3. Fundamentals Loader (Graceful Fallback)
# ---------------------------------------------------------------------------
def load_fundamentals_snapshot(symbol: str) -> dict:
    # Try importing fund_loader if present
    cards = {"sales": None, "net_profit": None, "opm": None, "roce": None}
    annual = {"labels": [], "sales": [], "net_profit": []}
    quarterly = {"labels": [], "net_profit": []}
    
    try:
        import fund_loader
        data = fund_loader.load_stock(symbol)
        pnl = data.get("pnl")
        q = data.get("quarterly")
        rat = data.get("ratios")
        
        if pnl is not None and not pnl.empty:
            annual["labels"] = [d.strftime("%Y") for d in pnl.index]
            for col in pnl.columns:
                if "sales" in col.lower(): annual["sales"] = [float(v) if pd.notna(v) else None for v in pnl[col]]
                if "net profit" in col.lower(): annual["net_profit"] = [float(v) if pd.notna(v) else None for v in pnl[col]]
                
        if q is not None and not q.empty:
            sub = q.tail(12)
            quarterly["labels"] = [d.strftime("%b-%y") for d in sub.index]
            for col in sub.columns:
                if "net profit" in col.lower():
                    quarterly["net_profit"] = [float(v) if pd.notna(v) else None for v in sub[col]]
    except Exception:
        pass
        
    return {
        "symbol": symbol,
        "cards": cards,
        "annual": annual,
        "quarterly": quarterly
    }

# ---------------------------------------------------------------------------
# 4. HTTP Request Handler
# ---------------------------------------------------------------------------
class EquityServerHandler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass  # Quiet logging

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        query = urllib.parse.parse_qs(parsed.query)
        
        if path == "/":
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(HTML_TEMPLATE.encode("utf-8"))
            return
            
        elif path == "/api/stocks":
            self.send_json({"total": len(ALL_SYMBOLS), "symbols": ALL_SYMBOLS})
            return
            
        elif path.startswith("/api/equity/"):
            sym = path.replace("/api/equity/", "").strip().upper()
            date_param = query.get("date", [None])[0]
            data = load_equity_spot_data(sym, date_param)
            self.send_json(data)
            return
            
        elif path.startswith("/api/fundamentals/"):
            sym = path.replace("/api/fundamentals/", "").strip().upper()
            data = load_fundamentals_snapshot(sym)
            self.send_json(data)
            return
            
        self.send_error(404, "Not Found")

    def send_json(self, payload: dict):
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(json.dumps(payload).encode("utf-8"))

# ---------------------------------------------------------------------------
# 5. Front-End HTML5/CSS3/JS Web UI
# ---------------------------------------------------------------------------
HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>211 Equity Stocks Browser & Live Charting</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<style>
  :root {
    --bg-main: #0f172a;
    --bg-card: #1e293b;
    --accent: #2563eb;
    --accent-hover: #1d4ed8;
    --text-main: #f8fafc;
    --text-muted: #94a3b8;
    --border: #334155;
    --green: #22c55e;
    --red: #ef4444;
  }
  body {
    margin: 0; padding: 0;
    font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
    background: var(--bg-main); color: var(--text-main);
  }
  header {
    background: #090d16; padding: 15px 25px; display: flex; align-items: center;
    justify-content: space-between; border-bottom: 1px solid var(--border);
  }
  h1 { font-size: 1.3rem; margin: 0; color: #60a5fa; }
  .badge { background: var(--accent); color: white; padding: 4px 10px; border-radius: 12px; font-size: 0.8rem; font-weight: bold; }
  
  .container { display: flex; height: calc(100vh - 65px); }
  
  /* Sidebar Stock List */
  .sidebar { width: 280px; background: #111827; border-right: 1px solid var(--border); display: flex; flex-direction: column; }
  .search-box { padding: 12px; border-bottom: 1px solid var(--border); }
  .search-box input {
    width: 90%; background: #1f2937; border: 1px solid var(--border); color: white;
    padding: 8px 12px; border-radius: 6px; outline: none; font-size: 0.9rem;
  }
  .stock-list { flex: 1; overflow-y: auto; list-style: none; margin: 0; padding: 0; }
  .stock-item {
    padding: 12px 18px; border-bottom: 1px solid #1f2937; cursor: pointer;
    font-weight: 500; font-size: 0.95rem; transition: background 0.15s;
  }
  .stock-item:hover { background: #1f2937; }
  .stock-item.active { background: var(--accent); color: white; }
  
  /* Main Content Area */
  .main-panel { flex: 1; padding: 20px; overflow-y: auto; display: flex; flex-direction: column; gap: 20px; }
  .control-bar { display: flex; align-items: center; justify-content: space-between; background: var(--bg-card); padding: 15px 20px; border-radius: 8px; border: 1px solid var(--border); }
  .stock-title { font-size: 1.4rem; font-weight: bold; color: white; }
  select { background: #0f172a; color: white; border: 1px solid var(--border); padding: 6px 12px; border-radius: 6px; outline: none; }
  
  .chart-container { background: var(--bg-card); padding: 20px; border-radius: 8px; border: 1px solid var(--border); height: 420px; position: relative; }
  
  .stats-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; }
  .stat-card { background: var(--bg-card); padding: 15px; border-radius: 8px; border: 1px solid var(--border); text-align: center; }
  .stat-val { font-size: 1.2rem; font-weight: bold; margin-top: 5px; color: #38bdf8; }
  .stat-lbl { font-size: 0.8rem; color: var(--text-muted); }
</style>
</head>
<body>

<header>
  <h1>📈 211 EQUITY STOCKS INTRADAY BROWSER</h1>
  <span class="badge" id="totalCount">211 STOCKS AVAILABLE</span>
</header>

<div class="container">
  <div class="sidebar">
    <div class="search-box">
      <input type="text" id="searchInput" placeholder="Search 211 stocks..." oninput="filterStocks()">
    </div>
    <ul class="stock-list" id="stockList"></ul>
  </div>
  
  <div class="main-panel">
    <div class="control-bar">
      <div>
        <span class="stock-title" id="selectedSymbol">RELIANCE</span>
        <span style="color: var(--text-muted); margin-left: 10px;" id="dateCount"></span>
      </div>
      <div>
        <label for="dateSelect" style="color: var(--text-muted); font-size: 0.9rem;">Trading Date: </label>
        <select id="dateSelect" onchange="onDateChange()"></select>
      </div>
    </div>
    
    <div class="stats-row">
      <div class="stat-card"><div class="stat-lbl">DAY OPEN PRICE</div><div class="stat-val" id="valOpen">—</div></div>
      <div class="stat-card"><div class="stat-lbl">DAY HIGH PRICE</div><div class="stat-val" id="valHigh" style="color: var(--green);">—</div></div>
      <div class="stat-card"><div class="stat-lbl">DAY LOW PRICE</div><div class="stat-val" id="valLow" style="color: var(--red);">—</div></div>
      <div class="stat-card"><div class="stat-lbl">DAY CLOSE / LAST PRICE</div><div class="stat-val" id="valClose">—</div></div>
    </div>
    
    <div class="chart-container">
      <canvas id="equityChart"></canvas>
    </div>
  </div>
</div>

<script>
let allSymbols = [];
let currentSymbol = "RELIANCE";
let myChart = null;

async function init() {
  const res = await fetch('/api/stocks');
  const data = await res.json();
  allSymbols = data.symbols;
  document.getElementById('totalCount').innerText = `${allSymbols.length} STOCKS AVAILABLE`;
  renderStockList(allSymbols);
  if(allSymbols.length > 0) {
    selectStock(allSymbols[0]);
  }
}

function renderStockList(list) {
  const ul = document.getElementById('stockList');
  ul.innerHTML = '';
  list.forEach(sym => {
    const li = document.createElement('li');
    li.className = 'stock-item' + (sym === currentSymbol ? ' active' : '');
    li.innerText = sym;
    li.onclick = () => selectStock(sym);
    ul.appendChild(li);
  });
}

function filterStocks() {
  const q = document.getElementById('searchInput').value.toUpperCase();
  const filtered = allSymbols.filter(s => s.includes(q));
  renderStockList(filtered);
}

async function selectStock(sym) {
  currentSymbol = sym;
  document.getElementById('selectedSymbol').innerText = sym;
  renderStockList(allSymbols.filter(s => s.includes(document.getElementById('searchInput').value.toUpperCase())));
  await loadStockData(sym);
}

async function loadStockData(sym, dateStr = null) {
  let url = `/api/equity/${sym}`;
  if (dateStr) url += `?date=${dateStr}`;
  
  const res = await fetch(url);
  const data = await res.json();
  
  document.getElementById('dateCount').innerText = `(${data.total_dates} Trading Days Available)`;
  
  const sel = document.getElementById('dateSelect');
  sel.innerHTML = '';
  if (data.dates) {
    data.dates.forEach(d => {
      const opt = document.createElement('option');
      opt.value = d; opt.innerText = d;
      if (d === data.selected_date) opt.selected = true;
      sel.appendChild(opt);
    });
  }
  
  renderChartAndCards(data.bars);
}

function onDateChange() {
  const dateStr = document.getElementById('dateSelect').value;
  loadStockData(currentSymbol, dateStr);
}

function renderChartAndCards(bars) {
  if (!bars || bars.length === 0) {
    document.getElementById('valOpen').innerText = '—';
    document.getElementById('valHigh').innerText = '—';
    document.getElementById('valLow').innerText = '—';
    document.getElementById('valClose').innerText = '—';
    if(myChart) myChart.destroy();
    return;
  }
  
  const openVal = bars[0].open;
  const closeVal = bars[bars.length - 1].close;
  const highVal = Math.max(...bars.map(b => b.high));
  const lowVal = Math.min(...bars.map(b => b.low));
  
  document.getElementById('valOpen').innerText = `₹${openVal.toFixed(2)}`;
  document.getElementById('valHigh').innerText = `₹${highVal.toFixed(2)}`;
  document.getElementById('valLow').innerText = `₹${lowVal.toFixed(2)}`;
  document.getElementById('valClose').innerText = `₹${closeVal.toFixed(2)}`;
  
  const labels = bars.map(b => b.time);
  const prices = bars.map(b => b.close);
  const volumes = bars.map(b => b.volume);
  
  if (myChart) myChart.destroy();
  
  const ctx = document.getElementById('equityChart').getContext('2d');
  myChart = new Chart(ctx, {
    type: 'line',
    data: {
      labels: labels,
      datasets: [{
        label: '1-Min Spot Price (₹)',
        data: prices,
        borderColor: '#38bdf8',
        backgroundColor: 'rgba(56, 189, 248, 0.1)',
        fill: true,
        tension: 0.1,
        pointRadius: 0
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        x: { grid: { color: '#1e293b' }, ticks: { color: '#94a3b8', maxTicksLimit: 12 } },
        y: { grid: { color: '#334155' }, ticks: { color: '#f8fafc' } }
      },
      plugins: {
        legend: { labels: { color: '#f8fafc' } }
      }
    }
  });
}

window.onload = init;
</script>
</body>
</html>
"""

# ---------------------------------------------------------------------------
# 6. Server Main Execution
# ---------------------------------------------------------------------------
def lan_ip() -> str:
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))
        return s.getsockname()[0]
    except Exception:
        return "127.0.0.1"
    finally:
        s.close()

def main():
    port = next((int(a) for a in sys.argv[1:] if a.isdigit()), 8000)
    lan = "--lan" in sys.argv
    
    host = "0.0.0.0" if lan else "127.0.0.1"
    ip = lan_ip() if lan else "localhost"
    url = f"http://{ip}:{port}"
    
    server = ThreadingHTTPServer((host, port), EquityServerHandler)
    print("=" * 70)
    print(f"  📈 211 EQUITY STOCKS BROWSER & SERVER")
    print(f"  Serving {len(ALL_SYMBOLS):,} Equity Stock Symbols from D:\\behaviour analysis\\OI_DATA")
    print(f"  Live Server URL: {url}" + ("   (Reachable on LAN)" if lan else ""))
    print("  Press Ctrl+C to stop.")
    print("=" * 70)
    
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nServer Stopped.")
        server.shutdown()

if __name__ == "__main__":
    main()
