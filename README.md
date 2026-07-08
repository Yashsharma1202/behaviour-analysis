# Behaviour Analysis — NSE Stock Browser & Corporate-Event Toolkit

Browse the fundamentals of **2,363 NSE stocks** one at a time, and pull each
company's **NSE event feeds** (announcements, board meetings, corporate actions,
financial results) on demand — in a clean, dark "trading-terminal" web UI.

## 🔗 Live demo (static snapshot)

**https://Yashsharma1202.github.io/behaviour-analysis/**

The live site is a static snapshot hosted on GitHub Pages: it has fundamentals
for **every** stock and event feeds for the stocks already downloaded. Live NSE
downloading only works when you run the app locally (GitHub Pages can't run a
server or reach NSE).

## Run it locally (full features)

```bash
pip install -r requirements.txt
python stock_server.py            # opens http://localhost:8000
```

Then:
- **Search** any of 2,363 symbols (type any part of the name).
- **← / →** to step through stocks one at a time.
- **Fundamentals** tab: sales, net profit, OPM, ROCE + annual & quarterly charts.
- **Announcements / Board Meetings / Corporate Actions / Financial Results** tabs:
  click **⬇ Download NSE feeds** to fetch that stock's data live from NSE (cached
  to a `<SYMBOL>/` folder). A green **✓ data** tag marks stocks already downloaded.

> The per-stock fundamentals CSVs (`pnl/`, `quarterly/`, `ratios/`,
> `balance_sheet/`, `cash_flow/`) are your own local dataset and are **not** in
> this repo (≈59 MB). Point the app at those folders to browse fundamentals
> locally; the published static site has the baked-down JSON instead.

## Rebuild the static site

```bash
python build_static.py            # regenerates docs/ from your local data
python -m http.server -d docs 9000   # preview at http://localhost:9000
```

## What's in here

| File | Purpose |
|------|---------|
| `stock_server.py` | The web dashboard (stdlib HTTP server) + live NSE downloader |
| `stock_browser_gui.py` | Desktop (Tkinter) version of the browser |
| `build_static.py` | Bakes the dashboard into `docs/` for GitHub Pages |
| `fund_loader.py` | Parses the Screener-style fundamental CSVs for any stock |
| `clean_events.py`, `event_study.py`, `dividend_recovery.py`, `earnings_surprise.py`, `backtest.py`, `playbook.py`, `ml_model.py`, `xbrl_parser.py`, `event_records.py`, `dashboard_gui.py` | RELIANCE corporate-event behaviour-analysis pipeline |
| `straddle_backtest.py`, `inspect_csv.py` | NIFTY intraday long-straddle options backtest |

Data source: [NSE India](https://www.nseindia.com/) public corporate filings.
Educational / research use.
