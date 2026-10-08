DATA EVENT — consolidated copy of all data worked on this session
====================================================================
Created: 2026-09-30

IMPORTANT: everything in this folder is a COPY. The originals are untouched
and remain the live source of truth in their original locations:
  - dashboard_data/      -> still feeds the live dashboard (local + GitHub Pages host)
  - scraped_parquet/     -> still the source for every backtest/report script
  - the report files     -> still sit in the project root as before

Do NOT edit files inside "DATA EVENT" expecting them to affect the live
dashboard or re-run any script — they are a snapshot for easy reference only.
If you need to regenerate anything, use the original build_*.py scripts in
the project root, which read/write the ORIGINAL locations, then re-copy here
if you want this snapshot refreshed.

FOLDER CONTENTS
---------------
dashboard_data/   (17 files)
  The live "event data" the dashboard reads: holidays, RBI MPC policy,
  quarterly earnings, F&O universe (211 stocks), the new F&O October
  seasonality live-execution file, holiday direction backtest, EOD returns,
  economics backtest, etc. (.json = data, .js = window-global bundles).

scraped_parquet/  (10 top-level files + nifty50_stocks/ subfolder)
  Raw price history used to build every backtest in this project:
    - nifty50_all_stocks_daily_2000_2026.parquet   (50 Nifty stocks, 2000-2026, cash EQ)
    - nifty50_stocks/                               (same data, split per symbol)
    - fo_futures_daily_all_contracts.parquet        (211 F&O stocks + 5 indices, all futures contracts, 2019-2026)
    - fo_futures_near_month_continuous.parquet      (same, stitched to one continuous series)
    - fo_futures_coverage.csv                       (per-symbol history coverage map)
    - nifty_futures_daily_all_contracts.parquet / nifty_futures_near_month_continuous.parquet (Nifty index futures)
    - nifty_spot_daily.parquet                      (Nifty index spot, 2010-2026)
    - rbi_monetary_policy_dates.parquet / rbi_policy_market_reaction_2000_2026.parquet (162 MPC events, 2000-2026)

reports/          (23 files)
  Every Excel/HTML/PDF report generated this session:
    - Quarterly futures & 1% ITM options masters (12 quarters), and their
      comparisons (Futures vs 1% ITM vs ATM, Rs 10L/Rs 20L performance, trade logs)
    - RBI MPC index playbook + top-15 stock comparative report
    - Holiday performance reports (Gandhi Jayanti, Dussehra, RBI Policy)
    - Gandhi missed-positions salvage report
    - Backtest economics report
    - F&O October seasonality report

For the exact backtest years / methodology behind each dataset, see the
Overview/Assumptions sheet inside each report, or the corresponding
build_*.py script in the project root.
