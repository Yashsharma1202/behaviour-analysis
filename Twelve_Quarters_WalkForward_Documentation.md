# 12-Quarters Walk-Forward Backtest System Documentation

This document explains the methodology, data pipeline, and logic used in the **12-Quarters Walk-Forward Backtest System** we have built. 

## 1. Objective
The goal of this system is to dynamically test our Event-Based Trading Strategy (Long vs. Short 8x8 Grid Optimization) across a continuous 12-quarter rolling window. Instead of optimizing on the entire dataset at once (which causes look-ahead bias), we use a **Walk-Forward** approach. This means for any given quarter we want to trade in, the system *only* trains on the historical earnings events that occurred *prior* to that quarter, and then tests those optimal rules on the unseen upcoming quarter.

## 2. Financial Year Mapping & Announcement Quarters
To accurately track the Indian corporate earnings cycle, we mapped the quarters according to the Indian Financial Year (FY), which begins in **April** and ends in **March**. 

Because earnings results are announced in the quarter *following* the actual business quarter, our 12 sequential quarters represent the **Announcement Windows**:
*   **Q1 Announcements (Jul - Sep):** Results for the Apr-Jun business quarter.
*   **Q2 Announcements (Oct - Dec):** Results for the Jul-Sep business quarter.
*   **Q3 Announcements (Jan - Mar):** Results for the Oct-Dec business quarter.
*   **Q4 Announcements (Apr - Jun):** Results for the Jan-Mar business quarter.

The 12 continuous quarters tested in this simulation run chronologically:
1.  **Q3 2023-24** (Oct - Dec 2023)
2.  **Q4 2023-24** (Jan - Mar 2024)
3.  **Q1 2024-25** (Apr - Jun 2024)
4.  **Q2 2024-25** (Jul - Sep 2024)
5.  **Q3 2024-25** (Oct - Dec 2024)
6.  **Q4 2024-25** (Jan - Mar 2025)
7.  **Q1 2025-26** (Apr - Jun 2025)
8.  **Q2 2025-26** (Jul - Sep 2025)
9.  **Q3 2025-26** (Oct - Dec 2025)
10. **Q4 2025-26** (Jan - Mar 2026)
11. **Q1 2026-27** (Apr - Jun 2026)
12. **Q2 2026-27** (Jul - Sep 2026) *(Future)*

## 3. Data Ingestion Pipeline
To build a seamless historical timeline spanning over 3 years, the system dynamically merges dates from two separate sources:
*   **Historical Database (CSVs):** Fetches all earnings announcement dates prior to July 2024 directly from the raw `financial_results.csv` files stored in individual stock folders.
*   **Peer Model Workbooks:** Fetches recent and upcoming dates (July 2024 onwards) from the manual `Nifty50_Peer_Model_8Q_ALL50.xlsx` sheets.
*   **Live Pricing:** Automatically downloads historical daily close prices for all 50 stocks from Yahoo Finance starting from January 1, 2022, to present.

## 4. Optimization & Trading Logic (The Walk-Forward Step)
For every quarter (e.g., Q1 2025-26), the system performs the following independent steps for all 50 stocks:

1.  **Training (In-Sample):** 
    *   It filters out all event dates strictly *before* the start of the quarter.
    *   It tests a symmetric **8x8 Grid** (1 to 8 days before, 1 to 8 days after) for both **LONG** and **SHORT** strategies independently.
    *   It records the parameter combination that yields the highest historical win-rate and average expected return.
2.  **Prediction / Testing (Out-of-Sample):**
    *   It compares the best historical expected return for LONG vs. SHORT.
    *   It selects the strategy (LONG or SHORT) with the highest expected yield.
    *   It then strictly applies these optimal entry/exit days to the actual earnings event *inside* the current quarter to generate a real trade.
3.  **Simulation & Capital Slots:**
    *   All out-of-sample trades generated across the 50 stocks for that quarter are chronologically sorted.
    *   The system simulates parallel Capital Slots. If a new trade starts while the previous trade is still holding capital, it opens a new Slot. If a previous slot has exited, the capital is re-used, minimizing total capital blocked.

## 5. Output Reports
The script generates a separate, fully formatted Excel workbook for each of the 12 quarters in the `12_Quarters_Reports/` directory.

Each workbook contains 3 sheets:
*   **Summary:** Provides a high-level view of Portfolio Win Rate, Net Profit, Return on Capital, and Total Concurrent Capital Slots Utilized during that quarter.
*   **Detailed_Trades:** A chronological list of every trade taken in that quarter. It highlights whether the model dynamically picked LONG or SHORT, the entry/exit dates, the expected return based on training, and the actual realized return and profit in Rupees.
*   **Losing Trades Verification:** An audit sheet for any trade that hit a loss. It performs a side-by-side comparison proving that, based purely on historical probabilities, the model picked the mathematically superior direction (i.e. proving that if it had picked the opposite direction, the loss would have likely been worse or the strategy was entirely out of bounds).
