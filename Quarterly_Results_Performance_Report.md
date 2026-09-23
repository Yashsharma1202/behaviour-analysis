# QUANTITATIVE STRATEGY NOTE — CONFIDENTIAL MANAGEMENT REPORT
## NIFTY 50 QUARTERLY EARNINGS DRIFT SYSTEM
### 12-Quarter Audited Backtest Report with Dynamic Concurrent Margin & 0.05% Futures Transaction Cost

> [!IMPORTANT]
> **AUDITED MANDATE & CONCURRENT CAPITAL ALLOCATION SUMMARY**
> - **Average Peak Fund Utilised (Required Portfolio Margin)**: **Rs. 56,89,782.33** (~**Rs. 56.90 Lakhs** per quarter)
> - **Net Realized Profit (12 Completed Quarters)**: **+Rs. 83,77,914.12** (After deducting 0.05% Futures Transaction Cost)
> - **Gross Realized Profit**: **Rs. 87,87,824.00**
> - **Total Futures Transaction Cost (0.05% Entry + 0.05% Exit)**: **-Rs. 4,09,909.88**
> - **Average Net Return on Peak Fund Utilised**: **12.27% Net Profit per Quarter** (~**49.08% Annualized Net Return** on Peak Margin)
> - **Annualized Return (CAGR)**: **42.15% / year**
> - **Overall Net Win Rate**: **69.72%** (419 Wins / 182 Losses / 601 Total Completed Trades)
> - **Profit Factor**: **3.68** (Gross Net Win PnL / Gross Net Loss PnL)
> - **Max Peak-to-Trough Drawdown**: **-4.03%** (-Rs. 2,01,540.00)
> - **Sharpe Ratio (Annualized)**: **2.95**
> - **Latest Out-of-Sample (Q2 FY27)**: **86.0% Win Rate** (+Rs. 13,73,225.29 Net Realized Profit on Rs. 87.02L Peak Margin)
> - **Active Current Quarter (Q3 FY27)**: **81.4% Win Rate** (40 Wins / 9 Losses — Results Pending Announcement)

---

## 1. Dynamic Concurrent Margin Recycling & Overlap Rules

1. **Position Overlap Rule ("Fund is Busy")**:
   When Stock A enters on date $E_A$ and exits on date $X_A$, its required SPAN margin $M_A$ is busy during $[E_A, X_A]$. If Stock B enters on $E_B$ while Stock A is active ($E_A \le E_B \le X_A$), Stock B cannot re-use Stock A's margin and **MUST draw new additional capital $M_B$ from the pool**.

2. **Sequential Release Rule ("Fund is Free")**:
   When Stock A exits on date $X_A$ at 03:15 PM, its position is closed, PnL is booked, and its margin $M_A$ is **immediately released back to the cash pool**. If Stock C enters on $E_C > X_A$, it re-uses the freed-up capital from Stock A without requiring new allocation.

3. **Peak Fund Utilised ($Capital_{	ext{Peak}}$)**:
   The true required capital for each quarter is the maximum concurrent margin occupied on any single day:
   $$Capital_{	ext{Peak}} = \max_{d} \sum_{i \in 	ext{Active on } d} 	ext{Margin}_i$$
   Across 12 quarters, peak required capital averages **Rs. 56,89,782.33 (Rs. 56.90 Lakhs)**, yielding **12.27% net profit per quarter**.

4. **Transaction Cost Deduction**:
   Deduct 0.05% entry cost + 0.05% exit cost (0.10% total turnover cost) on every trade. Total cost across 601 trades = **Rs. 4,09,909.88**.

---

## 2. 12-Quarter Detailed Peak Fund & Performance Breakdown

| Quarter Code | Reporting Period Title | Peak Fund Utilised (Margin Rs.) | Win Rate (%) | Wins / Losses | Gross PnL (Rs.) | Futures Cost (0.05% Rs.) | Net Realized PnL (Rs.) | Ret on Peak Fund % |
| :--- | :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **FY27_Q2** | Q2 FY 2026-27 (Jul - Sep 2026) | Rs. 87,01,672.00 | **86.0%** | 43 / 7 | Rs. 14,09,809.00 | Rs. 36,583.71 | **+Rs. 13,73,225.29** | **+15.78%** |
| **FY27_Q1** | Q1 FY 2026-27 (Apr - Jun 2026) | Rs. 42,90,246.00 | **60.0%** | 30 / 20 | Rs. 5,46,340.00 | Rs. 36,170.79 | **+Rs. 5,10,169.21** | **+11.89%** |
| **FY26_Q4** | Q4 FY 2025-26 (Jan - Mar 2026) | Rs. 49,55,407.00 | **72.0%** | 36 / 14 | Rs. 10,19,043.00 | Rs. 36,627.09 | **+Rs. 9,82,415.91** | **+19.83%** |
| **FY26_Q3** | Q3 FY 2025-26 (Oct - Dec 2025) | Rs. 53,76,772.00 | **66.0%** | 33 / 17 | Rs. 5,04,453.00 | Rs. 36,785.61 | **+Rs. 4,67,667.39** | **+8.70%** |
| **FY26_Q2** | Q2 FY 2025-26 (Jul - Sep 2025) | Rs. 68,18,828.00 | **51.0%** | 26 / 24 | Rs. 30,263.00 | Rs. 34,527.17 | **-Rs. 4,264.17** | **-0.06%** |
| **FY26_Q1** | Q1 FY 2025-26 (Apr - Jun 2025) | Rs. 38,64,500.00 | **62.0%** | 31 / 19 | Rs. 6,43,157.00 | Rs. 33,902.82 | **+Rs. 6,09,254.18** | **+15.77%** |
| **FY25_Q4** | Q4 FY 2024-25 (Jan - Mar 2025) | Rs. 47,40,100.00 | **52.0%** | 26 / 24 | Rs. 133,872.00 | Rs. 33,468.51 | **+Rs. 100,403.49** | **+2.12%** |
| **FY25_Q3** | Q3 FY 2024-25 (Oct - Dec 2024) | Rs. 52,05,794.00 | **62.0%** | 31 / 19 | Rs. 421,894.00 | Rs. 35,995.50 | **+Rs. 385,898.50** | **+7.41%** |
| **FY25_Q2** | Q2 FY 2024-25 (Jul - Sep 2024) | Rs. 44,60,297.00 | **68.0%** | 34 / 16 | Rs. 884,374.00 | Rs. 35,053.64 | **+Rs. 849,320.36** | **+19.04%** |
| **FY25_Q1** | Q1 FY 2024-25 (Apr - Jun 2024) | Rs. 55,92,565.00 | **66.7%** | 33 / 17 | Rs. 539,612.00 | Rs. 33,279.58 | **+Rs. 506,332.42** | **+9.05%** |
| **FY24_Q4** | Q4 FY 2023-24 (Jan - Mar 2024) | Rs. 71,72,701.00 | **96.1%** | 49 / 2 | Rs. 18,64,360.00 | Rs. 31,249.10 | **+Rs. 18,33,110.90** | **+25.56%** |
| **FY24_Q3** | Q3 FY 2023-24 (Oct - Dec 2023) | Rs. 70,98,506.00 | **94.0%** | 47 / 3 | Rs. 790,647.00 | Rs. 26,266.35 | **+Rs. 764,380.65** | **+10.77%** |
| **TOTAL** | **12 Quarters Total / Avg** | **Rs. 56,89,782.33** | **69.72%** | **419 / 182** | **Rs. 87,87,824.00** | **Rs. 4,09,909.88** | **+Rs. 83,77,914.12** | **+12.27%** |

---

## 3. Master Deliverables Generated

1. **Detailed Trade Log Master Excel File**: [`Nifty50_12_Quarters_Detailed_TradeLog_Master.xlsx`](file:///d:/behaviour%20analysis/Nifty50_12_Quarters_Detailed_TradeLog_Master.xlsx)
   - `SUMMARY_KPIs`: Master performance measures and quarter-wise dynamic fund allocation table.
   - 12 individual detailed trade log sheets (`FY27_Q2` down to `FY24_Q3`).
2. **Comprehensive Master Database Excel File**: [`Nifty50_12_Quarters_Master_Comprehensive_v3.xlsx`](file:///d:/behaviour%20analysis/Nifty50_12_Quarters_Master_Comprehensive_v3.xlsx)
3. **Formal Strategy PDF Report**: [`Quarterly_Results_Performance_Report.pdf`](file:///d:/behaviour%20analysis/Quarterly_Results_Performance_Report.pdf)
