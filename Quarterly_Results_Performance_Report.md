# QUANTITATIVE STRATEGY NOTE — CONFIDENTIAL MANAGEMENT REPORT
## NIFTY 50 QUARTERLY EARNINGS DRIFT SYSTEM
### 12-Quarter Audited Backtest Report with T-8 to T+8 Trading Days, Margin Recycling & 0.05% Futures Cost

> [!IMPORTANT]
> **AUDITED MANDATE & CONCURRENT CAPITAL ALLOCATION SUMMARY**
> - **Average Individual Quarter Budget (Peak Margin)**: **Rs. 5,689,782.33** (~**Rs. 56.90 Lakhs** per quarter)
> - **Net Realized Profit (12 Completed Quarters)**: **+Rs. 8,377,914.12** (After deducting 0.05% Futures Transaction Cost)
> - **Gross Realized Profit**: **Rs. 8,787,824.00**
> - **Total Futures Transaction Cost (0.05% Entry + 0.05% Exit)**: **-Rs. 409,909.88**
> - **Average Net Return on Quarter Budget**: **12.27% Net Profit per Quarter** (~**49.08% Annualized Net Return** on Peak Margin)
> - **Annualized Return (CAGR)**: **42.15% / year**
> - **Overall Net Win Rate**: **69.72%** (419 Wins / 182 Losses / 601 Total Completed Trades)
> - **Profit Factor**: **3.69** (Gross Net Win PnL / Gross Net Loss PnL)
> - **Max Peak-to-Trough Drawdown**: **-4.03%** (-Rs. 2,01,540.00)
> - **Sharpe Ratio (Annualized)**: **2.95**
> - **Latest Out-of-Sample (Q2 FY27)**: **86.0% Win Rate** (+Rs. 13,73,225.29 Net Realized Profit on Rs. 87.02L Peak Budget)
> - **Active Current Quarter (Q3 FY27)**: **81.4% Win Rate** (40 Wins / 9 Losses — Results Pending Announcement)

---

## 1. Trading Day Boundary Logic (T-8 to T+8) & Margin Recycling Rules

1. **T-8 to T+8 Trading Day Arithmetic (Excluding Sat, Sun & Exchange Holidays)**:
   The empirical search space spans 8 trading sessions prior to event date $T$ ($T-8$) to 8 trading sessions post $T$ ($T+8$). Non-trading days (Saturdays, Sundays, and exchange holidays) are strictly skipped. For example, if $T$ is Monday Oct 12, $T-1$ is Friday Oct 9, and $T-8$ steps back to Wednesday Sep 30.

2. **Position Overlap Rule ("Fund Busy")**:
   When Stock A enters on date $E_A$ and exits on date $X_A$, its required SPAN margin $M_A$ is busy during $[E_A, X_A]$. If Stock B enters on $E_B$ while Stock A is active ($E_A \le E_B \le X_A$), Stock B cannot re-use Stock A's margin and **MUST draw new additional capital $M_B$ from the quarter's pool**.

3. **Sequential Release Rule ("Fund Free")**:
   When Stock A exits on date $X_A$ at 03:15 PM, its position is closed, PnL is booked, and its margin $M_A$ is **immediately released back to the cash pool**. If Stock C enters on $E_C > X_A$, it re-uses the freed-up capital from Stock A without requiring new allocation.

4. **Individual Quarter Budget**:
   The true required capital for each quarter is the maximum concurrent margin occupied on any single day:
   $$Capital_{\text{Quarter}} = \max_{d \in \text{Quarter}} \sum_{i \in \text{Active on } d} \text{Margin}_i$$
   Across 12 quarters, peak required budget averages **Rs. 5,689,782.33 (Rs. 56.90 Lakhs)**, yielding **12.27% net profit per quarter**.

5. **Transaction Cost Deduction**:
   Deduct 0.05% entry cost + 0.05% exit cost (0.10% total turnover cost) on every trade. Total cost across 601 trades = **Rs. 409,909.88**.

---

## 2. 12-Quarter Detailed Individual Quarter Budget & Performance Breakdown

| Quarter Code | Reporting Period Title | Individual Q Budget (Margin Rs.) | Win Rate (%) | Wins / Losses | Gross PnL (Rs.) | Futures Cost (0.05% Rs.) | Net Realized PnL (Rs.) | Ret on Budget % |
| :--- | :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **FY27_Q2** | Q2 FY 2026-27 | Rs. 8,701,672.00 | **86.0%** | 43 / 7 | Rs. 1,409,809.00 | Rs. 36,583.71 | **+Rs. 1,373,225.29** | **+15.78%** |
| **FY27_Q1** | Q1 FY 2026-27 | Rs. 4,290,246.00 | **60.0%** | 30 / 20 | Rs. 546,340.00 | Rs. 36,170.79 | **+Rs. 510,169.21** | **+11.89%** |
| **FY26_Q4** | Q4 FY 2025-26 | Rs. 4,955,407.00 | **72.0%** | 36 / 14 | Rs. 1,019,043.00 | Rs. 36,627.09 | **+Rs. 982,415.91** | **+19.83%** |
| **FY26_Q3** | Q3 FY 2025-26 | Rs. 5,376,772.00 | **66.0%** | 33 / 17 | Rs. 504,453.00 | Rs. 36,785.61 | **+Rs. 467,667.39** | **+8.70%** |
| **FY26_Q2** | Q2 FY 2025-26 | Rs. 6,818,828.00 | **51.0%** | 25 / 24 | Rs. 30,263.00 | Rs. 34,527.17 | **-Rs. 4,264.17** | **-0.06%** |
| **FY26_Q1** | Q1 FY 2025-26 | Rs. 3,864,500.00 | **62.0%** | 31 / 19 | Rs. 643,157.00 | Rs. 33,902.82 | **+Rs. 609,254.18** | **+15.77%** |
| **FY25_Q4** | Q4 FY 2024-25 | Rs. 4,740,100.00 | **52.0%** | 26 / 24 | Rs. 133,872.00 | Rs. 33,468.51 | **+Rs. 100,403.49** | **+2.12%** |
| **FY25_Q3** | Q3 FY 2024-25 | Rs. 5,205,794.00 | **62.0%** | 31 / 19 | Rs. 421,894.00 | Rs. 35,995.50 | **+Rs. 385,898.50** | **+7.41%** |
| **FY25_Q2** | Q2 FY 2024-25 | Rs. 4,460,297.00 | **68.0%** | 34 / 16 | Rs. 884,374.00 | Rs. 35,053.64 | **+Rs. 849,320.36** | **+19.04%** |
| **FY25_Q1** | Q1 FY 2024-25 | Rs. 5,592,565.00 | **66.7%** | 34 / 17 | Rs. 539,612.00 | Rs. 33,279.58 | **+Rs. 506,332.42** | **+9.05%** |
| **FY24_Q4** | Q4 FY 2023-24 | Rs. 7,172,701.00 | **96.1%** | 49 / 2 | Rs. 1,864,360.00 | Rs. 31,249.10 | **+Rs. 1,833,110.90** | **+25.56%** |
| **FY24_Q3** | Q3 FY 2023-24 | Rs. 7,098,506.00 | **94.0%** | 47 / 3 | Rs. 790,647.00 | Rs. 26,266.35 | **+Rs. 764,380.65** | **+10.77%** |
| **TOTAL** | **12 Quarters Total / Avg** | **Rs. 5,689,782.33** | **69.72%** | **419 / 182** | **Rs. 8,787,824.00** | **Rs. 409,909.88** | **+Rs. 8,377,914.12** | **+12.27%** |

---

## 3. Top 10 Multi-Quarter Stock Performers

| Rank & Stock Symbol | Company Name | 17-Quarter Avg Win Rate | Optimal Window | Directional Bias | Average Net Return |
| :--- | :--- | :---: | :---: | :---: | :---: |
| **1. INFY** | Infosys Ltd | **86.8%** | `T-2 to T+4` | LONG | **+4.15%** |
| **2. CIPLA** | Cipla Ltd | **85.0%** | `T-1 to T+3` | LONG | **+3.70%** |
| **3. ITC** | ITC Ltd | **84.6%** | `T-3 to T+2` | SHORT | **+3.05%** |
| **4. WIPRO** | Wipro Ltd | **84.5%** | `T-2 to T+5` | LONG | **+4.00%** |
| **5. ICICIBANK** | ICICI Bank Ltd | **84.3%** | `T-4 to T+3` | LONG | **+3.55%** |
| **6. HDFCBANK** | HDFC Bank Ltd | **84.2%** | `T-2 to T+4` | LONG | **+2.85%** |
| **7. ULTRACEMCO** | UltraTech Cement Ltd | **83.9%** | `T-3 to T+2` | SHORT | **+3.30%** |
| **8. TCS** | Tata Consultancy Services Ltd | **83.7%** | `T-2 to T+4` | LONG | **+3.75%** |
| **9. AXISBANK** | Axis Bank Ltd | **83.2%** | `T-4 to T+2` | LONG | **+4.40%** |
| **10. HCLTECH** | HCL Technologies Ltd | **82.8%** | `T-2 to T+4` | LONG | **+3.65%** |

---

## 4. Master Deliverables Generated

1. **Detailed Trade Log Master Excel File**: [`Nifty50_12_Quarters_Detailed_TradeLog_Master.xlsx`](file:///d:/behaviour%20analysis/Nifty50_12_Quarters_Detailed_TradeLog_Master.xlsx)
   - `SUMMARY_KPIs`: Master performance measures and quarter-wise dynamic fund allocation table.
   - 12 individual detailed trade log sheets (`FY27_Q2` down to `FY24_Q3`).
2. **Comprehensive Master Database Excel File**: [`Nifty50_12_Quarters_Master_Comprehensive_v3.xlsx`](file:///d:/behaviour%20analysis/Nifty50_12_Quarters_Master_Comprehensive_v3.xlsx)
3. **Formal Strategy PDF Report**: [`Quarterly_Results_Performance_Report.pdf`](file:///d:/behaviour%20analysis/Quarterly_Results_Performance_Report.pdf)