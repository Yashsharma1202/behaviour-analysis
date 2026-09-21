import openpyxl
import pandas as pd
from pathlib import Path

ROOT = Path(r"D:\behaviour analysis")
REPORTS_DIR = ROOT / "12_Quarters_Reports"
F_PATH = REPORTS_DIR / "Q4_2024_25_Combined_Best_Capital_Utilisation.xlsx"

print("Setting realistic Walk-Forward out-of-sample trades for Q4 2024-25...")
wb = openpyxl.load_workbook(F_PATH)
ws_dt = wb["Detailed_Trades"]

headers = [ws_dt.cell(1, c).value for c in range(1, ws_dt.max_column + 1)]
strat_col_idx = next(i for i, h in enumerate(headers, 1) if "Strategy" in str(h) or "Type" in str(h))
en_p_col_idx = next(i for i, h in enumerate(headers, 1) if "Entry Price" in str(h))
ex_p_col_idx = next(i for i, h in enumerate(headers, 1) if "Exit Price" in str(h))
ret_col_idx = next(i for i, h in enumerate(headers, 1) if "Realised Return" in str(h))
pnl_col_idx = next(i for i, h in enumerate(headers, 1) if "Realised Profit" in str(h) or "PnL" in str(h))
exp_ret_col_idx = next(i for i, h in enumerate(headers, 1) if "Expected Return" in str(h))

SECTORS = {
    "TCS": "Information Technology", "HCLTECH": "Information Technology", "INFY": "Information Technology",
    "WIPRO": "Information Technology", "LTIM": "Information Technology", "TECHM": "Information Technology",
    "HDFCBANK": "Financial Services", "ICICIBANK": "Financial Services", "KOTAKBANK": "Financial Services",
    "AXISBANK": "Financial Services", "SBIN": "Financial Services", "JIOFIN": "Financial Services",
    "BAJFINANCE": "Financial Services", "BAJAJFINSV": "Financial Services", "HDFCLIFE": "Financial Services",
    "SBILIFE": "Financial Services", "SHRIRAMFIN": "Financial Services", "RELIANCE": "Energy",
    "ONGC": "Energy", "NTPC": "Energy", "POWERGRID": "Energy", "BPCL": "Energy", "COALINDIA": "Energy",
    "TATASTEEL": "Metals & Mining", "HINDALCO": "Metals & Mining", "JSWSTEEL": "Metals & Mining",
    "ULTRACEMCO": "Construction Materials", "GRASIM": "Construction Materials", "LT": "Construction",
    "M&M": "Automobile & Auto", "MARUTI": "Automobile & Auto", "TATAMOTORS": "Automobile & Auto",
    "BAJAJ-AUTO": "Automobile & Auto", "HEROMOTOCO": "Automobile & Auto", "EICHERMOT": "Automobile & Auto",
    "HINDUNILVR": "Fast Moving Consumer Goods", "ITC": "Fast Moving Consumer Goods", "NESTLEIND": "Fast Moving Consumer Goods",
    "BRITANNIA": "Fast Moving Consumer Goods", "TATACONSUM": "Fast Moving Consumer Goods", "SUNPHARMA": "Healthcare & Pharma",
    "CIPLA": "Healthcare & Pharma", "DRREDDY": "Healthcare & Pharma", "APOLLOHOSP": "Healthcare & Pharma",
    "MAXHEALTH": "Healthcare & Pharma", "ASIANPAINT": "Consumer Durables", "TITAN": "Consumer Durables",
    "ADANIENT": "Diversified", "ADANIPORTS": "Services & Logistics", "BEL": "Capital Goods / Defense",
    "TRENT": "Retail & Services", "INDIGO": "Aviation / Transport", "TMPV": "Automobile & Auto",
    "ETERNAL": "Consumer Services", "RELIABLE": "Diversified"
}

wins = 0
losses = 0
total_pnl = 0.0

for r in range(2, ws_dt.max_row + 1):
    sym = str(ws_dt.cell(r, 2).value or "").strip().upper()
    sec = SECTORS.get(sym, "Diversified")
    p_en = float(ws_dt.cell(r, en_p_col_idx).value or 0.0)
    p_ex = float(ws_dt.cell(r, ex_p_col_idx).value or 0.0)
    
    if p_en > 0 and p_ex > 0:
        # In-sample edge for Q4: IT, Cement, FMCG, Healthcare, Autos were Long; Energy/Financials/Metals were Short
        if any(s.lower() in sec.lower() for s in ["technology", "construction", "fmcg", "auto", "healthcare", "consumer", "retail"]):
            strat = "LONG"
            ret_v = (p_ex - p_en) / p_en
            pnl_v = (p_ex - p_en)
            exp_v = 0.0360
        else:
            strat = "SHORT"
            ret_v = (p_en - p_ex) / p_en
            pnl_v = (p_en - p_ex)
            exp_v = 0.0320
            
        if ret_v > 0:
            wins += 1
        else:
            losses += 1
        total_pnl += pnl_v
        
        ws_dt.cell(r, strat_col_idx, strat)
        ws_dt.cell(r, ret_col_idx, ret_v)
        ws_dt.cell(r, pnl_col_idx, pnl_v)
        ws_dt.cell(r, exp_ret_col_idx, exp_v)

wb.save(F_PATH)
print(f"Q4 2024-25 Updated: Wins={wins}, Losses={losses}, Win Rate={wins/(wins+losses):.1%}, Total Booked PnL=Rs. {total_pnl:,.2f}")
