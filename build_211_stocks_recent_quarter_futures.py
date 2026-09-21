import os
import pathlib
import sys
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(errors='replace')

ROOT = pathlib.Path('D:/behaviour analysis')
OI_DIR = ROOT / 'OI_DATA'
PROC = ROOT / 'processed'
PRICE_CACHE = PROC / 'price_cache'

# 1. Discover all 211 symbols
SYMBOLS = sorted([e.name.strip().upper() for e in os.scandir(OI_DIR) if e.is_dir()])

print("==========================================================================================")
print("BUILDING 211 STOCKS RECENT QUARTER (Q2 2026-27) FUTURES MASTER EXCEL")
print("==========================================================================================")
print(f"Total Target Universe: {len(SYMBOLS)} Stocks")

# 2. Known Lot Sizes Mapping (Default to ~10 Lakhs contract value equivalent if missing)
LOT_SIZES = {
    "RELIANCE": 250, "TCS": 175, "INFY": 400, "HDFCBANK": 550, "ICICIBANK": 700,
    "BHARTIARTL": 950, "ITC": 1600, "SBIN": 1500, "LTIM": 150, "LT": 300,
    "HINDUNILVR": 300, "AXISBANK": 625, "KOTAKBANK": 400, "BAJFINANCE": 125,
    "M&M": 350, "MARUTI": 100, "SUNPHARMA": 350, "TATASTEEL": 5500,
    "NTPC": 1500, "POWERGRID": 1800, "TITAN": 175, "ADANIENT": 300,
    "ADANIPORTS": 625, "ULTRACEMCO": 100, "ASIANPAINT": 200, "COALINDIA": 2100,
    "BAJAJ-AUTO": 125, "JSWSTEEL": 675, "TATAMOTORS": 1425, "HCLTECH": 350,
    "GRASIM": 250, "HEROMOTOCO": 150, "EICHERMOT": 175, "CIPLA": 650,
    "HDFCLIFE": 1100, "SBILIFE": 375, "DRREDDY": 125, "BRITANNIA": 200,
    "APOLLOHOSP": 125, "TATACONSUM": 450, "HINDALCO": 1400, "BPCL": 1800,
    "INDUSINDBK": 500, "DIVISLAB": 200, "BAJAJFINSV": 500, "NESTLEIND": 200,
    "WIPRO": 1500, "ONGC": 3750, "TECHM": 600, "ASHOKLEY": 5000, "ADANIGREEN": 500
}

# 3. Load Position Taking Window & Result Dates for Q2 2026-27
trade_rows = []

for idx, sym in enumerate(SYMBOLS, 1):
    stock_folder = ROOT / sym
    fr_path = stock_folder / 'financial_results.csv'
    price_path = PRICE_CACHE / f"{sym}.csv"
    
    # Load prices
    df_prices = pd.DataFrame()
    if price_path.exists():
        try:
            df_prices = pd.read_csv(price_path)
            df_prices['date'] = pd.to_datetime(df_prices['date'])
            df_prices = df_prices.sort_values('date').reset_index(drop=True)
        except Exception:
            pass
            
    if df_prices.empty:
        continue
        
    # Get Q2 2026-27 Result Date
    result_date = pd.to_datetime("2026-07-25")
    if fr_path.exists():
        try:
            df_fr = pd.read_csv(fr_path)
            if not df_fr.empty and 'broadCastDate' in df_fr.columns:
                d_str = df_fr['broadCastDate'].iloc[0]
                dt = pd.to_datetime(d_str, errors='coerce')
                if pd.notna(dt):
                    result_date = dt
        except Exception:
            pass
            
    # Position Window (Default T-8 to T+1, customized per stock)
    pre_days = 8
    post_days = 1
    window_str = f"T-{pre_days} to T+{post_days}"
    
    # Find Trading Calendar Entry and Exit
    valid_dates = df_prices['date'].tolist()
    if not valid_dates:
        continue
        
    # Closest date to result
    res_idx = min(range(len(valid_dates)), key=lambda i: abs((valid_dates[i] - result_date).days))
    entry_idx = max(0, res_idx - pre_days)
    exit_idx = min(len(valid_dates) - 1, res_idx + post_days)
    
    entry_date = valid_dates[entry_idx]
    exit_date = valid_dates[exit_idx]
    
    buy_price = float(df_prices.loc[entry_idx, 'adj'])
    sell_price = float(df_prices.loc[exit_idx, 'adj'])
    
    lot_size = LOT_SIZES.get(sym, max(100, int(1000000 / buy_price)))
    
    buy_turnover = round(lot_size * buy_price, 2)
    sell_turnover = round(lot_size * sell_price, 2)
    combined_turnover = round(buy_turnover + sell_turnover, 2)
    txn_cost = round(combined_turnover * 0.0005, 2)
    gross_pnl = round(sell_turnover - buy_turnover, 2)
    net_pnl = round(gross_pnl - txn_cost, 2)
    pnl_pct = round((net_pnl / buy_turnover) * 100, 2)
    margin_req = round(0.20 * buy_turnover, 2)
    return_on_margin = round((net_pnl / margin_req) * 100, 2)
    
    trade_rows.append({
        "sym": sym,
        "quarter": "Q2 2026-27",
        "strategy": "FUTURE LONG",
        "window": window_str,
        "entry_date": entry_date.strftime("%Y-%m-%d"),
        "exit_date": exit_date.strftime("%Y-%m-%d"),
        "result_date": result_date.strftime("%Y-%m-%d"),
        "lot_size": lot_size,
        "buy_price": buy_price,
        "sell_price": sell_price,
        "buy_turnover": buy_turnover,
        "sell_turnover": sell_turnover,
        "combined_turnover": combined_turnover,
        "txn_cost": txn_cost,
        "gross_pnl": gross_pnl,
        "net_pnl": net_pnl,
        "pnl_pct": pnl_pct,
        "margin_req": margin_req,
        "return_on_margin": return_on_margin
    })

print(f"Total Trade Rows Prepared: {len(trade_rows)} Stocks")

# Calculate Cumulative PnL, Peak, and Drawdown
df_trades = pd.DataFrame(trade_rows)
df_trades['cum_pnl'] = df_trades['net_pnl'].cumsum()
df_trades['running_max'] = np.maximum(0.0, np.maximum.accumulate(df_trades['cum_pnl']))
df_trades['drawdown'] = df_trades['cum_pnl'] - df_trades['running_max']

# 4. Export to Formatted Master Openpyxl Excel
out_path = ROOT / 'Nifty211_Recent_Quarter_Futures_Equity_Window_Master.xlsx'
wb = openpyxl.Workbook()

# Sheet 1: Recent Quarter Q2 2026-27
ws = wb.active
ws.title = "Q2 2026-27"
ws.views.sheetView[0].showGridLines = True

# Formatting styles
title_font = Font(name="Calibri", size=16, bold=True, color="1F4E79")
hdr_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
hdr_fill = PatternFill("solid", fgColor="1F4E79")
subhdr_fill = PatternFill("solid", fgColor="2F5597")
border_thin = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                     top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
pos_fill = PatternFill("solid", fgColor="E2EFDA")
neg_fill = PatternFill("solid", fgColor="FCE4D6")

# Title Block
ws.merge_cells("A1:W1")
ws["A1"] = "NIFTY 211 STOCKS - RECENT QUARTER (Q2 2026-27) FUTURES EQUITY WINDOW MASTER"
ws["A1"].font = title_font
ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[1].height = 35

headers = [
    "Sr. No.", "Stock Symbol", "Quarter", "Strategy", "Position Taking Window",
    "Entry Date", "Exit Date", "Result Date", "Lot Size", "Buy Price (₹)",
    "Sell Price (₹)", "Turnover Buy (₹)", "Turnover Sell (₹)", "Combined Turnover (₹)",
    "Transaction Cost (0.05%)", "Gross P&L (₹)", "Net Realised P&L (₹)", "P&L %",
    "Margin Required (20% ₹)", "Booked Return on Margin %", "Cumulative P&L (₹)",
    "Peak Equity (₹)", "Drawdown (₹)"
]

ws.row_dimensions[3].height = 28
for col_idx, h in enumerate(headers, 1):
    cell = ws.cell(row=3, column=col_idx, value=h)
    cell.font = hdr_font
    cell.fill = hdr_fill
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

# Insert Data Rows
for r_idx, r in enumerate(df_trades.to_dict('records'), start=4):
    ws.cell(row=r_idx, column=1, value=r_idx - 3)
    ws.cell(row=r_idx, column=2, value=r['sym'])
    ws.cell(row=r_idx, column=3, value=r['quarter'])
    ws.cell(row=r_idx, column=4, value=r['strategy'])
    ws.cell(row=r_idx, column=5, value=r['window'])
    ws.cell(row=r_idx, column=6, value=r['entry_date'])
    ws.cell(row=r_idx, column=7, value=r['exit_date'])
    ws.cell(row=r_idx, column=8, value=r['result_date'])
    ws.cell(row=r_idx, column=9, value=r['lot_size'])
    ws.cell(row=r_idx, column=10, value=r['buy_price'])
    ws.cell(row=r_idx, column=11, value=r['sell_price'])
    ws.cell(row=r_idx, column=12, value=f"=I{r_idx}*J{r_idx}")
    ws.cell(row=r_idx, column=13, value=f"=I{r_idx}*K{r_idx}")
    ws.cell(row=r_idx, column=14, value=f"=L{r_idx}+M{r_idx}")
    ws.cell(row=r_idx, column=15, value=f"=ROUND(N{r_idx}*0.0005, 2)")
    ws.cell(row=r_idx, column=16, value=f"=M{r_idx}-L{r_idx}")
    ws.cell(row=r_idx, column=17, value=f"=P{r_idx}-O{r_idx}")
    ws.cell(row=r_idx, column=18, value=f"=ROUND((Q{r_idx}/L{r_idx})*100, 2)")
    ws.cell(row=r_idx, column=19, value=f"=ROUND(0.20*L{r_idx}, 2)")
    ws.cell(row=r_idx, column=20, value=f"=ROUND((Q{r_idx}/S{r_idx})*100, 2)")
    
    if r_idx == 4:
        ws.cell(row=r_idx, column=21, value=f"=Q{r_idx}")
        ws.cell(row=r_idx, column=22, value=f"=MAX(0, U{r_idx})")
        ws.cell(row=r_idx, column=23, value=f"=U{r_idx}-V{r_idx}")
    else:
        ws.cell(row=r_idx, column=21, value=f"=U{r_idx-1}+Q{r_idx}")
        ws.cell(row=r_idx, column=22, value=f"=MAX(V{r_idx-1}, U{r_idx})")
        ws.cell(row=r_idx, column=23, value=f"=U{r_idx}-V{r_idx}")

    # Formats & Alignments
    for c_idx in range(1, 24):
        cell = ws.cell(row=r_idx, column=c_idx)
        cell.border = border_thin
        if c_idx in (1, 3, 4, 5, 6, 7, 8):
            cell.alignment = Alignment(horizontal="center")
        elif c_idx == 2:
            cell.alignment = Alignment(horizontal="left")
        else:
            cell.alignment = Alignment(horizontal="right")
            
        if c_idx in (10, 11):
            cell.number_format = "#,##0.00"
        elif c_idx in (12, 13, 14, 15, 16, 17, 19, 21, 22, 23):
            cell.number_format = "₹#,##0.00"
        elif c_idx in (18, 20):
            cell.number_format = "0.00\"%\""

# Summary Row
tot_row = len(df_trades) + 4
ws.cell(row=tot_row, column=1, value="TOTAL / SUMMARY")
ws.cell(row=tot_row, column=17, value=f"=SUM(Q4:Q{tot_row-1})")
ws.cell(row=tot_row, column=17).number_format = "₹#,##0.00"
ws.cell(row=tot_row, column=17).font = Font(bold=True)

# Column Widths
for col in ws.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

wb.save(out_path)

print("==========================================================================================")
print(f"✅ SUCCESSFULLY SAVED MASTER EXCEL: {out_path.name}")
print("==========================================================================================")
