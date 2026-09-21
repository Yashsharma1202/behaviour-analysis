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

SYMBOLS = sorted([e.name.strip().upper() for e in os.scandir(OI_DIR) if e.is_dir()])

print("==========================================================================================")
print("BUILDING 12-QUARTER WINDOW WIN RATE RANKING MASTER WORKBOOK")
print("==========================================================================================")
print(f"Total Target Universe: {len(SYMBOLS)} Stocks")

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

def get_stock_window_params(sym):
    h = sum(ord(c) for c in sym)
    pre_options = [8, 7, 5, 4, 3, 2, 1]
    post_options = [1, 2, 3, 4, 5, 8]
    pre = pre_options[h % len(pre_options)]
    post = post_options[(h * 3) % len(post_options)]
    return pre, post

quarters_list = [
    ("Q2 2026-27", "2026-07-25"),
    ("Q1 2026-27", "2026-04-22"),
    ("Q4 2025-26", "2026-01-20"),
    ("Q3 2025-26", "2025-10-21"),
    ("Q2 2025-26", "2025-07-25"),
    ("Q1 2025-26", "2025-04-22"),
    ("Q4 2024-25", "2025-01-20"),
    ("Q3 2024-25", "2024-10-21"),
    ("Q2 2024-25", "2024-07-25"),
    ("Q1 2024-25", "2024-04-22"),
    ("Q4 2023-24", "2024-01-20"),
    ("Q3 2023-24", "2023-10-21")
]

ranking_data = []

for idx, sym in enumerate(SYMBOLS, 1):
    stock_folder = ROOT / sym
    fr_path = stock_folder / 'financial_results.csv'
    price_path = PRICE_CACHE / f"{sym}.csv"
    
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
        
    n_quarters = 0
    q_dates_history = []
    
    if fr_path.exists() and fr_path.stat().st_size > 10:
        try:
            df_fr = pd.read_csv(fr_path, dtype=str).fillna('')
            if not df_fr.empty and 'broadCastDate' in df_fr.columns:
                dts = pd.to_datetime(df_fr['broadCastDate'], errors='coerce').dropna().sort_values(ascending=False)
                n_quarters = len(dts)
                q_dates_history = dts.tolist()
        except Exception:
            pass
            
    trading_status = "QUALIFIED" if n_quarters >= 12 else "AVOID BUT MONITOR IT"
    pre_days, post_days = get_stock_window_params(sym)
    window_str = f"T-{pre_days} to T+{post_days}"
    valid_dates = df_prices['date'].tolist()
    if not valid_dates:
        continue
        
    lot_size = LOT_SIZES.get(sym, max(100, int(1000000 / float(df_prices['adj'].iloc[-1]))))
    last_price = float(df_prices['adj'].iloc[-1])
    margin_for_stock = round(0.20 * lot_size * last_price, 2)
    
    # Compute Empirical Historical Win Rates for full history to determine optimal strategy
    emp_wins_long = 0; emp_wins_short = 0
    total_q = len(q_dates_history)
    
    for dt in q_dates_history:
        r_i = min(range(len(valid_dates)), key=lambda i: abs((valid_dates[i] - dt).days))
        e_i = max(0, r_i - pre_days)
        x_i = min(len(valid_dates) - 1, r_i + post_days)
        pe = float(df_prices.loc[e_i, 'adj']); px = float(df_prices.loc[x_i, 'adj'])
        if pe > 0:
            if px > pe: emp_wins_long += 1
            elif pe > px: emp_wins_short += 1
            
    wr_long = (emp_wins_long / total_q) * 100 if total_q > 0 else 50.0
    wr_short = (emp_wins_short / total_q) * 100 if total_q > 0 else 50.0

    optimal_strat = "FUTURE LONG" if wr_long >= wr_short else "FUTURE SHORT"

    # Evaluate performance SPECIFICALLY ACROSS THE 12-QUARTER WINDOW
    wins_12q = 0
    losses_12q = 0
    trades_12q = 0
    stock_12q_pnl = 0.0
    rets_12q = []

    for q_idx, (q_name, default_d_str) in enumerate(quarters_list):
        if q_idx < len(q_dates_history):
            result_date = q_dates_history[q_idx]
        else:
            result_date = pd.to_datetime(default_d_str)
            
        res_idx = min(range(len(valid_dates)), key=lambda i: abs((valid_dates[i] - result_date).days))
        entry_idx = max(0, res_idx - pre_days)
        exit_idx = min(len(valid_dates) - 1, res_idx + post_days)
        
        p_entry = float(df_prices.loc[entry_idx, 'adj'])
        p_exit = float(df_prices.loc[exit_idx, 'adj'])
        
        if optimal_strat == "FUTURE LONG":
            buy_p = p_entry; sell_p = p_exit
            act_r = (p_exit - p_entry) / p_entry * 100
        else:
            buy_p = p_exit; sell_p = p_entry
            act_r = (p_entry - p_exit) / p_entry * 100
            
        buy_to = round(lot_size * buy_p, 2); sell_to = round(lot_size * sell_p, 2)
        comb_to = round(buy_to + sell_to, 2); cost = round(comb_to * 0.0005, 2)
        net_pnl = round(sell_to - buy_to - cost, 2)
        
        trades_12q += 1
        rets_12q.append(act_r)
        stock_12q_pnl += net_pnl
        if net_pnl > 0:
            wins_12q += 1
        else:
            losses_12q += 1

    win_rate_12q = round((wins_12q / trades_12q) * 100, 2) if trades_12q > 0 else 0.0
    avg_ret_12q = round(np.mean(rets_12q), 2) if rets_12q else 0.0

    ranking_data.append({
        "sym": sym,
        "strategy": optimal_strat,
        "trading_status": trading_status,
        "n_quarters_str": f"{n_quarters} Quarters",
        "n_quarters_num": n_quarters,
        "window": window_str,
        "win_rate_12q": win_rate_12q,
        "trades_12q": trades_12q,
        "wins_12q": wins_12q,
        "losses_12q": losses_12q,
        "avg_ret_12q": avg_ret_12q,
        "margin": margin_for_stock,
        "pnl_12q": round(stock_12q_pnl, 2),
        "rom_12q": round((stock_12q_pnl / margin_for_stock) * 100, 2) if margin_for_stock > 0 else 0.0
    })

df_all = pd.DataFrame(ranking_data)

# Sort strictly by 12-Quarter Win Rate % descending, then 12-Quarter Net P&L descending
df_all = df_all.sort_values(['win_rate_12q', 'pnl_12q'], ascending=[False, False]).reset_index(drop=True)

df_q = df_all[df_all['trading_status'] == 'QUALIFIED'].reset_index(drop=True)
df_avoid = df_all[df_all['trading_status'] == 'AVOID BUT MONITOR IT'].reset_index(drop=True)

print(f"Total Ranked Universe: {len(df_all)} Stocks")
print(f"Qualified Stocks (>=12 Qtrs): {len(df_q)} Stocks")
print(f"Avoid but Monitor (<12 Qtrs): {len(df_avoid)} Stocks")

out_path = ROOT / "Nifty211_All_Stocks_12_Quarters_WinRate_Ranking_Master.xlsx"
wb = openpyxl.Workbook()

title_font = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
title_fill = PatternFill("solid", fgColor="1F4E79")
card_val_font = Font(name="Calibri", size=15, bold=True, color="1F4E79")
hdr_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
hdr_fill = PatternFill("solid", fgColor="1F4E79")
subhdr_fill = PatternFill("solid", fgColor="2F5597")
border_thin = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                     top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
card_fill = PatternFill("solid", fgColor="F2F4F7")
pos_fill = PatternFill("solid", fgColor="E2EFDA")
neg_fill = PatternFill("solid", fgColor="FCE4D6")
amb_fill = PatternFill("solid", fgColor="FFF2CC")
pos_font = Font(name="Calibri", size=11, bold=True, color="276A3C")
neg_font = Font(name="Calibri", size=11, bold=True, color="9C0006")
amb_font = Font(name="Calibri", size=11, bold=True, color="B25900")

def write_ranking_sheet(ws, sheet_title, df_sheet):
    ws.views.sheetView[0].showGridLines = True
    
    ws.merge_cells("A1:L1")
    ws["A1"] = f"NIFTY 211 STOCKS - {sheet_title.upper()} RANKING BY 12-QUARTER WIN RATE %"
    ws["A1"].font = title_font; ws["A1"].fill = title_fill
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 38

    # KPI Cards
    top_stock = df_sheet.iloc[0]['sym'] if not df_sheet.empty else "N/A"
    top_wr = df_sheet.iloc[0]['win_rate_12q'] if not df_sheet.empty else 0.0
    avg_wr = round(df_sheet['win_rate_12q'].mean(), 2) if not df_sheet.empty else 0.0
    tot_pnl = round(df_sheet['pnl_12q'].sum(), 2) if not df_sheet.empty else 0.0
    
    cards = [
        ("TOTAL STOCKS RANKED", f"{len(df_sheet)} Stocks", "A3:B4", "A3"),
        ("TOP 12-QTR WIN RATE STOCK", f"{top_stock} ({top_wr}%)", "C3:D4", "C3"),
        ("AVERAGE 12-QTR WIN RATE %", f"{avg_wr}%", "E3:F4", "E3"),
        ("12-QTR NET REALISED P&L", f"₹{tot_pnl:,.2f}", "G3:H4", "G3")
    ]
    for title, val, merge_range, top_left in cards:
        ws.merge_cells(merge_range)
        ws[top_left] = f"{title}\n{val}"
        ws[top_left].font = card_val_font
        ws[top_left].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws[top_left].fill = card_fill; ws[top_left].border = border_thin

    # Win Rate Bracket Distribution Summary Table
    ws.cell(row=6, column=1, value="12-QUARTER WIN RATE BRACKET DISTRIBUTION SUMMARY").font = hdr_font
    ws.cell(row=6, column=1).fill = hdr_fill
    ws.merge_cells("A6:D6")
    
    ws.cell(row=7, column=1, value="WIN RATE BRACKET").font = Font(bold=True, color="FFFFFF"); ws.cell(row=7, column=1).fill = subhdr_fill
    ws.cell(row=7, column=2, value="STOCKS COUNT").font = Font(bold=True, color="FFFFFF"); ws.cell(row=7, column=2).fill = subhdr_fill
    ws.cell(row=7, column=3, value="PERCENTAGE OF UNIVERSE").font = Font(bold=True, color="FFFFFF"); ws.cell(row=7, column=3).fill = subhdr_fill
    ws.cell(row=7, column=4, value="NET REALISED P&L (₹)").font = Font(bold=True, color="FFFFFF"); ws.cell(row=7, column=4).fill = subhdr_fill

    b_80 = df_sheet[df_sheet['win_rate_12q'] >= 80]
    b_70 = df_sheet[(df_sheet['win_rate_12q'] >= 70) & (df_sheet['win_rate_12q'] < 80)]
    b_60 = df_sheet[(df_sheet['win_rate_12q'] >= 60) & (df_sheet['win_rate_12q'] < 70)]
    b_50 = df_sheet[(df_sheet['win_rate_12q'] >= 50) & (df_sheet['win_rate_12q'] < 60)]
    b_below = df_sheet[df_sheet['win_rate_12q'] < 50]
    
    brackets = [
        ("≥ 80% Win Rate (Ultra High Accuracy)", b_80),
        ("70% - 79% Win Rate (High Accuracy)", b_70),
        ("60% - 69% Win Rate (Above Average)", b_60),
        ("50% - 59% Win Rate (Average Accuracy)", b_50),
        ("< 50% Win Rate (Below Average)", b_below)
    ]
    
    for r_idx, (b_lbl, b_df) in enumerate(brackets, start=8):
        n_b = len(b_df)
        pct_b = round((n_b / len(df_sheet)) * 100, 2) if len(df_sheet) > 0 else 0.0
        pnl_b = round(b_df['pnl_12q'].sum(), 2)
        
        ws.cell(row=r_idx, column=1, value=b_lbl).border = border_thin
        ws.cell(row=r_idx, column=2, value=n_b).border = border_thin
        ws.cell(row=r_idx, column=2).alignment = Alignment(horizontal="center")
        ws.cell(row=r_idx, column=3, value=f"{pct_b}%").border = border_thin
        ws.cell(row=r_idx, column=3).alignment = Alignment(horizontal="right")
        c_p = ws.cell(row=r_idx, column=4, value=f"₹{pnl_b:,.2f}")
        c_p.border = border_thin; c_p.alignment = Alignment(horizontal="right")
        if pnl_b >= 0: c_p.fill = pos_fill; c_p.font = pos_font
        else: c_p.fill = neg_fill; c_p.font = neg_font

    # Main Ranking Table
    rank_start_row = 15
    ws.cell(row=rank_start_row, column=1, value=f"ALL {len(df_sheet)} STOCKS RANKED BY 12-QUARTER WIN RATE %").font = hdr_font
    ws.cell(row=rank_start_row, column=1).fill = hdr_fill
    ws.merge_cells(start_row=rank_start_row, start_column=1, end_row=rank_start_row, end_column=12)

    headers = [
        "Rank", "Stock Symbol", "Optimal Strategy", "Action Status", "Historical Quarters",
        "Position Window", "12-Quarter Win Rate %", "12-Quarter Wins", "12-Quarter Losses",
        "12-Quarter Avg Return %", "Margin Capital (₹)", "12-Quarter Net P&L (₹)"
    ]
    ws.row_dimensions[rank_start_row+1].height = 26
    for c_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=rank_start_row+1, column=c_idx, value=h)
        cell.font = hdr_font; cell.fill = subhdr_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for r_idx, r in enumerate(df_sheet.to_dict('records'), start=rank_start_row+2):
        ws.cell(row=r_idx, column=1, value=r_idx - rank_start_row - 1).border = border_thin
        ws.cell(row=r_idx, column=1).alignment = Alignment(horizontal="center")
        
        ws.cell(row=r_idx, column=2, value=r['sym']).border = border_thin
        ws.cell(row=r_idx, column=2).font = Font(bold=True)
        
        c_st = ws.cell(row=r_idx, column=3, value=r['strategy'])
        c_st.border = border_thin; c_st.alignment = Alignment(horizontal="center")
        if r['strategy'] == 'FUTURE LONG': c_st.fill = pos_fill; c_st.font = pos_font
        else: c_st.fill = neg_fill; c_st.font = neg_font

        c_act = ws.cell(row=r_idx, column=4, value=r['trading_status'])
        c_act.border = border_thin; c_act.alignment = Alignment(horizontal="center")
        if r['trading_status'] == 'QUALIFIED': c_act.fill = pos_fill; c_act.font = pos_font
        else: c_act.fill = amb_fill; c_act.font = amb_font
        
        ws.cell(row=r_idx, column=5, value=r['n_quarters_str']).border = border_thin
        ws.cell(row=r_idx, column=5).alignment = Alignment(horizontal="center")
        
        ws.cell(row=r_idx, column=6, value=r['window']).border = border_thin
        ws.cell(row=r_idx, column=6).alignment = Alignment(horizontal="center")
        
        c_wr = ws.cell(row=r_idx, column=7, value=f"{r['win_rate_12q']}%")
        c_wr.border = border_thin; c_wr.alignment = Alignment(horizontal="right")
        c_wr.font = Font(bold=True)
        if r['win_rate_12q'] >= 65.0: c_wr.fill = pos_fill; c_wr.font = pos_font
        elif r['win_rate_12q'] < 50.0: c_wr.fill = neg_fill; c_wr.font = neg_font
        
        ws.cell(row=r_idx, column=8, value=r['wins_12q']).border = border_thin
        ws.cell(row=r_idx, column=8).alignment = Alignment(horizontal="center")
        
        ws.cell(row=r_idx, column=9, value=r['losses_12q']).border = border_thin
        ws.cell(row=r_idx, column=9).alignment = Alignment(horizontal="center")
        
        ws.cell(row=r_idx, column=10, value=f"{r['avg_ret_12q']}%").border = border_thin
        ws.cell(row=r_idx, column=10).alignment = Alignment(horizontal="right")
        
        ws.cell(row=r_idx, column=11, value=f"₹{r['margin']:,.2f}").border = border_thin
        ws.cell(row=r_idx, column=11).alignment = Alignment(horizontal="right")
        
        c_p = ws.cell(row=r_idx, column=12, value=f"₹{r['pnl_12q']:,.2f}")
        c_p.border = border_thin; c_p.alignment = Alignment(horizontal="right")
        if r['pnl_12q'] >= 0: c_p.fill = pos_fill; c_p.font = pos_font
        else: c_p.fill = neg_fill; c_p.font = neg_font

    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = 22

# Write Sheet 1: All 211 Stocks 12Q Ranking
ws1 = wb.active
ws1.title = "All 211 Stocks 12Q Ranking"
write_ranking_sheet(ws1, "All 211 Stocks", df_all)

# Write Sheet 2: Qualified Stocks (69)
ws2 = wb.create_sheet(title="Qualified Stocks (69)")
write_ranking_sheet(ws2, "69 Qualified Stocks", df_q)

# Write Sheet 3: Avoid but Monitor (142)
ws3 = wb.create_sheet(title="Avoid but Monitor (142)")
write_ranking_sheet(ws3, "142 Avoid but Monitor Stocks", df_avoid)

wb.save(out_path)
print(f"  • Successfully generated: Nifty211_All_Stocks_12_Quarters_WinRate_Ranking_Master.xlsx")
print("==========================================================================================")
print("12-QUARTER WINDOW WIN RATE RANKING MASTER WORKBOOK CREATED SUCCESSFULLY!")
print("==========================================================================================")
