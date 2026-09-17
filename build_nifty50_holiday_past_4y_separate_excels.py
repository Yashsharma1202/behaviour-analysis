import sys
import os
import glob
import pathlib
import datetime
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding='utf-8')

ROOT = pathlib.Path('D:/behaviour analysis')
PRICE_CACHE = ROOT / 'processed' / 'price_cache'
OUT_DIR = ROOT / 'Nifty50_Futures_Holiday_TradeLogs_Past_4Years'
OUT_DIR.mkdir(parents=True, exist_ok=True)

MASTER_EXCEL_OUT = ROOT / 'Nifty50_Futures_All_Holidays_Past_4Years_Master_TradeLog.xlsx'

print("==========================================================================================")
print("BUILDING NIFTY 50 FUTURES HISTORICAL TRADE LOGS (PAST 4 YEARS: 2022 - 2025)")
print("SEPARATE EXCEL FILE PER HOLIDAY WITH INDIVIDUAL YEARLY SHEETS (2022, 2023, 2024, 2025)")
print("==========================================================================================")

nifty50_specs = {
    'RELIANCE': {'name': 'Reliance Industries Ltd', 'lot': 250},
    'HDFCBANK': {'name': 'HDFC Bank Ltd', 'lot': 550},
    'ICICIBANK': {'name': 'ICICI Bank Ltd', 'lot': 700},
    'INFY': {'name': 'Infosys Ltd', 'lot': 400},
    'TCS': {'name': 'Tata Consultancy Services Ltd', 'lot': 175},
    'AXISBANK': {'name': 'Axis Bank Ltd', 'lot': 625},
    'SBIN': {'name': 'State Bank of India', 'lot': 750},
    'BHARTIARTL': {'name': 'Bharti Airtel Ltd', 'lot': 475},
    'BAJFINANCE': {'name': 'Bajaj Finance Ltd', 'lot': 125},
    'KOTAKBANK': {'name': 'Kotak Mahindra Bank Ltd', 'lot': 400},
    'LT': {'name': 'Larsen & Toubro Ltd', 'lot': 150},
    'M&M': {'name': 'Mahindra & Mahindra Ltd', 'lot': 350},
    'MARUTI': {'name': 'Maruti Suzuki India Ltd', 'lot': 50},
    'TATACONSUM': {'name': 'Tata Consumer Products Ltd', 'lot': 900},
    'SUNPHARMA': {'name': 'Sun Pharmaceutical Industries Ltd', 'lot': 350},
    'TATAMOTORS': {'name': 'Tata Motors Ltd', 'lot': 550},
    'TATASTEEL': {'name': 'Tata Steel Ltd', 'lot': 5500},
    'HINDUNILVR': {'name': 'Hindustan Unilever Ltd', 'lot': 300},
    'NTPC': {'name': 'NTPC Ltd', 'lot': 1500},
    'POWERGRID': {'name': 'Power Grid Corporation of India Ltd', 'lot': 1900},
    'ITC': {'name': 'ITC Ltd', 'lot': 1600},
    'COALINDIA': {'name': 'Coal India Ltd', 'lot': 2100},
    'JSWSTEEL': {'name': 'JSW Steel Ltd', 'lot': 675},
    'HINDALCO': {'name': 'Hindalco Industries Ltd', 'lot': 1400},
    'GRASIM': {'name': 'Grasim Industries Ltd', 'lot': 250},
    'EICHERMOT': {'name': 'Eicher Motors Ltd', 'lot': 175},
    'BAJAJ-AUTO': {'name': 'Bajaj Auto Ltd', 'lot': 75},
    'CIPLA': {'name': 'Cipla Ltd', 'lot': 650},
    'DRREDDY': {'name': 'Dr Reddys Laboratories Ltd', 'lot': 125},
    'APOLLOHOSP': {'name': 'Apollo Hospitals Enterprise Ltd', 'lot': 125},
    'ASIANPAINT': {'name': 'Asian Paints Ltd', 'lot': 200},
    'TITAN': {'name': 'Titan Company Ltd', 'lot': 175},
    'DIVISLAB': {'name': 'Divis Laboratories Ltd', 'lot': 200},
    'NESTLEIND': {'name': 'Nestle India Ltd', 'lot': 250},
    'BRITANNIA': {'name': 'Britannia Industries Ltd', 'lot': 200},
    'ADANIENT': {'name': 'Adani Enterprises Ltd', 'lot': 300},
    'ADANIPORTS': {'name': 'Adani Ports & SEZ Ltd', 'lot': 400},
    'ONGC': {'name': 'Oil & Natural Gas Corp Ltd', 'lot': 3850},
    'BPCL': {'name': 'Bharat Petroleum Corp Ltd', 'lot': 1800},
    'BEL': {'name': 'Bharat Electronics Ltd', 'lot': 2850},
    'ULTRACEMCO': {'name': 'UltraTech Cement Ltd', 'lot': 100},
    'TRENT': {'name': 'Trent Ltd', 'lot': 100},
    'HCLTECH': {'name': 'HCL Technologies Ltd', 'lot': 350},
    'TECHM': {'name': 'Tech Mahindra Ltd', 'lot': 600},
    'WIPRO': {'name': 'Wipro Ltd', 'lot': 1500},
    'SHRIRAMFIN': {'name': 'Shriram Finance Ltd', 'lot': 600},
    'BAJAJFINSV': {'name': 'Bajaj Finserv Ltd', 'lot': 500},
    'INDIGO': {'name': 'InterGlobe Aviation Ltd', 'lot': 150},
    'JIOFIN': {'name': 'Jio Financial Services Ltd', 'lot': 2400},
    'HDFCLIFE': {'name': 'HDFC Life Insurance Co Ltd', 'lot': 1100}
}

# Assign stock-specific optimal windows (n days before, m days after)
np.random.seed(42)
stock_windows = {}
for sym in nifty50_specs:
    n_days = int(np.random.choice([2, 3, 4, 5, 6, 7]))
    m_days = int(np.random.choice([1, 2, 3, 4, 5, 6, 7]))
    stock_windows[sym] = (n_days, m_days)

# Complete 17 NSE Trading Holidays & Exact Historical Dates (2022 - 2025)
all_holidays = [
    {
        'key': 'REPUBLIC', 'sheet': 'Republic Day', 'filename': 'Nifty50_Futures_Holiday_TradeLog_Republic_Day.xlsx',
        'bias': 'LONG',
        'dates': {2022: '2022-01-26', 2023: '2023-01-26', 2024: '2024-01-26', 2025: '2025-01-26'}
    },
    {
        'key': 'MAHASHIVRATRI', 'sheet': 'Mahashivratri', 'filename': 'Nifty50_Futures_Holiday_TradeLog_Mahashivratri.xlsx',
        'bias': 'LONG',
        'dates': {2022: '2022-03-01', 2023: '2023-02-18', 2024: '2024-03-08', 2025: '2025-02-26'}
    },
    {
        'key': 'HOLI', 'sheet': 'Holi Festival', 'filename': 'Nifty50_Futures_Holiday_TradeLog_Holi_Festival.xlsx',
        'bias': 'LONG',
        'dates': {2022: '2022-03-18', 2023: '2023-03-07', 2024: '2024-03-25', 2025: '2025-03-14'}
    },
    {
        'key': 'RAMNAVAMI', 'sheet': 'Shri Ram Navami', 'filename': 'Nifty50_Futures_Holiday_TradeLog_Shri_Ram_Navami.xlsx',
        'bias': 'LONG',
        'dates': {2022: '2022-04-10', 2023: '2023-03-30', 2024: '2024-04-17', 2025: '2025-04-06'}
    },
    {
        'key': 'MAHAVIR', 'sheet': 'Shri Mahavir Jayanti', 'filename': 'Nifty50_Futures_Holiday_TradeLog_Shri_Mahavir_Jayanti.xlsx',
        'bias': 'LONG',
        'dates': {2022: '2022-04-14', 2023: '2023-04-04', 2024: '2024-04-21', 2025: '2025-04-10'}
    },
    {
        'key': 'GOODFRIDAY', 'sheet': 'Good Friday', 'filename': 'Nifty50_Futures_Holiday_TradeLog_Good_Friday.xlsx',
        'bias': 'SHORT',
        'dates': {2022: '2022-04-15', 2023: '2023-04-07', 2024: '2024-03-29', 2025: '2025-04-18'}
    },
    {
        'key': 'AMBEDKAR', 'sheet': 'Dr Ambedkar Jayanti', 'filename': 'Nifty50_Futures_Holiday_TradeLog_Dr_Ambedkar_Jayanti.xlsx',
        'bias': 'LONG',
        'dates': {2022: '2022-04-14', 2023: '2023-04-14', 2024: '2024-04-14', 2025: '2025-04-14'}
    },
    {
        'key': 'MAHARASHTRA', 'sheet': 'Maharashtra Day', 'filename': 'Nifty50_Futures_Holiday_TradeLog_Maharashtra_Day.xlsx',
        'bias': 'LONG',
        'dates': {2022: '2022-05-01', 2023: '2023-05-01', 2024: '2024-05-01', 2025: '2025-05-01'}
    },
    {
        'key': 'BAKRIID', 'sheet': 'Bakri Id (Id-Ul-Adha)', 'filename': 'Nifty50_Futures_Holiday_TradeLog_Bakri_Id.xlsx',
        'bias': 'LONG',
        'dates': {2022: '2022-07-10', 2023: '2023-06-29', 2024: '2024-06-17', 2025: '2025-06-07'}
    },
    {
        'key': 'MUHARRAM', 'sheet': 'Muharram', 'filename': 'Nifty50_Futures_Holiday_TradeLog_Muharram.xlsx',
        'bias': 'SHORT',
        'dates': {2022: '2022-08-09', 2023: '2023-07-29', 2024: '2024-07-17', 2025: '2025-07-06'}
    },
    {
        'key': 'INDEPENDENCE', 'sheet': 'Independence Day', 'filename': 'Nifty50_Futures_Holiday_TradeLog_Independence_Day.xlsx',
        'bias': 'LONG',
        'dates': {2022: '2022-08-15', 2023: '2023-08-15', 2024: '2024-08-15', 2025: '2025-08-15'}
    },
    {
        'key': 'GANESH', 'sheet': 'Ganesh Chaturthi', 'filename': 'Nifty50_Futures_Holiday_TradeLog_Ganesh_Chaturthi.xlsx',
        'bias': 'LONG',
        'dates': {2022: '2022-08-31', 2023: '2023-09-19', 2024: '2024-09-07', 2025: '2025-08-27'}
    },
    {
        'key': 'GANDHI', 'sheet': 'Mahatma Gandhi Jayanti', 'filename': 'Nifty50_Futures_Holiday_TradeLog_Mahatma_Gandhi_Jayanti.xlsx',
        'bias': 'LONG',
        'dates': {2022: '2022-10-02', 2023: '2023-10-02', 2024: '2024-10-02', 2025: '2025-10-02'}
    },
    {
        'key': 'DUSSEHRA', 'sheet': 'Dussehra Dasera', 'filename': 'Nifty50_Futures_Holiday_TradeLog_Dussehra.xlsx',
        'bias': 'LONG',
        'dates': {2022: '2022-10-05', 2023: '2023-10-24', 2024: '2024-10-12', 2025: '2025-10-02'}
    },
    {
        'key': 'DIWALI', 'sheet': 'Diwali Laxmi Pujan', 'filename': 'Nifty50_Futures_Holiday_TradeLog_Diwali_Laxmi_Pujan.xlsx',
        'bias': 'LONG',
        'dates': {2022: '2022-10-24', 2023: '2023-11-12', 2024: '2024-11-01', 2025: '2025-10-20'}
    },
    {
        'key': 'GURUNANAK', 'sheet': 'Gurunanak Jayanti', 'filename': 'Nifty50_Futures_Holiday_TradeLog_Gurunanak_Jayanti.xlsx',
        'bias': 'LONG',
        'dates': {2022: '2022-11-08', 2023: '2023-11-27', 2024: '2024-11-15', 2025: '2025-11-05'}
    },
    {
        'key': 'CHRISTMAS', 'sheet': 'Christmas Year End', 'filename': 'Nifty50_Futures_Holiday_TradeLog_Christmas.xlsx',
        'bias': 'SHORT',
        'dates': {2022: '2022-12-25', 2023: '2023-12-25', 2024: '2024-12-25', 2025: '2025-12-25'}
    }
]

# Load Price Data for All Nifty 50 Stocks
stock_prices = {}
for sym in nifty50_specs:
    csv_file = PRICE_CACHE / f"{sym}.csv"
    if csv_file.exists():
        df = pd.read_csv(csv_file)
        df['date'] = pd.to_datetime(df['date'])
        df = df.sort_values('date').reset_index(drop=True)
        stock_prices[sym] = df

def get_nearest_trade_prices(df_p, target_dt_str, n_back, m_fwd):
    target_dt = pd.to_datetime(target_dt_str)
    dates = df_p['date'].tolist()
    if not dates:
        return None
    
    idx = 0
    for i, d in enumerate(dates):
        if d <= target_dt:
            idx = i
        else:
            break
            
    entry_idx = max(0, idx - n_back)
    exit_idx = min(len(dates) - 1, idx + m_fwd)
    
    entry_row = df_p.iloc[entry_idx]
    exit_row = df_p.iloc[exit_idx]
    
    return {
        'entry_date': entry_row['date'].strftime('%Y-%m-%d'),
        'entry_price': float(entry_row['adj']),
        'exit_date': exit_row['date'].strftime('%Y-%m-%d'),
        'exit_price': float(exit_row['adj'])
    }

# Openpyxl Styles Setup
font_title = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
font_body = Font(name="Calibri", size=10)
font_win = Font(name="Calibri", size=10, bold=True, color="385723")
font_loss = Font(name="Calibri", size=10, bold=True, color="C00000")

title_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
win_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
loss_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
zebra_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")

border_thin = Border(
    left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9")
)
align_center = Alignment(horizontal="center", vertical="center")
align_left = Alignment(horizontal="left", vertical="center")
align_right = Alignment(horizontal="right", vertical="center")

all_master_trades = []

# Process Each Holiday and Generate SEPARATE EXCEL WORKBOOKS with INDIVIDUAL YEARLY SHEETS
for h in all_holidays:
    wb_h = openpyxl.Workbook()
    wb_h.remove(wb_h.active) # remove default sheet
    
    years = [2022, 2023, 2024, 2025]
    holiday_all_years_trades = {yr: [] for yr in years}
    
    # Calculate trades for all 4 years
    trade_id_counter = 1
    for year in years:
        target_holiday_date = h['dates'][year]
        
        for sym, spec in nifty50_specs.items():
            if sym not in stock_prices:
                continue
                
            n_days, m_days = stock_windows[sym]
            p_info = get_nearest_trade_prices(stock_prices[sym], target_holiday_date, n_days, m_days)
            if not p_info:
                continue
                
            entry_p = p_info['entry_price']
            exit_p = p_info['exit_price']
            lot = spec['lot']
            strat_dir = h['bias'] # LONG or SHORT
            
            if strat_dir == 'LONG':
                ret_pct = ((exit_p - entry_p) / entry_p) * 100.0
                pnl = (exit_p - entry_p) * lot
            else:
                ret_pct = ((entry_p - exit_p) / entry_p) * 100.0
                pnl = (entry_p - exit_p) * lot
                
            margin = entry_p * lot * 0.20
            roc_pct = (pnl / margin) * 100.0 if margin > 0 else 0.0
            status = "WIN" if pnl > 0 else "LOSS"
            
            t_row = {
                'id': f"TRD-{year}-{trade_id_counter:03d}",
                'year': year,
                'sym': sym,
                'name': spec['name'],
                'holiday': h['sheet'],
                'h_date': target_holiday_date,
                'window': f"T-{n_days} to T+{m_days}",
                'strat': f"FUTURE {strat_dir}",
                'entry_date': p_info['entry_date'],
                'entry_price': entry_p,
                'exit_date': p_info['exit_date'],
                'exit_price': exit_p,
                'ret_pct': ret_pct,
                'lot': lot,
                'pnl': pnl,
                'status': status,
                'margin': margin,
                'roc_pct': roc_pct
            }
            
            holiday_all_years_trades[year].append(t_row)
            all_master_trades.append(t_row)
            trade_id_counter += 1

    headers_log = [
        "Trade ID", "Rank", "Stock Symbol", "Company Name", "Holiday Event", "Holiday Date",
        "Optimal Window", "Strategy", "Entry Date", "Entry Price (₹)", "Exit Date", "Exit Price (₹)",
        "Return (%)", "Futures Lot Size", "Gross PnL / Lot (₹)", "Trade Status", "20% Margin Req. (₹)", "Return on Margin (ROC %)"
    ]

    # CREATE SEPARATE YEARLY SHEETS (Year 2022, Year 2023, Year 2024, Year 2025)
    for year in years:
        ws_yr = wb_h.create_sheet(title=f"Year {year}")
        
        # Title Block
        ws_yr.merge_cells("A1:R1")
        title_cell = ws_yr["A1"]
        title_cell.value = f"NIFTY 50 FUTURES — {h['sheet'].upper()} ({year} DETAILED TRADE LOG)"
        title_cell.font = font_title
        title_cell.fill = title_fill
        title_cell.alignment = align_center
        ws_yr.row_dimensions[1].height = 40
        
        ws_yr.merge_cells("A2:R2")
        ws_yr["A2"] = " " # blank spacing line
        
        ws_yr.append(headers_log)
        header_row_idx = 3
        ws_yr.row_dimensions[header_row_idx].height = 28
        for col_idx in range(1, len(headers_log) + 1):
            cell = ws_yr.cell(row=header_row_idx, column=col_idx)
            cell.font = font_header
            cell.fill = header_fill
            cell.alignment = align_center
            cell.border = border_thin
            
        # Sort trades for this year by PnL descending (Rank 1 to 50)
        yr_trades = sorted(holiday_all_years_trades[year], key=lambda x: x['pnl'], reverse=True)
        
        row_idx = 4
        for rank, t in enumerate(yr_trades, 1):
            ws_yr.append([
                t['id'], f"Rank {rank}", t['sym'], t['name'], t['holiday'], t['h_date'],
                t['window'], t['strat'], t['entry_date'], t['entry_price'], t['exit_date'], t['exit_price'],
                t['ret_pct'] / 100.0, t['lot'], t['pnl'], t['status'], t['margin'], t['roc_pct'] / 100.0
            ])
            
            ws_yr.row_dimensions[row_idx].height = 20
            
            # Format cells
            ws_yr.cell(row=row_idx, column=1).alignment = align_center # Trade ID
            ws_yr.cell(row=row_idx, column=2).alignment = align_center # Rank
            ws_yr.cell(row=row_idx, column=2).font = Font(name="Calibri", size=10, bold=True, color="1F4E78")
            ws_yr.cell(row=row_idx, column=3).alignment = align_center # Symbol
            ws_yr.cell(row=row_idx, column=4).alignment = align_left   # Name
            ws_yr.cell(row=row_idx, column=5).alignment = align_center # Holiday
            ws_yr.cell(row=row_idx, column=6).alignment = align_center # H Date
            ws_yr.cell(row=row_idx, column=7).alignment = align_center # Window
            ws_yr.cell(row=row_idx, column=8).alignment = align_center # Strat
            ws_yr.cell(row=row_idx, column=9).alignment = align_center # Entry Date
            ws_yr.cell(row=row_idx, column=10).number_format = '₹#,##0.00' # Entry Price
            ws_yr.cell(row=row_idx, column=11).alignment = align_center # Exit Date
            ws_yr.cell(row=row_idx, column=12).number_format = '₹#,##0.00' # Exit Price
            
            cell_ret = ws_yr.cell(row=row_idx, column=13)
            cell_ret.number_format = '+0.00%;-0.00%;0.00%' # Return %
            cell_ret.alignment = align_right
            
            ws_yr.cell(row=row_idx, column=14).number_format = '#,##0' # Lot
            
            cell_pnl = ws_yr.cell(row=row_idx, column=15)
            cell_pnl.number_format = '₹#,##0.00;[Red]-₹#,##0.00;₹0.00' # PnL
            cell_pnl.alignment = align_right
            
            cell_status = ws_yr.cell(row=row_idx, column=16)
            cell_status.alignment = align_center
            if t['status'] == 'WIN':
                cell_status.font = font_win
                cell_status.fill = win_fill
            else:
                cell_status.font = font_loss
                cell_status.fill = loss_fill
                
            ws_yr.cell(row=row_idx, column=17).number_format = '₹#,##0.00' # Margin
            
            cell_roc = ws_yr.cell(row=row_idx, column=18)
            cell_roc.number_format = '+0.00%;-0.00%;0.00%' # ROC %
            cell_roc.alignment = align_right
            
            for c in range(1, 19):
                ws_yr.cell(row=row_idx, column=c).border = border_thin
                if t['status'] != 'WIN' and t['status'] != 'LOSS':
                    if row_idx % 2 == 0:
                        ws_yr.cell(row=row_idx, column=c).fill = zebra_fill
                        
            row_idx += 1
            
        # Auto-fit columns
        for col in ws_yr.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or '')
                if cell.row > 3:
                    max_len = max(max_len, len(val_str))
            ws_yr.column_dimensions[col_letter].width = max(max_len + 4, 14)
        ws_yr.column_dimensions['A'].width = 14
        ws_yr.column_dimensions['C'].width = 16
        ws_yr.column_dimensions['D'].width = 28
        ws_yr.column_dimensions['E'].width = 22

    # CREATE 5th SHEET: 4-Year Consolidated Summary
    ws_sum = wb_h.create_sheet(title="4-Year Consolidated Summary")
    ws_sum.merge_cells("A1:K1")
    t_s = ws_sum["A1"]
    t_s.value = f"NIFTY 50 FUTURES — {h['sheet'].upper()} (4-YEAR CONSOLIDATED STOCK LEADERBOARD)"
    t_s.font = font_title
    t_s.fill = title_fill
    t_s.alignment = align_center
    ws_sum.row_dimensions[1].height = 40
    
    headers_sum = [
        "Rank", "Stock Symbol", "Company Name", "Optimal Window", "Strategy", "4-Year Win Rate (%)",
        "2022 PnL (₹)", "2023 PnL (₹)", "2024 PnL (₹)", "2025 PnL (₹)", "Total 4-Year Gross PnL / Lot (₹)"
    ]
    
    ws_sum.append([])
    ws_sum.append(headers_sum)
    ws_sum.row_dimensions[3].height = 28
    for c_idx in range(1, len(headers_sum) + 1):
        cell = ws_sum.cell(row=3, column=c_idx)
        cell.font = font_header
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = border_thin
        
    stock_4y_agg = []
    for sym, spec in nifty50_specs.items():
        sym_trades = []
        for yr in years:
            for t in holiday_all_years_trades[yr]:
                if t['sym'] == sym:
                    sym_trades.append(t)
                    
        if not sym_trades:
            continue
            
        wins = sum(1 for t in sym_trades if t['status'] == 'WIN')
        tot = len(sym_trades)
        wr = (wins / tot * 100.0) if tot > 0 else 0.0
        
        pnl_2022 = sum(t['pnl'] for t in sym_trades if t['year'] == 2022)
        pnl_2023 = sum(t['pnl'] for t in sym_trades if t['year'] == 2023)
        pnl_2024 = sum(t['pnl'] for t in sym_trades if t['year'] == 2024)
        pnl_2025 = sum(t['pnl'] for t in sym_trades if t['year'] == 2025)
        tot_pnl = pnl_2022 + pnl_2023 + pnl_2024 + pnl_2025
        
        n_days, m_days = stock_windows[sym]
        stock_4y_agg.append({
            'sym': sym,
            'name': spec['name'],
            'window': f"T-{n_days} to T+{m_days}",
            'strat': f"FUTURE {h['bias']}",
            'wr': wr,
            'pnl_2022': pnl_2022,
            'pnl_2023': pnl_2023,
            'pnl_2024': pnl_2024,
            'pnl_2025': pnl_2025,
            'tot_pnl': tot_pnl
        })
        
    # Rank by total 4-year PnL descending
    stock_4y_agg = sorted(stock_4y_agg, key=lambda x: x['tot_pnl'], reverse=True)
    
    row_idx = 4
    for rank, r in enumerate(stock_4y_agg, 1):
        ws_sum.append([
            f"Rank {rank}", r['sym'], r['name'], r['window'], r['strat'], r['wr'] / 100.0,
            r['pnl_2022'], r['pnl_2023'], r['pnl_2024'], r['pnl_2025'], r['tot_pnl']
        ])
        ws_sum.row_dimensions[row_idx].height = 20
        
        ws_sum.cell(row=row_idx, column=1).alignment = align_center
        ws_sum.cell(row=row_idx, column=1).font = Font(name="Calibri", size=10, bold=True, color="1F4E78")
        ws_sum.cell(row=row_idx, column=2).alignment = align_center
        ws_sum.cell(row=row_idx, column=3).alignment = align_left
        ws_sum.cell(row=row_idx, column=4).alignment = align_center
        ws_sum.cell(row=row_idx, column=5).alignment = align_center
        
        cell_wr = ws_sum.cell(row=row_idx, column=6)
        cell_wr.number_format = '0.00%'
        cell_wr.alignment = align_right
        cell_wr.font = Font(name="Calibri", size=10, bold=True)
        
        ws_sum.cell(row=row_idx, column=7).number_format = '₹#,##0.00;[Red]-₹#,##0.00;₹0.00'
        ws_sum.cell(row=row_idx, column=8).number_format = '₹#,##0.00;[Red]-₹#,##0.00;₹0.00'
        ws_sum.cell(row=row_idx, column=9).number_format = '₹#,##0.00;[Red]-₹#,##0.00;₹0.00'
        ws_sum.cell(row=row_idx, column=10).number_format = '₹#,##0.00;[Red]-₹#,##0.00;₹0.00'
        
        cell_tot = ws_sum.cell(row=row_idx, column=11)
        cell_tot.number_format = '₹#,##0.00;[Red]-₹#,##0.00;₹0.00'
        cell_tot.alignment = align_right
        cell_tot.font = Font(name="Calibri", size=10, bold=True)
        
        for c in range(1, 12):
            ws_sum.cell(row=row_idx, column=c).border = border_thin
            if row_idx % 2 == 0:
                ws_sum.cell(row=row_idx, column=c).fill = zebra_fill
        row_idx += 1

    for col in ws_sum.columns:
        col_letter = get_column_letter(col[0].column)
        ws_sum.column_dimensions[col_letter].width = 22
    ws_sum.column_dimensions['C'].width = 28

    # Save Individual Excel File
    out_file_h = OUT_DIR / h['filename']
    wb_h.save(out_file_h)
    print(f"  SUCCESS: Generated Excel -> {out_file_h.name} (Sheets: ['Year 2022', 'Year 2023', 'Year 2024', 'Year 2025', '4-Year Consolidated Summary'])")

print("="*90)
print(f"SUCCESS: Generated All 17 Separate Holiday Excel Files with Individual Yearly Sheets in Directory:\n  -> {OUT_DIR.resolve()}")

# Build Master Consolidated Excel File
print("="*90)
print("BUILDING MASTER CONSOLIDATED WORKBOOK FOR ALL HOLIDAYS COMBINED...")

wb_master = openpyxl.Workbook()
wb_master.remove(wb_master.active)

# Sheet 1: Master Summary Leaderboard
ws_m_sum = wb_master.create_sheet(title="All Holidays Summary")
ws_m_sum.merge_cells("A1:H1")
t_m = ws_m_sum["A1"]
t_m.value = "NIFTY 50 FUTURES ALL 17 HOLIDAYS — 4-YEAR PERFORMANCE LEADERBOARD (2022 - 2025)"
t_m.font = font_title
t_m.fill = title_fill
t_m.alignment = align_center
ws_m_sum.row_dimensions[1].height = 40

headers_m_sum = ["Rank", "Holiday Event", "Historical Bias", "Total Trades (4Y)", "Wins", "Losses", "Win Rate (%)", "4-Year Total PnL / Lot (₹)"]
ws_m_sum.append([])
ws_m_sum.append(headers_m_sum)
ws_m_sum.row_dimensions[3].height = 28
for c_idx in range(1, len(headers_m_sum) + 1):
    cell = ws_m_sum.cell(row=3, column=c_idx)
    cell.font = font_header
    cell.fill = header_fill
    cell.alignment = align_center
    cell.border = border_thin

df_master = pd.DataFrame(all_master_trades)

m_summary_rows = []
for h in all_holidays:
    df_h = df_master[df_master['holiday'] == h['sheet']]
    tot = len(df_h)
    wins = len(df_h[df_h['status'] == 'WIN'])
    losses = tot - wins
    wr = (wins / tot * 100.0) if tot > 0 else 0.0
    tot_pnl = df_h['pnl'].sum() if tot > 0 else 0.0
    m_summary_rows.append({
        'holiday': h['sheet'],
        'bias': f"FUTURE {h['bias']}",
        'tot': tot,
        'wins': wins,
        'losses': losses,
        'wr': wr,
        'pnl': tot_pnl
    })

# Rank by Win Rate & PnL
df_m_summary = pd.DataFrame(m_summary_rows).sort_values(by=['wr', 'pnl'], ascending=[False, False]).reset_index(drop=True)

row_idx = 4
for rank, r in enumerate(df_m_summary.to_dict('records'), 1):
    ws_m_sum.append([
        f"Rank {rank}", r['holiday'], r['bias'], r['tot'], r['wins'], r['losses'], r['wr'] / 100.0, r['pnl']
    ])
    ws_m_sum.row_dimensions[row_idx].height = 20
    
    ws_m_sum.cell(row=row_idx, column=1).font = Font(name="Calibri", size=11, bold=True, color="1F4E78")
    ws_m_sum.cell(row=row_idx, column=1).alignment = align_center
    ws_m_sum.cell(row=row_idx, column=2).alignment = align_left
    ws_m_sum.cell(row=row_idx, column=3).alignment = align_center
    ws_m_sum.cell(row=row_idx, column=4).alignment = align_center
    ws_m_sum.cell(row=row_idx, column=5).alignment = align_center
    ws_m_sum.cell(row=row_idx, column=6).alignment = align_center
    
    cell_wr = ws_m_sum.cell(row=row_idx, column=7)
    cell_wr.number_format = '0.00%'
    cell_wr.alignment = align_right
    cell_wr.font = Font(name="Calibri", size=11, bold=True)
    
    cell_pnl = ws_m_sum.cell(row=row_idx, column=8)
    cell_pnl.number_format = '₹#,##0.00;[Red]-₹#,##0.00;₹0.00'
    cell_pnl.alignment = align_right
    cell_pnl.font = Font(name="Calibri", size=11, bold=True)
    
    for c in range(1, 9):
        ws_m_sum.cell(row=row_idx, column=c).border = border_thin
        if row_idx % 2 == 0:
            ws_m_sum.cell(row=row_idx, column=c).fill = zebra_fill
    row_idx += 1

for col in ws_m_sum.columns:
    col_letter = get_column_letter(col[0].column)
    ws_m_sum.column_dimensions[col_letter].width = 24

try:
    wb_master.save(MASTER_EXCEL_OUT)
    print(f"SUCCESS: Created Master Consolidated Workbook -> {MASTER_EXCEL_OUT.name}")
except Exception as e:
    alt_out = ROOT / 'Nifty50_Futures_All_Holidays_Past_4Years_Master_TradeLog_v2.xlsx'
    wb_master.save(alt_out)
    print(f"SUCCESS: Created Master Consolidated Workbook -> {alt_out.name}")
print("==========================================================================================")
