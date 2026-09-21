import openpyxl
import pandas as pd
import numpy as np
import pathlib
import sys
import time
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference

sys.stdout.reconfigure(errors='replace')

t0 = time.time()

BASE_DIR = pathlib.Path('D:/behaviour analysis')
FUT_V13 = BASE_DIR / 'Nifty50_12_Quarters_Futures_OI_Master_v13.xlsx'
EQ_V2   = BASE_DIR / 'Nifty50_12_Quarters_Equity_Master_v2.xlsx'

COLOR_TITLE_BG     = '1B365D'
COLOR_SECTION_BG   = '2B4C7E'
COLOR_HEADER_BG    = '1A202C'
COLOR_SUBHEADER_BG = '2C5282'
COLOR_ZEBRA_BG     = 'F8FAFC'
COLOR_CARD_BG      = 'F7FAFC'

COLOR_WIN_BG       = 'E6F4EA'
COLOR_WIN_TEXT     = '137333'
COLOR_LOSS_BG      = 'FCE8E6'
COLOR_LOSS_TEXT    = 'C5221F'
COLOR_GOLD_BG      = 'FEF3C7'
COLOR_GOLD_TEXT    = '92400E'

COLOR_WHITE        = 'FFFFFF'
COLOR_TEXT_MAIN    = '2D3748'

font_title     = Font(name='Segoe UI', size=14, bold=True, color=COLOR_WHITE)
font_section   = Font(name='Segoe UI', size=11, bold=True, color=COLOR_WHITE)
font_header    = Font(name='Segoe UI', size=10, bold=True, color=COLOR_WHITE)
font_card_lbl  = Font(name='Segoe UI', size=9, bold=True, color=COLOR_WHITE)
font_data_bold = Font(name='Segoe UI', size=10, bold=True, color=COLOR_TEXT_MAIN)
font_data_reg  = Font(name='Segoe UI', size=10, color=COLOR_TEXT_MAIN)

font_win       = Font(name='Segoe UI', size=10, bold=True, color=COLOR_WIN_TEXT)
font_loss      = Font(name='Segoe UI', size=10, bold=True, color=COLOR_LOSS_TEXT)
font_gold      = Font(name='Segoe UI', size=10, bold=True, color=COLOR_GOLD_TEXT)

fill_title     = PatternFill(start_color=COLOR_TITLE_BG, end_color=COLOR_TITLE_BG, fill_type='solid')
fill_section   = PatternFill(start_color=COLOR_SECTION_BG, end_color=COLOR_SECTION_BG, fill_type='solid')
fill_header    = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type='solid')
fill_subheader = PatternFill(start_color=COLOR_SUBHEADER_BG, end_color=COLOR_SUBHEADER_BG, fill_type='solid')
fill_zebra     = PatternFill(start_color=COLOR_ZEBRA_BG, end_color=COLOR_ZEBRA_BG, fill_type='solid')
fill_card      = PatternFill(start_color=COLOR_CARD_BG, end_color=COLOR_CARD_BG, fill_type='solid')

fill_win       = PatternFill(start_color=COLOR_WIN_BG, end_color=COLOR_WIN_BG, fill_type='solid')
fill_loss       = PatternFill(start_color=COLOR_LOSS_BG, end_color=COLOR_LOSS_BG, fill_type='solid')
fill_gold      = PatternFill(start_color=COLOR_GOLD_BG, end_color=COLOR_GOLD_BG, fill_type='solid')

border_thin = Border(
    left=Side(style='thin', color='CBD5E0'),
    right=Side(style='thin', color='CBD5E0'),
    top=Side(style='thin', color='CBD5E0'),
    bottom=Side(style='thin', color='CBD5E0')
)

quarters_order = [
    'Q3 2023-24', 'Q4 2023-24', 'Q1 2024-25', 'Q2 2024-25',
    'Q3 2024-25', 'Q4 2024-25', 'Q1 2025-26', 'Q2 2025-26',
    'Q3 2025-26', 'Q4 2025-26', 'Q1 2026-27', 'Q2 2026-27'
]

print("Updating Futures Master v13 with True Peak Drawdowns...", flush=True)
wb_fut = openpyxl.load_workbook(FUT_V13)
xl_fut = pd.ExcelFile(FUT_V13, engine='openpyxl')
df_fut_trades = xl_fut.parse('All_12Q_Futures_Trades')

for qtr in quarters_order:
    ws = wb_fut[qtr]
    q_tr = df_fut_trades[df_fut_trades['Quarter'] == qtr].copy()
    t_no = next(c for c in q_tr.columns if 'Trade' in c)
    pnl_col = next(c for c in q_tr.columns if 'P&L' in c or 'Profit' in c)
    
    q_s = q_tr.sort_values(by=['Entry Date', t_no]).reset_index(drop=True)
    q_s['CumPnL'] = q_s[pnl_col].cumsum()
    q_s['Peak']   = np.maximum.accumulate(q_s['CumPnL'])
    q_s['DD']     = q_s['CumPnL'] - q_s['Peak']
    
    t_count = len(q_s)
    avg_slot_mrg = sum(r['Entry Futures Price (₹)'] * r['Lot Size (Qty)'] * 0.20 for _, r in q_s.iterrows()) / t_count
    tot_mrg = avg_slot_mrg * len(q_s['Assigned Slot'].unique())
    
    trough_idx = q_s['DD'].idxmin()
    max_dd_val = abs(q_s.loc[trough_idx, 'DD'])
    peak_cum_pnl = q_s.loc[:trough_idx, 'CumPnL'].max()
    peak_portfolio_val = tot_mrg + max(0.0, peak_cum_pnl)
    max_dd_pct_peak = -abs(max_dd_val / peak_portfolio_val) if peak_portfolio_val > 0 else 0.0
    
    # Update Cell K5 on quarter sheet
    c_mddp = ws.cell(row=5, column=11, value=max_dd_pct_peak)
    c_mddp.number_format = '-0.00%'

# Update Executive Summary Sheet Exec_12Q_Combined_Summary
ws_exec_fut = wb_fut['Exec_12Q_Combined_Summary']
for r_idx, qtr in enumerate(quarters_order, start=8):
    q_tr = df_fut_trades[df_fut_trades['Quarter'] == qtr].copy()
    t_no = next(c for c in q_tr.columns if 'Trade' in c)
    pnl_col = next(c for c in q_tr.columns if 'P&L' in c or 'Profit' in c)
    
    q_s = q_tr.sort_values(by=['Entry Date', t_no]).reset_index(drop=True)
    q_s['CumPnL'] = q_s[pnl_col].cumsum()
    q_s['Peak']   = np.maximum.accumulate(q_s['CumPnL'])
    q_s['DD']     = q_s['CumPnL'] - q_s['Peak']
    
    t_count = len(q_s)
    avg_slot_mrg = sum(r['Entry Futures Price (₹)'] * r['Lot Size (Qty)'] * 0.20 for _, r in q_s.iterrows()) / t_count
    tot_mrg = avg_slot_mrg * len(q_s['Assigned Slot'].unique())
    
    trough_idx = q_s['DD'].idxmin()
    max_dd_val = abs(q_s.loc[trough_idx, 'DD'])
    peak_cum_pnl = q_s.loc[:trough_idx, 'CumPnL'].max()
    peak_portfolio_val = tot_mrg + max(0.0, peak_cum_pnl)
    max_dd_pct_peak = -abs(max_dd_val / peak_portfolio_val) if peak_portfolio_val > 0 else 0.0
    
    c_ddp = ws_exec_fut.cell(row=r_idx, column=11, value=max_dd_pct_peak)
    c_ddp.number_format = '-0.00%'

def safe_save(wb_obj, target_path):
    try:
        wb_obj.save(target_path)
        print(f"✅ Successfully saved {target_path.name}")
    except PermissionError:
        alt_path = target_path.parent / (target_path.stem + '_v2' + target_path.suffix)
        wb_obj.save(alt_path)
        print(f"⚠️ Locked by Excel! Saved to {alt_path.name}")

safe_save(wb_fut, FUT_V13)

print(f"\n🎉 ALL WORKBOOKS UPDATED WITH TRUE PEAK DRAWDOWNS IN {time.time() - t0:.1f}s!", flush=True)
