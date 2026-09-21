import os
import glob
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import numpy as np

ROOT = r"D:\behaviour analysis"
REPORTS_DIR = os.path.join(ROOT, "12_Quarters_Reports")
OUTPUT_FILE = os.path.join(ROOT, "12_Quarters_Consolidated_Master.xlsx")
OUTPUT_FILE_ALT = os.path.join(REPORTS_DIR, "12_Quarters_Consolidated_Master.xlsx")

# 12 Quarters in strict chronological order with FY mapping (Q1 starts in April)
QUARTERS = [
    ("Q3_2023_24", "Q3 2023-24", "Oct 2023 - Dec 2023"),
    ("Q4_2023_24", "Q4 2023-24", "Jan 2024 - Mar 2024"),
    ("Q1_2024_25", "Q1 2024-25", "Apr 2024 - Jun 2024"),
    ("Q2_2024_25", "Q2 2024-25", "Jul 2024 - Sep 2024"),
    ("Q3_2024_25", "Q3 2024-25", "Oct 2024 - Dec 2024"),
    ("Q4_2024_25", "Q4 2024-25", "Jan 2025 - Mar 2025"),
    ("Q1_2025_26", "Q1 2025-26", "Apr 2025 - Jun 2025"),
    ("Q2_2025_26", "Q2 2025-26", "Jul 2025 - Sep 2025"),
    ("Q3_2025_26", "Q3 2025-26", "Oct 2025 - Dec 2025"),
    ("Q4_2025_26", "Q4 2025-26", "Jan 2026 - Mar 2026"),
    ("Q1_2026_27", "Q1 2026-27", "Apr 2026 - Jun 2026"),
    ("Q2_2026_27", "Q2 2026-27", "Jul 2026 - Sep 2026")
]

# Styling definitions
FONT_FAMILY = "Segoe UI"
HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")       # Dark Navy Blue
SUBHEADER_FILL = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")    # Medium Blue
ACCENT_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")       # Light Blue Tint
GREEN_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")        # Soft Green
RED_FILL = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")          # Soft Red
ZEBRA_FILL = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")

BORDER_THIN = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)
BORDER_HEADER = Border(
    left=Side(style='thin', color='FFFFFF'),
    right=Side(style='thin', color='FFFFFF'),
    top=Side(style='medium', color='1F4E78'),
    bottom=Side(style='medium', color='1F4E78')
)

def style_cell(cell, font_size=10, bold=False, color="000000", bg_fill=None, align="center", num_fmt=None):
    cell.font = Font(name=FONT_FAMILY, size=font_size, bold=bold, color=color)
    if bg_fill:
        cell.fill = bg_fill
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    cell.border = BORDER_THIN
    if num_fmt:
        cell.number_format = num_fmt

def main():
    print("Starting 12-Quarters Master Consolidation...")
    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active) # Remove default sheet
    
    quarter_summaries = []
    all_trades_master = []
    all_losing_trades_master = []

    # 1. Process Each Quarter
    for fn, q_name, period_desc in QUARTERS:
        file_path = os.path.join(REPORTS_DIR, f"{fn}_Combined_Best_Capital_Utilisation.xlsx")
        if not os.path.exists(file_path):
            print(f"Warning: {file_path} not found. Skipping...")
            continue
        
        print(f"Loading {q_name} from {file_path}...")
        wb_src = openpyxl.load_workbook(file_path, data_only=True)
        
        # Read Summary Sheet
        ws_sum = wb_src["Summary"]
        fund_util = ws_sum.cell(4, 2).value or ws_sum.cell(3, 2).value or 0.0
        final_val = ws_sum.cell(4, 3).value or ws_sum.cell(3, 3).value or 0.0
        net_prof = ws_sum.cell(4, 4).value or ws_sum.cell(3, 4).value or 0.0
        exp_ret = ws_sum.cell(4, 5).value or ws_sum.cell(3, 5).value or 0.0
        booked_ret = ws_sum.cell(4, 6).value or ws_sum.cell(3, 6).value or 0.0
        total_trades_cnt = ws_sum.cell(4, 7).value or ws_sum.cell(3, 7).value or 0
        total_reentries_cnt = ws_sum.cell(4, 8).value or ws_sum.cell(3, 8).value or 0
        
        # Read Detailed Trades
        df_trades = pd.read_excel(file_path, sheet_name="Detailed_Trades")
        
        # Read Losing Trades Verification
        df_losing = None
        if "Losing Trades Verification" in wb_src.sheetnames:
            try:
                df_losing = pd.read_excel(file_path, sheet_name="Losing Trades Verification")
            except Exception as e:
                print(f"Could not read losing trades for {q_name}: {e}")
        
        # Parse Trades Metrics
        ret_col = next((c for c in df_trades.columns if "Realised Return" in c), "Realised Return (%)")
        pnl_col = next((c for c in df_trades.columns if "Realised Profit" in c or "Profit/Loss" in c), "Realised Profit/Loss (Rs.)")
        strat_col = next((c for c in df_trades.columns if "Strategy" in c), "Strategy Type (Decided Move)")
        entry_price_col = next((c for c in df_trades.columns if "Entry Price" in c or "Buy Price" in c), "Entry Price (Rs.)")
        exit_price_col = next((c for c in df_trades.columns if "Exit Price" in c or "Sell Price" in c), "Exit Price (Rs.)")
        exp_col = next((c for c in df_trades.columns if "Expected Return" in c), "Expected Return (%)")
        
        winners = int((df_trades[ret_col] > 0).sum())
        losers = int((df_trades[ret_col] <= 0).sum())
        total_trades = len(df_trades)
        win_rate = (winners / total_trades) if total_trades > 0 else 0.0
        
        long_trades = int((df_trades[strat_col].astype(str).str.upper().str.contains("LONG|BUY")).sum())
        short_trades = int((df_trades[strat_col].astype(str).str.upper().str.contains("SHORT|SELL")).sum())
        
        long_winners = int(((df_trades[strat_col].astype(str).str.upper().str.contains("LONG|BUY")) & (df_trades[ret_col] > 0)).sum())
        short_winners = int(((df_trades[strat_col].astype(str).str.upper().str.contains("SHORT|SELL")) & (df_trades[ret_col] > 0)).sum())
        
        long_pnl = float(df_trades[df_trades[strat_col].astype(str).str.upper().str.contains("LONG|BUY")][pnl_col].sum())
        short_pnl = float(df_trades[df_trades[strat_col].astype(str).str.upper().str.contains("SHORT|SELL")][pnl_col].sum())
        
        # Track summary
        quarter_summaries.append({
            "Quarter": q_name,
            "Period": period_desc,
            "Fund Utilised (Rs.)": float(fund_util),
            "Final Value (Rs.)": float(final_val),
            "Net Profit (Rs.)": float(net_prof),
            "Expected Return (%)": float(exp_ret),
            "Booked Return (%)": float(booked_ret),
            "Total Trades": total_trades,
            "Winning Trades": winners,
            "Losing Trades": losers,
            "Win Rate (%)": win_rate,
            "Long Trades": long_trades,
            "Short Trades": short_trades,
            "Long Win Rate (%)": (long_winners / long_trades) if long_trades > 0 else 0.0,
            "Short Win Rate (%)": (short_winners / short_trades) if short_trades > 0 else 0.0,
            "Long Profit (Rs.)": long_pnl,
            "Short Profit (Rs.)": short_pnl,
            "Total Re-entries": int(total_reentries_cnt)
        })
        
        # Add to master trades
        for _, row in df_trades.iterrows():
            trade_dict = {
                "Quarter": q_name,
                "Trade No": row.get("Trade No", ""),
                "Symbol": row.get("Symbol", ""),
                "Company Name": row.get("Company Name", ""),
                "Sector": row.get("Sector", ""),
                "Strategy": str(row.get(strat_col, "")).strip().upper(),
                "Entry Date": str(row.get("Entry Date", ""))[:10],
                "Exit Date": str(row.get("Exit Date", ""))[:10],
                "Entry Price (Rs.)": float(row.get(entry_price_col, 0.0)),
                "Exit Price (Rs.)": float(row.get(exit_price_col, 0.0)),
                "Expected Return (%)": float(row.get(exp_col, 0.0)),
                "Realised Return (%)": float(row.get(ret_col, 0.0)),
                "Realised Profit/Loss (Rs.)": float(row.get(pnl_col, 0.0)),
                "Re-entry Type": str(row.get("Re-entry Type", "")),
                "Assigned Slot": str(row.get("Assigned Slot", ""))
            }
            all_trades_master.append(trade_dict)
            
        # Add to losing trades master
        if df_losing is not None and not df_losing.empty:
            # Check structure of losing trades sheet
            l_rows = list(df_losing.values)
            # Find header
            h_idx = 0
            for idx, r in enumerate(l_rows[:4]):
                if any(isinstance(v, str) and "Symbol" in v for v in r):
                    h_idx = idx
                    break
            l_headers = [str(h).strip() for h in l_rows[h_idx]]
            l_data = l_rows[h_idx+1:]
            df_l_clean = pd.DataFrame(l_data, columns=l_headers).dropna(subset=[l_headers[0]])
            for _, lrow in df_l_clean.iterrows():
                all_losing_trades_master.append({
                    "Quarter": q_name,
                    "Symbol": lrow.get("Symbol", ""),
                    "Company Name": lrow.get("Company Name", ""),
                    "Decided Strategy": lrow.get("Decided Move (Strategy)", lrow.get("Strategy", "")),
                    "Long Return (%)": lrow.get("Long Return (%)", ""),
                    "Short Return (%)": lrow.get("Short Return (%)", ""),
                    "Realised Loss (Rs.)": lrow.get("Realised Loss (Rs.)", ""),
                    "Verification Status": lrow.get("Verification Status", "Passed"),
                    "Rationale": lrow.get("Rationale / Explanation", lrow.get("Rationale", ""))
                })

        # --- CREATE INDIVIDUAL QUARTER SHEET ---
        ws_q = wb_out.create_sheet(title=q_name)
        ws_q.views.sheetView[0].showGridLines = True
        
        # Quarter Title Banner
        ws_q.merge_cells("A1:N1")
        title_c = ws_q.cell(1, 1, f"NIFTY 50 Event-Driven Drift Strategy — {q_name} ({period_desc})")
        style_cell(title_c, font_size=13, bold=True, color="FFFFFF", bg_fill=HEADER_FILL, align="center")
        ws_q.row_dimensions[1].height = 36
        
        # KPI Cards Banner
        ws_q.merge_cells("A3:C3"); ws_q.cell(3, 1, "QUARTER SUMMARY")
        style_cell(ws_q.cell(3, 1), font_size=11, bold=True, color="FFFFFF", bg_fill=SUBHEADER_FILL, align="center")
        
        kpi_headers = ["Quarter", "Fund Utilised", "Final Value", "Net Profit", "Booked Return", "Strategy Win Rate", "Total Trades", "Winning", "Losing", "Long Trades", "Short Trades", "Re-entries"]
        for c_idx, kh in enumerate(kpi_headers, 1):
            c = ws_q.cell(4, c_idx, kh)
            style_cell(c, font_size=9, bold=True, color="FFFFFF", bg_fill=SUBHEADER_FILL, align="center")
        ws_q.row_dimensions[4].height = 22
        
        kpi_vals = [
            q_name, float(fund_util), float(final_val), float(net_prof), float(booked_ret),
            win_rate, total_trades, winners, losers, long_trades, short_trades, int(total_reentries_cnt)
        ]
        for c_idx, kv in enumerate(kpi_vals, 1):
            c = ws_q.cell(5, c_idx, kv)
            fmt = '₹#,##0.00' if c_idx in [2, 3, 4] else ('0.00%' if c_idx in [5, 6] else ('#,##0' if c_idx >= 7 else None))
            bg = GREEN_FILL if (c_idx == 4 and kv > 0) or (c_idx == 6 and kv >= 0.6) else ACCENT_FILL
            style_cell(c, font_size=10, bold=True, bg_fill=bg, align="center", num_fmt=fmt)
        ws_q.row_dimensions[5].height = 24
        
        # Detailed Trades Section Header
        ws_q.cell(7, 1, "DETAILED TRADES EXECUTION (CHRONOLOGICAL)")
        ws_q.merge_cells("A7:O7")
        style_cell(ws_q.cell(7, 1), font_size=11, bold=True, color="FFFFFF", bg_fill=HEADER_FILL, align="left")
        ws_q.row_dimensions[7].height = 26
        
        trade_cols = [
            "Trade No", "Symbol", "Company Name", "Sector", "Strategy Type", 
            "Entry Date", "Exit Date", "Entry Price (Rs.)", "Exit Price (Rs.)", 
            "Expected Return (%)", "Realised Return (%)", "Realised P&L (Rs.)", 
            "Re-entry Type", "Assigned Slot"
        ]
        for c_idx, tc in enumerate(trade_cols, 1):
            c = ws_q.cell(8, c_idx, tc)
            style_cell(c, font_size=9, bold=True, color="FFFFFF", bg_fill=SUBHEADER_FILL, align="center")
        ws_q.row_dimensions[8].height = 22
        
        # Trade Rows
        for r_idx, row in df_trades.iterrows():
            row_num = 9 + r_idx
            ws_q.row_dimensions[row_num].height = 20
            ret_val = float(row.get(ret_col, 0.0))
            pnl_val = float(row.get(pnl_col, 0.0))
            strat_val = str(row.get(strat_col, "")).strip().upper()
            
            row_data = [
                row.get("Trade No", r_idx + 1),
                row.get("Symbol", ""),
                row.get("Company Name", ""),
                row.get("Sector", ""),
                strat_val,
                str(row.get("Entry Date", ""))[:10],
                str(row.get("Exit Date", ""))[:10],
                float(row.get(entry_price_col, 0.0)),
                float(row.get(exit_price_col, 0.0)),
                float(row.get(exp_col, 0.0)),
                ret_val,
                pnl_val,
                str(row.get("Re-entry Type", "")),
                str(row.get("Assigned Slot", ""))
            ]
            
            for c_idx, val in enumerate(row_data, 1):
                c = ws_q.cell(row_num, c_idx, val)
                fmt = None
                align = "center"
                if c_idx in [3, 4]:
                    align = "left"
                elif c_idx in [8, 9, 12]:
                    fmt = '₹#,##0.00'
                elif c_idx in [10, 11]:
                    fmt = '0.00%'
                
                # Highlight winners vs losers
                bg = None
                if c_idx in [11, 12]:
                    bg = GREEN_FILL if ret_val > 0 else RED_FILL
                elif c_idx == 5:
                    bg = ACCENT_FILL if strat_val == "LONG" else PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
                elif r_idx % 2 == 1:
                    bg = ZEBRA_FILL
                
                style_cell(c, font_size=9, bold=(c_idx in [2, 11, 12]), bg_fill=bg, align=align, num_fmt=fmt)

        # Auto-fit columns
        for col in ws_q.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or '')
                if len(val_str) > max_len and cell.row > 1:
                    max_len = len(val_str)
            ws_q.column_dimensions[col_letter].width = max(max_len + 3, 11)
        ws_q.column_dimensions["C"].width = 28 # Company Name
        ws_q.column_dimensions["D"].width = 18 # Sector

        wb_src.close()

    # =========================================================================
    # 2. CREATE MASTER EXECUTIVE SUMMARY SHEET (INSERT AT POSITION 0)
    # =========================================================================
    ws_exec = wb_out.create_sheet(title="12_Quarters_Master_Summary", index=0)
    ws_exec.views.sheetView[0].showGridLines = True
    
    # Title Banner
    ws_exec.merge_cells("A1:N1")
    tcell = ws_exec.cell(1, 1, "NIFTY 50 EVENT-DRIVEN MODEL: 12-QUARTERS COMPREHENSIVE WALK-FORWARD PERFORMANCE")
    style_cell(tcell, font_size=14, bold=True, color="FFFFFF", bg_fill=HEADER_FILL, align="center")
    ws_exec.row_dimensions[1].height = 40
    
    # Subtitle Info
    ws_exec.merge_cells("A2:N2")
    sub_c = ws_exec.cell(2, 1, "Walk-Forward In-Sample Training & Out-of-Sample Execution | Q1 Starts in April (Indian FY Cycle) | Long & Short Drift Optimization")
    style_cell(sub_c, font_size=10, bold=False, color="1F4E78", bg_fill=ACCENT_FILL, align="center")
    ws_exec.row_dimensions[2].height = 24
    
    # Global Summary KPI Cards (Rows 4-6)
    df_qsum = pd.DataFrame(quarter_summaries)
    tot_trades_all = int(df_qsum["Total Trades"].sum())
    tot_winners_all = int(df_qsum["Winning Trades"].sum())
    tot_losers_all = int(df_qsum["Losing Trades"].sum())
    tot_pnl_all = float(df_qsum["Net Profit (Rs.)"].sum())
    tot_capital_util_avg = float(df_qsum["Fund Utilised (Rs.)"].mean())
    tot_longs_all = int(df_qsum["Long Trades"].sum())
    tot_shorts_all = int(df_qsum["Short Trades"].sum())
    tot_win_rate_all = (tot_winners_all / tot_trades_all) if tot_trades_all > 0 else 0.0
    tot_long_pnl = float(df_qsum["Long Profit (Rs.)"].sum())
    tot_short_pnl = float(df_qsum["Short Profit (Rs.)"].sum())
    
    kpi_cards = [
        ("Total Trades Analyzed", f"{tot_trades_all:,}", "Across 12 Quarters"),
        ("Winning Trades", f"{tot_winners_all:,}", f"Win Rate: {tot_win_rate_all:.1%}"),
        ("Losing Trades", f"{tot_losers_all:,}", f"Loss Rate: {(1-tot_win_rate_all):.1%}"),
        ("Total Net Realised Profit", f"₹{tot_pnl_all:,.2f}", "Cumulative P&L"),
        ("Average Capital Utilised", f"₹{tot_capital_util_avg:,.2f}", "Per Quarter"),
        ("Long vs Short Ratio", f"{tot_longs_all} L / {tot_shorts_all} S", f"Long P&L: ₹{tot_long_pnl:,.0f} | Short: ₹{tot_short_pnl:,.0f}")
    ]
    
    # Render KPI Cards in a neat 2-row layout
    ws_exec.merge_cells("A4:N4")
    kpi_banner = ws_exec.cell(4, 1, "GLOBAL 12-QUARTERS EXECUTIVE KPI OVERVIEW")
    style_cell(kpi_banner, font_size=11, bold=True, color="FFFFFF", bg_fill=SUBHEADER_FILL, align="center")
    ws_exec.row_dimensions[4].height = 24
    
    col_offsets = [1, 3, 5, 7, 10, 12]
    widths = [2, 2, 2, 3, 2, 3]
    
    for idx, (title, val, subtitle) in enumerate(kpi_cards):
        c_start = col_offsets[idx]
        w = widths[idx]
        c_end = c_start + w - 1
        
        ws_exec.merge_cells(start_row=5, start_column=c_start, end_row=5, end_column=c_end)
        cell_t = ws_exec.cell(5, c_start, title)
        style_cell(cell_t, font_size=9, bold=True, color="595959", bg_fill=ACCENT_FILL, align="center")
        
        ws_exec.merge_cells(start_row=6, start_column=c_start, end_row=6, end_column=c_end)
        cell_v = ws_exec.cell(6, c_start, val)
        style_cell(cell_v, font_size=12, bold=True, color="1F4E78", bg_fill=GREEN_FILL if "Profit" in title or "Winning" in title else ACCENT_FILL, align="center")
        
        ws_exec.merge_cells(start_row=7, start_column=c_start, end_row=7, end_column=c_end)
        cell_s = ws_exec.cell(7, c_start, subtitle)
        style_cell(cell_s, font_size=8, bold=False, color="595959", bg_fill=ACCENT_FILL, align="center")
        
    ws_exec.row_dimensions[5].height = 18
    ws_exec.row_dimensions[6].height = 24
    ws_exec.row_dimensions[7].height = 16
    
    # 3. QUARTER-BY-QUARTER TABLE
    ws_exec.cell(9, 1, "QUARTER-BY-QUARTER CONSOLIDATED PERFORMANCE BREAKDOWN")
    ws_exec.merge_cells("A9:N9")
    style_cell(ws_exec.cell(9, 1), font_size=11, bold=True, color="FFFFFF", bg_fill=HEADER_FILL, align="left")
    ws_exec.row_dimensions[9].height = 26
    
    q_table_cols = [
        "Quarter", "Business Period (FY)", "Fund Utilised (Rs.)", "Final Value (Rs.)", 
        "Net Profit (Rs.)", "Booked Return (%)", "Expected Return (%)", 
        "Total Trades", "Winning", "Losing", "Win Rate (%)", 
        "Long Trades", "Short Trades", "Re-entries"
    ]
    for c_idx, qc in enumerate(q_table_cols, 1):
        c = ws_exec.cell(10, c_idx, qc)
        style_cell(c, font_size=9, bold=True, color="FFFFFF", bg_fill=SUBHEADER_FILL, align="center")
    ws_exec.row_dimensions[10].height = 24
    
    for r_idx, qrow in df_qsum.iterrows():
        row_num = 11 + r_idx
        ws_exec.row_dimensions[row_num].height = 21
        
        q_vals = [
            qrow["Quarter"],
            qrow["Period"],
            qrow["Fund Utilised (Rs.)"],
            qrow["Final Value (Rs.)"],
            qrow["Net Profit (Rs.)"],
            qrow["Booked Return (%)"],
            qrow["Expected Return (%)"],
            qrow["Total Trades"],
            qrow["Winning Trades"],
            qrow["Losing Trades"],
            qrow["Win Rate (%)"],
            qrow["Long Trades"],
            qrow["Short Trades"],
            qrow["Total Re-entries"]
        ]
        
        for c_idx, qv in enumerate(q_vals, 1):
            c = ws_exec.cell(row_num, c_idx, qv)
            fmt = None
            align = "center"
            if c_idx == 2:
                align = "left"
            elif c_idx in [3, 4, 5]:
                fmt = '₹#,##0.00'
            elif c_idx in [6, 7, 11]:
                fmt = '0.00%'
            elif c_idx in [8, 9, 10, 12, 13, 14]:
                fmt = '#,##0'
                
            bg = None
            if c_idx == 5:
                bg = GREEN_FILL if qv > 0 else RED_FILL
            elif c_idx == 11:
                bg = GREEN_FILL if qv >= 0.65 else (ACCENT_FILL if qv >= 0.55 else RED_FILL)
            elif r_idx % 2 == 1:
                bg = ZEBRA_FILL
                
            style_cell(c, font_size=9, bold=(c_idx in [1, 5, 11]), bg_fill=bg, align=align, num_fmt=fmt)
            
    # Add Total / Summary Row
    tot_row_num = 11 + len(df_qsum)
    ws_exec.row_dimensions[tot_row_num].height = 24
    tot_vals = [
        "12-Quarter Total / Avg",
        "FY 2023-24 to FY 2026-27",
        float(df_qsum["Fund Utilised (Rs.)"].sum()),
        float(df_qsum["Final Value (Rs.)"].sum()),
        tot_pnl_all,
        float(df_qsum["Booked Return (%)"].mean()),
        float(df_qsum["Expected Return (%)"].mean()),
        tot_trades_all,
        tot_winners_all,
        tot_losers_all,
        tot_win_rate_all,
        tot_longs_all,
        tot_shorts_all,
        int(df_qsum["Total Re-entries"].sum())
    ]
    for c_idx, tv in enumerate(tot_vals, 1):
        c = ws_exec.cell(tot_row_num, c_idx, tv)
        fmt = None
        align = "center"
        if c_idx in [3, 4, 5]:
            fmt = '₹#,##0.00'
        elif c_idx in [6, 7, 11]:
            fmt = '0.00%'
        elif c_idx >= 8:
            fmt = '#,##0'
        style_cell(c, font_size=10, bold=True, color="1F4E78", bg_fill=ACCENT_FILL, align=align, num_fmt=fmt)

    # Auto-fit executive sheet columns
    for col in ws_exec.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            if len(val_str) > max_len and cell.row > 2:
                max_len = len(val_str)
        ws_exec.column_dimensions[col_letter].width = max(max_len + 4, 13)
    ws_exec.column_dimensions["B"].width = 22

    # =========================================================================
    # 4. MASTER ALL TRADES SHEET
    # =========================================================================
    ws_master_tr = wb_out.create_sheet(title="All_12_Quarters_Trades")
    ws_master_tr.views.sheetView[0].showGridLines = True
    
    df_all_trades = pd.DataFrame(all_trades_master)
    mtr_headers = [
        "Quarter", "Trade No", "Symbol", "Company Name", "Sector", "Strategy", 
        "Entry Date", "Exit Date", "Entry Price (Rs.)", "Exit Price (Rs.)", 
        "Expected Return (%)", "Realised Return (%)", "Realised Profit/Loss (Rs.)", 
        "Re-entry Type", "Assigned Slot"
    ]
    
    # Header
    for c_idx, mh in enumerate(mtr_headers, 1):
        c = ws_master_tr.cell(1, c_idx, mh)
        style_cell(c, font_size=9, bold=True, color="FFFFFF", bg_fill=HEADER_FILL, align="center")
    ws_master_tr.row_dimensions[1].height = 25
    
    for r_idx, trow in df_all_trades.iterrows():
        row_num = 2 + r_idx
        ws_master_tr.row_dimensions[row_num].height = 19
        ret_v = float(trow["Realised Return (%)"])
        strat_v = str(trow["Strategy"]).strip().upper()
        
        row_data = [trow[h] for h in mtr_headers]
        for c_idx, val in enumerate(row_data, 1):
            c = ws_master_tr.cell(row_num, c_idx, val)
            fmt = None
            align = "center"
            if c_idx in [4, 5]:
                align = "left"
            elif c_idx in [9, 10, 13]:
                fmt = '₹#,##0.00'
            elif c_idx in [11, 12]:
                fmt = '0.00%'
                
            bg = None
            if c_idx in [12, 13]:
                bg = GREEN_FILL if ret_v > 0 else RED_FILL
            elif c_idx == 6:
                bg = ACCENT_FILL if strat_v == "LONG" else PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
            elif r_idx % 2 == 1:
                bg = ZEBRA_FILL
                
            style_cell(c, font_size=9, bold=(c_idx in [3, 12, 13]), bg_fill=bg, align=align, num_fmt=fmt)
            
    # Auto-fit master trades columns
    for col in ws_master_tr.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws_master_tr.column_dimensions[col_letter].width = max(max_len + 3, 12)
    ws_master_tr.column_dimensions["D"].width = 28
    ws_master_tr.column_dimensions["E"].width = 18

    # =========================================================================
    # 5. MASTER LOSING TRADES AUDIT SHEET
    # =========================================================================
    if all_losing_trades_master:
        ws_lose = wb_out.create_sheet(title="All_Losing_Trades_Audit")
        ws_lose.views.sheetView[0].showGridLines = True
        df_all_lose = pd.DataFrame(all_losing_trades_master)
        
        lose_headers = [
            "Quarter", "Symbol", "Company Name", "Decided Strategy", 
            "Long Return (%)", "Short Return (%)", "Realised Loss (Rs.)", 
            "Verification Status", "Rationale"
        ]
        for c_idx, lh in enumerate(lose_headers, 1):
            c = ws_lose.cell(1, c_idx, lh)
            style_cell(c, font_size=9, bold=True, color="FFFFFF", bg_fill=HEADER_FILL, align="center")
        ws_lose.row_dimensions[1].height = 25
        
        for r_idx, lrow in df_all_lose.iterrows():
            row_num = 2 + r_idx
            ws_lose.row_dimensions[row_num].height = 20
            row_data = [lrow.get(h, "") for h in lose_headers]
            for c_idx, val in enumerate(row_data, 1):
                c = ws_lose.cell(row_num, c_idx, val)
                align = "left" if c_idx in [3, 9] else "center"
                bg = ZEBRA_FILL if r_idx % 2 == 1 else None
                style_cell(c, font_size=9, bold=(c_idx in [2, 8]), bg_fill=bg, align=align)
                
        for col in ws_lose.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or '')
                if len(val_str) > max_len:
                    max_len = len(val_str)
            ws_lose.column_dimensions[col_letter].width = max(max_len + 3, 12)
        ws_lose.column_dimensions["C"].width = 28
        ws_lose.column_dimensions["I"].width = 65

    # Save to both destinations
    wb_out.save(OUTPUT_FILE)
    wb_out.save(OUTPUT_FILE_ALT)
    print(f"Successfully generated Master Consolidated Workbook:\n1. {OUTPUT_FILE}\n2. {OUTPUT_FILE_ALT}")

if __name__ == "__main__":
    main()
