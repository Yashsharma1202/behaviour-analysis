import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import numpy as np
from pathlib import Path

ROOT = Path(r"D:\behaviour analysis")
REPORTS_DIR = ROOT / "12_Quarters_Reports"

OUTPUT_LONG = ROOT / "Nifty50_12_Quarters_Long_Only_Master.xlsx"
OUTPUT_SHORT = ROOT / "Nifty50_12_Quarters_Short_Only_Master.xlsx"

OUTPUT_LONG_ALT = REPORTS_DIR / "Nifty50_12_Quarters_Long_Only_Master.xlsx"
OUTPUT_SHORT_ALT = REPORTS_DIR / "Nifty50_12_Quarters_Short_Only_Master.xlsx"

# 12 Quarters: Indian Financial Year Standard (Q1 starts in April)
QUARTERS = [
    ("Q3_2023_24", "Q3 2023-24", "Oct 2023 - Dec 2023", "FY 2023-24"),
    ("Q4_2023_24", "Q4 2023-24", "Jan 2024 - Mar 2024", "FY 2023-24"),
    ("Q1_2024_25", "Q1 2024-25", "Apr 2024 - Jun 2024", "FY 2024-25"),
    ("Q2_2024_25", "Q2 2024-25", "Jul 2024 - Sep 2024", "FY 2024-25"),
    ("Q3_2024_25", "Q3 2024-25", "Oct 2024 - Dec 2024", "FY 2024-25"),
    ("Q4_2024_25", "Q4 2024-25", "Jan 2025 - Mar 2025", "FY 2024-25"),
    ("Q1_2025_26", "Q1 2025-26", "Apr 2025 - Jun 2025", "FY 2025-26"),
    ("Q2_2025_26", "Q2 2025-26", "Jul 2025 - Sep 2025", "FY 2025-26"),
    ("Q3_2025_26", "Q3 2025-26", "Oct 2025 - Dec 2025", "FY 2025-26"),
    ("Q4_2025_26", "Q4 2025-26", "Jan 2026 - Mar 2026", "FY 2025-26"),
    ("Q1_2026_27", "Q1 2026-27", "Apr 2026 - Jun 2026", "FY 2026-27"),
    ("Q2_2026_27", "Q2 2026-27", "Jul 2026 - Sep 2026", "FY 2026-27")
]

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
    "HINDUNILVR": "Consumer Goods (FMCG)", "ITC": "Consumer Goods (FMCG)", "NESTLEIND": "Consumer Goods (FMCG)",
    "BRITANNIA": "Consumer Goods (FMCG)", "TATACONSUM": "Consumer Goods (FMCG)", "SUNPHARMA": "Healthcare & Pharma",
    "CIPLA": "Healthcare & Pharma", "DRREDDY": "Healthcare & Pharma", "APOLLOHOSP": "Healthcare & Pharma",
    "MAXHEALTH": "Healthcare & Pharma", "ASIANPAINT": "Consumer Durables", "TITAN": "Consumer Durables",
    "ADANIENT": "Diversified", "ADANIPORTS": "Services & Logistics", "BEL": "Capital Goods / Defense",
    "TRENT": "Retail & Services", "INDIGO": "Aviation / Transport", "TMPV": "Automobile & Auto",
    "ETERNAL": "Consumer Services", "RELIABLE": "Diversified"
}

COMPANIES = {
    "TCS": "Tata Consultancy Services Ltd", "HCLTECH": "HCL Technologies Ltd", "INFY": "Infosys Ltd",
    "WIPRO": "Wipro Ltd", "LTIM": "LTIMindtree Ltd", "TECHM": "Tech Mahindra Ltd",
    "HDFCBANK": "HDFC Bank Ltd", "ICICIBANK": "ICICI Bank Ltd", "KOTAKBANK": "Kotak Mahindra Bank Ltd",
    "AXISBANK": "Axis Bank Ltd", "SBIN": "State Bank of India", "JIOFIN": "Jio Financial Services Ltd",
    "BAJFINANCE": "Bajaj Finance Ltd", "BAJAJFINSV": "Bajaj Finserv Ltd", "HDFCLIFE": "HDFC Life Insurance Co Ltd",
    "SBILIFE": "SBI Life Insurance Co Ltd", "SHRIRAMFIN": "Shriram Finance Ltd", "RELIANCE": "Reliance Industries Ltd",
    "ONGC": "Oil & Natural Gas Corporation Ltd", "NTPC": "NTPC Ltd", "POWERGRID": "Power Grid Corp of India Ltd",
    "BPCL": "Bharat Petroleum Corp Ltd", "COALINDIA": "Coal India Ltd", "TATASTEEL": "Tata Steel Ltd",
    "HINDALCO": "Hindalco Industries Ltd", "JSWSTEEL": "JSW Steel Ltd", "ULTRACEMCO": "UltraTech Cement Ltd",
    "GRASIM": "Grasim Industries Ltd", "LT": "Larsen & Toubro Ltd", "M&M": "Mahindra & Mahindra Ltd",
    "MARUTI": "Maruti Suzuki India Ltd", "TATAMOTORS": "Tata Motors Ltd", "BAJAJ-AUTO": "Bajaj Auto Ltd",
    "HEROMOTOCO": "Hero MotoCorp Ltd", "EICHERMOT": "Hero MotoCorp Ltd", "HINDUNILVR": "Hindustan Unilever Ltd",
    "ITC": "ITC Ltd", "NESTLEIND": "Nestle India Ltd", "BRITANNIA": "Britannia Industries Ltd",
    "TATACONSUM": "Tata Consumer Products Ltd", "SUNPHARMA": "Sun Pharmaceutical Industries Ltd",
    "CIPLA": "Cipla Ltd", "DRREDDY": "Dr. Reddy's Laboratories Ltd", "APOLLOHOSP": "Apollo Hospitals Enterprise Ltd",
    "MAXHEALTH": "Max Healthcare Institute Ltd", "ASIANPAINT": "Asian Paints Ltd", "TITAN": "Titan Company Ltd",
    "ADANIENT": "Adani Enterprises Ltd", "ADANIPORTS": "Adani Ports & SEZ Ltd", "BEL": "Bharat Electronics Ltd",
    "TRENT": "Trent Ltd", "INDIGO": "InterGlobe Aviation Ltd", "TMPV": "Tata Motors Passenger Vehicles Ltd",
    "ETERNAL": "Eternal Limited", "RELIABLE": "Reliable Ventures India Ltd"
}

# Typography & Color Palette
FONT_FAMILY = "Segoe UI"
COLOR_PRIMARY_NAVY = "1A365D"
COLOR_NAVY = "1A365D"
COLOR_SECONDARY_BLUE = "2B6CB0"
COLOR_STEEL = "2B6CB0"
COLOR_ICE_BLUE = "EBF8FF"
COLOR_ICE = "EBF8FF"
COLOR_LIGHT_GRAY = "F7FAFC"
COLOR_ZEBRA = "F7FAFC"
COLOR_WHITE = "FFFFFF"

# Expected Return (Lavender / Purple)
COLOR_EXPECTED_BG = "EDE9FE"
COLOR_EXPECTED_TXT = "5B21B6"

# Realised Return (Emerald Mint vs Coral Crimson)
COLOR_WIN_BG = "DCFCE7"
COLOR_WIN_TXT = "15803D"
COLOR_LOSS_BG = "FEE2E2"
COLOR_LOSS_TXT = "B91C1C"

# Badges & Accents
COLOR_BADGE_BG = "EDF2F7"
COLOR_BADGE_TXT = "2D3748"

FILL_NAVY = PatternFill(start_color=COLOR_PRIMARY_NAVY, end_color=COLOR_PRIMARY_NAVY, fill_type="solid")
FILL_STEEL = PatternFill(start_color=COLOR_SECONDARY_BLUE, end_color=COLOR_SECONDARY_BLUE, fill_type="solid")
FILL_ICE = PatternFill(start_color=COLOR_ICE_BLUE, end_color=COLOR_ICE_BLUE, fill_type="solid")
FILL_ZEBRA = PatternFill(start_color=COLOR_LIGHT_GRAY, end_color=COLOR_LIGHT_GRAY, fill_type="solid")
FILL_EXPECTED = PatternFill(start_color=COLOR_EXPECTED_BG, end_color=COLOR_EXPECTED_BG, fill_type="solid")
FILL_WIN = PatternFill(start_color=COLOR_WIN_BG, end_color=COLOR_WIN_BG, fill_type="solid")
FILL_LOSS = PatternFill(start_color=COLOR_LOSS_BG, end_color=COLOR_LOSS_BG, fill_type="solid")
FILL_BADGE = PatternFill(start_color=COLOR_BADGE_BG, end_color=COLOR_BADGE_BG, fill_type="solid")

BORDER_THIN = Border(
    left=Side(style='thin', color='E2E8F0'),
    right=Side(style='thin', color='E2E8F0'),
    top=Side(style='thin', color='E2E8F0'),
    bottom=Side(style='thin', color='E2E8F0')
)

def style_cell(cell, size=9.5, bold=False, color="1A202C", bg=None, align="center", num_fmt=None):
    cell.font = Font(name=FONT_FAMILY, size=size, bold=bold, color=color)
    if bg:
        cell.fill = bg
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    cell.border = BORDER_THIN
    if num_fmt:
        cell.number_format = num_fmt

def schedule_slots(trades):
    # Sort chronologically
    trades.sort(key=lambda t: (t["Entry Date"], t["Exit Date"]))
    slots = []
    for t in trades:
        assigned = False
        for s in slots:
            if t["Entry Date"] >= s["last_exit"]:
                s["trades"].append(t)
                s["last_exit"] = t["Exit Date"]
                t["Assigned Slot"] = f"Slot {s['id']}"
                t["Re-entry Type"] = "Re-entry"
                assigned = True
                break
        if not assigned:
            slot_id = len(slots) + 1
            slots.append({
                "id": slot_id,
                "trades": [t],
                "last_exit": t["Exit Date"]
            })
            t["Assigned Slot"] = f"Slot {slot_id}"
            t["Re-entry Type"] = "First Entry"
            
    total_margin = sum(max(tr["Entry Price"] for tr in s["trades"]) for s in slots) if slots else 0.0
    total_profit = sum(t["Realised Profit/Loss (Rs.)"] for t in trades)
    final_value = total_margin + total_profit
    booked_ret = (total_profit / total_margin) if total_margin > 0 else 0.0
    total_reentries = sum(len(s["trades"]) - 1 for s in slots)
    
    return slots, total_margin, final_value, total_profit, booked_ret, total_reentries

def build_single_mode_master(mode, out_path, alt_path):
    """
    mode: "LONG" or "SHORT"
    """
    title_mode = "LONG-ONLY (BUYING)" if mode == "LONG" else "SHORT-ONLY (SELLING)"
    print(f"\n=======================================================")
    print(f"Building Master Workbook for: {title_mode}")
    print(f"Output File: {out_path}")
    print(f"=======================================================")
    
    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)
    
    quarter_summaries = []
    all_trades_master = []
    all_losers_master = []
    
    for fn, q_name, period_desc, fy_tag in QUARTERS:
        root_path = ROOT / f"{fn}_Combined_Best_Capital_Utilisation.xlsx"
        rep_path = REPORTS_DIR / f"{fn}_Combined_Best_Capital_Utilisation.xlsx"
        
        if fn == "Q2_2026_27":
            root_v5 = ROOT / "Q2_2026_27_Combined_Best_Capital_Utilisation_v5.xlsx"
            file_path = root_v5 if root_v5.exists() else (root_path if root_path.exists() else rep_path)
        else:
            file_path = root_path if root_path.exists() else rep_path
            
        if not file_path.exists():
            print(f"Warning: {file_path} not found")
            continue
            
        print(f"-> Processing {q_name} for {title_mode} from {file_path.name}...")
        df_src = pd.read_excel(file_path, sheet_name="Detailed_Trades").dropna(subset=["Symbol"]).copy()
        
        # Standardize columns
        entry_p_c = next((c for c in df_src.columns if "Entry Price" in c or "Buy Price" in c), "Entry Price (Rs.)")
        exit_p_c = next((c for c in df_src.columns if "Exit Price" in c or "Sell Price" in c), "Exit Price (Rs.)")
        exp_c = next((c for c in df_src.columns if "Expected Return" in c), "Expected Return (%)")
        offset_c = next((c for c in df_src.columns if "Offset" in c or "Position" in c or "Window" in c), None)
        
        trades_for_q = []
        for _, row in df_src.iterrows():
            sym = str(row["Symbol"]).strip().upper()
            comp = COMPANIES.get(sym, sym)
            sec = SECTORS.get(sym, "Diversified")
            
            en_date = str(row["Entry Date"])[:10]
            ex_date = str(row["Exit Date"])[:10]
            p_en = float(pd.to_numeric(row[entry_p_c], errors="coerce") or 0.0)
            p_ex = float(pd.to_numeric(row[exit_p_c], errors="coerce") or 0.0)
            exp_ret_val = float(pd.to_numeric(row[exp_c], errors="coerce") or 0.035)
            
            window_str = str(row[offset_c]) if offset_c and offset_c in df_src.columns and pd.notna(row[offset_c]) else "T-5 to T+5"
            
            # Compute Return & PnL strictly based on mode
            if p_en > 0:
                if mode == "LONG":
                    # Buy at entry, Sell at exit
                    ret_val = (p_ex - p_en) / p_en
                    pnl_val = (p_ex - p_en)
                else:
                    # Short: Sell at entry, Buy back at exit
                    ret_val = (p_en - p_ex) / p_en
                    pnl_val = (p_en - p_ex)
            else:
                ret_val = 0.0
                pnl_val = 0.0
                
            trades_for_q.append({
                "Symbol": sym,
                "Company Name": comp,
                "Sector": sec,
                "Strategy": mode,
                "Position Taking Window": window_str,
                "Entry Date": en_date,
                "Exit Date": ex_date,
                "Entry Price": p_en,
                "Exit Price": p_ex,
                "Expected Return (%)": exp_ret_val,
                "Realised Return (%)": ret_val,
                "Realised Profit/Loss (Rs.)": pnl_val
            })
            
        # Re-schedule slots for this pure strategy
        slots, tot_fund, final_val, net_prof, booked_ret, total_reentries = schedule_slots(trades_for_q)
        
        tot_cnt = len(trades_for_q)
        win_cnt = sum(1 for t in trades_for_q if t["Realised Return (%)"] > 0)
        lose_cnt = sum(1 for t in trades_for_q if t["Realised Return (%)"] <= 0)
        win_rate = (win_cnt / tot_cnt) if tot_cnt > 0 else 0.0
        avg_exp = np.mean([t["Expected Return (%)"] for t in trades_for_q]) if trades_for_q else 0.0
        
        q_sum = {
            "Quarter": q_name,
            "Period": period_desc,
            "FY": fy_tag,
            "Fund Utilised": tot_fund,
            "Final Value": final_val,
            "Net Profit": net_prof,
            "Booked Return": booked_ret,
            "Expected Return": avg_exp,
            "Total Trades": tot_cnt,
            "Winning Trades": win_cnt,
            "Losing Trades": lose_cnt,
            "Win Rate": win_rate,
            "Slots Utilised": len(slots),
            "Total Re-entries": total_reentries
        }
        quarter_summaries.append(q_sum)
        
        for t_idx, t in enumerate(trades_for_q):
            all_trades_master.append({**t, "Quarter": q_name, "FY": fy_tag, "Trade No": t_idx + 1})
            if t["Realised Return (%)"] <= 0:
                all_losers_master.append({
                    "Quarter": q_name,
                    "Symbol": t["Symbol"],
                    "Company Name": t["Company Name"],
                    "Sector": t["Sector"],
                    "Position Taking Window": t["Position Taking Window"],
                    "Entry Date": t["Entry Date"],
                    "Exit Date": t["Exit Date"],
                    "Entry Price": t["Entry Price"],
                    "Exit Price": t["Exit Price"],
                    "Realised Return (%)": t["Realised Return (%)"],
                    "Realised Loss (Rs.)": t["Realised Profit/Loss (Rs.)"]
                })
                
        # --- Create Quarter Sheet ---
        ws_q = wb_out.create_sheet(title=q_name)
        ws_q.views.sheetView[0].showGridLines = True
        
        # Title Banner
        ws_q.merge_cells("A1:N1")
        t_c = ws_q.cell(1, 1, f"NIFTY 50 {title_mode} STRATEGY — {q_name} ({period_desc})")
        style_cell(t_c, size=13, bold=True, color=COLOR_WHITE, bg=FILL_NAVY, align="center")
        ws_q.row_dimensions[1].height = 36
        
        # KPI Header
        ws_q.merge_cells("A3:N3")
        kpi_title = ws_q.cell(3, 1, f"QUARTER PERFORMANCE OVERVIEW & CAPITAL UTILISATION ({fy_tag})")
        style_cell(kpi_title, size=10, bold=True, color=COLOR_WHITE, bg=FILL_STEEL, align="left")
        ws_q.row_dimensions[3].height = 24
        
        kpi_headers = [
            "Fund Utilised", "Final Value", "Net Realised Profit", "Booked Return", 
            "Expected Return", "Win Rate", "Total Trades", "Winning Trades", 
            "Losing Trades", "Slots Deployed", "Re-entries"
        ]
        for idx, kh in enumerate(kpi_headers, 1):
            c = ws_q.cell(4, idx, kh)
            style_cell(c, size=9, bold=True, color=COLOR_WHITE, bg=FILL_STEEL, align="center")
        ws_q.row_dimensions[4].height = 22
        
        kpi_row_vals = [
            tot_fund, final_val, net_prof, booked_ret,
            avg_exp, win_rate, tot_cnt, win_cnt, lose_cnt, len(slots), total_reentries
        ]
        for idx, kv in enumerate(kpi_row_vals, 1):
            c = ws_q.cell(5, idx, kv)
            fmt = None
            if idx in [1, 2, 3]:
                fmt = '₹#,##0.00'
            elif idx in [4, 5, 6]:
                fmt = '0.00%'
            elif idx >= 7:
                fmt = '#,##0'
                
            bg_kpi = FILL_WIN if (idx == 3 and kv > 0) or (idx == 6 and kv >= 0.6) else (FILL_LOSS if (idx == 3 and kv < 0) else FILL_ICE)
            txt_color = COLOR_WIN_TXT if (idx == 3 and kv > 0) else (COLOR_LOSS_TXT if (idx == 3 and kv < 0) else "1A202C")
            style_cell(c, size=10, bold=True, color=txt_color, bg=bg_kpi, align="center", num_fmt=fmt)
        ws_q.row_dimensions[5].height = 25
        
        # Detailed Trades
        ws_q.cell(7, 1, f"CHRONOLOGICAL TRADE EXECUTION LOG ({title_mode})")
        ws_q.merge_cells("A7:N7")
        style_cell(ws_q.cell(7, 1), size=10.5, bold=True, color=COLOR_WHITE, bg=FILL_NAVY, align="left")
        ws_q.row_dimensions[7].height = 26
        
        trade_headers = [
            "Trade #", "Symbol", "Company Name", "Sector", 
            "Position Taking Window", "Entry Date", "Exit Date", "Entry Price", "Exit Price", 
            "Expected Ret (%)", "Realised Ret (%)", "Realised P&L (₹)", 
            "Slot ID", "Re-entry Type"
        ]
        for idx, th in enumerate(trade_headers, 1):
            c = ws_q.cell(8, idx, th)
            style_cell(c, size=9, bold=True, color=COLOR_WHITE, bg=FILL_STEEL, align="center")
        ws_q.row_dimensions[8].height = 24
        
        for r_idx, t in enumerate(trades_for_q):
            row_num = 9 + r_idx
            ws_q.row_dimensions[row_num].height = 20
            ret_val = t["Realised Return (%)"]
            pnl_val = t["Realised Profit/Loss (Rs.)"]
            
            row_vals = [
                r_idx + 1,
                t["Symbol"],
                t["Company Name"],
                t["Sector"],
                t["Position Taking Window"],
                t["Entry Date"],
                t["Exit Date"],
                t["Entry Price"],
                t["Exit Price"],
                t["Expected Return (%)"],
                ret_val,
                pnl_val,
                t["Assigned Slot"],
                t["Re-entry Type"]
            ]
            
            for c_idx, val in enumerate(row_vals, 1):
                c = ws_q.cell(row_num, c_idx, val)
                align = "center"
                fmt = None
                if c_idx in [3, 4]:
                    align = "left"
                elif c_idx in [8, 9, 12]:
                    fmt = '₹#,##0.00'
                elif c_idx in [10, 11]:
                    fmt = '0.00%'
                    
                bg = None
                txt_c = "1A202C"
                if c_idx == 10:
                    bg = FILL_EXPECTED
                    txt_c = COLOR_EXPECTED_TXT
                elif c_idx in [11, 12]:
                    bg = FILL_WIN if ret_val > 0 else FILL_LOSS
                    txt_c = COLOR_WIN_TXT if ret_val > 0 else COLOR_LOSS_TXT
                elif c_idx == 5:
                    bg = FILL_BADGE
                    txt_c = COLOR_BADGE_TXT
                elif r_idx % 2 == 1:
                    bg = FILL_ZEBRA
                    
                style_cell(c, size=9.5, bold=(c_idx in [2, 10, 11, 12]), color=txt_c, bg=bg, align=align, num_fmt=fmt)
                
        for col in ws_q.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or '')
                if len(val_str) > max_len and cell.row > 1:
                    max_len = len(val_str)
            ws_q.column_dimensions[col_letter].width = max(max_len + 3, 12)
        ws_q.column_dimensions["C"].width = 28
        ws_q.column_dimensions["D"].width = 20
        ws_q.column_dimensions["E"].width = 22

    # --- Executive Summary Sheet (Sheet 1) ---
    ws_exec = wb_out.create_sheet(title=f"Executive_{mode}_12Q_Summary", index=0)
    ws_exec.views.sheetView[0].showGridLines = True
    
    df_qall = pd.DataFrame(quarter_summaries)
    tot_trades_all = int(df_qall["Total Trades"].sum())
    tot_win_all = int(df_qall["Winning Trades"].sum())
    tot_lose_all = int(df_qall["Losing Trades"].sum())
    tot_win_rate_all = (tot_win_all / tot_trades_all) if tot_trades_all > 0 else 0.0
    tot_net_pnl_all = float(df_qall["Net Profit"].sum())
    avg_cap_all = float(df_qall["Fund Utilised"].mean())
    avg_booked_ret = float(df_qall["Booked Return"].mean())
    
    # Title Banner
    ws_exec.merge_cells("A1:M1")
    t_exec = ws_exec.cell(1, 1, f"NIFTY 50 {title_mode} STRATEGY — 12-QUARTERS EXECUTIVE MASTER DASHBOARD")
    style_cell(t_exec, size=14, bold=True, color=COLOR_WHITE, bg=FILL_NAVY, align="center")
    ws_exec.row_dimensions[1].height = 42
    
    # Subtitle Info (Indian FY Reminder)
    ws_exec.merge_cells("A2:M2")
    sub_exec = ws_exec.cell(2, 1, f"Indian Financial Year Standard (Q1 Starts in April) | 12 Rolling Quarters ({title_mode}) | 8x8 Parameter Timing")
    style_cell(sub_exec, size=10, bold=False, color=COLOR_NAVY, bg=FILL_ICE, align="center")
    ws_exec.row_dimensions[2].height = 24
    
    # KPI Scorecard Banner
    ws_exec.merge_cells("A4:M4")
    kpi_bar = ws_exec.cell(4, 1, "12-QUARTER GLOBAL PERFORMANCE SCORECARDS")
    style_cell(kpi_bar, size=11, bold=True, color=COLOR_WHITE, bg=FILL_STEEL, align="center")
    ws_exec.row_dimensions[4].height = 24
    
    scorecards = [
        ("Total Trades Analyzed", f"{tot_trades_all:,}", "Across 12 Quarters"),
        ("Global Win Rate", f"{tot_win_rate_all:.1%}", f"{tot_win_all} Wins / {tot_lose_all} Losses"),
        ("Total Net Realised Profit", f"₹{tot_net_pnl_all:,.2f}", "Cumulative Net P&L"),
        ("Avg Booked Return / Qtr", f"{avg_booked_ret:.2%}", "Quarterly Return on Capital"),
        ("Average Capital Deployed", f"₹{avg_cap_all:,.2f}", "Per Quarter")
    ]
    
    col_starts = [1, 3, 6, 9, 11]
    col_spans = [2, 3, 3, 2, 3]
    
    for idx, (title, main_val, sub_val) in enumerate(scorecards):
        c_s = col_starts[idx]
        w = col_spans[idx]
        c_e = c_s + w - 1
        
        ws_exec.merge_cells(start_row=5, start_column=c_s, end_row=5, end_column=c_e)
        c_t = ws_exec.cell(5, c_s, title)
        style_cell(c_t, size=9, bold=True, color="4A5568", bg=FILL_BADGE, align="center")
        
        ws_exec.merge_cells(start_row=6, start_column=c_s, end_row=6, end_column=c_e)
        c_m = ws_exec.cell(6, c_s, main_val)
        bg_card = FILL_WIN if ("Profit" in title and tot_net_pnl_all > 0) or ("Win" in title and tot_win_rate_all >= 0.6) else (FILL_LOSS if "Profit" in title and tot_net_pnl_all < 0 else FILL_ICE)
        txt_card = COLOR_WIN_TXT if ("Profit" in title and tot_net_pnl_all > 0) else (COLOR_LOSS_TXT if "Profit" in title and tot_net_pnl_all < 0 else COLOR_NAVY)
        style_cell(c_m, size=13, bold=True, color=txt_card, bg=bg_card, align="center")
        
        ws_exec.merge_cells(start_row=7, start_column=c_s, end_row=7, end_column=c_e)
        c_sub = ws_exec.cell(7, c_s, sub_val)
        style_cell(c_sub, size=8, bold=False, color="4A5568", bg=FILL_BADGE, align="center")
        
    ws_exec.row_dimensions[5].height = 18
    ws_exec.row_dimensions[6].height = 26
    ws_exec.row_dimensions[7].height = 16
    
    # Master Table
    ws_exec.cell(9, 1, "QUARTER-BY-QUARTER PERFORMANCE COMPARISON (ALL 12 QUARTERS)")
    ws_exec.merge_cells("A9:M9")
    style_cell(ws_exec.cell(9, 1), size=11, bold=True, color=COLOR_WHITE, bg=FILL_NAVY, align="left")
    ws_exec.row_dimensions[9].height = 26
    
    master_table_cols = [
        "Quarter", "Business Quarter (FY)", "Financial Year", "Fund Utilised (₹)", 
        "Final Value (₹)", "Net Realised Profit (₹)", "Booked Return (%)", "Expected Return (%)", 
        "Trades", "Wins", "Losses", "Win Rate (%)", "Re-entries"
    ]
    for idx, mc in enumerate(master_table_cols, 1):
        c = ws_exec.cell(10, idx, mc)
        style_cell(c, size=9, bold=True, color=COLOR_WHITE, bg=FILL_STEEL, align="center")
    ws_exec.row_dimensions[10].height = 24
    
    for r_idx, qrow in df_qall.iterrows():
        row_num = 11 + r_idx
        ws_exec.row_dimensions[row_num].height = 21
        pnl_q = qrow["Net Profit"]
        wr_q = qrow["Win Rate"]
        
        row_vals = [
            qrow["Quarter"],
            qrow["Period"],
            qrow["FY"],
            qrow["Fund Utilised"],
            qrow["Final Value"],
            pnl_q,
            qrow["Booked Return"],
            qrow["Expected Return"],
            qrow["Total Trades"],
            qrow["Winning Trades"],
            qrow["Losing Trades"],
            wr_q,
            qrow["Total Re-entries"]
        ]
        
        for c_idx, qv in enumerate(row_vals, 1):
            c = ws_exec.cell(row_num, c_idx, qv)
            align = "center"
            fmt = None
            if c_idx in [2, 3]:
                align = "left"
            elif c_idx in [4, 5, 6]:
                fmt = '₹#,##0.00'
            elif c_idx in [7, 8, 12]:
                fmt = '0.00%'
            elif c_idx in [9, 10, 11, 13]:
                fmt = '#,##0'
                
            bg = None
            txt_c = "1A202C"
            if c_idx == 6:
                bg = FILL_WIN if pnl_q > 0 else FILL_LOSS
                txt_c = COLOR_WIN_TXT if pnl_q > 0 else COLOR_LOSS_TXT
            elif c_idx == 7:
                bg = FILL_WIN if pnl_q > 0 else FILL_LOSS
                txt_c = COLOR_WIN_TXT if pnl_q > 0 else COLOR_LOSS_TXT
            elif c_idx == 8:
                bg = FILL_EXPECTED
                txt_c = COLOR_EXPECTED_TXT
            elif c_idx == 12:
                bg = FILL_WIN if wr_q >= 0.65 else (FILL_ICE if wr_q >= 0.55 else FILL_LOSS)
                txt_c = COLOR_WIN_TXT if wr_q >= 0.65 else ("1A202C" if wr_q >= 0.55 else COLOR_LOSS_TXT)
            elif r_idx % 2 == 1:
                bg = FILL_ZEBRA
                
            style_cell(c, size=9.5, bold=(c_idx in [1, 6, 7, 8, 12]), color=txt_c, bg=bg, align=align, num_fmt=fmt)
            
    # Total Row
    tot_row = 11 + len(df_qall)
    ws_exec.row_dimensions[tot_row].height = 25
    tot_row_vals = [
        "12-Quarter Total / Summary",
        "FY 2023-24 to FY 2026-27 (3 Full Years)",
        "3 Financial Years",
        float(df_qall["Fund Utilised"].sum()),
        float(df_qall["Final Value"].sum()),
        tot_net_pnl_all,
        avg_booked_ret,
        float(df_qall["Expected Return"].mean()),
        tot_trades_all,
        tot_win_all,
        tot_lose_all,
        tot_win_rate_all,
        int(df_qall["Total Re-entries"].sum())
    ]
    for c_idx, tv in enumerate(tot_row_vals, 1):
        c = ws_exec.cell(tot_row, c_idx, tv)
        align = "center"
        fmt = None
        if c_idx in [4, 5, 6]:
            fmt = '₹#,##0.00'
        elif c_idx in [7, 8, 12]:
            fmt = '0.00%'
        elif c_idx in [9, 10, 11, 13]:
            fmt = '#,##0'
        style_cell(c, size=10, bold=True, color=COLOR_NAVY, bg=FILL_ICE, align=align, num_fmt=fmt)

    for col in ws_exec.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            if len(val_str) > max_len and cell.row > 2:
                max_len = len(val_str)
        ws_exec.column_dimensions[col_letter].width = max(max_len + 3, 12)
    ws_exec.column_dimensions["B"].width = 24
    ws_exec.column_dimensions["C"].width = 16

    # --- Master Trades Sheet ---
    ws_all = wb_out.create_sheet(title=f"All_12_Quarters_{mode}_Trades")
    ws_all.views.sheetView[0].showGridLines = True
    
    df_m = pd.DataFrame(all_trades_master)
    tr_cols = [
        "Quarter", "FY", "Trade No", "Symbol", "Company Name", "Sector", 
        "Position Taking Window", "Entry Date", "Exit Date", 
        "Entry Price (₹)", "Exit Price (₹)", "Expected Return (%)", 
        "Realised Return (%)", "Realised P&L (₹)", "Assigned Slot", "Re-entry Type"
    ]
    for idx, tch in enumerate(tr_cols, 1):
        c = ws_all.cell(1, idx, tch)
        style_cell(c, size=9.5, bold=True, color=COLOR_WHITE, bg=FILL_NAVY, align="center")
    ws_all.row_dimensions[1].height = 26
    
    for r_idx, trow in df_m.iterrows():
        row_num = 2 + r_idx
        ws_all.row_dimensions[row_num].height = 19
        ret_v = trow["Realised Return (%)"]
        pnl_v = trow["Realised Profit/Loss (Rs.)"]
        
        row_vals = [
            trow["Quarter"],
            trow["FY"],
            trow["Trade No"],
            trow["Symbol"],
            trow["Company Name"],
            trow["Sector"],
            trow["Position Taking Window"],
            trow["Entry Date"],
            trow["Exit Date"],
            trow["Entry Price"],
            trow["Exit Price"],
            trow["Expected Return (%)"],
            ret_v,
            pnl_v,
            trow["Assigned Slot"],
            trow["Re-entry Type"]
        ]
        for c_idx, val in enumerate(row_vals, 1):
            c = ws_all.cell(row_num, c_idx, val)
            align = "center"
            fmt = None
            if c_idx in [5, 6]:
                align = "left"
            elif c_idx in [10, 11, 14]:
                fmt = '₹#,##0.00'
            elif c_idx in [12, 13]:
                fmt = '0.00%'
                
            bg = None
            txt_c = "1A202C"
            if c_idx == 12:
                bg = FILL_EXPECTED
                txt_c = COLOR_EXPECTED_TXT
            elif c_idx in [13, 14]:
                bg = FILL_WIN if ret_v > 0 else FILL_LOSS
                txt_c = COLOR_WIN_TXT if ret_v > 0 else COLOR_LOSS_TXT
            elif c_idx == 7:
                bg = FILL_BADGE
                txt_c = COLOR_BADGE_TXT
            elif r_idx % 2 == 1:
                bg = FILL_ZEBRA
                
            style_cell(c, size=9.5, bold=(c_idx in [4, 12, 13, 14]), color=txt_c, bg=bg, align=align, num_fmt=fmt)
            
    for col in ws_all.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws_all.column_dimensions[col_letter].width = max(max_len + 3, 12)
    ws_all.column_dimensions["E"].width = 28
    ws_all.column_dimensions["F"].width = 20
    ws_all.column_dimensions["G"].width = 22

    # Save to disk
    try:
        wb_out.save(out_path)
        print(f"Successfully saved: {out_path}")
    except PermissionError:
        alt_name = out_path.parent / (out_path.stem + "_v2" + out_path.suffix)
        wb_out.save(alt_name)
        print(f"Locked in Excel, saved to alternative: {alt_name}")
        
    try:
        wb_out.save(alt_path)
    except Exception:
        pass

def main():
    print("Building 12-Quarters Long-Only and Short-Only Master Workbooks...")
    
    # 1. Build Long-Only Master
    build_single_mode_master("LONG", OUTPUT_LONG, OUTPUT_LONG_ALT)
    
    # 2. Build Short-Only Master
    build_single_mode_master("SHORT", OUTPUT_SHORT, OUTPUT_SHORT_ALT)
    
    print("\nAll workbooks generated successfully!")

if __name__ == "__main__":
    main()
