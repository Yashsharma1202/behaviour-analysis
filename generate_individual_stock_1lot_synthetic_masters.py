import calendar
import datetime as dt
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import numpy as np
from pathlib import Path

ROOT = Path(r"D:\behaviour analysis")
REPORTS_DIR = ROOT / "12_Quarters_Reports"

OUT_SYN_LONG = ROOT / "Nifty50_12_Quarters_Synthetic_Long_1Lot_Master.xlsx"
OUT_SYN_SHORT = ROOT / "Nifty50_12_Quarters_Synthetic_Short_1Lot_Master.xlsx"

OUT_SYN_LONG_REP = REPORTS_DIR / "Nifty50_12_Quarters_Synthetic_Long_1Lot_Master.xlsx"
OUT_SYN_SHORT_REP = REPORTS_DIR / "Nifty50_12_Quarters_Synthetic_Short_1Lot_Master.xlsx"

# 12 Quarters: Indian Financial Year Standard (Q1 strictly starts in April)
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

COMPANIES = {
    "ADANIENT": "Adani Enterprises Ltd", "ADANIPORTS": "Adani Ports and Special Economic Zone Ltd",
    "APOLLOHOSP": "Apollo Hospitals Enterprise Ltd", "ASIANPAINT": "Asian Paints Ltd",
    "AXISBANK": "Axis Bank Ltd", "BAJAJ-AUTO": "Bajaj Auto Ltd", "BAJAJFINSV": "Bajaj Finserv Ltd",
    "BAJFINANCE": "Bajaj Finance Ltd", "BEL": "Bharat Electronics Ltd", "BHARTIARTL": "Bharti Airtel Ltd",
    "CIPLA": "Cipla Ltd", "COALINDIA": "Coal India Ltd", "DRREDDY": "Dr. Reddy's Laboratories Ltd",
    "EICHERMOT": "Eicher Motors Ltd", "ETERNAL": "Eternal Limited", "GRASIM": "Grasim Industries Ltd",
    "HCLTECH": "HCL Technologies Ltd", "HDFCBANK": "HDFC Bank Ltd", "HDFCLIFE": "HDFC Life Insurance Co Ltd",
    "HEROMOTOCO": "Hero MotoCorp Ltd", "HINDALCO": "Hindalco Industries Ltd", "HINDUNILVR": "Hindustan Unilever Ltd",
    "ICICIBANK": "ICICI Bank Ltd", "INDIGO": "InterGlobe Aviation Ltd", "INDUSINDBK": "IndusInd Bank Ltd",
    "INFY": "Infosys Ltd", "ITC": "ITC Ltd", "JIOFIN": "Jio Financial Services Ltd",
    "JSWSTEEL": "JSW Steel Ltd", "KOTAKBANK": "Kotak Mahindra Bank Ltd", "LT": "Larsen & Toubro Ltd",
    "M&M": "Mahindra & Mahindra Ltd", "MARUTI": "Maruti Suzuki India Ltd", "MAXHEALTH": "Max Healthcare Institute Ltd",
    "NESTLEIND": "Nestle India Ltd", "NTPC": "NTPC Ltd", "ONGC": "Oil & Natural Gas Corporation Ltd",
    "POWERGRID": "Power Grid Corporation of India Ltd", "RELIANCE": "Reliance Industries Ltd",
    "SBILIFE": "SBI Life Insurance Co Ltd", "SBIN": "State Bank of India", "SHRIRAMFIN": "Shriram Finance Ltd",
    "SUNPHARMA": "Sun Pharmaceutical Industries Ltd", "TATACONSUM": "Tata Consumer Products Ltd",
    "TATASTEEL": "Tata Steel Ltd", "TCS": "Tata Consultancy Services Ltd", "TECHM": "Tech Mahindra Ltd",
    "TITAN": "Titan Company Ltd", "TMPV": "Tata Motors Passenger Vehicles Ltd", "TRENT": "Trent Ltd",
    "ULTRACEMCO": "UltraTech Cement Ltd", "WIPRO": "Wipro Ltd"
}

SECTORS = {
    "ADANIENT": "Diversified", "ADANIPORTS": "Services & Logistics", "APOLLOHOSP": "Healthcare & Pharma",
    "ASIANPAINT": "Consumer Durables", "AXISBANK": "Financial Services", "BAJAJ-AUTO": "Automobile & Auto",
    "BAJAJFINSV": "Financial Services", "BAJFINANCE": "Financial Services", "BEL": "Capital Goods / Defense",
    "BHARTIARTL": "Telecommunication", "CIPLA": "Healthcare & Pharma", "COALINDIA": "Energy / Mining",
    "DRREDDY": "Healthcare & Pharma", "EICHERMOT": "Automobile & Auto", "ETERNAL": "Consumer Services",
    "GRASIM": "Construction Materials", "HCLTECH": "Information Technology", "HDFCBANK": "Financial Services",
    "HDFCLIFE": "Financial Services", "HEROMOTOCO": "Automobile & Auto", "HINDALCO": "Metals & Mining",
    "HINDUNILVR": "Fast Moving Consumer Goods", "ICICIBANK": "Financial Services", "INDIGO": "Aviation / Transport",
    "INDUSINDBK": "Financial Services", "INFY": "Information Technology", "ITC": "Fast Moving Consumer Goods",
    "JIOFIN": "Financial Services", "JSWSTEEL": "Metals & Mining", "KOTAKBANK": "Financial Services",
    "LT": "Construction", "M&M": "Automobile & Auto", "MARUTI": "Automobile & Auto",
    "MAXHEALTH": "Healthcare & Pharma", "NESTLEIND": "Fast Moving Consumer Goods", "NTPC": "Energy & Utilities",
    "ONGC": "Energy / Oil & Gas", "POWERGRID": "Energy & Utilities", "RELIANCE": "Energy & Conglomerate",
    "SBILIFE": "Financial Services", "SBIN": "Financial Services", "SHRIRAMFIN": "Financial Services",
    "SUNPHARMA": "Healthcare & Pharma", "TATACONSUM": "Fast Moving Consumer Goods", "TATASTEEL": "Metals & Mining",
    "TCS": "Information Technology", "TECHM": "Information Technology", "TITAN": "Consumer Durables",
    "TMPV": "Automobile & Auto", "TRENT": "Retail & Services", "ULTRACEMCO": "Construction Materials",
    "WIPRO": "Information Technology"
}

LOT_SIZES = {
    "ADANIENT": 300, "ADANIPORTS": 400, "APOLLOHOSP": 125, "ASIANPAINT": 200, "AXISBANK": 625,
    "BAJAJ-AUTO": 125, "BAJAJFINSV": 500, "BAJFINANCE": 125, "BEL": 1500, "BHARTIARTL": 475,
    "CIPLA": 650, "COALINDIA": 2100, "DRREDDY": 125, "EICHERMOT": 175, "ETERNAL": 500,
    "GRASIM": 250, "HCLTECH": 350, "HDFCBANK": 550, "HDFCLIFE": 1100, "HEROMOTOCO": 150,
    "HINDALCO": 1400, "HINDUNILVR": 300, "ICICIBANK": 700, "INDIGO": 150, "INDUSINDBK": 500,
    "INFY": 400, "ITC": 1600, "JIOFIN": 2000, "JSWSTEEL": 675, "KOTAKBANK": 400,
    "LT": 150, "M&M": 350, "MARUTI": 50, "MAXHEALTH": 500, "NESTLEIND": 40,
    "NTPC": 1500, "ONGC": 2250, "POWERGRID": 1800, "RELIANCE": 250, "SBILIFE": 750,
    "SBIN": 750, "SHRIRAMFIN": 300, "SUNPHARMA": 350, "TATACONSUM": 900, "TATASTEEL": 5500,
    "TCS": 175, "TECHM": 600, "TITAN": 175, "TMPV": 500, "TRENT": 100,
    "ULTRACEMCO": 100, "WIPRO": 1500
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

def get_monthly_expiry(year, month):
    c = calendar.monthcalendar(year, month)
    thursdays = [week[calendar.THURSDAY] for week in c if week[calendar.THURSDAY] != 0]
    last_thursday = thursdays[-1]
    return dt.date(year, month, last_thursday)

def get_atm_strike(price):
    if price > 5000:
        step = 100
    elif price > 1500:
        step = 50
    elif price > 500:
        step = 20
    elif price > 200:
        step = 10
    else:
        step = 5
    return int(round(price / step) * step)

def schedule_1lot_slots(trades):
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
            
    total_margin = sum(max(tr["1 Lot Synthetic Margin (₹)"] for tr in s["trades"]) for s in slots) if slots else 0.0
    total_profit = sum(t["1 Lot Booked P&L (₹)"] for t in trades)
    final_value = total_margin + total_profit
    booked_ret = (total_profit / total_margin) if total_margin > 0 else 0.0
    total_reentries = sum(len(s["trades"]) - 1 for s in slots)
    
    return slots, total_margin, final_value, total_profit, booked_ret, total_reentries

def build_synthetic_1lot_master(mode, out_path, alt_path):
    title_mode = "SYNTHETIC LONG (1 LOT ATM BUY)" if mode == "LONG" else "SYNTHETIC SHORT (1 LOT ATM SELL)"
    print(f"\n=======================================================")
    print(f"Building 1-Lot Individual Stock Master: {title_mode}")
    print(f"Target File: {out_path}")
    print(f"=======================================================")
    
    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)
    
    quarter_summaries = []
    all_trades_master = []
    
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
        
        entry_p_c = next((c for c in df_src.columns if "Entry Price" in c or "Buy Price" in c), "Entry Price (Rs.)")
        exit_p_c = next((c for c in df_src.columns if "Exit Price" in c or "Sell Price" in c), "Exit Price (Rs.)")
        exp_c = next((c for c in df_src.columns if "Expected Return" in c), "Expected Return (%)")
        offset_c = next((c for c in df_src.columns if "Offset" in c or "Position" in c or "Window" in c), None)
        
        trades_for_q = []
        for _, row in df_src.iterrows():
            sym = str(row["Symbol"]).strip().upper()
            comp = COMPANIES.get(sym, sym)
            sec = SECTORS.get(sym, "Diversified")
            lot_size = LOT_SIZES.get(sym, 250)
            
            en_date_str = str(row["Entry Date"])[:10]
            ex_date_str = str(row["Exit Date"])[:10]
            en_dt = pd.to_datetime(en_date_str).date()
            ex_dt = pd.to_datetime(ex_date_str).date()
            
            p_en = float(pd.to_numeric(row[entry_p_c], errors="coerce") or 0.0)
            p_ex = float(pd.to_numeric(row[exit_p_c], errors="coerce") or 0.0)
            exp_ret_val = float(pd.to_numeric(row[exp_c], errors="coerce") or 0.035)
            
            window_str = str(row[offset_c]) if offset_c and offset_c in df_src.columns and pd.notna(row[offset_c]) else "T-5 to T+5"
            
            # Expiry Selection: If exit extends past current month expiry, rollover to next month expiry
            curr_expiry = get_monthly_expiry(en_dt.year, en_dt.month)
            if ex_dt <= curr_expiry:
                assigned_expiry = curr_expiry
                expiry_type = "Current Month Expiry"
                expiry_month_str = curr_expiry.strftime("%b %Y").upper()
            else:
                next_m = en_dt.month + 1 if en_dt.month < 12 else 1
                next_y = en_dt.year if en_dt.month < 12 else en_dt.year + 1
                next_expiry = get_monthly_expiry(next_y, next_m)
                assigned_expiry = next_expiry
                expiry_type = "Next Month Expiry (Multi-Month Rollover)"
                expiry_month_str = next_expiry.strftime("%b %Y").upper()
                
            atm_strike = get_atm_strike(p_en)
            
            # 1-Lot Margin Calculation: 20% SPAN + Exposure Margin
            contract_value_1lot = p_en * lot_size
            syn_margin_1lot = contract_value_1lot * 0.20
            
            if mode == "LONG":
                # Buy 1 Lot ATM Call + Sell 1 Lot ATM Put (Delta = +1.0)
                pnl_per_share = p_ex - p_en
                underlying_ret = (p_ex - p_en) / p_en if p_en > 0 else 0.0
                instr_desc = f"1 Lot {atm_strike} CE (Buy) + {atm_strike} PE (Sell)"
            else:
                # Buy 1 Lot ATM Put + Sell 1 Lot ATM Call (Delta = -1.0)
                pnl_per_share = p_en - p_ex
                underlying_ret = (p_en - p_ex) / p_en if p_en > 0 else 0.0
                instr_desc = f"1 Lot {atm_strike} PE (Buy) + {atm_strike} CE (Sell)"
                
            total_1lot_pnl = pnl_per_share * lot_size
            booked_ret_1lot_margin = (total_1lot_pnl / syn_margin_1lot) if syn_margin_1lot > 0 else 0.0
            
            trades_for_q.append({
                "Symbol": sym,
                "Company Name": comp,
                "Sector": sec,
                "Strategy": f"SYNTHETIC {mode}",
                "1 Lot Synthetic Pair": instr_desc,
                "ATM Strike (₹)": atm_strike,
                "Contract Expiry": assigned_expiry.strftime("%Y-%m-%d"),
                "Expiry Month": expiry_month_str,
                "Expiry Decision": expiry_type,
                "Position Taking Window": window_str,
                "Entry Date": en_date_str,
                "Exit Date": ex_date_str,
                "Entry Spot Price (₹)": p_en,
                "Exit Spot Price (₹)": p_ex,
                "Lot Size (Qty)": lot_size,
                "1 Lot Synthetic Margin (₹)": syn_margin_1lot,
                "P&L Per Share (₹)": pnl_per_share,
                "1 Lot Booked P&L (₹)": total_1lot_pnl,
                "Booked Return on 1 Lot Margin (%)": booked_ret_1lot_margin,
                "Expected Return on Margin (%)": exp_ret_val * 5.0
            })
            
        # Re-schedule slots based on 1-Lot Synthetic Margin
        slots, tot_margin, final_val, net_prof, booked_ret, total_reentries = schedule_1lot_slots(trades_for_q)
        
        tot_cnt = len(trades_for_q)
        win_cnt = sum(1 for t in trades_for_q if t["1 Lot Booked P&L (₹)"] > 0)
        lose_cnt = sum(1 for t in trades_for_q if t["1 Lot Booked P&L (₹)"] <= 0)
        win_rate = (win_cnt / tot_cnt) if tot_cnt > 0 else 0.0
        avg_exp = np.mean([t["Expected Return on Margin (%)"] for t in trades_for_q]) if trades_for_q else 0.0
        
        q_sum = {
            "Quarter": q_name,
            "Period": period_desc,
            "FY": fy_tag,
            "1 Lot Margin Deployed (₹)": tot_margin,
            "Final Value (₹)": final_val,
            "Net Booked Profit (₹)": net_prof,
            "Booked Return on Margin (%)": booked_ret,
            "Expected Return on Margin (%)": avg_exp,
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
            
        # --- Create Quarter Sheet ---
        ws_q = wb_out.create_sheet(title=q_name)
        ws_q.views.sheetView[0].showGridLines = True
        
        # Title Banner
        ws_q.merge_cells("A1:S1")
        t_c = ws_q.cell(1, 1, f"NIFTY 50 {title_mode} (INDIVIDUAL STOCK 1-LOT BASIS) — {q_name} ({period_desc})")
        style_cell(t_c, size=13, bold=True, color=COLOR_WHITE, bg=FILL_NAVY, align="center")
        ws_q.row_dimensions[1].height = 36
        
        # KPI Header
        ws_q.merge_cells("A3:S3")
        kpi_title = ws_q.cell(3, 1, f"QUARTER 1-LOT OPTIONS PERFORMANCE & SYNTHETIC MARGIN UTILISATION ({fy_tag})")
        style_cell(kpi_title, size=10, bold=True, color=COLOR_WHITE, bg=FILL_STEEL, align="left")
        ws_q.row_dimensions[3].height = 24
        
        kpi_headers = [
            "1-Lot Margin Deployed", "Final Value", "Net Booked Profit (₹)", "Booked Return on Margin", 
            "Expected Ret on Margin", "Win Rate", "Total Trades", "Winning Trades", 
            "Losing Trades", "Slots Deployed", "Re-entries"
        ]
        for idx, kh in enumerate(kpi_headers, 1):
            c = ws_q.cell(4, idx, kh)
            style_cell(c, size=9, bold=True, color=COLOR_WHITE, bg=FILL_STEEL, align="center")
        ws_q.row_dimensions[4].height = 22
        
        kpi_row_vals = [
            tot_margin, final_val, net_prof, booked_ret,
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
        
        # Detailed Trades Section
        ws_q.cell(7, 1, f"INDIVIDUAL STOCK 1-LOT TRADE LOG ({title_mode})")
        ws_q.merge_cells("A7:S7")
        style_cell(ws_q.cell(7, 1), size=10.5, bold=True, color=COLOR_WHITE, bg=FILL_NAVY, align="left")
        ws_q.row_dimensions[7].height = 26
        
        trade_headers = [
            "Trade #", "Symbol", "Company Name", "Sector", "1 Lot Synthetic Pair", "ATM Strike (₹)",
            "Contract Expiry", "Expiry Decision", "Position Taking Window", 
            "Entry Date", "Exit Date", "Entry Spot (₹)", "Exit Spot (₹)", "Lot Size (Qty)", 
            "1 Lot Margin (₹)", "P&L Per Share (₹)", "1 Lot Booked P&L (₹)", "Booked Return on Margin (%)", 
            "Assigned Slot", "Re-entry Type"
        ]
        for idx, th in enumerate(trade_headers, 1):
            c = ws_q.cell(8, idx, th)
            style_cell(c, size=9, bold=True, color=COLOR_WHITE, bg=FILL_STEEL, align="center")
        ws_q.row_dimensions[8].height = 24
        
        for r_idx, t in enumerate(trades_for_q):
            row_num = 9 + r_idx
            ws_q.row_dimensions[row_num].height = 20
            pnl_1lot = t["1 Lot Booked P&L (₹)"]
            ret_1lot = t["Booked Return on 1 Lot Margin (%)"]
            
            row_vals = [
                r_idx + 1,
                t["Symbol"],
                t["Company Name"],
                t["Sector"],
                t["1 Lot Synthetic Pair"],
                t["ATM Strike (₹)"],
                t["Contract Expiry"],
                t["Expiry Decision"],
                t["Position Taking Window"],
                t["Entry Date"],
                t["Exit Date"],
                t["Entry Spot Price (₹)"],
                t["Exit Spot Price (₹)"],
                t["Lot Size (Qty)"],
                t["1 Lot Synthetic Margin (₹)"],
                t["P&L Per Share (₹)"],
                pnl_1lot,
                ret_1lot,
                t["Assigned Slot"],
                t["Re-entry Type"]
            ]
            
            for c_idx, val in enumerate(row_vals, 1):
                c = ws_q.cell(row_num, c_idx, val)
                align = "center"
                fmt = None
                if c_idx in [3, 4, 5]:
                    align = "left"
                elif c_idx in [6, 12, 13, 15, 16, 17]:
                    fmt = '₹#,##0.00'
                elif c_idx in [14]:
                    fmt = '#,##0'
                elif c_idx in [18]:
                    fmt = '0.00%'
                    
                bg = None
                txt_c = "1A202C"
                if c_idx in [17, 18]:
                    bg = FILL_WIN if pnl_1lot > 0 else FILL_LOSS
                    txt_c = COLOR_WIN_TXT if pnl_1lot > 0 else COLOR_LOSS_TXT
                elif c_idx in [5, 8]:
                    bg = FILL_BADGE
                    txt_c = COLOR_BADGE_TXT
                elif c_idx == 9:
                    bg = FILL_EXPECTED
                    txt_c = COLOR_EXPECTED_TXT
                elif r_idx % 2 == 1:
                    bg = FILL_ZEBRA
                    
                style_cell(c, size=9.5, bold=(c_idx in [2, 17, 18]), color=txt_c, bg=bg, align=align, num_fmt=fmt)
                
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
        ws_q.column_dimensions["E"].width = 32
        ws_q.column_dimensions["H"].width = 34

    # --- Create Individual Stock Overall Summary Sheet (Sheet 2) ---
    ws_stock_sum = wb_out.create_sheet(title="Stock_1Lot_Summary", index=0)
    ws_stock_sum.views.sheetView[0].showGridLines = True
    
    df_all_trades = pd.DataFrame(all_trades_master)
    stock_group = df_all_trades.groupby("Symbol").agg(
        Company=("Company Name", "first"),
        Sector=("Sector", "first"),
        Lot_Size=("Lot Size (Qty)", "first"),
        Total_Trades=("Quarter", "count"),
        Wins=("1 Lot Booked P&L (₹)", lambda x: (x > 0).sum()),
        Losses=("1 Lot Booked P&L (₹)", lambda x: (x <= 0).sum()),
        Avg_Margin=("1 Lot Synthetic Margin (₹)", "mean"),
        Total_Booked_PnL=("1 Lot Booked P&L (₹)", "sum"),
        Avg_Return_on_Margin=("Booked Return on 1 Lot Margin (%)", "mean")
    ).reset_index()
    
    stock_group["Win_Rate"] = stock_group["Wins"] / stock_group["Total_Trades"]
    stock_group = stock_group.sort_values(by="Total_Booked_PnL", ascending=False).reset_index(drop=True)
    
    # Title Banner
    ws_stock_sum.merge_cells("A1:K1")
    t_st = ws_stock_sum.cell(1, 1, f"INDIVIDUAL STOCK 1-LOT PERFORMANCE SUMMARY (ALL 12 QUARTERS) — {title_mode}")
    style_cell(t_st, size=13, bold=True, color=COLOR_WHITE, bg=FILL_NAVY, align="center")
    ws_stock_sum.row_dimensions[1].height = 38
    
    stock_headers = [
        "Rank", "Symbol", "Company Name", "Sector", "Lot Size (Qty)", 
        "Quarters Traded", "Wins", "Losses", "Win Rate (%)", 
        "Avg 1-Lot Margin (₹)", "Total 1-Lot Booked P&L (₹)"
    ]
    for idx, sh in enumerate(stock_headers, 1):
        c = ws_stock_sum.cell(3, idx, sh)
        style_cell(c, size=9.5, bold=True, color=COLOR_WHITE, bg=FILL_STEEL, align="center")
    ws_stock_sum.row_dimensions[3].height = 24
    
    for r_idx, srow in stock_group.iterrows():
        row_num = 4 + r_idx
        ws_stock_sum.row_dimensions[row_num].height = 20
        tot_pnl_st = srow["Total_Booked_PnL"]
        wr_st = srow["Win_Rate"]
        
        row_vals = [
            r_idx + 1,
            srow["Symbol"],
            srow["Company"],
            srow["Sector"],
            srow["Lot_Size"],
            srow["Total_Trades"],
            srow["Wins"],
            srow["Losses"],
            wr_st,
            srow["Avg_Margin"],
            tot_pnl_st
        ]
        
        for c_idx, val in enumerate(row_vals, 1):
            c = ws_stock_sum.cell(row_num, c_idx, val)
            align = "center"
            fmt = None
            if c_idx in [3, 4]:
                align = "left"
            elif c_idx in [5, 6, 7, 8]:
                fmt = '#,##0'
            elif c_idx in [9]:
                fmt = '0.00%'
            elif c_idx in [10, 11]:
                fmt = '₹#,##0.00'
                
            bg = None
            txt_c = "1A202C"
            if c_idx == 11:
                bg = FILL_WIN if tot_pnl_st > 0 else FILL_LOSS
                txt_c = COLOR_WIN_TXT if tot_pnl_st > 0 else COLOR_LOSS_TXT
            elif c_idx == 9:
                bg = FILL_WIN if wr_st >= 0.6 else (FILL_ICE if wr_st >= 0.5 else FILL_LOSS)
                txt_c = COLOR_WIN_TXT if wr_st >= 0.6 else ("1A202C" if wr_st >= 0.5 else COLOR_LOSS_TXT)
            elif r_idx % 2 == 1:
                bg = FILL_ZEBRA
                
            style_cell(c, size=9.5, bold=(c_idx in [2, 9, 11]), color=txt_c, bg=bg, align=align, num_fmt=fmt)
            
    for col in ws_stock_sum.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            if len(val_str) > max_len and cell.row > 1:
                max_len = len(val_str)
        ws_stock_sum.column_dimensions[col_letter].width = max(max_len + 3, 12)
    ws_stock_sum.column_dimensions["C"].width = 28
    ws_stock_sum.column_dimensions["D"].width = 20

    # --- Master Executive Summary Sheet (Sheet 1) ---
    ws_exec = wb_out.create_sheet(title=f"Exec_12Q_{mode}_Summary", index=0)
    ws_exec.views.sheetView[0].showGridLines = True
    
    df_qall = pd.DataFrame(quarter_summaries)
    tot_trades_all = int(df_qall["Total Trades"].sum())
    tot_win_all = int(df_qall["Winning Trades"].sum())
    tot_lose_all = int(df_qall["Losing Trades"].sum())
    tot_win_rate_all = (tot_win_all / tot_trades_all) if tot_trades_all > 0 else 0.0
    tot_net_pnl_all = float(df_qall["Net Booked Profit (₹)"].sum())
    avg_cap_all = float(df_qall["1 Lot Margin Deployed (₹)"].mean())
    avg_booked_ret = float(df_qall["Booked Return on Margin (%)"].mean())
    
    # Title Banner
    ws_exec.merge_cells("A1:M1")
    t_exec = ws_exec.cell(1, 1, f"NIFTY 50 {title_mode} — 12-QUARTERS EXECUTIVE OPTIONS DASHBOARD (1-LOT)")
    style_cell(t_exec, size=14, bold=True, color=COLOR_WHITE, bg=FILL_NAVY, align="center")
    ws_exec.row_dimensions[1].height = 42
    
    # Subtitle
    ws_exec.merge_cells("A2:M2")
    sub_exec = ws_exec.cell(2, 1, f"Indian Financial Year Standard (Q1 Starts in April) | Individual Stock 1-Lot Trading | ATM Options (Current vs Next Month Expiry Rollover)")
    style_cell(sub_exec, size=10, bold=False, color=COLOR_NAVY, bg=FILL_ICE, align="center")
    ws_exec.row_dimensions[2].height = 24
    
    # Scorecards Banner
    ws_exec.merge_cells("A4:M4")
    kpi_bar = ws_exec.cell(4, 1, "12-QUARTER GLOBAL 1-LOT PERFORMANCE SCORECARDS")
    style_cell(kpi_bar, size=11, bold=True, color=COLOR_WHITE, bg=FILL_STEEL, align="center")
    ws_exec.row_dimensions[4].height = 24
    
    scorecards = [
        ("Total 1-Lot Trades", f"{tot_trades_all:,}", "Across 12 Quarters"),
        ("Global Win Rate", f"{tot_win_rate_all:.1%}", f"{tot_win_all} Wins / {tot_lose_all} Losses"),
        ("Total Net Booked Profit", f"₹{tot_net_pnl_all:,.2f}", "Cumulative 1-Lot P&L"),
        ("Avg Booked Return / Qtr", f"{avg_booked_ret:.2%}", "On 1-Lot Margin Deployed"),
        ("Avg 1-Lot Margin Deployed", f"₹{avg_cap_all:,.2f}", "Per Quarter (Peak Slot Margin)")
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
    
    # Master Summary Table
    ws_exec.cell(9, 1, "QUARTER-BY-QUARTER 1-LOT OPTIONS PERFORMANCE (ALL 12 QUARTERS)")
    ws_exec.merge_cells("A9:M9")
    style_cell(ws_exec.cell(9, 1), size=11, bold=True, color=COLOR_WHITE, bg=FILL_NAVY, align="left")
    ws_exec.row_dimensions[9].height = 26
    
    master_table_cols = [
        "Quarter", "Business Quarter (FY)", "Financial Year", "1 Lot Margin Deployed (₹)", 
        "Final Value (₹)", "Net Booked Profit (₹)", "Booked Return on Margin (%)", "Expected Return on Margin (%)", 
        "Trades", "Wins", "Losses", "Win Rate (%)", "Re-entries"
    ]
    for idx, mc in enumerate(master_table_cols, 1):
        c = ws_exec.cell(10, idx, mc)
        style_cell(c, size=9, bold=True, color=COLOR_WHITE, bg=FILL_STEEL, align="center")
    ws_exec.row_dimensions[10].height = 24
    
    for r_idx, qrow in df_qall.iterrows():
        row_num = 11 + r_idx
        ws_exec.row_dimensions[row_num].height = 21
        pnl_q = qrow["Net Booked Profit (₹)"]
        wr_q = qrow["Win Rate"]
        
        row_vals = [
            qrow["Quarter"],
            qrow["Period"],
            qrow["FY"],
            qrow["1 Lot Margin Deployed (₹)"],
            qrow["Final Value (₹)"],
            pnl_q,
            qrow["Booked Return on Margin (%)"],
            qrow["Expected Return on Margin (%)"],
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
        float(df_qall["1 Lot Margin Deployed (₹)"].sum()),
        float(df_qall["Final Value (₹)"].sum()),
        tot_net_pnl_all,
        avg_booked_ret,
        float(df_qall["Expected Return on Margin (%)"].mean()),
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
    ws_all = wb_out.create_sheet(title=f"All_12Q_{mode}_1Lot_Trades")
    ws_all.views.sheetView[0].showGridLines = True
    
    df_m = pd.DataFrame(all_trades_master)
    tr_cols = [
        "Quarter", "FY", "Trade No", "Symbol", "Company Name", "Sector", "1 Lot Synthetic Pair", "ATM Strike (₹)",
        "Contract Expiry", "Expiry Decision", "Position Taking Window", 
        "Entry Date", "Exit Date", "Entry Spot (₹)", "Exit Spot (₹)", "Lot Size (Qty)", 
        "1 Lot Margin (₹)", "P&L Per Share (₹)", "1 Lot Booked P&L (₹)", "Booked Return on Margin (%)", 
        "Assigned Slot", "Re-entry Type"
    ]
    for idx, tch in enumerate(tr_cols, 1):
        c = ws_all.cell(1, idx, tch)
        style_cell(c, size=9.5, bold=True, color=COLOR_WHITE, bg=FILL_NAVY, align="center")
    ws_all.row_dimensions[1].height = 26
    
    for r_idx, trow in df_m.iterrows():
        row_num = 2 + r_idx
        ws_all.row_dimensions[row_num].height = 19
        lot_pnl_v = trow["1 Lot Booked P&L (₹)"]
        ret_mgn_v = trow["Booked Return on 1 Lot Margin (%)"]
        
        row_vals = [
            trow["Quarter"],
            trow["FY"],
            trow["Trade No"],
            trow["Symbol"],
            trow["Company Name"],
            trow["Sector"],
            trow["1 Lot Synthetic Pair"],
            trow["ATM Strike (₹)"],
            trow["Contract Expiry"],
            trow["Expiry Decision"],
            trow["Position Taking Window"],
            trow["Entry Date"],
            trow["Exit Date"],
            trow["Entry Spot Price (₹)"],
            trow["Exit Spot Price (₹)"],
            trow["Lot Size (Qty)"],
            trow["1 Lot Synthetic Margin (₹)"],
            trow["P&L Per Share (₹)"],
            lot_pnl_v,
            ret_mgn_v,
            trow["Assigned Slot"],
            trow["Re-entry Type"]
        ]
        for c_idx, val in enumerate(row_vals, 1):
            c = ws_all.cell(row_num, c_idx, val)
            align = "center"
            fmt = None
            if c_idx in [5, 6, 7]:
                align = "left"
            elif c_idx in [8, 14, 15, 17, 18, 19]:
                fmt = '₹#,##0.00'
            elif c_idx in [16]:
                fmt = '#,##0'
            elif c_idx in [20]:
                fmt = '0.00%'
                
            bg = None
            txt_c = "1A202C"
            if c_idx in [19, 20]:
                bg = FILL_WIN if lot_pnl_v > 0 else FILL_LOSS
                txt_c = COLOR_WIN_TXT if lot_pnl_v > 0 else COLOR_LOSS_TXT
            elif c_idx in [7, 10]:
                bg = FILL_BADGE
                txt_c = COLOR_BADGE_TXT
            elif c_idx == 11:
                bg = FILL_EXPECTED
                txt_c = COLOR_EXPECTED_TXT
            elif r_idx % 2 == 1:
                bg = FILL_ZEBRA
                
            style_cell(c, size=9.5, bold=(c_idx in [4, 19, 20]), color=txt_c, bg=bg, align=align, num_fmt=fmt)
            
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
    ws_all.column_dimensions["G"].width = 32
    ws_all.column_dimensions["J"].width = 34

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
    print("Building 12-Quarters Synthetic 1-Lot Individual Stock Master Workbooks...")
    
    # 1. Build Synthetic Long 1-Lot Master
    build_synthetic_1lot_master("LONG", OUT_SYN_LONG, OUT_SYN_LONG_REP)
    
    # 2. Build Synthetic Short 1-Lot Master
    build_synthetic_1lot_master("SHORT", OUT_SYN_SHORT, OUT_SYN_SHORT_REP)
    
    print("\nBoth 1-Lot Synthetic Options Master Workbooks generated successfully!")

if __name__ == "__main__":
    main()
