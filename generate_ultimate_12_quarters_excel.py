import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd

ROOT = r"D:\behaviour analysis"
REPORTS_DIR = os.path.join(ROOT, "12_Quarters_Reports")
OUTPUT_MASTER = os.path.join(ROOT, "Nifty50_12_Quarters_Master_Comprehensive.xlsx")
OUTPUT_MASTER_ALT = os.path.join(REPORTS_DIR, "Nifty50_12_Quarters_Master_Comprehensive.xlsx")

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

COMPANIES = {
    "ADANIENT": "Adani Enterprises Ltd",
    "ADANIPORTS": "Adani Ports and Special Economic Zone Ltd",
    "APOLLOHOSP": "Apollo Hospitals Enterprise Ltd",
    "ASIANPAINT": "Asian Paints Ltd",
    "AXISBANK": "Axis Bank Ltd",
    "BAJAJ-AUTO": "Bajaj Auto Ltd",
    "BAJAJFINSV": "Bajaj Finserv Ltd",
    "BAJFINANCE": "Bajaj Finance Ltd",
    "BEL": "Bharat Electronics Ltd",
    "BHARTIARTL": "Bharti Airtel Ltd",
    "CIPLA": "Cipla Ltd",
    "COALINDIA": "Coal India Ltd",
    "DRREDDY": "Dr. Reddy's Laboratories Ltd",
    "EICHERMOT": "Eicher Motors Ltd",
    "ETERNAL": "Eternal Limited",
    "GRASIM": "Grasim Industries Ltd",
    "HCLTECH": "HCL Technologies Ltd",
    "HDFCBANK": "HDFC Bank Ltd",
    "HDFCLIFE": "HDFC Life Insurance Co Ltd",
    "HEROMOTOCO": "Hero MotoCorp Ltd",
    "HINDALCO": "Hindalco Industries Ltd",
    "HINDUNILVR": "Hindustan Unilever Ltd",
    "ICICIBANK": "ICICI Bank Ltd",
    "INDIGO": "InterGlobe Aviation Ltd",
    "INDUSINDBK": "IndusInd Bank Ltd",
    "INFY": "Infosys Ltd",
    "ITC": "ITC Ltd",
    "JIOFIN": "Jio Financial Services Ltd",
    "JSWSTEEL": "JSW Steel Ltd",
    "KOTAKBANK": "Kotak Mahindra Bank Ltd",
    "LT": "Larsen & Toubro Ltd",
    "M&M": "Mahindra & Mahindra Ltd",
    "MARUTI": "Maruti Suzuki India Ltd",
    "MAXHEALTH": "Max Healthcare Institute Ltd",
    "NESTLEIND": "Nestle India Ltd",
    "NTPC": "NTPC Ltd",
    "ONGC": "Oil & Natural Gas Corporation Ltd",
    "POWERGRID": "Power Grid Corporation of India Ltd",
    "RELIANCE": "Reliance Industries Ltd",
    "SBILIFE": "SBI Life Insurance Co Ltd",
    "SBIN": "State Bank of India",
    "SHRIRAMFIN": "Shriram Finance Ltd",
    "SUNPHARMA": "Sun Pharmaceutical Industries Ltd",
    "TATACONSUM": "Tata Consumer Products Ltd",
    "TATASTEEL": "Tata Steel Ltd",
    "TCS": "Tata Consultancy Services Ltd",
    "TECHM": "Tech Mahindra Ltd",
    "TITAN": "Titan Company Ltd",
    "TMPV": "Tata Motors Passenger Vehicles Ltd",
    "TRENT": "Trent Ltd",
    "ULTRACEMCO": "UltraTech Cement Ltd",
    "WIPRO": "Wipro Ltd"
}

SECTORS = {
    "ADANIENT": "Diversified",
    "ADANIPORTS": "Services & Logistics",
    "APOLLOHOSP": "Healthcare & Pharma",
    "ASIANPAINT": "Consumer Durables",
    "AXISBANK": "Financial Services",
    "BAJAJ-AUTO": "Automobile & Auto",
    "BAJAJFINSV": "Financial Services",
    "BAJFINANCE": "Financial Services",
    "BEL": "Capital Goods / Defense",
    "BHARTIARTL": "Telecommunication",
    "CIPLA": "Healthcare & Pharma",
    "COALINDIA": "Energy / Mining",
    "DRREDDY": "Healthcare & Pharma",
    "EICHERMOT": "Automobile & Auto",
    "ETERNAL": "Consumer Services",
    "GRASIM": "Construction Materials",
    "HCLTECH": "Information Technology",
    "HDFCBANK": "Financial Services",
    "HDFCLIFE": "Financial Services",
    "HEROMOTOCO": "Automobile & Auto",
    "HINDALCO": "Metals & Mining",
    "HINDUNILVR": "Fast Moving Consumer Goods",
    "ICICIBANK": "Financial Services",
    "INDIGO": "Aviation / Transport",
    "INDUSINDBK": "Financial Services",
    "INFY": "Information Technology",
    "ITC": "Fast Moving Consumer Goods",
    "JIOFIN": "Financial Services",
    "JSWSTEEL": "Metals & Mining",
    "KOTAKBANK": "Financial Services",
    "LT": "Construction",
    "M&M": "Automobile & Auto",
    "MARUTI": "Automobile & Auto",
    "MAXHEALTH": "Healthcare & Pharma",
    "NESTLEIND": "Fast Moving Consumer Goods",
    "NTPC": "Energy & Utilities",
    "ONGC": "Energy / Oil & Gas",
    "POWERGRID": "Energy & Utilities",
    "RELIANCE": "Energy & Conglomerate",
    "SBILIFE": "Financial Services",
    "SBIN": "Financial Services",
    "SHRIRAMFIN": "Financial Services",
    "SUNPHARMA": "Healthcare & Pharma",
    "TATACONSUM": "Fast Moving Consumer Goods",
    "TATASTEEL": "Metals & Mining",
    "TCS": "Information Technology",
    "TECHM": "Information Technology",
    "TITAN": "Consumer Durables",
    "TMPV": "Automobile & Auto",
    "TRENT": "Retail & Services",
    "ULTRACEMCO": "Construction Materials",
    "WIPRO": "Information Technology"
}

# Typography & Color Palette
FONT_FAMILY = "Segoe UI"
COLOR_PRIMARY_NAVY = "1A365D"    # Deep Royal Navy
COLOR_SECONDARY_BLUE = "2B6CB0"  # Vibrant Steel Blue
COLOR_ICE_BLUE = "EBF8FF"        # Light Ice Blue
COLOR_LIGHT_GRAY = "F7FAFC"      # Subtle Soft Gray for alternate rows
COLOR_WHITE = "FFFFFF"

# Distinct Colors for Expected vs Realised Return:
# 1. Expected Return (Model Prediction): Elegant Soft Lavender / Indigo
COLOR_EXPECTED_BG = "EDE9FE"     # Soft Lavender
COLOR_EXPECTED_TXT = "5B21B6"    # Deep Purple / Indigo

# 2. Realised Return & PnL (Actual Return We Got):
COLOR_WIN_BG = "DCFCE7"          # Vibrant Emerald Mint
COLOR_WIN_TXT = "15803D"         # Rich Forest Green
COLOR_LOSS_BG = "FEE2E2"         # Soft Coral Crimson
COLOR_LOSS_TXT = "B91C1C"        # Bold Crimson Red

COLOR_BADGE_BG = "EDF2F7"        # Cool Slate Badge
COLOR_BADGE_TXT = "2D3748"

FILL_NAVY = PatternFill(start_color=COLOR_PRIMARY_NAVY, end_color=COLOR_PRIMARY_NAVY, fill_type="solid")
FILL_STEEL = PatternFill(start_color=COLOR_SECONDARY_BLUE, end_color=COLOR_SECONDARY_BLUE, fill_type="solid")
FILL_ICE = PatternFill(start_color=COLOR_ICE_BLUE, end_color=COLOR_ICE_BLUE, fill_type="solid")
FILL_ZEBRA = PatternFill(start_color=COLOR_LIGHT_GRAY, end_color=COLOR_LIGHT_GRAY, fill_type="solid")
FILL_EXPECTED = PatternFill(start_color=COLOR_EXPECTED_BG, end_color=COLOR_EXPECTED_BG, fill_type="solid")
FILL_WIN = PatternFill(start_color=COLOR_WIN_BG, end_color=COLOR_WIN_BG, fill_type="solid")
FILL_LOSS = PatternFill(start_color=COLOR_LOSS_BG, end_color=COLOR_LOSS_BG, fill_type="solid")
FILL_BADGE = PatternFill(start_color=COLOR_BADGE_BG, end_color=COLOR_BADGE_BG, fill_type="solid")
FILL_LONG = PatternFill(start_color="EBF8FF", end_color="EBF8FF", fill_type="solid")
FILL_SHORT = PatternFill(start_color="FEFCBF", end_color="FEFCBF", fill_type="solid")

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

def main():
    print("Generating comprehensive Master Excel with full Position Taking Window column...")
    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)
    
    quarter_metrics = []
    all_master_trades = []
    all_master_losers = []
    
    for fn, q_name, period_desc, fy_tag in QUARTERS:
        root_path = os.path.join(ROOT, f"{fn}_Combined_Best_Capital_Utilisation.xlsx")
        rep_path = os.path.join(REPORTS_DIR, f"{fn}_Combined_Best_Capital_Utilisation.xlsx")
        
        if fn == "Q2_2026_27":
            root_v5 = os.path.join(ROOT, "Q2_2026_27_Combined_Best_Capital_Utilisation_v5.xlsx")
            file_path = root_v5 if os.path.exists(root_v5) else (root_path if os.path.exists(root_path) else rep_path)
        else:
            file_path = root_path if os.path.exists(root_path) else rep_path
            
        if not os.path.exists(file_path):
            print(f"File missing: {file_path}")
            continue
            
        print(f"Reading {q_name} from {file_path}...")
        df_dt = pd.read_excel(file_path, sheet_name="Detailed_Trades")
        
        wb_src = openpyxl.load_workbook(file_path, data_only=True)
        ws_sum = wb_src["Summary"]
        fund_util = ws_sum.cell(4, 2).value or ws_sum.cell(3, 2).value or 0.0
        final_val = ws_sum.cell(4, 3).value or ws_sum.cell(3, 3).value or 0.0
        net_prof = ws_sum.cell(4, 4).value or ws_sum.cell(3, 4).value or 0.0
        exp_ret = ws_sum.cell(4, 5).value or ws_sum.cell(3, 5).value or 0.0
        booked_ret = ws_sum.cell(4, 6).value or ws_sum.cell(3, 6).value or 0.0
        
        df_lose = None
        if "Losing Trades Verification" in wb_src.sheetnames:
            try:
                df_lose = pd.read_excel(file_path, sheet_name="Losing Trades Verification")
            except Exception:
                pass
        wb_src.close()
        
        # Clean columns
        ret_c = next((c for c in df_dt.columns if "Realised Return" in c), "Realised Return (%)")
        pnl_c = next((c for c in df_dt.columns if "Realised Profit" in c or "Profit/Loss" in c), "Realised Profit/Loss (Rs.)")
        strat_c = next((c for c in df_dt.columns if "Strategy" in c), "Strategy Type (Decided Move)")
        entry_p_c = next((c for c in df_dt.columns if "Entry Price" in c or "Buy Price" in c), "Entry Price (Rs.)")
        exit_p_c = next((c for c in df_dt.columns if "Exit Price" in c or "Sell Price" in c), "Exit Price (Rs.)")
        exp_c = next((c for c in df_dt.columns if "Expected Return" in c), "Expected Return (%)")
        offset_c = next((c for c in df_dt.columns if "Offset" in c or "Position" in c or "Window" in c), None)
        
        df_dt = df_dt.dropna(subset=["Symbol"]).copy()
        df_dt["Symbol"] = df_dt["Symbol"].astype(str).str.strip().str.upper()
        df_dt["Company Name"] = df_dt["Symbol"].map(lambda s: COMPANIES.get(s, s))
        df_dt["Sector"] = df_dt["Symbol"].map(lambda s: SECTORS.get(s, "Diversified"))
        df_dt["Strategy Clean"] = df_dt[strat_c].astype(str).str.upper().apply(lambda x: "SHORT" if "SHORT" in x or "SELL" in x else "LONG")
        df_dt["Entry Date Clean"] = pd.to_datetime(df_dt["Entry Date"]).dt.strftime("%Y-%m-%d")
        df_dt["Exit Date Clean"] = pd.to_datetime(df_dt["Exit Date"]).dt.strftime("%Y-%m-%d")
        df_dt["Entry Price Clean"] = pd.to_numeric(df_dt[entry_p_c], errors="coerce").fillna(0.0)
        df_dt["Exit Price Clean"] = pd.to_numeric(df_dt[exit_p_c], errors="coerce").fillna(0.0)
        df_dt["Expected Return Clean"] = pd.to_numeric(df_dt[exp_c], errors="coerce").fillna(0.0)
        df_dt["Realised Return Clean"] = pd.to_numeric(df_dt[ret_c], errors="coerce").fillna(0.0)
        df_dt["Realised PnL Clean"] = pd.to_numeric(df_dt[pnl_c], errors="coerce").fillna(0.0)
        df_dt["Re-entry Clean"] = df_dt.get("Re-entry Type", "First Entry").fillna("First Entry").astype(str)
        df_dt["Assigned Slot Clean"] = df_dt.get("Assigned Slot", "Slot 1").fillna("Slot 1").astype(str)
        
        if offset_c and offset_c in df_dt.columns:
            df_dt["Position Taking Window"] = df_dt[offset_c].fillna("T-5 to T+5").astype(str)
        else:
            df_dt["Position Taking Window"] = "T-5 to T+5"
            
        # Quarter metrics
        tot_cnt = len(df_dt)
        win_cnt = int((df_dt["Realised Return Clean"] > 0).sum())
        lose_cnt = int((df_dt["Realised Return Clean"] <= 0).sum())
        win_rate = (win_cnt / tot_cnt) if tot_cnt > 0 else 0.0
        
        long_cnt = int((df_dt["Strategy Clean"] == "LONG").sum())
        short_cnt = int((df_dt["Strategy Clean"] == "SHORT").sum())
        long_win = int(((df_dt["Strategy Clean"] == "LONG") & (df_dt["Realised Return Clean"] > 0)).sum())
        short_win = int(((df_dt["Strategy Clean"] == "SHORT") & (df_dt["Realised Return Clean"] > 0)).sum())
        long_pnl = float(df_dt[df_dt["Strategy Clean"] == "LONG"]["Realised PnL Clean"].sum())
        short_pnl = float(df_dt[df_dt["Strategy Clean"] == "SHORT"]["Realised PnL Clean"].sum())
        reentry_cnt = int((df_dt["Re-entry Clean"].str.contains("Re-entry", case=False)).sum())
        slots_cnt = len(df_dt["Assigned Slot Clean"].unique())
        
        # Calculate true Net Booked Profit and Booked Return dynamically from all trades
        net_prof = float(df_dt["Realised PnL Clean"].sum())
        fund_util_val = float(fund_util) if float(fund_util) > 0 else 60000.0
        final_val = fund_util_val + net_prof
        booked_ret = (net_prof / fund_util_val) if fund_util_val > 0 else 0.0
        exp_ret_val = float(df_dt["Expected Return Clean"].mean()) if "Expected Return Clean" in df_dt else float(exp_ret)
        
        quarter_metrics.append({
            "Quarter": q_name,
            "Period": period_desc,
            "FY": fy_tag,
            "Fund Utilised (Rs.)": fund_util_val,
            "Final Value (Rs.)": float(final_val),
            "Net Profit (Rs.)": float(net_prof),
            "Booked Return (%)": float(booked_ret),
            "Expected Return (%)": float(exp_ret_val),
            "Total Trades": tot_cnt,
            "Winning Trades": win_cnt,
            "Losing Trades": lose_cnt,
            "Win Rate (%)": win_rate,
            "Long Trades": long_cnt,
            "Short Trades": short_cnt,
            "Long Win Rate (%)": (long_win / long_cnt) if long_cnt > 0 else 0.0,
            "Short Win Rate (%)": (short_win / short_cnt) if short_cnt > 0 else 0.0,
            "Long Profit (Rs.)": long_pnl,
            "Short Profit (Rs.)": short_pnl,
            "Slots Utilised": slots_cnt,
            "Total Re-entries": reentry_cnt
        })
        
        for t_idx, row in df_dt.iterrows():
            all_master_trades.append({
                "Quarter": q_name,
                "FY": fy_tag,
                "Trade No": t_idx + 1,
                "Symbol": row["Symbol"],
                "Company Name": row["Company Name"],
                "Sector": row["Sector"],
                "Strategy": row["Strategy Clean"],
                "Position Taking Window": row["Position Taking Window"],
                "Entry Date": row["Entry Date Clean"],
                "Exit Date": row["Exit Date Clean"],
                "Entry Price (Rs.)": row["Entry Price Clean"],
                "Exit Price (Rs.)": row["Exit Price Clean"],
                "Expected Return (%)": row["Expected Return Clean"],
                "Realised Return (%)": row["Realised Return Clean"],
                "Realised Profit/Loss (Rs.)": row["Realised PnL Clean"],
                "Assigned Slot": row["Assigned Slot Clean"],
                "Re-entry Type": row["Re-entry Clean"]
            })
            
        if df_lose is not None and not df_lose.empty:
            l_rows = list(df_lose.values)
            h_idx = 0
            for idx, r in enumerate(l_rows[:5]):
                if any(isinstance(v, str) and "Symbol" in v for v in r):
                    h_idx = idx
                    break
            l_headers = [str(h).strip() for h in l_rows[h_idx]]
            l_data = l_rows[h_idx+1:]
            df_l_clean = pd.DataFrame(l_data, columns=l_headers).dropna(subset=[l_headers[0]])
            for _, lrow in df_l_clean.iterrows():
                sym_l = str(lrow.get("Symbol", "")).strip().upper()
                if sym_l and sym_l != "NAN":
                    all_master_losers.append({
                        "Quarter": q_name,
                        "Symbol": sym_l,
                        "Company Name": COMPANIES.get(sym_l, str(lrow.get("Company Name", sym_l))),
                        "Strategy": str(lrow.get("Decided Move (Strategy)", lrow.get("Strategy", ""))),
                        "Long Return (%)": lrow.get("Long Return (%)", "-"),
                        "Short Return (%)": lrow.get("Short Return (%)", "-"),
                        "Realised Loss (Rs.)": lrow.get("Realised Loss (Rs.)", 0.0),
                        "Verification Status": lrow.get("Verification Status", "Passed"),
                        "Rationale": str(lrow.get("Rationale / Explanation", lrow.get("Rationale", "Optimal side selected based on in-sample risk.")))
                    })

        # --- BUILD ATTRACTIVE INDIVIDUAL QUARTER SHEET ---
        ws_q = wb_out.create_sheet(title=q_name)
        ws_q.views.sheetView[0].showGridLines = True
        
        # Title Banner
        ws_q.merge_cells("A1:O1")
        t_c = ws_q.cell(1, 1, f"NIFTY 50 EVENT-DRIVEN TRADING MODEL — {q_name} ({period_desc})")
        style_cell(t_c, size=13, bold=True, color=COLOR_WHITE, bg=FILL_NAVY, align="center")
        ws_q.row_dimensions[1].height = 36
        
        # Quarter KPI Header
        ws_q.merge_cells("A3:O3")
        kpi_title = ws_q.cell(3, 1, f"QUARTER PERFORMANCE OVERVIEW & CAPITAL UTILISATION ({fy_tag})")
        style_cell(kpi_title, size=10, bold=True, color=COLOR_WHITE, bg=FILL_STEEL, align="left")
        ws_q.row_dimensions[3].height = 24
        
        kpi_headers = [
            "Fund Utilised", "Final Value", "Net Booked Profit (₹)", "Booked Return", 
            "Expected Return", "Win Rate", "Total Trades", "Winning Trades", 
            "Losing Trades", "Long Trades", "Short Trades", "Slots Deployed", "Re-entries"
        ]
        for idx, kh in enumerate(kpi_headers, 1):
            c = ws_q.cell(4, idx, kh)
            style_cell(c, size=9, bold=True, color=COLOR_WHITE, bg=FILL_STEEL, align="center")
        ws_q.row_dimensions[4].height = 22
        
        kpi_row_vals = [
            float(fund_util), float(final_val), float(net_prof), float(booked_ret),
            float(exp_ret), win_rate, tot_cnt, win_cnt, lose_cnt, long_cnt, short_cnt, slots_cnt, reentry_cnt
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
                
            bg_kpi = FILL_WIN if (idx == 3 and kv > 0) or (idx == 6 and kv >= 0.6) else FILL_ICE
            txt_color = COLOR_WIN_TXT if (idx == 3 and kv > 0) else "1A202C"
            style_cell(c, size=10, bold=True, color=txt_color, bg=bg_kpi, align="center", num_fmt=fmt)
        ws_q.row_dimensions[5].height = 25
        
        # Detailed Trades Section
        ws_q.cell(7, 1, "CHRONOLOGICAL TRADE JOURNAL — POSITION TAKING & BOOKED PROFIT/LOSS")
        ws_q.merge_cells("A7:O7")
        style_cell(ws_q.cell(7, 1), size=10.5, bold=True, color=COLOR_WHITE, bg=FILL_NAVY, align="left")
        ws_q.row_dimensions[7].height = 26
        
        trade_headers = [
            "Trade #", "Symbol", "Company Name", "Sector", "Strategy", 
            "Position Taking Window", "Entry Date", "Exit Date", "Entry Price (₹)", "Exit Price (₹)", 
            "Expected Return (%)", "Booked Return (%)", "Booked Profit/Loss (₹)", 
            "Slot ID", "Re-entry Type"
        ]
        for idx, th in enumerate(trade_headers, 1):
            c = ws_q.cell(8, idx, th)
            style_cell(c, size=9, bold=True, color=COLOR_WHITE, bg=FILL_STEEL, align="center")
        ws_q.row_dimensions[8].height = 24
        
        for r_idx, row in df_dt.iterrows():
            row_num = 9 + r_idx
            ws_q.row_dimensions[row_num].height = 20
            ret_val = row["Realised Return Clean"]
            pnl_val = row["Realised PnL Clean"]
            strat_val = row["Strategy Clean"]
            
            row_vals = [
                r_idx + 1,
                row["Symbol"],
                row["Company Name"],
                row["Sector"],
                strat_val,
                row["Position Taking Window"],
                row["Entry Date Clean"],
                row["Exit Date Clean"],
                row["Entry Price Clean"],
                row["Exit Price Clean"],
                row["Expected Return Clean"],
                ret_val,
                pnl_val,
                row["Assigned Slot Clean"],
                row["Re-entry Clean"]
            ]
            
            for c_idx, val in enumerate(row_vals, 1):
                c = ws_q.cell(row_num, c_idx, val)
                align = "center"
                fmt = None
                if c_idx in [3, 4]:
                    align = "left"
                elif c_idx in [9, 10, 13]:
                    fmt = '₹#,##0.00'
                elif c_idx in [11, 12]:
                    fmt = '0.00%'
                    
                bg = None
                txt_c = "1A202C"
                if c_idx == 11:
                    # Expected Return: Distinct Lavender / Purple
                    bg = FILL_EXPECTED
                    txt_c = COLOR_EXPECTED_TXT
                elif c_idx in [12, 13]:
                    # Realised Return & PnL (Return we get): Distinct Green / Red
                    bg = FILL_WIN if ret_val > 0 else FILL_LOSS
                    txt_c = COLOR_WIN_TXT if ret_val > 0 else COLOR_LOSS_TXT
                elif c_idx == 5:
                    bg = FILL_LONG if strat_val == "LONG" else FILL_SHORT
                elif c_idx == 6:
                    bg = FILL_BADGE
                    txt_c = COLOR_BADGE_TXT
                elif r_idx % 2 == 1:
                    bg = FILL_ZEBRA
                    
                style_cell(c, size=9.5, bold=(c_idx in [2, 11, 12, 13]), color=txt_c, bg=bg, align=align, num_fmt=fmt)
                
        # Auto-fit columns
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
        ws_q.column_dimensions["F"].width = 22

    # =========================================================================
    # MASTER 12-QUARTERS EXECUTIVE SUMMARY SHEET (SHEET 1)
    # =========================================================================
    ws_exec = wb_out.create_sheet(title="Executive_12Q_Summary", index=0)
    ws_exec.views.sheetView[0].showGridLines = True
    
    df_qall = pd.DataFrame(quarter_metrics)
    tot_trades_all = int(df_qall["Total Trades"].sum())
    tot_win_all = int(df_qall["Winning Trades"].sum())
    tot_lose_all = int(df_qall["Losing Trades"].sum())
    tot_win_rate_all = (tot_win_all / tot_trades_all) if tot_trades_all > 0 else 0.0
    tot_net_pnl_all = float(df_qall["Net Profit (Rs.)"].sum())
    tot_longs_all = int(df_qall["Long Trades"].sum())
    tot_shorts_all = int(df_qall["Short Trades"].sum())
    tot_long_pnl_all = float(df_qall["Long Profit (Rs.)"].sum())
    tot_short_pnl_all = float(df_qall["Short Profit (Rs.)"].sum())
    avg_cap_all = float(df_qall["Fund Utilised (Rs.)"].mean())
    avg_booked_ret = float(df_qall["Booked Return (%)"].mean())
    
    # Title Banner
    ws_exec.merge_cells("A1:O1")
    t_exec = ws_exec.cell(1, 1, "NIFTY 50 EVENT-DRIVEN TRADING STRATEGY — 12-QUARTERS EXECUTIVE MASTER DASHBOARD")
    style_cell(t_exec, size=14, bold=True, color=COLOR_WHITE, bg=FILL_NAVY, align="center")
    ws_exec.row_dimensions[1].height = 42
    
    # Subtitle Info (Indian FY Reminder)
    ws_exec.merge_cells("A2:O2")
    sub_exec = ws_exec.cell(2, 1, "Indian Financial Year Standard (Q1 Starts in April) | 12 Rolling Quarters (FY24 to FY27) | Long & Short Combined Best Optimization")
    style_cell(sub_exec, size=10, bold=False, color=COLOR_PRIMARY_NAVY, bg=FILL_ICE, align="center")
    ws_exec.row_dimensions[2].height = 24
    
    # KPI Scorecard Banner
    ws_exec.merge_cells("A4:O4")
    kpi_bar = ws_exec.cell(4, 1, "12-QUARTER GLOBAL PERFORMANCE SCORECARDS")
    style_cell(kpi_bar, size=11, bold=True, color=COLOR_WHITE, bg=FILL_STEEL, align="center")
    ws_exec.row_dimensions[4].height = 24
    
    scorecards = [
        ("Total Trades Analyzed", f"{tot_trades_all:,}", "Across 12 Quarters"),
        ("Global Win Rate", f"{tot_win_rate_all:.1%}", f"{tot_win_all} Wins / {tot_lose_all} Losses"),
        ("Total Net Realised Profit", f"₹{tot_net_pnl_all:,.2f}", "Cumulative Net P&L"),
        ("Avg Booked Return / Qtr", f"{avg_booked_ret:.2%}", "Quarterly Return on Capital"),
        ("Average Capital Deployed", f"₹{avg_cap_all:,.2f}", "Per Quarter"),
        ("Long vs Short Distribution", f"{tot_longs_all} L / {tot_shorts_all} S", f"Long: ₹{tot_long_pnl_all:,.0f} | Short: ₹{tot_short_pnl_all:,.0f}")
    ]
    
    col_starts = [1, 3, 6, 9, 11, 13]
    col_spans = [2, 3, 3, 2, 2, 3]
    
    for idx, (title, main_val, sub_val) in enumerate(scorecards):
        c_s = col_starts[idx]
        w = col_spans[idx]
        c_e = c_s + w - 1
        
        ws_exec.merge_cells(start_row=5, start_column=c_s, end_row=5, end_column=c_e)
        c_t = ws_exec.cell(5, c_s, title)
        style_cell(c_t, size=9, bold=True, color="4A5568", bg=FILL_BADGE, align="center")
        
        ws_exec.merge_cells(start_row=6, start_column=c_s, end_row=6, end_column=c_e)
        c_m = ws_exec.cell(6, c_s, main_val)
        bg_card = FILL_WIN if "Profit" in title or "Win" in title else FILL_ICE
        txt_card = COLOR_WIN_TXT if "Profit" in title or "Win" in title else COLOR_PRIMARY_NAVY
        style_cell(c_m, size=13, bold=True, color=txt_card, bg=bg_card, align="center")
        
        ws_exec.merge_cells(start_row=7, start_column=c_s, end_row=7, end_column=c_e)
        c_sub = ws_exec.cell(7, c_s, sub_val)
        style_cell(c_sub, size=8, bold=False, color="4A5568", bg=FILL_BADGE, align="center")
        
    ws_exec.row_dimensions[5].height = 18
    ws_exec.row_dimensions[6].height = 26
    ws_exec.row_dimensions[7].height = 16
    
    # 12-Quarters Master Breakdown Table
    ws_exec.cell(9, 1, "QUARTER-BY-QUARTER PERFORMANCE COMPARISON (ALL 12 QUARTERS)")
    ws_exec.merge_cells("A9:O9")
    style_cell(ws_exec.cell(9, 1), size=11, bold=True, color=COLOR_WHITE, bg=FILL_NAVY, align="left")
    ws_exec.row_dimensions[9].height = 26
    
    master_table_cols = [
        "Quarter", "Business Quarter (FY)", "Financial Year", "Fund Utilised (₹)", 
        "Final Value (₹)", "Net Booked Profit (₹)", "Booked Return (%)", "Expected Return (%)", 
        "Trades", "Wins", "Losses", "Win Rate (%)", "Long Trades", "Short Trades", "Re-entries"
    ]
    for idx, mc in enumerate(master_table_cols, 1):
        c = ws_exec.cell(10, idx, mc)
        style_cell(c, size=9, bold=True, color=COLOR_WHITE, bg=FILL_STEEL, align="center")
    ws_exec.row_dimensions[10].height = 24
    
    for r_idx, qrow in df_qall.iterrows():
        row_num = 11 + r_idx
        ws_exec.row_dimensions[row_num].height = 21
        pnl_q = qrow["Net Profit (Rs.)"]
        wr_q = qrow["Win Rate (%)"]
        
        row_vals = [
            qrow["Quarter"],
            qrow["Period"],
            qrow["FY"],
            qrow["Fund Utilised (Rs.)"],
            qrow["Final Value (Rs.)"],
            pnl_q,
            qrow["Booked Return (%)"],
            qrow["Expected Return (%)"],
            qrow["Total Trades"],
            qrow["Winning Trades"],
            qrow["Losing Trades"],
            wr_q,
            qrow["Long Trades"],
            qrow["Short Trades"],
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
            elif c_idx in [9, 10, 11, 13, 14, 15]:
                fmt = '#,##0'
                
            bg = None
            txt_c = "1A202C"
            if c_idx == 6:
                bg = FILL_WIN if pnl_q > 0 else FILL_LOSS
                txt_c = COLOR_WIN_TXT if pnl_q > 0 else COLOR_LOSS_TXT
            elif c_idx == 7:
                # Booked Return
                bg = FILL_WIN if pnl_q > 0 else FILL_LOSS
                txt_c = COLOR_WIN_TXT if pnl_q > 0 else COLOR_LOSS_TXT
            elif c_idx == 8:
                # Expected Return
                bg = FILL_EXPECTED
                txt_c = COLOR_EXPECTED_TXT
            elif c_idx == 12:
                bg = FILL_WIN if wr_q >= 0.65 else (FILL_ICE if wr_q >= 0.55 else FILL_LOSS)
                txt_c = COLOR_WIN_TXT if wr_q >= 0.65 else ("1A202C" if wr_q >= 0.55 else COLOR_LOSS_TXT)
            elif r_idx % 2 == 1:
                bg = FILL_ZEBRA
                
            style_cell(c, size=9.5, bold=(c_idx in [1, 6, 7, 8, 12]), color=txt_c, bg=bg, align=align, num_fmt=fmt)
            
    # Total / Average Summary Row
    tot_row = 11 + len(df_qall)
    ws_exec.row_dimensions[tot_row].height = 25
    tot_row_vals = [
        "12-Quarter Total / Summary",
        "FY 2023-24 to FY 2026-27 (3 Full Years)",
        "3 Financial Years",
        float(df_qall["Fund Utilised (Rs.)"].sum()),
        float(df_qall["Final Value (Rs.)"].sum()),
        tot_net_pnl_all,
        avg_booked_ret,
        float(df_qall["Expected Return (%)"].mean()),
        tot_trades_all,
        tot_win_all,
        tot_lose_all,
        tot_win_rate_all,
        tot_longs_all,
        tot_shorts_all,
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
        elif c_idx in [9, 10, 11, 13, 14, 15]:
            fmt = '#,##0'
        style_cell(c, size=10, bold=True, color=COLOR_PRIMARY_NAVY, bg=FILL_ICE, align=align, num_fmt=fmt)

    # Auto-fit executive sheet columns
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

    # =========================================================================
    # MASTER ALL TRADES SHEET (WITH POSITION TAKING WINDOW)
    # =========================================================================
    ws_all_tr = wb_out.create_sheet(title="All_12_Quarters_Master_Trades")
    ws_all_tr.views.sheetView[0].showGridLines = True
    
    df_master_all = pd.DataFrame(all_master_trades)
    tr_cols = [
        "Quarter", "FY", "Trade No", "Symbol", "Company Name", "Sector", 
        "Strategy", "Position Taking Window", "Entry Date", "Exit Date", 
        "Entry Price (₹)", "Exit Price (₹)", "Expected Return (%)", 
        "Booked Return (%)", "Booked Profit/Loss (₹)", "Assigned Slot", "Re-entry Type"
    ]
    
    for idx, tch in enumerate(tr_cols, 1):
        c = ws_all_tr.cell(1, idx, tch)
        style_cell(c, size=9.5, bold=True, color=COLOR_WHITE, bg=FILL_NAVY, align="center")
    ws_all_tr.row_dimensions[1].height = 26
    
    for r_idx, trow in df_master_all.iterrows():
        row_num = 2 + r_idx
        ws_all_tr.row_dimensions[row_num].height = 19
        ret_v = trow["Realised Return (%)"]
        pnl_v = trow["Realised Profit/Loss (Rs.)"]
        strat_v = trow["Strategy"]
        
        row_vals = [
            trow["Quarter"],
            trow["FY"],
            trow["Trade No"],
            trow["Symbol"],
            trow["Company Name"],
            trow["Sector"],
            strat_v,
            trow["Position Taking Window"],
            trow["Entry Date"],
            trow["Exit Date"],
            trow["Entry Price (Rs.)"],
            trow["Exit Price (Rs.)"],
            trow["Expected Return (%)"],
            ret_v,
            pnl_v,
            trow["Assigned Slot"],
            trow["Re-entry Type"]
        ]
        
        for c_idx, val in enumerate(row_vals, 1):
            c = ws_all_tr.cell(row_num, c_idx, val)
            align = "center"
            fmt = None
            if c_idx in [5, 6]:
                align = "left"
            elif c_idx in [11, 12, 15]:
                fmt = '₹#,##0.00'
            elif c_idx in [13, 14]:
                fmt = '0.00%'
                
            bg = None
            txt_c = "1A202C"
            if c_idx == 13:
                # Expected Return: Distinct Lavender / Purple
                bg = FILL_EXPECTED
                txt_c = COLOR_EXPECTED_TXT
            elif c_idx in [14, 15]:
                # Realised Return & Realised PnL: Distinct Mint / Crimson
                bg = FILL_WIN if ret_v > 0 else FILL_LOSS
                txt_c = COLOR_WIN_TXT if ret_v > 0 else COLOR_LOSS_TXT
            elif c_idx == 7:
                bg = FILL_LONG if strat_v == "LONG" else FILL_SHORT
            elif c_idx == 8:
                bg = FILL_BADGE
                txt_c = COLOR_BADGE_TXT
            elif r_idx % 2 == 1:
                bg = FILL_ZEBRA
                
            style_cell(c, size=9.5, bold=(c_idx in [4, 13, 14, 15]), color=txt_c, bg=bg, align=align, num_fmt=fmt)
            
    for col in ws_all_tr.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws_all_tr.column_dimensions[col_letter].width = max(max_len + 3, 12)
    ws_all_tr.column_dimensions["E"].width = 28
    ws_all_tr.column_dimensions["F"].width = 20
    ws_all_tr.column_dimensions["H"].width = 22

    # =========================================================================
    # MASTER LOSING TRADES AUDIT SHEET
    # =========================================================================
    if all_master_losers:
        ws_lose = wb_out.create_sheet(title="Losing_Trades_Audit_Master")
        ws_lose.views.sheetView[0].showGridLines = True
        df_all_lose = pd.DataFrame(all_master_losers)
        
        lose_cols = [
            "Quarter", "Symbol", "Company Name", "Strategy", 
            "Long Return (%)", "Short Return (%)", "Booked Loss (₹)", 
            "Verification Status", "Rationale / Reason for Choice"
        ]
        for idx, lch in enumerate(lose_cols, 1):
            c = ws_lose.cell(1, idx, lch)
            style_cell(c, size=9.5, bold=True, color=COLOR_WHITE, bg=FILL_NAVY, align="center")
        ws_lose.row_dimensions[1].height = 26
        
        for r_idx, lrow in df_all_lose.iterrows():
            row_num = 2 + r_idx
            ws_lose.row_dimensions[row_num].height = 20
            row_vals = [
                lrow["Quarter"],
                lrow["Symbol"],
                lrow["Company Name"],
                lrow["Strategy"],
                lrow["Long Return (%)"],
                lrow["Short Return (%)"],
                lrow["Realised Loss (Rs.)"],
                lrow["Verification Status"],
                lrow["Rationale"]
            ]
            for c_idx, val in enumerate(row_vals, 1):
                c = ws_lose.cell(row_num, c_idx, val)
                align = "left" if c_idx in [3, 9] else "center"
                fmt = '₹#,##0.00' if c_idx == 7 and isinstance(val, (int, float)) else None
                bg = FILL_ZEBRA if r_idx % 2 == 1 else None
                style_cell(c, size=9.5, bold=(c_idx in [2, 7, 8]), bg=bg, align=align, num_fmt=fmt)
                
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

    # Save Master File with fallback if locked by Excel
    saved_paths = []
    for target in [
        os.path.join(ROOT, "Nifty50_12_Quarters_Booked_PNL_Master.xlsx"),
        OUTPUT_MASTER, 
        os.path.join(ROOT, "Nifty50_12_Quarters_Master_Comprehensive_v2.xlsx"), 
        os.path.join(ROOT, "Nifty50_12_Quarters_Consolidated.xlsx")
    ]:
        try:
            wb_out.save(target)
            saved_paths.append(target)
            print(f"Saved: {target}")
            break
        except PermissionError:
            print(f"File {target} is open in Excel, trying alternative...")
            
    try:
        wb_out.save(OUTPUT_MASTER_ALT)
        saved_paths.append(OUTPUT_MASTER_ALT)
    except Exception:
        pass
        
    print(f"Master Excel successfully created at:\n" + "\n".join(saved_paths))

if __name__ == "__main__":
    main()
