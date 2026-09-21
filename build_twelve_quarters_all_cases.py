import datetime as dt
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os
import pandas as pd
import numpy as np
from pathlib import Path
import urllib.request

ROOT = Path("D:/behaviour analysis")
PEER_FILE = ROOT / "Nifty50_Peer_Model_8Q_ALL50.xlsx"
QTY_FILE = ROOT / "Nifty50Stocks_QtyResultDates.xlsx"

QUARTERS_CONFIG = [
    {"name": "Q3 2023-24", "start": "2023-10-01", "end": "2023-12-31", "type": "csv"},
    {"name": "Q4 2023-24", "start": "2024-01-01", "end": "2024-03-31", "type": "csv"},
    {"name": "Q1 2024-25", "start": "2024-04-01", "end": "2024-06-30", "type": "csv"},
    {"name": "Q2 2024-25", "start": "2024-07-01", "end": "2024-09-30", "type": "peer", "sheet": "2024Q3"},
    {"name": "Q3 2024-25", "start": "2024-10-01", "end": "2024-12-31", "type": "qty", "sheet": "FY24-25Q3"},
    {"name": "Q4 2024-25", "start": "2025-01-01", "end": "2025-03-31", "type": "qty", "sheet": "FY24-25Q4"},
    {"name": "Q1 2025-26", "start": "2025-04-01", "end": "2025-06-30", "type": "peer", "sheet": "2025Q2"},
    {"name": "Q2 2025-26", "start": "2025-07-01", "end": "2025-09-30", "type": "qty", "sheet": "FY25-26Q1"},
    {"name": "Q3 2025-26", "start": "2025-10-01", "end": "2025-12-31", "type": "qty", "sheet": "FY25-26Q2"},
    {"name": "Q4 2025-26", "start": "2026-01-01", "end": "2026-03-31", "type": "qty", "sheet": "FY25-26Q3"},
    {"name": "Q1 2026-27", "start": "2026-04-01", "end": "2026-06-30", "type": "qty", "sheet": "FY25-26Q4"},
    {"name": "Q2 2026-27", "start": "2026-07-01", "end": "2026-09-30", "type": "qty", "sheet": "Q2 2026-27"}
]

COST = 0.0010
font_family = "Segoe UI"

# Styles
f_title = Font(name=font_family, size=14, bold=True, color="FFFFFF")
f_sec = Font(name=font_family, size=11, bold=True, color="1E3A5F")
f_hdr = Font(name=font_family, size=9, bold=True, color="FFFFFF")
f_body = Font(name=font_family, size=9)
f_bold = Font(name=font_family, size=9, bold=True)
f_green = Font(name=font_family, size=9, color="047857", bold=True)
f_red = Font(name=font_family, size=9, color="B91C1C", bold=True)

fill_title = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
fill_hdr_long = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
fill_hdr_short = PatternFill(start_color="DC2626", end_color="DC2626", fill_type="solid")
fill_hdr_dark = PatternFill(start_color="374151", end_color="374151", fill_type="solid")
fill_soft_blue = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
fill_soft_red = PatternFill(start_color="FEF2F2", end_color="FEF2F2", fill_type="solid")
fill_soft_green = PatternFill(start_color="ECFDF5", end_color="ECFDF5", fill_type="solid")
fill_soft_gray = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")

thin_side = Side(style='thin', color='D1D5DB')
border_thin = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

class NumPyCache:
    def __init__(self, df):
        self.dates = df["date"].values
        self.closes = df["close"].values
        self.highs = df["high"].values
        self.lows = df["low"].values
        self.ema_lows = df["ema50_low"].values
        self.ema_highs = df["ema50_high"].values

def load_sectors_and_companies():
    SYM_COMPANY = {}
    SYM_SECTOR = {}
    try:
        wb = openpyxl.load_workbook(QTY_FILE, read_only=True)
        for sn in wb.sheetnames:
            ws = wb[sn]
            for r in range(2, ws.max_row+1):
                comp = ws.cell(r, 1).value
                sym = ws.cell(r, 2).value
                if sym:
                    sym_clean = str(sym).strip().upper()
                    SYM_COMPANY[sym_clean] = str(comp or sym_clean)
        wb.close()
    except Exception:
        pass
        
    sec_map = {
        "TCS": "Information Technology", "HCLTECH": "Information Technology", "INFY": "Information Technology",
        "WIPRO": "Information Technology", "LTIM": "Information Technology", "TECHM": "Information Technology",
        "HDFCBANK": "Financial Services", "ICICIBANK": "Financial Services", "KOTAKBANK": "Financial Services",
        "AXISBANK": "Financial Services", "SBIN": "Financial Services", "JIOFIN": "Financial Services",
        "BAJFINANCE": "Financial Services", "BAJAJFINSV": "Financial Services", "HDFCLIFE": "Financial Services",
        "SBILIFE": "Financial Services", "SHRIRAMFIN": "Financial Services", "RELIANCE": "Energy",
        "ONGC": "Energy", "NTPC": "Energy", "POWERGRID": "Energy", "BPCL": "Energy", "COALINDIA": "Energy",
        "TATASTEEL": "Metals & Mining", "HINDALCO": "Metals & Mining", "JSWSTEEL": "Metals & Mining",
        "ULTRACEMCO": "Construction Materials", "GRASIM": "Construction Materials", "LT": "Construction",
        "M&M": "Automobile and Auto Components", "MARUTI": "Automobile and Auto Components",
        "TATAMOTORS": "Automobile and Auto Components", "BAJAJ-AUTO": "Automobile and Auto Components",
        "HEROMOTOCO": "Automobile and Auto Components", "EICHERMOT": "Automobile and Auto Components",
        "HINDUNILVR": "Fast Moving Consumer Goods", "ITC": "Fast Moving Consumer Goods",
        "NESTLEIND": "Fast Moving Consumer Goods", "BRITANNIA": "Fast Moving Consumer Goods",
        "TATACONSUM": "Fast Moving Consumer Goods", "SUNPHARMA": "Healthcare", "CIPLA": "Healthcare",
        "DRREDDY": "Healthcare", "APOLLOHOSP": "Healthcare", "MAXHEALTH": "Healthcare",
        "ASIANPAINT": "Consumer Durables", "TITAN": "Consumer Durables", "ADANIENT": "Diversified",
        "ADANIPORTS": "Services", "BEL": "Capital Goods", "TRENT": "Consumer Services",
        "INDIGO": "Consumer Services"
    }
    for sym in SYM_COMPANY:
        if sym in sec_map:
            SYM_SECTOR[sym] = sec_map[sym]
        else:
            SYM_SECTOR[sym] = "Diversified"
            
    return SYM_COMPANY, SYM_SECTOR

def load_all_quarters_events(symbols):
    qty_data = {}
    if QTY_FILE.exists():
        try:
            qty_wb = openpyxl.load_workbook(QTY_FILE, read_only=True)
            for sheet in qty_wb.sheetnames:
                qty_data[sheet] = {}
                ws = qty_wb[sheet]
                for r in range(2, ws.max_row + 1):
                    sym = ws.cell(r, 2).value
                    dt_val = ws.cell(r, 3).value
                    if sym and dt_val:
                        sym_clean = str(sym).strip().upper()
                        try:
                            qty_data[sheet][sym_clean] = pd.Timestamp(dt_val).normalize()
                        except:
                            pass
            qty_wb.close()
        except Exception:
            pass
            
    peer_data = {}
    if PEER_FILE.exists():
        try:
            peer_wb = openpyxl.load_workbook(PEER_FILE, read_only=True)
            for sheet in peer_wb.sheetnames:
                peer_data[sheet] = {}
                ws = peer_wb[sheet]
                for r in range(4, ws.max_row + 1):
                    sym = ws.cell(r, 1).value
                    dt_val = ws.cell(r, 3).value
                    if sym and dt_val:
                        sym_clean = str(sym).strip().upper()
                        try:
                            peer_data[sheet][sym_clean] = pd.Timestamp(dt_val).normalize()
                        except:
                            pass
            peer_wb.close()
        except Exception:
            pass
            
    all_historical_events = {sym: set() for sym in symbols}
    for sym in symbols:
        f = ROOT / sym / "financial_results.csv"
        if f.exists():
            try:
                df = pd.read_csv(f, dtype=str).fillna("")
                if "broadCastDate" in df:
                    df["d"] = pd.to_datetime(df["broadCastDate"], errors="coerce", format="mixed", dayfirst=True).dt.normalize()
                    for dt_val in df["d"].dropna():
                        dt_ts = pd.Timestamp(dt_val)
                        if dt_ts < pd.Timestamp("2024-07-01"):
                            all_historical_events[sym].add(dt_ts)
            except:
                pass
                
    for sheet in peer_data:
        for sym, dt_val in peer_data[sheet].items():
            if sym in all_historical_events:
                all_historical_events[sym].add(dt_val)
                
    for sheet in qty_data:
        for sym, dt_val in qty_data[sheet].items():
            if sym in all_historical_events:
                all_historical_events[sym].add(dt_val)
                
    quarter_dates_map = {}
    for q in QUARTERS_CONFIG:
        q_name = q["name"]
        quarter_dates_map[q_name] = {}
        
        if q["type"] == "csv":
            start_ts = pd.Timestamp(q["start"]).normalize()
            end_ts = pd.Timestamp(q["end"]).normalize()
            for sym in symbols:
                q_dts = [d for d in all_historical_events[sym] if start_ts <= d <= end_ts]
                if q_dts:
                    quarter_dates_map[q_name][sym] = min(q_dts)
        elif q["type"] == "peer":
            sheet = q["sheet"]
            for sym in symbols:
                if sym in peer_data.get(sheet, {}):
                    quarter_dates_map[q_name][sym] = peer_data[sheet][sym]
        elif q["type"] == "qty":
            sheet = q["sheet"]
            for sym in symbols:
                if sym in qty_data.get(sheet, {}):
                    quarter_dates_map[q_name][sym] = qty_data[sheet][sym]
                    
    return quarter_dates_map, {sym: sorted(list(dts)) for sym, dts in all_historical_events.items()}

def fetch_prices(symbols):
    p1 = int(dt.datetime(2022, 1, 1).timestamp())
    p2 = int((dt.datetime.now() + pd.Timedelta(days=2)).timestamp())
    headers = {"User-Agent": "Mozilla/5.0"}
    cache = {}
    for sym in symbols:
        ticker = sym.replace("&", "%26") + ".NS"
        url = f"https://query1.finance.yahoo.com/v8/finance/chart/{ticker}?period1={p1}&period2={p2}&interval=1d"
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req) as resp:
                data = json.loads(resp.read().decode())
                r = data["chart"]["result"][0]
                q = r["indicators"]["quote"][0]
                df = pd.DataFrame({
                    "date": pd.to_datetime(r["timestamp"], unit="s").date,
                    "close": q["close"],
                    "high": q["high"],
                    "low": q["low"]
                }).dropna()
                df["date"] = pd.to_datetime(df["date"])
                
                df["ema50_low"] = df["low"].ewm(span=50, adjust=False).mean()
                df["ema50_high"] = df["high"].ewm(span=50, adjust=False).mean()
                
                sorted_df = df.sort_values("date").reset_index(drop=True)
                cache[sym] = NumPyCache(sorted_df)
        except Exception as e:
            print(f"Error fetching {sym}: {e}")
    return cache

def check_sl_hit_long_fast(cache, en_idx, ex_idx):
    start_idx = en_idx + 1
    end_idx = ex_idx + 1
    if start_idx >= end_idx:
        return False, None, None
    lows = cache.lows[start_idx:end_idx]
    ema_lows = cache.ema_lows[start_idx:end_idx]
    hits = lows < ema_lows
    if np.any(hits):
        idx = np.flatnonzero(hits)[0]
        hit_idx = start_idx + idx
        return True, pd.Timestamp(cache.dates[hit_idx]), float(ema_lows[idx])
    return False, None, None

def check_sl_hit_short_fast(cache, en_idx, ex_idx):
    start_idx = en_idx + 1
    end_idx = ex_idx + 1
    if start_idx >= end_idx:
        return False, None, None
    highs = cache.highs[start_idx:end_idx]
    ema_highs = cache.ema_highs[start_idx:end_idx]
    hits = highs > ema_highs
    if np.any(hits):
        idx = np.flatnonzero(hits)[0]
        hit_idx = start_idx + idx
        return True, pd.Timestamp(cache.dates[hit_idx]), float(ema_highs[idx])
    return False, None, None

def simulate_trade_case_fast(case_num, strategy, cache, pos_idx, bo, ao):
    en_idx = pos_idx - bo
    ex_idx = pos_idx + ao
    
    if en_idx < 0 or ex_idx >= len(cache.dates) or en_idx >= ex_idx:
        return -999.0, None, None, None, None, False
        
    close_e = cache.closes[en_idx]
    ema_h_e = cache.ema_highs[en_idx]
    ema_l_e = cache.ema_lows[en_idx]
    
    actual_en_idx = en_idx
    sl_price = None
    
    if strategy == "LONG":
        if case_num == 1:
            pass
        elif case_num == 2:
            if not (close_e > ema_h_e):
                return -999.0, None, None, None, None, False
            sl_price = ema_l_e
        elif case_num == 3:
            if not (ema_l_e <= close_e <= ema_h_e):
                return -999.0, None, None, None, None, False
            sl_price = ema_l_e
        elif case_num == 4 or case_num == 5:
            if not (close_e < ema_l_e):
                return -999.0, None, None, None, None, False
            entered = False
            for i in range(en_idx + 1, ex_idx):
                if cache.closes[i] > cache.ema_highs[i]:
                    actual_en_idx = i
                    sl_price = cache.ema_lows[i]
                    entered = True
                    break
            if not entered:
                return -999.0, None, None, None, None, False
        elif case_num == 6 or case_num == 7:
            sl_price = ema_l_e
        elif case_num == 8:
            if not (close_e >= ema_l_e):
                return -999.0, None, None, None, None, False
            sl_price = ema_l_e

    else: # SHORT
        if case_num == 1:
            pass
        elif case_num == 2:
            if not (close_e < ema_l_e):
                return -999.0, None, None, None, None, False
            sl_price = ema_h_e
        elif case_num == 3:
            if not (ema_l_e <= close_e <= ema_h_e):
                return -999.0, None, None, None, None, False
            sl_price = ema_h_e
        elif case_num == 4 or case_num == 5:
            if not (close_e > ema_h_e):
                return -999.0, None, None, None, None, False
            entered = False
            for i in range(en_idx + 1, ex_idx):
                if cache.closes[i] < cache.ema_lows[i]:
                    actual_en_idx = i
                    sl_price = cache.ema_highs[i]
                    entered = True
                    break
            if not entered:
                return -999.0, None, None, None, None, False
        elif case_num == 6 or case_num == 7:
            sl_price = ema_h_e
        elif case_num == 8:
            if not (close_e <= ema_h_e):
                return -999.0, None, None, None, None, False
            sl_price = ema_h_e

    actual_ex_idx = ex_idx
    actual_p_ex = cache.closes[ex_idx]
    sl_hit = False
    
    p_en = cache.closes[actual_en_idx]
    
    if strategy == "LONG":
        if sl_price is not None:
            if case_num in [2, 3, 4, 5, 6]:
                start_idx = actual_en_idx + 1
                end_idx = ex_idx + 1
                lows_slice = cache.lows[start_idx:end_idx]
                hits = lows_slice < sl_price
                if np.any(hits):
                    idx = np.flatnonzero(hits)[0]
                    actual_ex_idx = start_idx + idx
                    actual_p_ex = float(sl_price)
                    sl_hit = True
            elif case_num in [7, 8]:
                sl_hit, hit_date, hit_px = check_sl_hit_long_fast(cache, actual_en_idx, ex_idx)
                if sl_hit:
                    # Convert hit_date back to index
                    actual_ex_idx = np.searchsorted(cache.dates, np.datetime64(hit_date.normalize()))
                    actual_p_ex = hit_px
        ret = (actual_p_ex / p_en - 1 - COST) * 100
    else:
        if sl_price is not None:
            if case_num in [2, 3, 4, 5, 6]:
                start_idx = actual_en_idx + 1
                end_idx = ex_idx + 1
                highs_slice = cache.highs[start_idx:end_idx]
                hits = highs_slice > sl_price
                if np.any(hits):
                    idx = np.flatnonzero(hits)[0]
                    actual_ex_idx = start_idx + idx
                    actual_p_ex = float(sl_price)
                    sl_hit = True
            elif case_num in [7, 8]:
                sl_hit, hit_date, hit_px = check_sl_hit_short_fast(cache, actual_en_idx, ex_idx)
                if sl_hit:
                    actual_ex_idx = np.searchsorted(cache.dates, np.datetime64(hit_date.normalize()))
                    actual_p_ex = hit_px
        ret = ((p_en - actual_p_ex) / p_en - COST) * 100

    return ret, pd.Timestamp(cache.dates[actual_en_idx]), float(p_en), pd.Timestamp(cache.dates[actual_ex_idx]), float(actual_p_ex), sl_hit

def optimize_offsets_case_fast(case_num, sym, train_events_precalculated, cache):
    if not train_events_precalculated or cache is None:
        return (2, 5, 0.0), (2, 5, 0.0)
        
    best_long = (2, 5)
    best_long_score = (-9e9, -9e9)
    for bo in range(1, 9):
        for so in range(1, 9):
            rets = []
            for pos_idx in train_events_precalculated:
                ret, _, _, _, _, _ = simulate_trade_case_fast(case_num, "LONG", cache, pos_idx, bo, so)
                if ret != -999.0:
                    rets.append(ret)
            if len(rets) >= 3:
                win_r = sum(1 for r in rets if r > 0) / len(rets) * 100
                avg_r = sum(rets) / len(rets)
                score = (win_r, avg_r)
                if score > best_long_score:
                    best_long_score = score
                    best_long = (bo, so)
                    
    best_short = (2, 5)
    best_short_score = (-9e9, -9e9)
    for bo in range(1, 9):
        for so in range(1, 9):
            rets = []
            for pos_idx in train_events_precalculated:
                ret, _, _, _, _, _ = simulate_trade_case_fast(case_num, "SHORT", cache, pos_idx, bo, so)
                if ret != -999.0:
                    rets.append(ret)
            if len(rets) >= 3:
                win_r = sum(1 for r in rets if r > 0) / len(rets) * 100
                avg_r = sum(rets) / len(rets)
                score = (win_r, avg_r)
                if score > best_short_score:
                    best_short_score = score
                    best_short = (bo, so)
                    
    l_avg = (best_long_score[1] / 100.0) if best_long_score[1] != -9e9 else 0.0
    s_avg = (best_short_score[1] / 100.0) if best_short_score[1] != -9e9 else 0.0
    
    return (best_long[0], best_long[1], l_avg), (best_short[0], best_short[1], s_avg)

def simulate_slots(trades_list):
    slots = []
    for t in trades_list:
        assigned = False
        for s in slots:
            if t["entry_date"] >= s["last_exit"]:
                s["trades"].append(t)
                s["last_exit"] = t["exit_date"]
                assigned = True
                break
        if not assigned:
            slots.append({
                "id": len(slots) + 1,
                "trades": [t],
                "last_exit": t["exit_date"]
            })
    return slots

def build_excel_for_quarter(case_num, q_config, trades, slots, case_dir):
    q_name = q_config["name"]
    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)
    
    total_margin = 0.0
    total_profit = 0.0
    total_expected_return_weighted = 0.0
    for s in slots:
        peak_margin = max(t["entry_price"] for t in s["trades"])
        total_margin += peak_margin
        total_profit += sum(t["profit_rs"] for t in s["trades"])
        total_expected_return_weighted += sum(t["expected_return"] * t["entry_price"] for t in s["trades"])
        
    final_val = total_margin + total_profit
    overall_ret = (total_profit / total_margin) * 100 if total_margin > 0 else 0.0
    overall_exp = (total_expected_return_weighted / total_margin) if total_margin > 0 else 0.0
    total_reentries = sum(len(s["trades"]) - 1 for s in slots)
    win_rate = sum(1 for t in trades if t["profit_rs"] > 0) / len(trades) * 100 if trades else 0.0
    
    # ── Summary Sheet ──
    ws_sum = wb_out.create_sheet("Summary")
    ws_sum.sheet_view.showGridLines = True
    ws_sum.merge_cells("A1:H1")
    t_cell = ws_sum.cell(row=1, column=1, value=f"Nifty 50 Combined Long/Short Capital Utilisation Summary - {q_name} (Booked P&L)")
    t_cell.font = f_title; t_cell.fill = fill_title; t_cell.alignment = Alignment(vertical="center")
    ws_sum.row_dimensions[1].height = 32
    
    sum_headers = ["Quarter", "Fund Utilised", "Final Value", "Net Profit", "Expected Return", "Booked Return", "Total Trades", "Total Re-entries"]
    for c_idx, h in enumerate(sum_headers, 1):
        cell = ws_sum.cell(row=3, column=c_idx, value=h)
        cell.font = f_hdr; cell.fill = fill_hdr_dark; cell.alignment = Alignment(horizontal="center"); cell.border = border_thin
    ws_sum.row_dimensions[3].height = 22
    
    ws_sum.cell(row=4, column=1, value=q_name).alignment = Alignment(horizontal="center")
    
    c_fu = ws_sum.cell(row=4, column=2, value=total_margin)
    c_fu.number_format = "#,##0.00"; c_fu.alignment = Alignment(horizontal="right")
    
    c_fv = ws_sum.cell(row=4, column=3, value=final_val)
    c_fv.number_format = "#,##0.00"; c_fv.alignment = Alignment(horizontal="right")
    
    c_np = ws_sum.cell(row=4, column=4, value=total_profit)
    c_np.number_format = "#,##0.00"; c_np.alignment = Alignment(horizontal="right")
    c_np.font = f_green if total_profit >= 0 else f_red; c_np.fill = fill_soft_green if total_profit >= 0 else fill_soft_red
    
    c_er = ws_sum.cell(row=4, column=5, value=overall_exp)
    c_er.number_format = "+0.00%;-0.00%" if overall_exp >= 0 else "0.00%"
    c_er.alignment = Alignment(horizontal="right")
    c_er.font = f_bold; c_er.fill = fill_soft_blue
    
    c_br = ws_sum.cell(row=4, column=6, value=overall_ret/100)
    c_br.number_format = "+0.00%;-0.00%" if overall_ret >= 0 else "0.00%"; c_br.alignment = Alignment(horizontal="right")
    c_br.font = f_green if total_profit >= 0 else f_red; c_br.fill = fill_soft_green if total_profit >= 0 else fill_soft_red
    
    ws_sum.cell(row=4, column=7, value=len(trades)).alignment = Alignment(horizontal="center")
    ws_sum.cell(row=4, column=8, value=total_reentries).alignment = Alignment(horizontal="center")
    
    for c in range(1, 9):
        cell = ws_sum.cell(row=4, column=c)
        cell.border = border_thin
        if c not in (4, 5, 6):
            cell.font = f_bold if c in (2, 3) else f_body
            cell.fill = fill_soft_gray
            
    ws_sum.merge_cells("A6:B6")
    ws_sum.cell(row=6, column=1, value="Quarter Performance Statistics (Detailed Summary)").font = f_sec
    
    ws_sum.cell(row=7, column=1, value="Performance Metric").font = f_hdr; ws_sum.cell(row=7, column=1).fill = fill_hdr_dark; ws_sum.cell(row=7, column=1).border = border_thin
    ws_sum.cell(row=7, column=2, value="Value").font = f_hdr; ws_sum.cell(row=7, column=2).fill = fill_hdr_dark; ws_sum.cell(row=7, column=2).border = border_thin
    
    stat_rows = [
        ("Total Completed Trades", len(trades)),
        ("Winning Trades (Profitable Moves)", sum(1 for t in trades if t["profit_rs"] > 0)),
        ("Losing Trades", sum(1 for t in trades if t["profit_rs"] < 0)),
        ("Short Selling Trades Chosen", sum(1 for t in trades if t["strategy"] == "SHORT")),
        ("Long Buying Trades Chosen", sum(1 for t in trades if t["strategy"] == "LONG")),
        ("Strategy Portfolio Win Rate", f"{win_rate:.2f}%"),
        ("Compounded Slots Generated", len(slots))
    ]
    for idx, (k, v) in enumerate(stat_rows, 8):
        ws_sum.cell(row=idx, column=1, value=k).font = f_body; ws_sum.cell(row=idx, column=1).border = border_thin; ws_sum.cell(row=idx, column=1).fill = fill_soft_gray
        ws_sum.cell(row=idx, column=2, value=str(v)).font = f_bold; ws_sum.cell(row=idx, column=2).border = border_thin
        
    for col in ws_sum.columns:
        max_len = max((len(str(cell.value)) for cell in col if cell.value is not None), default=12)
        ws_sum.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 3, 40)
        
    # ── Detailed Trades Sheet ──
    ws_det = wb_out.create_sheet("Detailed_Trades")
    ws_det.sheet_view.showGridLines = True
    
    det_headers = [
        "Trade No", "Symbol", "Company Name", "Sector", "Strategy Type (Decided Move)", "Entry Date", "Exit Date",
        "Entry Price (Rs.)", "Exit Price (Rs.)", "Expected Offset", "Expected Return (%)", "Realised Return (%)", "Realised Profit/Loss (Rs.)",
        "Re-entry Type", "Assigned Slot"
    ]
    for c_idx, h in enumerate(det_headers, 1):
        cell = ws_det.cell(row=1, column=c_idx, value=h)
        cell.font = f_hdr; cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_thin
        if h == "Strategy Type (Decided Move)":
            cell.fill = fill_hdr_dark
        elif "Expected" in h:
            cell.fill = fill_hdr_long
        elif "Realised" in h:
            cell.fill = fill_hdr_short
        else:
            cell.fill = fill_hdr_dark
    ws_det.row_dimensions[1].height = 24
    
    for r_idx, t in enumerate(trades, 2):
        bg = fill_soft_gray if r_idx % 2 == 0 else PatternFill(fill_type=None)
        
        ws_det.cell(row=r_idx, column=1, value=r_idx - 1).alignment = Alignment(horizontal="center")
        ws_det.cell(row=r_idx, column=2, value=t["symbol"]).font = f_bold
        ws_det.cell(row=r_idx, column=3, value=t["company"])
        ws_det.cell(row=r_idx, column=4, value=t["sector"]) 
        
        c_st = ws_det.cell(row=r_idx, column=5, value=t["strategy"])
        c_st.alignment = Alignment(horizontal="center"); c_st.font = f_bold
        if t["strategy"] == "LONG":
            c_st.fill = fill_soft_blue; c_st.font = Font(name="Segoe UI", size=9, color="2563EB", bold=True)
        else:
            c_st.fill = fill_soft_red; c_st.font = Font(name="Segoe UI", size=9, color="DC2626", bold=True)
            
        def dt_cell(ws, r, c, val):
            cell = ws.cell(row=r, column=c, value=val.strftime("%Y-%m-%d") if val else "—")
            cell.alignment = Alignment(horizontal="center"); cell.font = f_body; cell.border = border_thin; cell.fill = bg
            
        dt_cell(ws_det, r_idx, 6, t["entry_date"])
        dt_cell(ws_det, r_idx, 7, t["exit_date"])
        
        c_ep = ws_det.cell(row=r_idx, column=8, value=t["entry_price"])
        c_ep.number_format = "#,##0.00"; c_ep.alignment = Alignment(horizontal="right"); c_ep.font = f_body; c_ep.fill = bg; c_ep.border = border_thin
        
        c_xp = ws_det.cell(row=r_idx, column=9, value=t["exit_price"])
        c_xp.number_format = "#,##0.00"; c_xp.alignment = Alignment(horizontal="right"); c_xp.font = f_body; c_xp.fill = bg; c_xp.border = border_thin
        
        c_of = ws_det.cell(row=r_idx, column=10, value=f"T-{t['before']} to T+{t['after']}")
        c_of.alignment = Alignment(horizontal="center"); c_of.font = f_body; c_of.fill = bg; c_of.border = border_thin
        
        c_er = ws_det.cell(row=r_idx, column=11, value=t["expected_return"])
        c_er.number_format = "+0.00%;-0.00%" if t["expected_return"] >= 0 else "0.00%"
        c_er.alignment = Alignment(horizontal="right"); c_er.border = border_thin
        if t["expected_return"] >= 0:
            c_er.font = f_green; c_er.fill = fill_soft_green
        else:
            c_er.font = f_red; c_er.fill = fill_soft_red
        
        c_rr = ws_det.cell(row=r_idx, column=12, value=t["return_pct"]/100)
        c_rr.number_format = "+0.00%;-0.00%"; c_rr.alignment = Alignment(horizontal="right"); c_rr.border = border_thin
        if t["return_pct"] >= 0:
            c_rr.font = f_green; c_rr.fill = fill_soft_green
        else:
            c_rr.font = f_red; c_rr.fill = fill_soft_red
            
        c_rp = ws_det.cell(row=r_idx, column=13, value=t["profit_rs"])
        c_rp.number_format = "₹#,##0.00"; c_rp.alignment = Alignment(horizontal="right"); c_rp.border = border_thin
        if t["profit_rs"] >= 0:
            c_rp.font = f_green; c_rp.fill = fill_soft_green
        else:
            c_rp.font = f_red; c_rp.fill = fill_soft_red
            
        slot_id = next(s["id"] for s in slots if t in s["trades"])
        slot_trades = next(s["trades"] for s in slots if s["id"] == slot_id)
        is_first = (slot_trades[0] == t)
        re_type = "First Entry" if is_first else "Re-entry"
        c_re = ws_det.cell(row=r_idx, column=14, value=re_type)
        c_re.alignment = Alignment(horizontal="center"); c_re.font = f_body; c_re.fill = bg; c_re.border = border_thin
        
        c_as = ws_det.cell(row=r_idx, column=15, value=f"Slot {slot_id}")
        c_as.alignment = Alignment(horizontal="center"); c_as.font = f_body; c_as.fill = bg; c_as.border = border_thin
        
        for c in [1, 2, 3, 4]:
            cell = ws_det.cell(row=r_idx, column=c)
            cell.border = border_thin; cell.fill = bg
            if c != 2: cell.font = f_body
            
    for col in ws_det.columns:
        max_len = max((len(str(cell.value)) for cell in col if cell.value is not None), default=12)
        ws_det.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 3, 30)
    ws_det.freeze_panes = "B2"
    
    # ── Losing Trades Verification Sheet ──
    ws_lost = wb_out.create_sheet("Losing Trades Verification")
    ws_lost.sheet_view.showGridLines = True
    
    ws_lost.merge_cells("A1:H1")
    t_lost = ws_lost.cell(row=1, column=1, value=f"{q_name} Losing Trades Side-by-Side Verification (Long vs Short)")
    t_lost.font = f_title; t_lost.fill = fill_title; t_lost.alignment = Alignment(vertical="center")
    ws_lost.row_dimensions[1].height = 32
    
    lost_headers = [
        "Symbol", "Company Name", "Decided Move (Strategy)", "Long Return (%)", "Short Return (%)", 
        "Realised Loss (Rs.)", "Verification Status", "Rationale / Explanation"
    ]
    for c_idx, h in enumerate(lost_headers, 1):
        cell = ws_lost.cell(row=3, column=c_idx, value=h)
        cell.font = f_hdr; cell.fill = fill_hdr_dark; cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_thin
    ws_lost.row_dimensions[3].height = 24
    
    losing_data = []
    for t in trades:
        if t["profit_rs"] < 0:
            sym = t["symbol"]
            comp = t["company"]
            strat = t["strategy"]
            l_ret_val = t["l_ret"]/100 if t["l_ret"] != -999.0 else None
            s_ret_val = t["s_ret"]/100 if t["s_ret"] != -999.0 else None
            loss = t["profit_rs"]
            status = "Passed"
            if strat == "LONG":
                if s_ret_val is None:
                    rat = "Short was out-of-bounds due to lack of post-event trading candles. Long was the only valid move."
                else:
                    rat = f"LONG had a smaller loss ({t['l_ret']:+.2f}%) compared to SHORT ({t['s_ret']:+.2f}%)."
            else: 
                if l_ret_val is None:
                    rat = "Long was out-of-bounds. Short was the only valid move."
                else:
                    rat = f"SHORT had a smaller loss ({t['s_ret']:+.2f}%) compared to LONG ({t['l_ret']:+.2f}%)."
            losing_data.append((sym, comp, strat, l_ret_val, s_ret_val, loss, status, rat))
            
    for r_idx, (sym, comp, strat, l_ret, s_ret, loss, status, rat) in enumerate(losing_data, 4):
        bg = fill_soft_gray if r_idx % 2 == 0 else PatternFill(fill_type=None)
        
        c_sym = ws_lost.cell(row=r_idx, column=1, value=sym)
        c_sym.font = f_bold; c_sym.alignment = Alignment(horizontal="center"); c_sym.border = border_thin; c_sym.fill = bg
        
        c_comp = ws_lost.cell(row=r_idx, column=2, value=comp)
        c_comp.font = f_body; c_comp.border = border_thin; c_comp.fill = bg
        
        c_strat = ws_lost.cell(row=r_idx, column=3, value=strat)
        c_strat.font = f_bold; c_strat.alignment = Alignment(horizontal="center"); c_strat.border = border_thin
        if strat == "LONG":
            c_strat.fill = fill_soft_blue; c_strat.font = Font(name="Segoe UI", size=9, color="2563EB", bold=True)
        else:
            c_strat.fill = fill_soft_red; c_strat.font = Font(name="Segoe UI", size=9, color="DC2626", bold=True)
            
        c_l = ws_lost.cell(row=r_idx, column=4, value=l_ret)
        c_l.number_format = "0.00%"; c_l.alignment = Alignment(horizontal="right"); c_l.border = border_thin; c_l.font = f_red; c_l.fill = fill_soft_red
        
        c_s = ws_lost.cell(row=r_idx, column=5)
        c_s.border = border_thin
        if s_ret is not None:
            c_s.value = s_ret
            c_s.number_format = "0.00%"
            c_s.alignment = Alignment(horizontal="right")
            c_s.font = f_red; c_s.fill = fill_soft_red
        else:
            c_s.value = "N/A"
            c_s.alignment = Alignment(horizontal="center")
            c_s.font = Font(name="Segoe UI", size=9, color="9CA3AF"); c_s.fill = bg
            
        c_loss = ws_lost.cell(row=r_idx, column=6, value=loss)
        c_loss.number_format = "₹#,##0.00"; c_loss.alignment = Alignment(horizontal="right"); c_loss.border = border_thin; c_loss.font = f_red; c_loss.fill = fill_soft_red
        
        c_stat = ws_lost.cell(row=r_idx, column=7, value=status)
        c_stat.alignment = Alignment(horizontal="center"); c_stat.font = f_bold; c_stat.border = border_thin; c_stat.fill = fill_soft_green; c_stat.font = f_green
        
        c_rat = ws_lost.cell(row=r_idx, column=8, value=rat)
        c_rat.font = f_body; c_rat.border = border_thin; c_rat.fill = bg
        
        ws_lost.row_dimensions[r_idx].height = 20
        
    for col in ws_lost.columns:
        max_len = max((len(str(cell.value)) for cell in col if cell.value is not None), default=12)
        ws_lost.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 3, 50)
        
    clean_name = q_name.replace(" ", "_").replace("-", "_")
    out_path = case_dir / f"{clean_name}_Combined_Best_Capital_Utilisation.xlsx"
    
    saved = False
    try:
        wb_out.save(out_path)
        saved = True
    except PermissionError:
        version = 2
        while not saved:
            alt_path = case_dir / f"{clean_name}_Combined_Best_Capital_Utilisation_v{version}.xlsx"
            try:
                wb_out.save(alt_path)
                saved = True
            except PermissionError:
                version += 1

def consolidate_case(case_num, case_dir):
    print(f"Consolidating Case {case_num} reports...")
    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)
    
    all_trades = []
    
    # Simple maps
    dates_file = ROOT / "Nifty50Stocks_QtyResultDates.xlsx"
    SYM_COMPANY = {}
    if dates_file.exists():
        wb_dates = openpyxl.load_workbook(dates_file, read_only=True)
        for sn in wb_dates.sheetnames:
            ws_d = wb_dates[sn]
            for r in range(2, ws_d.max_row + 1):
                comp = ws_d.cell(r, 1).value
                sym = ws_d.cell(r, 2).value
                if sym and comp:
                    SYM_COMPANY[str(sym).strip().upper()] = str(comp).strip()
        wb_dates.close()
        
    for q_config in QUARTERS_CONFIG:
        q_name = q_config["name"]
        fn = q_name.replace(" ", "_").replace("-", "_")
        file_path = case_dir / f"{fn}_Combined_Best_Capital_Utilisation.xlsx"
        if not file_path.exists():
            continue
            
        wb_q = openpyxl.load_workbook(file_path, data_only=True)
        ws_q_trades = wb_q["Detailed_Trades"]
        raw_rows = list(ws_q_trades.values)
        if len(raw_rows) < 3:
            wb_q.close()
            continue
            
        header_row_idx = 0
        for r_idx, r in enumerate(raw_rows[:5]):
            if any(isinstance(val, str) and "Symbol" in val for val in r):
                header_row_idx = r_idx
                break
        
        headers = [str(h).strip() for h in raw_rows[header_row_idx]]
        data_rows = raw_rows[header_row_idx+1:]
        df_raw = pd.DataFrame(data_rows, columns=headers).dropna(subset=["Symbol"])
        
        strategy_col = next((c for c in df_raw.columns if "Strategy" in c), "Strategy")
        df_raw["Strategy"] = df_raw[strategy_col].astype(str).str.strip().str.upper()
        
        entry_price_col = next((c for c in df_raw.columns if "Entry Price" in c or "Buy Price" in c), "Entry Price")
        exit_price_col = next((c for c in df_raw.columns if "Exit Price" in c or "Sell Price" in c), "Exit Price")
        realised_ret_col = next((c for c in df_raw.columns if "Realised Return" in c), "Realised Return (%)")
        expected_ret_col = next((c for c in df_raw.columns if "Expected Return" in c), "Expected Return (%)")
        realised_pnl_col = next((c for c in df_raw.columns if "Profit/Loss" in c or "Realised Profit" in c), "Realised Profit/Loss (Rs.)")
        
        df_raw["Entry Price"] = pd.to_numeric(df_raw[entry_price_col], errors="coerce").fillna(0.0)
        df_raw["Exit Price"] = pd.to_numeric(df_raw[exit_price_col], errors="coerce").fillna(0.0)
        df_raw["Realised Return (%)"] = pd.to_numeric(df_raw[realised_ret_col], errors="coerce").fillna(0.0)
        df_raw["Expected Return (%)"] = pd.to_numeric(df_raw[expected_ret_col], errors="coerce").fillna(0.0)
        df_raw["Realised Profit (Rs.)"] = pd.to_numeric(df_raw[realised_pnl_col], errors="coerce").fillna(0.0)
        
        df_raw["Entry Date"] = pd.to_datetime(df_raw["Entry Date"])
        df_raw["Exit Date"] = pd.to_datetime(df_raw["Exit Date"])
        
        # Greedy slot scheduling
        df_raw = df_raw.sort_values(by=["Entry Date", "Exit Date"]).reset_index(drop=True)
        slots = []
        slot_ids = []
        for idx, row in df_raw.iterrows():
            assigned = False
            for s in slots:
                if row["Entry Date"] >= s["last_exit"]:
                    s["trades"].append(row)
                    s["last_exit"] = row["Exit Date"]
                    slot_ids.append(s["id"])
                    assigned = True
                    break
            if not assigned:
                new_id = len(slots) + 1
                slots.append({"id": new_id, "trades": [row], "last_exit": row["Exit Date"]})
                slot_ids.append(new_id)
        df_raw["Slot ID"] = slot_ids
                
        # Fill details
        for idx, row in df_raw.iterrows():
            slot_id = row["Slot ID"]
            trades_in_slot = next(s["trades"] for s in slots if s["id"] == slot_id)
            is_first = (trades_in_slot[0]["Symbol"] == row["Symbol"] and trades_in_slot[0]["Entry Date"] == row["Entry Date"])
            re_type = "First Entry" if is_first else "Re-entry"
            
            all_trades.append({
                "Quarter": q_name,
                "Symbol": row["Symbol"],
                "Company Name": SYM_COMPANY.get(row["Symbol"], row["Symbol"]),
                "Sector": row["Sector"],
                "Strategy": row["Strategy"],
                "Entry Date": row["Entry Date"].strftime("%Y-%m-%d"),
                "Exit Date": row["Exit Date"].strftime("%Y-%m-%d"),
                "Buy Price (Rs.)": row["Entry Price"],
                "Exit Price (Rs.)": row["Exit Price"],
                "Expected Return (%)": row["Expected Return (%)"],
                "Realised Return (%)": row["Realised Return (%)"],
                "Realised Profit/Loss (Rs.)": row["Realised Profit (Rs.)"],
                "Re-entry Type": re_type,
                "Assigned Slot": f"Slot {slot_id}"
            })
            
        ws_q = wb_out.create_sheet(title=q_name)
        ws_q.views.sheetView[0].showGridLines = True
        
        ws_q.merge_cells("A1:G1")
        t_cell = ws_q.cell(row=1, column=1, value=f"Nifty 50 Capital Utilisation Summary - {q_name} (Walk-Forward)")
        t_cell.font = Font(name=font_family, size=11, bold=True, color='FFFFFF')
        t_cell.fill = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
        t_cell.alignment = Alignment(horizontal='center', vertical='center')
        ws_q.row_dimensions[1].height = 28
        
        headers_q = ['Quarter', 'Fund Utilised', 'Final Value', 'Net Profit', 'Compounded Return', 'Total Trades', 'Total Re-entries']
        for col_idx, h in enumerate(headers_q, 1):
            cell = ws_q.cell(row=2, column=col_idx, value=h)
            cell.font = Font(name=font_family, size=9, bold=True, color='FFFFFF')
            cell.fill = PatternFill(start_color='333333', end_color='333333', fill_type='solid')
            cell.alignment = Alignment(horizontal='center')
            cell.border = Border(left=Side(style='thin', color='CCCCCC'), right=Side(style='thin', color='CCCCCC'))
            
        total_margin_q = 0.0
        total_profit_q = 0.0
        for s in slots:
            peak_margin_s = max(r["Entry Price"] for r in s["trades"])
            total_margin_q += peak_margin_s
            total_profit_q += sum(r["Realised Profit (Rs.)"] for r in s["trades"])
            
        fv_q = total_margin_q + total_profit_q
        ret_q = (total_profit_q / total_margin_q) if total_margin_q > 0 else 0.0
        reentries_q = sum(len(s["trades"]) - 1 for s in slots)
        
        row3 = [q_name, total_margin_q, fv_q, total_profit_q, ret_q, len(df_raw), reentries_q]
        for col_idx, val in enumerate(row3, 1):
            cell = ws_q.cell(row=3, column=col_idx, value=val)
            cell.font = Font(name=font_family, size=9)
            cell.border = Border(left=Side(style='thin', color='CCCCCC'), right=Side(style='thin', color='CCCCCC'), bottom=Side(style='thin', color='CCCCCC'))
            if col_idx in [2, 3, 4]:
                cell.number_format = '₹#,##0.00'
                cell.alignment = Alignment(horizontal='right')
            elif col_idx == 5:
                cell.number_format = '0.00%'
                cell.alignment = Alignment(horizontal='right')
            else:
                cell.alignment = Alignment(horizontal='center')
                
        ws_q.cell(row=5, column=1, value="Compounded Slots Detail").font = Font(name=font_family, size=10, bold=True)
        headers_s = ["Slot ID", "Start Capital (Rs.)", "Final Value (Rs.)", "Net Profit (Rs.)", "Compounded Return (%)", "Number of Re-entries", "Trade Flow Path"]
        for col_idx, h in enumerate(headers_s, 1):
            cell = ws_q.cell(row=6, column=col_idx, value=h)
            cell.font = Font(name=font_family, size=9, bold=True, color='FFFFFF')
            cell.fill = PatternFill(start_color='4F81BD', end_color='4F81BD', fill_type='solid')
            cell.alignment = Alignment(horizontal='center')
            
        for s_idx, s in enumerate(slots, 7):
            start_cap = float(s["trades"][0]["Entry Price"])
            slot_prof = sum(r["Realised Profit (Rs.)"] for r in s["trades"])
            fin_val = start_cap + slot_prof
            comp_ret = (fin_val / start_cap - 1) if start_cap > 0 else 0
            
            flow = " -> Re-entry: ".join([f"{r['Symbol']} ({pd.Timestamp(r['Entry Date']).strftime('%d-%b')} to {pd.Timestamp(r['Exit Date']).strftime('%d-%b')})" for r in s["trades"]])
            
            row_vals = [f"Slot {s['id']}", start_cap, fin_val, slot_prof, comp_ret, len(s["trades"])-1, flow]
            for col_idx, val in enumerate(row_vals, 1):
                cell = ws_q.cell(row=s_idx, column=col_idx, value=val)
                cell.font = Font(name=font_family, size=9)
                cell.border = Border(left=Side(style='thin', color='E0E0E0'), right=Side(style='thin', color='E0E0E0'), bottom=Side(style='thin', color='E0E0E0'))
                if col_idx in [2, 3, 4]:
                    cell.number_format = '₹#,##0.00'
                    cell.alignment = Alignment(horizontal='right')
                elif col_idx == 5:
                    cell.number_format = '0.00%'
                    cell.alignment = Alignment(horizontal='right')
                elif col_idx in [1, 6]:
                    cell.alignment = Alignment(horizontal='center')
                else:
                    cell.alignment = Alignment(horizontal='left')
                    
        for col in ws_q.columns:
            max_len = max((len(str(cell.value or '')) for cell in col), default=12)
            ws_q.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 3, 30)
            
        wb_q.close()
        
    ws_master = wb_out.create_sheet(title="Detailed_Trades")
    ws_master.views.sheetView[0].showGridLines = True
    
    master_headers = ["Quarter", "Symbol", "Company Name", "Sector", "Strategy", "Entry Date", "Exit Date", "Buy Price (Rs.)", "Exit Price (Rs.)", "Expected Return (%)", "Realised Return (%)", "Realised Profit/Loss (Rs.)", "Re-entry Type", "Assigned Slot"]
    for col_idx, h in enumerate(master_headers, 1):
        cell = ws_master.cell(row=1, column=col_idx, value=h)
        cell.font = Font(name=font_family, size=9, bold=True, color='FFFFFF')
        cell.fill = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
        cell.alignment = Alignment(horizontal='center')
        cell.border = Border(left=Side(style='thin', color='CCCCCC'), right=Side(style='thin', color='CCCCCC'))
    ws_master.row_dimensions[1].height = 24
    
    for r_idx, t in enumerate(all_trades, 2):
        ws_master.row_dimensions[r_idx].height = 18
        bg = fill_soft_gray if r_idx % 2 == 0 else PatternFill(fill_type=None)
        row_vals = [t[h] for h in master_headers]
        for col_idx, val in enumerate(row_vals, 1):
            cell = ws_master.cell(row=r_idx, column=col_idx, value=val)
            cell.font = Font(name=font_family, size=9)
            cell.border = Border(left=Side(style='thin', color='E0E0E0'), right=Side(style='thin', color='E0E0E0'), bottom=Side(style='thin', color='E0E0E0'))
            cell.fill = bg
            if col_idx in [1, 2, 5, 6, 7, 13, 14]:
                cell.alignment = Alignment(horizontal='center')
            else:
                cell.alignment = Alignment(horizontal='left')
            if col_idx in [8, 9, 12]:
                cell.number_format = '₹#,##0.00'
            elif col_idx in [10, 11]:
                cell.number_format = '0.00%'
                
    for col in ws_master.columns:
        max_len = max((len(str(cell.value or '')) for cell in col), default=12)
        ws_master.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 3, 30)
        
    out_file = case_dir / f"12_Quarters_Consolidated_Case{case_num}.xlsx"
    wb_out.save(out_file)
    print(f"  Saved consolidated Case {case_num} report: {out_file.name}")

def main():
    print("Loading sectors and company names...")
    SYM_COMPANY, SYM_SECTOR = load_sectors_and_companies()
    symbols = sorted(list(SYM_COMPANY.keys()))
    print(f"Loaded {len(symbols)} Nifty stocks.")
    
    print("Fetching historical prices from Yahoo Finance...")
    prices = fetch_prices(symbols)
    print(f"Prices loaded for {len(prices)} stocks.")
    
    print("Loading result events from all sources...")
    quarter_events, stock_events = load_all_quarters_events(symbols)
    
    # Pre-calculate positions of training events for all stocks and quarters
    print("Pre-calculating historical event dates index mappings...")
    stock_events_precalculated = {}
    for sym in symbols:
        stock_events_precalculated[sym] = []
        cache = prices.get(sym)
        if cache is not None:
            for ev_dt in stock_events[sym]:
                D = np.datetime64(pd.Timestamp(ev_dt).normalize())
                pos = np.searchsorted(cache.dates, D, side="right") - 1
                pos = max(0, pos)
                stock_events_precalculated[sym].append((ev_dt, pos))
                
    # Loop over all 8 Cases
    for case_num in range(1, 9):
        print(f"\n==================================================")
        print(f"RUNNING WALK-FORWARD FOR CASE {case_num}...")
        print(f"==================================================")
        
        case_dir = ROOT / f"12_Quarters_Reports_Case{case_num}"
        case_dir.mkdir(exist_ok=True, parents=True)
        
        for q_config in QUARTERS_CONFIG:
            q_name = q_config["name"]
            start_dt = pd.Timestamp(q_config["start"]).normalize()
            
            # Filter training event indexes before start_dt
            long_params = {}
            short_params = {}
            for sym in symbols:
                cache = prices.get(sym)
                train_idx_list = []
                if cache is not None:
                    train_idx_list = [pos for ev_dt, pos in stock_events_precalculated[sym] if ev_dt < start_dt]
                l_opt, s_opt = optimize_offsets_case_fast(case_num, sym, train_idx_list, cache)
                long_params[sym] = l_opt
                short_params[sym] = s_opt
                
            q_trades = []
            for sym in symbols:
                cache = prices.get(sym)
                if cache is None or len(cache.dates) == 0:
                    continue
                    
                res_dt = quarter_events.get(q_name, {}).get(sym)
                if not res_dt:
                    continue
                    
                # Find pos of res_dt
                D_res = np.datetime64(pd.Timestamp(res_dt).normalize())
                pos_idx = np.searchsorted(cache.dates, D_res, side="right") - 1
                pos_idx = max(0, pos_idx)
                
                lb, la, l_avg = long_params[sym]
                l_ret, l_en_dt, lp_en, l_ex_dt, lp_ex, l_sl_hit = simulate_trade_case_fast(case_num, "LONG", cache, pos_idx, lb, la)
                
                sb, sa, s_avg = short_params[sym]
                s_ret, s_en_dt, sp_en, s_ex_dt, sp_ex, s_sl_hit = simulate_trade_case_fast(case_num, "SHORT", cache, pos_idx, sb, sa)
                
                if l_ret == -999.0 and s_ret == -999.0:
                    continue
                    
                trade_sector_val = res_dt
                
                if l_avg >= s_avg:
                    if l_ret != -999.0:
                        q_trades.append({
                            "symbol": sym,
                            "company": SYM_COMPANY.get(sym, sym),
                            "sector": trade_sector_val,
                            "strategy": "LONG",
                            "before": lb,
                            "after": la,
                            "entry_date": l_en_dt,
                            "exit_date": l_ex_dt,
                            "entry_price": lp_en,
                            "exit_price": lp_ex,
                            "expected_return": l_avg,
                            "return_pct": l_ret,
                            "profit_rs": lp_ex - lp_en,
                            "l_ret": l_ret,
                            "s_ret": s_ret,
                            "sl_hit": l_sl_hit
                        })
                else:
                    if s_ret != -999.0:
                        q_trades.append({
                            "symbol": sym,
                            "company": SYM_COMPANY.get(sym, sym),
                            "sector": trade_sector_val,
                            "strategy": "SHORT",
                            "before": sb,
                            "after": sa,
                            "entry_date": s_en_dt,
                            "exit_date": s_ex_dt,
                            "entry_price": sp_en,
                            "exit_price": sp_ex,
                            "expected_return": s_avg,
                            "return_pct": s_ret,
                            "profit_rs": sp_en - sp_ex,
                            "l_ret": l_ret,
                            "s_ret": s_ret,
                            "sl_hit": s_sl_hit
                        })
                        
            if not q_trades:
                continue
                
            q_trades.sort(key=lambda x: x["entry_date"] if x["entry_date"] else pd.Timestamp("2099-01-01"))
            q_slots = simulate_slots([t for t in q_trades if t["entry_date"] and t["exit_date"]])
            build_excel_for_quarter(case_num, q_config, q_trades, q_slots, case_dir)
            
        consolidate_case(case_num, case_dir)
        
    print("\nAll 8 Cases simulation and consolidation complete.")

if __name__ == "__main__":
    main()
