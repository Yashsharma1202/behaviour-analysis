"""
generate_actual_vs_predicted_report.py
===============================================================================
This script fetches the latest stock prices from Yahoo Finance, evaluates Q1 2026
expected (predicted) returns versus actual returns for Nifty 50 stocks, and
generates a detailed report in both Excel and PDF formats.

If the planned exit date has not yet arrived (or has not come yet), the script
uses the latest available price (till date) to calculate the actual return.
===============================================================================
"""
from __future__ import annotations

import datetime as dt
import json
import os
import re
import subprocess
import sys
import time
import urllib.request
import warnings
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import ScatterChart, Reference, Series
import pandas as pd
import numpy as np

warnings.filterwarnings("ignore")
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

# Import sector mappings
try:
    import expectation_bar as EB
    SYM_SECTOR = EB.SYM_SECTOR
except Exception:
    SYM_SECTOR = {}

ROOT = Path(__file__).resolve().parent
SRC_EXCEL = ROOT / "Nifty50_Q1Results_BehaviourEngine_UPDATED_06Aug_fixed.xlsx"
OUT_EXCEL = ROOT / "Nifty50_Actual_vs_Predicted_Report.xlsx"
OUT_HTML = ROOT / "Nifty50_Actual_vs_Predicted_Report.html"
OUT_PDF = ROOT / "Nifty50_Actual_vs_Predicted_Report.pdf"
CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
TODAY = pd.Timestamp(dt.date.today())

def safe_date_str(v, fmt="%Y-%m-%d", default="—"):
    if v is None or pd.isna(v):
        return default
    if isinstance(v, (dt.datetime, dt.date, pd.Timestamp)):
        try:
            return v.strftime(fmt)
        except Exception:
            pass
    return str(v)

def fetch_yahoo_prices(sym: str, start_date: dt.datetime) -> pd.DataFrame | None:
    t = sym.replace("&", "%26") + ".NS"
    p1 = int(start_date.timestamp())
    p2 = int((dt.datetime.now() + pd.Timedelta(days=2)).timestamp())
    url = (f"https://query1.finance.yahoo.com/v8/finance/chart/{t}"
           f"?period1={p1}&period2={p2}&interval=1d")
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=25) as r:
            j = json.loads(r.read())
        res = j["chart"]["result"][0]
        idx = pd.to_datetime(res["timestamp"], unit="s").normalize()
        closes = res["indicators"]["quote"][0]["close"]
        lows = res["indicators"]["quote"][0]["low"]
        df = pd.DataFrame({"close": closes, "low": lows}, index=idx).dropna()
        return df
    except Exception as e:
        print(f"  ! fetch {sym}: {e}")
        return None

def get_price_on_date(s: pd.Series, d) -> float | None:
    ts = pd.Timestamp(d).normalize()
    if ts in s.index:
        return round(float(s.loc[ts]), 2)
    # Find first trading day on or after the date
    after = s[s.index >= ts]
    if len(after):
        return round(float(after.iloc[0]), 2)
    # Fallback to last available trading day before the date
    before = s[s.index < ts]
    if len(before):
        return round(float(before.iloc[-1]), 2)
    return None

def load_source_data(filepath: Path):
    print(f"Loading source data from {filepath.name}...")
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    rows = []
    # Data rows start at index 3
    for r in range(3, ws.max_row + 1):
        sym_val = ws.cell(r, 2).value
        if not sym_val:
            continue
        sym = str(sym_val).strip()
        comp = ws.cell(r, 1).value
        ltp = ws.cell(r, 3).value
        res_date = ws.cell(r, 4).value
        bb = ws.cell(r, 5).value
        sa = ws.cell(r, 6).value
        win = ws.cell(r, 7).value
        exp_ret = ws.cell(r, 9).value
        entry_date = ws.cell(r, 11).value
        buy_price = ws.cell(r, 12).value
        exit_date = ws.cell(r, 13).value
        sell_price = ws.cell(r, 14).value
        
        rows.append({
            "company": comp,
            "symbol": sym,
            "ltp": ltp,
            "result_date": res_date,
            "buy_before": bb,
            "sell_after": sa,
            "win_rate": win,
            "expected_return": exp_ret,
            "entry_date": entry_date,
            "buy_price": buy_price,
            "exit_date": exit_date,
            "sell_price": sell_price,
            "row_index": r
        })
    print(f"Loaded {len(rows)} stocks from source workbook.")
    return rows

def process_results(rows):
    processed = []
    print("Fetching online price data and evaluating returns...")
    start_date = dt.datetime(2026, 1, 1) # fetch prices from January 2026 onwards for EMA 50 warm up
    
    for i, row in enumerate(rows):
        sym = row["symbol"]
        print(f"[{i+1}/{len(rows)}] Processing {sym}...", end="", flush=True)
        
        # Download prices
        df_prices = fetch_yahoo_prices(sym, start_date)
        if df_prices is None or len(df_prices) == 0:
            print(" -> FAILED (No price data)")
            continue
            
        latest_date = df_prices.index[-1]
        entry_date = row["entry_date"]
        exit_date = row["exit_date"]
        
        if pd.isna(entry_date) or entry_date is None:
            print(" -> SKIP (No entry date)")
            continue
            
        entry_dt = pd.Timestamp(entry_date).normalize()
        exit_dt = pd.Timestamp(exit_date).normalize() if not pd.isna(exit_date) else None
        
        # Calculate EMA 50 Low Band
        df_prices["ema50_low"] = df_prices["low"].ewm(span=50, adjust=False).mean()
        
        # If entry date is in the future relative to latest online price data
        if entry_dt > latest_date:
            row["actual_buy_price"] = None
            row["actual_sell_price"] = None
            row["actual_return"] = None
            row["calc_exit_date"] = None
            row["status"] = "Upcoming"
            row["error"] = None
            row["entry_close"] = None
            row["entry_ema50_low"] = None
            print(" -> Upcoming")
            processed.append(row)
            continue
            
        # Get buy price
        buy = row["buy_price"]
        if not isinstance(buy, (int, float)) or pd.isna(buy) or buy <= 0:
            buy = get_price_on_date(df_prices["close"], entry_dt)
            
        if buy is None or buy <= 0:
            print(" -> SKIP (Invalid buy price)")
            continue
            
        # Check EMA 50 Low Band filter
        buy_close = get_price_on_date(df_prices["close"], entry_dt)
        ema_low = get_price_on_date(df_prices["ema50_low"], entry_dt)
        
        row["entry_close"] = buy_close
        row["entry_ema50_low"] = ema_low
        
        if buy_close is not None and ema_low is not None and buy_close < ema_low:
            row["actual_buy_price"] = None
            row["actual_sell_price"] = None
            row["actual_return"] = None
            row["calc_exit_date"] = None
            row["status"] = "Filter Out"
            row["error"] = None
            print(" -> Filter Out (EMA 50)")
            processed.append(row)
            continue
            
        row["actual_buy_price"] = buy
        
        # Determine exit status and price
        # If exit date has already come (is in the past/present relative to latest online data)
        if exit_dt is not None and exit_dt <= latest_date:
            sell = row["sell_price"]
            if not isinstance(sell, (int, float)) or pd.isna(sell) or sell <= 0 or sell == '-':
                sell = get_price_on_date(df_prices["close"], exit_dt)
            if sell is None or sell <= 0:
                sell = df_prices["close"].iloc[-1] # Fallback to latest
            row["actual_sell_price"] = sell
            row["calc_exit_date"] = exit_dt
            row["status"] = "Realised"
            act_ret = round((sell / buy - 1) * 100, 2)
            row["actual_return"] = act_ret
            print(f" -> Realised: buy {buy:.1f}, sell {sell:.1f}, ret {act_ret:+.2f}%")
        else:
            # Exit date has NOT come yet (or is missing) -> calculate till date
            sell = df_prices["close"].iloc[-1]
            row["actual_sell_price"] = sell
            row["calc_exit_date"] = latest_date
            row["status"] = "Open (Till Date)"
            act_ret = round((sell / buy - 1) * 100, 2)
            row["actual_return"] = act_ret
            print(f" -> Open (Till Date): buy {buy:.1f}, latest {sell:.1f}, ret {act_ret:+.2f}%")
            
        # Compute comparison error
        exp_ret = row["expected_return"]
        if exp_ret is not None:
            row["error"] = round(row["actual_return"] - exp_ret, 2)
        else:
            row["error"] = None
            
        processed.append(row)
        time.sleep(0.1) # Be nice to Yahoo Finance
        
    return pd.DataFrame(processed)

def calculate_metrics(df):
    # filter out rows without expected_return or actual_return (like Upcoming or skips)
    valid_df = df[df["expected_return"].notna() & df["actual_return"].notna()].copy()
    if valid_df.empty:
        return {}
        
    y_pred = valid_df["expected_return"].values
    y_true = valid_df["actual_return"].values
    errors = valid_df["error"].values
    
    mae = np.mean(np.abs(errors))
    mse = np.mean(errors ** 2)
    rmse = np.sqrt(mse)
    
    # Pearson Correlation
    if len(valid_df) > 1 and np.std(y_pred) > 0 and np.std(y_true) > 0:
        corr = np.corrcoef(y_pred, y_true)[0, 1]
    else:
        corr = 0.0
        
    # Directional Accuracy (percentage where signs match)
    dir_acc = np.mean((y_pred >= 0) == (y_true >= 0)) * 100
    
    # Realised vs Open split
    realised_sub = valid_df[valid_df["status"] == "Realised"]
    open_sub = valid_df[valid_df["status"] == "Open (Till Date)"]
    
    realised_mae = np.mean(np.abs(realised_sub["error"].values)) if not realised_sub.empty else None
    open_mae = np.mean(np.abs(open_sub["error"].values)) if not open_sub.empty else None
    
    return {
        "total_valid": len(valid_df),
        "total_realised": len(realised_sub),
        "total_open": len(open_sub),
        "avg_predicted": np.mean(y_pred),
        "avg_actual": np.mean(y_true),
        "mae": mae,
        "mse": mse,
        "rmse": rmse,
        "correlation": corr,
        "directional_accuracy": dir_acc,
        "realised_mae": realised_mae,
        "open_mae": open_mae,
        "avg_actual_realised": np.mean(realised_sub["actual_return"].values) if not realised_sub.empty else None,
        "avg_actual_open": np.mean(open_sub["actual_return"].values) if not open_sub.empty else None
    }

def generate_svg_scatter(df):
    valid_df = df[df["expected_return"].notna() & df["actual_return"].notna()].copy()
    if valid_df.empty:
        return "<!-- No data for scatter -->"
        
    x_vals = valid_df["expected_return"].values
    y_vals = valid_df["actual_return"].values
    
    # Range setting
    min_val = min(x_vals.min(), y_vals.min()) - 1
    max_val = max(x_vals.max(), y_vals.max()) + 1
    val_range = max_val - min_val if max_val != min_val else 1
    
    # SVG Dimensions
    w, h = 500, 260
    pad_l, pad_r, pad_t, pad_b = 45, 20, 20, 45
    plot_w = w - pad_l - pad_r
    plot_h = h - pad_t - pad_b
    
    # Helper to convert val to pixel coords
    def get_coords(x_val, y_val):
        cx = pad_l + (x_val - min_val) / val_range * plot_w
        cy = h - pad_b - (y_val - min_val) / val_range * plot_h
        return cx, cy
        
    svg_elements = []
    
    # Diagonal reference line y = x (dashed)
    x0, y0 = get_coords(min_val, min_val)
    x1, y1 = get_coords(max_val, max_val)
    svg_elements.append(f'<line x1="{x0}" y1="{y0}" x2="{x1}" y2="{y1}" stroke="#94a3b8" stroke-width="1.5" stroke-dasharray="4" />')
    
    # Zero axes
    if min_val <= 0 <= max_val:
        zx, _ = get_coords(0, 0)
        _, zy = get_coords(0, 0)
        svg_elements.append(f'<line x1="{zx}" y1="{pad_t}" x2="{zx}" y2="{h - pad_b}" stroke="#cbd5e1" stroke-width="1" />')
        svg_elements.append(f'<line x1="{pad_l}" y1="{zy}" x2="{w - pad_r}" y2="{zy}" stroke="#cbd5e1" stroke-width="1" />')
        
    # Draw points
    for r in valid_df.to_dict('records'):
        color = "#16803d" if r["status"] == "Realised" else "#b45309"
        cx, cy = get_coords(r["expected_return"], r["actual_return"])
        svg_elements.append(
            f'<circle cx="{cx}" cy="{cy}" r="4.5" fill="{color}" opacity="0.85" stroke="#ffffff" stroke-width="0.8">'
            f'<title>{r["symbol"]}\\nExpected: {r["expected_return"]:+.2f}%\\nActual: {r["actual_return"]:+.2f}%\\nStatus: {r["status"]}</title>'
            f'</circle>'
        )
        
    # Grid ticks (5 steps)
    for i in range(5):
        val = min_val + val_range * i / 4
        cx, cy = get_coords(val, val)
        # horizontal label (X axis)
        svg_elements.append(f'<text x="{cx}" y="{h - pad_b + 12}" font-size="7" fill="#64748b" text-anchor="middle">{val:.1f}%</text>')
        # vertical label (Y axis)
        svg_elements.append(f'<text x="{pad_l - 6}" y="{cy + 2.5}" font-size="7" fill="#64748b" text-anchor="end">{val:.1f}%</text>')
        
    # Labels
    svg_elements.append(f'<text x="{pad_l + plot_w/2}" y="{h - 10}" font-size="9" font-weight="bold" fill="#334155" text-anchor="middle">Expected Return (%)</text>')
    svg_elements.append(f'<text x="12" y="{pad_t + plot_h/2}" font-size="9" font-weight="bold" fill="#334155" text-anchor="middle" transform="rotate(-90 12 {pad_t + plot_h/2})">Actual Return (%)</text>')
    
    # Title & Legend inside SVG
    svg_elements.append(f'<text x="{pad_l + 10}" y="{pad_t + 15}" font-size="9" font-weight="bold" fill="#0f172a">Actual vs Expected Scatter Plot</text>')
    
    # Legend
    svg_elements.append(f'<circle cx="{w - pad_r - 95}" cy="{pad_t + 10}" r="4" fill="#16803d" />')
    svg_elements.append(f'<text x="{w - pad_r - 87}" y="{pad_t + 13}" font-size="7" fill="#334155">Realised</text>')
    svg_elements.append(f'<circle cx="{w - pad_r - 50}" cy="{pad_t + 10}" r="4" fill="#b45309" />')
    svg_elements.append(f'<text x="{w - pad_r - 42}" y="{pad_t + 13}" font-size="7" fill="#334155">Open</text>')
    
    svg_body = "\n".join(svg_elements)
    return f"""
    <svg width="100%" height="100%" viewBox="0 0 {w} {h}" style="background-color: #ffffff; display: block;">
        <rect x="0" y="0" width="{w}" height="{h}" fill="none" stroke="#e2e8f0" stroke-width="1" rx="4" />
        {svg_body}
    </svg>
    """

def generate_svg_bar(df):
    valid_df = df[df["expected_return"].notna() & df["actual_return"].notna()].copy()
    if valid_df.empty:
        return "<!-- No data for bar -->"
        
    # Top 15 movers by absolute actual return
    top_15 = valid_df.assign(abs_act=valid_df["actual_return"].abs()).sort_values("abs_act", ascending=False).head(15).to_dict('records')
    
    w, h = 500, 260
    pad_l, pad_r, pad_t, pad_b = 35, 15, 20, 35
    plot_w = w - pad_l - pad_r
    plot_h = h - pad_t - pad_b
    
    # Y Scale
    max_y = max(max(abs(r["expected_return"]), abs(r["actual_return"])) for r in top_15) + 1
    
    def get_y(val):
        zero_y = pad_t + plot_h / 2
        pixel_per_val = (plot_h / 2) / max_y
        return zero_y - val * pixel_per_val
        
    zero_y = get_y(0)
    svg_elements = []
    
    # Grid ticks (Y axis)
    for tick_val in [-max_y*0.8, -max_y*0.4, 0, max_y*0.4, max_y*0.8]:
        ty = get_y(tick_val)
        svg_elements.append(f'<line x1="{pad_l}" y1="{ty}" x2="{w - pad_r}" y2="{ty}" stroke="#f1f5f9" stroke-width="1" />')
        svg_elements.append(f'<text x="{pad_l - 6}" y="{ty + 2.5}" font-size="7" fill="#64748b" text-anchor="end">{tick_val:.1f}%</text>')
        
    # Draw zero line
    svg_elements.append(f'<line x1="{pad_l}" y1="{zero_y}" x2="{w - pad_r}" y2="{zero_y}" stroke="#94a3b8" stroke-width="1" />')
    
    # Bar sizes
    n = len(top_15)
    slot_w = plot_w / n
    bar_w = slot_w * 0.32
    
    for idx, r in enumerate(top_15):
        cx = pad_l + idx * slot_w + slot_w / 2
        
        # Expected bar (blue)
        ey = get_y(r["expected_return"])
        eh = abs(zero_y - ey)
        ey_rect = ey if r["expected_return"] >= 0 else zero_y
        svg_elements.append(f'<rect x="{cx - bar_w}" y="{ey_rect}" width="{bar_w}" height="{eh}" fill="#3b82f6" rx="0.5" />')
        
        # Actual bar (green if positive, red if negative)
        ay = get_y(r["actual_return"])
        ah = abs(zero_y - ay)
        ay_rect = ay if r["actual_return"] >= 0 else zero_y
        acolor = "#10b981" if r["actual_return"] >= 0 else "#ef4444"
        svg_elements.append(f'<rect x="{cx}" y="{ay_rect}" width="{bar_w}" height="{ah}" fill="{acolor}" rx="0.5" />')
        
        # Symbol label
        svg_elements.append(f'<text x="{cx}" y="{h - pad_b + 11}" font-size="7" font-weight="bold" fill="#334155" text-anchor="middle">{r["symbol"]}</text>')
        
    # Axis Title
    svg_elements.append(f'<text x="10" y="{pad_t + plot_h/2}" font-size="9" font-weight="bold" fill="#334155" text-anchor="middle" transform="rotate(-90 10 {pad_t + plot_h/2})">Return %</text>')
    svg_elements.append(f'<text x="{pad_l + 10}" y="{pad_t + 15}" font-size="9" font-weight="bold" fill="#0f172a">Top 15 Absolute Movers (Expected vs Actual)</text>')
    
    # Legends
    svg_elements.append(f'<rect x="{w - pad_r - 180}" y="{pad_t}" width="8" height="8" fill="#3b82f6" rx="1" />')
    svg_elements.append(f'<text x="{w - pad_r - 168}" y="{pad_t + 7}" font-size="7" fill="#334155">Expected</text>')
    svg_elements.append(f'<rect x="{w - pad_r - 120}" y="{pad_t}" width="8" height="8" fill="#10b981" rx="1" />')
    svg_elements.append(f'<text x="{w - pad_r - 108}" y="{pad_t + 7}" font-size="7" fill="#334155">Actual (+)</text>')
    svg_elements.append(f'<rect x="{w - pad_r - 70}" y="{pad_t}" width="8" height="8" fill="#ef4444" rx="1" />')
    svg_elements.append(f'<text x="{w - pad_r - 58}" y="{pad_t + 7}" font-size="7" fill="#334155">Actual (-)</text>')
    
    svg_body = "\n".join(svg_elements)
    return f"""
    <svg width="100%" height="100%" viewBox="0 0 {w} {h}" style="background-color: #ffffff; display: block;">
        <rect x="0" y="0" width="{w}" height="{h}" fill="none" stroke="#e2e8f0" stroke-width="1" rx="4" />
        {svg_body}
    </svg>
    """

def write_excel_report(df, stats):
    print(f"Generating Excel report {OUT_EXCEL.name}...")
    wb = openpyxl.Workbook()
    
    # 1. SUMMARY SHEET
    ws_sum = wb.active
    ws_sum.title = "Summary"
    ws_sum.views.sheetView[0].showGridLines = True
    
    # Styles
    title_font = Font(name="Segoe UI", size=16, bold=True, color="1F4E79")
    hdr_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    section_font = Font(name="Segoe UI", size=11, bold=True, color="1F4E79")
    bold_font = Font(name="Segoe UI", size=10, bold=True)
    normal_font = Font(name="Segoe UI", size=10)
    ital_font = Font(name="Segoe UI", size=9, italic=True, color="555555")
    
    navy_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    soft_fill = PatternFill(start_color="F2F4F7", end_color="F2F4F7", fill_type="solid")
    
    thin_side = Side(style='thin', color='D9D9D9')
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    
    # Header block
    ws_sum["A1"] = "Actual vs Predicted Return Performance Report"
    ws_sum["A1"].font = title_font
    ws_sum.merge_cells("A1:D1")
    
    ws_sum["A2"] = f"Generated: {dt.date.today().strftime('%d-%b-%Y')}  |  Nifty 50 Q1 Results Tracker"
    ws_sum["A2"].font = ital_font
    ws_sum.merge_cells("A2:D2")
    
    # Metrics Table
    ws_sum["A4"] = "Overall Performance Statistics"
    ws_sum["A4"].font = section_font
    
    headers = ["Metric", "Value", "Notes"]
    for c, h in enumerate(headers, 1):
        cell = ws_sum.cell(row=5, column=c, value=h)
        cell.font = hdr_font
        cell.fill = navy_fill
        cell.alignment = align_center
        cell.border = border_all
        
    metrics = [
        ("Total Analyzed Positions", stats["total_valid"], "Stocks with valid predictions and entry dates"),
        ("Realised Positions (Exit date reached)", stats["total_realised"], "Positions closed based on holding period"),
        ("Open Positions (Running till-date)", stats["total_open"], "Positions still open, calculated with current price"),
        ("Mean Absolute Error (MAE)", f"{stats['mae']:.2f}%", "Average magnitude of prediction error"),
        ("Root Mean Squared Error (RMSE)", f"{stats['rmse']:.2f}%", "Penalizes larger errors more heavily"),
        ("Directional Accuracy", f"{stats['directional_accuracy']:.1f}%", "How often predicted sign matches actual sign"),
        ("Correlation (Predicted vs Actual)", f"{stats['correlation']:.3f}", "Pearson correlation coefficient (-1 to +1)"),
        ("Average Predicted Return", f"{stats['avg_predicted']:.2f}%", "Mean return predicted by the model"),
        ("Average Actual Return", f"{stats['avg_actual']:.2f}%", "Mean return actually generated (realised + open)"),
        ("Avg Return (Realised Only)", f"{stats['avg_actual_realised']:.2f}%" if stats['avg_actual_realised'] is not None else "N/A", "Mean return of completed trades"),
        ("Avg Return (Open Till-Date)", f"{stats['avg_actual_open']:.2f}%" if stats['avg_actual_open'] is not None else "N/A", "Mean return of open running positions")
    ]
    
    for r_idx, (m, v, n) in enumerate(metrics, 6):
        ws_sum.cell(row=r_idx, column=1, value=m).font = bold_font
        ws_sum.cell(row=r_idx, column=1).border = border_all
        
        v_cell = ws_sum.cell(row=r_idx, column=2, value=v)
        v_cell.font = normal_font
        v_cell.alignment = align_right
        v_cell.border = border_all
        
        ws_sum.cell(row=r_idx, column=3, value=n).font = normal_font
        ws_sum.cell(row=r_idx, column=3).border = border_all
        
    # Auto-adjust column width for Summary sheet
    for col in ws_sum.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws_sum.column_dimensions[col_letter].width = max(max_len + 3, 10)
        
    # 2. DETAIL SHEET
    ws_det = wb.create_sheet(title="Details")
    ws_det.views.sheetView[0].showGridLines = True
    
    det_headers = [
        "Company Name", "SYMBOL", "Sector", "Result Date", "Before Candle (Buy)", "After Candle (Sell)",
        "Predicted Return (%)", "Actual Return (%)", "Status", "Entry Date", "Buy Price", 
        "Planned Exit Date", "Calc Exit Date (Used)", "Sell Price (Latest)", "LTP"
    ]
    
    for c, h in enumerate(det_headers, 1):
        cell = ws_det.cell(row=1, column=c, value=h)
        cell.font = hdr_font
        cell.fill = navy_fill
        cell.alignment = align_center
        cell.border = border_all
        
    # Fill Data
    # Sort by symbol
    df_sorted = df.sort_values("symbol")
    
    green_text = Font(name="Segoe UI", size=10, color="FF006100")
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red_text = Font(name="Segoe UI", size=10, color="FF9C0006")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    yellow_text = Font(name="Segoe UI", size=10, color="FF9C6500")
    yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    
    for r_idx, row in enumerate(df_sorted.to_dict('records'), 2):
        bg = soft_fill if r_idx % 2 == 0 else PatternFill(fill_type=None)
        
        # company
        c = ws_det.cell(row=r_idx, column=1, value=row["company"])
        c.font = normal_font; c.border = border_all; c.fill = bg
        
        # symbol
        c = ws_det.cell(row=r_idx, column=2, value=row["symbol"])
        c.font = bold_font; c.border = border_all; c.alignment = align_center; c.fill = bg
        
        # sector
        sect = SYM_SECTOR.get(row["symbol"], "N/A")
        c = ws_det.cell(row=r_idx, column=3, value=sect)
        c.font = normal_font; c.border = border_all; c.fill = bg
        
        # result_date
        res_date = safe_date_str(row["result_date"], '%Y-%m-%d')
        c = ws_det.cell(row=r_idx, column=4, value=res_date)
        c.font = normal_font; c.border = border_all; c.alignment = align_center; c.fill = bg
        
        # predicted
        pred = row["expected_return"]
        c = ws_det.cell(row=r_idx, column=5, value=pred)
        c.font = normal_font; c.border = border_all; c.alignment = align_right; c.fill = bg
        if isinstance(pred, (int, float)):
            c.number_format = "0.00"
            
        # actual
        act = row["actual_return"]
        c = ws_det.cell(row=r_idx, column=6, value=act)
        c.font = normal_font; c.border = border_all; c.alignment = align_right; c.fill = bg
        if isinstance(act, (int, float)):
            c.number_format = "0.00"
            if act >= 0:
                c.font = green_text; c.fill = green_fill
            else:
                c.font = red_text; c.fill = red_fill
                
        # buy_before
        c = ws_det.cell(row=r_idx, column=5, value=row["buy_before"])
        c.font = normal_font; c.border = border_all; c.alignment = align_center; c.fill = bg
        
        # sell_after
        c = ws_det.cell(row=r_idx, column=6, value=row["sell_after"])
        c.font = normal_font; c.border = border_all; c.alignment = align_center; c.fill = bg
        
        # predicted
        pred = row["expected_return"]
        c = ws_det.cell(row=r_idx, column=7, value=pred)
        c.font = normal_font; c.border = border_all; c.alignment = align_right; c.fill = bg
        if isinstance(pred, (int, float)):
            c.number_format = "0.00"
            
        # actual
        act = row["actual_return"]
        c = ws_det.cell(row=r_idx, column=8, value=act)
        c.font = normal_font; c.border = border_all; c.alignment = align_right; c.fill = bg
        if isinstance(act, (int, float)):
            c.number_format = "0.00"
            if act >= 0:
                c.font = green_text; c.fill = green_fill
            else:
                c.font = red_text; c.fill = red_fill
                
        # status
        stat = row["status"]
        c = ws_det.cell(row=r_idx, column=9, value=stat)
        c.font = normal_font; c.border = border_all; c.alignment = align_center; c.fill = bg
        if stat == "Realised":
            c.fill = green_fill; c.font = green_text
        elif "Open" in stat:
            c.fill = yellow_fill; c.font = yellow_text
            
        # entry_date
        ent_date = safe_date_str(row["entry_date"], '%Y-%m-%d')
        c = ws_det.cell(row=r_idx, column=10, value=ent_date)
        c.font = normal_font; c.border = border_all; c.alignment = align_center; c.fill = bg
        
        # buy_price
        c = ws_det.cell(row=r_idx, column=11, value=row["actual_buy_price"])
        c.font = normal_font; c.border = border_all; c.alignment = align_right; c.fill = bg
        if isinstance(row["actual_buy_price"], (int, float)):
            c.number_format = "#,##0.00"
            
        # planned_exit_date
        pl_ex_date = safe_date_str(row["exit_date"], '%Y-%m-%d')
        c = ws_det.cell(row=r_idx, column=12, value=pl_ex_date)
        c.font = normal_font; c.border = border_all; c.alignment = align_center; c.fill = bg
        
        # calc_exit_date
        calc_ex_date = safe_date_str(row["calc_exit_date"], '%Y-%m-%d')
        c = ws_det.cell(row=r_idx, column=13, value=calc_ex_date)
        c.font = normal_font; c.border = border_all; c.alignment = align_center; c.fill = bg
        if stat == "Open (Till Date)":
            c.font = ital_font
            
        # sell_price (latest)
        c = ws_det.cell(row=r_idx, column=14, value=row["actual_sell_price"])
        c.font = normal_font; c.border = border_all; c.alignment = align_right; c.fill = bg
        if isinstance(row["actual_sell_price"], (int, float)):
            c.number_format = "#,##0.00"
            
        # LTP (latest price in excel)
        c = ws_det.cell(row=r_idx, column=15, value=row["ltp"])
        c.font = normal_font; c.border = border_all; c.alignment = align_right; c.fill = bg
        if isinstance(row["ltp"], (int, float)):
            c.number_format = "#,##0.00"
            
    # Auto-fit detailed sheet columns
    for col in ws_det.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws_det.column_dimensions[col_letter].width = max(max_len + 3, 10)
        
    # Add native openpyxl Scatter Chart to Summary Sheet
    try:
        last_row = len(df) + 1
        sc = ScatterChart()
        sc.title = "Expected vs Actual Returns Scatter"
        sc.x_axis.title = "Expected Return %"
        sc.y_axis.title = "Actual Return %"
        sc.legend = None
        sc.width = 16
        sc.height = 11
        
        # col 7 = Expected Return, col 8 = Actual Return in Details sheet
        xvalues = Reference(ws_det, min_col=7, min_row=2, max_row=last_row)
        yvalues = Reference(ws_det, min_col=8, min_row=2, max_row=last_row)
        ser = Series(yvalues, xvalues, title_from_data=False)
        ser.marker.symbol = "circle"
        ser.marker.size = 5
        ser.graphicalProperties.line.noFill = True
        sc.series.append(ser)
        
        ws_sum.add_chart(sc, "E4")
        print("  Added native openpyxl scatter chart to Summary sheet.")
    except Exception as ec:
        print(f"  ! Could not add native Excel chart: {ec}")
        
    out = OUT_EXCEL
    try:
        wb.save(out)
    except PermissionError:
        out = OUT_EXCEL.with_name(OUT_EXCEL.stem + "_2.xlsx")
        wb.save(out)
        print(f"  (original open in Excel — saved to {out.name})")
    print(f"Excel report saved successfully to {out.name}")

def write_html_pdf_report(df, stats):
    print("Generating HTML report template with inline SVG charts...")
    
    # Sort data for HTML display
    df_sorted = df.sort_values("symbol")
    
    # Create HTML row templates
    rows_html = []
    for r in df_sorted.to_dict('records'):
        sym = r["symbol"]
        sect = SYM_SECTOR.get(sym, "N/A")
        res_date = safe_date_str(r["result_date"], '%d-%b')
        ent_date = safe_date_str(r["entry_date"], '%d-%b')
        ex_date = safe_date_str(r["calc_exit_date"], '%d-%b')
        
        pred = r["expected_return"]
        pred_str = f"{pred:+.2f}%" if isinstance(pred, (int, float)) else "—"
        pred_cls = "pos" if isinstance(pred, (int, float)) and pred >= 0 else "neg" if isinstance(pred, (int, float)) else "dash"
        
        act = r["actual_return"]
        act_str = f"{act:+.2f}%" if isinstance(act, (int, float)) else "—"
        act_cls = "pos" if isinstance(act, (int, float)) and act >= 0 else "neg" if isinstance(act, (int, float)) else "dash"
        
        bb = r["buy_before"] if r["buy_before"] is not None else "—"
        sa = r["sell_after"] if r["sell_after"] is not None else "—"
        
        status = r["status"]
        status_cls = "realised" if status == "Realised" else "open" if "Open" in status else "filtered" if status == "Filter Out" else "upcoming"
        
        buy = f"{r['actual_buy_price']:,.1f}" if isinstance(r['actual_buy_price'], (int, float)) else "—"
        sell = f"{r['actual_sell_price']:,.1f}" if isinstance(r['actual_sell_price'], (int, float)) else "—"
        
        row_str = f"""
        <tr class="{status_cls}-row">
            <td class="name">{r['company']}</td>
            <td class="sym">{sym}</td>
            <td class="c">{sect}</td>
            <td class="c">{res_date}</td>
            <td class="c">{bb}</td>
            <td class="c">{sa}</td>
            <td class="n {pred_cls}">{pred_str}</td>
            <td class="n {act_cls}">{act_str}</td>
            <td class="c"><span class="badge {status_cls}">{status}</span></td>
            <td class="c">{ent_date}</td>
            <td class="n">{buy}</td>
            <td class="c">{"<i>" + ex_date + "</i>" if status_cls == "open" else ex_date}</td>
            <td class="n">{sell}</td>
        </tr>
        """
        rows_html.append(row_str)
        
    rows_body = "\n".join(rows_html)
    
    # Generate SVG code
    svg_scatter = generate_svg_scatter(df)
    svg_bar = generate_svg_bar(df)
    
    # HTML template with CSS print rules
    html_content = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Actual vs Predicted Performance Report</title>
<style>
@page {{
    size: A4 landscape;
    margin: 8mm;
}}
* {{
    box-sizing: border-box;
}}
body {{
    margin: 0;
    font-family: 'Segoe UI', Arial, sans-serif;
    color: #1e293b;
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
    background-color: #ffffff;
}}
header {{
    background-color: #1e3a8a;
    color: #ffffff;
    padding: 12px 18px;
    border-radius: 6px;
    margin-bottom: 12px;
}}
header h1 {{
    margin: 0;
    font-size: 18px;
    font-weight: 700;
    letter-spacing: 0.5px;
}}
header .subtitle {{
    margin: 2px 0 0 0;
    font-size: 10px;
    color: #93c5fd;
    font-weight: 500;
}}
.grid {{
    display: flex;
    gap: 12px;
    margin-bottom: 12px;
}}
.card {{
    flex: 1;
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    border-radius: 6px;
    padding: 8px 12px;
}}
.card-title {{
    font-size: 8.5px;
    text-transform: uppercase;
    color: #64748b;
    font-weight: 700;
    margin-bottom: 2px;
}}
.card-value {{
    font-size: 15px;
    font-weight: 700;
    color: #0f172a;
}}
.card-desc {{
    font-size: 8px;
    color: #64748b;
    margin-top: 1px;
}}
.plots-container {{
    display: flex;
    gap: 12px;
    margin-bottom: 15px;
    height: 270px;
    page-break-after: always;
}}
.plot-box {{
    flex: 1;
    border: 1px solid #e2e8f0;
    border-radius: 6px;
    padding: 8px;
    background-color: #ffffff;
    display: flex;
    flex-direction: column;
}}
table {{
    width: 100%;
    border-collapse: collapse;
    margin-top: 10px;
}}
th, td {{
    border: 1px solid #cbd5e1;
    padding: 3px 5px;
    font-size: 8.2px;
    overflow: hidden;
    white-space: nowrap;
    text-overflow: ellipsis;
}}
thead tr {{
    background-color: #f1f5f9;
}}
thead th {{
    font-weight: 700;
    color: #334155;
    text-align: center;
    padding: 5px 3px;
}}
tr.realised-row {{
    background-color: #ffffff;
}}
tr.open-row {{
    background-color: #fffbeb;
}}
tr.upcoming-row {{
    background-color: #f8fafc;
    color: #94a3b8;
}}
tr.filtered-row {{
    background-color: #f8fafc;
    color: #94a3b8;
    font-style: italic;
}}
td.name {{
    text-align: left;
    font-weight: 600;
}}
td.sym {{
    text-align: center;
    font-weight: 700;
    color: #1e3a8a;
}}
td.c {{
    text-align: center;
}}
td.n {{
    text-align: right;
}}
td.pos {{
    color: #15803d;
    font-weight: 600;
}}
td.neg {{
    color: #b91c1c;
    font-weight: 600;
}}
td.dash {{
    color: #94a3b8;
    text-align: center;
}}
.badge {{
    display: inline-block;
    padding: 1px 4px;
    border-radius: 3px;
    font-size: 7.5px;
    font-weight: 700;
    text-transform: uppercase;
}}
.badge.realised {{
    background-color: #dcfce7;
    color: #15803d;
}}
.badge.open {{
    background-color: #fef3c7;
    color: #b45309;
}}
.badge.upcoming {{
    background-color: #f1f5f9;
    color: #64748b;
}}
.badge.filtered {{
    background-color: #ffe4e6;
    color: #b91c1c;
}}
.footer {{
    margin-top: 12px;
    font-size: 8px;
    color: #64748b;
    display: flex;
    justify-content: space-between;
}}
</style>
</head>
<body>

    <header>
        <h1>Nifty 50 Q1 Results: Actual vs Predicted Performance Report</h1>
        <div class="subtitle">Comparing expected model predictions with actual outcomes (realised & open till-date positions)</div>
    </header>

    <div class="grid">
        <div class="card">
            <div class="card-title">Total Active Stocks</div>
            <div class="card-value">{stats.get('total_valid', 0)}</div>
            <div class="card-desc">With valid entry dates & predictions</div>
        </div>
        <div class="card">
            <div class="card-title">Realised / Open Split</div>
            <div class="card-value">{stats.get('total_realised', 0)} R / {stats.get('total_open', 0)} O</div>
            <div class="card-desc">R = Exited, O = Running till-date</div>
        </div>
        <div class="card">
            <div class="card-title">Mean Absolute Error</div>
            <div class="card-value">{stats.get('mae', 0.0):.2f}%</div>
            <div class="card-desc">MAE (Realised: {stats.get('realised_mae', 0.0) or 0.0:.2f}% | Open: {stats.get('open_mae', 0.0) or 0.0:.2f}%)</div>
        </div>
        <div class="card">
            <div class="card-title">Directional Accuracy</div>
            <div class="card-value">{stats.get('directional_accuracy', 0.0):.1f}%</div>
            <div class="card-desc">Sign match of Predicted vs Actual</div>
        </div>
        <div class="card">
            <div class="card-title">Correlation</div>
            <div class="card-value">{stats.get('correlation', 0.0):.3f}</div>
            <div class="card-desc">Pearson correlation coefficient</div>
        </div>
        <div class="card">
            <div class="card-title">Avg Pred vs Actual</div>
            <div class="card-value">{stats.get('avg_predicted', 0.0):.2f}% / {stats.get('avg_actual', 0.0):.2f}%</div>
            <div class="card-desc">Expected return vs Actual return</div>
        </div>
    </div>

    <div class="plots-container">
        <div class="plot-box">
            {svg_scatter}
        </div>
        <div class="plot-box">
            {svg_bar}
        </div>
    </div>

    <!-- PAGE 2: DETAILED TABLE -->
    <header style="margin-top: 20px;">
        <h1>Detailed Performance Metrics Table</h1>
        <div class="subtitle">Per-stock actual returns vs expectations based on Q1 reporting window</div>
    </header>

    <table>
        <thead>
            <tr>
                <th style="width: 15%;">Company Name</th>
                <th style="width: 7%;">SYMBOL</th>
                <th style="width: 8%;">Sector</th>
                <th style="width: 8%;">Result Date</th>
                <th style="width: 7%;">Before (Buy)</th>
                <th style="width: 7%;">After (Sell)</th>
                <th style="width: 8%;">Expected Return</th>
                <th style="width: 8%;">Actual Return</th>
                <th style="width: 8%;">Status</th>
                <th style="width: 8%;">Entry Date</th>
                <th style="width: 8%;">Buy Price</th>
                <th style="width: 8%;">Exit Date</th>
                <th style="width: 8%;">Exit Price</th>
            </tr>
        </thead>
        <tbody>
            {rows_body}
        </tbody>
    </table>

    <div class="footer">
        <div>Nifty 50 Behaviour Engine · Q1 Performance Audit</div>
        <div>Generated on {dt.date.today().strftime('%d-%b-%Y')} · Page 2 of 2</div>
    </div>

</body>
</html>
"""
    
    OUT_HTML.write_text(html_content, encoding="utf-8")
    print(f"HTML report template written to {OUT_HTML.name}")
    
    # Run Chrome in Headless Mode to print HTML to PDF
    print(f"Compiling PDF report {OUT_PDF.name} using Google Chrome...")
    try:
        r = subprocess.run([
            CHROME, 
            "--headless", 
            "--disable-gpu", 
            "--no-pdf-header-footer", 
            f"--print-to-pdf={OUT_PDF}", 
            OUT_HTML.as_uri()
        ], capture_output=True, text=True, timeout=30)
        
        if OUT_PDF.exists() and OUT_PDF.stat().st_size > 0:
            print(f"PDF report successfully compiled: {OUT_PDF.name} ({OUT_PDF.stat().st_size:,} bytes)")
        else:
            print(f"  ! PDF compile FAILED. Chrome output: {r.stderr}")
    except Exception as e:
        print(f"  ! Error compiling PDF: {e}")

def write_ema_filter_excel(df):
    print("Generating EMA 50 Low Band Filter Details Excel...")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "EMA 50 Filter Details"
    ws.views.sheetView[0].showGridLines = True
    
    # Styles
    title_font = Font(name="Segoe UI", size=16, bold=True, color="1F4E79")
    hdr_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    bold_font = Font(name="Segoe UI", size=10, bold=True)
    normal_font = Font(name="Segoe UI", size=10)
    
    navy_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    soft_fill = PatternFill(start_color="F2F4F7", end_color="F2F4F7", fill_type="solid")
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    green_text = Font(name="Segoe UI", size=10, color="FF006100")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    red_text = Font(name="Segoe UI", size=10, color="FF9C0006")
    
    thin_side = Side(style='thin', color='D9D9D9')
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    
    align_center = Alignment(horizontal="center", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    
    # Title
    ws.cell(row=1, column=1, value="Nifty 50 Q1 FY27 - EMA 50 Low Band Filter Decision").font = title_font
    ws.row_dimensions[1].height = 30
    
    # Headers
    headers = [
        "Company Name", "SYMBOL", "Result Date", "Entry Date", 
        "Entry Close Price", "EMA 50 Low Band", 
        "Predicted Return (%)", "Actual Return (%)", "Status / Decision"
    ]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=c, value=h)
        cell.font = hdr_font
        cell.fill = navy_fill
        cell.alignment = align_center
        cell.border = border_all
    ws.row_dimensions[3].height = 24
    
    # Fill rows
    df_sorted = df.sort_values("symbol")
    for r_idx, row in enumerate(df_sorted.to_dict('records'), 4):
        bg = soft_fill if r_idx % 2 == 0 else PatternFill(fill_type=None)
        
        # company
        c = ws.cell(row=r_idx, column=1, value=row["company"])
        c.font = normal_font; c.border = border_all; c.fill = bg
        
        # symbol
        c = ws.cell(row=r_idx, column=2, value=row["symbol"])
        c.font = bold_font; c.border = border_all; c.alignment = align_center; c.fill = bg
        
        # result_date
        res_date = safe_date_str(row["result_date"], '%Y-%m-%d')
        c = ws.cell(row=r_idx, column=3, value=res_date)
        c.font = normal_font; c.border = border_all; c.alignment = align_center; c.fill = bg
        
        # entry_date
        ent_date = safe_date_str(row["entry_date"], '%Y-%m-%d')
        c = ws.cell(row=r_idx, column=4, value=ent_date)
        c.font = normal_font; c.border = border_all; c.alignment = align_center; c.fill = bg
        
        # entry_close
        c = ws.cell(row=r_idx, column=5, value=row.get("entry_close"))
        c.font = normal_font; c.border = border_all; c.alignment = align_right; c.fill = bg
        if isinstance(c.value, (int, float)):
            c.number_format = "#,##0.00"
            
        # ema 50 low
        c = ws.cell(row=r_idx, column=6, value=row.get("entry_ema50_low"))
        c.font = normal_font; c.border = border_all; c.alignment = align_right; c.fill = bg
        if isinstance(c.value, (int, float)):
            c.number_format = "#,##0.00"
            
        # predicted return
        c = ws.cell(row=r_idx, column=7, value=row.get("expected_return"))
        c.font = normal_font; c.border = border_all; c.alignment = align_right; c.fill = bg
        if isinstance(c.value, (int, float)):
            c.number_format = "0.00"
            
        # actual return
        c = ws.cell(row=r_idx, column=8, value=row.get("actual_return"))
        c.font = normal_font; c.border = border_all; c.alignment = align_right; c.fill = bg
        if isinstance(c.value, (int, float)):
            c.number_format = "0.00"
            if c.value >= 0:
                c.font = green_text; c.fill = green_fill
            else:
                c.font = red_text; c.fill = red_fill
        elif c.value is None and row["status"] == "Filter Out":
            c.value = "Filter Out"
            c.font = Font(name="Segoe UI", size=10, italic=True, color="777777"); c.alignment = align_center
            
        # Decision
        is_filtered = row["status"] == "Filter Out"
        decision_str = "Filter Out (No Buy)" if is_filtered else "Approved (Buy)"
        c = ws.cell(row=r_idx, column=9, value=decision_str)
        c.border = border_all; c.alignment = align_center; c.fill = bg
        if is_filtered:
            c.fill = red_fill; c.font = red_text
        else:
            c.fill = green_fill; c.font = green_text
            
    # Auto-fit columns
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 3, 10)
        
    out_name = "Nifty50_EMA_50_Filter_Details.xlsx"
    try:
        wb.save(out_name)
    except PermissionError:
        out_name = "Nifty50_EMA_50_Filter_Details_2.xlsx"
        wb.save(out_name)
        print(f"  (original open in Excel — saved to {out_name})")
    print(f"EMA 50 Filter Details Excel saved successfully to {out_name}")

def main():
    print("=" * 60)
    print("Nifty 50 Actual vs Predicted Performance Report Tool")
    print("=" * 60)
    
    if not SRC_EXCEL.exists():
        print(f"Error: Source workbook {SRC_EXCEL.name} not found!")
        sys.exit(1)
        
    rows = load_source_data(SRC_EXCEL)
    df = process_results(rows)
    
    # Save processed dataframe to csv for debugging/backup
    df.to_csv(ROOT / "actual_vs_predicted_processed.csv", index=False)
    
    stats = calculate_metrics(df)
    
    print("\n" + "-" * 40)
    print("SUMMARY STATISTICS")
    print("-" * 40)
    for k, v in stats.items():
        if isinstance(v, float):
            print(f"  {k:<30}: {v:+.3f}")
        else:
            print(f"  {k:<30}: {v}")
    print("-" * 40)
    
    write_excel_report(df, stats)
    write_html_pdf_report(df, stats)
    write_ema_filter_excel(df)
    print("\nAll done!")

if __name__ == "__main__":
    main()
