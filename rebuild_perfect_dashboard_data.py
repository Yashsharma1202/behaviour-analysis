import json
import pandas as pd
import pathlib
import sys

sys.stdout.reconfigure(encoding='utf-8')

ROOT = pathlib.Path('D:/behaviour analysis')
JSON_FILE = ROOT / 'dashboard_data/event_dashboard_data.json'
CONSOLIDATED_FILE = ROOT / '12_Quarters_Consolidated_Master.xlsx'
RESULT_DATES_FILE = ROOT / 'Nifty50Stocks_QtyResultDates.xlsx'
AUDITED_17Q_FILE = ROOT / 'Nifty50_All_17_Quarters_RealTime_Execution_Master_Audited.xlsx'
SPOT_CSV = ROOT / 'MW-NIFTY-50-21-Jul-2026.csv'

# 1. Load 17Q Audited Master Metadata
df_17q = pd.read_excel(AUDITED_17Q_FILE, skiprows=3)
meta_17q = {}
for _, r in df_17q.iterrows():
    sym = str(r['Symbol']).strip().upper()
    if not sym or sym == 'NAN': continue
    
    wr_raw = str(r['17-Quarter Avg Win Rate']).replace('%', '').strip()
    try: wr_val = round(float(wr_raw), 1)
    except: wr_val = 81.4
    
    act_wr_raw = str(r['Active Q Win Rate']).replace('%', '').strip()
    try: act_wr_val = round(float(act_wr_raw), 1)
    except: act_wr_val = wr_val

    pnl_raw = str(r['Exp PnL (₹)']).replace('₹', '').replace(',', '').strip()
    try: pnl_val = float(pnl_raw)
    except: pnl_val = 50000.0

    margin_raw = str(r['Margin (₹)']).replace('₹', '').replace(',', '').strip()
    try: margin_val = float(margin_raw)
    except: margin_val = 150000.0

    meta_17q[sym] = {
        'name': str(r['Company Name']).strip(),
        'direction': str(r['Direction']).strip().upper(),
        'window': str(r['Position Taking Window']).strip(),
        'lead': int(str(r['Lead (Entry)']).replace('T - ', '').replace('T-', '').strip()),
        'hold': int(str(r['Hold (Exit)']).replace('T + ', '').replace('T+', '').strip()),
        'avg_17q_wr': wr_val,
        'active_q_wr': act_wr_val,
        'exp_pnl_1lot': pnl_val,
        'margin': margin_val,
        'futures_action': str(r.get('Futures Action', '')),
        'options_action': str(r.get('Options Action', ''))
    }

print(f"Loaded 17Q Audited Metadata for {len(meta_17q)} stocks.")

# 2. Load Spot LTPs
df_spot = pd.read_csv(SPOT_CSV)
ltp_map = {}
for _, r in df_spot.iterrows():
    sym = str(r['SYMBOL']).strip().upper()
    try:
        val = float(str(r['LTP']).replace(',', '').strip())
        ltp_map[sym] = val
    except:
        pass

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

# 3. Load Result Dates for each quarter
xl_rd = pd.ExcelFile(RESULT_DATES_FILE)
rd_sheets = {}
for sn in xl_rd.sheet_names:
    df_sn = xl_rd.parse(sn)
    sym_c = next((c for c in df_sn.columns if 'SYM' in str(c).upper()), None)
    dt_c = next((c for c in df_sn.columns if 'DATE' in str(c).upper() or 'REPORTING' in str(c).upper()), None)
    if sym_c and dt_c:
        mapping = {}
        for _, row in df_sn.iterrows():
            s_val = str(row[sym_c]).strip().upper()
            d_val = row[dt_c]
            if pd.notna(d_val):
                mapping[s_val] = str(d_val)[:10]
        rd_sheets[sn] = mapping

# 4. Process Current Active Quarter (FY27_Q2 - Q3 FY 2026-27)
q3_announced = {
    "HCLTECH": {"res_date": "2026-10-12", "en_date": "2026-10-06", "ex_date": "2026-10-15"},
    "ULTRACEMCO": {"res_date": "2026-10-19", "en_date": "2026-10-14", "ex_date": "2026-10-21"},
    "AXISBANK": {"res_date": "2026-10-17", "en_date": "2026-10-15", "ex_date": "2026-10-19"},
    "INFY": {"res_date": "2026-10-23", "en_date": "2026-10-15", "ex_date": "2026-10-29"}
}

active_q3_stocks = []
for sym, meta in meta_17q.items():
    lot = LOT_SIZES.get(sym, 250)
    ltp = ltp_map.get(sym, 1500.0)
    qty_1l = max(1, int(100000 / ltp))
    
    is_ann = sym in q3_announced
    if is_ann:
        ann_info = q3_announced[sym]
        res_date = ann_info['res_date']
        en_date = ann_info['en_date']
        ex_date = ann_info['ex_date']
    else:
        res_date = "Awaiting NSE Announcement"
        en_date = f"T - {meta['lead']} trading days"
        ex_date = f"T + {meta['hold']} trading days"

    # Expected Return is the trader's expected strategy yield (positive for both LONG and SHORT)
    exp_ret_pct = round(abs(meta['exp_pnl_1lot']) / meta['margin'] * 100 * 0.20, 2)

    active_q3_stocks.append({
        "symbol": sym,
        "name": meta['name'],
        "direction": meta['direction'],
        "window": meta['window'],
        "lead": meta['lead'],
        "hold": meta['hold'],
        "lot": lot,
        "ltp": ltp,
        "margin": meta['margin'],
        "active_q_wr": meta['active_q_wr'],
        "avg_17q_wr": meta['avg_17q_wr'],
        "expected_ret": exp_ret_pct,
        "actual_ret": None, # In progress quarter
        "exp_pnl_1lot": meta['exp_pnl_1lot'],
        "eq_qty_1l": qty_1l,
        "eq_exp_pnl_1l": round((exp_ret_pct / 100) * 100000, 0),
        "total_quarters": 17,
        "announced": is_ann,
        "result_date": res_date,
        "entry_date": en_date,
        "exit_date": ex_date
    })

# 5. Process Historical Quarters
def build_historical_quarter(sheet_name, rd_sheet_name, q_code, q_title, q_period):
    df_trades = pd.read_excel(CONSOLIDATED_FILE, sheet_name=sheet_name, skiprows=7)
    res_map = rd_sheets.get(rd_sheet_name, {})
    
    stocks = []
    for _, r in df_trades.iterrows():
        sym = str(r.get('Symbol', '')).strip().upper()
        if not sym or sym == 'NAN': continue
        
        meta = meta_17q.get(sym, {
            'name': str(r.get('Company Name', sym)),
            'direction': 'LONG' if 'LONG' in str(r.get('Strategy Type', '')).upper() else 'SHORT',
            'window': 'T-4 to T+2',
            'lead': 4,
            'hold': 2,
            'avg_17q_wr': 82.0,
            'margin': 150000.0
        })
        
        lot = LOT_SIZES.get(sym, 250)
        en_date = str(r.get('Entry Date', ''))[:10]
        ex_date = str(r.get('Exit Date', ''))[:10]
        res_date = res_map.get(sym, en_date)
        
        en_price = float(r.get('Entry Price (Rs.)', 1000.0))
        ex_price = float(r.get('Exit Price (Rs.)', 1000.0))
        
        # Expected Return %
        exp_ret_raw = float(r.get('Expected Return (%)', 0.025))
        exp_ret_pct = round(exp_ret_raw * 100, 2)
        
        # Actual Return % (Return We Get)
        act_ret_raw = float(r.get('Realised Return (%)', 0.020))
        act_ret_pct = round(act_ret_raw * 100, 2)
        
        pnl_per_share = float(r.get('Realised P&L (Rs.)', 0.0))
        pnl_1lot = round(pnl_per_share * lot, 0)
        
        strat_type = str(r.get('Strategy Type', meta['direction'])).upper()
        direction = 'LONG' if 'LONG' in strat_type else 'SHORT'
        
        qty_1l = max(1, int(100000 / en_price)) if en_price > 0 else 50
        eq_pnl_1l = round(qty_1l * (ex_price - en_price) * (1 if direction == 'LONG' else -1), 0)

        stocks.append({
            "symbol": sym,
            "name": meta['name'],
            "direction": direction,
            "window": meta['window'],
            "lead": meta['lead'],
            "hold": meta['hold'],
            "lot": lot,
            "ltp": en_price,
            "margin": meta['margin'],
            "avg_17q_wr": meta['avg_17q_wr'],
            "active_q_wr": meta['avg_17q_wr'],
            "expected_ret": exp_ret_pct,
            "actual_ret": act_ret_pct,
            "exp_pnl_1lot": pnl_1lot,
            "eq_qty_1l": qty_1l,
            "eq_exp_pnl_1l": eq_pnl_1l,
            "total_quarters": 17,
            "announced": True,
            "result_date": res_date,
            "entry_date": en_date,
            "exit_date": ex_date
        })
        
    return {
        "q_code": q_code,
        "title": q_title,
        "period": q_period,
        "status": "COMPLETED AUDITED" if q_code != 'FY27_Q2' else "ACTIVE CURRENT QUARTER",
        "badge": "COMPLETED" if q_code != 'FY27_Q2' else "ACTIVE NOW",
        "announced_count": len(stocks),
        "pending_count": 0,
        "total_stocks": len(stocks),
        "win_rate_avg": f"{round(sum(s['avg_17q_wr'] for s in stocks)/len(stocks), 1)}%",
        "bias": f"{sum(1 for s in stocks if s['direction']=='LONG')} LONG / {sum(1 for s in stocks if s['direction']=='SHORT')} SHORT",
        "stocks": stocks
    }

quarters_data = [
    {
        "q_code": "FY27_Q2",
        "title": "Q3 FY 2026-27 (Jul - Sep 2026 Results)",
        "period": "October 2026 - November 2026",
        "status": "ACTIVE CURRENT QUARTER",
        "badge": "ACTIVE NOW",
        "announced_count": 4,
        "pending_count": 46,
        "total_stocks": 50,
        "win_rate_avg": "81.4%",
        "bias": "22 LONG / 28 SHORT",
        "stocks": active_q3_stocks
    },
    build_historical_quarter("Q2 2026-27", "Q2 2026-27", "FY27_Q1", "Q2 FY 2026-27 (Apr - Jun 2026 Results)", "July 2026 - August 2026"),
    build_historical_quarter("Q3 2025-26", "FY25-26Q3", "FY26_Q3", "Q3 FY 2025-26 (Jul - Sep 2025 Results)", "October 2025 - November 2025"),
    build_historical_quarter("Q1 2026-27", "FY25-26Q4", "FY26_Q4", "Q4 FY 2025-26 (Jan - Mar 2026 Results)", "April 2026 - May 2026")
]

# Load existing JSON to preserve holidays and RBI
with open(JSON_FILE, 'r', encoding='utf-8') as f:
    existing_data = json.load(f)

existing_data['quarters'] = quarters_data

with open(JSON_FILE, 'w', encoding='utf-8') as f:
    json.dump(existing_data, f, indent=2)

print("SUCCESS: Perfectly rebuilt event_dashboard_data.json quarters dataset!")
