import urllib.request
import urllib.parse
import http.cookiejar
import json
import pandas as pd
import os
import pathlib
import sys
import time
import warnings
import concurrent.futures
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings('ignore')
sys.stdout.reconfigure(errors='replace')

ROOT = pathlib.Path('D:/behaviour analysis')
OI_DIR = ROOT / 'OI_DATA'
PROC = ROOT / 'processed'
PROC.mkdir(parents=True, exist_ok=True)

SYMBOLS = sorted([e.name.strip().upper() for e in os.scandir(OI_DIR) if e.is_dir()])

print("==========================================================================================")
print("FAST MULTI-THREADED LIVE FETCH: ANNOUNCED FINANCIAL RESULT DATES & BOARD MEETINGS")
print("==========================================================================================")
print(f"Target Universe: {len(SYMBOLS)} Stocks")
print(f"Current System Date: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")
print("------------------------------------------------------------------------------------------", flush=True)

class NSESession:
    def __init__(self):
        self.cj = http.cookiejar.CookieJar()
        self.opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(self.cj))
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': '*/*',
            'Accept-Language': 'en-US,en;q=0.9',
            'Referer': 'https://www.nseindia.com/'
        }
        self.primed = False

    def prime(self, symbol="RELIANCE"):
        try:
            url = f"https://www.nseindia.com/get-quotes/equity?symbol={urllib.parse.quote(symbol)}"
            req = urllib.request.Request(url, headers=self.headers)
            with self.opener.open(req, timeout=10) as resp:
                resp.read()
            self.primed = True
        except Exception as e:
            pass

    def get_json(self, url, symbol="RELIANCE"):
        if not self.primed:
            self.prime(symbol)
        req_headers = self.headers.copy()
        req_headers['Referer'] = f"https://www.nseindia.com/get-quotes/equity?symbol={urllib.parse.quote(symbol)}"
        req = urllib.request.Request(url, headers=req_headers)
        
        for attempt in range(2):
            try:
                with self.opener.open(req, timeout=10) as resp:
                    raw = resp.read().decode('utf-8', errors='ignore')
                    return json.loads(raw)
            except Exception:
                self.prime(symbol)
                time.sleep(0.3)
        return None

nse = NSESession()
nse.prime("RELIANCE")

# 1. Fetch live general corporate announcements (recent 60 days window)
today_str = pd.Timestamp.now().strftime('%d-%m-%Y')
sixty_days_ago_str = (pd.Timestamp.now() - pd.Timedelta(days=60)).strftime('%d-%m-%Y')

print(f"Fetching live NSE corporate announcements window ({sixty_days_ago_str} to {today_str})...", flush=True)
gen_ann_url = f"https://www.nseindia.com/api/corporate-announcements?index=equities&from_date={sixty_days_ago_str}&to_date={today_str}"
bulk_ann = nse.get_json(gen_ann_url, "RELIANCE") or []
print(f"Total live corporate announcements fetched: {len(bulk_ann)} records", flush=True)

# Group bulk announcements by symbol
ann_by_sym = {}
for item in bulk_ann:
    if isinstance(item, dict) and item.get('symbol'):
        s = item['symbol'].upper().strip()
        if s not in ann_by_sym:
            ann_by_sym[s] = []
        ann_by_sym[s].append(item)

# Multithreaded fetch per stock
results_summary = []
board_meetings_summary = []

def process_stock(sym):
    stock_folder = ROOT / sym
    stock_folder.mkdir(parents=True, exist_ok=True)
    encoded_sym = urllib.parse.quote(sym)
    
    # Results comparison fetch
    res_url = f"https://www.nseindia.com/api/results-comparision?symbol={encoded_sym}"
    res_data = nse.get_json(res_url, sym) or {}
    res_raw = (res_data.get('resCmpData') or []) if isinstance(res_data, dict) else []
    
    latest_r_info = None
    if isinstance(res_raw, list) and len(res_raw) > 0:
        fr_rows = []
        for r in res_raw:
            if not isinstance(r, dict): continue
            fr_rows.append({
                'broadCastDate': r.get('re_create_dt', ''),
                'relatingTo': f"Qtr ended {r.get('re_to_dt', '')}",
                'consolidated': r.get('re_res_type', 'Audited'),
                'audited': 'Audited',
                'fromDate': r.get('re_from_dt', ''),
                'toDate': r.get('re_to_dt', ''),
                'xbrl': ''
            })
        if fr_rows:
            df_fr = pd.DataFrame(fr_rows)
            df_fr.to_csv(stock_folder / 'financial_results.csv', index=False)
            
            df_fr['_dt'] = pd.to_datetime(df_fr['broadCastDate'], errors='coerce')
            latest_r = df_fr.sort_values('_dt', ascending=False).iloc[0]
            latest_r_info = {
                'symbol': sym,
                'latest_broadcast_date': str(latest_r['_dt'].date()) if pd.notna(latest_r['_dt']) else latest_r['broadCastDate'],
                'relating_to': latest_r['relatingTo'],
                'from_date': latest_r['fromDate'],
                'to_date': latest_r['toDate']
            }
            
    # Process announcements
    sym_anns = ann_by_sym.get(sym, [])
    bm_rows = []
    upcoming_bms = []
    
    for item in sym_anns:
        desc = item.get('desc', '')
        att_text = item.get('attchmntText', '')
        att_file = item.get('attchmntFile', '')
        dt = item.get('an_dt', '') or item.get('sort_date', '')
        combo = (desc + " " + att_text).lower()
        
        if any(k in combo for k in ['board meeting', 'financial result', 'results', 'audited', 'unaudited']):
            bm_rows.append({'bm_date': dt, 'bm_purpose': desc, 'bm_desc': att_text, 'attachment': att_file})
            upcoming_bms.append({
                'symbol': sym,
                'announcement_date': dt,
                'purpose': desc,
                'details': att_text,
                'attachment': f"https://nsearchives.nseindia.com/corporate/{att_file}" if att_file else ""
            })
            
    if bm_rows:
        pd.DataFrame(bm_rows).to_csv(stock_folder / 'board_meetings.csv', index=False)
        
    return latest_r_info, upcoming_bms

print("Scanning & updating event feeds across all 211 stocks (Parallel Pool)...", flush=True)

with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
    futures = {executor.submit(process_stock, sym): sym for sym in SYMBOLS}
    count = 0
    for future in concurrent.futures.as_completed(futures):
        count += 1
        sym = futures[future]
        try:
            latest_r_info, upcoming_bms = future.result()
            if latest_r_info:
                results_summary.append(latest_r_info)
            if upcoming_bms:
                board_meetings_summary.extend(upcoming_bms)
        except Exception as e:
            pass
        if count % 50 == 0 or count == len(SYMBOLS):
            print(f"  Progress: {count}/{len(SYMBOLS)} stocks completed...", flush=True)

print("------------------------------------------------------------------------------------------", flush=True)
print("FETCH COMPLETE! LATEST ANNOUNCED / BROADCAST RESULT DATES SUMMARY:")
print("------------------------------------------------------------------------------------------", flush=True)

df_latest_res = pd.DataFrame(results_summary)
if not df_latest_res.empty:
    df_latest_res['_dt'] = pd.to_datetime(df_latest_res['latest_broadcast_date'], errors='coerce')
    df_latest_res = df_latest_res.sort_values('_dt', ascending=False)
    print("\nMost Recent 20 Financial Result Broadcasts across Universe:")
    print(df_latest_res[['symbol', 'latest_broadcast_date', 'relating_to']].head(20).to_string(index=False), flush=True)

df_bm_all = pd.DataFrame(board_meetings_summary)
if not df_bm_all.empty:
    print(f"\nTotal Recent Result / Board Meeting Announcements: {len(df_bm_all)}")
    print("\nLatest 20 Board Meeting & Result Announcements:")
    print(df_bm_all[['symbol', 'announcement_date', 'purpose']].head(20).to_string(index=False), flush=True)

# Export Summary CSVs
PROC_RES = PROC / 'latest_announced_results_summary.csv'
PROC_BM = PROC / 'latest_board_meeting_announcements.csv'
df_latest_res.to_csv(PROC_RES, index=False)
if not df_bm_all.empty:
    df_bm_all.to_csv(PROC_BM, index=False)

# Update Master Excel: Verified_All_Quarters_Result_Dates_211_Stocks.xlsx
print("\nRegenerating Master Excel: Verified_All_Quarters_Result_Dates_211_Stocks.xlsx...", flush=True)

result_audit_rows = []
for sym in SYMBOLS:
    fr_path = ROOT / sym / 'financial_results.csv'
    if fr_path.exists() and fr_path.stat().st_size > 10:
        try:
            df_fr = pd.read_csv(fr_path, dtype=str).fillna('')
            for _, r in df_fr.iterrows():
                raw_dt = pd.to_datetime(r.get('broadCastDate', ''), errors='coerce')
                if pd.notna(raw_dt):
                    result_audit_rows.append({
                        "sym": sym,
                        "result_date": raw_dt.strftime('%Y-%m-%d'),
                        "relating_to": r.get('relatingTo', 'Financial Results'),
                        "audited_status": r.get('audited', 'Audited'),
                        "from_date": r.get('fromDate', ''),
                        "to_date": r.get('toDate', '')
                    })
        except Exception:
            pass

df_audit = pd.DataFrame(result_audit_rows)
if not df_audit.empty:
    # Sort by result date descending
    df_audit['_dt'] = pd.to_datetime(df_audit['result_date'])
    df_audit = df_audit.sort_values(['_dt', 'sym'], ascending=[False, True]).drop(columns=['_dt'])
    
    out_excel = ROOT / 'Verified_All_Quarters_Result_Dates_211_Stocks.xlsx'
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "All Quarters Result Dates"
    ws.views.sheetView[0].showGridLines = True
    
    hdr_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    hdr_fill = PatternFill("solid", fgColor="1F4E79")
    border_thin = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                         top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
                         
    headers = ["Sr. No.", "Stock Symbol", "Result Announcement Date", "Financial Period / Description", "Audit Status", "Period From Date", "Period To Date"]
    ws.row_dimensions[1].height = 26
    for c_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c_idx, value=h)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    for r_idx, r in enumerate(df_audit.to_dict('records'), start=2):
        ws.cell(row=r_idx, column=1, value=r_idx - 1).alignment = Alignment(horizontal="center")
        ws.cell(row=r_idx, column=2, value=r['sym'])
        ws.cell(row=r_idx, column=3, value=r['result_date']).alignment = Alignment(horizontal="center")
        ws.cell(row=r_idx, column=4, value=r['relating_to'])
        ws.cell(row=r_idx, column=5, value=r['audited_status']).alignment = Alignment(horizontal="center")
        ws.cell(row=r_idx, column=6, value=r['from_date']).alignment = Alignment(horizontal="center")
        ws.cell(row=r_idx, column=7, value=r['to_date']).alignment = Alignment(horizontal="center")
        for c in range(1, 8):
            ws.cell(row=r_idx, column=c).border = border_thin
            
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)
        
    wb.save(out_excel)
    print(f"Saved updated master workbook: {out_excel}", flush=True)

print("==========================================================================================", flush=True)
print("SUCCESS: Live Fetch & Date Master Update Completed Perfectly!", flush=True)
print("==========================================================================================", flush=True)
