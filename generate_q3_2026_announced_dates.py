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
print("COMPREHENSIVE AUDIT & FETCH: Q3 ANNOUNCED RESULT DATES FOR YEAR 2026")
print("==========================================================================================")
print(f"Target Universe: {len(SYMBOLS)} Stocks")
print(f"Filter Target: Q3 Financial Results Announced / Broadcast in 2026")
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
        except Exception:
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

q3_2026_broadcasts = []
q3_2026_board_meetings = []

def process_stock(sym):
    stock_folder = ROOT / sym
    stock_folder.mkdir(parents=True, exist_ok=True)
    encoded_sym = urllib.parse.quote(sym)
    
    # 1. Fetch live results comparison feed
    res_url = f"https://www.nseindia.com/api/results-comparision?symbol={encoded_sym}"
    res_data = nse.get_json(res_url, sym) or {}
    res_raw = (res_data.get('resCmpData') or []) if isinstance(res_data, dict) else []
    
    stock_fr_rows = []
    if isinstance(res_raw, list) and len(res_raw) > 0:
        for r in res_raw:
            if not isinstance(r, dict): continue
            dt_raw = r.get('re_create_dt', '')
            dt = pd.to_datetime(dt_raw, errors='coerce')
            rel = f"Qtr ended {r.get('re_to_dt', '')}"
            to_d = r.get('re_to_dt', '')
            
            # Save all financial results
            stock_fr_rows.append({
                'broadCastDate': dt_raw,
                'relatingTo': rel,
                'consolidated': r.get('re_res_type', 'Audited'),
                'audited': 'Audited',
                'fromDate': r.get('re_from_dt', ''),
                'toDate': to_d,
                'xbrl': ''
            })
            
            # Q3 2026 check (announced in 2026 or relating to Q3 ended Dec 31)
            if pd.notna(dt) and dt.year == 2026:
                if '31-DEC' in rel.upper() or '31-DEC' in to_d.upper() or dt.month in (1, 2, 3):
                    q3_2026_broadcasts.append({
                        'symbol': sym,
                        'broadcast_date': dt.strftime('%Y-%m-%d'),
                        'announcement_time': dt.strftime('%H:%M:%S') if dt.hour != 0 else 'N/A',
                        'month_name': dt.strftime('%B'),
                        'quarter_name': 'Q3 (Oct-Dec)',
                        'relating_to': rel,
                        'from_date': r.get('re_from_dt', ''),
                        'to_date': to_d,
                        'res_type': r.get('re_res_type', 'Audited')
                    })
                    
        if stock_fr_rows:
            pd.DataFrame(stock_fr_rows).to_csv(stock_folder / 'financial_results.csv', index=False)

    # 2. Check local board meetings file & announcements file
    bm_path = stock_folder / 'board_meetings.csv'
    if bm_path.exists() and bm_path.stat().st_size > 10:
        try:
            df_bm = pd.read_csv(bm_path, dtype=str).fillna('')
            for _, r in df_bm.iterrows():
                dt_str = r.get('bm_date', '') or r.get('announcement_date', '')
                dt = pd.to_datetime(dt_str, errors='coerce')
                purpose = str(r.get('bm_purpose', '')) + " " + str(r.get('purpose', '')) + " " + str(r.get('bm_desc', ''))
                
                if pd.notna(dt) and dt.year == 2026:
                    if any(k in purpose.lower() for k in ['result', 'financial', 'q3', 'december', 'board meeting']):
                        q3_2026_board_meetings.append({
                            'symbol': sym,
                            'announcement_date': dt.strftime('%Y-%m-%d %H:%M:%S'),
                            'month_name': dt.strftime('%B'),
                            'purpose': r.get('bm_purpose') or r.get('purpose') or 'Board Meeting for Financial Results',
                            'details': r.get('bm_desc') or r.get('details') or ''
                        })
        except Exception:
            pass

print("Processing universe stocks in parallel pool...", flush=True)

with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
    futures = {executor.submit(process_stock, sym): sym for sym in SYMBOLS}
    count = 0
    for future in concurrent.futures.as_completed(futures):
        count += 1
        if count % 50 == 0 or count == len(SYMBOLS):
            print(f"  Progress: {count}/{len(SYMBOLS)} stocks completed...", flush=True)

print("------------------------------------------------------------------------------------------", flush=True)
print("EXTRACTED Q3 2026 ANNOUNCED RESULT DATES SUMMARY:")
print("------------------------------------------------------------------------------------------", flush=True)

df_q3_res = pd.DataFrame(q3_2026_broadcasts)
df_q3_bm = pd.DataFrame(q3_2026_board_meetings)

if not df_q3_res.empty:
    df_q3_res = df_q3_res.sort_values(['broadcast_date', 'symbol'], ascending=[False, True])
    print(f"\nTotal Q3 2026 Financial Result Broadcast Dates Found: {len(df_q3_res)}")
    print("\nQ3 2026 Financial Result Broadcast Dates across Universe:")
    print(df_q3_res[['symbol', 'broadcast_date', 'month_name', 'relating_to']].to_string(index=False), flush=True)

if not df_q3_bm.empty:
    df_q3_bm = df_q3_bm.sort_values(['announcement_date', 'symbol'], ascending=[False, True])
    print(f"\nTotal Q3 2026 Board Meeting Outcomes & Intimations: {len(df_q3_bm)}")
    print("\nSample Q3 2026 Board Meeting Outcomes:")
    print(df_q3_bm[['symbol', 'announcement_date', 'purpose']].head(20).to_string(index=False), flush=True)

# Build Stock-Wise Q3 2026 Summary
stock_q3_rows = []
all_symbols_with_q3 = set(df_q3_res['symbol'].tolist() if not df_q3_res.empty else []) | set(df_q3_bm['symbol'].tolist() if not df_q3_bm.empty else [])

for sym in sorted(all_symbols_with_q3):
    res_dates = df_q3_res[df_q3_res['symbol'] == sym]['broadcast_date'].tolist() if not df_q3_res.empty else []
    bm_dates = df_q3_bm[df_q3_bm['symbol'] == sym]['announcement_date'].tolist() if not df_q3_bm.empty else []
    
    stock_q3_rows.append({
        'symbol': sym,
        'q3_result_broadcast_dates': ", ".join(sorted(set(res_dates))),
        'q3_board_meeting_dates': ", ".join(sorted(set(bm_dates)))
    })

df_stock_q3_sum = pd.DataFrame(stock_q3_rows)

# Generate Master Excel Workbook
out_excel = ROOT / 'Q3_2026_Announced_Result_Dates_Master.xlsx'
wb = openpyxl.Workbook()

hdr_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
hdr_fill = PatternFill("solid", fgColor="1F4E79")
border_thin = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                     top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))

# Sheet 1: Q3 2026 Result Broadcast Dates
ws1 = wb.active
ws1.title = "Q3 2026 Result Broadcast Dates"
ws1.views.sheetView[0].showGridLines = True

headers1 = ["Sr. No.", "Stock Symbol", "Q3 Broadcast Date (2026)", "Month", "Quarter Name", "Financial Period Description", "Period From Date", "Period To Date", "Audit Status"]
ws1.row_dimensions[1].height = 26
for c_idx, h in enumerate(headers1, 1):
    cell = ws1.cell(row=1, column=c_idx, value=h)
    cell.font = hdr_font
    cell.fill = hdr_fill
    cell.alignment = Alignment(horizontal="center", vertical="center")

if not df_q3_res.empty:
    for r_idx, r in enumerate(df_q3_res.to_dict('records'), start=2):
        ws1.cell(row=r_idx, column=1, value=r_idx - 1).alignment = Alignment(horizontal="center")
        ws1.cell(row=r_idx, column=2, value=r['symbol'])
        ws1.cell(row=r_idx, column=3, value=r['broadcast_date']).alignment = Alignment(horizontal="center")
        ws1.cell(row=r_idx, column=4, value=r['month_name']).alignment = Alignment(horizontal="center")
        ws1.cell(row=r_idx, column=5, value=r['quarter_name']).alignment = Alignment(horizontal="center")
        ws1.cell(row=r_idx, column=6, value=r['relating_to'])
        ws1.cell(row=r_idx, column=7, value=r['from_date']).alignment = Alignment(horizontal="center")
        ws1.cell(row=r_idx, column=8, value=r['to_date']).alignment = Alignment(horizontal="center")
        ws1.cell(row=r_idx, column=9, value=r['res_type']).alignment = Alignment(horizontal="center")
        for c in range(1, 10):
            ws1.cell(row=r_idx, column=c).border = border_thin

for col in ws1.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws1.column_dimensions[col_letter].width = max(max_len + 3, 12)

# Sheet 2: Q3 2026 Board Meeting Outcomes
ws2 = wb.create_sheet(title="Q3 2026 Board Meetings")
ws2.views.sheetView[0].showGridLines = True

headers2 = ["Sr. No.", "Stock Symbol", "Announcement Date & Time", "Month", "Board Meeting Purpose", "Announcement Details"]
ws2.row_dimensions[1].height = 26
for c_idx, h in enumerate(headers2, 1):
    cell = ws2.cell(row=1, column=c_idx, value=h)
    cell.font = hdr_font
    cell.fill = PatternFill("solid", fgColor="2F5597")
    cell.alignment = Alignment(horizontal="center", vertical="center")

if not df_q3_bm.empty:
    for r_idx, r in enumerate(df_q3_bm.to_dict('records'), start=2):
        ws2.cell(row=r_idx, column=1, value=r_idx - 1).alignment = Alignment(horizontal="center")
        ws2.cell(row=r_idx, column=2, value=r['symbol'])
        ws2.cell(row=r_idx, column=3, value=r['announcement_date']).alignment = Alignment(horizontal="center")
        ws2.cell(row=r_idx, column=4, value=r['month_name']).alignment = Alignment(horizontal="center")
        ws2.cell(row=r_idx, column=5, value=r['purpose'])
        ws2.cell(row=r_idx, column=6, value=r['details'])
        for c in range(1, 7):
            ws2.cell(row=r_idx, column=c).border = border_thin

for col in ws2.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws2.column_dimensions[col_letter].width = max(max_len + 3, 12)

# Sheet 3: Stock-Wise Q3 2026 Summary
ws3 = wb.create_sheet(title="Stock-Wise Q3 2026 Summary")
ws3.views.sheetView[0].showGridLines = True

headers3 = ["Sr. No.", "Stock Symbol", "Q3 2026 Result Broadcast Dates", "Q3 2026 Board Meeting Dates"]
ws3.row_dimensions[1].height = 26
for c_idx, h in enumerate(headers3, 1):
    cell = ws3.cell(row=1, column=c_idx, value=h)
    cell.font = hdr_font
    cell.fill = PatternFill("solid", fgColor="333F48")
    cell.alignment = Alignment(horizontal="center", vertical="center")

if not df_stock_q3_sum.empty:
    for r_idx, r in enumerate(df_stock_q3_sum.to_dict('records'), start=2):
        ws3.cell(row=r_idx, column=1, value=r_idx - 1).alignment = Alignment(horizontal="center")
        ws3.cell(row=r_idx, column=2, value=r['symbol'])
        ws3.cell(row=r_idx, column=3, value=r['q3_result_broadcast_dates'])
        ws3.cell(row=r_idx, column=4, value=r['q3_board_meeting_dates'])
        for c in range(1, 5):
            ws3.cell(row=r_idx, column=c).border = border_thin

for col in ws3.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws3.column_dimensions[col_letter].width = max(max_len + 3, 12)

wb.save(out_excel)
print(f"\nSuccessfully generated Master Excel: {out_excel}", flush=True)

# Save CSV
if not df_q3_res.empty:
    df_q3_res.to_csv(PROC / 'q3_2026_announced_results_master.csv', index=False)

print("==========================================================================================", flush=True)
print("SUCCESS: Q3 2026 Announced Result Dates Audit & Export Completed!", flush=True)
print("==========================================================================================", flush=True)
