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
print("NSE LIVE FETCH & AUDIT: OCTOBER TO DECEMBER QUARTERLY RESULT DATES & BOARD MEETINGS")
print("==========================================================================================")
print(f"Target Universe: {len(SYMBOLS)} Stocks")
print(f"Filter Month Window: October (10), November (11), December (12)")
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

# 1. Live corporate announcements fetch from NSE
print("Fetching live corporate announcements from NSE site for Oct-Dec window...", flush=True)
gen_ann_url = f"https://www.nseindia.com/api/corporate-announcements?index=equities"
bulk_ann = nse.get_json(gen_ann_url, "RELIANCE") or []
print(f"Total live corporate announcements fetched: {len(bulk_ann)} records", flush=True)

ann_by_sym = {}
for item in bulk_ann:
    if isinstance(item, dict) and item.get('symbol'):
        s = item['symbol'].upper().strip()
        if s not in ann_by_sym:
            ann_by_sym[s] = []
        ann_by_sym[s].append(item)

# Multithreaded per-stock processing
oct_dec_results = []
oct_dec_board_meetings = []

def process_stock(sym):
    stock_folder = ROOT / sym
    stock_folder.mkdir(parents=True, exist_ok=True)
    encoded_sym = urllib.parse.quote(sym)
    
    # 1. Fetch official results comparison feed
    res_url = f"https://www.nseindia.com/api/results-comparision?symbol={encoded_sym}"
    res_data = nse.get_json(res_url, sym) or {}
    res_raw = (res_data.get('resCmpData') or []) if isinstance(res_data, dict) else []
    
    stock_res_rows = []
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
            
            # Filter Oct, Nov, Dec broadcast dates
            for _, r in df_fr.iterrows():
                dt = pd.to_datetime(r.get('broadCastDate', ''), errors='coerce')
                if pd.notna(dt) and dt.month in (10, 11, 12):
                    stock_res_rows.append({
                        'symbol': sym,
                        'broadcast_date': dt.strftime('%Y-%m-%d'),
                        'year': dt.year,
                        'month_name': dt.strftime('%B'),
                        'month_num': dt.month,
                        'day_of_month': dt.day,
                        'relating_to': r.get('relatingTo', ''),
                        'from_date': r.get('fromDate', ''),
                        'to_date': r.get('toDate', '')
                    })
                    
    # 2. Process announcements for Oct-Dec board meetings
    sym_anns = ann_by_sym.get(sym, [])
    stock_bm_rows = []
    
    for item in sym_anns:
        desc = item.get('desc', '')
        att_text = item.get('attchmntText', '')
        att_file = item.get('attchmntFile', '')
        dt_str = item.get('an_dt', '') or item.get('sort_date', '')
        dt = pd.to_datetime(dt_str, errors='coerce')
        combo = (desc + " " + att_text).lower()
        
        if any(k in combo for k in ['board meeting', 'financial result', 'results', 'audited', 'unaudited']):
            if pd.notna(dt) and dt.month in (10, 11, 12):
                stock_bm_rows.append({
                    'symbol': sym,
                    'announcement_date': dt.strftime('%Y-%m-%d %H:%M:%S'),
                    'year': dt.year,
                    'month_name': dt.strftime('%B'),
                    'purpose': desc,
                    'details': att_text,
                    'attachment_url': f"https://nsearchives.nseindia.com/corporate/{att_file}" if att_file else ""
                })
                
    return stock_res_rows, stock_bm_rows

print("Scanning & extracting Oct-Dec quarterly result dates across 211 stocks...", flush=True)

with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
    futures = {executor.submit(process_stock, sym): sym for sym in SYMBOLS}
    count = 0
    for future in concurrent.futures.as_completed(futures):
        count += 1
        sym = futures[future]
        try:
            res_rows, bm_rows = future.result()
            if res_rows:
                oct_dec_results.extend(res_rows)
            if bm_rows:
                oct_dec_board_meetings.extend(bm_rows)
        except Exception:
            pass
        if count % 50 == 0 or count == len(SYMBOLS):
            print(f"  Progress: {count}/{len(SYMBOLS)} stocks completed...", flush=True)

print("------------------------------------------------------------------------------------------", flush=True)
print("EXTRACTED OCTOBER TO DECEMBER QUARTERLY RESULT DATES:")
print("------------------------------------------------------------------------------------------", flush=True)

df_results = pd.DataFrame(oct_dec_results)
df_bms = pd.DataFrame(oct_dec_board_meetings)

if not df_results.empty:
    df_results = df_results.sort_values(['broadcast_date', 'symbol'], ascending=[False, True])
    print(f"\nTotal Oct-Dec Quarterly Result Broadcast Dates Found: {len(df_results)}")
    print("\nSample Oct-Dec Result Broadcast Dates:")
    print(df_results[['symbol', 'broadcast_date', 'month_name', 'year', 'relating_to']].head(25).to_string(index=False), flush=True)

# Build Stock-Wise Oct-Dec Result Schedule Summary
stock_summary_rows = []
if not df_results.empty:
    for sym, group in df_results.groupby('symbol'):
        oct_dates = group[group['month_num'] == 10]['broadcast_date'].tolist()
        nov_dates = group[group['month_num'] == 11]['broadcast_date'].tolist()
        dec_dates = group[group['month_num'] == 12]['broadcast_date'].tolist()
        
        avg_day = round(group['day_of_month'].mean(), 1) if not group.empty else None
        
        stock_summary_rows.append({
            'symbol': sym,
            'total_oct_dec_results': len(group),
            'avg_day_of_month': avg_day,
            'october_dates': ", ".join(sorted(set(oct_dates))),
            'november_dates': ", ".join(sorted(set(nov_dates))),
            'december_dates': ", ".join(sorted(set(dec_dates)))
        })

df_stock_summary = pd.DataFrame(stock_summary_rows)

# Export Master Excel Workbook
out_excel = ROOT / 'Oct_Dec_Quarterly_Result_Dates_Master.xlsx'
wb = openpyxl.Workbook()

# Sheet 1: Detailed Oct-Dec Result Dates
ws1 = wb.active
ws1.title = "Oct-Dec Quarterly Result Dates"
ws1.views.sheetView[0].showGridLines = True

hdr_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
hdr_fill = PatternFill("solid", fgColor="1F4E79")
border_thin = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                     top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))

headers1 = ["Sr. No.", "Stock Symbol", "Result Broadcast Date", "Year", "Month", "Day of Month", "Financial Period / Description", "Period From Date", "Period To Date"]
ws1.row_dimensions[1].height = 26
for c_idx, h in enumerate(headers1, 1):
    cell = ws1.cell(row=1, column=c_idx, value=h)
    cell.font = hdr_font
    cell.fill = hdr_fill
    cell.alignment = Alignment(horizontal="center", vertical="center")

if not df_results.empty:
    for r_idx, r in enumerate(df_results.to_dict('records'), start=2):
        ws1.cell(row=r_idx, column=1, value=r_idx - 1).alignment = Alignment(horizontal="center")
        ws1.cell(row=r_idx, column=2, value=r['symbol'])
        ws1.cell(row=r_idx, column=3, value=r['broadcast_date']).alignment = Alignment(horizontal="center")
        ws1.cell(row=r_idx, column=4, value=r['year']).alignment = Alignment(horizontal="center")
        ws1.cell(row=r_idx, column=5, value=r['month_name']).alignment = Alignment(horizontal="center")
        ws1.cell(row=r_idx, column=6, value=r['day_of_month']).alignment = Alignment(horizontal="center")
        ws1.cell(row=r_idx, column=7, value=r['relating_to'])
        ws1.cell(row=r_idx, column=8, value=r['from_date']).alignment = Alignment(horizontal="center")
        ws1.cell(row=r_idx, column=9, value=r['to_date']).alignment = Alignment(horizontal="center")
        for c in range(1, 10):
            ws1.cell(row=r_idx, column=c).border = border_thin

for col in ws1.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws1.column_dimensions[col_letter].width = max(max_len + 3, 12)

# Sheet 2: Stock-Wise Oct-Dec Schedule Summary
ws2 = wb.create_sheet(title="Stock Oct-Dec Schedule Summary")
ws2.views.sheetView[0].showGridLines = True

headers2 = ["Sr. No.", "Stock Symbol", "Total Oct-Dec Results", "Average Day of Month", "October Broadcast Dates", "November Broadcast Dates", "December Broadcast Dates"]
ws2.row_dimensions[1].height = 26
for c_idx, h in enumerate(headers2, 1):
    cell = ws2.cell(row=1, column=c_idx, value=h)
    cell.font = hdr_font
    cell.fill = PatternFill("solid", fgColor="2F5597")
    cell.alignment = Alignment(horizontal="center", vertical="center")

if not df_stock_summary.empty:
    for r_idx, r in enumerate(df_stock_summary.to_dict('records'), start=2):
        ws2.cell(row=r_idx, column=1, value=r_idx - 1).alignment = Alignment(horizontal="center")
        ws2.cell(row=r_idx, column=2, value=r['symbol'])
        ws2.cell(row=r_idx, column=3, value=r['total_oct_dec_results']).alignment = Alignment(horizontal="center")
        ws2.cell(row=r_idx, column=4, value=r['avg_day_of_month']).alignment = Alignment(horizontal="center")
        ws2.cell(row=r_idx, column=5, value=r['october_dates'])
        ws2.cell(row=r_idx, column=6, value=r['november_dates'])
        ws2.cell(row=r_idx, column=7, value=r['december_dates'])
        for c in range(1, 8):
            ws2.cell(row=r_idx, column=c).border = border_thin

for col in ws2.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws2.column_dimensions[col_letter].width = max(max_len + 3, 12)

# Sheet 3: Oct-Dec Board Meeting Announcements
ws3 = wb.create_sheet(title="Oct-Dec Board Meetings")
ws3.views.sheetView[0].showGridLines = True

headers3 = ["Sr. No.", "Stock Symbol", "Announcement Date", "Year", "Month", "Purpose", "Details", "Attachment URL"]
ws3.row_dimensions[1].height = 26
for c_idx, h in enumerate(headers3, 1):
    cell = ws3.cell(row=1, column=c_idx, value=h)
    cell.font = hdr_font
    cell.fill = PatternFill("solid", fgColor="333F48")
    cell.alignment = Alignment(horizontal="center", vertical="center")

if not df_bms.empty:
    for r_idx, r in enumerate(df_bms.to_dict('records'), start=2):
        ws3.cell(row=r_idx, column=1, value=r_idx - 1).alignment = Alignment(horizontal="center")
        ws3.cell(row=r_idx, column=2, value=r['symbol'])
        ws3.cell(row=r_idx, column=3, value=r['announcement_date']).alignment = Alignment(horizontal="center")
        ws3.cell(row=r_idx, column=4, value=r['year']).alignment = Alignment(horizontal="center")
        ws3.cell(row=r_idx, column=5, value=r['month_name']).alignment = Alignment(horizontal="center")
        ws3.cell(row=r_idx, column=6, value=r['purpose'])
        ws3.cell(row=r_idx, column=7, value=r['details'])
        ws3.cell(row=r_idx, column=8, value=r['attachment_url'])
        for c in range(1, 9):
            ws3.cell(row=r_idx, column=c).border = border_thin

for col in ws3.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws3.column_dimensions[col_letter].width = max(max_len + 3, 12)

wb.save(out_excel)
print(f"\nSuccessfully generated Excel Master Workbook: {out_excel}", flush=True)

# Export CSVs
df_results.to_csv(PROC / 'oct_dec_quarterly_results_master.csv', index=False)
df_stock_summary.to_csv(PROC / 'oct_dec_stock_schedule_summary.csv', index=False)

print("==========================================================================================", flush=True)
print("SUCCESS: October to December Quarterly Result Dates Processed & Saved!", flush=True)
print("==========================================================================================", flush=True)
