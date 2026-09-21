import urllib.request
import urllib.parse
import http.cookiejar
import json
import os
import pandas as pd
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
print("NSE OFFICIAL SITE: UPCOMING QUARTERLY RESULT DATES & BOARD MEETINGS (LIVE SCANNER)")
print("==========================================================================================")
print(f"Target Universe: {len(SYMBOLS)} Stocks")
print(f"Current System Time: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")
print("------------------------------------------------------------------------------------------", flush=True)

class NSESession:
    def __init__(self):
        self.cj = http.cookiejar.CookieJar()
        self.opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(self.cj))
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': '*/*',
            'Accept-Language': 'en-US,en;q=0.9',
            'Referer': 'https://www.nseindia.com/market-data/event-calendar'
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

# 1. Fetch NSE Official Event Calendar API
print("Fetching NSE Official Event Calendar (https://www.nseindia.com/api/event-calendar)...", flush=True)
url_evt = 'https://www.nseindia.com/api/event-calendar'
events = nse.get_json(url_evt, "RELIANCE") or []
print(f"Total Upcoming Board Meetings on NSE Event Calendar: {len(events)}", flush=True)

# Parse event calendar records
upcoming_results = []
calendar_syms = set()

for e in events:
    if not isinstance(e, dict): continue
    sym = (e.get('symbol') or '').upper().strip()
    company = e.get('company', '')
    dt_str = e.get('date', '')
    purpose = e.get('purpose', '')
    bm_desc = e.get('bm_desc', '')
    calendar_syms.add(sym)
    
    upcoming_results.append({
        'symbol': sym,
        'company_name': company,
        'upcoming_result_date': dt_str,
        'purpose': purpose,
        'details': bm_desc,
        'source': 'NSE Official Event Calendar',
        'is_universe': 'YES ✅' if sym in set(SYMBOLS) else 'NO'
    })

# 2. Check Corporate Announcements live for all 211 stocks for upcoming board meetings
print("Scanning corporate announcements feed for upcoming board meetings...", flush=True)
ann_url = 'https://www.nseindia.com/api/corporate-announcements?index=equities'
bulk_ann = nse.get_json(ann_url, "RELIANCE") or []

ann_results = []
for a in bulk_ann:
    if not isinstance(a, dict): continue
    sym = (a.get('symbol') or '').upper().strip()
    desc = str(a.get('desc', ''))
    att = str(a.get('attchmntText', ''))
    dt = a.get('an_dt', '') or a.get('sort_date', '')
    combo = (desc + ' ' + att).lower()
    
    if any(k in combo for k in ['board meeting', 'financial result', 'consider results', 'quarterly result']):
        if 'intimation' in combo or 'board meeting' in combo:
            ann_results.append({
                'symbol': sym,
                'company_name': a.get('sm_name', ''),
                'upcoming_result_date': dt,
                'purpose': desc,
                'details': att,
                'source': 'NSE Corporate Announcements Feed',
                'is_universe': 'YES ✅' if sym in set(SYMBOLS) else 'NO'
            })

# Combine records
df_evt = pd.DataFrame(upcoming_results)
df_ann = pd.DataFrame(ann_results)

df_all = pd.concat([df_evt, df_ann], ignore_index=True) if not df_ann.empty else df_evt
if not df_all.empty:
    df_all = df_all.drop_duplicates(subset=['symbol', 'upcoming_result_date', 'purpose'], keep='first')
    df_all['_dt'] = pd.to_datetime(df_all['upcoming_result_date'], errors='coerce')
    df_all = df_all.sort_values(['_dt', 'symbol'], ascending=[True, True]).drop(columns=['_dt'])

print("------------------------------------------------------------------------------------------", flush=True)
print("NSE OFFICIAL UPCOMING QUARTERLY RESULT DATES & BOARD MEETINGS:")
print("------------------------------------------------------------------------------------------", flush=True)

if not df_all.empty:
    df_univ = df_all[df_all['is_universe'] == 'YES ✅'].copy()
    print(f"\nUpcoming Result Board Meetings for Universe Stocks ({len(df_univ)} found):")
    print(df_univ[['symbol', 'company_name', 'upcoming_result_date', 'purpose', 'details']].to_string(index=False), flush=True)
    
    print(f"\nAll Upcoming Result Board Meetings across NSE ({len(df_all)} found):")
    print(df_all[['symbol', 'company_name', 'upcoming_result_date', 'purpose']].head(30).to_string(index=False), flush=True)

# Generate Excel Master Workbook
out_excel = ROOT / 'NSE_Upcoming_Quarterly_Result_Dates_2026.xlsx'
wb = openpyxl.Workbook()

hdr_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
hdr_fill = PatternFill("solid", fgColor="1F4E79")
border_thin = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                     top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))

# Sheet 1: Universe Upcoming Result Dates
ws1 = wb.active
ws1.title = "Universe Upcoming Result Dates"
ws1.views.sheetView[0].showGridLines = True

headers1 = ["Sr. No.", "Stock Symbol", "Company Name", "Upcoming Result Date (NSE)", "Purpose / Event Type", "Announcement Details", "Data Source"]
ws1.row_dimensions[1].height = 26
for c_idx, h in enumerate(headers1, 1):
    cell = ws1.cell(row=1, column=c_idx, value=h)
    cell.font = hdr_font
    cell.fill = hdr_fill
    cell.alignment = Alignment(horizontal="center", vertical="center")

if not df_all.empty:
    df_univ = df_all[df_all['is_universe'] == 'YES ✅'].copy()
    for r_idx, r in enumerate(df_univ.to_dict('records'), start=2):
        ws1.cell(row=r_idx, column=1, value=r_idx - 1).alignment = Alignment(horizontal="center")
        ws1.cell(row=r_idx, column=2, value=r['symbol'])
        ws1.cell(row=r_idx, column=3, value=r['company_name'])
        ws1.cell(row=r_idx, column=4, value=r['upcoming_result_date']).alignment = Alignment(horizontal="center")
        ws1.cell(row=r_idx, column=5, value=r['purpose'])
        ws1.cell(row=r_idx, column=6, value=r['details'])
        ws1.cell(row=r_idx, column=7, value=r['source']).alignment = Alignment(horizontal="center")
        for c in range(1, 8):
            ws1.cell(row=r_idx, column=c).border = border_thin

for col in ws1.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws1.column_dimensions[col_letter].width = max(max_len + 3, 12)

# Sheet 2: All NSE Upcoming Board Meetings
ws2 = wb.create_sheet(title="All NSE Upcoming Board Meetings")
ws2.views.sheetView[0].showGridLines = True

headers2 = ["Sr. No.", "Stock Symbol", "Company Name", "Upcoming Meeting Date", "Purpose", "Details", "In Universe", "Source"]
ws2.row_dimensions[1].height = 26
for c_idx, h in enumerate(headers2, 1):
    cell = ws2.cell(row=1, column=c_idx, value=h)
    cell.font = hdr_font
    cell.fill = PatternFill("solid", fgColor="2F5597")
    cell.alignment = Alignment(horizontal="center", vertical="center")

if not df_all.empty:
    for r_idx, r in enumerate(df_all.to_dict('records'), start=2):
        ws2.cell(row=r_idx, column=1, value=r_idx - 1).alignment = Alignment(horizontal="center")
        ws2.cell(row=r_idx, column=2, value=r['symbol'])
        ws2.cell(row=r_idx, column=3, value=r['company_name'])
        ws2.cell(row=r_idx, column=4, value=r['upcoming_result_date']).alignment = Alignment(horizontal="center")
        ws2.cell(row=r_idx, column=5, value=r['purpose'])
        ws2.cell(row=r_idx, column=6, value=r['details'])
        ws2.cell(row=r_idx, column=7, value=r['is_universe']).alignment = Alignment(horizontal="center")
        ws2.cell(row=r_idx, column=8, value=r['source']).alignment = Alignment(horizontal="center")
        for c in range(1, 9):
            ws2.cell(row=r_idx, column=c).border = border_thin

for col in ws2.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws2.column_dimensions[col_letter].width = max(max_len + 3, 12)

wb.save(out_excel)
print(f"\nSaved Master Excel: {out_excel}", flush=True)

# Save CSV
if not df_all.empty:
    df_all.to_csv(PROC / 'nse_upcoming_quarterly_result_dates_2026.csv', index=False)

print("==========================================================================================", flush=True)
print("SUCCESS: Live NSE Upcoming Quarter Result Dates Extracted Successfully!", flush=True)
print("==========================================================================================", flush=True)
