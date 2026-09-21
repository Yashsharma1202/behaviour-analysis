import urllib.request
import urllib.parse
import http.cookiejar
import json
import pandas as pd
import os
import pathlib
import sys
import time

sys.stdout.reconfigure(errors='replace')

ROOT = pathlib.Path('D:/behaviour analysis')
OI_DIR = ROOT / 'OI_DATA'

# 1. Discover all 211 stock symbols
SYMBOLS = sorted([e.name.strip().upper() for e in os.scandir(OI_DIR) if e.is_dir()])

print("==========================================================================================")
print(f"BULK DOWNLOADER FOR ALL 211 NSE STOCKS EVENT FEEDS (OFFICIAL NSE SITE)")
print("==========================================================================================")
print(f"Total Target Stocks: {len(SYMBOLS)} Stocks")
print(f"First 10: {SYMBOLS[:10]}")
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
            with self.opener.open(req, timeout=15) as resp:
                resp.read()
            self.primed = True
        except Exception as e:
            print(f"NSE session prime error: {e}")

    def get_json(self, url, symbol="RELIANCE"):
        if not self.primed:
            self.prime(symbol)
        req_headers = self.headers.copy()
        req_headers['Referer'] = f"https://www.nseindia.com/get-quotes/equity?symbol={urllib.parse.quote(symbol)}"
        req = urllib.request.Request(url, headers=req_headers)
        
        for attempt in range(2):
            try:
                with self.opener.open(req, timeout=15) as resp:
                    raw = resp.read().decode('utf-8', errors='ignore')
                    return json.loads(raw)
            except Exception as e:
                self.prime(symbol)
                time.sleep(1)
        return None

nse = NSESession()
nse.prime("RELIANCE")

def process_stock(sym):
    stock_folder = ROOT / sym
    stock_folder.mkdir(parents=True, exist_ok=True)
    
    # Check if already processed
    if (stock_folder / "announcements.csv").exists() and (stock_folder / "financial_results.csv").exists():
        df_ann = pd.read_csv(stock_folder / "announcements.csv")
        if len(df_ann) > 0:
            return sym, len(df_ann), 0, True
            
    encoded_sym = urllib.parse.quote(sym)
    
    # 1. Fetch Announcements
    ann_url = f"https://www.nseindia.com/api/corporate-announcements?index=equities&symbol={encoded_sym}"
    ann_raw = nse.get_json(ann_url, sym) or []
    
    # 2. Fetch Results
    res_url = f"https://www.nseindia.com/api/results-comparision?symbol={encoded_sym}"
    res_data = nse.get_json(res_url, sym) or {}
    res_raw = res_data.get('resCmpData', []) if isinstance(res_data, dict) else []
    
    ann_rows, bm_rows, ca_rows = [], [], []
    for item in ann_raw:
        if not isinstance(item, dict): continue
        desc = item.get('desc', '')
        att_text = item.get('attchmntText', '')
        att_file = item.get('attchmntFile', '')
        dt = item.get('an_dt', '') or item.get('sort_date', '')
        
        ann_rows.append({'sort_date': dt, 'desc': desc, 'attchmntText': att_text, 'attchmntFile': att_file})
        
        if 'board meeting' in desc.lower() or 'financial result' in desc.lower() or 'board meeting' in att_text.lower():
            bm_rows.append({'bm_date': dt, 'bm_purpose': desc, 'bm_desc': att_text, 'attachment': att_file})
            
        if 'dividend' in desc.lower() or 'dividend' in att_text.lower() or 'split' in desc.lower() or 'bonus' in desc.lower():
            ca_rows.append({'exDate': dt, 'recDate': dt, 'subject': desc or att_text, 'faceVal': '1'})
            
    pd.DataFrame(ann_rows).to_csv(stock_folder / 'announcements.csv', index=False)
    pd.DataFrame(bm_rows).to_csv(stock_folder / 'board_meetings.csv', index=False)
    pd.DataFrame(ca_rows).to_csv(stock_folder / 'corporate_actions.csv', index=False)
    
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
    pd.DataFrame(fr_rows).to_csv(stock_folder / 'financial_results.csv', index=False)
    
    return sym, len(ann_rows), len(fr_rows), False

print("Starting live fetch from NSE Official Site...")

success_count = 0
for idx, sym in enumerate(SYMBOLS, 1):
    try:
        s, c_ann, c_fr, skipped = process_stock(sym)
        status_str = "ALREADY DONE" if skipped else "DOWNLOADED"
        print(f"[{idx:3d}/{len(SYMBOLS)}] {sym:14s} ({status_str:12s}) -> Ann: {c_ann:4d} | Results: {c_fr:2d}", flush=True)
        success_count += 1
    except Exception as e:
        print(f"[{idx:3d}/{len(SYMBOLS)}] {sym:14s} -> ERROR: {e}", flush=True)
    time.sleep(0.5)

print("==========================================================================================")
print(f"COMPLETED NSE EVENT FEEDS DOWNLOAD FOR ALL {success_count}/{len(SYMBOLS)} STOCKS!")
print("==========================================================================================")
