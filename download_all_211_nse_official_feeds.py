import urllib.request
import urllib.parse
import http.cookiejar
import json
import csv
import os
import pathlib
import sys
import time
from concurrent.futures import ThreadPoolExecutor

sys.stdout.reconfigure(errors='replace')

ROOT = pathlib.Path('D:/behaviour analysis')
OI_DIR = ROOT / 'OI_DATA'

# 1. Discover all 211 stock symbols
SYMBOLS = sorted([e.name.strip().upper() for e in os.scandir(OI_DIR) if e.is_dir()])

print("==========================================================================================")
print(f"BULK DOWNLOADER FOR OFFICIAL NSE EVENT FEEDS (211 STOCKS)")
print("==========================================================================================")
print(f"Total Target Stocks: {len(SYMBOLS)} Stocks")
print(f"First 10: {SYMBOLS[:10]}")
print("------------------------------------------------------------------------------------------", flush=True)

# 2. Setup NSE Session Class with Cookie Handling & Priming
class NSESession:
    def __init__(self):
        self.cj = http.cookiejar.CookieJar()
        self.opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(self.cj))
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.9',
            'Referer': 'https://www.nseindia.com/'
        }
        self.primed = False

    def prime(self):
        try:
            req = urllib.request.Request("https://www.nseindia.com", headers=self.headers)
            with self.opener.open(req, timeout=15) as resp:
                resp.read()
            self.primed = True
        except Exception as e:
            print(f"NSE session prime error: {e}")

    def get_json(self, url):
        if not self.primed:
            self.prime()
        req_headers = self.headers.copy()
        req_headers['Accept'] = 'application/json, text/plain, */*'
        req_headers['X-Requested-With'] = 'XMLHttpRequest'
        req = urllib.request.Request(url, headers=req_headers)
        
        for attempt in range(2):
            try:
                with self.opener.open(req, timeout=15) as resp:
                    raw = resp.read().decode('utf-8', errors='ignore')
                    return json.loads(raw)
            except Exception as e:
                self.prime()
                time.sleep(1)
        return None

nse = NSESession()
nse.prime()

# 3. Download & Save Functions for 4 NSE Event Feeds
def save_feed_csv(folder_path, filename, data, default_cols):
    folder_path.mkdir(parents=True, exist_ok=True)
    file_path = folder_path / filename
    
    if not isinstance(data, list) or len(data) == 0:
        if not file_path.exists():
            with open(file_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(default_cols)
        return 0
        
    keys = list(data[0].keys())
    with open(file_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=keys)
        writer.writeheader()
        writer.writerows(data)
    return len(data)

def process_stock(sym):
    stock_folder = ROOT / sym
    stock_folder.mkdir(parents=True, exist_ok=True)
    
    encoded_sym = urllib.parse.quote(sym)
    
    # 1. Announcements
    ann_url = f"https://www.nseindia.com/api/corporate-announcements?index=equities&symbol={encoded_sym}"
    ann_data = nse.get_json(ann_url)
    c_ann = save_feed_csv(stock_folder, 'announcements.csv', ann_data, ['sort_date', 'desc', 'attchmntText', 'attchmntFile'])
    
    # 2. Board Meetings
    bm_url = f"https://www.nseindia.com/api/corporate-boardmeetings?index=equities&symbol={encoded_sym}"
    bm_data = nse.get_json(bm_url)
    c_bm = save_feed_csv(stock_folder, 'board_meetings.csv', bm_data, ['bm_date', 'bm_purpose', 'bm_desc', 'attachment'])
    
    # 3. Corporate Actions
    ca_url = f"https://www.nseindia.com/api/corporate-actions?index=equities&symbol={encoded_sym}"
    ca_data = nse.get_json(ca_url)
    c_ca = save_feed_csv(stock_folder, 'corporate_actions.csv', ca_data, ['exDate', 'recDate', 'subject', 'faceVal'])
    
    # 4. Financial Results
    fr_url = f"https://www.nseindia.com/api/corporate-financial-results?index=equities&symbol={encoded_sym}"
    fr_data = nse.get_json(fr_url)
    c_fr = save_feed_csv(stock_folder, 'financial_results.csv', fr_data, ['broadCastDate', 'relatingTo', 'consolidated', 'audited', 'fromDate', 'toDate', 'xbrl'])
    
    return sym, c_ann, c_bm, c_ca, c_fr

print("Starting live fetch from NSE Official Site...")

success_count = 0
for idx, sym in enumerate(SYMBOLS, 1):
    try:
        s, c_ann, c_bm, c_ca, c_fr = process_stock(sym)
        print(f"[{idx:3d}/{len(SYMBOLS)}] {sym:14s} -> Ann: {c_ann:3d} | BM: {c_bm:3d} | Actions: {c_ca:3d} | Results: {c_fr:3d}", flush=True)
        success_count += 1
    except Exception as e:
        print(f"[{idx:3d}/{len(SYMBOLS)}] {sym:14s} -> ERROR: {e}", flush=True)
    time.sleep(1.0) # Polite rate limit

print("==========================================================================================")
print(f"COMPLETED DOWNLOAD FOR {success_count}/{len(SYMBOLS)} STOCKS FROM NSE OFFICIAL SITE!")
print("==========================================================================================")
