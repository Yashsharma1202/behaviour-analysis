import urllib.request
import urllib.parse
import http.cookiejar
import json
import os
import pandas as pd
import pathlib
import sys
import warnings

warnings.filterwarnings('ignore')
sys.stdout.reconfigure(errors='replace')

ROOT = pathlib.Path('D:/behaviour analysis')
OI_DIR = ROOT / 'OI_DATA'
PROC = ROOT / 'processed'
PROC.mkdir(parents=True, exist_ok=True)

SYMBOLS = set([e.name.strip().upper() for e in os.scandir(OI_DIR) if e.is_dir()]) if (OI_DIR.exists()) else set()

print("==========================================================================================")
print("NSE OFFICIAL EVENT CALENDAR: UPCOMING BOARD MEETINGS & QUARTERLY RESULT DATES")
print("==========================================================================================")
print(f"Current System Time: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")
print("------------------------------------------------------------------------------------------")

cj = http.cookiejar.CookieJar()
opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': '*/*',
    'Accept-Language': 'en-US,en;q=0.9',
    'Referer': 'https://www.nseindia.com/market-data/event-calendar'
}

# Prime session with NSE homepage
try:
    req = urllib.request.Request('https://www.nseindia.com/', headers=headers)
    with opener.open(req, timeout=12) as r:
        r.read()
except Exception as e:
    print(f"Warning priming session: {e}")

# 1. Fetch NSE Live Event Calendar API
url_evt = 'https://www.nseindia.com/api/event-calendar'
events = []
try:
    with opener.open(urllib.request.Request(url_evt, headers=headers), timeout=15) as r:
        events = json.loads(r.read().decode('utf-8'))
    print(f"Total Upcoming Board Meetings on NSE Official Event Calendar: {len(events)}")
except Exception as e:
    print(f"Error fetching NSE event calendar: {e}")

# 2. Check for quote-equity corporate info for target stocks or general board meetings
quote_bm_list = []

# Filter events for financial results / board meetings
parsed_events = []
for e in events:
    if not isinstance(e, dict): continue
    sym = (e.get('symbol') or '').upper().strip()
    company = e.get('company', '')
    dt_str = e.get('date', '')
    purpose = e.get('purpose', '')
    bm_desc = e.get('bm_desc', '')
    
    parsed_events.append({
        'symbol': sym,
        'company': company,
        'upcoming_board_meeting_date': dt_str,
        'purpose': purpose,
        'details': bm_desc,
        'is_in_211_universe': 'YES ✅' if sym in SYMBOLS else 'NO'
    })

df_events = pd.DataFrame(parsed_events)
if not df_events.empty:
    print("\n------------------------------------------------------------------------------------------")
    print("UPCOMING BOARD MEETINGS FROM NSE OFFICIAL EVENT CALENDAR:")
    print("------------------------------------------------------------------------------------------")
    print(df_events.to_string(index=False))
    
    out_csv = PROC / 'nse_upcoming_event_calendar.csv'
    df_events.to_csv(out_csv, index=False)
    print(f"\nSaved upcoming event calendar to: {out_csv}")
