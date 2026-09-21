import urllib.request, http.cookiejar, json, pandas as pd, pathlib, sys

sys.stdout.reconfigure(errors='replace')

cj = http.cookiejar.CookieJar()
opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': '*/*',
    'Accept-Language': 'en-US,en;q=0.9',
    'Referer': 'https://www.nseindia.com/get-quotes/equity?symbol=ASHOKLEY'
}

req = urllib.request.Request("https://www.nseindia.com/get-quotes/equity?symbol=ASHOKLEY", headers=headers)
with opener.open(req, timeout=15) as resp:
    resp.read()

# Fetch Announcements
url_ann = "https://www.nseindia.com/api/corporate-announcements?index=equities&symbol=ASHOKLEY"
req = urllib.request.Request(url_ann, headers=headers)
with opener.open(req, timeout=15) as resp:
    ann_raw = json.loads(resp.read().decode('utf-8', errors='ignore'))

# Fetch Results
url_res = "https://www.nseindia.com/api/results-comparision?symbol=ASHOKLEY"
req = urllib.request.Request(url_res, headers=headers)
with opener.open(req, timeout=15) as resp:
    res_raw = json.loads(resp.read().decode('utf-8', errors='ignore')).get('resCmpData', [])

stock_dir = pathlib.Path('D:/behaviour analysis/ASHOKLEY')
stock_dir.mkdir(parents=True, exist_ok=True)

# 1. Announcements CSV
ann_rows = []
bm_rows = []
ca_rows = []

for item in ann_raw:
    desc = item.get('desc', '')
    att_text = item.get('attchmntText', '')
    att_file = item.get('attchmntFile', '')
    dt = item.get('an_dt', '') or item.get('sort_date', '')
    
    ann_rows.append({
        'sort_date': dt,
        'desc': desc,
        'attchmntText': att_text,
        'attchmntFile': att_file
    })
    
    # Check if Board Meeting / Result Intimation
    if 'board meeting' in desc.lower() or 'financial result' in desc.lower() or 'board meeting' in att_text.lower():
        bm_rows.append({
            'bm_date': dt,
            'bm_purpose': desc,
            'bm_desc': att_text,
            'attachment': att_file
        })
        
    # Check if Dividend / Split / Action
    if 'dividend' in desc.lower() or 'dividend' in att_text.lower() or 'split' in desc.lower() or 'bonus' in desc.lower():
        ca_rows.append({
            'exDate': dt,
            'recDate': dt,
            'subject': desc or att_text,
            'faceVal': '1'
        })

df_ann = pd.DataFrame(ann_rows)
df_ann.to_csv(stock_dir / 'announcements.csv', index=False)

df_bm = pd.DataFrame(bm_rows)
df_bm.to_csv(stock_dir / 'board_meetings.csv', index=False)

df_ca = pd.DataFrame(ca_rows)
df_ca.to_csv(stock_dir / 'corporate_actions.csv', index=False)

# 2. Financial Results CSV
fr_rows = []
for r in res_raw:
    fr_rows.append({
        'broadCastDate': r.get('re_create_dt', ''),
        'relatingTo': f"Qtr ended {r.get('re_to_dt', '')}",
        'consolidated': r.get('re_res_type', 'Audited'),
        'audited': 'Audited',
        'fromDate': r.get('re_from_dt', ''),
        'toDate': r.get('re_to_dt', ''),
        'xbrl': ''
    })

df_fr = pd.DataFrame(fr_rows)
df_fr.to_csv(stock_dir / 'financial_results.csv', index=False)

print("==========================================================================================")
print(f"GENERATED EVENT FEEDS FOR ASHOKLEY:")
print(f"  • Announcements   : {len(df_ann)} rows saved")
print(f"  • Board Meetings  : {len(df_bm)} rows saved")
print(f"  • Corporate Actions: {len(df_ca)} rows saved")
print(f"  • Financial Results: {len(df_fr)} rows saved")
print("==========================================================================================")
