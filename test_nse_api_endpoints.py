import urllib.request, http.cookiejar, json, sys

sys.stdout.reconfigure(errors='replace')

cj = http.cookiejar.CookieJar()
opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': '*/*',
    'Accept-Language': 'en-US,en;q=0.9',
    'Referer': 'https://www.nseindia.com/get-quotes/equity?symbol=ASHOKLEY'
}

print("==========================================================================================")
print("TESTING LIVE NSE OFFICIAL API ENDPOINTS FOR ASHOKLEY")
print("==========================================================================================")

# Prime session on get-quotes page
try:
    req = urllib.request.Request("https://www.nseindia.com/get-quotes/equity?symbol=ASHOKLEY", headers=headers)
    with opener.open(req, timeout=15) as resp:
        resp.read()
    print("✅ NSE Session Primed Successfully!")
except Exception as e:
    print(f"❌ Prime error: {e}")

# Endpoints to test
endpoints = [
    ("Corporate Announcements", "https://www.nseindia.com/api/corporate-announcements?index=equities&symbol=ASHOKLEY"),
    ("Corporate Actions", "https://www.nseindia.com/api/corporate-actions?index=equities&symbol=ASHOKLEY"),
    ("Board Meetings", "https://www.nseindia.com/api/corporate-boardmeetings?index=equities&symbol=ASHOKLEY"),
    ("Financial Results", "https://www.nseindia.com/api/corporate-financial-results?index=equities&symbol=ASHOKLEY"),
    ("Equity Corp Info", "https://www.nseindia.com/api/quote-equity?symbol=ASHOKLEY&section=corp_info")
]

for name, url in endpoints:
    try:
        req = urllib.request.Request(url, headers=headers)
        with opener.open(req, timeout=15) as resp:
            raw = resp.read().decode('utf-8', errors='ignore')
            data = json.loads(raw)
            count = len(data) if isinstance(data, list) else (len(data.get('data', [])) if isinstance(data, dict) else 0)
            print(f"  • {name:25s}: {count:4d} items returned from NSE!")
    except Exception as e:
        print(f"  • {name:25s}: ERROR ({e})")

print("==========================================================================================")
