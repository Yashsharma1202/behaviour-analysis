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

req = urllib.request.Request("https://www.nseindia.com/get-quotes/equity?symbol=ASHOKLEY", headers=headers)
with opener.open(req, timeout=15) as resp:
    resp.read()

endpoints = [
    ("Actions v2", "https://www.nseindia.com/api/corporate-actions?symbol=ASHOKLEY"),
    ("BoardMeetings v2", "https://www.nseindia.com/api/corporate-boardmeetings?symbol=ASHOKLEY"),
    ("Corp Info Data", "https://www.nseindia.com/api/quote-equity?symbol=ASHOKLEY&section=corp_info"),
    ("Results Comparision", "https://www.nseindia.com/api/results-comparision?symbol=ASHOKLEY"),
    ("Financial Results v2", "https://www.nseindia.com/api/financial-results?symbol=ASHOKLEY")
]

print("==========================================================================================")
print("TESTING ALTERNATIVE LIVE NSE API ENDPOINTS")
print("==========================================================================================")

for name, url in endpoints:
    try:
        req = urllib.request.Request(url, headers=headers)
        with opener.open(req, timeout=15) as resp:
            raw = resp.read().decode('utf-8', errors='ignore')
            data = json.loads(raw)
            print(f"  ✅ {name:22s}: Success! Data keys: {list(data.keys()) if isinstance(data, dict) else len(data)}")
    except Exception as e:
        print(f"  ❌ {name:22s}: {e}")

print("==========================================================================================")
