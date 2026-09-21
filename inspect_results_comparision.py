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

url = "https://www.nseindia.com/api/results-comparision?symbol=ASHOKLEY"
req = urllib.request.Request(url, headers=headers)
with opener.open(req, timeout=15) as resp:
    raw = resp.read().decode('utf-8', errors='ignore')
    data = json.loads(raw)

print("==========================================================================================")
print("INSPECTING NSE OFFICIAL RESULTS COMPARISON DATA FOR ASHOKLEY")
print("==========================================================================================")
res_data = data.get('resCmpData', [])
print(f"Total Quarterly Result Records Returned: {len(res_data)}")
if res_data:
    print(f"Sample Record 0: {res_data[0]}")
print("==========================================================================================")
