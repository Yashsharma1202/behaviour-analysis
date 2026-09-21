import urllib.request, json, sys

sys.stdout.reconfigure(errors='replace')

url = "http://localhost:9000/api/equity/RELIANCE"

print(f"Testing Equity Spot Parquet API Endpoint {url}...")

try:
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, timeout=5) as response:
        code = response.getcode()
        payload = json.loads(response.read().decode('utf-8'))
        print(f"✅ Success! HTTP Status Code: {code}")
        print(f"   Symbol: {payload['symbol']}")
        print(f"   Selected Date: {payload['selected_date']}")
        print(f"   Total Trading Days Available: {payload['total_dates']}")
        print(f"   1-Min Bars Loaded: {len(payload['bars'])} bars")
        if payload['bars']:
            print(f"   First Bar: {payload['bars'][0]}")
            print(f"   Last Bar : {payload['bars'][-1]}")
except Exception as e:
    print(f"❌ Error connecting to API: {e}")
