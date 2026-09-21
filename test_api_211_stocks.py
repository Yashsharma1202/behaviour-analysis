import urllib.request, json, sys

sys.stdout.reconfigure(errors='replace')

url = "http://localhost:9000/api/stocks"

print(f"Testing API Endpoint {url}...")

try:
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, timeout=5) as response:
        code = response.getcode()
        payload = json.loads(response.read().decode('utf-8'))
        print(f"✅ Success! HTTP Status Code: {code}")
        print(f"   Total Symbols Returned: {payload['total']}")
        print(f"   Sample Symbols (First 20): {payload['symbols'][:20]}")
except Exception as e:
    print(f"❌ Error connecting to API: {e}")
