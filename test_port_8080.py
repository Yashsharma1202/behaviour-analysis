import urllib.request, sys

sys.stdout.reconfigure(errors='replace')

url = "http://localhost:8080"

print(f"Testing HTTP GET request to {url}...")

try:
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, timeout=5) as response:
        code = response.getcode()
        html = response.read().decode('utf-8', errors='ignore')
        print(f"✅ Success! HTTP Status Code: {code}")
        print(f"   Page Title / Sample HTML (First 200 chars):")
        print(f"   {html[:200]}")
except Exception as e:
    print(f"❌ Error connecting to {url}: {e}")
