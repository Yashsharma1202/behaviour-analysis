import urllib.request, sys

sys.stdout.reconfigure(errors='replace')

ip = "10.10.7.70"
urls = [
    ("Long Dashboard (Port 8080)", f"http://{ip}:8080"),
    ("Short Dashboard (Port 8081)", f"http://{ip}:8081")
]

print("==========================================================================================")
print("TESTING LAN IP HOSTING FOR LONG & SHORT DASHBOARDS")
print("==========================================================================================")

for name, url in urls:
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=5) as response:
            code = response.getcode()
            print(f"✅ {name:30s} -> {url:25s} | HTTP Status Code: {code}")
    except Exception as e:
        print(f"❌ {name:30s} -> {url:25s} | Error: {e}")

print("==========================================================================================")
