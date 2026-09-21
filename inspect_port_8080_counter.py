import urllib.request, sys, re

sys.stdout.reconfigure(errors='replace')

url = "http://localhost:8080"

req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
with urllib.request.urlopen(req, timeout=5) as response:
    html = response.read().decode('utf-8', errors='ignore')

print("==========================================================================================")
print("INSPECTING PORT 8080 SERVED HTML")
print("==========================================================================================")

# Search for total stock count / SYMBOLS definition in JS
symbols_matches = re.findall(r'SYMBOLS\s*=\s*(\[.*?\]);', html, re.DOTALL)
if symbols_matches:
    print(f"Found SYMBOLS JS definitions: {len(symbols_matches)}")
    for sm in symbols_matches:
        # count items
        items = [x.strip(' "\' \t\r\n') for x in sm.strip('[]').split(',') if x.strip()]
        print(f"  -> Discovered {len(items)} Symbols! Sample: {items[:5]} ... {items[-5:]}")

# Search for 1 / 50 or 1 / 211 in HTML
counter_match = re.search(r'\d+\s*/\s*\d+', html)
if counter_match:
    print(f"\nCounter string found in HTML: {counter_match.group(0)}")

print("==========================================================================================")
