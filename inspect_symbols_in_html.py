import urllib.request, sys, re, json

sys.stdout.reconfigure(errors='replace')

url = "http://localhost:8080"

req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
with urllib.request.urlopen(req, timeout=5) as response:
    html = response.read().decode('utf-8', errors='ignore')

print("==========================================================================================")
print("INSPECTING SYMBOLS INJECTED IN HTML SERVED ON PORT 8080")
print("==========================================================================================")

# Search for JSON arrays or stock list in JS
matches = re.findall(r'const\s+SYMBOLS\s*=\s*(\[.*?\]);', html)
if not matches:
    matches = re.findall(r'SYMBOLS\s*:\s*(\[.*?\])', html)
if not matches:
    matches = re.findall(r'(\["360ONE".*?\])', html)

if matches:
    try:
        sym_list = json.loads(matches[0])
        print(f"✅ Discovered {len(sym_list)} Stocks Injected in Client HTML!")
        print(f"   First 10: {sym_list[:10]}")
        print(f"   Last 10 : {sym_list[-10:]}")
    except Exception as e:
        print(f"   JSON parse error: {e}")
else:
    # Print lines containing 360ONE or RELIANCE
    for l in html.splitlines():
        if '360ONE' in l or 'ADANIENT' in l:
            if len(l) < 300:
                print(f"  Line: {l}")
            else:
                print(f"  Line snippet: {l[:200]}...")

print("==========================================================================================")
