import urllib.request, sys, re

sys.stdout.reconfigure(errors='replace')

url = "http://localhost:8080"

print(f"Testing HTTP GET request to {url}...")

try:
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, timeout=5) as response:
        code = response.getcode()
        html = response.read().decode('utf-8', errors='ignore')
        print(f"✅ Success! HTTP Status Code: {code}")
        
        # Check stock count references in HTML
        matches = re.findall(r'(\d+)\s*stocks', html, re.IGNORECASE)
        print(f"   Stock Count references found in HTML: {matches}")
        
        # Find SYMBOLS array length or stock counter
        symbols_match = re.search(r'const SYMBOLS = (\[.*?\]);', html, re.DOTALL)
        if symbols_match:
            try:
                syms_arr = eval(symbols_match.group(1))
                print(f"   SYMBOLS Array Length in HTML: {len(syms_arr)} STOCKS!")
                print(f"   First 5 Symbols: {syms_arr[:5]}")
                print(f"   Last 5 Symbols : {syms_arr[-5:]}")
            except Exception as e:
                print(f"   Eval error: {e}")
        else:
            print("   SYMBOLS array regex not matched directly.")
except Exception as e:
    print(f"❌ Error connecting to {url}: {e}")
