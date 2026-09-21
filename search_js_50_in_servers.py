import re, pathlib, sys

sys.stdout.reconfigure(errors='replace')

p1 = pathlib.Path('D:/behaviour analysis/stock_server.py')
p2 = pathlib.Path('D:/behaviour analysis/short_stock_server.py')

def find_50_occurrences(path):
    print(f"\n==========================================================================================")
    print(f"SEARCHING FOR '50' IN {path.name}")
    print(f"==========================================================================================")
    txt = path.read_text(encoding='utf-8', errors='ignore')
    lines = txt.splitlines()
    for idx, line in enumerate(lines, start=1):
        if re.search(r'\b50\b|1\s*/\s*50|nifty50', line, re.IGNORECASE):
            print(f"  L{idx:4d}: {line.strip()[:120]}")

find_50_occurrences(p1)
find_50_occurrences(p2)
