import os, pathlib, sys

sys.stdout.reconfigure(errors='replace')

b_dir = pathlib.Path('D:/behaviour analysis')

print("==========================================================================================")
print("FINDING ALL WEB SERVER FILES & SYMBOL DISCOVERY LOCATIONS IN REPOSITORY")
print("==========================================================================================")

server_files = []

for root, dirs, files in os.walk(b_dir):
    if '.git' in root or '__pycache__' in root:
        continue
    for f in files:
        if f.endswith('.py'):
            fp = pathlib.Path(root) / f
            try:
                txt = fp.read_text(encoding='utf-8', errors='ignore')
                if 'HTTPServer' in txt or 'BaseHTTPRequestHandler' in txt or 'discover_symbols' in txt or '50' in txt:
                    server_files.append((str(fp.relative_to(b_dir)), len(txt), 'HTTPServer' in txt))
            except Exception as e:
                pass

print(f"Found {len(server_files)} server/symbol candidate python files:\n")
for f_path, size, is_http in sorted(server_files):
    print(f"  - {f_path:45s} (Size: {size:,} bytes | Is HTTP Server? {is_http})")

print("==========================================================================================")
