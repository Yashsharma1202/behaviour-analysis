import pathlib, sys

sys.stdout.reconfigure(errors='replace')

p1 = pathlib.Path('D:/behaviour analysis/stock_server.py')
txt = p1.read_text(encoding='utf-8', errors='ignore')
lines = txt.splitlines()

print("==========================================================================================")
print("INSPECTING SYMBOLS INITIALIZATION IN stock_server.py")
print("==========================================================================================")

for idx, line in enumerate(lines, start=1):
    if 'SYMBOLS =' in line or 'SYMBOLS:' in line or 'discover_symbols' in line:
        print(f"L{idx:4d}: {line}")

print("\nLines 2850-2886:")
for idx in range(2849, len(lines)):
    print(f"L{idx+1:4d}: {lines[idx]}")
print("==========================================================================================")
