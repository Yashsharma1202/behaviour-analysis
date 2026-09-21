import os, pathlib, sys

sys.stdout.reconfigure(errors='replace')

oi_dir = pathlib.Path(r'D:\behaviour analysis\OI_DATA')
symbols = sorted([e.name.strip().upper() for e in os.scandir(oi_dir) if e.is_dir()])

print("==========================================================================================")
print("EXACT 211 STOCK SYMBOLS FROM OI_DATA")
print("==========================================================================================")
print(f"Total Stock Symbols in OI_DATA: {len(symbols)} Stocks")
print(f"First 10: {symbols[:10]}")
print(f"Last 10 : {symbols[-10:]}")
print("==========================================================================================")
