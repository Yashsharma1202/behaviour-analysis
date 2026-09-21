import sys
import short_stock_server

sys.stdout.reconfigure(errors='replace')

syms = short_stock_server.discover_symbols()

print("==========================================================================================")
print(f"VERIFYING DISCOVERED SYMBOLS IN short_stock_server.py")
print("==========================================================================================")
print(f"Total Discovered Symbols: {len(syms)} Stocks")
print(f"First 20 Symbols        : {syms[:20]}")
print(f"Last 20 Symbols         : {syms[-20:]}")
print("==========================================================================================")
