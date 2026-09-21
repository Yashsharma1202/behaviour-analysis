import sys
import stock_server
import short_stock_server

sys.stdout.reconfigure(errors='replace')

syms1 = stock_server.discover_symbols()
syms2 = short_stock_server.discover_symbols()

print("==========================================================================================")
print("TESTING discover_symbols() IN BOTH SERVERS")
print("==========================================================================================")
print(f"stock_server.py discover_symbols()       : {len(syms1)} STOCKS")
print(f"short_stock_server.py discover_symbols() : {len(syms2)} STOCKS")
print(f"Sample Symbols (211 total): {syms1[:10]} ... {syms1[-10:]}")
print("==========================================================================================")
