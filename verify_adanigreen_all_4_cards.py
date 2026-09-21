import sys
import stock_server

sys.stdout.reconfigure(errors='replace')

events = stock_server.build_events('ADANIGREEN')

print("==========================================================================================")
print("VERIFYING ALL 4 EVENT FEEDS FOR ADANIGREEN IN stock_server.py")
print("==========================================================================================")
print(f"Symbol           : {events['symbol']}")
print(f"Available on Disk: {events['available']}")
for feed_name, feed_info in events['feeds'].items():
    print(f"  • Feed: {feed_info['label']:20s} | Total Records Saved: {feed_info['total']:4d}")
print("==========================================================================================")
