import os, pathlib, sys

sys.stdout.reconfigure(errors='replace')

d1 = pathlib.Path(r'D:\behaviour analysis\OI_DATA')
d2 = pathlib.Path(r'E:\OI_DATA')

print("==========================================================================================")
print("TOP-LEVEL DIRECTORY STRUCTURE INSPECTION")
print("==========================================================================================")

print(f"Top-level contents of D:\\behaviour analysis\\OI_DATA:")
if d1.exists():
    items1 = sorted([e.name for e in os.scandir(d1)])
    print(f"  Total items: {len(items1)}")
    for item in items1[:20]:
        print(f"   - {item}")
else:
    print("  DOES NOT EXIST!")

print(f"\nTop-level contents of E:\\OI_DATA:")
if d2.exists():
    items2 = sorted([e.name for e in os.scandir(d2)])
    print(f"  Total items: {len(items2)}")
    for item in items2[:20]:
        print(f"   - {item}")
else:
    print("  DOES NOT EXIST!")

print("==========================================================================================")
