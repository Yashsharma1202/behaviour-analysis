import os, pathlib, sys, time, hashlib

sys.stdout.reconfigure(errors='replace')
t0 = time.time()

dir1 = pathlib.Path(r'D:\behaviour analysis\OI_DATA')
dir2 = pathlib.Path(r'E:\OI_DATA')

print("==========================================================================================")
print("FAST DIRECTORY & FILE COMPARISON TOOL")
print("==========================================================================================")
print(f"Directory 1: {dir1}")
print(f"Directory 2: {dir2}")
print("------------------------------------------------------------------------------------------", flush=True)

if not dir1.exists():
    print(f"❌ Path D:\\behaviour analysis\\OI_DATA does NOT exist!")
if not dir2.exists():
    print(f"❌ Path E:\\OI_DATA does NOT exist!")

if not dir1.exists() or not dir2.exists():
    sys.exit(0)

print("Scanning Directory 1 (D:\\behaviour analysis\\OI_DATA)...", flush=True)
files1 = {}
for root, dirs, filenames in os.walk(dir1):
    for f in filenames:
        fp = pathlib.Path(root) / f
        rel = str(fp.relative_to(dir1)).lower()
        files1[rel] = (str(fp.relative_to(dir1)), fp.stat().st_size, fp)

print(f"  ✓ Found {len(files1):,} files in D:\\behaviour analysis\\OI_DATA")

print("\nScanning Directory 2 (E:\\OI_DATA)...", flush=True)
files2 = {}
for root, dirs, filenames in os.walk(dir2):
    for f in filenames:
        fp = pathlib.Path(root) / f
        rel = str(fp.relative_to(dir2)).lower()
        files2[rel] = (str(fp.relative_to(dir2)), fp.stat().st_size, fp)

print(f"  ✓ Found {len(files2):,} files in E:\\OI_DATA")

keys1 = set(files1.keys())
keys2 = set(files2.keys())

only_in_1 = keys1 - keys2
only_in_2 = keys2 - keys1
common    = keys1 & keys2

print("\n------------------------------------------------------------------------------------------")
print("SUMMARY OF FILE MANIFEST COMPARISON:")
print("------------------------------------------------------------------------------------------")
print(f"Total files in D:\\... : {len(files1):,}")
print(f"Total files in E:\\... : {len(files2):,}")
print(f"Common files in both  : {len(common):,}")
print(f"Files ONLY in D:\\...  : {len(only_in_1):,}")
print(f"Files ONLY in E:\\...  : {len(only_in_2):,}")

if only_in_1:
    print(f"\n⚠️ Sample files ONLY in D:\\behaviour analysis\\OI_DATA (Showing up to 10):")
    for k in list(only_in_1)[:10]:
        print(f"   - {files1[k][0]} ({files1[k][1]:,} bytes)")

if only_in_2:
    print(f"\n⚠️ Sample files ONLY in E:\\OI_DATA (Showing up to 10):")
    for k in list(only_in_2)[:10]:
        print(f"   - {files2[k][0]} ({files2[k][1]:,} bytes)")

print("\n------------------------------------------------------------------------------------------")
print("COMPARING FILE SIZES & MD5 CHECKSUMS FOR COMMON FILES:")
print("------------------------------------------------------------------------------------------", flush=True)

size_mismatches = []
hash_mismatches = []
identical = []

def get_md5(path):
    h = hashlib.md5()
    with open(path, 'rb') as f:
        while chunk := f.read(4096 * 1024):
            h.update(chunk)
    return h.hexdigest()

for idx, k in enumerate(sorted(common), start=1):
    rel_name, sz1, path1 = files1[k]
    _, sz2, path2 = files2[k]
    
    if sz1 != sz2:
        size_mismatches.append((rel_name, sz1, sz2))
    else:
        h1 = get_md5(path1)
        h2 = get_md5(path2)
        if h1 != h2:
            hash_mismatches.append((rel_name, sz1, h1, h2))
        else:
            identical.append(rel_name)
            
    if idx % 100 == 0 or idx == len(common):
        print(f"  Processed {idx:,} / {len(common):,} common files... ({time.time() - t0:.1f}s)", flush=True)

print("\n==========================================================================================")
print("FINAL COMPARISON VERDICT")
print("==========================================================================================")
print(f"✅ Identical Files (100% Exact Content & Size) : {len(identical):,}")
print(f"📏 Size Mismatches                            : {len(size_mismatches):,}")
print(f"🧬 Content Hash Mismatches                     : {len(hash_mismatches):,}")
print(f"📁 Missing / Extra Files                      : {len(only_in_1) + len(only_in_2):,}")
print("------------------------------------------------------------------------------------------")

if len(identical) == len(files1) == len(files2) and len(only_in_1) == 0 and len(only_in_2) == 0:
    print("🎉 PERFECT MATCH! Both directories contain EXACTLY THE SAME DATA across all files!")
else:
    print("⚠️ DIFFERENCES DETECTED BETWEEN THE TWO DIRECTORIES!")

print(f"Total Time: {time.time() - t0:.2f} seconds")
print("==========================================================================================")
