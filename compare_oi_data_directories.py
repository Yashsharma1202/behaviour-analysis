import os
import hashlib
import pathlib
import sys
import time

sys.stdout.reconfigure(errors='replace')

t0 = time.time()

dir1 = pathlib.Path(r'D:\behaviour analysis\OI_DATA')
dir2 = pathlib.Path(r'E:\OI_DATA')

print("==========================================================================================")
print("DIRECTORY DATA COMPARISON TOOL")
print("==========================================================================================")
print(f"Directory 1: {dir1}")
print(f"Directory 2: {dir2}")
print("------------------------------------------------------------------------------------------")

dir1_exists = dir1.exists()
dir2_exists = dir2.exists()

print(f"Directory 1 exists? {dir1_exists}")
print(f"Directory 2 exists? {dir2_exists}")

if not dir1_exists or not dir2_exists:
    print("\n❌ CANNOT COMPARE: One or both directories do not exist on the filesystem!")
    sys.exit(0)

def get_file_manifest(base_dir):
    manifest = {}
    for root, dirs, files in os.walk(base_dir):
        for f in files:
            full_path = pathlib.Path(root) / f
            rel_path = full_path.relative_to(base_dir)
            manifest[str(rel_path).lower()] = {
                'rel_path_orig': str(rel_path),
                'full_path': full_path,
                'size': full_path.stat().st_size
            }
    return manifest

manifest1 = get_file_manifest(dir1)
manifest2 = get_file_manifest(dir2)

keys1 = set(manifest1.keys())
keys2 = set(manifest2.keys())

only_in_1 = keys1 - keys2
only_in_2 = keys2 - keys1
common = keys1 & keys2

print(f"\n📁 Structure Summary:")
print(f"  Files in D:\\behaviour analysis\\OI_DATA: {len(manifest1)}")
print(f"  Files in E:\\OI_DATA                  : {len(manifest2)}")
print(f"  Common files in both                : {len(common)}")
print(f"  Files only in D:\\...                : {len(only_in_1)}")
print(f"  Files only in E:\\...                : {len(only_in_2)}")

if only_in_1:
    print(f"\n⚠️ Sample files ONLY in D:\\behaviour analysis\\OI_DATA (Total {len(only_in_1)}):")
    for k in list(only_in_1)[:10]:
        print(f"   - {manifest1[k]['rel_path_orig']} ({manifest1[k]['size']:,} bytes)")

if only_in_2:
    print(f"\n⚠️ Sample files ONLY in E:\\OI_DATA (Total {len(only_in_2)}):")
    for k in list(only_in_2)[:10]:
        print(f"   - {manifest2[k]['rel_path_orig']} ({manifest2[k]['size']:,} bytes)")

def calc_file_hash(filepath):
    hasher = hashlib.md5()
    with open(filepath, 'rb') as f:
        while chunk := f.read(8192 * 1024):
            hasher.update(chunk)
    return hasher.hexdigest()

size_mismatches = []
content_mismatches = []
identical_files = []

print(f"\n🔍 Comparing {len(common)} common files byte-by-byte (MD5 hash check)...")

for idx, k in enumerate(sorted(common), start=1):
    f1 = manifest1[k]
    f2 = manifest2[k]
    
    if f1['size'] != f2['size']:
        size_mismatches.append((f1['rel_path_orig'], f1['size'], f2['size']))
    else:
        # Check MD5 hash
        h1 = calc_file_hash(f1['full_path'])
        h2 = calc_file_hash(f2['full_path'])
        if h1 != h2:
            content_mismatches.append((f1['rel_path_orig'], h1, h2))
        else:
            identical_files.append(f1['rel_path_orig'])

print("\n==========================================================================================")
print("COMPARISON RESULTS VERDICT")
print("==========================================================================================")
print(f"✅ Identical Files (Exact same content & size) : {len(identical_files)}")
print(f"📏 Size Mismatch Files                           : {len(size_mismatches)}")
print(f"🧬 Content Hash Mismatch Files (Same size, diff): {len(content_mismatches)}")
print(f"📁 Missing / Extra Files                         : {len(only_in_1) + len(only_in_2)}")
print("------------------------------------------------------------------------------------------")

if len(identical_files) == len(manifest1) == len(manifest2) and len(only_in_1) == 0 and len(only_in_2) == 0:
    print("🎉 PERFECT MATCH! Both directories contain EXACTLY THE SAME DATA across all files!")
else:
    print("⚠️ DIFFERENCES DETECTED!")
    
    if size_mismatches:
        print("\n📏 Sample Size Mismatches:")
        for rel_p, s1, s2 in size_mismatches[:10]:
            print(f"   - {rel_p}: D={s1:,} bytes vs E={s2:,} bytes (Diff: {s2-s1:+,} bytes)")
            
    if content_mismatches:
        print("\n🧬 Sample Content Mismatches (Same size but different data):")
        for rel_p, h1, h2 in content_mismatches[:10]:
            print(f"   - {rel_p}: Hash D={h1[:8]} vs Hash E={h2[:8]}")

print(f"\nExecution Time: {time.time() - t0:.2f} seconds")
print("==========================================================================================")
