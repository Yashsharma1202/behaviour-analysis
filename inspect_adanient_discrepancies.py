import os
import pathlib
import sys

sys.stdout.reconfigure(errors='replace')

d1 = pathlib.Path(r'D:\behaviour analysis\OI_DATA\ADANIENT')
d2 = pathlib.Path(r'E:\OI_DATA\ADANIENT')

print("==========================================================================================")
print("INSPECTING EXACT FILE SIZE DISCREPANCIES IN 'ADANIENT' STOCK FOLDER")
print("==========================================================================================")
print(f"Path 1: {d1}")
print(f"Path 2: {d2}")
print("------------------------------------------------------------------------------------------")

files1 = {}
for root, _, fnames in os.walk(d1):
    for fn in fnames:
        fp = pathlib.Path(root) / fn
        rel = str(fp.relative_to(d1)).lower()
        files1[rel] = (str(fp.relative_to(d1)), fp.stat().st_size)

files2 = {}
for root, _, fnames in os.walk(d2):
    for fn in fnames:
        fp = pathlib.Path(root) / fn
        rel = str(fp.relative_to(d2)).lower()
        files2[rel] = (str(fp.relative_to(d2)), fp.stat().st_size)

mismatches = []
for k in sorted(files1.keys() & files2.keys()):
    rel1, s1 = files1[k]
    rel2, s2 = files2[k]
    if s1 != s2:
        mismatches.append((rel1, s1, s2, s2 - s1))

print(f"Total files in ADANIENT (D) : {len(files1):,}")
print(f"Total files in ADANIENT (E) : {len(files2):,}")
print(f"Total size mismatch files   : {len(mismatches):,}")
print("------------------------------------------------------------------------------------------")

print("\nSample 20 Size Mismatches in ADANIENT:")
for rel_path, s1, s2, diff in mismatches[:20]:
    print(f"  - {rel_path}: D={s1:,} bytes | E={s2:,} bytes | Diff: {diff:+,} bytes")

# Check which subdirectory category the mismatches belong to
cat_counts = {}
for rel_path, s1, s2, diff in mismatches:
    subcat = rel_path.split(os.sep)[0]
    cat_counts[subcat] = cat_counts.get(subcat, 0) + 1

print("\n------------------------------------------------------------------------------------------")
print("Breakdown of ADANIENT Mismatches by Category:")
for cat, count in cat_counts.items():
    print(f"  - {cat}: {count:,} mismatched files")

print("==========================================================================================")
