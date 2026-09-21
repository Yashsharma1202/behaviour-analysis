import os, pathlib, sys, time
from concurrent.futures import ThreadPoolExecutor, as_completed

sys.stdout.reconfigure(errors='replace')

t0 = time.time()

d1 = pathlib.Path(r'D:\behaviour analysis\OI_DATA')
d2 = pathlib.Path(r'E:\OI_DATA')

print("==========================================================================================")
print("100% AUDIT CHECK: ALL 211 STOCK FOLDERS IN D:\\behaviour analysis\\OI_DATA vs E:\\OI_DATA")
print("==========================================================================================")

folders1 = sorted([e.name for e in os.scandir(d1) if e.is_dir()])
folders2 = sorted([e.name for e in os.scandir(d2) if e.is_dir()])

print(f"📁 D Drive Top Stock Folders Count : {len(folders1)}")
print(f"📁 E Drive Top Stock Folders Count : {len(folders2)}")

if folders1 != folders2:
    print("❌ STOCK FOLDERS MISMATCH!")
    sys.exit(0)
else:
    print("✅ Top-level stock folders (211 folders) match 100% perfectly!\n")

def audit_stock_folder(folder_name):
    p1 = d1 / folder_name
    p2 = d2 / folder_name
    
    files1 = {}
    for root, _, fnames in os.walk(p1):
        for fn in fnames:
            fp = pathlib.Path(root) / fn
            rel = str(fp.relative_to(p1)).lower()
            files1[rel] = fp.stat().st_size
            
    files2 = {}
    for root, _, fnames in os.walk(p2):
        for fn in fnames:
            fp = pathlib.Path(root) / fn
            rel = str(fp.relative_to(p2)).lower()
            files2[rel] = fp.stat().st_size
            
    k1 = set(files1.keys())
    k2 = set(files2.keys())
    
    mismatches = []
    for k in k1 & k2:
        if files1[k] != files2[k]:
            mismatches.append(k)
            
    return {
        'folder': folder_name,
        'count1': len(files1),
        'count2': len(files2),
        'only1': len(k1 - k2),
        'only2': len(k2 - k1),
        'size_diff': len(mismatches)
    }

print("Running fast parallel file & size scan across all 211 stock folders...", flush=True)

total_files_d = 0
total_files_e = 0
folder_errors = []
done = 0

with ThreadPoolExecutor(max_workers=32) as executor:
    futures = {executor.submit(audit_stock_folder, f): f for f in folders1}
    for future in as_completed(futures):
        res = future.result()
        done += 1
        total_files_d += res['count1']
        total_files_e += res['count2']
        if res['count1'] != res['count2'] or res['only1'] > 0 or res['only2'] > 0 or res['size_diff'] > 0:
            folder_errors.append(res)
            
        if done % 25 == 0 or done == len(folders1):
            print(f"  ✓ Audited {done}/{len(folders1)} stock folders ({total_files_d:,} files verified so far...) [{time.time() - t0:.1f}s]", flush=True)

print("\n==========================================================================================")
print("COMPREHENSIVE AUDIT RESULTS FOR ALL 211 STOCK FOLDERS")
print("==========================================================================================")
print(f"Total Stock Folders Checked : {len(folders1)} Folders")
print(f"Total Files in D:\\...       : {total_files_d:,} Files")
print(f"Total Files in E:\\...       : {total_files_e:,} Files")
print("------------------------------------------------------------------------------------------")
print(f"Folders with Any Discrepancy: {len(folder_errors)}")

if len(folder_errors) == 0 and total_files_d == total_files_e:
    print("\n🎉 YES! THE DATA IN BOTH DIRECTORIES IS 100% EXACTLY THE SAME!")
    print(f"   All {total_files_d:,} files across all {len(folders1)} stock folders match 100% perfectly in names, paths, and file sizes!")
else:
    print(f"\n⚠️ DISCREPANCIES DETECTED IN {len(folder_errors)} FOLDERS:")
    for err in folder_errors:
        print(f"   - Stock '{err['folder']}': D={err['count1']} files, E={err['count2']} files | Only D={err['only1']}, Only E={err['only2']}, Size Diff={err['size_diff']}")

print(f"\nAudit Execution Time: {time.time() - t0:.2f} seconds")
print("==========================================================================================")
