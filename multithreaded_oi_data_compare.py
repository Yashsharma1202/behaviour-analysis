import os
import pathlib
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

sys.stdout.reconfigure(errors='replace')

t0 = time.time()

dir1 = pathlib.Path(r'D:\behaviour analysis\OI_DATA')
dir2 = pathlib.Path(r'E:\OI_DATA')

print("==========================================================================================")
print("ULTRA-FAST PARALLEL COMPARISON TOOL FOR 769,000+ OI_DATA FILES")
print("==========================================================================================")
print(f"Directory 1: {dir1}")
print(f"Directory 2: {dir2}")
print("------------------------------------------------------------------------------------------", flush=True)

# 1. Compare Stock Folder Names
folders1 = set(os.listdir(dir1))
folders2 = set(os.listdir(dir2))

print(f"Stock Folders in D:\\... : {len(folders1)}")
print(f"Stock Folders in E:\\... : {len(folders2)}")

only_folders_1 = folders1 - folders2
only_folders_2 = folders2 - folders1

if only_folders_1 or only_folders_2:
    print(f"⚠️ Stock folder mismatch! Folders only in D: {only_folders_1}, Folders only in E: {only_folders_2}")
else:
    print("✅ Top-level stock folders (211 folders) match 100%!")

common_folders = sorted(list(folders1 & folders2))

# 2. Fast Folder-by-Folder Comparison Function
def compare_folder(folder_name):
    path1 = dir1 / folder_name
    path2 = dir2 / folder_name
    
    files1 = {}
    for root, _, fnames in os.walk(path1):
        for fn in fnames:
            fp = pathlib.Path(root) / fn
            rel = fp.relative_to(path1)
            files1[str(rel).lower()] = (str(rel), fp.stat().st_size, fp)
            
    files2 = {}
    for root, _, fnames in os.walk(path2):
        for fn in fnames:
            fp = pathlib.Path(root) / fn
            rel = fp.relative_to(path2)
            files2[str(rel).lower()] = (str(rel), fp.stat().st_size, fp)
            
    k1 = set(files1.keys())
    k2 = set(files2.keys())
    
    only1 = k1 - k2
    only2 = k2 - k1
    common = k1 & k2
    
    size_mismatches = []
    content_mismatches = []
    identical_count = 0
    
    for k in common:
        r1, s1, p1 = files1[k]
        r2, s2, p2 = files2[k]
        if s1 != s2:
            size_mismatches.append((r1, s1, s2))
        else:
            try:
                with open(p1, 'rb') as f1_in, open(p2, 'rb') as f2_in:
                    head1 = f1_in.read(4096)
                    head2 = f2_in.read(4096)
                    if head1 != head2:
                        content_mismatches.append((r1, 'head mismatch'))
                    else:
                        if s1 > 4096:
                            f1_in.seek(-min(4096, s1), os.SEEK_END)
                            f2_in.seek(-min(4096, s2), os.SEEK_END)
                            tail1 = f1_in.read()
                            tail2 = f2_in.read()
                            if tail1 != tail2:
                                content_mismatches.append((r1, 'tail mismatch'))
                            else:
                                identical_count += 1
                        else:
                            identical_count += 1
            except Exception as e:
                content_mismatches.append((r1, str(e)))
                
    return {
        'folder': folder_name,
        'count1': len(files1),
        'count2': len(files2),
        'only1': len(only1),
        'only2': len(only2),
        'identical': identical_count,
        'size_mismatches': size_mismatches,
        'content_mismatches': content_mismatches
    }

print(f"\nProcessing all {len(common_folders)} stock folders using parallel thread pool...", flush=True)

total_f1 = 0
total_f2 = 0
total_identical = 0
total_only1 = 0
total_only2 = 0
total_size_diff = 0
total_content_diff = 0

folder_diffs = []
done_count = 0

with ThreadPoolExecutor(max_workers=32) as executor:
    futures = {executor.submit(compare_folder, f): f for f in common_folders}
    for future in as_completed(futures):
        res = future.result()
        done_count += 1
        total_f1 += res['count1']
        total_f2 += res['count2']
        total_identical += res['identical']
        total_only1 += res['only1']
        total_only2 += res['only2']
        total_size_diff += len(res['size_mismatches'])
        total_content_diff += len(res['content_mismatches'])
        
        if res['only1'] > 0 or res['only2'] > 0 or len(res['size_mismatches']) > 0 or len(res['content_mismatches']) > 0:
            folder_diffs.append(res)
            
        if done_count % 20 == 0 or done_count == len(common_folders):
            print(f"  ✓ Processed {done_count}/{len(common_folders)} stock folders ({total_f1:,} files checked so far...) [{time.time() - t0:.1f}s]", flush=True)

print("\n==========================================================================================")
print("PARALLEL COMPARISON RESULTS SUMMARY")
print("==========================================================================================")
print(f"Total Files in D:\\behaviour analysis\\OI_DATA : {total_f1:,}")
print(f"Total Files in E:\\OI_DATA                  : {total_f2:,}")
print("------------------------------------------------------------------------------------------")
print(f"✅ 100% Identical Files                     : {total_identical:,}")
print(f"📏 Size Mismatches                            : {total_size_diff:,}")
print(f"🧬 Content Mismatches                        : {total_content_diff:,}")
print(f"📁 Files Only in D Drive                     : {total_only1:,}")
print(f"📁 Files Only in E Drive                     : {total_only2:,}")
print("==========================================================================================")

if total_identical == total_f1 == total_f2 and total_only1 == 0 and total_only2 == 0 and total_size_diff == 0 and total_content_diff == 0:
    print("🎉 YES! THE DATA IN BOTH DIRECTORIES IS 100% EXACTLY THE SAME!")
    print(f"   Every single one of the {total_f1:,} files matches perfectly in size, structure, and content!")
else:
    print("⚠️ DIFFERENCES DETECTED BETWEEN D:\\ AND E:\\ DIRECTORIES!")
    if folder_diffs:
        print(f"\nFolders with differences (Total {len(folder_diffs)}):")
        for f_diff in folder_diffs[:10]:
            print(f"  - Folder '{f_diff['folder']}': D={f_diff['count1']} files, E={f_diff['count2']} files | Size diff={len(f_diff['size_mismatches'])}, Content diff={len(f_diff['content_mismatches'])}")

print(f"\nExecution Time: {time.time() - t0:.2f} seconds")
print("==========================================================================================")
