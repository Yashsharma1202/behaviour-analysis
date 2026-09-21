import os, pathlib, sys, time

sys.stdout.reconfigure(errors='replace')
t0 = time.time()

d1 = pathlib.Path(r'D:\behaviour analysis\OI_DATA')
d2 = pathlib.Path(r'E:\OI_DATA')

print("==========================================================================================")
print("FAST SAMPLED COMPARISON REPORT — D:\\behaviour analysis\\OI_DATA vs E:\\OI_DATA")
print("==========================================================================================")

items1 = sorted([e.name for e in os.scandir(d1) if e.is_dir()])
items2 = sorted([e.name for e in os.scandir(d2) if e.is_dir()])

print(f"📁 D Drive Top Stock Folders Count : {len(items1)}")
print(f"📁 E Drive Top Stock Folders Count : {len(items2)}")
print(f"✅ Top Stock Folders Match         : {'YES (100% Match)' if items1 == items2 else 'NO'}")

# Sample 30 representative stock folders across the alphabet
sampled_folders = items1[::7] # Every 7th stock folder (31 stock folders sampled)

mismatches = []
total_files_sampled_d = 0
total_files_sampled_e = 0

for s_folder in sampled_folders:
    p1 = d1 / s_folder
    p2 = d2 / s_folder
    
    files1 = {e.name: e.stat().st_size for e in os.scandir(p1) if e.is_file()}
    files2 = {e.name: e.stat().st_size for e in os.scandir(p2) if e.is_file()}
    
    total_files_sampled_d += len(files1)
    total_files_sampled_e += len(files2)
    
    if files1 != files2:
        mismatches.append(s_folder)

print("\n------------------------------------------------------------------------------------------")
print(f"🔍 SAMPLED 31 STOCK FOLDERS ({total_files_sampled_d:,} files in D vs {total_files_sampled_e:,} files in E):")
print("------------------------------------------------------------------------------------------")
print(f"Sampled Stock Folders: {sampled_folders[:5]} ... {sampled_folders[-5:]}")
print(f"Folder Mismatches Found in Sample : {len(mismatches)}")

if len(mismatches) == 0 and total_files_sampled_d == total_files_sampled_e:
    print("🎉 ALL SAMPLED FOLDERS MATCH 100% PERFECTLY IN FILE NAMES, COUNTS & SIZES!")
else:
    print(f"⚠️ Mismatches found in folders: {mismatches}")

print(f"\nExecution Time: {time.time() - t0:.2f} seconds")
print("==========================================================================================")
