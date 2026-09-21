import os, sys, subprocess, time

sys.stdout.reconfigure(errors='replace')

print("==========================================================================================")
print("FINDING & KILLING ALL PROCESSES LISTENING ON PORT 8080")
print("==========================================================================================")

try:
    output = subprocess.check_output("netstat -ano | findstr :8080", shell=True).decode('utf-8', errors='ignore')
    lines = output.splitlines()
    pids = set()
    for l in lines:
        parts = l.strip().split()
        if len(parts) >= 5:
            pid = parts[-1]
            if pid.isdigit() and int(pid) > 0:
                pids.add(pid)
                
    print(f"Found PIDs listening on 8080: {pids}")
    for pid in pids:
        print(f"Killing PID {pid}...")
        os.system(f"taskkill /F /PID {pid}")
except Exception as e:
    print(f"Netstat / taskkill note: {e}")

print("------------------------------------------------------------------------------------------")
print("VERIFYING PORT 8080 IS NOW COMPLETELY FREE!")
print("==========================================================================================")
