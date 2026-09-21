import socket, os, pathlib, sys

sys.stdout.reconfigure(errors='replace')

print("==========================================================================================")
print("SYSTEM & NETWORK HOST INFORMATION")
print("==========================================================================================")

# 1. System Hostname & Local IP Addresses
hostname = socket.gethostname()
print(f"System Hostname : {hostname}")

try:
    local_ip = socket.gethostbyname(hostname)
    print(f"Primary Local IP: {local_ip}")
except Exception as e:
    print(f"Local IP Error  : {e}")

try:
    ip_list = socket.gethostbyname_ex(hostname)[2]
    print(f"All Active IPs  : {ip_list}")
except Exception as e:
    pass

# 2. Windows Hosts File Entries (C:\Windows\System32\drivers\etc\hosts)
hosts_file_path = pathlib.Path(r'C:\Windows\System32\drivers\etc\hosts')
print("\n------------------------------------------------------------------------------------------")
print(f"Windows Hosts File ({hosts_file_path}):")
print("------------------------------------------------------------------------------------------")

if hosts_file_path.exists():
    try:
        hosts_content = hosts_file_path.read_text(encoding='utf-8', errors='ignore')
        active_lines = [line.strip() for line in hosts_content.splitlines() if line.strip() and not line.strip().startswith('#')]
        if active_lines:
            for al in active_lines:
                print(f"  - {al}")
        else:
            print("  (No custom active host mappings in hosts file - default standard settings)")
    except Exception as e:
        print(f"  Error reading hosts file: {e}")

# 3. Python Web Application Host Servers in Project Directory
print("\n------------------------------------------------------------------------------------------")
print("Python Web Application Host Servers in `D:\\behaviour analysis`:")
print("------------------------------------------------------------------------------------------")

servers = [
    ('stock_server.py', 5000, 'Local Web Application Dashboard'),
    ('short_stock_server.py', 5001, 'Short Selling Dashboard Server'),
    ('ml_host.py', 5002, 'ML Prediction Host Server'),
    ('exp_bar_host.py', 5003, 'Expectation Bar Host Server')
]

for s_file, s_port, desc in servers:
    s_path = pathlib.Path('D:/behaviour analysis') / s_file
    exists = s_path.exists()
    print(f"  • Host Server: {s_file:25s} | Default Port: {s_port} | URL: http://localhost:{s_port} | Exists: {exists}")
    print(f"    Description: {desc}")

print("==========================================================================================")
