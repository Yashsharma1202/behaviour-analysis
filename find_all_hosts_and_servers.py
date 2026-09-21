import os, re, pathlib, sys

sys.stdout.reconfigure(errors='replace')

b_dir = pathlib.Path('D:/behaviour analysis')

print("==========================================================================================")
print("EXTRACTING ALL HOST CONFIGURATIONS & LOCAL WEB SERVERS IN CODEBASE")
print("==========================================================================================")

host_findings = []

for root, dirs, files in os.walk(b_dir):
    if '.git' in root or '__pycache__' in root:
        continue
    for f in files:
        if f.endswith(('.py', '.html', '.json', '.txt', '.bat', '.sh')):
            fp = pathlib.Path(root) / f
            try:
                content = fp.read_text(encoding='utf-8', errors='ignore')
                
                # Check for host=, port=, http://, https://, ip regex
                lines = content.splitlines()
                for line_idx, line in enumerate(lines, start=1):
                    if re.search(r'host\s*=|port\s*=|http://|https://|127\.0\.0\.1|localhost|0\.0\.0\.0', line, re.IGNORECASE):
                        host_findings.append({
                            'file': str(fp.relative_to(b_dir)),
                            'line': line_idx,
                            'content': line.strip()
                        })
            except Exception as e:
                pass

print(f"Total Host / Port / URL configuration entries found: {len(host_findings)}\n")

grouped_files = {}
for entry in host_findings:
    f_name = entry['file']
    if f_name not in grouped_files:
        grouped_files[f_name] = []
    grouped_files[f_name].append(entry)

for f_name, entries in sorted(grouped_files.items()):
    print(f"📄 File: {f_name}")
    for e in entries[:8]:
        print(f"   L{e['line']:4d}: {e['content']}")
    if len(entries) > 8:
        print(f"   ... ({len(entries)-8} more entries in file)")
    print()

print("==========================================================================================")
