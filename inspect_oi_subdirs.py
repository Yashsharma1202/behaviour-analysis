import os, pathlib, sys

sys.stdout.reconfigure(errors='replace')

p1 = pathlib.Path(r'D:\behaviour analysis\OI_DATA\360ONE')
p2 = pathlib.Path(r'E:\OI_DATA\360ONE')

print("==========================================================================================")
print("INSPECTING SUBDIRECTORY TREE FOR 360ONE")
print("==========================================================================================")

def get_tree(base_path):
    tree = {}
    for root, dirs, files in os.walk(base_path):
        rel = os.path.relpath(root, base_path)
        tree[rel] = (len(dirs), len(files), [f for f in files[:5]])
    return tree

tree1 = get_tree(p1)
tree2 = get_tree(p2)

print(f"Subdirectories in D:\\behaviour analysis\\OI_DATA\\360ONE : {len(tree1)}")
print(f"Subdirectories in E:\\OI_DATA\\360ONE                  : {len(tree2)}")

print("\nSample Subdirectories & File Counts in D:")
for k, v in list(tree1.items())[:10]:
    print(f"  Folder '{k}': {v[0]} dirs, {v[1]} files. Sample files: {v[2]}")

print("\nSample Subdirectories & File Counts in E:")
for k, v in list(tree2.items())[:10]:
    print(f"  Folder '{k}': {v[0]} dirs, {v[1]} files. Sample files: {v[2]}")

print("\n------------------------------------------------------------------------------------------")
print(f"Tree structure 100% identical for 360ONE? {tree1 == tree2}")
print("==========================================================================================")
