import pathlib, sys

sys.stdout.reconfigure(errors='replace')

OI_ROOT = pathlib.Path('D:/behaviour analysis/OI_DATA')
oi_stocks = sorted([d.name for d in OI_ROOT.iterdir() if d.is_dir()])
print(f"Total folders: {len(oi_stocks)}")
print(oi_stocks)
