import os, pathlib, sys, pandas as pd

sys.stdout.reconfigure(errors='replace')

base_dir = pathlib.Path(r'D:\behaviour analysis\OI_DATA')

print("==========================================================================================")
print("READING & INSPECTING D:\\behaviour analysis\\OI_DATA")
print("==========================================================================================")

# Find sample parquet files for futures, options, and spot
sample_stock = '360ONE'
stock_dir = base_dir / sample_stock

print(f"Stock Directory: {stock_dir}")

futures_dir = stock_dir / 'futures_parquet'
options_dir = stock_dir / 'options_parquet'
spot_dir    = stock_dir / 'spot_parquet'

def read_sample_file(category_name, folder_path):
    print(f"\n------------------------------------------------------------------------------------------")
    print(f"📁 Category: {category_name}")
    print(f"   Folder: {folder_path}")
    
    if not folder_path.exists():
        print("   ❌ Folder does not exist!")
        return
        
    p_files = sorted([f for f in os.listdir(folder_path) if f.endswith('.parquet')])
    print(f"   Total Parquet Files: {len(p_files)}")
    
    if not p_files:
        print("   ⚠️ No parquet files found in folder.")
        return
        
    sample_file = folder_path / p_files[0]
    print(f"   Sample File: {sample_file.name} (File Size: {sample_file.stat().st_size:,} bytes)")
    
    try:
        df = pd.read_parquet(sample_file)
        print(f"   Shape: {df.shape[0]:,} rows x {df.shape[1]} columns")
        print("   Column Names & Types:")
        for col in df.columns:
            print(f"     - {col:30s} : {str(df[col].dtype)}")
            
        print("\n   Sample Data Rows (Top 5):")
        print(df.head(5).to_string(index=False))
    except Exception as e:
        print(f"   ❌ Error reading parquet file: {e}")

read_sample_file('Futures Parquet', futures_dir)
read_sample_file('Options Parquet', options_dir)
read_sample_file('Spot Parquet', spot_dir)

print("\n==========================================================================================")
