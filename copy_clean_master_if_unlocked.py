import shutil, pathlib, sys

sys.stdout.reconfigure(errors='replace')

b_dir = pathlib.Path('D:/behaviour analysis')

master_clean = b_dir / 'Nifty50_12_Quarters_Futures_0.5PCT_Total_Turnover_Master_CLEAN.xlsx'
master_orig  = b_dir / 'Nifty50_12_Quarters_Futures_0.5PCT_Total_Turnover_Master.xlsx'

summary_clean = b_dir / 'Nifty50_12_Quarters_Futures_0.5PCT_Total_Turnover_Summary_CLEAN.xlsx'
summary_orig  = b_dir / 'Nifty50_12_Quarters_Futures_0.5PCT_Total_Turnover_Summary.xlsx'

try:
    shutil.copy2(master_clean, master_orig)
    print(f"✅ Successfully updated {master_orig.name}")
except Exception as e:
    print(f"⚠️ Primary master file locked in Excel: {e}")

try:
    shutil.copy2(summary_clean, summary_orig)
    print(f"✅ Successfully updated {summary_orig.name}")
except Exception as e:
    print(f"⚠️ Primary summary file locked in Excel: {e}")
