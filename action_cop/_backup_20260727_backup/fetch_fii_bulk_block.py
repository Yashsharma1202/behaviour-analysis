import os
import requests
import pandas as pd
import time

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "*/*",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.nseindia.com/reports/fii-dii",
}

OUTPUT_DIR = r"d:\action_coop\output_excels"

def clean_dataframe(df):
    if df is None or df.empty:
        return df
    df = df.dropna(how='all', axis=1)
    for col in df.columns:
        df[col] = df[col].apply(lambda x: str(x) if isinstance(x, (dict, list)) else x)
    return df

def save_to_excel(df, filename, category_name, dedupe_subset=None):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    filepath = os.path.join(OUTPUT_DIR, filename)
    df = clean_dataframe(df)

    if os.path.exists(filepath):
        try:
            existing_df = pd.read_excel(filepath)
            existing_df = clean_dataframe(existing_df)
            if "Message" in existing_df.columns:
                existing_df = existing_df[~existing_df["Message"].astype(str).str.contains("No records found", na=False)]
            
            if df is not None and not df.empty:
                if "Message" in df.columns:
                    df = df[~df["Message"].astype(str).str.contains("No records found", na=False)]
                if not df.empty and not existing_df.empty:
                    combined_df = pd.concat([existing_df, df], ignore_index=True)
                elif not existing_df.empty:
                    combined_df = existing_df
                else:
                    combined_df = df
            else:
                combined_df = existing_df

            if not combined_df.empty:
                if dedupe_subset:
                    valid_subset = [c for c in dedupe_subset if c in combined_df.columns]
                    if valid_subset:
                        combined_df = combined_df.drop_duplicates(subset=valid_subset, keep="last")
                    else:
                        combined_df = combined_df.drop_duplicates(keep="last")
                else:
                    combined_df = combined_df.drop_duplicates(keep="last")
                df = combined_df
        except Exception as e:
            print(f"[WARN] Could not merge with existing {filename}: {e}")

    if df is None or df.empty:
        df = pd.DataFrame([{"Message": f"No records found on NSE for {category_name}."}])
    else:
        if len(df.columns) > 1 and "Message" in df.columns:
            df = df.drop(columns=["Message"])

    for attempt in range(3):
        try:
            with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name=category_name[:31], index=False)
                worksheet = writer.sheets[category_name[:31]]
                for col in worksheet.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = col[0].column_letter
                    worksheet.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 60)
            print(f"[SUCCESS] Saved {len(df)} total accumulated rows to {filepath}")
            return filepath
        except PermissionError:
            time.sleep(1.0)
        except Exception as e:
            print(f"[WARN] Could not save {filename}: {e}")
            break
    return None

session = requests.Session()
session.get("https://www.nseindia.com", headers=HEADERS, timeout=10)
time.sleep(1)

print("Fetching FII/DII Daily Trading Activity...")
try:
    r = session.get("https://www.nseindia.com/api/fiidiiTradeReact", headers=HEADERS, timeout=10)
    data = r.json() if r.status_code == 200 else []
    df_fiidii = pd.DataFrame(data)
except Exception as e:
    df_fiidii = pd.DataFrame()
save_to_excel(df_fiidii, "33_FII_DII_Trading_Activity.xlsx", "FII DII Trading Activity", dedupe_subset=["date", "category"])

print("Fetching Bulk Deals...")
try:
    df_bulk = pd.read_csv("https://nsearchives.nseindia.com/content/equities/bulk.csv", storage_options={"User-Agent": HEADERS["User-Agent"]})
except Exception as e:
    df_bulk = pd.DataFrame()
save_to_excel(df_bulk, "34_Bulk_Deals.xlsx", "Bulk Deals", dedupe_subset=["Date", "Symbol", "Client Name", "Quantity Traded"])

print("Fetching Block Deals...")
try:
    df_block = pd.read_csv("https://nsearchives.nseindia.com/content/equities/block.csv", storage_options={"User-Agent": HEADERS["User-Agent"]})
except Exception as e:
    df_block = pd.DataFrame()
save_to_excel(df_block, "35_Block_Deals.xlsx", "Block Deals", dedupe_subset=["Date", "Symbol", "Client Name", "Quantity Traded"])
