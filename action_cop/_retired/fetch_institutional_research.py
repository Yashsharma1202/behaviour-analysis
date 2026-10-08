import os
import time
import requests
import pandas as pd

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "*/*",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.nseindia.com/companies-listing/corporate-filings-announcements",
}

OUTPUT_DIR = r"d:\action_coop\output_excels"

def get_session():
    session = requests.Session()
    try:
        session.get("https://www.nseindia.com", headers=HEADERS, timeout=10)
        time.sleep(1)
    except Exception as e:
        print(f"[WARN] Session init warning: {e}")
    return session

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

def load_nifty50_symbols(csv_path=r"d:\action_coop\ind_nifty50list.csv"):
    df = pd.read_csv(csv_path)
    symbols = df["Symbol"].str.strip().tolist()
    symbol_map = dict(zip(df["Symbol"].str.strip(), df["Company Name"].str.strip()))
    return symbols, symbol_map

def run_institutional_scraper():
    print("=== Starting Earnings Calls & Institutional Research Scraper ===")
    symbols, symbol_map = load_nifty50_symbols()
    session = get_session()
    nifty_set = set(symbols)

    # 1. Concall Transcripts & Investor Presentations
    print("\n--- [36/39] Fetching Concall Transcripts & Investor Presentation Links ---")
    concall_records = []
    for sym in symbols[:15]: # batch sample across Nifty 50
        url = f"https://www.nseindia.com/api/corporate-announcements?index=equities&symbol={sym}"
        try:
            r = session.get(url, headers=HEADERS, timeout=10)
            if r.status_code == 200:
                ann_list = r.json()
                for a in ann_list:
                    desc = str(a.get("desc", ""))
                    att_text = str(a.get("attchmntText", ""))
                    combined_text = (desc + " " + att_text).lower()
                    
                    if any(kw in combined_text for kw in ["transcript", "presentation", "concall", "earnings call", "analyst meet"]):
                        file_name = a.get("attchmntFile", "")
                        pdf_link = f"https://nsearchives.nseindia.com/corporate/{file_name}" if file_name and not file_name.startswith("http") else file_name
                        concall_records.append({
                            "Symbol": sym,
                            "Company Name": symbol_map.get(sym, sym),
                            "Date": a.get("an_dt"),
                            "Category": "Presentation" if "presentation" in combined_text else "Transcript",
                            "Description": desc,
                            "Attachment Text": att_text,
                            "PDF Download Link": pdf_link,
                            "BroadCast Time": a.get("exchdisstime")
                        })
        except Exception as e:
            pass
        time.sleep(0.15)
    df_concall = pd.DataFrame(concall_records) if concall_records else pd.DataFrame()
    save_to_excel(df_concall, "36_Concall_Transcripts_Presentations.xlsx", "Concalls & Presentations", dedupe_subset=["Symbol", "Date", "Description"])

    # 2. Credit Rating Reports
    print("\n--- [37/39] Fetching Credit Rating Reports (CRISIL, ICRA, CARE, India Ratings) ---")
    try:
        r_cr = session.get("https://www.nseindia.com/api/credit-rating-sdd-reg30?index=equities", headers=HEADERS, timeout=12)
        if r_cr.status_code == 200:
            cr_data = r_cr.json()
            if isinstance(cr_data, list):
                df_cr = pd.DataFrame(cr_data)
                if "symbol" in df_cr.columns:
                    df_cr = df_cr[df_cr["symbol"].isin(nifty_set)]
            else:
                df_cr = pd.DataFrame()
        else:
            df_cr = pd.DataFrame()
    except Exception as e:
        df_cr = pd.DataFrame()
    save_to_excel(df_cr, "37_Credit_Rating_Reports.xlsx", "Credit Rating Reports", dedupe_subset=["symbol", "dateOfCurrentCredit", "creditRating"])

    # 3. Shareholding Master (Mutual Funds, FIIs & Public > 1%)
    print("\n--- [38/39] Fetching Shareholding Disclosures (Super Investors & Institutional Holdings) ---")
    try:
        r_shp = session.get("https://www.nseindia.com/api/corporate-share-holdings-master?index=equities", headers=HEADERS, timeout=12)
        if r_shp.status_code == 200:
            shp_data = r_shp.json()
            if isinstance(shp_data, list):
                df_shp = pd.DataFrame(shp_data)
                if "symbol" in df_shp.columns:
                    df_shp = df_shp[df_shp["symbol"].isin(nifty_set)]
            elif isinstance(shp_data, dict) and "data" in shp_data:
                df_shp = pd.DataFrame(shp_data["data"])
                if "symbol" in df_shp.columns:
                    df_shp = df_shp[df_shp["symbol"].isin(nifty_set)]
            else:
                df_shp = pd.DataFrame()
        else:
            df_shp = pd.DataFrame()
    except Exception as e:
        df_shp = pd.DataFrame()
    save_to_excel(df_shp, "38_Super_Investor_Institutional_Holdings.xlsx", "Super Investor & Inst Holdings", dedupe_subset=["symbol", "asOnDate"])

    print("\n=======================================================")
    print(f"INSTITUTIONAL RESEARCH SCRAPING COMPLETE: {OUTPUT_DIR}")
    print("=======================================================")

if __name__ == "__main__":
    run_institutional_scraper()
