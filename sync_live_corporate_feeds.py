r"""
sync_live_corporate_feeds.py
===============================================================================
Synchronizes real-time and upcoming corporate events from D:\share_live\action_cop\output_excels
into the web dashboard data bundles (JSON & JS).

Ingests:
  - 04_Board_Meetings.xlsx
  - 05_Corporate_Actions.xlsx
  - 12_Event_Calendar.xlsx
  - 02_Announcements.xlsx (recent filings with PDF links)
  - 16_Insider_Trading_PIT.xlsx (recent promoter/insider transactions)
  - 34_Bulk_Deals.xlsx
  - 35_Block_Deals.xlsx
  - _run_status.json

Exports:
  - dashboard_data/live_corporate_events.json
  - dashboard_data/live_corporate_events.js (window.LIVE_CORPORATE_EVENTS)
  - docs/dashboard_data/live_corporate_events.json
  - docs/dashboard_data/live_corporate_events.js
  - Synchronizes FY27_Q3 announced earnings dates into event_dashboard_data.json / .js
===============================================================================
"""

import os
import json
from datetime import datetime, timedelta
import pandas as pd

SOURCE_DIR = r"D:\share_live\action_cop\output_excels"
BASE_DIR = r"d:\behaviour analysis"
DASH_DATA_DIR = os.path.join(BASE_DIR, "dashboard_data")
DOCS_DATA_DIR = os.path.join(BASE_DIR, "docs", "dashboard_data")

NIFTY50_SYMBOLS = {
    "ADANIENT", "ADANIPORTS", "APOLLOHOSP", "ASIANPAINT", "AXISBANK",
    "BAJAJ-AUTO", "BAJFINANCE", "BAJAJFINSV", "BEL", "BPCL",
    "BHARTIARTL", "BRITANNIA", "CIPLA", "COALINDIA", "DRREDDY",
    "EICHERMOT", "GRASIM", "HCLTECH", "HDFCBANK", "HDFCLIFE",
    "HEROMOTOCO", "HINDALCO", "HINDUNILVR", "ICICIBANK", "ITC",
    "INDUSINDBK", "INFY", "JSWSTEEL", "KOTAKBANK", "LT",
    "M&M", "MARUTI", "NTPC", "NESTLEIND", "ONGC",
    "POWERGRID", "RELIANCE", "SBILIFE", "SHRIRAMFIN", "SBIN",
    "SUNPHARMA", "TCS", "TATACONSUM", "TATAMOTORS", "TATASTEEL",
    "TECHM", "TITAN", "TRENT", "ULTRACEMCO", "WIPRO"
}

MARKET_HOLIDAYS_2026 = {
    "2026-01-26", "2026-03-03", "2026-03-25", "2026-04-03", "2026-04-14",
    "2026-05-01", "2026-06-17", "2026-08-15", "2026-08-26", "2026-10-02",
    "2026-10-20", "2026-11-08", "2026-11-10", "2026-11-24", "2026-12-25"
}

def is_trading_day(dt):
    if dt.weekday() >= 5: # Saturday or Sunday
        return False
    dt_str = dt.strftime("%Y-%m-%d")
    return dt_str not in MARKET_HOLIDAYS_2026

def shift_trading_days(dt, offset_days):
    curr = dt
    step = 1 if offset_days > 0 else -1
    remaining = abs(offset_days)
    while remaining > 0:
        curr += timedelta(days=step)
        if is_trading_day(curr):
            remaining -= 1
    return curr

def format_date_str(dt, with_day=False):
    if not dt or pd.isna(dt):
        return ""
    if with_day:
        return dt.strftime("%d-%b-%Y (%a)")
    return dt.strftime("%d-%b-%Y")

def sync_feeds():
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Starting Live Corporate Feeds Sync from: {SOURCE_DIR}")
    
    if not os.path.exists(SOURCE_DIR):
        print(f"Error: Source directory {SOURCE_DIR} does not exist!")
        return False

    now = datetime.now()
    now_str = now.strftime("%Y-%m-%d %H:%M:%S")

    # 1. Parse _run_status.json
    status_file = os.path.join(SOURCE_DIR, "_run_status.json")
    scraper_status = {}
    if os.path.exists(status_file):
        try:
            with open(status_file, "r", encoding="utf-8") as f:
                scraper_status = json.load(f)
        except Exception as e:
            print(f"Warning reading _run_status.json: {e}")

    # 2. Ingest 04_Board_Meetings.xlsx & 12_Event_Calendar.xlsx
    bm_file = os.path.join(SOURCE_DIR, "04_Board_Meetings.xlsx")
    ec_file = os.path.join(SOURCE_DIR, "12_Event_Calendar.xlsx")

    upcoming_bm_list = []
    announced_results_map = {} # symbol -> result_date_str

    if os.path.exists(bm_file):
        try:
            df_bm = pd.read_excel(bm_file)
            df_bm["dt"] = pd.to_datetime(df_bm["bm_date"], errors="coerce")
            
            # Filter from 01-Sep-2026 onwards for upcoming & active cycle
            cutoff_date = pd.to_datetime("2026-09-01")
            df_active_bm = df_bm[df_bm["dt"] >= cutoff_date].sort_values("dt", ascending=True)

            for _, row in df_active_bm.iterrows():
                sym = str(row.get("bm_symbol") or row.get("symbol") or "").strip().upper()
                comp = str(row.get("company_name") or row.get("sm_name") or sym).strip()
                purp = str(row.get("bm_purpose") or "").strip()
                desc = str(row.get("bm_desc") or "").strip()
                dt_val = row["dt"]
                dt_str = format_date_str(dt_val)

                is_nifty50 = sym in NIFTY50_SYMBOLS
                is_result = "result" in purp.lower() or "financial" in purp.lower() or "result" in desc.lower()
                is_dividend = "dividend" in purp.lower() or "dividend" in desc.lower()

                # Track confirmed earnings dates for Nifty 50 and F&O
                if is_result and sym and dt_val >= pd.to_datetime("2026-10-01"):
                    announced_results_map[sym] = dt_str

                status_label = "CONFIRMED"
                if dt_val < now:
                    status_label = "COMPLETED"
                elif (dt_val - now).days <= 3:
                    status_label = "IMMINENT"

                upcoming_bm_list.append({
                    "symbol": sym,
                    "company_name": comp,
                    "date": dt_str,
                    "raw_date": dt_val.strftime("%Y-%m-%d"),
                    "purpose": purp,
                    "description": desc if desc and desc != "nan" else purp,
                    "is_nifty50": is_nifty50,
                    "is_result": is_result,
                    "is_dividend": is_dividend,
                    "status": status_label,
                    "fetched_at": str(row.get("fetched_at") or "")
                })
        except Exception as e:
            print(f"Error parsing 04_Board_Meetings.xlsx: {e}")

    # Also check 12_Event_Calendar.xlsx for supplementary records
    if os.path.exists(ec_file):
        try:
            df_ec = pd.read_excel(ec_file)
            df_ec["dt"] = pd.to_datetime(df_ec["date"], errors="coerce")
            df_active_ec = df_ec[df_ec["dt"] >= pd.to_datetime("2026-09-01")].sort_values("dt", ascending=True)

            for _, row in df_active_ec.iterrows():
                sym = str(row.get("symbol") or "").strip().upper()
                comp = str(row.get("company") or sym).strip()
                purp = str(row.get("purpose") or "").strip()
                dt_val = row["dt"]
                dt_str = format_date_str(dt_val)

                # Avoid duplicates
                if not any(b["symbol"] == sym and b["date"] == dt_str and b["purpose"] == purp for b in upcoming_bm_list):
                    is_nifty50 = sym in NIFTY50_SYMBOLS
                    is_result = "result" in purp.lower() or "financial" in purp.lower()
                    is_dividend = "dividend" in purp.lower()

                    if is_result and sym and dt_val >= pd.to_datetime("2026-10-01"):
                        announced_results_map[sym] = dt_str

                    upcoming_bm_list.append({
                        "symbol": sym,
                        "company_name": comp,
                        "date": dt_str,
                        "raw_date": dt_val.strftime("%Y-%m-%d"),
                        "purpose": purp,
                        "description": purp,
                        "is_nifty50": is_nifty50,
                        "is_result": is_result,
                        "is_dividend": is_dividend,
                        "status": "CONFIRMED" if dt_val >= now else "COMPLETED",
                        "fetched_at": str(row.get("fetched_at") or "")
                    })
        except Exception as e:
            print(f"Error parsing 12_Event_Calendar.xlsx: {e}")

    upcoming_bm_list.sort(key=lambda x: x["raw_date"])

    # 3. Ingest 05_Corporate_Actions.xlsx (Dividends, Splits, Bonuses)
    ca_list = []
    ca_file = os.path.join(SOURCE_DIR, "05_Corporate_Actions.xlsx")
    if os.path.exists(ca_file):
        try:
            df_ca = pd.read_excel(ca_file)
            df_ca["dt"] = pd.to_datetime(df_ca["exDate"], errors="coerce")
            df_active_ca = df_ca[df_ca["dt"] >= pd.to_datetime("2026-08-01")].sort_values("dt", ascending=False)

            for _, row in df_active_ca.iterrows():
                sym = str(row.get("symbol") or "").strip().upper()
                comp = str(row.get("comp") or row.get("company_name") or sym).strip()
                subj = str(row.get("subject") or "").strip()
                dt_val = row["dt"]
                dt_str = format_date_str(dt_val)
                rec_val = pd.to_datetime(row.get("recDate"), errors="coerce")
                rec_str = format_date_str(rec_val) if not pd.isna(rec_val) else "-"

                is_nifty50 = sym in NIFTY50_SYMBOLS
                
                # Classify type
                action_type = "CORPORATE ACTION"
                if "dividend" in subj.lower():
                    action_type = "DIVIDEND"
                elif "split" in subj.lower():
                    action_type = "STOCK SPLIT"
                elif "bonus" in subj.lower():
                    action_type = "BONUS ISSUE"
                elif "rights" in subj.lower():
                    action_type = "RIGHTS ISSUE"

                ca_list.append({
                    "symbol": sym,
                    "company_name": comp,
                    "subject": subj,
                    "action_type": action_type,
                    "ex_date": dt_str,
                    "record_date": rec_str,
                    "raw_date": dt_val.strftime("%Y-%m-%d"),
                    "is_nifty50": is_nifty50,
                    "face_value": str(row.get("faceVal") or "-"),
                    "fetched_at": str(row.get("fetched_at") or "")
                })
        except Exception as e:
            print(f"Error parsing 05_Corporate_Actions.xlsx: {e}")

    # 4. Ingest 02_Announcements.xlsx (Top 60 real-time regulatory filings with PDF links)
    announcements_list = []
    an_file = os.path.join(SOURCE_DIR, "02_Announcements.xlsx")
    if os.path.exists(an_file):
        try:
            df_an = pd.read_excel(an_file, nrows=60)
            for _, row in df_an.iterrows():
                txt = str(row.get("attchmntText") or "").strip()
                pdf = str(row.get("attchmntFile") or "").strip()
                dt_raw = str(row.get("an_dt") or "").strip()
                sym = str(row.get("symbol") or "").strip().upper()
                if sym == "NAN" or not sym:
                    # Infer symbol from text or attachment URL if possible
                    if " " in txt:
                        first_words = txt.split()[:4]
                        sym = first_words[0].upper().replace(",", "").replace(".", "")
                
                if txt and txt != "nan":
                    announcements_list.append({
                        "symbol": sym if sym != "NAN" else "NSE",
                        "time": dt_raw,
                        "text": txt,
                        "pdf_url": pdf if pdf.startswith("http") else "",
                        "is_nifty50": sym in NIFTY50_SYMBOLS
                    })
        except Exception as e:
            print(f"Error parsing 02_Announcements.xlsx: {e}")

    # 5. Ingest 16_Insider_Trading_PIT.xlsx (Recent 40 transactions)
    insider_trades = []
    pit_file = os.path.join(SOURCE_DIR, "16_Insider_Trading_PIT.xlsx")
    if os.path.exists(pit_file):
        try:
            df_pit = pd.read_excel(pit_file, nrows=50)
            for _, row in df_pit.iterrows():
                comp = str(row.get("company") or "").strip()
                sym = str(row.get("symbol") or "").strip().upper()
                mode = str(row.get("acqMode") or "").strip()
                cat = str(row.get("personCategory") or "").strip()
                val = row.get("secVal") or row.get("buyValue") or row.get("sellValue") or 0
                dt_raw = str(row.get("date") or row.get("intimDt") or "").strip()
                if comp and comp != "nan":
                    insider_trades.append({
                        "company": comp,
                        "symbol": sym,
                        "person_category": cat if cat != "nan" else "Insider",
                        "mode": mode if mode != "nan" else "Market Acquisition",
                        "value": float(val) if pd.notna(val) and str(val).replace('.','').isdigit() else 0,
                        "date": dt_raw,
                        "is_nifty50": sym in NIFTY50_SYMBOLS
                    })
        except Exception as e:
            print(f"Error parsing 16_Insider_Trading_PIT.xlsx: {e}")

    # 6. Ingest 34_Bulk_Deals.xlsx & 35_Block_Deals.xlsx
    deals_list = []
    bulk_file = os.path.join(SOURCE_DIR, "34_Bulk_Deals.xlsx")
    block_file = os.path.join(SOURCE_DIR, "35_Block_Deals.xlsx")

    for f_path, deal_type in [(bulk_file, "BULK DEAL"), (block_file, "BLOCK DEAL")]:
        if os.path.exists(f_path):
            try:
                df_d = pd.read_excel(f_path)
                df_d["dt"] = pd.to_datetime(df_d["Date"], format="mixed", errors="coerce")
                df_sorted = df_d.sort_values("dt", ascending=False).head(25)
                for _, row in df_sorted.iterrows():
                    sym = str(row.get("Symbol") or "").strip().upper()
                    sec = str(row.get("Security Name") or sym).strip()
                    client = str(row.get("Client Name") or "").strip()
                    bs = str(row.get("Buy/Sell") or "").strip().upper()
                    qty = row.get("Quantity Traded")
                    px = row.get("Trade Price / Wght. Avg. Price")
                    dt_val = row["dt"]

                    deals_list.append({
                        "deal_type": deal_type,
                        "date": format_date_str(dt_val),
                        "raw_date": dt_val.strftime("%Y-%m-%d") if not pd.isna(dt_val) else "",
                        "symbol": sym,
                        "security_name": sec,
                        "client_name": client,
                        "buy_sell": bs,
                        "quantity": int(qty) if pd.notna(qty) and str(qty).replace('.','').isdigit() else str(qty),
                        "price": float(px) if pd.notna(px) else "-",
                        "is_nifty50": sym in NIFTY50_SYMBOLS
                    })
            except Exception as e:
                print(f"Error parsing {f_path}: {e}")

    deals_list.sort(key=lambda x: x["raw_date"], reverse=True)

    # 7. Assemble Live Master Payload
    live_payload = {
        "generated_at": now_str,
        "source": "D:\\share_live\\action_cop",
        "scraper_run": scraper_status.get("run", "2026-09-22 09:30:02"),
        "scraper_feeds_ok": scraper_status.get("ok", 26),
        "total_upcoming_board_meetings": len(upcoming_bm_list),
        "total_corporate_actions": len(ca_list),
        "total_recent_deals": len(deals_list),
        "total_announcements": len(announcements_list),
        "announced_q3_count": len([b for b in upcoming_bm_list if b["is_result"] and b["raw_date"] >= "2026-10-01"]),
        "upcoming_board_meetings": upcoming_bm_list,
        "corporate_actions": ca_list,
        "market_deals": deals_list[:50],
        "announcements": announcements_list[:50],
        "insider_trades": insider_trades[:40]
    }

    # 8. Save JSON & Standalone JS Bundles
    os.makedirs(DASH_DATA_DIR, exist_ok=True)
    os.makedirs(DOCS_DATA_DIR, exist_ok=True)

    json_path = os.path.join(DASH_DATA_DIR, "live_corporate_events.json")
    js_path = os.path.join(DASH_DATA_DIR, "live_corporate_events.js")
    docs_json_path = os.path.join(DOCS_DATA_DIR, "live_corporate_events.json")
    docs_js_path = os.path.join(DOCS_DATA_DIR, "live_corporate_events.js")

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(live_payload, f, indent=2)

    with open(docs_json_path, "w", encoding="utf-8") as f:
        json.dump(live_payload, f, indent=2)

    js_code = "window.LIVE_CORPORATE_EVENTS = " + json.dumps(live_payload) + ";\n"
    with open(js_path, "w", encoding="utf-8") as f:
        f.write(js_code)

    with open(docs_js_path, "w", encoding="utf-8") as f:
        f.write(js_code)

    print(f"Saved live corporate events bundle successfully:")
    print(f"  - Upcoming Board Meetings: {len(upcoming_bm_list)}")
    print(f"  - Corporate Actions: {len(ca_list)}")
    print(f"  - Market Deals: {len(deals_list)}")
    print(f"  - Regulatory Announcements: {len(announcements_list)}")
    print(f"  - Insider Transactions: {len(insider_trades)}")

    # 9. Auto-synchronize FY27_Q3 Earnings dates in event_dashboard_data.json
    sync_q3_earnings(announced_results_map)

    return True

def sync_q3_earnings(announced_results_map):
    dash_json_path = os.path.join(DASH_DATA_DIR, "event_dashboard_data.json")
    docs_dash_json_path = os.path.join(DOCS_DATA_DIR, "event_dashboard_data.json")
    
    if not os.path.exists(dash_json_path):
        return

    try:
        with open(dash_json_path, "r", encoding="utf-8") as f:
            d = json.load(f)

        updated_count = 0
        for q in d.get("quarters", []):
            if "Q3" in q.get("q_code", ""):
                for s in q.get("stocks", []):
                    sym = s.get("symbol")
                    if sym in announced_results_map:
                        res_dt_str = announced_results_map[sym]
                        res_dt = datetime.strptime(res_dt_str, "%d-%b-%Y")
                        
                        lead_days = int(s.get("lead", s.get("entry_lead_days", 2)))
                        hold_days = int(s.get("hold", s.get("exit_hold_days", 4)))

                        en_dt = shift_trading_days(res_dt, -lead_days)
                        ex_dt = shift_trading_days(res_dt, hold_days)

                        s["result_date"] = res_dt_str
                        s["entry_date"] = format_date_str(en_dt, with_day=True)
                        s["exit_date"] = format_date_str(ex_dt, with_day=True)
                        s["announced"] = True
                        s["status"] = "CONFIRMED"
                        updated_count += 1

        if updated_count > 0:
            with open(dash_json_path, "w", encoding="utf-8") as f:
                json.dump(d, f, indent=2)
            with open(docs_dash_json_path, "w", encoding="utf-8") as f:
                json.dump(d, f, indent=2)

            dash_js_code = "window.EVENT_DASHBOARD_DATA = " + json.dumps(d) + ";\nwindow.DASHBOARD_DATA = window.EVENT_DASHBOARD_DATA;\n"
            with open(os.path.join(DASH_DATA_DIR, "event_dashboard_data.js"), "w", encoding="utf-8") as f:
                f.write(dash_js_code)
            with open(os.path.join(DOCS_DATA_DIR, "event_dashboard_data.js"), "w", encoding="utf-8") as f:
                f.write(dash_js_code)

            print(f"Synchronized {updated_count} announced Q3 results into event_dashboard_data.json & .js!")
    except Exception as e:
        print(f"Error synchronizing Q3 earnings: {e}")

if __name__ == "__main__":
    sync_feeds()
