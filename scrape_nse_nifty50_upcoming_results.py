"""
scrape_nse_nifty50_upcoming_results.py
=================================================================================
Live NSE Web Scraper - Upcoming Quarterly Result Dates for Nifty 50 Stocks
=================================================================================
Hits multiple NSE official API endpoints to get board meeting / financial result
dates announced for upcoming quarters, filters for Nifty 50 stocks only, and:

  1. Prints a clean summary table
  2. Saves NSE_Official_Upcoming_Result_Dates.xlsx  (master Excel)
  3. Saves dashboard_data/upcoming_nifty50_results.json  (dashboard overlay)
  4. Optionally patches quarters_dataset.json with real scraped dates

Usage:
    python scrape_nse_nifty50_upcoming_results.py
    python scrape_nse_nifty50_upcoming_results.py --patch-dashboard
    python scrape_nse_nifty50_upcoming_results.py --per-stock-scan
"""

import urllib.request
import urllib.parse
import http.cookiejar
import json
import os
import sys
import time
import argparse
import pathlib
import warnings
from datetime import datetime, timedelta

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings('ignore')
try:
    sys.stdout.reconfigure(errors='replace')
except Exception:
    pass

ROOT = pathlib.Path('D:/behaviour analysis')
DASHBOARD_DATA = ROOT / 'dashboard_data'
OUT_EXCEL = ROOT / 'NSE_Official_Upcoming_Result_Dates.xlsx'
OUT_JSON = DASHBOARD_DATA / 'upcoming_nifty50_results.json'
QUARTERS_JSON = DASHBOARD_DATA / 'quarters_dataset.json'

NIFTY50_SYMBOLS = {
    'ADANIENT', 'ADANIPORTS', 'APOLLOHOSP', 'ASIANPAINT', 'AXISBANK',
    'BAJAJ-AUTO', 'BAJAJFINSV', 'BAJFINANCE', 'BHARTIARTL', 'BPCL',
    'BRITANNIA', 'CIPLA', 'COALINDIA', 'DIVISLAB', 'DRREDDY',
    'EICHERMOT', 'ETERNAL', 'GRASIM', 'HCLTECH', 'HDFCBANK',
    'HDFCLIFE', 'HEROMOTOCO', 'HINDALCO', 'HINDUNILVR', 'ICICIBANK',
    'INDUSINDBK', 'INFY', 'ITC', 'JIOFIN', 'JSWSTEEL',
    'KOTAKBANK', 'LT', 'LTIM', 'M&M', 'MARUTI',
    'NESTLEIND', 'NTPC', 'ONGC', 'POWERGRID', 'RELIANCE',
    'SBILIFE', 'SBIN', 'SHRIRAMFIN', 'SUNPHARMA', 'TATACONSUM',
    'TATAMOTORS', 'TATASTEEL', 'TCS', 'TECHM', 'TITAN',
    'TRENT', 'ULTRACEMCO', 'WIPRO',
}

class NSESession:
    BASE = 'https://www.nseindia.com'

    def __init__(self):
        self.cj = http.cookiejar.CookieJar()
        self.opener = urllib.request.build_opener(
            urllib.request.HTTPCookieProcessor(self.cj)
        )
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
            'Accept': 'application/json, text/javascript, */*; q=0.01',
            'Accept-Language': 'en-US,en;q=0.9',
            'X-Requested-With': 'XMLHttpRequest',
            'Connection': 'keep-alive',
        }
        self.primed = False

    def _prime(self):
        for url in [
            f'{self.BASE}/companies-listing/corporate-filings-board-meetings',
            f'{self.BASE}/market-data/event-calendar',
            f'{self.BASE}/get-quotes/equity?symbol=RELIANCE',
        ]:
            try:
                req = urllib.request.Request(url, headers={**self.headers, 'Referer': self.BASE + '/'})
                with self.opener.open(req, timeout=12) as r:
                    r.read(4096)
                self.primed = True
                return
            except Exception:
                pass
        self.primed = True

    def get_json(self, url, referer=None, retries=3):
        if not self.primed:
            self._prime()
        ref = referer or f'{self.BASE}/companies-listing/corporate-filings-board-meetings'
        h = {**self.headers, 'Referer': ref}
        for attempt in range(retries):
            try:
                req = urllib.request.Request(url, headers=h)
                with self.opener.open(req, timeout=15) as resp:
                    raw = resp.read().decode('utf-8', errors='ignore')
                    return json.loads(raw)
            except Exception:
                if attempt < retries - 1:
                    self._prime()
                    time.sleep(1.2 * (attempt + 1))
        return None


def fetch_event_calendar(nse):
    url = 'https://www.nseindia.com/api/event-calendar'
    data = nse.get_json(url) or []
    results = []
    for ev in data:
        if not isinstance(ev, dict):
            continue
        sym = (ev.get('symbol') or '').upper().strip()
        purpose = ev.get('purpose', '')
        bm_desc = ev.get('bm_desc', '')
        combo = (purpose + ' ' + bm_desc).lower()
        if any(k in combo for k in ['financial result', 'quarterly result', 'half year', 'annual result']):
            results.append({'symbol': sym, 'company': ev.get('company', ''),
                'result_date': ev.get('date', ''), 'purpose': purpose,
                'details': bm_desc, 'source': 'NSE Event Calendar API'})
    print(f"  [Event Calendar] {len(results)} financial result events")
    return results


def fetch_board_meetings_bulk(nse):
    results = []
    for segment in ['equities', 'sme']:
        url = f'https://www.nseindia.com/api/corporate-board-meetings?index={segment}'
        data = nse.get_json(url) or []
        for ev in data:
            if not isinstance(ev, dict):
                continue
            sym = (ev.get('symbol') or '').upper().strip()
            purpose = ev.get('purpose', '')
            dt_str = ev.get('bm_date', '') or ev.get('date', '')
            details = ev.get('details', '')
            combo = (purpose + ' ' + details).lower()
            if any(k in combo for k in ['financial result', 'quarterly result', 'half year', 'annual']):
                results.append({'symbol': sym, 'company': ev.get('sm_name', '') or ev.get('company', ''),
                    'result_date': dt_str, 'purpose': purpose, 'details': details,
                    'source': f'NSE Board Meetings ({segment})'})
    print(f"  [Board Meetings API] {len(results)} financial result events")
    return results


def fetch_corporate_announcements(nse):
    url = 'https://www.nseindia.com/api/corporate-announcements?index=equities'
    data = nse.get_json(url) or []
    results = []
    for ann in data:
        if not isinstance(ann, dict):
            continue
        sym = (ann.get('symbol') or '').upper().strip()
        desc = str(ann.get('desc', ''))
        att = str(ann.get('attchmntText', '') or '')
        dt = ann.get('an_dt', '') or ann.get('sort_date', '')
        combo = (desc + ' ' + att).lower()
        if 'board meeting' in combo and any(k in combo for k in ['financial result', 'quarterly result', 'half year', 'intimation']):
            results.append({'symbol': sym, 'company': ann.get('sm_name', ''),
                'result_date': dt, 'purpose': desc, 'details': att[:300] if att else '',
                'source': 'NSE Corporate Announcements'})
    print(f"  [Corporate Announcements] {len(results)} board meeting intimations")
    return results


def fetch_per_stock_board_meetings(nse, symbols, batch_delay=0.4):
    url_tpl = 'https://www.nseindia.com/api/quote-equity?symbol={sym}&section=board_meetings'
    results = []
    found = 0
    for i, sym in enumerate(symbols):
        encoded = urllib.parse.quote(sym)
        url = url_tpl.format(sym=encoded)
        ref = f'https://www.nseindia.com/get-quotes/equity?symbol={encoded}'
        data = nse.get_json(url, referer=ref, retries=2) or []
        if isinstance(data, list):
            for bm in data:
                if not isinstance(bm, dict):
                    continue
                purpose = bm.get('purpose', '')
                dt_str = bm.get('bm_date', '') or bm.get('date', '')
                combo = (purpose + ' ' + bm.get('details', '')).lower()
                if any(k in combo for k in ['financial result', 'quarterly result', 'half year']):
                    try:
                        dt = datetime.strptime(dt_str.split()[0], '%d-%b-%Y')
                        if dt >= datetime.now() - timedelta(days=7):
                            results.append({'symbol': sym, 'company': bm.get('sm_name', ''),
                                'result_date': dt_str, 'purpose': purpose,
                                'details': bm.get('details', ''), 'source': 'NSE Per-Stock BM'})
                            found += 1
                    except Exception:
                        pass
        if (i + 1) % 10 == 0:
            print(f"    {i+1}/{len(symbols)} scanned, {found} found...", flush=True)
        time.sleep(batch_delay)
    print(f"  [Per-Stock BM] Scanned {len(symbols)} Nifty 50 stocks, {found} result dates found")
    return results


def deduplicate(records):
    if not records:
        return pd.DataFrame()
    df = pd.DataFrame(records)
    df['symbol'] = df['symbol'].str.upper().str.strip()
    df['_dt_parsed'] = pd.to_datetime(
        df['result_date'].str.extract(r'(\d{2}-[A-Za-z]{3}-\d{4})')[0],
        format='%d-%b-%Y', errors='coerce'
    )
    cutoff = pd.Timestamp.now() - pd.Timedelta(days=7)
    df = df[df['_dt_parsed'].isna() | (df['_dt_parsed'] >= cutoff)].copy()
    prio = {'NSE Per-Stock BM': 1, 'NSE Event Calendar API': 2,
            'NSE Board Meetings (equities)': 3, 'NSE Corporate Announcements': 4, 'NSE Board Meetings (sme)': 5}
    df['_priority'] = df['source'].map(prio).fillna(9)
    df = df.sort_values(['symbol', '_priority', '_dt_parsed'])
    df = df.drop_duplicates(subset=['symbol', '_dt_parsed'], keep='first')
    return df.sort_values('_dt_parsed', ascending=True)


def detect_quarter(dt):
    if pd.isna(dt):
        return 'Unknown'
    month = dt.month
    year = dt.year
    if 7 <= month <= 9:
        fy = year
        return f"Q1 FY{fy}-{str(fy+1)[-2:]} (Apr-Jun {fy})"
    elif 10 <= month <= 12:
        fy = year
        return f"Q2 FY{fy}-{str(fy+1)[-2:]} (Jul-Sep {fy})"
    elif 1 <= month <= 3:
        fy = year - 1
        return f"Q3 FY{fy}-{str(fy+1)[-2:]} (Oct-Dec {fy})"
    else:
        fy = year - 1
        return f"Q4 FY{fy}-{str(fy+1)[-2:]} (Jan-Mar {year})"


def write_excel(df_n, df_a, out_path):
    wb = openpyxl.Workbook()
    hdr_blue = PatternFill("solid", fgColor="1F4E79")
    hdr_teal = PatternFill("solid", fgColor="00695C")
    hdr_purple = PatternFill("solid", fgColor="4A148C")
    hdr_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    alt_fill = PatternFill("solid", fgColor="EFF6FF")
    nifty_fill = PatternFill("solid", fgColor="E8F5E9")
    thin = Border(left=Side(style='thin',color='CCCCCC'), right=Side(style='thin',color='CCCCCC'),
                  top=Side(style='thin',color='CCCCCC'), bottom=Side(style='thin',color='CCCCCC'))
    ctr = Alignment(horizontal='center', vertical='center', wrap_text=True)
    lft = Alignment(horizontal='left', vertical='center', wrap_text=True)
    today = pd.Timestamp.now().normalize()

    # Sheet 1: Nifty 50
    ws1 = wb.active
    ws1.title = "Nifty 50 Upcoming Result Dates"
    hdrs1 = ["Sr.", "Symbol", "Company Name", "Result Date (NSE)", "Quarter", "Days to Result", "Purpose", "Source"]
    ws1.row_dimensions[1].height = 28
    for ci, h in enumerate(hdrs1, 1):
        c = ws1.cell(row=1, column=ci, value=h)
        c.font = hdr_font; c.fill = hdr_blue; c.alignment = ctr; c.border = thin

    for ri, row in enumerate(df_n.to_dict('records'), start=2):
        dt = row.get('_dt_parsed')
        dt_str = pd.Timestamp(dt).strftime('%d-%b-%Y (%a)') if pd.notna(dt) else str(row.get('result_date',''))
        days_str = f"{int((pd.Timestamp(dt)-today).days)}d" if pd.notna(dt) else '?'
        vals = [ri-1, row.get('symbol',''), row.get('company',''), dt_str,
                detect_quarter(pd.Timestamp(dt) if pd.notna(dt) else pd.NaT),
                days_str, row.get('purpose',''), row.get('source','')]
        aligns = [ctr, ctr, lft, ctr, lft, ctr, lft, lft]
        fill = nifty_fill if ri % 2 == 0 else None
        for ci, (v, al) in enumerate(zip(vals, aligns), 1):
            c = ws1.cell(row=ri, column=ci, value=v)
            c.alignment = al; c.border = thin
            if fill: c.fill = fill

    for col in ws1.columns:
        mx = max((len(str(c.value or '')) for c in col), default=8)
        ws1.column_dimensions[get_column_letter(col[0].column)].width = min(mx + 4, 60)
    ws1.freeze_panes = ws1['A2']
    ws1.auto_filter.ref = ws1.dimensions

    # Sheet 2: All NSE
    ws2 = wb.create_sheet("All NSE Result Dates")
    hdrs2 = ["Sr.", "Symbol", "Company Name", "Result Date (NSE)", "Quarter", "Days to Result", "Nifty 50?", "Purpose", "Source"]
    ws2.row_dimensions[1].height = 28
    for ci, h in enumerate(hdrs2, 1):
        c = ws2.cell(row=1, column=ci, value=h)
        c.font = hdr_font; c.fill = hdr_teal; c.alignment = ctr; c.border = thin

    for ri, row in enumerate(df_a.to_dict('records'), start=2):
        dt = row.get('_dt_parsed')
        dt_str = pd.Timestamp(dt).strftime('%d-%b-%Y (%a)') if pd.notna(dt) else str(row.get('result_date',''))
        days_str = f"{int((pd.Timestamp(dt)-today).days)}d" if pd.notna(dt) else '?'
        in_n50 = row.get('symbol','') in NIFTY50_SYMBOLS
        vals = [ri-1, row.get('symbol',''), row.get('company',''), dt_str,
                detect_quarter(pd.Timestamp(dt) if pd.notna(dt) else pd.NaT),
                days_str, 'YES' if in_n50 else '-', row.get('purpose',''), row.get('source','')]
        aligns = [ctr, ctr, lft, ctr, lft, ctr, ctr, lft, lft]
        fill = nifty_fill if in_n50 else (alt_fill if ri % 2 == 0 else None)
        for ci, (v, al) in enumerate(zip(vals, aligns), 1):
            c = ws2.cell(row=ri, column=ci, value=v)
            c.alignment = al; c.border = thin
            if fill: c.fill = fill

    for col in ws2.columns:
        mx = max((len(str(c.value or '')) for c in col), default=8)
        ws2.column_dimensions[get_column_letter(col[0].column)].width = min(mx + 4, 60)
    ws2.freeze_panes = ws2['A2']
    ws2.auto_filter.ref = ws2.dimensions

    # Sheet 3: Quarter Summary
    ws3 = wb.create_sheet("Quarter Summary")
    ws3.row_dimensions[1].height = 28
    hdrs3 = ["Quarter", "Nifty 50 Stocks", "All NSE Stocks", "Date Range (NSE Announced)"]
    for ci, h in enumerate(hdrs3, 1):
        c = ws3.cell(row=1, column=ci, value=h)
        c.font = hdr_font; c.fill = hdr_purple; c.alignment = ctr; c.border = thin

    all_quarters = sorted(set(
        list(df_n['_quarter'].unique() if '_quarter' in df_n.columns else []) +
        list(df_a['_quarter'].unique() if '_quarter' in df_a.columns else [])
    ))

    if not df_n.empty:
        df_n['_quarter'] = df_n['_dt_parsed'].apply(detect_quarter)
    if not df_a.empty:
        df_a['_quarter'] = df_a['_dt_parsed'].apply(detect_quarter)

    q_groups_n = df_n.groupby('_quarter') if not df_n.empty else {}
    q_groups_a = df_a.groupby('_quarter') if not df_a.empty else {}

    quarters_seen = set()
    if not df_n.empty:
        quarters_seen.update(df_n['_quarter'].unique())
    if not df_a.empty:
        quarters_seen.update(df_a['_quarter'].unique())

    for ri, q in enumerate(sorted(quarters_seen), start=2):
        n50_cnt = len(df_n[df_n['_quarter'] == q]) if not df_n.empty else 0
        all_cnt = len(df_a[df_a['_quarter'] == q]) if not df_a.empty else 0
        q_df = df_a[df_a['_quarter'] == q] if not df_a.empty else pd.DataFrame()
        if not q_df.empty:
            mn = q_df['_dt_parsed'].min()
            mx = q_df['_dt_parsed'].max()
            dr = f"{pd.Timestamp(mn).strftime('%d %b')} -> {pd.Timestamp(mx).strftime('%d %b %Y')}"
        else:
            dr = '—'
        vals = [q, n50_cnt, all_cnt, dr]
        aligns = [lft, ctr, ctr, ctr]
        fill = alt_fill if ri % 2 == 0 else None
        for ci, (v, al) in enumerate(zip(vals, aligns), 1):
            c = ws3.cell(row=ri, column=ci, value=v)
            c.alignment = al; c.border = thin
            if fill: c.fill = fill

    for col in ws3.columns:
        mx = max((len(str(c.value or '')) for c in col), default=8)
        ws3.column_dimensions[get_column_letter(col[0].column)].width = min(mx + 4, 55)
    ws3.freeze_panes = ws3['A2']

    wb.save(out_path)
    print(f"  Saved: {out_path}")


def write_dashboard_json(df_n, out_path):
    today = pd.Timestamp.now().normalize()
    records = []
    for _, row in df_n.iterrows():
        dt = row.get('_dt_parsed')
        records.append({
            'symbol': row['symbol'],
            'company': row['company'],
            'result_date': pd.Timestamp(dt).strftime('%d-%b-%Y') if pd.notna(dt) else row.get('result_date', ''),
            'result_date_day': pd.Timestamp(dt).strftime('%A') if pd.notna(dt) else '',
            'quarter': detect_quarter(pd.Timestamp(dt)) if pd.notna(dt) else 'Unknown',
            'days_to_result': int((pd.Timestamp(dt) - today).days) if pd.notna(dt) else None,
            'purpose': row.get('purpose', ''),
            'source': row.get('source', ''),
            'scraped_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        })
    payload = {'generated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
               'total_found': len(records), 'data': records}
    out_path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding='utf-8')
    print(f"  Saved dashboard JSON: {out_path}")


def patch_quarters_dataset(df_n, json_path):
    if not json_path.exists():
        print("  WARNING: quarters_dataset.json not found, skipping patch.")
        return
    with open(json_path, 'r', encoding='utf-8') as f:
        quarters = json.load(f)
    scraped_map = {}
    for _, row in df_n.iterrows():
        dt = row.get('_dt_parsed')
        if pd.notna(dt):
            scraped_map[row['symbol']] = pd.Timestamp(dt).strftime('%d-%b-%Y')
    patched = 0
    for quarter in quarters:
        for stock in quarter.get('stocks', []):
            sym = stock.get('symbol', '')
            if sym in scraped_map:
                new_date = scraped_map[sym]
                if stock.get('result_declaration_date') != new_date:
                    stock['result_declaration_date'] = new_date
                    stock['date_source'] = 'NSE Official (Scraped)'
                    patched += 1
    if patched > 0:
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(quarters, f, indent=2, ensure_ascii=False)
        print(f"  Patched {patched} result dates in quarters_dataset.json")
    else:
        print("  No patches needed (dates already up to date)")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--patch-dashboard', action='store_true', help='Patch quarters_dataset.json with real NSE dates')
    ap.add_argument('--per-stock-scan', action='store_true', help='Do per-stock API scan (slower but thorough)')
    args = ap.parse_args()

    BAR = '=' * 90
    print(BAR)
    print(' NSE OFFICIAL QUARTERLY RESULT DATE SCRAPER - NIFTY 50 STOCKS')
    print(BAR)
    print(f' Scraped At : {datetime.now().strftime("%d-%b-%Y %H:%M:%S")}')
    print(f' Nifty 50   : {len(NIFTY50_SYMBOLS)} canonical symbols')
    print(BAR, flush=True)

    nse = NSESession()
    print('\n[1/4] Priming NSE session (acquiring cookies)...', flush=True)
    nse._prime()
    print('  Session primed\n')

    print('[2/4] Fetching result dates from NSE APIs...', flush=True)
    all_records = []
    all_records += fetch_event_calendar(nse)
    time.sleep(0.8)
    all_records += fetch_board_meetings_bulk(nse)
    time.sleep(0.8)
    all_records += fetch_corporate_announcements(nse)
    time.sleep(0.8)

    if args.per_stock_scan:
        print('\n  Per-stock scan requested (~2-3 minutes)...', flush=True)
        all_records += fetch_per_stock_board_meetings(nse, sorted(NIFTY50_SYMBOLS), batch_delay=0.4)

    print(f'\n  Total raw records: {len(all_records)}', flush=True)

    print('\n[3/4] Deduplicating and filtering...', flush=True)
    df_all = deduplicate(all_records)

    if df_all.empty:
        print('\n  WARNING: No upcoming result dates found from NSE APIs.')
        print('  Try: python scrape_nse_nifty50_upcoming_results.py --per-stock-scan')
        df_nifty50 = pd.DataFrame()
    else:
        df_nifty50 = df_all[df_all['symbol'].isin(NIFTY50_SYMBOLS)].drop_duplicates(subset=['symbol'], keep='first').copy()

    print(f'  Total unique upcoming dates : {len(df_all)}')
    print(f'  Nifty 50 stocks found       : {len(df_nifty50)}')

    print('\n' + BAR)
    print(' NIFTY 50 UPCOMING QUARTERLY RESULT DATES (Live from NSE)')
    print(BAR)
    if not df_nifty50.empty:
        today = pd.Timestamp.now().normalize()
        df_nifty50['_quarter'] = df_nifty50['_dt_parsed'].apply(detect_quarter)
        print(f"\n  {'Sr.':<4} {'Symbol':<14} {'Result Date':<22} {'Days Left':<11} {'Quarter':<40} Source")
        print('  ' + '-' * 110)
        for i, (_, row) in enumerate(df_nifty50.sort_values('_dt_parsed').iterrows(), 1):
            dt = row['_dt_parsed']
            dt_str = pd.Timestamp(dt).strftime('%d-%b-%Y (%a)') if pd.notna(dt) else str(row.get('result_date','N/A'))
            days_str = f"{int((pd.Timestamp(dt)-today).days)}d" if pd.notna(dt) else '?'
            print(f"  {i:<4} {row['symbol']:<14} {dt_str:<22} {days_str:<11} {row.get('_quarter',''):<40} {row.get('source','')}")
    else:
        print('\n  No Nifty 50 result dates found. Try --per-stock-scan for deeper scan.')

    print('\n' + BAR)
    print(' ALL NSE STOCKS WITH UPCOMING RESULT DATES')
    print(BAR)
    if not df_all.empty:
        df_all['_quarter'] = df_all['_dt_parsed'].apply(detect_quarter)
        print(f"\n  Total: {len(df_all)} stocks\n")
        print(f"  {'Sr.':<4} {'Symbol':<16} {'Result Date':<20} {'Quarter':<40} {'Nifty 50'}")
        print('  ' + '-' * 90)
        for i, (_, row) in enumerate(df_all.sort_values('_dt_parsed').iterrows(), 1):
            dt = row['_dt_parsed']
            dt_str = pd.Timestamp(dt).strftime('%d-%b-%Y') if pd.notna(dt) else str(row.get('result_date','N/A'))
            marker = '<-- NIFTY 50' if row['symbol'] in NIFTY50_SYMBOLS else ''
            print(f"  {i:<4} {row['symbol']:<16} {dt_str:<20} {row.get('_quarter','?'):<40} {marker}")

    print('\n[4/4] Saving outputs...', flush=True)
    if not df_all.empty:
        write_excel(df_nifty50, df_all, OUT_EXCEL)
    if not df_nifty50.empty:
        write_dashboard_json(df_nifty50, OUT_JSON)
    if args.patch_dashboard and not df_nifty50.empty:
        print('\n  Patching quarters_dataset.json with live NSE dates...')
        patch_quarters_dataset(df_nifty50, QUARTERS_JSON)

    print('\n' + BAR)
    print(' DONE - NSE Nifty 50 Quarterly Result Date Scraper Complete!')
    print(BAR + '\n')

if __name__ == '__main__':
    main()
