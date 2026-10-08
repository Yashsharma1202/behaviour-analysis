"""
EXTRACT REAL, NSE-SOURCED QUARTERLY RESULT (ANNOUNCEMENT) DATES for the 50
Nifty 50 stocks, for the 17 historical quarters (FY23_Q2 .. FY27_Q2) that
currently hold SYNTHETIC placeholder result dates in event_dashboard_data.json
(every stock was found to "report" on a fixed day-of-month, exactly 3 months
apart, for 17 straight quarters -- e.g. TCS always the 20th, HDFCBANK always
the 17th -- which is not how real companies report).

Two authentic NSE-sourced local feeds, both under D:\\share_live\\action_cop\\
output_excels\\, are combined:

  PRIMARY:   13_Financial_Results.xlsx -- NSE's XBRL financial-results
             archive (25,727 rows back to 2008; each row traces to a real
             nsearchives.nseindia.com URL). Clean fromDate/toDate reporting
             period columns + broadCastDate. Covers roughly FY23_Q2..FY25_Q3
             for most symbols (archive wasn't backfilled past ~late 2024).

  FALLBACK:  04_Board_Meetings.xlsx -- NSE board-meeting intimations (real
             range 2016-07-15 .. 2026-11-13). No clean period columns; the
             reporting period is parsed out of the free-text bm_desc (e.g.
             "...period ended December 2025..."). Used only to fill (symbol,
             quarter) slots the primary source doesn't have -- this covers
             the FY25_Q4..FY27_Q2 tail the primary source is missing.

For each (symbol, quarter) the EARLIEST matching date is kept (first public
announcement), preferring Consolidated over Non-Consolidated in the primary
source, and the 'Financial Results/Dividend' purpose row over 'Board Meeting
Intimation' in the fallback (same date either way, but avoids relying on an
advance-notice-only row with no matching purpose text).

ADDITIVE: reads the external NSE archive + fo_stocks_211.json; writes a
standalone record file to data_center/ (does not touch the dashboard yet --
that is a separate, deliberate follow-up step once this extraction is
verified).
"""
import json
import re
import openpyxl
from datetime import datetime

BASE = r'D:\behaviour analysis'
SRC_RESULTS = r'D:\share_live\action_cop\output_excels\13_Financial_Results.xlsx'
SRC_BOARD = r'D:\share_live\action_cop\output_excels\04_Board_Meetings.xlsx'

fo = json.load(open(BASE + r'\dashboard_data\fo_stocks_211.json', encoding='utf-8'))
n50 = sorted(f['symbol'] for f in fo if str(f.get('is_nifty50', '')).upper() == 'YES')

SYMBOL_ALIASES = {
    'TMPV': ['TMPV', 'TATAMOTORS'],   # Tata Motors Passenger Vehicles, renamed from TATAMOTORS in 2025
}

QUARTERS = [
    ('FY23_Q2', '01-Jul-2022', '30-Sep-2022'),
    ('FY23_Q3', '01-Oct-2022', '31-Dec-2022'),
    ('FY23_Q4', '01-Jan-2023', '31-Mar-2023'),
    ('FY24_Q1', '01-Apr-2023', '30-Jun-2023'),
    ('FY24_Q2', '01-Jul-2023', '30-Sep-2023'),
    ('FY24_Q3', '01-Oct-2023', '31-Dec-2023'),
    ('FY24_Q4', '01-Jan-2024', '31-Mar-2024'),
    ('FY25_Q1', '01-Apr-2024', '30-Jun-2024'),
    ('FY25_Q2', '01-Jul-2024', '30-Sep-2024'),
    ('FY25_Q3', '01-Oct-2024', '31-Dec-2024'),
    ('FY25_Q4', '01-Jan-2025', '31-Mar-2025'),
    ('FY26_Q1', '01-Apr-2025', '30-Jun-2025'),
    ('FY26_Q2', '01-Jul-2025', '30-Sep-2025'),
    ('FY26_Q3', '01-Oct-2025', '31-Dec-2025'),
    ('FY26_Q4', '01-Jan-2026', '31-Mar-2026'),
    ('FY27_Q1', '01-Apr-2026', '30-Jun-2026'),
    ('FY27_Q2', '01-Jul-2026', '30-Sep-2026'),
]
PERIOD_KEYS = {(f, t) for _, f, t in QUARTERS}
PERIOD_END_TO_Q = {t: q for q, f, t in QUARTERS}  # 'toDate' string -> q_code

_MONTH_END_FULL = {
    'january': ('31-Mar', 'Q4'), 'february': None, 'march': ('31-Mar', 'Q4'),
    'april': ('30-Jun', 'Q1'), 'may': None, 'june': ('30-Jun', 'Q1'),
    'july': ('30-Sep', 'Q2'), 'august': None, 'september': ('30-Sep', 'Q2'),
    'october': ('31-Dec', 'Q3'), 'november': None, 'december': ('31-Dec', 'Q3'),
}
MONTH_END = dict(_MONTH_END_FULL)
for _name, _val in list(_MONTH_END_FULL.items()):
    MONTH_END[_name[:3]] = _val  # also index by 3-letter abbreviation (jun, sep, dec, ...)

print(f"Loading {SRC_RESULTS} ...")
wb1 = openpyxl.load_workbook(SRC_RESULTS, data_only=True, read_only=True)
ws1 = wb1['Financial Results']

target_syms = set()
for sym in n50:
    target_syms.update(SYMBOL_ALIASES.get(sym, [sym]))

headers = None
idx = {}
rows_by_key = {}
for i, row in enumerate(ws1.iter_rows(values_only=True)):
    if i == 0:
        headers = list(row)
        idx = {h: j for j, h in enumerate(headers)}
        continue
    sym = row[idx['symbol']]
    if sym not in target_syms:
        continue
    if row[idx['period']] != 'Quarterly' or row[idx['cumulative']] != 'Non-cumulative':
        continue
    key = (row[idx['fromDate']], row[idx['toDate']])
    if key not in PERIOD_KEYS:
        continue
    bd = row[idx['broadCastDate']]
    if bd is None:
        continue
    try:
        dt = bd if hasattr(bd, 'year') else datetime.strptime(str(bd).split(' ')[0], '%d-%b-%Y')
    except Exception:
        continue
    rows_by_key.setdefault((sym, key), []).append((dt, row[idx['consolidated']]))

result = {sym: {} for sym in n50}
source_tag = {sym: {} for sym in n50}
for sym in n50:
    aliases = SYMBOL_ALIASES.get(sym, [sym])
    for q_code, f, t in QUARTERS:
        candidates = []
        for alias in aliases:
            candidates.extend(rows_by_key.get((alias, (f, t)), []))
        if not candidates:
            continue
        cons = [c for c in candidates if c[1] == 'Consolidated']
        pool = cons if cons else candidates
        best_dt = min(pool, key=lambda c: c[0])[0]
        result[sym][q_code] = best_dt.strftime('%d-%b-%Y')
        source_tag[sym][q_code] = '13_Financial_Results.xlsx'

print(f"Primary source filled {sum(len(v) for v in result.values())} / {len(n50)*len(QUARTERS)} slots.")

# ---- Fallback: 04_Board_Meetings.xlsx for remaining gaps ----
print(f"\nLoading {SRC_BOARD} for gap-fill ...")
wb2 = openpyxl.load_workbook(SRC_BOARD, data_only=True, read_only=True)
ws2 = wb2['Board Meetings']
headers2 = None
idx2 = {}
bm_rows = {}  # (symbol, q_code) -> list of (bm_date_dt, is_results_purpose)
PAT = re.compile(r'period ended\s+([A-Za-z]+)\s*(?:\d{1,2},?\s*)?(\d{4})', re.IGNORECASE)

for i, row in enumerate(ws2.iter_rows(values_only=True)):
    if i == 0:
        headers2 = list(row)
        idx2 = {h: j for j, h in enumerate(headers2)}
        continue
    sym = row[idx2['symbol']]
    if sym not in target_syms:
        continue
    desc = row[idx2['bm_desc']] or ''
    purpose = row[idx2['bm_purpose']] or ''
    m = PAT.search(desc)
    if not m:
        continue
    month_name, year_s = m.group(1).lower(), m.group(2)
    me = MONTH_END.get(month_name)
    if not me:
        continue  # not a quarter-end month -> skip (annual/odd filings)
    end_str = f"{me[0]}-{year_s}"
    q_code = PERIOD_END_TO_Q.get(end_str)
    if not q_code:
        continue
    bm_date = row[idx2['bm_date']]
    try:
        dt = bm_date if hasattr(bm_date, 'year') else datetime.strptime(str(bm_date), '%d-%b-%Y')
    except Exception:
        continue
    is_results = 'financial results' in purpose.lower()
    # resolve alias back to canonical n50 symbol
    canon = sym
    for k, aliases in SYMBOL_ALIASES.items():
        if sym in aliases:
            canon = k
            break
    bm_rows.setdefault((canon, q_code), []).append((dt, is_results))

filled_by_fallback = 0
for (sym, q_code), candidates in bm_rows.items():
    if q_code in result.get(sym, {}):
        continue  # already covered by primary source
    results_only = [c for c in candidates if c[1]]
    pool = results_only if results_only else candidates
    best_dt = min(pool, key=lambda c: c[0])[0]
    result[sym][q_code] = best_dt.strftime('%d-%b-%Y')
    source_tag[sym][q_code] = '04_Board_Meetings.xlsx'
    filled_by_fallback += 1

print(f"Fallback source filled {filled_by_fallback} additional slots.")

total_slots = len(n50) * len(QUARTERS)
filled = sum(1 for sym in n50 for q, _, _ in QUARTERS if q in result[sym])
print(f"\nFinal coverage: {filled} / {total_slots} (symbol, quarter) slots have a real date.")

missing = [(sym, q) for sym in n50 for q, _, _ in QUARTERS if q not in result[sym]]
if missing:
    print(f"\n{len(missing)} still-missing slots:")
    by_sym = {}
    for sym, q in missing:
        by_sym.setdefault(sym, []).append(q)
    for sym, qs in sorted(by_sym.items()):
        print(f"  {sym}: {', '.join(qs)}")
else:
    print("\nFull coverage -- no missing slots.")

out = {
    'generated': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
    'sources': {
        'primary': {'file': SRC_RESULTS, 'description': "NSE XBRL financial-results archive, Quarterly/Non-cumulative filings, earliest broadcastDate per (symbol, reporting-period), Consolidated preferred."},
        'fallback': {'file': SRC_BOARD, 'description': "NSE board-meeting intimations; reporting period parsed from bm_desc free text ('period ended <Month> <Year>'); used only where the primary source has no entry."},
    },
    'scope': 'Nifty 50 (50 symbols), FY23_Q2 through FY27_Q2 (17 quarters)',
    'coverage': f'{filled} / {total_slots}',
    'quarters': [{'q_code': q, 'from': f, 'to': t} for q, f, t in QUARTERS],
    'result_dates': result,
    'source_per_slot': source_tag,
}
out_path = BASE + r'\data_center\nse_real_quarterly_result_dates.json'
json.dump(out, open(out_path, 'w', encoding='utf-8'), indent=2)
print(f"\nSaved: {out_path}")
