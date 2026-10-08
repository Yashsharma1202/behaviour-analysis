"""
NIFTY-50 RECONSTITUTION: WIPRO OUT, BSE IN (per official NSE download,
MW-NIFTY-50-01-Oct-2026.csv, verified exact 1-for-1 swap against the
existing 50-stock universe).

ADDITIVE policy: WIPRO is never deleted anywhere \u2014 its historical quarterly
entries (FY23_Q2 .. FY27_Q2) are untouched (it WAS Nifty-50 then). Only the
CURRENT/upcoming quarter (FY27_Q3) and the 4 upcoming 2026 holiday playbooks
get BSE added and WIPRO's current-quarter is_nifty50 flag flipped to "NO"
(legacy badge, record kept). fo_stocks_211.json's is_nifty50 flag (the single
source the holiday badges read from) was already corrected in a prior step.

BSE data sources used:
  - Holiday events (Gandhi/Dussehra/Diwali/Christmas): holidays_dataset.json's
    existing 4-year (2022-2025) F&O-211 backtest for BSE \u2014 the SAME source
    already used for the JIOFIN/BEL Dussehra fallback entries earlier this
    session (verified identical values), not a thin/placeholder number.
  - Q3 quarterly earnings win rate: computed fresh here, split-robust T-2/T+4
    window over BSE's 7 own announced result dates with real price data
    (2025-02-06 .. 2026-08-04, from fo_futures_near_month_continuous.parquet,
    BSE listed on F&O since 2024-11-29) \u2014 n=7, LONG 85.7% WR, avg +5.06%.
  - Spot price: 3,094.80 (today's live LTP from the NSE CSV), not the stale
    3,530.30 sitting in holidays_dataset.json.
"""
import json
from datetime import date, timedelta

BASE = r'D:\behaviour analysis'
SPOT = 3094.80
LOT = 500
MARGIN = round(SPOT * LOT * 0.20, 2)
SECTOR = 'Capital Markets & Exchanges'

NSE_2026_HOLIDAYS = {
    # '2026-03-04' removed (not a real NSE holiday -- leftover from Holi's date
    # once being wrongly typed as 04-Mar; the real Holi holiday is 03-Mar, kept
    # below). '2026-11-08'/'2026-11-10' (Diwali) added -- were missing entirely.
    '2026-01-26', '2026-03-03', '2026-03-26', '2026-03-31', '2026-04-03',
    '2026-04-14', '2026-05-01', '2026-05-28', '2026-06-26', '2026-09-14', '2026-10-02',
    '2026-10-20', '2026-11-08', '2026-11-10', '2026-11-24', '2026-12-25',
}
def is_trading_day(d):
    return d.weekday() < 5 and d.isoformat() not in NSE_2026_HOLIDAYS
def shift_trading_days(anchor, n):
    d = anchor
    step = 1 if n >= 0 else -1
    cnt = abs(n)
    while cnt > 0:
        d += timedelta(days=step)
        if is_trading_day(d): cnt -= 1
    return d

TODAY = date(2026, 10, 1)

path = BASE + r'\dashboard_data\event_dashboard_data.json'
d = json.load(open(path, encoding='utf-8'))

# ---------------- 1. FY27_Q3 quarterly table ----------------
q3 = next(q for q in d['quarters'] if q['q_code'] == 'FY27_Q3')
wipro = next(s for s in q3['stocks'] if s['symbol'] == 'WIPRO')
wipro['is_nifty50'] = 'NO'
wipro['legacy_note'] = 'No longer a Nifty-50 constituent as of 01-Oct-2026 reconstitution (replaced by BSE). Entry kept for F&O-universe visibility, excluded from Nifty-50-only views.'

bse_q3 = {
    "symbol": "BSE", "name": "BSE Ltd", "is_nifty50": "YES",
    "spot_ltp": SPOT, "lot_size": LOT, "margin_20pct": MARGIN,
    "direction": "LONG", "window": "T-2 to T+4", "lead": 2, "hold": 4,
    "result_date": "Yet to come", "result_declaration_date": "Yet to come",
    "entry_date": "Yet to come", "exit_date": "Yet to come",
    "announced": False, "result_status": "Yet to come",
    "purpose": "Awaiting NSE Announcement", "source": "Awaiting NSE Announcement",
    "entry_px": SPOT, "exit_px": None, "ltp": SPOT, "lot": LOT, "margin": MARGIN,
    "eq_qty_1l": max(1, int(100000 / SPOT)),
    "expected_ret": 5.06, "expected_return": 5.06,
    "actual_ret": None, "actual_return": None, "actual_pnl": None, "pnl": None,
    "exp_pnl_1lot": round(5.06 / 100 * SPOT * LOT, 2),
    "outcome": "PENDING",
    "avg_17q_wr": 85.7, "best_strategy": "Pre-Quarterly Run-up (LONG)",
    "option_play": "1% ITM CALL / PUT Option (30% SL)", "sector": SECTOR,
    "data_basis_note": "Win rate from n=7 real announced-result quarters (2025-02-06 to 2026-08-04, split-robust T-2/T+4 backtest on BSE's own F&O futures history since its 2024-11-29 listing) \u2014 not the full 17-quarter depth other stocks have; added for the 01-Oct-2026 Nifty-50 reconstitution (replaces WIPRO).",
}

# Replace an existing BSE record instead of blindly appending -- this script
# used to append unconditionally (and increment pending_count unconditionally)
# on every rerun, which had already quadrupled the BSE entry here before this
# fix (4 identical records found).
q3['stocks'] = [s for s in q3['stocks'] if s.get('symbol') != 'BSE']
q3['stocks'].append(bse_q3)
q3['announced_count'] = sum(1 for s in q3['stocks'] if s.get('announced'))
q3['pending_count'] = len(q3['stocks']) - q3['announced_count']
print(f"FY27_Q3: WIPRO marked legacy, BSE added. Total stocks now: {len(q3['stocks'])} "
      f"(announced={q3['announced_count']}, pending={q3['pending_count']})")

# ---------------- 2. Holiday playbooks ----------------
hds = json.load(open(BASE + r'\dashboard_data\holidays_dataset.json', encoding='utf-8'))
hds_by_id = {h['id']: h for h in hds}

HOLIDAY_MAP = {
    'gandhi_2026':    ('gandhi',    date(2026, 10, 2)),
    'dussehra_2026':  ('dussehra',  date(2026, 10, 20)),
    # was hand-typed as 21-Oct-2026 (2025's Diwali date, shifted); NSE's live
    # holiday calendar confirms Diwali (Laxmi Pujan) 2026 is 08-Nov-2026.
    'diwali_2026':    ('diwali',    date(2026, 11, 8)),
    'christmas_2026': ('christmas', date(2026, 12, 25)),
}

for hol_id, (hds_key, anchor) in HOLIDAY_MAP.items():
    hol = next(h for h in d['holidays'] if h['id'] == hol_id)
    bse_stat = next(s for s in hds_by_id[hds_key]['top_stocks'] if s['symbol'] == 'BSE')
    lead, hold = bse_stat['entry_lead_days'], bse_stat['exit_hold_days']
    direction = bse_stat['direction']
    wr = bse_stat['win_rate']

    en = shift_trading_days(anchor, -lead)
    ex = shift_trading_days(anchor, hold)
    entry_passed = en < TODAY

    stock_rec = {
        "symbol": "BSE", "name": "BSE Ltd", "direction": direction,
        "window": f"T-{lead} to T+{hold}", "lead": lead, "hold": hold,
        "entry_date": en.strftime('%d-%b-%Y (%a)'), "exit_date": ex.strftime('%d-%b-%Y (%a)'),
        "total_years": 4, "full_19y_wr": wr, "full_19y_avg_ret": None,
        "last_4y_wr": wr, "last_4y_avg_ret": None,
        "lot": LOT, "margin": MARGIN,
        "fut_action": f"{'BUY' if direction=='LONG' else 'SELL'} FUTURES ({'LONG' if direction=='LONG' else 'SHORT'} {LOT} qty)",
        "opt_action": f"BUY {'CALL' if direction=='LONG' else 'PUT'} OPTION",
        "sector": SECTOR,
        "direction_basis": f"Real 4yr F&O-211 backtest (n=4, {wr}% WR, 2022-2025 PNL series) \u2014 same data source/tier already used for the JIOFIN/BEL Dussehra fallback entries. Avg return %% left blank (source gives cumulative PNL per year, not a clean per-trade %% without each year's trade price) \u2014 win rate and direction are the real, usable numbers. Added: current Nifty-50 constituent (01-Oct-2026 reconstitution, replaces WIPRO).",
    }
    if entry_passed:
        stock_rec["position_status"] = "ENTRY WINDOW ALREADY PASSED (missed) \u2014 newly added stock, was not in the playbook when its own entry date occurred"

    # Replace an existing BSE record instead of blindly appending -- this
    # script used to append unconditionally, which duplicated BSE on every
    # rerun (found: 2 identical BSE entries in each of the 4 holidays).
    hol['stocks'] = [s for s in hol['stocks'] if s.get('symbol') != 'BSE']
    hol['stocks'].append(stock_rec)
    hol['stocks_count'] = len(hol['stocks'])
    print(f"{hol_id}: BSE added ({direction}, {wr}% WR, entry {en}, exit {ex}){' [PASSED]' if entry_passed else ''}. Total stocks now: {hol['stocks_count']}")

json.dump(d, open(path, 'w', encoding='utf-8'), indent=2)
print("\nSaved:", path)
