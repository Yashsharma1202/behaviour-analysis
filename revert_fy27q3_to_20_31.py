import json

PATH = "dashboard_data/event_dashboard_data.json"

# Exact pre-fix values (captured from commit f821bdf8 before the earlier revert).
REVERT = {
    "WIPRO":      dict(direction="LONG",  window="T-3 to T+3", lead=3, hold=3, entry_date="12-Oct-2026 (Mon)", exit_date="21-Oct-2026 (Wed)"),
    "HINDALCO":   dict(direction="LONG",  window="T-7 to T+5", lead=7, hold=5, entry_date="28-Oct-2026 (Wed)", exit_date="16-Nov-2026 (Mon)"),
    "BAJAJFINSV": dict(direction="LONG",  window="T-2 to T+3", lead=2, hold=3, entry_date="30-Oct-2026 (Fri)", exit_date="06-Nov-2026 (Fri)"),
    "CIPLA":      dict(direction="LONG",  window="T-1 to T+8", lead=1, hold=8, entry_date="26-Oct-2026 (Mon)", exit_date="06-Nov-2026 (Fri)"),
    "HDFCLIFE":   dict(direction="SHORT", window="T-4 to T+8", lead=4, hold=8, entry_date="09-Oct-2026 (Fri)", exit_date="28-Oct-2026 (Wed)"),
    "ICICIBANK":  dict(direction="LONG",  window="T-6 to T+2", lead=6, hold=2, entry_date="09-Oct-2026 (Fri)", exit_date="21-Oct-2026 (Wed)"),
    "HDFCBANK":   dict(direction="LONG",  window="T-2 to T+3", lead=2, hold=3, entry_date="15-Oct-2026 (Thu)", exit_date="22-Oct-2026 (Thu)"),
    "ADANIPORTS": dict(direction="LONG",  window="T-1 to T+1", lead=1, hold=1, entry_date="27-Oct-2026 (Tue)", exit_date="29-Oct-2026 (Thu)"),
    "NESTLEIND":  dict(direction="LONG",  window="T-3 to T+4", lead=3, hold=4, entry_date="12-Oct-2026 (Mon)", exit_date="22-Oct-2026 (Thu)"),
    "INFY":       dict(direction="SHORT", window="T-6 to T+1", lead=6, hold=1, entry_date="14-Oct-2026 (Wed)", exit_date="26-Oct-2026 (Mon)"),
    "JSWSTEEL":   dict(direction="LONG",  window="T-1 to T+8", lead=1, hold=8, entry_date="22-Oct-2026 (Thu)", exit_date="04-Nov-2026 (Wed)"),
    "TECHM":      dict(direction="LONG",  window="T-6 to T+8", lead=6, hold=8, entry_date="07-Oct-2026 (Wed)", exit_date="28-Oct-2026 (Wed)"),
    "M&M":        dict(direction="LONG",  window="T-8 to T+1", lead=8, hold=1, entry_date="26-Oct-2026 (Mon)", exit_date="06-Nov-2026 (Fri)"),
    "DRREDDY":    dict(direction="LONG",  window="T-3 to T+4", lead=3, hold=4, entry_date="19-Oct-2026 (Mon)", exit_date="29-Oct-2026 (Thu)"),
    "ASIANPAINT": dict(direction="SHORT", window="T-1 to T+2", lead=1, hold=2, entry_date="28-Oct-2026 (Wed)", exit_date="02-Nov-2026 (Mon)"),
    "ULTRACEMCO": dict(direction="LONG",  window="T-6 to T+7", lead=6, hold=7, entry_date="09-Oct-2026 (Fri)", exit_date="29-Oct-2026 (Thu)"),
    "HINDUNILVR": dict(direction="SHORT", window="T-1 to T+2", lead=1, hold=2, entry_date="27-Oct-2026 (Tue)", exit_date="30-Oct-2026 (Fri)"),
    "MARUTI":     dict(direction="LONG",  window="T-7 to T+7", lead=7, hold=7, entry_date="15-Oct-2026 (Thu)", exit_date="05-Nov-2026 (Thu)"),
    "BAJAJ-AUTO": dict(direction="LONG",  window="T-5 to T+1", lead=5, hold=1, entry_date="19-Oct-2026 (Mon)", exit_date="28-Oct-2026 (Wed)"),
    "BAJFINANCE": dict(direction="LONG",  window="T-4 to T+7", lead=4, hold=7, entry_date="13-Oct-2026 (Tue)", exit_date="29-Oct-2026 (Thu)"),
}

with open(PATH, encoding="utf-8") as f:
    data = json.load(f)

q3 = next(q for q in data["quarters"] if q["q_code"] == "FY27_Q3")
q3["bias"] = "20 SHORT / 31 LONG"

n = 0
for s in q3["stocks"]:
    r = REVERT.get(s["symbol"])
    if not r:
        continue
    s["direction"] = r["direction"]
    s["window"] = r["window"]
    s["lead"] = r["lead"]
    s["hold"] = r["hold"]
    s["entry_date"] = r["entry_date"]
    s["exit_date"] = r["exit_date"]
    s["entry_lead_days"] = r["lead"]
    s["exit_hold_days"] = r["hold"]
    s["taking_window_raw"] = r["window"]
    s["pos_window"] = r["window"]
    s["taking_window_full"] = f"Entry: {r['entry_date']} ({r['window']}) ➔ Exit: {r['exit_date']}"
    s.pop("date_fix_note", None)
    n += 1

with open(PATH, "w", encoding="utf-8") as f:
    json.dump(data, f, indent=2, ensure_ascii=False)
    f.write("\n")

print(f"Reverted {n} stocks; bias set to 20 SHORT / 31 LONG")
