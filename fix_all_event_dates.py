"""
Script to fix incorrect event dates across all JSON, JS, and HTML dashboard files.
Primary fix:
- Diwali (Laxmi Pujan & Muhurat) 2026 date fixed from 21-Oct-2026 (Wed) [which was 2025's date] to 08-Nov-2026 (Sunday).
- T-8 First Entry Trigger for Diwali 2026 updated to Wednesday, 28-Oct-2026.
- Holi 2026 date fixed from 14-Mar-2026 (Sat) to 04-Mar-2026 (Wed).
"""

import json
import os
import re

BASE_DIR = r"D:\behaviour analysis"

# 1. Update nifty_futures_holiday_behaviour.json & docs
h_json_path = os.path.join(BASE_DIR, "dashboard_data", "nifty_futures_holiday_behaviour.json")
docs_h_json_path = os.path.join(BASE_DIR, "docs", "dashboard_data", "nifty_futures_holiday_behaviour.json")

with open(h_json_path, 'r', encoding='utf-8') as f:
    h_data = json.load(f)

for h in h_data.get('holidays', []):
    if h.get('id') == 'diwali':
        h['2026_date_str'] = '08-Nov-2026 (Sun)'
        h['upcoming_entry'] = '28-Oct-2026 (Wed)'
        h['upcoming_exit'] = '16-Nov-2026 (Mon)'
    elif h.get('id') == 'holi':
        h['2026_date_str'] = '04-Mar-2026 (Wed)'
        h['upcoming_entry'] = '23-Feb-2026 (Mon)'
        h['upcoming_exit'] = '11-Mar-2026 (Wed)'

with open(h_json_path, 'w', encoding='utf-8') as f:
    json.dump(h_data, f, indent=2)

with open(docs_h_json_path, 'w', encoding='utf-8') as f:
    json.dump(h_data, f, indent=2)

print("Updated nifty_futures_holiday_behaviour.json in main & docs!")

# 2. Update nifty_futures_holiday_behaviour.js & docs
js_path = os.path.join(BASE_DIR, "dashboard_data", "nifty_futures_holiday_behaviour.js")
docs_js_path = os.path.join(BASE_DIR, "docs", "dashboard_data", "nifty_futures_holiday_behaviour.js")

with open(js_path, 'w', encoding='utf-8') as f:
    f.write(f"window.NIFTY_HOLIDAY_BEHAVIOUR_DATA = {json.dumps(h_data)};")

with open(docs_js_path, 'w', encoding='utf-8') as f:
    f.write(f"window.NIFTY_HOLIDAY_BEHAVIOUR_DATA = {json.dumps(h_data)};")

print("Updated nifty_futures_holiday_behaviour.js in main & docs!")

# 3. Update event_dashboard_data.json & docs
e_json_path = os.path.join(BASE_DIR, "dashboard_data", "event_dashboard_data.json")
docs_e_json_path = os.path.join(BASE_DIR, "docs", "dashboard_data", "event_dashboard_data.json")

if os.path.exists(e_json_path):
    with open(e_json_path, 'r', encoding='utf-8') as f:
        e_data = json.load(f)
    
    for h in e_data.get('holidays', []):
        if h.get('id') == 'diwali_2026':
            h['date_str'] = '08-Nov-2026 (Sunday)'
            h['opt_month'] = 'NOV'
            for s in h.get('stocks', []):
                # Update entry date strings if based on 21-Oct
                if 'Oct' in s.get('entry_date_str', ''):
                    s['entry_date_str'] = s['entry_date_str'].replace('21-Oct', '08-Nov').replace('Oct', 'Nov')
        elif h.get('id') == 'holi_2026':
            h['date_str'] = '04-Mar-2026 (Wednesday)'
            h['opt_month'] = 'MAR'
            
    with open(e_json_path, 'w', encoding='utf-8') as f:
        json.dump(e_data, f, indent=2)
        
    with open(docs_e_json_path, 'w', encoding='utf-8') as f:
        json.dump(e_data, f, indent=2)

# Update event_dashboard_data.js & docs
e_js_path = os.path.join(BASE_DIR, "dashboard_data", "event_dashboard_data.js")
docs_e_js_path = os.path.join(BASE_DIR, "docs", "dashboard_data", "event_dashboard_data.js")

if os.path.exists(e_json_path):
    with open(e_json_path, 'r', encoding='utf-8') as f:
        e_data = json.load(f)
    with open(e_js_path, 'w', encoding='utf-8') as f:
        f.write(f"window.EVENT_DASHBOARD_DATA = {json.dumps(e_data)};")
    with open(docs_e_js_path, 'w', encoding='utf-8') as f:
        f.write(f"window.EVENT_DASHBOARD_DATA = {json.dumps(e_data)};")

print("Updated event_dashboard_data.json/js in main & docs!")

# 4. Update event_trading_dashboard.html & docs
html_path = os.path.join(BASE_DIR, "event_trading_dashboard.html")
docs_html_path = os.path.join(BASE_DIR, "docs", "event_trading_dashboard.html")

with open(html_path, 'r', encoding='utf-8') as f:
    html = f.read()

# Replace Diwali row details in table
old_diwali_snippet = """<tr class="event-master-row" id="hrow-diwali_2026" onclick="selectHolidayRow('diwali_2026', this)">
                <td><strong>Diwali (Laxmi Pujan & Muhurat)</strong></td>
                <td><span class="date-badge">21-Oct-2026 (Wednesday)</span></td>
                <td>Festival Holiday</td>
                <td><span style="color:#94A3B8; font-size:0.75rem;">UPCOMING IN OCT</span></td>
                <td><strong style="color:#38BDF8;">Friday, 09-Oct (T-8)</strong></td>
                <td><span class="badge-long">50 LONG</span></td>
                <td><button class="btn-select-event" onclick="event.stopPropagation(); selectHolidayRow('diwali_2026', this.closest('tr'))">View Playbook</button></td>
              </tr>"""

new_diwali_snippet = """<tr class="event-master-row" id="hrow-diwali_2026" onclick="selectHolidayRow('diwali_2026', this)">
                <td><strong>Diwali (Laxmi Pujan & Muhurat)</strong></td>
                <td><span class="date-badge">08-Nov-2026 (Sunday)</span></td>
                <td>Festival Holiday</td>
                <td><span style="color:#94A3B8; font-size:0.75rem;">UPCOMING IN NOV</span></td>
                <td><strong style="color:#38BDF8;">Wednesday, 28-Oct (T-8)</strong></td>
                <td><span class="badge-long">50 LONG</span></td>
                <td><button class="btn-select-event" onclick="event.stopPropagation(); selectHolidayRow('diwali_2026', this.closest('tr'))">View Playbook</button></td>
              </tr>"""

if old_diwali_snippet in html:
    html = html.replace(old_diwali_snippet, new_diwali_snippet)
    print("Replaced Diwali snippet in HTML!")
else:
    # Regex fallback replacement
    html = re.sub(
        r'21-Oct-2026 \(Wednesday\)',
        '08-Nov-2026 (Sunday)',
        html
    )
    html = re.sub(
        r'Friday, 09-Oct \(T-8\)',
        'Wednesday, 28-Oct (T-8)',
        html
    )

with open(html_path, 'w', encoding='utf-8') as f:
    f.write(html)

with open(docs_html_path, 'w', encoding='utf-8') as f:
    f.write(html)

print("Updated HTML files successfully!")
