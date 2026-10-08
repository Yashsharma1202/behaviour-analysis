"""
rbi_mpc_calendar.py
===============================================================================
Single source of truth for RBI Monetary Policy Committee (MPC) decision dates
-- past, present and future. Every other script in this repo that needs an
RBI date should import from here instead of typing a date into its own copy
of a list.

Why this exists: this repo used to have the same RBI_POLICY_EVENTS list
hand-typed into several scripts (get_rbi_nifty50_table.py, enrich_rbi_json.py,
sync_live_corporate_feeds.py, build_rbi_mpc_detailed_tradelog_master.py,
build_rbi_policy_combined_master_excel.py). Two of those copies had the
Oct-2026 decision date wrong (09-Oct instead of the real 07-Oct), the Aug-2025
date wrong (08-Aug instead of the real 06-Aug), and all of them were missing
six meetings (Oct-2025 through Aug-2026) that had already happened.

Confirmed dates are pulled from RBI's own press-release RSS feed:
    https://www.rbi.org.in/pressreleases_rss.xml
by matching the "... Resolution of the Monetary Policy Committee ..." title
RBI publishes the moment a decision is announced, and reading its pubDate.
That is the actual decision date, not the pre-announced schedule -- the two
can differ (Aug-2025 was scheduled for 5-7 Aug but the meeting was actually
held 4-6 Aug, decision on 6-Aug).

The RSS feed only carries the ~10 most recent releases, so it cannot see
further back than that. The seed list below fills in everything older,
each date verified against RBI's own resolution / governor's statement
(not just the schedule) as of 2026-10-07. The seed is a one-time backfill,
not a replacement for fetching: refresh_from_rbi() is still the preferred
path and will keep the cache current without anyone editing this file again.
The one not-yet-held meeting in the seed is marked "scheduled" (from RBI's
official FY meeting-schedule press release) and will be promoted to
"confirmed" -- with its date corrected if it shifted -- the first time
refresh_from_rbi() sees the matching resolution in the live feed.

Everything here is additive: refresh_from_rbi() only adds new records or
promotes "scheduled" -> "confirmed"; it never deletes a confirmed one.

Usage:
    from rbi_mpc_calendar import refresh_from_rbi, load_calendar, get_next_meeting, get_most_recent_meeting

    refresh_from_rbi()           # pulls the live RSS feed, updates processed/rbi_mpc_calendar.json
    events = load_calendar()     # list of dicts, oldest first
    nxt = get_next_meeting()     # next meeting on/after today (falls back to the most recent if none)

Run directly (`python rbi_mpc_calendar.py`) to refresh and print the calendar.
===============================================================================
"""
import json
import pathlib
import re
import urllib.request

import pandas as pd

ROOT = pathlib.Path(__file__).resolve().parent
PROC = ROOT / 'processed'
PROC.mkdir(parents=True, exist_ok=True)
CALENDAR_PATH = PROC / 'rbi_mpc_calendar.json'

RSS_URL = 'https://www.rbi.org.in/pressreleases_rss.xml'
_UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'

# One-time verified backfill -- see module docstring. Do not add future
# meetings here by hand; let refresh_from_rbi() pick them up, or at most
# mark a not-yet-held one "scheduled" from RBI's own schedule press release.
_SEED_EVENTS = [
    {"code": "RBI_Oct_2023", "label": "RBI MPC Policy Oct 2023", "date_str": "2023-10-06", "period": "FY 2023-24 Q3 Policy", "status": "confirmed", "ann": "Repo Rate Decision & Stance"},
    {"code": "RBI_Dec_2023", "label": "RBI MPC Policy Dec 2023", "date_str": "2023-12-08", "period": "FY 2023-24 Q3 Policy", "status": "confirmed", "ann": "Repo Rate Decision & Stance"},
    {"code": "RBI_Feb_2024", "label": "RBI MPC Policy Feb 2024", "date_str": "2024-02-08", "period": "FY 2023-24 Q4 Policy", "status": "confirmed", "ann": "Repo Rate Decision & Stance"},
    {"code": "RBI_Apr_2024", "label": "RBI MPC Policy Apr 2024", "date_str": "2024-04-05", "period": "FY 2024-25 Q1 Policy", "status": "confirmed", "ann": "Repo Rate Decision & Stance"},
    {"code": "RBI_Jun_2024", "label": "RBI MPC Policy Jun 2024", "date_str": "2024-06-07", "period": "FY 2024-25 Q1 Policy", "status": "confirmed", "ann": "Repo Rate Decision & Stance"},
    {"code": "RBI_Aug_2024", "label": "RBI MPC Policy Aug 2024", "date_str": "2024-08-08", "period": "FY 2024-25 Q2 Policy", "status": "confirmed", "ann": "Repo Rate Decision & Stance"},
    {"code": "RBI_Oct_2024", "label": "RBI MPC Policy Oct 2024", "date_str": "2024-10-09", "period": "FY 2024-25 Q3 Policy", "status": "confirmed", "ann": "Repo Rate Decision & Stance"},
    {"code": "RBI_Dec_2024", "label": "RBI MPC Policy Dec 2024", "date_str": "2024-12-06", "period": "FY 2024-25 Q3 Policy", "status": "confirmed", "ann": "Repo Rate Decision & Stance"},
    {"code": "RBI_Feb_2025", "label": "RBI MPC Policy Feb 2025", "date_str": "2025-02-07", "period": "FY 2024-25 Q4 Policy", "status": "confirmed", "ann": "Repo Rate Cut Announcement"},
    {"code": "RBI_Apr_2025", "label": "RBI MPC Policy Apr 2025", "date_str": "2025-04-09", "period": "FY 2025-26 Q1 Policy", "status": "confirmed", "ann": "Repo Rate Decision & Stance"},
    {"code": "RBI_Jun_2025", "label": "RBI MPC Policy Jun 2025", "date_str": "2025-06-06", "period": "FY 2025-26 Q1 Policy", "status": "confirmed", "ann": "Repo Rate Decision & Stance"},
    # was hand-typed as 2025-08-08 across the repo. Actual meeting: 4-6 Aug 2025
    # (56th MPC meeting), decision 2025-08-06. Fixed here.
    {"code": "RBI_Aug_2025", "label": "RBI MPC Policy Aug 2025", "date_str": "2025-08-06", "period": "FY 2025-26 Q2 Policy", "status": "confirmed", "ann": "Repo Rate Decision & Stance"},
    # missing from every hand-typed list in the repo. Added here.
    {"code": "RBI_Oct_2025", "label": "RBI MPC Policy Oct 2025", "date_str": "2025-10-01", "period": "FY 2025-26 Q3 Policy", "status": "confirmed", "ann": "Repo Rate Decision & Stance"},
    {"code": "RBI_Dec_2025", "label": "RBI MPC Policy Dec 2025", "date_str": "2025-12-05", "period": "FY 2025-26 Q3 Policy", "status": "confirmed", "ann": "Repo Rate Decision & Stance"},
    {"code": "RBI_Feb_2026", "label": "RBI MPC Policy Feb 2026", "date_str": "2026-02-06", "period": "FY 2025-26 Q4 Policy", "status": "confirmed", "ann": "Repo Rate Decision & Stance"},
    {"code": "RBI_Apr_2026", "label": "RBI MPC Policy Apr 2026", "date_str": "2026-04-08", "period": "FY 2026-27 Q1 Policy", "status": "confirmed", "ann": "Repo Rate Decision & Stance"},
    {"code": "RBI_Jun_2026", "label": "RBI MPC Policy Jun 2026", "date_str": "2026-06-05", "period": "FY 2026-27 Q1 Policy", "status": "confirmed", "ann": "Repo Rate Decision & Stance"},
    {"code": "RBI_Aug_2026", "label": "RBI MPC Policy Aug 2026", "date_str": "2026-08-05", "period": "FY 2026-27 Q2 Policy", "status": "confirmed", "ann": "Repo Rate Decision & Stance"},
    # was hand-typed as 2026-10-09 across the repo (meeting announced for 5-7 Oct;
    # decision actually fell on 2026-10-07, the meeting's real last day). Fixed here.
    {"code": "RBI_Oct_2026", "label": "RBI MPC Policy Oct 2026", "date_str": "2026-10-07", "period": "FY 2026-27 Q3 Policy", "status": "confirmed", "ann": "Repo Rate Decision & Stance"},
    # Not yet held. RBI's official FY2026-27 schedule (press release 2025-2026/2306,
    # 23-Mar-2026) says 2-4 Dec 2026. The decision date can shift by a day or two
    # (see Aug-2025 above) -- refresh_from_rbi() will promote this to "confirmed"
    # and correct the date the moment RBI's own resolution appears in the feed.
    {"code": "RBI_Dec_2026", "label": "RBI MPC Policy Dec 2026", "date_str": "2026-12-04", "period": "FY 2026-27 Q3 Policy", "status": "scheduled", "ann": "Repo Rate Decision & Stance"},
    {"code": "RBI_Feb_2027", "label": "RBI MPC Policy Feb 2027", "date_str": "2027-02-05", "period": "FY 2026-27 Q4 Policy", "status": "scheduled", "ann": "Repo Rate Decision & Stance"},
]


def _period_for(dt):
    fy_start = dt.year if dt.month >= 4 else dt.year - 1
    q = {4: 1, 5: 1, 6: 1, 7: 2, 8: 2, 9: 2, 10: 3, 11: 3, 12: 3, 1: 4, 2: 4, 3: 4}[dt.month]
    return f"FY {fy_start}-{str(fy_start + 1)[2:]} Q{q} Policy"


def load_calendar():
    """Load the cached calendar, seeding it from _SEED_EVENTS on first use."""
    if CALENDAR_PATH.exists():
        try:
            events = json.loads(CALENDAR_PATH.read_text(encoding='utf-8'))
            if events:
                return sorted(events, key=lambda e: e['date_str'])
        except Exception:
            pass
    return sorted(_SEED_EVENTS, key=lambda e: e['date_str'])


def _save(events_by_code):
    events = sorted(events_by_code.values(), key=lambda e: e['date_str'])
    CALENDAR_PATH.write_text(json.dumps(events, indent=2), encoding='utf-8')
    return events


def refresh_from_rbi(timeout=15, verbose=True):
    """Pull RBI's live press-release RSS feed and reconcile it against the
    cached calendar. Any item titled '... Resolution of the Monetary Policy
    Committee ...' confirms/corrects that meeting's decision date (its
    pubDate). Never deletes an existing record; only adds or corrects."""
    events = {e['code']: e for e in load_calendar()}
    try:
        req = urllib.request.Request(RSS_URL, headers={'User-Agent': _UA})
        body = urllib.request.urlopen(req, timeout=timeout).read().decode('utf-8', errors='ignore')
        items = re.findall(r'<item>(.*?)</item>', body, re.S)
        found = 0
        for it in items:
            title_m = re.search(r'<title>(?:<!\[CDATA\[)?(.*?)(?:\]\]>)?</title>', it, re.S)
            pub_m = re.search(r'<pubDate>(.*?)</pubDate>', it, re.S)
            if not title_m or not pub_m:
                continue
            title = title_m.group(1).strip()
            if 'Resolution of the Monetary Policy Committee' not in title:
                continue
            decision_dt = pd.to_datetime(pub_m.group(1), errors='coerce')
            if pd.isna(decision_dt):
                continue
            decision_dt = decision_dt.normalize()
            code = f"RBI_{decision_dt.strftime('%b_%Y')}"
            events[code] = {
                "code": code,
                "label": f"RBI MPC Policy {decision_dt.strftime('%b %Y')}",
                "date_str": decision_dt.strftime('%Y-%m-%d'),
                "period": _period_for(decision_dt),
                "status": "confirmed",
                "ann": "Repo Rate Decision & Stance",
                "source": "RBI pressreleases_rss.xml (live)",
            }
            found += 1
        if verbose:
            print(f"[rbi_mpc_calendar] RSS scan: {found} MPC resolution(s) matched in the live feed.")
    except Exception as e:
        if verbose:
            print(f"[rbi_mpc_calendar] WARNING: could not reach RBI's RSS feed ({e}); keeping cached calendar as-is.")
    return _save(events)


def get_next_meeting(today=None):
    """The next meeting on/after `today` (confirmed or scheduled); falls back
    to the most recent past meeting if nothing upcoming is cached."""
    today = pd.Timestamp(today) if today is not None else pd.Timestamp.now().normalize()
    events = load_calendar()
    upcoming = [e for e in events if pd.to_datetime(e['date_str']) >= today]
    if upcoming:
        return upcoming[0]
    return events[-1] if events else None


def get_most_recent_meeting(today=None):
    """The most recent meeting on/before `today`."""
    today = pd.Timestamp(today) if today is not None else pd.Timestamp.now().normalize()
    events = [e for e in load_calendar() if pd.to_datetime(e['date_str']) <= today]
    return events[-1] if events else None


if __name__ == '__main__':
    refresh_from_rbi()
    print()
    for e in load_calendar():
        print(f"{e['date_str']}  {e['label']:26} {e['period']:22} [{e['status']}]")
