Retired 2026-07-27.

fetch_fii_bulk_block.py       -> feeds 33/34/35, now in scrap.py
fetch_institutional_research.py -> feeds 36/37/38, now in scrap.py

Both wrote to the SAME files as scrap.py but were never scheduled: the
Task Scheduler job runs run_daily_scraper.bat, which only calls scrap.py.

They were not harmless dead code. For feeds 33/34/35 they carried the CORRECT
dedupe keys while the live path in scrap.py carried none:

    feed            scrap.py (scheduled)   helper (never ran)
    33_FII_DII      -- none --             ["date","category"]
    34_Bulk_Deals   -- none --             ["Date","Symbol","Client Name","Quantity Traded"]
    35_Block_Deals  -- none --             same

So the correct implementation was the one that never executed. Those keys have
been folded into DEDUPE_KEYS in scrap.py; these files are kept here only for
reference and can be deleted once you are happy with a few scheduled runs.
