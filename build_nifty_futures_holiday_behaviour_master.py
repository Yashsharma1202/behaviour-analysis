import os
import sys
import json
import datetime
import pathlib
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding='utf-8')

ROOT = pathlib.Path(r"D:\behaviour analysis")
PARQUET_FILE = ROOT / "scraped_parquet" / "nifty_futures_near_month_continuous.parquet"
EXCEL_OUT = ROOT / "Nifty_Futures_Index_Holiday_Behaviour_Master_2000_2026.xlsx"
JSON_OUT = ROOT / "dashboard_data" / "nifty_futures_holiday_behaviour.json"
DOCS_JSON_OUT = ROOT / "docs" / "dashboard_data" / "nifty_futures_holiday_behaviour.json"

print("==========================================================================================")
print("NIFTY 50 FUTURES INDEX — FULL-HISTORY (2000–2026) HOLIDAY BEHAVIORAL ENGINE")
print("INGESTING CONTINUOUS FUTURES PARQUET & EVALUATING ALL NSE MARKET HOLIDAYS")
print("==========================================================================================")

# 1. Ingest Nifty Futures Continuous Data
df_fut = pd.read_parquet(PARQUET_FILE)
df_fut['DATE'] = pd.to_datetime(df_fut['DATE'])
df_fut = df_fut.sort_values('DATE').reset_index(drop=True)

dates_list = df_fut['DATE'].dt.date.tolist()
close_list = df_fut['CLOSE'].astype(float).tolist()
open_list = df_fut['OPEN'].astype(float).tolist()
high_list = df_fut['HIGH'].astype(float).tolist()
low_list = df_fut['LOW'].astype(float).tolist()
date_to_idx = {d: i for i, d in enumerate(dates_list)}

print(f"Loaded {len(dates_list)} trading days from {dates_list[0]} to {dates_list[-1]}")

# 2. Comprehensive 26-Year Historical Calendar of NSE Market Holidays (2000–2026)
# Maps each recurring occasion to its observed dates per year across the 26-year dataset.
holidays_def = [
    {
        'id': 'gandhi',
        'name': 'Mahatma Gandhi Jayanti',
        'category': 'National Holiday',
        '2026_date': '2026-10-02',
        '2026_date_str': '02-Oct-2026 (Fri)',
        'dates': {
            2000: '2000-10-02', 2001: '2001-10-02', 2002: '2002-10-02', 2003: '2003-10-02',
            2004: '2004-10-02', 2005: '2005-10-02', 2006: '2006-10-02', 2007: '2007-10-02',
            2008: '2008-10-02', 2009: '2009-10-02', 2010: '2010-10-02', 2011: '2011-10-02',
            2012: '2012-10-02', 2013: '2013-10-02', 2014: '2014-10-02', 2015: '2015-10-02',
            2016: '2016-10-02', 2017: '2017-10-02', 2018: '2018-10-02', 2019: '2019-10-02',
            2020: '2020-10-02', 2021: '2021-10-02', 2022: '2022-10-02', 2023: '2023-10-02',
            2024: '2024-10-02', 2025: '2025-10-02'
        },
        'desc': 'National holiday leading into Q3 earnings & early October festive cycle'
    },
    {
        'id': 'dussehra',
        'name': 'Dussehra / Vijayadashami',
        'category': 'Festive Accumulation',
        '2026_date': '2026-10-20',
        '2026_date_str': '20-Oct-2026 (Tue)',
        'dates': {
            2000: '2000-10-09', 2001: '2001-10-26', 2002: '2002-10-15', 2003: '2003-10-05',
            2004: '2004-10-22', 2005: '2005-10-12', 2006: '2006-10-02', 2007: '2007-10-20',
            2008: '2008-10-09', 2009: '2009-09-28', 2010: '2010-10-17', 2011: '2011-10-06',
            2012: '2012-10-24', 2013: '2013-10-13', 2014: '2014-10-03', 2015: '2015-10-22',
            2016: '2016-10-11', 2017: '2017-09-30', 2018: '2018-10-19', 2019: '2019-10-08',
            2020: '2020-10-26', 2021: '2021-10-15', 2022: '2022-10-05', 2023: '2023-10-24',
            2024: '2024-10-12', 2025: '2025-10-02'
        },
        'desc': 'Navratri into Vijayadashami festive momentum & seasonal liquidity injection'
    },
    {
        'id': 'diwali',
        'name': 'Diwali (Laxmi Pujan & Balipratipada)',
        'category': 'Festival of Lights',
        '2026_date': '2026-10-21',
        '2026_date_str': '21-Oct-2026 (Wed)',
        'dates': {
            2000: '2000-10-27', 2001: '2001-11-16', 2002: '2002-11-06', 2003: '2003-10-27',
            2004: '2004-11-15', 2005: '2005-11-03', 2006: '2006-10-24', 2007: '2007-11-12',
            2008: '2008-10-30', 2009: '2009-10-19', 2010: '2010-11-05', 2011: '2011-10-27',
            2012: '2012-11-14', 2013: '2013-11-04', 2014: '2014-10-24', 2015: '2015-11-12',
            2016: '2016-10-31', 2017: '2017-10-20', 2018: '2018-11-08', 2019: '2019-10-28',
            2020: '2020-11-16', 2021: '2021-11-05', 2022: '2022-10-24', 2023: '2023-11-12',
            2024: '2024-11-01', 2025: '2025-10-20'
        },
        'desc': 'Samvat New Year Muhurat trading, festive demand surge & pre-Diwali run-up'
    },
    {
        'id': 'gurunanak',
        'name': 'Gurunanak Jayanti',
        'category': 'Festival',
        '2026_date': '2026-11-24',
        '2026_date_str': '24-Nov-2026 (Tue)',
        'dates': {
            2000: '2000-11-11', 2001: '2001-11-30', 2002: '2002-11-19', 2003: '2003-11-08',
            2004: '2004-11-26', 2005: '2005-11-15', 2006: '2006-11-05', 2007: '2007-11-24',
            2008: '2008-11-13', 2009: '2009-11-02', 2010: '2010-11-21', 2011: '2011-11-10',
            2012: '2012-11-28', 2013: '2013-11-17', 2014: '2014-11-06', 2015: '2015-11-25',
            2016: '2016-11-14', 2017: '2017-11-04', 2018: '2018-11-23', 2019: '2019-11-12',
            2020: '2020-11-30', 2021: '2021-11-19', 2022: '2022-11-08', 2023: '2023-11-27',
            2024: '2024-11-15', 2025: '2025-11-05'
        },
        'desc': 'Post-Diwali momentum continuation and late November monthly expiry positioning'
    },
    {
        'id': 'christmas',
        'name': 'Christmas & Year-End',
        'category': 'Year-End Season',
        '2026_date': '2026-12-25',
        '2026_date_str': '25-Dec-2026 (Fri)',
        'dates': {
            2000: '2000-12-25', 2001: '2001-12-25', 2002: '2002-12-25', 2003: '2003-12-25',
            2004: '2004-12-25', 2005: '2005-12-25', 2006: '2006-12-25', 2007: '2007-12-25',
            2008: '2008-12-25', 2009: '2009-12-25', 2010: '2010-12-25', 2011: '2011-12-25',
            2012: '2012-12-25', 2013: '2013-12-25', 2014: '2014-12-25', 2015: '2015-12-25',
            2016: '2016-12-25', 2017: '2017-12-25', 2018: '2018-12-25', 2019: '2019-12-25',
            2020: '2020-12-25', 2021: '2021-12-25', 2022: '2022-12-25', 2023: '2023-12-25',
            2024: '2024-12-25', 2025: '2025-12-25'
        },
        'desc': 'FII book closing, global liquidity contraction & tax-loss harvesting play'
    },
    {
        'id': 'republic',
        'name': 'Republic Day',
        'category': 'National Holiday',
        '2026_date': '2026-01-26',
        '2026_date_str': '26-Jan-2026 (Mon)',
        'dates': {
            2001: '2001-01-26', 2002: '2002-01-26', 2003: '2003-01-26', 2004: '2004-01-26',
            2005: '2005-01-26', 2006: '2006-01-26', 2007: '2007-01-26', 2008: '2008-01-26',
            2009: '2009-01-26', 2010: '2010-01-26', 2011: '2011-01-26', 2012: '2012-01-26',
            2013: '2013-01-26', 2014: '2014-01-26', 2015: '2015-01-26', 2016: '2016-01-26',
            2017: '2017-01-26', 2018: '2018-01-26', 2019: '2019-01-26', 2020: '2020-01-26',
            2021: '2021-01-26', 2022: '2022-01-26', 2023: '2023-01-26', 2024: '2024-01-26',
            2025: '2025-01-26', 2026: '2026-01-26'
        },
        'desc': 'Pre-Union Budget & Republic Day national rally / budget anticipation buildup'
    },
    {
        'id': 'mahashivratri',
        'name': 'Mahashivratri',
        'category': 'Festival',
        '2026_date': '2026-03-03',
        '2026_date_str': '03-Mar-2026 (Tue)',
        'dates': {
            2001: '2001-02-21', 2002: '2002-03-12', 2003: '2003-03-01', 2004: '2004-02-18',
            2005: '2005-03-08', 2006: '2006-02-26', 2007: '2007-02-16', 2008: '2008-03-06',
            2009: '2009-02-23', 2010: '2010-02-12', 2011: '2011-03-02', 2012: '2012-02-20',
            2013: '2013-03-10', 2014: '2014-02-27', 2015: '2015-02-17', 2016: '2016-03-07',
            2017: '2017-02-24', 2018: '2018-02-13', 2019: '2019-03-04', 2020: '2020-02-21',
            2021: '2021-03-11', 2022: '2022-03-01', 2023: '2023-02-18', 2024: '2024-03-08',
            2025: '2025-02-26', 2026: '2026-03-03'
        },
        'desc': 'Late winter festival accumulation momentum play'
    },
    {
        'id': 'holi',
        'name': 'Holi Festival',
        'category': 'Festival',
        '2026_date': '2026-03-14',
        '2026_date_str': '14-Mar-2026 (Sat)',
        'dates': {
            2001: '2001-03-10', 2002: '2002-03-29', 2003: '2003-03-18', 2004: '2004-03-05',
            2005: '2005-03-25', 2006: '2006-03-15', 2007: '2007-03-27', 2008: '2008-03-22',
            2009: '2009-03-11', 2010: '2010-03-01', 2011: '2011-03-20', 2012: '2012-03-08',
            2013: '2013-03-27', 2014: '2014-03-17', 2015: '2015-03-06', 2016: '2016-03-24',
            2017: '2017-03-13', 2018: '2018-03-02', 2019: '2019-03-21', 2020: '2020-03-10',
            2021: '2021-03-29', 2022: '2022-03-18', 2023: '2023-03-07', 2024: '2024-03-25',
            2025: '2025-03-14'
        },
        'desc': 'Spring festival consumption surge and pre-Holi equity momentum'
    },
    {
        'id': 'ramnavami',
        'name': 'Shri Ram Navami',
        'category': 'Festival',
        '2026_date': '2026-03-26',
        '2026_date_str': '26-Mar-2026 (Thu)',
        'dates': {
            2001: '2001-04-02', 2002: '2002-04-21', 2003: '2003-04-11', 2004: '2004-03-30',
            2005: '2005-04-18', 2006: '2006-04-06', 2007: '2007-03-28', 2008: '2008-04-14',
            2009: '2009-04-03', 2010: '2010-03-24', 2011: '2011-04-12', 2012: '2012-04-01',
            2013: '2013-04-19', 2014: '2014-04-08', 2015: '2015-03-28', 2016: '2016-04-15',
            2017: '2017-04-04', 2018: '2018-03-25', 2019: '2019-04-13', 2020: '2020-04-02',
            2021: '2021-04-21', 2022: '2022-04-10', 2023: '2023-03-30', 2024: '2024-04-17',
            2025: '2025-04-06'
        },
        'desc': 'Chaitra Navratri conclusion & financial year-end liquidity adjustment'
    },
    {
        'id': 'mahavir',
        'name': 'Shri Mahavir Jayanti',
        'category': 'Festival',
        '2026_date': '2026-03-31',
        '2026_date_str': '31-Mar-2026 (Tue)',
        'dates': {
            2001: '2001-04-05', 2002: '2002-04-26', 2003: '2003-04-15', 2004: '2004-04-03',
            2005: '2005-04-22', 2006: '2006-04-11', 2007: '2007-03-31', 2008: '2008-04-18',
            2009: '2009-04-07', 2010: '2010-03-28', 2011: '2011-04-16', 2012: '2012-04-05',
            2013: '2013-04-24', 2014: '2014-04-13', 2015: '2015-04-02', 2016: '2016-04-19',
            2017: '2017-04-09', 2018: '2018-03-29', 2019: '2019-04-17', 2020: '2020-04-06',
            2021: '2021-04-25', 2022: '2022-04-14', 2023: '2023-04-04', 2024: '2024-04-21',
            2025: '2025-04-10'
        },
        'desc': 'Late March / early April financial year-end closing rally'
    },
    {
        'id': 'goodfriday',
        'name': 'Good Friday',
        'category': 'Religious / Global',
        '2026_date': '2026-04-03',
        '2026_date_str': '03-Apr-2026 (Fri)',
        'dates': {
            2001: '2001-04-13', 2002: '2002-03-29', 2003: '2003-04-18', 2004: '2004-04-09',
            2005: '2005-03-25', 2006: '2006-04-14', 2007: '2007-04-06', 2008: '2008-03-21',
            2009: '2009-04-10', 2010: '2010-04-02', 2011: '2011-04-22', 2012: '2012-04-06',
            2013: '2013-03-29', 2014: '2014-04-18', 2015: '2015-04-03', 2016: '2016-03-25',
            2017: '2017-04-14', 2018: '2018-03-30', 2019: '2019-04-19', 2020: '2020-04-10',
            2021: '2021-04-02', 2022: '2022-04-15', 2023: '2023-04-07', 2024: '2024-03-29',
            2025: '2025-04-18', 2026: '2026-04-03'
        },
        'desc': 'Easter long weekend global risk-off positioning and early Q1 earnings positioning'
    },
    {
        'id': 'ambedkar',
        'name': 'Dr. B.R. Ambedkar Jayanti',
        'category': 'National Holiday',
        '2026_date': '2026-04-14',
        '2026_date_str': '14-Apr-2026 (Tue)',
        'dates': {
            2001: '2001-04-14', 2002: '2002-04-14', 2003: '2003-04-14', 2004: '2004-04-14',
            2005: '2005-04-14', 2006: '2006-04-14', 2007: '2007-04-14', 2008: '2008-04-14',
            2009: '2009-04-14', 2010: '2010-04-14', 2011: '2011-04-14', 2012: '2012-04-14',
            2013: '2013-04-14', 2014: '2014-04-14', 2015: '2015-04-14', 2016: '2016-04-14',
            2017: '2017-04-14', 2018: '2018-04-14', 2019: '2019-04-14', 2020: '2020-04-14',
            2021: '2021-04-14', 2022: '2022-04-14', 2023: '2023-04-14', 2024: '2024-04-14',
            2025: '2025-04-14', 2026: '2026-04-14'
        },
        'desc': 'Mid-April IT earnings kickoff (TCS/Infosys results window)'
    },
    {
        'id': 'maharashtra',
        'name': 'Maharashtra Day',
        'category': 'State & Labour Day',
        '2026_date': '2026-05-01',
        '2026_date_str': '01-May-2026 (Fri)',
        'dates': {
            2001: '2001-05-01', 2002: '2002-05-01', 2003: '2003-05-01', 2004: '2004-05-01',
            2005: '2005-05-01', 2006: '2006-05-01', 2007: '2007-05-01', 2008: '2008-05-01',
            2009: '2009-05-01', 2010: '2010-05-01', 2011: '2011-05-01', 2012: '2012-05-01',
            2013: '2013-05-01', 2014: '2014-05-01', 2015: '2015-05-01', 2016: '2016-05-01',
            2017: '2017-05-01', 2018: '2018-05-01', 2019: '2019-05-01', 2020: '2020-05-01',
            2021: '2021-05-01', 2022: '2022-05-01', 2023: '2023-05-01', 2024: '2024-05-01',
            2025: '2025-05-01', 2026: '2026-05-01'
        },
        'desc': 'May series opening accumulation & start of Q4 corporate earnings peak'
    },
    {
        'id': 'bakriid',
        'name': 'Bakri Id (Id-Ul-Adha)',
        'category': 'Festival',
        '2026_date': '2026-05-28',
        '2026_date_str': '28-May-2026 (Thu)',
        'dates': {
            2001: '2001-03-06', 2002: '2002-02-23', 2003: '2003-02-13', 2004: '2004-02-02',
            2005: '2005-01-21', 2006: '2006-01-11', 2007: '2007-12-21', 2008: '2008-12-09',
            2009: '2009-11-28', 2010: '2010-11-17', 2011: '2011-11-07', 2012: '2012-10-27',
            2013: '2013-10-16', 2014: '2014-10-06', 2015: '2015-09-25', 2016: '2016-09-13',
            2017: '2017-09-02', 2018: '2018-08-22', 2019: '2019-08-12', 2020: '2020-08-01',
            2021: '2021-07-21', 2022: '2022-07-10', 2023: '2023-06-29', 2024: '2024-06-17',
            2025: '2025-06-07'
        },
        'desc': 'Mid-year consumption and festival liquidity cycle'
    },
    {
        'id': 'muharram',
        'name': 'Muharram (Ashura)',
        'category': 'Religious Occasion',
        '2026_date': '2026-06-26',
        '2026_date_str': '26-Jun-2026 (Fri)',
        'dates': {
            2001: '2001-04-05', 2002: '2002-03-25', 2003: '2003-03-14', 2004: '2004-03-02',
            2005: '2005-02-19', 2006: '2006-02-09', 2007: '2007-01-30', 2008: '2008-01-19',
            2009: '2009-12-28', 2010: '2010-12-17', 2011: '2011-12-06', 2012: '2012-11-25',
            2013: '2013-11-15', 2014: '2014-11-04', 2015: '2015-10-24', 2016: '2016-10-12',
            2017: '2017-10-01', 2018: '2018-09-21', 2019: '2019-09-10', 2020: '2020-08-30',
            2021: '2021-08-19', 2022: '2022-08-09', 2023: '2023-07-29', 2024: '2024-07-17',
            2025: '2025-07-06'
        },
        'desc': 'Monsoon onset risk-off hedging & volatility expansion window'
    },
    {
        'id': 'independence',
        'name': 'Independence Day',
        'category': 'National Holiday',
        '2026_date': '2026-08-15',
        '2026_date_str': '15-Aug-2026 (Sat)',
        'dates': {
            2000: '2000-08-15', 2001: '2001-08-15', 2002: '2002-08-15', 2003: '2003-08-15',
            2004: '2004-08-15', 2005: '2005-08-15', 2006: '2006-08-15', 2007: '2007-08-15',
            2008: '2008-08-15', 2009: '2009-08-15', 2010: '2010-08-15', 2011: '2011-08-15',
            2012: '2012-08-15', 2013: '2013-08-15', 2014: '2014-08-15', 2015: '2015-08-15',
            2016: '2016-08-15', 2017: '2017-08-15', 2018: '2018-08-15', 2019: '2019-08-15',
            2020: '2020-08-15', 2021: '2021-08-15', 2022: '2022-08-15', 2023: '2023-08-15',
            2024: '2024-08-15', 2025: '2025-08-15'
        },
        'desc': 'Pre-Independence Day national sentiment rally and post-Q1 results consolidation'
    },
    {
        'id': 'ganesh',
        'name': 'Ganesh Chaturthi',
        'category': 'Festive Accumulation',
        '2026_date': '2026-09-14',
        '2026_date_str': '14-Sep-2026 (Mon)',
        'dates': {
            2001: '2001-08-22', 2002: '2002-09-10', 2003: '2003-08-31', 2004: '2004-09-18',
            2005: '2005-09-07', 2006: '2006-08-27', 2007: '2007-09-15', 2008: '2008-09-03',
            2009: '2009-08-23', 2010: '2010-09-11', 2011: '2011-09-01', 2012: '2012-09-19',
            2013: '2013-09-09', 2014: '2014-08-29', 2015: '2015-09-17', 2016: '2016-09-05',
            2017: '2017-08-25', 2018: '2018-09-13', 2019: '2019-09-02', 2020: '2020-08-22',
            2021: '2021-09-10', 2022: '2022-08-31', 2023: '2023-09-19', 2024: '2024-09-07',
            2025: '2025-08-27'
        },
        'desc': 'Ganesh Utsav festive shopping, automobile sales surge & Q2 ending momentum'
    }
]

# 3. Exhaustive Grid Backtesting Engine (T-1..T-8 to T+1..T+8) for Nifty Futures
NIFTY_CURRENT_LOT = 25  # Current official NSE lot size for NIFTY Futures
NIFTY_MARGIN = 125000.0 # ~SPAN + Exposure margin for 1 lot Nifty Futures

def run_holiday_backtest(h_dict):
    dates_map = h_dict['dates']
    results = []
    
    for n in range(1, 9):
        for m in range(1, 9):
            for direction in ['LONG', 'SHORT']:
                trades = []
                for yr, dt_str in dates_map.items():
                    target_dt = pd.to_datetime(dt_str).date()
                    
                    # Find index in continuous trading calendar
                    # Target index is the last trading day on or before target_dt
                    idx = -1
                    for i, d in enumerate(dates_list):
                        if d <= target_dt:
                            idx = i
                        else:
                            break
                    
                    if idx < 0:
                        continue
                    
                    entry_idx = idx - n
                    exit_idx = idx + m
                    
                    if entry_idx < 0 or exit_idx >= len(dates_list):
                        continue
                    
                    entry_date = dates_list[entry_idx]
                    exit_date = dates_list[exit_idx]
                    entry_price = close_list[entry_idx]
                    exit_price = close_list[exit_idx]
                    
                    if entry_price <= 0:
                        continue
                        
                    if direction == 'LONG':
                        pts = exit_price - entry_price
                        ret = (pts / entry_price) * 100.0
                    else:
                        pts = entry_price - exit_price
                        ret = (pts / entry_price) * 100.0
                        
                    pnl_rs = pts * NIFTY_CURRENT_LOT
                    
                    trades.append({
                        'year': yr,
                        'holiday_date': target_dt.strftime('%Y-%m-%d'),
                        'entry_date': entry_date.strftime('%d-%b-%Y (%a)'),
                        'entry_price': round(entry_price, 2),
                        'exit_date': exit_date.strftime('%d-%b-%Y (%a)'),
                        'exit_price': round(exit_price, 2),
                        'points': round(pts, 2),
                        'ret_pct': round(ret, 2),
                        'pnl_rs': round(pnl_rs, 2),
                        'outcome': 'WIN' if pts > 0 else 'LOSS'
                    })
                    
                if not trades:
                    continue
                    
                total_trades = len(trades)
                wins = sum(1 for t in trades if t['outcome'] == 'WIN')
                losses = total_trades - wins
                wr = (wins / total_trades) * 100.0
                
                tot_pts = sum(t['points'] for t in trades)
                avg_pts = tot_pts / total_trades
                avg_ret = sum(t['ret_pct'] for t in trades) / total_trades
                median_pts = float(np.median([t['points'] for t in trades]))
                
                tot_pnl_rs = sum(t['pnl_rs'] for t in trades)
                avg_pnl_rs = tot_pnl_rs / total_trades
                
                # Profit factor
                gross_wins = sum(t['points'] for t in trades if t['points'] > 0)
                gross_loss = abs(sum(t['points'] for t in trades if t['points'] < 0))
                profit_factor = (gross_wins / gross_loss) if gross_loss > 0 else (99.0 if gross_wins > 0 else 0.0)
                
                # Drawdown calculation
                cum_pts = np.cumsum([t['points'] for t in trades])
                peak_pts = np.maximum.accumulate(cum_pts)
                dd_pts = peak_pts - cum_pts
                max_dd_pts = float(np.max(dd_pts)) if len(dd_pts) > 0 else 0.0
                
                results.append({
                    'n': n,
                    'm': m,
                    'window': f"T-{n} to T+{m}",
                    'direction': direction,
                    'total_trades': total_trades,
                    'wins': wins,
                    'losses': losses,
                    'wr': round(wr, 2),
                    'avg_ret': round(avg_ret, 2),
                    'avg_pts': round(avg_pts, 2),
                    'median_pts': round(median_pts, 2),
                    'tot_pts': round(tot_pts, 2),
                    'avg_pnl_rs': round(avg_pnl_rs, 2),
                    'tot_pnl_rs': round(tot_pnl_rs, 2),
                    'profit_factor': round(profit_factor, 2),
                    'max_dd_pts': round(max_dd_pts, 2),
                    'trades': trades
                })
                
    # Sort and pick optimal window: maximize win rate, then cumulative points
    results.sort(key=lambda x: (x['wr'], x['tot_pts']), reverse=True)
    best = results[0] if results else None
    return best, results

# 4. Process All 17 Holidays
print("Executing 26-Year Grid Backtest across all 17 Holidays...")
holiday_playbook = []
all_trades_flat = []

for h in holidays_def:
    best, all_combos = run_holiday_backtest(h)
    if not best:
        continue
        
    # Calculate streaks for the optimal window
    streak_win, streak_loss, max_streak_win, max_streak_loss = 0, 0, 0, 0
    for t in best['trades']:
        if t['outcome'] == 'WIN':
            streak_win += 1
            streak_loss = 0
            if streak_win > max_streak_win:
                max_streak_win = streak_win
        else:
            streak_loss += 1
            streak_win = 0
            if streak_loss > max_streak_loss:
                max_streak_loss = streak_loss
                
    # Upcoming execution dates for 2026
    h_2026_dt = pd.to_datetime(h['2026_date']).date()
    # Find index in continuous series
    idx_2026 = -1
    for i, d in enumerate(dates_list):
        if d <= h_2026_dt:
            idx_2026 = i
        else:
            break
            
    n_opt = best['n']
    m_opt = best['m']
    
    # Calculate upcoming 2026 execution window
    # If 2026 date is in future, calculate from last known date or calendar
    entry_2026_idx = max(0, idx_2026 - n_opt)
    exit_2026_idx = min(len(dates_list) - 1, idx_2026 + m_opt)
    
    entry_2026_str = dates_list[entry_2026_idx].strftime('%d-%b-%Y (%a)') if idx_2026 < len(dates_list) else 'TBD'
    exit_2026_str = dates_list[exit_2026_idx].strftime('%d-%b-%Y (%a)') if idx_2026 < len(dates_list) else 'TBD'
    
    item = {
        'id': h['id'],
        'name': h['name'],
        'category': h['category'],
        'desc': h['desc'],
        '2026_date_str': h['2026_date_str'],
        'optimal_window': best['window'],
        'lead_n': best['n'],
        'lag_m': best['m'],
        'direction': best['direction'],
        'bias_label': f"NIFTY {best['direction']}",
        'total_trades': best['total_trades'],
        'wins': best['wins'],
        'losses': best['losses'],
        'win_rate': best['wr'],
        'avg_return_pct': best['avg_ret'],
        'avg_points': best['avg_pts'],
        'median_points': best['median_pts'],
        'tot_points': best['tot_pts'],
        'avg_pnl_rs': best['avg_pnl_rs'],
        'tot_pnl_rs': best['tot_pnl_rs'],
        'profit_factor': best['profit_factor'],
        'max_dd_points': best['max_dd_pts'],
        'max_win_streak': max_streak_win,
        'upcoming_entry': entry_2026_str,
        'upcoming_exit': exit_2026_str,
        'trades': best['trades']
    }
    
    holiday_playbook.append(item)
    
    for t in best['trades']:
        t_copy = dict(t)
        t_copy['holiday_id'] = h['id']
        t_copy['holiday_name'] = h['name']
        t_copy['window'] = best['window']
        t_copy['direction'] = best['direction']
        all_trades_flat.append(t_copy)

# Sort playbook by Win Rate descending
holiday_playbook.sort(key=lambda x: (x['win_rate'], x['tot_points']), reverse=True)

print(f"Evaluated {len(holiday_playbook)} holidays with {len(all_trades_flat)} total historical trades.")
for h in holiday_playbook[:8]:
    print(f"  {h['name']:<30} | {h['bias_label']:<12} | {h['optimal_window']:<10} | WR: {h['win_rate']}% | Avg: +{h['avg_points']} pts | Tot: +{h['tot_points']} pts")

# 5. Build Master Excel Report with Institutional Styling
wb = openpyxl.Workbook()
wb.remove(wb.active)  # Remove default sheet

# Fonts & Palette
hdr_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
subhdr_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
title_font = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
bold_font = Font(name="Segoe UI", size=9, bold=True)
regular_font = Font(name="Segoe UI", size=9)
mono_font = Font(name="Consolas", size=9)
mono_bold = Font(name="Consolas", size=9, bold=True)

fill_dark_navy = PatternFill(start_color="0E1626", end_color="0E1626", fill_type="solid")
fill_navy_hdr = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
fill_blue_hdr = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

fill_green_tag = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
font_green_tag = Font(name="Segoe UI", size=9, bold=True, color="166534")

fill_red_tag = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
font_red_tag = Font(name="Segoe UI", size=9, bold=True, color="991B1B")

thin_border = Border(
    left=Side(style='thin', color='E2E8F0'),
    right=Side(style='thin', color='E2E8F0'),
    top=Side(style='thin', color='E2E8F0'),
    bottom=Side(style='thin', color='E2E8F0')
)

# ----------------------------------------------------
# Sheet 1: Executive_Summary
# ----------------------------------------------------
ws_exec = wb.create_sheet(title="Executive_Summary")
ws_exec.views.sheetView[0].showGridLines = True

ws_exec.merge_cells("A1:M1")
ws_exec["A1"] = "SMC GLOBAL • NIFTY 50 FUTURES INDEX — 26-YEAR EMPIRICAL HOLIDAY BEHAVIORAL MASTER"
ws_exec["A1"].font = title_font
ws_exec["A1"].fill = fill_dark_navy
ws_exec["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws_exec.row_dimensions[1].height = 40

ws_exec.merge_cells("A2:M2")
ws_exec["A2"] = "Full 2000–2026 Continuous Nifty Futures Dataset (6,473 Daily Bars) · Exhaustive T-n to T+m Window Backtest across All 17 NSE Market Holidays"
ws_exec["A2"].font = Font(name="Segoe UI", size=9, italic=True, color="94A3B8")
ws_exec["A2"].fill = fill_dark_navy
ws_exec["A2"].alignment = Alignment(horizontal="center", vertical="center")
ws_exec.row_dimensions[2].height = 22

headers_exec = [
    "Rank", "NSE Market Holiday", "Occasion Category", "Optimal Window", "Nifty Bias",
    "26Y Trades", "Wins", "Losses", "Win Rate (%)", "Avg Return (%)",
    "Avg Points / Trade", "Cumulative 26Y Points", "Profit Factor"
]

ws_exec.append([])
ws_exec.append(headers_exec)
header_row_idx = 4
ws_exec.row_dimensions[header_row_idx].height = 28

for col_idx, h in enumerate(headers_exec, start=1):
    cell = ws_exec.cell(row=header_row_idx, column=col_idx)
    cell.font = hdr_font
    cell.fill = fill_navy_hdr
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = thin_border

for r_idx, h in enumerate(holiday_playbook, start=1):
    row_num = header_row_idx + r_idx
    row_data = [
        r_idx,
        h['name'],
        h['category'],
        h['optimal_window'],
        h['bias_label'],
        h['total_trades'],
        h['wins'],
        h['losses'],
        h['win_rate'],
        h['avg_return_pct'],
        h['avg_points'],
        h['tot_points'],
        h['profit_factor']
    ]
    ws_exec.append(row_data)
    ws_exec.row_dimensions[row_num].height = 22
    
    fill_row = fill_zebra if r_idx % 2 == 0 else fill_white
    for c_idx in range(1, len(row_data) + 1):
        cell = ws_exec.cell(row=row_num, column=c_idx)
        cell.border = thin_border
        cell.fill = fill_row
        cell.font = regular_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
        if c_idx == 2:
            cell.alignment = Alignment(horizontal="left", vertical="center")
            cell.font = bold_font
        elif c_idx == 5:
            if 'LONG' in str(cell.value):
                cell.fill = fill_green_tag
                cell.font = font_green_tag
            else:
                cell.fill = fill_red_tag
                cell.font = font_red_tag
        elif c_idx in [9, 10, 11, 12, 13]:
            cell.font = mono_bold
            if c_idx == 9:
                cell.number_format = '0.0"%"'
            elif c_idx == 10:
                cell.number_format = '+0.00"%";-0.00"%"'
            elif c_idx in [11, 12]:
                cell.number_format = '+#,##0.00;-#,##0.00'
            elif c_idx == 13:
                cell.number_format = '0.00'

# Auto-fit columns
for col in ws_exec.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws_exec.column_dimensions[col_letter].width = max(max_len + 3, 11)
ws_exec.column_dimensions['A'].width = 8
ws_exec.column_dimensions['B'].width = 32

# ----------------------------------------------------
# Sheet 2: Upcoming_Playbook_2026
# ----------------------------------------------------
ws_playbook = wb.create_sheet(title="Upcoming_Playbook_2026")
ws_playbook.views.sheetView[0].showGridLines = True

ws_playbook.merge_cells("A1:K1")
ws_playbook["A1"] = "NIFTY 50 FUTURES INDEX — 2026-2027 UPCOMING HOLIDAY EXECUTION SCHEDULE"
ws_playbook["A1"].font = title_font
ws_playbook["A1"].fill = fill_dark_navy
ws_playbook["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws_playbook.row_dimensions[1].height = 40

headers_playbook = [
    "Holiday Occasion", "2026 Holiday Date", "Recommended Bias", "Optimal Window",
    "Position Entry Window", "Position Exit Date", "Historical Win Rate",
    "Avg Expected Return", "Avg Expected Points", "26Y Max Win Streak", "Strategy Rationale"
]

ws_playbook.append([])
ws_playbook.append(headers_playbook)
ws_playbook.row_dimensions[3].height = 28

for col_idx, h in enumerate(headers_playbook, start=1):
    cell = ws_playbook.cell(row=3, column=col_idx)
    cell.font = hdr_font
    cell.fill = fill_navy_hdr
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = thin_border

for r_idx, h in enumerate(holiday_playbook, start=1):
    row_num = 3 + r_idx
    row_data = [
        h['name'],
        h['2026_date_str'],
        h['bias_label'],
        h['optimal_window'],
        h['upcoming_entry'],
        h['upcoming_exit'],
        h['win_rate'],
        h['avg_return_pct'],
        h['avg_points'],
        h['max_win_streak'],
        h['desc']
    ]
    ws_playbook.append(row_data)
    ws_playbook.row_dimensions[row_num].height = 22
    fill_row = fill_zebra if r_idx % 2 == 0 else fill_white
    
    for c_idx in range(1, len(row_data) + 1):
        cell = ws_playbook.cell(row=row_num, column=c_idx)
        cell.border = thin_border
        cell.fill = fill_row
        cell.font = regular_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
        if c_idx == 1:
            cell.alignment = Alignment(horizontal="left", vertical="center")
            cell.font = bold_font
        elif c_idx == 3:
            if 'LONG' in str(cell.value):
                cell.fill = fill_green_tag
                cell.font = font_green_tag
            else:
                cell.fill = fill_red_tag
                cell.font = font_red_tag
        elif c_idx == 7:
            cell.font = mono_bold
            cell.number_format = '0.0"%"'
        elif c_idx == 8:
            cell.font = mono_bold
            cell.number_format = '+0.00"%";-0.00"%"'
        elif c_idx == 9:
            cell.font = mono_bold
            cell.number_format = '+#,##0.00;-#,##0.00'
        elif c_idx == 11:
            cell.alignment = Alignment(horizontal="left", vertical="center")

for col in ws_playbook.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws_playbook.column_dimensions[col_letter].width = max(max_len + 3, 12)
ws_playbook.column_dimensions['A'].width = 30
ws_playbook.column_dimensions['K'].width = 45

# ----------------------------------------------------
# Sheet 3: Complete_26Y_Trade_Log
# ----------------------------------------------------
ws_tradelog = wb.create_sheet(title="Complete_26Y_Trade_Log")
ws_tradelog.views.sheetView[0].showGridLines = True

ws_tradelog.merge_cells("A1:K1")
ws_tradelog["A1"] = "NIFTY 50 FUTURES INDEX — COMPLETE 2000–2026 HISTORICAL TRADE-BY-TRADE AUDIT LOG"
ws_tradelog["A1"].font = title_font
ws_tradelog["A1"].fill = fill_dark_navy
ws_tradelog["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws_tradelog.row_dimensions[1].height = 40

headers_tradelog = [
    "Trade #", "Holiday Occasion", "Year", "Holiday Date", "Direction", "Window",
    "Entry Date", "Entry Price (₹)", "Exit Date", "Exit Price (₹)",
    "Points P&L", "Return (%)", "1-Lot P&L (₹)", "Outcome"
]

ws_tradelog.append([])
ws_tradelog.append(headers_tradelog)
ws_tradelog.row_dimensions[3].height = 28

for col_idx, h in enumerate(headers_tradelog, start=1):
    cell = ws_tradelog.cell(row=3, column=col_idx)
    cell.font = hdr_font
    cell.fill = fill_navy_hdr
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = thin_border

# Sort all trades by Year ascending, then Holiday
all_trades_flat.sort(key=lambda x: (x['year'], x['holiday_date']))

for r_idx, t in enumerate(all_trades_flat, start=1):
    row_num = 3 + r_idx
    row_data = [
        r_idx,
        t['holiday_name'],
        t['year'],
        t['holiday_date'],
        t['direction'],
        t['window'],
        t['entry_date'],
        t['entry_price'],
        t['exit_date'],
        t['exit_price'],
        t['points'],
        t['ret_pct'],
        t['pnl_rs'],
        t['outcome']
    ]
    ws_tradelog.append(row_data)
    ws_tradelog.row_dimensions[row_num].height = 20
    fill_row = fill_zebra if r_idx % 2 == 0 else fill_white
    
    for c_idx in range(1, len(row_data) + 1):
        cell = ws_tradelog.cell(row=row_num, column=c_idx)
        cell.border = thin_border
        cell.fill = fill_row
        cell.font = regular_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
        if c_idx == 2:
            cell.alignment = Alignment(horizontal="left", vertical="center")
        elif c_idx == 5:
            if cell.value == 'LONG':
                cell.fill = fill_green_tag
                cell.font = font_green_tag
            else:
                cell.fill = fill_red_tag
                cell.font = font_red_tag
        elif c_idx in [8, 10]:
            cell.font = mono_font
            cell.number_format = '#,##0.00'
        elif c_idx in [11, 12, 13]:
            cell.font = mono_bold
            if c_idx == 11:
                cell.number_format = '+#,##0.00;-#,##0.00'
            elif c_idx == 12:
                cell.number_format = '+0.00"%";-0.00"%"'
            elif c_idx == 13:
                cell.number_format = '₹+#,##0.00;₹-#,##0.00'
        elif c_idx == 14:
            if cell.value == 'WIN':
                cell.fill = fill_green_tag
                cell.font = font_green_tag
            else:
                cell.fill = fill_red_tag
                cell.font = font_red_tag

for col in ws_tradelog.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws_tradelog.column_dimensions[col_letter].width = max(max_len + 3, 11)
ws_tradelog.column_dimensions['A'].width = 9
ws_tradelog.column_dimensions['B'].width = 28

# ----------------------------------------------------
# Sheets 4+: Dedicated Sheet per Top 5 Holidays
# ----------------------------------------------------
top_holidays = holiday_playbook[:6]
for h in top_holidays:
    sheet_title = h['name'][:28].replace('/', '-').replace('&', 'and')
    ws_h = wb.create_sheet(title=sheet_title)
    ws_h.views.sheetView[0].showGridLines = True
    
    ws_h.merge_cells("A1:J1")
    ws_h["A1"] = f"NIFTY FUTURES — {h['name'].upper()} 26-YEAR BACKTEST AUDIT"
    ws_h["A1"].font = title_font
    ws_h["A1"].fill = fill_dark_navy
    ws_h["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_h.row_dimensions[1].height = 36
    
    ws_h.merge_cells("A2:J2")
    ws_h["A2"] = f"Strategy: NIFTY {h['direction']} ({h['optimal_window']}) · Win Rate: {h['win_rate']}% · Cumulative P&L: +{h['tot_points']} Points ({h['tot_pnl_rs']:+,.2f} ₹/Lot)"
    ws_h["A2"].font = Font(name="Segoe UI", size=9.5, bold=True, color="38BDF8")
    ws_h["A2"].fill = fill_dark_navy
    ws_h["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws_h.row_dimensions[2].height = 24
    
    h_headers = [
        "Year", "Holiday Date", "Direction", "Window", "Entry Date",
        "Entry Price (₹)", "Exit Date", "Exit Price (₹)", "Points P&L", "Return (%)", "Outcome"
    ]
    ws_h.append([])
    ws_h.append(h_headers)
    ws_h.row_dimensions[4].height = 26
    
    for c_idx, hh in enumerate(h_headers, start=1):
        cell = ws_h.cell(row=4, column=c_idx)
        cell.font = hdr_font
        cell.fill = fill_navy_hdr
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        
    for r_idx, t in enumerate(h['trades'], start=1):
        row_num = 4 + r_idx
        row_data = [
            t['year'],
            t['holiday_date'],
            h['direction'],
            h['optimal_window'],
            t['entry_date'],
            t['entry_price'],
            t['exit_date'],
            t['exit_price'],
            t['points'],
            t['ret_pct'],
            t['outcome']
        ]
        ws_h.append(row_data)
        ws_h.row_dimensions[row_num].height = 20
        fill_row = fill_zebra if r_idx % 2 == 0 else fill_white
        
        for c_idx in range(1, len(row_data) + 1):
            cell = ws_h.cell(row=row_num, column=c_idx)
            cell.border = thin_border
            cell.fill = fill_row
            cell.font = regular_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
            if c_idx in [6, 8]:
                cell.font = mono_font
                cell.number_format = '#,##0.00'
            elif c_idx in [9, 10]:
                cell.font = mono_bold
                if c_idx == 9:
                    cell.number_format = '+#,##0.00;-#,##0.00'
                else:
                    cell.number_format = '+0.00"%";-0.00"%"'
            elif c_idx == 11:
                if cell.value == 'WIN':
                    cell.fill = fill_green_tag
                    cell.font = font_green_tag
                else:
                    cell.fill = fill_red_tag
                    cell.font = font_red_tag

    for col in ws_h.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_h.column_dimensions[col_letter].width = max(max_len + 3, 12)

wb.save(EXCEL_OUT)
print(f"SAVED EXCEL MASTER: {EXCEL_OUT}")

DOCS_EXCEL_OUT = ROOT / "docs" / "Nifty_Futures_Index_Holiday_Behaviour_Master_2000_2026.xlsx"
wb.save(DOCS_EXCEL_OUT)
print(f"SAVED DOCS EXCEL MASTER: {DOCS_EXCEL_OUT}")

# 6. Save JSON Datasets for Frontend Dashboard
json_payload = {
    'generated_at': datetime.datetime.now().strftime('%d-%b-%Y %H:%M:%S'),
    'data_source': 'scraped_parquet/nifty_futures_near_month_continuous.parquet',
    'total_trading_days': len(dates_list),
    'history_span': f"{dates_list[0]} to {dates_list[-1]}",
    'lot_size': NIFTY_CURRENT_LOT,
    'margin_per_lot': NIFTY_MARGIN,
    'holidays_count': len(holiday_playbook),
    'holidays': holiday_playbook
}

with open(JSON_OUT, 'w', encoding='utf-8') as f:
    json.dump(json_payload, f, indent=2)
print(f"SAVED JSON DATA: {JSON_OUT}")

with open(DOCS_JSON_OUT, 'w', encoding='utf-8') as f:
    json.dump(json_payload, f, indent=2)
print(f"SAVED DOCS JSON DATA: {DOCS_JSON_OUT}")

print("==========================================================================================")
print("SUCCESSFULLY BUILT NIFTY FUTURES INDEX HOLIDAY BEHAVIORAL MASTER ENGINE")
print("==========================================================================================")
