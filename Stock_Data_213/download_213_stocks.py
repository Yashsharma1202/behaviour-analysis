import json, pathlib, time, concurrent.futures as cf
import pandas as pd, yfinance as yf

ROOT = pathlib.Path('d:/behaviour analysis')
OUT = ROOT / 'Stock_Data_213'
DAILY = OUT / 'daily_ohlc'
DAILY.mkdir(parents=True, exist_ok=True)

# Symbol list as pasted by user (NSE trading symbol, underlying name)
SYMBOLS = """360ONE ABB APLAPOLLO AUBANK ADANIENSOL ADANIENT ADANIGREEN ADANIPORTS ADANIPOWER ABCAPITAL ALKEM AMBER AMBUJACEM ANANDRATHI ANGELONE APOLLOHOSP ASHOKLEY ASIANPAINT ASTRAL ATHERENERG AUROPHARMA DMART AXISBANK BSE BAJAJ-AUTO BAJFINANCE BAJAJFINSV BAJAJHLDNG BANDHANBNK BANKBARODA BANKINDIA MAHABANK BDL BEL BHARATFORG BHEL BPCL BHARTIARTL BIOCON BLUESTARCO BOSCHLTD BRITANNIA CGPOWER CANBK CDSL CHOLAFIN CIPLA COALINDIA COCHINSHIP COFORGE COLPAL CAMS CONCOR CROMPTON CUMMINSIND DLF DABUR DELHIVERY DIVISLAB DIXON DRREDDY ETERNAL EICHERMOT FORCEMOT NYKAA FORTIS GAIL GVT&D GMRAIRPORT GLENMARK GODFRYPHLP GODREJCP GODREJPROP GRASIM HCLTECH HDFCAMC HDFCBANK HDFCLIFE HAVELLS HEROMOTOCO HINDALCO HAL HINDPETRO HINDUNILVR HINDZINC POWERINDIA HYUNDAI ICICIBANK ICICIGI ICICIPRULI IDFCFIRSTB ITC INDIANB IEX IOC IRFC IREDA INDUSTOWER INDUSINDBK NAUKRI INFY INOXWIND INDIGO JINDALSTEL JSWENERGY JSWSTEEL JIOFIN JUBLFOOD KEI KPITTECH KALYANKJIL KAYNES KFINTECH KOTAKBANK LTF LICHSGFIN LTM LT LAURUSLABS LICI LODHA LUPIN M&M MANAPPURAM MANKIND MARICO MARUTI MFSL MAXHEALTH MAZDOCK MOTILALOFS MPHASIS MCX MUTHOOTFIN NBCC NHPC NMDC NTPC NATIONALUM NESTLEIND NAM-INDIA OBEROIRLTY ONGC OIL PAYTM OFSS POLICYBZR PGEL PIIND PNBHOUSING PAGEIND PATANJALI PERSISTENT PETRONET PIDILITIND POLYCAB PFC POWERGRID PREMIERENE PRESTIGE PNB RBLBANK RECLTD RADICO RVNL RELIANCE SAGILITY SBICARD SBILIFE SHREECEM SRF MOTHERSON SHRIRAMFIN ENRIN SIEMENS SOLARINDS SONACOMS SBIN SAIL SUNPHARMA SUPREMEIND SUZLON SWIGGY TATACONSUM TVSMOTOR TCS TATAELXSI TMPV TATAPOWER TATASTEEL TECHM FEDERALBNK INDHOTEL PHOENIXLTD TITAN TORNTPHARM TRENT TIINDIA UNOMINDA UPL UJJIVANSFB ULTRACEMCO UNIONBANK UNITDSPR VBL VEDL VMM IDEA VOLTAS WAAREEENER WIPRO YESBANK ZYDUSLIFE""".split()

# Yahoo tickers that differ from NSE symbol (fallbacks tried in order)
YF_ALIASES = {
    'TMPV': ['TMPV.NS', 'TATAMOTORS.NS'],          # Tata Motors PV (demerged; old TATAMOTORS history)
    'MAHABANK': ['MAHABANK.NS'],
    'GVT&D': ['GVT&D.NS', 'GVTD.NS'],
    'ENRIN': ['ENRIN.NS'],
    'POWERINDIA': ['POWERINDIA.NS'],
    'HYUNDAI': ['HYUNDAI.NS'],
    'NAM-INDIA': ['NAM-INDIA.NS'],
    'WAAREEENER': ['WAAREEENER.NS'],
    'PAYTM': ['PAYTM.NS'],
    'IDEA': ['IDEA.NS'],
}

def candidates(sym):
    return YF_ALIASES.get(sym, [f'{sym}.NS'])

def fetch(sym):
    for tk in candidates(sym):
        try:
            h = yf.Ticker(tk).history(period='max', auto_adjust=False, actions=True)
        except Exception as e:
            h = None
        if h is not None and len(h):
            h = h.reset_index()
            h['Date'] = pd.to_datetime(h['Date']).dt.tz_localize(None).dt.normalize()
            h.insert(0, 'Symbol', sym)
            h.insert(1, 'YahooTicker', tk)
            cols = ['Symbol', 'YahooTicker', 'Date', 'Open', 'High', 'Low', 'Close', 'Adj Close', 'Volume', 'Dividends', 'Stock Splits']
            h = h[cols]
            h.to_csv(DAILY / f'{sym}.csv', index=False)
            return sym, tk, h
        time.sleep(0.2)
    return sym, None, None

results, frames = [], []
with cf.ThreadPoolExecutor(max_workers=4) as ex:
    for sym, tk, h in ex.map(fetch, SYMBOLS):
        if h is None:
            results.append({'Symbol': sym, 'YahooTicker': '', 'Status': 'NO DATA', 'Rows': 0, 'FirstDate': '', 'LastDate': ''})
            print(f'[NO DATA] {sym}', flush=True)
        else:
            frames.append(h)
            results.append({'Symbol': sym, 'YahooTicker': tk, 'Status': 'OK', 'Rows': len(h),
                            'FirstDate': h['Date'].min().date(), 'LastDate': h['Date'].max().date()})
            print(f'[OK] {sym} {tk} rows={len(h)}', flush=True)

rep = pd.DataFrame(results)
rep.to_csv(OUT / 'download_report.csv', index=False)
if frames:
    allp = pd.concat(frames, ignore_index=True)
    allp.to_parquet(OUT / 'all_stocks_daily_ohlc.parquet', index=False)
    allp.to_csv(OUT / 'all_stocks_daily_ohlc.csv', index=False)

# Lot sizes: from local F&O master (dashboard_data/fo_stocks_211.json); NSE fo_mktlots.csv was unreachable
lots = {d['symbol']: d.get('lot_size') for d in json.load(open(ROOT / 'dashboard_data' / 'fo_stocks_211.json'))}
lot_df = pd.DataFrame({'Symbol': SYMBOLS})
lot_df['LotSize'] = lot_df['Symbol'].map(lots)
lot_df['LotSourceNote'] = lot_df['LotSize'].apply(lambda x: 'local fo_stocks_211.json (verify vs NSE fo_mktlots.csv)' if pd.notna(x) else 'MISSING - not in local master')
lot_df.to_csv(OUT / 'lot_sizes.csv', index=False)
print('OK:', (rep.Status == 'OK').sum(), 'NO DATA:', (rep.Status != 'OK').sum(), 'lots missing:', lot_df.LotSize.isna().sum())
