"""
build_universe_csv.py
===============================================================================
Generates `fno_derivatives_list.csv` — the scraper's universe.

The scraper used to run on `ind_nifty50list.csv` (50 names). The universe is now
the ~208 NSE F&O / derivatives-eligible underlyings, because those are the only
symbols on which a straddle actually exists: an event on a stock with no option
chain cannot move a straddle we display.

Industry and ISIN are carried over from ind_nifty50list.csv for the 50 symbols we
already had them for. The other 158 are left BLANK rather than guessed — a wrong
ISIN silently mis-joins downstream, and neither column is used by any feed except
the 08_Company_Directory fallback, which tolerates blanks.

    python build_universe_csv.py          # writes fno_derivatives_list.csv
    python build_universe_csv.py --check  # report only, write nothing
===============================================================================
"""

import csv
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
NIFTY50_CSV = os.path.join(HERE, "ind_nifty50list.csv")
OUT_CSV = os.path.join(HERE, "fno_derivatives_list.csv")

# Company Name <TAB> Symbol, exactly as published in the NSE derivatives-eligible
# securities list. Keep this block as the single source of truth; re-run the
# script after editing it.
RAW = """
360 ONE WAM LIMITED	360ONE
ABB India Limited	ABB
APL Apollo Tubes Limited	APLAPOLLO
AU Small Finance Bank Limited	AUBANK
Adani Energy Solutions Limited	ADANIENSOL
Adani Enterprises Limited	ADANIENT
Adani Green Energy Limited	ADANIGREEN
Adani Ports and Special Economic Zone Limited	ADANIPORTS
Adani Power Limited	ADANIPOWER
Aditya Birla Capital Limited	ABCAPITAL
Alkem Laboratories Limited	ALKEM
Amber Enterprises India Limited	AMBER
Ambuja Cements Limited	AMBUJACEM
Angel One Limited	ANGELONE
Apollo Hospitals Enterprise Limited	APOLLOHOSP
Ashok Leyland Limited	ASHOKLEY
Asian Paints Limited	ASIANPAINT
Astral Limited	ASTRAL
Aurobindo Pharma Limited	AUROPHARMA
Avenue Supermarts Limited	DMART
Axis Bank Limited	AXISBANK
BSE Limited	BSE
Bajaj Auto Limited	BAJAJ-AUTO
Bajaj Finance Limited	BAJFINANCE
Bajaj Finserv Limited	BAJAJFINSV
Bajaj Holdings & Investment Limited	BAJAJHLDNG
Bandhan Bank Limited	BANDHANBNK
Bank of Baroda	BANKBARODA
Bank of India	BANKINDIA
Bharat Dynamics Limited	BDL
Bharat Electronics Limited	BEL
Bharat Forge Limited	BHARATFORG
Bharat Heavy Electricals Limited	BHEL
Bharat Petroleum Corporation Limited	BPCL
Bharti Airtel Limited	BHARTIARTL
Biocon Limited	BIOCON
Blue Star Limited	BLUESTARCO
Bosch Limited	BOSCHLTD
Britannia Industries Limited	BRITANNIA
CG Power and Industrial Solutions Limited	CGPOWER
Canara Bank	CANBK
Central Depository Services (India) Limited	CDSL
Cholamandalam Investment and Finance Company Limited	CHOLAFIN
Cipla Limited	CIPLA
Coal India Limited	COALINDIA
Cochin Shipyard Limited	COCHINSHIP
Coforge Limited	COFORGE
Colgate Palmolive (India) Limited	COLPAL
Computer Age Management Services Limited	CAMS
Container Corporation of India Limited	CONCOR
Crompton Greaves Consumer Electricals Limited	CROMPTON
Cummins India Limited	CUMMINSIND
DLF Limited	DLF
Dabur India Limited	DABUR
Dalmia Bharat Limited	DALBHARAT
Delhivery Limited	DELHIVERY
Divi's Laboratories Limited	DIVISLAB
Dixon Technologies (India) Limited	DIXON
Dr. Reddy's Laboratories Limited	DRREDDY
ETERNAL LIMITED	ETERNAL
Eicher Motors Limited	EICHERMOT
FORCE MOTORS LTD	FORCEMOT
FSN E-Commerce Ventures Limited	NYKAA
Fortis Healthcare Limited	FORTIS
GAIL (India) Limited	GAIL
GE Vernova T&D India Limited	GVT&D
GMR AIRPORTS LIMITED	GMRAIRPORT
Glenmark Pharmaceuticals Limited	GLENMARK
Godfrey Phillips India Limited	GODFRYPHLP
Godrej Consumer Products Limited	GODREJCP
Godrej Properties Limited	GODREJPROP
Grasim Industries Limited	GRASIM
HCL Technologies Limited	HCLTECH
HDFC Asset Management Company Limited	HDFCAMC
HDFC Bank Limited	HDFCBANK
HDFC Life Insurance Company Limited	HDFCLIFE
Havells India Limited	HAVELLS
Hero MotoCorp Limited	HEROMOTOCO
Hindalco Industries Limited	HINDALCO
Hindustan Aeronautics Limited	HAL
Hindustan Petroleum Corporation Limited	HINDPETRO
Hindustan Unilever Limited	HINDUNILVR
Hindustan Zinc Limited	HINDZINC
Hitachi Energy India Limited	POWERINDIA
Hyundai Motor India Limited	HYUNDAI
ICICI Bank Limited	ICICIBANK
ICICI Lombard General Insurance Company Limited	ICICIGI
ICICI Prudential Life Insurance Company Limited	ICICIPRULI
IDFC First Bank Limited	IDFCFIRSTB
ITC Limited	ITC
Indian Bank	INDIANB
Indian Energy Exchange Limited	IEX
Indian Oil Corporation Limited	IOC
Indian Railway Finance Corporation Limited	IRFC
Indian Renewable Energy Development Agency Limited	IREDA
Indus Towers Limited	INDUSTOWER
IndusInd Bank Limited	INDUSINDBK
Info Edge (India) Limited	NAUKRI
Infosys Limited	INFY
Inox Wind Limited	INOXWIND
InterGlobe Aviation Limited	INDIGO
JINDAL STEEL LIMITED	JINDALSTEL
JSW Energy Limited	JSWENERGY
JSW Steel Limited	JSWSTEEL
Jio Financial Services Limited	JIOFIN
Jubilant Foodworks Limited	JUBLFOOD
KEI Industries Limited	KEI
KPIT Technologies Limited	KPITTECH
Kalyan Jewellers India Limited	KALYANKJIL
Kaynes Technology India Limited	KAYNES
Kfin Technologies Limited	KFINTECH
Kotak Mahindra Bank Limited	KOTAKBANK
L&T Finance Limited	LTF
LIC Housing Finance Limited	LICHSGFIN
LTM Limited	LTM
Larsen & Toubro Limited	LT
Laurus Labs Limited	LAURUSLABS
Life Insurance Corporation Of India	LICI
Lodha Developers Limited	LODHA
Lupin Limited	LUPIN
Mahindra & Mahindra Limited	M&M
Manappuram Finance Limited	MANAPPURAM
Mankind Pharma Limited	MANKIND
Marico Limited	MARICO
Maruti Suzuki India Limited	MARUTI
Max Financial Services Limited	MFSL
Max Healthcare Institute Limited	MAXHEALTH
Mazagon Dock Shipbuilders Limited	MAZDOCK
Motilal Oswal Financial Services Limited	MOTILALOFS
MphasiS Limited	MPHASIS
Multi Commodity Exchange of India Limited	MCX
Muthoot Finance Limited	MUTHOOTFIN
NBCC (India) Limited	NBCC
NHPC Limited	NHPC
NMDC Limited	NMDC
NTPC Limited	NTPC
National Aluminium Company Limited	NATIONALUM
Nestle India Limited	NESTLEIND
Nippon Life India Asset Management Limited	NAM-INDIA
Oberoi Realty Limited	OBEROIRLTY
Oil & Natural Gas Corporation Limited	ONGC
Oil India Limited	OIL
One 97 Communications Limited	PAYTM
Oracle Financial Services Software Limited	OFSS
PB Fintech Limited	POLICYBZR
PG Electroplast Limited	PGEL
PI Industries Limited	PIIND
PNB Housing Finance Limited	PNBHOUSING
Page Industries Limited	PAGEIND
Patanjali Foods Limited	PATANJALI
Persistent Systems Limited	PERSISTENT
Petronet LNG Limited	PETRONET
Pidilite Industries Limited	PIDILITIND
Polycab India Limited	POLYCAB
Power Finance Corporation Limited	PFC
Power Grid Corporation of India Limited	POWERGRID
Premier Energies Limited	PREMIERENE
Prestige Estates Projects Limited	PRESTIGE
Punjab National Bank	PNB
RBL Bank Limited	RBLBANK
REC Limited	RECLTD
Radico Khaitan Limited	RADICO
Rail Vikas Nigam Limited	RVNL
Reliance Industries Limited	RELIANCE
SBI Cards and Payment Services Limited	SBICARD
SBI Life Insurance Company Limited	SBILIFE
SHREE CEMENT LIMITED	SHREECEM
SRF Limited	SRF
Samvardhana Motherson International Limited	MOTHERSON
Shriram Finance Limited	SHRIRAMFIN
Siemens Limited	SIEMENS
Solar Industries India Limited	SOLARINDS
Sona BLW Precision Forgings Limited	SONACOMS
State Bank of India	SBIN
Steel Authority of India Limited	SAIL
Sun Pharmaceutical Industries Limited	SUNPHARMA
Supreme Industries Limited	SUPREMEIND
Suzlon Energy Limited	SUZLON
Swiggy Limited	SWIGGY
TATA CONSUMER PRODUCTS LIMITED	TATACONSUM
TVS Motor Company Limited	TVSMOTOR
Tata Consultancy Services Limited	TCS
Tata Elxsi Limited	TATAELXSI
Tata Motors Passenger Vehicles Limited	TMPV
Tata Power Company Limited	TATAPOWER
Tata Steel Limited	TATASTEEL
Tech Mahindra Limited	TECHM
The Federal Bank Limited	FEDERALBNK
The Indian Hotels Company Limited	INDHOTEL
The Phoenix Mills Limited	PHOENIXLTD
Titan Company Limited	TITAN
Torrent Pharmaceuticals Limited	TORNTPHARM
Trent Limited	TRENT
Tube Investments of India Limited	TIINDIA
UNO Minda Limited	UNOMINDA
UPL Limited	UPL
UltraTech Cement Limited	ULTRACEMCO
Union Bank of India	UNIONBANK
United Spirits Limited	UNITDSPR
Varun Beverages Limited	VBL
Vedanta Limited	VEDL
Vishal Mega Mart Limited	VMM
Vodafone Idea Limited	IDEA
Voltas Limited	VOLTAS
Waaree Energies Limited	WAAREEENER
Wipro Limited	WIPRO
Yes Bank Limited	YESBANK
Zydus Lifesciences Limited	ZYDUSLIFE
"""


def parse_raw():
    rows, seen = [], set()
    for line in RAW.strip().splitlines():
        line = line.strip()
        if not line:
            continue
        if "\t" not in line:
            raise ValueError(f"no tab separator in: {line!r}")
        name, symbol = line.rsplit("\t", 1)
        name, symbol = name.strip(), symbol.strip().upper()
        if symbol in seen:
            raise ValueError(f"duplicate symbol {symbol}")
        seen.add(symbol)
        rows.append((name, symbol))
    return rows


def load_known():
    """Industry + ISIN we already hold, keyed by symbol."""
    known = {}
    if not os.path.exists(NIFTY50_CSV):
        return known
    with open(NIFTY50_CSV, newline="", encoding="utf-8-sig") as fh:
        for r in csv.DictReader(fh):
            sym = (r.get("Symbol") or "").strip().upper()
            if sym:
                known[sym] = ((r.get("Industry") or "").strip(),
                              (r.get("ISIN Code") or "").strip())
    return known


def main():
    check_only = "--check" in sys.argv
    rows = parse_raw()
    known = load_known()

    out, with_isin = [], 0
    for name, symbol in rows:
        industry, isin = known.get(symbol, ("", ""))
        if isin:
            with_isin += 1
        out.append({"Company Name": name, "Industry": industry, "Symbol": symbol,
                    "Series": "EQ", "ISIN Code": isin})

    dropped = sorted(set(known) - {s for _, s in rows})

    print(f"universe        : {len(out)} symbols")
    print(f"ISIN carried    : {with_isin} (from ind_nifty50list.csv)")
    print(f"ISIN blank      : {len(out) - with_isin}")
    if dropped:
        print(f"in Nifty50 CSV but NOT in the new list: {', '.join(dropped)}")

    if check_only:
        print("--check: nothing written")
        return

    with open(OUT_CSV, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=["Company Name", "Industry", "Symbol",
                                           "Series", "ISIN Code"])
        w.writeheader()
        w.writerows(out)
    print(f"wrote -> {OUT_CSV}")


if __name__ == "__main__":
    main()
