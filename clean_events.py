"""
clean_events.py
===============================================================================
FOUNDATION STEP of the "behaviour analysis" project.

WHAT THIS DOES (in plain English)
---------------------------------
The four raw CSV files under RELIANCE/ come straight from NSE's website. They are
messy: dates are written five different ways, the "subject"/"desc" fields are free
text, one file has two columns swapped on some rows, and the results file contains
duplicate rows (standalone + consolidated + old/new format) for the same quarter.

This script turns that mess into CLEAN, TYPED tables and then stitches every event
into ONE MASTER TIMELINE (`unified_events.csv`) with a common shape:

        date | symbol | isin | source | event_type | title | amount | ratio | url

That master timeline is the thing every later step (XBRL parsing, event study)
will build on.

HOW TO RUN
----------
    python clean_events.py

Outputs land in the ./processed/ folder next to this script.

DEPENDENCIES: pandas (data wrangling), re (regex text parsing), pathlib/datetime.
===============================================================================
"""

from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

# ---------------------------------------------------------------------------
# 0. CONFIG — where things live
# ---------------------------------------------------------------------------
ROOT = Path(__file__).resolve().parent          # d:\behaviour analysis
RAW_DIR = ROOT / "RELIANCE"                      # the raw CSVs
OUT_DIR = ROOT / "processed"                     # cleaned outputs go here
OUT_DIR.mkdir(exist_ok=True)

SYMBOL = "RELIANCE"
ISIN = "INE002A01018"


# ---------------------------------------------------------------------------
# 1. DATE HANDLING
# ---------------------------------------------------------------------------
# The data mixes these formats:
#   05-Jun-2026                 (day-month-year)
#   17-Apr-2026 19:06:21        (with time)
#   2026-07-06 17:17:52         (ISO, already clean)
#   06072026171752              (packed ddmmyyyyhhmmss)
# We try each known format in order and keep the first that parses.
_DATE_FORMATS = [
    "%d-%b-%Y %H:%M:%S",
    "%d-%b-%Y %H:%M",
    "%d-%b-%Y",
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d",
    "%d%m%Y%H%M%S",
]


def parse_date(value) -> pd.Timestamp | pd.NaT:
    """Turn one messy date string into a real Timestamp (or NaT if unparseable).

    Input:  a string like '05-Jun-2026' or '2026-07-06 17:17:52' or '-'.
    Output: a pandas Timestamp, or pandas NaT for blanks / placeholders.
    """
    if value is None:
        return pd.NaT
    s = str(value).strip()
    if s in ("", "-", "nan", "None"):
        return pd.NaT
    for fmt in _DATE_FORMATS:
        try:
            return pd.to_datetime(s, format=fmt)
        except (ValueError, TypeError):
            continue
    # last resort: let pandas guess (day-first, since this is Indian data)
    return pd.to_datetime(s, dayfirst=True, errors="coerce")


def clean_text(value) -> str:
    """Fix broken UTF-8 (curly quotes showed up as '�') and trim whitespace."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    s = str(value)
    s = s.replace("�", "'")          # replacement char -> apostrophe
    s = re.sub(r"\s+", " ", s).strip()    # collapse runs of whitespace/newlines
    return s


# ---------------------------------------------------------------------------
# 2. PARSE THE FREE-TEXT CORPORATE-ACTION "subject" INTO TYPED FIELDS
# ---------------------------------------------------------------------------
# Examples we must handle:
#   "Dividend - Rs 6 Per Share"          -> DIVIDEND, amount=6
#   "Dividend Rs. 9.50 Per Equity Share" -> DIVIDEND, amount=9.50
#   "Bonus 1:1"                          -> BONUS,    ratio="1:1"
#   "Rights 1:15 @ Premium Rs 1247"      -> RIGHTS,   ratio="1:15", amount=1247
#   "Demerger"                           -> DEMERGER
#   "Annual General Meeting/Dividend - Rs 6.50 Per Share" -> DIVIDEND, amount=6.50
def parse_corporate_action(subject: str) -> dict:
    """Extract {event_type, amount, ratio} from a free-text corporate-action subject.

    Input:  raw subject string.
    Output: dict with keys event_type (str), amount (float|None), ratio (str|None).
    """
    text = clean_text(subject)
    low = text.lower()

    amount = None
    ratio = None

    # a ratio like 1:1 or 1:15  (used by bonus / rights / splits)
    m_ratio = re.search(r"(\d+\s*:\s*\d+)", text)
    if m_ratio:
        ratio = m_ratio.group(1).replace(" ", "")

    # a rupee amount after the word Dividend or after 'Rs'
    m_amt = re.search(r"(?:rs\.?|rupees|dividend)[^0-9]*([0-9]+(?:\.[0-9]+)?)", low)
    if m_amt:
        amount = float(m_amt.group(1))

    # classify by keyword priority (a row can mention AGM + Dividend; dividend wins)
    if "bonus" in low:
        event_type = "BONUS"
    elif "rights" in low:
        event_type = "RIGHTS"
        # for rights the premium (Rs 1247) is the amount; ratio already captured
    elif "split" in low or "face value" in low:
        event_type = "SPLIT"
    elif "demerger" in low or "demerged" in low:
        event_type = "DEMERGER"
    elif "buyback" in low or "buy back" in low:
        event_type = "BUYBACK"
    elif "dividend" in low:
        event_type = "DIVIDEND"
    elif "annual general meeting" in low or "agm" in low:
        event_type = "AGM"
    else:
        event_type = "OTHER"

    return {"event_type": event_type, "amount": amount, "ratio": ratio}


# ---------------------------------------------------------------------------
# 3. LOADERS — one per file. Each returns a cleaned DataFrame AND a slice of the
#    unified timeline (same columns for every source).
# ---------------------------------------------------------------------------
UNIFIED_COLS = ["date", "symbol", "isin", "source",
                "event_type", "title", "amount", "ratio", "url"]


def load_corporate_actions() -> tuple[pd.DataFrame, pd.DataFrame]:
    df = pd.read_csv(RAW_DIR / "corporate_actions.csv", dtype=str).fillna("")
    df["subject"] = df["subject"].map(clean_text)
    df["exDate"] = df["exDate"].map(parse_date)
    df["recDate"] = df["recDate"].map(parse_date)

    parsed = df["subject"].apply(parse_corporate_action).apply(pd.Series)
    clean = pd.concat([df, parsed], axis=1)

    uni = pd.DataFrame({
        "date": clean["exDate"],
        "symbol": clean["symbol"].replace("", SYMBOL),
        "isin": clean["isin"],
        "source": "corporate_action",
        "event_type": clean["event_type"],
        "title": clean["subject"],
        "amount": clean["amount"],
        "ratio": clean["ratio"],
        "url": "",
    })
    return clean, uni


def load_board_meetings() -> tuple[pd.DataFrame, pd.DataFrame]:
    df = pd.read_csv(RAW_DIR / "board_meetings.csv", dtype=str).fillna("")

    # --- FIX THE KNOWN COLUMN-SWAP BUG -------------------------------------
    # On some rows the long description landed in `bm_purpose` and the short
    # label ("Board Meeting Intimation") landed in `bm_desc`. Detect and swap.
    swapped = df["bm_purpose"].str.contains("informed the Exchange", case=False, na=False)
    df.loc[swapped, ["bm_purpose", "bm_desc"]] = (
        df.loc[swapped, ["bm_desc", "bm_purpose"]].values
    )

    df["bm_purpose"] = df["bm_purpose"].map(clean_text)
    df["bm_desc"] = df["bm_desc"].map(clean_text)
    df["bm_date"] = df["bm_date"].map(parse_date)

    def classify_purpose(p: str) -> str:
        low = p.lower()
        if "bonus" in low:
            return "BONUS_MEETING"
        if "dividend" in low:
            return "RESULTS_DIVIDEND_MEETING"
        if "result" in low:
            return "RESULTS_MEETING"
        if "fund" in low:
            return "FUNDRAISE_MEETING"
        return "BOARD_MEETING"

    df["meeting_type"] = df["bm_purpose"].map(classify_purpose)

    uni = pd.DataFrame({
        "date": df["bm_date"],
        "symbol": df["symbol"].replace("", SYMBOL),
        "isin": df["sm_isin"],
        "source": "board_meeting",
        "event_type": df["meeting_type"],
        "title": df["bm_purpose"],
        "amount": None,
        "ratio": None,
        "url": df["attachment"],
    })
    return df, uni


def load_financial_results() -> tuple[pd.DataFrame, pd.DataFrame]:
    df = pd.read_csv(RAW_DIR / "financial_results.csv", dtype=str).fillna("")
    df["broadCastDate"] = df["broadCastDate"].map(parse_date)
    df["fromDate"] = df["fromDate"].map(parse_date)
    df["toDate"] = df["toDate"].map(parse_date)

    # --- DEDUP ------------------------------------------------------------
    # Keep only the latest filing per (period-end, consolidated, audited) combo,
    # and drop revised-supersession noise using seqNumber as tie-breaker.
    df["seqNumber_int"] = pd.to_numeric(df["seqNumber"], errors="coerce")
    df = (df.sort_values("seqNumber_int")
            .drop_duplicates(subset=["toDate", "consolidated", "audited"], keep="last"))

    uni = pd.DataFrame({
        "date": df["broadCastDate"],
        "symbol": df["symbol"].replace("", SYMBOL),
        "isin": df["isin"],
        "source": "financial_result",
        "event_type": df["relatingTo"].map(clean_text).radd("RESULT_"),
        "title": (df["consolidated"] + " | " + df["audited"] + " | "
                  + df["relatingTo"]).map(clean_text),
        "amount": None,
        "ratio": None,
        "url": df["xbrl"],
    })
    return df, uni


# Only these announcement categories are treated as meaningful signal; the rest
# (e.g. "Loss of Share Certificates") are dropped from the master timeline.
SIGNAL_DESC = {
    "Press Release", "Media Release", "Acquisition", "Financial Result Updates",
    "Investor Presentation", "Credit Rating", "Analysts/Institutional Investor "
    "Meet/Con. Call Updates", "Allotment of Securities",
    "Disclosure under SEBI Takeover Regulations", "Board Meeting",
    "Outcome of Board Meeting", "Dividend", "Bonus issue",
}


def load_announcements() -> tuple[pd.DataFrame, pd.DataFrame]:
    df = pd.read_csv(RAW_DIR / "announcements.csv", dtype=str).fillna("")
    df["desc"] = df["desc"].map(clean_text)
    df["attchmntText"] = df["attchmntText"].map(clean_text)
    df["event_date"] = df["sort_date"].map(parse_date)

    signal = df[df["desc"].isin(SIGNAL_DESC)].copy()

    uni = pd.DataFrame({
        "date": signal["event_date"],
        "symbol": signal["symbol"].replace("", SYMBOL),
        "isin": signal["sm_isin"],
        "source": "announcement",
        "event_type": signal["desc"].str.upper().str.replace(r"[^A-Z]+", "_", regex=True),
        "title": signal["attchmntText"].str.slice(0, 120),
        "amount": None,
        "ratio": None,
        "url": signal["attchmntFile"],
    })
    return df, uni


# ---------------------------------------------------------------------------
# 4. MAIN — run every loader, save cleaned files + the unified timeline
# ---------------------------------------------------------------------------
def main() -> None:
    print("Cleaning raw NSE event files for", SYMBOL, "...\n")

    ca_clean, ca_uni = load_corporate_actions()
    bm_clean, bm_uni = load_board_meetings()
    fr_clean, fr_uni = load_financial_results()
    an_clean, an_uni = load_announcements()

    # save the per-file cleaned versions
    ca_clean.to_csv(OUT_DIR / "corporate_actions_clean.csv", index=False)
    bm_clean.to_csv(OUT_DIR / "board_meetings_clean.csv", index=False)
    fr_clean.to_csv(OUT_DIR / "financial_results_clean.csv", index=False)
    an_clean.to_csv(OUT_DIR / "announcements_clean.csv", index=False)

    # build the master timeline
    unified = pd.concat([ca_uni, bm_uni, fr_uni, an_uni], ignore_index=True)
    unified = unified[UNIFIED_COLS]
    unified = unified.dropna(subset=["date"]).sort_values("date").reset_index(drop=True)
    unified.to_csv(OUT_DIR / "unified_events.csv", index=False)

    # --- console summary ---------------------------------------------------
    print(f"  corporate_actions : {len(ca_clean):4d} rows -> {len(ca_uni)} events")
    print(f"  board_meetings    : {len(bm_clean):4d} rows -> {len(bm_uni)} events "
          f"({int(bm_clean['bm_purpose'].str.contains('informed', case=False).sum())} swap-repaired)")
    print(f"  financial_results : {len(fr_clean):4d} rows (deduped) -> {len(fr_uni)} events")
    print(f"  announcements     : {len(an_clean):4d} rows -> {len(an_uni)} signal events "
          f"({len(an_clean) - len(an_uni)} noise rows dropped)")
    print(f"\n  MASTER TIMELINE   : {len(unified)} dated events "
          f"({unified['date'].min().date()} -> {unified['date'].max().date()})")
    print("\n  Event types in master timeline:")
    for etype, n in unified["event_type"].value_counts().head(15).items():
        print(f"      {n:4d}  {etype}")
    print(f"\nDone. Cleaned files written to: {OUT_DIR}")


if __name__ == "__main__":
    main()
