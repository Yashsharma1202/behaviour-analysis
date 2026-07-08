"""
inspect_csv.py
===============================================================================
STEP 1 — Data Schema Inspector for Options Straddle Backtest
-------------------------------------------------------------------------------
PURPOSE:
    Load the user-provided CSV file and produce a comprehensive summary of its
    structure so the backtest logic can be written with the EXACT column names
    and data shape confirmed.

HOW TO USE:
    1. Set CSV_PATH below to the full path of your data file.
    2. Run:  python inspect_csv.py
       OR:   python inspect_csv.py  path/to/your/file.csv
    3. Review the printed output and confirm the schema before the backtest
       script is written.
===============================================================================
"""

import sys
import textwrap
from pathlib import Path

import pandas as pd

# ─────────────────────────────────────────────────────────────────────────────
# ▶  SET YOUR CSV PATH HERE
# ─────────────────────────────────────────────────────────────────────────────
CSV_PATH = r""          # <- paste your file path between the quotes

# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────
SEP = "=" * 80

def banner(title):
    print(f"\n{SEP}")
    print(f"  {title}")
    print(SEP)

def section(title):
    print(f"\n{'─' * 70}")
    print(f"  {title}")
    print("─" * 70)

def _find_col(df, candidates):
    """Find a column by trying a list of candidate names (case-insensitive fallback)."""
    for c in candidates:
        if c in df.columns:
            return c
    lower_map = {col.lower(): col for col in df.columns}
    for c in candidates:
        if c.lower() in lower_map:
            return lower_map[c.lower()]
    return None


# ─────────────────────────────────────────────────────────────────────────────
# Main inspection routine
# ─────────────────────────────────────────────────────────────────────────────
def inspect(path):
    p = Path(path)
    if not p.exists():
        sys.exit(f"[ERROR] File not found: {path}")

    # ── 1. Load ────────────────────────────────────────────────────────────
    banner("OPTIONS DATA CSV — SCHEMA INSPECTION REPORT")
    print(f"  File : {p.resolve()}")
    print(f"  Size : {p.stat().st_size / 1_048_576:.2f} MB")

    df = pd.read_csv(path, low_memory=False)
    print(f"  Rows : {len(df):,}")
    print(f"  Cols : {df.shape[1]}")

    # ── 2. Column names & dtypes ───────────────────────────────────────────
    section("2. COLUMN NAMES & DATA TYPES")
    dtype_df = pd.DataFrame({
        "column":     df.columns.tolist(),
        "dtype":      [str(d) for d in df.dtypes],
        "non_null":   df.notna().sum().values,
        "null_count": df.isna().sum().values,
        "null_%":     (df.isna().mean() * 100).round(2).values,
    })
    print(dtype_df.to_string(index=False))

    # ── 3. Sample rows ─────────────────────────────────────────────────────
    section("3. FIRST 5 ROWS")
    pd.set_option("display.max_columns", None)
    pd.set_option("display.width", 200)
    print(df.head(5).to_string(index=False))

    section("3b. LAST 5 ROWS")
    print(df.tail(5).to_string(index=False))

    section("3c. 5 RANDOM ROWS")
    print(df.sample(min(5, len(df)), random_state=42).to_string(index=False))

    # ── 4. Date range ──────────────────────────────────────────────────────
    section("4. DATE RANGE")
    date_col = _find_col(df, ["date", "Date", "DATE", "trade_date", "trading_date"])
    if date_col:
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
        print(f"  Date column detected : '{date_col}'")
        print(f"  Earliest date        : {df[date_col].min().date()}")
        print(f"  Latest date          : {df[date_col].max().date()}")
        print(f"  Unique trading days  : {df[date_col].nunique():,}")
    else:
        print("  [WARN] No obvious date column found — check column names above.")

    # ── 5. Unique symbols ─────────────────────────────────────────────────
    section("5. UNIQUE SYMBOLS")
    sym_col = _find_col(df, ["symbol", "Symbol", "SYMBOL", "ticker", "underlying"])
    if sym_col:
        print(f"  Symbol column : '{sym_col}'")
        print(f"  Unique values : {df[sym_col].unique().tolist()}")
    else:
        print("  [WARN] No obvious symbol column found.")

    # ── 6. option_type breakdown ───────────────────────────────────────────
    section("6. OPTION TYPE BREAKDOWN (CE / PE / Spot detection)")
    ot_col = _find_col(df, ["option_type", "OptionType", "option type",
                             "CE_PE", "type", "instrument_type", "INSTRUMENT_TYPE"])
    if ot_col:
        vc = df[ot_col].value_counts(dropna=False)
        print(f"  option_type column : '{ot_col}'")
        print(vc.to_string())
        null_count = df[ot_col].isna().sum() + (df[ot_col].astype(str) == "").sum()
        print(f"\n  Rows with NULL/blank option_type (possible SPOT rows) : {null_count}")
        if null_count > 0:
            print("  -> SPOT rows appear to be MIXED into the SAME file with "
                  "NULL option_type.")
        else:
            print("  -> All rows have a non-null option_type. "
                  "Spot may be in a SEPARATE file.")
    else:
        print("  [WARN] No option_type column found. "
              "Cannot auto-detect CE vs PE vs Spot rows.")

    # ── 7. Expiry info ─────────────────────────────────────────────────────
    section("7. EXPIRY DATES")
    exp_col = _find_col(df, ["expiry", "Expiry", "EXPIRY", "expiry_date",
                              "ExpiryDate", "expiration"])
    if exp_col:
        df[exp_col] = pd.to_datetime(df[exp_col], errors="coerce")
        n_expiries = df[exp_col].nunique()
        print(f"  Expiry column    : '{exp_col}'")
        print(f"  Unique expiries  : {n_expiries}")
        all_exp = sorted(df[exp_col].dropna().unique())
        print(f"  All expiry dates : {[str(e)[:10] for e in all_exp[:30]]}")
        if n_expiries > 30:
            print(f"  ... and {n_expiries - 30} more")

        if date_col and ot_col:
            section("7b. UNIQUE EXPIRIES PER TRADE DATE (first 10 dates)")
            sample_dates = sorted(df[date_col].dropna().unique())[:10]
            for d in sample_dates:
                exps = sorted(df[df[date_col] == d][exp_col].dropna().unique())
                print(f"    {str(d)[:10]}  ->  {[str(e)[:10] for e in exps]}")
    else:
        print("  [WARN] No expiry column found.")

    # ── 8. Strike range ────────────────────────────────────────────────────
    section("8. STRIKE PRICE RANGE")
    strike_col = _find_col(df, ["strike", "Strike", "STRIKE", "strike_price",
                                 "StrikePrice"])
    if strike_col:
        df[strike_col] = pd.to_numeric(df[strike_col], errors="coerce")
        print(f"  Strike column : '{strike_col}'")
        print(f"  Min strike    : {df[strike_col].min()}")
        print(f"  Max strike    : {df[strike_col].max()}")
        print(f"  Unique strikes: {df[strike_col].nunique()}")
    else:
        print("  [WARN] No strike column found.")

    # ── 9. OHLC column stats ───────────────────────────────────────────────
    section("9. PRICE COLUMNS (OHLC)")
    for label in ["open", "Open", "OPEN", "high", "High", "HIGH",
                  "low", "Low", "LOW", "close", "Close", "CLOSE",
                  "volume", "Volume", "VOLUME", "oi", "OI", "open_interest"]:
        if label in df.columns:
            col_data = pd.to_numeric(df[label], errors="coerce")
            print(f"  '{label}'  ->  "
                  f"min={col_data.min():.2f}  "
                  f"max={col_data.max():.2f}  "
                  f"mean={col_data.mean():.2f}  "
                  f"nulls={col_data.isna().sum()}")

    # ── 10. Spot row sample (if mixed) ─────────────────────────────────────
    if ot_col and date_col:
        spot_mask = df[ot_col].isna() | (df[ot_col].astype(str).str.strip() == "")
        spot_rows = df[spot_mask]
        if len(spot_rows) > 0:
            section("10. SAMPLE SPOT ROWS (option_type is NULL/blank)")
            print(spot_rows.head(5).to_string(index=False))
        else:
            section("10. SPOT ROWS")
            print("  No null option_type rows found — spot data likely in separate file.")

    # ── 11. CE/PE samples ──────────────────────────────────────────────────
    if ot_col:
        for ot in ["CE", "PE"]:
            mask = df[ot_col].astype(str).str.upper() == ot
            sub = df[mask]
            if not sub.empty:
                section(f"11. SAMPLE {ot} ROWS (first 3)")
                print(sub.head(3).to_string(index=False))

    # ── 12. Close price distribution for options (to understand premium range)
    section("12. OPTION CLOSE PRICE DISTRIBUTION (CE & PE combined)")
    close_col = _find_col(df, ["close", "Close", "CLOSE"])
    if close_col and ot_col:
        opt_mask = df[ot_col].astype(str).str.upper().isin(["CE", "PE"])
        opt_close = pd.to_numeric(df.loc[opt_mask, close_col], errors="coerce").dropna()
        if not opt_close.empty:
            print(f"  Percentiles of option close prices:")
            for pct in [1, 5, 10, 25, 50, 75, 90, 95, 99]:
                print(f"    P{pct:>2}  :  Rs {opt_close.quantile(pct/100):.2f}")
            print(f"\n  Rows with close >= 15  :  {(opt_close >= 15).sum():,}")
            print(f"  Rows with close >= 20  :  {(opt_close >= 20).sum():,}")
            print(f"  Rows with close >= 25  :  {(opt_close >= 25).sum():,}")
            print(f"\n  [NOTE] Strategy will pick strikes with close NEAREST to and")
            print(f"         ABOVE Rs 20 (i.e., smallest value >= TARGET_PREMIUM).")

    # ── 13. Final checklist ────────────────────────────────────────────────
    section("13. COLUMN MAPPING CHECKLIST FOR BACKTEST SCRIPT")
    detected = {
        "date_col":              date_col or "[NOT FOUND]",
        "symbol_col":            sym_col or "[NOT FOUND]",
        "option_type_col":       ot_col or "[NOT FOUND]",
        "expiry_col":            exp_col or "[NOT FOUND]",
        "strike_col":            strike_col or "[NOT FOUND]",
        "close_col":             close_col or "[NOT FOUND]",
        "open_col":              _find_col(df, ["open", "Open", "OPEN"]) or "[NOT FOUND]",
        "high_col":              _find_col(df, ["high", "High", "HIGH"]) or "[NOT FOUND]",
        "low_col":               _find_col(df, ["low", "Low", "LOW"]) or "[NOT FOUND]",
        "volume_col":            _find_col(df, ["volume", "Volume", "VOLUME"]) or "[NOT FOUND]",
        "oi_col":                _find_col(df, ["oi", "OI", "open_interest"]) or "[NOT FOUND]",
        "spot_mixed_same_file":  (
            "YES (null option_type rows)" if ot_col and
            (df[ot_col].isna().any() or (df[ot_col].astype(str) == "").any())
            else "NO — likely a separate file"
        ),
    }
    for k, v in detected.items():
        status = "OK     " if "[NOT FOUND]" not in str(v) else "MISSING"
        print(f"  [{status}]  {k:<30}  ->  {v}")

    banner("INSPECTION COMPLETE")
    print(textwrap.dedent("""
      NEXT STEPS:
        1. Share this output to confirm or correct column name mappings.
        2. Confirm whether spot rows are in this file (null option_type) or
           a separate file.
        3. Confirm date range and symbol(s) to backtest.
        4. The full straddle backtest script will then be written with exact
           column names from your file.

      STRIKE SELECTION NOTE (corrected per user):
        -> Strikes will be selected as: NEAREST close price >= TARGET_PREMIUM
           (default Rs 20). This ensures the premium is never below the floor.
           If no strike >= TARGET_PREMIUM exists, the day is logged to skipped_days.csv.
    """))

    return df


# ─────────────────────────────────────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    if len(sys.argv) > 1:
        CSV_PATH = sys.argv[1]

    if not CSV_PATH:
        sys.exit(
            "[ERROR] No CSV path provided.\n"
            "  Option A: Set CSV_PATH at the top of inspect_csv.py\n"
            "  Option B: python inspect_csv.py  path/to/your/file.csv"
        )

    inspect(CSV_PATH)
