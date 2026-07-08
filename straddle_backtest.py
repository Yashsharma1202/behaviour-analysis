r"""
nifty_straddle_backtest.py
===============================================================================
NIFTY LONG STRADDLE BACKTEST  —  MINUTE-LEVEL DATA MODE
-------------------------------------------------------------------------------
DATA SOURCE  : D:\AMAN SIR SMC\options_parquet\  (one parquet per day)
               D:\AMAN SIR SMC\spot_parquet\     (one parquet per day)

STRATEGY     : LONG STRADDLE (BUY CE + BUY PE)
  Entry      : 3:25 PM on Day T  —  use 15:25:59 minute bar CLOSE price
  Strike     : nearest ABOVE Rs TARGET_PREMIUM (smallest close >= 20)
  Expiry     : Near = current week expiry | Next = following week expiry
               Nifty weekly expiry is TUESDAY (changed from Thursday)

TRACKING (Day T+1):
  9:15 AM    —  first minute bar OPEN  (earliest exit proxy)
  3:25 PM    —  last minute bar CLOSE  (EOD exit proxy)
  Peak       —  max combined straddle premium during the day
  Trough     —  min combined straddle premium during the day

OUTPUTS  (all saved to OUTPUT_DIR):
  trade_log_near.csv / trade_log_next.csv
  summary_overall_near.csv / summary_overall_next.csv
  summary_weekly.csv / summary_monthly.csv
  drawdown_top5.csv  (top 5 drawdowns: trough→peak tenure)
  expiry_comparison.csv
  cumulative_pnl.png / drawdown_chart.png / comparison_bar.png / spot_entries.png

RUN:
  python nifty_straddle_backtest.py
===============================================================================
"""

from __future__ import annotations
import warnings
import sys
from pathlib import Path
from datetime import timedelta

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.patches as mpatches

warnings.filterwarnings("ignore")

# ╔══════════════════════════════════════════════════════════════════════════╗
# ║                        CONFIGURATION                                    ║
# ╚══════════════════════════════════════════════════════════════════════════╝

OPTIONS_DIR     = Path(r"D:\AMAN SIR SMC\options_parquet")
SPOT_DIR        = Path(r"D:\AMAN SIR SMC\spot_parquet")
OUTPUT_DIR      = Path(r"D:\AMAN SIR SMC\analysis")

TARGET_PREMIUM  = 20.0          # Rs — select strike with close >= this
TOLERANCE       = 10.0          # Rs — max excess: accept [20, 30] range
                                #      set large (e.g. 9999) to disable cap

ENTRY_TIME      = "15:25:59"    # 3:25 PM bar  — entry candle
OPEN_TIME       = "09:15:59"    # 9:15 AM bar  — next-day open tracking

# Column names in parquet files
C_DATE     = "Date"
C_TIME     = "Time"
C_TYPE     = "Type"             # CE / PE
C_STRIKE   = "Strike"
C_EXPIRY   = "ExpiryDate"
C_CONTRACT = "ContractType"     # W=weekly, M=monthly
C_OPEN     = "Open"
C_HIGH     = "High"
C_LOW      = "Low"
C_CLOSE    = "Close"
C_INSTR    = "Instrument"

# Chart palette
BG       = "#0f131b"
PANEL    = "#171d2b"
GRID     = "#2a3346"
INK      = "#e6edf3"
MUTE     = "#8b98ad"
BLUE     = "#58a6ff"
GREEN    = "#3fb950"
AMBER    = "#d29922"
RED      = "#f85149"
PURPLE   = "#bc8cff"


# ╔══════════════════════════════════════════════════════════════════════════╗
# ║                      HELPER UTILITIES                                   ║
# ╚══════════════════════════════════════════════════════════════════════════╝

def _hdr(msg: str):
    print(f"\n{'═'*72}\n  {msg}\n{'═'*72}")

def _sec(msg: str):
    print(f"\n{'─'*60}\n  {msg}\n{'─'*60}")


def load_day_options(date_str: str) -> pd.DataFrame | None:
    """Load one day's options parquet. Return None if file missing."""
    p = OPTIONS_DIR / f"{date_str}.parquet"
    if not p.exists():
        return None
    df = pd.read_parquet(p)
    df[C_EXPIRY] = pd.to_datetime(df[C_EXPIRY], errors="coerce")
    df[C_STRIKE] = pd.to_numeric(df[C_STRIKE], errors="coerce")
    return df


def load_day_spot(date_str: str) -> pd.DataFrame | None:
    """Load one day's spot parquet. Return None if file missing."""
    p = SPOT_DIR / f"{date_str}.parquet"
    if not p.exists():
        return None
    return pd.read_parquet(p)


def get_entry_bar(day_df: pd.DataFrame, opt_type: str, expiry: pd.Timestamp,
                  time: str = ENTRY_TIME) -> pd.DataFrame:
    """Filter to a specific option type, expiry, and time bar."""
    mask = (
        (day_df[C_TYPE]   == opt_type) &
        (day_df[C_EXPIRY] == expiry) &
        (day_df[C_TIME]   == time)
    )
    return day_df[mask].copy()


def get_all_bars(day_df: pd.DataFrame, opt_type: str, expiry: pd.Timestamp,
                 strike: float) -> pd.DataFrame:
    """All minute bars for a specific strike/expiry/type."""
    mask = (
        (day_df[C_TYPE]   == opt_type) &
        (day_df[C_EXPIRY] == expiry) &
        (day_df[C_STRIKE] == strike)
    )
    return day_df[mask].copy()


def get_bar_at_time(day_df: pd.DataFrame, opt_type: str, expiry: pd.Timestamp,
                    strike: float, time: str) -> pd.Series | None:
    """Single bar for a strike/expiry/type at a specific time."""
    mask = (
        (day_df[C_TYPE]   == opt_type) &
        (day_df[C_EXPIRY] == expiry) &
        (day_df[C_STRIKE] == strike) &
        (day_df[C_TIME]   == time)
    )
    rows = day_df[mask]
    return rows.iloc[0] if not rows.empty else None


# ╔══════════════════════════════════════════════════════════════════════════╗
# ║                    EXPIRY SELECTION                                     ║
# ╚══════════════════════════════════════════════════════════════════════════╝

def get_weekly_expiries(day_df: pd.DataFrame, trade_date: pd.Timestamp
                        ) -> tuple[pd.Timestamp | None, pd.Timestamp | None]:
    """
    Return (near_expiry, next_expiry) from weekly options on this trade date.
    Only WEEKLY (ContractType == 'W') expiries are considered.
    Nifty weekly expiry changed to Tuesday — we just read what's in the data.
    """
    weekly = day_df[day_df[C_CONTRACT] == "W"]
    expiries = sorted(weekly[C_EXPIRY].dropna().unique())
    future   = [e for e in expiries if pd.Timestamp(e) >= trade_date]
    near = pd.Timestamp(future[0]) if len(future) > 0 else None
    nxt  = pd.Timestamp(future[1]) if len(future) > 1 else None
    return near, nxt


# ╔══════════════════════════════════════════════════════════════════════════╗
# ║                    ENTRY SELECTION (3:25 PM)                            ║
# ╚══════════════════════════════════════════════════════════════════════════╝

def select_entry(day_df: pd.DataFrame, expiry: pd.Timestamp
                 ) -> dict | None:
    """
    At 3:25 PM (ENTRY_TIME), find:
      CE strike: smallest close >= TARGET_PREMIUM (nearest above floor)
      PE strike: same rule

    Returns entry dict or None if no valid strike found.
    """
    result = {}
    for opt in ["CE", "PE"]:
        bars = get_entry_bar(day_df, opt, expiry, ENTRY_TIME)
        bars = bars.dropna(subset=[C_CLOSE, C_STRIKE])
        # Keep only strikes with close >= TARGET_PREMIUM
        valid = bars[bars[C_CLOSE] >= TARGET_PREMIUM].copy()
        if valid.empty:
            return None
        # Apply tolerance cap
        valid = valid[valid[C_CLOSE] <= TARGET_PREMIUM + TOLERANCE]
        if valid.empty:
            return None
        # Nearest above floor = minimum close among valid
        best = valid.loc[valid[C_CLOSE].idxmin()]
        result[opt] = {
            "strike":      float(best[C_STRIKE]),
            "entry_close": float(best[C_CLOSE]),
        }

    if "CE" not in result or "PE" not in result:
        return None

    return {
        "ce_strike":       result["CE"]["strike"],
        "ce_entry_price":  result["CE"]["entry_close"],
        "pe_strike":       result["PE"]["strike"],
        "pe_entry_price":  result["PE"]["entry_close"],
        "entry_combined":  round(result["CE"]["entry_close"] + result["PE"]["entry_close"], 2),
    }


# ╔══════════════════════════════════════════════════════════════════════════╗
# ║                  NEXT-DAY TRACKING                                      ║
# ╚══════════════════════════════════════════════════════════════════════════╝

def track_next_day(next_df: pd.DataFrame, expiry: pd.Timestamp,
                   ce_strike: float, pe_strike: float
                   ) -> dict | None:
    """
    For the locked CE and PE strikes, extract from Day T+1:
      - 9:15 AM open
      - 3:25 PM close
      - Intraday peak  (max of all minute-bar highs)
      - Intraday trough (min of all minute-bar lows)
    """
    data = {}
    for opt, strike in [("CE", ce_strike), ("PE", pe_strike)]:
        # 9:15 bar
        bar_open = get_bar_at_time(next_df, opt, expiry, strike, OPEN_TIME)
        # 3:25 bar
        bar_close = get_bar_at_time(next_df, opt, expiry, strike, ENTRY_TIME)
        # All bars (for peak / trough)
        all_bars = get_all_bars(next_df, opt, expiry, strike)

        if bar_open is None or bar_close is None or all_bars.empty:
            return None     # missing data for this strike

        data[opt] = {
            "open_915":   float(bar_open[C_OPEN]),
            "close_325":  float(bar_close[C_CLOSE]),
            "peak":       float(all_bars[C_HIGH].max()),    # best seen
            "trough":     float(all_bars[C_LOW].min()),     # worst seen
        }

    # Spot data
    spot_915  = float("nan")
    spot_325  = float("nan")
    spot_peak = float("nan")
    spot_trough = float("nan")
    spot_rows = next_df[next_df[C_TYPE] == "Index_Spot"] if "Index_Spot" in next_df[C_TYPE].values else pd.DataFrame()
    if not spot_rows.empty:
        s915 = spot_rows[spot_rows[C_TIME] == OPEN_TIME]
        s325 = spot_rows[spot_rows[C_TIME] == ENTRY_TIME]
        spot_915    = float(s915[C_OPEN].iloc[0])   if not s915.empty  else float("nan")
        spot_325    = float(s325[C_CLOSE].iloc[0])  if not s325.empty  else float("nan")
        spot_peak   = float(spot_rows[C_HIGH].max())
        spot_trough = float(spot_rows[C_LOW].min())

    return {
        "ce_open_915":   data["CE"]["open_915"],
        "ce_close_325":  data["CE"]["close_325"],
        "ce_peak":       data["CE"]["peak"],
        "ce_trough":     data["CE"]["trough"],
        "pe_open_915":   data["PE"]["open_915"],
        "pe_close_325":  data["PE"]["close_325"],
        "pe_peak":       data["PE"]["peak"],
        "pe_trough":     data["PE"]["trough"],
        "spot_open_915": spot_915,
        "spot_close_325": spot_325,
        "spot_peak":     spot_peak,
        "spot_trough":   spot_trough,
    }


# ╔══════════════════════════════════════════════════════════════════════════╗
# ║                    COMPUTE TRADE METRICS                                ║
# ╚══════════════════════════════════════════════════════════════════════════╝

def compute_metrics(entry: dict, tracking: dict) -> dict:
    """
    Compute all PnL points for one trade.

    LONG STRADDLE (buyer):
      PnL = next_day_combined - entry_combined
      Positive = profit (straddle gained value)
      Negative = loss  (straddle decayed / market flat)

    Peak PnL   = best combined premium seen next day - entry
    Trough PnL = worst combined premium next day - entry (max loss if sold there)
    """
    ec = entry["entry_combined"]

    combined_915  = round(tracking["ce_open_915"]  + tracking["pe_open_915"],  2)
    combined_325  = round(tracking["ce_close_325"] + tracking["pe_close_325"], 2)
    combined_peak = round(tracking["ce_peak"]      + tracking["pe_peak"],      2)
    combined_trough = round(tracking["ce_trough"]  + tracking["pe_trough"],    2)

    pnl_915   = round(combined_915   - ec, 2)
    pnl_325   = round(combined_325   - ec, 2)
    pnl_peak  = round(combined_peak  - ec, 2)   # best achievable
    pnl_trough = round(combined_trough - ec, 2)  # worst case in day

    pnl_pct_915   = round(pnl_915   / ec * 100, 2) if ec else 0
    pnl_pct_325   = round(pnl_325   / ec * 100, 2) if ec else 0
    pnl_pct_peak  = round(pnl_peak  / ec * 100, 2) if ec else 0
    pnl_pct_trough = round(pnl_trough / ec * 100, 2) if ec else 0

    win_loss = "WIN" if pnl_325 > 0 else ("LOSS" if pnl_325 < 0 else "FLAT")

    return {
        # Combined premiums
        "combined_entry":        ec,
        "combined_open_915":     combined_915,
        "combined_close_325":    combined_325,
        "combined_peak":         combined_peak,
        "combined_trough":       combined_trough,
        # PnL in Rs
        "pnl_at_915_rs":         pnl_915,
        "pnl_at_325_rs":         pnl_325,
        "pnl_peak_rs":           pnl_peak,
        "pnl_trough_rs":         pnl_trough,
        # PnL in %
        "pnl_at_915_pct":        pnl_pct_915,
        "pnl_at_325_pct":        pnl_pct_325,
        "pnl_peak_pct":          pnl_pct_peak,
        "pnl_trough_pct":        pnl_pct_trough,
        # Win/Loss
        "win_loss":              win_loss,
    }


# ╔══════════════════════════════════════════════════════════════════════════╗
# ║                     MAIN BACKTEST LOOP                                  ║
# ╚══════════════════════════════════════════════════════════════════════════╝

def run_backtest(expiry_rank: int, label: str
                 ) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Run the full backtest for one expiry type.

    Parameters
    ----------
    expiry_rank : 0 = near expiry, 1 = next expiry
    label       : 'Near' or 'Next'

    Returns
    -------
    results_df, skipped_df
    """
    _hdr(f"RUNNING BACKTEST — {label.upper()} EXPIRY")

    # Discover all trading days from options_parquet folder
    all_files   = sorted(OPTIONS_DIR.glob("*.parquet"))
    trade_dates = [f.stem for f in all_files]   # e.g. '2026-07-08'

    results  = []
    skipped  = []
    prev_exp = None  # carry-over expiry when T+1 file is from next day

    for i, date_str in enumerate(trade_dates):
        trade_date = pd.Timestamp(date_str)

        # ── Load Day T options ──────────────────────────────────────────
        day_df = load_day_options(date_str)
        if day_df is None:
            skipped.append({"entry_date": date_str, "expiry_type": label,
                            "reason": "options file missing"})
            continue

        # ── Get expiry ──────────────────────────────────────────────────
        near_exp, next_exp = get_weekly_expiries(day_df, trade_date)
        expiry = near_exp if expiry_rank == 0 else next_exp

        if expiry is None:
            skipped.append({"entry_date": date_str, "expiry_type": label,
                            "reason": f"no {'near' if expiry_rank==0 else 'next'} weekly expiry found"})
            continue

        # ── Select strikes at 3:25 PM ───────────────────────────────────
        entry = select_entry(day_df, expiry)
        if entry is None:
            skipped.append({
                "entry_date":  date_str,
                "expiry_type": label,
                "expiry_date": str(expiry.date()),
                "reason":      f"No CE/PE with close >= Rs{TARGET_PREMIUM} within ±Rs{TOLERANCE}",
            })
            continue

        # ── Spot price at 3:25 PM ───────────────────────────────────────
        spot_df_t = load_day_spot(date_str)
        spot_entry = float("nan")
        if spot_df_t is not None:
            sp_bar = spot_df_t[spot_df_t[C_TIME] == ENTRY_TIME]
            if not sp_bar.empty:
                spot_entry = float(sp_bar[C_CLOSE].iloc[0])

        # ── Find T+1 date ───────────────────────────────────────────────
        future = trade_dates[i + 1:]
        if not future:
            skipped.append({"entry_date": date_str, "expiry_type": label,
                            "reason": "no T+1 trading day (last file)"})
            continue
        next_date_str = future[0]
        next_date     = pd.Timestamp(next_date_str)

        # ── Load T+1 options ────────────────────────────────────────────
        next_opt_df = load_day_options(next_date_str)
        if next_opt_df is None:
            skipped.append({"entry_date": date_str, "expiry_type": label,
                            "reason": "T+1 options file missing"})
            continue

        # ── Track T+1 prices ────────────────────────────────────────────
        # Load spot for T+1 to embed in tracking
        next_spot_df = load_day_spot(next_date_str)
        # Merge spot into options df so track_next_day can use it
        if next_spot_df is not None:
            next_combined = pd.concat([next_opt_df, next_spot_df], ignore_index=True)
        else:
            next_combined = next_opt_df

        tracking = track_next_day(next_combined, expiry,
                                  entry["ce_strike"], entry["pe_strike"])
        if tracking is None:
            skipped.append({
                "entry_date":  date_str,
                "expiry_type": label,
                "expiry_date": str(expiry.date()),
                "reason":      "T+1 strike data missing (9:15/3:25 bar not found)",
            })
            continue

        # ── Compute metrics ─────────────────────────────────────────────
        metrics = compute_metrics(entry, tracking)

        is_expiry_day = next_date >= expiry

        # ── Assemble row ────────────────────────────────────────────────
        row = {
            "entry_date":           date_str,
            "next_date":            next_date_str,
            "expiry_type":          label,
            "expiry_date":          str(expiry.date()),
            "is_expiry_day":        is_expiry_day,
            "ce_strike":            entry["ce_strike"],
            "ce_entry_price_325":   entry["ce_entry_price"],
            "pe_strike":            entry["pe_strike"],
            "pe_entry_price_325":   entry["pe_entry_price"],
            # T+1 individual leg prices
            "ce_open_915":          tracking["ce_open_915"],
            "ce_close_325":         tracking["ce_close_325"],
            "ce_peak":              tracking["ce_peak"],
            "ce_trough":            tracking["ce_trough"],
            "pe_open_915":          tracking["pe_open_915"],
            "pe_close_325":         tracking["pe_close_325"],
            "pe_peak":              tracking["pe_peak"],
            "pe_trough":            tracking["pe_trough"],
            # Spot
            "spot_at_entry_325":    round(spot_entry, 2),
            "spot_open_915":        round(tracking["spot_open_915"],   2),
            "spot_close_325":       round(tracking["spot_close_325"],  2),
            "spot_peak":            round(tracking["spot_peak"],        2),
            "spot_trough":          round(tracking["spot_trough"],      2),
            **metrics,
        }
        results.append(row)

        # Progress
        if (i + 1) % 100 == 0 or i == len(trade_dates) - 1:
            print(f"  [{label}]  Day {i+1}/{len(trade_dates)}  |  "
                  f"Trades: {len(results)}  |  Skipped: {len(skipped)}")

    results_df = pd.DataFrame(results)
    skipped_df = pd.DataFrame(skipped)
    print(f"\n  DONE — {len(results_df):,} trades | {len(skipped_df):,} skipped")
    return results_df, skipped_df


# ╔══════════════════════════════════════════════════════════════════════════╗
# ║                     SUMMARIZE                                           ║
# ╚══════════════════════════════════════════════════════════════════════════╝

def summarize_overall(df: pd.DataFrame, label: str) -> pd.DataFrame:
    """Overall aggregate statistics."""
    if df.empty:
        return pd.DataFrame()

    pnl   = df["pnl_at_325_rs"]
    wins  = (df["win_loss"] == "WIN").sum()
    total = len(df)

    stats = {
        "expiry_type":              label,
        "total_trades":             total,
        "wins":                     int(wins),
        "losses":                   int((df["win_loss"] == "LOSS").sum()),
        "flat":                     int((df["win_loss"] == "FLAT").sum()),
        "win_rate_%":               round(wins / total * 100, 2),
        "avg_pnl_325_rs":           round(pnl.mean(), 2),
        "median_pnl_325_rs":        round(pnl.median(), 2),
        "total_cum_pnl_rs":         round(pnl.sum(), 2),
        "best_day_pnl_rs":          round(pnl.max(), 2),
        "worst_day_pnl_rs":         round(pnl.min(), 2),
        "std_pnl_rs":               round(pnl.std(), 2),
        "avg_pnl_at_915_rs":        round(df["pnl_at_915_rs"].mean(), 2),
        "avg_peak_pnl_rs":          round(df["pnl_peak_rs"].mean(), 2),
        "avg_trough_pnl_rs":        round(df["pnl_trough_rs"].mean(), 2),
        "avg_entry_combined_rs":    round(df["combined_entry"].mean(), 2),
        "expiry_day_trades":        int(df["is_expiry_day"].sum()),
        "data_start":               str(df["entry_date"].min()),
        "data_end":                 str(df["entry_date"].max()),
    }

    _sec(f"OVERALL SUMMARY — {label.upper()} EXPIRY")
    for k, v in stats.items():
        print(f"  {k:<35}  {v}")
    return pd.DataFrame([stats])


def summarize_weekly(df: pd.DataFrame, label: str) -> pd.DataFrame:
    """Week-by-week aggregation."""
    if df.empty:
        return pd.DataFrame()
    df = df.copy()
    df["entry_date"] = pd.to_datetime(df["entry_date"])
    df["week"] = df["entry_date"].dt.to_period("W")
    grp = df.groupby("week").agg(
        trades      =("pnl_at_325_rs", "count"),
        wins        =("win_loss",       lambda x: (x == "WIN").sum()),
        total_pnl   =("pnl_at_325_rs", "sum"),
        avg_pnl     =("pnl_at_325_rs", "mean"),
        best_day    =("pnl_at_325_rs", "max"),
        worst_day   =("pnl_at_325_rs", "min"),
        avg_peak    =("pnl_peak_rs",   "mean"),
        avg_trough  =("pnl_trough_rs", "mean"),
    ).reset_index()
    grp["win_rate_%"] = (grp["wins"] / grp["trades"] * 100).round(2)
    grp["expiry_type"] = label
    grp["week"] = grp["week"].astype(str)
    return grp


def summarize_monthly(df: pd.DataFrame, label: str) -> pd.DataFrame:
    """Month-by-month aggregation."""
    if df.empty:
        return pd.DataFrame()
    df = df.copy()
    df["entry_date"] = pd.to_datetime(df["entry_date"])
    df["month"] = df["entry_date"].dt.to_period("M")
    grp = df.groupby("month").agg(
        trades      =("pnl_at_325_rs", "count"),
        wins        =("win_loss",       lambda x: (x == "WIN").sum()),
        total_pnl   =("pnl_at_325_rs", "sum"),
        avg_pnl     =("pnl_at_325_rs", "mean"),
        best_day    =("pnl_at_325_rs", "max"),
        worst_day   =("pnl_at_325_rs", "min"),
        avg_peak    =("pnl_peak_rs",   "mean"),
        avg_trough  =("pnl_trough_rs", "mean"),
    ).reset_index()
    grp["win_rate_%"] = (grp["wins"] / grp["trades"] * 100).round(2)
    grp["expiry_type"] = label
    grp["month"] = grp["month"].astype(str)
    return grp


# ╔══════════════════════════════════════════════════════════════════════════╗
# ║                  DRAWDOWN ANALYSIS (TOP 5)                              ║
# ╚══════════════════════════════════════════════════════════════════════════╝

def _dd_color(pct: float) -> str:
    """Color grade based on drawdown depth %."""
    if pct <= 10:   return "GREEN    — Small"
    if pct <= 25:   return "YELLOW   — Moderate"
    if pct <= 50:   return "ORANGE   — High"
    if pct <= 75:   return "RED      — Severe"
    return          "DARK_RED — Extreme"


def drawdown_analysis(df: pd.DataFrame, label: str, top_n: int = 5) -> pd.DataFrame:
    """
    Compute cumulative PnL drawdown series.
    Identify top N drawdown episodes:
      - Start  : last peak before the trough
      - Trough : lowest cumulative PnL point
      - Recovery : date when cumulative PnL exceeds the prior peak again

    Parameters
    ----------
    df     : results dataframe with pnl_at_325_rs column
    label  : Near / Next
    top_n  : how many worst drawdowns to return (default 5)
    """
    if df.empty:
        return pd.DataFrame()

    df = df.copy()
    df["entry_date"] = pd.to_datetime(df["entry_date"])
    df = df.sort_values("entry_date").reset_index(drop=True)
    df["cum_pnl"] = df["pnl_at_325_rs"].cumsum()

    # Running maximum (peak of equity curve)
    df["running_peak"] = df["cum_pnl"].cummax()
    df["drawdown_rs"]  = df["cum_pnl"] - df["running_peak"]   # always <= 0
    df["drawdown_pct"] = (df["drawdown_rs"] / df["running_peak"].abs().replace(0, np.nan) * 100)

    # Identify drawdown periods
    episodes = []
    in_dd    = False
    start_i  = None
    peak_val = 0.0

    for idx, row in df.iterrows():
        if not in_dd:
            if row["drawdown_rs"] < 0:
                in_dd    = True
                start_i  = idx
                peak_val = row["running_peak"]
        else:
            if row["cum_pnl"] >= peak_val:
                # Recovery found
                trough_i = df.loc[start_i:idx, "cum_pnl"].idxmin()
                episodes.append({
                    "peak_date":      str(df.loc[start_i, "entry_date"].date()),
                    "trough_date":    str(df.loc[trough_i, "entry_date"].date()),
                    "recovery_date":  str(row["entry_date"].date()),
                    "peak_cum_pnl":   round(peak_val, 2),
                    "trough_cum_pnl": round(df.loc[trough_i, "cum_pnl"], 2),
                    "dd_rs":          round(df.loc[trough_i, "drawdown_rs"], 2),
                    "dd_pct":         round(abs(df.loc[trough_i, "drawdown_pct"]), 2)
                                      if not pd.isna(df.loc[trough_i, "drawdown_pct"]) else 0.0,
                    "tenure_days":    (df.loc[trough_i, "entry_date"] -
                                       df.loc[start_i, "entry_date"]).days,
                    "recovery_days":  (row["entry_date"] -
                                       df.loc[trough_i, "entry_date"]).days,
                    "total_dd_days":  (row["entry_date"] -
                                       df.loc[start_i, "entry_date"]).days,
                })
                in_dd = False

    # Handle open drawdown at end of data
    if in_dd:
        trough_i = df.loc[start_i:, "cum_pnl"].idxmin()
        episodes.append({
            "peak_date":      str(df.loc[start_i, "entry_date"].date()),
            "trough_date":    str(df.loc[trough_i, "entry_date"].date()),
            "recovery_date":  "NOT YET RECOVERED",
            "peak_cum_pnl":   round(peak_val, 2),
            "trough_cum_pnl": round(df.loc[trough_i, "cum_pnl"], 2),
            "dd_rs":          round(df.loc[trough_i, "drawdown_rs"], 2),
            "dd_pct":         round(abs(df.loc[trough_i, "drawdown_pct"]), 2)
                              if not pd.isna(df.loc[trough_i, "drawdown_pct"]) else 0.0,
            "tenure_days":    (df.loc[trough_i, "entry_date"] -
                               df.loc[start_i, "entry_date"]).days,
            "recovery_days":  None,
            "total_dd_days":  None,
        })

    if not episodes:
        return pd.DataFrame()

    dd_df = pd.DataFrame(episodes)
    dd_df = dd_df.sort_values("dd_rs").head(top_n).reset_index(drop=True)
    dd_df["rank"]         = dd_df.index + 1
    dd_df["color_grade"]  = dd_df["dd_pct"].apply(_dd_color)
    dd_df["expiry_type"]  = label

    _sec(f"TOP {top_n} DRAWDOWNS — {label.upper()} EXPIRY")
    cols_show = ["rank", "peak_date", "trough_date", "recovery_date",
                 "dd_rs", "dd_pct", "tenure_days", "recovery_days",
                 "total_dd_days", "color_grade"]
    print(dd_df[cols_show].to_string(index=False))
    return dd_df


# ╔══════════════════════════════════════════════════════════════════════════╗
# ║                     COMPARE EXPIRIES                                    ║
# ╚══════════════════════════════════════════════════════════════════════════╝

def compare_expiries(near_s: pd.DataFrame, next_s: pd.DataFrame) -> pd.DataFrame:
    """Side-by-side near vs next comparison."""
    if near_s.empty or next_s.empty:
        return pd.DataFrame()
    comp = pd.merge(
        near_s.T.reset_index().rename(columns={"index": "metric", 0: "near_expiry"}),
        next_s.T.reset_index().rename(columns={"index": "metric", 0: "next_expiry"}),
        on="metric", how="outer"
    )
    _sec("NEAR vs NEXT EXPIRY — SIDE BY SIDE COMPARISON")
    print(comp.to_string(index=False))
    return comp


# ╔══════════════════════════════════════════════════════════════════════════╗
# ║                         CHARTS                                          ║
# ╚══════════════════════════════════════════════════════════════════════════╝

def make_charts(near_df: pd.DataFrame, next_df: pd.DataFrame,
                near_dd: pd.DataFrame, next_dd: pd.DataFrame,
                out_dir: Path) -> None:
    """Generate and save all 4 charts."""
    _sec("GENERATING CHARTS")
    plt.rcParams.update({"font.family": "Segoe UI", "axes.unicode_minus": False})

    def _style_ax(ax):
        ax.set_facecolor(PANEL)
        ax.tick_params(colors=MUTE, labelsize=8)
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%b-%y"))
        for spine in ax.spines.values():
            spine.set_color(GRID)
        ax.grid(axis="y", color=GRID, linewidth=0.5, linestyle="--")

    # ── 1. Cumulative PnL ─────────────────────────────────────────────────
    fig, ax = plt.subplots(figsize=(14, 5))
    fig.patch.set_facecolor(BG)
    _style_ax(ax)

    for df, col, lbl in [(near_df, BLUE, "Near Expiry"),
                          (next_df, GREEN, "Next Expiry")]:
        if not df.empty:
            dates   = pd.to_datetime(df["entry_date"])
            cum_pnl = df["pnl_at_325_rs"].cumsum()
            ax.plot(dates, cum_pnl, color=col, linewidth=2, label=lbl, zorder=3)
            ax.fill_between(dates, cum_pnl, alpha=0.07, color=col)

    ax.axhline(0, color=MUTE, linewidth=1, linestyle="--")
    ax.set_title("Cumulative PnL — NIFTY Long Straddle | Near vs Next Expiry",
                 color=INK, fontsize=12, pad=10)
    ax.set_xlabel("Date", color=MUTE, fontsize=9)
    ax.set_ylabel("Cumulative PnL (₹)", color=MUTE, fontsize=9)
    ax.legend(facecolor=PANEL, edgecolor=GRID, labelcolor=INK, fontsize=9)
    fig.autofmt_xdate()
    plt.tight_layout()
    fig.savefig(out_dir / "cumulative_pnl.png", dpi=150, bbox_inches="tight",
                facecolor=BG)
    plt.close(fig)
    print("  Saved: cumulative_pnl.png")

    # ── 2. Drawdown Chart ─────────────────────────────────────────────────
    fig, axes = plt.subplots(2, 1, figsize=(14, 8), sharex=False)
    fig.patch.set_facecolor(BG)

    for ax, df, col, lbl in zip(axes,
                                  [near_df, next_df],
                                  [BLUE, GREEN],
                                  ["Near Expiry", "Next Expiry"]):
        _style_ax(ax)
        if not df.empty:
            dates    = pd.to_datetime(df["entry_date"])
            cum_pnl  = df["pnl_at_325_rs"].cumsum()
            peak_eq  = cum_pnl.cummax()
            dd_curve = cum_pnl - peak_eq
            ax.fill_between(dates, dd_curve, 0, color=RED, alpha=0.4)
            ax.plot(dates, dd_curve, color=RED, linewidth=1)
            ax.set_title(f"Drawdown Curve — {lbl}", color=INK, fontsize=10)
            ax.set_ylabel("Drawdown (₹)", color=MUTE, fontsize=9)
            ax.xaxis.set_major_formatter(mdates.DateFormatter("%b-%y"))
        fig.autofmt_xdate()

    fig.suptitle("Drawdown Analysis — NIFTY Long Straddle", color=INK,
                 fontsize=13, y=1.01)
    plt.tight_layout()
    fig.savefig(out_dir / "drawdown_chart.png", dpi=150, bbox_inches="tight",
                facecolor=BG)
    plt.close(fig)
    print("  Saved: drawdown_chart.png")

    # ── 3. Comparison Bar ─────────────────────────────────────────────────
    def _safe_stat(df, col):
        try:
            return float(df[col].iloc[0])
        except Exception:
            return 0.0

    # Recompute quick stats inline
    def _quick(df):
        if df.empty:
            return {}
        pnl = df["pnl_at_325_rs"]
        wins = (df["win_loss"] == "WIN").sum()
        return {
            "Win Rate (%)":      round(wins / len(df) * 100, 2),
            "Avg PnL (₹)":      round(pnl.mean(), 2),
            "Avg Trough (₹)":   round(df["pnl_trough_rs"].mean(), 2),
            "Avg Peak (₹)":     round(df["pnl_peak_rs"].mean(), 2),
            "Total PnL (₹)":    round(pnl.sum(), 2),
        }

    near_stats = _quick(near_df)
    next_stats = _quick(next_df)
    metrics    = list(near_stats.keys())

    fig, axes = plt.subplots(1, len(metrics), figsize=(16, 5))
    fig.patch.set_facecolor(BG)

    for ax, m in zip(axes, metrics):
        _style_ax(ax)
        nv = near_stats.get(m, 0)
        xv = next_stats.get(m, 0)
        bars = ax.bar(["Near", "Next"], [nv, xv],
                      color=[BLUE, GREEN], width=0.5, edgecolor=GRID)
        ax.bar_label(bars, fmt="%.1f", color=INK, fontsize=8, padding=3)
        ax.set_title(m, color=INK, fontsize=9)
        ax.axhline(0, color=GRID, linewidth=0.8)
        ax.tick_params(colors=MUTE, labelsize=8)

    fig.suptitle("Near vs Next Expiry — Key Metrics", color=INK, fontsize=13)
    plt.tight_layout()
    fig.savefig(out_dir / "comparison_bar.png", dpi=150, bbox_inches="tight",
                facecolor=BG)
    plt.close(fig)
    print("  Saved: comparison_bar.png")

    # ── 4. Spot Price + Entry Dates ───────────────────────────────────────
    fig, ax = plt.subplots(figsize=(14, 5))
    fig.patch.set_facecolor(BG)
    _style_ax(ax)

    # Spot curve from near_df spot prices
    if not near_df.empty:
        sp_dates  = pd.to_datetime(near_df["entry_date"])
        sp_prices = near_df["spot_at_entry_325"]
        ax.plot(sp_dates, sp_prices, color=AMBER, linewidth=1.5, label="Spot (3:25 PM)", zorder=2)
        # Entry markers — coloured by win/loss
        for _, row in near_df.iterrows():
            col = GREEN if row["win_loss"] == "WIN" else (RED if row["win_loss"] == "LOSS" else MUTE)
            ax.scatter(pd.Timestamp(row["entry_date"]), row["spot_at_entry_325"],
                       color=col, s=18, zorder=4, alpha=0.8)

    legend_patches = [
        mpatches.Patch(color=GREEN, label="WIN entry"),
        mpatches.Patch(color=RED,   label="LOSS entry"),
        mpatches.Patch(color=AMBER, label="Spot Price"),
    ]
    ax.legend(handles=legend_patches, facecolor=PANEL, edgecolor=GRID,
              labelcolor=INK, fontsize=9)
    ax.set_title("Nifty Spot with Straddle Entry Dates (Near Expiry) — Green=WIN, Red=LOSS",
                 color=INK, fontsize=11, pad=10)
    ax.set_xlabel("Date", color=MUTE, fontsize=9)
    ax.set_ylabel("Spot Price (₹)", color=MUTE, fontsize=9)
    fig.autofmt_xdate()
    plt.tight_layout()
    fig.savefig(out_dir / "spot_entries.png", dpi=150, bbox_inches="tight",
                facecolor=BG)
    plt.close(fig)
    print("  Saved: spot_entries.png")


# ╔══════════════════════════════════════════════════════════════════════════╗
# ║                     SAVE ALL OUTPUTS                                    ║
# ╚══════════════════════════════════════════════════════════════════════════╝

def save_outputs(near_results, next_results,
                 near_summary, next_summary,
                 near_weekly,  next_weekly,
                 near_monthly, next_monthly,
                 near_dd,      next_dd,
                 comparison,   all_skipped,
                 out_dir: Path) -> None:
    _sec("SAVING CSV FILES")
    out_dir.mkdir(parents=True, exist_ok=True)

    files = {
        "trade_log_near.csv":          near_results,
        "trade_log_next.csv":          next_results,
        "summary_overall_near.csv":    near_summary,
        "summary_overall_next.csv":    next_summary,
        "summary_weekly_near.csv":     near_weekly,
        "summary_weekly_next.csv":     next_weekly,
        "summary_monthly_near.csv":    near_monthly,
        "summary_monthly_next.csv":    next_monthly,
        "drawdown_top5_near.csv":      near_dd,
        "drawdown_top5_next.csv":      next_dd,
        "expiry_comparison.csv":       comparison,
        "skipped_days.csv":            all_skipped,
    }
    for fname, df in files.items():
        if df is not None and not (isinstance(df, pd.DataFrame) and df.empty):
            p = out_dir / fname
            df.to_csv(p, index=False)
            kb = p.stat().st_size / 1024
            print(f"  Saved: {fname:<45}  {kb:>7.1f} KB")


# ╔══════════════════════════════════════════════════════════════════════════╗
# ║                           MAIN                                          ║
# ╚══════════════════════════════════════════════════════════════════════════╝

def main():
    _hdr("NIFTY LONG STRADDLE BACKTEST — INTRADAY MINUTE DATA")
    print(f"""
  ┌─────────────────────────────────────────────────────────────────┐
  │  DATA MODE   : Intraday minute bars                             │
  │  STRATEGY    : Long Straddle (BUY CE + BUY PE)                  │
  │  ENTRY       : 3:25 PM close (15:25:59 bar)                     │
  │  TRACKING    : T+1 at 9:15 AM open, 3:25 PM close,             │
  │                intraday peak and trough                         │
  │  EXPIRY      : Weekly (Nifty Tuesday) — Near & Next             │
  │  STRIKE RULE : Nearest close >= Rs {TARGET_PREMIUM:<6}  max Rs {TARGET_PREMIUM+TOLERANCE:.0f}    │
  └─────────────────────────────────────────────────────────────────┘
  Options folder : {OPTIONS_DIR}
  Spot folder    : {SPOT_DIR}
  Output folder  : {OUTPUT_DIR}
    """)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # ── Backtest runs ──────────────────────────────────────────────────────
    near_results, near_skipped = run_backtest(0, "Near")
    next_results, next_skipped = run_backtest(1, "Next")
    all_skipped = pd.concat([near_skipped, next_skipped], ignore_index=True)

    # ── Summaries ──────────────────────────────────────────────────────────
    near_summary = summarize_overall(near_results, "Near")
    next_summary = summarize_overall(next_results, "Next")
    near_weekly  = summarize_weekly(near_results,  "Near")
    next_weekly  = summarize_weekly(next_results,  "Next")
    near_monthly = summarize_monthly(near_results, "Near")
    next_monthly = summarize_monthly(next_results, "Next")

    # ── Drawdown analysis ─────────────────────────────────────────────────
    near_dd = drawdown_analysis(near_results, "Near", top_n=5)
    next_dd = drawdown_analysis(next_results, "Next", top_n=5)

    # ── Comparison ────────────────────────────────────────────────────────
    comparison = compare_expiries(near_summary, next_summary)

    # ── Skipped days ──────────────────────────────────────────────────────
    if not all_skipped.empty:
        _sec("SKIPPED DAYS SAMPLE (first 20)")
        print(all_skipped.head(20).to_string(index=False))
        print(f"\n  Total skipped: {len(all_skipped)}")

    # ── Charts ────────────────────────────────────────────────────────────
    make_charts(near_results, next_results, near_dd, next_dd, OUTPUT_DIR)

    # ── Save CSVs ─────────────────────────────────────────────────────────
    save_outputs(
        near_results, next_results,
        near_summary, next_summary,
        near_weekly,  next_weekly,
        near_monthly, next_monthly,
        near_dd,      next_dd,
        comparison,   all_skipped,
        OUTPUT_DIR
    )

    _hdr("ALL DONE")
    print(f"\n  All files saved to: {OUTPUT_DIR.resolve()}\n")
    print("  Output files:")
    for f in sorted(OUTPUT_DIR.iterdir()):
        kb = f.stat().st_size / 1024
        print(f"    {f.name:<48}  {kb:>8.1f} KB")


if __name__ == "__main__":
    main()
