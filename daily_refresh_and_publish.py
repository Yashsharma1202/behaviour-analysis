"""
daily_refresh_and_publish.py
===============================================================================
ONE command that refreshes every live data source this project has, rebuilds
the dashboard, and publishes it -- the full pipeline this session built up
step by step, now chained into a single run.

This is the "refresh" button: there's no way to safely put a working refresh
button on the PUBLIC website itself (it would need push credentials embedded
in a page anyone can view -- that's not caution, it's a real hole). This
script is the real thing, meant to be triggered two ways:

  1. Automatically, daily -- action_cop/run_daily_scraper.bat already calls
     this after scrap.py, and that .bat is already wired into Windows Task
     Scheduler (9:30 AM / 9:00 PM, see action_cop/setup_task.ps1). No button
     needed for this path; it just happens.

  2. On demand -- double-click daily_refresh_and_publish.bat (same folder)
     any time you want an immediate refresh instead of waiting for the
     schedule. That IS a literal one-click button, it just lives on your
     machine, not the public page, because it needs this machine's Python
     environment, your quant_db access, and your git push credentials.

Steps, each independently error-handled (one failing step does not abort the
rest, and the final publish is skipped if the dashboard fails its own
consistency check):

    1. sync_live_corporate_feeds.py     (reads action_cop's fresh scrape)
    2. fetch_nifty50_result_dates.py    (live NSE announcement + event-calendar)
    3. sync_q3_live_nse_results.py      (pushes those into the dashboard)
    4. fix_q3_position_window.py        (per-stock windows for newly-announced)
    5. refresh_price_cache.py           (Yahoo, ~8 min for 211 stocks)
    6. fetch_futures_from_quant_db.py   (quant_db futures backfill, incremental)
    7. build_dashboard.py               (regenerate .js, sync docs/)
    8. build_dashboard.py --check       (gate: only publish if this passes)
    9. git add -A && commit && push     (both remotes) -- ONLY if step 8 passed

Usage:
    python daily_refresh_and_publish.py              # full run + publish
    python daily_refresh_and_publish.py --no-publish  # refresh only, no git push
    python daily_refresh_and_publish.py --skip-scraper-wait
                                                       # don't re-run scrap.py
                                                       # (assumes action_cop's
                                                       # own scrape already ran)
===============================================================================
"""
import argparse
import subprocess
import sys
import time
import pathlib
from datetime import datetime

ROOT = pathlib.Path(__file__).resolve().parent
LOG_PATH = ROOT / "daily_refresh_and_publish.log"


def log(msg):
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    print(line, flush=True)
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def run_step(label, cmd, cwd=None, timeout=1800):
    log(f"START  {label}  ({' '.join(cmd)})")
    t0 = time.time()
    try:
        result = subprocess.run(
            cmd, cwd=cwd or ROOT, timeout=timeout,
            capture_output=True, text=True, encoding="utf-8", errors="replace",
        )
        elapsed = time.time() - t0
        if result.returncode == 0:
            log(f"OK     {label}  ({elapsed:.0f}s)")
            return True
        else:
            log(f"FAILED {label}  (exit {result.returncode}, {elapsed:.0f}s)")
            log(f"  stderr tail: {result.stderr[-500:]}")
            return False
    except subprocess.TimeoutExpired:
        log(f"TIMEOUT {label} (> {timeout}s)")
        return False
    except Exception as e:
        log(f"ERROR  {label}  {type(e).__name__}: {e}")
        return False


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--no-publish", action="store_true", help="refresh everything but skip the git commit/push")
    args = ap.parse_args()

    log("=" * 70)
    log("DAILY REFRESH + PUBLISH -- starting")
    log("=" * 70)

    py = sys.executable

    # NOTE: fix_q3_position_window.py is intentionally NOT in this chain.
    # It re-optimizes each stock's T-n/T+m window from a walk-forward search,
    # which silently overwrote an intentionally-set window (TCS: T-2/T+4,
    # changed to T-6/T+1) the first time this pipeline ran it automatically.
    # Window choice is a strategy decision, not something that should get
    # silently rewritten by a daily cron job. Run it by hand if you want a
    # fresh optimization pass, and review what it changes before it's live.
    steps = [
        ("sync_live_corporate_feeds.py", [py, "sync_live_corporate_feeds.py"]),
        ("fetch_nifty50_result_dates.py", [py, "fetch_nifty50_result_dates.py"]),
        ("sync_q3_live_nse_results.py", [py, "sync_q3_live_nse_results.py"]),
        ("refresh_price_cache.py", [py, "refresh_price_cache.py"]),
        ("fetch_futures_from_quant_db.py", [py, "fetch_futures_from_quant_db.py"]),
        ("build_dashboard.py", [py, "build_dashboard.py"]),
    ]

    results = {}
    for label, cmd in steps:
        results[label] = run_step(label, cmd, timeout=1800)

    log("-" * 70)
    log("Step results: " + ", ".join(f"{k}={'OK' if v else 'FAIL'}" for k, v in results.items()))

    # Gate: only publish if the dashboard actually passes its own consistency check.
    check_ok = run_step("build_dashboard.py --check", [py, "build_dashboard.py", "--check"])

    if args.no_publish:
        log("--no-publish given: refreshed but not committing/pushing. Done.")
        return

    if not check_ok:
        log("build_dashboard.py --check FAILED -- refusing to publish a broken dashboard. "
            "Local files ARE updated; nothing was pushed. Investigate before re-running.")
        return

    log("Consistency check passed. Committing and pushing...")
    ts = datetime.now().strftime("%Y-%m-%d %H:%M")
    commit_msg = f"chore(auto-refresh): daily data refresh {ts}"

    add_ok = run_step("git add -A", ["git", "add", "-A"])
    if not add_ok:
        log("git add failed -- aborting publish.")
        return

    commit_result = subprocess.run(
        ["git", "commit", "-m", commit_msg], cwd=ROOT,
        capture_output=True, text=True, encoding="utf-8", errors="replace",
    )
    if commit_result.returncode != 0:
        if "nothing to commit" in (commit_result.stdout + commit_result.stderr).lower():
            log("Nothing changed since the last refresh -- nothing to publish. Done.")
            return
        log(f"git commit failed: {commit_result.stderr[-500:]}")
        return
    log("Committed: " + commit_msg)

    for remote in ("backup-origin", "origin"):
        push_ok = run_step(f"git push {remote} main", ["git", "push", remote, "main"])
        if not push_ok:
            log(f"WARNING: push to {remote} failed -- check network/auth. "
                f"The commit exists locally; retry with: git push {remote} main")

    log("=" * 70)
    log("DAILY REFRESH + PUBLISH -- finished")
    log("=" * 70)


if __name__ == "__main__":
    main()
