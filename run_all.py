"""
run_all.py
===============================================================================
Run the whole Nifty-50 analysis suite in order, with timing.

    python run_all.py               # everything
    python run_all.py --quick       # 2,000 permutations instead of 10,000
    python run_all.py --desc-only   # just the descriptive study (7s)
    python run_all.py --perms 50000 # heavier permutation testing

Every module defaults to "all Nifty 50 stocks that have downloaded feeds", so
no symbol list is needed. GPU is used automatically where it helps; the
descriptive study is I/O-bound and runs on CPU by design.
===============================================================================
"""

from __future__ import annotations

import argparse
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent

STEPS = [
    # (script, args, label, uses_gpu, note)
    ("results_dividend_behaviour.py", [], "Results & dividend behaviour",
     False, "descriptive: reaction, drop ratio, recovery curve"),
    ("event_stats.py", ["--perms", "{perms}"], "Statistical validation",
     True, "market model, BMP, Corrado, permutation, DSR, PBO"),
    ("event_ml.py", [], "ML surface model",
     True, "triple barrier, purged CV, CPCV, MDA"),
    ("disclosure_timing.py", ["--perms", "{perms}"], "Disclosure metadata",
     True, "filing-time burial, reporting lag, notice period"),
]


def gpu_banner():
    try:
        import torch
        if torch.cuda.is_available():
            p = torch.cuda.get_device_properties(0)
            return (f"{torch.cuda.get_device_name(0)}  sm_{p.major}{p.minor}  "
                    f"{p.total_memory / 1e9:.1f} GB  |  torch {torch.__version__}")
        return "CUDA NOT AVAILABLE — everything will fall back to CPU"
    except ImportError:
        return "torch not installed — GPU steps will fail"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--perms", type=int, default=10000)
    ap.add_argument("--quick", action="store_true")
    ap.add_argument("--desc-only", action="store_true")
    args = ap.parse_args()
    perms = 2000 if args.quick else args.perms

    steps = STEPS[:1] if args.desc_only else STEPS

    print("=" * 78)
    print("  NIFTY-50 ANALYSIS SUITE")
    print(f"  {gpu_banner()}")
    print(f"  permutations: {perms:,}")
    print("=" * 78)

    results, t_all = [], time.time()
    for i, (script, extra, label, gpu, note) in enumerate(steps, 1):
        cmd = [sys.executable, str(ROOT / script)] + \
              [a.format(perms=perms) for a in extra]
        tag = "GPU" if gpu else "CPU"
        print(f"\n[{i}/{len(steps)}] {label}  [{tag}]")
        print(f"        {note}")
        print(f"        $ {' '.join(cmd[1:])}")
        t0 = time.time()
        r = subprocess.run(cmd, cwd=ROOT)
        el = time.time() - t0
        ok = r.returncode == 0
        print(f"        {'done' if ok else 'FAILED'} in {el:.1f}s")
        results.append((label, ok, el))

    print("\n" + "=" * 78)
    print("  SUMMARY")
    print("=" * 78)
    for label, ok, el in results:
        print(f"  {'OK  ' if ok else 'FAIL'}  {label:<34} {el:>8.1f}s")
    print(f"  {'':6}{'total':<34} {time.time() - t_all:>8.1f}s")
    print("\n  Outputs in processed/:")
    for f in ("results_behaviour.csv", "dividend_behaviour.csv",
              "event_stats_summary.csv", "event_ml_cpcv.csv",
              "event_ml_mda.csv", "disclosure_timing.csv"):
        p = ROOT / "processed" / f
        mark = "yes" if p.exists() else " - "
        print(f"    [{mark}] {f}")


if __name__ == "__main__":
    main()
