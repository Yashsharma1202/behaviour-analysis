import pandas as pd, pathlib, sys

sys.stdout.reconfigure(errors='replace')

p_summary = pathlib.Path('D:/behaviour analysis/processed/event_behaviour_summary.csv')

print("==========================================================================================")
print("INSPECTING STOCK-SPECIFIC OPTIMAL POSITION TAKING WINDOWS")
print("==========================================================================================")

if p_summary.exists():
    df_sum = pd.read_csv(p_summary)
    print(f"Total Rows in summary: {len(df_sum)}")
    df_best = df_sum[df_sum['is_best'] == True]
    print(f"Total Best Window Rows: {len(df_best)}")
    print("\nSample Best Windows per Stock:")
    for _, r in df_best.head(15).iterrows():
        sym = r.get('symbol', 'UNKNOWN')
        etype = r.get('event_type', 'RESULTS')
        b_days = r.get('days_before_entry', 8)
        a_days = r.get('post_event_day', 1)
        wr = r.get('win_rate_pct', 55.0)
        print(f"  • {str(sym):12s} | Event: {etype:18s} | Window: T-{b_days} to T+{a_days} | Win Rate: {wr}%")

print("==========================================================================================")
