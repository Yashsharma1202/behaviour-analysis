"""
RBI Monetary Policy Event — Pre/Post Window Backtest Engine
Analyses T-3 to T+5 day windows around each RBI policy date
Generates: Excel report + JSON data for dashboard
"""

import pandas as pd
import numpy as np
import json
import os
from datetime import datetime, timedelta

# ── Paths ──────────────────────────────────────────────────────────────────────
BASE = r"D:\behaviour analysis"
PARQUET_DIR = os.path.join(BASE, "scraped_parquet")
OUT_DIR = BASE
DOCS_DATA = os.path.join(BASE, "docs", "dashboard_data")

# ── Load Data ──────────────────────────────────────────────────────────────────
print("Loading data...")
rbi_dates = pd.read_parquet(os.path.join(PARQUET_DIR, "rbi_monetary_policy_dates.parquet"))
rbi_reaction = pd.read_parquet(os.path.join(PARQUET_DIR, "rbi_policy_market_reaction_2000_2026.parquet"))
fut = pd.read_parquet(os.path.join(PARQUET_DIR, "nifty_futures_near_month_continuous.parquet"))

# ── Prep Futures OHLC ─────────────────────────────────────────────────────────
fut['DATE'] = pd.to_datetime(fut['DATE'])
fut = fut.sort_values('DATE').reset_index(drop=True)
fut_close = fut.set_index('DATE')['CLOSE']
fut_open = fut.set_index('DATE')['OPEN']
trading_days = fut['DATE'].sort_values().reset_index(drop=True)

def get_nth_trading_day(ref_date, n):
    """Get Nth trading day relative to ref_date (n can be negative)"""
    ref_date = pd.Timestamp(ref_date)
    idx_list = trading_days[trading_days == ref_date].index
    if len(idx_list) == 0:
        # Find nearest trading day
        nearest = trading_days[trading_days >= ref_date]
        if len(nearest) == 0:
            return None
        idx = nearest.index[0]
    else:
        idx = idx_list[0]
    target_idx = idx + n
    if target_idx < 0 or target_idx >= len(trading_days):
        return None
    return trading_days.iloc[target_idx]

def get_price(date, price_series, use_open=False):
    """Get close (or open) price for a date"""
    if date is None:
        return np.nan
    date = pd.Timestamp(date)
    if date in price_series.index:
        return price_series[date]
    # Try ±1 day
    for delta in [1, -1, 2, -2]:
        d = date + timedelta(days=delta)
        if d in price_series.index:
            return price_series[d]
    return np.nan

# ── Categorize Actions ────────────────────────────────────────────────────────
def categorize(action):
    a = str(action).upper()
    if 'CUT' in a: return 'CUT'
    elif 'HIKE' in a: return 'HIKE'
    elif 'STATUS QUO' in a or 'NO CHANGE' in a or 'UNCHANGED' in a: return 'STATUS QUO'
    else: return 'OTHER'

# ── Window Definitions ────────────────────────────────────────────────────────
# (entry_offset, exit_offset) relative to T=0 (policy date)
WINDOWS = [
    (-3, +1), (-3, +2), (-3, +3), (-3, +5),
    (-2, +1), (-2, +2), (-2, +3), (-2, +5),
    (-1, +1), (-1, +2), (-1, +3), (-1, +5),
    ( 0, +1), ( 0, +2), ( 0, +3), ( 0, +5),
    (+1, +2), (+1, +3), (+1, +5),
]

# ── Build Per-Event Per-Window Trade Log ──────────────────────────────────────
print("Running pre/post window backtest...")
rbi_reaction['POLICY_DATE'] = pd.to_datetime(rbi_reaction['POLICY_DATE'])
rbi_reaction['CATEGORY'] = rbi_reaction['ACTION'].apply(categorize)
rbi_reaction['EFFECTIVE_TRADING_DATE'] = pd.to_datetime(rbi_reaction['EFFECTIVE_TRADING_DATE'])

all_trades = []

for _, row in rbi_reaction.iterrows():
    eff_date = row['EFFECTIVE_TRADING_DATE']  # T=0 (effective market day)
    policy_date = row['POLICY_DATE']
    category = row['CATEGORY']
    action = row['ACTION']
    repo_rate = row['REPO_RATE']
    year = row['YEAR']
    spot_day_chg = row['SPOT_DAY_CHANGE_PCT']
    fut_day_chg = row['FUT_DAY_CHANGE_PCT']

    for (entry_off, exit_off) in WINDOWS:
        entry_date = get_nth_trading_day(eff_date, entry_off)
        exit_date = get_nth_trading_day(eff_date, exit_off)

        # Entry: use open price on entry day; Exit: use close on exit day
        entry_price = get_price(entry_date, fut_open if entry_off < 0 else fut_close)
        exit_price = get_price(exit_date, fut_close)

        if pd.isna(entry_price) or pd.isna(exit_price) or entry_price == 0:
            pnl_pct = np.nan
        else:
            pnl_pct = ((exit_price - entry_price) / entry_price) * 100

        all_trades.append({
            'POLICY_DATE': policy_date.strftime('%Y-%m-%d'),
            'EFF_DATE': eff_date.strftime('%Y-%m-%d') if pd.notna(eff_date) else '',
            'YEAR': year,
            'ACTION': action,
            'CATEGORY': category,
            'REPO_RATE': repo_rate,
            'WINDOW': f'T{entry_off:+d} to T{exit_off:+d}',
            'ENTRY_OFFSET': entry_off,
            'EXIT_OFFSET': exit_off,
            'ENTRY_DATE': entry_date.strftime('%Y-%m-%d') if entry_date is not None else '',
            'EXIT_DATE': exit_date.strftime('%Y-%m-%d') if exit_date is not None else '',
            'ENTRY_PRICE': round(entry_price, 2) if not pd.isna(entry_price) else None,
            'EXIT_PRICE': round(exit_price, 2) if not pd.isna(exit_price) else None,
            'RETURN_PCT': round(pnl_pct, 2) if not pd.isna(pnl_pct) else None,
            'SPOT_DAY_CHANGE_PCT': round(spot_day_chg, 2) if not pd.isna(spot_day_chg) else None,
            'FUT_DAY_CHANGE_PCT': round(fut_day_chg, 2) if not pd.isna(fut_day_chg) else None,
        })

trades_df = pd.DataFrame(all_trades)
print(f"Total trade records: {len(trades_df)}")

# ── Window-Level Summary ──────────────────────────────────────────────────────
print("Computing window summaries...")

def window_summary(df):
    valid = df.dropna(subset=['RETURN_PCT'])
    if len(valid) == 0:
        return pd.Series({'count': 0, 'win_rate': np.nan, 'avg_return': np.nan,
                          'median_return': np.nan, 'max_return': np.nan,
                          'min_return': np.nan, 'std': np.nan, 'score': np.nan})
    wins = (valid['RETURN_PCT'] > 0).sum()
    wr = round(wins / len(valid) * 100, 1)
    avg = round(valid['RETURN_PCT'].mean(), 2)
    med = round(valid['RETURN_PCT'].median(), 2)
    mx = round(valid['RETURN_PCT'].max(), 2)
    mn = round(valid['RETURN_PCT'].min(), 2)
    std = round(valid['RETURN_PCT'].std(), 2)
    score = round(wr * avg / 10, 2)  # composite score
    return pd.Series({'count': len(valid), 'win_rate': wr, 'avg_return': avg,
                      'median_return': med, 'max_return': mx, 'min_return': mn,
                      'std': std, 'score': score})

# Overall summary
overall_summary = trades_df.groupby('WINDOW').apply(window_summary).reset_index()
overall_summary['CATEGORY'] = 'ALL'

# By category
cat_summary = trades_df.groupby(['CATEGORY', 'WINDOW']).apply(window_summary).reset_index()

# Combined
full_summary = pd.concat([overall_summary, cat_summary], ignore_index=True)

# ── Find Optimal Windows ──────────────────────────────────────────────────────
print("Finding optimal windows per category...")
optimal = {}
for cat in ['ALL', 'CUT', 'HIKE', 'STATUS QUO']:
    sub = full_summary[full_summary['CATEGORY'] == cat].copy()
    sub = sub.dropna(subset=['win_rate', 'avg_return'])
    sub = sub[sub['count'] >= 5]
    if len(sub) == 0:
        optimal[cat] = None
        continue
    # Best by win rate (min 55%), then avg return
    high_wr = sub[sub['win_rate'] >= 55].sort_values(['avg_return'], ascending=False)
    if len(high_wr) > 0:
        best = high_wr.iloc[0]
    else:
        best = sub.sort_values(['win_rate', 'avg_return'], ascending=False).iloc[0]
    optimal[cat] = {
        'window': best['WINDOW'],
        'win_rate': best['win_rate'],
        'avg_return': best['avg_return'],
        'count': int(best['count']),
        'score': best['score']
    }
    print(f"  {cat}: Best={best['WINDOW']} | WR={best['win_rate']}% | Avg={best['avg_return']}%")

# ── Build JSON Output ─────────────────────────────────────────────────────────
print("Building JSON...")

# Event list (T=0 day data)
events = rbi_reaction.to_dict('records')
for e in events:
    for k, v in e.items():
        if pd.isna(v) if isinstance(v, float) else False:
            e[k] = None
        elif isinstance(v, (pd.Timestamp,)):
            e[k] = str(v.date())
        elif hasattr(v, 'item'):
            e[k] = v.item()

# Window summary for JSON
def df_to_json_records(df):
    recs = []
    for _, row in df.iterrows():
        r = {}
        for k, v in row.items():
            if isinstance(v, float) and np.isnan(v):
                r[k] = None
            elif hasattr(v, 'item'):
                r[k] = v.item()
            else:
                r[k] = v
        recs.append(r)
    return recs

# Overall stats by category
category_stats = {}
for cat in ['ALL', 'CUT', 'HIKE', 'STATUS QUO']:
    if cat == 'ALL':
        sub = rbi_reaction
    else:
        sub = rbi_reaction[rbi_reaction['CATEGORY'] == cat]
    sub_valid = sub.dropna(subset=['SPOT_DAY_CHANGE_PCT'])
    wr = round((sub_valid['SPOT_DAY_CHANGE_PCT'] > 0).sum() / len(sub_valid) * 100, 1) if len(sub_valid) > 0 else 0
    category_stats[cat] = {
        'count': int(len(sub)),
        'win_rate_day': wr,
        'avg_day_change': round(float(sub['SPOT_DAY_CHANGE_PCT'].mean()), 2) if len(sub) > 0 else 0,
        'avg_fut_change': round(float(sub['FUT_DAY_CHANGE_PCT'].mean()), 2) if len(sub) > 0 else 0,
        'avg_intraday_range': round(float(sub['SPOT_RANGE_PCT'].mean()), 2) if len(sub) > 0 else 0,
        'optimal_window': optimal.get(cat),
    }

# Window summary records
window_recs = df_to_json_records(full_summary)

# Upcoming events (future dates)
today = pd.Timestamp.today()
upcoming = rbi_dates[pd.to_datetime(rbi_dates['Date']) > today].copy()
upcoming_list = []
for _, r in upcoming.iterrows():
    upcoming_list.append({
        'date': str(r['Date'].date()),
        'year': int(r['Year']),
        'meeting_type': str(r['Meeting_Type']),
        'repo_rate': float(r['Repo_Rate']),
        'action': str(r['Action']),
    })

# Recent events (last 10)
recent = rbi_reaction.sort_values('POLICY_DATE', ascending=False).head(12)
recent_list = []
for _, r in recent.iterrows():
    recent_list.append({
        'policy_date': str(r['POLICY_DATE'])[:10],
        'year': int(r['YEAR']),
        'action': str(r['ACTION']),
        'category': str(r['CATEGORY']),
        'repo_rate': float(r['REPO_RATE']),
        'spot_day_change': round(float(r['SPOT_DAY_CHANGE_PCT']), 2) if not pd.isna(r['SPOT_DAY_CHANGE_PCT']) else None,
        'fut_day_change': round(float(r['FUT_DAY_CHANGE_PCT']), 2) if not pd.isna(r['FUT_DAY_CHANGE_PCT']) else None,
        'spot_intraday_pct': round(float(r['SPOT_INTRADAY_PCT']), 2) if not pd.isna(r['SPOT_INTRADAY_PCT']) else None,
        'spot_range_pct': round(float(r['SPOT_RANGE_PCT']), 2) if not pd.isna(r['SPOT_RANGE_PCT']) else None,
        'fut_oi': int(r['FUT_OPEN_INTEREST']) if not pd.isna(r['FUT_OPEN_INTEREST']) else None,
        'fut_chg_oi': int(r['FUT_CHANGE_IN_OI']) if not pd.isna(r['FUT_CHANGE_IN_OI']) else None,
        'is_weekend': bool(r['IS_WEEKEND_ANNOUNCEMENT']),
    })

# Full events list
all_events_list = []
for _, r in rbi_reaction.sort_values('POLICY_DATE', ascending=False).iterrows():
    all_events_list.append({
        'policy_date': str(r['POLICY_DATE'])[:10],
        'eff_date': str(r['EFFECTIVE_TRADING_DATE'])[:10],
        'year': int(r['YEAR']),
        'meeting_type': str(r['MEETING_TYPE']),
        'action': str(r['ACTION']),
        'category': str(r['CATEGORY']),
        'repo_rate': float(r['REPO_RATE']),
        'spot_day_change': round(float(r['SPOT_DAY_CHANGE_PCT']), 2) if not pd.isna(r['SPOT_DAY_CHANGE_PCT']) else None,
        'fut_day_change': round(float(r['FUT_DAY_CHANGE_PCT']), 2) if not pd.isna(r['FUT_DAY_CHANGE_PCT']) else None,
        'spot_range_pct': round(float(r['SPOT_RANGE_PCT']), 2) if not pd.isna(r['SPOT_RANGE_PCT']) else None,
        'fut_oi': float(r['FUT_OPEN_INTEREST']) if not pd.isna(r['FUT_OPEN_INTEREST']) else None,
        'fut_chg_oi': float(r['FUT_CHANGE_IN_OI']) if not pd.isna(r['FUT_CHANGE_IN_OI']) else None,
        'is_weekend': bool(r['IS_WEEKEND_ANNOUNCEMENT']),
    })

json_out = {
    'generated_at': datetime.now().strftime('%Y-%m-%dT%H:%M:%S'),
    'total_events': int(len(rbi_reaction)),
    'date_range': '2000–2026',
    'category_stats': category_stats,
    'optimal_windows': optimal,
    'window_summary': window_recs,
    'upcoming_events': upcoming_list,
    'recent_events': recent_list,
    'all_events': all_events_list,
}

json_path = os.path.join(OUT_DIR, "dashboard_data", "rbi_policy_behaviour.json")
os.makedirs(os.path.dirname(json_path), exist_ok=True)
with open(json_path, 'w') as f:
    json.dump(json_out, f, indent=2, default=str)
print(f"JSON saved: {json_path} ({os.path.getsize(json_path)//1024} KB)")

# Also save to docs
docs_json_path = os.path.join(DOCS_DATA, "rbi_policy_behaviour.json")
os.makedirs(DOCS_DATA, exist_ok=True)
with open(docs_json_path, 'w') as f:
    json.dump(json_out, f, indent=2, default=str)
print(f"JSON (docs) saved: {docs_json_path}")

# ── Build Excel Report ─────────────────────────────────────────────────────────
print("Building Excel report...")
excel_path = os.path.join(OUT_DIR, "RBI_Policy_Behaviour_Master_2000_2026.xlsx")

with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
    # Sheet 1: Summary by Category × Window
    pivot = full_summary.pivot_table(
        index='WINDOW', columns='CATEGORY',
        values=['win_rate', 'avg_return', 'count'],
        aggfunc='first'
    )
    pivot.to_excel(writer, sheet_name='Window_Summary')

    # Sheet 2: Optimal Windows
    opt_df = pd.DataFrame([
        {'Category': k, **v} for k, v in optimal.items() if v
    ])
    opt_df.to_excel(writer, sheet_name='Optimal_Windows', index=False)

    # Sheet 3: Full Trade Log (T-3 to T+5 all events)
    trades_df.to_excel(writer, sheet_name='Full_Trade_Log', index=False)

    # Sheet 4: Event-Level Summary (T=0 day data)
    rbi_reaction_out = rbi_reaction.copy()
    rbi_reaction_out.to_excel(writer, sheet_name='Event_Day_Data', index=False)

    # Sheet 5: Year-wise
    yr_summary = rbi_reaction.groupby('YEAR').agg(
        Events=('POLICY_DATE', 'count'),
        Avg_Spot_Change=('SPOT_DAY_CHANGE_PCT', 'mean'),
        Avg_Range=('SPOT_RANGE_PCT', 'mean'),
        Avg_OI_Change=('FUT_CHANGE_IN_OI', 'mean')
    ).round(2).reset_index()
    yr_summary.to_excel(writer, sheet_name='Year_Wise', index=False)

    # Sheet 6: Category Stats
    cat_df = pd.DataFrame(category_stats).T.reset_index()
    cat_df.columns = ['Category'] + list(cat_df.columns[1:])
    cat_df.to_excel(writer, sheet_name='Category_Stats', index=False)

print(f"Excel saved: {excel_path}")
print("\n✅ ALL OUTPUTS GENERATED SUCCESSFULLY!")
print(f"  JSON: {json_path}")
print(f"  Excel: {excel_path}")
print(f"  Events: {len(rbi_reaction)} | Windows per event: {len(WINDOWS)} | Total trades: {len(trades_df)}")
