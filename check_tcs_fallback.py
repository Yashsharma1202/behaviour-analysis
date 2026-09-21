import pandas as pd
import sys
import pathlib
import datetime as dt
import calendar

sys.stdout.reconfigure(encoding='utf-8')

base = pathlib.Path('E:/OI_DATA')

# ----- Test TCS Q4 2025-26 -----
sym = 'TCS'
en_date = '2026-01-01'   # Entry date from trade log
ex_date = '2026-01-14'   # Exit date from trade log

opt_dir = base / sym / 'options_parquet'
spot_dir = base / sym / 'spot_parquet'

print(f'=== Diagnosing {sym} | Entry={en_date} | Exit={ex_date} ===')
print()

# Check what parquet files exist around entry date
all_opt = sorted(opt_dir.glob('*.parquet'))
dates_around = [f.stem for f in all_opt if '2025-12' <= f.stem <= '2026-01-20']
print(f'Parquet files around entry date: {dates_around[:10]}')

# Load entry date parquet
path = opt_dir / f'{en_date}.parquet'
print(f'\nEntry date parquet exists: {path.exists()}')

if path.exists():
    df = pd.read_parquet(path)
    print(f'Shape: {df.shape}')
    expiries = sorted(df['ExpiryDate'].unique())
    print(f'Available expiries: {expiries}')

    # Spot price
    sp = base / sym / 'spot_parquet' / f'{en_date}.parquet'
    if sp.exists():
        dfs = pd.read_parquet(sp)
        close_s = dfs[dfs['Time'] == '15:29:59']
        if close_s.empty:
            close_s = dfs.sort_values('Time').tail(1)
        actual_spot = float(close_s.iloc[0]['Close'])
        print(f'Actual spot price on {en_date}: {actual_spot:.2f}')

        # Scale factor
        # Assume sheet price might be ~3227 (from fallback data shown)
        sheet_price = 3227.4
        ratio = actual_spot / sheet_price
        print(f'Sheet price (proxy): {sheet_price}')
        print(f'Scale ratio (actual/sheet): {ratio:.4f}')

        # ATM strike in actual terms
        standard_ratios = [1, 2, 5, 10, 20]
        sf = min(standard_ratios, key=lambda x: abs(x - ratio))
        print(f'Nearest standard scale factor: {sf}')

        actual_atm = round(actual_spot / 50) * 50
        print(f'ATM Strike (actual): {actual_atm}')

        # Check if this strike exists in the parquet
        near_strikes = sorted(df[df['ExpiryDate'] == expiries[0]]['Strike'].unique())
        print(f'\nAvailable strikes for expiry {expiries[0]}:')
        print(near_strikes[:20])

        # Try to find CE and PE
        for exp in expiries:
            exp_dt = dt.date.fromisoformat(exp)
            ex_dt = dt.date.fromisoformat(ex_date)
            if exp_dt >= ex_dt:
                print(f'\nUsing expiry: {exp} (exit date: {ex_date})')
                ce_rows = df[(df['Strike'] == float(actual_atm)) & (df['Type'] == 'CE') & (df['ExpiryDate'] == exp)]
                pe_rows = df[(df['Strike'] == float(actual_atm)) & (df['Type'] == 'PE') & (df['ExpiryDate'] == exp)]
                print(f'CE rows for strike {actual_atm}: {len(ce_rows)}')
                print(f'PE rows for strike {actual_atm}: {len(pe_rows)}')
                if not ce_rows.empty:
                    close_ce = ce_rows[ce_rows['Time'] == '15:29:59']
                    print(f'CE close: {float(close_ce.iloc[0]["Close"]):.2f}' if not close_ce.empty else 'CE close: no 15:29:59 row')
                if not pe_rows.empty:
                    close_pe = pe_rows[pe_rows['Time'] == '15:29:59']
                    print(f'PE close: {float(close_pe.iloc[0]["Close"]):.2f}' if not close_pe.empty else 'PE close: no 15:29:59 row')
                break

# ----- Also check HEROMOTOCO -----
print()
print('=== HEROMOTOCO parquet check ===')
hero_dir = base / 'HEROMOTOCO'
print(f'Folder exists: {hero_dir.exists()}')
print(f'options_parquet exists: {(hero_dir / "options_parquet").exists()}')

# ----- Also check INDUSINDBK -----
print()
print('=== INDUSINDBK parquet check ===')
indus_dir = base / 'INDUSINDBK'
print(f'Folder exists: {indus_dir.exists()}')
print(f'options_parquet exists: {(indus_dir / "options_parquet").exists()}')

# Check if HEROMOTOCO / INDUSINDBK data is stored under a different name
print()
print('All folders in E:/OI_DATA:')
all_dirs = sorted([d.name for d in base.iterdir() if d.is_dir()])
print(all_dirs)
