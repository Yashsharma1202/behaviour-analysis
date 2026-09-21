import pandas as pd, numpy as np, pathlib, sys

sys.stdout.reconfigure(errors='replace')

BASE_DIR = pathlib.Path('D:/behaviour analysis')
FUT_V13 = BASE_DIR / 'Nifty50_12_Quarters_Futures_OI_Master_v13.xlsx'
EQ_V2   = BASE_DIR / 'Nifty50_12_Quarters_Equity_Master_v2.xlsx'

xl_fut = pd.ExcelFile(FUT_V13, engine='openpyxl')
df_fut_exec = xl_fut.parse('Exec_12Q_Combined_Summary')

xl_eq = pd.ExcelFile(EQ_V2, engine='openpyxl')
df_eq_exec = xl_eq.parse('Exec_12Q_Combined_Summary')

print("==========================================================================================")
print("AUDITING EXACT TOTALS VS AVERAGES FOR 12 QUARTERS")
print("==========================================================================================")

# Read 12 rows (rows index 5 to 16 in dataframe)
# Column 4: Fund Utilised, Column 5: Net PnL, Column 6: Booked Return %, Column 8: Max DD val, Column 9: Max DD %

eq_rows = []
for r in range(6, 18):
    qtr = df_eq_exec.iloc[r, 0]
    trades = float(df_eq_exec.iloc[r, 1])
    cap = float(df_eq_exec.iloc[r, 4])
    pnl = float(df_eq_exec.iloc[r, 5])
    ret = float(df_eq_exec.iloc[r, 6])
    dd_v = float(df_eq_exec.iloc[r, 8])
    dd_p = float(df_eq_exec.iloc[r, 9])
    wr = float(df_eq_exec.iloc[r, 2]) if df_eq_exec.iloc[r, 2] is not None else 0.0
    eq_rows.append({
        'Quarter': qtr, 'Trades': trades, 'Capital': cap, 'PnL': pnl, 'Return': ret,
        'MaxDD_Val': dd_v, 'MaxDD_Pct': dd_p, 'WinRate': wr
    })

df_eq = pd.DataFrame(eq_rows)

fut_rows = []
for r in range(6, 18):
    qtr = df_fut_exec.iloc[r, 0]
    trades = float(df_fut_exec.iloc[r, 1])
    cap = float(df_fut_exec.iloc[r, 4])
    pnl = float(df_fut_exec.iloc[r, 5])
    ret = float(df_fut_exec.iloc[r, 6])
    dd_v = float(df_fut_exec.iloc[r, 8])
    dd_p = float(df_fut_exec.iloc[r, 9])
    wr = float(df_fut_exec.iloc[r, 2]) if df_fut_exec.iloc[r, 2] is not None else 0.0
    fut_rows.append({
        'Quarter': qtr, 'Trades': trades, 'Margin': cap, 'PnL': pnl, 'Return': ret,
        'MaxDD_Val': dd_v, 'MaxDD_Pct': dd_p, 'WinRate': wr
    })

df_fut = pd.DataFrame(fut_rows)

print("\n--- SPOT EQUITY (1 SHARE) ---")
print("12-Quarter Total Trades       :", int(df_eq['Trades'].sum()))
print("Average Trades Per Quarter    :", df_eq['Trades'].mean())
print("Average Capital Pool Utilised : ₹", f"{df_eq['Capital'].mean():,.2f}")
print("TOTAL 12-Quarter Net Equity P&L: ₹", f"{df_eq['PnL'].sum():,.2f}")
print("AVERAGE Quarterly Equity P&L  : ₹", f"{df_eq['PnL'].mean():,.2f}")
print("AVERAGE Booked Return %       :", f"{df_eq['Return'].mean():.2%}")
print("AVERAGE Max Drawdown ₹        : ₹", f"{df_eq['MaxDD_Val'].mean():,.2f}")
print("PEAK Max Drawdown ₹           : ₹", f"{df_eq['MaxDD_Val'].min():,.2f}")
print("AVERAGE Max Drawdown %        :", f"{df_eq['MaxDD_Pct'].mean():.2%}")
print("PEAK Max Drawdown %           :", f"{df_eq['MaxDD_Pct'].min():.2%}")
print("Global Win Rate               :", f"{df_eq['WinRate'].mean():.2%}")

print("\n--- FUTURES (1 LOT) ---")
print("12-Quarter Total Trades       :", int(df_fut['Trades'].sum()))
print("Average Trades Per Quarter    :", df_fut['Trades'].mean())
print("Average Margin Deployed       : ₹", f"{df_fut['Margin'].mean():,.2f}")
print("TOTAL 12-Quarter Net Futures P&L: ₹", f"{df_fut['PnL'].sum():,.2f}")
print("AVERAGE Quarterly Futures P&L  : ₹", f"{df_fut['PnL'].mean():,.2f}")
print("AVERAGE Booked Return %       :", f"{df_fut['Return'].mean():.2%}")
print("AVERAGE Max Drawdown ₹        : ₹", f"{df_fut['MaxDD_Val'].mean():,.2f}")
print("PEAK Max Drawdown ₹           : ₹", f"{df_fut['MaxDD_Val'].min():,.2f}")
print("AVERAGE Max Drawdown %        :", f"{df_fut['MaxDD_Pct'].mean():.2%}")
print("PEAK Max Drawdown %           :", f"{df_fut['MaxDD_Pct'].min():.2%}")
print("Global Win Rate               :", f"{df_fut['WinRate'].mean():.2%}")

print("==========================================================================================")
