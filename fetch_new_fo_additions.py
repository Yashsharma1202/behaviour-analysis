import pathlib, yfinance as yf, pandas as pd
OUT = pathlib.Path('d:/behaviour analysis/Stock_Data_213/new_fo_additions_20Nov2026')
for s in "COROMANDEL CUB DATAPATTNS ICICILIFE KARURVYSYA NAVINFLUOR SYRMA".split():
    df = yf.Ticker(f'{s}.NS').history(period='max', auto_adjust=False).reset_index()
    if df.empty: print(s,'EMPTY'); continue
    df['Date'] = pd.to_datetime(df['Date']).dt.tz_localize(None).dt.strftime('%Y-%m-%d')
    df.insert(0,'YahooTicker',f'{s}.NS'); df.insert(0,'Symbol',s)
    df.to_csv(OUT/f'{s}.csv', index=False)
    print(s, len(df), df.Date.iloc[0], df.Date.iloc[-1])
