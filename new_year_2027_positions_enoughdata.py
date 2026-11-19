"""1-Jan-2027 position book, TOP-50 restricted to stocks/windows with >=12 valid New Year years (enough data).
Same method as new_year_2027_positions.py, live ranking uses MINP=12. New outputs only (…_Top50_EnoughData…)."""
s = open('D:/behaviour analysis/new_year_2027_positions.py', encoding='utf-8').read()
s = s.replace('exec(src, globals())', "exec(src.replace('TOPN, MINP, Y0 = 50, 6, 2012', 'TOPN, MINP, Y0 = 50, 12, 2012'), globals())", 1)
s = s.replace('New_Year_1Jan2027_Position_Book', 'New_Year_1Jan2027_Position_Book_Top50_EnoughData')
s = s.replace('The backtest of this rule averaged +1.83% a year (2012-26, 14 of 15 years positive).', 'Live ranking additionally requires >=12 valid years per stock-window (enough data). The original walk-forward backtest (min 6 prior years) averaged +1.83% a year (2012-26, 14 of 15 years positive); it was not re-run with the 12-year filter.')
s = s.replace('backtest of this rule: +1.83%/yr, 14 of 15 years positive', 'all stocks have >=12 yrs of data; walk-forward backtest (min-6-yr rule) +1.83%/yr, 14/15 yrs +')
exec(compile(s, 'enoughdata', 'exec'), globals())
