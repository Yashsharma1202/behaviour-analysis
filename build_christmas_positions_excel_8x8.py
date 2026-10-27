"""Christmas_Positions_8x8_Nifty_FO.xlsx - position sheets for Nifty 50, F&O non-Nifty and all 213, from the validated 8x8 JSON. No trade/no-trade decisions. Additive."""
import json, sys
import pandas as pd, numpy as np
sys.stdout.reconfigure(encoding='utf-8')
ROOT = 'D:/behaviour analysis/'
d = json.load(open(ROOT + 'Christmas_FO_All_Stocks_Data_8x8.json', encoding='utf-8')); st = d['stocks']
rows = []; trades = []
for s, v in st.items():
    sp = v['selected_position']; bl = v['best_long']; bs = v['best_short']; lp = v['low_sample_indication']
    g = 'NIFTY 50' if v['is_nifty50'] else 'F&O (non-Nifty)'
    base = dict(Group=g, Symbol=s, Name=v['name'], Sector=v['sector'], Lot_Size=v['lot_size'], Last_Close=v['last_close'], Contract_Value_1Lot=v['contract_value_1lot'], Margin_20pct_1Lot=v['margin_20pct_1lot'],
                Data_From=v['first_data_date'], Max_Valid_Years=v['max_valid_years_any_window'],
                Best_LONG_Window=bl and bl['window'], Best_LONG_WinRate=bl and bl['win_rate_pct'], Best_SHORT_Window=bs and bs['window'], Best_SHORT_WinRate=bs and bs['win_rate_pct'])
    if sp:
        wf = sp['walk_forward']; d26 = sp['dates_2026']
        rows.append(dict(base, Side=sp['side'], Window=sp['window'], Lead_N=sp['lead'], Hold_M=sp['hold'], Entry_2026=d26['entry_date'], Exit_2026=d26['exit_date'],
                         Entry_Already_Passed='YES' if d26['entry_date_already_passed'] else 'NO', Valid_Years=sp['valid_years'], Win_Rate_pct=sp['win_rate_pct'], Avg_Signed_Return_pct=sp['avg_signed_return_pct'],
                         Median_pct=sp['median_signed_return_pct'], Best_Year_pct=sp['best_year_pct'], Worst_Year_pct=sp['worst_year_pct'], Last10_WinRate_pct=sp['last_10_win_rate_pct'],
                         WF_Trades=wf['trades'], WF_WinRate_pct=wf['win_rate_pct'], WF_Avg_Return_pct=wf['avg_return_pct'], Futures_Action=sp['futures_action'], Option_Play=sp['option_play'],
                         Note='TIE: best LONG and best SHORT win rates are equal (LONG window shown)' if sp['side'] == 'TIE' else ''))
        for t in sp['yearly_trades']:
            trades.append(dict(Group=g, Symbol=s, Side=sp['side'], Window=sp['window'], Year=t['year'], Entry_Date=t['entry_date'], Exit_Date=t['exit_date'], Entry_Raw_Close=t['entry_raw_close'],
                               Exit_Raw_Close=t['exit_raw_close'], Signed_Return_pct=t['signed_return_pct'], Result=t['result']))
    else:
        rows.append(dict(base, Side='NO RULE-BASED PICK', Valid_Years=v['max_valid_years_any_window'], Note=v['selection_note'], Low_Sample_Side=lp and lp['side'], Low_Sample_Window=lp and lp['window'],
                         Low_Sample_Years=lp and lp['valid_years'], Low_Sample_WinRate_pct=lp and lp['win_rate_pct'], Low_Sample_Avg_pct=lp and lp['avg_signed_return_pct']))
df = pd.DataFrame(rows); df['_o'] = df.Side.eq('NO RULE-BASED PICK').astype(int)
df = df.sort_values(['_o', 'Win_Rate_pct', 'Avg_Signed_Return_pct'], ascending=[True, False, False]).drop(columns='_o')
tr = pd.DataFrame(trades)
COLS = ['Rank', 'Symbol', 'Name', 'Sector', 'Side', 'Window', 'Lead_N', 'Hold_M', 'Entry_2026', 'Exit_2026', 'Entry_Already_Passed', 'Valid_Years', 'Win_Rate_pct', 'Avg_Signed_Return_pct', 'Median_pct',
        'Best_Year_pct', 'Worst_Year_pct', 'Last10_WinRate_pct', 'WF_Trades', 'WF_WinRate_pct', 'WF_Avg_Return_pct', 'Best_LONG_Window', 'Best_LONG_WinRate', 'Best_SHORT_Window', 'Best_SHORT_WinRate',
        'Lot_Size', 'Last_Close', 'Contract_Value_1Lot', 'Margin_20pct_1Lot', 'Futures_Action', 'Option_Play', 'Data_From', 'Low_Sample_Side', 'Low_Sample_Window', 'Low_Sample_Years',
        'Low_Sample_WinRate_pct', 'Low_Sample_Avg_pct', 'Note']
def sheet(sub):
    sub = sub.copy().reset_index(drop=True)
    sub.insert(0, 'Rank', [None if sd == 'NO RULE-BASED PICK' else i + 1 for i, sd in enumerate(sub.Side)])
    for c in COLS:
        if c not in sub: sub[c] = None
    return sub[COLS]
nif, fo, al = sheet(df[df.Group == 'NIFTY 50']), sheet(df[df.Group != 'NIFTY 50']), sheet(df)
OUT = ROOT + 'Christmas_Positions_8x8_Nifty_FO.xlsx'
with pd.ExcelWriter(OUT, engine='xlsxwriter') as w:
    wb = w.book
    hdr = wb.add_format({'bold': True, 'font_color': '#FFFFFF', 'bg_color': '#1F3864', 'border': 1, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center'})
    grn = wb.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'}); red = wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
    amb = wb.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C5700'}); gry = wb.add_format({'font_color': '#7F7F7F', 'italic': True})
    summ = []
    for nm, sub in (('Nifty 50', nif), ('F&O non-Nifty', fo), ('All 213', al)):
        p = sub[sub.Side != 'NO RULE-BASED PICK']
        summ.append(dict(Group=nm, Stocks=len(sub), With_Rule_Based_Position=len(p), LONG=int((p.Side == 'LONG').sum()), SHORT=int((p.Side == 'SHORT').sum()), TIE=int((p.Side == 'TIE').sum()),
                         No_Pick_Under_7_Years=int((sub.Side == 'NO RULE-BASED PICK').sum()), Avg_Win_Rate_pct=round(float(p.Win_Rate_pct.mean()), 1),
                         Avg_WalkForward_WinRate_pct=round(float(p.WF_WinRate_pct.astype(float).mean()), 1), Entry_Already_Passed=int((p.Entry_Already_Passed == 'YES').sum())))
    sm = pd.DataFrame(summ); sm.to_excel(w, sheet_name='Summary', index=False)
    for nm, sub in (('Nifty50_Positions', nif), ('FO_NonNifty_Positions', fo), ('All_213_Positions', al)):
        sub.to_excel(w, sheet_name=nm, index=False); ws = w.sheets[nm]
        for i, c in enumerate(sub.columns):
            ws.write(0, i, c.replace('_', ' '), hdr); ws.set_column(i, i, 28 if c in ('Name', 'Note', 'Sector') else 14)
        ws.set_row(0, 34); ws.freeze_panes(1, 3); ws.autofilter(0, 0, len(sub), len(sub.columns) - 1); L = len(sub); cols = list(sub.columns)
        ci = cols.index('Side')
        for val, f in (('LONG', grn), ('SHORT', red), ('TIE', amb), ('NO RULE-BASED PICK', gry)):
            ws.conditional_format(1, ci, L, ci, {'type': 'cell', 'criteria': '==', 'value': f'"{val}"', 'format': f})
        for key, lo, mid, hi in (('Win_Rate_pct', 50, 70, 90), ('WF_WinRate_pct', 30, 55, 80)):
            k = cols.index(key)
            ws.conditional_format(1, k, L, k, {'type': '3_color_scale', 'min_type': 'num', 'min_value': lo, 'mid_type': 'num', 'mid_value': mid, 'max_type': 'num', 'max_value': hi,
                                               'min_color': '#F8696B', 'mid_color': '#FFFFFF', 'max_color': '#63BE7B'})
    tr.to_excel(w, sheet_name='Yearly_Trades_Selected', index=False); w.sheets['Yearly_Trades_Selected'].set_column(0, 10, 16); w.sheets['Yearly_Trades_Selected'].freeze_panes(1, 2)
    pd.DataFrame({'Note': ['PROJECT 8x8 GRID: entry T-N and exit T+M with N,M = 1..8; LONG and SHORT; Christmas (25-Dec) 2000-2025; minimum 7 valid years.',
                           'Side = the side with the higher best-window win rate; TIE = equal (LONG window shown). No trade / no-trade decision is made in this workbook.',
                           'Entry at close of T-N, exit at close of T+M. 2026: T = Friday 25-Dec-2026 (NSE closed). Entry_Already_Passed is as of 27-Oct-2026.',
                           'In-sample win rates are optimistic (128 windows scanned per stock); WF = walk-forward (pick redone each year from earlier years only) is the out-of-sample view.',
                           'Sector is only available for the 50 Nifty stocks. ANANDRATHI, ENRIN and UJJIVANSFB have no lot size in the source.',
                           'Source: Christmas_FO_All_Stocks_Data_8x8.json (validated, 0 errors). Additive: new file only.']}).to_excel(w, sheet_name='Notes', index=False)
    w.sheets['Summary'].set_column(0, 10, 24); w.sheets['Notes'].set_column(0, 0, 160)
x = pd.ExcelFile(OUT); print({n: x.parse(n).shape for n in x.sheet_names}); print(sm.T.to_string())
print(nif[nif.Side != 'NO RULE-BASED PICK'][['Rank', 'Symbol', 'Side', 'Window', 'Entry_2026', 'Exit_2026', 'Win_Rate_pct', 'WF_WinRate_pct']].head(12).to_string(index=False))
