import sys; sys.stdout.reconfigure(encoding='utf-8')
src=open('build_diwali_christmas_perstock_reports_v4.py',encoding='utf-8').read()
head=src.split('# ---------------------------------------------------------------- per-stock analysis')[0].replace('json.dump(EVENT_WINDOW','(lambda *a,**k:None)(EVENT_WINDOW').replace(r'open(ROOT + r"\Diwali_Christmas_event_windows_v4.json", "w")', 'open(__import__("os").devnull, "w")')
ns={}; exec(compile(head,'v4','exec'),ns)
np=ns['np']; pd=ns['pd']; data=ns['data']['Diwali']; N=ns['NIFTY50']; M=ns['MIN_YEARS']; meta=ns['meta']; Y=ns['YEARS']
rows=[]
for s in sorted(data):
    C=data[s]['R'][:, :8, :8]; v=~np.isnan(C); n=v.sum(0)
    with np.errstate(all='ignore'):
        wl=((C>0)&v).sum(0)/n*100; wsh=((C<0)&v).sum(0)/n*100; avg=np.nanmean(C,0)
    ok=n>=M; g='NIFTY 50' if s in N else 'F&O'
    base=dict(Group=g,Symbol=s,Name=meta.get(s,{}).get('name',''),Max_Valid_Years=int(n.max()))
    if not ok.any(): rows.append({**base,'Verdict':'DATA INSUFFICIENT (<7 yrs)'}); continue
    sl=np.where(ok,wl+1e-6*avg,-1); ss=np.where(ok,wsh-1e-6*avg,-1)
    a=np.unravel_index(np.argmax(sl),sl.shape); b=np.unravel_index(np.argmax(ss),ss.shape)
    L=(wl[a],avg[a],n[a]); S=(wsh[b],-avg[b],n[b]); d=L[0]-S[0]
    verdict='LONG' if d>0.05 else ('SHORT' if d<-0.05 else 'TIE')
    base.update(Best_LONG_Window=f"T-{a[0]+1} to T+{a[1]+1}",LONG_WinRate=round(L[0],1),LONG_AvgRet=round(L[1],2),LONG_Years=int(L[2]),
      Best_SHORT_Window=f"T-{b[0]+1} to T+{b[1]+1}",SHORT_WinRate=round(S[0],1),SHORT_AvgRet=round(S[1],2),SHORT_Years=int(S[2]),
      WinRate_Gap_LONG_minus_SHORT=round(d,1),Verdict=verdict)
    rows.append(base)
df=pd.DataFrame(rows)
yr=[]
for g,sel in (('NIFTY 50',lambda s:s in N),('F&O (non-Nifty)',lambda s:s not in N),('ALL 213',lambda s:True)):
    st=np.stack([data[s]['R'] for s in data if sel(s)]); B=np.nanmean(st,0)
    for i,y in enumerate(Y):
        x=B[i,2,2]
        if not np.isnan(x): yr.append(dict(Group=g,Year=y,Basket_Ret_T3_to_T3_pct=round(float(x),2),Result='UP' if x>0 else 'DOWN'))
yd=pd.DataFrame(yr)
summ=[]
for g,d_ in df.groupby('Group'):
    t=d_[d_.Verdict.isin(['LONG','SHORT','TIE'])]
    summ.append(dict(Group=g,Stocks=len(d_),Tested=len(t),LONG_better=int((t.Verdict=='LONG').sum()),SHORT_better=int((t.Verdict=='SHORT').sum()),Tie=int((t.Verdict=='TIE').sum()),
      Data_insufficient=int(d_.Verdict.str.startswith('DATA').sum()),Avg_best_LONG_WR=round(t.LONG_WinRate.mean(),1),Avg_best_SHORT_WR=round(t.SHORT_WinRate.mean(),1),
      Basket_UP_years=int(((yd.Group.str.startswith(g[:3]))&(yd.Result=='UP')).sum()),Basket_years=int(yd.Group.str.startswith(g[:3]).sum())))
sd=pd.DataFrame(summ)
out='Diwali_Long_vs_Short_Check_Nifty_FO_8x8.xlsx'
with pd.ExcelWriter(out,engine='xlsxwriter') as w:
    sd.to_excel(w,sheet_name='Summary',index=False)
    for nm,sub in (('Nifty50',df[df.Group=='NIFTY 50']),('FO_NonNifty',df[df.Group=='F&O']),('All_213',df)):
        sub.sort_values('WinRate_Gap_LONG_minus_SHORT',ascending=False).to_excel(w,sheet_name=nm,index=False)
    yd.to_excel(w,sheet_name='Basket_By_Year',index=False)
    pd.DataFrame({'Note':['Per stock: best LONG and best SHORT window over PROJECT 8x8 GRID lead 1-8 x hold 1-8, Diwali (Laxmi Pujan) 2000-2025, min 7 valid years.',
     'Verdict = LONG/SHORT if win-rate gap > 0.05 pts, else TIE (the original build silently picks LONG on a tie).',
     'In-sample win rates are optimistic (200 combos scanned); walk-forward numbers are in the Diwali_Christmas v4 workbooks.',
     'Basket_By_Year: equal-weight average return over T-3 to T+3.','Additive: new file, nothing else changed.']}).to_excel(w,sheet_name='Notes',index=False)
    for sh in w.sheets.values(): sh.set_column(0,14,16); sh.freeze_panes(1,0)
print('saved',out); print(sd.T)
