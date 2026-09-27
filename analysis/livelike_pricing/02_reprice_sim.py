import pandas as pd, numpy as np
s=pd.read_pickle('single.pkl')
r=s[(s.date>='2025-10-01')&(s.date<'2026-07-01')&(s.cr<2)].copy()
FV,FALL,K,M=0.18,0.24,300,2.5
def rnd(p): return np.ceil(p/100)*100-10
r['M']=r.sales*(1-FV)/(r.cost+K)
print(r.M.describe(percentiles=[.1,.25,.5,.75,.9]).round(2))
r['Pt']=np.maximum(rnd((r.cost+K)/(1-FV)*M),1990)
r['P1']=np.clip(r.Pt, r.sales*0.8, r.sales*1.2); r['P1']=rnd(r.P1)
# only change if deviation >10%
dev=r.Pt/r.sales-1
r.loc[dev.abs()<=0.10,'P1']=r.sales
r['chg']=r.P1/r.sales-1
cls=np.select([r.chg>0.001,r.chg<-0.001],['値上げ','値下げ'],'据え置き')
print(pd.Series(cls).value_counts(normalize=True).round(3))
print('avg up',r.chg[r.chg>0].mean().round(3),'avg down',r.chg[r.chg<0].mean().round(3))
for e in [0.8,1.2,1.6,2.0]:
    q=(r.P1/r.sales)**(-e)
    p0=(r.sales*(1-FALL)-r.cost).sum(); p1=((r.P1*(1-FALL)-r.cost)*q).sum()
    print(f'e={e}: orders {q.sum()/len(r)-1:+.1%} revenue {(r.P1*q).sum()/r.sales.sum()-1:+.1%} 入金後粗利 {p1/p0-1:+.1%}')
bins=[0,500,700,900,1200,1600,2200,3000,5000,1e9]
r['cb']=pd.cut(r.cost,bins)
t=r.groupby('cb',observed=True).agg(注文=('sales','size'),原価中央=('cost','median'),現売価中央=('sales','median'),原価率=('cr','median'),推奨売価=('Pt','median'),
   値上げ率=('chg',lambda x:(x>0).mean()),値下げ率=('chg',lambda x:(x<0).mean()))
t['入金後粗利_現']=t.現売価中央*(1-FALL)-t.原価中央
t['入金後粗利_推']=t.推奨売価*(1-FALL)-t.原価中央
t['推奨原価率']=t.原価中央/t.推奨売価
t['差']=t.推奨売価/t.現売価中央-1
print(t.round(3).to_string())
r.to_pickle('sim2.pkl'); t.to_pickle('bandtable.pkl')
