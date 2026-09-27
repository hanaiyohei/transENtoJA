import pandas as pd, numpy as np, re
df=pd.read_excel('tsugou.xlsx',sheet_name='AEorders')
df['注文日']=pd.to_datetime(df['注文日'])
ok=~df['状態'].isin(['Canceled','Expired','AEキャンセル(未出荷)'])
d=df[ok & df['受注番号'].notna()].copy()
def ssum(s):
    t=0
    for p in str(s).split('/'):
        p=p.strip().replace(',','')
        try: t+=float(p)
        except: pass
    return t
g=d.groupby('受注番号').agg(cost=('金額','sum'),n_ae=('金額','size'),sales=('売上金額','first'),mall=('モール','first'),date=('注文日','min'),store=('仕入先ストア','first'),item=('商品','first'),cp=('CPマーカー','first'),grade=('確度','first'))
g['sales']=g['sales'].map(ssum)
g['n_ord']=g.index.map(lambda s:len(str(s).split('/')))
g=g[(g.sales>0)]
g['rak']=g['mall'].map(lambda m: set(x.strip() for x in str(m).split('/'))=={'楽天'})
g['cr']=g.cost/g.sales
g['ym']=g.date.dt.to_period('M')
g.to_pickle('bundles.pkl')
print(len(g), g.rak.sum())
r=g[g.rak]
print(r.groupby('ym').agg(n=('cr','size'),sales=('sales','sum'),cost=('cost','sum'),cr_med=('cr','median')).assign(cr_w=lambda x:x.cost/x.sales).tail(24).to_string())
