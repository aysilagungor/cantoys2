import io
import math
import re
from datetime import datetime
from pathlib import Path
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from scipy.stats import norm

st.set_page_config(page_title='CanToys | Akıllı Stok Karar Merkezi', page_icon='🧸', layout='wide', initial_sidebar_state='expanded')
BRANCHES=['CanToys PLUS','CanToys XL','CanToys EXTRA']
COLORS=['#fa6b9b','#7c6df1','#20bfa9']
st.markdown('''<style>
.stApp{background:linear-gradient(135deg,#fff8fb 0%,#f3f7ff 55%,#f0fffa 100%);color:#26324a}
[data-testid="stMetric"]{background:white;border-radius:18px;padding:16px;border:1px solid #f3dce8;box-shadow:0 5px 16px #e3d9e333}
h1,h2,h3{color:#6844a4!important}.stButton>button{background:#f9d5e6;border-radius:16px;border:0;color:#553379;font-weight:700}
[data-testid="stSidebar"]{background:#fff1f7}
</style>''',unsafe_allow_html=True)

ALIASES={
 'date':['date','tarih','belge tarihi','islem tarihi','işlem tarihi','fatura tarihi'],
 'branch':['branch','sube','şube','magaza','mağaza','depo'],
 'sku':['sku','stok kodu','urun kodu','ürün kodu','barkod','stok_kodu'],
 'product':['product','urun','ürün','urun adi','ürün adı','stok adi','stok adı'],
 'qty':['qty','quantity','miktar','adet','satis adedi','satış adedi'],
 'unit_price':['unit_price','birim fiyat','satis fiyati','satış fiyatı','fiyat'],
 'revenue':['revenue','ciro','tutar','net tutar','toplam tutar','satis tutari','satış tutarı'],
 'vat_rate':['vat_rate','kdv orani','kdv oranı','kdv'],
 'transaction':['transaction','islem tipi','işlem tipi','hareket tipi','tip'],
 'stock':['stock','stok','stok miktari','stok miktarı','mevcut stok'],
 'unit_cost':['unit_cost','birim maliyet','alis fiyati','alış fiyatı','maliyet'],
 'lead_days':['lead_days','tedarik suresi','tedarik süresi','termin suresi','termin süresi']}
def keynorm(x):
 import unicodedata
 x=str(x).strip().lower().replace('ı','i').replace('_',' ')
 x=''.join(c for c in unicodedata.normalize('NFKD',x) if not unicodedata.combining(c))
 return re.sub(r'\s+',' ',x)
def normalize_columns(df):
 df=df.copy(); rename={}
 for canonical, names in ALIASES.items():
  opts={keynorm(s) for s in names+[canonical]}
  for col in df.columns:
   if keynorm(col) in opts and col not in rename and canonical not in rename.values():rename[col]=canonical;break
 return df.rename(columns=rename)
def number(s):
 if pd.api.types.is_numeric_dtype(s):return pd.to_numeric(s,errors='coerce')
 def conv(v):
  if pd.isna(v):return np.nan
  v=re.sub(r'[^\d,.\-]','',str(v).strip())
  if ',' in v and '.' in v:v=v.replace('.','').replace(',','.') if v.rfind(',')>v.rfind('.') else v.replace(',','')
  elif ',' in v:v=v.replace(',','.')
  try:return float(v)
  except:return np.nan
 return s.map(conv)
def load_file(f):
 data=f.getvalue()
 if f.name.lower().endswith('.csv'):
  for enc in ['utf-8-sig','cp1254','latin1']:
   try:
    txt=data.decode(enc)
    return pd.read_csv(io.StringIO(txt),sep=None,engine='python')
   except (UnicodeDecodeError,pd.errors.ParserError):pass
  raise ValueError('CSV ayrıştırılamadı.')
 return pd.read_excel(io.BytesIO(data))
def parse_cantoys_erp(file_or_path):
 """Parse CanToys hierarchical transaction export. Returns sales and movement ledger."""
 if isinstance(file_or_path,(str,Path)):
  raw=pd.read_excel(file_or_path,header=None,sheet_name=0)
 else:
  raw=pd.read_excel(io.BytesIO(file_or_path.getvalue()),header=None,sheet_name=0)
 rows=[];sku=None;product=None;header={};source=str(getattr(file_or_path,'name',Path(file_or_path).name if isinstance(file_or_path,(str,Path)) else 'ERP'))
 for _,r in raw.iterrows():
  v=r.tolist();first=v[0] if len(v)>0 else None;second=v[1] if len(v)>1 else None
  if pd.isna(first) and pd.notna(second) and pd.to_numeric(second,errors='coerce')==pd.to_numeric(second,errors='coerce') and len(v)>2 and isinstance(v[2],str) and v[2].strip():
   sku=str(int(float(second))) if float(second).is_integer() else str(second).strip();product=str(v[2]).strip();header={};continue
  if str(first).strip().upper()=='ALISSATIS_NO':
   header={str(val).strip().upper():j for j,val in enumerate(v) if pd.notna(val)};continue
  if not sku or not header or pd.isna(first) or pd.to_numeric(first,errors='coerce')!=pd.to_numeric(first,errors='coerce'):continue
  def get(k):
   j=header.get(k);return v[j] if j is not None and j<len(v) else None
  typ=str(get('ISLEM_ADI') or '').strip();kind=keynorm(typ)
  if not any(t in kind for t in ['satis','alis','iade']):continue
  dateval=get('TARIH')
  if isinstance(dateval,(int,float,np.number)) and not isinstance(dateval,bool):dt=pd.to_datetime(float(dateval),unit='D',origin='1899-12-30',errors='coerce')
  else:dt=pd.to_datetime(dateval,dayfirst=True,errors='coerce')
  qty=number(pd.Series([get('MIKTAR')])).iloc[0];price=number(pd.Series([get('UYGULAMA_FIYATI_KDV')])).iloc[0]
  if pd.isna(dt) or pd.isna(qty) or qty==0:continue
  location=str(get('STOK_YERI_ADI') or '').upper();customer=str(get('CARI_UNVANI') or '').upper()
  context=location+' '+customer
  if 'EXTRA' in context:branch='CanToys EXTRA'
  elif 'PLUS' in context:branch='CanToys PLUS'
  elif 'XL' in context:branch='CanToys XL'
  else:branch=None
  is_return='iade' in kind;is_sale='satis' in kind;is_purchase='alis' in kind
  vat=number(pd.Series([get('KDV_ORANI')])).iloc[0];vat=20.0 if pd.isna(vat) else float(vat)
  rows.append({'date':dt,'branch':branch,'sku':sku,'product':product,'qty':(-abs(qty) if is_return else abs(qty)) if is_sale else 0.,'revenue':((-abs(qty) if is_return else abs(qty))*price if pd.notna(price) else np.nan) if is_sale else 0.,'vat_rate':vat,'transaction':typ,'is_sale':is_sale,'is_purchase':is_purchase,'raw_qty':float(qty),'price_gross':price,'source':source,'transaction_id':str(first),'location':location,'customer':customer})
 ledger=pd.DataFrame(rows)
 if ledger.empty:raise ValueError(f'{source}: ERP hareketleri okunamadı.')
 sales=ledger[ledger.is_sale & ledger.branch.notna()].copy()
 if sales.empty:raise ValueError(f'{source}: Şubeye atanabilen satış kaydı bulunamadı.')
 return sales,ledger

def ledger_stock_proxy(ledger):
 """Only a movement-based lower-bound/proxy; NOT verified on-hand inventory."""
 b=ledger[ledger.branch.notna()].copy()
 b['movement']=np.where(b.is_purchase,np.where(b.transaction.map(keynorm).str.contains('iade'),-b.raw_qty,b.raw_qty),np.where(b.is_sale,-b.qty,0))
 out=b.groupby(['branch','sku'],as_index=False)['movement'].sum().rename(columns={'movement':'stock'})
 out['stock']=out['stock'].clip(lower=0)
 out['unit_cost']=np.nan;out['lead_days']=np.nan
 return out

def prepare_sales(raw):
 df=normalize_columns(raw)
 needed={'date','branch','sku','qty'}
 if not needed.issubset(df.columns):raise ValueError('Eksik zorunlu sütunlar: '+', '.join(sorted(needed-set(df.columns))))
 df=df.copy();df['date']=pd.to_datetime(df['date'],errors='coerce',dayfirst=True)
 for c in ['qty','unit_price','revenue','vat_rate','unit_cost']:
  if c in df:df[c]=number(df[c])
 df=df.dropna(subset=['date','branch','sku','qty']);df['branch']=df['branch'].astype(str).str.strip();df['sku']=df['sku'].astype(str).str.strip()
 df=df[(df['branch']!='')&(df['sku']!='')]
 df['product']=df['product'].fillna(df['sku']).astype(str) if 'product' in df else df['sku']
 if 'transaction' in df:
  returns=df['transaction'].astype(str).str.lower().str.contains('iade|return|refund',regex=True)
  df.loc[returns,'qty']=-df.loc[returns,'qty'].abs()
  if 'revenue' in df:df.loc[returns,'revenue']=-df.loc[returns,'revenue'].abs()
 if 'vat_rate' not in df:df['vat_rate']=0.
 df['vat_rate']=df['vat_rate'].fillna(0)
 df['vat_rate']=np.where(df['vat_rate']>1,df['vat_rate']/100,df['vat_rate'])
 df['vat_rate']=df['vat_rate'].clip(0,1)
 if 'revenue' not in df:
  if 'unit_price' not in df:raise ValueError('revenue (ciro) veya unit_price (birim fiyat) gerekli.')
  df['revenue']=df['qty']*df['unit_price']
 df['net_revenue']=df['revenue']/(1+df['vat_rate'])
 df['month']=df['date'].dt.to_period('M').dt.to_timestamp()
 return df

def prepare_stock(raw):
 df=normalize_columns(raw)
 if not {'branch','sku','stock'}.issubset(df):raise ValueError('Stok dosyası branch, sku, stock alanlarını içermeli.')
 for c in ['stock','unit_cost','lead_days']:
  if c in df:df[c]=number(df[c])
 df=df.dropna(subset=['branch','sku','stock']).copy()
 df['branch']=df['branch'].astype(str).str.strip();df['sku']=df['sku'].astype(str).str.strip()
 for c in ['unit_cost','lead_days']:
  if c not in df:df[c]=np.nan
 return df.groupby(['branch','sku'],as_index=False).agg(stock=('stock','sum'),unit_cost=('unit_cost','mean'),lead_days=('lead_days','mean'))

def monthly_panel(sales):
 start=sales['month'].min();end=sales['month'].max();months=pd.date_range(start,end,freq='MS')
 agg=sales.groupby(['branch','sku','month'],as_index=False).agg(qty=('qty','sum'),revenue=('net_revenue','sum'))
 keys=sales[['branch','sku']].drop_duplicates();full=keys.merge(pd.DataFrame({'month':months}),how='cross')
 return full.merge(agg,on=['branch','sku','month'],how='left').fillna({'qty':0,'revenue':0})

def classify(sales,panel,stock,asof,fast_cut=0.65):
 g=panel.groupby(['branch','sku'])
 a=g.agg(total_qty=('qty','sum'),total_revenue=('revenue','sum'),mean_monthly=('qty','mean'),std_monthly=('qty','std'),active_months=('qty',lambda x:int((x>0).sum())),months=('qty','size')).reset_index()
 a['cv']=(a['std_monthly'].fillna(0)/a['mean_monthly'].abs().replace(0,np.nan)).fillna(0)
 slopes=g.apply(lambda x:float(np.polyfit(np.arange(len(x)),x['qty'],1)[0]) if len(x)>1 else 0,include_groups=False).reset_index(name='trend_slope')
 a=a.merge(slopes,on=['branch','sku'])
 a['continuity']=a['active_months']/a['months'].clip(lower=1)
 a['fsn']=np.select([a['total_qty']<=0,a['continuity']>=fast_cut],['N','F'],default='S')
 a=a.sort_values('total_revenue',ascending=False)
 a['revenue_positive']=a['total_revenue'].clip(lower=0)
 a['cum_share']=a['revenue_positive'].cumsum()/max(a['revenue_positive'].sum(),1e-9)
 a['abc']=np.select([a['cum_share']<=.8,a['cum_share']<=.95],['A','B'],default='C')
 last=sales[sales['qty']>0].groupby(['branch','sku'])['date'].max().rename('last_sale').reset_index()
 a=a.merge(last,on=['branch','sku'],how='left')
 a['days_since_sale']=(pd.Timestamp(asof)-a['last_sale']).dt.days.fillna((pd.Timestamp(asof)-sales['date'].min()).days).clip(lower=0)
 names=sales.groupby(['branch','sku'])['product'].first().reset_index()
 a=a.merge(names,on=['branch','sku'],how='left').merge(stock,on=['branch','sku'],how='left')
 return a

def forecast_methods(y,h=1,window=3,alpha=.3):
 y=np.asarray(y,dtype=float); n=len(y)
 if n==0:return {}
 ma=float(np.mean(y[-min(window,n):]));level=float(y[0]);
 for v in y[1:]:level=alpha*v+(1-alpha)*level
 def croston(sba=False):
  nz=np.flatnonzero(y>0)
  if len(nz)==0:return 0.
  z=float(y[nz[0]]);p=float(nz[0]+1);last=int(nz[0]);
  for i in nz[1:]:
   z=alpha*y[i]+(1-alpha)*z;p=alpha*(i-last)+(1-alpha)*p;last=int(i)
  return max(0,z/max(p,1e-9)*(1-alpha/2 if sba else 1))
 return {'Hareketli Ortalama':max(0,ma),'Üstel Düzeltme':max(0,level),'Croston':croston(),'SBA':croston(True)}

def backtest(y,window=3,alpha=.3,min_train=4):
 y=np.asarray(y,float);records=[]
 if len(y)<=min_train:return pd.DataFrame()
 for i in range(min_train,len(y)):
  preds=forecast_methods(y[:i],window=window,alpha=alpha)
  for name,pred in preds.items():records.append((name,y[i],pred))
 r=pd.DataFrame(records,columns=['method','actual','pred'])
 denom=np.mean(np.abs(np.diff(y[:max(min_train,2)])))
 if denom<1e-9:denom=np.mean(np.abs(np.diff(y))) if len(y)>1 else 0
 out=r.groupby('method').apply(lambda d:pd.Series({'MAE':np.mean(np.abs(d.actual-d.pred)),'RMSE':np.sqrt(np.mean((d.actual-d.pred)**2)),'MASE':np.mean(np.abs(d.actual-d.pred))/denom if denom>1e-9 else np.nan}),include_groups=False).reset_index()
 return out

def merec_weights(matrix,benefit):
 x=np.asarray(matrix,float);m=x.shape[1]; eps=1e-9
 if len(x)<2:return np.ones(m)/m
 z=np.zeros_like(x)
 for j in range(m):
  v=x[:,j];lo,hi=np.min(v),np.max(v)
  if hi-lo<eps:z[:,j]=1
  elif benefit[j]:z[:,j]=(v-lo)/(hi-lo)+.01
  else:z[:,j]=(hi-v)/(hi-lo)+.01
 z=np.clip(z,eps,None)
 base=np.log(1+np.mean(np.abs(np.log(z)),axis=1))
 effects=[]
 for j in range(m):
  reduced=np.log(1+np.sum(np.abs(np.log(np.delete(z,j,axis=1))),axis=1)/m)
  effects.append(np.sum(np.abs(reduced-base)))
 effects=np.asarray(effects)
 return effects/effects.sum() if effects.sum()>eps else np.ones(m)/m

def marcos(matrix,weights,benefit):
 x=np.asarray(matrix,float);eps=1e-9
 ideal=np.array([x[:,j].max() if benefit[j] else x[:,j].min() for j in range(x.shape[1])])
 anti=np.array([x[:,j].min() if benefit[j] else x[:,j].max() for j in range(x.shape[1])])
 extended=np.vstack([anti,x,ideal]);normal=np.zeros_like(extended)
 for j in range(x.shape[1]):
  normal[:,j]=extended[:,j]/max(ideal[j],eps) if benefit[j] else max(ideal[j],eps)/np.maximum(extended[:,j],eps)
 S=normal@weights;ka=S[1:-1]/max(S[0],eps);ki=S[1:-1]/max(S[-1],eps)
 fki=ki/(ki+ka+eps);fka=ka/(ki+ka+eps)
 utility=(ki+ka)/(1+(1-fki)/np.maximum(fki,eps)+(1-fka)/np.maximum(fka,eps))
 return utility

def risk_table(a):
 r=a.copy();r['coverage_days']=r['stock'].clip(lower=0)/(r['mean_monthly'].clip(lower=0)/30+1e-6)
 r['capital']=r['stock'].clip(lower=0)*r['unit_cost'].clip(lower=0)
 r['discontinuity']=1-r['continuity']
 cols=['coverage_days','days_since_sale','discontinuity','cv','capital']
 x=r[cols].replace([np.inf,-np.inf],np.nan).fillna(0).to_numpy(float)
 # All criteria increase excess/obsolescence risk. No-sales coverage capped to avoid unbounded values.
 x[:,0]=np.minimum(x[:,0],3650)
 w=merec_weights(x,[True]*len(cols));r['risk_score']=marcos(x,w,[True]*len(cols))
 r['risk_rank']=r['risk_score'].rank(ascending=False,method='min').astype(int)
 return r.sort_values('risk_rank'),pd.DataFrame({'criterion':cols,'weight':w})

def policies(a,service,holding,order_cost,default_lead,budget,capacity,transfer_cost,discount):
 d=a.copy();z=norm.ppf(service)
 d['lead_used']=d['lead_days'].fillna(default_lead).clip(lower=1)
 d['unit_cost_used']=d['unit_cost'].fillna(0).clip(lower=0)
 d['daily_demand']=d['mean_monthly'].clip(lower=0)/30
 d['daily_sigma']=d['std_monthly'].fillna(0).clip(lower=0)/np.sqrt(30)
 d['safety_stock']=z*d['daily_sigma']*np.sqrt(d['lead_used'])
 d['rop']=np.ceil(d['daily_demand']*d['lead_used']+d['safety_stock']).astype(int)
 annual=d['daily_demand']*365
 h=(d['unit_cost_used']*holding).replace(0,np.nan)
 d['eoq']=np.sqrt(2*annual*order_cost/h)
 d['eoq']=d['eoq'].replace([np.inf,-np.inf],np.nan)
 d['order_qty']=np.where(d['stock']<=d['rop'],np.maximum(0,np.ceil(np.maximum(d['eoq'].fillna(0),d['rop']-d['stock']))),0).astype(int)
 d['action']=np.select([(d['total_qty']<=0)&(d['stock']>0),(d['coverage_days']>180)&(d['stock']>0),(d['order_qty']>0)],['İndirimli tasfiye / kampanya','Siparişi durdur / transfer adayı','Yeni sipariş ver'],default='Stoku koru')
 # Transfer from surplus branches to shortages of same SKU; allocate greedily.
 d['transfer_in']=0;d['transfer_out']=0;transfers=[]
 for sku,group in d.groupby('sku'):
  donors=group[group['stock']>group['rop']+group['mean_monthly']].sort_values('stock',ascending=False)
  receivers=group[group['order_qty']>0].sort_values('order_qty',ascending=False)
  for ri,rv in receivers.iterrows():
   need=int(rv['order_qty']-d.at[ri,'transfer_in'])
   for di,dv in donors.iterrows():
    if di==ri or need<=0:continue
    surplus=max(0,int(d.at[di,'stock']-d.at[di,'rop']-d.at[di,'mean_monthly']-d.at[di,'transfer_out']))
    q=min(need,surplus)
    if q>0:
     d.at[di,'transfer_out']+=q;d.at[ri,'transfer_in']+=q;need-=q
     transfers.append({'sku':sku,'from':d.at[di,'branch'],'to':d.at[ri,'branch'],'qty':q,'cost':q*transfer_cost})
 d['order_qty']=(d['order_qty']-d['transfer_in']).clip(lower=0)
 # Enforce global capacity and budget with risk of stockout priority.
 d['shortage_priority']=((d['rop']-d['stock']-d['transfer_in']).clip(lower=0)*d['daily_demand']).fillna(0)
 remaining_budget=float(budget);remaining_capacity=int(capacity);approved=[]
 for i,row in d.sort_values('shortage_priority',ascending=False).iterrows():
  q=int(row['order_qty']);cost=float(row['unit_cost_used']);max_budget=int(remaining_budget/cost) if cost>0 else q
  take=max(0,min(q,remaining_capacity,max_budget));approved.append((i,take));remaining_capacity-=take;remaining_budget-=take*cost
 d['approved_qty']=0
 for i,q in approved:d.at[i,'approved_qty']=q
 d.loc[d['transfer_in']>0,'action']='Şubeler arası transfer'
 d.loc[d['approved_qty']>0,'action']='Yeni sipariş ver'
 d['purchase_cost']=d['approved_qty']*d['unit_cost_used']
 d['holding_cost_annual']=(d['stock'].clip(lower=0)+d['approved_qty']/2)*d['unit_cost_used']*holding
 d['order_cost_annual']=np.where(d['approved_qty']>0,annual/np.maximum(d['approved_qty'],1)*order_cost,0)
 d['coverage_after_days']=(d['stock']+d['approved_qty']+d['transfer_in']-d['transfer_out']).clip(lower=0)/d['daily_demand'].replace(0,np.nan)
 d['stockout_proxy']=np.maximum(0,d['rop']-(d['stock']+d['approved_qty']+d['transfer_in']-d['transfer_out']))
 d['discount_recovery']=np.where(d['action']=='İndirimli tasfiye / kampanya',d['stock'].clip(lower=0)*d['unit_cost_used']*(1-discount),0)
 return d,pd.DataFrame(transfers,columns=['sku','from','to','qty','cost'])

# Türkçe arayüz sözlüğü: soru işaretleri, grafikler, tablolar ve Excel raporları aynı terimleri kullanır.
TERIMLER = {
 'branch': ('Şube', 'Ürünün satıldığı veya stoklandığı CanToys mağazası.'),
 'sku': ('Ürün kodu', 'ERP sisteminde ürünü benzersiz tanımlayan stok kodu.'),
 'product': ('Ürün adı', 'ERP raporundaki ürün açıklaması.'),
 'abc': ('ABC sınıfı', 'Net ciroya katkı temelinde A: yüksek, B: orta, C: düşük önem sınıfı.'),
 'fsn': ('FSN sınıfı', 'F: sık satılan; S: seyrek satılan; N: gözlenen dönemde net satışı olmayan ürün.'),
 'total_qty': ('Net satış adedi', 'Satış miktarından satış iadeleri düşüldükten sonraki adet.'),
 'total_revenue': ('KDV hariç net ciro (₺)', 'Satış ve iadelerin KDV arındırılmış toplam parasal değeri.'),
 'continuity': ('Satış sürekliliği', 'Satış yapılan ay sayısının incelenen toplam ay sayısına oranı. 1, her ay satış anlamına gelir.'),
 'cv': ('Talep değişkenliği (CV)', 'Aylık satış standart sapmasının aylık ortalama satışa oranı. Yüksek değer daha düzensiz talebi gösterir.'),
 'trend_slope': ('Aylık satış eğilimi', 'Doğrusal regresyon eğimi: pozitifse artış, negatifse azalış eğilimi.'),
 'days_since_sale': ('Son satıştan geçen gün', 'Ürünün son pozitif satış tarihinden analiz tarihine kadar geçen gün.'),
 'risk_rank': ('Risk sırası', 'MARCOS stok fazlası/eskime riskine göre sıra. 1 en yüksek riskli üründür.'),
 'risk_score': ('MARCOS risk puanı', 'MEREC ağırlıklarıyla MARCOS yönteminden hesaplanan göreli stok fazlası/eskime riski puanı.'),
 'coverage_days': ('Stok karşılama süresi (gün)', 'Mevcut stokun ortalama günlük talebi kaç gün karşılayabileceğine ilişkin tahmin.'),
 'discontinuity': ('Satış düzensizliği', '1 eksi satış sürekliliği. Yüksek değer satışın daha seyrek olduğunu gösterir.'),
 'capital': ('Stoka bağlı sermaye (₺)', 'Mevcut stok adedi × birim alış maliyeti.'),
 'stock': ('Stok adedi', 'Doğrulanmış envanter yüklenmezse yalnızca hareketlerden türetilen güvenilmez yaklaşık değer.'),
 'rop': ('Yeniden sipariş noktası (ROP)', 'Tedarik süresindeki beklenen talep ile emniyet stokunun toplamı.'),
 'safety_stock': ('Emniyet stoku', 'Talep belirsizliği ve hedef hizmet düzeyine göre hesaplanan güvenlik miktarı.'),
 'eoq': ('Ekonomik sipariş miktarı (EOQ)', 'Sipariş verme ve stok taşıma maliyetlerini dengeleyen teorik sipariş adedi.'),
 'order_qty': ('Önerilen sipariş adedi', 'Transfer sonrası hesaplanan, bütçe ve kapasite onayından önceki sipariş adedi.'),
 'approved_qty': ('Sınırlar sonrası sipariş adedi', 'Bütçe ve adet üst sınırları uygulandıktan sonra önerilen miktar.'),
 'transfer_in': ('Gelen transfer adedi', 'Diğer CanToys şubelerinden bu şubeye önerilen transfer miktarı.'),
 'transfer_out': ('Giden transfer adedi', 'Bu şubeden başka şubeye önerilen transfer miktarı.'),
 'action': ('Önerilen işlem', 'Hesaplanan duruma göre yeni sipariş, siparişi durdurma, tasfiye veya transfer önerisi.'),
 'purchase_cost': ('Sipariş tutarı (₺)', 'Onaylanan sipariş adedi × birim maliyet. Eksik maliyetlerde güvenilir değildir.'),
 'criterion': ('Risk ölçütü', 'MEREC yönteminin ağırlıklandırdığı stok riski ölçütü.'),
 'weight': ('MEREC ağırlığı', 'Ölçütün veri dağılımından nesnel olarak hesaplanan göreli önemi. Ağırlıkların toplamı 1’dir.'),
 'method': ('Tahmin yöntemi', 'Geçmiş satışlardan gelecek ay talebini hesaplamak için kullanılan yöntem.'),
 'MAE': ('Ortalama mutlak hata (MAE)', 'Gerçek ve tahmin edilen adet farklarının mutlak değerlerinin ortalaması. Düşük olması iyidir.'),
 'RMSE': ('Kök ortalama kare hata (RMSE)', 'Büyük tahmin hatalarını daha fazla cezalandıran hata ölçüsü. Düşük olması iyidir.'),
 'MASE': ('Ölçeklenmiş mutlak hata (MASE)', 'Tahmin hatasının basit naif tahmin hatasına göre ölçeklenmiş değeri. Düşük olması iyidir.'),
 'month': ('Ay', 'Satış hareketinin ait olduğu takvim ayı.'),
 'qty': ('Adet', 'İlgili dönemde net satılan ürün miktarı.'),
 'revenue': ('Net ciro (₺)', 'KDV hariç ve iadeler düşülmüş satış tutarı.'),
 'from': ('Gönderen şube', 'Transfer önerisinde stok çıkışı yapacak şube.'),
 'to': ('Alıcı şube', 'Transfer önerisinde stok alacak şube.'),
 'cost': ('Transfer maliyeti (₺)', 'Transfer edilen adet × birim transfer maliyeti.')
}

def turkce_tablo(df):
    """Tablo başlıklarına Türkçe adlar ve başlık üzerinde ? açıklamaları ekler."""
    config = {}
    for col in df.columns:
        if col in TERIMLER:
            baslik, aciklama = TERIMLER[col]
            config[col] = st.column_config.Column(baslik, help=aciklama)
    return config

def tablo(df, **kwargs):
    st.dataframe(df, column_config=turkce_tablo(df), hide_index=True, use_container_width=True, **kwargs)

def grafik(fig):
    fig.update_layout(font=dict(family='Arial, sans-serif'),
                      paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                      legend_title_text='Şube', hovermode='closest')
    st.plotly_chart(fig, use_container_width=True)

st.title('🧸 CanToys • Akıllı Stok Karar Merkezi')
st.caption('🌈 PLUS • XL • EXTRA | Veriyi anla → Talebi tahmin et → Riski değerlendir → Karar ver')
st.caption('☁️ GitHub ve Streamlit Community Cloud üzerinde çalışmaya uygun sürüm • Dosyalar uygulama klasöründen okunur.')
st.info('💡 **Nasıl kullanılır?** Alanların yanındaki **ⓘ** işaretinin veya tablolardaki **sütun başlıklarının** üzerine fareyi getirerek açıklamalarını okuyabilirsin. Grafiklerdeki nokta ve çubukların üzerine gelerek ayrıntılı değerleri görebilirsin.')
with st.expander('📖 Yöntemleri ve kavramları öğren — Türkçe mini sözlük'):
 st.markdown('**ABC:** Ürünlerin KDV hariç cirosuna göre önem sıralaması. **FSN:** Satışın ne kadar sık gerçekleştiğine göre hareket sınıflandırması. **CV:** Talebin ortalamaya göre ne kadar değişken olduğu.')
 st.markdown('**Hareketli ortalama:** Son birkaç ayın ortalama satışı. **Üstel düzeltme:** Yakın dönem satışlarına daha fazla ağırlık veren tahmin. **Croston / SBA:** Satış olmayan ayların sık görüldüğü aralıklı talep için tahmin; SBA, Croston tahminindeki yanlılığı azaltmaya çalışır.')
 st.markdown('**MAE, RMSE, MASE:** Tahmin hatası ölçütleri; daha küçük değer daha iyidir. **MEREC:** Her ölçüt çıkarıldığında karar değerlendirmesinin ne kadar değiştiğini kullanarak ölçüt ağırlığı hesaplar. **MARCOS:** Bu ağırlıklarla ürünleri ideal ve ideal olmayan durumlara göre sıralar.')
 st.markdown('**ROP:** Stokun yeniden sipariş edilmesi gereken seviye. **EOQ:** Sipariş ve elde bulundurma maliyetlerini dengeleyen teorik miktar. **Hizmet düzeyi:** Talebin stoktan karşılanabilmesi için hedeflenen olasılık; gerçekleşmiş hizmet oranı değildir.')
 st.warning('Bu dosyalardaki ERP hareketleri güncel envanterin kesin kanıtı değildir. Doğrulanmış stok ve maliyet verisi olmadan stok riski, sipariş ve transfer önerilerini uygulamayın.')
with st.sidebar:
 st.header('🎠 CanToys ERP Veri Merkezi')
 embedded=[Path(__file__).resolve().parent/'stoklar1.xlsx',Path(__file__).resolve().parent/'stokk2.xlsx']
 has_embedded=all(x.is_file() for x in embedded)
 use_embedded=st.checkbox('📚 Projeye entegre ERP dosyalarını kullan',value=has_embedded,disabled=not has_embedded,help='Programın içindeki stoklar1.xlsx ve stokk2.xlsx hareket raporlarını otomatik okur. İşlemleri iki kez saymamak için yinelenen hareketler ayıklanır.')
 sf=None if use_embedded else st.file_uploader('Satış hareketleri dosyası (CSV / Excel)',type=['csv','xlsx'],key='sales',help='Satış, iade, işlem tarihi, şube ve ürün kodu içeren ERP hareket raporunu yükleyin.')
 tf=st.file_uploader('📦 Doğrulanmış güncel stok dosyası (isteğe bağlı)',type=['csv','xlsx'],key='stocks',help='Şube, ürün kodu, gerçek stok adedi; mümkünse birim maliyet ve tedarik süresi alanları olan envanter dosyası. Bu dosya olmadan sipariş ve risk önerileri kesin değildir.')
 st.caption('Güncel stok dosyası sütunları: branch (şube), sku (ürün kodu), stock (stok adedi); isteğe bağlı: unit_cost (birim maliyet), lead_days (tedarik süresi).')
 st.divider();st.header('⚙️ Hesaplama ayarları')
 service=st.slider('Hedef hizmet düzeyi',.80,.99,.95,.01,help='Tedarik süresindeki talebi stoktan karşılayabilmek için hedeflenen olasılık. Örneğin %95; gerçekleşen satış karşılama oranı değildir.')
 holding=st.slider('Yıllık stok taşıma maliyeti oranı',.05,.60,.25,.01,help='Stokta tutulan ürünün yıllık sermaye, depolama ve benzeri taşıma maliyetinin birim maliyete oranı.')
 order_cost=st.number_input('Sipariş başına işlem maliyeti (₺)',min_value=0.,value=150.,step=10.,help='Her siparişte katlanılan sabit işlem/lojistik maliyeti. EOQ hesabında kullanılır.')
 default_lead=st.number_input('Varsayılan tedarik süresi (gün)',min_value=1,value=14,help='Ürüne özel tedarik süresi bilinmiyorsa siparişin mağazaya ulaşması için varsayılan gün sayısı.')
 budget=st.number_input('Toplam sipariş bütçesi (₺)',min_value=0.,value=100000.,step=1000.,help='Tüm şubeler için önerilen satın almaların toplam harcama üst sınırı. Maliyet eksikse bütçe kontrolü güvenilir değildir.')
 capacity=st.number_input('Toplam sipariş adet üst sınırı',min_value=0,value=5000,step=100,help='Üç şube genelinde satın alınabilecek en fazla ürün adedi.')
 transfer_cost=st.number_input('Şubeler arası birim transfer maliyeti (₺)',min_value=0.,value=5.,help='Bir ürünün başka şubeye taşınmasının varsayılan birim maliyeti.')
 discount=st.slider('Tasfiye indirimi oranı',0.,.80,.30,.05,help='Yavaş veya hiç satılmayan stoklar için düşünülen indirim oranı. Öneri niteliğindedir.')
 window=st.slider('Hareketli ortalama dönemi (ay)',2,6,3,help='Tahmin için geriye dönük kaç aylık satışın ortalaması alınacak?')
 alpha=st.slider('Üstel düzeltme katsayısı (α)',.05,.95,.30,.05,help='Son satışlara verilen ağırlık. Büyük değerler yakın dönemdeki değişikliklere daha hızlı tepki verir.')

if not use_embedded and sf is None:
 st.info('📚 ERP dosyaları bulunamadı. Sol panelden ERP satış dosyası yükleyin.')
 st.stop()
try:
 if use_embedded:
  parts=[parse_cantoys_erp(path) for path in embedded]
  raw_sales=pd.concat([part[0] for part in parts],ignore_index=True)
  ledger=pd.concat([part[1] for part in parts],ignore_index=True)
  # Deduplicate overlapping exports by transaction ID and SKU.
  ledger=ledger.drop_duplicates(subset=['transaction_id','sku','transaction','raw_qty'],keep='first')
  raw_sales=ledger[ledger.is_sale & ledger.branch.notna()].copy()
  sales=prepare_sales(raw_sales)
 elif sf.name.lower().endswith('.xlsx'):
  try:
   raw_sales,ledger=parse_cantoys_erp(sf);sales=prepare_sales(raw_sales)
  except ValueError:
   ledger=None;sales=prepare_sales(load_file(sf))
 else:
  ledger=None;sales=prepare_sales(load_file(sf))
 stock_verified=tf is not None
 if stock_verified:stock=prepare_stock(load_file(tf))
 else:
  stock=ledger_stock_proxy(ledger) if ledger is not None else pd.DataFrame(columns=['branch','sku','stock','unit_cost','lead_days'])
 if sales.empty:st.error('Geçerli satış kaydı yok.');st.stop()
 panel=monthly_panel(sales);asof=max(pd.Timestamp.today().normalize(),sales['date'].max().normalize())
 a=classify(sales,panel,stock,asof)
 missing=a['stock'].isna().sum();a['stock']=a['stock'].fillna(0)
 risk,w=risk_table(a)
 plan,transfers=policies(risk,service,holding,order_cost,default_lead,budget,capacity,transfer_cost,discount)
except Exception as e:
 st.error(f'Veri işleme hatası: {e}');st.info('GitHub deposunda stoklar1.xlsx ve stokk2.xlsx dosyalarının bulunduğunu kontrol edin. Excel dosyaları eksikse sol panelden kendi verinizi yükleyebilirsiniz.');st.stop()
if not stock_verified:st.warning('⚠️ Güncel stok doğrulanmadı. Gösterilen stoklar yalnızca gözlenen şube hareketlerinden türetilmiş tahmini bakiyelerdir; başlangıç stokları ve tüm transferler bilinmediği için ROP, EOQ, MEREC–MARCOS stok riski ve sipariş/transfer kararları OPERASYONEL OLARAK KULLANILMAMALIDIR. Gerçek stok dosyasını yükleyin.')
if missing:st.warning(f'{missing} ürün-şube kaydı stok dosyasında bulunamadı; bu satırlarda stok 0 varsayıldı.')
if a['unit_cost'].isna().any():st.warning('Bazı ürünlerin birim maliyeti eksik; bu ürünlerde EOQ ve maliyet sonuçları güvenilir değildir.')
if (sales['date'].max()<pd.Timestamp.today()-pd.Timedelta(days=90)):st.warning('Satış verisi 90 günden eski; son satıştan geçen gün ve risk puanları bundan etkilenir.')
branches=sorted(a['branch'].unique());chosen=st.multiselect('🏬 İncelenecek şubeler',branches,default=branches,help='Grafik ve tablolarda görüntülenecek şubeleri seçin. Birden fazla şube seçilebilir.')
if not chosen:st.stop()
sub=plan[plan.branch.isin(chosen)];pnl=panel[panel.branch.isin(chosen)]
tabs=st.tabs(['🏠 Genel Bakış','🎯 ABC • FSN','🔮 Talep Tahmini','🚨 MEREC • MARCOS','📦 Stok Politikası','🧪 Senaryo Laboratuvarı'])
with tabs[0]:
 st.caption('📚 Veri kaynağı: '+(' + '.join(p.name for p in embedded) if use_embedded else sf.name)+' | Stok: '+('doğrulanmış dosya' if stock_verified else 'tahmini hareket bakiyesi'))
 c=st.columns(4);c[0].metric('KDV hariç net ciro',f"₺{sub.total_revenue.sum():,.0f}");c[1].metric('Net satılan adet',f"{sub.total_qty.sum():,.0f}");c[2].metric('Stok (doğrulanmadı)' if not stock_verified else 'Mevcut stok',f"{sub.stock.sum():,.0f}");c[3].metric('Taslak sipariş önerisi',f"{sub.approved_qty.sum():,.0f}")
 monthly=pnl.groupby(['month','branch'],as_index=False).agg(qty=('qty','sum'),revenue=('revenue','sum'))
 grafik(px.line(monthly,x='month',y='revenue',color='branch',markers=True,title='Aylık net ciro',color_discrete_sequence=COLORS,labels={'month':'Ay','revenue':'KDV hariç net ciro (₺)','branch':'Şube','stock':'Stok adedi'}))
 grafik(px.bar(sub.groupby('branch',as_index=False)['stock'].sum(),x='branch',y='stock',color='branch',title='Şubelere göre stok',color_discrete_sequence=COLORS,labels={'month':'Ay','revenue':'KDV hariç net ciro (₺)','branch':'Şube','stock':'Stok adedi'}))
with tabs[1]:
 c1,c2=st.columns(2)
 c1.plotly_chart(px.histogram(sub,x='abc',color='branch',title='Ciro önem sınıflarına göre ürün sayısı',labels={'abc':'ABC sınıfı','branch':'Şube','count':'Ürün sayısı'},category_orders={'abc':['A','B','C']}))
 c2.plotly_chart(px.histogram(sub,x='fsn',color='branch',title='Satış sıklığına göre ürün sayısı',labels={'fsn':'FSN sınıfı','branch':'Şube','count':'Ürün sayısı'},category_orders={'fsn':['F','S','N']}))
 tablo(sub[['branch','sku','product','abc','fsn','total_qty','total_revenue','continuity','cv','trend_slope','days_since_sale']])
with tabs[2]:
 opts=sub[['branch','sku','product']].drop_duplicates();labels=[f'{r.branch} | {r.sku} | {r.product}' for r in opts.itertuples()]
 selected=st.selectbox('Tahmin edilecek ürün',range(len(labels)),format_func=lambda i:labels[i],help='Tahmin ve hata analizi yapılacak ürün ve şubeyi seçin.');rr=opts.iloc[selected]
 ydf=pnl[(pnl.branch==rr.branch)&(pnl.sku==rr.sku)].sort_values('month');y=ydf.qty.to_numpy(float)
 bt=backtest(y,window,alpha)
 preds=forecast_methods(y,window=window,alpha=alpha)
 if not bt.empty:
  tablo(bt.sort_values('MAE'))
  best=bt.sort_values('MAE').iloc[0]['method']
 else:
  best='SBA' if np.mean(y==0)>.5 else 'Üstel Düzeltme'
  st.warning('En az 5 aylık veri olmadan geriye dönük hata karşılaştırması yapılamaz. Varsayılan yöntem gösteriliyor.')
 st.metric('En düşük MAE değerli yöntem',best,help='Geriye dönük tahminlerde ortalama mutlak hatası en küçük olan yöntem. Yeterli veri yoksa geçici varsayılan yöntem kullanılır.');st.metric('Gelecek ay tahmini',f'{preds[best]:.1f} adet',help='Seçilen yöntemle gelecek ay için tahmin edilen net satış miktarıdır; kesin satış değildir.')
 fig=go.Figure();fig.add_trace(go.Scatter(x=ydf.month,y=y,mode='lines+markers',name='Gerçek satış'))
 nextmonth=ydf.month.max()+pd.offsets.MonthBegin(1)
 for method,v in preds.items():fig.add_trace(go.Scatter(x=[ydf.month.max(),nextmonth],y=[y[-1],v],mode='lines+markers',name=method,line={'dash':'dot'}))
 fig.update_layout(title='Aylık satış ve bir dönem ileri tahmin',xaxis_title='Ay',yaxis_title='Adet');grafik(fig)
 st.caption('📘 MAE: ortalama mutlak hata; RMSE: kök ortalama kare hata; MASE: ölçeklenmiş mutlak hata. Üçünde de düşük değer tercih edilir.')
 st.caption('MASE: eğitim serisindeki naif farklara göre ölçeklenir; sabit serilerde tanımsız olabilir. Croston/SBA pozitif aralıklı talep için uygundur.')
with tabs[3]:
 st.caption('📘 Grafikteki ölçütlerin ayrıntılı açıklamaları yukarıdaki Türkçe mini sözlükte ve tablodaki sütun başlıklarının üzerindedir.')
 st.caption('⚠️ Stok doğrulanmadıysa bu sıralama yalnızca deneme amaçlıdır. Bu ekran yüksek stok / eskime riskini sıralar. Stok tükenmesi riski ayrı olarak stok politikası ekranında ele alınır.')
 grafik(px.bar(w,x='criterion',y='weight',title='MEREC yöntemiyle hesaplanan risk ölçütü ağırlıkları',color='criterion',labels={'criterion':'Risk ölçütü','weight':'Ağırlık'}))
 tablo(sub.sort_values('risk_rank')[['risk_rank','branch','sku','product','risk_score','coverage_days','days_since_sale','discontinuity','cv','capital']])
 grafik(px.bar(sub.nsmallest(15,'risk_rank'),x='risk_score',y='product',color='branch',orientation='h',title='Stok fazlası / eskime riski en yüksek 15 ürün',labels={'risk_score':'Göreli risk puanı','product':'Ürün adı','branch':'Şube'}))
with tabs[4]:
 c=st.columns(4);c[0].metric('ROP toplamı',f'{sub.rop.sum():,.0f}');c[1].metric('Onaylanan sipariş',f'{sub.approved_qty.sum():,.0f}');c[2].metric('Sipariş maliyeti',f"₺{sub.purchase_cost.sum():,.0f}");c[3].metric('Transfer adedi',f"{sub.transfer_in.sum():,.0f}")
 tablo(sub[['branch','sku','product','stock','rop','safety_stock','eoq','order_qty','approved_qty','transfer_in','transfer_out','action','purchase_cost']])
 st.subheader('🚚 Şubeler arası transferler')
 tablo(transfers[transfers['from'].isin(chosen)&transfers['to'].isin(chosen)])
 st.caption('⚠️ Güncel stok dosyası yoksa kararları uygulamayın. EOQ sabit talep/sipariş maliyeti varsayar; sıfır veya eksik maliyetlerde hesaplanmaz. Transferler aynı SKU üzerinden önerilir.')
with tabs[5]:
 st.subheader('🎮 Parametrelerle kararları karşılaştır')
 alt_service=st.slider('Alternatif hedef hizmet düzeyi',.80,.99,min(.99,round(service+.02,2)),.01,key='alt_service',help='Alternatif senaryoda hedeflenen talebi karşılama olasılığı.')
 alt_budget=st.number_input('Alternatif bütçe (₺)',min_value=0.,value=float(budget),step=1000.,key='alt_budget',help='Alternatif senaryoda kullanılabilecek toplam satın alma bütçesi.')
 alt_capacity=st.number_input('Alternatif adet sınırı',min_value=0,value=int(capacity),step=100,key='alt_capacity',help='Alternatif senaryoda verilebilecek toplam sipariş adedi üst sınırı.')
 alt,_=policies(risk,alt_service,holding,order_cost,default_lead,alt_budget,alt_capacity,transfer_cost,discount)
 alt=alt[alt.branch.isin(chosen)]
 def summary(d):return {'Sipariş adedi':int(d.approved_qty.sum()),'Satın alma ₺':float(d.purchase_cost.sum()),'Yıllık elde tutma ₺':float(d.holding_cost_annual.sum()),'Yıllık sipariş ₺':float(d.order_cost_annual.sum()),'ROP açığı (adet)':float(d.stockout_proxy.sum())}
 compare=pd.DataFrame({'Mevcut':summary(sub),'Alternatif':summary(alt)}).reset_index(names='Gösterge')
 tablo(compare)
 grafik(px.bar(compare.melt('Gösterge',var_name='Senaryo',value_name='Değer'),x='Gösterge',y='Değer',color='Senaryo',barmode='group',title='Senaryo karşılaştırması',labels={'Gösterge':'Ölçüt','Değer':'Hesaplanan değer','Senaryo':'Senaryo'}))
 st.caption('ROP açığı, gerçekleşmiş stok tükenmesi veya ölçülmüş hizmet seviyesi değildir. Satın alma harcaması ve yıllık maliyetleri aynı toplamda karşılaştırmayın.')

st.divider();st.subheader('📤 Analiz çıktıları')
out=io.BytesIO()
with pd.ExcelWriter(out,engine='openpyxl') as writer:
 sub.rename(columns={k:v[0] for k,v in TERIMLER.items()}).to_excel(writer,sheet_name='Kararlar',index=False)
 w.rename(columns={k:v[0] for k,v in TERIMLER.items()}).to_excel(writer,sheet_name='MEREC_Agirlik',index=False)
 transfers.rename(columns={k:v[0] for k,v in TERIMLER.items()}).to_excel(writer,sheet_name='Transferler',index=False)
 pnl.rename(columns={k:v[0] for k,v in TERIMLER.items()}).to_excel(writer,sheet_name='Aylik_Satis',index=False)
 pd.DataFrame([(baslik,aciklama) for baslik,aciklama in TERIMLER.values()],columns=['Terim','Türkçe açıklama']).to_excel(writer,sheet_name='Terimler_Sozlugu',index=False)
st.download_button('📥 Excel karar raporunu indir',out.getvalue(),'CanToys_KDS_Rapor.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
st.caption('⚠️ Analitik karar desteğidir. Sipariş/transferler ERP sistemine otomatik işlenmez. Gerçek hizmet düzeyi için stok tükenme ve karşılanamayan talep verileri gerekir.')
