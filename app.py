import re
from io import BytesIO
from pathlib import Path
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title='입고 예정 PLT Dashboard', page_icon='📦', layout='wide')
st.title('📦 입고 예정 PLT Dashboard')
st.caption('PU팀 입고 예정 PCS → 일자별·창고별 PLT 자동 환산')

ALIASES = {
    'sku': ['SKU ID','SKU','SKU_ID'],
    'product': ['제품 모델명','제품명','상품명','Product'],
    'line': ['라인','라인명','Line'],
    'pallet': ['팔레트 당 제품 수량','팔레트당 제품 수량','팔레트 당 PCS','Pallet Qty'],
}
EXCLUDE = ['협의중','합계','total','s&op','sop','누계']
UNMAPPED = '미매핑'

def clean(x):
    return re.sub(r'\s+', ' ', str(x).replace('\n',' ')).strip()

def find_col(df, aliases):
    cols = {clean(c): c for c in df.columns}
    for a in aliases:
        if a in cols: return cols[a]
    for c0,c in cols.items():
        for a in aliases:
            if a in c0: return c
    return None

def parse_header(x):
    s = clean(x); lo = s.lower()
    if any(k in lo for k in EXCLUDE): return None
    m = re.search(r'(20\d{2})\s*[-./]\s*(\d{1,2})\s*[-./]\s*(\d{1,2})', s)
    if m: return int(m.group(2)), int(m.group(3))
    m = re.search(r'(\d{1,2})\s*월\s*(\d{1,2})\s*일?', s)
    if m: return int(m.group(1)), int(m.group(2))
    m = re.search(r'(?<!\d)(\d{1,2})\s*[./-]\s*(\d{1,2})(?!\d)', s)
    if m:
        mo,day = map(int,m.groups())
        if 1 <= mo <= 12 and 1 <= day <= 31: return mo,day
    p = pd.to_datetime(x, errors='coerce')
    if pd.notna(p): return int(p.month),int(p.day)
    return None

def date_cols(df, month=None):
    out=[]
    for c in df.columns:
        p=parse_header(c)
        if p and (month is None or p[0]==month): out.append((c,p[0],p[1]))
    return out

def numeric(s):
    return pd.to_numeric(s.astype(str).str.replace(',','',regex=False).str.replace(' ','',regex=False),errors='coerce')

def read_file(f):
    ext=Path(f.name).suffix.lower()
    if ext in ('.xlsx','.xls'):
        return {s:pd.read_excel(f,sheet_name=s) for s in pd.ExcelFile(f).sheet_names}
    raw=f.getvalue()
    for enc in ('utf-8-sig','cp949','euc-kr','utf-8'):
        try: return {'CSV':pd.read_csv(BytesIO(raw),encoding=enc)}
        except Exception: pass
    raise ValueError('CSV 인코딩을 읽을 수 없습니다.')

def transform(raw, year, month, mapping=None, roundup=True):
    df=raw.copy(); df.columns=[clean(c) for c in df.columns]
    sku=find_col(df,ALIASES['sku']); product=find_col(df,ALIASES['product']); line=find_col(df,ALIASES['line']); pallet=find_col(df,ALIASES['pallet'])
    if not sku: raise ValueError('SKU ID 컬럼을 찾지 못했습니다.')
    if not pallet: raise ValueError('팔레트 당 제품 수량 컬럼을 찾지 못했습니다.')
    dcols=date_cols(df,month)
    if not dcols: raise ValueError(f'{month}월 날짜 컬럼을 찾지 못했습니다.')
    ids=[sku,pallet]+([product] if product else [])+([line] if line else [])
    ids=list(dict.fromkeys(ids)); names=[x[0] for x in dcols]
    out=df[ids+names].melt(id_vars=ids,value_vars=names,var_name='입고일자_raw',value_name='입고 PCS')
    mapping_date={c:pd.Timestamp(year=year,month=mo,day=day) for c,mo,day in dcols}
    out['입고일자']=out['입고일자_raw'].map(mapping_date); out['입고 PCS']=numeric(out['입고 PCS']); out['팔레트당 PCS']=numeric(out[pallet])
    out=out[out['입고 PCS'].notna() & (out['입고 PCS']!=0)].copy()
    out['SKU ID']=out[sku].astype(str).str.strip(); out['제품명']=out[product].astype(str).str.strip() if product else ''; out['라인']=out[line].astype(str).str.strip() if line else ''
    if mapping is not None and {'라인','창고'}.issubset(mapping.columns):
        mp=mapping[['라인','창고']].drop_duplicates('라인').copy(); out=out.merge(mp,on='라인',how='left')
    else: out['창고']=UNMAPPED
    out['창고']=out['창고'].fillna(UNMAPPED)
    valid=out['팔레트당 PCS'].gt(0)
    ratio=out['입고 PCS']/out['팔레트당 PCS']
    out['입고 PLT']=ratio.where(valid)
    if roundup: out.loc[valid,'입고 PLT']=ratio.loc[valid].apply(lambda x:int(-(-x//1)))
    out['입고 PLT']=pd.to_numeric(out['입고 PLT'],errors='coerce')
    return out.sort_values(['입고일자','창고','SKU ID']).reset_index(drop=True), dcols

with st.sidebar:
    st.header('⚙️ 설정')
    year=st.number_input('입고 연도',2020,2035,2026)
    month=st.selectbox('분석 월',range(1,13),index=8,format_func=lambda x:f'{x}월')
    roundup=st.checkbox('부분 팔레트 올림 (ROUNDUP)',True)
    st.caption('현재 원본에는 창고 컬럼이 명확하지 않아 LINE → 창고 매핑을 사용합니다.')

uploaded=st.file_uploader('PU 입고 예정 파일을 업로드하세요',type=['xlsx','xls','csv'])
if uploaded is None:
    st.info('월별 PU Excel/CSV를 업로드하면 날짜 컬럼을 자동 인식하고 PLT Dashboard를 생성합니다.')
    st.markdown('**자동 인식 대상:** SKU ID / 제품 모델명 / 라인 / 팔레트 당 제품 수량 / 날짜 컬럼\n\n**자동 제외:** 협의중 / 합계 / S&OP / 누계')
    st.stop()

try: sheets=read_file(uploaded)
except Exception as e: st.error(f'파일을 읽지 못했습니다: {e}'); st.stop()
sheet=st.selectbox('분석할 시트',list(sheets.keys())); raw=sheets[sheet].copy(); raw.columns=[clean(c) for c in raw.columns]
all_dates=date_cols(raw); target_dates=date_cols(raw,month)

with st.expander('🔎 자동 인식 결과',expanded=True):
    c1,c2,c3,c4=st.columns(4)
    c1.metric('SKU 컬럼',find_col(raw,ALIASES['sku']) or '미검출'); c2.metric('제품명',find_col(raw,ALIASES['product']) or '미검출'); c3.metric('라인',find_col(raw,ALIASES['line']) or '미검출'); c4.metric('팔레트당 PCS',find_col(raw,ALIASES['pallet']) or '미검출')
    st.write(f'전체 날짜 컬럼 {len(all_dates)}개 / {month}월 날짜 컬럼 {len(target_dates)}개')
    if target_dates: st.code(', '.join(str(x[0]) for x in target_dates))

mapping_file=st.file_uploader('선택: 라인 → 창고 매핑 파일 (라인, 창고)',type=['xlsx','xls','csv'],key='mapping')
mapping=None
if mapping_file:
    try:
        ms=read_file(mapping_file); mn=st.selectbox('매핑 시트',list(ms.keys()),key='mapping_sheet'); mapping=ms[mn].copy(); mapping.columns=[clean(c) for c in mapping.columns]
        if not {'라인','창고'}.issubset(mapping.columns): st.error('매핑 파일에는 라인, 창고 컬럼이 필요합니다.'); mapping=None
        else: st.dataframe(mapping[['라인','창고']].drop_duplicates(),use_container_width=True,hide_index=True)
    except Exception as e: st.error(f'매핑 파일 오류: {e}'); mapping=None
else: st.warning('창고 매핑이 없으므로 모든 데이터가 미매핑으로 표시됩니다.')

try: data,dcols=transform(raw,int(year),int(month),mapping,roundup)
except Exception as e: st.error(f'데이터 변환 실패: {e}'); st.stop()

raw_pcs=sum(numeric(raw[c]).sum() for c,_,_ in dcols); trans_pcs=data['입고 PCS'].sum(); diff=raw_pcs-trans_pcs
st.subheader('✅ 데이터 검증')
v1,v2,v3,v4,v5=st.columns(5); v1.metric('원본 날짜 PCS',f'{raw_pcs:,.0f}'); v2.metric('변환 후 PCS',f'{trans_pcs:,.0f}'); v3.metric('차이',f'{diff:,.0f}'); v4.metric('SKU 수',f"{data['SKU ID'].nunique():,}"); v5.metric('데이터 행',f'{len(data):,}')
if abs(diff)>0.01: st.warning('⚠️ 원본 날짜 PCS와 변환 후 PCS가 다릅니다. 중복 헤더/날짜 범위/원본 값을 확인하세요.')
else: st.success('원본 날짜 PCS와 변환 후 PCS가 일치합니다.')

st.divider(); st.subheader('📊 입고 예정 현황')
kp1,kp2,kp3,kp4=st.columns(4); kp1.metric('총 입고 예정 PLT',f"{data['입고 PLT'].sum():,.0f}"); kp2.metric('총 입고 예정 PCS',f"{data['입고 PCS'].sum():,.0f}"); kp3.metric('입고 SKU 수',f"{data['SKU ID'].nunique():,}"); kp4.metric('최대 일일 PLT',f"{data.groupby('입고일자')['입고 PLT'].sum().max():,.0f}")

st.subheader('🔎 Dashboard 필터')
f1,f2,f3=st.columns(3); whs=sorted(data['창고'].dropna().unique()); lines=sorted(data['라인'].dropna().unique())
sel_wh=f1.multiselect('창고',whs,default=whs); sel_line=f2.multiselect('라인',lines,default=lines)
mi,ma=data['입고일자'].min().date(),data['입고일자'].max().date(); dates=f3.date_input('입고일자',value=(mi,ma),min_value=mi,max_value=ma)
start,end=(pd.Timestamp(dates[0]),pd.Timestamp(dates[1])) if isinstance(dates,tuple) and len(dates)==2 else (pd.Timestamp(mi),pd.Timestamp(ma))
f=data[data['창고'].isin(sel_wh)&data['라인'].isin(sel_line)&data['입고일자'].between(start,end)].copy()

st.subheader('① 일자별 전체 입고 예정 PLT')
daily=f.groupby('입고일자',as_index=False)['입고 PLT'].sum(); fig=px.line(daily,x='입고일자',y='입고 PLT',markers=True); fig.update_layout(margin=dict(l=20,r=20,t=20,b=20)); st.plotly_chart(fig,use_container_width=True)

st.subheader('② 창고별 일자별 입고 예정 PLT')
wd=f.groupby(['입고일자','창고'],as_index=False)['입고 PLT'].sum(); fig=px.bar(wd,x='입고일자',y='입고 PLT',color='창고',barmode='stack'); fig.update_layout(margin=dict(l=20,r=20,t=20,b=20)); st.plotly_chart(fig,use_container_width=True)

st.subheader('③ 창고별 총 입고 예정 PLT')
wt=f.groupby('창고',as_index=False)['입고 PLT'].sum().sort_values('입고 PLT'); fig=px.bar(wt,x='입고 PLT',y='창고',orientation='h',text_auto='.0f'); fig.update_layout(margin=dict(l=20,r=20,t=20,b=20)); st.plotly_chart(fig,use_container_width=True)

st.subheader('④ 품목별 입고 예정 PLT TOP 10')
top=f.groupby(['SKU ID','제품명'],as_index=False)['입고 PLT'].sum().sort_values('입고 PLT',ascending=False).head(10).sort_values('입고 PLT'); top['표시명']=top['SKU ID']+' | '+top['제품명']; fig=px.bar(top,x='입고 PLT',y='표시명',orientation='h',text_auto='.0f'); fig.update_layout(margin=dict(l=20,r=20,t=20,b=20)); st.plotly_chart(fig,use_container_width=True)

st.subheader('⑤ 일자별 창고 입고 집중도')
heat=f.groupby(['창고','입고일자'],as_index=False)['입고 PLT'].sum().pivot(index='창고',columns='입고일자',values='입고 PLT').fillna(0); fig=px.imshow(heat,aspect='auto',text_auto='.0f',labels={'x':'입고일자','y':'창고','color':'입고 PLT'}); fig.update_layout(margin=dict(l=20,r=20,t=20,b=20)); st.plotly_chart(fig,use_container_width=True)

with st.expander('📋 변환된 입고 DATA 보기'):
    cols=['입고일자','창고','라인','SKU ID','제품명','입고 PCS','팔레트당 PCS','입고 PLT']; st.dataframe(f[cols],use_container_width=True,hide_index=True); st.download_button('⬇️ CSV 다운로드',f[cols].to_csv(index=False,encoding='utf-8-sig'),file_name=f'inbound_data_{year}_{month:02d}.csv',mime='text/csv')
