import config as cfg
import pandas as pd, re, difflib, numpy as np
exec(open('src/load_data.py',encoding='utf-8').read())
m=pd.read_excel(x,'상품정보',dtype=str,header=1)
m['sn']=m['스타일넘버'].fillna('').str.upper().str.replace(r'\(.*?\)','',regex=True).str.strip()
BASE=r'^([A-Z]{2}\d{2}[A-Z]{2}\d{2})'
m['base']=m['sn'].str.extract(BASE)[0].fillna(m['sn'])
PRI={'판매중':1,'품절':2,'등록중':3,'검수반려':4,'판매중지':5,'삭제':6}
m['pri']=m['무신사'].map(PRI).fillna(9)

# ---- 컬러 사전 ----
KO={'블랙':'BLACK','화이트':'WHITE','아이보리':'IVORY','베이지':'BEIGE','차콜':'CHARCOAL','그레이':'GREY','회색':'GREY','네이비':'NAVY',
'블루':'BLUE','스카이블루':'SKYBLUE','소라':'SORA','카키':'KHAKI','브라운':'BROWN','핑크':'PINK','레드':'RED','퍼플':'PURPLE','올리브':'OLIVE',
'크림':'CREAM','다크베이지':'DARKBEIGE','다크네이비':'DARKNAVY','다크그레이':'DARKGREY','연베이지':'LIGHTBEIGE','연그레이':'LIGHTGREY',
'진청':'DEEPBLUE','연청':'LIGHTBLUE','중청':'MIDBLUE','흑청':'BLACKDENIM','멜란지':'MELANGE','오트밀':'OATMEAL','모카':'MOCHA','카멜':'CAMEL',
'그린':'GREEN','옐로우':'YELLOW','민트':'MINT','라벤더':'LAVENDER','와인':'WINE','버건디':'BURGUNDY','실버':'SILVER','골드':'GOLD',
'오렌지':'ORANGE','인디고':'INDIGO','데님':None,'딥블루':'DEEPBLUE','라이트블루':'LIGHTBLUE','라이트그레이':'LIGHTGREY','머스타드':'MUSTARD','코코아':'COCOA'}
KO={k:v for k,v in KO.items() if v}
EN=set(KO.values())|{'BLACK','WHITE','IVORY','BEIGE','CHARCOAL','GREY','GRAY','NAVY','BLUE','SKY','KHAKI','BROWN','PINK','RED','PURPLE','OLIVE','CREAM',
'LIGHT','DARK','DEEP','PALE','PIGMENT','WASHING','MELANGE','OATMEAL','MOCHA','MOCA','CAMEL','GREEN','YELLOW','MINT','LAVENDER','WINE','SILVER','GOLD',
'ORANGE','INDIGO','BUTTER','SALAD','SORA','STITCH','COATING','CAMO','MULTI','CHECK','SALT'}
COL0=re.compile(r'[\s_\-]*\(?\d+\s*COL(OR)?S?\)?\s*$',re.I)
# 데이터에서 컬러 단어 수집: 대소문자 섞인 상품명의 끝 대문자 단어들
def _upper_tail(t):
    out=[]
    for w in reversed(t):
        if re.fullmatch(r'[A-Z][A-Z0-9/]*',w): out.insert(0,w)
        else: break
    return out
from collections import Counter
tailc=Counter(); midc=Counter()
for nm in pd.concat([s['상품명'],m['상품명']]).dropna().unique():
    nm=COL0.sub('',re.sub(r'^\s*(\[[^\]]*\]\s*)+','',nm))
    t=re.sub(r'[_\-/]',' ',nm).split()
    if any(re.search(r'[a-z가-힣]',w) for w in t):
        tl=_upper_tail(t)
        if 0<len(tl)<len(t) and len(tl)<=3: tailc.update(tl)
        midc.update(w.upper() for w in t[:len(t)-len(tl)])
    else: midc.update(w.upper() for w in t[:-1])
EN.update(w for w,k in tailc.items() if k>midc[w] and not re.search(r'\d',w))
NOTCOL={'AV','T','SET','UNISEX','HOODIE','HOODY','SHIRT','PANTS','TOP','CROP','TEE','SKIRT','KNIT','JACKET'}
EN-= {'AV','T','SET','UNISEX','HOODIE','HOODY','SHIRT','PANTS','TOP','CROP','TEE','SKIRT','KNIT','JACKET'}
COL=re.compile(r'[\s_\-]*\(?\d+\s*COL(OR)?S?\)?\s*$',re.I)

def split_name(n):
    """반환: (stem, color_tuple, multi)  multi=컬러 묶음/컬러 표기 없음"""
    n=str(n).strip()
    n=re.sub(r'^\s*(\[[^\]]*\]\s*)+','',n)          # [SET] [패키지] 등 머리말
    multi=bool(COL.search(n)); n=COL.sub('',n)
    t=re.sub(r'[_\-/]',' ',n).split()
    mixed=any(re.search(r'[a-z]',w) for w in t)
    col=[]
    while len(t)>1:
        w=t[-1]; W=w.upper()
        if w in KO: col.insert(0,KO[w]); t.pop(); continue
        if re.fullmatch(r'[A-Z][A-Z0-9]*',w) and ((mixed and W not in NOTCOL) or W in EN): col.insert(0,W); t.pop(); continue
        break
    stem=re.sub(r'[^0-9a-z가-힣]','',' '.join(t).lower())
    c=''.join(col).replace('GRAY','GREY').replace('MOCA','MOCHA')
    return stem, c, (multi or not c)
def same_color(a,b): return a==b or (a and b and (a.endswith(b) or b.endswith(a)) and min(len(a),len(b))>=4)

m[['stem','color','multi']]=m['상품명'].apply(lambda v: pd.Series(split_name(v)))
by_sn=m.groupby('sn').groups; by_base=m.groupby('base').groups; by_stem=m.groupby('stem').groups
stems=[k for k in by_stem if len(k)>=5]

s['재고']=pd.to_numeric(s['현재재고'],errors='coerce').fillna(0).astype(int)
s=s[s['상품코드'].ne('nan')&s['상품코드'].ne('')]
def brand(c):
    if re.match(cfg.TARGET_CODE_PATTERN,c): return cfg.BRAND_KR
    if c.startswith('VV') or re.match(r'^\d{2}(SS|FW)VC',c): return '비바셔스'
    if c.startswith('ZD'): return '젠더스트'
    return '미분류'
g=s.groupby('상품코드',sort=False).agg(상품명=('상품명','first'),상품분류=('상품분류','first'),상품연도=('상품연도','first'),
   판매여부=('판매여부',lambda v: '판매' if (v=='판매').any() else v.mode()[0]),옵션수=('옵션코드','size'),
   재고합계=('재고','sum'),대표판매가=('대표판매가','first'),입고예정일=('입고예정일','max')).reset_index()
g['브랜드']=g['상품코드'].map(brand)
g['base']=g['상품코드'].str.extract(BASE)[0].fillna(g['상품코드'])

def color_filter(idx,sc,smulti,label_multi,label_color):
    """같은 상품 후보 중 컬러가 맞는 것만. 반환 (hits{idx:label}, other_color_idx)"""
    hits={}; other=[]
    for i in idx:
        if smulti or m.at[i,'multi']: hits[i]=label_multi
        elif same_color(sc,m.at[i,'color']): hits[i]=label_color
        else: other.append(i)
    return hits,other

def match(r):
    c,b,n=r['상품코드'],r['base'],r['상품명']
    stem,sc,smulti=split_name(n)
    cand={}
    # 1) 스타일넘버
    if c in by_sn: cand.update({i:'스타일넘버 일치' for i in by_sn[c]})
    if c!=b and b in by_sn:
        for i in by_sn[b]: cand.setdefault(i,'스타일넘버 일치(통합상품)')
    other=[]
    if b in by_base:
        idx=[i for i in by_base[b] if m.at[i,'sn'] not in (b,c)]
        h,o=color_filter(idx,sc,c==b and smulti,'스타일넘버 일치(컬러별 등록)','스타일넘버+컬러 일치')
        for i,v in h.items(): cand.setdefault(i,v)
        other+=o
    # 2) 상품명 (컬러/컬러묶음 제외한 이름으로) - 스타일넘버가 다르게 등록된 무신사 상품까지 함께 확인
    if len(stem)>=5 and stem in by_stem:
        h,o=color_filter(list(by_stem[stem]),sc,smulti,'상품명 일치(컬러묶음 상품)','상품명+컬러 일치')
        for i,v in h.items(): cand.setdefault(i,v)
        if not cand: other+=o
    if cand: return cand,[i for i in other if i not in cand]
    if len(stem)>=5:
        # 3) 유사 상품명 (오타·띄어쓰기·단어 차이)
        best=difflib.get_close_matches(stem,stems,n=3,cutoff=0.88)
        for bs in best:
            h,o=color_filter(list(by_stem[bs]),sc,smulti,'유사 상품명(확인필요)','유사 상품명+컬러(확인필요)')
            if h: return h,o
            other+=o
    return {},other

res=[]; used=set()
for _,r in g.iterrows():
    cand,other=match(r)
    if cand:
        idx=list(cand); used.update(idx)
        best=m.loc[idx].sort_values(['pri','등록 일시'],ascending=[True,False]).iloc[0]
        st=best['무신사']; others=', '.join(sorted(set(m.loc[idx,'무신사'])-{st}))
        res.append(dict(무신사상태=st,판정='등록·노출' if st=='판매중' else ('미등록' if st=='삭제' else '등록·미노출'),
            매칭방식=cand[best.name],무신사상품번호=best['상품 번호'],무신사상품명=best['상품명'],무신사스타일넘버=best['스타일넘버'],
            무신사판매가=best['판매가(원)'],무신사재고=best['판매 가능 재고'],다른매칭상태=others,
            무신사링크=f"https://www.musinsa.com/products/{best['상품 번호']}"))
    else:
        d=dict(무신사상태='-',판정='미등록',매칭방식='미매칭')
        if other:
            o=m.loc[other].sort_values('pri').iloc[0]
            d.update(매칭방식='미매칭(같은 상품 다른 컬러만 등록)',무신사상품번호=o['상품 번호'],무신사상품명=o['상품명'],무신사스타일넘버=o['스타일넘버'],
                     다른매칭상태='참고: 다른 컬러 상태 '+o['무신사'],무신사링크=f"https://www.musinsa.com/products/{o['상품 번호']}")
        res.append(d)
out=pd.concat([g.reset_index(drop=True),pd.DataFrame(res)],axis=1)
act=(out['재고합계']>0)   # 셀메이트 판매여부(판매/이월)와 상관없이 재고가 있으면 조치 대상
out['조치필요']=np.select([act&out['판정'].eq('미등록'),act&out['무신사상태'].eq('품절'),act&out['판정'].eq('등록·미노출')],
  ['미등록(재고있음)','무신사 품절표시(재고있음)','판매중지 등(재고있음)'],'')
m['셀메이트매칭']=m.index.isin(used)
out.to_pickle(cfg.WORK+'out.pkl'); m.to_pickle(cfg.WORK+'m.pkl')
u=out[out['브랜드']==cfg.BRAND_KR]
print(u.groupby(['판매여부','판정']).size().unstack(fill_value=0))
print(u['매칭방식'].value_counts()); print(u.groupby('조치필요')['재고합계'].agg(['size','sum']))
print(m[~m['셀메이트매칭']]['무신사'].value_counts())
