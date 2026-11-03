import config as cfg
import pandas as pd, re, numpy as np
exec(open('src/load_data.py',encoding='utf-8').read())          # s = 셀메이트
m=pd.read_excel(cfg.MUSINSA_SELLING,'상품정보',dtype=str,header=1)
src=open('src/step1_listing_match.py',encoding='utf-8').read()
exec(src[src.index('# ---- 컬러 사전'):src.index("m[['stem','color','multi']]")])
BASE=r'^([A-Z]{2}\d{2}[A-Z]{2}\d{2})'
KO2=dict(KO); KO2.update({'라이트':'LIGHT','다크':'DARK','딥':'DEEP','피그먼트':'PIGMENT','워싱':'WASHING','스카이':'SKY','민트':'MINT',
   '라벤더':'LAVENDER','소라':'SORA','멜란지':'MELANGE','차콜':'CHARCOAL','페일':'PALE','파스텔':'PASTEL','빈티지':'VINTAGE','인디고':'INDIGO','연':'LIGHT','진':'DEEP'})
KEYS=sorted(KO2,key=len,reverse=True)
def canon(c):
    c=str(c).upper().strip()
    for k in KEYS: c=c.replace(k,KO2[k])
    return re.sub(r'[^A-Z0-9가-힣]','',c).replace('GRAY','GREY').replace('MOCA','MOCHA')
def is_color(tok):
    c=canon(tok); return bool(c) and re.fullmatch(r'[A-Z]+',c) is not None
SZ=r'(?:XXS|XS|S|M|L|XL|XXL|2XL|XXXL|3XL|FREE|F|OS|\d{1,3})(?:\([^)]*\)[A-Z0-9]?)?|단일상품'
def parse_opt(v):
    v=re.sub(r'[\x00-\x1f\x7f]',' ',str(v)).strip()
    v=re.sub(r'\s+\(','(',v)
    if v in ('nan',''): return '',''
    t=[x for x in re.split(r'[\^\s/,]+',v) if x]
    size=''
    if t and re.fullmatch(SZ,t[-1],flags=re.I): size=t.pop().upper()
    elif len(t)>1 and re.fullmatch(SZ,t[0],flags=re.I): size=t.pop(0).upper()
    return ' '.join(t), size
def size_keys(z):
    z=str(z).upper().replace(' ','')
    if not z: return set()
    if z in ('단일상품','F','OS','FREE'): return {'FREE'}
    mm=re.fullmatch(r'([^(]+)\(([^)]*)\)',z)
    if mm:
        k={mm.group(1)}
        if re.fullmatch(SZ,mm.group(2),flags=re.I): k.add(mm.group(2))
        return k
    return {z}
PREF=['PIGMENT','WASHING','VINTAGE','LIGHT','DARK','DEEP','PALE','PASTEL','SKY','WHITE','BLACK']
TYPO={'GERY':'GREY','GRAY':'GREY','MOCA':'MOCHA','CHACOAL':'CHARCOAL','CHA':'CHARCOAL'}
def ctoks(c):
    """컬러 문자열 -> 정규화된 토큰 집합"""
    out=[]
    for t in re.split(r'[^0-9A-Za-z가-힣]+|(?<=[A-Za-z])X(?=[A-Z])',str(c).upper()):
        if not t or t=='X': continue
        t=canon(t); t=TYPO.get(t,t)
        if not re.fullmatch(r'[A-Z]+',t) or len(t)<2: continue
        changed=True
        while changed:
            changed=False
            for pf in PREF:
                if t.startswith(pf) and len(t)>len(pf)+2 and re.fullmatch(r'[A-Z]+',t[len(pf):]):
                    out.append(pf); t=t[len(pf):]; t=TYPO.get(t,t); changed=True; break
        out.append(t)
    return frozenset(out)
def color_score(want,have):
    """0=완전일치, 1=포함관계, None=불일치"""
    if not want or not have: return None
    if want==have: return 0
    if want<=have or have<=want: return 1+len(want^have)
    return None
# ---- 셀메이트 옵션 ----
s=s[s['상품코드'].ne('nan')&s['상품코드'].ne('')].copy()
s['재고']=pd.to_numeric(s['현재재고'],errors='coerce').fillna(0).astype(int)
s['base']=s['상품코드'].str.extract(BASE)[0].fillna(s['상품코드'])
def prod_color(n):
    n=re.sub(r'[\s_\-]*\d+\s*COL(OR)?S?\s*$','',str(n),flags=re.I)
    n=re.sub(r'\s-\s*Z$','',n)
    t=[x for x in re.split(r'[\s_]+',n.strip()) if x]
    mixed=any(re.search(r'[a-z가-힣]',w) for w in t)
    out=[]
    while len(t)>1:
        w=t[-1]; core=re.sub(r'\(.*?\)','',w)
        if is_color(core) and re.fullmatch(r'[가-힣]+',core): out.insert(0,core); t.pop(); continue
        if re.fullmatch(r'[A-Z][A-Z0-9*]*',core) and core not in NOTCOL and (mixed or core in EN or core.split('*')[0] in EN): out.insert(0,core); t.pop(); continue
        break
    return ctoks(' '.join(out))
s=s[s['상품코드'].notna()&~s['상품코드'].astype(str).isin(['nan',''])].copy()
_blank=s['옵션코드'].isna()|s['옵션코드'].astype(str).isin(['nan',''])
s['옵션코드']=np.where(_blank,s['상품코드'].astype(str)+' / '+s['옵션명'].astype(str)+' (옵션코드 없음)',s['옵션코드'])
s['pcol']=s['상품명'].map(prod_color)
po=s['옵션명'].apply(lambda v: pd.Series(parse_opt(v)))
s['ocol']=po[0].map(ctoks); s['osize']=po[1]
s['col']=[oc if oc else pc for oc,pc in zip(s['ocol'],s['pcol'])]
s['stem']=s['상품명'].map(lambda v: split_name(v)[0])
by_code={k:list(v) for k,v in s.groupby('상품코드').groups.items()}; by_base=s.groupby('base').groups; by_stem=s.groupby('stem').groups

# ---- 무신사 옵션 ----
o=pd.read_excel(cfg.MUSINSA_OPTION,dtype=str)
for c in ['국내출고/브랜드배송','국내출고/무신사배송','일본출고/무신사배송','총합 판매가능']: o[c]=pd.to_numeric(o[c],errors='coerce').fillna(0).astype(int)
po=o['옵션'].apply(lambda v: pd.Series(parse_opt(v)))
o['옵션컬러']=po[0]; o['옵션사이즈']=po[1]
o['sn']=o['스타일넘버'].fillna('').str.upper().str.replace(r'\(.*?\)','',regex=True).str.strip()
o['base']=o['sn'].str.extract(BASE)[0].fillna(o['sn'])
L=o.drop_duplicates('상품번호').set_index('상품번호')[['sn','base','상품명']].copy()
L['stem']=L['상품명'].map(lambda v: split_name(v)[0]); L['multi']=L['상품명'].map(lambda v: split_name(v)[2] and bool(re.search(r'\d\s*COL',v,re.I)))
L['lcol']=L['상품명'].map(prod_color)
L['set']=L['상품명'].str.contains(r'^\s*\[(?:SET|패키지|\d\s*PACK)',case=False,regex=True)
def cand_rows(lid):
    r=L.loc[lid]; idx=set()
    if r['sn'] in by_code: idx|=set(by_code[r['sn']])
    if r['base'] in by_base: idx|=set(by_base[r['base']])
    if len(r['stem'])>=5 and r['stem'] in by_stem: idx|=set(by_stem[r['stem']])
    return sorted(idx)
CAND={lid:cand_rows(lid) for lid in L.index}
# 스타일넘버 오입력 의심 상품 → 상품명으로 찾은 셀메이트 상품 기준으로 다시 연결
OVR={}
try:
    NC=pd.read_pickle(cfg.WORK+'namechk.pkl')
    for _,x in NC[NC['대조결과'].str.contains('상품명은 다른 셀메이트 상품과 같음')].iterrows():
        OVR[x['무신사상품번호']]=[c for c in x['상품명일치_셀메이트'].split(', ') if c]
except FileNotFoundError: pass
OVR.update(cfg.MANUAL_LINKS)                 # 사람이 확인한 수동 연결 (config.py)
SHORT_OVR=cfg.MANUAL_SHORT_LINKS            # 긴팔 상품 안의 (short) 옵션 → 반팔 상품
code_rows=s.groupby('상품코드').groups
def rows_of(codes): return sorted(set(i for c in codes for i in code_rows.get(c,[])))
RELINK=set()
for lid,codes in OVR.items():
    if lid in CAND and rows_of(codes): CAND[lid]=rows_of(codes); RELINK.add(lid)

def match_opt(r):
    lid=r['상품번호']; idx=CAND[lid]; lr=L.loc[lid]
    if lid in SHORT_OVR and re.search(r'\(short\)',str(r['옵션']),re.I):
        idx=rows_of(SHORT_OVR[lid]); r=r.copy(); r['옵션컬러']=re.sub(r'\(short\)','',r['옵션컬러'],flags=re.I).strip()
    if lr['set']: return None,'세트·패키지 상품(대조 제외)'
    if not idx: return None,'무신사 상품이 셀메이트에 없음'
    C=s.loc[idx]
    want=ctoks(r['옵션컬러']) if r['옵션컬러'] else frozenset()
    if not want:
        if lr['multi'] and len(set(C['col']))>1: return None,'컬러 미지정 옵션'
        want=lr['lcol']
    if want:
        sc=C['col'].map(lambda h: color_score(want,h))
        ok=sc.dropna()
        if ok.empty:
            import difflib
            w=''.join(sorted(want)); fz=C['col'].map(lambda h: difflib.SequenceMatcher(None,w,''.join(sorted(h))).ratio() if h else 0)
            if fz.max()>=0.85: return_hit=C.loc[fz[fz==fz.max()].index]
            else: return_hit=None
            if return_hit is not None: hit=return_hit
            elif r['sn'] in by_code and not r['옵션컬러']: hit=s.loc[[i for i in by_code[r['sn']] if i in s.index]]
            else: return None,'셀메이트에 해당 컬러 없음'
        else: hit=C.loc[ok[ok==ok.min()].index]
    else: hit=C
    if r['sn'] in set(hit['상품코드']): hit=hit[hit['상품코드']==r['sn']]
    z=r['옵션사이즈']
    if not z:
        if len(hit)==1: return hit.index[0],'일치'
        return None,'사이즈 정보 없음'
    zz=z.upper().replace(' ','')
    ex=hit[hit['osize'].str.upper().str.replace(' ','')==zz]
    if len(ex)==0:
        k=size_keys(z); ex=hit[hit['osize'].map(lambda v: bool(size_keys(v)&k))]
    if len(ex)==0 and len(hit)==1 and hit.iloc[0]['osize']=='': ex=hit
    if len(ex)==0: return None,'셀메이트에 해당 사이즈 없음'
    if ex['옵션코드'].nunique()>1:
        if ex['상품코드'].nunique()>1: return ex.sort_values('재고',ascending=False).index[0],'일치(동일컬러 상품 여러 개·확인필요)'
        return ex.index[0],'일치'
    return ex.index[0],'일치'
mm=o.apply(match_opt,axis=1)
o['_si']=mm.str[0]; o['옵션매칭']=mm.str[1]
rel=o['상품번호'].isin(RELINK)|(o['상품번호'].isin(SHORT_OVR)&o['옵션'].str.contains(r'\(short\)',case=False,regex=True))
o.loc[rel&o['옵션매칭'].str.startswith('일치'),'옵션매칭']=o.loc[rel&o['옵션매칭'].str.startswith('일치'),'옵션매칭']+' (스타일넘버 오입력 → 상품명 기준 재연결)'
S=s.loc[o['_si'].dropna().astype(int)]
for c_new,c_old in [('셀메이트상품코드','상품코드'),('셀메이트옵션코드','옵션코드'),('셀메이트상품명','상품명'),('셀메이트옵션명','옵션명'),('셀메이트판매여부','판매여부'),('셀메이트재고','재고')]:
    o[c_new]=o['_si'].map(lambda i: s.at[int(i),c_old] if pd.notna(i) else None)
o.to_pickle(cfg.WORK+'optres.pkl')
matched=set(o['셀메이트옵션코드'].dropna())
linked=set(o['셀메이트상품코드'].dropna())
miss=s[s['상품코드'].isin(linked)&~s['옵션코드'].isin(matched)].copy()
# 연결된 무신사 상품번호/명
lk=o.dropna(subset=['셀메이트상품코드']).drop_duplicates('셀메이트상품코드').set_index('셀메이트상품코드')[['상품번호','상품명','스타일넘버']].rename(columns={'상품번호':'무신사상품번호','상품명':'무신사상품명','스타일넘버':'무신사스타일넘버'})
miss=miss.join(lk,on='상품코드'); miss.to_pickle(cfg.WORK+'optmiss.pkl')
L.assign(ncand=[len(CAND[i]) for i in L.index]).to_pickle(cfg.WORK+'optL.pkl')
print(o['옵션매칭'].value_counts()); print('missing selmate opts',len(miss),'with stock',(miss['재고']>0).sum())
