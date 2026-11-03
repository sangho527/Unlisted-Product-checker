import config as cfg
import pandas as pd, re, numpy as np, difflib
exec(open('src/step3_option_match.py',encoding='utf-8').read().split('# ---- 무신사 옵션 ----')[0])   # s(셀메이트), split_name 등
def stem(n):
    n=str(n)
    n=re.sub(r'\((?:N|\d)\)','',n)
    n=re.sub(r'\s*-\s*Z\s*$','',n)
    n=re.sub(r'[\s_\-]*\(?\d+\s*COL(OR)?S?\)?(\(N\))?\s*$','',n,flags=re.I)
    n=re.sub(r'\s*-\s*[A-Z]+\s+\d+\s*COL(OR)?S?(\(N\))?\s*$','',n,flags=re.I)  # "- BLACK 4COL(N)"
    st,c,mu=split_name(n)
    # 남은 끝 한글 컬러 제거
    t=re.split(r'\s+',n.strip())
    return st
s['stem']=s['상품명'].map(stem)
P=s.groupby('상품코드').agg(상품명=('상품명','first'),stem=('stem','first'),base=('base','first'),재고=('재고','sum'),
    판매여부=('판매여부',lambda v:'판매' if (v=='판매').any() else v.iloc[0]),브랜드=('상품코드',lambda v: v.iloc[0][:2])).reset_index()
P['어반드레스']=P['상품코드'].str.match(cfg.TARGET_CODE_PATTERN)
m=pd.read_excel(cfg.MUSINSA_SELLING,'상품정보',dtype=str,header=1)
m['sn']=m['스타일넘버'].fillna('').str.upper().str.replace(r'\(.*?\)','',regex=True).str.strip()
m['base']=m['sn'].str.extract(BASE)[0].fillna(m['sn'])
m['stem']=m['상품명'].map(stem)
m['set']=m['상품명'].str.contains(r'^\s*\[(?:SET|패키지|\d\s*PACK)',case=False,regex=True)
by_code=P.groupby('상품코드').groups; by_base=P.groupby('base').groups; by_stem=P.groupby('stem').groups
stems=[k for k in by_stem if len(k)>=4]
def lang(x): return 'ko' if re.search(r'[가-힣]',x) else 'en'
rows=[]
for _,r in m.iterrows():
    code=set()
    if r['sn'] in by_code: code|=set(by_code[r['sn']])
    if r['base'] in by_base: code|=set(by_base[r['base']])
    name=set(by_stem.get(r['stem'],[])) if len(r['stem'])>=4 else set()
    fuzzy=set()
    if len(r['stem'])>=6:
        for f in difflib.get_close_matches(r['stem'],stems,n=3,cutoff=0.86):
            if f!=r['stem']: fuzzy|=set(by_stem[f])
    cs=set(P.loc[list(code),'stem']) if code else set()
    KW=['short','long','반팔','긴팔','sleeveless','나시','mini','maxi','미니','롱','half','하프','crop','크롭','zipup','집업','hood','후드','skirt','스커트','pants','팬츠','shorts','쇼츠','cap','캡','bag','백']
    mk={k for k in KW if k in r['stem']}
    kw_conflict=bool(code) and bool(mk) and any(lang(x)==lang(r['stem']) for x in cs) and all(not ({k for k in KW if k in x}>=mk) for x in cs)
    if r['set']: res='세트상품(제외)'
    elif code and (code&name) and not kw_conflict: res='정상(스타일넘버·상품명 일치)'
    elif code and (code&name): res='확인: 상품 종류 키워드 다름(반팔/긴팔 등)'
    elif code and name: res='확인: 상품명은 다른 셀메이트 상품과 같음(스타일넘버 오입력 의심)'
    elif code and kw_conflict: res='확인: 상품 종류 키워드 다름(반팔/긴팔 등)'
    elif code and (code&fuzzy): res='정상(스타일넘버 일치, 상품명 표기 차이)'
    elif code:
        sim=max([difflib.SequenceMatcher(None,r['stem'],x).ratio() for x in cs],default=0)
        same_lang=any(lang(x)==lang(r['stem']) for x in cs)
        res='확인: 스타일넘버만 연결(상품명 다름)' if (same_lang and sim<0.75) else '정상(스타일넘버 일치, 상품명 표기 차이)'
    elif name: res='확인: 상품명으로만 찾음(스타일넘버 불일치)'
    elif fuzzy: name=fuzzy; res='확인: 비슷한 상품명으로만 찾음(스타일넘버 불일치)'
    else: res='셀메이트에서 못 찾음'
    sim=max([difflib.SequenceMatcher(None,r['stem'],x).ratio() for x in cs],default=0)
    def fmt(ix): 
        d=P.loc[list(ix)].sort_values('상품코드'); return ', '.join(d['상품코드'])
    def fmtn(ix):
        return ' / '.join(sorted(set(P.loc[list(ix),'상품명'].map(lambda v: re.sub(r'\s+[A-Z가-힣]+$','',v))))[:3])
    rows.append(dict(대조결과=res,무신사상품번호=r['상품 번호'],무신사상품명=r['상품명'],무신사스타일넘버=r['스타일넘버'],비교용상품명=r['stem'],
        스타일넘버연결_셀메이트=fmt(code),스타일넘버연결_상품명=fmtn(code) if code else '',상품명일치_셀메이트=fmt(name),
        상품명일치_상품명=fmtn(name) if name else '',유사도=round(sim,2) if code else None,무신사재고=r['판매 가능 재고'],
        링크='https://www.musinsa.com/products/'+r['상품 번호']))
R=pd.DataFrame(rows)
# 같은 셀메이트 스타일에 상품명이 다른 무신사 상품이 여러 개 연결된 경우
R['_b']=m['base'].values
grp=R[~R['대조결과'].str.startswith('세트')].groupby('_b')['비교용상품명'].nunique()
multi=set(grp[grp>1].index)
R['같은스타일_다른이름_무신사상품']=R['_b'].isin(multi)
R.to_pickle(cfg.WORK+'namechk.pkl')
# 셀메이트 쪽 누락: 재고 있는 판매 상품 중 무신사 판매중 상품과 스타일넘버·상품명 어느 쪽으로도 연결 안 됨
linked=set()
for x in R['스타일넘버연결_셀메이트'].tolist()+R['상품명일치_셀메이트'].tolist():
    linked|=set(t for t in x.split(', ') if t)
Q=P[P['어반드레스']&~P['상품코드'].isin(linked)&(P['재고']>0)&(P['판매여부']=='판매')].copy()
# 이전 점검(전체 상태 파일)에서의 무신사 상태 참고
old=pd.read_pickle(cfg.WORK+'out.pkl').set_index('상품코드')
Q['전체파일_무신사상태']=Q['상품코드'].map(old['무신사상태']); Q['전체파일_무신사상품명']=Q['상품코드'].map(old['무신사상품명']); Q['전체파일_상품번호']=Q['상품코드'].map(old['무신사상품번호'])
Q.to_pickle(cfg.WORK+'namechk_sel.pkl')
print(R['대조결과'].value_counts()); print('multi-name styles',R['같은스타일_다른이름_무신사상품'].sum())
print(len(Q), Q['재고'].sum()); print(Q['전체파일_무신사상태'].value_counts(dropna=False))
