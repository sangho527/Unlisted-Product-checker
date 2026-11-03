import config as cfg
import pandas as pd, numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
r=pd.read_pickle(cfg.WORK+'judged.pkl'); miss=pd.read_pickle(cfg.WORK+'optmiss.pkl')
import re
OTHER=cfg.EXCLUDE_CODE_PATTERN   # 다른 브랜드 제외 (config.py)
def other_brand(code): return bool(re.match(OTHER,str(code).strip().upper()))
r=r[~r['스타일넘버'].map(other_brand)&~r['셀메이트상품코드'].map(lambda c: isinstance(c,str) and other_brand(c))].copy()
miss=miss[~miss['상품코드'].map(other_brand)&~miss['무신사스타일넘버'].map(other_brand)].copy()
PROB=['무신사 품절(셀메이트 재고 있음)','판매 제한(재고 입력됐는데 판매가능 0)']
ORDER=PROB+['품절(다른 상품번호에서는 판매중)','품절(같은 상품 안 다른 옵션명으로 판매중)','초과판매 위험(셀메이트 재고 없음)','확인필요(옵션 매칭 안 됨)','정상 판매','정상 판매(무신사 풀필먼트 출고)','정상 품절(양쪽 재고 없음)','대조 제외(세트상품)']
OTH=ORDER[2:4]; OVR_RISK='초과판매 위험(셀메이트 재고 없음)'; CHK='확인필요(옵션 매칭 안 됨)'
FILL={ORDER[0]:'F8CBAD',ORDER[1]:'F8CBAD',ORDER[2]:'FFE699',ORDER[3]:'FFE699',OVR_RISK:'FCE4D6',CHK:'E7E6E6','정상 판매':'C6EFCE','정상 판매(무신사 풀필먼트 출고)':'DDEBF7'}
r['_o']=r['판정'].map({k:i for i,k in enumerate(ORDER)})
r['링크']='https://www.musinsa.com/products/'+r['상품번호'].astype(str)
r['무신사배송재고']=r['국내출고/무신사배송']+r['일본출고/무신사배송']
R=pd.read_pickle(cfg.WORK+'namechk.pkl'); flag=R[R['대조결과'].str.startswith('확인')].set_index('무신사상품번호')['대조결과']
r['상품명대조']=r['상품번호'].map(flag).fillna('')
r=r.sort_values(['_o','셀메이트재고','상품번호'],ascending=[True,False,True])
F='Arial'; H=PatternFill('solid',fgColor='1F3864'); HF=Font(name=F,bold=True,color='FFFFFF',size=10)
wb=Workbook(); ws0=wb.active; ws0.title='요약'
DESC_FILL=PatternFill('solid',fgColor='EEF3FA')
def desc_height(desc,widths):
    """병합된 설명 셀에 줄바꿈된 글자가 다 보이도록 행 높이 계산 (한글은 폭 1.8로 계산)"""
    import math
    total=max(sum(widths),20)
    lines=0
    for ln in desc.split('\n'):
        units=sum(1.8 if ord(ch)>0x1100 else 1.0 for ch in ln)
        lines+=max(1,math.ceil(units*1.05/total))
    return max(30,15*lines+10)
def sheet(name,df,cols,heads,widths,colorcol=None,freeze='C3',desc=''):
    """1행: 시트 설명(왜 만들었는지/읽는 법) · 2행: 헤더 · 3행~: 데이터"""
    ws=wb.create_sheet(name)
    n=len(cols)
    ws.cell(1,1,desc).font=Font(name=F,size=10,color='1F3864')
    ws.cell(1,1).alignment=Alignment(wrap_text=True,vertical='top'); ws.cell(1,1).fill=DESC_FILL
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=n)
    for j,h in enumerate(heads,1):
        x=ws.cell(2,j,h); x.font=HF; x.fill=H; x.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    for i,row in enumerate(df[cols].itertuples(index=False),3):
        for j,v in enumerate(row,1):
            if v is None or (isinstance(v,float) and np.isnan(v)): v=None
            elif isinstance(v,(np.integer,)): v=int(v)
            elif isinstance(v,float) and v.is_integer(): v=int(v)
            x=ws.cell(i,j,v); x.font=Font(name=F,size=10)
            if isinstance(v,str) and v.startswith('https://'): x.hyperlink=v; x.font=Font(name=F,size=10,color='0563C1',underline='single')
            if isinstance(v,int): x.number_format='#,##0'
            if colorcol and cols[j-1]==colorcol and v in FILL: x.fill=PatternFill('solid',fgColor=FILL[v])
    for j,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(j)].width=w
    ws.row_dimensions[1].height=desc_height(desc,widths)
    ws.freeze_panes=freeze; ws.auto_filter.ref=f"A2:{get_column_letter(n)}{max(len(df)+2,3)}"; ws.row_dimensions[2].height=32
    return ws
# 2) 옵션별 점검
c1=['판정','상품번호','상품명','옵션','셀메이트옵션코드','셀메이트상품명','셀메이트옵션명','셀메이트재고','국내출고/브랜드배송','무신사배송재고','총합 판매가능','상품상태','다른상품번호에서판매중','셀메이트판매여부','상품명대조','링크','옵션매칭']
h1=['판정','무신사 상품번호','무신사 상품명','무신사 옵션','셀메이트 옵션코드','셀메이트 상품명','셀메이트 옵션명','셀메이트 재고','무신사 브랜드배송 재고','무신사배송(MFS) 재고','무신사 총합 판매가능','무신사 상품 상태','같은 옵션 판매중인 다른 상품번호','셀메이트 판매여부','상품명 대조(스타일넘버 확인 필요 상품)','무신사 링크','옵션 매칭']
sheet('옵션별_점검',r,c1,h1,[30,11,38,22,20,36,14,9,10,10,10,12,16,9,30,38,34],colorcol='판정',freeze='D3',desc='[왜 만들었나] 셀메이트에는 재고가 있는데 무신사에서 재고 입력 실수로 품절된 옵션, 컬러 묶음 상품(_7COL 등)에서 일부 컬러·사이즈만 품절된 옵션을 찾기 위해 무신사 옵션 1개를 셀메이트 옵션 1개와 1:1로 대조한 상세 시트입니다.\n[읽는 법] A열 판정 순서(문제 → 정상)로 정렬. 판정 근거는 맨 끝 "옵션 매칭" 열. 브랜드배송 0 + 무신사배송(MFS) 재고가 있으면 셀메이트 재고와 무관하게 정상입니다.')
# ---- 무신사 검색 노출 결과 (search/musinsa_brand_search.js 결과, 선택) ----
# 형식: 상품번호|검색결과수|순위|검색결과 중 브랜드 상품 수  (순위는 해당 상품번호가 '브랜드명 + 상품명' 검색 결과에
#       브랜드=어반드레스로 나왔을 때만 기록, 안 나오면 빈칸)
import os
HAS_SEARCH=os.path.exists(cfg.SEARCH_RESULT)
if HAS_SEARCH:
    SR=pd.DataFrame([x.strip().split('|')[:4] for x in open(cfg.SEARCH_RESULT,encoding='utf-8').read().replace('\n',';').split(';') if x.strip()],
                    columns=['상품번호','검색결과수','순위','브랜드결과수'])
    SR['검색결과수']=pd.to_numeric(SR['검색결과수'],errors='coerce').fillna(0).astype(int)
    SR['브랜드결과수']=pd.to_numeric(SR['브랜드결과수'],errors='coerce').fillna(0).astype(int)
    SR['순위']=pd.to_numeric(SR['순위'],errors='coerce')
    SR['검색여부']=np.where(SR['순위'].notna(),f'검색됨 ({cfg.BRAND_KR} 브랜드 확인)','검색 안 됨')
    SRm=SR.set_index('상품번호')
else:
    SRm=pd.DataFrame()
try: KW=pd.read_pickle(cfg.WORK+'kw.pkl').set_index('무신사 상품번호')['검색어']
except FileNotFoundError: KW=pd.Series(dtype=str)
def search_label(lid):
    lid=str(lid)
    if not HAS_SEARCH or lid not in SRm.index: return '미확인'
    x=SRm.loc[lid]
    return x['검색여부'] if pd.isna(x['순위']) else f"검색됨 ({x['검색결과수']}건 중 {int(x['순위'])}위)"
# 3) 상품별 한눈에 보기 (무신사 상품번호 1개 = 1줄)
import re as _re
def opt_text(d,val=None,fmt='{}'):
    """옵션들을 컬러별로 묶어 'BLACK: S(12)·M(5) | GREY: L(3)' 형태로"""
    if d.empty: return ''
    parts={}
    for _,x in d.iterrows():
        col=str(x['옵션컬러']).strip() or '-'; sz=str(x['옵션사이즈']).strip() or str(x['옵션'])
        if col=='-' and not str(x['옵션사이즈']).strip(): sz=str(x['옵션'])
        v='' if val is None or pd.isna(x[val]) else f"({fmt.format(int(x[val]))})"
        parts.setdefault(col,[]).append(f"{sz}{v}")
    return ' | '.join(f"{c}: {'·'.join(v)}" if c!='-' else '·'.join(v) for c,v in parts.items())
rows=[]
for lid,d in r.groupby('상품번호',sort=False):
    pb=d[d['판정'].isin(PROB)]; lim=d[d['판정']==PROB[1]]
    ov=d[d['판정']==OVR_RISK]; oth=d[d['판정'].isin(OTH)]; chk=d[d['판정']==CHK]
    sold=(d['총합 판매가능']==0)
    is_set=(d['판정']=='대조 제외(세트상품)').all()
    if is_set: res='대조 제외(세트상품)'
    elif len(pb): res='① 셀메이트 재고 있는 옵션 품절'
    elif len(oth): res='② 품절 옵션이 다른 옵션명·상품번호로 판매중'
    elif len(chk): res='③ 확인필요 옵션 있음'
    elif sold.any(): res='정상(품절 옵션은 셀메이트도 재고 없음)'
    else: res='정상(전 옵션 판매중)'
    codes=sorted(set(d['셀메이트상품코드'].dropna().astype(str)))
    sel_sum=d.drop_duplicates('셀메이트옵션코드')['셀메이트재고'].fillna(0).sum()
    rows.append(dict(점검결과=res,상품번호=lid,상품명=d['상품명'].iloc[0],스타일넘버=d['스타일넘버'].iloc[0],
        셀메이트상품코드=', '.join(codes[:8])+(' 외' if len(codes)>8 else ''),
        옵션수=len(d),판매중옵션=int((~sold).sum()),품절옵션=int(sold.sum()),
        문제옵션수=len(pb),문제재고=int(pb['셀메이트재고'].fillna(0).sum()),
        문제옵션=opt_text(pb,'셀메이트재고'),판매제한=opt_text(lim),
        초과판매=opt_text(ov,'국내출고/브랜드배송'),다른번호=opt_text(oth,'셀메이트재고'),
        확인필요=opt_text(chk),셀메이트재고합=int(sel_sum),무신사판매가능합=int(d['총합 판매가능'].sum()),
        MFS옵션=int((d['판정']=='정상 판매(무신사 풀필먼트 출고)').sum()),
        검색노출=search_label(lid),
        상품명대조=d['상품명대조'].iloc[0],링크='https://www.musinsa.com/products/'+str(lid)))
L=pd.DataFrame(rows)
RK={'① 셀메이트 재고 있는 옵션 품절':0,'② 품절 옵션이 다른 옵션명·상품번호로 판매중':1,'③ 확인필요 옵션 있음':2}
L=L[L['점검결과'].isin(RK)].copy()          # 정상 판매·정상 품절·세트 제외
L['_k']=L['점검결과'].map(RK)
L=L.sort_values(['_k','문제재고','품절옵션'],ascending=[True,False,False])
c2=['점검결과','상품번호','상품명','스타일넘버','셀메이트상품코드','옵션수','판매중옵션','품절옵션','문제옵션수','문제재고','문제옵션','판매제한','다른번호','확인필요','셀메이트재고합','무신사판매가능합','MFS옵션','검색노출','상품명대조','링크']
h2=['점검결과','무신사 상품번호','무신사 상품명','스타일넘버','연결된 셀메이트 상품코드','옵션 수','판매중 옵션','품절 옵션',
    '셀메이트 재고 있는데 품절인 옵션 수','그 옵션들 셀메이트 재고','품절 옵션 상세 (셀메이트 재고)','그중 판매 제한 옵션(재고 입력됐는데 판매가능 0)',
    '품절이지만 다른 옵션명·상품번호로 판매중 (셀메이트 재고)','확인필요 옵션',
    '셀메이트 재고 합계','무신사 판매가능 합계','MFS로 판매중인 옵션 수',f'무신사 검색 ("{cfg.BRAND_KR} + 상품명")','상품명 대조(스타일넘버 확인)','무신사 링크']
ws=sheet('상품별_한눈에보기',L,c2,h2,[30,11,38,13,24,7,8,8,11,11,48,30,30,30,10,10,9,30,30,38],freeze='D3',desc='[왜 만들었나] 옵션별로 나뉜 결과는 한 상품이 여러 줄이라 확인이 어려워서, 무신사 상품 1개 = 1줄로 묶고 조치가 필요한 상품만 남긴 시트입니다. 실제 작업은 이 시트부터 보시면 됩니다.\n[읽는 법] 정상 판매·정상 품절·초과판매 위험·세트상품은 제외. 품절 옵션은 "BLACK: S(12)·M(5)"처럼 컬러별로 묶었고 괄호 안은 셀메이트 재고. 상품명·스타일넘버로 Ctrl+F 검색하거나 A열 필터를 쓰세요.')
pf={'① 셀메이트 재고 있는 옵션 품절':'F8CBAD','② 품절 옵션이 다른 옵션명·상품번호로 판매중':'FFE699','③ 확인필요 옵션 있음':'E7E6E6'}
for i in range(3,len(L)+3):
    c=ws.cell(i,1); c.fill=PatternFill('solid',fgColor=pf.get(c.value,'FFFFFF'))
    for j in (11,12,13,14): ws.cell(i,j).alignment=Alignment(wrap_text=True,vertical='top')
wb.move_sheet('상품별_한눈에보기',offset=-(len(wb.sheetnames)-2))
if HAS_SEARCH:
    S=L[['상품번호','상품명','점검결과']].copy(); S['상품번호']=S['상품번호'].astype(str)
    S=S.join(SRm,on='상품번호'); S['검색여부']=S['검색여부'].fillna('미확인')
    S['검색어']=S['상품번호'].map(KW).map(lambda v: f'{cfg.BRAND_KR} {v}' if isinstance(v,str) else '')
    from urllib.parse import quote
    S['검색링크']='https://www.musinsa.com/search/goods?keyword='+S['검색어'].map(lambda v: quote(str(v)))
    S['상품링크']='https://www.musinsa.com/products/'+S['상품번호']
    def _note(row):
        if row['검색여부']!='검색 안 됨': return ''
        if '유니섹스' in str(row['검색어']):
            return "무신사 웹 검색은 검색어에 '유니섹스'가 들어가면 결과 화면이 비어 보임(사이트 전체 현상, '유니섹스'만 검색해도 동일). '유니섹스'를 빼고 검색하면 노출됨 → 상품명에서 '유니섹스' 제외 검토"
        return '검색 결과 화면에 해당 상품이 없음 → 무신사 상품 노출 상태 확인 필요'
    S['비고']=S.apply(_note,axis=1)
    S['_k']=S['검색여부'].map(lambda v: 0 if v=='검색 안 됨' else (1 if v=='미확인' else 2))
    S=S.sort_values(['_k','순위'],ascending=[True,False])
    c6=['검색여부','상품번호','상품명','검색어','검색결과수','브랜드결과수','순위','비고','점검결과','검색링크','상품링크']
    h6=['검색 결과','무신사 상품번호','무신사 상품명','검색어 (브랜드명 + 상품명, 컬러·COL 제외)','검색 결과 수 (페이지 표시 총 N건)',f'화면에 보인 {cfg.BRAND_KR} 상품 수','화면 순위','비고','옵션 점검결과','무신사 검색 링크','무신사 상품 링크']
    D_S=(f'[왜 만들었나] 무신사에 정상 등록돼 있어도 고객이 검색했을 때 우리 브랜드 상품으로 나오지 않는 경우를 찾기 위한 시트입니다.\n'
         f'[읽는 법] 무신사 검색(비로그인)에 "{cfg.BRAND_KR} + 상품명(컬러명·_3COL 등 컬러 묶음 표시 제외)"을 넣어, 실제 검색 결과 화면에 해당 상품번호가 브랜드 "{cfg.BRAND_KR}"로 나오면 "검색됨". 순위는 화면에서의 위치입니다. "검색 안 됨"은 비고에 원인을 적었습니다.')
    ws=sheet('무신사_검색노출',S,c6,h6,[26,11,40,40,11,11,9,60,30,40,38],freeze='D3',desc=D_S)
    for i in range(3,len(S)+3): ws.cell(i,8).alignment=Alignment(wrap_text=True,vertical='top')
    for i in range(3,len(S)+3):
        c=ws.cell(i,1); c.fill=PatternFill('solid',fgColor={'검색 안 됨':'F8CBAD','미확인':'E7E6E6'}.get(c.value,'C6EFCE'))
    wb.move_sheet('무신사_검색노출',offset=-(len(wb.sheetnames)-3))
# 4) 무신사에 옵션 없음
mi=miss.copy(); mi['링크']='https://www.musinsa.com/products/'+mi['무신사상품번호'].astype(str)
mi=mi.sort_values('재고',ascending=False)
c3=['옵션코드','상품명','옵션명','재고','판매여부','무신사상품번호','무신사상품명','링크']
h3=['셀메이트 옵션코드','셀메이트 상품명','셀메이트 옵션명','셀메이트 재고','셀메이트 판매여부','연결된 무신사 상품번호','연결된 무신사 상품명','무신사 링크']
sheet('무신사에_옵션없음',mi,c3,h3,[26,40,16,9,9,11,38,38],desc='[왜 만들었나] 셀메이트에는 재고가 있는 옵션인데, 연결된 무신사 상품에 그 옵션(컬러·사이즈) 자체가 없어서 고객이 고를 수 없는 경우를 찾기 위한 시트입니다. 품절이 아니라 옵션 누락이라 옵션별_점검에는 나오지 않습니다.\n[읽는 법] 셀메이트 재고 많은 순. 무신사 상품에 옵션을 추가할지 확인하세요.')
# 6) 상품명 대조 (무신사 판매중 → 셀메이트)
R=pd.read_pickle(cfg.WORK+'namechk.pkl')
R=R[~R['무신사스타일넘버'].map(other_brand)&~R['대조결과'].str.startswith('정상')].copy()   # 정상 일치 상품은 기록하지 않음
NOTE=cfg.LISTING_NOTES
R['메모']=R['무신사상품번호'].map(NOTE).fillna('')
R['_k']=R['대조결과'].map(lambda v: 0 if v.startswith('확인') else (1 if v.startswith('셀메이트에서') else (3 if v.startswith('세트') else 2)))
R['무신사재고']=pd.to_numeric(R['무신사재고'].astype(str).str.replace(',',''),errors='coerce')
R=R.sort_values(['_k','대조결과','무신사재고'],ascending=[True,True,False])
c4=['대조결과','무신사상품번호','무신사상품명','무신사스타일넘버','비교용상품명','스타일넘버연결_셀메이트','스타일넘버연결_상품명','상품명일치_셀메이트','상품명일치_상품명','무신사재고','메모','링크']
h4=['대조 결과','무신사 상품번호','무신사 상품명','무신사 스타일넘버','비교용 상품명(컬러·COL 제외)','스타일넘버로 연결된 셀메이트 상품코드','그 셀메이트 상품명','상품명으로 찾은 셀메이트 상품코드','그 셀메이트 상품명','무신사 판매가능 재고','메모','무신사 링크']
ws=sheet('상품명_대조',R,c4,h4,[40,11,40,14,30,30,36,30,36,10,46,38],freeze='C3',desc='[왜 만들었나] 스타일넘버만으로는 놓치는 경우가 있어, 무신사 판매중 상품을 스타일넘버와 상품명(컬러명·컬러 묶음 표시 제외)으로 각각 셀메이트와 교차 확인한 시트입니다. 스타일넘버·상품명이 모두 맞는 정상 상품은 기록하지 않고, 확인이 필요한 상품만 남겼습니다.\n[읽는 법] "확인:"으로 시작하면 무신사 스타일넘버 오입력 의심 → 무신사에서 스타일넘버를 고쳐주세요. 한글/영문으로 다르게 등록된 상품은 이름 비교가 안 돼 메모에 직접 적었습니다.')
NF={'확인':'FFE699','셀메이트에서':'E7E6E6','정상':'C6EFCE','세트':'FFFFFF'}
for i in range(3,len(R)+3):
    c=ws.cell(i,1); c.fill=PatternFill('solid',fgColor=next(v for k,v in NF.items() if str(c.value).startswith(k)))
# 1) 요약
O="'옵션별_점검'"; P="'상품별_한눈에보기'"; ws=ws0; b=Font(name=F,bold=True,size=10)
ws['A1']='무신사 옵션 재고 점검 — 어반드레스 (판매중 상품, 젠더스트·비바셔스·미발매 제외)'; ws['A1'].font=Font(name=F,bold=True,size=14)
ws['A3']=('[이 시트] 파일 전체 결과를 숫자로 모은 시트입니다. 다른 시트를 수식으로 집계해서 다른 시트를 고치면 숫자도 같이 바뀝니다. '
          '작업 순서: 상품별_한눈에보기 → 상품명_대조(스타일넘버 수정) → 무신사에_옵션없음.')
ws['A3'].font=Font(name=F,size=10,color='1F3864')
ws['A2']=f'기준: {cfg.fname(cfg.MUSINSA_OPTION)} × {cfg.fname(cfg.SELLMATE_STOCK)} × {cfg.fname(cfg.MUSINSA_SELLING)} · 작성 {cfg.TODAY}'; ws['A2'].font=Font(name=F,size=9,color='808080')
def hdr(row,titles):
    for j,t in enumerate(titles,1):
        x=ws.cell(row,j,t); x.font=HF; x.fill=H; x.alignment=Alignment(horizontal='center',wrap_text=True)
hdr(4,['옵션 판정','옵션 수','셀메이트 재고 합계','무신사 상품 수'])
for k,j in enumerate(ORDER,5):
    ws.cell(k,1,j).font=b
    if j in FILL: ws.cell(k,1).fill=PatternFill('solid',fgColor=FILL[j])
    ws.cell(k,2,f'=COUNTIFS({O}!$A:$A,$A{k})'); ws.cell(k,3,f'=SUMIFS({O}!$H:$H,{O}!$A:$A,$A{k})')
ws.cell(k+1,1,'합계').font=b; ws.cell(k+1,2,f'=SUM(B5:B{k})').font=b; ws.cell(k+1,3,f'=SUM(C5:C{k})').font=b
# 상품 수(문제 판정만 유니크) - 파이썬 계산값을 근거와 함께 기재
for kk,j in enumerate(ORDER,5):
    ws.cell(kk,4,int(r[r['판정']==j]['상품번호'].nunique()))
r0=k+3
hdr(r0,['상품별_한눈에보기 (점검 필요 상품만)','상품 수'])
for i,t in enumerate(list(RK),r0+1):
    ws.cell(i,1,t).font=b; ws.cell(i,1).fill=PatternFill('solid',fgColor=pf[t]); ws.cell(i,2,f'=COUNTIFS({P}!$A:$A,$A{i})')
r1=r0+len(RK)+2
ws.cell(r1,1,'셀메이트 재고 있는데 무신사에 옵션 자체가 없음').font=b
ws.cell(r1,2,"=COUNTIFS('무신사에_옵션없음'!$D:$D,\">0\")"); ws.cell(r1,3,"=SUMIFS('무신사에_옵션없음'!$D:$D,'무신사에_옵션없음'!$D:$D,\">0\")")
ws.cell(r1+1,1,'상품명 대조: 스타일넘버 확인 필요한 무신사 상품').font=b
ws.cell(r1+1,2,"=COUNTIFS('상품명_대조'!$A:$A,\"확인*\")")
r2=r1+3
if HAS_SEARCH:
    hdr(r2,[f'무신사 검색 ("{cfg.BRAND_KR} + 상품명", 점검 필요 상품 대상)','상품 수'])
    for i,t in enumerate([f'검색됨 ({cfg.BRAND_KR} 브랜드 확인)','검색 안 됨','미확인'],r2+1):
        ws.cell(i,1,t).font=b; ws.cell(i,2,f"=COUNTIFS('무신사_검색노출'!$A:$A,$A{i})")
r1=r1+6
notes=['판정 기준',
'· 무신사 품절(셀메이트 재고 있음): 셀메이트에 재고가 있는데 무신사 옵션 판매가능 재고가 0 → 재고 입력 누락 가능성이 가장 큰 옵션입니다.',
'· 판매 제한: 무신사 브랜드배송 재고는 입력돼 있는데 총합 판매가능이 0 → 허용판매수량 등 판매 제한 설정을 확인해 주세요.',
'· 품절(다른 상품번호에서는 판매중): 같은 셀메이트 옵션이 다른 무신사 상품에서는 팔리고 있는 경우라 우선순위는 낮습니다(중복 등록 상품 정리 대상).',
'· 품절(같은 상품 안 다른 옵션명으로 판매중): 같은 무신사 상품 안에서 "(Short) PINK S"처럼 다른 옵션명으로 같은 셀메이트 옵션을 팔고 있어 고객은 구매 가능합니다(품절로 남은 중복 옵션 정리 대상).',
'· 초과판매 위험: 셀메이트 재고 0인데 무신사 브랜드배송 재고가 남아있는 옵션입니다.',
'· 정상 판매(무신사 풀필먼트 출고): 브랜드배송 재고는 0이고 무신사배송(MFS) 재고로 판매 중인 옵션 → 셀메이트 재고가 없어도 정상입니다.',
'· 스타일넘버가 잘못 입력된 무신사 상품은 상품명으로 찾은 셀메이트 상품 기준으로 옵션을 다시 연결했습니다(옵션 매칭 열에 "재연결" 표시).',
'· 확인필요: 컬러 없는 잔여 옵션(_5COL 상품의 "S", "FREE" 등), (기모)·(LONG) 같은 변형 옵션, 셀메이트에 해당 컬러·사이즈가 없는 옵션 등입니다. O열 "옵션 매칭"에 사유가 있습니다.',
'· 매칭: 무신사 옵션(컬러^사이즈)을 스타일넘버·상품명으로 찾은 셀메이트 상품의 컬러(상품명 끝/옵션명)와 사이즈(옵션명)에 맞췄습니다. 블랙=BLACK 등 한/영 컬러명, S(1)=1(S) 같은 표기 차이를 감안합니다.',
'· 셀메이트 재고는 셀메이트 현재재고, 무신사 재고는 무신사 총합 판매가능(브랜드배송+무신사배송, 허용판매수량 적용) 기준입니다.',
'· 젠더스트(ZD)·비바셔스(VV) 상품은 모든 시트에서 제외했습니다.',
'· 셀메이트 옵션코드가 비어 있는 상품은 "상품코드 / 옵션명 (옵션코드 없음)"으로 표시했습니다.']
for i,t in enumerate(notes,r1+6): ws.cell(i,1,t).font=b if t=='판정 기준' else Font(name=F,size=10)
ws.cell(4,4).comment=None
ws.cell(r1+4,1,'※ D열 "무신사 상품 수"는 판정별 고유 상품번호 수(작성 시점 계산값)입니다.').font=Font(name=F,size=9,color='808080')
for c,w in zip('ABCD',[46,10,16,13]): ws.column_dimensions[c].width=w
for row in ws.iter_rows():
    for x in row:
        if x.value is not None and x.font.name!=F: x.font=Font(name=F,size=10)
        if isinstance(x.value,(int,str)) and x.column>1 and (isinstance(x.value,int) or str(x.value).startswith('=')): x.number_format='#,##0'
wb.save(cfg.OUT+f'무신사_옵션재고점검_{cfg.TODAY}.xlsx')
print(L['점검결과'].value_counts())
