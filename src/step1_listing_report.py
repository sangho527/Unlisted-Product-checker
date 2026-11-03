import config as cfg
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
o=pd.read_pickle(cfg.WORK+'out.pkl'); m=pd.read_pickle(cfg.WORK+'m.pkl')
u=o[o['브랜드']==cfg.BRAND_KR].copy()
u['_a']=u['조치필요'].eq('').astype(int)
u['_j']=u['판정'].map({'등록·미노출':0,'미등록':1,'등록·노출':2})
u['_s']=u['판매여부'].ne('판매').astype(int)
u=u.sort_values(['_j','_a','_s','재고합계'],ascending=[True,True,True,False])
cols=['상품코드','상품명','판매여부','옵션수','재고합계','입고예정일','판정','무신사상태','조치필요',
      '무신사상품번호','무신사상품명','무신사재고','다른매칭상태','무신사링크','매칭방식']
hdr={'판매여부':'셀메이트 판매여부','재고합계':'셀메이트 재고','판정':'무신사 판정'}
F='Arial'; thin=Side(style='thin',color='D9D9D9')
H=PatternFill('solid',fgColor='1F3864'); HF=Font(name=F,bold=True,color='FFFFFF',size=10)
fills={'미등록':'F8CBAD','등록·미노출':'FFE699','등록·노출':'C6EFCE'}
wb=Workbook()
ws0=wb.active; ws0.title='요약'
def desc_height(desc,widths):
    """병합된 설명 셀에 줄바꿈된 글자가 다 보이도록 행 높이 계산 (한글은 폭 1.8로 계산)"""
    import math
    total=max(sum(widths),20)
    lines=0
    for ln in desc.split('\n'):
        units=sum(1.8 if ord(ch)>0x1100 else 1.0 for ch in ln)
        lines+=max(1,math.ceil(units*1.05/total))
    return max(30,15*lines+10)
def sheet(name,df,cols,widths,desc=''):
    """1행: 시트 설명 · 2행: 헤더 · 3행~: 데이터"""
    ws=wb.create_sheet(name)
    ws.cell(1,1,desc).font=Font(name=F,size=10,color='1F3864')
    ws.cell(1,1).alignment=Alignment(wrap_text=True,vertical='top'); ws.cell(1,1).fill=PatternFill('solid',fgColor='EEF3FA')
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=len(cols))
    for j,c in enumerate(cols,1):
        x=ws.cell(2,j,hdr.get(c,c)); x.font=HF; x.fill=H; x.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    for i,row in enumerate(df[cols].itertuples(index=False),3):
        for j,v in enumerate(row,1):
            if pd.isna(v): v=None
            if cols[j-1] in ('옵션수','재고합계') and v is not None: v=int(v)
            if cols[j-1] in ('무신사판매가','무신사재고') and v: 
                try: v=int(str(v).replace(',',''))
                except: pass
            x=ws.cell(i,j,v); x.font=Font(name=F,size=10)
            if cols[j-1]=='무신사링크' and v: x.hyperlink=v; x.font=Font(name=F,size=10,color='0563C1',underline='single')
            if cols[j-1] in ('판정','무신사 판정') and v in fills: x.fill=PatternFill('solid',fgColor=fills[v])
            if cols[j-1]=='조치필요' and v: x.font=Font(name=F,size=10,bold=True,color='C00000')
            if isinstance(v,int): x.number_format='#,##0'
    for j,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(j)].width=w
    ws.row_dimensions[1].height=desc_height(desc,widths)
    ws.freeze_panes='C3'; ws.auto_filter.ref=f"A2:{get_column_letter(len(cols))}{len(df)+2}"
    ws.row_dimensions[2].height=30
    return ws
sheet('무신사_어반드레스_점검',u,cols,[15,42,10,7,9,11,11,10,22,11,42,9,24,40,30],desc='[왜 만들었나] 셀메이트에 있는 우리 브랜드 상품이 무신사에 등록돼 있는지, 등록돼 있다면 판매중인지(품절·판매중지 등으로 묶이지 않았는지) 찾기 위한 시트입니다. 셀메이트 상품(컬러) 1개 = 1줄.\n[읽는 법] 무신사 판정: 등록·미노출(품절·판매중지 등) → 미등록(매칭 없음·삭제) → 등록·노출(판매중) 순서. "조치필요"는 셀메이트 재고가 있는데(판매·이월 상관없음) 무신사에서 안 팔리는 상품. 매칭 근거는 맨 끝 "매칭방식" 열(스타일넘버 → 상품명 순으로 대조).')
mo=m[(~m['셀메이트매칭'])&m['무신사'].isin(['판매중','품절','등록중','검수반려'])].copy()
mo['무신사링크']='https://www.musinsa.com/products/'+mo['상품 번호']
mo=mo.sort_values(['pri','등록 일시'],ascending=[True,False])
c2=['무신사','상품 번호','상품명','스타일넘버','판매가(원)','판매 가능 재고','등록 일시','무신사링크']
sheet('무신사에만_있는상품',mo,c2,[10,11,55,18,11,11,16,40],desc='[왜 만들었나] "셀메이트엔 없는데 무신사에만 있는 상품"을 확인하기 위한 시트입니다. 무신사 판매중·품절 등 상품 중 셀메이트 상품과 스타일넘버·상품명 어느 쪽으로도 연결되지 않은 상품입니다.\n[읽는 법] 대부분 셀메이트엔 단품으로만 있는 [SET]·[패키지] 상품입니다. 세트가 아닌데 여기 있으면 셀메이트 미등록이거나 무신사 스타일넘버가 틀린 경우라 주문·재고 연동을 확인하세요.').freeze_panes='A3'
# 요약 (formulas)
D="'무신사_어반드레스_점검'"
ws=ws0; b=Font(name=F,bold=True,size=10); n=Font(name=F,size=10)
ws['A1']='무신사 등록 현황 점검 — 어반드레스'; ws['A1'].font=Font(name=F,bold=True,size=14)
ws['A3']='[이 시트] 무신사 등록 현황을 숫자로 모은 시트입니다. 다른 시트를 수식으로 집계합니다.'; ws['A3'].font=Font(name=F,size=10,color='1F3864')
ws['A2']=f'기준: {cfg.fname(cfg.SELLMATE_STOCK)} × {cfg.fname(cfg.MUSINSA_ALL)} · 상품코드(컬러) 단위 · 작성 {cfg.TODAY}'; ws['A2'].font=Font(name=F,size=9,color='808080')
r=4
for j,t in enumerate(['무신사 판정','셀메이트 판매','셀메이트 이월/기타','합계'],1):
    x=ws.cell(r,j,t); x.font=HF; x.fill=H; x.alignment=Alignment(horizontal='center')
for k,j in enumerate(['등록·노출','등록·미노출','미등록'],r+1):
    ws.cell(k,1,j).font=b; ws.cell(k,1).fill=PatternFill('solid',fgColor=fills[j])
    ws.cell(k,2,f'=COUNTIFS({D}!$G:$G,$A{k},{D}!$C:$C,"판매")')
    ws.cell(k,3,f'=COUNTIFS({D}!$G:$G,$A{k})-B{k}')
    ws.cell(k,4,f'=B{k}+C{k}')
ws.cell(r+4,1,'합계').font=b
for c in 'BCD': ws[f'{c}{r+4}']=f'=SUM({c}{r+1}:{c}{r+3})'; ws[f'{c}{r+4}'].font=b
r=10
for j,t in enumerate(['무신사 상태 (셀메이트 판매 상품)','상품 수'],1):
    x=ws.cell(r,j,t); x.font=HF; x.fill=H
for k,st in enumerate(['판매중','품절','판매중지','등록중','검수반려','삭제','-'],r+1):
    ws.cell(k,1,st if st!='-' else '-  (매칭 없음)')
    ws.cell(k,2,f'=COUNTIFS({D}!$H:$H,"{st}",{D}!$C:$C,"판매")')
r=19
for j,t in enumerate(['우선 조치 대상 (셀메이트 재고 있음, 판매·이월 무관)','상품 수','재고 합계'],1):
    x=ws.cell(r,j,t); x.font=HF; x.fill=H
for k,st in enumerate(['미등록(재고있음)','판매중지 등(재고있음)','무신사 품절표시(재고있음)'],r+1):
    ws.cell(k,1,st).font=Font(name=F,size=10,bold=True,color='C00000')
    ws.cell(k,2,f'=COUNTIFS({D}!$I:$I,$A{k})'); ws.cell(k,3,f'=SUMIFS({D}!$E:$E,{D}!$I:$I,$A{k})')
ws.cell(24,1,'무신사에만 있는 상품(판매중/품절 등, 셀메이트 매칭 안 됨)').font=b
ws['B24']="=COUNTA('무신사에만_있는상품'!B:B)-1"
notes=['읽는 법',
 '· 브랜드 구분: 상품코드 AD·AV로 시작(구코드 yySS/FW+AD 포함) = 어반드레스. 무신사 avan 스토어에 AD·AV 코드가 함께 등록돼 있어 둘 다 어반드레스로 봤습니다.',
 '· 매칭 순서: ① 스타일넘버(상품코드 전체 / 앞 8자리 기본코드) ② 상품명 — 스타일넘버가 달라도 상품명이 같으면 같은 상품으로 봅니다. 한 상품에 여러 무신사 상품이 걸리면 가장 좋은 상태(판매중>품절>…>삭제)를 표시합니다.',
 '· 상품명 비교 시 [SET]·[패키지] 머리말, "_3COL / - 4COL / 4 COLOR" 같은 컬러 묶음 표시, 끝의 컬러명(BLACK, LIGHT BLUE, 블랙, 다크베이지 등)을 떼고 비교합니다.',
 '· 컬러 처리: 무신사가 컬러 묶음 상품(_3COL 등)이면 모든 컬러 등록으로 보고, 컬러별 상품이면 셀메이트 컬러와 같은 컬러가 있어야 일치로 봅니다(다른 컬러만 있으면 미등록 + 참고 표시).',
 '· "유사 상품명(확인필요)"은 오타·단어 차이가 있는 이름을 맞춘 것이라 한 번 확인해 주세요.',
 '· 무신사 상태가 "삭제"인 것은 미등록으로 분류했습니다.',
 '· 통합상품(예: _3COL)은 스타일 단위로 판정하므로, 통합상품 안에서 특정 컬러 옵션만 빠진 경우는 잡히지 않습니다.',
 '· "무신사 품절표시(재고있음)"은 셀메이트 재고가 있는데 무신사는 품절 → 재고 연동 확인이 필요한 상품입니다.']
for i,t in enumerate(notes,27):
    ws.cell(i,1,t).font=b if i==27 else Font(name=F,size=10)
for c,w in zip('ABCD',[44,14,18,10]): ws.column_dimensions[c].width=w
for row in ws.iter_rows():
    for x in row:
        if x.font.name!=F: x.font=Font(name=F,size=10)
        if x.column>1 and isinstance(x.value,str) and x.value.startswith('='): x.number_format='#,##0'
wb.save(cfg.OUT+f'무신사_어반드레스_등록점검_{cfg.TODAY}.xlsx')
