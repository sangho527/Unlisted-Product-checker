import config as cfg
import pandas as pd, numpy as np
r=pd.read_pickle(cfg.WORK+'optres.pkl'); miss=pd.read_pickle(cfg.WORK+'optmiss.pkl')
r['셀메이트재고']=pd.to_numeric(r['셀메이트재고'])
ok=r['옵션매칭'].str.startswith('일치')
# 변형 옵션((기모), (LONG), (short) 등)이 기본 옵션과 같은 셀메이트 옵션에 붙은 경우 → 별도 확인
r['_tag']=r['옵션'].astype(str).str.contains(r'\(',regex=True)&~r['옵션매칭'].str.contains('재연결')
plain=set(r[ok&~r['_tag']]['셀메이트옵션코드'])
r.loc[ok&r['_tag']&r['셀메이트옵션코드'].isin(plain),'옵션매칭']='변형 옵션(기모·롱 등) - 셀메이트에 별도 옵션 없음'
ok=r['옵션매칭'].str.startswith('일치')
# 같은 무신사 상품 안에서 여러 옵션이 같은 셀메이트 옵션에 연결되면, 대표 1개(판매가능 재고 많은 것→컬러 표기 있는 것)만 남김
r['_rk']=list(zip(-r['총합 판매가능'],r['옵션컬러'].eq('')))
d=r[ok].sort_values(['상품번호','셀메이트옵션코드','총합 판매가능','옵션컬러'],ascending=[True,True,False,False])
extra=d[d.duplicated(['상품번호','셀메이트옵션코드'])].index
r.loc[extra,'옵션매칭']='같은 상품 안 중복 옵션(같은 셀메이트 옵션)'
ok=r['옵션매칭'].str.startswith('일치')
sel=r['셀메이트재고'].fillna(0); tot=r['총합 판매가능']; br=r['국내출고/브랜드배송']; mfs=r['국내출고/무신사배송']+r['일본출고/무신사배송']
J=np.select([
  ok&(sel>0)&(tot==0)&(br>0),
  ok&(sel>0)&(tot==0),
  ok&(sel==0)&(br>0),
  ok&(br==0)&(mfs>0),
  ~ok&r['옵션매칭'].str.startswith('세트'),
  ~ok,
  (tot>0),
  ],[
  '판매 제한(재고 입력됐는데 판매가능 0)',
  '무신사 품절(셀메이트 재고 있음)',
  '초과판매 위험(셀메이트 재고 없음)',
  '정상 판매(무신사 풀필먼트 출고)',
  '대조 제외(세트상품)',
  '확인필요(옵션 매칭 안 됨)',
  '정상 판매',
  ],'정상 품절(양쪽 재고 없음)')
r['판정']=J
# 같은 셀메이트 옵션이 다른 무신사 상품번호에서는 판매중인지
okr=r[ok]
sell=okr[okr['총합 판매가능']>0].groupby('셀메이트옵션코드')['상품번호'].apply(lambda v: ', '.join(sorted(set(v))))
def other(row):
    if not isinstance(row['셀메이트옵션코드'],str) or row['셀메이트옵션코드'] not in sell.index: return ''
    o=[x for x in sell[row['셀메이트옵션코드']].split(', ') if x!=row['상품번호']]
    return ', '.join(o)
r['다른상품번호에서판매중']=r.apply(other,axis=1)
r.loc[r['판정'].isin(['무신사 품절(셀메이트 재고 있음)','판매 제한(재고 입력됐는데 판매가능 0)'])&r['다른상품번호에서판매중'].ne(''),'판정']='품절(다른 상품번호에서는 판매중)'
# 상품(무신사 상품번호) 단위 상태
g=r.groupby('상품번호')
st=pd.DataFrame({'옵션수':g.size(),'판매가능옵션':g['총합 판매가능'].apply(lambda v:(v>0).sum()),
   '품절옵션':g['총합 판매가능'].apply(lambda v:(v==0).sum()),
   '품절인데셀메이트재고있음':g['판정'].apply(lambda v: v.isin(['무신사 품절(셀메이트 재고 있음)','판매 제한(재고 입력됐는데 판매가능 0)']).sum())})
st['상품상태']=np.select([st['판매가능옵션']==0, st['품절옵션']>0],['전체 품절(판매중 표시)','부분 품절'],'전 옵션 판매중')
r=r.join(st['상품상태'],on='상품번호')
r.to_pickle(cfg.WORK+'judged.pkl'); st.to_pickle(cfg.WORK+'liststat.pkl')
print(r['판정'].value_counts()); print(st['상품상태'].value_counts())
x=r[r['판정'].isin(['무신사 품절(셀메이트 재고 있음)','판매 제한(재고 입력됐는데 판매가능 0)'])]
print('문제 옵션',len(x),'상품',x['상품번호'].nunique(),'셀메이트재고',x['셀메이트재고'].sum())
print(pd.crosstab(x['판정'],x['상품상태']))
print('miss with stock',(miss['재고']>0).sum(), miss[miss['재고']>0]['재고'].sum())
print(x.groupby(['상품명']).agg(n=('옵션','size'),q=('셀메이트재고','sum')).sort_values('q',ascending=False).head(15))
