import config as cfg
import pandas as pd, re, json
exec(open('src/step3_option_match.py',encoding='utf-8').read().split('# ---- 셀메이트 옵션 ----')[0])
def clean(n):
    n=str(n).strip()
    n=re.sub(r'^\s*(\[[^\]]*\]\s*)+','',n)
    n=re.sub(r'[\s_\-]*\(?\d+\s*COL(OR)?S?\)?(\(N\))?\s*$','',n,flags=re.I)
    n=re.sub(r'\s+-\s*$','',n)
    t=n.split()
    mixed=any(re.search(r'[a-z가-힣]',w) for w in t)
    while len(t)>2:
        w=t[-1]; core=re.sub(r'\(.*?\)','',w)
        if re.fullmatch(r'[가-힣]+',core) and is_color(core): t.pop(); continue
        if re.fullmatch(r'[A-Z][A-Z0-9*]*',core) and core not in NOTCOL and (mixed or core in EN): t.pop(); continue
        break
    return re.sub(r'[_]+',' ',' '.join(t)).strip(' -_')
d=pd.read_excel(cfg.OUT+f'무신사_옵션재고점검_{cfg.TODAY}.xlsx','상품별_한눈에보기',dtype=str)
d['검색어']=d['무신사 상품명'].map(clean)
d[['무신사 상품번호','무신사 상품명','검색어']].to_pickle(cfg.WORK+'kw.pkl')
json.dump([[a,b] for a,b in zip(d['무신사 상품번호'],d['검색어'])],open(cfg.WORK+'kw.json','w',encoding='utf-8'),ensure_ascii=False)
# 브라우저 콘솔에 붙여넣을 검색 스크립트 생성 (키워드 포함)
tpl=open('search/musinsa_brand_search.js',encoding='utf-8').read()
kwjs=json.dumps([[a,b] for a,b in zip(d['무신사 상품번호'],d['검색어'])],ensure_ascii=False)
tpl=tpl.replace('const BRAND = "어반드레스";',f'const BRAND = {json.dumps(cfg.BRAND_KR,ensure_ascii=False)};').replace('const KEYWORDS = [];',f'const KEYWORDS = {kwjs};')
open(cfg.OUT+'musinsa_brand_search_ready.js','w',encoding='utf-8').write(tpl)
print(f'검색어 {len(d)}개 → {cfg.OUT}musinsa_brand_search_ready.js')
