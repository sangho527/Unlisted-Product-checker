"""
29CM ↔ 셀메이트 상품명·옵션 정리 규칙
- 상품명: 컬러 묶음 표기(- 3COL, _3Colors, 7 COLOR)·끝 컬러명 제거, 괄호 속 다른 언어 이름도 대조 키로 사용
- 한글 상품명 ↔ 영문 상품명: 의류 용어 사전 + 발음 골격 비교
- 옵션: [색상]/[사이즈]/[컬러] 같은 꼬리표 제거, 컬러·사이즈 분리, 한글 컬러 → 영문
"""
import re, difflib
KO={'블랙':'BLACK','화이트':'WHITE','아이보리':'IVORY','베이지':'BEIGE','차콜':'CHARCOAL','그레이':'GREY','회색':'GREY','네이비':'NAVY',
'블루':'BLUE','스카이블루':'SKYBLUE','소라':'SORA','카키':'KHAKI','브라운':'BROWN','핑크':'PINK','레드':'RED','퍼플':'PURPLE','올리브':'OLIVE',
'크림':'CREAM','다크베이지':'DARKBEIGE','다크네이비':'DARKNAVY','다크그레이':'DARKGREY','연베이지':'LIGHTBEIGE','연그레이':'LIGHTGREY',
'진청':'DEEPBLUE','연청':'LIGHTBLUE','중청':'MIDBLUE','흑청':'BLACKDENIM','멜란지':'MELANGE','오트밀':'OATMEAL','모카':'MOCHA','카멜':'CAMEL',
'그린':'GREEN','옐로우':'YELLOW','민트':'MINT','라벤더':'LAVENDER','와인':'WINE','버건디':'BURGUNDY','실버':'SILVER','골드':'GOLD',
'오렌지':'ORANGE','인디고':'INDIGO','딥블루':'DEEPBLUE','라이트블루':'LIGHTBLUE','라이트그레이':'LIGHTGREY','머스타드':'MUSTARD','코코아':'COCOA',
'라이트':'LIGHT','다크':'DARK','딥':'DEEP','피그먼트':'PIGMENT','워싱':'WASHING','스카이':'SKY','페일':'PALE','파스텔':'PASTEL','빈티지':'VINTAGE','연':'LIGHT','진':'DEEP',
'흑':'BLACK','백':'WHITE','청':'BLUE','차':'CHARCOAL','크림':'CREAM','버터':'BUTTER','샌드':'SAND','토프':'TAUPE','라임':'LIME','스틸':'STEEL'}
EN={'BLACK','WHITE','IVORY','BEIGE','CHARCOAL','GREY','GRAY','NAVY','BLUE','SKY','KHAKI','BROWN','PINK','RED','PURPLE','OLIVE','CREAM','LIGHT','DARK','DEEP',
'PALE','PIGMENT','WASHING','WASHED','MELANGE','OATMEAL','MOCHA','MOCA','CAMEL','GREEN','YELLOW','MINT','LAVENDER','WINE','SILVER','GOLD','ORANGE','INDIGO',
'BUTTER','SORA','SOLID','BURGUNDY','MUSTARD','COCOA','SAND','TAUPE','LIME','STEEL','SNOW','MIDDLE','MID','DENIM','BLUEBLACK','MULTI','LEOPARD','CHECK','STRIPE','SUEDE','STITCH','STONE','LIGHTBLUE','ASH','COTTON','VINTAGE','CAMO','SALT','INDIE','DUSTY','POWDER','PALEPINK','POWDERBLUE','LEOPARD','BLUEGREEN'}
NOTCOL={'MTM','CD','OPS','ZIP-UP','T-SHIRT','AD','AV','T','SET','UNISEX','HOODIE','HOODY','SHIRT','PANTS','TOP','CROP','TEE','SKIRT','KNIT','JACKET','BAG','CAP','ZIP','UP','ZIPUP','DRESS','SHORTS','BLOUSE','VEST','COAT','LONG','SHORT','SLEEVE','MINI','WIDE','DENIM','SUEDE','CHECK','STRIPE'}
COLT=re.compile(r'[\s_\-]*[\(\[]?\d+\s*COL(OR)?S?[\)\]]?(\(N\))?\s*$',re.I)
def base_clean(n):
    n=str(n).strip()
    n=re.sub(r'^\s*(\[[^\]]*\]\s*)+','',n)
    n=re.sub(r'\((N|\d)\)','',n)
    n=COLT.sub('',n); n=re.sub(r'\s*-\s*Z\s*$','',n); n=COLT.sub('',n)
    return n.strip(' -_')
def is_ko_color(w):
    c=w
    for k in sorted(KO,key=len,reverse=True): c=c.replace(k,'#')
    return len(w)>=2 and set(c)<={'#'}
def strip_color(n):
    """returns (core name, color tokens string)"""
    t=re.sub(r'[_/]',' ',n).replace(' - ',' ').split()
    mixed=any(re.search(r'[a-z가-힣]',w) for w in t)
    col=[]
    while len(t)>1:
        w=t[-1]; core=re.sub(r'\(.*?\)','',w).strip('-')
        if core=='' : t.pop(); continue
        if re.fullmatch(r'[가-힣]+',core) and is_ko_color(core): col.insert(0,core); t.pop(); continue
        U=core.upper()
        if re.fullmatch(r'[A-Z][A-Z0-9]*',core) and ((mixed and U not in NOTCOL) or U in EN) : col.insert(0,core); t.pop(); continue
        break
    return ' '.join(t), ' '.join(col)
def key(n): return re.sub(r'[^0-9a-z가-힣]','',str(n).lower())
def names(n):
    """대조용 키 목록: 본 이름 + 괄호 속 다른 언어 이름"""
    b=base_clean(n)
    out=[]
    paren=re.findall(r'\(([^()]*[A-Za-z가-힣][^()]*)\)',b)
    main=re.sub(r'\([^()]*\)','',b).strip()
    for x in [main]+paren:
        bx=base_clean(x); core,_=strip_color(bx)
        for k in (key(core),key(bx)):
            if len(k)>=3 and k not in out: out.append(k)
    return out, strip_color(main)[1]

DICT={'크롭':'crop','후드집업':'hood zipup','후드':'hood','후디':'hoody','집업':'zipup','팬츠':'pants','스커트':'skirt','셔츠':'shirt','티셔츠':'tshirt','티':'tee',
'맨투맨':'mtm','니트':'knit','가디건':'cardigan','자켓':'jacket','재킷':'jacket','점퍼':'jumper','코트':'coat','원피스':'dress','드레스':'dress','블라우스':'blouse',
'베스트':'vest','조끼':'vest','슬랙스':'slacks','데님':'denim','진':'jeans','쇼츠':'shorts','반바지':'shorts','캡':'cap','모자':'cap','백':'bag','가방':'bag','숄더백':'shoulder bag',
'와이드':'wide','세미':'semi','부츠컷':'bootscut','롱':'long','숏':'short','미니':'mini','미디':'midi','맥시':'maxi','오버핏':'overfit','유니섹스':'unisex','로고':'logo',
'버뮤다':'bermuda','카고':'cargo','다트':'dart','레오파드':'leopard','디스트로이드':'destroyed','스트라이프':'stripe','체크':'check','슬리브':'sleeve','반팔':'short sleeve',
'긴팔':'long sleeve','슬리브리스':'sleeveless','나시':'sleeveless','플리츠':'pleats','레이어드':'layered','투턱':'two tuck','원턱':'one tuck','트랙':'track','나일론':'nylon',
'스웨이드':'suede','레더':'leather','페이크':'fake','트위드':'tweed','블레이저':'blazer','후드티':'hoody','스웨트셔츠':'sweatshirt','벨벳':'velvet','벨루어':'velour',
'코듀로이':'corduroy','린넨':'linen','니트웨어':'knit','카라':'collar','브이넥':'vneck','라운드':'round','밴딩':'banding','포켓':'pocket','빅':'big','슬림핏':'slimfit',
'워싱':'washing','피그먼트':'pigment','스티치':'stitch','그래픽':'graphic','하프':'half','윙':'wing','다크':'dark','라인':'line','셔링':'shirring','컷아웃':'cutout',
'바시티':'varsity','바람막이':'windbreaker','조거':'jogger','트레이닝':'training','스냅':'snap','버튼':'button','투버튼':'two button','더블':'double','싱글':'single'}
CHO='g kk n d tt r m b pp s ss _ j jj ch k t p h'.split(' ')
JUNG='a ae ya yae eo e yeo ye o wa wae oe yo u wo we wi yu eu ui i'.split(' ')
JONG=['','k','k','k','n','n','n','t','l','l','l','l','l','l','l','l','m','p','p','t','t','ng','t','t','k','t','p','t']
def rom(w):
    o=''
    for ch in w:
        c=ord(ch)-0xAC00
        if 0<=c<11172:
            a,b,d=c//588,(c%588)//28,c%28
            o+=CHO[a].replace('_','')+JUNG[b]+JONG[d]
        else: o+=ch.lower()
    return o
def skel_en(w):
    w=w.lower(); w=re.sub(r'[^a-z]','',w)
    for a,b in [('ph','p'),('sh','s'),('th','t'),('ch','j'),('ck','k'),('x','ks'),('q','k'),('c','k'),('f','p'),('v','p'),('b','p'),('d','t'),('g','k'),('z','j'),('w',''),('y',''),('h','')]: w=w.replace(a,b)
    w=w.replace('r','l').replace('ng','n')
    w=re.sub(r'[aeiou]','',w); return re.sub(r'(.)\1+',r'\1',w)
def skel_ko(w):
    w=rom(w)
    for a,b in [('ch','j'),('kk','k'),('tt','t'),('pp','p'),('ss','s'),('jj','j'),('ng','n'),('g','k'),('d','t'),('b','p'),('h',''),('w',''),('y','')]: w=w.replace(a,b)
    w=w.replace('r','l')
    w=re.sub(r'[aeiou]','',w); return re.sub(r'(.)\1+',r'\1',w)
def ko_tokens(name):
    """한글 상품명 → [(영문 후보, 골격)]"""
    out=[]
    for t in re.split(r'[\s_\-]+',name):
        if not t: continue
        if re.search(r'[가-힣]',t):
            if t in DICT: out.extend(DICT[t].split())
            else:
                # 붙여쓴 합성어(후드집업 등) 사전 분해
                done=False
                for k in sorted(DICT,key=len,reverse=True):
                    if t.endswith(k) and len(t)>len(k) and t[:-len(k)] in DICT:
                        out.extend((DICT[t[:-len(k)]]+' '+DICT[k]).split()); done=True; break
                if not done: out.append('#'+t)
        else: out.append(t.lower())
    return out
def ko_en_score(ko_name, en_name):
    """0~1. 한글 상품명 단어와 영문 상품명 단어를 순서대로 맞춘 비율"""
    kt=ko_tokens(ko_name); et=[w for w in re.split(r'[^a-z0-9]+',en_name.lower()) if w]
    if not kt or not et: return 0
    et=[w for w in re.split(r'[^a-z0-9]+',re.sub(r'zip[\s\-]*up','zipup',re.sub(r't[\s\-]shirt','tshirt',en_name.lower()))) if w]
    used=set(); hit=0
    for t in kt:
        best=0;bi=None
        for i,e in enumerate(et):
            if i in used: continue
            if t.startswith('#'):
                a,b=skel_ko(t[1:]),skel_en(e)
                r=difflib.SequenceMatcher(None,a,b).ratio() if a and b else 0
            else:
                r=1.0 if (t==e or t.rstrip('s')==e.rstrip('s') or (t in('hood','hoody','hoodie') and e in('hood','hoody','hoodie'))) else (0.8 if (len(t)>3 and (e.startswith(t) or t.startswith(e))) else 0)
            if r>best: best,bi=r,i
        if best>=0.6: used.add(bi); hit+=best
    return hit/max(len(kt),len(et))

KO2=dict(KO); KO2.update({'라이트':'LIGHT','다크':'DARK','딥':'DEEP','피그먼트':'PIGMENT','워싱':'WASHING','스카이':'SKY','민트':'MINT','라벤더':'LAVENDER','소라':'SORA','멜란지':'MELANGE','차콜':'CHARCOAL','페일':'PALE','파스텔':'PASTEL','빈티지':'VINTAGE','연':'LIGHT','진':'DEEP'})
KEYS=sorted(KO2,key=len,reverse=True)
def canon(c):
    c=str(c).upper().strip()
    for k in KEYS: c=c.replace(k,KO2[k])
    return re.sub(r'[^A-Z0-9가-힣]','',c).replace('GRAY','GREY').replace('MOCA','MOCHA')
SZ=r'(?:XXS|XS|S|M|L|XL|XXL|2XL|XXXL|3XL|FREE|F|OS|\d{1,3})(?:\([^)]*\)[A-Z0-9]?)?|단일상품'
def parse_opt(v):
    v=re.sub(r'[\x00-\x1f\x7f]',' ',str(v)).strip(); v=re.sub(r'\s+\(','(',v); v=re.sub(r'\)\s+([0-9N])$',r')\1',v)
    if v in ('nan',''): return '',''
    t=[x for x in re.split(r'[\^\s/,:]+',v) if x]
    size=''
    if t and re.fullmatch(SZ,t[-1],flags=re.I): size=t.pop().upper()
    elif len(t)>1 and re.fullmatch(SZ,t[0],flags=re.I): size=t.pop(0).upper()
    return ' '.join(t), size
def size_keys(z):
    z=re.sub(r'(?<=\))[0-9N]$','',str(z).upper().replace(' ',''))
    if not z: return set()
    if z in ('단일상품','F','OS','FREE'): return {'FREE'}
    mm=re.fullmatch(r'([^(]+)\(([^)]*)\)',z)
    if mm:
        k={mm.group(1)}
        if re.fullmatch(SZ,mm.group(2),flags=re.I): k.add(mm.group(2))
        return k
    return {z}
PREF=['PIGMENT','WASHING','VINTAGE','LIGHT','DARK','DEEP','PALE','PASTEL','SKY','WHITE','BLACK','SALT','SOLID','SUEDE']
TYPO={'GERY':'GREY','GRAY':'GREY','MOCA':'MOCHA','CHACOAL':'CHARCOAL','CHA':'CHARCOAL'}
def ctoks(c):
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
    if not want or not have: return None
    if want==have: return 0
    if want<=have or have<=want: return 1+len(want^have)
    return None
def opt_parts(v):
    """29CM 옵션명 → [(구성품 라벨, 옵션값)]  세트는 2개"""
    v=str(v)
    segs=[x.strip() for x in re.split(r'\[[^\]]*\]',v) if x.strip()]
    if len(segs)>=2 and all(re.match(r'^[A-Za-z가-힣]+\s*:',x) for x in segs):
        return [tuple(p.strip() for p in x.split(':',1)) for x in segs]
    if len(segs)>=2 and re.search(r'\[선택\d\]',v): return [('',x) for x in segs]
    return [('',' '.join(segs))]
