"""
무신사 판매중 상품 검색 노출 점검

입력: 무신사 파트너센터 → 재고 관리(v2) 표를 그대로 복사해 엑셀에 붙여넣은 파일
      (머리글 행은 없어도 되고, 어디에 있어도 됨 / 18개 열 또는 필요한 14개 열)
흐름: read_stock_table → make_listings → write_search_script → (브라우저에서 검색) → write_report
"""
import os, re, json, datetime
from collections import Counter
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties
from urllib.parse import quote

# 재고 관리(v2) 화면의 열 (왼쪽부터)
COLUMNS_18 = ['매장명', '업체', '품목', '브랜드', '스타일넘버', '상품구분', '상품번호', '상품명', '상태', '옵션', '옵션번호', '사용유무',
              '국내출고/브랜드배송', '국내출고/무신사배송', '일본출고/무신사배송', '허용판매수량', '총합 판매가능 재고수', '재입고알림요청수']
# 점검에 쓰는 열
KEEP = ['매장명', '업체', '품목', '브랜드', '스타일넘버', '상품번호', '상품명', '상태', '옵션', '옵션번호',
        '국내출고/브랜드배송', '국내출고/무신사배송', '일본출고/무신사배송', '총합 판매가능']
COLUMNS_14 = KEEP
ALIASES = {'총합 판매가능 재고수': '총합 판매가능', '총합 판매가능재고수': '총합 판매가능'}
NUM_COLS = ['국내출고/브랜드배송', '국내출고/무신사배송', '일본출고/무신사배송', '총합 판매가능']


# ------------------------------------------------------------------ 입력 읽기
def _is_header(row):
    vals = {str(v).strip() for v in row if pd.notna(v)}
    return '상품번호' in vals and ('옵션' in vals or '상품명' in vals)


def read_stock_table(path):
    """재고 관리(v2) 복사본 → 필요한 14개 열 DataFrame (판매중만)"""
    raw = pd.read_excel(path, header=None, dtype=str, sheet_name=0)
    raw = raw.dropna(how='all').dropna(axis=1, how='all')
    hdr_idx = [i for i, row in raw.iterrows() if _is_header(row.values)]
    if hdr_idx:
        names = [ALIASES.get(str(v).strip(), str(v).strip()) for v in raw.loc[hdr_idx[0]].values]
        df = raw.drop(index=hdr_idx)
        df.columns = names
    elif raw.shape[1] == len(COLUMNS_18):
        df = raw.copy(); df.columns = COLUMNS_18
    elif raw.shape[1] == len(COLUMNS_14):
        df = raw.copy(); df.columns = COLUMNS_14
    else:
        raise ValueError(f'열이 {raw.shape[1]}개입니다. 재고 관리(v2) 표 전체(18개 열)를 그대로 붙여넣은 파일인지 확인해 주세요.\n'
                         f'(맨 위나 아래에 머리글 행이 있으면 열 개수와 상관없이 읽습니다)')
    df = df.rename(columns=ALIASES)
    need = ['스타일넘버', '상품번호', '상품명', '옵션', '총합 판매가능']
    miss = [c for c in need if c not in df.columns]
    if miss:
        raise ValueError(f'필요한 열이 없습니다: {", ".join(miss)}')
    df = df[[c for c in KEEP if c in df.columns]].copy()
    for c in df.columns:
        df[c] = df[c].astype(str).str.strip().replace({'nan': '', 'None': ''})
    df = df[df['상품번호'].str.fullmatch(r'\d+')]
    if '상태' in df.columns:
        df = df[df['상태'].isin(['판매중', ''])]
    for c in NUM_COLS:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c].str.replace(',', ''), errors='coerce').fillna(0).astype(int)
    return df.reset_index(drop=True)


# ------------------------------------------------------------------ 상품명 정리 (컬러·컬러묶음 표기 제거)
KO = {'블랙': 'BLACK', '화이트': 'WHITE', '아이보리': 'IVORY', '베이지': 'BEIGE', '차콜': 'CHARCOAL', '그레이': 'GREY', '회색': 'GREY',
      '네이비': 'NAVY', '블루': 'BLUE', '스카이블루': 'SKYBLUE', '소라': 'SORA', '카키': 'KHAKI', '브라운': 'BROWN', '핑크': 'PINK',
      '레드': 'RED', '퍼플': 'PURPLE', '올리브': 'OLIVE', '크림': 'CREAM', '다크베이지': 'DARKBEIGE', '다크네이비': 'DARKNAVY',
      '다크그레이': 'DARKGREY', '연베이지': 'LIGHTBEIGE', '연그레이': 'LIGHTGREY', '진청': 'DEEPBLUE', '연청': 'LIGHTBLUE',
      '중청': 'MIDBLUE', '흑청': 'BLACKDENIM', '멜란지': 'MELANGE', '오트밀': 'OATMEAL', '모카': 'MOCHA', '카멜': 'CAMEL',
      '그린': 'GREEN', '옐로우': 'YELLOW', '민트': 'MINT', '라벤더': 'LAVENDER', '와인': 'WINE', '버건디': 'BURGUNDY',
      '실버': 'SILVER', '골드': 'GOLD', '오렌지': 'ORANGE', '인디고': 'INDIGO', '딥블루': 'DEEPBLUE', '라이트블루': 'LIGHTBLUE',
      '라이트그레이': 'LIGHTGREY', '머스타드': 'MUSTARD', '코코아': 'COCOA'}
KO_PART = dict(KO, **{'라이트': 'LIGHT', '다크': 'DARK', '딥': 'DEEP', '피그먼트': 'PIGMENT', '워싱': 'WASHING', '스카이': 'SKY',
                      '페일': 'PALE', '파스텔': 'PASTEL', '빈티지': 'VINTAGE', '연': 'LIGHT', '진': 'DEEP'})
EN_BASE = set(KO.values()) | {'BLACK', 'WHITE', 'IVORY', 'BEIGE', 'CHARCOAL', 'GREY', 'GRAY', 'NAVY', 'BLUE', 'SKY', 'KHAKI', 'BROWN',
                              'PINK', 'RED', 'PURPLE', 'OLIVE', 'CREAM', 'LIGHT', 'DARK', 'DEEP', 'PALE', 'PIGMENT', 'WASHING',
                              'MELANGE', 'OATMEAL', 'MOCHA', 'MOCA', 'CAMEL', 'GREEN', 'YELLOW', 'MINT', 'LAVENDER', 'WINE',
                              'SILVER', 'GOLD', 'ORANGE', 'INDIGO', 'BUTTER', 'SALAD', 'SORA', 'STITCH', 'COATING', 'CAMO',
                              'MULTI', 'CHECK', 'SALT'}
NOT_COLOR = {'AV', 'T', 'SET', 'UNISEX', 'HOODIE', 'HOODY', 'SHIRT', 'PANTS', 'TOP', 'CROP', 'TEE', 'SKIRT', 'KNIT', 'JACKET',
             'BLOUSE', 'MTM', 'DRESS', 'VEST', 'COAT', 'CD', 'OPS', 'BAG', 'CAP', 'BEANIE', 'SHORTS', 'HOOD', 'ZIP-UP', 'T-SHIRT',
             'CARDIGAN', 'SLACKS', 'JUMPER', 'SWEATSHIRT', 'SLEEVELESS', 'LONG', 'SHORT', 'SLEEVE'}
_COL_TAIL = re.compile(r'[\s_\-]*\(?\d+\s*COL(OR)?S?\)?(\(N\))?\s*$', re.I)


def _is_ko_color(tok):
    c = tok
    for k in sorted(KO_PART, key=len, reverse=True):
        c = c.replace(k, KO_PART[k])
    return bool(c) and re.fullmatch(r'[A-Z]+', c) is not None


def color_words(names):
    """상품명 끝에 자주 붙는 대문자 단어를 컬러로 수집 (예: SALT, PIGMENT)"""
    en = set(EN_BASE)
    tail, mid = Counter(), Counter()
    for nm in pd.Series(names).dropna().unique():
        nm = _COL_TAIL.sub('', re.sub(r'^\s*(\[[^\]]*\]\s*)+', '', str(nm)))
        t = re.sub(r'[_\-/]', ' ', nm).split()
        if any(re.search(r'[a-z가-힣]', w) for w in t):
            tl = []
            for w in reversed(t):
                if re.fullmatch(r'[A-Z][A-Z0-9/]*', w): tl.insert(0, w)
                else: break
            if 0 < len(tl) < len(t) and len(tl) <= 3: tail.update(tl)
            mid.update(w.upper() for w in t[:len(t) - len(tl)])
        else:
            mid.update(w.upper() for w in t[:-1])
    en.update(w for w, k in tail.items() if k > mid[w] and not re.search(r'\d', w))
    return en - NOT_COLOR


def clean_name(n, en):
    """'모로우 투 턱 슬랙스_2COL' → '모로우 투 턱 슬랙스', 'Thin Bootscut Pants LIGHT BLUE' → 'Thin Bootscut Pants'"""
    n = str(n).strip()
    n = re.sub(r'^\s*(\[[^\]]*\]\s*)+', '', n)
    n = _COL_TAIL.sub('', n)
    n = re.sub(r'\s+-\s*$', '', n)
    t = n.split()
    mixed = any(re.search(r'[a-z가-힣]', w) for w in t)
    while len(t) > 2:
        core = re.sub(r'\(.*?\)', '', t[-1])
        if re.fullmatch(r'[가-힣]+', core) and _is_ko_color(core): t.pop(); continue
        if re.fullmatch(r'[A-Z][A-Z0-9*]*', core) and core not in NOT_COLOR and (mixed or core in en): t.pop(); continue
        break
    return re.sub(r'[_]+', ' ', ' '.join(t)).strip(' -_')


# ------------------------------------------------------------------ 상품 목록
def make_listings(df, exclude_pattern=None):
    """옵션 1줄 → 무신사 상품번호 1줄"""
    d = df.copy()
    if exclude_pattern:
        d = d[~d['스타일넘버'].str.upper().str.match(exclude_pattern)]
    g = d.groupby('상품번호', sort=False)
    L = pd.DataFrame({
        '상품명': g['상품명'].first(),
        '스타일넘버': g['스타일넘버'].first(),
        '브랜드': g['브랜드'].first() if '브랜드' in d.columns else '',
        '옵션수': g.size(),
        '판매가능옵션수': g['총합 판매가능'].apply(lambda v: int((v > 0).sum())),
        '총합판매가능': g['총합 판매가능'].sum(),
    }).reset_index()
    for c, col in [('브랜드배송', '국내출고/브랜드배송'), ('MFS', '국내출고/무신사배송')]:
        L[c] = g[col].sum().values if col in d.columns else 0
    en = color_words(L['상품명'])
    L['검색어'] = L['상품명'].map(lambda v: clean_name(v, en))
    excluded = df['상품번호'].nunique() - L['상품번호'].nunique()
    return L, excluded


def write_search_script(L, brand, template_path, out_path):
    tpl = open(template_path, encoding='utf-8').read()
    kw = json.dumps([[a, b] for a, b in zip(L['상품번호'], L['검색어'])], ensure_ascii=False)
    tpl = tpl.replace('const BRAND = "어반드레스";', f'const BRAND = {json.dumps(brand, ensure_ascii=False)};')
    tpl = tpl.replace('const KEYWORDS = [];', f'const KEYWORDS = {kw};')
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(tpl)
    return out_path


def parse_search_result(text):
    rows = [l.strip().split('|')[:5] for l in text.replace(';', '\n').splitlines() if l.strip()]
    if not rows or any(len(r) < 4 for r in rows):
        raise ValueError('검색 결과 형식이 아닙니다. 무신사 콘솔에서 복사한 결과를 그대로 붙여넣어 주세요.')
    rows = [r + [''] * (5 - len(r)) for r in rows]
    sr = pd.DataFrame(rows, columns=['상품번호', '검색결과수', '순위', '브랜드결과수', '재검색어'])
    for c in ['검색결과수', '브랜드결과수']:
        sr[c] = pd.to_numeric(sr[c], errors='coerce').fillna(0).astype(int)
    sr['순위'] = pd.to_numeric(sr['순위'], errors='coerce')
    return sr.drop_duplicates('상품번호').set_index('상품번호')


# ------------------------------------------------------------------ 결과 엑셀
F = 'Arial'
H = PatternFill('solid', fgColor='1F3864'); HF = Font(name=F, bold=True, color='FFFFFF', size=10)
DESC = PatternFill('solid', fgColor='EEF3FA')
RETRY_LABEL = '검색됨 (UNISEX·유니섹스 빼고 재검색)'
FILL = {'검색 안 됨': 'F8CBAD', '미확인': 'E7E6E6', RETRY_LABEL: 'FFEB9C'}


def _note(row):
    if row['검색여부'] == RETRY_LABEL:
        return (f"원래 검색어로는 안 나오고, UNISEX·유니섹스를 뺀 \"{row['재검색어']}\"로 검색하면 나옴 "
                "→ 상품명에서 UNISEX·유니섹스 제외 검토")
    if row['검색여부'] != '검색 안 됨':
        return '판매가능 재고 0 (전 옵션 품절)' if row['총합판매가능'] == 0 else ''
    if re.search(r'유니섹스|unisex', str(row['검색어']), re.I):
        return 'UNISEX·유니섹스를 빼고 다시 검색해도 안 나옴 → 무신사에서 상품 전시 상태 확인 필요'
    if row['총합판매가능'] == 0:
        return '판매가능 재고 0 (전 옵션 품절) — 품절 상품이라 검색에서 빠졌을 수 있음'
    if row['검색결과수'] == 0:
        return '검색 결과 0건 → 상품명·검색 키워드 확인 필요'
    return '검색 결과에 다른 상품만 나옴 → 무신사에서 상품 전시 상태 확인 필요'


def _desc_height(desc, widths):
    import math
    total = max(sum(widths), 20); n = 0
    for ln in desc.split('\n'):
        n += max(1, math.ceil(sum(1.8 if ord(ch) > 0x1100 else 1.0 for ch in ln) * 1.05 / total))
    return max(30, 15 * n + 10)


def _sheet(wb, name, df, cols, heads, widths, desc, freeze='D3'):
    ws = wb.create_sheet(name)
    ws.cell(1, 1, desc).font = Font(name=F, size=10, color='1F3864')
    ws.cell(1, 1).alignment = Alignment(wrap_text=True, vertical='top'); ws.cell(1, 1).fill = DESC
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(cols))
    for j, h in enumerate(heads, 1):
        x = ws.cell(2, j, h); x.font = HF; x.fill = H
        x.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for i, row in enumerate(df[cols].itertuples(index=False), 3):
        for j, v in enumerate(row, 1):
            if v is None or (isinstance(v, float) and np.isnan(v)): v = None
            elif isinstance(v, np.integer): v = int(v)
            elif isinstance(v, float) and v.is_integer(): v = int(v)
            x = ws.cell(i, j, v); x.font = Font(name=F, size=10)
            if isinstance(v, str) and v.startswith('https://'):
                x.hyperlink = v; x.font = Font(name=F, size=10, color='0563C1', underline='single')
            elif isinstance(v, str) and cols[j - 1] == '캡처':
                x.value = '캡처 보기'; x.hyperlink = v.replace(os.sep, '/')
                x.font = Font(name=F, size=10, color='0563C1', underline='single')
            if isinstance(v, int): x.number_format = '#,##0'
            if j == 1:
                x.fill = PatternFill('solid', fgColor=FILL.get(v, 'C6EFCE'))
    for j, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.row_dimensions[1].height = _desc_height(desc, widths)
    ws.row_dimensions[2].height = 32
    ws.freeze_panes = freeze
    ws.auto_filter.ref = f'A2:{get_column_letter(len(cols))}{max(len(df) + 2, 3)}'
    return ws


def write_report(L, sr, brand, out_path, source_name='', excluded=0):
    """sr: parse_search_result 결과 또는 None(검색 전)"""
    R = L.copy()
    if sr is not None:
        R = R.join(sr, on='상품번호')
        R['재검색어'] = R['재검색어'].fillna('')
        R['검색여부'] = np.where(R['순위'].notna(), np.where(R['재검색어'] != '', RETRY_LABEL, f'검색됨 ({brand} 브랜드 확인)'),
                              np.where(R['검색결과수'].notna(), '검색 안 됨', '미확인'))
    else:
        R['검색결과수'] = np.nan; R['브랜드결과수'] = np.nan; R['순위'] = np.nan; R['검색여부'] = '미확인'; R['재검색어'] = ''
    R['검색결과수'] = R['검색결과수'].fillna(0)
    R['검색어(브랜드 포함)'] = brand + ' ' + R['검색어']
    R['비고'] = R.apply(_note, axis=1)
    R['검색링크'] = 'https://www.musinsa.com/search/goods?keyword=' + R['검색어(브랜드 포함)'].map(lambda v: quote(str(v)))
    R['상품링크'] = 'https://www.musinsa.com/products/' + R['상품번호'].astype(str)
    if '캡처' not in R.columns:
        R['캡처'] = None
    R['_k'] = R['검색여부'].map(lambda v: 0 if v == '검색 안 됨' else (1 if v == RETRY_LABEL else (2 if v == '미확인' else 3)))
    R = R.sort_values(['_k', '총합판매가능'], ascending=[True, False])

    cols = ['검색여부', '상품번호', '상품명', '스타일넘버', '검색어(브랜드 포함)', '검색결과수', '브랜드결과수', '순위', '비고',
            '캡처', '옵션수', '판매가능옵션수', '총합판매가능', '브랜드배송', 'MFS', '검색링크', '상품링크']
    heads = ['검색 결과', '무신사 상품번호', '무신사 상품명', '스타일넘버', '검색어 (브랜드명 + 상품명, 컬러·COL 제외)', '검색 결과 수',
             f'화면에 보인 {brand} 상품 수', '해당 상품 화면 순위', '비고', '검색 화면 캡처 (미노출_캡처 폴더)', '옵션 수', '판매가능 옵션 수', '총합 판매가능 재고',
             '브랜드배송 재고', '무신사배송(MFS) 재고', '무신사 검색 링크', '무신사 상품 링크']
    widths = [24, 11, 40, 14, 40, 9, 11, 9, 50, 12, 7, 9, 10, 10, 10, 38, 36]

    wb = Workbook(); ws = wb.active; ws.title = '요약'
    miss = R[R['검색여부'] == '검색 안 됨']
    _sheet(wb, '검색_미노출', miss, cols, heads, widths,
           f'[왜 만들었나] 무신사 판매중 상품인데, 무신사 검색창에 "{brand} + 상품명"을 넣었을 때 검색 결과 화면에 나오지 않는 상품만 모은 시트입니다. '
           f'실제 작업은 이 시트부터 보시면 됩니다.\n[읽는 법] 비고에 원인(전 옵션 품절, 다른 상품만 나옴 등)을 적었습니다. UNISEX·유니섹스가 들어간 상품명은 그 단어를 빼고 한 번 더 검색한 결과입니다. 링크를 눌러 무신사 화면에서 직접 확인할 수 있고, "캡처 보기"를 누르면 검색했을 때의 화면 캡처가 열립니다 (엑셀 파일과 미노출_캡처 폴더가 같은 폴더에 있어야 함).')
    _sheet(wb, '전체_검색결과', R, cols, heads, widths,
           f'[왜 만들었나] 판매중 상품 전체의 검색 결과입니다. 검색이 안 된 상품이 맨 위에 옵니다.\n'
           f'[읽는 법] 검색어는 상품명에서 컬러명·"_3COL" 같은 컬러 묶음 표시를 뗀 이름 앞에 "{brand}"를 붙인 것입니다. '
           f'"검색됨"은 실제 검색 결과 화면에 해당 상품번호가 브랜드 "{brand}"로 나왔다는 뜻입니다. "{RETRY_LABEL}"은 원래 검색어로는 안 나오고 UNISEX·유니섹스를 빼고 다시 검색하니 나온 상품입니다. "미확인"은 아직 브라우저 검색을 하지 않은 상품입니다.')
    b = Font(name=F, bold=True, size=10)
    ws['A1'] = f'무신사 검색 노출 점검 — {brand} 판매중 상품'; ws['A1'].font = Font(name=F, bold=True, size=14)
    ws['A2'] = f'기준: {source_name} · 작성 {datetime.date.today():%Y-%m-%d}'; ws['A2'].font = Font(name=F, size=9, color='808080')
    ws['A3'] = '[이 시트] 전체 결과를 숫자로 모은 시트입니다 (다른 시트를 수식으로 집계). 작업은 "검색_미노출" 시트부터 보세요.'
    ws['A3'].font = Font(name=F, size=10, color='1F3864')
    for j, t in enumerate(['검색 결과', '상품 수'], 1):
        x = ws.cell(5, j, t); x.font = HF; x.fill = H
    labels = [f'검색됨 ({brand} 브랜드 확인)', RETRY_LABEL, '검색 안 됨', '미확인']
    for i, t in enumerate(labels, 6):
        ws.cell(i, 1, t).font = b
        ws.cell(i, 1).fill = PatternFill('solid', fgColor=FILL.get(t, 'C6EFCE'))
        ws.cell(i, 2, f"=COUNTIFS('전체_검색결과'!$A:$A,$A{i})").number_format = '#,##0'
    ws.cell(10, 1, '합계 (판매중 상품 수)').font = b; ws.cell(10, 2, '=SUM(B6:B9)').font = b
    r = 12
    ws.cell(r, 1, '검색 안 됨 — 원인별').font = b
    if len(miss):
        for i, (k, v) in enumerate(miss['비고'].str.split(' — | →|\\(사이트').str[0].value_counts().items(), r + 1):
            ws.cell(i, 1, k).font = Font(name=F, size=10); ws.cell(i, 2, int(v))
        r = r + 1 + miss['비고'].nunique()
    else:
        ws.cell(r + 1, 1, '없음').font = Font(name=F, size=10); r += 2
    if excluded:
        ws.cell(r + 1, 1, f'제외한 다른 브랜드 상품 (설정의 제외 코드): {excluded}개').font = Font(name=F, size=9, color='808080')
    ws.column_dimensions['A'].width = 60; ws.column_dimensions['B'].width = 12
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    wb.save(out_path)
    return R
