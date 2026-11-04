"""
29CM 옵션 재고 점검 — 29CM 옵션재고 일괄변경 파일(페이지별 여러 개) ↔ 셀메이트 재고목록

29CM 파일에는 셀메이트 상품코드가 없어서 상품명으로 셀메이트 상품을 찾고, 옵션(컬러·사이즈)을 1:1로 맞춥니다.
  read_29cm(paths)            페이지 파일 여러 개 → 옵션 목록 1개 (같은 옵션은 한 번만)
  read_sellmate(path)         셀메이트 재고목록 CSV
  match(a, s)                 상품 매칭 → 옵션 매칭 → 판정
  write_report(...)           29CM_옵션재고점검 엑셀
"""
import os, re, difflib
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from cm29_names import (names, strip_color, base_clean, key, ko_en_score, ctoks, parse_opt, size_keys, color_score,
                        opt_parts, COLT)

BASE_RE = r'^([A-Z]{2}\d{2}[A-Z]{2}\d{2})'
PRODUCT_URL = 'https://www.29cm.co.kr/products/'
NEED = ['스타일 번호', '상품 번호', '상품명', '옵션명', '옵션 판매 상태']


# ------------------------------------------------------------------ 입력
def _read_one(path):
    raw = pd.read_excel(path, sheet_name=0, header=None, dtype=str)
    hdr = None
    for i in range(min(len(raw), 15)):
        vals = {str(v).replace('\n', ' ').strip() for v in raw.iloc[i].values if pd.notna(v)}
        if '스타일 번호' in vals and '상품명' in vals:
            hdr = i; break
    if hdr is None:
        raise ValueError(f'{os.path.basename(path)}: 29CM 옵션재고 일괄변경 파일이 아닙니다 ("스타일 번호", "상품명" 열을 찾지 못함).')
    d = raw.iloc[hdr + 1:].copy()
    d.columns = [str(c).replace('\n', ' ').strip() for c in raw.iloc[hdr].values]
    d = d.loc[:, [c for c in d.columns if c and c != 'nan']]
    miss = [c for c in NEED if c not in d.columns]
    if miss:
        raise ValueError(f'{os.path.basename(path)}: 필요한 열이 없습니다 ({", ".join(miss)})')
    d = d.rename(columns={'재고 수량 (변경 전)': '재고 수량', '옵션 추가 가격 (변경 전)': '옵션 추가 가격',
                          '재입고 알림 여부 (변경 전)': '재입고 알림 여부'})
    d = d[d['스타일 번호'].notna() & d['상품명'].notna()]
    d['원본 파일'] = os.path.basename(path)
    return d


def read_29cm(paths):
    """29CM_옵션재고_일괄변경(파트너)_페이지 파일 여러 개 → 하나로 (스타일 번호 기준 중복 제거)"""
    if isinstance(paths, str):
        paths = [p for p in paths.split(';') if p.strip()]
    if not paths:
        raise FileNotFoundError('29CM 옵션재고 일괄변경 파일을 선택해 주세요.')
    ds = [_read_one(p) for p in paths]
    a = pd.concat(ds, ignore_index=True)
    n0 = len(a)
    a = a.drop_duplicates('스타일 번호').reset_index(drop=True)
    for c in ['상품 번호', '옵션 번호']:
        if c in a.columns:
            a[c] = a[c].astype(str).str.replace(r'\.0$', '', regex=True)
    for c in ['재고 수량', '판매가']:
        a[c] = pd.to_numeric(a.get(c, 0), errors='coerce').fillna(0).astype(int)
    if '브랜드' not in a.columns:
        a['브랜드'] = ''
    a['옵션 판매 상태'] = a['옵션 판매 상태'].fillna('').str.strip()
    return a, n0 - len(a)


def read_sellmate(path, target_pattern=None):
    s = pd.read_csv(path, encoding='cp949', dtype=str)
    s = s.loc[:, ~s.columns.str.startswith('Unnamed')]
    for c in s.columns:
        s[c] = s[c].astype(str).str.replace(r'^=?"*|"*$', '', regex=True).str.replace('=', '').str.strip('"')
    s = s.fillna('').astype(str)
    s = s[s['상품코드'].ne('nan') & s['상품코드'].ne('')].copy()
    s = s.replace({'nan': ''})
    s['재고'] = pd.to_numeric(s['현재재고'], errors='coerce').fillna(0).astype(int)
    s['가격'] = pd.to_numeric(s.get('대표판매가', 0), errors='coerce').fillna(0).astype(int)
    s['우리브랜드'] = s['상품코드'].str.match(target_pattern) if target_pattern else True
    s['base'] = s['상품코드'].str.extract(BASE_RE)[0].fillna(s['상품코드'])
    return s.reset_index(drop=True)


# ------------------------------------------------------------------ 상품 매칭
def _words(n, strip=True):
    n = strip_color(base_clean(n))[0] if strip else base_clean(n)
    return set(w for w in re.split(r'[^0-9a-z가-힣]+', re.sub(r'zip[\s\-]*up', 'zipup', n.lower())) if w)


class _Finder:
    def __init__(self, s):
        P = s.groupby('상품코드').agg(상품명=('상품명', 'first'), base=('base', 'first'), 가격=('가격', 'max'),
                                     우리=('우리브랜드', 'first')).reset_index()
        P['keys'] = P['상품명'].map(lambda v: names(v)[0])
        P['words'] = P['상품명'].map(lambda n: _words(n, strip=False))
        self.idx = {}
        for c, ks in zip(P['상품코드'], P['keys']):
            for k in ks:
                self.idx.setdefault(k, set()).add(c)
        self.allkeys = list(self.idx)
        self.P = P
        self.EN = P[~P['상품명'].str.contains('[가-힣]')]
        self.info = P.set_index('상품코드')

    def find(self, name, price, sib=None):
        idx, P = self.idx, self.P
        ks, _ = names(name)
        hit = set().union(*[idx.get(k, set()) for k in ks]) if ks else set()
        if hit:
            return hit, '상품명 일치'
        for k in ks:   # 셀메이트 상품명 = 29CM 상품명 + 컬러
            if len(k) < 6:
                continue
            pre = set()
            for kk in self.allkeys:
                rest = kk[len(k):]
                if kk.startswith(k) and rest and ctoks(rest) and re.fullmatch(r'[a-z가-힣]*', rest):
                    pre |= idx[kk]
            if pre:
                return pre, '상품명 일치(컬러 표기 차이)'
        W = _words(name)
        if len(W) >= 2:   # 29CM 상품명 단어가 모두 들어 있는 셀메이트 상품
            sup = P[P['words'].map(lambda x: W <= x)]
            if sib is not None and len(sup):
                sw = _words(sib) - W
                pref = sup[sup['words'].map(lambda x: bool(x & sw))]
                if len(pref):
                    sup = pref
            if len(sup) and sup['base'].nunique() == 1:
                return set(sup['상품코드']), '상품명 단어 포함(확인필요)'
        best = None
        for k in ks:
            if len(k) < 6:
                continue
            for f in difflib.get_close_matches(k, self.allkeys, n=5, cutoff=0.86):
                r = difflib.SequenceMatcher(None, k, f).ratio()
                for c in idx[f]:
                    sc = r + (0.05 if self.info.at[c, '가격'] == price else 0)
                    if best is None or sc > best[0]:
                        best = (sc, f)
        if best:
            return set(idx[best[1]]), '비슷한 상품명(확인필요)'
        if re.search('[가-힣]', name) and len(self.EN):
            core = strip_color(base_clean(name))[0]
            E = self.EN.assign(_s=self.EN['상품명'].map(
                lambda e: max(ko_en_score(core, strip_color(base_clean(e))[0]), ko_en_score(core, base_clean(e)))))
            E['_s'] = E['_s'] + np.where(E['가격'] == price, 0.05, 0)
            top = E[E['_s'] >= 0.8]
            if len(top):
                m = top['_s'].max(); b = set(top[top['_s'] >= m - 0.01]['base'])
                return set(P[P['base'].isin(b)]['상품코드']), '한글↔영문 상품명(확인필요)'
        return set(), '셀메이트에서 못 찾음'

    def narrow(self, codes, price):
        """여러 스타일이 걸리면 가격 같은 스타일 → 우리 브랜드 순으로 좁힘 (최종 선택은 옵션 컬러로)"""
        if not codes:
            return codes
        d = self.info.loc[sorted(codes)]
        if d['base'].nunique() > 1:
            same = d[d['가격'] == price]
            if len(same) and same['base'].nunique() < d['base'].nunique():
                d = same
        if d['base'].nunique() > 1:
            u = d[d['우리'].astype(bool)]
            if len(u):
                d = u
        return set(d.index)


def _components(n):
    """[SET] A + B / A & B → [A, B]"""
    if re.match(r'^\s*\[(SET|세트)\]', n, re.I):
        b = COLT.sub('', re.sub(r'^\s*\[[^\]]*\]\s*', '', n))
        parts = [x.strip() for x in re.split(r'\s+[+&]\s+', b) if x.strip()]
        if len(parts) >= 2:
            return parts
    return [n]


# ------------------------------------------------------------------ 옵션 매칭
def _variant(z):
    v = re.findall(r'\(([^)]*[가-힣][^)]*)\)', str(z))
    return v[0] if v else ''


class _OptMatcher:
    def __init__(self, s):
        self.s = s
        s['pcol'] = s['상품명'].map(lambda v: ctoks(names(v)[1]))
        po = s['옵션명'].map(parse_opt)
        s['ocol'] = po.str[0].map(ctoks); s['osize'] = po.str[1]
        s['col'] = [oc if oc else pc for oc, pc in zip(s['ocol'], s['pcol'])]
        self.rows_by_code = s.groupby('상품코드').groups
        self.codes_by_base = s.groupby('base')['상품코드'].unique().to_dict()

    def match(self, codes, val, name_col, pname=''):
        s = self.s
        if not codes:
            return None, '상품 못 찾음'
        bs = {re.match(BASE_RE, c).group(1) for c in codes if re.match(BASE_RE, c)}
        codes = set(codes) | {c for b in bs for c in self.codes_by_base.get(b, [])}   # 같은 스타일(앞 8자리) 전체
        C = s.loc[sorted(i for c in codes for i in self.rows_by_code.get(c, []))]
        col, size = parse_opt(val)
        want = ctoks(col) if col else frozenset()
        raw = re.sub(r'\s', '', col)
        if not want and re.search('[가-힣]', raw):   # 사전에 없는 한글 컬러(반다이크브라운 등)는 글자 그대로 찾기
            txt = (C['상품명'] + C['옵션명']).str.replace(r'\s', '', regex=True)
            C = C[txt.str.contains(re.escape(raw))]
            if C.empty:
                return None, '셀메이트에 해당 컬러 없음'
            C = C.assign(col=[frozenset()] * len(C))
        if not want and not raw:
            want = name_col
        if want:
            sc = C['col'].map(lambda h: color_score(want, h)); ok = sc.dropna()
            if ok.empty:
                w = ''.join(sorted(want))
                fz = C['col'].map(lambda h: difflib.SequenceMatcher(None, w, ''.join(sorted(h))).ratio() if h else 0)
                if len(fz) and fz.max() >= 0.85:
                    hit = C.loc[fz[fz == fz.max()].index]
                elif C['col'].map(len).eq(0).all():
                    hit = C
                else:
                    return None, '셀메이트에 해당 컬러 없음'
            else:
                hit = C.loc[ok[ok == ok.min()].index]
        else:
            if C['col'].nunique() > 1:
                return None, '컬러 정보 없음'
            hit = C
        if not size:
            if hit['옵션코드'].nunique() == 1:
                return hit.index[0], '일치'
            return None, '사이즈 정보 없음'
        zz = size.upper().replace(' ', '')
        vt = _variant(zz)   # (기모)·(쭈리) 같은 구분은 같아야 같은 옵션
        hit = hit[hit['osize'].map(_variant) == vt]
        if hit.empty:
            return None, (f'셀메이트에 해당 옵션 없음({vt})' if vt else '셀메이트에 해당 사이즈 없음')
        zz = re.sub(r'(?<=\))[0-9N]$', '', zz)
        ex = hit[hit['osize'].str.upper().str.replace(' ', '').str.replace(r'(?<=\))[0-9N]$', '', regex=True) == zz]
        if len(ex) == 0:
            k = size_keys(size); ex = hit[hit['osize'].map(lambda v: bool(size_keys(v) & k))]
        if len(ex) == 0 and hit['osize'].eq('').all() and hit['옵션코드'].nunique() == 1:
            ex = hit
        if len(ex) == 0:
            return None, '셀메이트에 해당 사이즈 없음'
        if ex['상품코드'].nunique() > 1:
            me = key(strip_color(base_clean(pname))[0])
            sim = ex['상품명'].map(lambda v: difflib.SequenceMatcher(None, me, key(strip_color(base_clean(v))[0])).ratio())
            top = ex[sim >= sim.max() - 0.001]
            if top['상품코드'].nunique() == 1:
                return top.sort_values('재고', ascending=False).index[0], '일치'
            return top.sort_values('재고', ascending=False).index[0], '일치(같은 컬러·사이즈 상품 여러 개·확인필요)'
        return ex.sort_values('재고', ascending=False).index[0], '일치'


JUDGE_ORDER = ['셀메이트 재고 있는데 29CM 품절', '셀메이트 재고 있는데 29CM 판매중지', '셀메이트 옵션 못 찾음',
               '29CM 판매중 · 셀메이트 재고 0', '정상', '정상(양쪽 재고 없음)']


def _judge(st, q):
    if pd.isna(q):
        return '셀메이트 옵션 못 찾음'
    if st == '판매중':
        return '정상' if q > 0 else '29CM 판매중 · 셀메이트 재고 0'
    if q > 0:
        return '셀메이트 재고 있는데 29CM 품절' if st == '일시품절' else '셀메이트 재고 있는데 29CM 판매중지'
    return '정상(양쪽 재고 없음)'


def match(a, s, log=print):
    """반환: R(옵션별 결과), A(상품별 매칭), res({상품번호: [구성품별 {codes, how}]})"""
    F = _Finder(s)
    A = a.drop_duplicates('상품 번호')[['상품 번호', '상품명', '판매가']].copy()
    res = {}
    for _, r in A.iterrows():
        cs = _components(r['상품명']); out = []
        for i, c in enumerate(cs):
            price = r['판매가'] if len(cs) == 1 else -1
            h, how = F.find(c, price, cs[1 - i] if len(cs) == 2 else None)
            out.append(dict(codes=F.narrow(h, price), how=how))
        res[r['상품 번호']] = out
    A['how'] = A['상품 번호'].map(lambda k: ' / '.join(x['how'] for x in res[k]))
    A['set'] = A['상품 번호'].map(lambda k: len(res[k]) > 1)
    log(f'      상품 {len(A):,}개 중 셀메이트에서 찾음 {int((~A["how"].str.contains("못 찾음")).sum()):,}개'
        f' (확인필요 {int(A["how"].str.contains("확인필요").sum()):,}개)')
    M = _OptMatcher(s)
    rows = []
    for ix, r in a.iterrows():
        comp = res[r['상품 번호']]
        parts = opt_parts(r['옵션명'])
        if len(comp) == 1:
            pairs = [(comp[0], v) for _, v in parts]
        else:
            pairs = list(zip(comp, [v for _, v in parts])) if len(parts) == len(comp) else [(c, '') for c in comp]
        nm = [r['상품명']] if len(comp) == 1 else _components(r['상품명'])
        got = []
        for j, (cp, v) in enumerate(pairs):
            pn = nm[min(j, len(nm) - 1)]
            got.append(M.match(cp['codes'], v, ctoks(names(pn)[1]), pn))
        rec = {'_i': ix}
        if all(i is not None for i, _ in got):
            S = s.loc[[i for i, _ in got]]
            rec.update(셀메이트상품코드=' + '.join(S['상품코드']), 셀메이트옵션코드=' + '.join(S['옵션코드']),
                       셀메이트상품명=' + '.join(S['상품명']), 셀메이트옵션명=' + '.join(S['옵션명']),
                       셀메이트재고=int(S['재고'].min()), 셀메이트판매여부=' + '.join(S['판매여부'].unique()),
                       옵션매칭=' / '.join(sorted(set(w for _, w in got))))
        else:
            rec.update(옵션매칭=' / '.join(w for i, w in got if i is None), 셀메이트재고=np.nan)
        rec['_codes'] = [s.at[i, '옵션코드'] for i, _ in got if i is not None]
        rows.append(rec)
    O = pd.DataFrame(rows).set_index('_i')
    R = a.join(O)
    R['매칭방식'] = R['상품 번호'].map(A.set_index('상품 번호')['how'])
    R['판정'] = [_judge(st, q) for st, q in zip(R['옵션 판매 상태'], R['셀메이트재고'])]
    R['링크'] = PRODUCT_URL + R['상품 번호'].astype(str)
    return R, A, res


# ------------------------------------------------------------------ 결과 엑셀
F_NAME = '맑은 고딕'
HEAD_FILL = PatternFill('solid', fgColor='1F2937')
NOTE_FILL = PatternFill('solid', fgColor='FFF7E6')


def write_sheet(wb, name, df, note, widths, link_col='링크', link_text='29CM 보기', extra_links=None):
    ws = wb.create_sheet(name); n = max(len(df.columns), 1)
    c = ws.cell(1, 1, note); c.font = Font(name=F_NAME, size=10); c.fill = NOTE_FILL
    c.alignment = Alignment(wrap_text=True, vertical='top')
    ws.merge_cells(start_row=1, end_row=1, start_column=1, end_column=n)
    ws.row_dimensions[1].height = 15 * (note.count('\n') + 1) + 8
    for j, col in enumerate(df.columns, 1):
        x = ws.cell(2, j, col); x.font = Font(name=F_NAME, bold=True, color='FFFFFF', size=10); x.fill = HEAD_FILL
        x.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[2].height = 30
    links = {}
    if link_col in df.columns:
        links[list(df.columns).index(link_col) + 1] = link_text
    for col, txt in (extra_links or {}).items():
        if col in df.columns:
            links[list(df.columns).index(col) + 1] = txt
    for i, row in enumerate(df.itertuples(index=False), 3):
        for j, v in enumerate(row, 1):
            if v is None or (isinstance(v, float) and pd.isna(v)):
                v = None
            elif isinstance(v, (np.integer,)):
                v = int(v)
            elif isinstance(v, (np.floating,)):
                v = float(v)
            x = ws.cell(i, j, v); x.font = Font(name=F_NAME, size=10)
            if isinstance(v, (int, float)):
                x.number_format = '#,##0'
        for j, txt in links.items():
            x = ws.cell(i, j)
            if x.value:
                x.hyperlink = str(x.value); x.value = txt
                x.font = Font(name=F_NAME, size=10, color='0563C1', underline='single')
    for j, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = 'A3'
    ws.auto_filter.ref = f'A2:{get_column_letter(n)}{max(2, len(df) + 2)}'
    return ws


def _optlist(d):
    return ', '.join(re.sub(r'\[[^\]]*\]', '', v).strip() for v in d['옵션명'])


def product_summary(R, res):
    rows = []
    for pid, d in R.groupby('상품 번호', sort=False):
        sold = d[d['판정'] == JUDGE_ORDER[0]]; stop = d[d['판정'] == JUDGE_ORDER[1]]
        nf = d[d['판정'] == JUDGE_ORDER[2]]; z = d[d['판정'] == JUDGE_ORDER[3]]
        codes = sorted({c for x in res[pid] for c in x['codes']})
        bases = sorted({(re.match(BASE_RE, c).group(1) if re.match(BASE_RE, c) else c) for c in codes})
        rows.append({'29CM 상품명': d['상품명'].iloc[0], '옵션 수': len(d), '판매중 옵션': int((d['옵션 판매 상태'] == '판매중').sum()),
                     '① 재고 있는데 29CM 품절': len(sold), '① 옵션': _optlist(sold), '① 셀메이트 재고': int(sold['셀메이트재고'].sum()),
                     '② 재고 있는데 29CM 판매중지': len(stop), '② 옵션': _optlist(stop), '② 셀메이트 재고': int(stop['셀메이트재고'].sum()),
                     '③ 셀메이트 옵션 못 찾음': len(nf), '③ 옵션': _optlist(nf), '④ 29CM 판매중·셀메이트 0': len(z),
                     '29CM 재고 합': int(d['재고 수량'].sum()), '셀메이트 재고 합': int(d['셀메이트재고'].fillna(0).sum()),
                     '연결된 셀메이트 스타일': ', '.join(bases[:6]) + (' 외' if len(bases) > 6 else ''),
                     '상품 매칭방식': d['매칭방식'].iloc[0], '29CM 상품번호': str(pid), '링크': d['링크'].iloc[0]})
    return pd.DataFrame(rows)


def write_report(R, A, res, s, out_path, source_label, search=None):
    """search: DataFrame(index=29CM 상품번호, 컬럼 '검색여부') — 검색 노출 점검 후 상품별 시트에 붙임"""
    PS = product_summary(R, res)
    if search is not None:
        PS.insert(1, '29CM 검색', PS['29CM 상품번호'].map(search['검색여부']).fillna('미확인'))
    act = PS[(PS['① 재고 있는데 29CM 품절'] + PS['② 재고 있는데 29CM 판매중지'] + PS['③ 셀메이트 옵션 못 찾음']) > 0].copy()
    act = act.assign(_s=act['① 셀메이트 재고'] + act['② 셀메이트 재고']).sort_values(['_s', '③ 셀메이트 옵션 못 찾음'], ascending=False).drop(columns='_s')
    order = {k: i for i, k in enumerate(JUDGE_ORDER)}
    OC = ['상품명', '브랜드', '옵션명', '옵션 판매 상태', '재고 수량', '셀메이트상품코드', '셀메이트옵션코드', '셀메이트상품명', '셀메이트옵션명',
          '셀메이트재고', '셀메이트판매여부', '판정', '옵션매칭', '매칭방식', '상품 번호', '링크']
    OPT = R.assign(_o=R['판정'].map(order)).sort_values(['_o', '셀메이트재고'], ascending=[True, False])[OC].rename(columns={
        '상품명': '29CM 상품명', '브랜드': '29CM 브랜드', '옵션명': '29CM 옵션명', '옵션 판매 상태': '29CM 옵션 판매상태', '재고 수량': '29CM 재고',
        '셀메이트상품코드': '셀메이트 상품코드', '셀메이트옵션코드': '셀메이트 옵션코드', '셀메이트상품명': '셀메이트 상품명',
        '셀메이트옵션명': '셀메이트 옵션명', '셀메이트재고': '셀메이트 재고', '셀메이트판매여부': '셀메이트 판매여부', '옵션매칭': '옵션 매칭',
        '매칭방식': '상품 매칭방식', '상품 번호': '29CM 상품번호'})
    # 29CM에 옵션 없음: 연결된 셀메이트 스타일 중 재고는 있는데 29CM 어느 옵션과도 연결 안 된 셀메이트 옵션
    used = set(c for cs in R['_codes'] for c in cs)
    linked = set(s[s['옵션코드'].isin(used)]['base'])
    miss = s[s['base'].isin(linked) & ~s['옵션코드'].isin(used) & (s['재고'] > 0)].copy()
    b2n = R.assign(b=R['셀메이트상품코드'].fillna('').str.split(r' \+ ')).explode('b')
    b2n['b'] = b2n['b'].str.extract(BASE_RE)[0].fillna(b2n['b'])
    b2n = b2n.groupby('b')['상품명'].agg(lambda v: ' / '.join(sorted(set(v))[:3]))
    miss['연결된 29CM 상품'] = miss['base'].map(b2n)
    MISS = miss.sort_values('재고', ascending=False)[['상품코드', '상품명', '옵션코드', '옵션명', '재고', '판매여부', '연결된 29CM 상품']].rename(
        columns={'상품코드': '셀메이트 상품코드', '상품명': '셀메이트 상품명', '옵션코드': '셀메이트 옵션코드', '옵션명': '셀메이트 옵션명',
                 '재고': '셀메이트 재고', '판매여부': '셀메이트 판매여부'})
    P = s.drop_duplicates('상품코드').set_index('상품코드')
    CK = A[A['how'].str.contains('확인필요|못 찾음')].copy()
    CK['연결된 셀메이트 상품'] = CK['상품 번호'].map(lambda k: ' + '.join(
        ' / '.join(sorted(set(P.loc[sorted(x['codes']), '상품명']))[:4]) if x['codes'] else '(못 찾음)' for x in res[k]))
    CK['연결된 셀메이트 상품코드'] = CK['상품 번호'].map(lambda k: ' + '.join(', '.join(sorted(x['codes'])[:6]) for x in res[k]))
    CK = CK[['상품명', 'how', '연결된 셀메이트 상품', '연결된 셀메이트 상품코드', '상품 번호']].rename(
        columns={'상품명': '29CM 상품명', 'how': '상품 매칭방식', '상품 번호': '29CM 상품번호'})
    CK['링크'] = PRODUCT_URL + CK['29CM 상품번호'].astype(str)

    wb = Workbook(); wb.remove(wb.active)
    ws = wb.create_sheet('요약')
    ws['A1'] = '29CM × 셀메이트 옵션 재고 점검'; ws['A1'].font = Font(name=F_NAME, bold=True, size=13)
    ws['A2'] = ('29CM 옵션재고 일괄변경 파일(페이지별)을 합친 뒤, 상품명으로 셀메이트 상품을 찾고 옵션(컬러·사이즈)을 1:1로 맞췄습니다.  '
                f'입력: {source_label}')
    ws['A2'].font = Font(name=F_NAME, size=10)
    vc = R['판정'].value_counts()
    lines = [('29CM 상품 수', R['상품 번호'].nunique()), ('29CM 옵션 수', len(R)), ('', None), ('옵션 판정', '옵션 수')]
    lines += [(k, int(vc.get(k, 0))) for k in JUDGE_ORDER]
    lines += [('', None), ('조치가 필요한 29CM 상품 (①②③ 중 하나라도 있음)', len(act)),
              ('셀메이트 재고 있는데 29CM에 옵션 자체가 없음', len(MISS)), ('상품 매칭 확인필요', int(A['how'].str.contains('확인필요').sum()))]
    if search is not None:
        sv = search['검색여부'].value_counts()
        lines += [('', None), ('29CM 검색 (상품)', '상품 수')] + [(k, int(v)) for k, v in sv.items()]
    lines += [('', None), ('상품 매칭방식', '상품 수')] + [(k, int(v)) for k, v in A['how'].value_counts().items()]
    r = 4
    for k, v in lines:
        bold = k in ('옵션 판정', '상품 매칭방식', '29CM 검색 (상품)')
        ws.cell(r, 1, k).font = Font(name=F_NAME, size=10, bold=bold)
        c = ws.cell(r, 2, v); c.font = Font(name=F_NAME, size=10, bold=bold); c.number_format = '#,##0'; r += 1
    r += 1
    for t in ['판정 기준',
              '· 셀메이트 재고 있는데 29CM 품절: 29CM 옵션이 일시품절인데 셀메이트 현재재고가 1개 이상 → 29CM 재고 반영 필요',
              '· 셀메이트 재고 있는데 29CM 판매중지: 29CM 옵션이 판매중지인데 셀메이트 재고가 있음 → 판매 재개 검토',
              '· 셀메이트 옵션 못 찾음: 상품은 찾았지만 같은 컬러·사이즈(기모 등 구분 포함) 옵션이 셀메이트에 없음',
              '· 29CM 판매중 · 셀메이트 재고 0: 29CM에서는 팔리는데 셀메이트 재고가 없음 → 품절 처리 또는 재고 확인',
              '· 셀메이트 (품절)·(이월) 표시와 관계없이 현재재고가 있으면 재고 있음으로 봅니다.',
              '· 세트 상품은 구성품 각각을 셀메이트에서 찾고, 셀메이트 재고는 구성품 중 적은 쪽으로 봅니다.',
              '상품명 매칭 순서',
              '① 컬러 묶음 표기(- 3COL, _3Colors 등)와 끝 컬러명을 떼고 같은 상품명  ② 셀메이트 상품명 = 29CM 상품명 + 컬러',
              '③ 29CM 상품명 단어가 모두 들어 있는 셀메이트 상품(확인필요)  ④ 비슷한 상품명(확인필요)  ⑤ 한글↔영문 상품명 발음·용어 대조(확인필요)',
              '· 같은 상품명이 여러 스타일에 걸리면 옵션 컬러로 해당 스타일을 고르고, 같은 스타일 코드(앞 8자리)의 다른 컬러도 함께 봅니다.']:
        ws.cell(r, 1, t).font = Font(name=F_NAME, size=10, bold=t in ('판정 기준', '상품명 매칭 순서')); r += 1
    ws.column_dimensions['A'].width = 62; ws.column_dimensions['B'].width = 12

    srch = ['29CM 검색'] if search is not None else []
    write_sheet(wb, '상품별_한눈에보기', act,
                '29CM 상품 1개 = 1줄. 셀메이트에는 재고가 있는데 29CM에서 고를 수 없는 옵션이 있는 상품만 모았습니다. 작업은 이 시트부터 보시면 됩니다.\n'
                '① 29CM 일시품절인데 셀메이트 재고 있음  ② 29CM 판매중지인데 셀메이트 재고 있음  ③ 셀메이트에서 같은 컬러·사이즈를 못 찾은 옵션.  ①+② 셀메이트 재고가 많은 순.'
                + ('\n29CM 검색: "브랜드명 + 상품명"으로 29CM에서 검색했을 때 결과 화면에 나왔는지 (29CM_검색노출점검 엑셀 참고)' if srch else ''),
                [42] + [14] * len(srch) + [7, 8, 9, 40, 9, 9, 40, 9, 9, 30, 9, 9, 9, 22, 22, 12, 10])
    write_sheet(wb, '옵션별_대조', OPT,
                f'29CM 옵션 전체({len(OPT):,}개)와 연결된 셀메이트 옵션. 왼쪽 5개 열은 29CM 파일 값(상품명·브랜드·옵션명·옵션 판매상태·재고 수량), 그 오른쪽은 셀메이트 값입니다.\n'
                '판정 순서로 정렬했습니다(재고 있는데 29CM 품절 → 판매중지 → 못 찾음 → 29CM 판매중·셀메이트 0 → 정상). 세트 상품의 셀메이트 값은 "구성품1 + 구성품2"로 표시합니다.',
                [40, 11, 34, 10, 8, 16, 18, 36, 16, 9, 9, 26, 18, 20, 11, 10])
    write_sheet(wb, '29CM에_옵션없음', MISS,
                '29CM에 연결된 셀메이트 스타일 중, 셀메이트 재고는 있는데 29CM 어느 옵션과도 연결되지 않은 셀메이트 옵션입니다.\n'
                '29CM에 해당 컬러·사이즈가 등록돼 있지 않거나, 29CM 옵션명이 셀메이트와 달라 연결되지 않은 경우입니다. 재고 많은 순.',
                [16, 42, 18, 18, 9, 9, 50], link_col=None)
    write_sheet(wb, '상품매칭_확인필요', CK,
                '상품명이 정확히 같지 않아 비슷한 이름·단어 포함·한글↔영문 대조로 연결한(또는 못 찾은) 29CM 상품입니다. 연결된 셀메이트 상품이 맞는지 확인해 주세요.',
                [50, 30, 60, 40, 12, 10])
    wb.save(out_path)
    return PS, act
