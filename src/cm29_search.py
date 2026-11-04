"""
29CM 검색 노출 점검 — 크롬을 띄워 29CM 검색창에 "브랜드명 + 상품명"을 직접 입력하고, 검색 결과 화면에 우리 상품이 보이는지 확인

흐름 (검색어 1개마다)
  1. 화면 위쪽 검색창(검색 결과 화면) 또는 오른쪽 위 돋보기 버튼(검색 레이어)을 마우스로 클릭
  2. "브랜드명 + 상품명"(컬러·_3Colors 같은 컬러 묶음 표기 제외)을 입력하고 Enter
  3. 검색 결과 화면(/store/search)의 상품 카드를 읽음 — 카드의 상품번호(catalog/번호)와 브랜드명(카드의 브랜드 링크)
  4. 찾는 상품이 우리 브랜드로 안 보이면 화면을 캡처해 '29CM_미노출_캡처' 폴더에 저장

무신사 검색(ui_search.py)과 같은 크롬 원격 제어(Browser)를 씁니다. 확인 범위는 비로그인 웹 검색 결과 첫 페이지(최대 50개)입니다.
"""
import os, re, json, time, shutil, urllib.parse
import pandas as pd

from ui_search import Browser, _safe, retry_keyword, RETRY_LABEL
import search_check as sc
from cm29_match import write_sheet, F_NAME, PRODUCT_URL
from openpyxl import Workbook
from openpyxl.styles import Font

BASE_URL = 'https://www.29cm.co.kr'
HOME_PATH = '/'
SEARCH_PATH = '/store/search'

# 검색창: 검색 결과 화면에서는 위쪽에 보이는 입력칸, 그 외 화면에서는 오른쪽 위 돋보기 버튼(검색 레이어 열기)
JS_TRIGGER = r"""(() => {
  window.scrollTo(0, 0);
  const vis = e => { const r = e.getBoundingClientRect(); return r.width > 40 && r.height > 0 && r.y >= 0 && r.y < 400; };
  const inp = [...document.querySelectorAll('input[name="keyword"]')].find(vis);
  if (inp) { const r = inp.getBoundingClientRect(); return { kind: 'input', x: r.x + r.width / 2, y: r.y + r.height / 2 }; }
  const btns = [...document.querySelectorAll('button, a')].filter(b => {
    const r = b.getBoundingClientRect();
    return r.width > 0 && r.y < 200 && r.x > window.innerWidth - 400 && b.querySelector('svg') && !b.textContent.trim() && !b.getAttribute('href');
  });
  const lab = btns.find(b => /검색|search/i.test(b.getAttribute('aria-label') || ''));
  const el = lab || btns.sort((a, b) => b.getBoundingClientRect().x - a.getBoundingClientRect().x)[0];
  if (!el) return null;
  const r = el.getBoundingClientRect();
  return { kind: 'button', x: r.x + r.width / 2, y: r.y + r.height / 2 };
})()"""

JS_FOCUSED = r"""(() => {
  const a = document.activeElement;
  if (a && a.tagName === 'INPUT' && !a.readOnly) { a.select(); return true; }
  return false;
})()"""

JS_STATE = r"""(() => {
  const kw = new URLSearchParams(location.search).get('keyword');
  const t = document.body ? document.body.innerText : '';
  const m = t.match(/총\s*([\d,]+)\s*개\s*상품/);
  const n = document.querySelectorAll('a[href*="/catalog/"]').length;
  return { path: location.pathname, kw, total: m ? Number(m[1].replace(/,/g, '')) : null, n, ready: document.readyState };
})()"""

JS_COLLECT = r"""((BRAND) => {
  const got = new Map();
  let idx = 0;
  for (const li of document.querySelectorAll('li')) {
    const p = li.querySelector('a[href*="/catalog/"]');
    if (!p) continue;
    const m = p.href.match(/catalog\/(\d+)/);
    if (!m || got.has(m[1])) continue;
    const b = li.querySelector('a[href*="/store/brand/"]');
    const bm = b ? b.href.match(/brand\/(\d+)/) : null;
    const lines = (li.innerText || '').split('\n').map(s => s.trim()).filter(Boolean);
    const brand = b ? b.textContent.trim() : '';
    const name = lines.find(s => s !== brand && !/^\d+%$|^[\d,]+$|원$|^(무료배송|품절|쿠폰|단독|스페셜)$/.test(s)) || '';
    got.set(m[1], { id: m[1], idx: ++idx, brand, brandId: bm ? bm[1] : '', bt: brand === BRAND, name, soldout: lines.includes('품절') });
  }
  const t = document.body.innerText;
  const mm = t.match(/총\s*([\d,]+)\s*개\s*상품/);
  return { items: [...got.values()].slice(0, 60), total: mm ? Number(mm[1].replace(/,/g, '')) : null };
})"""


def _same_kw(a, b):
    n = lambda s: re.sub(r'\s+', ' ', str(s or '').replace('+', ' ')).strip().lower()
    return n(a) == n(b)


def search_once(br, keyword, base=BASE_URL):
    """검색창에 입력해서 검색. 실패하면 검색 결과 주소로 직접 이동(예비). 반환: 'ui' | 'url'"""
    how = 'url'
    pos = br.eval(JS_TRIGGER)
    if pos:
        time.sleep(0.2)
        pos = br.eval(JS_TRIGGER) or pos
        br.click(pos['x'], pos['y'])
        for _ in range(15):
            time.sleep(0.2)
            if br.eval(JS_FOCUSED):
                br.type_text(keyword); time.sleep(0.25); br.enter(); how = 'ui'; break
    if how == 'url':
        br.call('Page.navigate', url=base + SEARCH_PATH + '?' + urllib.parse.urlencode(
            {'keyword': keyword, 'sort': 'RECOMMENDED', 'page': 1}, quote_via=urllib.parse.quote))
    t = time.time(); last = None
    while time.time() - t < 15:
        time.sleep(0.3)
        try:
            s = br.eval(JS_STATE)
        except Exception:
            continue
        if s and s['path'].startswith(SEARCH_PATH) and _same_kw(s['kw'], keyword) and s['ready'] == 'complete' \
                and s['total'] is not None and (s['n'] or s['total'] == 0):
            if last == (s['total'], s['n']):   # 두 번 연속 같으면 다 그려진 것
                break
            last = (s['total'], s['n'])
    time.sleep(0.4)
    return how


def collect(br, brand):
    return br.eval(JS_COLLECT + f'({json.dumps(brand)})')


# ------------------------------------------------------------------ 검색 목록
def make_listings(R):
    """29CM 옵션별 결과 → 상품 1줄 (판매중 옵션이 있는 상품만 검색)"""
    g = R.groupby('상품 번호', sort=False)
    L = pd.DataFrame({
        '상품번호': list(g.groups),
        '상품명': g['상품명'].first().values,
        '옵션수': g.size().values,
        '판매중옵션수': g['옵션 판매 상태'].apply(lambda v: int((v == '판매중').sum())).values,
        '29CM재고': g['재고 수량'].sum().values,
        '셀메이트재고': g['셀메이트재고'].apply(lambda v: int(v.fillna(0).sum())).values,
    })
    L['상품번호'] = L['상품번호'].astype(str)
    en = sc.color_words(L['상품명'])
    L['검색어'] = L['상품명'].map(lambda v: sc.clean_name(re.sub(r'\s+[+&]\s+', ' ', v), en))
    return L[L['판매중옵션수'] > 0].reset_index(drop=True)


# ------------------------------------------------------------------ 자동 검색
def run(L, brand, out_dir, log=print, stop=None, progress=None, chrome_path=None, headless=False, delay=0.5, base=BASE_URL,
        work_name='work_29cm', cap_name='29CM_미노출_캡처', drop_words=None):
    """반환: {상품번호: {total, rank, ours, empty, soldout, cap, seen, retry}}  (중간에 멈춰도 다시 누르면 이어서)
    상품이 안 나오면 drop_words(UNISEX·유니섹스 등)를 뺀 검색어로 한 번 더 검색합니다."""
    work = os.path.join(out_dir, work_name); cap_dir = os.path.join(out_dir, cap_name)
    os.makedirs(work, exist_ok=True); os.makedirs(cap_dir, exist_ok=True)
    prog_path = os.path.join(work, 'ui_progress.json')
    try:
        prog = json.load(open(prog_path, encoding='utf-8'))
    except Exception:
        prog = {}
    groups = {}
    for no, name, kw in zip(L['상품번호'], L['상품명'], L['검색어']):
        groups.setdefault(f'{brand} {kw}'.strip(), []).append((no, name))
    todo = [k for k in groups if k not in prog]
    log(f'검색어 {len(groups):,}개 (상품 {len(L):,}개)' + (f' — 이전에 한 {len(groups) - len(todo):,}개는 건너뜀' if len(todo) < len(groups) else ''))
    brand_ids = set()
    if todo:
        br = Browser(os.path.join(work, 'chrome_profile'), chrome_path, headless)
        try:
            br.goto(base + HOME_PATH); time.sleep(1.5)
            t0 = time.time(); fallback = 0
            for i, kw in enumerate(todo, 1):
                if stop is not None and stop.is_set():
                    log('중지했습니다. 다시 누르면 이어서 검색합니다.'); break
                try:
                    how = search_once(br, kw, base); res = collect(br, brand)
                except Exception as e:
                    log(f'  [{kw}] 오류: {e} → 한 번 더 시도'); time.sleep(2)
                    br.goto(base + HOME_PATH); time.sleep(1)
                    how = search_once(br, kw, base); res = collect(br, brand)
                fallback += how == 'url'
                for it in res['items']:
                    if it['bt'] and it['brandId']:
                        brand_ids.add(it['brandId'])
                ours = [it for it in res['items'] if it['bt'] or (it['brandId'] and it['brandId'] in brand_ids)]
                total = res['total'] if res['total'] is not None else len(res['items'])
                rec = {'total': total, 'ours': len(ours), 'rank': {}, 'soldout': {}, 'cap': {},
                       'seen': [f"{it['id']}:{it['brand']} {it['name']}" for it in res['items'][:10]]}
                ids = [it['id'] for it in res['items']]
                miss = []
                for no, name in groups[kw]:
                    hit = next((it for it in ours if it['id'] == no), None)
                    if hit:
                        rec['rank'][no] = ids.index(no) + 1
                        if hit['soldout']:
                            rec['soldout'][no] = True
                    else:
                        miss.append((no, name))
                rec['retry'] = {}
                kw2 = retry_keyword(kw, drop_words, brand) if miss else None
                if kw2:
                    try:
                        search_once(br, kw2, base); r2 = collect(br, brand)
                    except Exception as e:
                        log(f'  [{kw2}] 재검색 오류: {e}'); r2 = None
                    if r2:
                        for it in r2['items']:
                            if it['bt'] and it['brandId']:
                                brand_ids.add(it['brandId'])
                        ours2 = [it for it in r2['items'] if it['bt'] or (it['brandId'] and it['brandId'] in brand_ids)]
                        ids2 = [it['id'] for it in r2['items']]
                        tot2 = r2['total'] if r2['total'] is not None else len(r2['items'])
                        still = []
                        for no, name in miss:
                            hit = next((it for it in ours2 if it['id'] == no), None)
                            if hit:
                                rec['rank'][no] = ids2.index(no) + 1
                                rec['retry'][no] = {'kw': kw2, 'total': tot2}
                                if hit['soldout']:
                                    rec['soldout'][no] = True
                            else:
                                still.append((no, name))
                        miss = still
                if miss:
                    p = os.path.join(cap_dir, f'{miss[0][0]}_{_safe(miss[0][1])}.png')
                    br.screenshot(p)
                    for no, name in miss:
                        q = os.path.join(cap_dir, f'{no}_{_safe(name)}.png')
                        if q != p:
                            shutil.copyfile(p, q)
                        rec['cap'][no] = os.path.relpath(q, out_dir)
                prog[kw] = rec
                if i % 10 == 0 or i == len(todo):
                    json.dump(prog, open(prog_path, 'w', encoding='utf-8'), ensure_ascii=False)
                    el = time.time() - t0; left = el / i * (len(todo) - i)
                    log(f'  진행 {i:,}/{len(todo):,}  (남은 시간 약 {left / 60:.0f}분)')
                if progress:
                    progress(i, len(todo))
                time.sleep(delay)
            json.dump(prog, open(prog_path, 'w', encoding='utf-8'), ensure_ascii=False)
            if fallback:
                log(f'  ※ {fallback}개는 검색창 입력이 안 돼 검색 결과 주소로 직접 열어 확인했습니다.')
        finally:
            br.close()
    out = {}
    for kw, items in groups.items():
        rec = prog.get(kw)
        if not rec:
            continue
        for no, _ in items:
            rt = rec.get('retry', {}).get(no)
            out[no] = {'검색어': kw, 'total': rt['total'] if rt else rec['total'], 'rank': rec['rank'].get(no), 'ours': rec['ours'],
                       'retry': rt['kw'] if rt else '',
                       'soldout': rec.get('soldout', {}).get(no, False), 'cap': rec['cap'].get(no), 'seen': ' / '.join(rec['seen'][:5])}
    return out


# ------------------------------------------------------------------ 결과 엑셀
def judge(L, found):
    rows = []
    for _, r in L.iterrows():
        f = found.get(r['상품번호'])
        if not f:
            st, note = '미확인', '아직 검색 안 함'
        elif f['rank'] and f.get('retry'):
            st = RETRY_LABEL
            note = (f'원래 검색어로는 안 나오고, UNISEX·유니섹스를 뺀 "{f["retry"]}"로 검색하면 나옴 → 상품명에서 UNISEX·유니섹스 제외 검토'
                    + (' (카드에 "품절" 표시)' if f['soldout'] else ''))
        elif f['rank']:
            st = '검색됨' + (' (품절 표시)' if f['soldout'] else '')
            note = '검색 결과에는 나오지만 카드에 "품절"로 표시됨' if f['soldout'] else ''
        else:
            st = '검색 안 됨'
            if f['total'] == 0:
                note = '검색 결과 0건 → 상품명·검색 키워드 확인 필요'
            elif r['29CM재고'] == 0:
                note = '29CM 재고 0 (전 옵션 품절) → 검색에서 빠졌을 수 있음'
            elif f['ours'] == 0:
                note = '검색 결과에 우리 브랜드 상품이 없음 → 29CM 전시 상태 확인 필요'
            else:
                note = '검색 결과에 우리 브랜드의 다른 상품만 나옴 → 29CM 전시 상태 확인 필요'
        rows.append({'검색여부': st, '29CM 상품번호': r['상품번호'], '29CM 상품명': r['상품명'],
                     '검색어': (f or {}).get('검색어', ''), '재검색어': (f or {}).get('retry', ''), '검색 결과 수': (f or {}).get('total'),
                     '화면 순위': (f or {}).get('rank'), '판매중 옵션 수': r['판매중옵션수'], '옵션 수': r['옵션수'],
                     '29CM 재고 합': r['29CM재고'], '셀메이트 재고 합': r['셀메이트재고'], '비고': note,
                     '검색 결과 앞쪽 상품': (f or {}).get('seen', ''),
                     '검색 링크': BASE_URL + SEARCH_PATH + '?' + urllib.parse.urlencode({'keyword': (f or {}).get('검색어', '')}, quote_via=urllib.parse.quote) if f else '',
                     '링크': PRODUCT_URL + r['상품번호'], '캡처': (f or {}).get('cap') or ''})
    return pd.DataFrame(rows)


def write_report(S, out_path, brand, source_label):
    order = {'검색 안 됨': 0, RETRY_LABEL: 1, '검색됨 (품절 표시)': 2, '미확인': 3, '검색됨': 4}
    S = S.assign(_o=S['검색여부'].map(order)).sort_values(['_o', '셀메이트 재고 합'], ascending=[True, False]).drop(columns='_o')
    wb = Workbook(); wb.remove(wb.active)
    ws = wb.create_sheet('요약')
    ws['A1'] = '29CM 검색 노출 점검'; ws['A1'].font = Font(name=F_NAME, bold=True, size=13)
    ws['A2'] = f'29CM 검색창에 "{brand} + 상품명"(컬러·컬러 묶음 표기 제외)을 입력해, 검색 결과 첫 화면(최대 50개)에 우리 브랜드 상품으로 나오는지 확인했습니다.  입력: {source_label}'
    ws['A2'].font = Font(name=F_NAME, size=10)
    r = 4
    for k, v in [('판매중 상품 (판매중 옵션이 1개 이상)', len(S))] + [(k, int((S['검색여부'] == k).sum())) for k in order] + \
                [('', None), ('검색 안 됨 — 원인', '상품 수')] + [(k, int(v)) for k, v in S[S['검색여부'] == '검색 안 됨']['비고'].value_counts().items()]:
        ws.cell(r, 1, k).font = Font(name=F_NAME, size=10, bold=k == '검색 안 됨 — 원인')
        c = ws.cell(r, 2, v); c.font = Font(name=F_NAME, size=10); c.number_format = '#,##0'; r += 1
    r += 1
    for t in ['· 검색됨: 검색 결과 화면에 그 상품번호가 우리 브랜드(카드의 브랜드명)로 나옴',
              '· 검색 안 됨: 결과 화면 첫 페이지에 그 상품이 없음 → "캡처 보기"로 화면 확인',
              '· 확인 범위: 로그인하지 않은 29CM 웹 검색, 추천순 첫 페이지. 앱·카테고리·브랜드관 노출은 확인하지 않습니다.',
              '· 같은 검색어(예: 컬러별로 따로 등록된 상품)는 한 번만 검색합니다.',
              f'· {RETRY_LABEL}: 원래 검색어로는 안 나와서 UNISEX·유니섹스를 빼고 다시 검색하니 나온 상품 (재검색어 열 참고).']:
        ws.cell(r, 1, t).font = Font(name=F_NAME, size=10); r += 1
    ws.column_dimensions['A'].width = 70; ws.column_dimensions['B'].width = 12
    W = [22, 12, 42, 40, 34, 9, 8, 9, 7, 9, 9, 44, 50, 10, 10, 10]
    links = {'검색 링크': '검색 열기', '캡처': '캡처 보기'}
    write_sheet(wb, '검색_미노출', S[S['검색여부'] == '검색 안 됨'],
                '29CM 판매중 상품 중, "브랜드명 + 상품명"으로 검색했을 때 결과 화면에 안 나오는 상품입니다. 작업은 이 시트부터 보시면 됩니다.\n'
                '"캡처 보기"로 검색했을 때의 화면을, "검색 열기"로 지금 검색 결과를 볼 수 있습니다. 셀메이트 재고가 많은 순.',
                W, extra_links=links)
    write_sheet(wb, '전체_검색결과', S,
                '판매중 상품 전체의 검색 결과. 검색 안 됨 → 재검색으로 검색됨 → 검색됨(품절 표시) → 미확인 → 검색됨 순.', W, extra_links=links)
    wb.save(out_path)
    return S
