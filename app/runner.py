"""
점검 실행 모듈 (화면·명령줄 공용)

prepare(stock_copy_path, sellmate_path, out_dir)
    재고 관리(v2) 복사본 + 셀메이트 재고목록
    → ① 옵션 재고 점검 엑셀 (셀메이트 재고 있는데 무신사 옵션 품절 등)
      ② 판매중 상품 전체 검색어 + 브라우저 검색 스크립트 + 검색 노출 엑셀(검색 전)
apply_search(search_text, out_dir)     브라우저 검색 결과 → 두 엑셀에 반영
"""
import os, sys, io, json, datetime, traceback, runpy, contextlib, time

# 단계별 스크립트(src/step*.py)가 쓰는 라이브러리 — PyInstaller가 exe에 포함하도록 미리 import
import pandas, numpy, openpyxl, difflib, re, math, collections  # noqa: F401
import urllib.parse  # noqa: F401
from openpyxl.workbook.properties import CalcProperties  # noqa: F401

BASE = getattr(sys, '_MEIPASS', os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
for p in (BASE, os.path.join(BASE, 'src')):
    if p not in sys.path:
        sys.path.insert(0, p)
import config as cfg      # noqa: E402
import search_check as sc  # noqa: E402

SCRIPT_TEMPLATE = os.path.join(BASE, 'search', 'musinsa_brand_search.js')


def _console(msg):
    out = sys.__stdout__
    if out:
        out.write(msg + '\n'); out.flush()


def _paths(out_dir):
    work = os.path.join(out_dir, 'work')
    os.makedirs(work, exist_ok=True)
    today = datetime.date.today().strftime('%Y%m%d')
    return {
        'listings': os.path.join(work, 'listings.pkl'),
        'meta': os.path.join(work, 'meta.json'),
        'script': os.path.join(out_dir, 'musinsa_brand_search_ready.js'),
        'search': os.path.join(out_dir, 'search_result.txt'),
        'report': os.path.join(out_dir, f'무신사_검색노출점검_{today}.xlsx'),
        'option_report': os.path.join(out_dir, f'무신사_옵션재고점검_{today}.xlsx'),
        'opt_in': os.path.join(work, 'musinsa_option_from_copy.xlsx'),
        'sel_in': os.path.join(work, 'musinsa_selling_from_copy.xlsx'),
    }


OPTION_STEPS = [
    ('step2_name_check.py',    '    상품명 대조 (스타일넘버·상품명)'),
    ('step3_option_match.py',  '    무신사 옵션 ↔ 셀메이트 옵션 매칭'),
    ('step4_option_judge.py',  '    옵션 판정 (품절·판매 제한·MFS 등)'),
    ('step6_option_report.py', '    옵션 재고 점검 엑셀 작성'),
]


def _run_steps(steps, log):
    cwd = os.getcwd()
    os.chdir(BASE)
    try:
        for f, label in steps:
            t = time.time()
            log(f'{label} ...')
            with contextlib.redirect_stdout(io.StringIO()):   # 단계 스크립트의 디버그 출력은 숨김
                runpy.run_path(os.path.join('src', f), run_name='__main__')
            log(f'      완료 ({time.time() - t:.0f}초)')
    finally:
        os.chdir(cwd)


def _write_option_inputs(df, P):
    """재고 관리(v2) 복사본 → 옵션 대조 단계가 읽는 형태(옵션 목록 / 상품 목록)로 저장"""
    df.to_excel(P['opt_in'], index=False)
    g = df.groupby('상품번호', sort=False)
    sel = pandas.DataFrame({'상품 번호': list(g.groups), '상품명': g['상품명'].first().values,
                            '스타일넘버': g['스타일넘버'].first().values, '판매 가능 재고': g['총합 판매가능'].sum().values})
    with pandas.ExcelWriter(P['sel_in']) as w:
        sel.to_excel(w, sheet_name='상품정보', index=False, startrow=1)


def _set_cfg(out_dir, meta, P, search_path=None):
    cfg.load_settings()
    cfg.set_output(out_dir)
    cfg.SELLMATE_STOCK = meta['sellmate_path']
    cfg.MUSINSA_OPTION = P['opt_in']
    cfg.MUSINSA_SELLING = P['sel_in']
    cfg.MUSINSA_ALL = ''
    cfg.SOURCE_LABEL = f"{meta['source']} × {meta['sellmate']}"
    cfg.SEARCH_RESULT = search_path if search_path and os.path.exists(search_path) else os.path.join(cfg.WORK, '__no_search__.txt')


def prepare(stock_copy_path, sellmate_path, out_dir, log=_console):
    if not stock_copy_path or not os.path.exists(stock_copy_path):
        raise FileNotFoundError('무신사 재고 관리(v2) 복사본 엑셀 파일을 선택해 주세요.')
    if not sellmate_path or not os.path.exists(sellmate_path):
        raise FileNotFoundError('셀메이트 재고목록 파일(stk_stockList_*.csv)을 선택해 주세요.')
    cfg.load_settings()
    P = _paths(out_dir)
    log(f'브랜드: {cfg.BRAND_KR}')
    log('1/4 재고 관리(v2) 복사본 읽는 중 ...')
    df = sc.read_stock_table(stock_copy_path)
    log(f'      옵션 {len(df):,}줄 (판매중), 필요한 열 {len(df.columns)}개만 사용')
    L, excluded = sc.make_listings(df, cfg.EXCLUDE_CODE_PATTERN)
    log(f'      판매중 상품 {len(L):,}개' + (f' (다른 브랜드 {excluded}개 제외)' if excluded else ''))
    meta = {'source': os.path.basename(stock_copy_path), 'excluded': int(excluded), 'brand': cfg.BRAND_KR,
            'sellmate': os.path.basename(sellmate_path), 'sellmate_path': os.path.abspath(sellmate_path)}
    with open(P['meta'], 'w', encoding='utf-8') as f:
        json.dump(meta, f, ensure_ascii=False)
    L.to_pickle(P['listings'])

    log('2/4 셀메이트 재고와 무신사 옵션 대조 중 ...')
    _write_option_inputs(df, P)
    _set_cfg(out_dir, meta, P)
    L[['상품번호', '상품명', '검색어']].rename(columns={'상품번호': '무신사 상품번호', '상품명': '무신사 상품명'}).to_pickle(
        os.path.join(cfg.WORK, 'kw.pkl'))
    _run_steps(OPTION_STEPS, log)
    log(f'      → {os.path.basename(P["option_report"])}')

    # 검색 대상: 옵션 재고 점검에서 조치가 필요하다고 나온 상품(상품별_한눈에보기)만
    hv = pandas.read_excel(P['option_report'], '상품별_한눈에보기', header=1, dtype=str)
    targets = set(hv.iloc[:, 1].dropna().astype(str))
    L = L[L['상품번호'].astype(str).isin(targets)].reset_index(drop=True)
    L.to_pickle(P['listings'])
    log(f'3/4 브라우저 검색 스크립트 만드는 중 (상품별_한눈에보기 상품 {len(L):,}개) ...')
    sc.write_search_script(L, cfg.BRAND_KR, SCRIPT_TEMPLATE, P['script'])
    log('4/4 검색 노출 엑셀(검색 전) 만드는 중 ...')
    sc.write_report(L, None, cfg.BRAND_KR, P['report'], meta['source'], excluded)
    mins = max(1, round(len(L) * 0.55 / 60))
    log(f'준비 완료. 옵션 재고 점검 엑셀은 이미 만들어졌습니다.')
    log(f'이제 ② 무신사 검색을 실행하세요 (상품 {len(L):,}개, 약 {mins}분).')
    return P


def apply_search(search_text, out_dir, log=_console):
    P = _paths(out_dir)
    if not os.path.exists(P['listings']):
        raise FileNotFoundError('먼저 [① 옵션 재고 점검 + 검색 준비]를 실행해 주세요.')
    import pandas as pd
    L = pd.read_pickle(P['listings'])
    with open(P['meta'], encoding='utf-8') as f:
        meta = json.load(f)
    sr = sc.parse_search_result(search_text)
    with open(P['search'], 'w', encoding='utf-8') as f:
        f.write(search_text.strip())
    R = sc.write_report(L, sr, meta['brand'], P['report'], meta['source'], meta['excluded'])
    if meta.get('sellmate_path') and os.path.exists(os.path.join(out_dir, 'work', 'judged.pkl')):
        _set_cfg(out_dir, meta, P, P['search'])
        _run_steps([('step6_option_report.py', '옵션 재고 점검 엑셀에 검색 결과 반영')], log)
    n_ok = int(R['검색여부'].str.startswith('검색됨').sum())
    n_miss = int((R['검색여부'] == '검색 안 됨').sum())
    n_unk = int((R['검색여부'] == '미확인').sum())
    log(f'검색됨 {n_ok:,}개 / 검색 안 됨 {n_miss:,}개' + (f' / 미확인 {n_unk:,}개' if n_unk else ''))
    log(f'결과 엑셀: {P["report"]}')
    log(f'          {P["option_report"]}')
    return P['report']


def main():
    """명령줄: python app/runner.py <재고관리_복사본.xlsx> <셀메이트_재고목록.csv> [결과폴더] [검색결과.txt]"""
    a = sys.argv[1:]
    if len(a) < 2:
        print(main.__doc__); sys.exit(1)
    out = os.path.abspath(a[2]) if len(a) > 2 else os.path.join(cfg.ROOT, 'output')
    try:
        P = prepare(os.path.abspath(a[0]), os.path.abspath(a[1]), out)
        print(' → 옵션 재고 점검:', P['option_report'])
        if len(a) > 3:
            with open(a[3], encoding='utf-8') as f:
                apply_search(f.read(), out)
        else:
            print(' → 검색 스크립트:', P['script'])
        print(' → 결과 엑셀:', P['report'])
    except Exception:
        traceback.print_exc(); sys.exit(1)


if __name__ == '__main__':
    main()
