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

    # 이전 검색 진행 기록·캡처 초기화 (새 파일로 다시 점검)
    for p in (os.path.join(cfg.WORK, 'ui_progress.json'),):
        if os.path.exists(p):
            os.remove(p)
    cap = os.path.join(out_dir, '미노출_캡처')
    if os.path.isdir(cap):
        import shutil; shutil.rmtree(cap, ignore_errors=True)
    log(f'3/4 검색 목록 만드는 중 (판매중 상품 {len(L):,}개, 검색어 {L["검색어"].nunique():,}개) ...')
    sc.write_search_script(L, cfg.BRAND_KR, SCRIPT_TEMPLATE, P['script'])   # 예비(수동) 콘솔 방식용
    log('4/4 검색 노출 엑셀(검색 전) 만드는 중 ...')
    sc.write_report(L, None, cfg.BRAND_KR, P['report'], meta['source'], excluded)
    log(f'옵션 재고 점검 완료 → {os.path.basename(P["option_report"])}')
    return P


def auto_search(out_dir, log=_console, stop=None, progress=None, chrome_path=None, headless=False, base=None):
    """크롬을 띄워 무신사 검색창에 직접 검색 → 결과 엑셀 완성 (중간에 멈추면 다음에 이어서)"""
    import ui_search
    P = _paths(out_dir)
    if not os.path.exists(P['listings']):
        raise FileNotFoundError('먼저 옵션 재고 점검(①)을 실행해 주세요.')
    L = pandas.read_pickle(P['listings'])
    with open(P['meta'], encoding='utf-8') as f:
        meta = json.load(f)
    log('무신사 자동 검색: 크롬을 띄웁니다. 검색이 끝날 때까지 그 크롬 창은 건드리지 말아 주세요 (최소화는 괜찮습니다).')
    kw = dict(log=log, stop=stop, progress=progress, chrome_path=chrome_path, headless=headless,
              drop_words=getattr(cfg, 'SEARCH_DROP_WORDS', ['UNISEX', '유니섹스']))
    if base:
        kw['base'] = base
    text, caps = ui_search.run(L, meta['brand'], out_dir, **kw)
    if not text:
        log('검색한 상품이 없습니다.'); return None
    n_cap = len(caps)
    if n_cap:
        log(f'검색 안 된 상품 화면 캡처 {n_cap:,}개 → {os.path.join(out_dir, "미노출_캡처")}')
    return apply_search(text, out_dir, log, caps)


def apply_search(search_text, out_dir, log=_console, caps=None):
    P = _paths(out_dir)
    if not os.path.exists(P['listings']):
        raise FileNotFoundError('먼저 [① 옵션 재고 점검 + 검색 준비]를 실행해 주세요.')
    L = pandas.read_pickle(P['listings'])
    with open(P['meta'], encoding='utf-8') as f:
        meta = json.load(f)
    sr = sc.parse_search_result(search_text)
    if caps:
        sr['캡처'] = pandas.Series(caps)
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


# ================================================================== 29CM
def _paths_29cm(out_dir):
    work = os.path.join(out_dir, 'work_29cm')
    os.makedirs(work, exist_ok=True)
    today = datetime.date.today().strftime('%Y%m%d')
    return {
        'work': work,
        'state': os.path.join(work, 'match.pkl'),
        'meta': os.path.join(work, 'meta.json'),
        'option_report': os.path.join(out_dir, f'29CM_옵션재고점검_{today}.xlsx'),
        'report': os.path.join(out_dir, f'29CM_검색노출점검_{today}.xlsx'),
        'cap': os.path.join(out_dir, '29CM_미노출_캡처'),
    }


def prepare_29cm(cm_paths, sellmate_path, out_dir, log=_console):
    """29CM 옵션재고 일괄변경 파일(여러 개) + 셀메이트 재고목록 → 29CM_옵션재고점검 엑셀 + 검색 목록"""
    import pickle, shutil
    import cm29_match as cm, cm29_search as cs
    if isinstance(cm_paths, str):
        cm_paths = [p.strip() for p in cm_paths.split(';') if p.strip()]
    cm_paths = [p for p in cm_paths if p]
    missing = [p for p in cm_paths if not os.path.exists(p)]
    if not cm_paths or missing:
        raise FileNotFoundError('29CM 옵션재고 일괄변경 파일을 선택해 주세요.' + (f'\n없는 파일: {missing[0]}' if missing else ''))
    if not sellmate_path or not os.path.exists(sellmate_path):
        raise FileNotFoundError('셀메이트 재고목록 파일(stk_stockList_*.csv)을 선택해 주세요.')
    cfg.load_settings()
    P = _paths_29cm(out_dir)
    log(f'브랜드: {cfg.BRAND_KR}  (29CM)')
    log(f'1/4 29CM 파일 {len(cm_paths)}개 읽는 중 ...')
    a, dup = cm.read_29cm(cm_paths)
    log(f'      옵션 {len(a):,}줄, 상품 {a["상품 번호"].nunique():,}개' + (f' (겹친 옵션 {dup:,}줄은 한 번만)' if dup else ''))
    s = cm.read_sellmate(sellmate_path, cfg.TARGET_CODE_PATTERN)
    log(f'      셀메이트 옵션 {len(s):,}줄')
    log('2/4 상품명으로 셀메이트 상품 찾고 옵션(컬러·사이즈) 대조 중 ... (1분 정도)')
    R, A, res = cm.match(a, s, log=log)
    src = f'29CM 파일 {len(cm_paths)}개 × {os.path.basename(sellmate_path)}'
    with open(P['state'], 'wb') as f:
        pickle.dump((R, A, res, s), f)
    meta = {'source': src, 'brand': cfg.BRAND_KR, 'files': [os.path.abspath(p) for p in cm_paths],
            'sellmate_path': os.path.abspath(sellmate_path)}
    with open(P['meta'], 'w', encoding='utf-8') as f:
        json.dump(meta, f, ensure_ascii=False)
    log('3/4 옵션 재고 점검 엑셀 작성 중 ...')
    cm.write_report(R, A, res, s, P['option_report'], src)
    vc = R['판정'].value_counts()
    log(f'      셀메이트 재고 있는데 29CM 품절 {int(vc.get(cm.JUDGE_ORDER[0], 0)):,}개 / 판매중지 {int(vc.get(cm.JUDGE_ORDER[1], 0)):,}개 (옵션)')
    log(f'      → {os.path.basename(P["option_report"])}')
    # 새 파일로 다시 점검: 이전 검색 진행 기록·캡처 초기화
    pp = os.path.join(P['work'], 'ui_progress.json')
    if os.path.exists(pp):
        os.remove(pp)
    if os.path.isdir(P['cap']):
        shutil.rmtree(P['cap'], ignore_errors=True)
    L = cs.make_listings(R)
    L.to_pickle(os.path.join(P['work'], 'listings.pkl'))
    log(f'4/4 검색 목록: 판매중 상품 {len(L):,}개, 검색어 {L["검색어"].nunique():,}개')
    cs.write_report(cs.judge(L, {}), P['report'], cfg.BRAND_KR, src)
    log(f'옵션 재고 점검 완료 → {os.path.basename(P["option_report"])}')
    return P


def auto_search_29cm(out_dir, log=_console, stop=None, progress=None, chrome_path=None, headless=False, base=None):
    """크롬을 띄워 29CM 검색창에 직접 검색 → 29CM_검색노출점검 엑셀 + 옵션 점검 엑셀에 검색 결과 반영"""
    import pickle
    import cm29_match as cm, cm29_search as cs
    P = _paths_29cm(out_dir)
    lp = os.path.join(P['work'], 'listings.pkl')
    if not os.path.exists(lp):
        raise FileNotFoundError('먼저 29CM 옵션 재고 점검(▶ 점검 시작)을 실행해 주세요.')
    L = pandas.read_pickle(lp)
    with open(P['meta'], encoding='utf-8') as f:
        meta = json.load(f)
    log('29CM 자동 검색: 크롬을 띄웁니다. 검색이 끝날 때까지 그 크롬 창은 건드리지 말아 주세요 (최소화는 괜찮습니다).')
    kw = dict(log=log, stop=stop, progress=progress, chrome_path=chrome_path, headless=headless,
              drop_words=getattr(cfg, 'SEARCH_DROP_WORDS', ['UNISEX', '유니섹스']))
    if base:
        kw['base'] = base
    found = cs.run(L, meta['brand'], out_dir, **kw)
    S = cs.judge(L, found)
    cs.write_report(S, P['report'], meta['brand'], meta['source'])
    n_cap = int((S['캡처'] != '').sum())
    if n_cap:
        log(f'검색 안 된 상품 화면 캡처 {n_cap:,}개 → {P["cap"]}')
    with open(P['state'], 'rb') as f:
        R, A, res, s = pickle.load(f)
    cm.write_report(R, A, res, s, P['option_report'], meta['source'], search=S.set_index('29CM 상품번호'))
    n_ok = int(S['검색여부'].str.startswith('검색됨').sum()); n_miss = int((S['검색여부'] == '검색 안 됨').sum())
    n_unk = int((S['검색여부'] == '미확인').sum())
    log(f'검색됨 {n_ok:,}개 / 검색 안 됨 {n_miss:,}개' + (f' / 미확인 {n_unk:,}개' if n_unk else ''))
    log(f'결과 엑셀: {P["report"]}')
    log(f'          {P["option_report"]}')
    return P['report']


def main():
    """명령줄: python app/runner.py <재고관리_복사본.xlsx> <셀메이트_재고목록.csv> [결과폴더] [검색결과.txt | --auto]
       29CM : python app/runner.py --29cm <29CM파일1.xlsx;29CM파일2.xlsx;...> <셀메이트_재고목록.csv> [결과폴더] [--auto]"""
    a = sys.argv[1:]
    if a and a[0] == '--29cm':
        if len(a) < 3:
            print(main.__doc__); sys.exit(1)
        out = os.path.abspath(a[3]) if len(a) > 3 else os.path.join(cfg.ROOT, 'output')
        try:
            P = prepare_29cm([os.path.abspath(p) for p in a[1].split(';') if p], os.path.abspath(a[2]), out)
            if len(a) > 4 and a[4] == '--auto':
                auto_search_29cm(out)
            print(' → 옵션 재고 점검:', P['option_report'])
        except Exception:
            traceback.print_exc(); sys.exit(1)
        return
    if len(a) < 2:
        print(main.__doc__); sys.exit(1)
    out = os.path.abspath(a[2]) if len(a) > 2 else os.path.join(cfg.ROOT, 'output')
    try:
        P = prepare(os.path.abspath(a[0]), os.path.abspath(a[1]), out)
        print(' → 옵션 재고 점검:', P['option_report'])
        if len(a) > 3 and a[3] == '--auto':
            auto_search(out)
        elif len(a) > 3:
            with open(a[3], encoding='utf-8') as f:
                apply_search(f.read(), out)
        else:
            print(' → 검색 스크립트:', P['script'])
        print(' → 결과 엑셀:', P['report'])
    except Exception:
        traceback.print_exc(); sys.exit(1)


if __name__ == '__main__':
    main()
