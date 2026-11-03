"""
검색 노출 점검 실행 모듈 (화면·명령줄 공용)

prepare(stock_copy_path, out_dir)      재고 관리(v2) 복사본 → 상품 목록 + 브라우저 검색 스크립트 + 결과 엑셀(검색 전)
apply_search(search_text, out_dir)     브라우저 검색 결과 → 결과 엑셀 완성
"""
import os, sys, json, datetime, traceback

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
    }


def prepare(stock_copy_path, out_dir, log=_console):
    if not stock_copy_path or not os.path.exists(stock_copy_path):
        raise FileNotFoundError('재고 관리(v2) 복사본 엑셀 파일을 선택해 주세요.')
    cfg.load_settings()
    P = _paths(out_dir)
    log(f'브랜드: {cfg.BRAND_KR}')
    log('1/3 재고 관리(v2) 복사본 읽는 중 ...')
    df = sc.read_stock_table(stock_copy_path)
    log(f'      옵션 {len(df):,}줄 (판매중), 필요한 열 {len(df.columns)}개만 사용')
    log('2/3 상품별로 묶고 검색어 만드는 중 ...')
    L, excluded = sc.make_listings(df, cfg.EXCLUDE_CODE_PATTERN)
    log(f'      판매중 상품 {len(L):,}개' + (f' (다른 브랜드 {excluded}개 제외)' if excluded else ''))
    L.to_pickle(P['listings'])
    meta = {'source': os.path.basename(stock_copy_path), 'excluded': int(excluded), 'brand': cfg.BRAND_KR}
    with open(P['meta'], 'w', encoding='utf-8') as f:
        json.dump(meta, f, ensure_ascii=False)
    log('3/3 브라우저 검색 스크립트와 결과 엑셀(검색 전) 만드는 중 ...')
    sc.write_search_script(L, cfg.BRAND_KR, SCRIPT_TEMPLATE, P['script'])
    sc.write_report(L, None, cfg.BRAND_KR, P['report'], meta['source'], excluded)
    mins = max(1, round(len(L) * 0.55 / 60))
    log(f'준비 완료. 이제 무신사 검색을 실행하세요 (상품 {len(L):,}개, 약 {mins}분).')
    return P


def apply_search(search_text, out_dir, log=_console):
    P = _paths(out_dir)
    if not os.path.exists(P['listings']):
        raise FileNotFoundError('먼저 [상품 목록 만들기]를 실행해 주세요.')
    import pandas as pd
    L = pd.read_pickle(P['listings'])
    with open(P['meta'], encoding='utf-8') as f:
        meta = json.load(f)
    sr = sc.parse_search_result(search_text)
    with open(P['search'], 'w', encoding='utf-8') as f:
        f.write(search_text.strip())
    R = sc.write_report(L, sr, meta['brand'], P['report'], meta['source'], meta['excluded'])
    n_ok = int(R['검색여부'].str.startswith('검색됨').sum())
    n_miss = int((R['검색여부'] == '검색 안 됨').sum())
    n_unk = int((R['검색여부'] == '미확인').sum())
    log(f'검색됨 {n_ok:,}개 / 검색 안 됨 {n_miss:,}개' + (f' / 미확인 {n_unk:,}개' if n_unk else ''))
    log(f'결과 엑셀: {P["report"]}')
    return P['report']


def main():
    """명령줄: python app/runner.py <재고관리_복사본.xlsx> [결과폴더] [검색결과.txt]"""
    a = sys.argv[1:]
    if not a:
        print(main.__doc__); sys.exit(1)
    out = os.path.abspath(a[1]) if len(a) > 1 else os.path.join(cfg.ROOT, 'output')
    try:
        P = prepare(os.path.abspath(a[0]), out)
        if len(a) > 2:
            with open(a[2], encoding='utf-8') as f:
                apply_search(f.read(), out)
        else:
            print(' → 검색 스크립트:', P['script'])
        print(' → 결과 엑셀:', P['report'])
    except Exception:
        traceback.print_exc(); sys.exit(1)


if __name__ == '__main__':
    main()
