"""
점검 실행 모듈 (GUI·명령줄 공용)

run_check(files, out_dir, listing_only=False, log=print)
    files: {'sellmate': ..., 'musinsa_all': ..., 'musinsa_selling': ..., 'musinsa_option': ..., 'search': ...(선택)}
rebuild_with_search(search_text, out_dir, log=print)
    브라우저 검색 결과(텍스트)를 저장하고 옵션 재고 점검 엑셀만 다시 작성
"""
import os, sys, io, runpy, time, contextlib, traceback

# 단계별 스크립트(src/)가 쓰는 라이브러리 — PyInstaller가 exe에 포함하도록 여기서 미리 import
import pandas, numpy, openpyxl, difflib, json, re, math, collections, datetime  # noqa: F401
import urllib.parse  # noqa: F401
from openpyxl.workbook.properties import CalcProperties  # noqa: F401

# exe면 압축 해제 폴더, 아니면 레포 폴더 (src/, search/ 가 있는 곳)
BASE = getattr(sys, '_MEIPASS', os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if BASE not in sys.path:
    sys.path.insert(0, BASE)
import config as cfg  # noqa: E402

LISTING_STEPS = [
    ('step1_listing_match.py',  '1/6 셀메이트 ↔ 무신사 전체 상품 매칭'),
    ('step1_listing_report.py', '    등록 현황 엑셀 작성'),
]
OPTION_STEPS = [
    ('step2_name_check.py',     '2/6 상품명 대조 (컬러·COL 표기 제외)'),
    ('step3_option_match.py',   '3/6 무신사 옵션 ↔ 셀메이트 옵션 매칭'),
    ('step4_option_judge.py',   '4/6 옵션 판정'),
    ('step6_option_report.py',  '5/6 옵션 재고 점검 엑셀 작성'),
    ('step5_search_keywords.py','6/6 무신사 검색 스크립트 생성'),
]


def _console(msg):
    # print 대신 원래 콘솔에 직접 씀 (단계 실행 중에는 stdout을 로그로 돌려받기 때문)
    out = sys.__stdout__
    if out:
        out.write(msg + '\n'); out.flush()


class _LogWriter(io.TextIOBase):
    def __init__(self, log):
        self.log, self.buf = log, ''
    def write(self, s):
        self.buf += s
        while '\n' in self.buf:
            line, self.buf = self.buf.split('\n', 1)
            if line.strip():
                self.log('      ' + line)
        return len(s)


def _run_steps(steps, log):
    cwd = os.getcwd()
    os.chdir(BASE)
    try:
        for f, label in steps:
            t = time.time()
            log(f'{label} ...')
            # 단계 스크립트의 디버그 출력은 진행 로그에 섞지 않음
            with contextlib.redirect_stdout(io.StringIO()):
                runpy.run_path(os.path.join('src', f), run_name='__main__')
            log(f'      완료 ({time.time() - t:.0f}초)')
    finally:
        os.chdir(cwd)


def _check(path, label):
    if not path or not os.path.exists(path):
        raise FileNotFoundError(f'{label} 파일을 선택해 주세요.')


def run_check(files, out_dir, listing_only=False, log=_console):
    _check(files.get('sellmate'), '셀메이트 재고목록')
    has_all = bool(files.get('musinsa_all')) and os.path.exists(files['musinsa_all'])
    if listing_only:
        _check(files.get('musinsa_all'), '무신사 상품정보조회(전체)')
    else:
        _check(files.get('musinsa_selling'), '무신사 상품정보조회(판매중)')
        _check(files.get('musinsa_option'), '무신사 재고관리 옵션 목록')
    cfg.load_settings()
    cfg.set_output(out_dir)
    cfg.SELLMATE_STOCK = files['sellmate']
    cfg.MUSINSA_ALL = files.get('musinsa_all') or ''
    cfg.MUSINSA_SELLING = files.get('musinsa_selling') or ''
    cfg.MUSINSA_OPTION = files.get('musinsa_option') or ''
    search = files.get('search')
    cfg.SEARCH_RESULT = search if search and os.path.exists(search) else os.path.join(cfg.WORK, '__no_search__.txt')
    log(f'브랜드: {cfg.BRAND_KR} / 결과 폴더: {out_dir}')
    if listing_only:
        steps = LISTING_STEPS
    elif has_all:
        steps = LISTING_STEPS + OPTION_STEPS
    else:
        log('상품정보조회(전체) 파일이 없어 등록 현황 엑셀(미등록·판매중지)은 건너뜁니다.')
        steps = OPTION_STEPS
    _run_steps(steps, log)
    if not listing_only and search and os.path.exists(search):
        log('검색 결과 반영해서 엑셀 다시 작성 ...')
        _run_steps([('step6_option_report.py', '    옵션 재고 점검 엑셀')], log)
    log('끝났습니다.')
    return outputs(out_dir)


def rebuild_with_search(search_text, out_dir, log=_console):
    lines = [l.strip() for l in search_text.replace(';', '\n').splitlines() if l.strip()]
    bad = [l for l in lines if l.count('|') < 3]
    if not lines or bad:
        raise ValueError('검색 결과 형식이 아닙니다. 무신사 콘솔에서 복사한 결과를 그대로 붙여넣어 주세요.')
    cfg.set_output(out_dir)
    path = os.path.join(out_dir, 'search_brand.txt')
    with open(path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    cfg.SEARCH_RESULT = path
    miss = sum(1 for l in lines if l.split('|')[2] == '')
    log(f'검색 결과 {len(lines)}개 저장 (검색 안 됨 {miss}개) → {path}')
    _run_steps([('step6_option_report.py', '옵션 재고 점검 엑셀 다시 작성')], log)
    log('끝났습니다.')
    return path


def outputs(out_dir):
    return sorted(os.path.join(out_dir, f) for f in os.listdir(out_dir)
                  if f.endswith('.xlsx') or f.endswith('.js'))


def main():
    """명령줄: python app/runner.py 셀메이트.csv 무신사전체.xlsx(없으면 -) 무신사판매중.xlsx 옵션목록.xlsx [결과폴더] [검색결과.txt]"""
    a = sys.argv[1:]
    if len(a) < 4:
        print(main.__doc__); sys.exit(1)
    out = os.path.abspath(a[4]) if len(a) > 4 else os.path.join(cfg.ROOT, 'output')
    files = dict(zip(['sellmate', 'musinsa_all', 'musinsa_selling', 'musinsa_option'], [os.path.abspath(p) if p != '-' else '' for p in a[:4]]))
    if len(a) > 5:
        files['search'] = os.path.abspath(a[5])
    try:
        for f in run_check(files, out):
            print(' →', f)
    except Exception:
        traceback.print_exc(); sys.exit(1)


if __name__ == '__main__':
    main()
