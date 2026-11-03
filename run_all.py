"""
전체 점검 실행
    python run_all.py            # 1~6단계 전체
    python run_all.py --listing  # 1단계(등록 현황)만
결과: output/ 폴더
"""
import os, sys, runpy, time

ROOT = os.path.dirname(os.path.abspath(__file__))
os.chdir(ROOT)
sys.path.insert(0, ROOT)
import config as cfg

STEPS = [
    ('step1_listing_match.py',  '1. 셀메이트 ↔ 무신사 전체 상품 매칭 (스타일넘버·상품명)'),
    ('step1_listing_report.py', '   → 등록 현황 엑셀'),
    ('step2_name_check.py',     '2. 상품명 대조 (컬러·COL 표기 제외)'),
    ('step3_option_match.py',   '3. 무신사 옵션 ↔ 셀메이트 옵션 매칭'),
    ('step4_option_judge.py',   '4. 옵션 판정 (품절·판매 제한·MFS 등)'),
    ('step6_option_report.py',  '5. 옵션 재고 점검 엑셀'),
    ('step5_search_keywords.py','6. 무신사 검색어 + 브라우저 검색 스크립트 생성'),
    ('step6_option_report.py',  '   → 엑셀 다시 작성 (검색 결과 반영)'),
]

def need(path, label):
    if not os.path.exists(path):
        sys.exit(f'[입력 파일 없음] {label}: {path}\n  README의 "준비할 파일"을 확인하세요.')

need(cfg.SELLMATE_STOCK, '셀메이트 재고목록')
if '--listing' in sys.argv:
    need(cfg.MUSINSA_ALL, '무신사 상품정보조회(전체)')
    steps = STEPS[:2]
elif os.path.exists(cfg.MUSINSA_ALL):
    steps = STEPS
else:
    print('상품정보조회(전체) 파일이 없어 등록 현황 엑셀은 건너뜁니다.')
    steps = STEPS[2:]
if '--listing' not in sys.argv:
    need(cfg.MUSINSA_SELLING, '무신사 상품정보조회(판매중)')
    need(cfg.MUSINSA_OPTION, '무신사 재고관리 옵션 목록')

for f, label in steps:
    t = time.time()
    print(f'{label} ...', end=' ', flush=True)
    runpy.run_path(os.path.join('src', f), run_name='__main__')
    print(f'{time.time()-t:.0f}s')

print('\n완료 →', cfg.OUT)
if not os.path.exists(cfg.SEARCH_RESULT) and '--listing' not in sys.argv:
    print('※ 무신사 검색 노출 확인: output/musinsa_brand_search_ready.js 를 무신사 페이지 콘솔에서 실행 → '
          '결과를 data/search_brand.txt 로 저장 → run_all.py 다시 실행')
