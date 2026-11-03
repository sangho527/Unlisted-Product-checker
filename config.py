"""
경로·브랜드·수동 연결 설정
- 입력 파일은 data/ 폴더에 아래 이름으로 넣어주세요 (이름이 다르면 여기서 바꾸면 됩니다).
- data/, output/ 폴더는 .gitignore 처리되어 있어 레포에 올라가지 않습니다 (재고·원가 등 내부 데이터 보호).
"""
import os, datetime

ROOT = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(ROOT, 'data') + os.sep
OUT  = os.path.join(ROOT, 'output') + os.sep
WORK = os.path.join(OUT, 'work') + os.sep
os.makedirs(WORK, exist_ok=True)
TODAY = datetime.date.today().strftime('%Y%m%d')

# ---- 입력 파일 ----
SELLMATE_STOCK  = DATA + 'sellmate_stock.csv'          # 셀메이트 재고목록 (stk_stockList_*.csv, cp949)
MUSINSA_ALL     = DATA + 'musinsa_products_all.xlsx'   # 무신사 상품정보조회 - 상품상태 '전체' (PO_상품정보조회_*.xlsx)
MUSINSA_SELLING = DATA + 'musinsa_products_selling.xlsx'  # 무신사 상품정보조회 - 상품상태 '판매중'
MUSINSA_OPTION  = DATA + 'musinsa_option_stock.xlsx'   # 무신사 재고관리(v2) 옵션 목록 (화면 복사 → 엑셀 저장)
SEARCH_RESULT   = DATA + 'search_brand.txt'            # search/musinsa_brand_search.js 결과 (없으면 검색 시트 생략)

# ---- 브랜드 ----
BRAND_KR = '어반드레스'
# 점검 대상 상품코드: AD·AV로 시작 + 구코드(yySS/FW + AD)
TARGET_CODE_PATTERN  = r'^(AD|AV|\d{2}(SS|FW)AD)'
# 같은 무신사 스토어에 섞여 있지만 이번 점검에서 제외할 다른 브랜드 (젠더스트 ZD, 비바셔스 VV)
EXCLUDE_CODE_PATTERN = r'^(ZD|VV)'

# ---- 사람이 확인한 수동 연결 ----
# 무신사 스타일넘버가 잘못 입력돼 상품명으로도 자동 연결이 안 되는 상품 (한글/영문 상품명이 서로 다른 경우)
# {무신사 상품번호: [셀메이트 상품코드, ...]}
MANUAL_LINKS = {
    '6618666': ['AD62TS06BK', 'AD62TS06BR', 'AD62TS06CH', 'AD62TS06GR'],   # Sally ... Short Sleeve → 샐리 반팔 티
    '6586867': ['AV61PT06BK', 'AV61PT06WH'],                                # 테사 세미 와이드 팬츠 → Tessa Semi Wide Pants
}
# 긴팔 상품 안에 '(short)' 옵션으로 섞여 있는 반팔 옵션
MANUAL_SHORT_LINKS = {
    '5442146': ['AD62TS06BK', 'AD62TS06BR', 'AD62TS06CH', 'AD62TS06GR'],
}
# 결과 시트 메모
LISTING_NOTES = {
    '6618666': '반팔 상품 → 셀메이트 AD62TS06(샐리 브이컷 스트라이프 레이어드 반팔 티)로 추정. 스타일넘버 수정 필요',
    '5442146': '긴팔 상품. 옵션 중 "(short)"는 반팔(AD62TS06) 옵션이 섞인 것으로 추정',
    '6586867': '테사 세미 와이드 팬츠 → 셀메이트 AV61PT06(Tessa Semi Wide Pants)로 추정. AD61PT06은 사이드 라인 플레어 팬츠',
}
SELLMATE_NOTES = {   # 셀메이트 상품코드 앞 8자리 기준
    'AD62TS06': '무신사 6618666(Sally ... Short Sleeve)으로 등록된 것으로 추정 - 스타일넘버가 AD53TS18로 잘못 입력',
    'AV61PT06': '무신사 6586867(테사 세미 와이드 팬츠)로 등록된 것으로 추정 - 스타일넘버가 AD61PT06으로 잘못 입력',
}

def fname(p):
    return os.path.basename(p)
