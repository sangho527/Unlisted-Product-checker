"""
시험용 가짜 입력 파일 생성 (실제 회사 데이터 없이 exe/파이프라인이 끝까지 도는지 확인용)
    python tests/make_fixture.py <폴더>
→ sellmate_stock.csv, musinsa_products_all.xlsx, musinsa_products_selling.xlsx, musinsa_option_stock.xlsx, search_brand.txt
"""
import os, sys
import pandas as pd
from openpyxl import Workbook

out = sys.argv[1] if len(sys.argv) > 1 else 'tests/fixture'
os.makedirs(out, exist_ok=True)

# ---- 셀메이트 재고목록 (cp949 CSV) ----
S_COLS = ['상품분류', '상품시즌', '상품연도', '상품코드', '바코드번호(서식)', '상품명', '옵션코드', '옵션명', '현재재고',
          '대표판매가', '원가', '판매여부', '입고예정일', '관리등급']
rows = []
def sm(code, name, opts, sale='판매'):
    for size, qty in opts:
        rows.append(['#상의-반팔', '1-S/S-봄여름', '26', code, '1', name, code + size, size, qty, '30000', '1', sale, '', '일반관리'])
sm('AD99TS01BK', 'Test Crop Tee BLACK', [('S', 10), ('M', 5)])
sm('AD99TS01WH', 'Test Crop Tee WHITE', [('S', 7), ('M', 0)])
sm('AV99PT01BL', '테스트 와이드 팬츠 BLUE', [('S', 3), ('M', 4)], sale='이월')
sm('AD99OP01BK', 'Test Dress BLACK', [('FREE', 9)])          # 무신사 미등록
sm('ZD99TS01BK', 'Other Brand Tee BLACK', [('S', 1)])         # 다른 브랜드(제외)
with open(os.path.join(out, 'sellmate_stock.csv'), 'w', encoding='cp949', newline='') as f:
    f.write(','.join(f'"{c}"' for c in S_COLS) + ',\n')
    for r in rows:
        f.write(','.join(f'="{v}"' if i in (3, 4) else f'"{v}"' for i, v in enumerate(r)) + ',\n')

# ---- 무신사 상품정보조회 (2행 헤더, 시트 '상품정보') ----
P_COLS = ['무신사', '상품 번호', '상품명', '대표이미지', '플랫폼명', '조조타운', '티몰글로벌', '브랜드', '판매 지역', '품목 카테고리',
          '기준 카테고리', '상품 구분', '판매분 매입', '원산지', '상품 유형', '스타일넘버', '정상가(원)', '판매가(원)', '세일', '할인율(%)',
          '수수료(%)', '쿠폰 할인', '회원 할인', '제휴 할인', '적립금 선할인', '보유 적립금 할인', '적립금 금액(원)', '적립금 지급(%)',
          '판매 가능 재고', '배송 방법', '출고 예정 구분', '과세 구분', '옵션 사용', '성별', '아울렛 전환', '시즌 연도/타입', '단독 상품 구분',
          '판매 기간', '스타일', '수입확인증필', '시험성적서 등록', '시험성적서 필수 브랜드', '전문관', '상품정보제공고시 타입', '맞춤 제작',
          'K-커넥트 상품', '사이즈/백필 가능', '통관 서류', '속성', '색상', '출고지/반품지', '배송비 정책', '배송 형태', 'API 등록', '품목',
          '상단 홍보글 기간 설정', '상단 홍보글', '글로벌 전용', '한국 HS CODE', '히어로 상품', '오프라인 판매', '세트 상품', '세트 상품 개수', '등록 일시']
LISTINGS = [  # 상태, 상품번호, 상품명, 스타일넘버, 재고
    ('판매중', '9000001', 'Test Crop Tee - 2COL', 'AD99TS01', '20'),
    ('품절',   '9000002', '테스트 와이드 팬츠 BLUE', 'AV99PT01BL', '0'),
    ('판매중', '9000003', 'Other Brand Tee BLACK', 'ZD99TS01BK', '1'),
]
def po(path, statuses):
    wb = Workbook(); ws = wb.active; ws.title = '상품정보'
    ws.append(['판매/검수 상태'] + [''] * (len(P_COLS) - 1))
    ws.append(P_COLS)
    for st, no, name, sn, q in LISTINGS:
        if st not in statuses: continue
        r = {c: '' for c in P_COLS}
        r.update({'무신사': st, '상품 번호': no, '상품명': name, '스타일넘버': sn, '판매 가능 재고': q, '판매가(원)': '30,000',
                  '브랜드': 'test', '옵션 사용': '예', '등록 일시': '2026-01-01 10:00'})
        ws.append([r[c] for c in P_COLS])
    wb.create_sheet('__verify')
    wb.save(path)
po(os.path.join(out, 'musinsa_products_all.xlsx'), {'판매중', '품절'})
po(os.path.join(out, 'musinsa_products_selling.xlsx'), {'판매중'})

# ---- 무신사 재고관리(v2) 옵션 목록 ----
O_COLS = ['매장명', '업체', '품목', '브랜드', '스타일넘버', '상품번호', '상품명', '상태', '옵션', '옵션번호',
          '국내출고/브랜드배송', '국내출고/무신사배송', '일본출고/무신사배송', '총합 판매가능']
orows = []
def op(sn, no, name, opt, oid, brand, mfs, total):
    orows.append(['온라인 스토어', 'test', '의류', '테스트', sn, no, name, '판매중', opt, oid, brand, mfs, 0, total])
op('AD99TS01', '9000001', 'Test Crop Tee - 2COL', 'BLACK^S', '1', 10, 0, 10)
op('AD99TS01', '9000001', 'Test Crop Tee - 2COL', 'BLACK^M', '2', 0, 0, 0)     # 셀메이트 재고 5인데 품절 → 문제
op('AD99TS01', '9000001', 'Test Crop Tee - 2COL', 'WHITE^S', '3', 7, 0, 7)
op('AD99TS01', '9000001', 'Test Crop Tee - 2COL', 'WHITE^M', '4', 0, 3, 3)     # MFS 출고
op('ZD99TS01BK', '9000003', 'Other Brand Tee BLACK', 'S', '5', 1, 0, 1)
pd.DataFrame(orows, columns=O_COLS).to_excel(os.path.join(out, 'musinsa_option_stock.xlsx'), index=False)

with open(os.path.join(out, 'search_brand.txt'), 'w', encoding='utf-8') as f:
    f.write('9000001|1|1|1\n')
print('fixture →', out)
