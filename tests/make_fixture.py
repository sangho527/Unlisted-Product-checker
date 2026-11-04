"""
시험용 가짜 입력 파일 생성 (실제 회사 데이터 없이 exe가 끝까지 도는지 확인용)
    python tests/make_fixture.py <폴더>
→ 재고 관리(v2) 복사본 3가지 형태 + 셀메이트 재고목록(sellmate_stock.csv) + 브라우저 검색 결과(search_result.txt)
   stock_copy_noheader.xlsx       18개 열, 머리글 없음 (화면을 그대로 복사한 경우)
   stock_copy_header_bottom.xlsx  18개 열, 맨 아래에 머리글
   stock_copy_header_top14.xlsx   필요한 14개 열만, 맨 위에 머리글
"""
import os, sys
import pandas as pd

out = sys.argv[1] if len(sys.argv) > 1 else 'tests/fixture'
os.makedirs(out, exist_ok=True)
C18 = ['매장명', '업체', '품목', '브랜드', '스타일넘버', '상품구분', '상품번호', '상품명', '상태', '옵션', '옵션번호', '사용유무',
       '국내출고/브랜드배송', '국내출고/무신사배송', '일본출고/무신사배송', '허용판매수량', '총합 판매가능 재고수', '재입고알림요청수']
rows = []
def op(sn, no, name, opt, oid, brand, mfs):
    tot = brand + mfs
    rows.append(['온라인 스토어', 'test', '의류', '어반드레스', sn, '입점/위탁', no, name, '판매중', opt, oid, 'True',
                 brand, mfs, 0, 2000000000, tot, 0])
op('AD99TS01', '9000001', 'Test Crop Tee_2COL', 'BLACK^S', '1', 10, 0)
op('AD99TS01', '9000001', 'Test Crop Tee_2COL', 'BLACK^M', '2', 0, 0)
op('AD99TS01', '9000001', 'Test Crop Tee_2COL', 'WHITE^S', '3', 7, 0)
op('AV99PT01BL', '9000002', '테스트 와이드 팬츠 BLUE', 'S', '4', 3, 0)
op('AV99CT01BK', '9000003', '유니섹스 테스트 롱 코트 블랙', 'FREE', '5', 5, 0)
op('AV99SK01BK', '9000004', 'Test Mini Skirt BLACK', 'S', '6', 0, 0)
op('ZD99TS01BK', '9000005', 'Other Brand Tee BLACK', 'S', '7', 1, 0)   # 다른 브랜드(제외)
df = pd.DataFrame(rows, columns=C18)
df.to_excel(os.path.join(out, 'stock_copy_noheader.xlsx'), index=False, header=False)
pd.concat([df, pd.DataFrame([C18], columns=C18)]).to_excel(os.path.join(out, 'stock_copy_header_bottom.xlsx'), index=False, header=False)
keep = ['매장명', '업체', '품목', '브랜드', '스타일넘버', '상품번호', '상품명', '상태', '옵션', '옵션번호',
        '국내출고/브랜드배송', '국내출고/무신사배송', '일본출고/무신사배송', '총합 판매가능 재고수']
df[keep].rename(columns={'총합 판매가능 재고수': '총합 판매가능'}).to_excel(os.path.join(out, 'stock_copy_header_top14.xlsx'), index=False)
# ---- 셀메이트 재고목록 (cp949 CSV, 값 앞에 = 가 붙는 형식) ----
S_COLS = ['상품분류', '상품시즌', '상품연도', '상품코드', '바코드번호(서식)', '상품명', '옵션코드', '옵션명', '현재재고',
          '대표판매가', '원가', '판매여부', '입고예정일', '관리등급']
srows = []
def sm(code, name, opts, sale='판매'):
    for size, qty in opts:
        srows.append(['#상의-반팔', '1-S/S-봄여름', '26', code, '1', name, code + size, size, qty, '30000', '1', sale, '', '일반관리'])
sm('AD99TS01BK', 'Test Crop Tee BLACK', [('S', 10), ('M', 5)])     # 무신사 BLACK^M 품절인데 재고 5 → 문제
sm('AD99TS01WH', 'Test Crop Tee WHITE', [('S', 7), ('M', 0)])
sm('AV99PT01BL', '테스트 와이드 팬츠 BLUE', [('S', 3), ('M', 4)], sale='이월')
sm('AV99CT01BK', '유니섹스 테스트 롱 코트 블랙', [('FREE', 5)])
sm('AV99SK01BK', 'Test Mini Skirt BLACK', [('S', 0)])
sm('ZD99TS01BK', 'Other Brand Tee BLACK', [('S', 1)])
sm('AV99TS02BK', 'UNISEX Test Hoodie BLACK', [('M', 3)])                # 29CM: UNISEX 빼고 재검색
with open(os.path.join(out, 'sellmate_stock.csv'), 'w', encoding='cp949', newline='') as f:
    f.write(','.join(f'"{c}"' for c in S_COLS) + ',\n')
    for r in srows:
        f.write(','.join(f'="{v}"' if i in (3, 4) else f'"{v}"' for i, v in enumerate(r)) + ',\n')

with open(os.path.join(out, 'search_result.txt'), 'w', encoding='utf-8') as f:
    f.write('9000001|1|1|1\n9000002|1|1|1\n9000003|1||0\n9000004|3||2\n')
print('fixture →', out)

# ---- 29CM 옵션재고 일괄변경 파일 (페이지 2개, 1줄 겹침) — 위쪽 안내문 2줄 + 머리글 ----
C29 = ['스타일 번호', '상품 번호', '상품명', '브랜드', '옵션 번호', '옵션명', '옵션 판매 상태', '재고 버전', '재고 수량\n(변경 전)',
       '재고 수량\n(변경 후)', '옵션 추가 가격\n(변경 전)', '옵션 추가 가격\n(변경 후)', '판매가', '누적 판매량', '재입고 알림 여부\n(변경 전)',
       '재입고 알림 여부\n(변경 후)']
def r29(pno, name, oid, opt, st, qty):
    return [f'{pno}-{oid}', pno, name, 'AVANDRESS', oid, opt, st, '', qty, '', 0, '', 39000, 0, 'N', '']
p1 = [r29('3000001', 'Test Crop Tee_2Colors', '1', '[색상]BLACK [사이즈]S', '판매중', 10),
      r29('3000001', 'Test Crop Tee_2Colors', '2', '[색상]BLACK [사이즈]M', '일시품절', 0),     # 셀메이트 5개 → ①
      r29('3000001', 'Test Crop Tee_2Colors', '3', '[색상]WHITE [사이즈]S', '판매중', 7),
      r29('3000002', '테스트 와이드 팬츠 BLUE', '4', '[사이즈]S', '판매중지', 0)]                 # 셀메이트 3개 → ②
p2 = [r29('3000002', '테스트 와이드 팬츠 BLUE', '4', '[사이즈]S', '판매중지', 0),                 # 1페이지와 겹침
      r29('3000002', '테스트 와이드 팬츠 BLUE', '5', '[사이즈]M', '판매중', 4),
      r29('3000003', 'Test Mini Skirt BLACK', '6', '[사이즈]S', '판매중', 2),                       # 셀메이트 0 → ④
      r29('3000004', 'UNISEX Test Hoodie - 2COL', '7', '[색상]BLACK [사이즈]M', '판매중', 3)]       # UNISEX 빼고 재검색하면 검색됨
for fn, rows_ in [('cm29_page1.xlsx', p1), ('cm29_page2.xlsx', p2)]:
    top = [['[자동 입력 / 수정 불가]'] * len(C29), ['기존 정보 확인'] + [None] * (len(C29) - 1), C29]
    pd.DataFrame(top + rows_).to_excel(os.path.join(out, fn), index=False, header=False, sheet_name='옵션')
print('29CM fixture →', out)
