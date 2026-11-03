"""
시험용 가짜 입력 파일 생성 (실제 회사 데이터 없이 exe가 끝까지 도는지 확인용)
    python tests/make_fixture.py <폴더>
→ 재고 관리(v2) 복사본 3가지 형태 + 브라우저 검색 결과(search_result.txt)
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
with open(os.path.join(out, 'search_result.txt'), 'w', encoding='utf-8') as f:
    f.write('9000001|1|1|1\n9000002|1|1|1\n9000003|1||0\n9000004|3||2\n')
print('fixture →', out)
