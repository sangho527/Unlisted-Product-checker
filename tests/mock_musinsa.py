"""
무신사 검색 화면을 흉내 낸 시험용 로컬 서버 (크롬 자동 검색 시험용, 실제 무신사 접속 없음)
    python tests/mock_musinsa.py <port>
- 메인: 상단 [검색] 버튼(aria-label="검색 레이어 열기") → 검색창 → Enter → /search/goods?keyword=...
- 검색 결과: 상품 카드(a[data-item-id][data-section-name="goods_list"]), 화면에 보이는 카드만 그리는 목록,
  추천 영역(goods_29cm), 결과 없음 화면
"""
import sys, json, urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

BRAND = '어반드레스'
# 검색어 → [(상품번호, 브랜드, 상품명)]
DB = {
    '어반드레스 Test Crop Tee': [('9000001', BRAND, 'Test Crop Tee_2COL'), ('8000001', '다른브랜드', 'Crop Tee')],
    '어반드레스 테스트 와이드 팬츠': [('8000002', '다른브랜드', '와이드 팬츠')],                 # 우리 상품 안 나옴
    '어반드레스 유니섹스 테스트 롱 코트': [],                                                # 결과 없음
    '어반드레스 테스트 롱 코트': [('9000003', BRAND, '유니섹스 테스트 롱 코트 블랙')],          # 유니섹스 빼고 재검색하면 나옴
    '어반드레스 Test Mini Skirt': [(str(7000000 + i), BRAND, f'Other Skirt {i}') for i in range(40)]
                                 + [('9000004', BRAND, 'Test Mini Skirt BLACK')],          # 41번째(스크롤해야 보임)
}

HEAD = """<!doctype html><meta charset=utf-8><title>%s</title>
<style>body{margin:0;font:14px sans-serif}header{height:56px;background:#000;color:#fff;display:flex;align-items:center;padding:0 20px}
#layer{display:none;position:fixed;top:0;left:0;right:0;background:#fff;padding:10px;border-bottom:1px solid #ccc}
.grid{display:grid;grid-template-columns:repeat(4,200px);gap:10px;padding:20px}.card{height:300px;border:1px solid #ddd}</style>
<header>MUSINSA <button aria-label="검색 레이어 열기" style="margin-left:auto">검색</button></header>
<div id=layer><input class="SearchBar__KeywordInput" id=kw style="width:600px"></div>
<script>
document.querySelector('[aria-label="검색 레이어 열기"]').onclick=()=>{const l=document.getElementById('layer');l.style.display='block';document.getElementById('kw').focus();};
document.getElementById('kw').addEventListener('keydown',e=>{if(e.key==='Enter'){location.href='/search/goods?keyword='+encodeURIComponent(e.target.value).replace(/%%20/g,'+')+'&keywordType=keyword&gf=A';}});
</script>"""


class H(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def do_GET(self):
        u = urllib.parse.urlparse(self.path)
        if u.path.startswith('/search/goods'):
            kw = urllib.parse.parse_qs(u.query).get('keyword', [''])[0]
            items = DB.get(kw, [])
            body = HEAD % f"'{kw}' 관련 상품 찾기 | 총 {len(items) + 7}건 | 무신사"
            if not items:
                body += '<p>검색 결과가 없습니다.<br>다른 검색어를 입력해 보세요.</p>'
            else:
                body += f'<p>{len(items)}개</p><div class=grid id=g></div>'
                body += '<script>const D=%s;' % json.dumps(items, ensure_ascii=False)
                body += """const g=document.getElementById('g');
function draw(){const top=Math.floor(window.scrollY/310)*4-8;g.innerHTML='';g.style.paddingTop=Math.max(0,top/4)*310+'px';
 D.forEach((d,i)=>{if(i<top||i>top+24)return;const div=document.createElement('div');div.className='GoodsItem-x card';
 div.innerHTML=`<a href="/products/${d[0]}" data-item-id="${d[0]}" data-item-list-index="${i+1}" data-section-name="goods_list" data-item-brand="${d[1]==='어반드레스'?'avan':'x'}">img</a><div>${d[1]}</div><div>${d[2]}</div><div>30,000원</div>`;g.appendChild(div);});
 g.style.height=(Math.ceil(D.length/4)*310)+'px';}
draw();window.addEventListener('scroll',draw);</script>
<div><a href="/products/9000002" data-item-id="9000002" data-section-name="goods_29cm">추천</a></div>"""
        else:
            body = HEAD % '무신사'
        b = body.encode('utf-8')
        self.send_response(200); self.send_header('Content-Type', 'text/html; charset=utf-8'); self.send_header('Content-Length', str(len(b)))
        self.end_headers(); self.wfile.write(b)


if __name__ == '__main__':
    ThreadingHTTPServer(('127.0.0.1', int(sys.argv[1]) if len(sys.argv) > 1 else 8765), H).serve_forever()
