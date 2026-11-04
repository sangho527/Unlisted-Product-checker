"""
29CM 검색 화면을 흉내 낸 시험용 로컬 서버 (크롬 자동 검색 시험용, 실제 29CM 접속 없음)
    python tests/mock_29cm.py <port>
- 메인: 오른쪽 위 돋보기 버튼(글자 없는 svg 버튼) → 검색 레이어의 input[name=keyword] → Enter → /store/search?keyword=...
- 검색 결과: 위쪽 검색창(input[name=keyword]), "총 N개 상품", li 카드(상품 링크 catalog/번호 + 브랜드 링크 /store/brand/번호)
"""
import sys, json, urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

BRAND = '어반드레스'
DB = {
    '어반드레스 Test Crop Tee': [('3000001', BRAND, 'Test Crop Tee_2Colors', False), ('8000001', '다른브랜드', 'Crop Tee', False)],
    '어반드레스 테스트 와이드 팬츠': [('8000002', '다른브랜드', '와이드 팬츠', False)],          # 우리 상품 안 나옴
    '어반드레스 UNISEX Test Hoodie': [],                                                          # 결과 없음
    '어반드레스 Test Hoodie': [('3000004', BRAND, 'UNISEX Test Hoodie - 2COL', False)],           # UNISEX 빼고 재검색하면 나옴
    '어반드레스 Test Mini Skirt': [(str(7000000 + i), BRAND, f'Other Skirt {i}', False) for i in range(40)]
                                + [('3000003', BRAND, 'Test Mini Skirt BLACK', True)],          # 41번째, 품절 표시
}
HEAD = """<!doctype html><meta charset=utf-8><title>29CM</title>
<style>body{margin:0;font:14px sans-serif;width:1300px}header{height:120px;position:relative}
#sbtn{position:absolute;right:30px;top:60px;width:44px;height:44px;border:0;background:none}
#layer{display:none;position:fixed;top:0;right:0;width:700px;background:#fff;padding:30px}
#layer input{width:600px;font-size:30px}
.top input{width:700px;font-size:24px;margin:20px 300px}
ul{display:grid;grid-template-columns:repeat(4,220px);gap:10px;list-style:none}li{height:330px;border:1px solid #ddd}</style>
<header>29CM <button id=sbtn><svg viewBox="0 0 24 24" width=24 height=24><circle cx=10 cy=10 r=7 /></svg></button></header>
<div id=layer><input name=keyword placeholder=Search></div>
<script>
const go=v=>{location.href='/store/search?keyword='+encodeURIComponent(v)+'&sort=RECOMMENDED&page=1';};
document.getElementById('sbtn').onclick=()=>{const l=document.getElementById('layer');l.style.display='block';l.querySelector('input').focus();};
document.addEventListener('keydown',e=>{if(e.key==='Enter'&&e.target.name==='keyword'){go(e.target.value);}});
</script>"""


class H(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def do_GET(self):
        u = urllib.parse.urlparse(self.path)
        body = HEAD
        if u.path.startswith('/store/search'):
            kw = urllib.parse.parse_qs(u.query).get('keyword', [''])[0]
            items = DB.get(kw, [])
            body += f'<div class=top><input name=keyword placeholder=Search value="{kw}"></div><p>총 {len(items)}개 상품</p><ul id=g></ul>'
            body += '<script>const D=%s;setTimeout(()=>{const g=document.getElementById("g");' % json.dumps(items, ensure_ascii=False)
            body += """D.forEach(d=>{const li=document.createElement('li');
 li.innerHTML=`<a href="https://product.29cm.co.kr/catalog/${d[0]}?search_keyword=x">img</a><div><a href="/store/brand/${d[1]==='어반드레스'?7564:1}">${d[1]}</a>
 <a href="https://product.29cm.co.kr/catalog/${d[0]}">${d[2]}</a></div><div>30,000</div>${d[3]?'<div>품절</div>':''}`;g.appendChild(li);});},300);</script>"""
            if not items:
                body += '<p>선택한 필터의 검색 결과가 없습니다.</p>'
        b = body.encode('utf-8')
        self.send_response(200); self.send_header('Content-Type', 'text/html; charset=utf-8'); self.send_header('Content-Length', str(len(b)))
        self.end_headers(); self.wfile.write(b)


if __name__ == '__main__':
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8766
    ThreadingHTTPServer(('127.0.0.1', port), H).serve_forever()
