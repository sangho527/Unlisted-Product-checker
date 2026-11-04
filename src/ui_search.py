"""
크롬을 띄워 무신사 검색창에 사용자처럼 직접 검색하고, 검색 결과 화면에서 상품이 보이는지 확인

흐름 (검색어 1개마다)
  1. 상단 메뉴의 [검색] 버튼(검색 레이어 열기)을 마우스로 클릭
  2. 검색창에 "브랜드명 + 상품명"을 입력하고 Enter
  3. 검색 결과 화면이 뜨면 상품 카드를 읽음 (화면을 내리며 최대 60개)
     - 카드의 상품번호와 브랜드명(카드에 적힌 글자)을 확인
     - '함께 본 29CM 상품' 같은 추천 영역은 제외
  4. 찾는 상품이 우리 브랜드로 안 보이면 화면을 캡처해 '미노출_캡처' 폴더에 저장

크롬은 이 프로그램 전용 프로필(결과폴더/work/chrome_profile)로 따로 띄우므로 평소 쓰는 크롬과 섞이지 않고,
로그인하지 않은 일반 고객 화면으로 확인합니다. 크롬 원격 제어(DevTools Protocol)만 쓰고 별도 드라이버는 필요 없습니다.
"""
import os, re, sys, json, time, base64, socket, shutil, subprocess, urllib.request, urllib.parse
import websocket  # websocket-client

BASE_URL = 'https://www.musinsa.com'
HOME_PATH = '/main/musinsa/recommend'

JS_TRIGGER = r"""(() => {
  const vis = e => { const r = e.getBoundingClientRect(); return r.width > 0 && r.height > 0 && r.y >= 0 && r.y < 300; };
  const c = [...document.querySelectorAll('button[aria-label="검색 레이어 열기"], button[aria-label*="검색"], [class*="earchbar"] button, [class*="earchBar"] button, [class*="HeaderSearch"]')];
  const el = c.find(vis);
  if (!el) return null;
  const r = el.getBoundingClientRect();
  return { x: r.x + Math.min(r.width / 2, 200), y: r.y + r.height / 2 };
})()"""

JS_FOCUSED = r"""(() => {
  const a = document.activeElement;
  if (a && a.tagName === 'INPUT' && !a.readOnly) { a.select(); return true; }
  return false;
})()"""

JS_STATE = r"""(() => {
  const kw = new URLSearchParams(location.search).get('keyword');
  const n = document.querySelectorAll('a[data-item-id][data-section-name="goods_list"]').length;
  const empty = /검색 결과가 없습니다/.test(document.body ? document.body.innerText : '');
  return { path: location.pathname, kw, n, empty, ready: document.readyState };
})()"""

JS_COLLECT = r"""(async (targets, BRAND) => {
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const got = new Map();
  const grab = () => {
    for (const a of document.querySelectorAll('a[data-item-id][data-section-name="goods_list"]')) {
      const id = a.dataset.itemId;
      const c = a.closest('[class*="GoodsItem"]') || a.parentElement;
      const lines = ((c && c.innerText) || '').split('\n').map(s => s.trim());
      const bt = lines.includes(BRAND);
      const name = lines.find(s => s && s !== BRAND && !/^\d+%$|원$|^\(|^[\d.,]+$|^(패션위크|아울렛|옵션|남성|여성|쿠폰|단독|무료배송|멤버스데이)$/.test(s)) || '';
      const p = got.get(id);
      if (!p) got.set(id, { id, idx: Number(a.dataset.itemListIndex) || got.size + 1, slug: a.dataset.itemBrand || '', bt, name });
      else { if (bt) p.bt = true; if (!p.name && name) p.name = name; }
    }
  };
  grab();
  let still = 0;
  while (got.size < 60 && !targets.every(t => got.has(t)) && still < 3) {
    const n = got.size;
    window.scrollBy(0, 1400); await sleep(450); grab();
    still = got.size === n ? still + 1 : 0;
  }
  window.scrollTo(0, 0);
  const txt = document.body.innerText;
  const empty = /검색 결과가 없습니다/.test(txt);
  const m = txt.match(/(?:^|\n)\s*([\d,]+)개\s*(?:\n|$)/);
  return { items: [...got.values()].sort((a, b) => a.idx - b.idx).slice(0, 60), empty, shown: m ? Number(m[1].replace(/,/g, '')) : null };
})"""


def find_chrome():
    c = []
    for env in ('PROGRAMFILES', 'PROGRAMFILES(X86)', 'LOCALAPPDATA'):
        b = os.environ.get(env)
        if b:
            c.append(os.path.join(b, 'Google', 'Chrome', 'Application', 'chrome.exe'))
    for env in ('PROGRAMFILES(X86)', 'PROGRAMFILES'):
        b = os.environ.get(env)
        if b:
            c.append(os.path.join(b, 'Microsoft', 'Edge', 'Application', 'msedge.exe'))
    c += ['/Applications/Google Chrome.app/Contents/MacOS/Google Chrome']
    for n in ('google-chrome', 'chrome', 'chromium', 'chromium-browser', 'msedge'):
        p = shutil.which(n)
        if p:
            c.append(p)
    for p in c:
        if p and os.path.exists(p):
            return p
    return None


def _free_port():
    s = socket.socket(); s.bind(('127.0.0.1', 0)); p = s.getsockname()[1]; s.close(); return p


_NOPROXY = urllib.request.build_opener(urllib.request.ProxyHandler({}))


def _get_json(url, timeout=2):
    with _NOPROXY.open(url, timeout=timeout) as r:
        return json.loads(r.read().decode('utf-8'))


class Browser:
    def __init__(self, profile_dir, chrome_path=None, headless=False):
        exe = chrome_path or find_chrome()
        if not exe:
            raise RuntimeError('크롬(또는 엣지)을 찾지 못했습니다. 크롬을 설치한 뒤 다시 실행해 주세요.')
        os.makedirs(profile_dir, exist_ok=True)
        self.port = _free_port()
        args = [exe, f'--remote-debugging-port={self.port}', '--remote-allow-origins=*', f'--user-data-dir={profile_dir}',
                '--no-first-run', '--no-default-browser-check', '--disable-popup-blocking', '--window-size=1400,1000',
                '--lang=ko-KR', 'about:blank']
        if headless:
            args[1:1] = ['--headless=new', '--no-sandbox']
        flags = 0x08000000 if sys.platform.startswith('win') else 0   # CREATE_NO_WINDOW (콘솔 창 안 뜨게)
        self.proc = subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=flags)
        ws = None
        for _ in range(60):
            try:
                pages = [p for p in _get_json(f'http://127.0.0.1:{self.port}/json') if p.get('type') == 'page']
                if pages:
                    ws = pages[0]['webSocketDebuggerUrl']; break
            except Exception:
                pass
            time.sleep(0.5)
        if not ws:
            self.close()
            raise RuntimeError('크롬을 원격 제어 모드로 열지 못했습니다. 열려 있는 이 프로그램용 크롬 창을 닫고 다시 시도해 주세요.')
        self.ws = websocket.create_connection(ws, timeout=60, suppress_origin=True)
        self._id = 0
        self.call('Page.enable'); self.call('Runtime.enable')

    def call(self, method, **params):
        self._id += 1; mid = self._id
        self.ws.send(json.dumps({'id': mid, 'method': method, 'params': params}))
        while True:
            msg = json.loads(self.ws.recv())
            if msg.get('id') == mid:
                if 'error' in msg:
                    raise RuntimeError(f"{method}: {msg['error'].get('message')}")
                return msg.get('result', {})

    def eval(self, expr, await_promise=False):
        r = self.call('Runtime.evaluate', expression=expr, returnByValue=True, awaitPromise=await_promise)
        if 'exceptionDetails' in r:
            raise RuntimeError(r['exceptionDetails'].get('text', 'JS 오류'))
        return r.get('result', {}).get('value')

    def goto(self, url, wait=15):
        self.call('Page.navigate', url=url)
        t = time.time()
        while time.time() - t < wait:
            time.sleep(0.4)
            try:
                if self.eval('document.readyState') == 'complete':
                    return True
            except Exception:
                pass
        return False

    def click(self, x, y):
        for t in ('mouseMoved', 'mousePressed', 'mouseReleased'):
            self.call('Input.dispatchMouseEvent', type=t, x=x, y=y, button='left', clickCount=1)

    def type_text(self, text):
        self.call('Input.insertText', text=text)

    def enter(self):
        k = dict(key='Enter', code='Enter', windowsVirtualKeyCode=13, nativeVirtualKeyCode=13)
        self.call('Input.dispatchKeyEvent', type='keyDown', text='\r', **k)
        self.call('Input.dispatchKeyEvent', type='keyUp', **k)

    def screenshot(self, path):
        data = self.call('Page.captureScreenshot', format='png')['data']
        with open(path, 'wb') as f:
            f.write(base64.b64decode(data))

    def close(self):
        try:
            self.ws.close()
        except Exception:
            pass
        try:
            self.proc.terminate()
        except Exception:
            pass


def _same_kw(a, b):
    n = lambda s: re.sub(r'\s+', ' ', str(s or '').replace('+', ' ')).strip().lower()
    return n(a) == n(b)


def search_once(br, keyword, base=BASE_URL):
    """검색창에 입력해서 검색. 실패하면 검색 결과 주소로 직접 이동(예비). 반환: 'ui' | 'url'"""
    how = 'url'
    pos = br.eval(JS_TRIGGER)
    if pos:
        br.click(pos['x'], pos['y'])
        for _ in range(15):
            time.sleep(0.2)
            if br.eval(JS_FOCUSED):
                br.type_text(keyword); time.sleep(0.25); br.enter(); how = 'ui'; break
    if how == 'url':
        br.call('Page.navigate', url=base + '/search/goods?' + urllib.parse.urlencode(
            {'keyword': keyword, 'keywordType': 'keyword', 'gf': 'A'}))
    # 결과 화면이 뜰 때까지
    t = time.time()
    while time.time() - t < 15:
        time.sleep(0.3)
        try:
            s = br.eval(JS_STATE)
        except Exception:
            continue
        if s and s['path'].startswith('/search') and _same_kw(s['kw'], keyword) and s['ready'] == 'complete' and (s['n'] or s['empty']):
            break
    time.sleep(0.6)
    return how


def collect(br, targets, brand):
    return br.eval(JS_COLLECT + f'({json.dumps(list(targets))}, {json.dumps(brand)})', await_promise=True)


DROP_WORDS = ['UNISEX', '유니섹스']
RETRY_LABEL = '검색됨 (UNISEX·유니섹스 빼고 재검색)'


def retry_keyword(kw, words=None, brand=''):
    """검색어에서 words(UNISEX·유니섹스 등)를 뺀 재검색어. 뺄 단어가 없거나 상품명이 비면 None"""
    words = words if words is not None else DROP_WORDS
    if not words:
        return None
    pat = r'(?<![0-9A-Za-z가-힣])(?:' + '|'.join(re.escape(w) for w in words) + r')(?![0-9A-Za-z가-힣])'
    k2 = re.sub(r'\s+', ' ', re.sub(pat, ' ', str(kw), flags=re.I)).strip(' -_')
    if not k2 or k2 == str(kw).strip() or k2.strip() == str(brand).strip():
        return None
    return k2


def _safe(s, n=50):
    return re.sub(r'[\\/:*?"<>|\s]+', '_', str(s)).strip('_')[:n]


def run(L, brand, out_dir, log=print, stop=None, progress=None, chrome_path=None, headless=False, delay=0.5, base=BASE_URL,
        drop_words=None):
    """
    L: 상품 목록 (상품번호, 상품명, 검색어)
    반환: (검색결과 텍스트 '상품번호|결과수|순위|브랜드상품수|재검색어' 줄들, {상품번호: 캡처 파일 상대경로})
    상품이 안 나오면 drop_words(UNISEX·유니섹스 등)를 뺀 검색어로 한 번 더 검색합니다.
    작업 중 결과는 work/ui_progress.json 에 저장해, 중간에 멈춰도 다시 누르면 이어서 합니다.
    """
    work = os.path.join(out_dir, 'work'); cap_dir = os.path.join(out_dir, '미노출_캡처')
    os.makedirs(work, exist_ok=True); os.makedirs(cap_dir, exist_ok=True)
    prog_path = os.path.join(work, 'ui_progress.json')
    try:
        prog = json.load(open(prog_path, encoding='utf-8'))
    except Exception:
        prog = {}
    groups = {}
    for no, name, kw in zip(L['상품번호'].astype(str), L['상품명'], L['검색어']):
        groups.setdefault(f'{brand} {kw}'.strip(), []).append((no, name))
    todo = [k for k in groups if k not in prog]
    log(f'검색어 {len(groups):,}개 (상품 {len(L):,}개)' + (f' — 이전에 한 {len(groups) - len(todo):,}개는 건너뜀' if len(todo) < len(groups) else ''))
    br = Browser(os.path.join(work, 'chrome_profile'), chrome_path, headless)
    slugs = set()
    try:
        br.goto(base + HOME_PATH)
        time.sleep(1.5)
        t0 = time.time(); fallback = 0
        for i, kw in enumerate(todo, 1):
            if stop is not None and stop.is_set():
                log('중지했습니다. 다시 누르면 이어서 검색합니다.'); break
            targets = [no for no, _ in groups[kw]]
            try:
                how = search_once(br, kw, base)
                res = collect(br, targets, brand)
            except Exception as e:
                log(f'  [{kw}] 오류: {e} → 한 번 더 시도'); time.sleep(2)
                br.goto(base + HOME_PATH); time.sleep(1)
                how = search_once(br, kw, base); res = collect(br, targets, brand)
            fallback += how == 'url'
            for it in res['items']:
                if it['bt'] and it['slug']:
                    slugs.add(it['slug'])
            ours = [it for it in res['items'] if it['bt'] or (it['slug'] and it['slug'] in slugs)]
            total = res['shown'] if res['shown'] is not None else (0 if res['empty'] else len(res['items']))
            rec = {'total': total, 'empty': res['empty'], 'ours': len(ours), 'rank': {}, 'cap': {},
                   'seen': [f"{it['id']}:{it['name']}" for it in res['items'][:10]]}
            ids = [it['id'] for it in res['items']]
            miss = []
            for no, name in groups[kw]:
                hit = next((it for it in ours if it['id'] == no), None)
                if hit:
                    rec['rank'][no] = ids.index(no) + 1
                else:
                    miss.append((no, name))
            rec['retry'] = {}
            kw2 = retry_keyword(kw, drop_words, brand) if miss else None
            if kw2:
                try:
                    search_once(br, kw2, base); r2 = collect(br, [no for no, _ in miss], brand)
                except Exception as e:
                    log(f'  [{kw2}] 재검색 오류: {e}'); r2 = None
                if r2:
                    for it in r2['items']:
                        if it['bt'] and it['slug']:
                            slugs.add(it['slug'])
                    ours2 = [it for it in r2['items'] if it['bt'] or (it['slug'] and it['slug'] in slugs)]
                    ids2 = [it['id'] for it in r2['items']]
                    tot2 = r2['shown'] if r2['shown'] is not None else (0 if r2['empty'] else len(r2['items']))
                    still = []
                    for no, name in miss:
                        if any(it['id'] == no for it in ours2):
                            rec['rank'][no] = ids2.index(no) + 1
                            rec['retry'][no] = {'kw': kw2, 'total': tot2, 'ours': len(ours2)}
                        else:
                            still.append((no, name))
                    miss = still
            if miss:
                p = os.path.join(cap_dir, f'{miss[0][0]}_{_safe(miss[0][1])}.png')
                br.screenshot(p)
                for no, name in miss:
                    q = os.path.join(cap_dir, f'{no}_{_safe(name)}.png')
                    if q != p:
                        shutil.copyfile(p, q)
                    rec['cap'][no] = os.path.relpath(q, out_dir)
            prog[kw] = rec
            if i % 10 == 0 or i == len(todo):
                json.dump(prog, open(prog_path, 'w', encoding='utf-8'), ensure_ascii=False)
                el = time.time() - t0; left = el / i * (len(todo) - i)
                log(f'  진행 {i:,}/{len(todo):,}  (남은 시간 약 {left / 60:.0f}분)')
            if progress:
                progress(i, len(todo))
            time.sleep(delay)
        json.dump(prog, open(prog_path, 'w', encoding='utf-8'), ensure_ascii=False)
        if fallback:
            log(f'  ※ {fallback}개는 검색창 입력이 안 돼 검색 결과 주소로 직접 열어 확인했습니다.')
    finally:
        br.close()
    lines, caps = [], {}
    for kw, items in groups.items():
        rec = prog.get(kw)
        if not rec:
            continue
        for no, _ in items:
            rt = rec.get('retry', {}).get(no)
            if rt:
                lines.append(f"{no}|{rt['total']}|{rec['rank'].get(no, '')}|{rt['ours']}|{rt['kw']}")
            else:
                lines.append(f"{no}|{rec['total']}|{rec['rank'].get(no, '')}|{rec['ours']}")
            if no in rec['cap']:
                caps[no] = rec['cap'][no]
    return '\n'.join(lines), caps
