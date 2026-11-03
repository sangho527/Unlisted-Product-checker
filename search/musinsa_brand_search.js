/*
 * 무신사 검색 노출 확인 스크립트 ("브랜드명 + 상품명" 검색)
 *
 * 사용법
 *   1) MusinsaChecker.exe 의 [① 상품 목록 만들기] (또는 python app/runner.py <재고관리_복사본.xlsx>)를 실행하면
 *      결과 폴더에 musinsa_brand_search_ready.js 가 만들어집니다. (검색할 상품번호·검색어가 들어가 있음)
 *      exe 에서는 [② 검색 스크립트 복사 + 무신사 열기] 버튼이 이 내용을 복사해 줍니다.
 *   2) Chrome 에서 https://www.musinsa.com 을 열고 F12 → Console 탭을 엽니다. (로그인 필요 없음)
 *   3) musinsa_brand_search_ready.js 내용을 통째로 붙여넣고 Enter.
 *   4) 끝나면 결과가 클립보드에 복사됩니다. (콘솔에도 출력)
 *      → exe 의 [③ 검색 결과 붙여넣기] (또는 텍스트 파일로 저장해 runner.py 세 번째 인자로)
 *
 * 판정 기준
 *   - 검색어: "<BRAND> <상품명>"  (상품명은 컬러명·_3COL 같은 컬러 묶음 표시를 뗀 이름)
 *   - 해당 무신사 상품번호가 검색 결과에 있고, 그 상품의 브랜드명이 BRAND 일 때만 '검색됨'
 *   - 고객이 보는 무신사 검색 결과 페이지(www.musinsa.com/search/goods)를 그대로 불러와,
 *     화면에 표시되는 상품 목록(첫 페이지, 최대 60개)에서 상품번호·브랜드명을 확인합니다.
 *     (검색 API 결과와 실제 화면이 다를 수 있어 화면 기준으로 확인합니다.
 *      예: 검색어에 '유니섹스'가 들어가면 API에는 결과가 있어도 웹 화면은 '검색 결과가 없습니다'로 표시됨)
 *   - 반드시 www.musinsa.com 페이지의 콘솔에서 실행하세요.
 *
 * 출력 형식 (한 줄에 한 상품)
 *   상품번호|검색결과수(페이지 제목의 '총 N건')|화면 순위(못 찾으면 빈칸)|화면에 보인 BRAND 상품 수
 */
(async () => {
  const BRAND = "어반드레스";
  const KEYWORDS = [];            // [["무신사 상품번호", "검색어(브랜드명 제외)"], ...]  ← exe [① 상품 목록 만들기]가 채워줍니다
  const DELAY_MS = 400;

  if (!/musinsa\.com$/.test(location.hostname)) { console.warn('www.musinsa.com 페이지에서 실행하세요.'); return; }
  if (!KEYWORDS.length) { console.warn('KEYWORDS 가 비어 있습니다. 결과 폴더의 musinsa_brand_search_ready.js 를 사용하세요.'); return; }
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const cache = {}, lines = [];

  for (let n = 0; n < KEYWORDS.length; n++) {
    const [goodsNo, name] = KEYWORDS[n];
    const kw = `${BRAND} ${name}`;
    if (!cache[kw]) {
      const html = await fetch('/search/goods?keyword=' + encodeURIComponent(kw) + '&gf=A').then((r) => r.text());
      const seen = new Set(), all = [];
      for (const m of html.matchAll(/\\?"goodsNo\\?":(\d+)[^{}]*?\\?"brandName\\?":\\?"([^"\\]+)/g)) {
        if (!seen.has(m[1])) { seen.add(m[1]); all.push({ no: m[1], brand: m[2] }); }
      }
      const t = (html.match(/총 ([\d,]+)건/) || [])[1];
      cache[kw] = { all, total: t ? Number(t.replace(/,/g, '')) : all.length };
      await sleep(DELAY_MS);
    }
    const c = cache[kw];
    const i = c.all.findIndex((x) => x.no === String(goodsNo) && x.brand === BRAND);
    const brandCnt = c.all.filter((x) => x.brand === BRAND).length;
    lines.push([goodsNo, c.total, i >= 0 ? i + 1 : '', brandCnt].join('|'));
    if ((n + 1) % 20 === 0) console.log(`진행 ${n + 1}/${KEYWORDS.length}`);
  }

  const text = lines.join('\n');
  window.__musinsaSearchResult = text;
  try { await navigator.clipboard.writeText(text); console.log('결과를 클립보드에 복사했습니다 → exe 에서 [③ 검색 결과 붙여넣기]를 누르세요.'); }
  catch (e) { console.log('클립보드 복사 실패 — 아래 결과를 직접 복사한 뒤 exe 에서 [③ 검색 결과 붙여넣기]를 누르세요.'); }
  const miss = lines.filter((l) => l.split('|')[2] === '');
  console.log(`완료: ${lines.length}개 중 검색 안 됨 ${miss.length}개`);
  console.log(text);
})();
