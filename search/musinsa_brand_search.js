/*
 * 무신사 검색 노출 확인 스크립트 ("브랜드명 + 상품명" 검색)
 *
 * 사용법
 *   1) python run_all.py 를 실행하면 output/musinsa_brand_search_ready.js 가 만들어집니다.
 *      (이 파일에 검색할 상품번호·검색어가 이미 들어가 있습니다)
 *   2) Chrome 에서 https://www.musinsa.com 을 열고 F12 → Console 탭을 엽니다. (로그인 필요 없음)
 *   3) musinsa_brand_search_ready.js 내용을 통째로 붙여넣고 Enter.
 *   4) 끝나면 결과가 클립보드에 복사됩니다. (콘솔에도 출력)
 *      → data/search_brand.txt 로 저장한 뒤 python run_all.py 를 다시 실행하면
 *        결과 엑셀에 '무신사_검색노출' 시트가 생깁니다.
 *
 * 판정 기준
 *   - 검색어: "<BRAND> <상품명>"  (상품명은 컬러명·_3COL 같은 컬러 묶음 표시를 뗀 이름)
 *   - 해당 무신사 상품번호가 검색 결과에 있고, 그 상품의 브랜드명이 BRAND 일 때만 '검색됨'
 *   - 무신사 검색 화면이 쓰는 것과 같은 검색 API(api.musinsa.com/api2/dp/v1/plp/goods)를 호출합니다.
 *     페이지당 60개, 최대 5페이지(300개)까지 확인합니다.
 *
 * 출력 형식 (한 줄에 한 상품)
 *   상품번호|검색결과수|순위(못 찾으면 빈칸)|검색결과 중 BRAND 상품 수
 */
(async () => {
  const BRAND = "어반드레스";
  const KEYWORDS = [];            // [["무신사 상품번호", "검색어(브랜드명 제외)"], ...]  ← step5 가 채워줍니다
  const MAX_PAGES = 5, SIZE = 60, DELAY_MS = 300;

  if (!KEYWORDS.length) { console.warn('KEYWORDS 가 비어 있습니다. output/musinsa_brand_search_ready.js 를 사용하세요.'); return; }
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const cache = {}, lines = [];

  for (let n = 0; n < KEYWORDS.length; n++) {
    const [goodsNo, name] = KEYWORDS[n];
    const kw = `${BRAND} ${name}`;
    if (!cache[kw]) {
      let all = [], total = 0;
      for (let page = 1; page <= MAX_PAGES; page++) {
        const url = 'https://api.musinsa.com/api2/dp/v1/plp/goods?gf=A&keyword=' + encodeURIComponent(kw) +
                    `&isUsed=false&page=${page}&size=${SIZE}&caller=SEARCH`;
        const j = await fetch(url, { credentials: 'include' }).then((r) => r.json());
        const d = j.data || {};
        total = d.pagination ? d.pagination.totalCount : total;
        all = all.concat((d.list || []).map((x) => ({ no: String(x.goodsNo), brand: x.brandName })));
        if (!(d.pagination && d.pagination.hasNext)) break;
        await sleep(DELAY_MS);
      }
      cache[kw] = { all, total };
      await sleep(DELAY_MS);
    }
    const c = cache[kw];
    const i = c.all.findIndex((x) => x.no === String(goodsNo) && x.brand === BRAND);
    const brandCnt = c.all.filter((x) => x.brand === BRAND).length;
    lines.push([goodsNo, c.total, i >= 0 ? i + 1 : '', brandCnt].join('|'));
    if ((n + 1) % 20 === 0) console.log(`진행 ${n + 1}/${KEYWORDS.length}`);
  }

  const text = lines.join('\n');
  try { await navigator.clipboard.writeText(text); console.log('결과를 클립보드에 복사했습니다 → data/search_brand.txt 로 저장하세요.'); }
  catch (e) { console.log('클립보드 복사 실패 — 아래 내용을 복사해 data/search_brand.txt 로 저장하세요.'); }
  const miss = lines.filter((l) => l.split('|')[2] === '');
  console.log(`완료: ${lines.length}개 중 검색 안 됨 ${miss.length}개`);
  console.log(text);
})();
