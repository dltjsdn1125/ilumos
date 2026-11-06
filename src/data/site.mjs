/**
 * site.mjs — single source of truth for every generated page.
 *
 * THEME: 위생용지 제조·공급 — (주)이루모스 / ILUMOS, 충남 계룡.
 * 키친타월 · 점보롤티슈 · 두루마리티슈.
 *
 * The LAYOUT, MOTION and DESIGN SYSTEM come from the cinetica.studio analysis in
 * ../../cineticaanalysis — that is what was measured and rebuilt. Only the
 * CONTENT was replaced for this renewal.
 *
 * ✔ 회사 식별 정보(상호·대표자·전화·팩스·이메일·소재지·취급 품목)는
 *   original/CONTENT.md 의 원본 값을 그대로 옮긴 것이며 임의로 바꾸지 않습니다.
 *
 * ⚠ 제품 규격 수치(길이·겹수·박스 입수)와 사용처별 제안의 숫자는 업계 일반
 *   규격을 기준으로 한 예시입니다. 실제 사양은 견적 단계에서 확정됩니다.
 *   See MISSING.md > D.
 */

/* ── brand / SEO ────────────────────────────────────────────────────────────── */

export const site = {
  name: '(주)이루모스',
  nameLatin: 'ILUMOS',
  title: '충남 계룡 위생용지 키친타월 · 점보롤티슈 | (주)이루모스',
  description:
    '충남 계룡의 위생용지 공급사 (주)이루모스입니다. 키친타월, 점보롤티슈, 두루마리티슈를 가정용 소량부터 업소용 대량 납품까지 공급하고, 로고 인쇄 주문제작도 상담해 드립니다.',
  /* 지금 실제로 열리는 주소. 정식 도메인(www.ilumos.co.kr)을 연결하면 이 한 줄만
     바꾸면 canonical · og:image · sitemap · JSON-LD 가 같이 따라옵니다. */
  origin: 'https://ilumos.pages.dev',
  /* 사이트맵 lastmod 기준일 — 내용을 실제로 고친 날에만 바꿉니다(빌드 날짜 아님). */
  lastmod: '2026-11-06',
  /* 영업시간 구조화 값 — `hours` 문구와 같은 내용입니다. */
  openingHours: { days: ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'], opens: '09:00', closes: '18:00' },
  locale: 'ko_KR',
  lang: 'ko',
  ogImage: '/assets/images/og-share.webp', // 1200×630 사본 (원본 og-cover.webp 는 세로형)
  city: '충남 계룡',
  ceo: '김상우',
  address: { street: '두마면 입암길 76-18 (입암리 653)', locality: '계룡시', region: '충청남도', country: 'KR' },
  email: 'edward1105@naver.com',
  phone: '063-243-3240',
  fax: '063-243-7730',
  hours: '평일 09:00–18:00 · 주말·공휴일 휴무',
  social: [],
}

/* ── navigation ─────────────────────────────────────────────────────────────── */

export const nav = [
  { label: '회사소개', href: '/#about' },
  { label: '제품소개', href: '/#services' },
  { label: '온라인 주문', href: '/shop/' },
  { label: '사용처 안내', href: '/portfolio/' },
  { label: '주문 · 배송', href: '/#process' },
  { label: '고객센터', href: '/faq/' },
  { label: '위생용지 저널', href: '/blog/' },
]

export const languages = [
  { code: 'KO', href: '/', current: true },
  { code: 'EN', href: '/#', current: false, note: 'EN 라우트는 이번 범위 밖 — MISSING.md 참고' },
]

/* ── hero ───────────────────────────────────────────────────────────────────── */

export const hero = {
  eyebrow: '(Kitchen towel · Jumbo roll · Toilet tissue)',
  words: ['EVERYDAY', 'ESSENTIAL', 'PAPER'],
  scrambleWords: ['ESSENTIAL', 'ABSORBENT', 'RELIABLE', 'ILUMOS'],
  lead: '충남 계룡의 위생용지 공급사 (주)이루모스입니다. 키친타월, 점보롤티슈, 두루마리티슈를 한 박스부터 정기 납품까지 필요한 만큼 제때 보내 드립니다.',
  cta: { label: '견적 요청', href: '/#contact' },
  ctaSecondary: { label: '제품 보기', href: '/#services' },
  scrollHint: 'SCROLL',
}

/* ── about ──────────────────────────────────────────────────────────────────── */

export const about = {
  eyebrow: '(Who we are)',
  heading: 'WE MAKE\nWHAT YOU USE\nEVERY DAY',
  lead: '',
  body: [
    '(주)이루모스는 키친타월, 점보롤티슈, 두루마리티슈를 공급하는 위생용지 기업입니다. 충청남도 계룡시에서 주방과 화장실, 매장과 사무실에서 매일 소비되는 종이를 다룹니다.',
    '매일 쓰는 물건일수록 차이는 사소한 곳에서 드러납니다. 한 장으로 닦이는지, 젖었을 때 찢어지지 않는지, 물에 제대로 풀리는지 — 품질을 판단하는 기준은 결국 이런 것들입니다.',
    '필요한 품목과 수량, 쓰시는 곳만 알려 주시면 규격과 단가를 정리해 회신드립니다. 한 박스로 시작하시든 정기 대량 납품이든 같은 기준으로 대응합니다.',
  ],
  stats: [
    { value: 3, suffix: '종', label: '주력 품목', labelEn: 'PRODUCT LINES' },
    { value: 24, suffix: '시간', label: '견적 회신', labelEn: 'QUOTE RESPONSE' },
    { value: 1, suffix: '박스', label: '최소 주문', labelEn: 'MIN ORDER' },
    { value: 365, suffix: '일', label: '연중 공급', labelEn: 'YEAR-ROUND' },
  ],
}

/* ── where it is used (logo wall) ───────────────────────────────────────────
   업종 카테고리로 표기합니다. 특정 거래처를 표기하지 않는 것은 위생용지 납품이
   대부분 유통 계약으로 묶여 있기 때문이며, 제3자 상표를 빌드에 넣지 않기
   위해서이기도 합니다. */

export const clients = {
  eyebrow: '(Where it is used)',
  heading: 'WHERE OUR\nPAPER GOES',
  /* 4열 × 2행 = 8칸, 한 칸에 2장씩 무한 순환하므로 정확히 16개입니다.
     순서가 곧 배치입니다 — 앞에서부터 두 개씩 묶여 한 칸이 됩니다. */
  /* `labelEn` 이 카드에 보이는 이름이고, `label` 은 alt 로 들어갑니다.
     사이트 전체가 영문 표기 + 한글 병기 규칙이라(C15) 카드도 같은 규칙을 따릅니다 —
     한글은 alt 에 남아 의미와 접근성이 사라지지 않습니다. */
  logos: [
    { labelEn: 'RESTAURANT', label: '외식 프랜차이즈', image: 'where-restaurant' },
    { labelEn: 'CATERING', label: '단체 급식', image: 'where-catering' },
    { labelEn: 'HEALTHCARE', label: '요양 · 병원', image: 'where-care' },
    { labelEn: 'OFFICE', label: '사무 빌딩', image: 'where-office' },
    { labelEn: 'SCHOOL', label: '학교 · 학원', image: 'where-school' },
    { labelEn: 'HOTEL', label: '호텔 · 펜션', image: 'where-hotel' },
    { labelEn: 'CAFE & BAKERY', label: '카페 · 베이커리', image: 'where-cafe' },
    { labelEn: 'LOGISTICS', label: '물류 · 공장', image: 'where-logistics' },
    { labelEn: 'PUBLIC SECTOR', label: '공공기관', image: 'where-public' },
    { labelEn: 'RETAIL', label: '마트 · 유통', image: 'where-market' },
    { labelEn: 'SPORTS', label: '스포츠 시설', image: 'where-sports' },
    { labelEn: 'CAMPING', label: '캠핑 · 레저', image: 'where-camping' },
    { labelEn: 'KINDERGARTEN', label: '어린이집 · 유치원', image: 'where-kindergarten' },
    { labelEn: 'SALON', label: '미용실 · 살롱', image: 'where-salon' },
    { labelEn: 'FITNESS', label: '헬스장', image: 'where-fitness' },
    { labelEn: 'CONVENIENCE', label: '편의점 · 소매', image: 'where-store' },
  ],
}

/* ── products (8) ───────────────────────────────────────────────────────────── */

export const servicesMeta = {
  eyebrow: '(What we supply)',
  heading: 'OUR\nPRODUCTS',
  lead: '주방에서 화장실까지, 여덟 개 품목을 다룹니다. 규격과 겹수는 쓰시는 곳에 맞춰 골라 드립니다.',
}

export const services = [
  {
    slug: 'kitchen-towel',
    summaryEn: 'Kitchen towel that absorbs in one sheet and holds together when wet.',
    num: '01',
    title: '키친타월',
    titleLatin: 'KITCHEN TOWEL',
    tagline: '한 장으로 끝나야 진짜 흡수력입니다',
    summary:
      '가정용부터 업소용 대용량까지, 기름과 물을 빠르게 빨아들이는 키친타월을 공급합니다.',
    body: [
      '키친타월은 두께가 아니라 구조로 흡수합니다. 엠보싱으로 만들어진 미세한 공간이 물과 기름을 빨아들이기 때문에, 같은 무게라도 엠보싱이 살아 있는 쪽이 훨씬 빨리 흡수합니다.',
      '젖은 상태에서 찢어지면 흡수력이 좋아도 소용이 없습니다. 습윤 강도를 함께 확인해 튀김 기름을 걷어내거나 젖은 조리대를 닦을 때 한 장이 끝까지 버티도록 맞춥니다.',
      '식품에 닿는 용도이므로 형광증백제를 쓰지 않은 원지를 기본으로 안내드립니다. 사용량이 많은 업소는 대용량 롤로 바꾸는 것만으로도 교체 부담이 크게 줄어듭니다.',
    ],
    deliverables: ['가정용 롤형', '업소용 대용량 롤', '2겹 · 3겹 선택', '박스 단위 납품'],
    tools: ['엠보싱 흡수 구조', '습윤 강도 보강', '무형광 원지', '낱개 포장 옵션'],
    image: 'service-parts',
  },
  {
    slug: 'jumbo-roll',
    summaryEn: 'High-capacity jumbo rolls that cut how often someone has to refill.',
    num: '02',
    title: '점보롤티슈',
    titleLatin: 'JUMBO ROLL TISSUE',
    tagline: '교체 횟수를 줄이는 가장 확실한 방법',
    summary:
      '화장실 관리 부담을 줄이는 업소용 대용량 점보롤입니다. 디스펜서 규격에 맞춰 공급합니다.',
    body: [
      '점보롤의 목적은 하나입니다. 누군가 화장실에 들어가 휴지를 갈아 끼우는 횟수를 줄이는 것. 일반 두루마리보다 권취 길이가 훨씬 길기 때문에 이용이 잦은 곳일수록 효과가 분명합니다.',
      '디스펜서와 규격이 맞지 않으면 아무리 좋은 롤이라도 걸리거나 헛돕니다. 기존에 설치된 디스펜서의 내경과 폭을 먼저 확인하고 그에 맞는 코어 규격으로 보내 드립니다.',
      '심지 없는 무심지 타입은 폐기물이 줄고 마지막 한 장까지 쓸 수 있습니다. 다만 전용 디스펜서가 필요하므로 신규 설치 현장에 권해 드립니다.',
    ],
    deliverables: ['업소용 점보롤', '디스펜서 호환 규격', '유심지 · 무심지', '박스 12롤 단위'],
    tools: ['대용량 권취', '표준 코어 내경', '2겹 엠보싱', '낱개 비닐 포장'],
    image: 'service-clean',
  },
  {
    slug: 'toilet-tissue',
    summaryEn: 'Toilet tissue balanced between strength in hand and breaking down in water.',
    num: '03',
    title: '두루마리티슈',
    titleLatin: 'TOILET TISSUE',
    tagline: '매일 쓰는 것이라 기준이 더 분명해야 합니다',
    summary:
      '가정용·업소용 두루마리티슈를 공급합니다. 쓸 때는 부서지지 않고 물에서는 잘 풀리도록 맞춥니다.',
    body: [
      '두루마리티슈는 상반된 두 가지를 동시에 만족해야 합니다. 손에 쥐었을 때는 버텨야 하고, 변기에 들어가면 빨리 풀려야 합니다. 어느 한쪽으로 치우치면 불편하거나 막힙니다.',
      '피부에 직접 닿는 품목이라 원지 선택이 중요합니다. 무형광 천연펄프 제품과 사용량이 많은 곳을 위한 재생지 제품을 함께 취급하니, 용도와 예산에 맞춰 골라 드립니다.',
      '가정용과 업소용은 롤 지름과 권취 길이가 다릅니다. 화장실 홀더 크기에 맞지 않으면 매번 걸리므로, 쓰시는 홀더 형태를 알려 주시면 맞는 규격으로 안내드립니다.',
    ],
    deliverables: ['가정용 30m급', '업소용 대용량', '무형광 · 재생지 선택', '박스 30롤 단위'],
    tools: ['수분 해리성 관리', '2겹 · 3겹 구성', '엠보싱 부드러움', '개별 비닐 포장'],
    image: 'service-ceramic',
  },
  {
    slug: 'facial-napkin',
    summaryEn: 'Facial tissue and table napkins for counters, tables and guest areas.',
    num: '04',
    title: '미용티슈 · 냅킨',
    titleLatin: 'FACIAL TISSUE & NAPKIN',
    tagline: '손님 손에 가장 먼저 닿는 종이입니다',
    summary:
      '테이블과 카운터에 두는 각티슈, 미용티슈, 식탁 냅킨을 공급합니다.',
    body: [
      '미용티슈와 냅킨은 매장의 인상을 만드는 소모품입니다. 얇아서 두세 장씩 뽑아 쓰게 되면 결국 비용도 인상도 함께 나빠집니다.',
      '한 장으로 충분한 두께와, 뽑을 때 다음 장이 따라 나오는 팝업 구조를 함께 봅니다. 카운터에 놓는 각티슈는 케이스 크기와 인출 방향까지 확인해 드립니다.',
      '식탁 냅킨은 접이 방식과 크기 선택지가 넓습니다. 테이블 세팅에 맞춰 골라 드리고, 로고 인쇄가 필요하면 주문제작으로 진행합니다.',
    ],
    deliverables: ['각티슈 · 미용티슈', '식탁용 냅킨', '팝업 인출 구조', '접이 · 규격 선택'],
    tools: ['2겹 부드러움', '팝업 적층', '무형광 원지', '로고 인쇄 가능'],
    image: 'service-inspection',
  },
  {
    slug: 'hand-towel',
    summaryEn: 'Folded hand towels sized to the dispenser, not the other way round.',
    num: '05',
    title: '핸드타월',
    titleLatin: 'HAND TOWEL',
    tagline: '한 번에 한 장만 나와야 낭비가 없습니다',
    summary:
      '화장실과 주방 세면대에 두는 접이식 핸드타월을 공급합니다. 디스펜서 규격에 맞춰 납품합니다.',
    body: [
      '핸드타월에서 가장 자주 생기는 문제는 한 번에 여러 장이 딸려 나오는 것입니다. 접이 방식과 디스펜서 조합이 맞지 않을 때 생기며, 그대로 두면 소비량이 배로 늘어납니다.',
      '쓰시는 디스펜서가 어떤 접이 방식용인지 확인하고 그에 맞는 제품을 보내 드립니다. 디스펜서가 없는 현장은 그냥 쌓아 두고 쓸 수 있는 형태로 안내드립니다.',
      '젖은 손을 닦는 용도이므로 습윤 강도가 중요합니다. 한 장으로 손이 마르지 않으면 두 장을 쓰게 되고, 결국 저렴한 제품이 더 비싸집니다.',
    ],
    deliverables: ['접이식 핸드타월', '디스펜서 호환 규격', '평판 적층형', '박스 단위 납품'],
    tools: ['습윤 강도 관리', '접이 방식 선택', '한 장 인출 구조', '무형광 원지'],
    image: 'service-coating',
  },
  {
    slug: 'baking-paper',
    summaryEn: 'Baking parchment and food-contact paper for kitchens that use it daily.',
    num: '06',
    title: '주방용 종이호일',
    titleLatin: 'BAKING PAPER',
    tagline: '눌어붙지 않는 것이 전부인 품목입니다',
    summary:
      '오븐과 팬에 쓰는 종이호일, 식품 포장용 유산지를 주방 단위로 공급합니다.',
    body: [
      '종이호일은 기능이 단순한 만큼 실패도 분명합니다. 눌어붙거나, 열에 갈변하거나, 기름이 배어 나오면 그 자리에서 못 쓰게 됩니다.',
      '실리콘 처리가 고르게 되어 있는지가 대부분을 결정합니다. 사용하시는 조리 온도와 시간을 알려 주시면 그 범위에서 문제가 없는 제품으로 안내드립니다.',
      '업소는 롤형이, 소량 사용처는 시트형이 편합니다. 오븐팬 규격에 맞춰 재단된 시트형은 매번 자르는 시간을 아껴 줍니다.',
    ],
    deliverables: ['롤형 종이호일', '재단 시트형', '식품 포장용 유산지', '오븐팬 규격 재단'],
    tools: ['양면 실리콘 처리', '내열 온도 확인', '식품 접촉 적합', '주방 규격 재단'],
    image: 'service-vacuum',
  },
  {
    slug: 'private-label',
    summaryEn: 'Private-label printing on the wrap and the box, from one design file.',
    num: '07',
    title: 'OEM · 주문제작',
    titleLatin: 'PRIVATE LABEL',
    tagline: '같은 제품이라도 포장이 브랜드를 만듭니다',
    summary:
      '낱개 포장과 박스에 로고·브랜드를 인쇄해 납품합니다. 도안만 주시면 시안부터 잡아 드립니다.',
    body: [
      '프랜차이즈나 숙박업소처럼 손님 손에 제품이 그대로 놓이는 곳에서는 포장이 곧 브랜드입니다. 같은 두루마리티슈라도 인쇄 하나로 인상이 달라집니다.',
      '보유하신 로고 파일만 보내 주시면 인쇄 위치와 크기를 잡은 시안을 먼저 보여 드립니다. 확정된 뒤에 생산에 들어가므로 받아 보고 놀라는 일이 없습니다.',
      '인쇄 도수와 수량에 따라 단가와 최소 수량이 달라집니다. 처음 진행하시는 경우 도수를 줄이는 쪽이 비용 대비 만족도가 높은 편입니다.',
    ],
    deliverables: ['낱개 포장 인쇄', '박스 인쇄', '시안 제작 · 확인', '전용 규격 제작'],
    tools: ['1~4도 인쇄', '로고 데이터 검수', '샘플 선출고', '최소 수량 안내'],
    image: 'service-packaging',
  },
  {
    slug: 'bulk-supply',
    summaryEn: 'Standing orders sized to real consumption, so nothing runs out.',
    num: '08',
    title: '대량 공급 · 정기 배송',
    titleLatin: 'BULK SUPPLY',
    tagline: '떨어질 때마다 주문하지 않아도 되게',
    summary:
      '사용량에 맞춰 주기와 수량을 정해 두고 정기적으로 보내 드립니다. 전국 배송합니다.',
    body: [
      '위생용지는 떨어지면 바로 곤란해지는 물건입니다. 그래서 대부분의 현장이 필요 이상으로 쌓아 두거나, 반대로 급하게 사러 나가는 일을 반복합니다.',
      '두세 번 납품하는 동안 실제 사용량이 잡힙니다. 그 수치로 주기와 수량을 정해 두면 주문하실 필요가 없어지고 보관 공간도 함께 줄어듭니다.',
      '사용량이 계절이나 행사에 따라 크게 바뀌는 곳은 미리 알려 주시면 그 회차만 조정합니다. 정기 배송이라고 해서 고정된 수량에 묶이지는 않습니다.',
    ],
    deliverables: ['정기 배송 계약', '사용량 기반 수량 산정', '전국 배송', '회차별 수량 조정'],
    tools: ['납품 주기 설정', '품목 혼합 배송', '박스 단위 출고', '담당자 직통 응대'],
    image: 'service-reliability',
  },
]

/* ── use cases ──────────────────────────────────────────────────────────────── */

export const portfolioMeta = {
  eyebrow: '(Where it is used)',
  heading: 'USE\nCASES',
  lead: '업종별로 무엇을 어디에 두면 되는지 정리했습니다. 숫자는 일반적인 규격을 기준으로 한 예시이며, 실제 사양은 견적 단계에서 확정됩니다.',
}

export const portfolio = [
  {
    slug: 'restaurant-kitchen',
    title: '외식업장 주방',
    titleLatin: 'RESTAURANT KITCHEN',
    sector: '외식 · 급식',
    scale: '홀 60석 · 주방 4인 기준',
    category: '키친타월 · 핸드타월',
    tags: ['주방', '기름', '대용량'],
    summary:
      '기름과 물을 동시에 다루는 주방에서 어떤 규격을 어디에 두면 되는지 정리한 구성안.',
    challenge:
      '주방은 마른 것을 닦는 일과 기름을 걷어내는 일이 하루 종일 섞입니다. 얇은 타월을 두면 두세 장씩 쓰게 되고, 세면대와 조리대에 같은 제품을 두면 어느 한쪽이 항상 부족합니다.',
    approach:
      '조리대에는 흡수 속도와 습윤 강도가 높은 3겹 업소용 대용량 키친타월을, 세면대에는 한 장씩 뽑히는 접이식 핸드타월을 나눠 배치하는 구성을 권해 드립니다. 튀김기 주변처럼 소모가 집중되는 자리는 롤을 하나 더 두는 편이 결과적으로 저렴합니다.',
    outcome:
      '자리마다 맞는 제품이 놓이면 한 장으로 끝나는 일이 늘고 총 사용량이 줄어듭니다. 다만 조리 품목에 따라 소모 비율이 달라지므로 첫 두 번은 넉넉히 잡고 실제 사용량을 확인하시는 편이 좋습니다.',
    stats: [
      { value: 3, suffix: '겹', label: '조리대 권장', labelEn: 'PLY' },
      { value: 2, suffix: '주', label: '납품 주기', labelEn: 'DELIVERY CYCLE' },
      { value: 2, suffix: '개소', label: '배치 지점', labelEn: 'PLACEMENT' },
    ],
    services: ['kitchen-towel', 'hand-towel'],
    image: 'work-etch',
  },
  {
    slug: 'care-facility',
    title: '요양 · 의료시설',
    titleLatin: 'CARE FACILITY',
    sector: '요양 · 병원',
    scale: '입소 80명 · 3개 층 기준',
    category: '두루마리티슈 · 핸드타월',
    tags: ['요양', '무형광', '정기 납품'],
    summary:
      '피부에 닿는 품목과 소모가 많은 품목을 나눠 구성한, 요양·의료시설용 공급안.',
    challenge:
      '입소자가 직접 쓰는 품목과 직원이 관리하는 품목은 기준이 다릅니다. 앞쪽은 부드러움과 무형광이 중요하고, 뒤쪽은 교체 부담과 단가가 중요합니다.',
    approach:
      '거주 공간 화장실에는 무형광 천연펄프 두루마리티슈를, 공용 화장실과 직원 동선에는 점보롤과 접이식 핸드타월을 두는 구성을 권해 드립니다. 층별 사용량이 크게 다르므로 첫 납품은 층 단위로 나눠 확인합니다.',
    outcome:
      '품목을 나누면 사용자 만족도를 유지하면서 전체 비용을 낮출 수 있습니다. 다만 감염 관리 기준이 있는 시설은 해당 기준을 먼저 확인해 주셔야 합니다.',
    stats: [
      { value: 3, suffix: '개 층', label: '층별 구성', labelEn: 'FLOORS' },
      { value: 2, suffix: '종', label: '품목 분리', labelEn: 'SPLIT LINES' },
      { value: 4, suffix: '주', label: '납품 주기', labelEn: 'DELIVERY CYCLE' },
    ],
    services: ['toilet-tissue', 'hand-towel'],
    image: 'work-cleaning',
  },
  {
    slug: 'office-building',
    title: '사무실 · 공공기관',
    titleLatin: 'OFFICE & PUBLIC',
    sector: '오피스 · 공공',
    scale: '상주 200명 · 화장실 8개소',
    category: '점보롤티슈 · 미용티슈',
    tags: ['점보롤', '디스펜서', '관리'],
    summary:
      '화장실 관리 인력이 한정된 건물에서 교체 횟수를 줄이는 데 초점을 맞춘 구성안.',
    challenge:
      '이용 인원이 많은 건물은 두루마리티슈로 버티면 하루에도 몇 번씩 갈아 끼워야 합니다. 관리 인력이 그 시간에 묶이고, 비어 있는 시간대도 생깁니다.',
    approach:
      '화장실에는 대용량 점보롤과 디스펜서를, 회의실과 안내 데스크에는 각티슈를 두는 구성을 권해 드립니다. 기존 디스펜서가 있으면 내경과 폭을 먼저 확인해 맞는 코어 규격으로 보내 드립니다.',
    outcome:
      '권취 길이가 늘어난 만큼 교체 주기가 길어지고 관리 동선이 단순해집니다. 다만 디스펜서 규격이 맞지 않으면 효과가 사라지므로 확인이 먼저입니다.',
    stats: [
      { value: 8, suffix: '개소', label: '화장실 수', labelEn: 'RESTROOMS' },
      { value: 12, suffix: '롤', label: '박스 입수', labelEn: 'PER BOX' },
      { value: 4, suffix: '주', label: '납품 주기', labelEn: 'DELIVERY CYCLE' },
    ],
    services: ['jumbo-roll', 'facial-napkin'],
    image: 'work-foundry',
  },
  {
    slug: 'cafe-lodging',
    title: '카페 · 숙박업소',
    titleLatin: 'CAFE & LODGING',
    sector: '카페 · 숙박',
    scale: '객실 24실 · 카페 40석',
    category: '냅킨 · 두루마리티슈',
    tags: ['냅킨', '로고 인쇄', '객실'],
    summary:
      '손님 손에 제품이 그대로 놓이는 공간을 위해 인쇄 주문제작까지 함께 잡은 구성안.',
    challenge:
      '카페 테이블과 객실은 제품이 그대로 보이는 자리입니다. 포장이나 냅킨이 어울리지 않으면 공간을 아무리 잘 꾸며도 그 부분에서 눈에 걸립니다.',
    approach:
      '테이블 냅킨과 객실 두루마리티슈에 로고를 인쇄하는 주문제작으로 진행하고, 백오피스와 직원 화장실은 인쇄 없는 일반 제품으로 나눕니다. 인쇄 도수를 낮추면 비용 부담 없이 통일감을 만들 수 있습니다.',
    outcome:
      '손님 동선에만 인쇄 제품을 쓰면 비용을 크게 늘리지 않고도 공간이 정돈됩니다. 다만 인쇄 주문은 최소 수량이 있으므로 회전이 느린 품목은 신중히 정하셔야 합니다.',
    stats: [
      { value: 24, suffix: '실', label: '객실 기준', labelEn: 'ROOMS' },
      { value: 2, suffix: '도', label: '권장 인쇄', labelEn: 'PRINT COLORS' },
      { value: 6, suffix: '주', label: '납품 주기', labelEn: 'DELIVERY CYCLE' },
    ],
    services: ['facial-napkin', 'private-label'],
    image: 'work-hbm',
  },
]

/* ── process ────────────────────────────────────────────────────────────────── */

export const process = {
  eyebrow: '(How we work)',
  heading: 'FROM CALL\nTO DOOR',
  lead: '네 단계. 견적부터 정기 배송까지 담당자가 바뀌지 않습니다.',
  steps: [
    {
      num: '01',
      title: '용도 확인',
      titleLatin: 'WHAT YOU NEED',
      body: '어디에서 몇 분이 쓰시는지부터 여쭙습니다. 같은 두루마리티슈라도 가정과 업소는 필요한 규격이 다릅니다.',
    },
    {
      num: '02',
      title: '샘플 · 견적',
      titleLatin: 'SAMPLE & QUOTE',
      body: '해당 품목의 샘플과 규격표를 먼저 보내 드립니다. 직접 써 보신 뒤에 수량을 정하시면 됩니다.',
    },
    {
      num: '03',
      title: '생산 · 포장',
      titleLatin: 'PRODUCTION',
      body: '주문 수량에 맞춰 생산하고 박스 단위로 포장합니다. 로고 인쇄가 있으면 이 단계에서 시안을 확정합니다.',
    },
    {
      num: '04',
      title: '배송 · 정기 공급',
      titleLatin: 'DELIVERY',
      body: '납품일에 맞춰 전국으로 출고합니다. 사용량이 파악되면 정기 배송으로 전환해 재고 부담을 덜어 드립니다.',
    },
  ],
}

/* ── FAQ ────────────────────────────────────────────────────────────────────── */

export const faqMeta = {
  eyebrow: '(FAQ)',
  // 04-pages/home.json > a11y: on the analysed page the single <h1> sits on the
  // FAQ CTA block, not the visual hero. Reproduced here.
  heading: 'Got Questions?',
  lead: '견적을 요청하시기 전에 가장 많이 받는 질문들입니다.',
}

export const faq = [
  {
    q: '최소 주문 수량이 어떻게 되나요?',
    a: '품목에 따라 다르지만 대부분 1박스부터 주문하실 수 있습니다. 처음에는 소량으로 써 보시고 수량을 늘리시는 편을 권해 드립니다. 다만 로고 인쇄가 들어가는 주문제작은 최소 수량이 별도로 발생합니다.',
  },
  {
    q: '샘플을 먼저 받아볼 수 있나요?',
    a: '가능합니다. 사용하시는 곳과 관심 있는 품목을 알려 주시면 해당 규격의 샘플과 규격표를 함께 보내 드립니다. 매일 쓰는 물건이라 직접 써 보시고 결정하시는 것이 가장 확실합니다.',
  },
  {
    q: '견적은 어떻게 요청하나요?',
    a: '홈페이지의 견적 요청 양식, 전화 063-243-3240, 이메일 edward1105@naver.com 중 편하신 방법으로 주시면 됩니다. 품목·수량·사용처만 적어 주셔도 회신이 가능하며, 보통 24시간 안에 답변드립니다.',
  },
  {
    q: '배송은 어디까지 되나요? 기간은 얼마나 걸리나요?',
    a: '전국으로 배송합니다. 재고가 있는 품목은 주문 확인 후 통상 2~3일 내 출고되며, 주문 생산이나 로고 인쇄가 포함되면 일정이 더 필요합니다. 정확한 일정은 견적 회신 시 함께 안내드립니다.',
  },
  {
    q: '로고나 브랜드를 넣은 주문제작도 되나요?',
    a: '가능합니다. 낱개 포장이나 박스에 인쇄하는 형태로 진행하며, 보유하신 로고 파일을 보내 주시면 인쇄 위치와 크기를 잡은 시안을 먼저 보여 드립니다. 인쇄 도수와 수량에 따라 단가와 최소 수량이 달라집니다.',
  },
  {
    q: '점보롤과 두루마리티슈 중 어느 쪽을 써야 하나요?',
    a: '이용 인원과 디스펜서 유무로 나뉩니다. 화장실 이용이 많고 디스펜서를 설치할 수 있다면 점보롤이 교체 부담을 크게 줄여 줍니다. 가정이나 소규모 매장은 두루마리티슈가 더 적합합니다. 판단이 어려우시면 사용 인원만 알려 주세요.',
  },
  {
    q: '천연펄프와 재생지 중 무엇을 고르면 되나요?',
    a: '용도에 따라 다릅니다. 피부에 직접 닿는 미용티슈와 두루마리티슈는 형광증백제를 쓰지 않은 천연펄프를 권해 드리고, 소모가 많은 화장실 점보롤이나 청소용은 재생지 제품이 비용 면에서 유리합니다. 두 가지 모두 취급합니다.',
  },
  {
    q: '정기적으로 받아볼 수 있나요?',
    a: '가능합니다. 두세 번 납품하면서 실제 사용량이 파악되면 주기와 수량을 정해 정기 배송으로 전환해 드립니다. 매번 주문하실 필요가 없어지고 재고가 떨어지는 일도 줄어듭니다. 회차별 수량 조정도 가능합니다.',
  },
]

/* ── contact / footer ───────────────────────────────────────────────────────── */

export const contact = {
  eyebrow: '(Get a quote)',
  heading: 'REQUEST\nA QUOTE',
  lead: '필요한 품목과 수량, 쓰시는 곳만 알려 주세요. 규격과 단가를 정리해 회신드립니다.',
  cta: { label: '견적 요청', href: `mailto:${site.email}` },
}

export const marquee = [
  'KITCHEN TOWEL', 'JUMBO ROLL', 'TOILET TISSUE', 'FACIAL TISSUE', 'NAPKIN', 'HAND TOWEL', 'BAKING PAPER', 'PRIVATE LABEL',
]

/* ── journal ────────────────────────────────────────────────────────────────── */

export const blogMeta = {
  eyebrow: '(Paper journal)',
  heading: 'JOURNAL',
  lead: '위생용지를 고르고 쓰는 데 실제로 도움이 되는 것들만 적습니다.',
}

export const posts = [
  {
    slug: 'jumbo-vs-roll',
    title: '점보롤과 두루마리, 어디에 무엇을 놓아야 하나',
    titleLatin: 'JUMBO OR ROLL',
    date: '2025-11-27',
    category: '선택 기준',
    image: 'service-clean',
    readingTime: '7분',
    excerpt:
      '둘의 차이는 품질이 아니라 권취 길이와 디스펜서입니다. 이용 인원으로 갈라 보는 기준을 정리했습니다.',
    body: [
      { h: '같은 휴지, 다른 물건', p: ['점보롤과 두루마리티슈는 원지도 겹수도 비슷할 수 있습니다. 결정적인 차이는 한 롤에 감겨 있는 길이이고, 그 길이가 "누가 얼마나 자주 갈아 끼우는가"를 정합니다.', '그래서 이 선택은 품질 비교가 아니라 관리 방식의 선택에 가깝습니다.'] },
      { h: '이용 인원으로 갈린다', p: ['하루 이용자가 많지 않은 가정이나 소규모 매장에서는 두루마리티슈로 충분합니다. 홀더가 이미 있고 규격이 표준이라 추가 설치도 필요 없습니다.', '반대로 상주 인원이 수십 명을 넘어가는 사무실이나 공공시설에서는 두루마리로 버티면 하루에도 몇 번씩 채워야 합니다. 이때부터 점보롤이 의미를 갖습니다.'] },
      { h: '디스펜서가 전제 조건', p: ['점보롤은 전용 디스펜서가 있어야 제대로 걸립니다. 내경과 폭이 맞지 않으면 걸리거나 헛돌고, 결국 손으로 뜯어 쓰게 되어 아무 효과가 없습니다.', '기존 디스펜서가 있다면 안쪽 지름과 롤이 들어가는 폭을 재서 알려 주세요. 신규 설치라면 무심지 타입까지 고려해 볼 수 있습니다.'] },
      { h: '섞어 쓰는 것이 보통', p: ['한 건물 안에서도 자리마다 답이 다릅니다. 공용 화장실은 점보롤, 임원실이나 손님용 화장실은 두루마리로 나누는 구성이 실제로는 가장 흔합니다.', '전부 하나로 통일하려는 시도가 오히려 불편을 만드는 경우를 자주 봅니다.'] },
    ],
  },
  {
    slug: 'kitchen-towel-absorption',
    title: '키친타월 흡수력은 두께가 아니라 구조가 정한다',
    titleLatin: 'ABSORPTION BY STRUCTURE',
    date: '2025-10-14',
    category: '제품',
    image: 'service-parts',
    readingTime: '6분',
    excerpt:
      '두꺼운데 잘 안 빨아들이는 제품이 있는 이유. 엠보싱이 만든 공간이 실제 흡수량을 결정합니다.',
    body: [
      { h: '두께와 흡수는 다른 이야기', p: ['키친타월을 고를 때 두께부터 보게 되지만, 두껍다고 더 빨아들이지는 않습니다. 종이가 물을 머금는 것은 섬유 사이의 빈 공간이고, 그 공간은 두께가 아니라 엠보싱 구조가 만듭니다.', '눌려서 납작해진 타월은 두꺼워도 흡수가 느립니다.'] },
      { h: '한 장으로 끝나는가', p: ['현장에서 의미 있는 기준은 딱 하나입니다. 쏟은 것을 한 장으로 닦아낼 수 있는가. 두 장을 쓰게 되는 순간, 장당 단가가 아무리 싸도 총비용은 올라갑니다.', '샘플을 받으시면 물보다 기름으로 먼저 시험해 보시길 권합니다. 차이가 훨씬 분명하게 드러납니다.'] },
      { h: '젖었을 때 버티는가', p: ['흡수와 함께 봐야 하는 것이 습윤 강도입니다. 물을 머금은 순간 찢어지면 그 타월은 조리대에서 쓸 수 없습니다.', '젖은 상태로 살짝 문질러 보면 바로 알 수 있습니다. 보풀이 뭉치듯 떨어져 나오면 주방용으로는 부적합합니다.'] },
      { h: '식품에 닿는다는 전제', p: ['키친타월은 조리 중인 음식에 직접 닿는 일이 많습니다. 형광증백제를 쓰지 않은 원지를 기본으로 보시는 편이 안전합니다.', '유난히 새하얀 제품이 항상 좋은 제품은 아니라는 점만 기억해 두시면 됩니다.'] },
    ],
  },
  {
    slug: 'ply-difference',
    title: '2겹과 3겹, 실제로 달라지는 것',
    titleLatin: 'TWO PLY OR THREE',
    date: '2025-08-05',
    category: '규격',
    image: 'service-ceramic',
    readingTime: '6분',
    excerpt:
      '겹수는 품질 등급이 아니라 용도 구분에 가깝습니다. 늘려서 이득인 자리와 그렇지 않은 자리.',
    body: [
      { h: '겹수는 등급이 아니다', p: ['3겹이 2겹보다 항상 좋은 제품인 것은 아닙니다. 겹수는 두께와 부드러움을 만드는 방식일 뿐, 용도에 따라 유리한 쪽이 갈립니다.', '실제로 겹수를 올렸다가 오히려 불편해지는 경우도 있습니다.'] },
      { h: '올려서 이득인 자리', p: ['기름을 걷어내는 조리대, 손을 닦는 세면대처럼 한 번에 많은 양을 흡수해야 하는 자리에서는 겹수를 올리는 편이 확실히 낫습니다. 장수가 줄어 총 사용량이 오히려 감소합니다.', '손님이 직접 만지는 냅킨과 객실 티슈도 촉감 차이가 바로 느껴지는 자리입니다.'] },
      { h: '올릴 필요가 없는 자리', p: ['소모가 많은 공용 화장실은 겹수를 올려도 사용 장수가 크게 줄지 않는 경우가 많습니다. 단가만 오르고 체감은 크지 않습니다.', '변기에 흘려보내는 품목은 겹수가 두꺼울수록 풀리는 데 시간이 더 걸린다는 점도 함께 봐야 합니다.'] },
      { h: '정하기 어려울 때', p: ['같은 원지의 2겹과 3겹을 함께 받아 보시고 한 주만 같은 자리에 두어 보시면 답이 나옵니다. 어느 쪽이 더 빨리 줄어드는지가 곧 결론입니다.', '샘플은 두 규격을 함께 보내 드립니다.'] },
    ],
  },
  {
    slug: 'dispenser-placement',
    title: '디스펜서 위치가 소비량을 바꾼다',
    titleLatin: 'WHERE THE DISPENSER GOES',
    date: '2025-06-18',
    category: '운영',
    image: 'service-coating',
    readingTime: '5분',
    excerpt:
      '한 번에 여러 장이 딸려 나오는 대부분의 원인은 제품이 아니라 설치 높이와 방향입니다.',
    body: [
      { h: '제품 탓이 아닐 때가 많다', p: ['핸드타월이 한 번에 서너 장씩 뽑힌다는 이야기를 자주 듣습니다. 제품을 바꿔도 그대로인 경우가 많은데, 원인이 설치 쪽에 있기 때문입니다.'] },
      { h: '높이와 각도', p: ['디스펜서가 너무 높으면 아래로 당기는 힘이 세져 여러 장이 함께 나옵니다. 세면대 상단에서 눈높이 아래 정도가 무난합니다.', '벽에서 떨어져 있거나 기울어 있으면 다음 장이 제자리를 못 잡아 같은 문제가 생깁니다.'] },
      { h: '접이 방식이 맞아야 한다', p: ['디스펜서마다 전제하는 접이 방식이 정해져 있습니다. 그 방식이 아닌 제품을 넣으면 한 장씩 인출되는 구조 자체가 작동하지 않습니다.', '쓰시는 디스펜서 사진 한 장만 보내 주셔도 맞는 제품을 골라 드릴 수 있습니다.'] },
      { h: '점보롤도 마찬가지', p: ['점보롤 디스펜서는 내경과 폭이 맞아야 합니다. 억지로 끼우면 회전이 뻑뻑해지고, 결국 이용자가 손으로 뜯어 쓰게 되어 소비량이 늘어납니다.'] },
    ],
  },
  {
    slug: 'pulp-and-recycled',
    title: '천연펄프와 재생지, 무엇을 기준으로 고를까',
    titleLatin: 'PULP OR RECYCLED',
    date: '2025-03-11',
    category: '소재',
    image: 'service-inspection',
    readingTime: '8분',
    excerpt:
      '피부에 닿는가, 소모가 많은가. 두 질문만으로 대부분의 선택이 정리됩니다.',
    body: [
      { h: '두 가지를 다 쓴다', p: ['천연펄프와 재생지는 우열의 문제가 아니라 용도의 문제입니다. 한 건물 안에서도 자리에 따라 다른 쪽이 맞습니다.', '전부 천연펄프로 통일하면 비용이 부담되고, 전부 재생지로 통일하면 만족도가 떨어지는 자리가 반드시 생깁니다.'] },
      { h: '피부에 닿는 자리', p: ['미용티슈, 두루마리티슈, 냅킨처럼 피부에 직접 닿는 품목은 형광증백제를 쓰지 않은 천연펄프를 권해 드립니다. 촉감 차이가 분명하고, 민감한 분들이 쓰는 시설이라면 더 그렇습니다.', '요양시설이나 어린이집처럼 사용자가 정해져 있는 곳에서는 이 부분을 먼저 정하고 나머지를 맞추는 편이 낫습니다.'] },
      { h: '소모가 많은 자리', p: ['공용 화장실 점보롤, 청소용 타월처럼 소모가 많고 촉감이 크게 중요하지 않은 자리는 재생지 제품이 비용 면에서 유리합니다.', '요즘 재생지는 품질 편차가 예전보다 훨씬 작아져서, 이 자리에 굳이 천연펄프를 쓸 이유를 찾기 어렵습니다.'] },
      { h: '하얀 정도로 판단하지 않기', p: ['재생지가 더 어둡고 천연펄프가 더 하얗다는 인식이 있지만, 표백 정도에 따라 얼마든지 달라집니다. 색만 보고 원지를 판단하기는 어렵습니다.', '무형광 여부가 필요하시면 그 항목을 명시해 주세요. 해당 조건에 맞는 제품으로 안내드립니다.'] },
    ],
  },
  {
    slug: 'storage-humidity',
    title: '대량으로 들여놓은 휴지가 눅눅해지는 이유',
    titleLatin: 'KEEPING PAPER DRY',
    date: '2024-12-09',
    category: '보관',
    image: 'service-reliability',
    readingTime: '6분',
    excerpt:
      '싸게 많이 사두고 절반을 버리는 일이 실제로 생깁니다. 보관 조건 몇 가지면 대부분 막을 수 있습니다.',
    body: [
      { h: '한 번에 많이 사는 것의 함정', p: ['단가를 낮추려고 반년 치를 한 번에 들여놓는 곳이 있습니다. 보관 조건이 나쁘면 뒤쪽 물량은 눅눅해지거나 냄새가 배어 결국 쓰지 못하게 됩니다.', '아낀 단가보다 버린 양이 큰 경우를 드물지 않게 봅니다.'] },
      { h: '바닥과 벽에서 띄운다', p: ['종이는 콘크리트 바닥과 외벽에서 올라오는 습기를 그대로 빨아들입니다. 팔레트나 선반으로 바닥에서 띄우고, 벽에서 한 뼘 정도 떨어뜨려 두는 것만으로 상당 부분 해결됩니다.'] },
      { h: '비닐 포장은 뜯지 않는다', p: ['낱개 비닐 포장은 운반용이 아니라 보호용입니다. 쓸 만큼만 뜯고 나머지는 그대로 두는 편이 좋습니다.', '박스도 마찬가지로, 미리 개봉해 쌓아 두면 먼지와 습기가 함께 들어갑니다.'] },
      { h: '차라리 주기를 나눈다', p: ['보관 공간이 마땅치 않다면 한 번에 많이 받기보다 주기를 나누는 편이 낫습니다. 정기 배송으로 돌리면 단가는 유지하면서 재고 부담만 덜어낼 수 있습니다.', '실제 사용량이 잡히면 주기와 수량을 함께 잡아 드립니다.'] },
    ],
  },
]

/* ── legal ──────────────────────────────────────────────────────────────────── */

export const legal = [
  {
    slug: 'terms-conditions',
    title: '이용약관',
    titleLatin: 'TERMS & CONDITIONS',
    updated: '2026-01-01',
    sections: [
      {
        h: '1. 목적',
        p: [
          '본 약관은 주식회사 이루모스(이하 "회사")가 운영하는 웹사이트 및 관련 서비스의 이용 조건과 절차를 규정합니다.',
          '이용자는 본 사이트에 접속함으로써 본 약관에 동의한 것으로 봅니다.',
        ],
      },
      {
        h: '2. 거래 조건의 우선순위',
        p: [
          '본 사이트에 게시된 제품 정보와 규격은 일반적인 안내를 목적으로 하며 청약의 유인에 해당합니다.',
          '실제 공급의 범위, 규격, 수량, 납기, 대금, 품질보증은 회사와 고객이 확정한 견적서 및 발주서에 따르며, 해당 문서가 본 약관과 사이트 게시 내용에 우선합니다.',
        ],
      },
      {
        h: '3. 제품 규격에 관한 고지',
        p: [
          '본 사이트에 게시된 겹수, 권취 길이, 박스 입수, 납품 주기 등의 수치는 일반적인 규격을 기준으로 한 참고값입니다.',
          '실제 공급 사양은 품목과 생산 로트에 따라 달라질 수 있으며, 견적 단계에서 확정됩니다.',
        ],
      },
      {
        h: '4. 주문제작 및 인쇄물',
        p: [
          '로고 인쇄 등 주문제작 건은 고객이 제공한 도안과 회사가 제시한 시안을 고객이 확인·승인한 뒤 생산에 착수합니다.',
          '승인 이후 발생한 도안 변경 또는 주문 취소에 대해서는 진행 단계에 따른 비용이 발생할 수 있습니다.',
          '고객이 제공한 상표·도안의 사용 권한은 고객에게 있으며, 이로 인한 제3자와의 분쟁에 대해 회사는 책임을 지지 않습니다.',
        ],
      },
      {
        h: '5. 지식재산권',
        p: [
          '본 사이트의 텍스트, 이미지, 제품 사진 등 모든 콘텐츠의 권리는 회사 또는 정당한 권리자에게 있습니다.',
          '사전 서면 동의 없이 복제, 배포, 2차적 저작물 작성을 할 수 없습니다.',
        ],
      },
      {
        h: '6. 준거법 및 관할',
        p: ['본 약관은 대한민국 법령에 따라 해석되며, 분쟁은 대전지방법원 논산지원을 제1심 관할 법원으로 합니다.'],
      },
    ],
  },
  {
    slug: 'privacy-notice',
    title: '개인정보 처리방침',
    titleLatin: 'PRIVACY NOTICE',
    updated: '2026-01-01',
    sections: [
      {
        h: '1. 수집하는 개인정보',
        p: [
          '견적 요청 및 문의 시 회사명(상호), 담당자 성명, 연락처, 이메일, 문의 내용을 수집합니다.',
          '배송이 필요한 경우 수령지 주소와 수령인 연락처를 추가로 수집합니다.',
          '사이트 이용 과정에서 접속 로그, 브라우저 정보가 자동으로 기록될 수 있습니다.',
        ],
      },
      {
        h: '2. 이용 목적',
        p: [
          '수집한 정보는 문의 응대, 견적 제공, 계약 체결 및 이행, 제품 배송, 사후 관리 목적으로만 이용합니다.',
          '별도 동의를 받은 경우에 한해 신제품 소식과 할인 안내를 발송합니다.',
        ],
      },
      {
        h: '3. 보유 및 파기',
        p: [
          '견적·상담 정보는 목적 달성 후 3년간 보관 후 파기합니다.',
          '거래 관련 정보는 전자상거래법, 국세기본법 등 관계 법령이 정한 기간 동안 보관합니다.',
        ],
      },
      {
        h: '4. 제3자 제공 및 처리 위탁',
        p: [
          '법령에 근거한 경우를 제외하고 개인정보를 제3자에게 제공하지 않습니다.',
          '배송, 메일 발송, 호스팅 등 일부 업무는 외부 사업자에게 위탁될 수 있으며, 위탁 시 안전한 처리를 계약으로 보장합니다.',
        ],
      },
      {
        h: '5. 이용자의 권리',
        p: [
          '이용자는 언제든지 자신의 개인정보에 대한 열람, 정정, 삭제, 처리 정지를 요구할 수 있습니다.',
          `요청은 ${site.email} 또는 ${site.phone} 으로 접수합니다.`,
        ],
      },
    ],
  },
]

/* ── shop / 온라인 주문 ──────────────────────────────────────────────────────
   자사몰 성격의 제품 구매 화면입니다. 카탈로그(20종) · 장바구니 · 주문서까지
   실제로 동작하지만 결제와 재고 연동은 **없습니다** — 정적 사이트라 서버가 없기
   때문입니다. 주문서는 문의 폼과 같은 방식으로 메일 초안을 만들어 넘깁니다
   (ui.js > initContactForm 과 같은 원칙: 받지 않은 주문을 받았다고 말하지 않습니다).

   ⚠ 아래 **가격 · 규격 · 배송 조건은 전부 시안용 예시**입니다. 원본 사이트에 판매가
   표기가 없어 실제 값을 알 수 없고, 위생용지는 수량과 납품 조건에 따라 단가가
   달라집니다. 화면에도 그렇게 밝혀 두었습니다. See MISSING.md > D7. */

export const shopMeta = {
  eyebrow: '(Online store)',
  heading: 'ORDER\nONLINE',
  lead: '자주 나가는 규격을 골라 바로 담으실 수 있게 정리했습니다. 수량이 많거나 규격이 다르면 견적으로 진행하시는 편이 더 저렴합니다.',
  notice:
    '표시된 가격 · 규격 · 배송 조건은 화면 구성을 보여 드리기 위한 예시입니다. 실제 판매가는 수량과 납품 조건에 따라 달라지며, 주문서를 보내 주시면 확정 견적으로 회신드립니다.',
  /* 장바구니 합계에 쓰입니다. 이 값들도 예시입니다. */
  shipping: { flat: 3000, freeOver: 100000 },
}

export const shopCategories = [
  { slug: 'kitchen-towel', label: '키친타월', labelEn: 'KITCHEN TOWEL' },
  { slug: 'jumbo-roll', label: '점보롤티슈', labelEn: 'JUMBO ROLL' },
  { slug: 'toilet-tissue', label: '두루마리티슈', labelEn: 'TOILET TISSUE' },
  { slug: 'facial-napkin', label: '미용티슈 · 냅킨', labelEn: 'FACIAL & NAPKIN' },
  { slug: 'hand-towel', label: '핸드타월', labelEn: 'HAND TOWEL' },
  { slug: 'baking-paper', label: '종이호일', labelEn: 'BAKING PAPER' },
]

/* 20종. `category` 는 shopCategories 의 slug 이면서 services 의 slug 이기도 해서,
   상세 화면에서 해당 품목 소개 페이지로 그대로 이어집니다.

   `externalUrl` (선택) — 카페24 등 호스팅형 커머스에 상품을 등록한 뒤 회수한
   그 상품의 주소. 값이 있으면 카드와 상세의 「담기」가 「구매하기」 외부 링크로
   바뀌고 `Product` JSON-LD 의 offers.url 도 그쪽을 가리킵니다. 없으면 기존
   장바구니 → 주문서(견적) 흐름 그대로입니다. 품목마다 따로 켤 수 있어 소량은
   외부 결제, 대량은 견적으로 갈라 받는 병행 운영이 됩니다.
   ⚠ 주소 형태는 스킨·설정마다 달라 실제 등록된 상품을 열어 확인한 뒤 넣으십시오.
   See docs/cafe24-integration-guidebook.html > 04. */
export const products = [
  {
    slug: 'kitchen-towel-home-150',
    sku: 'ILM-KT-150H',
    name: '가정용 키친타월 2겹 150매',
    nameLatin: 'HOME KITCHEN TOWEL 2PLY',
    category: 'kitchen-towel',
    price: 14900,
    unit: '팩',
    pack: '6롤 / 팩',
    badge: 'BEST',
    image: 'service-parts',
    summary: '엠보싱 2겹으로 기름과 물을 한 장에 걷어내는 가정용 표준 규격입니다.',
    specs: [
      { k: '겹수', v: '2겹 엠보싱' },
      { k: '매수', v: '롤당 150매' },
      { k: '구성', v: '6롤 / 팩' },
      { k: '원지', v: '무형광 천연펄프' },
      { k: '용도', v: '가정 주방 · 소형 매장' },
    ],
    tags: ['가정용', '무형광', '2겹'],
  },
  {
    slug: 'kitchen-towel-pro-300',
    sku: 'ILM-KT-300P',
    name: '업소용 대용량 키친타월 300m',
    nameLatin: 'PRO KITCHEN TOWEL 300M',
    category: 'kitchen-towel',
    price: 28000,
    unit: '박스',
    pack: '2롤 / 박스',
    badge: '',
    image: 'service-reliability',
    summary: '롤 하나가 오래 가는 업소용 대용량 규격입니다. 교체 횟수를 줄여 줍니다.',
    specs: [
      { k: '겹수', v: '2겹 엠보싱' },
      { k: '길이', v: '롤당 300m' },
      { k: '구성', v: '2롤 / 박스' },
      { k: '습윤 강도', v: '보강 처리' },
      { k: '용도', v: '외식업장 주방 · 급식소' },
    ],
    tags: ['업소용', '대용량', '2겹'],
  },
  {
    slug: 'kitchen-towel-3ply',
    sku: 'ILM-KT-120T',
    name: '3겹 프리미엄 키친타월 120매',
    nameLatin: 'PREMIUM KITCHEN TOWEL 3PLY',
    category: 'kitchen-towel',
    price: 18500,
    unit: '팩',
    pack: '4롤 / 팩',
    badge: 'NEW',
    image: 'circuit-macro',
    summary: '깊은 엠보싱 3겹입니다. 튀김 기름처럼 흡수량이 많은 자리에 씁니다.',
    specs: [
      { k: '겹수', v: '3겹 딥 엠보싱' },
      { k: '매수', v: '롤당 120매' },
      { k: '구성', v: '4롤 / 팩' },
      { k: '원지', v: '무형광 천연펄프' },
      { k: '용도', v: '튀김 · 조리대 물기 제거' },
    ],
    tags: ['3겹', '고흡수', '무형광'],
  },
  {
    slug: 'kitchen-towel-box-12',
    sku: 'ILM-KT-B12',
    name: '키친타월 12롤 박스',
    nameLatin: 'KITCHEN TOWEL CASE 12',
    category: 'kitchen-towel',
    price: 46000,
    unit: '박스',
    pack: '12롤 / 박스',
    badge: '',
    image: 'server-rack',
    summary: '사용량이 꾸준한 곳을 위한 박스 단위 구성입니다. 정기 배송으로 전환할 수 있습니다.',
    specs: [
      { k: '겹수', v: '2겹 엠보싱' },
      { k: '매수', v: '롤당 150매' },
      { k: '구성', v: '12롤 / 박스' },
      { k: '보관', v: '박스 적재 · 직사광선 피할 것' },
      { k: '용도', v: '급식소 · 프랜차이즈 매장' },
    ],
    tags: ['박스 단위', '정기 배송 가능'],
  },
  {
    slug: 'jumbo-roll-300',
    sku: 'ILM-JR-300',
    name: '업소용 점보롤 2겹 300m',
    nameLatin: 'JUMBO ROLL 2PLY 300M',
    category: 'jumbo-roll',
    price: 42000,
    unit: '박스',
    pack: '12롤 / 박스',
    badge: 'BEST',
    image: 'service-clean',
    summary: '가장 많이 나가는 표준 점보롤입니다. 일반 디스펜서 내경에 맞춥니다.',
    specs: [
      { k: '겹수', v: '2겹 엠보싱' },
      { k: '길이', v: '롤당 300m' },
      { k: '구성', v: '12롤 / 박스' },
      { k: '코어', v: '유심지 · 표준 내경' },
      { k: '용도', v: '다중이용시설 화장실' },
    ],
    tags: ['업소용', '디스펜서 호환', '2겹'],
  },
  {
    slug: 'jumbo-roll-coreless',
    sku: 'ILM-JR-250C',
    name: '무심지 점보롤 2겹 250m',
    nameLatin: 'CORELESS JUMBO ROLL 250M',
    category: 'jumbo-roll',
    price: 45000,
    unit: '박스',
    pack: '12롤 / 박스',
    badge: '',
    image: 'cleanroom-bay',
    summary: '심지가 없어 마지막 한 장까지 쓰고 폐기물도 줄어듭니다. 전용 디스펜서가 필요합니다.',
    specs: [
      { k: '겹수', v: '2겹 엠보싱' },
      { k: '길이', v: '롤당 250m' },
      { k: '구성', v: '12롤 / 박스' },
      { k: '코어', v: '무심지 (전용 디스펜서)' },
      { k: '용도', v: '신규 설치 현장' },
    ],
    tags: ['무심지', '폐기물 절감'],
  },
  {
    slug: 'jumbo-roll-eco',
    sku: 'ILM-JR-400E',
    name: '재생 점보롤 1겹 400m',
    nameLatin: 'RECYCLED JUMBO ROLL 400M',
    category: 'jumbo-roll',
    price: 33000,
    unit: '박스',
    pack: '12롤 / 박스',
    badge: '',
    image: 'chip-die',
    summary: '사용량이 아주 많은 곳을 위한 재생지 규격입니다. 길이가 길어 교체 주기가 깁니다.',
    specs: [
      { k: '겹수', v: '1겹' },
      { k: '길이', v: '롤당 400m' },
      { k: '구성', v: '12롤 / 박스' },
      { k: '원지', v: '재생지' },
      { k: '용도', v: '고사용 공공시설 · 공장' },
    ],
    tags: ['재생지', '고사용처', '경제형'],
  },
  {
    slug: 'jumbo-roll-premium',
    sku: 'ILM-JR-200T',
    name: '점보롤 3겹 엠보싱 200m',
    nameLatin: 'PREMIUM JUMBO ROLL 3PLY',
    category: 'jumbo-roll',
    price: 52000,
    unit: '박스',
    pack: '12롤 / 박스',
    badge: 'NEW',
    image: 'hero-poster',
    summary: '점보롤의 편의성에 3겹의 촉감을 더한 구성입니다. 객실과 라운지 화장실에 씁니다.',
    specs: [
      { k: '겹수', v: '3겹 엠보싱' },
      { k: '길이', v: '롤당 200m' },
      { k: '구성', v: '12롤 / 박스' },
      { k: '원지', v: '무형광 천연펄프' },
      { k: '용도', v: '호텔 · 라운지 · 프리미엄 매장' },
    ],
    tags: ['3겹', '무형광', '프리미엄'],
  },
  {
    slug: 'toilet-tissue-home-30',
    sku: 'ILM-TT-30H',
    name: '가정용 두루마리 2겹 30m',
    nameLatin: 'HOME TOILET TISSUE 30M',
    category: 'toilet-tissue',
    price: 21000,
    unit: '박스',
    pack: '30롤 / 박스',
    badge: 'BEST',
    image: 'service-ceramic',
    summary: '일반 가정용 홀더에 맞는 표준 지름입니다. 개별 비닐 포장이라 보관이 깔끔합니다.',
    specs: [
      { k: '겹수', v: '2겹 엠보싱' },
      { k: '길이', v: '롤당 30m' },
      { k: '구성', v: '30롤 / 박스' },
      { k: '포장', v: '개별 비닐 포장' },
      { k: '용도', v: '가정 · 소형 사무실' },
    ],
    tags: ['가정용', '무형광', '개별 포장'],
  },
  {
    slug: 'toilet-tissue-pro-50',
    sku: 'ILM-TT-50P',
    name: '업소용 두루마리 2겹 50m',
    nameLatin: 'PRO TOILET TISSUE 50M',
    category: 'toilet-tissue',
    price: 27500,
    unit: '박스',
    pack: '24롤 / 박스',
    badge: '',
    image: 'chip-array',
    summary: '가정용보다 권취 길이가 길어 교체 부담이 줄어듭니다. 업소용 홀더 규격입니다.',
    specs: [
      { k: '겹수', v: '2겹 엠보싱' },
      { k: '길이', v: '롤당 50m' },
      { k: '구성', v: '24롤 / 박스' },
      { k: '해리성', v: '수분 해리 관리' },
      { k: '용도', v: '식당 · 사무실 · 학원' },
    ],
    tags: ['업소용', '2겹'],
  },
  {
    slug: 'toilet-tissue-3ply',
    sku: 'ILM-TT-25T',
    name: '3겹 천연펄프 두루마리 25m',
    nameLatin: 'NATURAL PULP TISSUE 3PLY',
    category: 'toilet-tissue',
    price: 34000,
    unit: '박스',
    pack: '30롤 / 박스',
    badge: '',
    image: 'silicon-wafer',
    summary: '피부에 닿는 느낌을 우선한 3겹 천연펄프 구성입니다.',
    specs: [
      { k: '겹수', v: '3겹 엠보싱' },
      { k: '길이', v: '롤당 25m' },
      { k: '구성', v: '30롤 / 박스' },
      { k: '원지', v: '무형광 천연펄프 100%' },
      { k: '용도', v: '호텔 객실 · 가정' },
    ],
    tags: ['3겹', '천연펄프', '무형광'],
  },
  {
    slug: 'toilet-tissue-eco',
    sku: 'ILM-TT-50E',
    name: '재생 두루마리 1겹 50m',
    nameLatin: 'RECYCLED TOILET TISSUE 50M',
    category: 'toilet-tissue',
    price: 24000,
    unit: '박스',
    pack: '40롤 / 박스',
    badge: '',
    image: 'silicon-ingot',
    summary: '사용량 대비 단가를 낮춘 재생지 구성입니다. 공용 화장실에 적합합니다.',
    specs: [
      { k: '겹수', v: '1겹' },
      { k: '길이', v: '롤당 50m' },
      { k: '구성', v: '40롤 / 박스' },
      { k: '원지', v: '재생지' },
      { k: '용도', v: '공용 화장실 · 현장 사무소' },
    ],
    tags: ['재생지', '경제형', '대량'],
  },
  {
    slug: 'facial-tissue-200',
    sku: 'ILM-FT-200',
    name: '각티슈 2겹 200매',
    nameLatin: 'FACIAL TISSUE BOX 200',
    category: 'facial-napkin',
    price: 26000,
    unit: '박스',
    pack: '20갑 / 박스',
    badge: 'BEST',
    image: 'service-inspection',
    summary: '카운터와 테이블에 두는 표준 각티슈입니다. 다음 장이 따라 올라오는 팝업 구조입니다.',
    specs: [
      { k: '겹수', v: '2겹' },
      { k: '매수', v: '갑당 200매' },
      { k: '구성', v: '20갑 / 박스' },
      { k: '인출', v: '팝업 적층' },
      { k: '용도', v: '카운터 · 객실 · 사무실' },
    ],
    tags: ['각티슈', '팝업', '2겹'],
  },
  {
    slug: 'napkin-table-6fold',
    sku: 'ILM-NP-6F',
    name: '식탁 냅킨 6절 5,000매',
    nameLatin: 'TABLE NAPKIN 6-FOLD',
    category: 'facial-napkin',
    price: 31000,
    unit: '박스',
    pack: '5,000매 / 박스',
    badge: '',
    image: 'nano-texture',
    summary: '식탁 디스펜서에 들어가는 6절 냅킨입니다. 색상을 고르실 수 있습니다.',
    specs: [
      { k: '접이', v: '6절 접이' },
      { k: '매수', v: '5,000매 / 박스' },
      { k: '규격', v: '펼침 기준 약 230 × 230mm' },
      { k: '색상', v: '백색 기본 · 컬러 선택 가능' },
      { k: '용도', v: '식당 테이블 · 푸드코트' },
    ],
    tags: ['냅킨', '컬러 선택', '식탁용'],
  },
  {
    slug: 'napkin-printed',
    sku: 'ILM-NP-OEM',
    name: '로고 인쇄 냅킨 (주문제작)',
    nameLatin: 'PRINTED NAPKIN — MADE TO ORDER',
    category: 'facial-napkin',
    price: 89000,
    unit: '박스',
    pack: '10,000매 / 박스',
    badge: 'MADE TO ORDER',
    image: 'service-packaging',
    summary: '보유하신 로고를 인쇄해 만듭니다. 시안을 먼저 확인하신 뒤 생산에 들어갑니다.',
    specs: [
      { k: '인쇄', v: '1~4도' },
      { k: '매수', v: '10,000매 / 박스' },
      { k: '최소 수량', v: '1박스부터 (도수에 따라 상이)' },
      { k: '진행', v: '로고 파일 접수 → 시안 확인 → 생산' },
      { k: '납기', v: '시안 확정 후 별도 안내' },
    ],
    tags: ['주문제작', '로고 인쇄', 'OEM'],
  },
  {
    slug: 'napkin-kraft',
    sku: 'ILM-NP-KR',
    name: '무표백 크라프트 냅킨 3,000매',
    nameLatin: 'UNBLEACHED KRAFT NAPKIN',
    category: 'facial-napkin',
    price: 24500,
    unit: '박스',
    pack: '3,000매 / 박스',
    badge: 'NEW',
    image: 'fab-equipment',
    summary: '표백하지 않은 크라프트 원지 냅킨입니다. 매장 톤을 살리는 자리에 씁니다.',
    specs: [
      { k: '원지', v: '무표백 크라프트' },
      { k: '겹수', v: '1겹' },
      { k: '매수', v: '3,000매 / 박스' },
      { k: '규격', v: '펼침 기준 약 300 × 300mm' },
      { k: '용도', v: '카페 · 베이커리 · 브런치' },
    ],
    tags: ['무표백', '크라프트', '카페'],
  },
  {
    slug: 'hand-towel-multifold',
    sku: 'ILM-HT-MF',
    name: '멀티폴드 핸드타월 200매',
    nameLatin: 'MULTIFOLD HAND TOWEL',
    category: 'hand-towel',
    price: 38000,
    unit: '박스',
    pack: '20속 / 박스',
    badge: 'BEST',
    image: 'service-coating',
    summary: '한 번에 한 장만 나오는 접이 구조입니다. 디스펜서 접이 방식을 확인하고 보내 드립니다.',
    specs: [
      { k: '접이', v: '멀티폴드' },
      { k: '매수', v: '속당 200매' },
      { k: '구성', v: '20속 / 박스' },
      { k: '습윤 강도', v: '보강 처리' },
      { k: '용도', v: '화장실 · 주방 세면대' },
    ],
    tags: ['핸드타월', '한 장 인출', '디스펜서 호환'],
  },
  {
    slug: 'hand-towel-dispenser',
    sku: 'ILM-HT-DSP',
    name: '핸드타월 디스펜서 (벽걸이)',
    nameLatin: 'HAND TOWEL DISPENSER',
    category: 'hand-towel',
    price: 19000,
    unit: '개',
    pack: '1개',
    badge: '',
    image: 'probe-station',
    summary: '멀티폴드 규격용 벽걸이 디스펜서입니다. 나사 고정과 양면테이프 설치 모두 가능합니다.',
    specs: [
      { k: '재질', v: 'ABS' },
      { k: '호환', v: '멀티폴드 핸드타월' },
      { k: '설치', v: '벽면 고정 (나사 · 양면테이프)' },
      { k: '수납', v: '약 400매' },
      { k: '용도', v: '화장실 · 주방 세면대' },
    ],
    tags: ['디스펜서', '벽걸이', '비품'],
  },
  {
    slug: 'baking-paper-roll',
    sku: 'ILM-BP-R50',
    name: '종이호일 롤 30cm × 50m',
    nameLatin: 'BAKING PAPER ROLL 50M',
    category: 'baking-paper',
    price: 12000,
    unit: '롤',
    pack: '1롤',
    badge: '',
    image: 'service-vacuum',
    summary: '양면 실리콘 처리 롤형입니다. 필요한 길이만 잘라 쓰는 주방용 기본 규격입니다.',
    specs: [
      { k: '규격', v: '폭 30cm × 길이 50m' },
      { k: '코팅', v: '양면 실리콘' },
      { k: '내열', v: '조리 온도 확인 후 안내' },
      { k: '적합성', v: '식품 접촉 적합' },
      { k: '용도', v: '오븐 · 팬 · 식품 포장' },
    ],
    tags: ['롤형', '양면 실리콘'],
  },
  {
    slug: 'baking-paper-sheet',
    sku: 'ILM-BP-S46',
    name: '종이호일 재단 시트 40 × 60cm',
    nameLatin: 'BAKING PAPER SHEET 40x60',
    category: 'baking-paper',
    price: 23000,
    unit: '박스',
    pack: '500매 / 박스',
    badge: '',
    image: 'lab-optics',
    summary: '오븐팬 규격에 맞춰 미리 재단된 시트형입니다. 매번 자르는 시간을 아낍니다.',
    specs: [
      { k: '규격', v: '400 × 600mm' },
      { k: '매수', v: '500매 / 박스' },
      { k: '코팅', v: '양면 실리콘' },
      { k: '적합성', v: '식품 접촉 적합' },
      { k: '용도', v: '베이커리 · 단체 급식 오븐' },
    ],
    tags: ['시트형', '오븐팬 규격', '베이커리'],
  },
]

/* ── helpers ────────────────────────────────────────────────────────────────── */

export const serviceBySlug = (slug) => services.find((s) => s.slug === slug)
export const workBySlug = (slug) => portfolio.find((p) => p.slug === slug)
export const productBySlug = (slug) => products.find((p) => p.slug === slug)
export const productsIn = (category) => products.filter((p) => p.category === category)
export const categoryBySlug = (slug) => shopCategories.find((c) => c.slug === slug)
