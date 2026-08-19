/**
 * fetch-images.mjs — source, convert and store every image locally.
 *
 * TWO SOURCES, one output convention:
 *
 *   1. `from:` — a photo from `original/images/`, i.e. (주)이루모스 자신의 제품
 *      사진. 실제 취급 품목이므로 제품·서비스 자리에는 이쪽을 우선합니다.
 *      원본이 600px 급이라 확대 대신 **캔버스를 넓혀서**(fit: contain, 가장자리
 *      색 샘플링) 목표 비율을 맞춥니다. 흰 배경 제품컷이라 이어붙인 자리가
 *      보이지 않고, 해상도 손실도 없습니다.
 *
 *   2. `queries:` — 스톡 검색 (Pixabay → Pexels, 상업적 이용 무료).
 *      히어로·분위기 컷처럼 원본에 쓸 만한 사진이 없는 자리만 채웁니다.
 *      히어로는 2560px 이상만 통과시킵니다(`minReal`).
 *
 * 규칙:
 *   · 영문 검색어, 주제에 맞춰 작성
 *   · search -> download -> WebP -> assets/images -> 내부 경로로 참조
 *   · 같은 사진을 두 자리에 쓰지 않음 (provider id AND url 로 전역 중복 제거)
 *   · 다운로드가 진짜 실패한 자리에만 플레이스홀더를 씁니다
 *
 * assets/images/manifest.json 을 씁니다: { slot: {src,width,height,alt,credit} }.
 * 템플릿이 이 파일을 읽어 모든 <img> 에 intrinsic size 와 alt 를 붙입니다.
 */

import { mkdir, writeFile, readFile } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import sharp from 'sharp'

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..')
const OUT = join(ROOT, 'assets', 'images')

const PIXABAY_KEY = process.env.PIXABAY_KEY
const PEXELS_KEY = process.env.PEXELS_KEY

/* ── slot definitions ───────────────────────────────────────────────────── */

/**
 * THEME: 위생용지 — 키친타월 · 점보롤티슈 · 두루마리티슈.
 *
 * Art direction: 제품 그 자체와 그것이 놓이는 공간. 흰 종이, 롤, 엠보싱 표면,
 * 주방·화장실·카페·사무실 같은 사용 현장, 그리고 원지와 물류. 연출된 모델컷과
 * 일러스트·3D 렌더는 배제합니다.
 *
 * ⚠ 스톡 API 는 태그 순위로 정렬하므로 가끔 어긋난 컷이 섞일 수 있습니다.
 * assets/images/CREDITS.md 를 확인한 뒤 해당 슬롯의 `queries` 를 고치거나
 * `FORCE=1 npm run images` 로 다시 뽑으세요. See MISSING.md > D4.
 */

/* Anchor vocabulary — 제공자 태그에 이 중 하나는 있어야 통과합니다. */
const PAPER = ['paper', 'tissue', 'towel', 'napkin', 'roll', 'hygiene', 'toilet', 'kitchen', 'pulp', 'sanitary', 'wipe', 'cellulose', 'white', 'soft', 'clean']
const PLACE = ['restaurant', 'kitchen', 'cafe', 'coffee', 'office', 'hotel', 'room', 'interior', 'restroom', 'bathroom', 'washroom', 'hospital', 'clinic', 'building', 'table', 'counter', 'workplace', 'desk']
const MILL = ['factory', 'industry', 'industrial', 'production', 'manufacturing', 'machine', 'warehouse', 'logistics', 'pallet', 'box', 'carton', 'storage', 'mill', 'paper']

/* 사진이 아닌 것, 그리고 이 업종에서 쓸 수 없는 것. 위생용지는 검색어가
   생활 밀착형이라 일러스트·3D 렌더·의료 폐기물 계열이 잘 섞여 들어옵니다.
   슬롯이 `ban` 을 직접 선언하면 이 목록 대신 그것이 쓰입니다. */
const BAN = [
  'illustration', 'drawing', 'cartoon', 'vector', 'clipart', 'render', '3d',
  'sketch', 'painting', 'icon', 'logo', 'text', 'origami', 'wallpaper',
  'newspaper', 'money', 'banknote', 'document', 'book', 'autumn', 'leaf',
  'blood', 'wound', 'syringe', 'garbage', 'trash', 'waste',
]

const SLOTS = [
  /* ── 히어로 / 소셜 — 첫인상용이라 반드시 스톡 고해상도 ─────────────────── */
  {
    name: 'hero-poster',
    w: 2560,
    h: 1440,
    must: PAPER,
    minReal: 2400, // 실제 내려받은 파일 기준. 업스케일을 막습니다.
    preferPexels: true, // Pixabay 무료 API 는 1920px 로 잘려 나옵니다.
    queries: ['paper towel roll kitchen', 'toilet paper rolls white', 'white tissue paper soft', 'paper roll stack white'],
    alt: '흰 위생용지 롤',
  },
  {
    name: 'og-cover',
    w: 1200,
    h: 630,
    must: PAPER,
    queries: ['kitchen paper towel roll', 'toilet paper roll white background', 'tissue roll closeup'],
    alt: '(주)이루모스 대표 이미지 — 위생용지',
  },

  /* 히어로 배경 — 별필드 뒤에 깔리는 울창한 숲.
     종이의 출발점이 나무라는 점에서 업종과도 맞고, 숲 위로 별이 뜨는 구도라
     기존 별 효과를 그대로 살릴 수 있습니다.

     ⚠ 톤 창 25~85 는 두 번 실패한 끝에 나온 값입니다.

     1차(상한 110): luma 75 짜리 안개 숲. 그 위에 별 100만 개가 **가산 합성**으로
        얹히면서 사진이 회색 안개로 뭉개졌습니다. → 너무 밝고 대비가 없었습니다.
     2차(상한 60): 밤하늘 실루엣. 별은 완벽히 살았지만 **나무가 순수한 검정**이라
        아무리 밝기·대비를 만져도 되살아나지 않았습니다. 실루엣 사진의 나무는
        luma 0.04 근처의 빈 픽셀이라 복원할 디테일 자체가 없습니다.

     그래서 이 자리는 "실루엣"도 "안개"도 아닌 **빛을 받은 울창한 숲**을 찾습니다.
     하한 25 는 나무에 실제 계조가 있을 것, 상한 85 는 페이지가 어두울 것. */
  {
    name: 'hero-forest',
    w: 2560,
    h: 1440,
    must: ['forest', 'tree', 'trees', 'woods', 'pine', 'fir', 'spruce', 'woodland', 'canopy'],
    /* 전역 BAN 의 'autumn'·'leaf' 는 숲 사진 대부분을 걷어내므로 여기서는 뺍니다.
       대신 초원·인물·도시 쪽을 막습니다. */
    ban: ['illustration', 'drawing', 'cartoon', 'vector', 'render', '3d', 'painting', 'sketch',
      'wallpaper', 'meadow', 'field', 'flower', 'beach', 'desert', 'city', 'road', 'people', 'woman', 'man', 'snow'],
    tone: { min: 25, max: 85 },
    minReal: 2400,
    preferPexels: true, // 2560px 이상은 Pexels 원본에서만 나옵니다
    queries: ['dense green forest tall trees', 'dark green forest dense canopy', 'coniferous forest dense trunks'],
    alt: '빛을 받은 울창한 숲',
  },

  /* About 서브 히어로 — 밝은 대낮의 건물이 카피 뒤에 전면으로 깔립니다.
     사이트 전체의 어두운 톤 규칙을 일부러 깨는 자리라 톤 창을 따로 갖습니다. */
  {
    name: 'about-building',
    w: 1920,
    h: 1200,
    must: ['warehouse', 'building', 'architecture', 'sky', 'facade', 'office', 'hall'],
    /* 첫 패스에서 굴뚝이 늘어선 화학 플랜트가 걸렸습니다. 위생용지 회사의
       전경으로는 정반대의 인상이라 중공업 어휘를 통째로 막았습니다. */
    ban: ['vintage', 'antique', 'ruin', 'abandoned', 'night', 'graffiti', 'illustration', 'render',
      'chimney', 'smoke', 'smokestack', 'refinery', 'oil', 'chemical', 'power', 'pollution', 'steel', 'coal'],
    tone: { min: 95, max: 235 }, // 밝은 주광이 요점
    queries: ['modern warehouse building white exterior', 'modern industrial building facade daylight', 'contemporary office building sky daylight'],
    alt: '한낮의 생산 · 물류 시설 외관',
  },

  /* ── 사용처 4곳 — 실제 공간 사진이라 스톡 ─────────────────────────────── */
  {
    name: 'work-etch',
    w: 1600,
    h: 1200,
    must: PLACE,
    queries: ['commercial restaurant kitchen stainless', 'professional kitchen interior', 'restaurant kitchen counter'],
    alt: '외식업장 주방',
  },
  {
    name: 'work-cleaning',
    w: 1600,
    h: 1200,
    must: PLACE,
    queries: ['hospital corridor clean interior', 'nursing home hallway interior', 'clinic waiting room interior'],
    alt: '요양 · 의료시설 복도',
  },
  {
    name: 'work-foundry',
    w: 1600,
    h: 1200,
    must: ['office', 'interior', 'desk', 'workplace', 'lobby', 'hall', 'room', 'workspace'],
    ban: [...BAN, 'night', 'exterior', 'facade', 'street'], // 첫 패스에서 야경 외관이 걸렸습니다
    queries: ['open plan office interior desks', 'bright office workspace interior', 'office lobby interior daylight'],
    alt: '사무실 · 공공기관 실내',
  },
  {
    name: 'work-hbm',
    w: 1600,
    h: 1200,
    must: PLACE,
    queries: ['cafe interior table wooden', 'coffee shop interior counter', 'hotel room interior bed'],
    alt: '카페 · 숙박업소 실내',
  },

  /* ── 제품 8종 — 전부 (주)이루모스 원본 제품 사진 ───────────────────────── */
  { name: 'service-parts', w: 1200, h: 675, from: '03-content.jpg', fit: 'contain', alt: '키친타월 — 액체 흡수 시험' },
  { name: 'service-clean', w: 1200, h: 675, from: '04-content.jpg', fit: 'contain', alt: '점보롤티슈 — 대용량 롤' },
  { name: 'service-ceramic', w: 1200, h: 675, from: '06-content.jpg', fit: 'contain', alt: '두루마리티슈 — 엠보싱 롤' },
  { name: 'service-inspection', w: 1200, h: 675, from: '11-content.png', fit: 'cover', alt: '미용티슈 · 냅킨 — 적층 묶음' },
  { name: 'service-coating', w: 1200, h: 675, from: '02-content.jpg', fit: 'contain', alt: '핸드타월 — 접이식 적층' },
  { name: 'service-vacuum', w: 1200, h: 675, from: '08-content.png', fit: 'cover', alt: '주방용 종이호일' },
  /* 이 자리만 스톡입니다. 원본 사이트의 남은 사진(10-content, 밀랍 인장과 만년필)은
     문구류 연출컷이라 위생용지 주문제작을 설명하지 못해 쓰지 않았습니다. */
  {
    name: 'service-packaging',
    w: 1200,
    h: 675,
    must: [...MILL, 'print', 'printing', 'packaging', 'label', 'color', 'ink', 'paper'],
    /* 타사 상표가 크게 박힌 설비 클로즈업이 걸리지 않도록, 기계 자체보다
       "인쇄된 종이"와 잉크 쪽으로 검색어를 옮겼습니다. */
    ban: [...BAN, 'brand', 'sign', 'signage'],
    queries: ['printed paper color ink sheets', 'cardboard packaging boxes kraft', 'color ink printing paper stack'],
    alt: 'OEM · 주문제작 — 인쇄와 포장',
  },
  { name: 'service-reliability', w: 1200, h: 675, from: '05-content.jpg', fit: 'contain', alt: '대량 공급 — 적재된 롤' },

  /* ── 캐러셀 · 갤러리 ─────────────────────────────────────────────────── */

  // 원본 제품 사진에서
  { name: 'silicon-wafer', w: 900, h: 1200, from: '07-content.jpg', fit: 'contain', alt: '두루마리티슈 2롤' },
  { name: 'cleanroom-bay', w: 1200, h: 800, from: '12-content.png', fit: 'cover', alt: '점보롤 적재' },
  { name: 'chip-array', w: 1200, h: 800, from: '13-content.png', fit: 'cover', alt: '두루마리티슈 묶음' },
  { name: 'nano-texture', w: 1000, h: 1000, from: '09-content.png', fit: 'cover', alt: '컬러 인쇄 · 색상 옵션' },

  // 스톡으로 채우는 자리 — 원본에 대응 사진이 없는 주제들
  {
    /* PAPER 앵커가 넓어 'kitchen' 만 걸린 식기건조대 사진이 통과했었습니다.
       이 자리는 제품 자체가 보여야 하므로 종이 어휘로만 좁혔습니다. */
    name: 'chip-die',
    w: 1200,
    h: 1600,
    /* 'towel' 을 앵커로 두면 천 행주 사진이 통과합니다(실제로 걸렸습니다).
       태그에 'paper' 가 반드시 있어야 하고, 천·인물 어휘는 막습니다. */
    must: ['paper'],
    ban: [...BAN, 'cloth', 'fabric', 'textile', 'dish', 'woman', 'girl', 'man', 'people',
      'child', 'family', 'holder', 'stand', 'rack', 'toilet', 'bathroom', 'basket'],
    searchLandscape: true, // 세로 검색 풀이 너무 얕아 가로에서 찾아 크롭합니다
    minReal: 1400,
    queries: ['paper towel roll white kitchen', 'kitchen paper roll towel', 'rolled kitchen paper'],
    // 최종 채택본은 인쇄 전 무지 롤입니다. 캡션도 그에 맞춰 두었습니다.
    alt: '인쇄 전 무지 롤',
  },
  {
    name: 'circuit-macro',
    w: 1200,
    h: 1200,
    must: ['tissue', 'paper', 'towel', 'napkin', 'texture', 'soft'],
    ban: [...BAN, 'outdoor', 'tree', 'nature', 'grass', 'christmas', 'flower'],
    queries: ['white paper texture macro soft', 'tissue paper closeup white', 'soft paper surface macro'],
    alt: '엠보싱 표면 확대',
  },
  {
    /* MILL 은 'machine' 만으로도 통과되어 금속 연삭기가 걸렸습니다.
       태그에 종이가 반드시 있어야 통과하도록 앵커를 바꿨습니다. */
    name: 'fab-equipment',
    w: 1600,
    h: 1067,
    must: ['paper', 'pulp', 'mill', 'cardboard', 'carton', 'tissue'],
    ban: [...BAN, 'metal', 'steel', 'grinding', 'welding', 'lathe', 'cnc', 'iron'],
    queries: ['paper mill production rolls', 'paper factory rolls machine', 'cardboard production factory'],
    alt: '크라프트 원지 롤',
  },
  {
    name: 'probe-station',
    w: 1600,
    h: 1067,
    must: PAPER,
    queries: ['drying hands paper towel', 'paper towel dispenser wall', 'hand towel dispenser bathroom'],
    alt: '핸드타월 디스펜서',
  },
  {
    name: 'lab-optics',
    w: 1600,
    h: 1067,
    must: PAPER,
    queries: ['baking paper parchment tray', 'parchment paper baking sheet', 'baking paper oven tray'],
    alt: '주방용 종이호일',
  },
  {
    name: 'semiconductor-line',
    w: 1600,
    h: 1067,
    must: MILL,
    queries: ['production line conveyor factory', 'packaging line factory automation', 'industrial conveyor belt production'],
    alt: '포장 · 생산 라인',
  },
  {
    name: 'server-rack',
    w: 1600,
    h: 1067,
    must: MILL,
    queries: ['warehouse boxes shelves storage', 'cardboard boxes warehouse pallet', 'logistics warehouse interior boxes'],
    alt: '박스 단위 보관 창고',
  },
  {
    name: 'silicon-ingot',
    w: 1200,
    h: 1200,
    must: ['paper', 'pulp', 'cardboard', 'carton', 'mill', 'roll'],
    ban: [...BAN, 'metal', 'steel', 'toilet', 'bathroom', 'bag', 'shopping', 'gift', 'handle'],
    queries: ['paper rolls stacked warehouse industry', 'paper reel mill production', 'cardboard rolls industrial storage'],
    alt: '원지 롤 보관',
  },

  /* ── 사용처 로고월 16칸 ────────────────────────────────────────────────
     "이런 곳에서 씁니다" 그리드. 4열 × 2행 = 8칸이고 칸마다 2장이 무한 순환하므로
     정확히 16장이 필요합니다. 그리드 한 칸은 화면에서 작으니 800×600 이면 충분하고,
     16장을 한 번에 받는 자리라 해상도를 키우면 페이지만 무거워집니다.

     톤 창을 넓게(30~190) 잡은 것은 실내 사진이 밝기 편차가 크기 때문입니다.
     어두운 페이지와의 조화는 CSS 쪽에서 처리합니다 — 평소에는 살짝 탈색·감광했다가
     hover 에서 원래 색으로 돌아옵니다(components.css > .logo-card). */
  ...[
    ['where-restaurant', '외식 프랜차이즈', ['restaurant interior dining tables', 'restaurant dining room interior']],
    ['where-catering', '단체 급식', ['cafeteria canteen interior tables', 'canteen food service counter']],
    /* 첫 패스에서 수술실(환자가 시트에 덮인 채)이 걸렸습니다. 위생용지 사용처 벽에
       놓기에 부적절해 시술·환자 어휘를 통째로 막았습니다. */
    ['where-care', '요양 · 병원', ['hospital room empty beds', 'nursing home bedroom interior', 'clinic waiting room chairs'],
      ['surgery', 'surgical', 'operation', 'operating', 'patient', 'blood', 'doctor', 'nurse', 'emergency', 'icu']],
    ['where-office', '사무 빌딩', ['office desks workspace interior', 'open office interior desks']],
    /* 첫 패스: 창밖을 보는 아이 혼자. 교실이 아니라 정서적인 인물 사진이었습니다. */
    ['where-school', '학교 · 학원', ['empty classroom desks chairs', 'school classroom interior desks', 'lecture hall seats rows'],
      ['child', 'girl', 'boy', 'kid', 'student', 'person', 'people', 'portrait', 'sad', 'alone']],
    ['where-hotel', '호텔 · 펜션', ['hotel room bed interior', 'guest room hotel interior']],
    ['where-cafe', '카페 · 베이커리', ['cafe counter interior coffee', 'bakery shop interior counter']],
    ['where-logistics', '물류 · 공장', ['warehouse racking logistics interior', 'factory production hall interior']],
    ['where-public', '공공기관', ['library reading room shelves interior', 'public building lobby hall interior'],
      ['night', 'neon', 'dark', 'club']],
    ['where-market', '마트 · 유통', ['supermarket aisle shelves', 'grocery store aisle interior']],
    /* 첫 패스: 붉은 조명의 계단·터널. 스포츠 시설로 읽히지 않았습니다. */
    ['where-sports', '스포츠 시설', ['gymnasium basketball court indoor', 'sports hall indoor court', 'indoor swimming pool facility'],
      ['stair', 'stairs', 'tunnel', 'neon', 'club', 'night', 'corridor']],
    ['where-camping', '캠핑 · 레저', ['campsite tents outdoor daylight', 'camping tent forest campsite'],
      ['aurora', 'northern lights', 'galaxy', 'milky']],
    /* 첫 패스: 빈 회의실처럼 보이는 방. 장난감·놀이 어휘로 좁혔습니다. */
    ['where-kindergarten', '어린이집 · 유치원', ['kindergarten playroom toys colorful', 'daycare room toys interior', 'children play area toys'],
      ['meeting', 'conference', 'office', 'boardroom']],
    ['where-salon', '미용실 · 살롱', ['hair salon interior chairs mirror', 'beauty salon interior']],
    ['where-fitness', '헬스장', ['gym equipment interior fitness', 'fitness studio interior equipment']],
    /* 첫 패스: 핸드백 편집숍. 편의점이 아니었습니다. */
    ['where-store', '편의점 · 소매', ['convenience store shelves drinks interior', 'small grocery shop shelves interior'],
      ['bag', 'handbag', 'fashion', 'boutique', 'clothing', 'clothes', 'shoe']],
  ].map(([name, label, queries, extraBan = []]) => ({
    name,
    w: 800,
    h: 600,
    must: [...PLACE, 'shop', 'store', 'gym', 'salon', 'camp', 'tent', 'library', 'classroom', 'school', 'supermarket', 'warehouse', 'canteen', 'cafeteria', 'court', 'playroom', 'toy'],
    ban: [...BAN, 'sketch', 'mockup', ...extraBan],
    tone: { min: 30, max: 190 },
    queries,
    alt: `${label} — 위생용지 사용처`,
  })),
]

/* ── providers ──────────────────────────────────────────────────────────── */

const usedIds = new Set()
const usedUrls = new Set()
const rejected = { tone: 0, offTopic: 0, tooSmall: 0 }

async function searchPixabay(query, landscape) {
  if (!PIXABAY_KEY) return []
  const url =
    `https://pixabay.com/api/?key=${PIXABAY_KEY}` +
    `&q=${encodeURIComponent(query)}&image_type=photo&safesearch=true&per_page=50` +
    `&orientation=${landscape ? 'horizontal' : 'vertical'}&min_width=1600&order=popular`
  const res = await fetch(url)
  if (!res.ok) return []
  const json = await res.json()
  return (json.hits || []).map((h) => ({
    id: `pixabay:${h.id}`,
    url: h.fullHDURL || h.largeImageURL,
    width: h.imageWidth,
    height: h.imageHeight,
    credit: `Pixabay / ${h.user}`,
    provider: 'Pixabay',
    tags: String(h.tags || '').toLowerCase(),
  }))
}

async function searchPexels(query, landscape) {
  if (!PEXELS_KEY) return []
  const url =
    `https://api.pexels.com/v1/search?query=${encodeURIComponent(query)}` +
    `&per_page=50&orientation=${landscape ? 'landscape' : 'portrait'}&size=large`
  const res = await fetch(url, { headers: { Authorization: PEXELS_KEY } })
  if (!res.ok) return []
  const json = await res.json()
  return (json.photos || []).map((p) => ({
    id: `pexels:${p.id}`,
    url: p.src.original,
    width: p.width,
    height: p.height,
    credit: `Pexels / ${p.photographer}`,
    provider: 'Pexels',
    // Pexels has no tag list; its alt text is the closest usable description.
    tags: String(p.alt || '').toLowerCase(),
  }))
}

/**
 * Topical relevance guard.
 *
 * 검색 순위만으로는 부족합니다. "paper" 한 단어로도 벽지·서류·낙엽이 딸려
 * 옵니다. 그래서 슬롯마다 `must` 를 선언해, 제공자가 붙인 태그/alt 안에 그
 * 단어가 하나라도 있어야 통과시킵니다. 순위가 아무리 높아도 주제에서 벗어나면
 * 건너뜁니다.
 */
function isRelevant(hit, must, ban) {
  if (ban && hit.tags && ban.some((word) => hit.tags.includes(word))) return false
  if (!must || !must.length) return true
  if (!hit.tags) return false
  return must.some((word) => hit.tags.includes(word))
}

/* ── pipeline ───────────────────────────────────────────────────────────── */

async function download(url) {
  const res = await fetch(url, { headers: { 'User-Agent': 'Mozilla/5.0 ilumos-build' } })
  if (!res.ok) throw new Error(`HTTP ${res.status}`)
  return Buffer.from(await res.arrayBuffer())
}

/**
 * 가장자리 색 샘플링.
 *
 * `fit: 'contain'` 으로 캔버스를 넓힐 때 남는 여백을 무엇으로 채울지 정합니다.
 * 원본 제품컷은 흰 배경 스튜디오 촬영이라 좌상단 모서리 평균색을 그대로 쓰면
 * 이어붙인 경계가 보이지 않습니다.
 */
async function edgeColor(buffer) {
  const meta = await sharp(buffer).metadata()
  const s = Math.max(2, Math.round(Math.min(meta.width, meta.height) * 0.04))
  const stats = await sharp(buffer).extract({ left: 0, top: 0, width: s, height: s }).stats()
  const [r, g, b] = stats.channels.map((c) => Math.round(c.mean))
  return { r, g, b, alpha: 1 }
}

async function toWebp(buffer, slot) {
  const fit = slot.fit || 'cover'
  const pipe = sharp(buffer).resize(slot.w, slot.h, {
    fit,
    position: fit === 'cover' ? 'attention' : 'centre',
    background: fit === 'contain' ? await edgeColor(buffer) : undefined,
    kernel: 'lanczos3',
  })
  // 원본 제품컷은 600px 급이라 확대가 조금 들어갑니다. 가벼운 샤픈으로
  // 엠보싱 결이 뭉개지지 않게 잡아 줍니다.
  if (slot.from) pipe.sharpen({ sigma: 0.7 })
  return pipe.webp({ quality: 86, effort: 5 }).toBuffer()
}

/**
 * Tone guard.
 *
 * 사이트는 하나의 어두운 톤으로 흐르지만, 위생용지는 품목 자체가 흽니다.
 * 그래서 반도체 테마에서 쓰던 132 상한을 205 까지 올려 잡았습니다. 완전히
 * 날아간 프레임(디테일 없음)과 새까만 프레임만 걸러냅니다. 원본 제품 사진은
 * 애초에 흰 배경이 정상이므로 이 검사를 건너뜁니다.
 */
const TONE = { min: 18, max: 205 }

const toneWindow = (slot) => ({ ...TONE, ...(slot.tone || {}) })

async function toneOf(buffer) {
  const { channels } = await sharp(buffer).stats()
  const [r, g, b] = channels
  // Rec. 601 luma — close enough for a threshold test.
  return 0.299 * r.mean + 0.587 * g.mean + 0.114 * b.mean
}

/**
 * Monochrome placeholder, in the site's own visual language. Only used when a
 * slot genuinely could not be filled; each one is reported at the end.
 */
async function placeholder(slot) {
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${slot.w}" height="${slot.h}">
    <rect width="100%" height="100%" fill="#0a0a0a"/>
    <g stroke="rgba(255,255,255,0.16)" stroke-width="1">
      ${Array.from({ length: 12 }, (_, i) => `<line x1="0" y1="${(slot.h / 12) * i}" x2="${slot.w}" y2="${(slot.h / 12) * i}"/>`).join('')}
    </g>
    <circle cx="${slot.w / 2}" cy="${slot.h / 2}" r="${Math.min(slot.w, slot.h) * 0.16}" fill="none" stroke="rgba(255,255,255,0.5)" stroke-width="2"/>
    <text x="50%" y="${slot.h / 2 + Math.min(slot.w, slot.h) * 0.28}" text-anchor="middle"
          font-family="monospace" font-size="${Math.min(slot.w, slot.h) * 0.035}"
          fill="rgba(255,255,255,0.55)" letter-spacing="4">IMAGE PENDING</text>
  </svg>`
  return sharp(Buffer.from(svg)).webp({ quality: 86 }).toBuffer()
}

/* ── 1. 원본 사진에서 채우는 슬롯 ───────────────────────────────────────── */

async function fillFromOriginal(slot) {
  const src = join(ROOT, 'original', 'images', slot.from)
  const raw = await readFile(src)
  const meta = await sharp(raw).metadata()
  const webp = await toWebp(raw, slot)
  await writeFile(join(OUT, `${slot.name}.webp`), webp)
  console.log(
    `  ✓ ${slot.name} <- original/${slot.from} (${meta.width}×${meta.height} -> ${slot.w}×${slot.h} ${slot.fit || 'cover'}, ${Math.round(webp.length / 1024)}kb)`
  )
  return {
    src: `/assets/images/${slot.name}.webp`,
    width: slot.w,
    height: slot.h,
    alt: slot.alt,
    credit: `(주)이루모스 원본 사진 (original/images/${slot.from})`,
    source: 'original',
  }
}

/* ── 2. 스톡 검색으로 채우는 슬롯 ───────────────────────────────────────── */

async function fillFromStock(slot) {
  /* 기본은 슬롯 비율을 따라갑니다. 다만 세로 슬롯인데 그 주제의 세로 사진이
     거의 없는 경우가 있어(스톡의 "paper towel" 세로컷이 그렇습니다),
     `searchLandscape` 로 가로에서 찾아 attention 크롭으로 세로를 만듭니다. */
  const landscape = slot.searchLandscape ?? slot.w >= slot.h
  const providers = slot.preferPexels ? [searchPexels, searchPixabay] : [searchPixabay, searchPexels]

  for (const query of slot.queries) {
    for (const search of providers) {
      let results = []
      try {
        results = await search(query, landscape)
      } catch (err) {
        console.warn(`  ! ${search.name} "${query}" failed: ${err.message}`)
        continue
      }

      for (const hit of results) {
        if (usedIds.has(hit.id) || usedUrls.has(hit.url)) continue
        if (hit.width < slot.w * 0.9) continue // reject upscales (reported size)
        if (!isRelevant(hit, slot.must, slot.ban ?? BAN)) { rejected.offTopic += 1; continue }

        try {
          const raw = await download(hit.url)

          /* 실제로 내려받은 파일의 해상도로 다시 검사합니다. 제공자가 보고하는
             크기와 무료 API 가 실제로 주는 파일 크기가 다를 수 있기 때문입니다
             (Pixabay 무료 엔드포인트는 1920px 로 잘립니다). 히어로처럼
             `minReal` 을 건 자리에서는 이 검사가 유일하게 믿을 수 있는 기준입니다. */
          const realMeta = await sharp(raw).metadata()
          const floor = slot.minReal || Math.round(slot.w * 0.9)
          if (realMeta.width < floor) { rejected.tooSmall += 1; continue }

          const webp = await toWebp(raw, slot)

          const luma = await toneOf(webp)
          const win = toneWindow(slot)
          if (luma > win.max || luma < win.min) {
            rejected.tone += 1
            continue // 완전히 날아갔거나 디테일이 없는 프레임
          }

          await writeFile(join(OUT, `${slot.name}.webp`), webp)
          usedIds.add(hit.id)
          usedUrls.add(hit.url)
          console.log(
            `  ✓ ${slot.name} <- ${hit.provider} "${query}" (${realMeta.width}px, ${Math.round(webp.length / 1024)}kb, luma ${luma.toFixed(0)})`
          )
          return {
            src: `/assets/images/${slot.name}.webp`,
            width: slot.w,
            height: slot.h,
            alt: slot.alt,
            credit: hit.credit,
            /* provider id 를 매니페스트에 남깁니다. 이게 없으면 캐시된 슬롯의
               사진이 usedIds 에 등록되지 않아, 한 슬롯만 다시 뽑을 때 이미
               다른 자리에 쓰인 사진이 또 걸립니다. 실제로 히어로와 같은
               촬영분이 세로 슬롯에 다시 들어온 적이 있습니다. */
            id: hit.id,
            url: hit.url,
            query,
            luma: Math.round(luma),
            source: 'stock',
          }
        } catch (err) {
          console.warn(`  ! download failed (${hit.id}): ${err.message}`)
        }
      }
    }
  }

  const webp = await placeholder(slot)
  await writeFile(join(OUT, `${slot.name}.webp`), webp)
  console.warn(`  ✗ ${slot.name} -> PLACEHOLDER (no usable result)`)
  return { src: `/assets/images/${slot.name}.webp`, width: slot.w, height: slot.h, alt: slot.alt, placeholder: true }
}

const fillSlot = (slot) => (slot.from ? fillFromOriginal(slot) : fillFromStock(slot))

/* ── run ────────────────────────────────────────────────────────────────── */

await mkdir(OUT, { recursive: true })
await mkdir(join(ROOT, 'assets', 'video'), { recursive: true })

if (!PIXABAY_KEY && !PEXELS_KEY) {
  console.warn('No PIXABAY_KEY / PEXELS_KEY in env — 스톡 슬롯은 전부 플레이스홀더가 됩니다.')
}

// Keep already-downloaded slots unless FORCE=1, so re-running the build is cheap.
const force = process.env.FORCE === '1'
const manifestPath = join(OUT, 'manifest.json')
const manifest = existsSync(manifestPath) && !force ? JSON.parse(await readFile(manifestPath, 'utf8')) : {}

for (const slot of SLOTS) {
  const have = manifest[slot.name]
  if (have && !have.placeholder && existsSync(join(OUT, `${slot.name}.webp`)) && !force) {
    console.log(`  = ${slot.name} (cached)`)
    if (have.id) usedIds.add(have.id)
    if (have.url) usedUrls.add(have.url)
    continue
  }
  manifest[slot.name] = await fillSlot(slot)
}

/* 이번 테마에서 더 이상 쓰지 않는 슬롯이 예전 매니페스트에 남아 있으면 지웁니다.
   남겨 두면 CREDITS.md 에 유령 항목이 생깁니다. */
const live = new Set(SLOTS.map((s) => s.name))
for (const key of Object.keys(manifest)) if (!live.has(key)) delete manifest[key]

await writeFile(manifestPath, JSON.stringify(manifest, null, 2), 'utf8')

const placeholders = Object.entries(manifest).filter(([, v]) => v.placeholder)
const fromOriginal = Object.values(manifest).filter((v) => v.source === 'original').length
console.log(`\n${SLOTS.length} slots -> assets/images (원본 ${fromOriginal} · 스톡 ${SLOTS.length - fromOriginal - placeholders.length} · 플레이스홀더 ${placeholders.length})`)
console.log(`guards rejected: ${rejected.offTopic} off-topic, ${rejected.tone} wrong tone, ${rejected.tooSmall} too small`)
if (placeholders.length) console.log('placeholders:', placeholders.map(([k]) => k).join(', '))

// Credits file — attribution is not required by either licence, but shipping it
// keeps the provenance of every asset checkable.
await writeFile(
  join(OUT, 'CREDITS.md'),
  `# 이미지 출처\n\n` +
    `제품 사진은 (주)이루모스 원본 사이트의 사진을 WebP 로 변환해 쓴 것입니다.\n` +
    `그 외 사진은 상업적 이용이 무료인 스톡(Pixabay / Pexels 라이선스)입니다.\n\n` +
    Object.entries(manifest)
      .map(([k, v]) => `- \`${k}.webp\` — ${v.credit || 'generated placeholder'}${v.query ? ` (query: "${v.query}")` : ''}`)
      .join('\n') +
    '\n',
  'utf8'
)
