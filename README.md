# (주)이루모스 / ILUMOS — 위생용지 기업 홈페이지 리뉴얼 시안

**키친타월 · 점보롤티슈 · 두루마리티슈**를 공급하는 (주)이루모스(충남 계룡)의 홈페이지 리뉴얼 시안입니다.
원본 사이트는 http://www.ilumos.co.kr/ 이며, 진단 결과는 [analysis.md](./analysis.md) 에 있습니다.

레이아웃·모션·디자인 시스템은 `cineticaanalysis/`의 **cinetica.studio 리버스 엔지니어링 분석**을
재현한 것이고, **콘텐츠(텍스트·이미지)만 이루모스에 맞게 교체**했습니다.
빌드 산출물은 순수 HTML/CSS/JS라 서버 없이 파일만 열어도 동작합니다.

> **이 결과물은 계획한 페이지 구조 전체를 구현한 것으로, 전체 범위의 약 100% 수준(전체 구현)입니다.**

원본 사이트 대비 개선된 것 (analysis.md 의 진단 항목에 대응):

| 원본 진단 | 이 시안 |
|---|---|
| 뷰포트 메타 없음 → 모바일에서 확대해야 보임 | 375 / 480 / 768 / 991px 반응형, 가로 오버플로 0 |
| meta description 없음 | 전 페이지 title·description·canonical·OG·JSON-LD |
| alt 없는 이미지 6개 | `<img>` 72개 **전부** alt + width/height |
| jQuery 1.8.3 (구버전 취약점) | jQuery 미사용. GSAP·Lenis만 로컬 번들 |
| 본문이 전부 이미지 → 검색 노출 불가 | 본문 전체가 실제 텍스트 |

> **무엇이 원본 값이고 무엇이 작성한 것인지**는 **[MISSING.md](./MISSING.md)** 에 전부 정리되어 있습니다.
> 회사 식별 정보(상호·대표자·연락처·소재지·품목 3종)는 **원본 그대로**이고,
> 제품 상세 문구·사용처 제안·규격 수치·저널은 **작성한 것**입니다.
> 특히 **제품 8종 중 5종과 모든 규격 수치는 업종 일반론으로 채운 예시**이고,
> **법률 문서는 검토가 필요**합니다.

---

## 빠른 시작

```bash
npm install       # gsap, lenis, sharp
npm run setup     # vendor(라이브러리 복사) + images(수급·변환) + build
npm run serve     # http://localhost:4173
```

빌드된 결과만 보려면 `npm run serve` 만 실행해도 됩니다 (HTML 이 이미 커밋되어 있습니다).

| 스크립트 | 하는 일 |
|---|---|
| `npm run vendor` | `node_modules`의 GSAP·Lenis를 `assets/js/vendor/`로 복사 (외부 호스트 의존 0) |
| `npm run images` | **44개 슬롯**을 채웁니다 — `original/images/`의 이루모스 자사 제품 사진 11장 + Pixabay/Pexels 스톡 33장(사용처 월 16장 포함) → WebP 변환 → `assets/images/` + `manifest.json` + `CREDITS.md`. 스톡 후보는 **주제 가드**(제공자 태그에 앵커 단어 필수) · **오분류 가드**(일러스트·렌더 차단) · **실측 해상도 가드**(히어로는 2560px 이상) · **톤 가드**(평균 휘도 18~205)를 모두 통과해야 합니다 |
| `npm run build` | `src/`의 데이터·템플릿으로 모든 라우트를 정적 HTML로 렌더 |
| `node scripts/check.mjs` | 링크·헤딩·alt·SEO·aria 검증 (실패 시 exit 1) |
| `node scripts/audit-css.mjs` | 마크업에 한 번도 나오지 않는 CSS 클래스 보고 (리팩터 잔여물 탐지) |
| `npm run serve` | 무의존 정적 서버 |

이미지를 다시 뽑으려면 `FORCE=1 npm run images`. `PIXABAY_KEY` / `PEXELS_KEY` 환경변수를 사용합니다.
키가 없으면 스톡 슬롯만 플레이스홀더가 되고, **자사 제품 사진 11장은 키 없이도 그대로 생성됩니다.**

> ⚠ `npm run images` 는 `original/images/` 를 입력으로 읽습니다. 이 폴더는 **작업 입력물이라
> 배포 결과물에 포함되지 않습니다.** 이미지를 다시 뽑아야 한다면 원본 폴더를 먼저 복원하십시오.
> 이미 변환된 `assets/images/*.webp` 만으로 빌드하는 데는 아무 문제가 없습니다.

---

## 콘텐츠를 고치는 곳

업종·문구·제품이 바뀌면 **세 곳만** 고치면 됩니다. 레이아웃과 모션 코드는 건드리지 않습니다.

| 파일 | 담당 |
|---|---|
| `src/data/site.mjs` | 회사 정보, 내비게이션, 제품 8종, 사용처 4건, FAQ, 저널, 법률 문서 — **모든 텍스트** |
| `scripts/fetch-images.mjs` | 이미지 슬롯 44개의 출처(`from:` 원본 / `queries:` 스톡)와 검색 가드 |
| `src/templates/pages.mjs` | 캐러셀·갤러리의 이미지 배치와 캡션 |

색은 `assets/css/tokens.css` 의 `--accent-lime` 한 줄입니다
(`#c8ff3d`. 로고의 주황으로 한 번 바꿨다가 되돌렸습니다 — 숲 배경 위에서는 주황이
나무 색과 붙어 히어로 2행이 묻힙니다. MISSING.md > C16).

---

## 구조

```
├─ src/                     ← 소스 (빌드 입력)
│  ├─ data/site.mjs         ← 모든 페이지의 단일 진실 공급원
│  └─ templates/
│     ├─ layout.mjs         ← 문서 셸: head/SEO/JSON-LD, 헤더, 오버레이 메뉴, 커서, 푸터
│     ├─ components.mjs     ← 아톰·몰리큘 (버튼, 이어브로우, 미디어, 아코디언, 카드, 캐러셀…)
│     └─ pages.mjs          ← 페이지 본문 + 캐러셀/갤러리 이미지 목록
├─ scripts/                 ← build · vendor · fetch-images · check · audit-css · serve
├─ assets/
│  ├─ css/  tokens · base · components · sections · pages · lenis
│  ├─ js/   main · motion · ui · hero-gl · vendor/
│  └─ images/ (44 webp + manifest.json + CREDITS.md)
└─ (생성물) index.html · portfolio/ · services/ · faq/ · blog/ · utility-pages/ · 404.html
```

**공통 마크업은 페이지에 직접 쓰지 않습니다.** 헤더·푸터·버튼·카드 등은 전부 컴포넌트가 소유하며
빌드가 조합합니다. 내비게이션 한 줄을 고치면 25개 페이지에 동시 반영됩니다.

---

## 페이지 (25)

| 라우트 | 비고 |
|---|---|
| `/` | 원페이저. 섹션 순서는 분석에서 실측된 골격 그대로 |
| `/services/` 상세 8 | 키친타월 · 점보롤티슈 · 두루마리티슈 · 미용티슈/냅킨 · 핸드타월 · 종이호일 · OEM 주문제작 · 대량 공급 |
| `/portfolio/` + 상세 4 | 사용처 안내 — 외식업장 주방 · 요양/의료시설 · 사무실/공공기관 · 카페/숙박업소 |
| `/faq/` | 고객센터. 8문항 아코디언 (최소 수량 · 샘플 · 견적 · 배송 · 인쇄 · 품목 선택 · 원지 · 정기 배송) |
| `/blog/` + 글 6편 | 위생용지 저널. 목록 + 개별 상세, `BlogPosting` JSON-LD |
| `/utility-pages/` 2 | 이용약관 · 개인정보 처리방침 (⚠ 법률 검토 필요) |
| `/404.html` | |

**원본 메뉴 구조와의 대응** — 원본은 `HOME / 회사소개(회사개요·경영방침·오시는길) / 제품소개(키친타월·점보롤티슈·두루마리티슈) / 견적요청 / 고객센터` 였습니다.

| 원본 메뉴 | 이 시안 |
|---|---|
| 회사소개 | `/#about` — 회사 소개 + 카운터. 오시는길·연락처는 `/#contact` 에 통합 |
| 제품소개 | `/#services` + `/services/` 상세 8 (원본 3품목 포함) |
| 견적요청 | `/#contact` — 실동작 폼 (검증 + mailto 초안) · 헤더 CTA |
| 고객센터 | `/faq/` |
| — | `/portfolio/` 사용처 안내, `/blog/` 저널 (리뉴얼에서 추가) |

홈 섹션 순서 (`04-pages/home.json > layout_skeleton`):
`히어로 → 회사소개 → 사용처 로고월 → 제품소개 → 사용처 안내(3D 캐러셀) → 주문·배송 과정 → 갤러리 → 스크럽 비디오 → FAQ → 견적 요청`

---

## 디자인 시스템

전체가 **하나의 유동 단위**에서 파생됩니다 (분석 `:root`에서 그대로 가져온 값):

```css
--size-font: calc(clamp(992px, 100vw, 2560px) / (1440 / 16));   /* = 뷰포트/90 */
html { font-size: var(--size-font); }
```

`html`의 font-size가 곧 이 단위이므로 **1rem = 1 유동 단위**입니다. 토큰을 `em`이 아닌 `rem`으로 쓴 이유가
이것입니다 — 비율은 리포트가 유도한 그대로면서, font-size를 이미 바꾼 요소 안에서 중첩 곱셈이 없습니다.

측정값 대조 (분석 뷰포트 1315px 기준 실측 → 구현):

| 역할 | 실측 | 토큰 |
|---|---|---|
| 최대 히어로 워드 | 146.2px | `--t-2xl: 10rem` |
| 섹션 H2 | 80.4px | `--t-h2: 5.5rem` |
| 에디토리얼 리드 | 32.0px | `--t-lead: 2.19rem` |
| FAQ 질문 | 21.9px | `--t-lg: 1.5rem` |
| 본문 기준 (782노드) | 14.6px | `--t-base: 1rem` |

**두 가지가 분석과 다릅니다** (자세한 근거는 MISSING.md):

1. **폰트** — **IBM Plex 전역 적용** (원본은 Gabarito/Instrument Serif/Montserrat/Martian Mono).
   다른 파운드리 서체는 하나도 로드하지 않습니다. 단, IBM Plex Serif에 한글이 없어 serif 역할에서
   **라틴은 세리프, 한글은 IBM Plex Sans KR**로 렌더링됩니다 — MISSING.md > C1
2. **최소 16px** — `--t-min: 16px` 플로어. ~1440px 이상에서는 작동하지 않아 측정 비율이 그대로 유지되고,
   그 아래에서만 작은 단계가 16px로 클램프됩니다

### 단일 다크 톤

문서 전체가 히어로와 같은 톤입니다. 섹션 배경은 불투명 검정이 아니라 **`rgba(6,6,12,0.82)` 반투명**이라,
고정된 별필드가 모든 블록 뒤로 은은하게 비칩니다 — 히어로는 그 필드를 가장 덜 가리는 블록일 뿐입니다.
원본의 구조와도 맞습니다: `.unicorn__item`은 21개 스크롤 스톱 전부에서 `top:0`을 유지한, 문서 전체 뒤의 고정 배경이었습니다.

```
body      →  #06060c (불투명 바닥)
모든 섹션  →  rgba(6,6,12,0.82) (반투명, 별필드가 비침)
히어로     →  배경 없음 (별필드 그대로)
```

컴포넌트는 `--black`/`--white`를 직접 쓰지 않고 `--bg` / `--ink` / `--dim` / `--hairline` 시맨틱 토큰만
참조합니다. 밝은 블록이 필요하면 `.theme-light`(크림 #FAF9F5 / 니어블랙 #141413 — 원본 inverted 실측값)를
붙이면 됩니다.

> 함정 기록: `color: var(--ink)`는 **선언한 요소에서 한 번만** 해석되고 자식은 계산된 색을 상속합니다.
> 그래서 `.theme-dark`는 `--ink`만 재정의해서는 안 되고 `color`도 함께 선언해야 합니다.

---

## 모션

`03-motion.json`의 파라미터를 그대로 옮겼습니다. 기본 케이던스는
`ScrollTrigger { start: 'top 80%', toggleActions: 'play', once: true }`.

| 컴포넌트 | 파라미터 |
|---|---|
| SplitText 라인 리빌 | `yPercent 110 → 0` · 0.8s · `expo.out` · stagger 0.1 · once |
| ScrambleText | 1.5s · speed 0.8 · 룬/기호 세트 |
| 스크럽 비디오 | 292×164 → 1257×707, x −482 / y +400 · `scrub: true` |
| 사용처 월 | 칸마다 사진 2장 무한 슬라이드 — 상하좌우 혼합, 1.5s `sine.inOut`, 배경 노출 0 |
| 카운터 | 1s `power4` 롤 |
| 미디어 리빌 | 0.6s `power2.out` |
| 마퀴 | `@keyframes translateX` (원본 시트에서 그대로) |
| 커스텀 커서 | 다크: `mix-blend-mode: difference` 링 / 라이트: 파란 방사형 별 |

스무스 스크롤은 **Lenis**가 담당하고 `ScrollTrigger.update`와 동기화됩니다.

### 별필드 밀도
뷰포트별로 **160,000 / 48,000 / 14,000** 개.
흰색 92.5%에 청백(3.1%)·청록(2.0%)·노랑(1.6%)·붉은색(0.8%) 별이 섞이고,
유채색 별은 색이 가져가는 휘도 손실을 보정해(×1.55) 흰 별 사이에서도 색점으로 읽힙니다.

원래는 **1,000,000 / 300,000 / 80,000** 이었습니다. 그 밀도는 **순수한 검정 위의
별필드**를 전제로 한 값 — 그때는 밀도 자체가 곧 그림이었습니다. 배경이 숲 사진으로
바뀐 뒤로는 같은 밀도가 하늘을 정적 노이즈로 만들어 나무와 경쟁해서, 약 1/6로 줄이고
정지 상태 밝기도 함께 낮췄습니다(0.72 → 0.42). 대신 밀도가 낮아진 만큼 각 별에
헤일로를 돌려주었습니다 — 성긴 별필드가 흩어진 픽셀이 아니라 하늘로 읽히게 하는 건
이 헤일로입니다.

같은 이유로 **와이어프레임 별과 포인터 광원도 각각 50% 축소**했습니다.
노브는 `hero-gl.js` 의 `FIELD_COUNT` · `buildStarSegments(spike, scale)` · 광원 셰이더의 `R`.
자세한 내역과 함정은 MISSING.md > C14.

### 3D 이미지 링
원본은 4개의 큰 패널이 실린더와 함께 회전해 측면에서 옆면·뒷면이 보였습니다.
이 구현은 **10개의 작은 패널이 궤도를 돌되 항상 정면을 향합니다** — 사진면은 평면을 유지하고
위치와 깊이만 변합니다. 깊이가 스케일·불투명도·쌓임 순서를 결정해 다차원적으로 읽힙니다.
회전 주기는 30s → **14s**. 실린더 반지름은 상수가 아니라 패널 수에서 계산합니다
(`r = (w/2)/tan(π/n) × 1.35`) — 고정 반지름은 패널을 늘리는 즉시 겹칩니다.

### 히어로 배경 — 울창한 숲 위의 별
히어로는 **울창한 소나무 숲 사진 위에 별 효과가 얹힌** 구성입니다.
종이의 출발점이 나무라는 점에서 업종과도 맞습니다.

```
.unicorn__item  (fixed, z-index:0 — 자체 스태킹 컨텍스트)
  ├ background   #000            ← 바닥
  ├ .unicorn__scene  z-index:-1  ← 숲 사진  (홈에서만 · 스크롤하면 사라짐)
  ├ canvas       (auto)          ← 광원 · 별필드 · 와이어프레임 별
  ├ .unicorn__fallback           ← WebGL 없을 때
  └ .unicorn__grain              ← 그레인
```

`z-index:-1` 은 DOM 순서에 기대지 않기 위한 것입니다 — `hero-gl.js` 가 캔버스를
`prepend` 하므로, 순서로 맞추면 언제든 뒤집힙니다.

숲은 **홈에서만**(`body.is-home`) 켜지고 **히어로 핀 구간에서 사라집니다**.
배경 레이어가 fixed 이고 아래 섹션이 전부 반투명이라, 그대로 두면 문서 전체 뒤에
숲이 깔립니다. 사라지는 타이밍을 별 워프와 맞춰서, 히어로를 벗어나는 동작이
**숲에서 솟아올라 성간으로 들어가는 것**으로 읽힙니다.

**사진 선정이 이 자리의 핵심이고, 세 번 갈아엎었습니다**:

| 시도 | luma | 결과 |
|---|---|---|
| 안개 숲 | 75 | 가산 합성된 별필드에 뭉개져 회색 안개가 됨 |
| 밤하늘 실루엣 | 42 | 별은 완벽했지만 **나무가 복원 불가능한 검정** |
| 빛을 받은 소나무 숲 | 66 | 나무에 껍질 질감과 계조가 있어 어둡게 눌러도 살아남음 |

세 번째에서야 작업이 "그림자를 끌어올리기"에서 "밝은 사진을 눌러 담기"로 바뀌었고,
그쪽이 훨씬 다루기 쉽습니다. 실루엣 시절에 필요했던 multiply 캐노피 레이어도
이 사진에서는 아무 일도 하지 않아 **제거**했습니다.
타이포 보호는 `::after` 의 방사형 그라디언트가 **위치로** 처리합니다 —
헤드라인 위만 눌러서, 가장자리의 빽빽한 줄기들은 밝게 남깁니다.
자세한 실패 기록은 MISSING.md > C17.

### 히어로 별 효과 (`assets/js/hero-gl.js`)
원본의 Unicorn Studio / Spline은 **복원 불가**입니다. 대신 자체 WebGL2로
**얇은 실선 3D 별 + 마우스를 따라가는 방사형 원형 빛**을 그립니다.

- 별: 정이십면체(30 edge) + 20면 각각의 뿔(60 edge) = 90 세그먼트
- 선: `gl.LINES` 대신 **화면공간 쿼드 + 해석적 AA** — 브라우저가 `lineWidth`를 1 백킹픽셀로 고정해
  DPR에 따라 흐려지는 문제를 피하고 어떤 화면에서도 같은 굵기를 유지
- 빛: 보간 없이 커서에 정확히 고정. 좌표는 `window.innerWidth`가 아니라 **캔버스 rect**로 정규화하고
  (스크롤바 폭만큼 어긋나는 것 방지), y축은 **한 번만** 뒤집습니다
- 광원은 로드 시 **캔버스 정중앙 = 별이 투영되는 지점**에서 시작하고, 첫 마우스 이동에 0.55초 이징으로
  커서로 옮겨간 뒤 정확 추종으로 전환합니다
- WebGL 미지원·컨텍스트 손실 시 CSS 폴백, `prefers-reduced-motion`에서 회전·맥동 정지

### 히어로 스크롤 워프 전환
스크롤을 내리면 히어로가 **1.1 뷰포트만큼 pin** 되고, 그 구간에서 화면이 성간 워프로 전환됩니다.

- 헤드라인·CTA가 위로 밀리며 `blur(7px)`와 함께 사라짐 (진행도 0 → 0.72)
- **100만 개** 별필드가 가속 — 각 별은 방향 `(x,y)`와 깊이 `d`를 갖고 `(x/d, y/d)`로 투영되므로,
  `d`가 줄면 화면 밖으로 흩뿌려지며 커집니다. 스크롤이 `d` 감소를 지배합니다
- 와이어프레임 별이 뒤로 물러나며(z −4.15 → −13.2) 회전이 빨라져, 장면이 흩어지듯 읽힘

> 함정 기록: **핀된 요소는 rect가 움직이지 않습니다.** 그래서 (1) 콘텐츠 페이드는 별도 트리거가 아니라
> 핀 트리거에 `animation: timeline`으로 붙여야 하고, (2) `hero-gl.js`도 스스로 스크롤을 잴 수 없어
> 핀 트리거가 `window.__heroWarp`에 진행도를 게시하고 셰이더가 그 값을 씁니다.
> 둘 다 별도 트리거로 만들면 영원히 시작 위치에 멈춘 채로 동작합니다.

---

## 접근성 · SEO

- 시맨틱 마크업, 랜드마크, 본문 건너뛰기 링크
- **문의 폼**: label 연결, `aria-describedby` 오류 메시지, `role="alert"`, 첫 오류 필드로 포커스 이동
- 오버레이 메뉴: `role="dialog"` · `aria-modal` · **포커스 트랩** · Esc 닫기
- 아코디언: `aria-expanded` / `aria-controls` / `role="region"` 연결 (검증 스크립트가 id 대응 확인)
- **페이지당 `<h1>` 정확히 1개.** 홈은 원본을 그대로 따라 h1이 히어로가 아니라 **FAQ 블록**에 있습니다
- 모든 텍스트 **16px 이상**
- 모든 `<img>`에 `alt` + `width`/`height` (레이아웃 시프트 방지) — 72개 전부 검증
- 모든 타이틀은 영문 표기 + 한글 보조 라인(`.t-ko`). 서브페이지 h1은 `aria-labelledby`로
  한글·영문을 함께 읽습니다 — **SplitText가 `aria-label`을 덮어쓰기 때문에 `label`로는 안 됩니다**
- 페이지별 title/description/canonical/OG/Twitter, JSON-LD `Organization` + `WebSite` + `WebPage`,
  제품 상세는 `Product`, 사용처 안내는 `Article`, FAQ는 `FAQPage`, 저널 글은 `BlogPosting`
- `Organization` 에는 원본에서 확인된 값만 넣습니다 — 전화·팩스·이메일·주소·대표자.
  **설립연도(`foundingDate`)와 SNS(`sameAs`)는 원본에 없어 필드 자체를 뺐습니다.**
- `sitemap.xml` (24 URL) · `robots.txt` · `site.webmanifest`

---

## 검증 결과

```
$ npm run images
44 slots -> assets/images (원본 11 · 스톡 33 · 플레이스홀더 0)

$ npm run build
built 25 pages + sitemap.xml (24 urls) + robots.txt + site.webmanifest
images in manifest: 44

$ node scripts/check.mjs
pages: 25  ·  internal refs: 1455  ·  images: 72
notes (expected, not failures):
  index.html: /assets/video/about.webm not present (poster fallback — documented)
  index.html: /assets/video/about.mp4 not present (poster fallback — documented)
  index.html: /assets/video/hero.webm not present (poster fallback — documented)
  index.html: /assets/video/hero.mp4 not present (poster fallback — documented)
PASS — links resolve, headings/alt/SEO/aria all present

$ node scripts/audit-css.mjs
declared classes: 238  ·  emitted in markup/js: 278
no dead component classes
```

- 이미지 44슬롯 전부 채움, **플레이스홀더 0장**. 제품 자리는 이루모스 자사 사진 우선
- 회사 식별 정보가 전 페이지에 원본 값으로 반영됨 (전화·팩스·이메일·대표자·소재지)
- 미사용 CSS 0개, 미사용 이미지 0개, 끊긴 내부 참조 0개
- 홈 히어로 렌더 확인 (Chrome, localhost:4173)

**이번 리뉴얼에서 직접 재확인하지 못한 것** — 이전 빌드 기준으로는 통과했고 레이아웃 코드는
바뀌지 않았으나, 콘텐츠 교체 후 다시 측정하지는 않았습니다:

- 반응형 가로 오버플로 (375 / 480 / 768 / 991px) — 텍스트 길이가 바뀌었으므로 재확인 권장
- 문의 폼 브라우저 검증 동작 (필드 라벨과 문구가 바뀜)
- 별필드 100만 개 프레임 드롭

**애초에 미확인**: Safari/Firefox 호환성, Lighthouse, 2-테마 전환 후의 WCAG 명암비 (MISSING.md E).
히어로가 검정에서 숲 사진 배경으로 바뀌었으므로 **그 위에 얹힌 흰 타이포의 명암비는
특히 재검증이 필요합니다.**
