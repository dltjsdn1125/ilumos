/**
 * build.mjs — renders every route to static HTML.
 *
 * Route list is 01-discovery.json > internal_routes, minus the /es locale
 * (this rebuild ships ko as default — see MISSING.md > i18n).
 * Output is plain HTML: no runtime, no hydration, opens straight from disk.
 *
 * 콘텐츠는 (주)이루모스 리뉴얼용으로 교체되어 있습니다 — src/data/site.mjs 참고.
 */

import { mkdir, writeFile, readFile, rm } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { site, services, portfolio, legal, faq, posts, products, shopCategories, shopMeta } from '../src/data/site.mjs'
import { layout } from '../src/templates/layout.mjs'
import * as P from '../src/templates/pages.mjs'
import { starfieldCss } from '../src/templates/starfield.mjs'

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..')

/* Image manifest (written by fetch-images.mjs). Missing manifest is not fatal —
   components fall back to a default path and dimensions. */
let manifest = {}
const manifestPath = join(ROOT, 'assets', 'images', 'manifest.json')
if (existsSync(manifestPath)) {
  manifest = JSON.parse(await readFile(manifestPath, 'utf8'))
}

/* ── FAQPage structured data (real Q&A pairs — eligible for rich results) ── */
const faqJsonLd = {
  '@type': 'FAQPage',
  mainEntity: faq.map((f) => ({
    '@type': 'Question',
    name: f.q,
    acceptedAnswer: { '@type': 'Answer', text: f.a },
  })),
}

/* ── route table ────────────────────────────────────────────────────────── */

const routes = [
  {
    path: '/',
    title: site.title,
    description: site.description,
    body: P.home(manifest),
    jsonLdExtra: [faqJsonLd],
    priority: '1.0',
    /* Switches on the forest behind the star field. Home only — sub-pages keep
       the plain field. See sections.css > .unicorn__scene. */
    bodyClass: 'is-home',
  },
  {
    path: '/portfolio/',
    title: '사용처 안내',
    description: '외식업장 주방, 요양·의료시설, 사무실·공공기관, 카페·숙박업소 — 업종별로 어떤 위생용지를 어디에 두면 되는지 정리한 구성안.',
    body: P.portfolioIndex(manifest),
    priority: '0.9',
  },
  ...portfolio.map((w) => ({
    path: `/portfolio/${w.slug}/`,
    title: `${w.title} — ${w.category}`,
    description: w.summary,
    body: P.workDetail(w, manifest),
    ogType: 'article',
    priority: '0.8',
    jsonLdExtra: [
      {
        '@type': 'Article',
        headline: w.title,
        alternativeHeadline: w.titleLatin,
        articleSection: w.category,
        about: w.sector,
        description: w.summary,
        inLanguage: site.lang,
        author: { '@id': `${site.origin}/#organization` },
        publisher: { '@id': `${site.origin}/#organization` },
        mainEntityOfPage: { '@type': 'WebPage', '@id': `${site.origin}/portfolio/${w.slug}/#webpage` },
      },
    ],
  })),
  ...services.map((s) => ({
    path: `/services/${s.slug}/`,
    title: s.title,
    description: s.summary,
    body: P.serviceDetail(s, manifest),
    priority: '0.8',
    jsonLdExtra: [
      {
        '@type': 'Product',
        name: s.title,
        alternateName: s.titleLatin,
        category: '위생용지',
        description: s.summary,
        brand: { '@id': `${site.origin}/#organization` },
        manufacturer: { '@id': `${site.origin}/#organization` },
      },
    ],
  })),
  /* ── shop ──────────────────────────────────────────────────────────────
     자사몰 라우트. 카탈로그 20종은 전부 정적으로 렌더되고, 장바구니만
     브라우저에 삽니다(shop.js). 결제·재고 연동은 없습니다 — MISSING.md > D7.

     상세 페이지의 Product JSON-LD 에는 Offer 를 넣되 `priceValidUntil` 이나
     재고 상태는 넣지 않습니다. 가격 자체가 시안용 예시라, 검색엔진에 확정된
     판매 조건인 것처럼 신호를 보내지 않기 위해서입니다. */
  {
    path: '/shop/',
    title: '온라인 주문',
    description: '키친타월, 점보롤티슈, 두루마리티슈부터 냅킨·핸드타월·종이호일까지 자주 나가는 규격 20종을 골라 바로 주문하실 수 있습니다.',
    body: P.shopIndex(manifest),
    priority: '0.9',
    jsonLdExtra: [
      {
        '@type': 'ItemList',
        name: '(주)이루모스 온라인 주문 카탈로그',
        numberOfItems: products.length,
        itemListElement: products.map((p, i) => ({
          '@type': 'ListItem',
          position: i + 1,
          url: `${site.origin}/shop/${p.slug}/`,
          name: p.name,
        })),
      },
    ],
  },
  {
    path: '/shop/cart/',
    title: '장바구니 · 주문서',
    description: '담으신 품목을 확인하고 주문서를 보내시면 확정 견적과 입금 안내로 회신드립니다.',
    body: P.cartPage(manifest),
    priority: '0.4',
  },
  ...products.map((p) => {
    const cat = shopCategories.find((c) => c.slug === p.category)
    return {
      path: `/shop/${p.slug}/`,
      title: `${p.name} — 온라인 주문`,
      description: `${p.summary} ${p.pack} · 품목 코드 ${p.sku}. ${site.name}에서 바로 주문하실 수 있습니다.`,
      body: P.productDetail(p, manifest),
      ogType: 'product',
      priority: '0.7',
      jsonLdExtra: [
        {
          '@type': 'Product',
          name: p.name,
          alternateName: p.nameLatin,
          sku: p.sku,
          category: cat ? cat.label : '위생용지',
          description: p.summary,
          image: `${site.origin}${(manifest[p.image] || {}).src || `/assets/images/${p.image}.webp`}`,
          brand: { '@id': `${site.origin}/#organization` },
          manufacturer: { '@id': `${site.origin}/#organization` },
          offers: {
            '@type': 'Offer',
            /* 외부 스토어에 올린 품목은 실제로 살 수 있는 곳을 가리켜야 검색
               결과의 링크가 맞습니다 (docs/cafe24-integration-guidebook.html > 4-3). */
            url: p.externalUrl || `${site.origin}/shop/${p.slug}/`,
            price: p.price,
            priceCurrency: 'KRW',
            eligibleQuantity: { '@type': 'QuantitativeValue', minValue: 1, unitText: p.unit },
            seller: { '@id': `${site.origin}/#organization` },
          },
        },
      ],
    }
  }),
  {
    path: '/faq/',
    title: '고객센터 · 자주 묻는 질문',
    description: '최소 주문 수량, 샘플 요청, 견적 방법, 배송 기간, 로고 인쇄, 점보롤과 두루마리 선택 기준까지 — 견적 전 가장 많이 받는 질문과 답변.',
    body: P.faqPage(),
    jsonLdExtra: [faqJsonLd],
    priority: '0.7',
  },
  {
    path: '/blog/',
    title: '위생용지 저널',
    description: '점보롤과 두루마리 선택 기준, 키친타월 흡수 구조, 겹수 차이, 디스펜서 배치, 원지 선택, 보관 방법까지. 고르고 쓰는 데 도움이 되는 기록.',
    body: P.blogIndex(),
    priority: '0.6',
  },
  ...posts.map((p) => ({
    path: `/blog/${p.slug}/`,
    title: p.title,
    description: p.excerpt,
    body: P.blogPost(p, manifest),
    ogType: 'article',
    priority: '0.5',
    jsonLdExtra: [
      {
        '@type': 'BlogPosting',
        headline: p.title,
        alternativeHeadline: p.titleLatin,
        datePublished: p.date,
        articleSection: p.category,
        description: p.excerpt,
        inLanguage: site.lang,
        author: { '@id': `${site.origin}/#organization` },
        publisher: { '@id': `${site.origin}/#organization` },
        mainEntityOfPage: { '@type': 'WebPage', '@id': `${site.origin}/blog/${p.slug}/#webpage` },
      },
    ],
  })),
  ...legal.map((doc) => ({
    path: `/utility-pages/${doc.slug}/`,
    title: doc.title,
    description: `${site.name}의 ${doc.title}입니다. ${doc.sections.map((s) => s.h.replace(/^\d+\.\s*/, '')).join(' · ')} 등을 안내합니다.`,
    body: P.legalPage(doc),
    priority: '0.3',
    noSitemapPriority: true,
  })),
  {
    path: '/404.html',
    title: '페이지를 찾을 수 없습니다',
    description: '요청하신 페이지가 존재하지 않습니다.',
    body: P.notFound(),
    skipSitemap: true,
  },
]

/* ── clean previously generated HTML (never touches assets/ or src/) ─────── */

for (const dir of ['portfolio', 'services', 'utility-pages', 'faq', 'blog', 'shop']) {
  await rm(join(ROOT, dir), { recursive: true, force: true })
}

/* ── render ─────────────────────────────────────────────────────────────── */

let count = 0

for (const route of routes) {
  const html = layout(route, { services, portfolio })
  const outPath = route.path.endsWith('.html')
    ? join(ROOT, route.path.replace(/^\//, ''))
    : join(ROOT, route.path.replace(/^\//, ''), 'index.html')

  await mkdir(dirname(outPath), { recursive: true })
  await writeFile(outPath, html, 'utf8')
  count += 1
}

/* ── sitemap.xml ────────────────────────────────────────────────────────── */

/* lastmod 는 빌드한 날이 아니라 내용을 실제로 고친 날(site.lastmod)입니다. */
const today = site.lastmod
const sitemap = `<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
${routes
  .filter((r) => !r.skipSitemap)
  .map(
    (r) => `  <url>
    <loc>${site.origin}${r.path}</loc>
    <lastmod>${today}</lastmod>
    <priority>${r.priority || '0.5'}</priority>
  </url>`
  )
  .join('\n')}
</urlset>
`
await writeFile(join(ROOT, 'sitemap.xml'), sitemap, 'utf8')

/* ── robots.txt ─────────────────────────────────────────────────────────── */

await writeFile(
  join(ROOT, 'robots.txt'),
  `User-agent: *\nAllow: /\n\nSitemap: ${site.origin}/sitemap.xml\n`,
  'utf8'
)

/* ── llms.txt (AI 검색용 안내 — 사이트에 적힌 사실만) ─────────────────────── */

await writeFile(
  join(ROOT, 'llms.txt'),
  `# ${site.name}

> 충남 계룡의 위생용지 공급사입니다. 키친타월, 점보롤티슈, 두루마리티슈를 가정용 소량부터 업소용 대량 납품까지 공급하고, 로고 인쇄 주문제작을 상담합니다.

- 주소: ${site.address.region} ${site.address.locality} ${site.address.street}
- 전화: ${site.phone}
- 이메일: ${site.email}
- 영업시간: ${site.hours}

## 페이지
- [첫 화면](${site.origin}/): 회사 소개와 견적 요청
- [온라인 주문](${site.origin}/shop/): 주문 가능한 품목 ${products.length}종
${services.map((s) => `- [${s.title}](${site.origin}/services/${s.slug}/): ${s.summary}`).join('\n')}
- [사용처 안내](${site.origin}/portfolio/): 업종별 위생용지 구성안
- [고객센터](${site.origin}/faq/): 자주 묻는 질문
- [위생용지 저널](${site.origin}/blog/): 선택 기준과 사용 방법
`,
  'utf8'
)

/* ── favicon ────────────────────────────────────────────────────────────────
   (주)이루모스 로고 마크를 옮긴 것입니다. original/images/14-logo.png 의
   짙은 자주 → 주황 그라디언트 라운드 스퀘어 위에, 시트가 풀린 롤 실루엣.
   16px 에서도 형태가 남도록 획을 굵게 잡았습니다. */

await writeFile(
  join(ROOT, 'assets', 'favicon.svg'),
  `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32">
  <defs>
    <linearGradient id="ilumos" x1="2" y1="30" x2="30" y2="2" gradientUnits="userSpaceOnUse">
      <stop offset="0" stop-color="#5c2b33"/>
      <stop offset="0.55" stop-color="#c25a3f"/>
      <stop offset="1" stop-color="#f5a44b"/>
    </linearGradient>
  </defs>
  <rect width="32" height="32" rx="9" fill="url(#ilumos)"/>
  <path d="M13.5 6.6h6.1a3.6 3.6 0 0 1 3.6 3.6v15.2h-4.4V10.2a3.6 3.6 0 0 0-3.6-3.6z" fill="#fff" fill-opacity="0.72"/>
  <rect x="6.8" y="6.6" width="8.4" height="18.8" rx="4.2" fill="#fff"/>
  <ellipse cx="11" cy="11.2" rx="1.7" ry="1.7" fill="#8d3a38"/>
</svg>
`,
  'utf8'
)

await writeFile(join(ROOT, 'assets', 'css', 'starfield.css'), starfieldCss(), 'utf8')

/* ── web app manifest ───────────────────────────────────────────────────── */

await writeFile(
  join(ROOT, 'site.webmanifest'),
  JSON.stringify(
    {
      name: `${site.name} — ${site.title}`,
      short_name: site.nameLatin,
      description: site.description,
      start_url: '/',
      scope: '/',
      display: 'standalone',
      lang: site.lang,
      background_color: '#06060c',
      theme_color: '#06060c',
      icons: [{ src: '/assets/favicon.svg', sizes: 'any', type: 'image/svg+xml', purpose: 'any' }],
    },
    null,
    2
  ),
  'utf8'
)

console.log(`built ${count} pages + sitemap.xml (${routes.filter((r) => !r.skipSitemap).length} urls) + robots.txt + site.webmanifest`)
console.log(`images in manifest: ${Object.keys(manifest).length}`)
