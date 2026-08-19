/**
 * layout.mjs — the document shell every page is poured into.
 *
 * SEO shape mirrors 01-discovery.json > seo: title / description / canonical /
 * OG / JSON-LD. Schema type is Organization (위생용지 제조·공급, 대한민국);
 * the analysed original used CreativeWorkAgency — see MISSING.md > C7.
 */

import { site, nav, marquee as marqueeWords } from '../data/site.mjs'
import { esc, icon, marquee } from './components.mjs'

/* ── head ───────────────────────────────────────────────────────────────── */

function fonts() {
  // Same delivery mechanism the source uses (Google Fonts). IBM Plex ONLY —
  // Sans KR, Sans, Mono, Serif. No other foundry is requested anywhere on the
  // site (Noto Serif KR used to cover Hangul in the serif role and was removed
  // so the stack is entirely IBM). See tokens.css > type families.
  return `<link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+KR:wght@300;400;500;600;700&family=IBM+Plex+Sans:ital,wght@0,300;0,400;0,500;0,600;0,700;1,400&family=IBM+Plex+Mono:wght@300;400;500;600&family=IBM+Plex+Serif:ital,wght@0,400;0,600;1,400&display=swap">`
}

function jsonLd(page) {
  const graph = [
    {
      '@type': 'Organization',
      '@id': `${site.origin}/#organization`,
      name: site.name,
      url: `${site.origin}/`,
      email: site.email,
      telephone: site.phone,
      faxNumber: site.fax,
      founder: { '@type': 'Person', name: site.ceo },
      description: site.description,
      address: {
        '@type': 'PostalAddress',
        streetAddress: site.address.street,
        addressLocality: site.address.locality,
        addressRegion: site.address.region,
        addressCountry: site.address.country,
      },
      /* 원본 사이트에 SNS 채널 표기가 없어 sameAs 는 비워 둡니다 — 없는 계정을
         만들어 넣지 않습니다. */
      ...(site.social.length ? { sameAs: site.social.map((s) => s.href) } : {}),
    },
    {
      '@type': 'WebSite',
      '@id': `${site.origin}/#website`,
      url: `${site.origin}/`,
      name: site.name,
      inLanguage: site.lang,
      publisher: { '@id': `${site.origin}/#organization` },
    },
    {
      '@type': 'WebPage',
      '@id': `${site.origin}${page.path}#webpage`,
      url: `${site.origin}${page.path}`,
      name: page.title,
      description: page.description,
      isPartOf: { '@id': `${site.origin}/#website` },
      inLanguage: site.lang,
    },
  ]

  if (page.jsonLdExtra) graph.push(...page.jsonLdExtra)

  return `<script type="application/ld+json">${JSON.stringify({ '@context': 'https://schema.org', '@graph': graph })}</script>`
}

function head(page) {
  const url = `${site.origin}${page.path}`
  const title = page.path === '/' ? page.title : `${page.title} | ${site.name}`
  return `<meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>${esc(title)}</title>
  <meta name="description" content="${esc(page.description)}">
  <link rel="canonical" href="${esc(url)}">
  <meta name="theme-color" content="#000000">
  <meta name="robots" content="index, follow">

  <meta property="og:type" content="${page.ogType || 'website'}">
  <meta property="og:site_name" content="${esc(site.name)}">
  <meta property="og:locale" content="${esc(site.locale)}">
  <meta property="og:title" content="${esc(title)}">
  <meta property="og:description" content="${esc(page.description)}">
  <meta property="og:url" content="${esc(url)}">
  <meta property="og:image" content="${esc(site.origin + site.ogImage)}">
  <meta name="twitter:card" content="summary_large_image">
  <meta name="twitter:title" content="${esc(title)}">
  <meta name="twitter:description" content="${esc(page.description)}">
  <meta name="twitter:image" content="${esc(site.origin + site.ogImage)}">

  <link rel="icon" href="/assets/favicon.svg" type="image/svg+xml">
  <link rel="apple-touch-icon" href="/assets/favicon.svg">
  <link rel="manifest" href="/site.webmanifest">
  ${fonts()}
  <link rel="stylesheet" href="/assets/css/lenis.css">
  <link rel="stylesheet" href="/assets/css/tokens.css">
  <link rel="stylesheet" href="/assets/css/base.css">
  <link rel="stylesheet" href="/assets/css/starfield.css">
  <link rel="stylesheet" href="/assets/css/components.css">
  <link rel="stylesheet" href="/assets/css/sections.css">
  <link rel="stylesheet" href="/assets/css/pages.css">
  ${jsonLd(page)}`
}

/* ── WebGL backdrop layer ───────────────────────────────────────────────── */

/* `.unicorn__scene` is the photographic ground the star field sits on — a dense
   stand of pines. It is painted BELOW the canvas (z-index:-1 inside the layer's
   own stacking context), so every star, the wireframe star and the pointer light
   all draw on top of it. Only the home hero switches it on (`body.is-home`), and
   it dissolves across the hero pin — see motion.js. */
function webglLayer() {
  return `<div class="unicorn__item" aria-hidden="true">
    <div class="unicorn__scene"></div>
    <div class="unicorn__fallback"></div>
    <div class="unicorn__grain"></div>
  </div>`
}

/* ── header + overlay menu ──────────────────────────────────────────────── */

function brand() {
  return `<a class="brand" href="/" aria-label="${esc(site.name)} 홈">
      <span>${esc(site.nameLatin)}</span><span class="brand__dot" aria-hidden="true"></span>
    </a>`
}

/* The analysed original carried an EN/ES switcher ([data-lang-hover]). This
   build ships Korean only, so the control was removed rather than left pointing
   at a route that does not exist. See MISSING.md > C9. */

function header() {
  return `<header class="header theme-dark padding-global" data-theme-watch>
    <div class="container header__inner">
      <div class="header__bar">
        ${brand()}
        <nav class="nav-desktop" aria-label="주요">
          ${nav.map((n) => `<a class="nav-desktop__link" href="${esc(n.href)}" data-underline-link>${esc(n.label)}</a>`).join('')}
        </nav>
        <div class="header__actions">
          <button class="menu-toggle" type="button" aria-expanded="false" aria-controls="menu-overlay">
            <span class="menu-toggle__bars" aria-hidden="true"><span></span><span></span><span></span></span>
            <span>MENU</span>
          </button>
        </div>
      </div>
    </div>
  </header>

  <div class="menu-overlay theme-dark padding-global" id="menu-overlay" role="dialog" aria-modal="true" aria-label="사이트 메뉴">
    <div class="container menu-overlay__head">
      ${brand()}
      <span class="eyebrow">(Menu)</span>
    </div>
    <div class="container menu-overlay__body">
      <nav class="menu-overlay__list" aria-label="전체">
        <ul>
          ${nav
            .map(
              (n, i) => `<li class="menu-overlay__item">
              <a class="menu-overlay__link" href="${esc(n.href)}" data-cursor="GO">
                <span class="menu-overlay__num">0${i + 1}</span><span>${esc(n.label)}</span>
              </a>
            </li>`
            )
            .join('')}
        </ul>
      </nav>
    </div>
    <div class="container menu-overlay__foot">
      <a href="mailto:${esc(site.email)}" data-underline-link>${esc(site.email)}</a>
      <div class="meta-row">
        <a href="tel:${esc(site.phone.replace(/[^0-9+]/g, ''))}" data-underline-link>TEL ${esc(site.phone)}</a>
        <span>FAX ${esc(site.fax)}</span>
        ${site.social.map((s) => `<a href="${esc(s.href)}" target="_blank" rel="noopener noreferrer" data-underline-link>${esc(s.label)}</a>`).join('')}
      </div>
    </div>
  </div>`
}

/* ── footer ─────────────────────────────────────────────────────────────── */

function footer(services, portfolio) {
  return `${marquee(marqueeWords)}
  <footer class="footer theme-dark padding-global">
    <div class="container">
      <div class="footer__top">
        <div>
          <p class="footer__col-title">${esc(site.name)}</p>
          <p class="u-dim" style="max-width:34ch">${esc(site.description)}</p>
          <p class="u-dim" style="margin-top:var(--space-sm)">대표자 ${esc(site.ceo)}<br>${esc(site.address.region)} ${esc(site.address.locality)} ${esc(site.address.street)}<br>TEL ${esc(site.phone)} · FAX ${esc(site.fax)}</p>
          <p style="margin-top:var(--space-sm)"><a href="mailto:${esc(site.email)}" data-underline-link>${esc(site.email)}</a></p>
        </div>
        <div>
          <p class="footer__col-title">제품소개</p>
          <div class="footer__links">
            ${services.slice(0, 6).map((s) => `<a href="/services/${esc(s.slug)}/">${esc(s.title)}</a>`).join('')}
          </div>
        </div>
        <div>
          <p class="footer__col-title">바로가기</p>
          <div class="footer__links">
            <a href="/portfolio/">사용처 안내</a>
            <a href="/faq/">고객센터</a>
            <a href="/blog/">위생용지 저널</a>
            <a href="/#contact">견적 요청</a>
            ${site.social.map((s) => `<a href="${esc(s.href)}" target="_blank" rel="noopener noreferrer">${esc(s.label)}</a>`).join('')}
          </div>
        </div>
      </div>
      <p class="footer__wordmark" aria-hidden="true">${esc(site.nameLatin)}</p>
      <div class="footer__bottom">
        <span>© <span data-year>2026</span> ${esc(site.name)} — ${esc(site.address.region)} ${esc(site.address.locality)}</span>
        <span class="footer__links" style="flex-direction:row;gap:var(--space-sm)">
          <a href="/utility-pages/privacy-notice/">개인정보 처리방침</a>
          <a href="/utility-pages/terms-conditions/">이용약관</a>
        </span>
      </div>
    </div>
  </footer>`
}

/* ── cursor nodes ───────────────────────────────────────────────────────── */

function cursor() {
  // Two cursors, swapped by theme: the difference-blended ring over the dark
  // hero/footer, a small blue radial star over the light body — a scaled-down
  // echo of the hero's star-and-light backdrop.
  return `<div class="cursor-dot" aria-hidden="true"></div>
  <div class="cursor-ring" aria-hidden="true"></div>
  <div class="cursor-star" aria-hidden="true">
    <span class="cursor-star__glow"></span>
    <svg class="cursor-star__mark" viewBox="0 0 24 24" focusable="false">
      <path d="M12 0 C12.6 6.5 17.5 11.4 24 12 C17.5 12.6 12.6 17.5 12 24 C11.4 17.5 6.5 12.6 0 12 C6.5 11.4 11.4 6.5 12 0 Z"/>
    </svg>
  </div>`
}

/* ── vendor scripts (local, no external hosts) ──────────────────────────── */

function scripts() {
  const vendor = [
    'lenis.min.js',
    'gsap.min.js',
    'ScrollTrigger.min.js',
    'SplitText.min.js',
    'ScrambleTextPlugin.min.js',
    'Flip.min.js',
    'Observer.min.js',
    'InertiaPlugin.min.js',
    'CustomEase.min.js',
  ]
  return `${vendor.map((f) => `<script src="/assets/js/vendor/${f}" defer></script>`).join('\n  ')}
  <script type="module" src="/assets/js/main.js"></script>`
}

/* ── document ───────────────────────────────────────────────────────────── */

export function layout(page, { services, portfolio }) {
  return `<!DOCTYPE html>
<html lang="${site.lang}">
<head>
  ${head(page)}
</head>
<body${page.bodyClass ? ` class="${page.bodyClass}"` : ''}>
  <a class="skip-link" href="#main">본문으로 건너뛰기</a>
  ${webglLayer()}
  ${cursor()}
  ${header()}
  <main class="page-main" id="main">
${page.body}
  </main>
  ${footer(services, portfolio)}
  ${scripts()}
</body>
</html>
`
}
