/**
 * pages.mjs — page bodies, assembled purely from components.mjs.
 *
 * Home reproduces the analysed section order (04-pages/home.json > layout_skeleton):
 *   header/nav (fixed) > WebGL hero + scramble keywords -> about/who-we-are
 *   (mono tag + serif lead + counters) -> clients logo wall -> services (tablist)
 *   -> 3D image carousel portfolio -> process -> scaling-video reveal
 *   -> FAQ accordion -> contact/footer
 */

import {
  site, hero, about, clients, servicesMeta, services, portfolioMeta, portfolio,
  process, faqMeta, faq, contact, blogMeta, posts, marquee as marqueeWords,
} from '../data/site.mjs'

import {
  esc, icon, button, eyebrow, media, stat, sectionHead, pageHero, accordion,
  logoWall, workCard, servicesTabs, carousel3D, pager, contactSection, tagList,
  koDate, breadcrumb,
} from './components.mjs'

/* Panels for the 3D cylinder. Many small images rather than a few large ones —
   see components.mjs > carousel3D. */
const CAROUSEL = [
  { image: 'silicon-wafer', caption: '두루마리티슈' },
  { image: 'chip-die', caption: '무지 롤 — 인쇄 전' },
  { image: 'cleanroom-bay', caption: '점보롤 적재' },
  { image: 'circuit-macro', caption: '엠보싱 표면' },
  { image: 'fab-equipment', caption: '크라프트 원지' },
  { image: 'work-etch', caption: '외식업장 주방' },
  { image: 'probe-station', caption: '핸드타월 디스펜서' },
  { image: 'nano-texture', caption: '컬러 인쇄 옵션' },
  { image: 'work-hbm', caption: '카페 · 숙박' },
  { image: 'lab-optics', caption: '주방용 종이호일' },
]

/* Full-bleed gallery strip, used on the home page and the case pages. */
const GALLERY = [
  { image: 'semiconductor-line', ratio: '4x3', alt: '포장 · 생산 라인' },
  { image: 'silicon-ingot', ratio: '1x1', alt: '원지 롤 보관' },
  { image: 'server-rack', ratio: '4x3', alt: '박스 단위 보관 창고' },
  { image: 'circuit-macro', ratio: '1x1', alt: '엠보싱 표면 확대' },
  { image: 'chip-die', ratio: '3x4', alt: '인쇄 전 무지 롤' },
  { image: 'chip-array', ratio: '4x3', alt: '두루마리티슈 묶음' },
  { image: 'nano-texture', ratio: '1x1', alt: '컬러 인쇄 · 색상 옵션' },
  { image: 'lab-optics', ratio: '4x3', alt: '주방용 종이호일' },
]

/* ── HOME ───────────────────────────────────────────────────────────────── */

export function home(manifest) {
  return `
    <!-- hero: WebGL backdrop (see .unicorn__item) + scrambling keywords -->
    <section class="hero theme-dark padding-global" id="top">
      <div class="container hero__inner">
        <div class="hero__eyebrow">${eyebrow(hero.eyebrow)}</div>
        <h2 class="hero__title">
          <span>${esc(hero.words[0])}</span>
          <span class="hero__scramble" data-word-cycle='${JSON.stringify(hero.scrambleWords)}'>${esc(hero.words[1])}</span>
          <span>${esc(hero.words[2])}</span>
        </h2>
        <div class="hero__foot">
          <p class="hero__lead u-lead" data-split>${esc(hero.lead)}</p>
          <div class="hero__foot-right">
            <p class="hero__scroll">${icon.arrowDown}<span>${esc(hero.scrollHint)}</span></p>
            <div class="hero__actions">
              ${button({ label: hero.cta.label, href: hero.cta.href, variant: 'solid', cursor: 'TALK' })}
              ${button({ label: hero.ctaSecondary.label, href: hero.ctaSecondary.href, variant: 'ghost' })}
            </div>
          </div>
        </div>
      </div>
    </section>

    <!-- about: full-bleed sub-hero.
         A bright daylight building fills the block; the copy sits on it in the
         inverted (near-black) ink. The stats panel is the one exception — it
         numbers are filled with the same drifting star field as the headline.
         The still is a <video> poster: drop
         assets/video/about.webm|mp4 in and the block plays footage instead,
         with no markup change. -->
    <section class="section about-hero theme-light" id="about">
      <div class="about-hero__media" aria-hidden="true">
        <video poster="${esc((manifest['about-building'] || {}).src || '/assets/images/about-building.webp')}"
               autoplay muted loop playsinline preload="none">
          <source src="/assets/video/about.webm" type="video/webm">
          <source src="/assets/video/about.mp4" type="video/mp4">
        </video>
        <span class="about-hero__scrim"></span>
      </div>

      <div class="container padding-global about-hero__inner">
        <div class="section-head__eyebrow">${eyebrow(about.eyebrow)}</div>
        <!-- data-reveal, NOT data-split: SplitText rewrites the heading into
             transformed child divs, and background-clip:text cannot clip a
             parent's background to glyphs that live inside a transformed
             descendant — the star fill would vanish. The block reveal gives this
             heading its entrance instead. See base.css > .type-starfield. -->
        <h2 class="about-hero__title type-starfield" data-reveal>${esc(about.heading)}</h2>
        <div class="about-hero__body" data-reveal-group>
          ${about.body.map((p) => `<p>${esc(p)}</p>`).join('')}
        </div>
        <div style="margin-top:var(--space-md)" data-reveal>
          ${button({ label: '사용처 안내 보기', href: '/portfolio/', variant: 'ghost' })}
        </div>
        <div class="about-hero__stats" data-reveal-group>
          ${about.stats.map((s) => stat(s, { starfield: true })).join('')}
        </div>
      </div>
    </section>

    <!-- clients logo wall (GSAP Flip swap loop) -->
    <section class="section padding-global padding-section" id="clients">
      <div class="container">
        ${sectionHead({ tag: clients.eyebrow, title: clients.heading, stack: true })}
        ${logoWall(clients.logos, { manifest })}
      </div>
    </section>

    <!-- services tablist -->
    <section class="section padding-global padding-section" id="services">
      <div class="container">
        ${sectionHead({
          tag: servicesMeta.eyebrow,
          title: servicesMeta.heading,
          aside: `<p class="u-lead" data-split>${esc(servicesMeta.lead)}</p>`,
        })}
        ${servicesTabs(services)}
      </div>
    </section>

    <!-- portfolio: 4 panels on a 3D cylinder, 30s linear infinite -->
    <section class="section padding-global padding-section" id="portfolio">
      <div class="container">
        ${sectionHead({
          tag: portfolioMeta.eyebrow,
          title: portfolioMeta.heading,
          aside: `<p class="u-lead" data-split>${esc(portfolioMeta.lead)}</p>`,
        })}
        ${carousel3D(CAROUSEL, manifest)}
        <div style="margin-top:var(--space-md)" data-reveal>
          ${button({ label: '사용처 전체 보기', href: '/portfolio/', variant: 'ghost' })}
        </div>
      </div>
    </section>

    <!-- process -->
    <section class="section padding-global padding-section" id="process">
      <div class="container">
        ${sectionHead({
          tag: process.eyebrow,
          title: process.heading,
          aside: `<p class="u-lead" data-split>${esc(process.lead)}</p>`,
        })}
        <ol class="process__list">
          ${process.steps
            .map(
              (s) => `<li class="process__item" data-reveal>
              <span class="process__num">${esc(s.num)}</span>
              <h3 class="process__title" lang="en">${esc(s.titleLatin)}<span class="t-ko" lang="ko">${esc(s.title)}</span></h3>
              <p class="process__body">${esc(s.body)}</p>
            </li>`
            )
            .join('')}
        </ol>
      </div>
    </section>

    <!-- gallery strip -->
    <section class="section padding-global padding-section" id="gallery">
      <div class="container">
        ${sectionHead({
          tag: '(In the box)',
          title: 'IN THE\nBOX',
          aside: `<p class="u-lead" data-split>원지에서 박스까지, 매일 쓰는 종이가 만들어지는 자리들.</p>`,
        })}
        <div class="gallery-grid" data-reveal-group>
          ${GALLERY.map((g) => media(g.image, { ratio: g.ratio, alt: g.alt, zoom: true, manifest })).join('')}
        </div>
      </div>
    </section>

    <!-- scroll-scrubbed expanding video -->
    <section class="section" id="reel">
      <div class="scaling-video__wrapper">
        <div class="scaling-video">
          <video poster="/assets/images/hero-poster.webp" autoplay muted loop playsinline preload="none" aria-hidden="true">
            <source src="/assets/video/hero.webm" type="video/webm">
            <source src="/assets/video/hero.mp4" type="video/mp4">
          </video>
          <span class="scaling-video__overlay" aria-hidden="true"></span>
        </div>
      </div>
    </section>

    <!-- FAQ: the live page's single <h1> lives on this block, not the hero -->
    <section class="section padding-global padding-section" id="faq">
      <div class="container">
        <div class="faq__grid">
          <div class="faq__aside">
            <div class="section-head__eyebrow">${eyebrow(faqMeta.eyebrow)}</div>
            <h1 class="faq__h1" data-split>${esc(faqMeta.heading)}</h1>
            <p class="u-dim" style="margin-top:var(--space-sm);max-width:32ch">${esc(faqMeta.lead)}</p>
            <div style="margin-top:var(--space-md)">${button({ label: '전체 FAQ 보기', href: '/faq/', variant: 'ghost' })}</div>
          </div>
          <div>${accordion(faq.slice(0, 5), { idPrefix: 'home-faq' })}</div>
        </div>
      </div>
    </section>

    ${contactSection(contact, site)}
  `
}

/* ── PORTFOLIO INDEX ────────────────────────────────────────────────────── */

export function portfolioIndex(manifest) {
  return `
    ${pageHero({
      crumbs: [{ label: 'HOME', href: '/' }, { label: 'USE CASES' }],
      tag: portfolioMeta.eyebrow,
      title: portfolioMeta.heading,
      lead: portfolioMeta.lead,
      foot: `<span class="eyebrow">${portfolio.length} USE CASES</span>${button({ label: '견적 요청', href: '/#contact', variant: 'ghost' })}`,
    })}

    <section class="section padding-global padding-section">
      <div class="container">
        <div class="work-list">
          ${portfolio
            .map(
              (w, i) => `<article class="work-row">
              <a class="work-row__media" href="/portfolio/${esc(w.slug)}/" data-cursor="VIEW" aria-label="${esc(w.title)} 구성안 자세히 보기">
                ${media(w.image, { ratio: '4x3', alt: `${w.title} — ${w.sector}`, zoom: true, manifest })}
              </a>
              <div class="work-row__body" data-reveal>
                <span class="work-row__latin">(${String(i + 1).padStart(2, '0')})</span>
                <h2 class="work-row__title" lang="en"><a href="/portfolio/${esc(w.slug)}/">${esc(w.titleLatin)}</a><span class="t-ko" lang="ko">${esc(w.title)}</span></h2>
                <div class="meta-row"><span>${esc(w.sector)}</span><span>${esc(w.scale)}</span><span>${esc(w.category)}</span></div>
                <p class="u-dim" style="max-width:48ch">${esc(w.summary)}</p>
                ${tagList(w.tags)}
                <div>${button({ label: '자세히 보기', href: `/portfolio/${w.slug}/`, variant: 'ghost' })}</div>
              </div>
            </article>`
            )
            .join('')}
        </div>
      </div>
    </section>

    ${contactSection(contact, site)}
  `
}

/* ── CASE STUDY ─────────────────────────────────────────────────────────── */

export function workDetail(work, manifest) {
  const i = portfolio.findIndex((w) => w.slug === work.slug)
  const related = work.services.map((slug) => services.find((s) => s.slug === slug)).filter(Boolean)

  return `
    ${pageHero({
      crumbs: [{ label: 'HOME', href: '/' }, { label: 'USE CASES', href: '/portfolio/' }, { label: work.titleLatin }],
      tag: `(${work.category})`,
      title: work.titleLatin,
      subtitle: work.title,
      lead: work.summary,
      foot: `<div class="meta-row"><span>${esc(work.sector)}</span><span>${esc(work.scale)}</span><span>${esc(work.category)}</span></div>${tagList(work.tags)}`,
    })}

    <section class="section padding-global" style="padding-top:var(--space-md)">
      <div class="container">
        ${media(work.image, { ratio: '16x9', alt: `${work.title} 대표 이미지`, priority: true, manifest })}

        <dl class="case__facts" style="margin-top:var(--space-md)" data-reveal-group>
          <div class="case__fact"><dt>적용 업종</dt><dd>${esc(work.sector)}</dd></div>
          <div class="case__fact"><dt>기준 규모</dt><dd>${esc(work.scale)}</dd></div>
          <div class="case__fact"><dt>주요 품목</dt><dd>${esc(work.category)}</dd></div>
          <div class="case__fact"><dt>관련 제품</dt><dd>${esc(related.map((s) => s.title).join(', ') || '위생용지 전 품목')}</dd></div>
        </dl>

        <div class="case__block">
          <h2 class="case__block-title">현장 상황</h2>
          <p class="case__block-body" data-split>${esc(work.challenge)}</p>
        </div>

        <div class="case__block case__block--prose">
          <h2 class="case__block-title">제안 구성</h2>
          <p class="case__block-body" data-reveal>${esc(work.approach)}</p>
        </div>

        <div class="case__block case__block--prose">
          <h2 class="case__block-title">기대 효과</h2>
          <p class="case__block-body" data-reveal>${esc(work.outcome)}</p>
        </div>

        <div class="case__stats" data-reveal-group>
          ${work.stats.map(stat).join('')}
        </div>

        <div class="case__gallery">
          ${media('cleanroom-bay', { ratio: '16x9', alt: `${work.title} — 점보롤 적재`, manifest })}
          ${media('silicon-wafer', { ratio: '3x4', alt: `${work.title} — 두루마리티슈`, zoom: true, manifest })}
          ${media('fab-equipment', { ratio: '4x3', alt: `${work.title} — 원지 입고`, zoom: true, manifest })}
          ${media('probe-station', { ratio: '4x3', alt: `${work.title} — 핸드타월 디스펜서`, zoom: true, manifest })}
          ${media('silicon-ingot', { ratio: '1x1', alt: `${work.title} — 원지 롤 보관`, zoom: true, manifest })}
          ${media('chip-array', { ratio: '4x3', alt: `${work.title} — 두루마리티슈 묶음`, zoom: true, manifest })}
        </div>

        ${related.length
          ? `<section style="padding:var(--space-lg) 0">
              <h2 class="case__block-title" style="margin-bottom:var(--space-sm)">여기에 쓰이는 제품</h2>
              ${servicesTabs(related)}
            </section>`
          : ''}

        <div style="padding-bottom:var(--space-lg)">
          ${pager({
            prev: portfolio[i - 1],
            next: portfolio[i + 1],
            base: '/portfolio/',
          })}
        </div>
      </div>
    </section>

    ${contactSection(contact, site)}
  `
}

/* ── SERVICE DETAIL ─────────────────────────────────────────────────────── */

export function serviceDetail(service, manifest) {
  const i = services.findIndex((s) => s.slug === service.slug)
  const cases = portfolio.filter((w) => w.services.includes(service.slug))

  return `
    ${pageHero({
      crumbs: [{ label: 'HOME', href: '/' }, { label: 'PRODUCTS', href: '/#services' }, { label: service.titleLatin }],
      tag: `(${service.num})`,
      title: service.titleLatin,
      subtitle: service.title,
      lead: service.tagline,
      foot: `<p class="u-dim" style="max-width:52ch">${esc(service.summary)}</p>${button({ label: '이 품목 견적 요청', href: '/#contact', variant: 'ghost' })}`,
    })}

    <section class="section padding-global padding-section">
      <div class="container">
        ${media(service.image, { ratio: '16x9', alt: `${service.title} 제품 이미지`, priority: true, manifest })}

        <div class="service__grid" style="margin-top:var(--space-lg)">
          <div class="service__prose" data-reveal-group>
            ${service.body.map((p) => `<p>${esc(p)}</p>`).join('')}
          </div>

          <aside class="service__panel" data-reveal>
            <div>
              <p class="service__panel-title">취급 규격</p>
              <ul class="service__checklist">
                ${service.deliverables.map((d) => `<li>${esc(d)}</li>`).join('')}
              </ul>
            </div>
            <div>
              <p class="service__panel-title">확인 항목</p>
              ${tagList(service.tools)}
            </div>
            ${button({ label: '견적 요청', href: '/#contact', variant: 'solid', cursor: 'MAIL' })}
          </aside>
        </div>

        ${cases.length
          ? `<section style="padding-top:var(--space-section)">
              ${sectionHead({ tag: '(Where it is used)', title: '이 품목을 쓰는 곳', stack: true })}
              <div class="grid grid--2" data-reveal-group>
                ${cases.map((w, idx) => workCard(w, idx, manifest)).join('')}
              </div>
            </section>`
          : ''}

        <section style="padding-top:var(--space-section)">
          ${sectionHead({ tag: '(All products)', title: '다른 품목', stack: true })}
          <div class="service__nav-list">
            ${services
              .map(
                (s) => `<a class="service__nav-item" href="/services/${esc(s.slug)}/"${s.slug === service.slug ? ' aria-current="page"' : ''}>
                <span lang="en"><span class="service__nav-num">(${esc(s.num)})</span> ${esc(s.titleLatin)}<span class="t-ko" lang="ko">${esc(s.title)}</span></span>
                <span aria-hidden="true">${icon.arrow}</span>
              </a>`
              )
              .join('')}
          </div>
        </section>

        <div style="padding-top:var(--space-lg)">
          ${pager({ prev: services[i - 1], next: services[i + 1], base: '/services/' })}
        </div>
      </div>
    </section>

    ${contactSection(contact, site)}
  `
}

/* ── FAQ PAGE ───────────────────────────────────────────────────────────── */

export function faqPage() {
  return `
    ${pageHero({
      crumbs: [{ label: 'HOME', href: '/' }, { label: 'FAQ' }],
      tag: faqMeta.eyebrow,
      title: faqMeta.heading,
      lead: faqMeta.lead,
    })}

    <section class="section padding-global padding-section">
      <div class="container">
        <div class="faq__grid">
          <div class="faq__aside">
            <h2 class="section-head__title" style="font-size:var(--t-h4)" data-split>${'무엇이든\n물어보세요'}</h2>
            <p class="u-dim" style="margin-top:var(--space-sm);max-width:30ch">답이 여기 없다면 전화나 이메일로 주세요. 보통 하루 안에 회신드립니다.</p>
            <div style="margin-top:var(--space-md)">${button({ label: site.email, href: `mailto:${site.email}`, variant: 'ghost', cursor: 'MAIL' })}</div>
          </div>
          <div>${accordion(faq, { idPrefix: 'faq' })}</div>
        </div>
      </div>
    </section>

    ${contactSection(contact, site)}
  `
}

/* ── JOURNAL ────────────────────────────────────────────────────────────── */

export function blogIndex() {
  return `
    ${pageHero({
      crumbs: [{ label: 'HOME', href: '/' }, { label: 'JOURNAL' }],
      tag: blogMeta.eyebrow,
      title: blogMeta.heading,
      lead: blogMeta.lead,
      foot: `<span class="eyebrow">${posts.length} ENTRIES</span>`,
    })}

    <section class="section padding-global padding-section">
      <div class="container">
        <div class="post-list">
          ${posts
            .map(
              (p) => `<article class="post-row" data-reveal>
              <time class="post-row__date" datetime="${esc(p.date)}">${esc(koDate(p.date))}</time>
              <div>
                <span class="post-row__cat">${esc(p.category)}</span>
                <h2 class="post-row__title" lang="en"><a href="/blog/${esc(p.slug)}/" data-cursor="READ">${esc(p.titleLatin)}</a><span class="t-ko" lang="ko">${esc(p.title)}</span></h2>
                <p class="post-row__excerpt">${esc(p.excerpt)}</p>
              </div>
              <span class="post-row__time">${esc(p.readingTime)}</span>
            </article>`
            )
            .join('')}
        </div>
      </div>
    </section>

    ${contactSection(contact, site)}
  `
}

/* ── JOURNAL POST ───────────────────────────────────────────────────────── */

export function blogPost(post, manifest) {
  const i = posts.findIndex((p) => p.slug === post.slug)
  // Reading order runs newest-first, so "previous" is the newer entry.
  const newer = posts[i - 1]
  const older = posts[i + 1]

  return `
    ${pageHero({
      crumbs: [{ label: 'HOME', href: '/' }, { label: 'JOURNAL', href: '/blog/' }, { label: post.titleLatin }],
      tag: `(${post.category})`,
      title: post.titleLatin,
      subtitle: post.title,
      lead: post.excerpt,
      foot: `<div class="meta-row"><time datetime="${esc(post.date)}">${esc(koDate(post.date))}</time><span>${esc(post.category)}</span><span>읽는 데 ${esc(post.readingTime)}</span></div>`,
    })}

    <section class="section padding-global padding-section">
      <div class="container">
        ${media(post.image, { ratio: '16x9', alt: `${post.title} 관련 이미지`, priority: true, manifest })}

        <article class="prose" style="margin-top:var(--space-lg)">
          ${post.body
            .map(
              (block) => `<section class="prose__block" data-reveal>
              <h2 class="prose__h">${esc(block.h)}</h2>
              ${block.p.map((p) => `<p>${esc(p)}</p>`).join('')}
            </section>`
            )
            .join('')}
        </article>

        <aside class="prose__cta" data-reveal>
          <p class="u-dim">읽으시다 지금 쓰시는 곳이 떠오르셨다면, 사용처와 인원만 알려 주세요.</p>
          ${button({ label: '견적 요청', href: '/#contact', variant: 'solid', cursor: 'MAIL' })}
        </aside>

        <div style="padding-top:var(--space-lg)">
          ${pager({ prev: newer, next: older, base: '/blog/' })}
        </div>
      </div>
    </section>

    ${contactSection(contact, site)}
  `
}

/* ── LEGAL ──────────────────────────────────────────────────────────────── */

export function legalPage(doc) {
  return `
    ${pageHero({
      crumbs: [{ label: 'HOME', href: '/' }, { label: doc.titleLatin }],
      tag: '(Legal)',
      title: doc.titleLatin,
      subtitle: doc.title,
      foot: `<span class="eyebrow">최종 개정 ${esc(koDate(doc.updated))}</span>`,
    })}

    <section class="section padding-global padding-section">
      <div class="container legal">
        <p class="legal__notice">
          이 문서는 원본 사이트에 해당 문서가 없어 리뉴얼 시안용으로 작성한 표준 문안입니다.
          실제 서비스에 사용하기 전 법률 검토가 필요합니다. (MISSING.md 참고)
        </p>
        ${doc.sections
          .map(
            (s) => `<section class="legal__section" data-reveal>
            <h2>${esc(s.h)}</h2>
            ${s.p.map((p) => `<p>${esc(p)}</p>`).join('')}
          </section>`
          )
          .join('')}
      </div>
    </section>
  `
}

/* ── 404 ────────────────────────────────────────────────────────────────── */

export function notFound() {
  return `
    <section class="notfound padding-global">
      <div class="container">
        <div class="section-head__eyebrow" style="justify-content:center;display:flex">${eyebrow('(404)')}</div>
        <h1 class="notfound__code">LOST</h1>
        <p class="u-lead" style="margin-top:var(--space-sm)">찾으시는 페이지가 없습니다.</p>
        <div style="margin-top:var(--space-md);display:flex;gap:var(--space-2xs);justify-content:center">
          ${button({ label: '홈으로', href: '/', variant: 'solid' })}
          ${button({ label: '제품 보기', href: '/#services', variant: 'ghost' })}
        </div>
      </div>
    </section>
  `
}
