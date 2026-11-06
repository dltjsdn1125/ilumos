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
  shopMeta, shopCategories, products,
} from '../data/site.mjs'

import {
  esc, icon, button, eyebrow, media, stat, sectionHead, pageHero, accordion,
  logoWall, workCard, servicesTabs, carousel3D, pager, contactSection, tagList,
  koDate, breadcrumb,
  won, priceBlock, qtyStepper, productCard, shopNotice, shopToolbar,
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
        <h1 class="hero__title">
          <span>${esc(hero.words[0])}</span>
          <span class="hero__scramble" data-word-cycle='${JSON.stringify(hero.scrambleWords)}'>${esc(hero.words[1])}</span>
          <span>${esc(hero.words[2])}</span>
        </h1>
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

    <!-- FAQ: 홈의 대제목(h1)은 히어로입니다. 이 블록 제목은 h2(글자 모양은 .faq__h1 그대로). -->
    <section class="section padding-global padding-section" id="faq">
      <div class="container">
        <div class="faq__grid">
          <div class="faq__aside">
            <div class="section-head__eyebrow">${eyebrow(faqMeta.eyebrow)}</div>
            <h2 class="faq__h1" data-split>${esc(faqMeta.heading)}</h2>
            <p class="u-dim" style="margin-top:var(--space-sm);max-width:32ch">${esc(faqMeta.lead)}</p>
            <div style="margin-top:var(--space-md)">${button({ label: '전체 FAQ 보기', href: '/faq/', variant: 'ghost' })}</div>
          </div>
          <div>${accordion(faq.slice(0, 5), { idPrefix: 'home-faq' })}</div>
        </div>
      </div>
    </section>

    <!-- 온라인 주문 품목 바로가기: 품목 상세 페이지 20곳이 첫 화면에서 링크 한 번으로 닿게 하는 글 링크 목록 -->
    <section class="section padding-global padding-section shop-links" id="shop-links" aria-labelledby="shop-links-title">
      <div class="container">
        <h2 class="shop-links__title" id="shop-links-title">온라인 주문 품목 ${products.length}종</h2>
        <div class="shop-links__grid">
          ${shopCategories
            .map((c) => {
              const items = products.filter((p) => p.category === c.slug)
              if (!items.length) return ''
              return `<div class="shop-links__group">
            <h3 class="shop-links__group-title"><a href="/services/${esc(c.slug)}/">${esc(c.label)}</a></h3>
            <ul>${items.map((p) => `<li><a href="/shop/${esc(p.slug)}/">${esc(p.name)}</a></li>`).join('')}</ul>
          </div>`
            })
            .join('')}
        </div>
        <p class="shop-links__note"><a href="/shop/">온라인 주문 전체 보기</a> · 영업시간 ${esc(site.hours)} · ${Number(site.lastmod.slice(0, 4))}년 ${Number(site.lastmod.slice(5, 7))}월 기준</p>
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

/* ── SHOP INDEX ─────────────────────────────────────────────────────────────
   자사몰 카탈로그. 20종을 전부 서버에서 그려 두고, 분류·정렬은 이미 있는 카드를
   숨기거나 순서만 바꾸는 방식으로 동작합니다(shop.js). 자바스크립트가 꺼져 있어도
   제품 20종과 가격이 그대로 보이고 상세 페이지로도 들어갈 수 있습니다. */

export function shopIndex(manifest) {
  const catOf = (slug) => shopCategories.find((c) => c.slug === slug)

  return `
    ${pageHero({
      crumbs: [{ label: 'HOME', href: '/' }, { label: 'SHOP' }],
      tag: shopMeta.eyebrow,
      title: shopMeta.heading,
      subtitle: '온라인 주문',
      lead: shopMeta.lead,
      foot: `<span class="eyebrow">${products.length} PRODUCTS</span>${button({ label: '장바구니 보기', href: '/shop/cart/', variant: 'ghost' })}`,
    })}

    <section class="section padding-global padding-section">
      <div class="container">
        ${shopNotice(shopMeta.notice)}
        ${shopToolbar(shopCategories, products)}
        <p class="shop-count" data-shop-count role="status" aria-live="polite">전체 ${products.length}개 제품</p>

        <div class="product-grid" data-product-grid data-reveal-group>
          ${products.map((p, i) => productCard(p, catOf(p.category), i, manifest)).join('')}
        </div>

        <p class="shop-empty" data-shop-empty hidden>해당 분류에 제품이 없습니다.</p>

        <div class="shop-info" data-reveal-group>
          <div class="shop-info__item">
            <p class="shop-info__title" lang="en">SHIPPING<span class="t-ko" lang="ko">배송</span></p>
            <p class="shop-info__body">전국 배송합니다. ${won(shopMeta.shipping.freeOver)}원 이상 주문은 배송비가 없고, 그 아래는 ${won(shopMeta.shipping.flat)}원이 붙습니다. 지역과 수량에 따라 화물 배송으로 전환될 수 있습니다.</p>
          </div>
          <div class="shop-info__item">
            <p class="shop-info__title" lang="en">PAYMENT<span class="t-ko" lang="ko">결제</span></p>
            <p class="shop-info__body">주문서를 보내 주시면 확정 견적과 함께 입금 계좌를 안내드립니다. 사업자는 세금계산서 발행이 가능합니다.</p>
          </div>
          <div class="shop-info__item">
            <p class="shop-info__title" lang="en">BULK ORDER<span class="t-ko" lang="ko">대량 주문</span></p>
            <p class="shop-info__body">박스 단위가 여러 개거나 규격이 다른 경우, 낱개로 담기보다 견적으로 진행하시는 편이 단가가 낮습니다.</p>
            ${button({ label: '견적 요청', href: '/#contact', variant: 'ghost' })}
          </div>
        </div>
      </div>
    </section>

    ${contactSection(contact, site)}
  `
}

/* ── PRODUCT DETAIL ─────────────────────────────────────────────────────── */

export function productDetail(product, manifest) {
  const i = products.findIndex((p) => p.slug === product.slug)
  const cat = shopCategories.find((c) => c.slug === product.category)
  const service = services.find((s) => s.slug === product.category)
  const siblings = products.filter((p) => p.category === product.category && p.slug !== product.slug)

  return `
    ${pageHero({
      crumbs: [{ label: 'HOME', href: '/' }, { label: 'SHOP', href: '/shop/' }, { label: product.nameLatin }],
      tag: `(${cat.labelEn})`,
      title: product.nameLatin,
      subtitle: product.name,
      lead: product.summary,
      foot: `<div class="meta-row"><span>${esc(product.sku)}</span><span>${esc(product.pack)}</span><span>${esc(cat.label)}</span></div>${tagList(product.tags)}`,
    })}

    <section class="section padding-global padding-section">
      <div class="container">
        <div class="product__layout">
          <div class="product__media">
            ${media(product.image, { ratio: '4x3', alt: `${product.name} 제품 이미지`, priority: true, manifest })}
            ${product.badge ? `<span class="product-card__badge">${esc(product.badge)}</span>` : ''}
          </div>

          <div class="product__panel" data-reveal>
            <p class="product__panel-cat" lang="en">${esc(cat.labelEn)}</p>
            <h2 class="product__panel-name">${esc(product.name)}</h2>
            ${priceBlock(product, { className: 'price--lg' })}
            <p class="product__panel-pack">${esc(product.pack)} · 최소 1${esc(product.unit)}부터</p>

            ${product.externalUrl
              /* 외부 스토어에 등록된 품목 — 수량·결제는 그쪽 상세에서 고르므로
                 스테퍼와 장바구니를 내지 않습니다. 대량 견적 경로는 그대로 둡니다. */
              ? `<div class="product__buy">
              ${button({ label: '구매하기', href: product.externalUrl, variant: 'solid', size: 'lg', cursor: 'BUY', attrs: ' target="_blank" rel="noopener noreferrer"' })}
            </div>
            <p class="product__status product__status--note" role="note">온라인 스토어의 상품 페이지가 새 창으로 열립니다.</p>

            <div class="product__links">
              ${button({ label: '수량 견적 문의', href: '/#contact', variant: 'ghost', cursor: 'MAIL' })}
            </div>`
              : `<div class="product__buy">
              ${qtyStepper({ id: `qty-${product.slug}`, label: `${product.name} 수량` })}
              ${button({ label: '장바구니에 담기', variant: 'solid', size: 'lg', cursor: 'ADD', attrs: ` data-add-to-cart="${esc(product.slug)}" data-qty-source` })}
            </div>
            <p class="product__status" role="status" aria-live="polite" data-add-status></p>

            <div class="product__links">
              ${button({ label: '장바구니 보기', href: '/shop/cart/', variant: 'ghost' })}
              ${button({ label: '수량 견적 문의', href: '/#contact', variant: 'ghost', cursor: 'MAIL' })}
            </div>`}

            <ul class="product__notes">
              <li>표시 가격은 시안용 예시입니다. 실제 단가는 수량과 납품 조건에 따라 달라집니다.</li>
              <li>${esc(won(shopMeta.shipping.freeOver))}원 이상 무료배송 · 미만 ${esc(won(shopMeta.shipping.flat))}원</li>
              <li>평일 오후 2시 이전 주문 건은 당일 출고를 기준으로 합니다.</li>
            </ul>
          </div>
        </div>

        <div class="product__detail">
          <div>
            <h2 class="case__block-title">제품 규격</h2>
            <dl class="product__spec-list" data-reveal-group>
              ${product.specs.map((s) => `<div class="product__spec"><dt>${esc(s.k)}</dt><dd>${esc(s.v)}</dd></div>`).join('')}
              <div class="product__spec"><dt>품목 코드</dt><dd>${esc(product.sku)}</dd></div>
            </dl>
          </div>

          ${service
            ? `<aside class="product__about" data-reveal>
                <h2 class="case__block-title">이 품목에 대해</h2>
                <p class="product__about-body">${esc(service.body[0])}</p>
                ${button({ label: `${service.title} 자세히 보기`, href: `/services/${service.slug}/`, variant: 'ghost' })}
              </aside>`
            : ''}
        </div>

        ${siblings.length
          ? `<section style="padding-top:var(--space-section)">
              ${sectionHead({ tag: `(${cat.labelEn})`, title: '같은 품목의 다른 규격', stack: true })}
              <div class="product-grid product-grid--related" data-reveal-group>
                ${siblings.map((p, idx) => productCard(p, cat, idx, manifest)).join('')}
              </div>
            </section>`
          : ''}

        <div style="padding-top:var(--space-lg)">
          ${pager({
            prev: products[i - 1] ? { slug: products[i - 1].slug, title: products[i - 1].name, titleLatin: products[i - 1].nameLatin } : undefined,
            next: products[i + 1] ? { slug: products[i + 1].slug, title: products[i + 1].name, titleLatin: products[i + 1].nameLatin } : undefined,
            base: '/shop/',
          })}
        </div>
      </div>
    </section>

    ${contactSection(contact, site)}
  `
}

/* ── CART / ORDER ───────────────────────────────────────────────────────────
   담긴 내용은 브라우저(localStorage)에만 있습니다. 그래서 이 페이지의 목록은
   서버에서 그릴 수 없고 shop.js 가 채웁니다 — 대신 카탈로그를 페이지 안에
   JSON 으로 실어 보내, 담을 때 저장한 값이 아니라 **현재 카탈로그의 값**으로
   이름과 가격을 다시 읽습니다. 오래된 장바구니가 옛 가격을 들고 있지 않습니다.

   주문서는 결제하지 않습니다. 문의 폼과 같은 방식으로 메일 초안을 만들어
   넘깁니다 — 정적 호스팅에서 실제로 동작하고, 접수되지 않은 주문을 접수됐다고
   말하지 않는 유일한 방법입니다. See MISSING.md > D7. */

export function cartPage(manifest) {
  const catalogue = products.map((p) => {
    const m = manifest[p.image] || {}
    const cat = shopCategories.find((c) => c.slug === p.category)
    return {
      slug: p.slug,
      sku: p.sku,
      name: p.name,
      nameLatin: p.nameLatin,
      price: p.price,
      unit: p.unit,
      pack: p.pack,
      category: cat ? cat.label : '',
      src: m.src || `/assets/images/${p.image}.webp`,
      w: m.width || 1200,
      h: m.height || 900,
    }
  })

  const field = ({ name, label, type = 'text', required = false, autocomplete = '', placeholder = '' }) => `
        <p class="field">
          <label class="field__label" for="of-${name}">${esc(label)}${required ? '<span class="field__req" aria-hidden="true">*</span>' : ''}</label>
          <input class="field__input" id="of-${name}" name="${name}" type="${type}"
                 ${required ? 'required' : ''}${autocomplete ? ` autocomplete="${autocomplete}"` : ''}${placeholder ? ` placeholder="${esc(placeholder)}"` : ''}
                 aria-describedby="of-${name}-err">
          <span class="field__err" id="of-${name}-err" data-error-for="${name}" role="alert"></span>
        </p>`

  return `
    ${pageHero({
      crumbs: [{ label: 'HOME', href: '/' }, { label: 'SHOP', href: '/shop/' }, { label: 'CART' }],
      tag: '(Cart & order)',
      title: 'CART',
      subtitle: '장바구니 · 주문서',
      lead: '담으신 품목을 확인하고 주문서를 보내 주시면, 확정 견적과 입금 안내로 회신드립니다.',
      foot: `${button({ label: '제품 더 보기', href: '/shop/', variant: 'ghost' })}`,
    })}

    <section class="section padding-global padding-section">
      <div class="container">
        ${shopNotice(shopMeta.notice)}

        <div class="cart" data-cart>
          <div class="cart__main">
            <div class="cart__list" data-cart-list></div>

            <div class="cart__empty" data-cart-empty>
              <p class="u-lead">장바구니가 비어 있습니다.</p>
              <p class="u-dim" style="margin-top:var(--space-2xs)">담으신 품목이 이 브라우저에만 저장되기 때문에, 다른 기기에서 담은 내용은 보이지 않습니다.</p>
              <div style="margin-top:var(--space-md)">${button({ label: '제품 보러 가기', href: '/shop/', variant: 'solid' })}</div>
            </div>
          </div>

          <aside class="cart__summary" data-cart-summary hidden>
            <p class="cart__summary-title" lang="en">ORDER SUMMARY<span class="t-ko" lang="ko">주문 요약</span></p>
            <dl class="cart__sums">
              <div class="cart__sum"><dt>상품 합계</dt><dd data-sum-goods>0원</dd></div>
              <div class="cart__sum"><dt>배송비</dt><dd data-sum-shipping>0원</dd></div>
              <div class="cart__sum cart__sum--total"><dt>합계</dt><dd data-sum-total>0원</dd></div>
            </dl>
            <p class="cart__summary-note">${esc(won(shopMeta.shipping.freeOver))}원 이상 주문은 배송비가 없습니다. 부가세 포함 여부와 화물 배송 여부는 확정 견적에서 안내드립니다.</p>
            <button class="cart__clear" type="button" data-cart-clear>장바구니 비우기</button>
          </aside>
        </div>

        <section class="order" data-order-section hidden>
          ${sectionHead({ tag: '(Order form)', title: '주문서 보내기', stack: true })}
          <p class="u-dim" style="max-width:60ch;margin-bottom:var(--space-md)">
            이 화면에서는 결제가 이루어지지 않습니다. 아래 내용을 담은 메일 초안이 열리며, 보내 주시면 확정 견적과 입금 계좌로 회신드립니다.
          </p>

          <form class="order-form" id="order-form" novalidate data-mailto="${esc(site.email)}">
            <div class="contact-form__row">
              ${field({ name: 'company', label: '상호 · 사용처', required: true, autocomplete: 'organization' })}
              ${field({ name: 'name', label: '담당자 성함', required: true, autocomplete: 'name' })}
            </div>
            <div class="contact-form__row">
              ${field({ name: 'phone', label: '연락처', type: 'tel', required: true, autocomplete: 'tel' })}
              ${field({ name: 'email', label: '이메일', type: 'email', required: true, autocomplete: 'email' })}
            </div>
            ${field({ name: 'address', label: '배송지 주소', required: true, autocomplete: 'street-address', placeholder: '건물명 · 층 · 하차 조건까지 적어 주시면 좋습니다' })}

            <p class="field">
              <label class="field__label" for="of-note">요청 사항</label>
              <textarea class="field__input field__input--area" id="of-note" name="note" rows="4"
                        placeholder="희망 납품일, 세금계산서 발행 여부, 하차 조건 등 알려 주실 내용이 있으면 적어 주세요."></textarea>
            </p>

            <p class="field field--check">
              <input class="field__check" id="of-consent" name="consent" type="checkbox" required
                     aria-describedby="of-consent-err">
              <label for="of-consent">
                <a href="/utility-pages/privacy-notice/" data-underline-link>개인정보 처리방침</a>에 따른 수집·이용에 동의합니다.<span class="field__req" aria-hidden="true">*</span>
              </label>
              <span class="field__err" id="of-consent-err" data-error-for="consent" role="alert"></span>
            </p>

            <div class="contact-form__actions">
              ${button({ label: '주문서 보내기', variant: 'solid', size: 'lg', cursor: 'SEND', attrs: ' type="submit"' })}
              <p class="contact-form__status" role="status" aria-live="polite"></p>
            </div>
          </form>
        </section>
      </div>
    </section>

    <!-- 장바구니 한 줄의 마크업. shop.js 가 이것을 복제해 값만 채웁니다 —
         마크업은 템플릿 층이 갖고 자바스크립트는 동작만 갖는다는 이 프로젝트의
         규칙(components.mjs 머리말)을 장바구니에서도 지키기 위해서입니다. -->
    <template id="cart-row-template">
      <article class="cart-row" data-cart-row>
        <div class="cart-row__media">
          <img src="/assets/images/og-cover.webp" alt="" width="1200" height="900" loading="lazy" decoding="async">
        </div>
        <div class="cart-row__body">
          <p class="cart-row__cat"></p>
          <h3 class="cart-row__title"><a href="/shop/"></a></h3>
          <p class="cart-row__meta"></p>
        </div>
        <div class="cart-row__ctrl">
          ${qtyStepper({ id: 'cart-qty-template', label: '수량' })}
          <p class="cart-row__price"></p>
          <button class="cart-row__remove" type="button" data-cart-remove>삭제</button>
        </div>
      </article>
    </template>

    <script type="application/json" id="shop-catalogue">${JSON.stringify(catalogue).replace(/</g, '\\u003c')}</script>
    <script type="application/json" id="shop-shipping">${JSON.stringify(shopMeta.shipping)}</script>
  `
}
