/**
 * components.mjs — atoms & molecules.
 *
 * Every page is assembled from these functions. Pages never hand-write markup
 * that a component owns; if two pages need the same block, it lives here.
 */

/* ── helpers ────────────────────────────────────────────────────────────── */

export const esc = (s = '') =>
  String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

export const attr = (name, value) => (value === undefined || value === false || value === null ? '' : ` ${name}="${esc(value)}"`)

export const cx = (...parts) => parts.filter(Boolean).join(' ')

export const koDate = (iso) => {
  const [y, m, d] = iso.split('-')
  return `${y}.${m}.${d}`
}

/* ── icons (inline SVG, currentColor) ───────────────────────────────────── */

export const icon = {
  arrow: `<svg class="button__arrow" viewBox="0 0 16 16" fill="none" aria-hidden="true"><path d="M3 8h10M9 4l4 4-4 4" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round"/></svg>`,
  arrowDown: `<svg viewBox="0 0 16 16" fill="none" aria-hidden="true"><path d="M8 3v10M4 9l4 4 4-4" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round"/></svg>`,
  diagonal: `<svg viewBox="0 0 16 16" fill="none" aria-hidden="true"><path d="M5 11L11 5M6 5h5v5" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round"/></svg>`,
}

/* ── atoms ──────────────────────────────────────────────────────────────── */

export function button({ label, href, variant = '', size = '', arrow = true, cursor = '', attrs = '' }) {
  const cls = cx('button', variant && `button--${variant}`, size && `button--${size}`)
  const inner = `<span class="button__label">${esc(label)}</span>${arrow ? icon.arrow : ''}`
  const cursorAttr = cursor ? ` data-cursor="${esc(cursor)}"` : ''
  return href
    ? `<a class="${cls}" href="${esc(href)}"${cursorAttr}${attrs}>${inner}</a>`
    : `<button class="${cls}" type="button"${cursorAttr}${attrs}>${inner}</button>`
}

export function eyebrow(text, { scramble = true, className = '' } = {}) {
  return `<span class="${cx('eyebrow', className)}"${scramble ? ' data-scramble' : ''}>${esc(text)}</span>`
}

export function tag(text) {
  return `<li class="tag">${esc(text)}</li>`
}

export function tagList(items = []) {
  if (!items.length) return ''
  return `<ul class="tag-list">${items.map(tag).join('')}</ul>`
}

/**
 * Image slot. `name` maps to assets/images/<name>.webp produced by
 * scripts/fetch-images.mjs; the manifest it writes supplies width/height/alt so
 * every <img> ships intrinsic dimensions (no layout shift) and real alt text.
 */
export function media(name, { ratio = '16x9', alt = '', className = '', reveal = true, zoom = false, priority = false, manifest = {} } = {}) {
  const m = manifest[name] || {}
  const src = m.src || `/assets/images/${name}.webp`
  const w = m.width || 1600
  const h = m.height || 900
  const altText = alt || m.alt || ''
  return `<div class="${cx('media', `media--${ratio}`, className)}"${reveal ? ' data-media-reveal' : ''}>
        <img src="${esc(src)}" alt="${esc(altText)}" width="${w}" height="${h}"
             class="${zoom ? 'media__zoom' : ''}"
             loading="${priority ? 'eager' : 'lazy'}" decoding="async"${priority ? ' fetchpriority="high"' : ''}>
      </div>`
}

export function stat({ value, suffix = '', label, labelEn = '', format = '' }, { starfield = false } = {}) {
  return `<div class="stat">
        <span class="${starfield ? 'stat__value type-starfield' : 'stat__value'}"><span data-counter="${value}"${format ? ` data-counter-format="${format}"` : ''}>0</span>${suffix ? `<span class="stat__suffix">${esc(suffix)}</span>` : ''}</span>
        <span class="stat__label"${labelEn ? ` lang="en" aria-label="${esc(label)}"` : ''}>${esc(labelEn || label)}</span>
      </div>`
}

/* ── molecules ──────────────────────────────────────────────────────────── */

export function sectionHead({ tag: eyebrowText, title, aside = '', stack = false, id = '' }) {
  return `<header class="${cx('section-head', stack && 'section-head--stack')}"${id ? ` id="${esc(id)}"` : ''}>
        <div>
          ${eyebrowText ? `<div class="section-head__eyebrow">${eyebrow(eyebrowText)}</div>` : ''}
          <h2 class="section-head__title" data-split>${esc(title)}</h2>
        </div>
        ${aside ? `<div class="section-head__aside" data-reveal>${aside}</div>` : ''}
      </header>`
}

export function breadcrumb(items = []) {
  const last = items.length - 1
  return `<nav aria-label="브레드크럼">
        <ol class="breadcrumb">
          ${items
            .map((it, i) =>
              i === last
                ? `<li aria-current="page">${esc(it.label)}</li>`
                : `<li><a href="${esc(it.href)}">${esc(it.label)}</a></li>`
            )
            .join('')}
        </ol>
      </nav>`
}

/**
 * Sub-page hero.
 *
 * `title` is set in English on every page — the display line is Latin across the
 * site, matching the services tablist and the home headings. `subtitle` carries
 * the Korean name directly beneath it so the page still reads in Korean and the
 * h1's accessible name includes both.
 */
export function pageHero({ crumbs = [], tag: eyebrowText, title, subtitle = '', lead = '', foot = '' }) {
  /* aria-labelledby, not aria-label: GSAP SplitText rewrites the heading's
     children and stamps its own aria-label with the visible (English) text,
     which would silently drop the Korean name from the accessible name.
     labelledby outranks label, so this survives SplitText and the heading
     announces "정밀 세정, PRECISION CLEANING". */
  const ids = subtitle ? ' id="page-title" aria-labelledby="page-title-ko page-title"' : ''
  return `<section class="page-hero theme-dark padding-global">
      <div class="container page-hero__inner">
        ${crumbs.length ? breadcrumb(crumbs) : ''}
        ${eyebrowText ? `<div class="section-head__eyebrow">${eyebrow(eyebrowText)}</div>` : ''}
        <h1 class="page-hero__title" lang="en"${ids} data-split>${esc(title)}</h1>
        ${subtitle ? `<p class="t-ko" id="page-title-ko" lang="ko" style="margin-top:var(--space-2xs);font-size:var(--t-md)" data-reveal>${esc(subtitle)}</p>` : ''}
        ${lead ? `<p class="page-hero__lead u-lead" data-split>${esc(lead)}</p>` : ''}
        ${foot ? `<div class="page-hero__foot" data-reveal>${foot}</div>` : ''}
      </div>
    </section>`
}

export function accordion(items = [], { idPrefix = 'faq', single = true } = {}) {
  return `<div class="accordion" data-accordion-single="${single}">
        ${items
          .map((item, i) => {
            const id = `${idPrefix}-${i + 1}`
            return `<div class="accordion-css__item">
            <h3>
              <button class="accordion__trigger" type="button" id="${id}-trigger" aria-expanded="false" aria-controls="${id}-panel">
                <span class="accordion__question">${esc(item.q)}</span>
                <span class="accordion__icon" aria-hidden="true"></span>
              </button>
            </h3>
            <div class="accordion__panel" id="${id}-panel" role="region" aria-labelledby="${id}-trigger" data-open="false">
              <div class="accordion__panel-inner">
                <p class="accordion__answer">${esc(item.a)}</p>
              </div>
            </div>
          </div>`
          })
          .join('')}
      </div>`
}

export function marquee(items = []) {
  return `<div class="marquee theme-dark" aria-hidden="true">
        <div class="hs-track">
          ${items.map((t) => `<span class="hs-track__item">${esc(t)}</span>`).join('')}
        </div>
      </div>`
}

/**
 * Use-case wall.
 *
 * Each cell holds `perCell` image cards stacked on top of one another and cycles
 * through them forever (motion.js > logoWall). The list is chunked in order, so
 * `logos` order IS the layout: items 1-2 are cell one, 3-4 are cell two, and so on.
 *
 * Previously this was a wall of text labels with a spare pool swapped in. Photos
 * of the actual environments say "이런 곳에서 씁니다" far faster than a wordmark
 * can, so the label now rides along as a caption instead of being the content —
 * which also keeps the meaning (and the alt text) intact.
 */
export function logoWall(logos = [], { perCell = 2, manifest = {} } = {}) {
  const cells = []
  for (let i = 0; i < logos.length; i += perCell) cells.push(logos.slice(i, i + perCell))

  return `<div class="logo-wall">
        ${cells
          .map(
            (cell) => `<div class="logo-wall__cell">
          ${cell
            .map((it, i) => {
              const m = manifest[it.image] || {}
              /* The first card of every cell is the CSS resting state, so the wall
                 is complete and legible before GSAP runs — and stays that way under
                 reduced motion, where the cycle never starts.

                 The caption is the Latin name; the Korean rides in `alt`, so the
                 meaning and the accessible name both survive. */
              return `<figure class="logo-card${i === 0 ? ' is-active' : ''}">
              <img src="${esc(m.src || `/assets/images/${it.image}.webp`)}" alt="${esc(it.label)}"
                   width="${m.width || 800}" height="${m.height || 600}" loading="lazy" decoding="async">
              <figcaption class="logo-card__label" lang="en">${esc(it.labelEn || it.label)}</figcaption>
            </figure>`
            })
            .join('')}
        </div>`
          )
          .join('')}
      </div>`
}

export function workCard(work, index, manifest) {
  return `<a class="card" href="/portfolio/${esc(work.slug)}/" data-cursor="VIEW">
        <div class="card__media">
          <span class="card__index">${String(index + 1).padStart(2, '0')}</span>
          ${media(work.image, { ratio: '4x3', alt: `${work.title} — ${work.sector} 사용처 이미지`, zoom: true, manifest })}
        </div>
        <div class="stack stack--sm">
          <div class="meta-row"><span>${esc(work.sector)}</span><span>${esc(work.scale)}</span><span>${esc(work.category)}</span></div>
          <h3 class="card__title" lang="en">${esc(work.titleLatin)}<span class="t-ko" lang="ko">${esc(work.title)}</span></h3>
          <p class="card__excerpt">${esc(work.summary)}</p>
        </div>
      </a>`
}

/**
 * Services tablist.
 *
 * Set in English: the tab titles are the Latin service names and the one-liners
 * are the English summaries, matching the display-type convention used across
 * the site (English display headings, Korean running copy). The Korean name is
 * kept as a small second line so the list is still scannable in Korean, and
 * `aria-label` carries it for screen readers.
 */
export function servicesTabs(services, { limit } = {}) {
  const list = limit ? services.slice(0, limit) : services
  return `<div class="services-tab" role="tablist" aria-label="Products">
        ${list
          .map(
            (s) => `<a class="services-tab__item" role="tab" aria-selected="false" tabindex="0"
             href="/services/${esc(s.slug)}/" data-cursor="OPEN" aria-label="${esc(s.title)} — ${esc(s.titleLatin)}">
            <span class="services-tab__num">(${esc(s.num)})</span>
            <span class="services-tab__title" lang="en">${esc(s.titleLatin)}
              <span class="t-ko" lang="ko">${esc(s.title)}</span>
              <span class="services-tab__desc" lang="en">${esc(s.summaryEn || s.summary)}</span>
            </span>
            <span class="services-tab__svg-wrapper" aria-hidden="true">${icon.diagonal}</span>
          </a>`
          )
          .join('')}
      </div>`
}

/**
 * 3D cylinder carousel.
 *
 * The analysed original span 4 large panels (03-motion > "3D image carousel").
 * This build feeds it many more, smaller panels so the whole ring reads as one
 * rotating object rather than a few big slabs passing by — a deliberate change,
 * recorded in MISSING.md > C8. The spin itself keeps the measured parameters
 * (rotationY -=360, 30s, linear, infinite) and the list keeps its 5° tilt.
 *
 * `items`: [{ image, caption }]
 */
export function carousel3D(items, manifest) {
  return `<div class="img-carousel" data-panels="${items.length}">
        <div class="img-carousel__list">
          <div class="img-carousel__stage">
            ${items
              .map((it) => {
                const m = manifest[it.image] || {}
                return `<div class="img-carousel__panel">
                <img src="${esc(m.src || `/assets/images/${it.image}.webp`)}"
                     alt="${esc(it.caption)}" width="${m.width || 900}" height="${m.height || 1200}"
                     loading="lazy" decoding="async">
                <span class="img-carousel__caption">${esc(it.caption)}</span>
              </div>`
              })
              .join('')}
          </div>
        </div>
      </div>`
}

export function pager({ prev, next, base }) {
  if (!prev && !next) return ''
  return `<nav class="pager" aria-label="문서 이동">
        ${prev
          ? `<a class="pager__item" href="${esc(base + prev.slug)}/"><span class="pager__label">이전</span><span class="pager__title" lang="en">${esc(prev.titleLatin || prev.title)}<span class="t-ko" lang="ko">${esc(prev.title)}</span></span></a>`
          : '<span></span>'}
        ${next
          ? `<a class="pager__item pager__item--next" href="${esc(base + next.slug)}/"><span class="pager__label">다음</span><span class="pager__title" lang="en">${esc(next.titleLatin || next.title)}<span class="t-ko" lang="ko">${esc(next.title)}</span></span></a>`
          : '<span></span>'}
      </nav>`
}

/* ── contact block (shared by home + every sub-page footer CTA) ─────────── */

/**
 * Contact block: copy, a real enquiry form, and the direct details.
 *
 * ⚠ There is no backend. The form validates fully in the browser and then hands
 * off by composing a mail draft (`mailto:`) with the answers filled in — which
 * genuinely works from a static host and never pretends a submission was
 * received. To POST to a real endpoint instead, set `data-endpoint` on the form
 * and ui.js will send JSON there. See MISSING.md > D6.
 */
export function contactSection(contact, site, { id = 'contact' } = {}) {
  const field = ({ name, label, type = 'text', required = false, autocomplete = '', placeholder = '' }) => `
        <p class="field">
          <label class="field__label" for="cf-${name}">${esc(label)}${required ? '<span class="field__req" aria-hidden="true">*</span>' : ''}</label>
          <input class="field__input" id="cf-${name}" name="${name}" type="${type}"
                 ${required ? 'required' : ''}${autocomplete ? ` autocomplete="${autocomplete}"` : ''}${placeholder ? ` placeholder="${esc(placeholder)}"` : ''}
                 aria-describedby="cf-${name}-err">
          <span class="field__err" id="cf-${name}-err" data-error-for="${name}" role="alert"></span>
        </p>`

  return `<section class="section contact padding-global padding-section" id="${esc(id)}">
      <div class="container">
        <div class="section-head__eyebrow">${eyebrow(contact.eyebrow)}</div>
        <h2 class="contact__title" data-split>${esc(contact.heading)}</h2>
        <p class="u-lead contact__lead" data-split>${esc(contact.lead)}</p>

        <div class="contact__cols">
          <form class="contact-form" id="contact-form" novalidate data-reveal-group
                data-mailto="${esc(site.email)}">
            <div class="contact-form__row">
              ${field({ name: 'company', label: '상호 · 사용처', required: true, autocomplete: 'organization' })}
              ${field({ name: 'name', label: '담당자 성함', required: true, autocomplete: 'name' })}
            </div>
            <div class="contact-form__row">
              ${field({ name: 'email', label: '이메일', type: 'email', required: true, autocomplete: 'email' })}
              ${field({ name: 'phone', label: '연락처', type: 'tel', autocomplete: 'tel', placeholder: '선택' })}
            </div>

            <p class="field">
              <label class="field__label" for="cf-topic">문의 품목</label>
              <select class="field__input" id="cf-topic" name="topic">
                <option value="키친타월">키친타월</option>
                <option value="점보롤티슈">점보롤티슈</option>
                <option value="두루마리티슈">두루마리티슈</option>
                <option value="미용티슈 · 냅킨 · 핸드타월">미용티슈 · 냅킨 · 핸드타월</option>
                <option value="OEM · 주문제작">OEM · 주문제작</option>
                <option value="대량 공급 · 정기 배송">대량 공급 · 정기 배송</option>
                <option value="기타">기타</option>
              </select>
            </p>

            <p class="field">
              <label class="field__label" for="cf-message">문의 내용<span class="field__req" aria-hidden="true">*</span></label>
              <textarea class="field__input field__input--area" id="cf-message" name="message" rows="5" required
                        aria-describedby="cf-message-err"
                        placeholder="쓰시는 곳(주방·화장실·객실 등), 대략적인 사용 인원, 희망 수량과 납품 시기를 적어 주시면 견적이 빨라집니다. 샘플이 필요하시면 함께 적어 주세요."></textarea>
              <span class="field__err" id="cf-message-err" data-error-for="message" role="alert"></span>
            </p>

            <p class="field field--check">
              <input class="field__check" id="cf-consent" name="consent" type="checkbox" required
                     aria-describedby="cf-consent-err">
              <label for="cf-consent">
                <a href="/utility-pages/privacy-notice/" data-underline-link>개인정보 처리방침</a>에 따른 수집·이용에 동의합니다.<span class="field__req" aria-hidden="true">*</span>
              </label>
              <span class="field__err" id="cf-consent-err" data-error-for="consent" role="alert"></span>
            </p>

            <div class="contact-form__actions">
              ${button({ label: '견적 요청 보내기', variant: 'solid', size: 'lg', cursor: 'SEND', attrs: ' type="submit"' })}
              <p class="contact-form__status" role="status" aria-live="polite"></p>
            </div>
          </form>

          <div class="contact__aside" data-reveal-group>
            <dl class="contact__detail-list">
              <div class="contact__detail"><dt>전화</dt><dd><a href="tel:${esc(site.phone.replace(/[^0-9+]/g, ''))}" data-underline-link>${esc(site.phone)}</a></dd></div>
              <div class="contact__detail"><dt>팩스</dt><dd>${esc(site.fax)}</dd></div>
              <div class="contact__detail"><dt>이메일</dt><dd><a href="mailto:${esc(site.email)}" data-underline-link>${esc(site.email)}</a></dd></div>
              <div class="contact__detail"><dt>주소</dt><dd>${esc(site.address.region)} ${esc(site.address.locality)} ${esc(site.address.street)}</dd></div>
              <div class="contact__detail"><dt>운영 시간</dt><dd>${esc(site.hours)}</dd></div>
            </dl>
          </div>
        </div>
      </div>
    </section>`
}
