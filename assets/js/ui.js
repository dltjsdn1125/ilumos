/* ==========================================================================
   ui.js — cursor, nav overlay, accordion, header behaviour
   Sources: 06-components.json (custom cursor, header, FAQ accordion) and
   02-source > hover_selectors_stylesheet / state_selector_patterns
   (156 state rules, driven by data-attributes rather than classes).
   ========================================================================== */

const REDUCED = window.matchMedia('(prefers-reduced-motion: reduce)').matches
const FINE_POINTER = window.matchMedia('(hover: hover) and (pointer: fine)').matches

/* ── custom cursor ──────────────────────────────────────────────────────────
   06-components: [data-cursor] driven, mix-blend-mode difference, inertia
   follow with a custom 1-pow(1-t,r) ease, scales on [data-cursor]:hover.
   Implemented with a per-frame lerp — the same asymptotic curve that ease
   describes — so it works with or without InertiaPlugin present. */

export function initCursor() {
  if (!FINE_POINTER || REDUCED) return

  const dot = document.querySelector('.cursor-dot')
  const ring = document.querySelector('.cursor-ring')
  const star = document.querySelector('.cursor-star')
  if (!dot || !ring) return

  const pos = { x: innerWidth / 2, y: innerHeight / 2 }
  const dotPos = { ...pos }
  const ringPos = { ...pos }
  const starPos = { ...pos }
  let scale = 1
  let scaleTarget = 1
  let raf = 0
  let shown = false
  let onDark = true

  /* Which cursor to show is decided by the block under the pointer: the
     difference-blended ring inside the hero, where it sits directly on the star
     field, and the blue star everywhere else. The test is the HERO BLOCK, not
     the theme — the whole page runs one dark tone now, so a theme test would
     never switch. elementFromPoint is safe here because every cursor node is
     pointer-events: none, so it never returns one of them. */
  function themeAt(x, y) {
    const el = document.elementFromPoint(x, y)
    return el ? !!el.closest('.hero, .page-hero, .menu-overlay') : true
  }

  function applyTheme() {
    dot.style.opacity = shown && onDark ? '1' : '0'
    ring.style.opacity = shown && onDark ? '1' : '0'
    star?.classList.toggle('is-visible', shown && !onDark)
  }

  function onMove(e) {
    pos.x = e.clientX
    pos.y = e.clientY
    shown = true
    const nowDark = themeAt(e.clientX, e.clientY)
    if (nowDark !== onDark) {
      onDark = nowDark
    }
    applyTheme()
  }

  function onOver(e) {
    const target = e.target.closest('[data-cursor]')
    if (target) {
      scaleTarget = Number(target.dataset.cursorScale || 2.1)
      ring.classList.add('is-active')
      star?.classList.add('is-active')
    } else {
      scaleTarget = 1
      ring.classList.remove('is-active')
      star?.classList.remove('is-active')
    }
  }

  function onLeave() {
    shown = false
    applyTheme()
  }

  function tick() {
    raf = requestAnimationFrame(tick)
    // Three follow rates: the dot is nearly locked to the pointer, the ring
    // trails, and the star trails a touch less so it reads as attached.
    dotPos.x += (pos.x - dotPos.x) * 0.9
    dotPos.y += (pos.y - dotPos.y) * 0.9
    ringPos.x += (pos.x - ringPos.x) * 0.14
    ringPos.y += (pos.y - ringPos.y) * 0.14
    starPos.x += (pos.x - starPos.x) * 0.22
    starPos.y += (pos.y - starPos.y) * 0.22
    scale += (scaleTarget - scale) * 0.16

    dot.style.transform = `translate3d(${dotPos.x}px, ${dotPos.y}px, 0)`
    ring.style.transform = `translate3d(${ringPos.x}px, ${ringPos.y}px, 0) scale(${scale})`
    if (star) star.style.transform = `translate3d(${starPos.x}px, ${starPos.y}px, 0) scale(${0.85 + (scale - 1) * 0.5})`
  }

  window.addEventListener('pointermove', onMove, { passive: true })
  window.addEventListener('pointerover', onOver, { passive: true })
  document.addEventListener('pointerleave', onLeave)
  window.addEventListener('blur', onLeave)
  raf = requestAnimationFrame(tick)
}

/* ── header: hide on scroll down, show on scroll up ─────────────────────────
   The sweep recorded header top:0 across all 21 stops, i.e. always present.
   The hide-on-scroll-down refinement is ours and never traps the header
   off-screen (it always returns on upward scroll or at the top). */

export function initHeader() {
  const header = document.querySelector('.header')
  if (!header) return

  let last = window.scrollY
  let ticking = false

  /* Theme switch. The page runs dark inside the hero and light everywhere below,
     so the fixed header has to flip as it crosses the boundary. An
     IntersectionObserver on the hero would only tell us when the WHOLE hero has
     left; what matters is whether the hero still covers the header's own band,
     so we compare the hero's bottom edge against the header's height. */
  const heroBlock = document.querySelector('.hero, .page-hero')

  function updateTheme() {
    if (!heroBlock) {
      header.classList.remove('theme-dark')
      return
    }
    const bottom = heroBlock.getBoundingClientRect().bottom
    header.classList.toggle('theme-dark', bottom > header.offsetHeight * 0.65)
  }

  function update() {
    const y = window.scrollY
    const goingDown = y > last && y > 260
    header.classList.toggle('is-hidden', goingDown && !document.body.classList.contains('is-menu-open'))
    updateTheme()
    last = y
    ticking = false
  }

  updateTheme()
  window.addEventListener('resize', updateTheme)

  window.addEventListener(
    'scroll',
    () => {
      if (!ticking) {
        ticking = true
        requestAnimationFrame(update)
      }
    },
    { passive: true }
  )
}

/* ── fullscreen overlay menu ────────────────────────────────────────────── */

export function initMenu() {
  const toggle = document.querySelector('.menu-toggle')
  const overlay = document.querySelector('.menu-overlay')
  if (!toggle || !overlay) return

  const links = Array.from(overlay.querySelectorAll('.menu-overlay__link'))
  const focusables = () =>
    Array.from(overlay.querySelectorAll('a[href], button:not([disabled])')).filter((el) => el.offsetParent !== null)

  let open = false

  function setLinks(visible) {
    const gsap = window.gsap
    if (!gsap || REDUCED) {
      links.forEach((l) => { l.style.transform = visible ? 'translateY(0)' : 'translateY(110%)' })
      return
    }
    // Same reveal signature as the page: yPercent 110 -> 0, expo.out, 0.1 stagger.
    gsap.to(links, {
      yPercent: visible ? 0 : 110,
      duration: visible ? 0.8 : 0.4,
      ease: visible ? 'expo.out' : 'power3.in',
      stagger: visible ? 0.06 : 0.02,
      delay: visible ? 0.12 : 0,
    })
  }

  function setOpen(next) {
    open = next
    toggle.setAttribute('aria-expanded', String(open))
    overlay.classList.toggle('is-open', open)
    overlay.setAttribute('aria-hidden', String(!open))
    document.body.classList.toggle('is-menu-open', open)

    const lenis = window.__lenis
    if (lenis) open ? lenis.stop() : lenis.start()
    document.documentElement.style.overflow = open ? 'hidden' : ''

    setLinks(open)
    if (open) focusables()[0]?.focus()
    else toggle.focus()
  }

  setLinks(false)
  overlay.setAttribute('aria-hidden', 'true')

  toggle.addEventListener('click', () => setOpen(!open))

  overlay.addEventListener('click', (e) => {
    if (e.target.closest('a')) setOpen(false)
  })

  document.addEventListener('keydown', (e) => {
    if (!open) return
    if (e.key === 'Escape') { setOpen(false); return }
    if (e.key !== 'Tab') return
    // focus trap
    const items = focusables()
    if (!items.length) return
    const first = items[0]
    const last = items[items.length - 1]
    if (e.shiftKey && document.activeElement === first) { e.preventDefault(); last.focus() }
    else if (!e.shiftKey && document.activeElement === last) { e.preventDefault(); first.focus() }
  })
}

/* ── FAQ accordion ──────────────────────────────────────────────────────────
   06-components: display:grid item, grid-template-rows animates open/closed.
   The CSS owns the transition; JS only flips state and ARIA. */

export function initAccordion() {
  document.querySelectorAll('.accordion').forEach((root) => {
    const single = root.dataset.accordionSingle !== 'false'
    const triggers = Array.from(root.querySelectorAll('.accordion__trigger'))

    triggers.forEach((trigger) => {
      const panel = document.getElementById(trigger.getAttribute('aria-controls'))
      if (!panel) return

      trigger.addEventListener('click', () => {
        const willOpen = trigger.getAttribute('aria-expanded') !== 'true'

        if (single && willOpen) {
          triggers.forEach((other) => {
            if (other === trigger) return
            other.setAttribute('aria-expanded', 'false')
            const p = document.getElementById(other.getAttribute('aria-controls'))
            if (p) p.dataset.open = 'false'
          })
        }

        trigger.setAttribute('aria-expanded', String(willOpen))
        panel.dataset.open = String(willOpen)
        window.ScrollTrigger?.refresh()
      })
    })
  })
}

/* ── marquee: duplicate the track so the -100% keyframe loops seamlessly ─── */

export function initMarquee() {
  document.querySelectorAll('.marquee').forEach((m) => {
    const track = m.querySelector('.hs-track')
    if (!track || track.dataset.cloned === 'true') return
    const clone = track.cloneNode(true)
    clone.setAttribute('aria-hidden', 'true')
    m.appendChild(clone)
    track.dataset.cloned = 'true'
  })
}

/* ── current-route marking ──────────────────────────────────────────────── */

export function markCurrentNav() {
  const here = location.pathname.replace(/index\.html$/, '')
  document.querySelectorAll('.nav-desktop__link, .menu-overlay__link').forEach((a) => {
    const href = a.getAttribute('href') || ''
    if (href.startsWith('/#') || href === '/') return
    const path = href.replace(/index\.html$/, '')
    if (path !== '/' && here.startsWith(path)) a.setAttribute('aria-current', 'page')
  })
}

/* ── contact form ────────────────────────────────────────────────────────────
   Full client-side validation, then a handoff that is honest about the fact that
   this build has no server:

     · `data-endpoint` present  -> POST JSON there
     · otherwise                -> open a pre-filled mail draft (mailto:)

   The mailto path is the default because it actually works from a static host.
   The form never claims a submission was received when nothing received it. */

export function initContactForm() {
  const form = document.querySelector('.contact-form')
  if (!form) return

  const status = form.querySelector('.contact-form__status')
  const setStatus = (msg, kind = '') => {
    if (!status) return
    status.textContent = msg
    status.dataset.kind = kind
  }

  const showError = (name, msg) => {
    const slot = form.querySelector(`[data-error-for="${name}"]`)
    const input = form.elements[name]
    if (slot) slot.textContent = msg
    if (input) {
      input.setAttribute('aria-invalid', msg ? 'true' : 'false')
      input.classList.toggle('is-invalid', Boolean(msg))
    }
  }

  const RULES = {
    company: (v) => (v.trim() ? '' : '상호 또는 사용처를 입력해 주세요.'),
    name: (v) => (v.trim() ? '' : '담당자 성함을 입력해 주세요.'),
    email: (v) => {
      if (!v.trim()) return '이메일을 입력해 주세요.'
      // Deliberately permissive: the only reliable test of an address is sending to it.
      return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(v.trim()) ? '' : '이메일 형식을 확인해 주세요.'
    },
    phone: (v) => (!v.trim() || /^[0-9+\-\s()]{7,}$/.test(v.trim()) ? '' : '연락처 형식을 확인해 주세요.'),
    message: (v) => (v.trim().length >= 10 ? '' : '문의 내용을 10자 이상 적어 주세요.'),
    consent: (v, el) => (el.checked ? '' : '개인정보 수집·이용 동의가 필요합니다.'),
  }

  const validateField = (name) => {
    const el = form.elements[name]
    if (!el || !RULES[name]) return true
    const msg = RULES[name](el.value ?? '', el)
    showError(name, msg)
    return !msg
  }

  // Validate on blur, and clear the error live once the user starts fixing it.
  for (const name of Object.keys(RULES)) {
    const el = form.elements[name]
    if (!el) continue
    el.addEventListener('blur', () => validateField(name))
    el.addEventListener('input', () => {
      if (el.getAttribute('aria-invalid') === 'true') validateField(name)
    })
    el.addEventListener('change', () => {
      if (el.type === 'checkbox') validateField(name)
    })
  }

  form.addEventListener('submit', async (e) => {
    e.preventDefault()
    setStatus('')

    const names = Object.keys(RULES)
    const ok = names.map(validateField).every(Boolean)
    if (!ok) {
      const firstBad = names.find((n) => form.elements[n]?.getAttribute('aria-invalid') === 'true')
      form.elements[firstBad]?.focus()
      setStatus('입력을 확인해 주세요.', 'error')
      return
    }

    const data = Object.fromEntries(new FormData(form).entries())
    const endpoint = form.dataset.endpoint

    if (endpoint) {
      setStatus('보내는 중…')
      try {
        const res = await fetch(endpoint, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(data),
        })
        if (!res.ok) throw new Error(`HTTP ${res.status}`)
        form.reset()
        setStatus('문의가 접수되었습니다. 영업일 기준 1일 내 회신드립니다.', 'ok')
      } catch (err) {
        setStatus(`전송에 실패했습니다 (${err.message}). 이메일로 보내 주세요.`, 'error')
      }
      return
    }

    // No endpoint configured: compose a mail draft instead of faking a send.
    const to = form.dataset.mailto || ''
    const subject = `[견적 요청] ${data.company || ''} — ${data.topic || '문의'}`
    const body = [
      `상호 · 사용처: ${data.company || ''}`,
      `담당자: ${data.name || ''}`,
      `이메일: ${data.email || ''}`,
      `연락처: ${data.phone || '-'}`,
      `문의 품목: ${data.topic || ''}`,
      '',
      data.message || '',
    ].join('\n')

    window.location.href = `mailto:${to}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`
    setStatus('메일 작성 창을 열었습니다. 전송까지 완료해 주세요.', 'ok')
  })
}
