/* ==========================================================================
   main.js — entry point
   Boot order: WebGL backdrop -> UI -> motion. GSAP/Lenis are vendored locally
   (assets/js/vendor) and loaded as classic scripts before this module, so the
   site has no external runtime dependency.
   ========================================================================== */

import { initHeroGL } from './hero-gl.js'
import { initCursor, initHeader, initMenu, initAccordion, initMarquee, markCurrentNav, initContactForm } from './ui.js'
import { initMotion } from './motion.js'
import { initShop } from './shop.js'

function boot() {
  document.documentElement.classList.add('js')

  const glHost = document.querySelector('.unicorn__item')
  if (glHost) initHeroGL(glHost)

  initHeader()
  initMenu()
  initAccordion()
  initMarquee()
  initCursor()
  markCurrentNav()
  initContactForm()
  /* 헤더 배지는 모든 페이지에서 필요하므로 shop 라우트가 아니어도 부릅니다.
     나머지(분류·장바구니 화면)는 각자 자기 훅이 없으면 즉시 빠져나갑니다. */
  initShop()

  initMotion()

  // Footer year, kept out of the generated HTML so pages don't go stale.
  document.querySelectorAll('[data-year]').forEach((el) => {
    el.textContent = String(new Date().getFullYear())
  })
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', boot, { once: true })
} else {
  boot()
}

// Fonts shift line breaks; re-measure once they land so ScrollTrigger positions hold.
if (document.fonts?.ready) {
  document.fonts.ready.then(() => window.ScrollTrigger?.refresh())
}

/* Landing straight on an anchor (/#portfolio) jumps the scroller before
   ScrollTrigger has measured anything, which can leave that section stuck at its
   pre-reveal state. A refresh once everything has loaded re-evaluates every
   trigger against the real scroll position and plays whatever is already in view. */
window.addEventListener('load', () => {
  window.ScrollTrigger?.refresh()
}, { once: true })
