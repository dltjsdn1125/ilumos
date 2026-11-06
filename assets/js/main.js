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

/* 모션 준비가 끝나기 전에 ScrollTrigger.refresh 를 여러 번 부르면 같은 계산이
   겹쳐서 메인 스레드를 오래 잡습니다. 글꼴·로드 알림은 한 번으로 모아서,
   모션이 다 만들어진 뒤 브라우저가 한가할 때 한 번만 다시 잽니다. */
let motionReady = Promise.resolve()
let refreshTimer = 0
function refreshLater() {
  clearTimeout(refreshTimer)
  refreshTimer = setTimeout(() => {
    motionReady.then(() => {
      const run = () => window.ScrollTrigger?.refresh()
      if ('requestIdleCallback' in window) window.requestIdleCallback(run, { timeout: 1500 })
      else run()
    })
  }, 250)
}

function boot() {
  document.documentElement.classList.add('js')

  /* WebGL 배경은 첫 화면이 그려지고 페이지가 다 로드된 뒤, 브라우저가 한가할 때 켭니다.
     첫 화면(LCP)과 같은 시간대에 셰이더 컴파일·별 데이터 생성이 몰려 누름 반응을
     막지 않게 하기 위해서입니다. 배경은 정적 이미지가 먼저 깔려 있어 모습은 같습니다. */
  const glHost = document.querySelector('.unicorn__item')
  if (glHost) {
    const startGL = () => {
      const run = () => initHeroGL(glHost)
      if ('requestIdleCallback' in window) window.requestIdleCallback(run, { timeout: 2000 })
      else setTimeout(run, 200)
    }
    if (document.readyState === 'complete') startGL()
    else window.addEventListener('load', startGL, { once: true })
  }

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

  motionReady = Promise.resolve(initMotion()).catch(() => {})

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
  document.fonts.ready.then(refreshLater)
}

/* Landing straight on an anchor (/#portfolio) jumps the scroller before
   ScrollTrigger has measured anything, which can leave that section stuck at its
   pre-reveal state. A refresh once everything has loaded re-evaluates every
   trigger against the real scroll position and plays whatever is already in view. */
window.addEventListener('load', refreshLater, { once: true })
