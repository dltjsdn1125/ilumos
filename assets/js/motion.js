/* ==========================================================================
   motion.js — the choreography

   Every duration, ease, distance, stagger and trigger position below comes from
   03-motion.json (45-child global timeline census + 7 ScrollTriggers + 21-stop
   scroll sweep) and 05-design-system.json > motion_tokens. Where the analysis
   recorded a value it is used literally; where it recorded behaviour but not a
   number, the reconstruction is called out in a comment.

   Default cadence on the source site:
       ScrollTrigger { start: 'top 80%', toggleActions: 'play', once: true }
   ========================================================================== */

const REDUCED = window.matchMedia('(prefers-reduced-motion: reduce)').matches

/* Scramble glyph set — the report describes "a runic/symbol glyph set"; the exact
   characters were not captured, so this is a matched stand-in. */
const SCRAMBLE_CHARS = 'ᚠᚢᚦᚨᚱᚲᚷᚹᚺᚾᛁᛃᛇᛈᛉᛊᛏᛒᛖᛗᛚᛜᛞᛟ▓▒░#$%&@*+=<>/\\|'

const D = {
  media: 0.6,
  reveal: 0.8,
  logo: 0.8,
  counter: 1,
  carouselIn: 1.2,
  scramble: 1.5,
  spin: 30,
}

export function initMotion() {
  const { gsap } = window
  if (!gsap) {
    console.warn('[motion] GSAP missing — content is already visible, nothing to do')
    document.querySelectorAll('[data-reveal], [data-split]').forEach((el) => el.classList.add('is-revealed'))
    return
  }

  gsap.registerPlugin(
    window.ScrollTrigger,
    window.SplitText,
    window.ScrambleTextPlugin,
    window.Flip,
    window.Observer,
    window.InertiaPlugin,
    window.CustomEase
  )

  const ScrollTrigger = window.ScrollTrigger

  const lenis = initLenis(ScrollTrigger)

  // Reduced motion: resolve everything to its end state and stop here.
  if (REDUCED) {
    document.querySelectorAll('[data-reveal], [data-split]').forEach((el) => el.classList.add('is-revealed'))
    document.querySelectorAll('[data-scramble]').forEach((el) => {
      el.textContent = el.dataset.scrambleText || el.textContent
      el.classList.add('is-revealed')
    })
    document.querySelectorAll('[data-counter]').forEach((el) => {
      el.textContent = formatCount(Number(el.dataset.counter), el.dataset.counterFormat)
    })
    return { lenis }
  }

  heroWarp(gsap, ScrollTrigger)
  splitReveals(gsap, ScrollTrigger)
  fadeReveals(gsap, ScrollTrigger)
  scrambleLabels(gsap, ScrollTrigger)
  counters(gsap, ScrollTrigger)
  mediaReveals(gsap, ScrollTrigger)
  carousel3D(gsap, ScrollTrigger)
  scalingVideo(gsap, ScrollTrigger)
  logoWall(gsap, ScrollTrigger)
  heroWordCycle(gsap)

  ScrollTrigger.refresh()

  return { lenis }
}

/* ── Lenis 1.2.3: drives all scrolling and syncs ScrollTrigger ───────────── */

function initLenis(ScrollTrigger) {
  if (!window.Lenis || REDUCED) return null

  const lenis = new window.Lenis({
    duration: 1.1,
    smoothWheel: true,
    // Touch keeps native scrolling — matches the source's mobile behaviour.
    syncTouch: false,
  })

  lenis.on('scroll', ScrollTrigger.update)

  window.gsap.ticker.add((time) => lenis.raf(time * 1000))
  window.gsap.ticker.lagSmoothing(0)

  document.documentElement.classList.add('lenis')

  // Anchor links have to go through Lenis or they fight the smooth scroller.
  document.addEventListener('click', (e) => {
    const a = e.target.closest('a[href^="#"], a[href*="/#"]')
    if (!a) return
    const hash = a.getAttribute('href').split('#')[1]
    if (!hash) return
    const target = document.getElementById(hash)
    if (!target) return
    e.preventDefault()
    lenis.scrollTo(target, { offset: -96 })
  })

  window.__lenis = lenis
  return lenis
}

/* ── hero scroll-out: the star field warp ───────────────────────────────────
   The hero is pinned for one extra viewport. Through that pin the headline
   lifts and blurs away while hero-gl.js accelerates its star field into a warp
   and pushes the wireframe star back — so leaving the hero reads as the scene
   scattering into deep space rather than as a section simply scrolling by.

   The pin is what makes it a transition: without it the effect would only be
   visible in the shrinking sliver of hero still on screen, because every block
   below is opaque.

   The pin's own progress is published on window.__heroWarp; hero-gl.js reads it
   (a pinned element's rect never moves, so it cannot measure scroll itself).
   This is authored motion — the analysis records no such transition. */

function heroWarp(gsap, ScrollTrigger) {
  const hero = document.querySelector('.hero')
  if (!hero) return
  const inner = hero.querySelector('.hero__inner')
  const sceneLayers = document.querySelectorAll('.unicorn__scene')

  /* The fade has to ride the SAME trigger that does the pinning. A pinned
     element's rect stops moving, so a second ScrollTrigger aimed at `.hero`
     would sit at its start position forever and never advance. Attaching a
     timeline via `animation` maps the pin's scroll range onto the timeline
     instead, which is the only thing that tracks correctly while pinned. */
  const tl = gsap.timeline()

  if (inner) {
    gsap.set(inner, { filter: 'blur(0px)' })
    tl.to(inner, {
      yPercent: -12,
      opacity: 0,
      filter: 'blur(7px)',
      ease: 'none',
      duration: 0.72, // content is gone by 72% of the pin, leaving pure warp
    }, 0)
  }

  /* The forest dissolves on the same range as the copy, so leaving the hero
     reads as rising up out of the trees into the warp — not as a photo
     scrolling away. It also has to go: the backdrop layer is fixed and every
     section below is translucent, so a forest left at full strength would sit
     behind the whole document. */
  // NodeList is truthy even when empty, so test length, not the list.
  if (sceneLayers.length) {
    tl.to(sceneLayers, { opacity: 0, ease: 'none', duration: 0.62 }, 0)
  }

  tl.to({}, { duration: 0.28 }, 0.72) // hold so the timeline spans the full range

  ScrollTrigger.create({
    trigger: hero,
    start: 'top top',
    end: '+=110%',
    pin: true,
    pinSpacing: true,
    scrub: 0.6,
    animation: inner || sceneLayers.length ? tl : undefined,
    onUpdate: (self) => { window.__heroWarp = self.progress; window.__heroPast = false },
    // Past the pin the field settles to an ambient drift — see hero-gl.js.
    // Leaving it at full warp would streak behind every section below, since
    // those sections are translucent over the same canvas.
    onLeave: () => { window.__heroWarp = 1; window.__heroPast = true },
    onLeaveBack: () => { window.__heroWarp = 0; window.__heroPast = false },
    onEnterBack: () => { window.__heroPast = false },
  })
}

/* ── SplitText line reveals ─────────────────────────────────────────────────
   yPercent 110 -> 0 · 0.8s · expo.out · stagger 0.1 · once · start 'top 80%' */

function splitReveals(gsap, ScrollTrigger) {
  document.querySelectorAll('[data-split]').forEach((el) => {
    const split = new window.SplitText(el, {
      type: 'lines',
      linesClass: 'split-line',
      mask: 'lines', // the line-clip mask the census saw as negative margins
    })

    el.classList.add('is-revealed')

    gsap.from(split.lines, {
      yPercent: 110,
      duration: D.reveal,
      ease: 'expo.out',
      stagger: 0.1,
      scrollTrigger: {
        trigger: el,
        start: 'top 80%',
        toggleActions: 'play none none none',
        once: true,
      },
    })
  })
}

/* ── generic block reveal ───────────────────────────────────────────────── */

function fadeReveals(gsap, ScrollTrigger) {
  document.querySelectorAll('[data-reveal]').forEach((el) => {
    const delay = Number(el.dataset.revealDelay || 0)
    el.classList.add('is-revealed')

    /* fromTo, not from: heroWarp's pin triggers a ScrollTrigger.refresh() after
       these tweens are built, and a refresh re-records a from-tween's end value
       from whatever is on the element AT THAT MOMENT — which is the from-state
       itself. The block then animates 0 -> 0 and stays invisible. Spelling out
       both ends makes the destination immune to re-recording. */
    gsap.fromTo(el, {
      opacity: 0,
      y: 28,
    }, {
      opacity: 1,
      y: 0,
      duration: D.reveal,
      ease: 'expo.out',
      delay,
      scrollTrigger: {
        trigger: el,
        start: 'top 80%',
        toggleActions: 'play none none none',
        once: true,
      },
    })
  })

  // stagger groups: children animate in sequence, parent is the trigger
  document.querySelectorAll('[data-reveal-group]').forEach((group) => {
    const kids = Array.from(group.children)
    kids.forEach((k) => k.classList.add('is-revealed'))

    gsap.fromTo(kids, {
      opacity: 0,
      y: 28,
    }, {
      opacity: 1,
      y: 0,
      duration: D.reveal,
      ease: 'expo.out',
      stagger: 0.1,
      scrollTrigger: {
        trigger: group,
        start: 'top 80%',
        toggleActions: 'play none none none',
        once: true,
      },
    })
  })
}

/* ── ScrambleText labels ────────────────────────────────────────────────────
   1.5s · scrambleText { speed: 0.8, chars: runic/symbol } · once at top 80%
   Observed on eyebrows "(Who we are)", "(Our Clients)", "(Portfolio)",
   "(Our process)" and on the hero keywords. */

function scrambleLabels(gsap, ScrollTrigger) {
  document.querySelectorAll('[data-scramble]').forEach((el) => {
    const text = el.dataset.scrambleText || el.textContent.trim()
    el.dataset.scrambleText = text
    el.classList.add('is-revealed')

    gsap.to(el, {
      duration: D.scramble,
      scrambleText: { text, chars: SCRAMBLE_CHARS, speed: 0.8, revealDelay: 0.2 },
      ease: 'power4.out',
      scrollTrigger: {
        trigger: el,
        start: 'top 80%',
        toggleActions: 'play none none none',
        once: true,
      },
    })
  })
}

/* ── counters ───────────────────────────────────────────────────────────────
   SPAN.ds-num · 1s · power4 number roll */

function formatCount(value, format) {
  if (format === 'year') return String(Math.round(value))
  return Math.round(value).toLocaleString('ko-KR')
}

function counters(gsap, ScrollTrigger) {
  document.querySelectorAll('[data-counter]').forEach((el) => {
    const end = Number(el.dataset.counter)
    const format = el.dataset.counterFormat
    // Year counters roll from a nearby value, not from zero — rolling 0→2019 reads as noise.
    const startVal = format === 'year' ? end - 24 : 0
    const obj = { v: startVal }

    el.textContent = formatCount(startVal, format)

    gsap.to(obj, {
      v: end,
      duration: D.counter,
      ease: 'power4.out',
      onUpdate: () => { el.textContent = formatCount(obj.v, format) },
      scrollTrigger: {
        trigger: el,
        start: 'top 80%',
        toggleActions: 'play none none none',
        once: true,
      },
    })
  })
}

/* ── media reveal — IMG.ds-mediaimg · 0.6s · power2.out ─────────────────── */

function mediaReveals(gsap, ScrollTrigger) {
  document.querySelectorAll('[data-media-reveal]').forEach((el) => {
    const img = el.querySelector('img')

    gsap.fromTo(
      el,
      { clipPath: 'inset(0% 0% 100% 0%)' },
      {
        clipPath: 'inset(0% 0% 0% 0%)',
        duration: D.media,
        ease: 'power2.out',
        scrollTrigger: { trigger: el, start: 'top 85%', toggleActions: 'play none none none', once: true },
      }
    )

    if (img) {
      gsap.from(img, {
        scale: 1.14,
        duration: 1.1,
        ease: 'power2.out',
        scrollTrigger: { trigger: el, start: 'top 85%', toggleActions: 'play none none none', once: true },
      })
    }
  })
}

/* ── 3D image carousel ──────────────────────────────────────────────────────
   4 panels on a cylinder · rotationY '-=360' · 30s · ease none · repeat -1
   entrance 1.2s expo.inOut at top 80% · list holds a static 5deg tilt (CSS) */

/**
 * Orbiting image ring.
 *
 * The analysed original rode 4 panels on a rotating cylinder, so each panel
 * turned with the cylinder and went edge-on (and backwards) as it passed the
 * sides. This build instead ORBITS the panels while keeping every one of them
 * facing the viewer — the picture plane stays flat, only its position and depth
 * change. Ten panels, and depth drives scale, opacity and stacking, which is
 * what gives the ring its dimensional read.
 *
 * Kept from the measurement: linear easing, infinite repeat, the list's static
 * 5° tilt, and the entrance at ScrollTrigger `top 80%`.
 * Changed: panel count (4 -> 10) and duration (30s -> 14s). See MISSING.md > C5.
 */
function carousel3D(gsap, ScrollTrigger) {
  const stage = document.querySelector('.img-carousel__stage')
  if (!stage) return

  const panels = Array.from(stage.querySelectorAll('.img-carousel__panel'))
  if (!panels.length) return

  const n = panels.length
  const SPIN_SECONDS = 14

  // Radius from panel count: n panels of width w just touch at (w/2)/tan(180/n);
  // the factor opens a gap so the ring reads as a ring.
  const radius = () => ((stage.clientWidth || 180) / 2 / Math.tan(Math.PI / n)) * 1.35

  const angle = { v: 0 }

  function render() {
    const r = radius()
    for (let i = 0; i < n; i++) {
      const a = ((angle.v + (i * 360) / n) * Math.PI) / 180
      const x = Math.sin(a) * r
      const z = Math.cos(a) * r - r // front of the ring sits at z = 0

      // depth 1 at the front, 0 at the back
      const depth = (z + 2 * r) / (2 * r)
      const scale = 0.55 + depth * 0.45
      const opacity = 0.18 + depth * 0.82

      panels[i].style.transform = `translate3d(${x}px, 0px, ${z}px) scale(${scale})`
      panels[i].style.opacity = String(opacity)
      panels[i].style.zIndex = String(Math.round(depth * 100))
    }
  }

  render()
  window.addEventListener('resize', render)

  const spin = gsap.to(angle, {
    v: 360,
    duration: SPIN_SECONDS,
    ease: 'none',
    repeat: -1,
    paused: true,
    onUpdate: render,
  })

  gsap.from(stage, {
    scale: 0.72,
    opacity: 0,
    duration: D.carouselIn,
    ease: 'expo.inOut',
    scrollTrigger: {
      trigger: '.img-carousel__list',
      start: 'top 80%',
      end: 'bottom top',
      toggleActions: 'play resume play play',
      onEnter: () => spin.play(),
      onLeave: () => spin.pause(),
      onEnterBack: () => spin.play(),
      onLeaveBack: () => spin.pause(),
    },
  })

  // Slow down on hover so a panel can be looked at — ours, no measurement for it.
  const wrap = document.querySelector('.img-carousel')
  if (wrap) {
    wrap.addEventListener('pointerenter', () => spin.timeScale(0.25))
    wrap.addEventListener('pointerleave', () => spin.timeScale(1))
  }
}

/* ── scroll-scrubbed scaling video ──────────────────────────────────────────
   Measured: width 292 -> 1257, height 164 -> 707, x 0 -> -482, y 0 -> +400,
   ease none, scrub true, trigger .scaling-video__wrapper.
   The x is measured but deliberately not used — see targets() below.

   The report logs start AND end as 'center center' (a zero-length range, which
   is how ScrollTrigger reported it) while the 21-stop sweep shows the expansion
   actually running between ~85% and 100% of page scroll. We reproduce the
   OBSERVED behaviour: it grows as the section crosses the middle of the viewport.

   The measured pixel targets are treated as ratios against the analysed 1315x674
   viewport so the geometry survives at any window size. */

function scalingVideo(gsap, ScrollTrigger) {
  const wrapper = document.querySelector('.scaling-video__wrapper')
  const el = document.querySelector('.scaling-video')
  if (!wrapper || !el) return

  const REF_W = 1315
  const REF_H = 674
  const M = { w0: 292, h0: 164, w1: 1257, h1: 707, x1: -482, y1: 400 }

  function targets() {
    const sx = window.innerWidth / REF_W
    const sy = window.innerHeight / REF_H
    return {
      w0: M.w0 * sx,
      h0: M.h0 * sx,
      w1: Math.min(M.w1 * sx, window.innerWidth * 0.96),
      h1: Math.min(M.h1 * sx, window.innerHeight * 1.05),
      /* The measured x of -482 is NOT reproduced — decided, see MISSING.md > C.
         It only reads as a drift on the reference viewport, where the grown clip
         still left ~60px of margin either side. Here the width cap takes it to
         96vw, so any horizontal translate simply pushes one edge off-screen and
         leaves a gap on the other. Held centred; the vertical drift stays. */
      x1: 0,
      y1: M.y1 * sy * 0.28,
    }
  }

  let t = targets()

  gsap.set(el, { width: t.w0, height: t.h0 })

  const tween = gsap.fromTo(
    el,
    { width: t.w0, height: t.h0, x: 0, y: 0 },
    {
      width: () => t.w1,
      height: () => t.h1,
      x: () => t.x1,
      y: () => t.y1,
      ease: 'none',
      scrollTrigger: {
        trigger: wrapper,
        start: 'top 85%',
        end: 'center center',
        scrub: true,
        invalidateOnRefresh: true,
      },
    }
  )

  ScrollTrigger.addEventListener('refreshInit', () => {
    t = targets()
    gsap.set(el, { width: t.w0, height: t.h0 })
  })

  return tween
}

/* ── use-case wall — seamless sliding cycle ─────────────────────────────────
   Each cell holds N cards and cycles them forever. The cards are laid out edge
   to edge (card k parked at xPercent k*100) and the whole set slides by -100 per
   step, so at every instant the outgoing and incoming cards TOGETHER cover the
   cell — the backdrop is never exposed.

   The wrap modifier is what makes it endless: the moment a card passes fully off
   the left it is remapped to the right edge and slides back in. No card is ever
   faded, hidden, or repositioned by hand.

   This replaced a flip-and-fade swap driven by a setInterval. That version built
   a throwaway timeline per swap, and any interruption between its "fade out" and
   "fade in" halves left BOTH cards at opacity 0 — an empty cell showing the page
   background. Observed on three of eight cells. Here GSAP owns the whole state
   machine: one repeating timeline per cell, so pausing merely freezes a slide
   mid-way, with both cards still covering the cell. */

function logoWall(gsap, ScrollTrigger) {
  const wall = document.querySelector('.logo-wall')
  if (!wall) return

  const cells = Array.from(wall.querySelectorAll('.logo-wall__cell'))
    .map((cell) => Array.from(cell.querySelectorAll('.logo-card')))
    .filter((cards) => cards.length > 1)
  if (!cells.length) return

  /* Slow travel on a sine ease. `sine.inOut` is the gentlest of the standard
     curves — it leaves and arrives with almost no acceleration spike, so the card
     appears to drift rather than snap. A shorter, punchier slide made the wall
     read as eight things flicking at you.

     SLIDE deliberately exceeds the per-cell stagger below (cycle/8 ≈ 1.2s), so
     transitions overlap slightly and something is always gently in motion —
     which is what makes the wall feel continuous instead of metronomic. */
  const HOLD = 3.4   // dwell on a card before the next slide
  const SLIDE = 1.5  // travel time

  /* Four travel directions, dealt out across the grid so neighbouring cells never
     move the same way. The wall then reads as a set of independent panels rather
     than one sheet sliding sideways. Order is hand-set, not `i % 4`: with four
     columns, a plain modulo gives every column a single direction and the whole
     thing stripes. */
  const DIRECTIONS = [
    { axis: 'xPercent', step: '-=100' }, // left
    { axis: 'yPercent', step: '-=100' }, // up
    { axis: 'xPercent', step: '+=100' }, // right
    { axis: 'yPercent', step: '+=100' }, // down
    { axis: 'yPercent', step: '-=100' }, // up
    { axis: 'xPercent', step: '+=100' }, // right
    { axis: 'yPercent', step: '+=100' }, // down
    { axis: 'xPercent', step: '-=100' }, // left
  ]

  const loops = cells.map((cards, i) => {
    const n = cards.length
    const dir = DIRECTIONS[i % DIRECTIONS.length]

    /* Every card is visible from here on — the cell's overflow does the hiding.
       This overrides the CSS resting state, which only had to cover the moment
       before GSAP ran. Both axes are zeroed first so the parked offset lands on
       the travel axis only. */
    gsap.set(cards, { autoAlpha: 1, xPercent: 0, yPercent: 0 })
    gsap.set(cards, { [dir.axis]: (k) => k * 100 })

    const tl = gsap.timeline({ repeat: -1, paused: true, defaults: { duration: SLIDE, ease: 'sine.inOut' } })
    for (let k = 0; k < n; k++) {
      tl.to(
        cards,
        {
          [dir.axis]: dir.step,
          /* Range is [-100, (n-1)*100): a card only jumps once it is completely
             off-frame, so the wrap is never visible. Works in both directions —
             wrap folds whichever end the card runs off. */
          modifiers: { [dir.axis]: gsap.utils.wrap(-100, (n - 1) * 100) },
        },
        `+=${HOLD}`
      )
    }
    return tl
  })

  // Spread the cells across one full cycle so the grid never flips in unison.
  const cycle = cells[0].length * (HOLD + SLIDE)
  loops.forEach((tl, i) => tl.time((cycle / loops.length) * i))

  // Only runs while the wall is on screen.
  ScrollTrigger.create({
    trigger: wall,
    start: 'top 90%',
    end: 'bottom top',
    onToggle: (self) => loops.forEach((tl) => (self.isActive ? tl.play() : tl.pause())),
  })
}

/* ── hero keyword cycle ─────────────────────────────────────────────────────
   The live hero scrambles between keywords ("immersive", "interactive",
   "real-time", "Cinetica"). Cadence was not timed, so the 3.2s dwell is ours. */

function heroWordCycle(gsap) {
  const el = document.querySelector('[data-word-cycle]')
  if (!el) return

  const words = JSON.parse(el.dataset.wordCycle || '[]')
  if (words.length < 2) return

  let i = 0
  gsap.delayedCall(2.4, function next() {
    i = (i + 1) % words.length
    gsap.to(el, {
      duration: D.scramble * 0.7,
      scrambleText: { text: words[i], chars: SCRAMBLE_CHARS, speed: 0.9 },
      ease: 'none',
    })
    gsap.delayedCall(3.2, next)
  })
}
