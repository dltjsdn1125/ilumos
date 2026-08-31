/* ==========================================================================
   shop.js — 자사몰 동작 (분류 · 정렬 · 장바구니 · 주문서)

   전제 하나만 붙잡고 씁니다: **서버가 없습니다.**

   그래서 카탈로그 20종은 전부 정적 HTML 로 먼저 나가 있고, 이 파일은 이미
   그려진 것을 숨기거나 순서만 바꿉니다. 자바스크립트가 꺼져 있어도 제품과
   가격은 그대로 보이고 상세 페이지로도 들어갑니다 — 사라지는 것은 분류·정렬과
   장바구니뿐입니다.

   장바구니는 localStorage 에만 삽니다. 즉 이 브라우저 밖으로 나가지 않고,
   다른 기기에서는 보이지 않습니다. 그래서 저장하는 값은 {slug, qty} 뿐이고
   이름·가격은 담을 때가 아니라 **볼 때** 카탈로그에서 다시 읽습니다. 한 달 전
   장바구니가 옛 가격을 들고 있는 일이 없습니다.

   주문서 역시 결제하지 않습니다. 문의 폼과 같은 방식으로 메일 초안을 만들어
   넘깁니다(ui.js > initContactForm). 정적 호스팅에서 실제로 동작하는 유일한
   경로이고, 받지 않은 주문을 받았다고 말하지 않는 유일한 방법입니다.
   ========================================================================== */

const KEY = 'ilumos-cart-v1'
const MAX_QTY = 99

/* 빌드(components.mjs > won)와 **같은 문자열**을 만들어야 합니다. 화면에 이미
   찍혀 있는 가격과 자바스크립트가 계산한 합계가 다른 규칙으로 찍히면 그 자체가
   버그로 읽힙니다. 그래서 Intl 로케일 데이터에 기대지 않습니다. */
const won = (n) => String(Math.round(Number(n) || 0)).replace(/\B(?=(\d{3})+(?!\d))/g, ',')

const clampQty = (n) => Math.min(MAX_QTY, Math.max(1, Math.round(Number(n) || 1)))

/* ── 저장소 ─────────────────────────────────────────────────────────────────
   사파리 프라이빗 모드나 저장소 차단 설정에서는 localStorage 접근 자체가
   throw 합니다. 장바구니가 안 되는 것과 페이지 전체가 죽는 것은 전혀 다른
   문제라, 읽기·쓰기를 모두 감쌉니다. 실패하면 이번 세션만 비어 있는 장바구니로
   동작합니다. */

function readCart() {
  try {
    const raw = localStorage.getItem(KEY)
    if (!raw) return []
    const parsed = JSON.parse(raw)
    if (!Array.isArray(parsed)) return []
    return parsed
      .filter((it) => it && typeof it.slug === 'string')
      .map((it) => ({ slug: it.slug, qty: clampQty(it.qty) }))
  } catch {
    return []
  }
}

function writeCart(items) {
  try {
    localStorage.setItem(KEY, JSON.stringify(items))
  } catch {
    /* 저장은 못 해도 이번 화면은 계속 동작합니다. */
  }
  document.dispatchEvent(new CustomEvent('cart:change', { detail: { items } }))
}

const cartCount = (items) => items.reduce((n, it) => n + it.qty, 0)

function addToCart(slug, qty = 1) {
  const items = readCart()
  const found = items.find((it) => it.slug === slug)
  if (found) found.qty = clampQty(found.qty + qty)
  else items.push({ slug, qty: clampQty(qty) })
  writeCart(items)
  return items
}

/* ── 헤더 배지 ────────────────────────────────────────────────────────────
   서버는 담긴 수량을 모르므로 배지는 `hidden` 인 채로 나갑니다. 여기서 값을
   채우면서 벗기기 때문에, 비어 있을 때 "0" 이 깜빡였다 사라지지 않습니다. */

export function initCartBadge() {
  const paint = () => {
    const n = cartCount(readCart())
    document.querySelectorAll('[data-cart-count]').forEach((el) => {
      el.textContent = String(n)
      el.hidden = n === 0
    })
    document.querySelectorAll('[data-cart-link]').forEach((el) => {
      el.setAttribute('aria-label', n ? `장바구니 — ${n}개 담김` : '장바구니 — 비어 있음')
    })
  }

  paint()
  document.addEventListener('cart:change', paint)
  /* 다른 탭에서 담았을 때도 맞춥니다. 같은 사이트를 두 탭에 열어 두는 일은
     장보기에서 아주 흔합니다. */
  window.addEventListener('storage', (e) => {
    if (e.key === KEY) paint()
  })
}

/* ── 수량 스테퍼 ──────────────────────────────────────────────────────────
   input 이 진짜 상태이고 버튼은 그 값을 바꾸기만 합니다. 위임으로 한 번만
   묶으므로 장바구니처럼 나중에 생기는 줄에도 그대로 걸립니다. */

export function initQtySteppers() {
  document.addEventListener('click', (e) => {
    const btn = e.target.closest('[data-qty-step]')
    if (!btn) return
    const box = btn.closest('[data-qty]')
    const input = box && box.querySelector('.qty__input')
    if (!input) return
    input.value = String(clampQty(Number(input.value) + Number(btn.dataset.qtyStep)))
    input.dispatchEvent(new Event('change', { bubbles: true }))
  })

  document.addEventListener('change', (e) => {
    if (!e.target.classList || !e.target.classList.contains('qty__input')) return
    e.target.value = String(clampQty(e.target.value))
  })
}

/* ── 담기 ───────────────────────────────────────────────────────────────── */

export function initAddToCart() {
  document.addEventListener('click', (e) => {
    const btn = e.target.closest('[data-add-to-cart]')
    if (!btn) return
    e.preventDefault()

    /* 상세 페이지의 담기 버튼만 수량 스테퍼를 함께 봅니다. 카탈로그 카드의
       담기는 언제나 1개 — 목록에서 수량을 정하게 하면 스무 개의 입력칸이
       생기고, 그 값을 아무도 다시 확인하지 않습니다. */
    let qty = 1
    if (btn.hasAttribute('data-qty-source')) {
      const input = btn.closest('.product__panel, .product__buy')?.querySelector('.qty__input')
      if (input) qty = clampQty(input.value)
    }

    const items = addToCart(btn.dataset.addToCart, qty)

    btn.classList.add('is-added')
    setTimeout(() => btn.classList.remove('is-added'), 900)

    const status = document.querySelector('[data-add-status]')
    if (status) {
      status.textContent = `장바구니에 담았습니다 (총 ${cartCount(items)}개).`
      setTimeout(() => { status.textContent = '' }, 4000)
    }
  })
}

/* ── 카탈로그: 분류 · 정렬 ────────────────────────────────────────────────
   DOM 을 다시 만들지 않습니다. 카드에 붙은 GSAP 인라인 스타일과 ScrollTrigger
   위치가 그대로 살아 있어야 하므로, 걸러내기는 클래스 하나로, 정렬은 CSS
   `order` 로만 합니다. */

export function initShopFilter() {
  const grid = document.querySelector('[data-product-grid]')
  if (!grid) return

  const cards = Array.from(grid.querySelectorAll('[data-product]'))
  const chips = Array.from(document.querySelectorAll('[data-filter]'))
  const sort = document.querySelector('[data-shop-sort]')
  const countEl = document.querySelector('[data-shop-count]')
  const emptyEl = document.querySelector('[data-shop-empty]')

  let filter = 'all'

  const comparators = {
    default: (a, b) => Number(a.dataset.order) - Number(b.dataset.order),
    'price-asc': (a, b) => Number(a.dataset.price) - Number(b.dataset.price),
    'price-desc': (a, b) => Number(b.dataset.price) - Number(a.dataset.price),
    name: (a, b) => a.dataset.name.localeCompare(b.dataset.name, 'ko'),
  }

  function applyView() {
    const visible = cards.filter((c) => filter === 'all' || c.dataset.category === filter)
    const hidden = cards.filter((c) => !visible.includes(c))

    visible.forEach((c) => c.classList.remove('is-filtered'))
    hidden.forEach((c) => c.classList.add('is-filtered'))

    const cmp = comparators[(sort && sort.value) || 'default'] || comparators.default
    visible.slice().sort(cmp).forEach((c, i) => { c.style.order = String(i) })

    if (countEl) {
      const label = filter === 'all' ? '전체' : chips.find((c) => c.dataset.filter === filter)?.textContent.trim().split('\n')[0] || ''
      countEl.textContent = `${label} ${visible.length}개 제품`
    }
    if (emptyEl) emptyEl.hidden = visible.length > 0

    /* 카드가 사라지면 페이지 높이가 바뀝니다. 다시 재지 않으면 아래쪽 섹션의
       스크롤 트리거가 옛 좌표에 묶여 영영 재생되지 않습니다. */
    window.ScrollTrigger?.refresh()
  }

  chips.forEach((chip) => {
    chip.addEventListener('click', () => {
      filter = chip.dataset.filter
      chips.forEach((c) => {
        const on = c === chip
        c.classList.toggle('is-active', on)
        c.setAttribute('aria-pressed', String(on))
      })
      applyView()
    })
  })

  if (sort) sort.addEventListener('change', applyView)
}

/* ── 장바구니 화면 ──────────────────────────────────────────────────────── */

function readJson(id, fallback) {
  const el = document.getElementById(id)
  if (!el) return fallback
  try {
    return JSON.parse(el.textContent)
  } catch {
    return fallback
  }
}

export function initCartPage() {
  const root = document.querySelector('[data-cart]')
  if (!root) return

  const list = root.querySelector('[data-cart-list]')
  const empty = root.querySelector('[data-cart-empty]')
  const summary = root.querySelector('[data-cart-summary]')
  const orderSection = document.querySelector('[data-order-section]')
  const template = document.getElementById('cart-row-template')

  const catalogue = readJson('shop-catalogue', [])
  const shipping = readJson('shop-shipping', { flat: 0, freeOver: Infinity })
  const bySlug = new Map(catalogue.map((p) => [p.slug, p]))

  /* 카탈로그에 없는 slug 는 버립니다 — 단종되거나 이름이 바뀐 뒤에도 남아 있던
     오래된 장바구니가 빈 줄로 나타나지 않게. */
  const liveItems = () => readCart().filter((it) => bySlug.has(it.slug))

  function render() {
    const items = liveItems()
    list.textContent = ''

    for (const item of items) {
      const p = bySlug.get(item.slug)
      const row = template.content.firstElementChild.cloneNode(true)

      row.dataset.slug = p.slug

      const img = row.querySelector('img')
      img.src = p.src
      img.alt = `${p.name} 제품 이미지`
      img.width = p.w
      img.height = p.h

      row.querySelector('.cart-row__cat').textContent = p.category
      const link = row.querySelector('.cart-row__title a')
      link.href = `/shop/${p.slug}/`
      link.textContent = p.name
      row.querySelector('.cart-row__meta').textContent = `${p.sku} · ${p.pack} · ${won(p.price)}원 / ${p.unit}`

      /* 템플릿 안의 id 는 스무 줄이 복제되면 그대로 중복됩니다. label 과 함께
         줄마다 고유한 값으로 바꿔 줍니다. */
      const input = row.querySelector('.qty__input')
      const label = row.querySelector('label[for="cart-qty-template"]')
      const id = `cart-qty-${p.slug}`
      input.id = id
      if (label) {
        label.setAttribute('for', id)
        label.textContent = `${p.name} 수량`
      }
      input.value = String(item.qty)

      row.querySelector('.cart-row__price').textContent = `${won(p.price * item.qty)}원`
      row.querySelector('[data-cart-remove]').setAttribute('aria-label', `${p.name} 삭제`)

      list.appendChild(row)
    }

    const goods = items.reduce((sum, it) => sum + bySlug.get(it.slug).price * it.qty, 0)
    const ship = goods === 0 || goods >= shipping.freeOver ? 0 : shipping.flat

    const has = items.length > 0
    if (empty) empty.hidden = has
    if (summary) summary.hidden = !has
    if (orderSection) orderSection.hidden = !has

    const set = (sel, text) => {
      const el = root.querySelector(sel)
      if (el) el.textContent = text
    }
    set('[data-sum-goods]', `${won(goods)}원`)
    set('[data-sum-shipping]', ship === 0 ? (has ? '무료' : '0원') : `${won(ship)}원`)
    set('[data-sum-total]', `${won(goods + ship)}원`)

    window.ScrollTrigger?.refresh()
  }

  /* 수량 변경 · 삭제 · 비우기 — 전부 저장소를 고치고 다시 그립니다. 화면과
     저장소가 각자 상태를 갖지 않게 하는 것이 요점입니다. */
  list.addEventListener('change', (e) => {
    if (!e.target.classList.contains('qty__input')) return
    const row = e.target.closest('[data-cart-row]')
    if (!row) return
    const items = readCart()
    const found = items.find((it) => it.slug === row.dataset.slug)
    if (found) found.qty = clampQty(e.target.value)
    writeCart(items)
  })

  list.addEventListener('click', (e) => {
    const btn = e.target.closest('[data-cart-remove]')
    if (!btn) return
    const row = btn.closest('[data-cart-row]')
    if (!row) return
    writeCart(readCart().filter((it) => it.slug !== row.dataset.slug))
  })

  const clear = root.querySelector('[data-cart-clear]')
  if (clear) {
    clear.addEventListener('click', () => {
      /* confirm() 은 Lenis 스크롤과 커스텀 커서를 얼려 버려서 쓰지 않습니다.
         대신 한 번 더 누르게 합니다 — 되돌릴 수 있는 동작이라 이 정도면 충분. */
      if (clear.dataset.armed === 'true') {
        writeCart([])
        clear.dataset.armed = 'false'
        clear.textContent = '장바구니 비우기'
        return
      }
      clear.dataset.armed = 'true'
      clear.textContent = '한 번 더 누르면 비웁니다'
      setTimeout(() => {
        clear.dataset.armed = 'false'
        clear.textContent = '장바구니 비우기'
      }, 4000)
    })
  }

  document.addEventListener('cart:change', render)
  window.addEventListener('storage', (e) => { if (e.key === KEY) render() })
  render()

  initOrderForm(bySlug, shipping)
}

/* ── 주문서 ─────────────────────────────────────────────────────────────────
   문의 폼과 같은 규칙(ui.js > initContactForm)입니다. 검증은 브라우저에서
   끝까지 하고, 마지막에 담긴 내용을 본문에 적은 메일 초안을 엽니다. */

function initOrderForm(bySlug, shipping) {
  const form = document.getElementById('order-form')
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
    phone: (v) => (/^[0-9+\-\s()]{7,}$/.test(v.trim()) ? '' : '연락 가능한 번호를 입력해 주세요.'),
    email: (v) => (/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(v.trim()) ? '' : '이메일 형식을 확인해 주세요.'),
    address: (v) => (v.trim().length >= 5 ? '' : '배송지 주소를 입력해 주세요.'),
    consent: (v, el) => (el.checked ? '' : '개인정보 수집·이용 동의가 필요합니다.'),
  }

  const validateField = (name) => {
    const el = form.elements[name]
    if (!el || !RULES[name]) return true
    const msg = RULES[name](el.value ?? '', el)
    showError(name, msg)
    return !msg
  }

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

  form.addEventListener('submit', (e) => {
    e.preventDefault()
    setStatus('')

    const items = readCart().filter((it) => bySlug.has(it.slug))
    if (!items.length) {
      setStatus('장바구니가 비어 있습니다.', 'error')
      return
    }

    const names = Object.keys(RULES)
    const ok = names.map(validateField).every(Boolean)
    if (!ok) {
      const firstBad = names.find((n) => form.elements[n]?.getAttribute('aria-invalid') === 'true')
      form.elements[firstBad]?.focus()
      setStatus('입력을 확인해 주세요.', 'error')
      return
    }

    const data = Object.fromEntries(new FormData(form).entries())
    const goods = items.reduce((sum, it) => sum + bySlug.get(it.slug).price * it.qty, 0)
    const ship = goods >= shipping.freeOver ? 0 : shipping.flat

    const lines = items.map((it) => {
      const p = bySlug.get(it.slug)
      return `· ${p.name} (${p.sku}) — ${it.qty}${p.unit} × ${won(p.price)}원 = ${won(p.price * it.qty)}원`
    })

    const body = [
      `상호 · 사용처: ${data.company || ''}`,
      `담당자: ${data.name || ''}`,
      `연락처: ${data.phone || ''}`,
      `이메일: ${data.email || ''}`,
      `배송지: ${data.address || ''}`,
      '',
      '주문 품목',
      ...lines,
      '',
      `상품 합계: ${won(goods)}원`,
      `배송비: ${ship === 0 ? '무료' : `${won(ship)}원`}`,
      `합계: ${won(goods + ship)}원`,
      '',
      '요청 사항',
      data.note || '-',
      '',
      '※ 홈페이지 주문서로 작성된 내용입니다. 표시 금액은 홈페이지 예시가 기준이며 확정 견적으로 회신 부탁드립니다.',
    ].join('\n')

    const subject = `[주문서] ${data.company || ''} — ${items.length}개 품목`
    window.location.href = `mailto:${form.dataset.mailto || ''}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`
    setStatus('메일 작성 창을 열었습니다. 전송까지 완료해 주세요.', 'ok')
  })
}

/* ── 부팅 ───────────────────────────────────────────────────────────────── */

export function initShop() {
  initCartBadge()
  initQtySteppers()
  initAddToCart()
  initShopFilter()
  initCartPage()
}
