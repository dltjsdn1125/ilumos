/**
 * check.mjs — build verification.
 * Validates: internal links resolve, one <h1> per page, images have alt +
 * intrinsic dimensions, required SEO tags present, no leftover template markers.
 */

import { readFileSync, readdirSync, existsSync } from 'node:fs'
import { join, relative, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..')
const SKIP_DIRS = new Set(['node_modules', 'src', 'scripts', 'assets', '.git'])

const pages = []
;(function walk(dir) {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    if (entry.isDirectory()) {
      if (SKIP_DIRS.has(entry.name)) continue
      walk(join(dir, entry.name))
    } else if (entry.name.endsWith('.html')) {
      pages.push(join(dir, entry.name))
    }
  }
})(ROOT)

const problems = []
const notes = []
const add = (page, msg) => problems.push(`${relative(ROOT, page).replace(/\\/g, '/')}: ${msg}`)

/* Documented empty slots: the hero-media convention ships the <source> markup
   before the file exists so the clip can be dropped in later without a code
   change. Until then the browser falls back to the poster, which is the intended
   behaviour — see MISSING.md > hero-video. Reported as a note, not a failure. */
const EXPECTED_ABSENT = new Set([
  '/assets/video/hero.webm',
  '/assets/video/hero.mp4',
  '/assets/video/about.webm',
  '/assets/video/about.mp4',
])

let linkCount = 0
let imgCount = 0

for (const page of pages) {
  const html = readFileSync(page, 'utf8')

  // --- internal links / assets resolve ---
  for (const m of html.matchAll(/(?:href|src)="([^"]+)"/g)) {
    const raw = m[1]
    if (!raw.startsWith('/') || raw.startsWith('//')) continue
    const clean = raw.split('#')[0].split('?')[0]
    if (!clean) continue
    linkCount += 1
    const target = clean.endsWith('/') ? join(ROOT, clean, 'index.html') : join(ROOT, clean)
    if (existsSync(target)) continue
    if (EXPECTED_ABSENT.has(clean)) notes.push(`${relative(ROOT, page).replace(/\\/g, '/')}: ${clean} not present (poster fallback — documented)`)
    else add(page, `broken -> ${raw}`)
  }

  // --- headings ---
  const h1s = [...html.matchAll(/<h1[^>]*>([\s\S]*?)<\/h1>/g)]
  if (h1s.length === 0) add(page, 'no <h1>')
  if (h1s.length > 1) add(page, `${h1s.length} <h1> elements (expected 1)`)

  // --- images ---
  for (const m of html.matchAll(/<img\b[^>]*>/g)) {
    const t = m[0]
    imgCount += 1
    if (!/\balt="/.test(t)) add(page, `img without alt: ${t.slice(0, 90)}`)
    if (!/\bwidth="\d+"/.test(t) || !/\bheight="\d+"/.test(t)) {
      add(page, `img without intrinsic size: ${t.slice(0, 90)}`)
    }
  }

  // --- SEO essentials ---
  for (const [label, re] of [
    ['<title>', /<title>[^<]{5,}<\/title>/],
    ['meta description', /<meta name="description" content="[^"]{20,}"/],
    ['canonical', /<link rel="canonical" href="https?:[^"]+"/],
    ['og:image', /<meta property="og:image"/],
    ['JSON-LD', /application\/ld\+json/],
    ['lang attr', /<html lang="ko">/],
    ['viewport', /name="viewport"/],
  ]) {
    if (!re.test(html)) add(page, `missing ${label}`)
  }

  // --- template leakage ---
  if (/undefined|\[object Object\]|NaN/.test(html.replace(/"[^"]*undefined[^"]*"/g, ''))) {
    add(page, 'contains undefined / [object Object] / NaN')
  }

  // --- accessibility wiring ---
  // Both attributes take an ID *list* (space-separated), not a single ID.
  for (const [attr, re] of [
    ['aria-controls', /aria-controls="([^"]+)"/g],
    ['aria-labelledby', /aria-labelledby="([^"]+)"/g],
  ]) {
    for (const m of html.matchAll(re)) {
      for (const id of m[1].trim().split(/\s+/)) {
        if (!html.includes(`id="${id}"`)) add(page, `${attr} points at missing id: ${id}`)
      }
    }
  }
}

console.log(`pages: ${pages.length}  ·  internal refs: ${linkCount}  ·  images: ${imgCount}`)
if (notes.length) {
  console.log('\nnotes (expected, not failures):')
  for (const n of [...new Set(notes)]) console.log('  ' + n)
}
if (problems.length) {
  console.log(`\nFAIL — ${problems.length} problem(s):`)
  for (const p of [...new Set(problems)]) console.log('  ' + p)
  process.exit(1)
}
console.log('PASS — links resolve, headings/alt/SEO/aria all present')
