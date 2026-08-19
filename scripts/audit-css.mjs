/**
 * audit-css.mjs — report CSS class selectors that no built page ever emits.
 *
 * Not a linter and not run by the build: dead rules are cheap, and a few
 * deliberately-unused design-system utilities are fine to keep. This exists so
 * "is anything left over from a refactor?" is answerable with evidence instead
 * of a guess.
 *
 *   node scripts/audit-css.mjs
 */

import { readFileSync, readdirSync } from 'node:fs'
import { join, dirname, relative } from 'node:path'
import { fileURLToPath } from 'node:url'

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..')
const SKIP = new Set(['node_modules', 'src', 'scripts', 'assets', '.git'])

/* Utilities that are part of the design-system surface rather than any one
   component. Unused today is not a defect — they exist to be reached for. */
const KEEP = [
  /^u-/, /^container--/, /^grid--/, /^stack--/, /^padding-section--/,
  /^media--/, /^button--/, /^theme-/, /^is-/, /^has-/, /^section--/,
  /^lenis/, /^field--/, /^pager__item--/, /^hairline$/,
]

// ── collect built HTML ──────────────────────────────────────────────────────
let html = ''
;(function walk(dir) {
  for (const e of readdirSync(dir, { withFileTypes: true })) {
    if (e.isDirectory()) {
      if (SKIP.has(e.name)) continue
      walk(join(dir, e.name))
    } else if (e.name.endsWith('.html')) {
      html += readFileSync(join(dir, e.name), 'utf8')
    }
  }
})(ROOT)

// Classes actually present in the markup, plus classes any script can add.
const emitted = new Set()
for (const m of html.matchAll(/class="([^"]*)"/g)) {
  for (const c of m[1].trim().split(/\s+/)) if (c) emitted.add(c)
}

let js = ''
for (const f of readdirSync(join(ROOT, 'assets', 'js'))) {
  if (f.endsWith('.js')) js += readFileSync(join(ROOT, 'assets', 'js', f), 'utf8')
}
for (const m of js.matchAll(/['"`]([a-z][a-z0-9-]*(?:__[a-z0-9-]+)?(?:--[a-z0-9-]+)?)['"`]/gi)) {
  emitted.add(m[1])
}

// ── collect declared class selectors ────────────────────────────────────────
const declared = new Map() // class -> file
for (const f of readdirSync(join(ROOT, 'assets', 'css'))) {
  if (!f.endsWith('.css')) continue
  const css = readFileSync(join(ROOT, 'assets', 'css', f), 'utf8')
  // Strip comments AND url() payloads before scanning. An inlined SVG data URI
  // carries "www.w3.org", which would otherwise be reported as .w3 and .org.
  const clean = css
    .replace(/\/\*[\s\S]*?\*\//g, '')
    .replace(/url\((?:[^()]|\([^()]*\))*\)/g, 'url()')
  for (const m of clean.matchAll(/\.(-?[_a-zA-Z][\w-]*)/g)) {
    if (!declared.has(m[1])) declared.set(m[1], f)
  }
}

const dead = [...declared.entries()]
  .filter(([c]) => !emitted.has(c))
  .filter(([c]) => !KEEP.some((re) => re.test(c)))
  .sort((a, b) => a[1].localeCompare(b[1]) || a[0].localeCompare(b[0]))

console.log(`declared classes: ${declared.size}  ·  emitted in markup/js: ${emitted.size}`)
if (!dead.length) {
  console.log('no dead component classes')
} else {
  console.log(`\n${dead.length} class(es) declared but never emitted:`)
  for (const [c, f] of dead) console.log(`  .${c.padEnd(28)} ${f}`)
}
