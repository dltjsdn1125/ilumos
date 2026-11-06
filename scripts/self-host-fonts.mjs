/**
 * self-host-fonts.mjs — Google Fonts 의 IBM Plex 를 자기 서버(assets/fonts)로 옮깁니다.
 * 사이트 글(HTML·데이터·스크립트)에 실제로 나오는 글자가 들어 있는 조각(unicode-range)만
 * 내려받고, 같은 모양의 @font-face 를 assets/css/fonts.css 로 씁니다. 한 번만 돌리면 됩니다.
 *   node scripts/self-host-fonts.mjs
 * 글을 크게 바꿔 새 글자가 생기면 다시 돌립니다.
 */
import { readFile, writeFile, mkdir, readdir } from 'node:fs/promises'
import { join, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..')
const UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36'
// IBM Plex Sans(라틴) 는 Sans KR 이 같은 글자를 이미 갖고 있어 받지 않습니다.
const CSS_URL =
  'https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+KR:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&family=IBM+Plex+Serif:ital,wght@0,400;0,600;1,400&display=swap'

async function walk(dir, out = []) {
  for (const e of await readdir(dir, { withFileTypes: true })) {
    if (['node_modules', '.git', 'HM-BACKUP', 'HM-SNAPSHOTS', 'export', 'fonts', 'images', 'vendor', 'HM-REQUIREMENTS', '_threads'].includes(e.name)) continue
    const p = join(dir, e.name)
    if (e.isDirectory()) await walk(p, out)
    else if (/\.(html|mjs|js|css|json)$/.test(e.name)) out.push(p)
  }
  return out
}

const used = new Set()
for (const f of await walk(ROOT)) {
  if (f.includes('docs')) continue
  for (const ch of await readFile(f, 'utf8')) used.add(ch.codePointAt(0))
}

function inRange(rangeText) {
  return rangeText.split(',').some((r) => {
    const m = r.trim().match(/^U\+([0-9A-Fa-f?]+)(?:-([0-9A-Fa-f]+))?$/)
    if (!m) return false
    let lo, hi
    if (m[1].includes('?')) { lo = parseInt(m[1].replace(/\?/g, '0'), 16); hi = parseInt(m[1].replace(/\?/g, 'F'), 16) }
    else { lo = parseInt(m[1], 16); hi = m[2] ? parseInt(m[2], 16) : lo }
    for (const c of used) if (c >= lo && c <= hi) return true
    return false
  })
}

const css = await (await fetch(CSS_URL, { headers: { 'User-Agent': UA } })).text()
// 한글 조각은 앞에 주석이 없어서, 주석은 있어도 없어도 됩니다.
const faces = [...css.matchAll(/(?:\/\*\s*([^*]+?)\s*\*\/\s*)?@font-face\s*\{([^}]+)\}/g)].map((m, i) => {
  const b = m[2]
  const g = (re) => (b.match(re) || [])[1]
  return {
    subset: m[1] ? m[1].replace(/[^\w-]/g, '') : `ko${i}`,
    family: g(/font-family:\s*'([^']+)'/),
    style: g(/font-style:\s*(\w+)/),
    weight: g(/font-weight:\s*(\d+)/),
    url: g(/url\(([^)]+)\)/),
    range: g(/unicode-range:\s*([^;]+);/),
  }
})

await mkdir(join(ROOT, 'assets', 'fonts'), { recursive: true })
const slug = (s) => s.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '')
const rules = []
let n = 0, bytes = 0
for (const f of faces) {
  if (!f.range || !inRange(f.range)) continue
  const file = `${slug(f.family)}-${f.weight}${f.style === 'italic' ? 'i' : ''}-${slug(f.subset)}.woff2`
  const buf = Buffer.from(await (await fetch(f.url, { headers: { 'User-Agent': UA } })).arrayBuffer())
  await writeFile(join(ROOT, 'assets', 'fonts', file), buf)
  bytes += buf.length
  n += 1
  rules.push(`@font-face {
  font-family: '${f.family}';
  font-style: ${f.style};
  font-weight: ${f.weight};
  font-display: swap;
  src: url(/assets/fonts/${file}) format('woff2');
  unicode-range: ${f.range};
}`)
}
await writeFile(
  join(ROOT, 'assets', 'css', 'fonts.css'),
  `/* 자기 서버 글꼴 — scripts/self-host-fonts.mjs 가 만든 파일입니다. 직접 고치지 마세요. */\n${rules.join('\n')}\n`,
  'utf8'
)
console.log(`fonts: ${n} files, ${(bytes / 1024).toFixed(0)}KB`)
