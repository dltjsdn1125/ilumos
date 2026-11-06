/**
 * make-small-images.mjs — 속도용 사본 만들기 (원본은 건드리지 않음).
 *   - where-*.webp  → where-*-400.webp  (로고 월 카드는 172px 로 표시 → 2배 선명도용 400px)
 *   - og-cover.webp → og-share.webp     (공유 대표 사진 1200×630)
 * 사용: node scripts/make-small-images.mjs
 */
import sharp from 'sharp'
import { readdirSync } from 'node:fs'
import { join, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'

const DIR = join(dirname(fileURLToPath(import.meta.url)), '..', 'assets', 'images')

for (const f of readdirSync(DIR)) {
  if (!/^where-[a-z]+\.webp$/.test(f)) continue
  const out = f.replace('.webp', '-400.webp')
  const info = await sharp(join(DIR, f)).resize(400, 300, { fit: 'cover' }).webp({ quality: 82, effort: 6 }).toFile(join(DIR, out))
  console.log(out, info.size)
}

const og = await sharp(join(DIR, 'og-cover.webp')).resize(1200, 630, { fit: 'cover', position: 'attention' }).webp({ quality: 82, effort: 6 }).toFile(join(DIR, 'og-share.webp'))
console.log('og-share.webp', og.size)
