/**
 * export-products.mjs — 카탈로그 20종을 외부 쇼핑몰(카페24 등)에 올릴 수 있는
 * 형태로 내보냅니다.
 *
 *   node scripts/export-products.mjs
 *
 * 만드는 것 (export/ 폴더, git 에는 올리지 않습니다):
 *   export/products.csv        상품 일괄등록용 표
 *   export/products.json       API 로 밀어 넣을 때 쓸 원본 (카페24 MCP/Admin API 등)
 *   export/images/             상품 이미지 20장 (자체 상품코드로 이름 붙임)
 *
 * ⚠ CSV 의 열 이름은 **카페24 관리자에서 내려받은 일괄등록 양식과 맞춰야 합니다.**
 *   카페24는 양식 버전에 따라 열 구성과 순서가 달라지고, 인코딩도 EUC-KR 을
 *   요구하는 경우가 있습니다. 여기서는 UTF-8 BOM 으로 씁니다 — 엑셀에서 한글이
 *   깨지지 않는 기본값입니다. 업로드가 거부되면 양식을 먼저 내려받아
 *   `COLUMNS` 를 그 헤더로 바꾸십시오. 값 계산은 그대로 쓸 수 있습니다.
 *
 * ⚠ 판매가는 **시안용 예시값**입니다(MISSING.md > D7). 실제 등록 전에
 *   src/data/site.mjs 의 `products[].price` 를 확정가로 바꾸고 다시 뽑으십시오.
 */

import { mkdir, writeFile, copyFile, rm } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { site, products, shopCategories, services } from '../src/data/site.mjs'

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..')
const OUT = join(ROOT, 'export')

/* CSV 한 칸. 쉼표·따옴표·줄바꿈이 들어가면 통째로 감싸고 따옴표는 두 번 씁니다
   (RFC 4180). 상세설명에 HTML 과 줄바꿈이 들어가므로 반드시 필요합니다. */
const cell = (v) => {
  const s = String(v ?? '')
  return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s
}

/* 카페24 일괄등록에서 흔히 쓰는 열들입니다. 양식과 대조해 이름만 바꾸면 됩니다. */
const COLUMNS = [
  '상품명',
  '영문 상품명',
  '자체 상품코드',
  '상품 분류',
  '판매가',
  '과세구분',
  '진열상태',
  '판매상태',
  '제조사',
  '공급사',
  '원산지',
  '상품 요약설명',
  '상품 상세설명',
  '검색어',
  '대표이미지 파일명',
]

/* 상세설명은 쇼핑몰 에디터에 그대로 들어가는 HTML 입니다. 이 사이트의 상세
   페이지에서 **사실만** 옮깁니다 — 규격표, 요약, 품목 설명 첫 문단. 사이트에
   없는 문구를 여기서 새로 지어내지 않습니다. */
function detailHtml(p, cat, service) {
  const rows = [...p.specs, { k: '품목 코드', v: p.sku }]
    .map((s) => `    <tr><th style="text-align:left;padding:6px 16px 6px 0;white-space:nowrap">${s.k}</th><td style="padding:6px 0">${s.v}</td></tr>`)
    .join('\n')

  return [
    `<p>${p.summary}</p>`,
    '<table style="border-collapse:collapse;margin:16px 0">',
    '  <tbody>',
    rows,
    '  </tbody>',
    '</table>',
    service ? `<p>${service.body[0]}</p>` : '',
    `<p>${site.name} · TEL ${site.phone} · ${site.email}</p>`,
  ]
    .filter(Boolean)
    .join('\n')
}

await rm(OUT, { recursive: true, force: true })
await mkdir(join(OUT, 'images'), { recursive: true })

let manifest = {}
const manifestPath = join(ROOT, 'assets', 'images', 'manifest.json')
if (existsSync(manifestPath)) manifest = JSON.parse(await (await import('node:fs/promises')).readFile(manifestPath, 'utf8'))

const rows = []
const json = []
let copied = 0

for (const p of products) {
  const cat = shopCategories.find((c) => c.slug === p.category)
  const service = services.find((s) => s.slug === p.category)
  const imageFile = `${p.sku}.webp`

  rows.push([
    p.name,
    p.nameLatin,
    p.sku,
    cat ? cat.label : '',
    p.price,
    'A', // 과세
    'Y', // 진열
    'Y', // 판매
    site.name,
    site.name,
    '대한민국',
    `${p.summary} (${p.pack})`,
    detailHtml(p, cat, service),
    p.tags.join(','),
    imageFile,
  ])

  json.push({
    sku: p.sku,
    name: p.name,
    nameEn: p.nameLatin,
    category: cat ? cat.label : '',
    price: p.price,
    unit: p.unit,
    pack: p.pack,
    summary: p.summary,
    specs: p.specs,
    tags: p.tags,
    detailHtml: detailHtml(p, cat, service),
    image: imageFile,
    sourceUrl: `${site.origin}/shop/${p.slug}/`,
  })

  const src = join(ROOT, 'assets', 'images', `${p.image}.webp`)
  if (existsSync(src)) {
    await copyFile(src, join(OUT, 'images', imageFile))
    copied += 1
  }
}

/* BOM. 없으면 엑셀이 UTF-8 을 못 알아보고 한글이 전부 깨집니다. */
const csv = '﻿' + [COLUMNS, ...rows].map((r) => r.map(cell).join(',')).join('\r\n') + '\r\n'

await writeFile(join(OUT, 'products.csv'), csv, 'utf8')
await writeFile(join(OUT, 'products.json'), JSON.stringify(json, null, 2), 'utf8')
await writeFile(
  join(OUT, 'README.txt'),
  [
    `${site.name} 상품 등록 자료 — ${new Date().toISOString().slice(0, 10)} 생성`,
    '',
    `products.csv    상품 ${products.length}종 (일괄등록용)`,
    `products.json   같은 내용, API 로 밀어 넣을 때 사용`,
    `images/         대표이미지 ${copied}장, 파일명 = 자체 상품코드`,
    '',
    '주의',
    '  · 판매가는 홈페이지 시안용 예시값입니다. 등록 전 확정가로 교체하십시오.',
    '  · CSV 열 이름은 카페24 관리자에서 내려받은 일괄등록 양식과 대조가 필요합니다.',
    '  · 인코딩은 UTF-8(BOM)입니다. 양식이 EUC-KR 을 요구하면 변환하십시오.',
    '  · 등록 후 전자상거래법상 고지(교환·반품·청약철회)는 쇼핑몰 쪽에서 별도 설정해야 합니다.',
  ].join('\n'),
  'utf8'
)

console.log(`export/products.csv  — ${products.length}종 (${COLUMNS.length}개 열)`)
console.log(`export/products.json — API 등록용`)
console.log(`export/images/       — 대표이미지 ${copied}장`)
console.log(`\n⚠ 판매가는 시안용 예시입니다. 등록 전 src/data/site.mjs 의 price 를 확정가로 바꾸십시오.`)
