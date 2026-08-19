/**
 * vendor.mjs — copy runtime libraries out of node_modules into assets/js/vendor
 * so the built site runs offline with zero external script hosts.
 *
 * Library selection comes straight from 03-motion.json > stack:
 *   Lenis 1.2.3 · GSAP 3.15.0 + ScrollTrigger, SplitText, ScrambleTextPlugin,
 *   Flip, Observer, Draggable, InertiaPlugin, CustomEase
 * (installed range is ^3.13 — the first fully-free GSAP; all plugins above ship in it)
 */
import { mkdir, copyFile, readFile, writeFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..')
const OUT = join(ROOT, 'assets', 'js', 'vendor')

const GSAP_FILES = [
  'gsap.min.js',
  'ScrollTrigger.min.js',
  'SplitText.min.js',
  'ScrambleTextPlugin.min.js',
  'Flip.min.js',
  'Observer.min.js',
  'Draggable.min.js',
  'InertiaPlugin.min.js',
  'CustomEase.min.js',
]

await mkdir(OUT, { recursive: true })

for (const f of GSAP_FILES) {
  await copyFile(join(ROOT, 'node_modules', 'gsap', 'dist', f), join(OUT, f))
}

// Lenis ships UMD as lenis.min.js (global `Lenis`).
await copyFile(join(ROOT, 'node_modules', 'lenis', 'dist', 'lenis.min.js'), join(OUT, 'lenis.min.js'))
await copyFile(join(ROOT, 'node_modules', 'lenis', 'dist', 'lenis.css'), join(ROOT, 'assets', 'css', 'lenis.css'))

const gsapPkg = JSON.parse(await readFile(join(ROOT, 'node_modules', 'gsap', 'package.json'), 'utf8'))
const lenisPkg = JSON.parse(await readFile(join(ROOT, 'node_modules', 'lenis', 'package.json'), 'utf8'))

await writeFile(
  join(OUT, 'VERSIONS.txt'),
  `gsap ${gsapPkg.version}\nlenis ${lenisPkg.version}\n\n` +
    `Analysis (03-motion.json) recorded gsap 3.15.0 / lenis 1.2.3 on the live site.\n` +
    `Vendored here from npm; plugin set matches the analysed list exactly.\n`,
  'utf8'
)

console.log(`vendored ${GSAP_FILES.length + 1} libs -> assets/js/vendor (gsap ${gsapPkg.version}, lenis ${lenisPkg.version})`)
