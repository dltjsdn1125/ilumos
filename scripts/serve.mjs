/**
 * serve.mjs — zero-dependency static server for local preview.
 * Resolves /path/ -> /path/index.html and falls back to /404.html.
 */

import { createServer } from 'node:http'
import { readFile, stat } from 'node:fs/promises'
import { join, extname, normalize, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..')
const PORT = Number(process.env.PORT || 4173)

const TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.webp': 'image/webp',
  '.svg': 'image/svg+xml',
  '.xml': 'application/xml; charset=utf-8',
  '.txt': 'text/plain; charset=utf-8',
  '.webm': 'video/webm',
  '.mp4': 'video/mp4',
  '.woff2': 'font/woff2',
}

async function resolve(urlPath) {
  const clean = decodeURIComponent(urlPath.split('?')[0])
  let p = join(ROOT, normalize(clean).replace(/^(\.\.[/\\])+/, ''))
  try {
    const s = await stat(p)
    if (s.isDirectory()) p = join(p, 'index.html')
    return p
  } catch {
    return p
  }
}

createServer(async (req, res) => {
  const file = await resolve(req.url === '/' ? '/index.html' : req.url)
  try {
    const body = await readFile(file)
    res.writeHead(200, { 'Content-Type': TYPES[extname(file)] || 'application/octet-stream' })
    res.end(body)
  } catch {
    try {
      const body = await readFile(join(ROOT, '404.html'))
      res.writeHead(404, { 'Content-Type': TYPES['.html'] })
      res.end(body)
    } catch {
      res.writeHead(404).end('Not found')
    }
  }
}).listen(PORT, () => console.log(`http://localhost:${PORT}`))
