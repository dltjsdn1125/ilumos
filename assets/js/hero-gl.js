/* ==========================================================================
   hero-gl.js — hero backdrop: a 3D wireframe star + a radial light that
   follows the pointer.

   Two passes on one WebGL2 canvas:
     1. light pass — fullscreen triangle: a glowing disc with radial rays,
                     locked to the exact pointer position
     2. star pass  — stellated icosahedron drawn as screen-space quads with
                     analytic antialiasing, so the hairlines stay crisp at any
                     devicePixelRatio (gl.LINES cannot do this — browsers clamp
                     lineWidth to one backing pixel)

   ⚠ This replaces, not reproduces, the original backdrop. The source runs
   Unicorn Studio (compiled shaders, scenes id-wbf1qz4dw1aw8s1klbwd /
   id-l5ox3rrhz1vbx0l7hyp1) and two opaque Spline .splinecode scenes, all marked
   UNRECOVERABLE in 07-audit.json. Honoured from the analysis: webgl2, alpha,
   full-viewport fixed canvas behind all content, role="img" + alt text.
   See MISSING.md.
   ========================================================================== */

/* ── tiny mat4 (column-major) ───────────────────────────────────────────── */

const mat4 = () => new Float32Array([1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1])

function perspective(out, fovy, aspect, near, far) {
  const f = 1 / Math.tan(fovy / 2)
  out.fill(0)
  out[0] = f / aspect
  out[5] = f
  out[10] = (far + near) / (near - far)
  out[11] = -1
  out[14] = (2 * far * near) / (near - far)
  return out
}

function multiply(out, a, b) {
  const r = new Float32Array(16)
  for (let i = 0; i < 4; i++) {
    for (let j = 0; j < 4; j++) {
      r[i * 4 + j] =
        a[0 * 4 + j] * b[i * 4 + 0] +
        a[1 * 4 + j] * b[i * 4 + 1] +
        a[2 * 4 + j] * b[i * 4 + 2] +
        a[3 * 4 + j] * b[i * 4 + 3]
    }
  }
  out.set(r)
  return out
}

function rotationY(out, rad) {
  const c = Math.cos(rad)
  const s = Math.sin(rad)
  out.set([c, 0, -s, 0, 0, 1, 0, 0, s, 0, c, 0, 0, 0, 0, 1])
  return out
}

function rotationX(out, rad) {
  const c = Math.cos(rad)
  const s = Math.sin(rad)
  out.set([1, 0, 0, 0, 0, c, s, 0, 0, -s, c, 0, 0, 0, 0, 1])
  return out
}

function translation(out, x, y, z) {
  out.set([1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, x, y, z, 1])
  return out
}

/* ── geometry ───────────────────────────────────────────────────────────────
   Icosahedron body (30 edges) + each of the 20 faces pushed out into a spike
   (60 edges) = 90 segments. Reads unmistakably as a 3D star and stays a
   hairline drawing at any size. */

/* `scale` shrinks the whole solid uniformly. Halved (0.5) once the forest photo
   went in behind: at full size the wireframe spanned most of the frame and read
   as the subject, competing with the trees. Scaling the geometry rather than
   pushing it further back in `view` keeps its depth — and so its parallax against
   the star field — exactly where it was. */
function buildStarSegments(spike = 1.55, scale = 0.5) {
  const t = (1 + Math.sqrt(5)) / 2
  const len = Math.hypot(1, t) / scale
  const verts = [
    [-1, t, 0], [1, t, 0], [-1, -t, 0], [1, -t, 0],
    [0, -1, t], [0, 1, t], [0, -1, -t], [0, 1, -t],
    [t, 0, -1], [t, 0, 1], [-t, 0, -1], [-t, 0, 1],
  ].map(([x, y, z]) => [x / len, y / len, z / len])

  const faces = [
    [0, 11, 5], [0, 5, 1], [0, 1, 7], [0, 7, 10], [0, 10, 11],
    [1, 5, 9], [5, 11, 4], [11, 10, 2], [10, 7, 6], [7, 1, 8],
    [3, 9, 4], [3, 4, 2], [3, 2, 6], [3, 6, 8], [3, 8, 9],
    [4, 9, 5], [2, 4, 11], [6, 2, 10], [8, 6, 7], [9, 8, 1],
  ]

  const seg = []
  const push = (a, b) => seg.push(...a, ...b)

  const seen = new Set()
  for (const [a, b, c] of faces) {
    for (const [i, j] of [[a, b], [b, c], [c, a]]) {
      const key = i < j ? `${i}-${j}` : `${j}-${i}`
      if (seen.has(key)) continue
      seen.add(key)
      push(verts[i], verts[j])
    }
  }

  for (const [a, b, c] of faces) {
    const cx = (verts[a][0] + verts[b][0] + verts[c][0]) / 3
    const cy = (verts[a][1] + verts[b][1] + verts[c][1]) / 3
    const cz = (verts[a][2] + verts[b][2] + verts[c][2]) / 3
    const m = Math.hypot(cx, cy, cz)
    // `spike` is an absolute radius, so it has to carry the same scale as the
    // body above — otherwise shrinking the solid leaves the spikes full length
    // and the star turns into a sea urchin.
    const s = spike * scale
    const tip = [(cx / m) * s, (cy / m) * s, (cz / m) * s]
    push(verts[a], tip)
    push(verts[b], tip)
    push(verts[c], tip)
  }

  return seg
}

/** Turn each segment into two triangles carrying (p0, p1, side, end). */
function expandToQuads(segments) {
  const p0 = []
  const p1 = []
  const side = []
  const end = []
  const corners = [[-1, 0], [1, 0], [-1, 1], [-1, 1], [1, 0], [1, 1]]

  for (let i = 0; i < segments.length; i += 6) {
    const a = segments.slice(i, i + 3)
    const b = segments.slice(i + 3, i + 6)
    for (const [s, e] of corners) {
      p0.push(...a)
      p1.push(...b)
      side.push(s)
      end.push(e)
    }
  }

  return {
    p0: new Float32Array(p0),
    p1: new Float32Array(p1),
    side: new Float32Array(side),
    end: new Float32Array(end),
    count: side.length,
  }
}

/* ── shaders ────────────────────────────────────────────────────────────── */

const LIGHT_VS = `#version 300 es
in vec2 a_pos;
void main() { gl_Position = vec4(a_pos, 0.0, 1.0); }`

/* Radial light: a disc that glows evenly across its whole face, a wide bloom
   around it, and slow radial rays fanning outward. */
const LIGHT_FS = `#version 300 es
precision highp float;
out vec4 outColor;
uniform vec2  u_res;
uniform vec2  u_mouse;   // backing-store pixels, y already flipped
uniform float u_time;


void main() {
  vec2  px     = gl_FragCoord.xy;
  float aspect = u_res.x / max(u_res.y, 1.0);

  vec2 uv = px / u_res;
  vec2 mo = u_mouse / u_res;
  vec2 d  = (uv - mo) * vec2(aspect, 1.0);
  float r = length(d);
  float a = atan(d.y, d.x);

  float pulse = 0.96 + 0.04 * sin(u_time * 1.1);

  /* Halved along with the wireframe star, for the same reason: over the forest
     photo the light was the largest object on screen. Every radial falloff has to
     halve WITH the radius — the exponents are per-unit-distance, so leaving them
     alone would shrink the core and keep the old bloom, which reads as a small
     disc floating in an unexplained haze. */

  // 1. the disc itself — near-flat brightness inside, soft shoulder outward
  float R    = 0.0575;
  float disc = smoothstep(R, R * 0.34, r);

  // 2. wide bloom so the glow reads as a light source, not a sticker
  float bloom = exp(-r * 5.2) * 0.52 + exp(-r * 14.0) * 0.34;

  // 3. radial rays — the "방사형" fan, drifting slowly so it never looks stamped
  float fan   = 0.5 + 0.5 * sin(a * 9.0 + u_time * 0.35);
  float fan2  = 0.5 + 0.5 * sin(a * 15.0 - u_time * 0.22 + 1.7);
  float rays  = (fan * 0.6 + fan2 * 0.4) * exp(-r * 10.4) * 0.42;
  rays       *= smoothstep(R * 0.5, R * 1.6, r);   // start rays outside the core

  // 4. a thin corona ring at the disc edge
  float ring = smoothstep(R * 1.28, R * 1.02, r) * smoothstep(R * 0.86, R * 1.04, r) * 0.5;

  vec3 hot  = vec3(0.92, 0.94, 1.00);
  vec3 cool = vec3(0.42, 0.46, 1.00);

  vec3 col  = hot  * (disc * 0.95 + ring * 0.7);
  col      += cool * (bloom + rays);
  col      *= pulse;

  // faint centre wash keeps the star from floating on pure black
  col += vec3(0.035, 0.035, 0.07) * exp(-length((uv - 0.5) * vec2(aspect, 1.0)) * 2.4);

  // dither against banding across the wide bloom
  float dt = fract(sin(dot(px, vec2(12.9898, 78.233))) * 43758.5453);
  col += (dt - 0.5) / 255.0;

  /* Alpha tracks the light's OWN brightness instead of being a flat 1.0.
     This pass runs with blending disabled, so a constant alpha made the whole
     canvas opaque and nothing behind it could ever show — which is fine over a
     plain black ground, but hides .unicorn__scene (the forest).

     The canvas is premultiplied (the default), so the browser composites
     result = col + backdrop * (1 - a). Setting a to the light's own coverage
     therefore keeps the disc and its bloom looking exactly as before while
     letting the backdrop read through wherever the light contributes nothing.
     The two passes after this one are additive, so stars still add on top.

     Named cover, not a: a is already the polar angle above, and redeclaring it
     fails the whole program to compile - which drops the canvas to the static
     fallback with no console error of its own. */
  float cover = clamp(max(max(col.r, col.g), col.b) * 1.15, 0.0, 1.0);

  outColor = vec4(col, cover);
}`

/* Screen-space quad lines: width set in pixels, edges antialiased analytically. */
const STAR_VS = `#version 300 es
in vec3  a_p0;
in vec3  a_p1;
in float a_side;
in float a_end;

uniform mat4  u_mvp;
uniform vec2  u_res;
uniform float u_halfWidth;   // backing-store pixels

out float v_side;
out float v_depth;
out vec3  v_model;

void main() {
  vec4 c0 = u_mvp * vec4(a_p0, 1.0);
  vec4 c1 = u_mvp * vec4(a_p1, 1.0);
  vec4 c  = mix(c0, c1, a_end);

  vec2 halfRes = u_res * 0.5;
  vec2 s0 = (c0.xy / max(abs(c0.w), 1e-5)) * halfRes;
  vec2 s1 = (c1.xy / max(abs(c1.w), 1e-5)) * halfRes;

  vec2 delta = s1 - s0;
  float lenD = length(delta);
  vec2 dir = lenD > 1e-5 ? delta / lenD : vec2(1.0, 0.0);
  vec2 nrm = vec2(-dir.y, dir.x);

  // offset perpendicular, in pixels, then back into clip space
  c.xy += (nrm * a_side * u_halfWidth / halfRes) * c.w;

  gl_Position = c;
  v_side  = a_side;
  v_depth = c.z / c.w;
  v_model = mix(a_p0, a_p1, a_end);
}`

/* ── warp starfield ─────────────────────────────────────────────────────────
   The scroll-out transition: a field of stars streaming outward from the centre
   of the screen. Each point carries a fixed direction (xy) and a phase (z); the
   phase advances with time and, much faster, with scroll progress out of the
   hero — so scrolling down accelerates the field into a warp and the hero
   dissolves into scattering stars.

   Projection is deliberately hand-rolled rather than run through the mvp: a star
   at direction (x, y) and depth d lands at (x/d, y/d), which is the classic
   "flying through starfield" mapping — stars near the camera fly outward fast
   and grow, distant ones crawl near the centre. */
const FIELD_VS = `#version 300 es
in vec4  a_star;      // xy: direction; z: phase seed [0,1); w: tint seed [0,1)
uniform float u_time;
uniform float u_scroll;   // 0 at the top of the hero, 1 once it has scrolled away
uniform float u_aspect;
uniform float u_dpi;
out float v_bright;
out vec3  v_tint;

/* Stellar colour. Overwhelmingly white, with thin bands of teal, yellow and red
   so the field reads as a real sky rather than a uniform dust of pixels — the
   coloured ones are rare enough to register as individual stars. */
vec3 tintFor(float t) {
  if (t < 0.9250) return vec3(1.00, 1.00, 1.00);   // white          ~92.5%
  if (t < 0.9560) return vec3(0.94, 0.96, 1.00);   // blue-white      ~3.1%
  if (t < 0.9760) return vec3(0.20, 1.00, 0.92);   // teal / cyan     ~2.0%
  if (t < 0.9920) return vec3(1.00, 0.82, 0.16);   // yellow          ~1.6%
  return vec3(1.00, 0.26, 0.20);                   // red             ~0.8%
}

/* A coloured star at the same brightness as its white neighbours disappears into
   the field — colour costs luminance. The boost buys it back so the rare ones
   register as deliberate points of colour rather than dirt. */
float tintBoost(float t) {
  if (t < 0.9250) return 1.00;
  if (t < 0.9560) return 1.12;
  return 1.55;
}

void main() {
  // Depth cycles 1 -> 0. Scroll dominates the term, which is what turns a slow
  // drift into a warp as the user scrolls.
  float advance = u_time * 0.035 + u_scroll * 1.15;
  float d = fract(a_star.z - advance);
  d = max(d, 0.0035);

  vec2 p = a_star.xy / d;
  gl_Position = vec4(p.x / u_aspect, p.y, 0.0, 1.0);

  float near = 1.0 - d;

  // Size is kept tight: at this density fat points merge into a wash. Distant
  // stars stay a single pixel and only the closest few grow.
  gl_PointSize = clamp(pow(near, 3.0) * 6.0 * u_dpi, 1.0 * u_dpi, 13.0 * u_dpi);

  // Fade in at the far plane so stars never pop.
  float born = smoothstep(1.0, 0.90, d);

  // Faint at rest, and it only gains weight once the warp engages. Over the
  // forest photo the resting field should read as atmosphere, not as a second
  // subject — the earlier 0.72 base made it the brightest thing on screen.
  v_bright = pow(near, 1.15) * born * (0.42 + 0.6 * u_scroll) * tintBoost(a_star.w);
  v_tint = tintFor(a_star.w);
}`

const FIELD_FS = `#version 300 es
precision highp float;
in float v_bright;
in vec3  v_tint;
out vec4 outColor;

void main() {
  vec2 c = gl_PointCoord - 0.5;
  float r = length(c);
  if (r > 0.5) discard;

  /* Softer falloff than before. The hard core existed because a gaussian at a
     million points turned the field into fog — at a sixth of that count there is
     room for each star to have a halo again, which is what makes a sparse field
     read as sky rather than as scattered pixels. */
  float a = smoothstep(0.5, 0.06, r);

  // Colour is preserved rather than washed to white, so the rare teal/yellow/red
  // stars actually read as coloured.
  outColor = vec4(v_tint * v_bright * 1.1, a * v_bright);
}`

const STAR_FS = `#version 300 es
precision highp float;
in float v_side;
in float v_depth;
in vec3  v_model;
out vec4 outColor;

uniform vec2  u_res;
uniform vec2  u_mouse;
uniform float u_halfWidth;

void main() {
  // analytic AA across the quad: one backing pixel of feather, no more
  float feather = clamp(1.0 / max(u_halfWidth, 0.5), 0.05, 0.95);
  float edge = 1.0 - smoothstep(1.0 - feather, 1.0, abs(v_side));
  if (edge <= 0.001) discard;

  // far edges recede — what makes a flat wireframe read as 3D
  float depth = clamp(v_depth * 0.5 + 0.5, 0.0, 1.0);
  float fade  = mix(1.0, 0.42, smoothstep(0.35, 0.88, depth));

  // the cursor light spills onto the lines it passes over
  float aspect = u_res.x / max(u_res.y, 1.0);
  vec2 uv = gl_FragCoord.xy / u_res;
  vec2 mo = u_mouse / u_res;
  float r = length((uv - mo) * vec2(aspect, 1.0));
  float lit = exp(-r * 3.6);

  float tip = smoothstep(0.85, 1.5, length(v_model));   // spike tips run hotter

  vec3 col = mix(vec3(0.80, 0.84, 1.0), vec3(1.0), tip);
  col *= 0.92 + 1.15 * lit;

  float alpha = edge * fade * (0.72 + 0.5 * lit + 0.22 * tip);
  outColor = vec4(col, clamp(alpha, 0.0, 1.0));
}`

/* ── gl helpers ─────────────────────────────────────────────────────────── */

function compile(gl, type, src, label) {
  const sh = gl.createShader(type)
  gl.shaderSource(sh, src)
  gl.compileShader(sh)
  if (!gl.getShaderParameter(sh, gl.COMPILE_STATUS)) {
    console.warn(`[hero-gl] ${label} compile failed:`, gl.getShaderInfoLog(sh))
    gl.deleteShader(sh)
    return null
  }
  return sh
}

function program(gl, vsSrc, fsSrc, label) {
  const vs = compile(gl, gl.VERTEX_SHADER, vsSrc, `${label} vs`)
  const fs = compile(gl, gl.FRAGMENT_SHADER, fsSrc, `${label} fs`)
  if (!vs || !fs) return null
  const p = gl.createProgram()
  gl.attachShader(p, vs)
  gl.attachShader(p, fs)
  gl.linkProgram(p)
  if (!gl.getProgramParameter(p, gl.LINK_STATUS)) {
    console.warn(`[hero-gl] ${label} link failed:`, gl.getProgramInfoLog(p))
    return null
  }
  return p
}

function attribBuffer(gl, prog, name, data, size) {
  const loc = gl.getAttribLocation(prog, name)
  if (loc < 0) return
  const buf = gl.createBuffer()
  gl.bindBuffer(gl.ARRAY_BUFFER, buf)
  gl.bufferData(gl.ARRAY_BUFFER, data, gl.STATIC_DRAW)
  gl.enableVertexAttribArray(loc)
  gl.vertexAttribPointer(loc, size, gl.FLOAT, false, 0, 0)
}

/* ── main ───────────────────────────────────────────────────────────────── */

export function initHeroGL(host) {
  if (!host) return null

  const reduced = window.matchMedia('(prefers-reduced-motion: reduce)').matches

  const canvas = document.createElement('canvas')
  canvas.setAttribute('role', 'img')
  canvas.setAttribute('aria-label', 'gradient background')
  host.prepend(canvas)

  const gl = canvas.getContext('webgl2', {
    alpha: true,
    depth: false,
    antialias: true,
    powerPreference: 'default',
  })

  if (!gl) {
    canvas.remove()
    return null
  }

  const lightProg = program(gl, LIGHT_VS, LIGHT_FS, 'light')
  const starProg = program(gl, STAR_VS, STAR_FS, 'star')
  const fieldProg = program(gl, FIELD_VS, FIELD_FS, 'field')
  if (!lightProg || !starProg || !fieldProg) {
    canvas.remove()
    return null
  }

  /* GPU 가 없는 환경(SwiftShader·llvmpipe 같은 소프트웨어 렌더러)에서는 이 장면을
     CPU 로 그리게 돼 메인 스레드가 수 초씩 막힙니다. 그런 곳은 정적 배경(.unicorn__fallback
     + 숲 사진)만 보여 주고 WebGL 은 켜지 않습니다. GPU 가 있는 일반 기기는 그대로입니다. */
  try {
    const dbg = gl.getExtension('WEBGL_debug_renderer_info')
    const gpu = dbg ? String(gl.getParameter(dbg.UNMASKED_RENDERER_WEBGL) || '') : ''
    if (/swiftshader|llvmpipe|software|basic render/i.test(gpu)) {
      canvas.remove()
      return null
    }
  } catch (_) { /* 확인 못 하면 그대로 진행 */ }

  // fullscreen triangle
  const quadVao = gl.createVertexArray()
  gl.bindVertexArray(quadVao)
  attribBuffer(gl, lightProg, 'a_pos', new Float32Array([-1, -1, 3, -1, -1, 3]), 2)

  // star quads
  const star = expandToQuads(buildStarSegments())
  const starVao = gl.createVertexArray()
  gl.bindVertexArray(starVao)
  attribBuffer(gl, starProg, 'a_p0', star.p0, 3)
  attribBuffer(gl, starProg, 'a_p1', star.p1, 3)
  attribBuffer(gl, starProg, 'a_side', star.side, 1)
  attribBuffer(gl, starProg, 'a_end', star.end, 1)

  /* Warp starfield.
     Each star: direction on a disc (sqrt keeps the disc evenly covered), a phase
     seed for its depth cycle, and a tint seed. Deterministic pseudo-random — no
     Math.random — so the field is identical on every load.

     Count is tiered by viewport. One gl.POINTS draw call handles these numbers;
     the real cost is fragments, which is why point size is clamped tight in the
     shader.

     Measured on the dev machine: 1.4M sat exactly on the frame budget — median
     16.7ms but dropping up to 10 frames in 120. 1.0M holds a clean 60fps with
     headroom, which is where this is set: a machine weaker than the dev box
     should still hold. This is THE tuning knob — raise it if the target hardware
     is known to be strong, lower it if the field stutters. The visual degrades
     gradually either way. See MISSING.md > C14 for why this is not literally the
     10,000x that was asked for. */
  /* Cut roughly 6x from the original counts (1.0M / 300k / 80k). Those were tuned
     for a star field on pure black, where sheer density WAS the image. It now sits
     over a photographed forest, and at that density the sky read as static noise
     competing with the trees rather than as a sky above them. Fewer, fainter
     points let the forest carry the frame. Still the one tuning knob. */
  const FIELD_COUNT = (() => {
    const w = window.innerWidth
    if (w < 768) return 14000
    if (w < 1280) return 48000
    return 160000
  })()

  const fieldData = new Float32Array(FIELD_COUNT * 4)
  let seed = 20240917
  const rnd = () => {
    seed = (seed * 1664525 + 1013904223) % 4294967296
    return seed / 4294967296
  }
  for (let i = 0; i < FIELD_COUNT; i++) {
    const a = rnd() * Math.PI * 2
    const r = Math.sqrt(rnd()) * 1.3
    fieldData[i * 4] = Math.cos(a) * r
    fieldData[i * 4 + 1] = Math.sin(a) * r
    fieldData[i * 4 + 2] = rnd()
    fieldData[i * 4 + 3] = rnd()
  }

  const fieldVao = gl.createVertexArray()
  gl.bindVertexArray(fieldVao)
  attribBuffer(gl, fieldProg, 'a_star', fieldData, 4)

  gl.bindVertexArray(null)

  const U = {
    lightRes: gl.getUniformLocation(lightProg, 'u_res'),
    lightMouse: gl.getUniformLocation(lightProg, 'u_mouse'),
    lightTime: gl.getUniformLocation(lightProg, 'u_time'),
    starMvp: gl.getUniformLocation(starProg, 'u_mvp'),
    starRes: gl.getUniformLocation(starProg, 'u_res'),
    starMouse: gl.getUniformLocation(starProg, 'u_mouse'),
    starHalfWidth: gl.getUniformLocation(starProg, 'u_halfWidth'),
    fieldTime: gl.getUniformLocation(fieldProg, 'u_time'),
    fieldScroll: gl.getUniformLocation(fieldProg, 'u_scroll'),
    fieldAspect: gl.getUniformLocation(fieldProg, 'u_aspect'),
    fieldDpi: gl.getUniformLocation(fieldProg, 'u_dpi'),
  }

  /* Scroll progress out of the hero. Measured from the hero block itself rather
     than a fixed pixel count, so it works on the home page and every sub-page
     hero regardless of their different heights. */
  const heroBlock = document.querySelector('.hero, .page-hero')

  /* Once the hero is behind us the field must NOT stay locked at full warp —
     every section below is translucent over it, so a permanent warp would
     streak behind the whole page. It settles to a slow ambient drift instead,
     which is the calm star backdrop the rest of the document sits on. */
  const AMBIENT = 0.05

  function scrollProgress() {
    if (window.__heroPast) return AMBIENT
    // motion.js pins the home hero and publishes the pin's own progress here.
    // While pinned the hero's rect never moves, so the fallback below cannot see
    // the scroll — the published value is the authority when it exists.
    if (typeof window.__heroWarp === 'number') return window.__heroWarp
    if (!heroBlock) return AMBIENT
    const r = heroBlock.getBoundingClientRect()
    const h = r.height || 1
    // 0 while the hero is in place, 1 once it has fully left the top.
    return Math.max(AMBIENT, Math.min(1, -r.top / h))
  }

  /* Pointer state.
     `light` is the exact cursor position — the glow must sit under the cursor,
     so it is not smoothed. `tilt` is a lagged copy, used only to sway the star.
     Normalisation uses the canvas rect, NOT window.innerWidth: the canvas
     excludes the classic scrollbar, so innerWidth would offset the light by the
     scrollbar width and it would never quite line up with the pointer. */
  const light = { x: 0.5, y: 0.5 }
  const tilt = { x: 0.5, y: 0.5 }
  let hasPointer = false

  // Render one backing pixel per device pixel — the condition for crisp hairlines.
  const dpi = () => Math.min(Math.max(window.devicePixelRatio || 1, 0.75), 2)

  function resize() {
    const ratio = dpi()
    const w = Math.max(1, Math.round(host.clientWidth * ratio))
    const h = Math.max(1, Math.round(host.clientHeight * ratio))
    if (canvas.width === w && canvas.height === h) return
    canvas.width = w
    canvas.height = h
    gl.viewport(0, 0, w, h)
  }

  const proj = mat4()
  const view = mat4()
  const rotY = mat4()
  const rotX = mat4()
  const model = mat4()
  const mvp = mat4()

  let raf = 0
  let visible = true
  let lost = false
  let intro = 0    // 0..1 — one-time ease from the star's centre out to the pointer
  let scroll = 0   // 0..1 — smoothed hero scroll-out, drives the warp transition
  let last = 0
  const INTRO_SECONDS = 0.55
  const start = performance.now()

  /* 느린 기기 보호: 프레임 간격이 계속 길면(25fps 미만) 별 개수를 절반씩 줄이고,
     그래도 느리면(15fps 미만) 정적 배경으로 물러납니다. 빠른 기기는 1.0 그대로. */
  let fieldShare = 1
  let tickAt = 0
  let tickSum = 0
  let tickN = 0

  function adapt(now) {
    if (tickAt) {
      const dt = now - tickAt
      if (dt < 250) { tickSum += dt; tickN += 1 }
    }
    tickAt = now
    if (tickN < 40) return
    const avg = tickSum / tickN
    tickSum = 0
    tickN = 0
    if (avg > 66) {
      lost = true
      host.removeAttribute('data-gl')
      canvas.remove()
    } else if (avg > 40 && fieldShare > 0.25) {
      fieldShare *= 0.5
    }
  }

  function frame(now) {
    raf = requestAnimationFrame(frame)
    if (!visible || lost) { tickAt = 0; return }
    adapt(now)

    resize()
    // Under reduced motion the clock freezes: the star stops spinning and the
    // light stops breathing. The cursor light still tracks, because that is a
    // direct response to input rather than autonomous animation.
    const t = reduced ? 0 : (now - start) / 1000

    tilt.x += (light.x - tilt.x) * 0.06
    tilt.y += (light.y - tilt.y) * 0.06

    /* `light` is already stored in GL's bottom-origin space (see onPointer), so
       it scales straight to backing pixels — flipping again here would cancel
       the first flip and put the glow at the vertical mirror of the cursor.

       On load the light sits at the centre of the canvas, which is exactly where
       the star projects, so the hero opens lit from inside the star. The first
       pointer move eases it out to the cursor over INTRO_SECONDS; after that it
       tracks the pointer exactly, with no smoothing. */
    if (hasPointer && intro < 1) {
      const dt = last ? Math.min((now - last) / 1000, 0.05) : 0
      intro = Math.min(1, intro + dt / INTRO_SECONDS)
    }
    last = now

    const ease = intro * intro * (3 - 2 * intro) // smoothstep
    const lx = 0.5 + (light.x - 0.5) * ease
    const ly = 0.5 + (light.y - 0.5) * ease

    const mx = lx * canvas.width
    const my = ly * canvas.height

    // How far the hero has scrolled away — drives the whole transition.
    const sp = reduced ? 0 : scrollProgress()
    scroll += (sp - scroll) * 0.18 // light smoothing so wheel steps don't judder

    gl.disable(gl.BLEND)
    gl.clearColor(0, 0, 0, 0)
    gl.clear(gl.COLOR_BUFFER_BIT)

    // pass 1 — the radial light. Dims as the hero leaves so the field takes over.
    gl.useProgram(lightProg)
    gl.bindVertexArray(quadVao)
    gl.uniform2f(U.lightRes, canvas.width, canvas.height)
    gl.uniform2f(U.lightMouse, mx, my)
    gl.uniform1f(U.lightTime, t)
    gl.drawArrays(gl.TRIANGLES, 0, 3)

    // pass 1.5 — warp starfield, additive over the light
    const aspectField = canvas.width / Math.max(canvas.height, 1)
    gl.useProgram(fieldProg)
    gl.bindVertexArray(fieldVao)
    gl.enable(gl.BLEND)
    gl.blendFunc(gl.SRC_ALPHA, gl.ONE)
    gl.uniform1f(U.fieldTime, t)
    gl.uniform1f(U.fieldScroll, scroll)
    gl.uniform1f(U.fieldAspect, aspectField)
    gl.uniform1f(U.fieldDpi, dpi())
    gl.drawArrays(gl.POINTS, 0, Math.max(1, Math.floor(FIELD_COUNT * fieldShare)))
    gl.disable(gl.BLEND)

    // pass 2 — wireframe star, additive so the lines glow rather than occlude.
    // It recedes and spins up as the field warps, so the star reads as breaking
    // apart into the scattering stars rather than simply fading out.
    const aspect = canvas.width / Math.max(canvas.height, 1)
    perspective(proj, (46 * Math.PI) / 180, aspect, 0.1, 100)
    translation(view, 0, 0, -4.15 - scroll * 9.0)

    const spin = reduced ? 0.6 : t * 0.16 + scroll * 2.2
    rotationY(rotY, spin + (tilt.x - 0.5) * 0.5)
    rotationX(rotX, -0.34 + (tilt.y - 0.5) * 0.45)

    multiply(model, rotX, rotY)
    multiply(mvp, view, model)
    multiply(mvp, proj, mvp)

    gl.useProgram(starProg)
    gl.bindVertexArray(starVao)
    gl.enable(gl.BLEND)
    gl.blendFunc(gl.SRC_ALPHA, gl.ONE)
    gl.uniformMatrix4fv(U.starMvp, false, mvp)
    gl.uniform2f(U.starRes, canvas.width, canvas.height)
    gl.uniform2f(U.starMouse, mx, my)
    // 1.4 CSS px, converted to backing pixels -> same apparent weight everywhere
    gl.uniform1f(U.starHalfWidth, 0.7 * dpi())
    gl.drawArrays(gl.TRIANGLES, 0, star.count)

    gl.bindVertexArray(null)
  }

  function onPointer(e) {
    // Normalise against the CANVAS rect, not window.innerWidth: the canvas
    // excludes the classic scrollbar, so innerWidth would drift the light by the
    // scrollbar width and it would never sit exactly under the pointer.
    // y is flipped once, here, into GL's bottom-origin space.
    const r = canvas.getBoundingClientRect()
    if (r.width < 1 || r.height < 1) return
    light.x = (e.clientX - r.left) / r.width
    light.y = 1 - (e.clientY - r.top) / r.height
    if (!hasPointer) {
      hasPointer = true
      tilt.x = light.x
      tilt.y = light.y
    }
  }

  function onVisibility() {
    visible = document.visibilityState === 'visible'
  }

  function onLost(e) {
    e.preventDefault()
    lost = true
    host.removeAttribute('data-gl')
  }

  window.addEventListener('pointermove', onPointer, { passive: true })
  window.addEventListener('resize', resize)
  document.addEventListener('visibilitychange', onVisibility)
  canvas.addEventListener('webglcontextlost', onLost)

  resize()
  host.setAttribute('data-gl', 'on')
  raf = requestAnimationFrame(frame)

  return {
    destroy() {
      cancelAnimationFrame(raf)
      window.removeEventListener('pointermove', onPointer)
      window.removeEventListener('resize', resize)
      document.removeEventListener('visibilitychange', onVisibility)
      canvas.removeEventListener('webglcontextlost', onLost)
      canvas.remove()
      host.removeAttribute('data-gl')
    },
  }
}
