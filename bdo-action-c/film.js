// BDO Automotive – Action C «DATA». 60 s, 128 bpm, 128 beats. Every shot is a flat data visualisation
// the big word sits in or is made of. Told from the customer's side: «du har tallene; vi viser hva de betyr».
// Pure function of time: no state between frames, no Math.random (seeded hashes only).
import * as M from './lib/motion.js';
import * as K from './lib/kit.js';
import * as world from './scenes/world.js';

const { W, H, C, clamp, lerp, prog, E, rgba } = K;
const P = 0.46875, END = 116, N = 2400, DISPLAY = K.DISPLAY, MONO = K.MONO;
const hh = (i, s) => K.hash01(i, s);

// ------------------------------------------------------------------ looks
const LK = {
  N: { id: 'N', k: 0, bg: C.night, ink: C.white, data: C.cloud, dim: C.slate, acc: C.red },
  A: { id: 'A', k: 1, bg: C.arctic, ink: C.blue, data: C.blue, dim: C.slate, acc: C.red },
  B: { id: 'B', k: 0, bg: C.blue, ink: C.white, data: C.white, dim: C.pale, acc: C.red },
  R: { id: 'R', k: 0, bg: C.red, ink: C.white, data: C.white, dim: C.white, acc: C.white },
};
function paintBg(ctx, L, u) {
  ctx.fillStyle = L.bg; ctx.fillRect(-120, -120, W + 240, H + 240);
  if (L.id === 'N') K.grid(ctx, 0, 1, u * 9, u * 5);
  else if (L.id === 'A') K.grid(ctx, 1, 1, -u * 9, u * 5);
  else K.grid(ctx, 0, L.id === 'R' ? 1.4 : 1.2, u * 9, -u * 5);
}
const disc = (ctx, x, y, r, col) => { ctx.fillStyle = col; ctx.beginPath(); ctx.arc(x, y, Math.max(0.1, r), 0, 7); ctx.fill(); };
const ringS = (ctx, x, y, r, w, col) => { ctx.lineWidth = w; ctx.strokeStyle = col; ctx.beginPath(); ctx.arc(x, y, Math.max(0.1, r), 0, 7); ctx.stroke(); };
function mono(ctx, str, x, y, { size = 34, col = '#fff', align = 'center', q = 1, track = 0.12 } = {}) {
  if (q <= 0) return;
  ctx.save(); ctx.font = M.font(size, 500, MONO); ctx.letterSpacing = `${size * track}px`; ctx.textAlign = align; ctx.fillStyle = col;
  const w = ctx.measureText(str).width; const x0 = align === 'center' ? x - w / 2 : align === 'right' ? x - w : x;
  ctx.beginPath(); ctx.rect(x0 - 10, y - size * 1.05, w + 20, size * 1.5); ctx.clip();
  ctx.fillText(str, x + (align === 'center' ? track * size / 2 : 0), y + (1 - q) * size * 1.3); ctx.restore();
}

// ------------------------------------------------------------------ formations (2400 dots; index 0 = the signature dot)
const BOX = { low: [96, 1824, 520, 1000], up: [96, 1824, 96, 600], full: [96, 1824, 96, 984] };
const mapOf = (b) => (u, v) => [b[0] + u * (b[1] - b[0]), b[2] + v * (b[3] - b[2])];
function colsGen(box, wts, cw, step, r) {
  const o = [], S = wts.reduce((a, b) => a + b, 0), slot = (box[1] - box[0]) / wts.length, per = Math.max(1, Math.floor(slot * cw / step));
  wts.forEach((w, j) => {
    const nj = Math.round(2250 * w / S), x0 = box[0] + slot * (j + 0.5) - per * step / 2 + step / 2;
    for (let k = 0; k < nj; k++) o.push([x0 + (k % per) * step, box[3] - Math.floor(k / per) * step - step / 2, r, j]);
  });
  return o;
}
function chartLine(m, seed, steep, thick) {
  const rr = K.mulberry32(seed), pts = [];
  for (let i = 0; i < 28; i++) { const t = i / 27; pts.push(m(t, 1 - clamp(0.08 + Math.pow(t, steep) * 0.84 + (rr() - 0.5) * 0.1 * (1 - 0.5 * t)))); }
  const o = [];
  for (let i = 0; i < 27; i++) { const [a, b] = [pts[i], pts[i + 1]], len = Math.hypot(b[0] - a[0], b[1] - a[1]), n = Math.ceil(len / 5);
    for (let k = 0; k < n; k++) for (let q = 0; q < thick; q++) { const f = k / n, jit = (q - (thick - 1) / 2) * 5.5; o.push([lerp(a[0], b[0], f), lerp(a[1], b[1], f) + jit, 3.1, 0]); } }
  return { o, sig: [...pts[27], 11] };
}
const GEN = {
  grid(m) { const o = []; for (let j = 0; j < 40; j++) for (let i = 0; i < 60; i++) o.push([...m((i + 0.5) / 60, (j + 0.5) / 40), 5, 0]); return { o, sig: [...m(0.62, 0.55), 10] }; },
  burst(m) { const o = []; for (let i = 0; i < 2300; i++) { const a = hh(i, 1) * 6.283, d = Math.pow(hh(i, 2), 0.65); o.push([...m(0.5 + 0.5 * d * Math.cos(a), 0.5 + 0.5 * d * Math.sin(a)), 1.8 + hh(i, 3) * 3.2, 0]); } return { o, sig: [...m(0.5, 0.5), 11] }; },
  field(m) { const o = []; for (let i = 1; i < K.NPTS; i++) { const p = K.PTS[i]; o.push([...m(p.x / 1920, p.y / 1080), p.r * 1.1, 0]); } return { o, sig: [...m(K.PTS[0].x / 1920, K.PTS[0].y / 1080), 11] }; },
  band(m) { const o = []; for (let i = 0; i < 2300; i++) { const u = hh(i, 4), v = 0.5 + ((hh(i, 5) + hh(i, 6) + hh(i, 7)) / 3 - 0.5) * 0.85 + 0.14 * Math.sin(u * 9); o.push([...m(u, clamp(v, 0.02, 0.98)), 1.8 + hh(i, 8) * 3, 0]); } return { o, sig: [...m(0.72, 0.55), 12] }; },
  waves(m) { const o = []; for (let i = 0; i < 2300; i++) { const w = i % 4, u = Math.floor(i / 4) / 575; o.push([...m(u, (w + 0.5) / 4 + 0.1 * Math.sin(u * 12 + w * 1.3)), 3, 0]); } return { o, sig: [...m(0.8, 0.5 + 0.1 * Math.sin(9.6 + 2.6)), 11] }; },
  ledger(m) { const o = []; for (let r = 0; r < 22; r++) { let u = 0; while (u < 1) { const seg = 0.04 + hh(r * 40 + Math.floor(u * 90), 3) * 0.22, on = hh(r * 31 + Math.floor(u * 70), 9) > 0.28; if (on) for (let x = u; x < Math.min(1, u + seg); x += 1 / 115) o.push([...m(x, (r + 0.5) / 22), 3.1, 0]); u += seg + 0.025; } } return { o, sig: [...m(0.3, 13.5 / 22), 10] }; },
  cols6(m, b) { return { o: colsGen(b, [0.13, 0.17, 0.2, 0.15, 0.14, 0.19], 0.7, 14, 5.4), sig: [...m(0.5, 0.08), 0] }; },
  sizes(m, b) { return { o: colsGen(b, [0.12, 0.3, 0.58], 0.75, 10, 4.2), sig: [...m(0.5, 0.02), 0] }; },
  stairs(m, b) { return { o: colsGen(b, [1, 2, 3, 4, 5, 6, 7, 8], 0.8, 12, 5), sig: [...m(0.94, 0.04), 12] }; },
  hist(m, b) { const w = Array.from({ length: 16 }, (_, j) => (j === 12 ? 1 : 0.3 + 0.6 * hh(j, 12))); return { o: colsGen(b, w, 0.8, 10, 4.2), sig: [...m(0.8, 0.05), 0] }; },
  line(m) { return chartLine(m, 4, 2.2, 3); },
  twolines(m) { const a = chartLine(m, 4, 2.2, 2), b = chartLine(m, 9, 0.8, 2); return { o: a.o.concat(b.o), sig: b.sig }; },
  ring(m, b) { const o = [], cx = (b[0] + b[1]) / 2, cy = (b[2] + b[3]) / 2, rx = (b[1] - b[0]) * 0.34, ry = (b[3] - b[2]) * 0.42;
    for (let i = 0; i < 2200; i++) { const a = i / 2200 * 6.283 * 3, ring = i % 3, k = 1 - ring * 0.07; o.push([cx + Math.cos(a) * rx * k, cy + Math.sin(a) * ry * k, 3.4, 0]); } return { o, sig: [cx, cy, 11] }; },
  target(m, b) { const o = [], cx = (b[0] + b[1]) / 2, cy = (b[2] + b[3]) / 2; let n = 0;
    for (let k = 1; k <= 6; k++) { const R = k / 6, c = Math.round(2200 * k / 21); for (let i = 0; i < c; i++) { const a = i / c * 6.283; o.push([cx + Math.cos(a) * (b[1] - b[0]) * 0.32 * R, cy + Math.sin(a) * (b[3] - b[2]) * 0.46 * R, 3, 0]); } }
    return { o, sig: [cx, cy, 12] }; },
  conv(m, b) { const o = [], cx = (b[0] + b[1]) / 2, cy = (b[2] + b[3]) / 2; for (let i = 0; i < 2300; i++) { const a = hh(i, 1) * 6.283, d = 70 * Math.sqrt(hh(i, 2)); o.push([cx + Math.cos(a) * d, cy + Math.sin(a) * d, 2.4, 0]); } return { o, sig: [cx, cy, 14] }; },
  row1(m) { const o = []; for (let i = 0; i < 700; i++) o.push([...m(i / 699, 0.5), 3.4, 0]); return { o, sig: [...m(0.12, 0.5), 12] }; },
  rows1(m) { const o = []; for (const v of [0.4, 0.5, 0.6]) for (let i = 0; i < 700; i++) o.push([...m(i / 699, v), 3.4, 0]); return { o, sig: [...m(0.5, 0.5), 12] }; },
  rows3(m) { const o = []; for (const v of [0.12, 0.55, 0.92]) for (let i = 0; i < 740; i++) o.push([...m(i / 739, v), 3.4, 0]); return { o, sig: [...m(0.5, 0.55), 12] }; },
  bar1(m, b) { const o = [], cx = (b[0] + b[1]) / 2, w = 420, h = 440, s = 9; for (let y = 0; y < h; y += s) for (let x = 0; x < w; x += s) o.push([cx - w / 2 + x + s / 2, b[3] - h + y + s / 2, 4.2, 0]); return { o, sig: [cx, b[3] - h / 2, 12] }; },
  car() {
    const EL = world.EL, gx = 960, gy = 950, s = 330, Q = ([x, y]) => [gx + x * s, gy - y * s], lines = [EL.body, EL.glass, ...EL.lines];
    for (const w of EL.wheels) { lines.push(...w.rings); if (w.spokes) lines.push(...w.spokes); }
    const segs = []; let tot = 0;
    for (const pl of lines) { const q = pl.map(Q); for (let i = 0; i < q.length - 1; i++) { const l = Math.hypot(q[i + 1][0] - q[i][0], q[i + 1][1] - q[i][1]); if (l > 0) { segs.push([tot, l, q[i], q[i + 1]]); tot += l; } } }
    const o = []; let si = 0;
    for (let k = 0; k < 2250; k++) { const t = (k + 0.5) / 2250 * tot; while (si < segs.length - 1 && segs[si][0] + segs[si][1] < t) si++; const [a, l, p, q] = segs[si], f = (t - a) / l; o.push([lerp(p[0], q[0], f), lerp(p[1], q[1], f), 2.7, 0]); }
    return { o, sig: [...Q(EL.FA), 10] };
  },
};
const FORM = {};
function finalize(o, sig) {
  let pts = o.slice();
  if (pts.length > N - 1) { const st = pts.length / (N - 1); pts = Array.from({ length: N - 1 }, (_, k) => pts[Math.floor(k * st)]); }
  const n0 = pts.length; for (let k = 0; pts.length < N - 1; k++) { const q = pts[Math.floor(hh(k, 3) * n0)]; pts.push([q[0], q[1], 0, q[3]]); }
  pts.sort((a, b) => (a[0] + a[1] * 0.2) - (b[0] + b[1] * 0.2));
  const F = { x: new Float32Array(N), y: new Float32Array(N), r: new Float32Array(N), g: new Uint8Array(N) };
  F.x[0] = sig[0]; F.y[0] = sig[1]; F.r[0] = sig[2];
  pts.forEach((p, i) => { F.x[i + 1] = p[0]; F.y[i + 1] = p[1]; F.r[i + 1] = p[2]; F.g[i + 1] = p[3]; });
  return F;
}
function build(key) {
  if (FORM[key]) return FORM[key];
  const [name, bx] = key.split('|'), b = BOX[bx || 'full'], res = GEN[name](mapOf(b), b);
  let sig = res.sig; if (!sig[2]) sig = [sig[0], sig[1], 12];
  return (FORM[key] = finalize(res.o, sig));
}

// ------------------------------------------------------------------ the word
function geom(ctx, sh) {
  if (sh._g) return sh._g;
  const str = sh.word, size = sh.size || M.fitSize(ctx, str, 800, DISPLAY, sh.maxW || 1728, 420, -0.03);
  const f = M.font(size, 800, DISPLAY), lay = M.layout(ctx, str, f, -0.03 * size), width = lay.width;
  const pos = sh.pos || 'top', cy = (pos === 'top' ? 96 + 0.36 * size : pos === 'bot' ? 984 - 0.36 * size : H / 2) + (sh.dy || 0);
  const x0 = (W - width) / 2 + (sh.dx || 0), base = cy + 0.36 * size;
  return (sh._g = { size, f, lay, width, x0, base, cy, period: str.endsWith('.') });
}
function wordForm(ctx, sh) {
  const g = geom(ctx, sh), cv = (wordForm.cv ||= document.createElement('canvas'));
  cv.width = W; cv.height = H; const c = cv.getContext('2d', { willReadFrequently: true });
  c.clearRect(0, 0, W, H); c.font = g.f; c.fillStyle = '#fff'; c.letterSpacing = `${-0.03 * g.size}px`; c.textBaseline = 'alphabetic';
  const body = g.period ? sh.word.slice(0, -1) : sh.word; c.fillText(body, g.x0, g.base);
  const d = c.getImageData(0, 0, W, H).data; let cnt = 0;
  for (let y = 0; y < H; y += 3) for (let x = 0; x < W; x += 3) if (d[(y * W + x) * 4 + 3] > 127) cnt++;
  const s = clamp(Math.sqrt(cnt * 9 / 2000), 3.2, 14), o = [];
  for (let j = 0, y = 4; y < H; j++, y = 4 + j * s * 0.866) for (let x = (j & 1) * s / 2 + 2; x < W; x += s) if (d[((y | 0) * W + (x | 0)) * 4 + 3] > 127) o.push([x, y, s * 0.52, 0]);
  let sig = null;
  if (g.period) { const gl = g.lay.glyphs[g.lay.glyphs.length - 1], r = g.size * 0.088; sig = [g.x0 + gl.cx, g.base - r, r]; }
  if (!sig) { const q = o[Math.floor(o.length * 0.4)]; sig = [q[0], q[1], 12]; }
  return finalize(o, sig);
}

function drawWord(ctx, sh, c, g) {
  const L = c.L, ls = c.ls;
  let sc = 1, oy = 0;
  const q = E.outBack(prog(ls, 0, 0.14), 2.6);
  if (sh.anim === 'slam') sc = 1 + 0.22 * (1 - q); else if (sh.anim === 'U') oy = (1 - E.outExpo(prog(ls, 0, 0.16))) * 90;
  ctx.save(); ctx.translate(W / 2, g.cy + oy); ctx.scale(sc, sc); ctx.translate(-W / 2, -g.cy);
  if (sh.plate) { ctx.fillStyle = L.bg; ctx.fillRect(g.x0 - 36, g.base - 0.8 * g.size, g.width + 72, g.size * 1.02); }
  if (sh.wordFn) sh.wordFn(ctx, c, g);
  else {
    const body = g.period ? sh.word.slice(0, -1) : sh.word;
    ctx.font = g.f; ctx.letterSpacing = `${-0.03 * g.size}px`; ctx.fillStyle = L.ink; ctx.textAlign = 'left'; ctx.fillText(body, g.x0, g.base); ctx.letterSpacing = '0px';
    if (g.period) { const gl = g.lay.glyphs[g.lay.glyphs.length - 1], r = g.size * 0.088; disc(ctx, g.x0 + gl.cx, g.base - r, r, L.acc); }
  }
  ctx.restore();
}

// ------------------------------------------------------------------ dots
const PX = new Float32Array(N), PY = new Float32Array(N), PR = new Float32Array(N), BOW = new Float32Array(N), SD = new Float32Array(N);
for (let i = 0; i < N; i++) { BOW[i] = (hh(i, 9) - 0.5) * 220; SD[i] = hh(i, 5) * 6.283; }
function drawDots(ctx, sh, c, u) {
  const L = c.L, lt = c.lt, fs = sh.fs; let j = 0;
  for (let k = 0; k < fs.length; k++) if (fs[k][0] <= lt) j = k;
  const at = fs[j][0], B = FORM[fs[j][1]], dur = (fs[j][2] || 0.6);
  const Akey = j > 0 ? fs[j - 1][1] : sh.prevKey, A = Akey ? FORM[Akey] : null;
  const m = sh.settled && j === 0 ? 1 : clamp((lt - at) / dur), moving = m < 1 && A && A !== B, popping = m < 1 && !A;
  const drift = sh.calm ? 1.2 : 2.6, ub = sh.cols ? 0.25 : 1;
  const hi = sh.hi, soft = !sh.dotWord && !sh.hard;
  let colsMode = !!sh.cols, base = 1000;
  for (let i = 0; i < N; i++) {
    let x = B.x[i], y = B.y[i], r = B.r[i];
    if (moving) {
      const d = 0.4 * i / N, q = E.inOutCubic(clamp((m - d) / 0.6)), ax = A.x[i], ay = A.y[i], dx = x - ax, dy = y - ay, len = Math.hypot(dx, dy) || 1, bw = BOW[i] * Math.sin(Math.PI * q);
      x = ax + dx * q - dy / len * bw; y = ay + dy * q + dx / len * bw; r = lerp(A.r[i], r, q);
    } else if (popping) r *= E.outBack(clamp((m - 0.4 * i / N) / 0.6), 2.2);
    if (i > 0) { x += drift * ub * Math.sin(u * 1.7 + SD[i]); y += drift * ub * Math.cos(u * 1.3 + SD[i]); }
    if (colsMode && i > 0) {
      const g = B.g[i], sc = 1 + 0.03 * Math.sin(u * 1.8 + g * 1.3) + (g === hi ? 0.05 + 0.12 * (1 - E.outExpo(prog(c.ls, 0, 0.35))) : 0);
      y = base + (y - base) * sc;
    }
    PX[i] = x; PY[i] = y; PR[i] = r;
  }
  const col = L.data;
  for (let pass = 0; pass < (soft ? 2 : 1); pass++) {
    ctx.fillStyle = col; ctx.globalAlpha = soft && pass === 1 ? 0.5 : 1; ctx.beginPath();
    for (let i = 1; i < N; i++) { if (soft && (i % 3 === 0) !== (pass === 1)) continue; const r = PR[i]; if (r < 0.3 || (hi !== undefined && FORM[fs[j][1]].g[i] === hi)) continue; ctx.moveTo(PX[i] + r, PY[i]); ctx.arc(PX[i], PY[i], r, 0, 7); }
    ctx.fill();
  }
  ctx.globalAlpha = 1;
  if (hi !== undefined) { ctx.fillStyle = L.acc; ctx.beginPath(); for (let i = 1; i < N; i++) { const r = PR[i]; if (r < 0.3 || B.g[i] !== hi) continue; ctx.moveTo(PX[i] + r, PY[i]); ctx.arc(PX[i], PY[i], r, 0, 7); } ctx.fill(); }
  const sr = PR[0] * (1 + 0.12 * Math.sin(u * 3.1)); disc(ctx, PX[0], PY[0], sr, L.acc);
  return [PX[0], PY[0]];
}

// ------------------------------------------------------------------ custom data images
const LROWS = 28, LP = 34, LY0 = 56, KR = 22, ROWY = LY0 + KR * LP - 6;
function ledgerWall(ctx, c) {
  const L = c.L, ls = c.ls, kn = E.outExpo(prog(c.lt, 1.0, 1.5));
  for (let r = 0; r < LROWS; r++) {
    const y = LY0 + r * LP, q = E.outExpo(prog(c.lt, r * 0.012, r * 0.012 + 0.5)), sh = (ls * 26) % 1728, a = (1 - 0.5 * kn) * 0.3;
    if (r === KR && kn > 0) continue;
    ctx.fillStyle = rgba(L.data, a);
    let x = 96; const w = [110, 120 + hh(r, 1) * 520, 90 + hh(r, 2) * 160];
    ctx.fillRect(x, y - 6, w[0] * q, 12); x += w[0] + 30;
    ctx.fillRect(x, y - 6, w[1] * q, 12); x = 1824 - w[2] * q;
    ctx.fillRect(x, y - 6, w[2] * q, 12);
    ctx.fillStyle = rgba(L.data, a * 0.5); ctx.fillRect(96 + ((r * 97 + sh) % 1500), y - 2, 36 * q, 4);
  }
  if (kn > 0) { ctx.fillStyle = C.red; ctx.fillRect(96, ROWY, 1728 * kn, 16); }
}
let CELLS = null;
function buildCells() {
  const cw = 1728 / 43, ch = 33.6, cnt = new Float32Array(43 * 25);
  for (const car of world.LOT) { const ci = Math.floor((car.x + 85) / 4), ri = Math.floor((car.y + 50) / 4); if (ci >= 0 && ci < 43 && ri >= 0 && ri < 25) cnt[ri * 43 + ci]++; }
  CELLS = [];
  for (let ri = 0; ri < 25; ri++) for (let ci = 0; ci < 43; ci++) {
    let v = clamp(cnt[ri * 43 + ci] / 2.2) + 0.12 * hh(ri * 43 + ci, 2);
    if (ri >= 1 && ri <= 6 && ci >= 7 && ci <= 35) v = 0.55 + 0.45 * hh(ri * 43 + ci, 4);
    CELLS.push({ x: 96 + ci * cw, y: 118 + ri * ch, ci, ri, v: clamp(v), b: Math.min(3, Math.floor(v * 4)), hi: ci === 21 && ri === 15 });
  }
}
function heatGrid(ctx, c, noWave) {
  const L = c.L, cw = 1728 / 43 - 4, ch = 29.6;
  for (let b = 0; b < 4; b++) {
    ctx.fillStyle = rgba(L.data, [0.1, 0.3, 0.6, 1][b]); ctx.beginPath();
    for (const e of CELLS) {
      if (e.b !== b || e.hi) continue;
      const q = noWave ? 1 : E.outBack(prog(c.lt, 0.05 + (e.ci / 43) * 1.1 + hh(e.ri * 43 + e.ci, 6) * 0.15, 0.4 + (e.ci / 43) * 1.1 + hh(e.ri * 43 + e.ci, 6) * 0.15));
      if (q <= 0) continue; const w = cw * q, h = ch * q; ctx.rect(e.x + (cw - w) / 2, e.y + (ch - h) / 2, w, h);
    }
    ctx.fill();
  }
  const hc = CELLS.find((e) => e.hi);
  if (noWave) {
    const q = E.outBack(prog(c.lt, 0.1, 0.5));
    ctx.fillStyle = C.red; ctx.fillRect(hc.x + cw / 2 - cw * q * 0.5, hc.y + ch / 2 - ch * q * 0.5, cw * q, ch * q);
    ringS(ctx, hc.x + cw / 2, hc.y + ch / 2, 20 + 60 * (1 - E.outExpo(prog(c.lt, 0.3, 1.1))), 4, C.red);
  }
}
function netLines(ctx, c, box) {
  const m = mapOf(box), CL = K.CLUSTERS, pairs = [[0, 1], [1, 2], [2, 3], [3, 4], [4, 5], [5, 6], [6, 0], [0, 5], [1, 5], [2, 4], [2, 5], [1, 6], [3, 5], [0, 2]];
  ctx.save(); ctx.lineJoin = 'round'; ctx.strokeStyle = c.L.data; ctx.lineWidth = 3;
  pairs.forEach(([a, b], i) => { const q = E.outExpo(prog(c.lt, 0.1 + i * 0.07, 0.1 + i * 0.07 + 0.5)); K.drawOn(ctx, [m(CL[a].x / 1920, CL[a].y / 1080), m(CL[b].x / 1920, CL[b].y / 1080)], q); });
  ctx.restore();
  CL.forEach((e, i) => { const q = prog(c.lt, i * 0.06, i * 0.06 + 0.3); if (q > 0) { const [x, y] = m(e.x / 1920, e.y / 1080); disc(ctx, x, y, 14 * E.outBack(q, 3), c.L.data); } });
}
function matrix(ctx, c) {
  const L = c.L; const un = [], fl = [];
  for (let j = 0; j < 25; j++) for (let i = 0; i < 52; i++) {
    const x = 96 + (i + 0.5) * 1728 / 52, y = 508 + j * 19.6, q = E.outBack(prog(c.lt, 0.1 + (i / 52) * 1.5 + 0.06 * Math.sin(j * 0.7 + i * 0.2), 0.5 + (i / 52) * 1.5 + 0.06 * Math.sin(j * 0.7 + i * 0.2)), 2);
    (q < 0.5 ? un : fl).push([x, y, q]);
  }
  ctx.fillStyle = rgba(L.dim, 0.7); ctx.beginPath(); for (const [x, y] of un) { ctx.moveTo(x + 2, y); ctx.arc(x, y, 2, 0, 7); } ctx.fill();
  ctx.fillStyle = L.data; ctx.beginPath(); for (const [x, y, q] of fl) { const r = 2 + q * 5.6; ctx.moveTo(x + r, y); ctx.arc(x, y, r, 0, 7); } ctx.fill();
  disc(ctx, 96 + 26.5 * 1728 / 52, 508 + 12 * 19.6, 8 * E.outBack(prog(c.lt, 1.8, 2.2)), C.red);
}
const BARS = [[360, 340, 300, '1,5 %', '2023'], [780, 340, 345, '1,7 %', '2024'], [1200, 340, 450, '2,5 %', 'BRUKT 2024']];
function marginBars(ctx, c, stage) {
  const L = c.L;
  BARS.forEach(([x, w, h, val, yr], k) => {
    if (stage === 1 && k === 2) return;
    const t0 = stage === 1 ? [0.2, 1.0][k] : k === 2 ? 0.1 : -9, q = E.outExpo(prog(c.lt, t0, t0 + 0.6)); if (q <= 0) return;
    const hh_ = h * q, top = 1000 - hh_, strong = stage === 1 || k === 2, col = k === 2 ? (L.id === 'N' ? C.cloud : L.data) : rgba(L.data, strong ? (k ? 1 : 0.55) : 0.35);
    ctx.fillStyle = col; ctx.fillRect(x, top, w, hh_);
    if (k === 1) { ctx.fillStyle = C.red; ctx.fillRect(x, top, w, 14); }
    mono(ctx, val, x + w / 2, top - 22, { size: 34, col: L.data, q: q });
    mono(ctx, yr, x + w / 2, 1000 - 24, { size: 28, col: L.bg, q: q });
  });
}
const ODO = {};
function prepOdo(ctx, sh, g) {
  const cv = document.createElement('canvas'), step = Math.round(g.size * 0.98); ctx.font = g.f; ctx.letterSpacing = '0px';
  const wD = Math.ceil(ctx.measureText('7').width) + 16; cv.width = wD; cv.height = step * 5;
  const c = cv.getContext('2d'); c.font = g.f; c.fillStyle = sh.L.ink; c.textAlign = 'left';
  for (let d = 4; d <= 8; d++) c.fillText(String(d), 4, (d - 4) * step + Math.round(0.82 * g.size));
  ODO.cv = cv; ODO.step = step; ODO.w = wD;
}
function odometer(ctx, c, g) {
  const L = c.L; ctx.font = g.f; ctx.letterSpacing = `${-0.03 * g.size}px`; ctx.fillStyle = L.ink; ctx.textAlign = 'left';
  const wPre = ctx.measureText('1,').width, wD = ctx.measureText('7').width;
  ctx.fillText('1,', g.x0, g.base);
  const val = 5 + 2 * E.inOutCubic(prog(c.lt, 1.6, 3.2)), x = g.x0 + wPre;
  ctx.drawImage(ODO.cv, 0, (val - 4) * ODO.step, ODO.w, ODO.step * 0.9, x - 4, g.base - 0.82 * g.size, ODO.w, ODO.step * 0.9);
  ctx.fillText(' %', x + wD, g.base); ctx.letterSpacing = '0px';
}
function timeline(ctx, c) {
  const L = c.L, xs = Array.from({ length: 10 }, (_, i) => 160 + i * 1600 / 9);
  ctx.strokeStyle = rgba(L.dim, 0.9); ctx.lineWidth = 3; K.drawOn(ctx, [[96, 900], [1824, 900]], E.outExpo(prog(c.lt, 0, 0.6)));
  xs.forEach((x, i) => { const q = E.outExpo(prog(c.lt, 0.1 + i * 0.08, 0.5 + i * 0.08)); ctx.strokeStyle = L.data; ctx.lineWidth = 4; K.drawOn(ctx, [[x, 900], [x, 900 + 34]], q); });
  // ten year marks only: the analysis' values are not reproduced here, so no curve is drawn
  const p = E.outCubic(prog(c.lt, 0.2, 2.2)), fi = 9 * p;
  ctx.font = M.font(30, 500, K.MONO); ctx.textAlign = 'center'; ctx.fillStyle = L.data;
  xs.forEach((x, i) => { if (i <= fi) { disc(ctx, x, 900, 9, L.data); ctx.fillText(String(2015 + i), x, 985); } });
  ctx.textAlign = 'left';
  if (p > 0.02) disc(ctx, lerp(xs[0], xs[9], p), 900, 16, C.red);
}

// ------------------------------------------------------------------ shot table
const SH = [];
const S = (b, word, look, o = {}) => SH.push(Object.assign({ b, word, L: LK[look] }, o));
// cold open: du har tallene
S(0, 'DU.', 'N', { F: 'self', settled: 1, dotWord: 1, pos: 'mid', sfx: 'impact', shake: 14 });
S(2, 'HAR.', 'A', { im: ledgerWall, pos: 'top', sfx: 'thud', tail: { type: 'bar', rect: [96, ROWY, 1728, 16], b: 0.6 }, hits: [[1.0, 'red row', 'pop']] });
S(4, 'TALLENE.', 'B', { F: [[0, 'ledger|low'], [2, 'band|low']], sfx: 'thud' });
S(8, 'TUSENVIS.', 'N', { F: [[0, 'burst|full'], [2, 'field|full']], pos: 'mid', plate: 1, sfx: 'impact', shake: 16 });
S(12, 'HVER.', 'A', { F: 'grid|full', pos: 'mid', plate: 1, sfx: 'click' });
S(14, 'DAG.', 'N', { F: 'ring|up', pos: 'bot', sfx: 'thud', shake: 10 });
// six systems as six pulsing columns
[['NYBIL.', 'A'], ['BRUKTBIL.', 'N'], ['VERKSTED.', 'B'], ['DELER.', 'A'], ['GARANTI.', 'N']].forEach(([w, l], i) => S(16 + i, w, l, { F: 'cols6|low', cols: 1, hard: 1, hi: i, sfx: 'click', shake: 8 }));
S(21, 'SERVICE.', 'B', { F: 'cols6|low', cols: 1, hard: 1, hi: 5, sfx: 'click', shake: 8 });
S(22, 'SEKS.', 'N', { F: 'cols6|low', cols: 1, hard: 1, sfx: 'thud', shake: 10 });
S(24, 'SYSTEMER.', 'A', { F: [[0, 'cols6|low'], [1, 'rows3|low']], cols: 1, hard: 1, sfx: 'thud' });
S(26, 'ETT.', 'B', { F: 'bar1|low', sfx: 'thud', shake: 14 });
S(28, 'BILDE?', 'N', { im: (ctx, c) => heatGrid(ctx, c, false), pos: 'mid', plate: 1, sfx: 'swell' });
// technology
S(32, 'VI.', 'B', { F: 'field|low', sfx: 'thud', shake: 12 });
S(33, 'HENTER.', 'N', { F: [[0, 'burst|full'], [1, 'conv|low']], sfx: 'whoosh' });
S(35, 'KOBLER.', 'A', { F: 'field|low', im: (ctx, c) => netLines(ctx, c, BOX.low), sfx: 'click' });
S(38, 'ANALYSERER.', 'B', { F: [[0, 'ledger|low'], [2, 'hist|low'], [4, 'bar1|low']], sfx: 'thud', tail: { type: 'bar', rect: [750, 560, 420, 440], b: 0.7 } });
S(44, 'ALT.', 'R', { F: 'rows3|low', sfx: 'impact', shake: 26 });
S(45, 'PÅ.', 'R', { F: 'rows3|low', sfx: 'click' });
S(46, 'ÉN.', 'R', { F: 'rows1|low', sfx: 'impact', shake: 18 });
S(47, 'PLATTFORM.', 'R', { F: 'rows1|low', sfx: 'thud', shake: 18, kick: 'BDO DATAPLATTFORM' });
// industry knowledge
S(48, 'MEN.', 'N', { F: 'rows1|low', sfx: 'thud' });
S(49, 'TALL.', 'A', { F: 'band|up', pos: 'bot', sfx: 'click' });
S(50, 'TRENGER.', 'B', { F: 'band|up', pos: 'bot', sfx: 'click' });
S(51, 'KONTEKST.', 'N', { F: 'field|full', pos: 'mid', plate: 1, sfx: 'thud', shake: 12 });
S(53, 'TITLE', 'N', { F: 'hist|low', title: 1, sfx: 'impact', shake: 14 });
S(56, '1 300+', 'A', { im: matrix, label: 'BILBRANSJEBEDRIFTER ANALYSERT', sfx: 'thud', hits: [[1.8, 'red dot', 'pop']] });
S(59, '1,7 %', 'B', { im: (ctx, c) => marginBars(ctx, c, 1), size: 360, wordFn: odometer, label: 'DRIFTSMARGIN MERKEFORHANDLERE 2023 → 2024', kick2: 'ETTERLENGTET ØKNING', sfx: 'thud', hits: [[1.0, 'bar 2', 'click'], [1.6, 'odometer', 'riser'], [3.2, 'lands', 'bell']] });
S(63, '2,5 %', 'N', { size: 340, im: (ctx, c) => marginBars(ctx, c, 2), label: 'DRIFTSMARGIN BRUKTBILFORHANDLERE 2024', sfx: 'thud', tail: { type: 'bar', rect: [1200, 560, 340, 440], b: 0.6 } });
S(65, 'SMÅ.', 'A', { F: 'sizes|low', cols: 1, hard: 1, hi: 0, label: 'FORHANDLERE SAMMENLIGNET ETTER STØRRELSE', sfx: 'click' });
S(66, 'MELLOMSTORE.', 'B', { F: 'sizes|low', cols: 1, hard: 1, hi: 1, label: 'FORHANDLERE SAMMENLIGNET ETTER STØRRELSE', sfx: 'click' });
S(68, 'STORE.', 'N', { F: 'sizes|low', cols: 1, hard: 1, hi: 2, label: 'FORHANDLERE SAMMENLIGNET ETTER STØRRELSE', sfx: 'thud' });
S(69, '10 ÅR.', 'A', { im: timeline, label: 'OMSETNING OG DRIFTSMARGIN I ET 10-ÅRS PERSPEKTIV', sfx: 'thud', tail: { type: 'line', b: 0.7 } });
S(72, 'HVOR.', 'B', { F: 'band|low', sfx: 'thud', shake: 10 });
S(73, 'STÅR.', 'N', { F: 'band|low', ringSig: 1, sfx: 'click' });
S(74, 'DU?', 'A', { F: 'band|low', ringSig: 1, big: 1, sfx: 'swell', tail: { type: 'ring', b: 0.7 } });
S(76, 'VI.', 'R', { F: 'self', dotWord: 1, pos: 'mid', sfx: 'impact', shake: 28 });
S(77, 'KJENNER.', 'R', { F: 'self', dotWord: 1, pos: 'mid', sfx: 'thud', shake: 14 });
S(78, 'BRANSJEN.', 'R', { F: 'self', dotWord: 1, pos: 'mid', sfx: 'thud', shake: 14 });
// efficient audit
S(80, 'MØNSTRE.', 'N', { F: [[0, 'band|low'], [0.4, 'waves|up']], sfx: 'thud', shake: 10 });
S(81, 'TREDER.', 'A', { F: 'waves|up', pos: 'bot', sfx: 'click' });
S(82, 'FRAM.', 'B', { F: [[0, 'waves|up'], [0.3, 'target|up']], pos: 'bot', sfx: 'thud' });
S(83, 'ÉN.', 'N', { F: 'band|low', sfx: 'click' });
S(84, 'TING.', 'A', { F: 'band|low', sfx: 'click' });
S(85, 'SKILLER.', 'B', { F: 'band|low', ringSig: 1, sfx: 'thud' });
S(86, 'SEG UT.', 'N', { F: 'band|low', ringSig: 2, big: 1, tag: 'Garantikrav uten arbeidsordre', sfx: 'thud', shake: 14 });
S(88, 'REVISOR.', 'N', { F: 'ring|low', calm: 1, sfx: 'swell' });
S(90, 'VURDERER.', 'A', { F: 'ring|up', pos: 'bot', calm: 1, sfx: 'bell' });
S(92, 'MÅLRETTET.', 'B', { F: 'target|up', pos: 'bot', sfx: 'thud', shake: 10 });
S(94, 'EFFEKTIVT.', 'N', { F: 'rows1|low', sfx: 'thud', tail: { type: 'slam', b: 0.6 } });
// value
S(96, 'DU.', 'A', { im: (ctx, c) => heatGrid(ctx, c, false), pos: 'mid', plate: 1, sfx: 'thud', shake: 12 });
S(97, 'FÅR.', 'B', { im: (ctx, c) => heatGrid(ctx, c, true), pos: 'mid', plate: 1, sfx: 'click' });
S(98, 'OVERSIKT.', 'N', { F: 'cols6|low', cols: 1, hard: 1, sfx: 'thud' });
S(100, 'DIALOG.', 'A', { F: 'twolines|low', sfx: 'thud', tail: { type: 'line', b: 0.5 } });
S(102, 'INNSIKT.', 'B', { F: 'line|low', sfx: 'thud', shake: 10 });
S(104, 'I.', 'N', { F: 'grid|full', pos: 'mid', plate: 1, sfx: 'click' });
S(105, 'DINE.', 'A', { F: 'field|full', pos: 'mid', plate: 1, sfx: 'click' });
S(106, 'EGNE.', 'B', { F: 'conv|low', sfx: 'click' });
S(107, 'DATA.', 'N', { F: 'self', dotWord: 1, pos: 'mid', sfx: 'thud', shake: 14 });
S(108, 'MER.', 'R', { F: 'stairs|low', sfx: 'impact', shake: 30 });
S(110, 'VERDI.', 'R', { F: 'stairs|low', sfx: 'impact', shake: 30 });
// BDO Automotive: three hard accents
S(112, 'BDO.', 'N', { F: 'self', dotWord: 1, pos: 'mid', sfx: 'impact', shake: 20 });
S(113, 'AUTOMOTIVE.', 'B', { F: 'self', dotWord: 1, pos: 'mid', sfx: 'impact', shake: 20 });
S(114, 'BDO AUTOMOTIVE.', 'A', { F: 'conv|low', sfx: 'impact', shake: 18 });
SH.forEach((s, i) => {
  s.n = (i + 1 < SH.length ? SH[i + 1].b : END) - s.b; s.i = i;
  if (typeof s.F === 'string') s.F = [[0, s.F]];
  if (s.F) s.F = s.F.map(([a, k, d]) => [a, k === 'self' ? 'W' + i : k, d]);
  const pv = SH[i - 1]; s.prevKey = pv && pv.F ? pv.F[pv.F.length - 1][1] : null;
  if (s.F) s.fs = s.F;
  if (s.title) s.dotWord = 0;
});
const STARTS = SH.map((s) => s.b);
const shotAt = (u) => { let i = 0; for (let k = 0; k < SH.length; k++) if (STARTS[k] <= u) i = k; else break; return SH[i]; };

// ------------------------------------------------------------------ tails (covers that end in the next shot's flat)
function tail(ctx, sh, lt, sigPos) {
  const T = sh.tail; if (!T) return; const nx = SH[sh.i + 1]; if (!nx) return;
  const q = prog(lt, sh.n - (T.b || 0.6), sh.n); if (q <= 0) return;
  ctx.save(); ctx.setTransform(1, 0, 0, 1, 0, 0); ctx.fillStyle = nx.L.bg;
  if (T.type === 'bar') { const e = E.inCubic(q), [x, y, w, h] = T.rect; ctx.fillRect(lerp(x, -2, e), lerp(y, -2, e), lerp(w, W + 4, e), lerp(h, H + 4, e)); }
  else if (T.type === 'ring') { const [x, y] = sigPos || [W / 2, H / 2]; disc(ctx, x, y, E.inQuad(q) * 2200 + 10, nx.L.bg); }
  else if (T.type === 'slam') { const e = E.inCubic(q); ctx.fillRect(0, 0, W, e * (H / 2 + 2)); ctx.fillRect(0, H - e * (H / 2 + 2), W, e * (H / 2 + 2)); }
  else if (T.type === 'line') {
    const e = E.inQuad(q), ex = (y) => e * (W + 500) - 250 + 90 * K.noise1(y / 70, 5), pts = [];
    for (let y = -40; y <= H + 40; y += 40) pts.push([ex(y), y]);
    ctx.beginPath(); ctx.moveTo(-400, -40); pts.forEach(([x, y]) => ctx.lineTo(x, y)); ctx.lineTo(-400, H + 40); ctx.closePath(); ctx.fill();
    ctx.lineJoin = 'round'; ctx.lineWidth = 12; ctx.strokeStyle = sh.L.ink === C.white ? C.white : sh.L.ink; ctx.beginPath(); pts.forEach(([x, y], i) => (i ? ctx.lineTo(x, y) : ctx.moveTo(x, y))); ctx.stroke();
  }
  ctx.restore();
}

// ------------------------------------------------------------------ frame
function titleWord(ctx, c) {
  const L = c.L, ls = c.ls;
  K.kicker(ctx, 'BDO', 96, 150, ls, { color: L.ink, size: 40 });
  const rise = (str, y, size, t0) => { const q = E.outExpo(prog(ls, t0, t0 + 0.3)); ctx.font = M.font(size, 800, DISPLAY); ctx.letterSpacing = `${-0.03 * size}px`; ctx.fillStyle = L.ink; ctx.textAlign = 'left'; ctx.fillText(str, 96, y + (1 - q) * 90); const w = ctx.measureText(str).width; ctx.letterSpacing = '0px'; return w; };
  rise('BILBRANSJEANALYSEN', 330, 150, 0);
  const w = rise('2025', 720, 400, 0.08);
  disc(ctx, 96 + w + 40, 720 - 24, 34, C.red);
}
function shot(ctx, u, t, sh) {
  const lt = u - sh.b, ls = lt * P, p = clamp(lt / sh.n), L = sh.L, c = { ls, lt, p, n: sh.n, L };
  const g = geom(ctx, sh); let sigPos = null;
  ctx.save();
  const fx = W / 2, fy = H / 2;
  const amp = sh.calm ? 1.012 : sh.big ? 1.09 : 1.05;
  const z = Math.exp(Math.log(amp) * p), punch = sh.calm ? 1 : 1 + 0.025 * (1 - E.outExpo(prog(ls, 0, 0.25))) * (sh.shake > 12 ? 1.6 : 1);
  ctx.translate(W / 2, H / 2); ctx.scale(z * punch, z * punch); ctx.translate(-fx, -fy);
  if (!sh.calm) M.shake(ctx, lt, 0, sh.shake ?? 6, (sh.b * 7 + 3) | 0, 0.12);
  paintBg(ctx, L, u);
  if (sh.im) sh.im(ctx, c);
  if (sh.F) sigPos = drawDots(ctx, sh, c, u);
  if (sh.ringSig && sigPos) {
    const k = sh.ringSig, q = E.outBack(prog(ls, 0.05, 0.3)), rr = (k === 2 ? 52 : 38) + 120 * (1 - q);
    ringS(ctx, sigPos[0], sigPos[1], rr, k === 2 ? 6 : 4, C.red);
    if (sh.tag) { const q2 = E.outCubic(prog(ls, 0.15, 0.4)); mono(ctx, sh.tag, clamp(sigPos[0], 96 + 330, 1824 - 330), sigPos[1] + 110, { size: 30, col: L.ink, q: q2 }); }
  }
  if (sh.title) titleWord(ctx, c);
  else if (!sh.dotWord) drawWord(ctx, sh, c, g);
  if (sh.label) {
    const q = E.outCubic(prog(ls, 0.15, 0.4)), yy = g.base + (/,/.test(sh.word) ? 112 : 70);
    mono(ctx, sh.label, W / 2, yy, { size: 40, col: L.ink, q });
  }
  if (sh.kick2) { const q = E.outCubic(prog(ls, 0.5, 0.8)); ctx.fillStyle = C.red; ctx.fillRect(W / 2 - 340, 545, 26 * q, 4); mono(ctx, sh.kick2, W / 2 + 20, 560, { size: 34, col: L.ink, q }); }
  if (sh.kick) { const q = E.outCubic(prog(ls, 0.3, 0.6)); mono(ctx, sh.kick, W / 2, g.base + 80, { size: 40, col: L.ink, q }); }
  ctx.restore();
  tail(ctx, sh, lt, sigPos);
}
function endCard(ctx, u) {
  const ls = (u - END) * P;
  ctx.fillStyle = C.white; ctx.fillRect(0, 0, W, H);
  K.line(ctx, 'Vi ser detaljene.', 120, 400, ls, { size: 140, weight: 300, color: C.charcoal, t0: -0.1, dur: 0.5 });
  K.line(ctx, 'Vi forstår helheten.', 120, 560, ls, { size: 140, weight: 300, color: C.charcoal, t0: 0.1, dur: 0.5 });
  const q = E.outCubic(prog(ls, 0.4, 0.8)); mono(ctx, 'BDO AUTOMOTIVE', 120, 720, { size: 40, col: C.blue, align: 'left', q });
  K.cornerstone(ctx, E.outExpo(prog(ls, 0.5, 1.4)));
}

// ------------------------------------------------------------------ hits
const HITS = [];
SH.forEach((s) => {
  HITS.push([s.b, s.word.toLowerCase(), s.sfx || 'thud']);
  (s.hits || []).forEach((h) => HITS.push([s.b + h[0], h[1], h[2]]));
  if (s.tag) HITS.push([s.b + 0.3, 'ring tag', 'blip']);
  if (s.tail) HITS.push([s.b + s.n - (s.tail.b || 0.6), 'cover', 'whoosh']);
});
HITS.push([END, 'end card', 'impact'], [END + 1.1, 'cornerstone', 'thud']);
HITS.sort((a, b) => a[0] - b[0]);

M.film({
  fonts: [M.font(100, 800, DISPLAY), M.font(100, 300, DISPLAY), M.font(40, 500, MONO)],
  hits: HITS,
  samples: (t) => (t / P >= END ? 1 : 3),
  async init(ctx) {
    await K.loadLogos();
    buildCells();
    for (const s of SH) {
      geom(ctx, s);
      if (s.wordFn === odometer) prepOdo(ctx, s, geom(ctx, s));
      if (s.F) for (const [, k] of s.F) { if (k === 'W' + s.i) FORM[k] = wordForm(ctx, s); else build(k); }
    }
  },
  draw(ctx, u) {
    if (u >= END) { endCard(ctx, u); return; }
    shot(ctx, Math.max(0, u), 0, shotAt(u));
  },
  post(ctx, u, t) { if (u < END && shotAt(u).L.id === 'N') K.grain(ctx, t, 0); },
});
