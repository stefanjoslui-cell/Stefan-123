// BDO Automotive – Fra støy til klarhet. A pure function of time: window.seek(t) paints frame t.
// ONE system of 6 000 particles carries the whole film. Every formation (noise, car, site plan, platform
// lattice, figures, calm bands, end ring) is a precomputed set of per-particle targets; each frame the
// particles travel between formations on curved, staggered, spring-eased paths. They are the data.
import * as M from './lib/motion.js';
import * as K from './lib/kit.js';
import { EL, GROUND, site } from './scenes/world.js';

const { W, H, C, clamp, lerp, prog, E, ss, rgba } = K;
const DBGSITE = false;
const N = 6000, PERIOD = 0.625, TAU = Math.PI * 2;
const rnd = K.mulberry32;

// ------------------------------------------------------------------ per-particle constants
const Z = new Float32Array(N), SB = new Float32Array(N), PH1 = new Float32Array(N), PH2 = new Float32Array(N);
const H2 = new Float32Array(N), H3 = new Float32Array(N), SG = new Float32Array(N), AMP = new Float32Array(N);
{
  const r = rnd(11);
  for (let i = 0; i < N; i++) {
    Z[i] = r(); SB[i] = 1.7 + r() * r() * 1.9; PH1[i] = r() * TAU; PH2[i] = r() * TAU;
    H2[i] = r(); H3[i] = r(); SG[i] = r() < 0.72 ? 1 : -1; AMP[i] = 0.1 + r() * 0.26;
  }
  Z[0] = 0.92; SB[0] = 2.6;
}

// ------------------------------------------------------------------ formations
// F[f] = { x, y, a (alpha mult), s (size mult), m (meta), ex, ey (extension vectors) } per particle
const NF = 9;
const F = [];
const mkF = () => ({ x: new Float32Array(N), y: new Float32Array(N), a: new Float32Array(N), s: new Float32Array(N), m: new Float32Array(N), ex: new Float32Array(N), ey: new Float32Array(N) });
for (let f = 0; f < NF; f++) F.push(mkF());

// transitions k = 1..8 (formation k-1 -> k): start, stagger spread, spring stiffness range, sort angle
const TR = [null,
  { ts: 4.2, spread: 1.5, lo: 9, hi: 15, ang: 0.3 },     // noise -> car          (lands ~6.2)
  { ts: 12.3, spread: 1.7, lo: 8, hi: 13, ang: 1.2 },    // car -> site plan      (lands ~14.5)
  { ts: 23.0, spread: 1.7, lo: 8, hi: 13, ang: 0.7 },    // site -> platform      (lands ~25.5, under the wipe)
  { ts: 32.6, spread: 1.5, lo: 9, hi: 15, ang: 2.0 },    // platform -> 1,7 %     (lands ~34.6)
  { ts: 36.4, spread: 0.55, lo: 26, hi: 40, ang: 0.2 },  // -> 2,5 %              (lands ~37.5)
  { ts: 38.5, spread: 0.55, lo: 26, hi: 40, ang: 1.6 },  // -> 1 300+             (lands ~39.5)
  { ts: 40.2, spread: 1.7, lo: 9, hi: 14, ang: 0.9 },    // -> calm bands         (lands ~42.4)
  { ts: 48.4, spread: 1.7, lo: 8, hi: 13, ang: 2.4 },    // -> end ring           (lands ~51)
];
const DL = [null]; for (let k = 1; k <= 8; k++) DL.push(new Float32Array(N));

// ---- F0: noise (analytic flow, a function of t only)
const BX = new Float32Array(N), BY = new Float32Array(N);
{
  const r = rnd(5);
  const g = () => Math.sqrt(-2 * Math.log(1 - r() * 0.999)) * Math.cos(TAU * r());
  for (let i = 0; i < N; i++) { const w = r() < 0.25 ? 1.8 : 1; BX[i] = clamp(960 + g() * 330 * w, -80, 2000); BY[i] = clamp(540 + g() * 200 * w, -60, 1140); }
  BX[0] = 1210; BY[0] = 400;
}
let TX = 0, TY = 0;
const linkGrow = (t, m) => prog(t, LT0 + (m - 1) * 0.17, LT0 + (m - 1) * 0.17 + 0.95);
function alphaAt(f, i, t) {
  if (f === 3 && F[3].m[i] > 0) return clamp((linkGrow(t, F[3].m[i]) * 1.15 - F[3].ex[i]) * 8);
  return F[f].a[i];
}
function noiseAt(i, t) {
  let x = BX[i], y = BY[i];
  let dx = x - 780, dy = y - 480, d = Math.hypot(dx, dy);
  let a = (0.5 * t + 0.6 * Math.sin(0.45 * t)) / (1 + d / 230);
  let c = Math.cos(a), s = Math.sin(a);
  x = 780 + dx * c - dy * s; y = 480 + dx * s + dy * c;
  dx = x - 1280; dy = y - 640; d = Math.hypot(dx, dy);
  a = -(0.3 * t + 0.4 * Math.sin(0.33 * t + 1)) / (1 + d / 260);
  c = Math.cos(a); s = Math.sin(a);
  x = 1280 + dx * c - dy * s; y = 640 + dx * s + dy * c;
  TX = x + 62 * Math.sin(0.8 * t + y * 0.009 + 1.3) + 28 * Math.sin(2.1 * t + y * 0.021 + PH1[i] * 0.2);
  TY = y + 56 * Math.sin(0.7 * t + x * 0.008) + 24 * Math.cos(1.9 * t + x * 0.02 + PH2[i] * 0.2);
}
const LT0 = 27.3;                       // first vertical link starts to grow
const RING = { cx: 960, cy: 480, rx: 700, ry: 325, w: 0.05, t0: 48 };
function tgt(f, i, t) {
  if (f === 0) { noiseAt(i, t); return; }
  const G = F[f];
  let x = G.x[i], y = G.y[i];
  if (f === 3) {
    y += 3.4 * Math.sin(x * 0.012 + t * 1.3 + y * 0.02);
  } else if (f === 8) {
    const th = G.m[i] + RING.w * (t - RING.t0), k = 1 + G.ex[i];
    x = RING.cx + RING.rx * k * Math.cos(th); y = RING.cy + RING.ry * k * Math.sin(th);
    if (G.ey[i] > 0) { x = G.x[i]; y = G.y[i]; }       // pinned (the red point)
  }
  TX = x; TY = y;
}

// ---- sampling helpers
function ambientPts(n, seed) {
  const r = rnd(900 + seed), out = [];
  for (let i = 0; i < n; i++) out.push({ x: 40 + r() * 1840, y: 40 + r() * 1000, a: 0.3, s: 0.75, m: 0, ex: 0, ey: 0 });
  return out;
}
function polyLen(p) { let L = 0; for (let i = 1; i < p.length; i++) L += Math.hypot(p[i][0] - p[i - 1][0], p[i][1] - p[i - 1][1]); return L; }
function alongPoly(p, n, r, jit, out, a = 1, s = 1) {
  if (n < 1) return;
  const cum = [0]; for (let i = 1; i < p.length; i++) cum.push(cum[i - 1] + Math.hypot(p[i][0] - p[i - 1][0], p[i][1] - p[i - 1][1]));
  const L = cum[cum.length - 1]; let seg = 0; const ph = r();
  for (let j = 0; j < n; j++) {
    const d = ((j + ph) / n) * L;
    while (seg < p.length - 2 && cum[seg + 1] < d) seg++;
    const f = (d - cum[seg]) / Math.max(1e-6, cum[seg + 1] - cum[seg]);
    out.push({ x: lerp(p[seg][0], p[seg + 1][0], f) + (r() - 0.5) * jit, y: lerp(p[seg][1], p[seg + 1][1], f) + (r() - 0.5) * jit, a, s, m: 0, ex: 0, ey: 0 });
  }
}
function pickPixels(data, w, h, test, n, r, ox = 0, oy = 0) {
  const idx = [];
  for (let p = 0, q = 0; p < w * h; p++, q += 4) if (test(data[q], data[q + 1], data[q + 2], data[q + 3])) idx.push(p);
  const out = [];
  for (let j = 0; j < Math.min(n, idx.length); j++) {
    const k = j + Math.floor(r() * (idx.length - j)); const t = idx[k]; idx[k] = idx[j]; idx[j] = t;
    out.push({ x: (t % w) + r() + ox, y: Math.floor(t / w) + r() + oy, a: 0.95, s: 1, m: 0, ex: 0, ey: 0 });
  }
  return out;
}
function nearest(pts, x, y) { let b = 0, bd = 1e12; for (let i = 0; i < pts.length; i++) { const d = (pts[i].x - x) ** 2 + (pts[i].y - y) ** 2; if (d < bd) { bd = d; b = i; } } return b; }

// ---- assignment: match particles to targets by locally sorted strips (approximate transport),
// so a formation change reads as a coherent pour, not a random shuffle.
function assign(f, pts, forced, ang) {
  const G = F[f], prevT = TR[f].ts, free = [];
  const px = new Float32Array(N), py = new Float32Array(N);
  for (let i = 0; i < N; i++) { tgt(f - 1, i, prevT); px[i] = TX; py[i] = TY; if (!forced[i]) free.push(i); }
  const r = rnd(300 + f);
  if (pts.length > free.length) { for (let j = pts.length - 1; j > 0; j--) { const k = Math.floor(r() * (j + 1)); [pts[j], pts[k]] = [pts[k], pts[j]]; } pts.length = free.length; }
  else if (pts.length < free.length) pts.push(...ambientPts(free.length - pts.length, f));
  const c = Math.cos(ang), s = Math.sin(ang);
  const ku = (x, y) => x * c + y * s, kv = (x, y) => -x * s + y * c;
  const A = free.slice().sort((a, b) => ku(px[a], py[a]) - ku(px[b], py[b]));
  const B = pts.map((_, i) => i).sort((a, b) => ku(pts[a].x, pts[a].y) - ku(pts[b].x, pts[b].y));
  const CH = 80;
  for (let o = 0; o < A.length; o += CH) {
    const a = A.slice(o, o + CH).sort((p, q) => kv(px[p], py[p]) - kv(px[q], py[q]));
    const b = B.slice(o, o + CH).sort((p, q) => kv(pts[p].x, pts[p].y) - kv(pts[q].x, pts[q].y));
    for (let j = 0; j < a.length; j++) {
      const i = a[j], P = pts[b[j]];
      G.x[i] = P.x; G.y[i] = P.y; G.a[i] = P.a; G.s[i] = P.s; G.m[i] = P.m; G.ex[i] = P.ex; G.ey[i] = P.ey;
    }
  }
  for (const i in forced) { const P = forced[i]; G.x[i] = P.x; G.y[i] = P.y; G.a[i] = P.a ?? 1; G.s[i] = P.s ?? 1; G.m[i] = P.m ?? 0; G.ex[i] = P.ex ?? 0; G.ey[i] = P.ey ?? 0; }
  // stagger: far travellers leave a little later, with individual jitter
  const D = DL[f], dist = new Float32Array(N); let mx = 1;
  for (let i = 0; i < N; i++) { tgt(f, i, prevT + 2); const d = Math.hypot(TX - px[i], TY - py[i]); dist[i] = d; if (d > mx) mx = d; }
  for (let i = 0; i < N; i++) D[i] = TR[f].spread * (0.5 * Math.pow(dist[i] / mx, 0.8) + 0.5 * H2[i]);
}

// ---- F0 base
for (let i = 0; i < N; i++) { F[0].a[i] = 1; F[0].s[i] = 1; }

// ---- F1: the car, side elevation, drawn in particles
const CAR = { s: 290, x: 960, gy: 600 };
const carX = (x) => CAR.x + x * CAR.s, carY = (y) => CAR.gy - y * CAR.s;
const DOT = { x: 0.2, y: 1.215 };
function buildCar() {
  const P = ([x, y]) => [carX(x), carY(y)], r = rnd(21), items = [];
  items.push([EL.body.map(P), 1, 1]); items.push([EL.glass.map(P), 1, 1]);
  EL.lines.forEach((l) => items.push([l.map(P), 0.8, 0.9]));
  for (const w of EL.wheels) { items.push([w.rings[0].map(P), 1, 1], [w.rings[1].map(P), 0.85, 0.9], [w.rings[2].map(P), 0.7, 0.9]); w.spokes.forEach((s) => items.push([s.map(P), 0.7, 0.9])); }
  items.push([[[carX(-2.95), carY(0)], [carX(2.95), carY(0)]], 0.55, 0.7]);
  const tot = items.reduce((a, [p, w]) => a + polyLen(p) * w, 0), n = 5350, out = [];
  for (const [p, w, al] of items) alongPoly(p, Math.round((n * polyLen(p) * w) / tot), r, 1.3, out, al, 1);
  return out;
}

// ---- F2: the dealership plan, sampled from site() rendered once
const PL = { s: 7.8, ox: 990, oy: 490 };            // world (0,0) at screen (ox, oy)
const planX = (wx) => PL.ox + wx * PL.s, planY = (wy) => PL.oy + wy * PL.s;
function buildSite() {
  const cv = document.createElement('canvas'); cv.width = W; cv.height = H;
  const c = cv.getContext('2d', { willReadFrequently: true });
  c.fillStyle = '#F5F7FA'; c.fillRect(0, 0, W, H);
  site(c, { k: 1, cx: (W / 2 - PL.ox) / PL.s, cy: (H / 2 - PL.oy) / PL.s, s: PL.s });
  window.__site = cv;
  const d = c.getImageData(0, 0, W, H).data, r = rnd(31);
  const lum = new Float32Array(W * H); for (let p = 0, q = 0; p < W * H; p++, q += 4) lum[p] = 0.3 * d[q] + 0.59 * d[q + 1] + 0.11 * d[q + 2];
  const edge = new Uint8Array(4 * W * H);          // line pixels: darker than a neighbour two px away, and not paper
  for (let y = 2; y < H - 2; y++) for (let x = 2; x < W - 2; x++) {
    const p = y * W + x, l = lum[p]; if (l > 228) continue;
    const m = Math.max(lum[p - 2], lum[p + 2], lum[p - 2 * W], lum[p + 2 * W]);
    if (l < 125) edge[p * 4 + 3] = 254; else if (m - l > 20) edge[p * 4 + 3] = 255;
  }
  const pts = pickPixels(edge, W, H, (R, G, B, A) => A === 254, 2300, r).concat(pickPixels(edge, W, H, (R, G, B, A) => A === 255, 3600, r));
  for (const p of pts) p.s = 1.2;
  for (const p of pts) { const wx = (p.x - PL.ox) / PL.s, wy = (p.y - PL.oy) / PL.s; p.m = wy > -46.4 && wy < -23.6 && wx > -57 && wx < 57 ? Math.floor((wx + 57) / 19) + 1 : 0; }
  return pts;
}

// ---- F3: the platform lattice (three layers + vertical links)
const LAY = { x0: 470, w: 1090, skew: 170, depth: 150, back: [190, 410, 630], cols: 88, rows: 15, gap: 220 };
const nodeX = (u, v) => LAY.x0 + LAY.skew * (1 - v) + u * LAY.w, nodeY = (l, v) => LAY.back[l] + v * LAY.depth;
const LAYERS = ['KJØRETØYDATA', 'TRANSAKSJONER', 'ØKONOMISKE DATA'];
const RED_NODE = { u: 0.42, v: 0.6 };
function buildLattice() {
  const out = [], r = rnd(41);
  for (let l = 0; l < 3; l++) for (let a = 0; a < LAY.cols; a++) for (let b = 0; b < LAY.rows; b++) {
    const u = a / (LAY.cols - 1), v = b / (LAY.rows - 1);
    out.push({ x: nodeX(u, v), y: nodeY(l, v), a: 0.95, s: 1.15, m: 0, ex: 0, ey: 0 });
  }
  const chains = [{ u: RED_NODE.u, v: RED_NODE.v }];
  while (chains.length < 24) chains.push({ u: Math.round(r() * (LAY.cols - 1)) / (LAY.cols - 1), v: Math.round(r() * (LAY.rows - 1)) / (LAY.rows - 1) });
  chains.forEach((ch, ci) => {
    const x = nodeX(ch.u, ch.v);
    for (let g = 0; g < 2; g++) {
      const y0 = nodeY(g, ch.v), n = 31;
      for (let j = 1; j <= n; j++) out.push({ x, y: y0 + (LAY.gap * j) / n, a: 0, s: 1.1, m: ci + 1, ex: j / n, ey: 0 });
    }
  });
  return out;
}
const redNode = () => ({ x: nodeX(RED_NODE.u, RED_NODE.v), y: nodeY(0, RED_NODE.v) });

// ---- F4–F6: the figures as particle numerals
const STATS = [
  { num: '1,7 %', l1: 'DRIFTSMARGIN MERKEFORHANDLERE 2024', l2: 'FRA 1,5 % I 2023', t0: 34.7, out: 37.2 },
  { num: '2,5 %', l1: 'DRIFTSMARGIN BRUKTBILFORHANDLERE 2024', l2: 'ANALYSERT AV BDO', t0: 37.5, out: 39.1 },
  { num: '1 300+', l1: 'BILBRANSJEBEDRIFTER ANALYSERT', l2: 'BDO BILBRANSJEANALYSEN 2025', t0: 39.4, out: 40.4 },
];
const NUM = { x: 112, base: 585, size: 440 };
function buildNumeral(str, seed) {
  const cv = document.createElement('canvas'); cv.width = W; cv.height = H;
  const c = cv.getContext('2d', { willReadFrequently: true });
  let size = NUM.size; c.font = M.font(size, 800, K.DISPLAY);
  const w = c.measureText(str).width; if (w > 1500) { size *= 1500 / w; c.font = M.font(size, 800, K.DISPLAY); }
  c.fillStyle = '#fff'; c.textBaseline = 'alphabetic'; c.fillText(str, NUM.x, NUM.base);
  const d = c.getImageData(0, 0, W, H).data;
  const pts = pickPixels(d, W, H, (R, G, B, A) => A > 128, 5400, rnd(seed));
  for (const p of pts) { p.s = 1.35; p.a = 1; }
  return pts;
}

// ---- F7: calm bands, a 12-month chart in four lanes
const LANE = { x0: 340, pitch: 165, base: [215, 380, 545, 710], bw: 62, step: 118, names: ['SALG', 'LAGER', 'VERKSTED', 'GARANTI'] };
const OUT = [{ i: 0, lane: 2, mo: 6, t0: 43.75, label: 'MND 07 · VERKSTED' }, { i: 1, lane: 0, mo: 2, t0: 45.0, label: 'MND 03 · SALG' }, { i: 2, lane: 3, mo: 9, t0: 46.25, label: 'MND 10 · GARANTI' }];
const barH = (l, m) => 58 + 24 * Math.sin(m * 0.55 + l * 1.7) + 9 * Math.sin(m * 1.3 + l * 0.6);
function buildBands(nAvail) {
  let g = 6.2, out;
  for (;;) {
    out = []; const r = rnd(51);
    for (let l = 0; l < 4; l++) for (let m = 0; m < 12; m++) {
      const cols = Math.floor(LANE.bw / g), rows = Math.round(barH(l, m) / g), bx = LANE.x0 + m * LANE.step + (LANE.bw - cols * g) / 2;
      for (let c = 0; c < cols; c++) for (let q = 0; q < rows; q++) out.push({ x: bx + (c + 0.5) * g + (r() - 0.5) * 0.8, y: LANE.base[l] - (q + 0.5) * g, a: 0.95, s: 1, m: 0, ex: 0, ey: 0 });
    }
    if (out.length <= nAvail) break; g += 0.15;
  }
  return out;
}
const outPos = (o) => ({ x: LANE.x0 + o.mo * LANE.step + LANE.bw / 2, y: LANE.base[o.lane] - 124 });

// ---- F8: the end ring
function buildRing(n) {
  const out = [], r = rnd(61);
  for (let i = 0; i < n; i++) { const th = r() * TAU, rad = (r() + r() + r() - 1.5) * 0.07; out.push({ x: 0, y: 0, a: 0.8, s: 0.9, m: th, ex: rad, ey: 0 }); }
  return out;
}
let END_LABEL = null;
const LOGO_Y = 335, TAG_Y = 545, MONO_Y = 625;

let READY = false;
async function init(ctx) {
  await K.loadLogo();
  ctx.save(); ctx.font = M.font(26, 500, K.MONO); ctx.letterSpacing = `${26 * 0.14}px`;
  const lw = ctx.measureText('TETT PÅ BILBRANSJEN').width; ctx.letterSpacing = '0px'; ctx.restore();
  END_LABEL = { x: 960 - lw / 2 - 30, w: lw };
  // F1
  let pts = buildCar();
  assign(1, pts, { 0: { x: carX(DOT.x), y: carY(DOT.y), a: 1, s: 1 } }, TR[1].ang);
  // F2
  pts = buildSite();
  const hero = { x: planX(0), y: planY(12) };
  assign(2, pts, { 0: { x: hero.x, y: hero.y, a: 1, s: 1 } }, TR[2].ang);
  // F3
  pts = buildLattice(); const rn = redNode();
  assign(3, pts, { 0: { x: rn.x, y: rn.y, a: 1, s: 1 } }, TR[3].ang);
  // F4–F6
  STATS.forEach((st, k) => {
    pts = buildNumeral(st.num, 70 + k);
    const j = nearest(pts, NUM.x + 380, NUM.base - 40); const red = pts.splice(j, 1)[0];
    assign(4 + k, pts, { 0: { x: red.x, y: red.y, a: 1, s: 1 } }, TR[4 + k].ang);
  });
  // F7: three outliers (particles 0,1,2) float above their lanes
  const forced = {}; OUT.forEach((o) => { const p = outPos(o); forced[o.i] = { x: p.x, y: p.y, a: 1, s: 1 }; });
  pts = buildBands(N - 3 - 360);
  assign(7, pts, forced, TR[7].ang);
  // F8: the ring; particle 0 rests beside the small label
  pts = buildRing(N - 1);
  F[8].ey[0] = 1;
  assign(8, pts, { 0: { x: END_LABEL.x, y: MONO_Y - 8, a: 1, s: 1, ey: 1 } }, TR[8].ang);
  READY = true;
}

// ------------------------------------------------------------------ camera
let cs = 1, sn = 0, czoom = 1, cpx = 0, cpy = 0;
function camSet(t) {
  const rot = 0.012 * Math.sin(0.13 * t + 1); cs = Math.cos(rot); sn = Math.sin(rot);
  czoom = 1 + 0.0011 * t; cpx = 16 * Math.sin(0.17 * t + 0.5); cpy = 10 * Math.cos(0.13 * t);
}
let CX = 0, CY = 0;
function cam(x, y, z) {
  const zz = czoom * (1 + (z - 0.5) * 0.025), dx = x - 960, dy = y - 540, par = 1 + (z - 0.5) * 0.7;
  CX = 960 + (cs * dx - sn * dy) * zz + cpx * par; CY = 540 + (sn * dx + cs * dy) * zz + cpy * par;
}
const camAt = (x, y) => { cam(x, y, 0.5); return [CX, CY]; };

// ------------------------------------------------------------------ background and the wipe
const WIPE = { t0: 24.0, t1: 25.6, ax: 0.94, ay: -0.34 };
const wipeE = (t) => lerp(-420, 1850, E.inOutCubic(prog(t, WIPE.t0, WIPE.t1)));
function background(ctx, t) {
  const e = wipeE(t), gx = -cpx * 0.5, gy = -cpy * 0.5;
  if (e > 1840) { K.sheet(ctx, 1); K.grid(ctx, 1, 1, gx, gy); return; }
  K.sheet(ctx, 0); K.grid(ctx, 0, 1, gx, gy);
  if (e <= -400) return;
  const nx = WIPE.ax, ny = WIPE.ay, dx = -ny, dy = nx, px = nx * e, py = ny * e, L = 3200;
  ctx.save(); ctx.beginPath();
  ctx.moveTo(px + dx * L, py + dy * L); ctx.lineTo(px - dx * L, py - dy * L);
  ctx.lineTo(px - dx * L - nx * L, py - dy * L - ny * L); ctx.lineTo(px + dx * L - nx * L, py + dy * L - ny * L); ctx.closePath(); ctx.clip();
  K.sheet(ctx, 1); K.grid(ctx, 1, 1, gx, gy);
  ctx.restore();
  ctx.save(); ctx.lineWidth = 2; ctx.strokeStyle = rgba(C.blue, 0.85); ctx.beginPath();
  ctx.moveTo(px + dx * L, py + dy * L); ctx.lineTo(px - dx * L, py - dy * L); ctx.stroke();
  ctx.lineWidth = 1; ctx.strokeStyle = rgba(C.light, 0.5); ctx.beginPath();
  ctx.moveTo(px + dx * L - nx * 16, py + dy * L - ny * 16); ctx.lineTo(px - dx * L - nx * 16, py - dy * L - ny * 16); ctx.stroke();
  ctx.restore();
}

// ------------------------------------------------------------------ the particles
const REG = [0, 0.5, 0.5, 0.8, 0.7, 0.7, 0.7, 1, 0.3], NB = 20;
const BI = Array.from({ length: NB }, () => new Int32Array(N)), BC = new Int32Array(NB), HOT = new Int32Array(N);
const SX = new Float32Array(N), SY = new Float32Array(N), SS = new Float32Array(N);
const DEPT_T = [0, 1, 2, 3, 4, 5].map((i) => 15.625 + 1.25 * i);
const SCAN = { t0: 8.6, t1: 11.4 };
const RGB = { dark: [C.light, C.cloud], bright: [C.light, C.blue] };
const REDP = [{ x: 0, y: 0, on: 1 }, { x: 0, y: 0, on: 0 }, { x: 0, y: 0, on: 0 }];
function particles(ctx, t) {
  camSet(t);
  BC.fill(0); let nHot = 0;
  const e = wipeE(t), allBright = e > 1840, noneBright = e <= -400;
  const fadeEnd = 1 - 0.55 * ss(53.5, 57, t);
  const scanX = lerp(260, 1660, E.inOutSine(prog(t, SCAN.t0, SCAN.t1)));
  const lastTs = TR[8].ts;
  for (let i = 0; i < N; i++) {
    let k = 0, tau = 0;
    for (let kk = 8; kk >= 1; kk--) { const tt = t - TR[kk].ts - DL[kk][i]; if (tt > 0) { k = kk; tau = tt; break; } }
    let x, y, a, s, cf = k;
    if (k === 0) { noiseAt(i, t); x = TX; y = TY; a = 1; s = 1; }
    else if (tau > 3.4) { tgt(k, i, t); x = TX; y = TY; a = alphaAt(k, i, t); s = F[k].s[i]; }
    else {
      const tr = TR[k], stf = lerp(tr.lo, tr.hi, H3[i]), w0 = Math.sqrt(stf), zeta = 0.68, wd = w0 * Math.sqrt(1 - zeta * zeta);
      const ee = 1 - Math.exp(-zeta * w0 * tau) * (Math.cos(wd * tau) + ((zeta * w0) / wd) * Math.sin(wd * tau));
      tgt(k - 1, i, t); const ax = TX, ay = TY; tgt(k, i, t);
      const bx = TX, by = TY, c = ee < 0 ? 0 : ee > 1 ? 1 : ee, ddx = bx - ax, ddy = by - ay, dd = Math.hypot(ddx, ddy) + 1e-3;
      const pp = SG[i] * AMP[i] * dd * 4 * c * (1 - c) / dd;
      x = ax + ddx * ee - ddy * pp; y = ay + ddy * ee + ddx * pp;
      a = lerp(alphaAt(k - 1, i, t), alphaAt(k, i, t), c); s = lerp(F[k - 1].s[i], F[k].s[i], c);
      if (tau < 0.6) cf = k - 1;
    }
    // idle breathing keeps the settled drawings alive
    if (k > 0) { x += 0.9 * Math.sin(t * 1.3 + PH1[i]); y += 0.9 * Math.cos(t * 1.1 + PH2[i]); }
    // highlights (hot = brighter, larger)
    let hot = 0;
    if (cf === 1 && t > SCAN.t0 - 0.4 && t < SCAN.t1 + 0.4) { if (Math.abs(x - scanX) < 42 && F[1].a[i] > 0.5) hot = 1; }
    else if (cf === 2) {
      const m = F[2].m[i];
      if (m > 0) { const d = t - DEPT_T[m - 1]; if (d > 0 && d < 1.1) hot = 1; else if (t > 22.4 && t < 23.3) hot = 1; }
    } else if (cf === 3) { const m = F[3].m[i]; if (m > 0) { const gr = linkGrow(t, m); if (gr > 0 && gr < 1 && Math.abs(gr * 1.15 - F[3].ex[i]) < 0.14) hot = 1; } }
    if (cf === 8) a *= fadeEnd;
    let reg = 0, zi = Z[i], sbi = SB[i];
    { const rk = k === 0 ? 0 : (tau > 3.4 ? REG[k] : lerp(REG[k - 1], REG[k], 0.5)); if (k > 0 && tau <= 3.4) reg = lerp(REG[k - 1], REG[k], clamp(tau / 2)); else reg = rk; }
    if (reg > 0) { zi = lerp(zi, 0.78, reg); sbi = lerp(sbi, 2.3, reg); }
    cam(x, y, zi); x = CX; y = CY;
    if (i < 3) { REDP[i].x = x; REDP[i].y = y; if (i > 0) REDP[i].on = t >= OUT[i].t0 && t < 50.2 ? 1 : 0; if (i === 0) REDP[0].on = 1; if (REDP[i].on) { continue; } }
    // bucket: side x colour (far = light blue) x alpha level
    const side = allBright ? 1 : noneBright ? 0 : (x * WIPE.ax + y * WIPE.ay < e ? 1 : 0);
    const z = zi, al = (0.5 + 0.5 * z) * a * (hot ? 1.6 : 1);
    let q = Math.round(al * 5); if (q < 1) { if (al < 0.08) continue; q = 1; } if (q > 5) q = 5;
    const far = z < 0.22 ? 0 : 1, b = side * 10 + far * 5 + (q - 1);
    BI[b][BC[b]++] = i;
    SX[i] = x; SY[i] = y; SS[i] = sbi * (0.62 + 0.8 * z) * s * (hot ? 1.5 : 1) * (side ? 1.25 : 1);
  }
  for (let b = 0; b < NB; b++) {
    const n = BC[b]; if (!n) continue;
    const side = b >= 10 ? 1 : 0, far = (b % 10) >= 5 ? 1 : 0, q = (b % 5) + 1;
    const col = side ? (far ? C.blue : C.light) : (far ? C.cloud : C.light);
    ctx.fillStyle = rgba(col, side ? [0.5, 0.68, 0.82, 0.93, 1][q - 1] : [0.3, 0.5, 0.68, 0.85, 1][q - 1]);
    ctx.beginPath(); const arr = BI[b];
    for (let j = 0; j < n; j++) { const i = arr[j], sz = SS[i]; ctx.rect(SX[i] - sz * 0.5, SY[i] - sz * 0.5, sz, sz); }
    ctx.fill();
  }
}

// the three followed points: the red signature, drawn over the field
function redPoints(ctx, t) {
  const dark = t < 24.8;
  // particle 0: always red, with a soft ring on dark / bright
  const p0 = REDP[0], r0 = t > 49.5 ? 5.2 : 6.5;
  K.redDot(ctx, p0.x, p0.y, r0, 1);
  if (t < 4.6) { ctx.strokeStyle = rgba(C.red, 0.55 * (1 - ss(3.5, 4.6, t))); ctx.lineWidth = 1.2; ctx.beginPath(); ctx.arc(p0.x, p0.y, 20 + 3 * Math.sin(t * 3), 0, TAU); ctx.stroke(); }
  for (let j = 1; j < 3; j++) {
    const P = REDP[j]; if (!P.on) continue;
    ctx.fillStyle = C.red; ctx.beginPath(); ctx.arc(P.x, P.y, 5.4, 0, TAU); ctx.fill();
  }
  void dark;
}

// ------------------------------------------------------------------ overlays
const ink = (t) => (t < 24.8 ? C.white : C.graphite);
const subInk = (t) => (t < 24.8 ? C.cloud : C.blue);
function cap(ctx, str, t, t0, out, { y = K.CAP.y, size = 60, x = K.CAP.x } = {}) {
  return K.line(ctx, str, x, y, t, { size, weight: 600, color: ink(t), t0, out });
}
function tag(ctx, str, x, y, t, t0, out, { size = 24, color = null, align = 'center' } = {}) {
  const pi = E.outQuint(prog(t, t0, t0 + 0.5)), po = prog(t, out - 0.35, out); if (pi <= 0 || po >= 1) return;
  ctx.save(); ctx.globalAlpha *= pi * (1 - po);
  ctx.font = M.font(size, 500, K.MONO); ctx.letterSpacing = `${size * 0.1}px`; ctx.textAlign = align; ctx.fillStyle = color || subInk(t);
  ctx.fillText(str, x, y - (1 - pi) * 10); ctx.restore();
}

function overlayCar(ctx, t) {
  if (t < 6.4 || t > 13.2) return;
  const out = 1 - ss(12.4, 13.2, t), lineY = carY(-0.62);
  const EV = [[-1.9, '12.01.2024', 'INNKJØP'], [-0.95, '02.02.2024', 'KLARGJØRING'], [0.2, '14.03.2024', 'SALG'], [1.1, '20.09.2024', 'SERVICE'], [1.95, '03.11.2024', 'GARANTI']];
  ctx.save(); ctx.globalAlpha = out; ctx.lineWidth = 1;
  ctx.strokeStyle = rgba(C.light, 0.8);
  const p = E.inOutCubic(prog(t, 7.2, 9.4)), [x0, y0] = camAt(carX(-2.35), lineY), [x1, y1] = camAt(carX(2.35), lineY);
  K.drawOn(ctx, [[x0, y0], [x1, y1]], p);
  EV.forEach(([x, d, l], i) => {
    const t0 = 7.5 + i * 0.625, q = E.outCubic(prog(t, t0, t0 + 0.6)); if (q <= 0) return;
    const main = i === 2, [px, py] = camAt(carX(x), lineY);
    ctx.strokeStyle = main ? C.red : rgba(C.light, 0.9); ctx.lineWidth = main ? 2 : 1.2;
    ctx.beginPath(); ctx.moveTo(px, py - 10 * q); ctx.lineTo(px, py + 14 * q); ctx.stroke();
    ctx.save(); ctx.globalAlpha = out * q; ctx.textAlign = 'center'; ctx.letterSpacing = '2px';
    ctx.font = M.font(24, 500, K.MONO); ctx.fillStyle = main ? C.white : rgba(C.cloud, 0.8); ctx.fillText(d, px, py + 52 + (1 - q) * 8);
    ctx.font = M.font(26, 600, K.MONO); ctx.fillStyle = main ? C.white : rgba(C.light, 1); ctx.fillText(l, px, py + 88 + (1 - q) * 8); ctx.restore();
    if (main) { const [dx, dy] = [REDP[0].x, REDP[0].y]; ctx.strokeStyle = rgba(C.red, 0.9); ctx.lineWidth = 1.5; K.drawOn(ctx, [[dx, dy + 6], [px, py - 10]], q); }
  });
  ctx.restore();
}

function overlayPlan(ctx, t) {
  if (t < 15 || t > 23.7) return;
  DEPT_T.forEach((t0, i) => {
    const [x, y] = camAt(planX(-57 + 19 * (i + 0.5)), planY(-46));
    tag(ctx, ['SALG', 'LAGER', 'VERKSTED', 'DELER', 'GARANTI', 'SERVICE'][i], x, y - 34, t, t0 - 0.1, 23.4, { size: 24, color: C.white });
    const q = E.outCubic(prog(t, t0 - 0.1, t0 + 0.4)) * (1 - prog(t, 23.0, 23.4));
    if (q > 0) { ctx.fillStyle = C.red; ctx.fillRect(x - 1, y - 26, 2, 18 * q); }
  });
  // caption: the words arrive one at a time, then all six stand together
  const words = ['Salg.', 'Lager.', 'Verksted.', 'Deler.', 'Garanti.', 'Service.'];
  ctx.save(); ctx.font = M.font(60, 600, K.DISPLAY); let x = K.CAP.x;
  words.forEach((w, i) => { K.line(ctx, w, x, K.CAP.y, t, { size: 60, weight: 600, color: C.white, t0: DEPT_T[i] - 0.1, out: 23.7 }); x += ctx.measureText(w + ' ').width; });
  ctx.restore();
}

function overlayLattice(ctx, t) {
  if (t < 25 || t > 34) return;
  tag(ctx, 'BDO DATAPLATTFORM', 120, 108, t, 25.0, 33.6, { size: 26, color: C.blue, align: 'left' });
  // layer outlines draw on, hairlines
  const out = 1 - ss(32.8, 33.6, t);
  for (let l = 0; l < 3; l++) {
    const p = E.inOutCubic(prog(t, 26.0 + l * 0.35, 27.2 + l * 0.35)); if (p <= 0) continue;
    const corners = [[nodeX(0, 0), nodeY(l, 0)], [nodeX(1, 0), nodeY(l, 0)], [nodeX(1, 1), nodeY(l, 1)], [nodeX(0, 1), nodeY(l, 1)], [nodeX(0, 0), nodeY(l, 0)]].map(([x, y]) => camAt(x, y));
    ctx.save(); ctx.globalAlpha = out; ctx.strokeStyle = rgba(C.blue, 0.45); ctx.lineWidth = 1; K.drawOn(ctx, corners, p); ctx.restore();
    const [lx, ly] = camAt(nodeX(0, 1) - 10, nodeY(l, 1));
    ctx.save(); ctx.globalAlpha = out; K.kicker(ctx, LAYERS[l], Math.min(120, lx - 360), ly + 4, t, { color: C.graphite, t0: 26.4 + l * 0.6, size: 26 }); ctx.restore();
  }
  cap(ctx, 'Samlet. Sammenstilt. Analysert.', t, 25.0, 33.4);
}

function overlayNumbers(ctx, t) {
  if (t < 34.2 || t > 40.5) return;
  tag(ctx, 'BDO BILBRANSJEANALYSEN 2025', 120, 108, t, 34.3, 39.9, { size: 26, color: C.blue, align: 'left' });
  STATS.forEach((s, k) => {
    K.line(ctx, s.l1, 118, 716, t, { size: 30, weight: 500, family: K.MONO, color: C.blue, t0: s.t0 + 0.2, out: s.out, track: 0.08, dur: 0.6 });
    K.line(ctx, s.l2, 118, 766, t, { size: 28, weight: 400, family: K.MONO, color: C.graphite, t0: s.t0 + 0.45, out: s.out, track: 0.08, dur: 0.6 });
  });
  cap(ctx, 'Bransjekunnskap gir forståelse.', t, 35.0, 40.2);
}

function overlayBands(ctx, t) {
  if (t < 40.4 || t > 50.4) return;
  const out = 1 - ss(49.2, 50.2, t);
  tag(ctx, 'KLARHET · 12 MÅNEDER', 120, 108, t, 40.6, 49.4, { size: 26, color: C.blue, align: 'left' });
  ctx.save(); ctx.globalAlpha = out;
  for (let l = 0; l < 4; l++) {
    const p = E.inOutCubic(prog(t, 41.6 + l * 0.25, 43.2 + l * 0.25)), by = LANE.base[l] + 6;
    ctx.strokeStyle = rgba(C.blue, 0.35); ctx.lineWidth = 1;
    const a = camAt(LANE.x0 - 24, by), b = camAt(LANE.x0 + 11 * LANE.step + LANE.bw + 24, by); K.drawOn(ctx, [a, b], p);
    const [lx, ly] = camAt(LANE.x0 - 24, LANE.base[l] - 6);
    K.line(ctx, LANE.names[l], 120, ly, t, { size: 24, weight: 500, family: K.MONO, color: C.blue, t0: 41.8 + l * 0.2, out: 49.4, track: 0.1, dur: 0.5 });
  }
  for (let m = 0; m < 12; m++) {
    const [mx, my] = camAt(LANE.x0 + m * LANE.step + LANE.bw / 2, LANE.base[3] + 40);
    tag(ctx, String(m + 1).padStart(2, '0'), mx, my, t, 42.6 + m * 0.05, 49.4, { size: 22, color: rgba(C.graphite, 0.8) });
  }
  ctx.restore();
  // the three points that stand outside: they turn red, get a ring and a short label
  OUT.forEach((o, j) => {
    const q = prog(t, o.t0, o.t0 + 0.9); if (q <= 0) return;
    const P = REDP[j], so = 1 - ss(49.2, 50.2, t), r = 9 + 26 * E.outBack(q, 2.4);
    ctx.save(); ctx.globalAlpha = so;
    ctx.strokeStyle = C.red; ctx.lineWidth = 1.6; ctx.beginPath(); ctx.arc(P.x, P.y, r, 0, TAU); ctx.stroke();
    ctx.lineWidth = 1; ctx.beginPath(); ctx.moveTo(P.x + r, P.y); ctx.lineTo(P.x + r + 40 * E.outCubic(q), P.y); ctx.stroke();
    ctx.restore();
    K.line(ctx, 'TIL VURDERING', P.x + r + 52, P.y - 4, t, { size: 26, weight: 600, family: K.MONO, color: C.graphite, t0: o.t0 + 0.3, out: 49.4, track: 0.08, dur: 0.5 });
    K.line(ctx, o.label, P.x + r + 52, P.y + 26, t, { size: 22, weight: 400, family: K.MONO, color: C.blue, t0: o.t0 + 0.5, out: 49.4, track: 0.08, dur: 0.5 });
  });
  cap(ctx, 'Teknologien viser hvor vi bør se.', t, 40.625, 49.4, { y: 884, size: 54 });
  cap(ctx, 'Revisor vurderer.', t, 47.5, 49.4, { y: 948, size: 54 });
}

function overlayEnd(ctx, t) {
  if (t < 50.6) return;
  K.drawLogo(ctx, 960, LOGO_Y, 150, { wipe: E.inOutCubic(prog(t, 51.3, 52.5)) });
  K.line(ctx, 'Vi ser detaljene. Vi forstår helheten.', 960, TAG_Y, t, { size: 60, weight: 600, color: C.graphite, t0: 52.5, align: 'center' });
  ctx.save(); ctx.font = M.font(26, 500, K.MONO); ctx.letterSpacing = `${26 * 0.14}px`; ctx.textAlign = 'left';
  const q = E.outQuint(prog(t, 53.1, 53.7)); ctx.globalAlpha = q; ctx.fillStyle = C.blue;
  ctx.fillText('TETT PÅ BILBRANSJEN', 960 - END_LABEL.w / 2, MONO_Y + (1 - q) * 12); ctx.restore();
}

function overlayNoise(ctx, t) { if (t < 6) cap(ctx, 'Tusenvis av transaksjoner.', t, 0.3, 5.7); }
function overlayCarCaption(ctx, t) { if (t > 6 && t < 14) cap(ctx, 'Hver bil har en historikk.', t, 6.25, 13.4); }

// ------------------------------------------------------------------ film
const HITS = [
  [0, 'noise', 'blip', { pitch: 'D5' }], [0.6, 'caption', 'tick'], [3.75, 'gather', 'riser', { len: 2.5 }], [6.25, 'car', 'pop', { pitch: 'A4' }],
  [7.5, 'history', 'tick'], [8.75, 'sale', 'blip', { pitch: 'F#5' }], [12.5, 'disperse', 'whoosh', { len: 2.2, from: 400, to: 1800 }],
  ...DEPT_T.map((d, i) => [d, ['salg', 'lager', 'verksted', 'deler', 'garanti', 'service'][i], 'blip', { pitch: ['D5', 'E5', 'F#5', 'A5', 'B5', 'D6'][i] }]),
  [22.5, 'all six', 'sub', { len: 1.2 }], [24.0, 'wipe', 'whoosh', { len: 1.6, from: 300, to: 2400 }], [25.0, 'platform', 'impact'],
  [27.5, 'links', 'swell', { len: 2.5 }], [32.5, 'to figures', 'whoosh', { len: 2.2, from: 500, to: 1600 }], [35.0, '1,7', 'bell'],
  [37.5, '2,5', 'bell'], [39.4, '1300', 'bell'], [40.0, 'bands', 'sub', { len: 1.2 }], [43.75, 'point 1', 'blip', { pitch: 'A5' }],
  [45.0, 'point 2', 'blip', { pitch: 'B5' }], [46.25, 'point 3', 'blip', { pitch: 'D6' }], [47.5, 'revisor', 'tick'],
  [50.0, 'to card', 'impact'], [51.25, 'logo', 'bell'], [52.5, 'tagline', 'tick'],
].sort((a, b) => a[0] - b[0]).map(([s, ...rest]) => [s / PERIOD, ...rest]);

M.film({
  fonts: [M.font(100, 600, K.DISPLAY), M.font(100, 300, K.DISPLAY), M.font(40, 500, K.MONO)],
  hits: HITS,
  samples: (t) => (Math.abs(t - 5) < 2.6 || Math.abs(t - 13.5) < 2.6 || (t > 22.5 && t < 26.5) || (t > 32.5 && t < 35.5) || (t > 36.3 && t < 38) || (t > 38.4 && t < 40) || (t > 40.2 && t < 43) || (t > 48.4 && t < 51.5) ? 5 : 3),
  init,
  draw(ctx, u, t) {
    if (!READY) { K.sheet(ctx, 0); return; }
    if (DBGSITE) { ctx.drawImage(window.__site, 0, 0); return; }
    background(ctx, t);
    particles(ctx, t);
    overlayCar(ctx, t); overlayPlan(ctx, t); overlayLattice(ctx, t); overlayNumbers(ctx, t); overlayBands(ctx, t);
    redPoints(ctx, t);
    overlayNoise(ctx, t); overlayCarCaption(ctx, t); overlayEnd(ctx, t);
  },
  post(ctx, u, t) { K.grain(ctx, t, t < 24.8 ? 0 : 1); },
});
