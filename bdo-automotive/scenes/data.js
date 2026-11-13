// scenes/data.js (50-120 s): network -> BDO data platform -> analysis -> the eight selected points.
import * as K from '../lib/kit.js';

const { C, E, ss, clamp, lerp, prog, mix, rgba, W, H } = K;
const N = K.NPTS, P = K.PTS, TAU = Math.PI * 2;

export const HITS = [
  [50.3, 'dawn', 'swell'], [53.2, 'sources named', 'tick'], [55.2, 'streams leave', 'whoosh'],
  [58.4, 'planes draw', 'click'], [62.0, 'lattice fills', 'blip'], [65.6, 'platform set', 'pop'],
  [67.0, 'link wave', 'tick'], [70.0, 'red chain', 'bell'], [74.5, 'quote', 'swell'],
  [80.3, 'reform', 'whoosh'], [84.4, 'axes', 'click'], [88.2, 'bands', 'tick'],
  [92.5, 'first flag', 'impact'], [96.0, 'flag', 'pop'], [99.4, 'flag', 'blip'], [101.2, 'eighth flag', 'bell'],
  [104.4, 'rail', 'click'], [106.8, 'recede', 'whoosh'], [109.4, 'stage', 'tick'], [111.9, 'stage', 'tick'],
  [114.4, 'departure', 'whoosh'], [118.6, 'arrive', 'swell'],
];

// the dawn is a diagonal wipe (51.2-53.4): the bright sheet scans in from the right, no grey mid-tone
const WIPE = (t) => E.inOutCubic(prog(t, 51.2, 53.4));
let KO = null;   // forced sheet value while the wipe paints its two halves
export const sheetAt = (t) => (KO !== null ? KO : t <= 50 ? 0 : WIPE(t) > 0.5 ? 1 : 0);

// ---------------------------------------------------------------- layout (computed once)
const OUTI = [0, 211, 702, 1033, 1450, 1799, 2102, 2333];
const OUTSPEC = [
  { lane: 4, u: 0.40, dir: -1, tag: 'Garantikrav uten arbeidsordre' },
  { lane: 0, u: 0.95, dir: 1, tag: 'Kreditnota nær periodeslutt' },
  { lane: 1, u: 0.30, dir: 1, tag: 'Bruktbil · lang lagertid' },
  { lane: 3, u: 0.12, dir: 1, tag: 'Delelager · telledifferanse' },
  { lane: 0, u: 0.30, dir: -1, tag: 'Manuell rabatt over fullmakt' },
  { lane: 1, u: 0.05, dir: 1, tag: 'Innbytte · verdivurdering' },
  { lane: 2, u: 0.68, dir: -1, tag: 'Arbeidsordre uten faktura' },
  { lane: 0, u: 0.50, dir: 1, tag: 'Periodisering nybilsalg' },
];
const OUTX = new Int8Array(N).fill(-1);
OUTI.forEach((idx, j) => { OUTX[idx] = j; });

// platform planes
const TOPS = [205, 410, 615], PH = 140, PX = 600, PW = 980, SKEW = 150, COLS = 50, ROWS = 16;
const planePts = (L, pad = 0) => [[PX + SKEW - pad, TOPS[L] - pad], [PX + SKEW + PW + pad, TOPS[L] - pad], [PX + PW + pad, TOPS[L] + PH + pad], [PX - pad, TOPS[L] + PH + pad], [PX + SKEW - pad, TOPS[L] - pad]];
const slotXY = (L, slot) => {
  const col = slot >> 4, row = slot & 15, u = 0.03 + 0.94 * col / (COLS - 1), v = 0.1 + 0.8 * row / (ROWS - 1);
  return [PX + u * PW + (1 - v) * SKEW, TOPS[L] + v * PH];
};
const SLOT = new Int32Array(N);
for (let L = 0; L < 3; L++) {
  const list = []; for (let i = L; i < N; i += 3) list.push(i);
  list.sort((a, b) => P[a].x - P[b].x || a - b);
  for (let c = 0; c < COLS; c++) {
    const chunk = list.slice(c * ROWS, c * ROWS + ROWS).sort((a, b) => P[a].y - P[b].y || a - b);
    chunk.forEach((idx, r) => { SLOT[idx] = c * ROWS + r; });
  }
}
const S0 = 22 * 16 + 7;
for (let L = 0; L < 3; L++) {
  let h = -1; for (let i = L; i < N; i += 3) if (SLOT[i] === S0) { h = i; break; }
  if (h >= 0 && h !== L) { SLOT[h] = SLOT[L]; SLOT[L] = S0; }
}

// scatter
const LANEY = [235, 343, 451, 559, 667], X0 = 330, XW = 1440, BW = 22, AXY = 724;
const LANES = ['Nybil', 'Bruktbil', 'Verksted', 'Deler', 'Garanti'];
const MONTHS = ['JAN', 'FEB', 'MAR', 'APR', 'MAI', 'JUN', 'JUL', 'AUG', 'SEP', 'OKT', 'NOV', 'DES'];
const trend = (k, u) => 9 * Math.sin(TAU * (u * (0.9 + 0.12 * k) + 0.17 * k));

const hx = new Float32Array(N), hy = new Float32Array(N), lx = new Float32Array(N), ly = new Float32Array(N), sx = new Float32Array(N), sy = new Float32Array(N);
const d1 = new Float32Array(N), D1 = new Float32Array(N), d2 = new Float32Array(N), D2 = new Float32Array(N), fa = new Float32Array(N);
const c1x = new Float32Array(N), c1y = new Float32Array(N), c2x = new Float32Array(N), c2y = new Float32Array(N);
const lcol = new Float32Array(N), lrow = new Float32Array(N);
const ctrl = (ax, ay, bx, by, curl) => [(ax + bx) / 2 - (by - ay) * curl, (ay + by) / 2 + (bx - ax) * curl];
for (let i = 0; i < N; i++) {
  const p = P[i], h = (s) => K.hash01(i, s), L = i % 3;
  hx[i] = p.x; hy[i] = p.y;
  [lx[i], ly[i]] = slotXY(L, SLOT[i]); lcol[i] = SLOT[i] >> 4; lrow[i] = SLOT[i] & 15;
  let lane = h(11) < 0.18 ? 0 : h(11) < 0.38 ? 1 : h(11) < 0.68 ? 2 : h(11) < 0.88 ? 3 : 4, u = h(12);
  let g = clamp((h(13) + h(14) + h(15) - 1.5) * 16, -21, 21);
  if (OUTX[i] >= 0) { const o = OUTSPEC[OUTX[i]]; lane = o.lane; u = o.u; g = o.dir * 46; }
  sx[i] = X0 + u * XW; sy[i] = LANEY[lane] - trend(lane, u) + g;
  // act 3 departure: red leads, others staggered cluster by cluster
  d1[i] = i === 0 ? 55.0 : 55.5 + 5.4 * (0.45 * (p.c / 6) + 0.55 * h(3));
  D1[i] = i === 0 ? 3.2 : 2.6 + 1.0 * h(4);
  const bx0 = hx[i] + Math.sin(d1[i] * 0.6 + p.seed * 20) * 4, by0 = hy[i] + Math.cos(d1[i] * 0.5 + p.seed * 17) * 4;
  [c1x[i], c1y[i]] = ctrl(bx0, by0, lx[i], ly[i], (p.c % 2 ? 1 : -1) * 0.28 + (h(6) - 0.5) * 0.3);
  d2[i] = 80.2 + 3.0 * ((sx[i] - X0) / XW) + 1.2 * h(7); D2[i] = 2.6;
  [c2x[i], c2y[i]] = ctrl(lx[i], ly[i], sx[i], sy[i], (i % 2 ? 1 : -1) * 0.22 + (h(8) - 0.5) * 0.25);
  fa[i] = 114.4 + 2.2 * h(9);
}
// travel of the eight to the hand-off column
const TS0 = (j) => 114.3 + j * 0.12, TE = (j) => 118.5 + j * 0.1;
const DEST = (j) => [300, 250 + j * 80];

// vertical chains between layers (same slot in each layer)
const CH = [];
for (let s = 0; s < COLS * ROWS; s++) {
  if (s === S0 || K.hash01(s, 5) > 0.2) continue;
  const [x, y0] = slotXY(0, s), [, y1] = slotXY(1, s), [, y2] = slotXY(2, s), col = s >> 4;
  CH.push({ x, y0, y1, y2, col, s0: 66.4 + (col / (COLS - 1)) * 8.0 + (s & 15) * 0.012 });
}

// bands
const NB = 49;
const BANDS = LANEY.map((cy, k) => { const c = [], tp = [], bt = []; for (let n = 0; n < NB; n++) { const u = n / (NB - 1), y = cy - trend(k, u), x = X0 + u * XW; c.push([x, y]); tp.push([x, y - BW]); bt.push([x, y + BW]); } return { c, tp, bt }; });

// ---------------------------------------------------------------- camera
function camAt(t) {
  const e = ss(50, 54, t) * (1 - ss(116.2, 118.4, t)), a = t - 50;
  const par = 1 - 0.88 * ss(58, 64, t) + 0.22 * ss(80, 84, t);
  return {
    e, par,
    px: e * 22 * (Math.sin(t * 0.23) * 0.7 + Math.sin(t * 0.51 + 1) * 0.3),
    py: e * 14 * (Math.sin(t * 0.19 + 2) * 0.7 + Math.sin(t * 0.43) * 0.3),
    s: 1 + e * (0.018 * Math.sin(a * 0.17) + 0.012 * (1 - Math.cos(a * 0.06))),
  };
}

// ---------------------------------------------------------------- per-point position (world space)
const bez2 = (a, c, b, u) => { const v = 1 - u; return v * v * a + 2 * v * u * c + u * u * b; };
const RS_LAT = 0.78, RS_SC = 0.62;
function raw(i, t, o) {
  const p = P[i], L = i % 3;
  let x, y, a = 1, rs = 1;
  const arr1 = d1[i] + D1[i], arr2 = d2[i] + D2[i];
  if (t <= d1[i]) {
    const e = ss(50, 52.5, t);
    x = hx[i] + Math.sin(t * 0.6 + p.seed * 20) * 4 * e; y = hy[i] + Math.cos(t * 0.5 + p.seed * 17) * 4 * e;
  } else if (t < arr1) {
    const u = E.inOutCubic(prog(t, d1[i], arr1));
    const bx0 = hx[i] + Math.sin(d1[i] * 0.6 + p.seed * 20) * 4, by0 = hy[i] + Math.cos(d1[i] * 0.5 + p.seed * 17) * 4;
    x = bez2(bx0, c1x[i], lx[i], u); y = bez2(by0, c1y[i], ly[i], u); rs = lerp(1, RS_LAT, u);
  } else if (t < d2[i]) {
    const w = ss(arr1, arr1 + 0.6, t) * (1 - ss(d2[i] - 0.5, d2[i], t));
    x = lx[i] + Math.sin(t * 0.9 + lcol[i] * 0.35 + L) * 1.0 * w; y = ly[i] + Math.cos(t * 1.1 + lrow[i] * 0.5 + L) * 1.2 * w; rs = RS_LAT;
  } else if (t < arr2) {
    const u = E.inOutCubic(prog(t, d2[i], arr2));
    x = bez2(lx[i], c2x[i], sx[i], u); y = bez2(ly[i], c2y[i], sy[i], u); rs = lerp(RS_LAT, RS_SC, u);
  } else {
    const w = ss(arr2, arr2 + 0.6, t);
    x = sx[i] + Math.sin(t * 0.7 + p.seed * 30) * 1.0 * w; y = sy[i] + Math.cos(t * 0.6 + p.seed * 25) * 1.0 * w; rs = RS_SC;
  }
  const j = OUTX[i];
  if (j >= 0) {
    const s0 = TS0(j);
    if (t > s0) {
      const e1 = TE(j), u = E.inOutCubic(prog(t, s0, e1));
      const ax = sx[i] + Math.sin(s0 * 0.7 + p.seed * 30), ay = sy[i] + Math.cos(s0 * 0.6 + p.seed * 25), [bx, by] = DEST(j);
      const [cx, cy] = ctrl(ax, ay, bx, by, (j % 2 ? 1 : -1) * 0.32);
      x = bez2(ax, cx, bx, u); y = bez2(ay, cy, by, u);
    }
  } else if (t >= 104) {
    const rc = ss(104, 106.5, t);
    a *= 1 - 0.92 * rc; const sh = 1 - 0.075 * rc * (1 - p.z);
    x = 960 + (x - 960) * sh; y = 500 + (y - 500) * sh; rs *= 1 - 0.15 * rc;
    if (t > fa[i]) { const dt = t - fa[i]; y += 520 * dt * dt; x += Math.sin(p.seed * 40) * 22 * dt; a *= 1 - ss(0, 0.9, dt); }
  }
  o[0] = x; o[1] = y; o[2] = a; o[3] = rs;
}
function proj(o, z, cam) {
  const f = (z * 1.6 - 0.4) * cam.par;
  o[0] = 960 + (o[0] - 960) * cam.s + cam.px * f; o[1] = 540 + (o[1] - 540) * cam.s + cam.py * f;
  o[3] *= 1 + cam.e * (z - 0.5) * 0.4;
}

const PS = Array.from({ length: N }, () => [0, 0, 1, 1]);
const TMP = [0, 0, 1, 1], TMP2 = [0, 0, 1, 1];

// ---------------------------------------------------------------- helpers
function part(ctx, pts, a, b) {
  if (b <= a) return;
  const n = pts.length - 1, A = a * n, B = b * n;
  const at = (s) => { const k = Math.min(Math.floor(s), n - 1), f = s - k; return [lerp(pts[k][0], pts[k + 1][0], f), lerp(pts[k][1], pts[k + 1][1], f)]; };
  const p0 = at(A); ctx.beginPath(); ctx.moveTo(p0[0], p0[1]);
  for (let k = Math.floor(A) + 1; k < B; k++) ctx.lineTo(pts[k][0], pts[k][1]);
  const p1 = at(B); ctx.lineTo(p1[0], p1[1]); ctx.stroke();
}
function rise(ctx, str, x, y, p, { size = 16, color = C.blue, align = 'left', track = 0.12, weight = 500 } = {}) {
  if (p <= 0.001) return;
  ctx.save();
  ctx.font = `${weight} ${size}px ${K.MONO}`; ctx.letterSpacing = `${size * track}px`; ctx.textAlign = align;
  const w = ctx.measureText(str).width, x0 = align === 'center' ? x - w / 2 : align === 'right' ? x - w : x;
  ctx.beginPath(); ctx.rect(x0 - 6, y - size * 1.1, w + 12, size * 1.5); ctx.clip();
  ctx.fillStyle = color; ctx.fillText(str, x, y + (1 - p) * size * 1.3);
  ctx.restore();
}
function tag(ctx, str, x, y, side, pa, pr) {
  if (pa <= pr) return;
  ctx.save();
  ctx.font = `500 17px ${K.MONO}`; ctx.letterSpacing = '1.4px'; ctx.textAlign = 'left';
  const w = ctx.measureText(str).width, x0 = side > 0 ? x + 26 : x - 26 - w;
  const va = side > 0 ? x0 + w * pr : x0 + w * (1 - pa), vb = side > 0 ? x0 + w * pa : x0 + w * (1 - pr);
  ctx.beginPath(); ctx.rect(va - 2, y - 22, vb - va + 4, 34); ctx.clip();
  ctx.lineJoin = 'round'; ctx.lineWidth = 5; ctx.strokeStyle = C.arctic; ctx.strokeText(str, x0, y + 5);
  ctx.fillStyle = C.graphite; ctx.fillText(str, x0, y + 5);
  ctx.restore();
  ctx.save(); ctx.strokeStyle = C.red; ctx.lineWidth = 1;
  const lp = clamp(pa * 3), lr = clamp(pr * 3);
  if (lp > lr) { ctx.beginPath(); ctx.moveTo(x + side * (16 + 8 * lr), y); ctx.lineTo(x + side * (16 + 8 * lp), y); ctx.stroke(); }
  ctx.restore();
}

// ---------------------------------------------------------------- chrome passes (world space, camera applied by caller)
function planes(ctx, t, pass) {
  for (let L = 0; L < 3; L++) {
    const pd = E.outCubic(prog(t, 57.0 + 0.4 * L, 59.4 + 0.4 * L)), r = E.inOutCubic(prog(t, 80.2 + 0.15 * L, 82.6 + 0.15 * L));
    if (pd <= 0 || r >= 1) continue;
    const pts = planePts(L, 14);
    if (pass === 0) {
      ctx.save(); ctx.beginPath(); ctx.rect(PX - 20 + (PW + SKEW + 40) * r, 0, (PW + SKEW + 40) * (pd - r) + 2, H); ctx.clip();
      ctx.fillStyle = rgba(C.cloud, 0.42); ctx.beginPath(); pts.forEach(([x, y], n) => (n ? ctx.lineTo(x, y) : ctx.moveTo(x, y))); ctx.fill();
      ctx.restore();
    } else {
      ctx.strokeStyle = rgba(C.blue, 0.85); ctx.lineWidth = 1.3; part(ctx, pts, r, pd);
      const e = [[pts[3][0], pts[3][1]], [pts[3][0], pts[3][1] + 7], [pts[2][0], pts[2][1] + 7], [pts[2][0], pts[2][1]]];
      ctx.lineWidth = 1; ctx.strokeStyle = rgba(C.blue, 0.5); part(ctx, e, r, pd);
      // leader from the layer label
      const ly0 = TOPS[L] + PH / 2, lp = E.outCubic(prog(t, 58.8 + 0.4 * L, 60.0 + 0.4 * L)) , lr = E.inOutCubic(prog(t, 80.2, 81.6));
      ctx.strokeStyle = rgba(C.blue, 0.45); part(ctx, [[430, ly0], [PX + SKEW / 2 - 16, ly0]], lr, lp);
    }
  }
}
function chains(ctx, t) {
  const rr = E.inOutCubic(prog(t, 79.3, 81.5));
  if (rr >= 1 || t < 66.4) return;
  const front = lerp(560, 1760, E.inOutSine(prog(t, 72.5, 78.5)));
  for (const c of CH) {
    const p = E.outCubic(prog(t, c.s0, c.s0 + 1.1));
    if (p <= 0) continue;
    const yA = lerp(c.y0, c.y2, rr), yB = lerp(c.y0, c.y2, p);
    if (yB <= yA) continue;
    const boost = Math.max(0, 1 - Math.abs(c.x - front) / 130) * (t > 72.5 && t < 79 ? 1 : 0);
    ctx.strokeStyle = rgba(C.blue, 0.34 + 0.5 * boost); ctx.lineWidth = 1;
    ctx.beginPath(); ctx.moveTo(c.x, yA); ctx.lineTo(c.x, yB); ctx.stroke();
    ctx.fillStyle = rgba(C.blue, 0.55 + 0.4 * boost);
    for (const y of [c.y0, c.y1, c.y2]) if (y <= yB + 0.5 && y >= yA - 0.5) { ctx.beginPath(); ctx.arc(c.x, y, 2.3, 0, TAU); ctx.fill(); }
  }
}
function scatterChrome(ctx, t, over) {
  const pa = E.outCubic(prog(t, 82, 85)), rr = E.inOutCubic(prog(t, 114, 116.8));
  if (pa <= 0 || rr >= 1) return;
  const ca = 1 - 0.72 * ss(104, 106.5, t);
  ctx.save(); ctx.globalAlpha *= ca;
  const ax = X0 + XW * rr - 2, bx = X0 + XW * pa + 2;
  if (!over) {
    ctx.save(); ctx.beginPath(); ctx.rect(ax, 0, bx - ax, H); ctx.clip();
    ctx.lineWidth = 1; ctx.strokeStyle = rgba(C.blue, 0.055); ctx.beginPath();
    for (let m = 0; m <= 12; m++) { const x = X0 + m * 120; ctx.moveTo(x, 181); ctx.lineTo(x, AXY); }
    ctx.stroke();
    ctx.strokeStyle = rgba(C.blue, 0.13); ctx.beginPath();
    for (let k = 0; k < 4; k++) { const y = LANEY[k] + 54; ctx.moveTo(X0, y); ctx.lineTo(X0 + XW, y); }
    ctx.stroke(); ctx.restore();
    for (let k = 0; k < 5; k++) {
      const pb = E.inOutCubic(prog(t, 86.3 + 0.5 * k, 89.8 + 0.5 * k)), r2 = E.inOutCubic(prog(t, 114, 116.8));
      if (pb <= 0) continue;
      const a0 = X0 + XW * r2, b0 = X0 + XW * pb;
      if (b0 <= a0) continue;
      const B = BANDS[k];
      ctx.save(); ctx.beginPath(); ctx.rect(a0, 0, b0 - a0, H); ctx.clip();
      ctx.fillStyle = rgba(C.cloud, 0.5); ctx.beginPath();
      B.tp.forEach(([x, y], n) => (n ? ctx.lineTo(x, y) : ctx.moveTo(x, y))); for (let n = NB - 1; n >= 0; n--) ctx.lineTo(B.bt[n][0], B.bt[n][1]);
      ctx.fill(); ctx.restore();
    }
  } else {
    ctx.save(); ctx.beginPath(); ctx.rect(ax, 0, bx - ax, H); ctx.clip();
    ctx.lineWidth = 1.2; ctx.strokeStyle = rgba(C.blue, 0.8);
    ctx.beginPath(); ctx.moveTo(X0, AXY); ctx.lineTo(X0 + XW, AXY); ctx.stroke();
    ctx.beginPath(); for (let m = 0; m <= 12; m++) { const x = X0 + m * 120; ctx.moveTo(x, AXY); ctx.lineTo(x, AXY + (m % 3 === 0 ? 12 : 7)); } ctx.stroke();
    ctx.restore();
    for (let k = 0; k < 5; k++) {
      const pb = E.inOutCubic(prog(t, 86.3 + 0.5 * k, 89.8 + 0.5 * k)), r2 = E.inOutCubic(prog(t, 114, 116.8));
      const a0 = X0 + XW * r2, b0 = X0 + XW * pb;
      if (pb <= 0 || b0 <= a0) continue;
      const B = BANDS[k];
      ctx.save(); ctx.beginPath(); ctx.rect(a0, 0, b0 - a0, H); ctx.clip();
      ctx.lineWidth = 1; ctx.strokeStyle = rgba(C.blue, 0.3);
      ctx.beginPath(); B.tp.forEach(([x, y], n) => (n ? ctx.lineTo(x, y) : ctx.moveTo(x, y))); ctx.stroke();
      ctx.beginPath(); B.bt.forEach(([x, y], n) => (n ? ctx.lineTo(x, y) : ctx.moveTo(x, y))); ctx.stroke();
      ctx.lineWidth = 1.5; ctx.strokeStyle = rgba(C.blue, 0.9);
      ctx.beginPath(); B.c.forEach(([x, y], n) => (n ? ctx.lineTo(x, y) : ctx.moveTo(x, y))); ctx.stroke();
      ctx.restore();
    }
    for (let k = 0; k < 5; k++) {
      const p = E.outQuint(prog(t, 82.3 + 0.18 * k, 83.1 + 0.18 * k)) * (1 - E.inCubic(prog(t, 114 + 0.1 * k, 114.8 + 0.1 * k)));
      rise(ctx, LANES[k], 120, LANEY[k] + 6, p, { size: 20, color: C.blue });
    }
    for (let m = 0; m < 12; m++) {
      const p = E.outQuint(prog(t, 82.6 + 0.07 * m, 83.4 + 0.07 * m)) * (1 - E.inCubic(prog(t, 114 + 0.05 * m, 114.8 + 0.05 * m)));
      rise(ctx, MONTHS[m], X0 + m * 120 + 60, AXY + 38, p, { size: 16, color: rgba(C.blue, 0.8), align: 'center' });
    }
  }
  ctx.restore();
}

// ---------------------------------------------------------------- the field
function drawField(ctx, t, k) {
  const col = mix(C.cloud, C.blue, ss(0.4, 0.6, k)), lcol_ = mix(C.light, C.blue, ss(0.4, 0.6, k));
  const links = 1 - ss(55.0, 56.8, t);
  if (links > 0.01) {
    ctx.lineWidth = 1; ctx.strokeStyle = rgba(lcol_, 0.22 * links); ctx.beginPath();
    for (const [a, b] of K.LINKS) { const A = PS[a], B = PS[b]; ctx.moveTo(A[0], A[1]); ctx.lineTo(B[0], B[1]); }
    ctx.stroke();
  }
  const skipOut = t >= 92;
  for (let i = 1; i < N; i++) {
    const p = P[i], q = PS[i];
    if (skipOut && OUTX[i] >= 0) continue;
    const a = q[2] * (0.35 + 0.65 * p.z);
    if (a <= 0.01) continue;
    ctx.fillStyle = rgba(col, a); ctx.beginPath(); ctx.arc(q[0], q[1], p.r * q[3], 0, 7); ctx.fill();
  }
  if (!skipOut) { const q = PS[0]; K.redDot(ctx, q[0], q[1], 7 * q[3], q[2]); }
}

// ---------------------------------------------------------------- frame
const SRC = [['DMS · salg', 0], ['Lagerstyring', 1], ['Verkstedsystem', 2], ['Delelager', 3], ['Garantiportal', 4], ['Servicebooking', 5], ['Kjøretøyregister', 6]];
const STOPS = [120, 560, 1000, 1440], RAILN = ['INNHENTING', 'SAMMENSTILLING', 'ANALYSE', 'REVISORS VURDERING'];

export function draw(ctx, t) {
  const w = WIPE(t);
  if (w > 0 && w < 1) {
    const ex = lerp(W + 320, -320, w), sl = 260;
    KO = 0; paint(ctx, t);
    ctx.save(); ctx.beginPath(); ctx.moveTo(ex + sl, 0); ctx.lineTo(W + 10, 0); ctx.lineTo(W + 10, H); ctx.lineTo(ex - sl, H); ctx.closePath(); ctx.clip();
    KO = 1; paint(ctx, t); ctx.restore(); KO = null;
    ctx.strokeStyle = C.blue; ctx.lineWidth = 2; ctx.beginPath(); ctx.moveTo(ex + sl, 0); ctx.lineTo(ex - sl, H); ctx.stroke();
    return;
  }
  paint(ctx, t);
}
function paint(ctx, t) {
  if (t >= 119.2) {
    K.sheet(ctx, 1);
    for (let i = 0; i < 8; i++) K.redDot(ctx, 300, 250 + i * 80, 8);
    return;
  }
  const k = sheetAt(t), cam = camAt(t);
  K.sheet(ctx, k);
  const ga = ss(50.4, 54, t) * (1 - ss(113.5, 117, t));
  if (ga > 0.01) K.grid(ctx, k, ga, cam.px * 0.15, cam.py * 0.15);

  const chrome = () => { ctx.translate(960 + cam.px * 0.4 * cam.par, 540 + cam.py * 0.4 * cam.par); ctx.scale(cam.s, cam.s); ctx.translate(-960, -540); };

  // positions
  for (let i = 0; i < N; i++) { raw(i, t, PS[i]); proj(PS[i], P[i].z, cam); }

  // under-chrome
  ctx.save(); chrome(); planes(ctx, t, 0); scatterChrome(ctx, t, false); ctx.restore();

  // stream trails
  if ((t > 55 && t < 67.5) || (t > 80.2 && t < 87.4)) {
    ctx.lineWidth = 1; ctx.strokeStyle = rgba(C.blue, 0.26); ctx.beginPath();
    for (let i = 3; i < N; i += 6) {
      const m1 = t > d1[i] && t < d1[i] + D1[i] + 0.1, m2 = t > d2[i] && t < d2[i] + D2[i] + 0.1;
      if (!m1 && !m2) continue;
      raw(i, t - 0.1, TMP); proj(TMP, P[i].z, cam);
      ctx.moveTo(TMP[0], TMP[1]); ctx.lineTo(PS[i][0], PS[i][1]);
    }
    ctx.stroke();
  }
  // red leader trail
  if (t > d1[0] && t < d1[0] + D1[0] + 0.8) {
    ctx.lineWidth = 1.5;
    let pv = null;
    for (let j = 0; j <= 14; j++) {
      raw(0, t - 0.9 * (1 - j / 14), TMP2); proj(TMP2, 0.5, cam);
      if (pv) { ctx.strokeStyle = rgba(C.red, 0.5 * (j / 14)); ctx.beginPath(); ctx.moveTo(pv[0], pv[1]); ctx.lineTo(TMP2[0], TMP2[1]); ctx.stroke(); }
      pv = [TMP2[0], TMP2[1]];
    }
  }

  drawField(ctx, t, k);

  // over-chrome
  ctx.save(); chrome(); planes(ctx, t, 1); chains(ctx, t); scatterChrome(ctx, t, true); ctx.restore();

  // red chain through the three layers
  if (t > 69.4 && t < 81.8) {
    const p = E.inOutCubic(prog(t, 69.6, 72.2)), rr = E.inOutCubic(prog(t, 79.6, 81.6));
    const a = PS[0], b = PS[2];
    const yA = lerp(a[1], b[1], rr), yB = lerp(a[1], b[1], p);
    if (yB > yA) {
      ctx.strokeStyle = C.red; ctx.lineWidth = 1.6; ctx.beginPath(); ctx.moveTo(lerp(a[0], b[0], rr), yA); ctx.lineTo(lerp(a[0], b[0], p), yB); ctx.stroke();
      for (const n of [1, 2]) {
        const q = PS[n], reached = (q[1] - a[1]) / (b[1] - a[1]) <= p + 0.001 && (q[1] - a[1]) / (b[1] - a[1]) >= rr - 0.001;
        if (!reached) continue;
        ctx.fillStyle = C.red; ctx.beginPath(); ctx.arc(q[0], q[1], 4.2, 0, TAU); ctx.fill();
        ctx.strokeStyle = C.red; ctx.lineWidth = 1.2; ctx.beginPath(); ctx.arc(q[0], q[1], 10, 0, TAU); ctx.stroke();
      }
      if (rr < 0.4) { ctx.lineWidth = 1.2; ctx.beginPath(); ctx.arc(a[0], a[1], 15, 0, TAU * clamp(p * 1.2)); ctx.stroke(); }
    }
  }

  // the eight
  if (t >= 92) {
    const pulseT = [104.6, 106.9, 109.3, 111.8];
    for (let j = 0; j < 8; j++) {
      const i = OUTI[j], q = PS[i], tr = 92.5 + 1.15 * j;
      const ramp = j === 0 ? 1 : ss(tr - 0.05, tr + 0.25, t);
      if (ramp < 1) { ctx.fillStyle = rgba(C.blue, (1 - ramp) * (0.35 + 0.65 * P[i].z) * q[2]); ctx.beginPath(); ctx.arc(q[0], q[1], P[i].r * q[3] + 0.6, 0, TAU); ctx.fill(); }
      let pul = 0; for (const s of pulseT) { const d = (t - (s + j * 0.07)) / 0.4; if (d > -1 && d < 1) pul += 0.35 * Math.cos(d * Math.PI / 2) ** 2; }
      const pop = j === 0 ? 0 : (t > tr ? 0.8 * Math.exp(-(t - tr) * 4.2) : 0);
      const grow = lerp(5.5, 8, E.inOutCubic(prog(t, TS0(j), TE(j))));
      K.redDot(ctx, q[0], q[1], grow * (1 + pop + pul), ramp);
      // ring + tag
      const rp = E.outCubic(prog(t, tr, tr + 0.8)), rrt = E.inCubic(prog(t, 113.3, 114.1));
      if (rp > rrt) {
        ctx.strokeStyle = C.red; ctx.lineWidth = 1.4; ctx.beginPath();
        ctx.arc(q[0], q[1], 14, -Math.PI / 2 + TAU * rrt, -Math.PI / 2 + TAU * rp); ctx.stroke();
      }
      const o = OUTSPEC[j], side = o.u > 0.78 ? -1 : 1;
      tag(ctx, o.tag, q[0], q[1], side, E.outCubic(prog(t, tr + 0.4, tr + 1.5)), E.inCubic(prog(t, 112.6, 113.6)));
    }
  }

  // ---------------- text
  const kk = sheetAt(t), ink = kk ? C.graphite : C.white, kc = kk ? C.blue : C.cloud;
  const X = K.CAP.x;
  K.line(ctx, 'DATA FRA FLERE SYSTEMER.', X, 930, t, { t0: 51, out: 56, color: ink });
  K.line(ctx, 'SAMMENSTILT I ÉN PLATTFORM.', X, 930, t, { t0: 58, out: 65, color: C.graphite });
  K.line(ctx, 'ANALYSERT PÅ TVERS AV VIRKSOMHETEN.', X, 930, t, { t0: 68, out: 74, color: C.graphite });
  K.line(ctx, 'Vi kobler sammen dataene.', X, 860, t, { t0: 74.5, out: 79.5, color: C.graphite, weight: 300 });
  K.line(ctx, 'Vi ser sammenhengene.', X, 930, t, { t0: 74.85, out: 79.5, color: C.graphite, weight: 300 });
  K.line(ctx, 'FRA DATA TIL FORSTÅELSE.', X, 930, t, { t0: 84, out: 90, color: C.graphite });
  K.line(ctx, 'Fra store datamengder', X, 860, t, { t0: 106, out: 113, color: C.graphite });
  K.line(ctx, 'til målrettet revisjon.', X, 930, t, { t0: 106.35, out: 113, color: C.graphite });

  if (t < 55.8) for (const [name, c] of SRC) {
    const cl = K.CLUSTERS[c];
    K.kicker(ctx, name, cl.x - 110, cl.y - 135, t, { color: kc, t0: 50.6 + 0.58 * c, out: 55.6 });
  }
  K.kicker(ctx, 'BDO DATAPLATTFORM', 120, 110, t, { color: C.blue, t0: 56.6, out: 80.4 });
  ['KJØRETØYDATA', 'TRANSAKSJONER', 'ØKONOMISKE DATA'].forEach((s, L) => K.kicker(ctx, s, 120, TOPS[L] + PH / 2 + 7, t, { color: C.blue, t0: 58.6 + 0.4 * L, out: 80.6 }));
  K.kicker(ctx, 'MØNSTRE OG SAMMENHENGER', 120, 110, t, { color: C.blue, t0: 81, out: 92 });
  K.kicker(ctx, 'KREVER NÆRMERE UNDERSØKELSE', 120, 110, t, { color: C.blue, t0: 92.3, out: 104.2 });

  // stage rail
  if (t >= 104.2 && t < 117.2) {
    const pr = E.outCubic(prog(t, 104.3, 105.8)), so = E.inOutCubic(prog(t, 115.0, 117.0));
    ctx.save(); ctx.translate(0, -so * 230);
    ctx.strokeStyle = rgba(C.light, 0.9); ctx.lineWidth = 1.2;
    ctx.beginPath(); ctx.moveTo(120, 140); ctx.lineTo(120 + 1680 * pr, 140); ctx.stroke();
    const u = (t - 104.6) / 2.2, s = clamp(Math.floor(u), 0, 3), fr = clamp(u - s, 0, 1);
    const tx = u < 0 ? 120 : STOPS[s] + (s < 3 ? E.inOutCubic(prog(fr, 0.65, 1)) * (STOPS[s + 1] - STOPS[s]) : 0);
    const uu = u + 0.2;
    for (let n = 0; n < 4; n++) {
      const act = ss(n - 0.05, n + 0.05, uu) * (n < 3 ? 1 - ss(n + 0.95, n + 1.05, uu) : 1);
      const col = mix(C.light, C.graphite, act);
      if (pr * 1680 + 120 > STOPS[n]) { ctx.fillStyle = col; ctx.beginPath(); ctx.arc(STOPS[n], 140, 3.2, 0, TAU); ctx.fill(); }
      K.kicker(ctx, RAILN[n], STOPS[n], 112, t, { color: col, t0: 104.5 + 0.14 * n, tick: false, size: 20 });
    }
    if (t > 104.5) { ctx.fillStyle = C.red; ctx.fillRect(tx - 1.5, 131, 3, 18); }
    ctx.restore();
  }
}
