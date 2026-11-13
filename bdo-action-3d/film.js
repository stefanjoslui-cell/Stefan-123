// BDO Automotive – Action 2-D (v2). 60 s at 128 bpm: one huge flat word per shot, varied motivated transitions.
// A pure function of time: window.seek(t) paints frame t. No state between frames, no Math.random.
import * as M from './lib/motion.js';
import * as K from './lib/kit.js';
import * as world from './scenes/world.js';

const { W, H, C, clamp, lerp, prog, E, ss, rgba } = K;
const P = 0.46875;                       // one beat in seconds (128 bpm)
const END = 116;                         // end card starts (bar 30, beat 116)
const DISPLAY = K.DISPLAY, MONO = K.MONO;
const { EL, GROUND, DOT, carPlan, rooms, DEPTS, B: BLD } = world;
const toScreen = world.toScreen;
const NF = (s) => s.replace(/ /g, ' ');

// ------------------------------------------------------------------ looks (full-bleed fills: night, arctic, blue, red)
const LK = {
  N: { id: 'N', k: 0, bg: C.night, fg: C.white, line: C.cloud, dot: C.red },
  A: { id: 'A', k: 1, bg: C.arctic, fg: C.blue, line: C.graphite, dot: C.red },
  B: { id: 'B', k: 0, bg: C.blue, fg: C.white, line: C.white, dot: C.red },
  R: { id: 'R', k: 0, bg: C.red, fg: C.white, line: C.white, dot: C.white },
};
const BGC = {};
function bgCache(L) {
  if (BGC[L.id]) return BGC[L.id];
  const cv = document.createElement('canvas'); cv.width = W; cv.height = H; const c = cv.getContext('2d');
  c.fillStyle = L.bg; c.fillRect(0, 0, W, H);
  if (L.id === 'N') K.sheet(c, 0); else if (L.id === 'A') K.sheet(c, 1);
  return (BGC[L.id] = cv);
}
function paintBg(ctx, L, u) {
  ctx.fillStyle = L.bg; ctx.fillRect(-80, -80, W + 160, H + 160);
  if (L.id === 'N') K.grid(ctx, 0, 1, u * 9, u * 5);
  else if (L.id === 'A') K.grid(ctx, 1, 1, -u * 9, u * 5);
  else K.grid(ctx, 0, L.id === 'R' ? 1.4 : 1.2, u * 9, -u * 5);
}

// ------------------------------------------------------------------ small drawing helpers
const ring = (ctx, x, y, r, lw, col) => { ctx.lineWidth = lw; ctx.strokeStyle = col; ctx.beginPath(); ctx.arc(x, y, Math.max(0.1, r), 0, 7); ctx.stroke(); };
const disc = (ctx, x, y, r, col) => { ctx.fillStyle = col; ctx.beginPath(); ctx.arc(x, y, Math.max(0.1, r), 0, 7); ctx.fill(); };
const poly = (ctx, pts, close = false) => { ctx.beginPath(); pts.forEach(([x, y], i) => (i ? ctx.lineTo(x, y) : ctx.moveTo(x, y))); if (close) ctx.closePath(); };
const lw = (ctx, w) => { ctx.lineWidth = w; };

// car in side elevation, ground-line at gy, origin x at gx, scale s px/m. p = draw-on progress.
function elev(ctx, { gx, gy, s, col, w = 2.4, p = 1, fillCol = null, ground = true }) {
  const Q = ([x, y]) => [gx + x * s, gy - y * s];
  ctx.save(); ctx.lineJoin = 'round'; ctx.lineCap = 'round'; ctx.strokeStyle = col;
  const body = EL.body.map(Q);
  if (fillCol) { poly(ctx, body, true); ctx.fillStyle = fillCol; ctx.fill(); }
  const glass = EL.glass.map(Q); poly(ctx, glass, true); ctx.fillStyle = rgba(col, 0.1); ctx.fill();
  lw(ctx, w); K.drawOn(ctx, glass, clamp(p * 1.4), true); K.drawOn(ctx, body, p, true);
  lw(ctx, w * 0.6); ctx.strokeStyle = rgba(col, 0.7);
  if (p > 0.5) for (const l of EL.lines) K.drawOn(ctx, l.map(Q), clamp((p - 0.5) * 2.4));
  for (const wh of EL.wheels) {
    const wp = clamp((p - 0.3) * 2);
    ctx.strokeStyle = col; lw(ctx, w); K.drawOn(ctx, wh.rings[0].map(Q), wp, true);
    ctx.strokeStyle = rgba(col, 0.7); lw(ctx, w * 0.6); K.drawOn(ctx, wh.rings[1].map(Q), wp, true); K.drawOn(ctx, wh.rings[2].map(Q), wp, true);
    for (const sp of wh.spokes) K.drawOn(ctx, sp.map(Q), wp);
  }
  if (ground) { ctx.strokeStyle = rgba(col, 0.5); lw(ctx, 1.5); K.drawOn(ctx, [[gx - 4 * s, gy], [gx + 4 * s, gy]], p); }
  ctx.restore();
}

// chart: line that shoots up; returns the tip position
function chartPts(seed, n, steep) {
  const r = K.mulberry32(seed), a = [];
  for (let i = 0; i < n; i++) { const t = i / (n - 1); a.push([t, clamp(0.1 + Math.pow(t, steep) * 0.82 + (r() - 0.5) * 0.12 * (1 - 0.5 * t), 0, 1)]); }
  return a;
}
function chart(ctx, { x, y, w, h, p, col, lwid = 6, seed = 4, steep = 2.2, axes = true, tipCol = C.red, tipR = 13, n = 28 }) {
  const pts = chartPts(seed, n, steep).map(([t, v]) => [x + t * w, y + h - v * h]);
  if (axes) {
    ctx.strokeStyle = rgba(col, 0.28); lw(ctx, 1.5); ctx.beginPath();
    for (let i = 0; i <= 5; i++) { ctx.moveTo(x, y + h * i / 5); ctx.lineTo(x + w, y + h * i / 5); }
    ctx.stroke();
  }
  ctx.save(); ctx.lineJoin = 'round'; ctx.lineCap = 'round'; ctx.strokeStyle = col; lw(ctx, lwid);
  K.drawOn(ctx, pts, p); ctx.restore();
  const fi = (pts.length - 1) * clamp(p), k = Math.min(pts.length - 2, Math.floor(fi)), f = fi - k;
  const tip = [lerp(pts[k][0], pts[k + 1][0], f), lerp(pts[k][1], pts[k + 1][1], f)];
  if (p > 0.02) { disc(ctx, tip[0], tip[1], tipR, tipCol); }
  return tip;
}

// ledger rows (mono), scrolling; hl = index of the highlighted row
const LTXT = ['INNKJØP NYBIL', 'KLARGJØRING', 'SALG NYBIL', 'INNBYTTE', 'SERVICE 30 000 KM', 'GARANTI', 'DELER UTTAK', 'ARBEIDSORDRE', 'LEIE VERKSTED'];
const ledgerRow = (i) => { const h = K.hash01(i, 5); return `${118000 + i * 37}  ${String(1 + (i * 7) % 28).padStart(2, '0')}.${String(1 + (i * 5) % 12).padStart(2, '0')}.2024  ${LTXT[i % LTXT.length]}  ${NF(((h * 480 + 6) * 1000).toFixed(0).replace(/\B(?=(\d{3})+(?!\d))/g, ' '))},00`; };
function ledgerPage(ctx, { off, col, alpha, size = 40, pitch = 78, hl = -1, hlCol = C.red, x = 96 }) {
  ctx.font = M.font(size, 500, MONO); ctx.textAlign = 'left'; ctx.letterSpacing = '0px';
  const i0 = Math.floor(off / pitch) - 1;
  for (let i = i0; i < i0 + H / pitch + 3; i++) {
    const y = i * pitch - off + 60;
    ctx.fillStyle = rgba(col, alpha * (i === hl ? 3 : 1)); ctx.fillText(ledgerRow(((i % 60) + 60) % 60), x, y);
    if (i === hl) { ctx.fillStyle = hlCol; ctx.fillRect(x, y + 12, 1728, 4); }
  }
}


// batched data field (same look as K.field, one path per alpha bucket: ~10x cheaper)
const BUCKET = K.PTS.map((p) => Math.min(3, Math.floor(p.z * 4)));
const FCACHE = {};
function fieldCache(k) {
  const key = k < 0.5 ? 0 : 1;
  if (FCACHE[key]) return FCACHE[key];
  const mk = () => { const cv = document.createElement('canvas'); cv.width = W; cv.height = H; return cv; };
  const lk = mk(), pk = mk(), a = lk.getContext('2d'), b = pk.getContext('2d'), PT = K.PTS, col = key === 0 ? C.cloud : C.blue;
  a.lineWidth = 1; a.strokeStyle = rgba(key === 0 ? C.light : C.blue, 0.22); a.beginPath();
  for (const [i, j] of K.LINKS) { a.moveTo(PT[i].x, PT[i].y); a.lineTo(PT[j].x, PT[j].y); }
  a.stroke();
  for (let bk = 0; bk < 4; bk++) {
    b.fillStyle = rgba(col, 0.35 + 0.65 * (bk + 0.5) / 4); b.beginPath();
    for (let i = 1; i < K.NPTS; i++) { if (BUCKET[i] !== bk) continue; const r = PT[i].r; b.moveTo(PT[i].x + r, PT[i].y); b.arc(PT[i].x, PT[i].y, r, 0, 7); }
    b.fill();
  }
  return (FCACHE[key] = { lk, pk });
}
function FF(ctx, k, { alpha = 1, pos = null, links = 1, red = 1 } = {}) {
  const col = k < 0.5 ? C.cloud : C.blue, PT = K.PTS;
  if (!pos) {
    const fc = fieldCache(k);
    ctx.save(); if (links > 0) { ctx.globalAlpha *= clamp(alpha * links); ctx.drawImage(fc.lk, 0, 0); } ctx.globalAlpha = ctx.globalAlpha; ctx.restore();
    ctx.save(); ctx.globalAlpha *= clamp(alpha); ctx.drawImage(fc.pk, 0, 0); ctx.restore();
    K.redDot(ctx, PT[0].x, PT[0].y, 7, red);
    return;
  }
  const Q = new Array(PT.length), at = (i) => Q[i] || (Q[i] = pos(PT[i]));
  if (links > 0) {
    ctx.lineWidth = 1; ctx.strokeStyle = rgba(k < 0.5 ? C.light : C.blue, 0.22 * alpha * links); ctx.beginPath();
    for (const [a, b] of K.LINKS) { const A = at(a), B = at(b); ctx.moveTo(A[0], A[1]); ctx.lineTo(B[0], B[1]); }
    ctx.stroke();
  }
  for (let b = 0; b < 4; b++) {
    ctx.fillStyle = rgba(col, alpha * (0.35 + 0.65 * (b + 0.5) / 4)); ctx.beginPath();
    for (let i = 1; i < K.NPTS; i += 2) { if (BUCKET[i] !== b) continue; const q = at(i), r = PT[i].r * (q[3] ?? 1) * 1.3; if (q[0] < -20 || q[0] > W + 20 || q[1] < -20 || q[1] > H + 20) continue; if (r < 2.6) ctx.rect(q[0] - r, q[1] - r, 2 * r, 2 * r); else { ctx.moveTo(q[0] + r, q[1]); ctx.arc(q[0], q[1], r, 0, 7); } }
    ctx.fill();
  }
  const q = at(0); K.redDot(ctx, q[0], q[1], 7 * (q[3] ?? 1), red * (q[2] ?? 1));
}

// -- field helpers
const CL = K.CLUSTERS;
function netLines(ctx, p, col, wd) {
  const pairs = [[0, 1], [1, 2], [2, 3], [3, 4], [4, 5], [5, 6], [6, 0], [0, 5], [1, 5], [2, 4], [2, 5], [1, 6], [3, 5], [0, 2]];
  ctx.save(); ctx.lineJoin = 'round'; ctx.strokeStyle = col; lw(ctx, wd);
  pairs.forEach(([a, b], i) => { const q = E.outExpo(prog(p, i * 0.055, i * 0.055 + 0.16)); K.drawOn(ctx, [[CL[a].x, CL[a].y], [CL[b].x, CL[b].y]], q); });
  ctx.restore();
  CL.forEach((c, i) => { const q = prog(p, i * 0.05, i * 0.05 + 0.1); if (q > 0 && i !== 6) { disc(ctx, c.x, c.y, 15 * E.outBack(q, 3), col); } });
}
const aboutXf = (ctx, s, ax, ay, tx = W / 2, ty = H / 2) => { ctx.translate(tx, ty); ctx.scale(s, s); ctx.translate(-ax, -ay); };

// -- lattice (the platform): three layers that slam together
function lattice(ctx, { gap, col, a = 1, cx = W / 2, cy = H / 2, rx = 560, ry = 170 }) {
  for (let l = 0; l < 3; l++) {
    const y0 = cy + (l - 1) * gap;
    const top = [cx, y0 - ry], rt = [cx + rx, y0], bt = [cx, y0 + ry], lf = [cx - rx, y0];
    poly(ctx, [top, rt, bt, lf], true); ctx.fillStyle = rgba(col, 0.1 * a); ctx.fill();
    ctx.strokeStyle = rgba(col, 0.9 * a); lw(ctx, l === 1 ? 4 : 2.5); ctx.stroke();
    ctx.strokeStyle = rgba(col, 0.45 * a); lw(ctx, 1.5); ctx.beginPath();
    for (let i = 1; i < 10; i++) {
      const f = i / 10;
      ctx.moveTo(lerp(top[0], lf[0], f), lerp(top[1], lf[1], f)); ctx.lineTo(lerp(rt[0], bt[0], f), lerp(rt[1], bt[1], f));
      ctx.moveTo(lerp(top[0], rt[0], f), lerp(top[1], rt[1], f)); ctx.lineTo(lerp(lf[0], bt[0], f), lerp(lf[1], bt[1], f));
    }
    ctx.stroke();
  }
  ctx.strokeStyle = rgba(col, 0.8 * a); lw(ctx, 2); ctx.beginPath();
  for (const f of [0.25, 0.5, 0.75]) { const x = cx + (f - 0.5) * rx * 1.2, y = cy + (f - 0.5) * ry * 0.4; ctx.moveTo(x, y - gap); ctx.lineTo(x, y + gap); }
  ctx.stroke();
}

// ------------------------------------------------------------------ images: im(ctx, c) with c = {ls (s), lt (beats), p (0..1), n (beats), L}
const IM = {
  dot(ctx, c) {
    const q = E.outBack(prog(c.ls, 0, 0.16), 2.6), r = 150 * q, x = 1420, y = 540;
    const beat = (c.lt % 1);
    ring(ctx, x, y, r * (1.5 + 0.6 * beat), 2, rgba(C.cloud, 0.5 * (1 - beat)));
    ring(ctx, x, y, r * 2.4, 1.5, rgba(C.cloud, 0.25));
    disc(ctx, x, y, r, C.red);
  },
  ledger(ctx, c) {
    const L = c.L, q = E.outExpo(prog(c.ls, 0, 0.2));
    ledgerPage(ctx, { off: 20 + c.ls * 40, col: L.line, alpha: 0.13, size: 36, pitch: 70 });
    // THE line
    const str = 'SALG NYBIL 489 900,00', size = 118;
    ctx.save(); ctx.font = M.font(size, 600, MONO); ctx.textAlign = 'left'; const w = ctx.measureText(NF(str)).width, x = (W - w) / 2, y = 905;
    ctx.beginPath(); ctx.rect(0, y - size, W, size * 1.4); ctx.clip();
    ctx.fillStyle = L.line; ctx.fillText(NF(str), x - (1 - q) * 900, y);
    ctx.restore();
    ctx.fillStyle = C.red; ctx.fillRect(x, y + 20, w * q, 5);
    disc(ctx, x - 40, y - size * 0.34, 10 * q, C.red);
  },
  wheel(ctx, c) {
    const s = 820 + c.p * 90;
    elev(ctx, { gx: W / 2 - 1.75 * s, gy: H / 2 + 0.55 * s, s, col: rgba(c.L.line, 0.8), w: 3, p: 1, ground: false });
    disc(ctx, W / 2 - 0.3 * s, H / 2 + 0.205 * s, 14, C.red);
  },
  rearwheel(ctx, c) {
    const s = 560 + c.p * 70;
    elev(ctx, { gx: W / 2 + 1.85 * s, gy: H / 2 + 0.6 * s, s, col: rgba(c.L.line, 0.8), w: 3, p: 1 });
  },
  mirror(ctx, c) {
    const s = 1050 + c.p * 120;
    elev(ctx, { gx: W / 2 - 0.2 * s, gy: H / 2 + 1.2 * s, s, col: rgba(c.L.line, 0.85), w: 3, p: 1, ground: false });
    const x = W / 2, y = H / 2 + 0.015 * s, q = E.outBack(prog(c.ls, 0.15, 0.3), 3);
    disc(ctx, x, y, 20 * q, C.red); ring(ctx, x, y, 20 + 70 * (1 - E.outExpo(prog(c.ls, 0.15, 0.5))) + 26, 3, rgba(C.red, 0.9 * q));
  },
  car(ctx, c) {
    const s = 330, p = E.inOutCubic(prog(c.ls, 0, 0.9));
    elev(ctx, { gx: W / 2, gy: 930, s, col: rgba(c.L.line, 0.9), w: 3.4, p, fillCol: rgba(c.L.id === 'A' ? C.cloud : C.cloud, c.L.id === 'A' ? 0.5 : 0.1) });
    disc(ctx, W / 2 + DOT.x * s, 930 - DOT.y * s, 12 * E.outBack(prog(c.ls, 0.8, 0.95), 3), C.red);
  },
  whip(ctx, c) {
    const e = E.inOutCubic(prog(c.p, 0, 1)), cx = lerp(-50, 55, e), v = Math.abs(55 + 50) * 6 * Math.sin(Math.PI * c.p) * 0.012;
    ctx.save();
    for (const g of [-1, 0, 1]) { if (g && v < 0.25) continue; ctx.globalAlpha = g ? 0.45 : 1; (g ? world.siteFast : world.site)(ctx, { k: 0, cx: cx + g * v * 4, cy: 18, s: 42 }); }
    ctx.restore();
  },
  plan(ctx, c) { world.siteFast(ctx, { k: c.L.k, cx: 0, cy: 13, s: 12.5 + c.p * 1.6, alpha: c.L.id === 'B' ? 0.7 : 0.9 }); },
  planClose(ctx, c) { world.site(ctx, { k: c.L.k, cx: -6 + c.p * 12, cy: 12 + c.p * 3, s: 64, alpha: c.L.id === 'B' ? 0.8 : 1 }); },
  dept(d) {
    return (ctx, c) => {
      const x0 = BLD.x0 + BLD.w * d + 9.5, s = 44 + c.p * 5;
      const cam = { cx: x0 + (d % 2 ? -3 : 3), cy: -35, s };
      ctx.save(); ctx.globalAlpha = 1; rooms(ctx, cam, c.L.k, 1, (i) => (i === d ? 1 : 0.0)); ctx.letterSpacing = '0px'; ctx.restore();
    };
  },
  fieldAll(ctx, c) {
    const s = 1 + c.p * 0.06; ctx.save(); aboutXf(ctx, s, W / 2, H / 2); FF(ctx, c.L.k, { alpha: c.L.id === 'B' ? 0.7 : 0.85 }); ctx.restore();
  },
  fieldPulse(ctx, c) {
    const s = 1 + 0.08 * Math.abs(Math.sin(c.lt * Math.PI)) + c.p * 0.04; ctx.save(); aboutXf(ctx, s, W / 2, H / 2);
    FF(ctx, 0, { alpha: 0.75 }); ctx.restore();
  },
  net(ctx, c) {
    ctx.save(); FF(ctx, 0, { alpha: 0.4, links: 0.8 }); netLines(ctx, clamp(c.ls / 0.9), c.L.line, 3); ctx.restore();
  },
  explode(ctx, c) {
    const q = E.outExpo(prog(c.ls, 0, 0.45)), f = q * 1.0 + c.p * 0.45;
    const pos = (p) => { const dx = p.x - W / 2, dy = p.y - H / 2; const j = 1 + p.seed * 0.5 * (1 - q); return [W / 2 + dx * f * j, H / 2 + dy * f * j, 1, 1 + (1 - q) * 0.5]; };
    FF(ctx, c.L.k, { pos, alpha: 0.9, links: q });
  },
  fly(ctx, c) {
    const f = lerp(1.7, 3.3, E.outCubic(c.p)); const pos = (p) => { const dx = p.x - 700, dy = p.y - 520; return [W / 2 + dx * f, H / 2 + dy * f, 1, 2.2]; };
    FF(ctx, c.L.k, { pos, alpha: 0.85, links: 0.5 });
  },
  zoomCluster(ctx, c) {
    ctx.save(); aboutXf(ctx, 3.1 + c.p * 0.5, CL[0].x, CL[0].y); FF(ctx, c.L.k, { alpha: 0.85, red: 0 }); ctx.restore();
  },
  labels(ctx, c) {
    FF(ctx, 0, { alpha: 0.6, links: 0.6 });
    ctx.font = M.font(30, 500, MONO); ctx.letterSpacing = '4px'; ctx.textAlign = 'left';
    CL.forEach((cl, i) => { const q = E.outCubic(prog(c.ls, i * 0.05, i * 0.05 + 0.2)); if (q <= 0) return; ctx.fillStyle = rgba(c.L.line, q); disc(ctx, cl.x, cl.y, 8, c.L.line); ctx.fillText(cl.name.toUpperCase(), cl.x + 20, cl.y - 20); });
    ctx.letterSpacing = '0px';
  },
  streams(ctx, c, mode = 0) {
    const col = c.L.line;
    for (let i = 0; i < 26; i++) {
      const h = K.hash01(i, 8), h2 = K.hash01(i, 9), y = 40 + i * 41 + h * 12, sp = 1100 + h2 * 2400, len = 90 + h * 280;
      const x = (((h2 * 3000 + sp * c.ls) % (W + 800)) + (W + 800)) % (W + 800) - 400;
      ctx.strokeStyle = rgba(col, 0.55); lw(ctx, 2 + h * 2.5); ctx.beginPath(); ctx.moveTo(x - len, y); ctx.lineTo(x, y); ctx.stroke();
      disc(ctx, x, y, 4 + h * 3, col);
    }
  },
  converge(ctx, c) {
    const f = 1 - E.inOutCubic(prog(c.p, 0, 0.85)), tx = W / 2, ty = 300;
    const pos = (p) => [lerp(tx, p.x, f), lerp(ty, p.y, f), 1, 1 + (1 - f) * 0.4];
    FF(ctx, c.L.k, { pos, alpha: 0.8, links: f });
    const r = 22 * (1 - f) + 6 * f;
    if (f < 0.05) { disc(ctx, tx, ty, 24 + (c.lt - 0.7) * 0, C.red); }
  },
  lattice(ctx, c) {
    const q = E.outExpo(prog(c.ls, 0, 0.24)), gap = lerp(380, 118, q);
    lattice(ctx, { gap, col: c.L.line, a: 0.5 });
  },
  rows(ctx, c) { ledgerPage(ctx, { off: c.ls * 120, col: c.L.line, alpha: 0.22, size: 40, pitch: 78 }); },
  scan(ctx, c) {
    const q = E.outExpo(prog(c.ls, 0, 0.55));
    ledgerPage(ctx, { off: lerp(900, 0, q) + c.ls * 20, col: c.L.line, alpha: 0.28, size: 44, pitch: 82, hl: 8 + 0, x: 96 });
  },
  chartUp(ctx, c) {
    const p = E.inOutCubic(prog(c.ls, 0, 0.5));
    chart(ctx, { x: 96, y: 330, w: 1728, h: 650, p, col: rgba(c.L.line, 0.95), lwid: 7 });
  },
  chartZoom(ctx, c) {
    ctx.save(); aboutXf(ctx, 2.2 + c.p * 0.4, 1824, 400, 1300, 640);
    chart(ctx, { x: 96, y: 330, w: 1728, h: 650, p: 1, col: rgba(c.L.line, 0.95), lwid: 6, tipR: 16 }); ctx.restore();
  },
  chartSteep(ctx, c) {
    const p = E.outExpo(prog(c.ls, 0, 0.5));
    chart(ctx, { x: -30, y: 80, w: W + 100, h: 900, p, col: c.L.line, lwid: 12, steep: 3.2, tipCol: c.L.line, tipR: 22, axes: false, seed: 9 });
  },
  bars(ctx, c) {
    for (let i = 0; i < 9; i++) {
      const q = E.outBack(prog(c.ls, i * 0.04, i * 0.04 + 0.2), 1.6) * 0.7 + 0.3, hgt = (260 + i * 62 + K.hash01(i, 3) * 40) * q, w = 150, x = 96 + i * 190;
      ctx.fillStyle = rgba(C.cloud, 0.95); ctx.fillRect(x, 1000 - hgt, w, hgt); ctx.strokeStyle = C.graphite; lw(ctx, 2); ctx.strokeRect(x, 1000 - hgt, w, hgt);
      if (i === 8) disc(ctx, x + w / 2, 1000 - hgt - 6, 14 * q, C.red);
    }
  },
  links(ctx, c) {
    const L = c.L, base = c.mode === 1 ? 1.5 : 1;
    ctx.save(); aboutXf(ctx, base + c.p * 0.15, W / 2, H / 2); FF(ctx, L.k, { alpha: 0.35, links: 1.2, red: 0.8 });
    const r = K.mulberry32(c.mode === 1 ? 31 : 17);
    ctx.lineCap = 'round'; ctx.strokeStyle = L.line; lw(ctx, 3);
    for (let i = 0; i < 16; i++) {
      const a = K.PTS[1 + Math.floor(r() * 2399)], b = K.PTS[1 + Math.floor(r() * 2399)], q = E.outExpo(prog(c.ls, i * 0.035, i * 0.035 + 0.18));
      K.drawOn(ctx, [[a.x, a.y], [b.x, b.y]], q); if (q > 0.9) { disc(ctx, a.x, a.y, 7, L.line); disc(ctx, b.x, b.y, 7, L.line); }
    }
    ctx.restore();
  },
  cross(ctx, c) {
    const q = E.outExpo(prog(c.ls, 0, 0.2)); ctx.strokeStyle = rgba(c.L.line, 0.55); lw(ctx, 2); ctx.beginPath();
    ctx.moveTo(0, H / 2); ctx.lineTo(W * q, H / 2); ctx.moveTo(W / 2, 0); ctx.lineTo(W / 2, H * q); ctx.stroke();
    ring(ctx, W / 2, H / 2, 520 * (1.5 - 0.5 * q), 2, rgba(c.L.line, 0.35));
  },
  dial(ctx, c) {
    ctx.save(); ctx.translate(W / 2, H / 2); ctx.rotate(c.ls * 0.25); ctx.strokeStyle = rgba(c.L.line, 0.5);
    ring(ctx, 0, 0, 470, 2, rgba(c.L.line, 0.5)); ring(ctx, 0, 0, 380, 1.5, rgba(c.L.line, 0.3));
    ctx.beginPath(); for (let i = 0; i < 72; i++) { const a = i * Math.PI / 36, r0 = i % 6 ? 470 : 440; ctx.moveTo(Math.cos(a) * r0, Math.sin(a) * r0); ctx.lineTo(Math.cos(a) * 490, Math.sin(a) * 490); } lw(ctx, 2); ctx.stroke();
    ctx.restore();
  },
  herField(ctx, c) {          // the field with the red point shifted to the right; ring snaps around it
    const hold = c.mode === 1, s = 1 + c.p * 0.04;
    ctx.save(); ctx.translate(1350, 640); ctx.scale(s, s); ctx.translate(-560, -690);
    FF(ctx, c.L.k, { alpha: hold ? 0.55 : 0.7, links: 0.8, red: 0 }); ctx.restore();
    const t0 = hold ? -1 : 0.28, q = hold ? 1 : E.outExpo(prog(c.ls, t0, t0 + 0.12)), r = hold ? 38 : lerp(420, 38, q) - (hold ? 0 : 5 * Math.sin(prog(c.ls, t0 + 0.1, t0 + 0.3) * Math.PI));
    const x = 1350, y = 640;
    disc(ctx, x, y, 9, C.red);
    if (!hold && c.ls < t0) return;
    ring(ctx, x, y, r, hold ? 3 : 6, C.red);
    if (!hold) {
      const tq = E.outCubic(prog(c.ls, t0 + 0.12, t0 + 0.4));
      if (tq > 0) {
        ctx.strokeStyle = C.red; lw(ctx, 2); ctx.beginPath(); ctx.moveTo(x, y + 44); ctx.lineTo(x, y + 44 + 96 * tq); ctx.stroke();
        ctx.save(); ctx.beginPath(); ctx.rect(1000, 740, 860, 70); ctx.clip();
        ctx.font = M.font(36, 500, MONO); ctx.textAlign = 'right'; ctx.letterSpacing = '0px'; ctx.fillStyle = c.L.fg;
        ctx.fillText('Garantikrav uten arbeidsordre', 1824, 796 + (1 - tq) * 60); ctx.restore();
      }
    }
  },
  target(ctx, c) {
    const q = E.outExpo(prog(c.ls, 0, 0.3)), L = c.L;
    for (const r of [180, 330, 480]) ring(ctx, W / 2, H / 2, r * (0.6 + 0.4 * q) * (1 + c.p * 0.06), 2, rgba(L.line, 0.5));
    ctx.strokeStyle = rgba(L.line, 0.5); lw(ctx, 2); ctx.beginPath(); ctx.moveTo(W / 2 - 560, H / 2); ctx.lineTo(W / 2 + 560, H / 2); ctx.moveTo(W / 2, H / 2 - 560); ctx.lineTo(W / 2, H / 2 + 560); ctx.stroke();
    disc(ctx, W / 2 + 330 * 0.707, H / 2 - 330 * 0.707, 14 * q, C.red);
  },
  picks(ctx, c) {
    FF(ctx, 0, { alpha: 0.3, links: 0.5, red: 0 });
    const r = K.mulberry32(53); ctx.lineCap = 'round';
    for (let i = 0; i < 18; i++) {
      const p = K.PTS[1 + Math.floor(r() * 2399)], q = E.outBack(prog(c.ls, i * 0.025, i * 0.025 + 0.15), 2);
      ring(ctx, p.x, p.y, 26 * q, 3, c.L.line); disc(ctx, p.x, p.y, 7, c.L.line);
    }
  },
  siteFor(ctx, c) { world.site(ctx, { k: c.L.k, cx: -6 + c.p * 10, cy: -29, s: 24 + c.p * 5 }); },
  siteBack(ctx, c) {
    const e = E.inOutCubic(prog(c.ls, 0, 0.95 * 4 * P)), s = lerp(34, 10.8, e);
    (s <= 18 ? world.siteFast : world.site)(ctx, { k: 1, cx: 0, cy: lerp(10, 0, e), s });
    const [x, y] = toScreen(0, lerp(10, 0, e), s, 0, 12); disc(ctx, x, y, 9, C.red); ring(ctx, x, y, 16 + 10 * e, 2, rgba(C.red, 0.8));
  },
  numDots(ctx, c) {   // 1 300 points, one per company
    const v = c.num01 ?? 1, n = Math.floor(v * 1300), hot = c.L.id === 'N' ? C.pale : C.blue, dim = c.L.id === 'N' ? C.slate : C.blue;
    for (const [on, col] of [[true, rgba(hot, 0.85)], [false, rgba(dim, 0.35)]]) {
      ctx.fillStyle = col;
      for (let i = 0; i < 1300; i++) { if ((i < n) !== on) continue; const x = 114 + (i % 52) * 33.2, y = 150 + Math.floor(i / 52) * 33.2, r = on ? 4.2 : 2.4; ctx.fillRect(x - r, y - r, 2 * r, 2 * r); }
    }
  },
  base(ctx, c) {
    const q = E.outExpo(prog(c.ls, 0, 0.4)); ctx.strokeStyle = rgba(c.L.line, 0.6); lw(ctx, 2); K.drawOn(ctx, [[96, 880], [1824, 880]], q);
    for (let i = 0; i <= 12; i++) { ctx.beginPath(); ctx.moveTo(96 + i * 144, 880); ctx.lineTo(96 + i * 144, 868); ctx.stroke(); }
    ctx.fillStyle = C.red; ctx.fillRect(1824 - 6, 850, 6, 30 * q);
  },
};
// ------------------------------------------------------------------ new images (flat, thin lines, bold bars)
const KICK = 'BDO Bilbransjeanalysen 2025';
const BASEY = 960;
// bars: 1,5 / 1,7 / 2,5 drawn at true proportions (230 px per percentage point)
IM.bars3 = (mode) => (ctx, c) => {
  const L = c.L, k = 230, w = mode === 2 ? 150 : 200, x0 = 1290, gap = 30;
  const spec = mode === 2 ? [[1.5, 0], [1.7, 0], [2.5, 1]] : [[1.5, 0], [1.7, 1]];
  ctx.strokeStyle = rgba(L.line, 0.6); lw(ctx, 2); ctx.beginPath(); ctx.moveTo(1250, BASEY); ctx.lineTo(1840, BASEY); ctx.stroke();
  spec.forEach(([v, hot], i) => {
    const q = E.outBack(prog(c.ls, i * 0.08 + (mode === 2 && i < 2 ? -1 : 0.02), i * 0.08 + 0.3), 1.5), h = v * k * Math.max(0, q), x = x0 + i * (w + gap);
    ctx.fillStyle = hot ? L.fg : rgba(L.line, 0.3); ctx.fillRect(x, BASEY - h, w, h);
    if (hot) { disc(ctx, x + w / 2, BASEY - h - 26, 16 * Math.min(1, q * 1.2), C.red); }
  });
};
IM.sizes = (hi) => (ctx, c) => {
  const L = c.L, hs = [150, 270, 420], w = 360, gap = 36, x0 = (W - (3 * w + 2 * gap)) / 2, by = 905;
  ctx.strokeStyle = rgba(L.line, 0.6); lw(ctx, 2); ctx.beginPath(); ctx.moveTo(96, by); ctx.lineTo(1824, by); ctx.stroke();
  hs.forEach((h, i) => {
    const q = E.outBack(prog(c.ls, i * 0.03, i * 0.03 + 0.22), 1.6), x = x0 + i * (w + gap), hh = h * q;
    if (i === hi) { ctx.fillStyle = L.fg; ctx.fillRect(x, by - hh, w, hh); disc(ctx, x + w / 2, by - hh - 24, 14, C.red); }
    else { ctx.strokeStyle = rgba(L.line, 0.55); lw(ctx, 3); ctx.strokeRect(x, by - hh, w, hh); }
    // dealer dots inside the column: more dots in larger columns
    const n = [6, 16, 36][i]; ctx.fillStyle = i === hi ? rgba(L.bg, 0.55) : rgba(L.line, 0.35);
    for (let d = 0; d < n; d++) { const cx = x + 28 + (d % 8) * 44, cy = by - 28 - Math.floor(d / 8) * 44; if (cy > by - hh + 18) { ctx.beginPath(); ctx.arc(cx, cy, 8, 0, 7); ctx.fill(); } }
  });
};
IM.years = (ctx, c) => {
  const L = c.L, y = 800, q = E.outExpo(prog(c.ls, 0, 0.5));
  ctx.lineCap = 'round';
  for (const [yy, a, lab] of [[742, 0.45, 'OMSETNING'], [858, 0.9, 'DRIFTSMARGIN']]) {
    ctx.strokeStyle = rgba(L.line, a); lw(ctx, a > 0.6 ? 5 : 3); K.drawOn(ctx, [[96, yy], [96 + 1728, yy]], q);
    ctx.font = M.font(26, 500, MONO); ctx.letterSpacing = '4px'; ctx.fillStyle = rgba(L.line, 0.8); ctx.textAlign = 'left'; ctx.fillText(lab, 96, yy - 16); ctx.letterSpacing = '0px';
  }
  for (let i = 0; i < 10; i++) {
    const x = 96 + 1728 * (i + 0.5) / 10, tq = E.outBack(prog(c.ls, 0.08 + i * 0.035, 0.08 + i * 0.035 + 0.2), 2);
    ctx.strokeStyle = L.line; lw(ctx, 3); ctx.beginPath(); ctx.moveTo(x, 700); ctx.lineTo(x, 700 + 200 * tq); ctx.stroke();
  }
  const mx = 96 + 1728 * (0.05 + 0.9 * E.inOutCubic(prog(c.ls, 0.1, 1.2)));
  disc(ctx, mx, 800, 13, C.red);
};
IM.scatter = (ctx, c) => {   // scatter that settles into bands, one point ringed in red
  const L = c.L, e = E.outExpo(prog(c.ls, 0, 0.5)), r = K.mulberry32(21);
  ctx.fillStyle = rgba(L.line, 0.7);
  for (let i = 0; i < 280; i++) {
    const hx = r() * (W - 192) + 96, hy = 120 + r() * 840, band = 150 + Math.floor(r() * 6) * 150 + (r() - 0.5) * 36;
    ctx.beginPath(); ctx.arc(hx, lerp(hy, band, e), 3 + r() * 3, 0, 7); ctx.fill();
  }
  ctx.strokeStyle = rgba(L.line, 0.25); lw(ctx, 1.5); ctx.beginPath(); for (let b = 0; b < 6; b++) { ctx.moveTo(96, 150 + b * 150); ctx.lineTo(1824, 150 + b * 150); } ctx.stroke();
  const rq = E.outExpo(prog(c.ls, 0.25, 0.4)); if (rq > 0) { disc(ctx, 1430, 450, 10, C.red); ring(ctx, 1430, 450, lerp(300, 36, rq), 4, C.red); }
};
// scattered dots -> ordered rows ("tall blir forståelse")
IM.order = (a0, a1) => (ctx, c) => {
  const L = c.L, e = E.inOutCubic(lerp(a0, a1, c.p)), r = K.mulberry32(8), cols = 24, rows = 10;
  ctx.fillStyle = L.line; ctx.strokeStyle = rgba(L.line, 0.5); lw(ctx, 2);
  for (let i = 0; i < cols * rows; i++) {
    const gx = 160 + (i % cols) * 66, gy = 140 + Math.floor(i / cols) * 90, hx = r() * W, hy = r() * H, j = (1 - e) * 14;
    const x = lerp(hx, gx, e) + Math.sin(c.ls * 9 + i) * j, y = lerp(hy, gy, e) + Math.cos(c.ls * 8 + i * 2) * j;
    ctx.beginPath(); ctx.arc(x, y, 6, 0, 7); ctx.fill();
  }
  if (e > 0.6) { ctx.beginPath(); for (let rr = 0; rr < rows; rr++) { ctx.moveTo(160, 140 + rr * 90); ctx.lineTo(160 + (cols - 1) * 66, 140 + rr * 90); } ctx.stroke(); }
};
const ringAt = (x, y) => (ctx, c) => { ring(ctx, x, y, 40 + 360 * (1 - E.outExpo(prog(c.ls, 0, 0.3))), 5, C.red); disc(ctx, x, y, 14, C.red); };

// ------------------------------------------------------------------ the shot list
const SH = [];
let cur = 0;
function S(word, n, look, anim, im, x = {}) { SH.push({ word, n, b: cur, L: LK[look], anim, im, ...x }); cur += n; }
const T = (t, d, o = {}) => ({ tr: { t, d, ...o } });
const odo = { from: NF('1,5 %'), to: NF('1,7 %'), t0: 0.28 };

// COLD OPEN (bars 1-4)
S('ÉN.', 2, 'N', 'slam', IM.dot, { align: 'l', dx: 40, sfx: 'impact', shake: 14 });
S('TRANSAKSJON.', 3, 'A', 'letters', IM.ledger, { dy: -90, sfx: 'thud', hits: [[0.5, 'amount', 'tick']], ...T('iris', 0.9, { at: [1420, 540], r0: 230 }) });
S('ÉN.', 2, 'N', 'crop', IM.wheel, { sfx: 'thud', shake: 12, ...T('stopDot', 0.7, { to: [714, 708], r: 14 }) });
S('BIL.', 3, 'A', 'U', IM.car, { dy: -190, sfx: 'thud', shake: 8, hits: [[1.8, 'car drawn', 'pop', { pitch: 'A4' }]], ...T('pushIris', 1.1, { at: [671, 738], r0: 333, smax: 8 }) });
S('TUSEN.', 2, 'N', 'R', IM.whip, { sfx: 'thud', ...T('whip', 0.6, { dir: [-1, 0] }) });
S('BILER.', 2, 'A', 'snap', IM.plan, { sfx: 'thud', shake: 10, ...T('diag', 0.6) });
S('BILER.', 2, 'N', 'echo', IM.planClose, { sfx: 'click', shake: 6 });
// THE BUSINESS (bars 5-8)
['SALG.', 'LAGER.', 'VERKSTED.', 'DELER.', 'GARANTI.', 'SERVICE.'].forEach((w, i) =>
  S(w, i === 5 ? 3 : 1, ['A', 'N', 'B', 'A', 'N', 'B'][i], ['slam', 'letters', 'mask', 'letters', 'snap', 'letters'][i], IM.dept(i), { sfx: 'click', shake: 7, ...(i === 0 ? {} : i % 2 ? T('slide', 0.5) : T('whip', 0.4, { dir: i === 2 ? [1, 0] : [-1, 0] })) }));
S('ALT.', 2, 'A', 'crop', IM.fieldAll, { sfx: 'thud', shake: 12, ...T('whip', 0.5, { dir: [0, -1] }) });
S('HENGER.', 2, 'B', 'R', IM.fieldPulse, { sfx: 'thud', shake: 8, ...T('lineWipe', 0.6) });
S('SAMMEN.', 4, 'N', 'cut', IM.net, { sfx: 'thud', shake: 10, ...T('diag', 0.6), hits: [0.15, 0.3, 0.45, 0.6, 0.8, 1.0, 1.2].map((h, i) => [h, 'link', 'tick', { pitch: 1200 + i * 110 }]) });
// TECHNOLOGY: DATA (bars 9-12)
S('MILLIONER.', 2, 'N', 'slam', IM.explode, { sfx: 'impact', shake: 20, ...T('stopDot', 0.6, { to: [W / 2, H / 2], r: 12 }) });
S('AV.', 1, 'B', 'slam', IM.fly, { sfx: 'thud', shake: 10, ...T('whip', 0.35, { dir: [-1, 0] }) });
S('DATAPUNKTER.', 2, 'N', 'letters', IM.labels, { sfx: 'click', shake: 6, ...T('slide', 0.5) });
S('HENTET.', 1, 'A', 'L', IM.streams, { sfx: 'click', shake: 6, ...T('whip', 0.35, { dir: [1, 0] }) });
S('SAMMENSTILT.', 2, 'B', 'snap', IM.converge, { sfx: 'thud', shake: 8, ...T('lineWipe', 0.5), pre: [-0.3, 'stream', 'whoosh', { len: 0.4, from: 800, to: 3000 }] });
S('ANALYSERT.', 2, 'N', 'mask', (ctx, c) => IM.links(ctx, c), { sfx: 'thud', shake: 8, ...T('diag', 0.5) });
S('ÉN.', 2, 'A', 'slam', IM.converge, { sfx: 'sub', shake: 8, dy: 160, align: 'l', dx: 60, pre: [-0.5, 'draw in', 'riser', { len: 0.45 }] });
S('PLATTFORM.', 4, 'R', 'crop', IM.lattice, { sfx: 'impact', shake: 28, hits: [[0.5, 'layer lock', 'sub', { len: 0.8 }]], kick: 'BDO Dataplattform', ...T('iris', 0.7, { at: [W / 2, 300], r0: 60 }) });
// BRANSJEKUNNSKAP (bars 13-20)
S('VI.', 1, 'A', 'slam', IM.rows, { sfx: 'thud', shake: 12, align: 'l' });
S('KJENNER.', 2, 'N', 'letters', IM.rearwheel, { sfx: 'thud', shake: 10, ...T('whip', 0.4, { dir: [-1, 0] }) });
S('BRANSJEN.', 3, 'A', 'echo', IM.planClose, { sfx: 'thud', shake: 10, ...T('diag', 0.6) });
S('', 3, 'B', 'title', IM.rows, { sfx: 'impact', shake: 14, kick: KICK, title: true, ...T('whip', 0.45, { dir: [0, 1] }) });
S(NF('1 300+'), 4, 'N', 'slam', IM.numDots, { fitStr: NF('1 300+'), numDots: true, nodot: true, label: 'Bilbransjebedrifter analysert', kick: KICK, sfx: 'thud', shake: 10, hits: [[0.8, 'count', 'tick']], dy: -40, ...T('iris', 0.8, { at: [960, 540], r0: 120 }) });
S(NF('1,7 %'), 4, 'A', 'slam', IM.bars3(1), { fitStr: NF('1,7 %'), align: 'l', max: 330, maxW: 1000, nodot: true, odo, label: 'Driftsmargin merkeforhandlere 2023 → 2024', labelX: 96, labelAlign: 'left', kick: KICK, kick2: 'Etterlengtet økning', sfx: 'thud', shake: 10, hits: [[0.8, 'roll', 'tick', { pitch: 1500 }]], dy: -70, ...T('lineWipe', 0.6) });
S(NF('2,5 %'), 3, 'N', 'slam', IM.bars3(2), { fitStr: NF('2,5 %'), align: 'l', max: 330, maxW: 1000, nodot: true, label: 'Driftsmargin bruktbilforhandlere 2024', labelX: 96, labelAlign: 'left', kick: KICK, sfx: 'thud', shake: 10, hits: [[0.8, 'bar', 'tick']], dy: -70, ...T('whip', 0.4, { dir: [-1, 0] }) });
['SMÅ.', 'MELLOMSTORE.', 'STORE.'].forEach((w, i) =>
  S(w, 1, ['B', 'A', 'N'][i], ['slam', 'letters', 'snap'][i], IM.sizes(i), { label: 'Forhandlere sammenlignet etter størrelse', labelY: 1010, kick: KICK, dy: -250, max: 300, sfx: 'click', shake: 8, ...(i === 0 ? T('diag', 0.5) : T('slide', 0.4)) }));
S(NF('10 ÅR.'), 3, 'A', 'crop', IM.years, { label: 'Omsetning og driftsmargin i et 10-årsperspektiv', kick: KICK, sfx: 'thud', shake: 10, dy: -190, hits: [[0.6, 'tick', 'tick', { pitch: 1300 }], [1.5, 'tick', 'tick', { pitch: 1600 }]], ...T('whip', 0.4, { dir: [1, 0] }) });
S('TALL.', 1, 'N', 'slam', IM.order(0, 0), { sfx: 'thud', shake: 10, ...T('whip', 0.35, { dir: [-1, 0] }) });
S('BLIR.', 1, 'A', 'cut', IM.order(0, 0.45), { sfx: 'click', shake: 8, ...T('slide', 0.4) });
S('FORSTÅELSE.', 4, 'R', 'crop', IM.order(0.45, 1), { sfx: 'impact', shake: 28, hits: [[1.2, 'ordered', 'sub', { len: 0.8 }]], ...T('pushIris', 0.9, { at: [0, 0], r0: 80, smax: 30, find: 'O' }) });
// EFFICIENT AUDIT (bars 21-24)
S('MØNSTRE.', 1, 'N', 'letters', IM.chartUp, { sfx: 'thud', shake: 8, ...T('diag', 0.4) });
S('SAMMENHENGER.', 2, 'A', 'mask', (ctx, c) => IM.links(ctx, c), { sfx: 'thud', shake: 8, ...T('whip', 0.4, { dir: [-1, 0] }) });
S('SE.', 1, 'B', 'slam', IM.cross, { sfx: 'thud', shake: 12, ...T('slide', 0.35) });
S('HER.', 2, 'N', 'slam', (ctx, c) => IM.herField(ctx, c), { align: 'l', dy: -90, sfx: 'thud', shake: 12, hits: [[0.6, 'ring snap', 'impact']], ...T('iris', 0.6, { at: [W / 2, H / 2], r0: 160 }) });
S('TEKNOLOGIEN.', 1, 'N', 'snap', IM.dial, { sfx: 'click', shake: 4, ...T('whip', 0.35, { dir: [0, -1] }) });
S('VISER.', 1, 'B', 'letters', IM.target, { sfx: 'click', shake: 6, ...T('pushIris', 0.5, { at: [0, 0], r0: 70, smax: 30, find: 'O' }) });
S('VEI.', 1, 'A', 'slam', IM.scatter, { sfx: 'click', shake: 6, ...T('lineWipe', 0.4) });
S('REVISOR.', 1, 'A', 'cut', (ctx, c) => IM.herField(ctx, { ...c, mode: 1 }), { sfx: 'click', shake: 4, hold: true, ...T('diag', 0.4) });
S('VURDERER.', 3, 'A', 'mask', (ctx, c) => IM.herField(ctx, { ...c, mode: 1 }), { sfx: 'sub', shake: 3, hold: true, ...T('slide', 0.5) });
S('MÅLRETTET.', 1, 'N', 'slam', IM.target, { sfx: 'thud', shake: 10, ...T('whip', 0.35, { dir: [-1, 0] }) });
S('EFFEKTIV.', 1, 'B', 'snap', IM.scan, { sfx: 'thud', shake: 8, ...T('slide', 0.35) });
S('REVISJON.', 1, 'A', 'crop', IM.picks, { sfx: 'thud', shake: 8, ...T('diag', 0.4) });
// INSIGHT FOR THE CUSTOMER (bars 25-28)
S('DINE.', 1, 'N', 'slam', IM.fieldPulse, { sfx: 'thud', shake: 8, pre: [-1, 'lift', 'riser', { len: 0.9 }] });
S('DATA.', 1, 'B', 'snap', IM.fieldAll, { sfx: 'click', shake: 8, ...T('slide', 0.35) });
S('DIN.', 1, 'A', 'L', IM.siteFor, { sfx: 'click', shake: 6, ...T('whip', 0.35, { dir: [-1, 0] }) });
S('INNSIKT.', 2, 'N', 'letters', IM.mirror, { sfx: 'thud', shake: 10, hits: [[0.3, 'insight', 'pop', { pitch: 'E6' }]], ...T('pushIris', 0.8, { at: [960, 540], r0: 40, smax: 60 }) });
S('OVERSIKT.', 2, 'A', 'mask', IM.siteBack, { label: 'Bedre oversikt over egen virksomhet', labelY: 1010, dy: -130, sfx: 'thud', shake: 8, ...T('diag', 0.5) });
S('DIALOG.', 2, 'N', 'cut', IM.picks, { label: 'Tydeligere dialog om funn', labelY: 1010, dy: -130, sfx: 'thud', shake: 8, ...T('whip', 0.4, { dir: [1, 0] }) });
S('UTVIKLING.', 2, 'A', 'echo', IM.chartUp, { label: 'Innsikt i utvikling og sammenhenger', labelY: 1010, dy: -130, sfx: 'thud', shake: 8, ...T('lineWipe', 0.5) });
S('MER.', 1, 'B', 'slam', IM.bars, { sfx: 'click', shake: 6, ...T('slide', 0.35) });
S('VERDI.', 2, 'R', 'crop', IM.chartSteep, { sfx: 'impact', shake: 28, hits: [[0.6, 'rise', 'sub', { len: 0.8 }]], ...T('iris', 0.6, { at: [W - 80, 140], r0: 60 }) });
S('FOR.', 1, 'N', 'slam', IM.siteFor, { sfx: 'click', shake: 6, ...T('stopDot', 0.4, { to: [W / 2, H / 2], r: 9 }) });
S('DEG.', 1, 'A', 'snap', IM.siteBack, { sfx: 'thud', shake: 12, ...T('diag', 0.4) });
// END stack
S('', 4, 'N', 'stack', IM.fieldAll, { sfx: 'thud', shake: 10, stack: true, ...T('whip', 0.45, { dir: [0, -1] }) });
if (cur !== END) throw new Error('shot list ends at beat ' + cur + ', expected ' + END);
const STARTS = SH.map((s) => s.b);
const idxAt = (u) => { let i = 0; for (let k = 0; k < SH.length; k++) if (STARTS[k] <= u) i = k; else break; return i; };

// ------------------------------------------------------------------ the word: kinetic type
function geom(ctx, sh) {
  if (sh._g) return sh._g;
  const str = sh.fitStr || sh.word;
  const size = sh.size || M.fitSize(ctx, str, 800, DISPLAY, sh.maxW || 1640, sh.max || 400, -0.03);
  const f = M.font(size, 800, DISPLAY), tr = -0.03 * size, lay = M.layout(ctx, str, f, tr), gl = lay.glyphs, last = gl[gl.length - 1];
  const x0 = (sh.align === 'l' ? 96 : sh.align === 'r' ? W - 96 - lay.width : (W - lay.width) / 2) + (sh.dx || 0);
  const base = H / 2 + 0.36 * size + (sh.dy || 0);
  const hasDot = !sh.nodot && last.ch === '.';
  const iO = gl.findIndex((g) => g.ch === 'O');
  return (sh._g = { str, size, f, tr, lay, width: lay.width, x0, base, hasDot,
    dot: hasDot ? { x: x0 + last.cx, y: base - size * 0.088, r: size * 0.088 } : null,
    counter: iO >= 0 ? { x: x0 + gl[iO].cx, y: base - 0.36 * size, r: size * 0.2 } : null });
}
const ODO = (sh) => { if (!sh.odo) return null; const a = [...sh.odo.from], b = [...sh.odo.to]; return a.map((ch, i) => (ch !== b[i] ? i : -1)).filter((i) => i >= 0); };
function glyphLoop(ctx, g, sh, c, anim, tw = {}) {
  const L = sh.L, ls = c.ls, gl = g.lay.glyphs, pivX = g.x0 + g.width / 2, odi = ODO(sh);
  gl.forEach((gp, i) => {
    let gx = g.x0 + gp.cx, gy = g.base, gs = 1, rot = 0, clip = false;
    if (anim === 'letters') { const q = E.outBack(prog(ls - 0.022 * i, 0, 0.2), 2.8); gy += (1 - q) * g.size * 0.75; gs = clamp(0.5 + 0.5 * q, 0.01, 1.4); rot = (1 - q) * (i % 2 ? 0.4 : -0.4); }
    else if (anim === 'snap') { const q = E.outBack(prog(ls, 0, 0.2), 2.2); gx = pivX + (gx - pivX) * (1 + (1 - q) * 0.6); gs = 1 + 0.16 * (1 - q); }
    else if (anim === 'mask') { const q = E.outQuint(prog(ls - 0.025 * i, 0, 0.3)); gy += (1 - q) * g.size * 1.15; clip = true; }
    if (clip) { ctx.save(); ctx.beginPath(); ctx.rect(gx - g.size, g.base - g.size * 0.95, g.size * 2, g.size * 1.25); ctx.clip(); }
    if (g.hasDot && i === gl.length - 1 && gp.ch === '.') { const r = g.size * 0.088; disc(ctx, gx, gy - r * gs, r * gs, L.dot); }
    else if (odi && odi.includes(i)) {
      const q = M.spring(ls - sh.odo.t0, { stiffness: 260, damping: 17 }), sz = g.size * 0.95;
      ctx.save(); ctx.beginPath(); ctx.rect(gx - g.size * 0.6, g.base - g.size * 0.85, g.size * 1.2, g.size * 1.0); ctx.clip();
      M.glyph(ctx, [...sh.odo.from][i], gx, g.base - q * sz, g.f, L.fg, 1, 1); M.glyph(ctx, [...sh.odo.to][i], gx, g.base + (1 - q) * sz, g.f, L.fg, 1, 1);
      ctx.restore();
    } else M.glyph(ctx, gp.ch, gx, gy, g.f, L.fg, gs, gs, rot);
    if (clip) ctx.restore();
  });
}
function drawWord(ctx, sh, c) {
  if (!sh.word) return;
  const g = geom(ctx, sh), L = sh.L, ls = c.ls, anim = sh.anim, pivX = g.x0 + g.width / 2, pivY = g.base - 0.36 * g.size;
  let sc = 1, ox = 0, oy = 0;
  if (anim === 'slam') { const q = E.outExpo(prog(ls, 0, 0.07)); sc = 1.34 - 0.34 * q - 0.05 * Math.sin(Math.PI * prog(ls, 0.05, 0.2)); }
  else if (anim === 'crop') { const q = E.outExpo(prog(ls, 0, 0.26)); sc = 1 + 0.75 * (1 - q) - 0.035 * Math.sin(Math.PI * prog(ls, 0.18, 0.34)); ox = 40 * (1 - q); }
  else if (anim === 'L' || anim === 'R') ox = (anim === 'L' ? -1 : 1) * (1 - M.spring(ls, { stiffness: 520, damping: 22 })) * 800;
  else if (anim === 'U' || anim === 'D') oy = (anim === 'U' ? 1 : -1) * (1 - M.spring(ls, { stiffness: 520, damping: 22 })) * 500;
  ctx.save(); ctx.translate(pivX + ox, pivY + oy); ctx.scale(sc, sc); ctx.translate(-pivX, -pivY);
  if (anim === 'cut') {
    const q = E.outExpo(prog(ls, 0.03, 0.3)), off = (1 - q) * 700, mid = pivY, lq = E.outExpo(prog(ls, 0, 0.08)), lf = 1 - prog(ls, 0.3, 0.45);
    for (const s of [-1, 1]) { ctx.save(); ctx.beginPath(); if (s < 0) ctx.rect(-200, -200, W + 400, mid + 200); else ctx.rect(-200, mid, W + 400, H); ctx.clip(); ctx.translate(s * off, 0); glyphLoop(ctx, g, sh, c, 'slam'); ctx.restore(); }
    if (lf > 0) { ctx.strokeStyle = L.fg; lw(ctx, 4); ctx.beginPath(); ctx.moveTo(0, mid); ctx.lineTo(W * lq, mid); ctx.stroke(); }
  } else if (anim === 'echo') {
    const sep = 1 - E.outExpo(prog(ls, 0, 0.35)), str = g.hasDot ? g.str.slice(0, -1) : g.str;
    ctx.font = g.f; ctx.letterSpacing = `${g.tr}px`; ctx.textAlign = 'left'; ctx.lineJoin = 'round'; ctx.strokeStyle = L.fg; ctx.lineWidth = 3;
    for (let k = 3; k >= 1; k--) ctx.strokeText(str, g.x0 - k * 34 * sep, g.base);
    ctx.fillStyle = L.fg; ctx.fillText(str, g.x0, g.base); ctx.letterSpacing = '0px';
    if (g.hasDot) disc(ctx, g.dot.x, g.dot.y, g.dot.r, L.dot);
  } else glyphLoop(ctx, g, sh, c, anim);
  ctx.restore();
  return g;
}
// the previous word leaves letter by letter (drawn over the new shot): word-to-word morph
function wordExit(ctx, sh, tl, col) {
  if (!sh.word) return;
  const g = geom(ctx, sh); ctx.save();
  g.lay.glyphs.forEach((gp, i) => {
    const q = E.inQuad(prog(tl, i * 0.02, i * 0.02 + 0.2)); if (q >= 1) return;
    if (g.hasDot && i === g.lay.glyphs.length - 1) disc(ctx, g.x0 + gp.cx, g.base - g.size * 0.088 - q * 1200, g.size * 0.088, C.red);
    else M.glyph(ctx, gp.ch, g.x0 + gp.cx, g.base - q * 1200, g.f, col, 1, 1 - 0.0 * q, q * (i % 2 ? 0.5 : -0.5));
  });
  ctx.restore();
}
// pseudo-shots for the title card and the closing stack
const TL = ['BILBRANSJE-', 'ANALYSEN', '2025'].map((w, i) => ({ word: w, L: LK.B, anim: ['letters', 'mask', 'snap'][i], nodot: true, align: 'l', max: 290, maxW: 1560, dy: -255 + i * 270 }));
const SK = ['BRANSJEKUNNSKAP.', '+ TEKNOLOGI.', '+ INNSIKT.'].map((w, i) => ({ word: w, L: LK.N, anim: ['slam', 'letters', 'crop'][i], align: 'l', size: 168, dy: -250 + i * 250 }));
SK[1].nodot = false;

// ------------------------------------------------------------------ one shot = bg + image layer + word layer
function shot(ctx, u, sh, o = {}) {
  const lt = u - sh.b, ls = lt * P, p = clamp(lt / sh.n), L = sh.L, c = { ls, lt, p: clamp(p), n: sh.n, L };
  if (sh.numDots) c.num01 = E.outCubic(prog(ls, 0.05, 0.5));
  ctx.save();
  const punch = 1 + 0.03 * (1 - E.outExpo(prog(ls, 0, 0.22))) * (sh.shake > 12 ? 1.7 : 1);
  ctx.translate(W / 2, H / 2); ctx.scale(punch, punch); ctx.translate(-W / 2, -H / 2);
  M.shake(ctx, lt, 0, sh.shake ?? 6, (sh.b * 7 + 3) | 0, 0.12);
  paintBg(ctx, L, u);
  const pp = Math.min(1.6, Math.max(0, p));
  ctx.save(); const zi = Math.exp(Math.log(1.06) * pp); ctx.translate(W / 2 - 16 * pp, H / 2 + 6 * pp); ctx.scale(zi, zi); ctx.translate(-W / 2, -H / 2); sh.im(ctx, c); ctx.restore();   // image layer: faster push
  ctx.save(); const zw = Math.exp(Math.log(1.022) * pp); ctx.translate(W / 2 + 14 * pp, H / 2 - 4 * pp); ctx.scale(zw, zw); ctx.translate(-W / 2, -H / 2);   // word layer: slower, opposite drift
  if (sh.title) { TL.forEach((w, i) => { const cc = { ...c, ls: ls - i * 0.12 }; if (cc.ls >= 0) drawWord(ctx, w, cc); }); }
  else if (sh.stack) { SK.forEach((w, i) => { const cc = { ...c, ls: ls - i * P }; if (cc.ls >= 0) drawWord(ctx, w, cc); }); }
  else {
    const wd = sh.tr && sh.tr.t === 'slide' ? 0.12 : 0, wg = ls >= wd ? drawWord(ctx, sh, { ...c, ls: ls - wd }) : geom(ctx, sh);
    if (sh.label) {
      ctx.font = M.font(34, 500, MONO); ctx.letterSpacing = '4px'; ctx.textAlign = sh.labelAlign || 'center'; ctx.fillStyle = L.fg;
      const q = E.outCubic(prog(ls, 0.15, 0.35)), ly = sh.labelY || wg.base + 94, lx = sh.labelX ?? W / 2 + 2;
      ctx.save(); ctx.beginPath(); ctx.rect(0, ly - 50, W, 90); ctx.clip();
      ctx.fillText(sh.label.toUpperCase(), lx, ly + (1 - q) * 70); ctx.restore(); ctx.letterSpacing = '0px';
    }
  }
  if (sh.kick) K.kicker(ctx, sh.kick, 96, 104, ls, { color: L.fg, size: 28 });
  if (sh.kick2) K.kicker(ctx, sh.kick2, 96, 154, ls, { color: L.fg, t0: 0.2, size: 28 });
  ctx.restore(); ctx.restore();
}

// ------------------------------------------------------------------ transitions (all on the beat; B is on screen at the cut frame)
const MAXR = 2400;
let OFFC = null;
function viaOff(ctx, u, sh, clipFn) {      // render B off-screen once, then blit it through the clip (one cheap clipped op)
  if (!OFFC) { OFFC = document.createElement('canvas'); OFFC.width = W; OFFC.height = H; }
  const o = OFFC.getContext('2d'); o.setTransform(1, 0, 0, 1, 0, 0); o.globalAlpha = 1; o.textBaseline = 'alphabetic';
  shot(o, u, sh);
  ctx.save(); clipFn(); ctx.drawImage(OFFC, 0, 0); ctx.restore();
}
const edge = (ctx, col, w) => { ctx.strokeStyle = col; ctx.lineWidth = w; };
const TR = {
  whip(ctx, u, A, B, lu, tr) {          // both on screen, sliding together; motion blur does the rest
    const e = 0.3 + 0.7 * E.outCubic(prog(lu, 0, tr.d)), [dx, dy] = tr.dir, ox = dx * W * e, oy = dy * H * e;
    for (const [sh, x, y] of [[A, ox, oy], [B, ox - dx * W, oy - dy * H]]) { ctx.save(); ctx.translate(x, y); ctx.beginPath(); ctx.rect(0, 0, W, H); ctx.clip(); shot(ctx, u, sh); ctx.restore(); }
  },
  lineWipe(ctx, u, A, B, lu, tr) {      // a red chart line draws across and pushes the frame
    const e = 0.3 + 0.7 * E.outCubic(prog(lu, 0, tr.d)), sx = W * (1 - e);
    ctx.save(); ctx.translate(-W * 0.3 * e, 0); shot(ctx, u, A); ctx.restore();
    const pts = []; for (let i = 0; i <= 18; i++) { const y = (H * i) / 18; pts.push([sx + 70 * Math.sin(i * 1.7 + 1) * (1 - 0.4 * e) + 40 * (K.hash01(i, 4) - 0.5), y]); }
    viaOff(ctx, u, B, () => { ctx.beginPath(); ctx.moveTo(W + 100, -10); pts.forEach(([x, y]) => ctx.lineTo(x, y)); ctx.lineTo(W + 100, H + 10); ctx.closePath(); ctx.clip(); });
    ctx.save(); ctx.lineJoin = 'round'; edge(ctx, C.red, 7); ctx.beginPath(); pts.forEach(([x, y], i) => (i ? ctx.lineTo(x, y) : ctx.moveTo(x, y))); ctx.stroke(); ctx.restore();
  },
  iris(ctx, u, A, B, lu, tr) {          // B opens from a point (the red dot of A) and becomes the frame
    const e = prog(lu, 0, tr.d), r = tr.r0 + (MAXR - tr.r0) * E.inQuad(e), [x, y] = tr.at;
    shot(ctx, u, A);
    viaOff(ctx, u, B, () => { ctx.beginPath(); ctx.arc(x, y, r, 0, 7); ctx.clip(); });
    if (r < 2000) ring(ctx, x, y, r, 8, C.red);
  },
  diag(ctx, u, A, B, lu, tr) {          // 35 degree wipe, the angle of the BDO cornerstone, growing from the bottom right
    const e = 0.3 + 0.7 * E.outCubic(prog(lu, 0, tr.d)), a = (35 * Math.PI) / 180, n = [Math.sin(a), Math.cos(a)], d = [Math.cos(a), -Math.sin(a)];
    const th = (W * n[0] + H * n[1]) * (1 - e) - 4, o = [n[0] * th, n[1] * th];   // a point on the edge line
    const P1 = [o[0] - d[0] * 3000, o[1] - d[1] * 3000], P2 = [o[0] + d[0] * 3000, o[1] + d[1] * 3000], Q2 = [P2[0] + n[0] * 3000, P2[1] + n[1] * 3000], Q1 = [P1[0] + n[0] * 3000, P1[1] + n[1] * 3000];
    shot(ctx, u, A);
    viaOff(ctx, u, B, () => { ctx.beginPath(); ctx.moveTo(...P1); ctx.lineTo(...P2); ctx.lineTo(...Q2); ctx.lineTo(...Q1); ctx.closePath(); ctx.clip(); });
    if (e < 0.98) { ctx.save(); edge(ctx, C.red, 6); ctx.beginPath(); ctx.moveTo(...P1); ctx.lineTo(...P2); ctx.stroke(); ctx.restore(); }
  },
  stopDot(ctx, u, A, B, lu, tr) {       // the full stop of A flies on and becomes the red point of B
    shot(ctx, u, B);
    const g = geom(ctx, A), e = E.inOutCubic(prog(lu, 0, tr.d)), d0 = g.dot || { x: W / 2, y: H / 2, r: 20 };
    disc(ctx, lerp(d0.x, tr.to[0], e), lerp(d0.y, tr.to[1], e) - 140 * Math.sin(Math.PI * e), lerp(d0.r, tr.r, e), C.red);
  },
  pushIris(ctx, u, A, B, lu, tr) {      // push through (log space) into the ring / the O: B lives inside it
    let at = tr.at, r0 = tr.r0;
    if (tr.find) { const g = geom(ctx, A), q = g.counter; if (q) { at = [q.x, q.y]; r0 = q.r; } else at = [W / 2, H / 2]; }
    const e = prog(lu, 0, tr.d), smax = Math.max(MAXR / r0 * 1.05, 2), s = Math.exp(Math.log(smax) * E.inQuad(e)), r = r0 * s;
    ctx.save(); ctx.translate(at[0], at[1]); ctx.scale(s, s); ctx.translate(-at[0], -at[1]); shot(ctx, u, A); ctx.restore();
    viaOff(ctx, u, B, () => { ctx.beginPath(); ctx.arc(at[0], at[1], r, 0, 7); ctx.clip(); });
    if (r < 2000) ring(ctx, at[0], at[1], r, Math.max(3, 7 * Math.sqrt(s) * 0.6), C.red);
  },
  slide(ctx, u, A, B, lu, tr) {         // same frame change, A's letters slide out while B's slide in
    shot(ctx, u, B); wordExit(ctx, A, lu * P, B.L.fg);
  },
};
const FRAMEY = () => {};

// ------------------------------------------------------------------ end card
function endCard(ctx, u) {
  const ls = (u - END) * P;
  ctx.fillStyle = C.white; ctx.fillRect(0, 0, W, H);
  K.kicker(ctx, 'BDO Automotive', 120, 330, ls, { color: C.blue, t0: 0.1, size: 30 });
  const lines = ['Vi ser detaljene', 'Vi forstår helheten'], ys = [500, 690];
  lines.forEach((s, i) => {
    const w = K.line(ctx, s, 120, ys[i], ls, { size: 150, weight: 300, color: C.charcoal, t0: 0.15 + 0.2 * i, dur: 0.6 });
    const sp = M.spring(ls - (0.5 + 0.2 * i), { stiffness: 420, damping: 15 });
    if (w && ls > 0.5 + 0.2 * i) { ctx.font = M.font(150, 300, DISPLAY); const tw = ctx.measureText(s).width; disc(ctx, 120 + tw + 20, ys[i] - 10 - (1 - sp) * 260, 10, C.red); }
  });
  K.cornerstone(ctx, E.outQuint(prog(ls, 0.5, 1.7)));
}

// ------------------------------------------------------------------ hits
const HITS = [];
const SEC = [16, 32, 48, 80, 96, 112];
SH.forEach((s) => {
  HITS.push([s.b, s.word || (s.title ? 'title' : 'stack'), s.sfx || 'thud', s.sfx === 'impact' ? { gain: 1 } : undefined].filter((x) => x !== undefined));
  if (s.tr && ['whip', 'lineWipe', 'pushIris', 'iris'].includes(s.tr.t)) HITS.push([s.b - 0.3, 'transition ' + s.tr.t, 'whoosh', { len: 0.4, from: 600, to: 2600 }]);
  if (s.tr && s.tr.t === 'diag') HITS.push([s.b - 0.15, 'wipe 35', 'whoosh', { len: 0.3, from: 1200, to: 3200 }]);
  if (s.tr && s.tr.t === 'stopDot') HITS.push([s.b + s.tr.d * 0.9, 'dot lands', 'tick', { pitch: 'A5' }]);
  if (s.pre) HITS.push([s.b + s.pre[0], s.pre[1], s.pre[2], s.pre[3]]);
  (s.hits || []).forEach((h) => HITS.push(Array.isArray(h) ? [s.b + h[0], h[1], h[2], h[3]].filter((x) => x !== undefined) : [s.b + h, 'hit', 'tick']));
});
SEC.forEach((b) => HITS.push([b - 1.5, 'riser', 'riser', { len: 1.4 }]));
HITS.push([END, 'card', 'impact'], [END + 1.9, 'cornerstone', 'impact'], [END + 1.15, 'red dot lands', 'tick', { pitch: 'A5' }]);
HITS.sort((a, b) => a[0] - b[0]);

let GRAIN = null;
const stage = (u) => { const k = idxAt(u), sh = SH[k], lu = u - sh.b; return { k, sh, lu, inTr: !!(sh.tr && k > 0 && lu < sh.tr.d) }; };
M.film({
  fonts: [M.font(100, 800, DISPLAY), M.font(100, 300, DISPLAY), M.font(40, 500, MONO)],
  hits: HITS,
  samples: (t) => { const u = t / P; if (u >= END) return 1; const s = stage(u); return s.inTr ? 4 : s.sh.hold ? 2 : 3; },
  async init() { await K.loadLogos(); const c = document.createElement('canvas').getContext('2d'); world.siteFast(c, { k: 0 }); world.siteFast(c, { k: 1 }); },
  draw(ctx, u, t) {
    if (u >= END) { endCard(ctx, u); return; }
    const s = stage(Math.max(0, u));
    if (s.inTr) TR[s.sh.tr.t](ctx, u, SH[s.k - 1], s.sh, s.lu, s.sh.tr); else shot(ctx, u, s.sh);
  },
  post(ctx, u, t) {
    if (u >= END || stage(u).sh.L.id !== 'N') return;
    const r = K.mulberry32(Math.floor(t * 30) + 1); ctx.fillStyle = 'rgba(255,255,255,0.05)'; for (let i = 0; i < 260; i++) ctx.fillRect(r() * W, r() * H, 2, 2);
  },
});
