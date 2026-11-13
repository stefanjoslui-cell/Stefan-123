// BDO Automotive – Action. 45 s trailer cut at 128 bpm: short hard cuts, ONE huge word per shot.
// A pure function of time: window.seek(t) paints frame t. No state between frames, no Math.random.
import * as M from './lib/motion.js';
import * as K from './lib/kit.js';
import * as world from './scenes/world.js';

const { W, H, C, clamp, lerp, prog, E, ss, rgba } = K;
const P = 0.46875;                       // one beat in seconds (128 bpm)
const END = 88;                          // beat where the end card starts (bar 23)
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
function paintBg(ctx, L, u) {
  ctx.fillStyle = L.bg; ctx.fillRect(-80, -80, W + 160, H + 160);
  if (L.id === 'N') { K.sheet(ctx, 0); K.grid(ctx, 0, 1, u * 9, u * 5); }
  else if (L.id === 'A') { K.sheet(ctx, 1); K.grid(ctx, 1, 1, -u * 9, u * 5); }
  else { K.grid(ctx, 0, L.id === 'R' ? 1.4 : 1.2, u * 9, -u * 5); }
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
    for (const g of [-1, 0, 1]) { if (g && v < 0.25) continue; ctx.globalAlpha = g ? 0.45 : 1; world.site(ctx, { k: 0, cx: cx + g * v * 4, cy: 18, s: 42 }); }
    ctx.restore();
  },
  plan(ctx, c) { world.site(ctx, { k: c.L.k, cx: 0, cy: 13, s: 12.5 + c.p * 1.6, alpha: c.L.id === 'B' ? 0.7 : 0.9 }); },
  planClose(ctx, c) { world.site(ctx, { k: c.L.k, cx: -6 + c.p * 12, cy: 12 + c.p * 3, s: 64, alpha: c.L.id === 'B' ? 0.8 : 1 }); },
  dept(d) {
    return (ctx, c) => {
      const x0 = BLD.x0 + BLD.w * d + 9.5, s = 44 + c.p * 5;
      const cam = { cx: x0 + (d % 2 ? -3 : 3), cy: -35, s };
      ctx.save(); ctx.globalAlpha = 1; rooms(ctx, cam, c.L.k, 1, (i) => (i === d ? 1 : 0.0)); ctx.letterSpacing = '0px'; ctx.restore();
    };
  },
  fieldAll(ctx, c) {
    const s = 1 + c.p * 0.06; ctx.save(); aboutXf(ctx, s, W / 2, H / 2); K.field(ctx, c.L.k, { alpha: c.L.id === 'B' ? 0.7 : 0.85 }); ctx.restore();
  },
  fieldPulse(ctx, c) {
    const s = 1 + 0.08 * Math.abs(Math.sin(c.lt * Math.PI)) + c.p * 0.04; ctx.save(); aboutXf(ctx, s, W / 2, H / 2);
    K.field(ctx, 0, { alpha: 0.75 }); ctx.restore();
  },
  net(ctx, c) {
    ctx.save(); K.field(ctx, 0, { alpha: 0.4, links: 0.8 }); netLines(ctx, clamp(c.ls / 0.9), C.white, 3); ctx.restore();
  },
  explode(ctx, c) {
    const q = E.outExpo(prog(c.ls, 0, 0.45)), f = q * 1.0 + c.p * 0.45;
    const pos = (p) => { const dx = p.x - W / 2, dy = p.y - H / 2; const j = 1 + p.seed * 0.5 * (1 - q); return [W / 2 + dx * f * j, H / 2 + dy * f * j, 1, 1 + (1 - q) * 0.5]; };
    K.field(ctx, c.L.k, { pos, alpha: 0.9, links: q });
  },
  fly(ctx, c) {
    const f = lerp(1.7, 3.3, E.outCubic(c.p)); const pos = (p) => { const dx = p.x - 700, dy = p.y - 520; return [W / 2 + dx * f, H / 2 + dy * f, 1, 2.2]; };
    K.field(ctx, c.L.k, { pos, alpha: 0.85, links: 0.5 });
  },
  zoomCluster(ctx, c) {
    ctx.save(); aboutXf(ctx, 3.1 + c.p * 0.5, CL[0].x, CL[0].y); K.field(ctx, c.L.k, { alpha: 0.85, red: 0 }); ctx.restore();
  },
  labels(ctx, c) {
    K.field(ctx, 0, { alpha: 0.6, links: 0.6 });
    ctx.font = M.font(30, 500, MONO); ctx.letterSpacing = '4px'; ctx.textAlign = 'left';
    CL.forEach((cl, i) => { const q = E.outCubic(prog(c.ls, i * 0.05, i * 0.05 + 0.2)); if (q <= 0) return; ctx.fillStyle = rgba(C.white, q); disc(ctx, cl.x, cl.y, 8, C.white); ctx.fillText(cl.name.toUpperCase(), cl.x + 20, cl.y - 20); });
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
    K.field(ctx, 0, { pos, alpha: 0.8, links: f });
    const r = 22 * (1 - f) + 6 * f;
    if (f < 0.05) { disc(ctx, tx, ty, 24 + (c.lt - 0.7) * 0, C.red); }
  },
  lattice(ctx, c) {
    const q = E.outExpo(prog(c.ls, 0, 0.24)), gap = lerp(380, 118, q);
    lattice(ctx, { gap, col: C.white, a: 0.5 });
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
    chart(ctx, { x: -30, y: 80, w: W + 100, h: 900, p, col: C.white, lwid: 12, steep: 3.2, tipCol: C.white, tipR: 22, axes: false, seed: 9 });
  },
  bars(ctx, c) {
    for (let i = 0; i < 9; i++) {
      const q = E.outBack(prog(c.ls, i * 0.05, i * 0.05 + 0.22), 1.6), hgt = (260 + i * 62 + K.hash01(i, 3) * 40) * q, w = 150, x = 96 + i * 190;
      ctx.fillStyle = rgba(C.cloud, 0.95); ctx.fillRect(x, 1000 - hgt, w, hgt); ctx.strokeStyle = C.graphite; lw(ctx, 2); ctx.strokeRect(x, 1000 - hgt, w, hgt);
      if (i === 8) disc(ctx, x + w / 2, 1000 - hgt - 6, 14 * q, C.red);
    }
  },
  links(ctx, c) {
    const L = c.L, base = c.mode === 1 ? 1.5 : 1;
    ctx.save(); aboutXf(ctx, base + c.p * 0.15, W / 2, H / 2); K.field(ctx, L.k, { alpha: 0.35, links: 1.2, red: 0.8 });
    const r = K.mulberry32(c.mode === 1 ? 31 : 17);
    ctx.lineCap = 'round'; ctx.strokeStyle = L.line; lw(ctx, 3);
    for (let i = 0; i < 16; i++) {
      const a = K.PTS[1 + Math.floor(r() * 2399)], b = K.PTS[1 + Math.floor(r() * 2399)], q = E.outExpo(prog(c.ls, i * 0.035, i * 0.035 + 0.18));
      K.drawOn(ctx, [[a.x, a.y], [b.x, b.y]], q); if (q > 0.9) { disc(ctx, a.x, a.y, 7, L.line); disc(ctx, b.x, b.y, 7, L.line); }
    }
    ctx.restore();
  },
  cross(ctx, c) {
    const q = E.outExpo(prog(c.ls, 0, 0.2)); ctx.strokeStyle = rgba(C.white, 0.55); lw(ctx, 2); ctx.beginPath();
    ctx.moveTo(0, H / 2); ctx.lineTo(W * q, H / 2); ctx.moveTo(W / 2, 0); ctx.lineTo(W / 2, H * q); ctx.stroke();
    ring(ctx, W / 2, H / 2, 520 * (1.5 - 0.5 * q), 2, rgba(C.white, 0.35));
  },
  dial(ctx, c) {
    ctx.save(); ctx.translate(W / 2, H / 2); ctx.rotate(c.ls * 0.25); ctx.strokeStyle = rgba(C.white, 0.5);
    ring(ctx, 0, 0, 470, 2, rgba(C.white, 0.5)); ring(ctx, 0, 0, 380, 1.5, rgba(C.white, 0.3));
    ctx.beginPath(); for (let i = 0; i < 72; i++) { const a = i * Math.PI / 36, r0 = i % 6 ? 470 : 440; ctx.moveTo(Math.cos(a) * r0, Math.sin(a) * r0); ctx.lineTo(Math.cos(a) * 490, Math.sin(a) * 490); } lw(ctx, 2); ctx.stroke();
    ctx.restore();
  },
  herField(ctx, c) {          // the field with the red point shifted to the right; ring snaps around it
    const hold = c.mode === 1, s = 1 + c.p * 0.04;
    ctx.save(); ctx.translate(1350, 640); ctx.scale(s, s); ctx.translate(-560, -690);
    K.field(ctx, c.L.k, { alpha: hold ? 0.55 : 0.7, links: 0.8, red: 0 }); ctx.restore();
    const t0 = hold ? -1 : 0.28, q = hold ? 1 : E.outExpo(prog(c.ls, t0, t0 + 0.12)), r = hold ? 38 : lerp(420, 38, q) - (hold ? 0 : 5 * Math.sin(prog(c.ls, t0 + 0.1, t0 + 0.3) * Math.PI));
    const x = 1350, y = 640;
    disc(ctx, x, y, 9, C.red);
    if (!hold && c.ls < t0) return;
    ring(ctx, x, y, r, hold ? 3 : 6, C.red);
    if (!hold) {
      const tq = E.outCubic(prog(c.ls, t0 + 0.12, t0 + 0.4));
      if (tq > 0) {
        ctx.strokeStyle = C.red; lw(ctx, 2); ctx.beginPath(); ctx.moveTo(x, y + 44); ctx.lineTo(x, y + 44 + 96 * tq); ctx.stroke();
        ctx.save(); ctx.beginPath(); ctx.rect(1300, 740, 560, 70); ctx.clip();
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
    K.field(ctx, 0, { alpha: 0.3, links: 0.5, red: 0 });
    const r = K.mulberry32(53); ctx.lineCap = 'round';
    for (let i = 0; i < 18; i++) {
      const p = K.PTS[1 + Math.floor(r() * 2399)], q = E.outBack(prog(c.ls, i * 0.025, i * 0.025 + 0.15), 2);
      ring(ctx, p.x, p.y, 26 * q, 3, C.white); disc(ctx, p.x, p.y, 7, C.white);
    }
  },
  siteFor(ctx, c) { world.site(ctx, { k: 0, cx: -6 + c.p * 10, cy: -29, s: 24 + c.p * 5 }); },
  siteBack(ctx, c) {
    const e = E.inOutCubic(prog(c.ls, 0, 0.95 * 4 * P)), s = lerp(34, 10.8, e);
    world.site(ctx, { k: 1, cx: 0, cy: lerp(10, 0, e), s });
    const [x, y] = toScreen(0, lerp(10, 0, e), s, 0, 12); disc(ctx, x, y, 9, C.red); ring(ctx, x, y, 16 + 10 * e, 2, rgba(C.red, 0.8));
  },
  numDots(ctx, c) {   // 1 300 points, one per company
    const v = c.num01 ?? 1, n = Math.floor(v * 1300);
    for (let i = 0; i < 1300; i++) { const col = i % 52, row = Math.floor(i / 52), x = 114 + col * 33.2, y = 150 + row * 33.2; disc(ctx, x, y, i < n ? 4.2 : 2.4, i < n ? rgba(C.blue, 0.55) : rgba(C.blue, 0.22)); }
  },
  base(ctx, c) {
    const q = E.outExpo(prog(c.ls, 0, 0.4)); ctx.strokeStyle = rgba(c.L.line, 0.6); lw(ctx, 2); K.drawOn(ctx, [[96, 880], [1824, 880]], q);
    for (let i = 0; i <= 12; i++) { ctx.beginPath(); ctx.moveTo(96 + i * 144, 880); ctx.lineTo(96 + i * 144, 868); ctx.stroke(); }
    ctx.fillStyle = C.red; ctx.fillRect(1824 - 6, 850, 6, 30 * q);
  },
};

// ------------------------------------------------------------------ the shot list
// S(word, beats, look, anim, image, extra). anim: slam | L | R | U | D | letters
const SH = [];
let cur = 0;
function S(word, n, look, anim, im, x = {}) { SH.push({ word, n, b: cur, L: LK[look], anim, im, ...x }); cur += n; }
const num = (f) => ({ num: f });

S('ÉN.', 2, 'N', 'slam', IM.dot, { align: 'l', sfx: 'impact', shake: 14 });
S('TRANSAKSJON.', 3, 'A', 'L', IM.ledger, { dy: -90, sfx: 'thud', hits: [[0.5, 'amount', 'tick']] });
S('ÉN.', 2, 'N', 'slam', IM.wheel, { sfx: 'thud', shake: 12 });
S('BIL.', 3, 'A', 'U', IM.car, { dy: -190, sfx: 'thud', shake: 8, hits: [[1.8, 'car drawn', 'pop', { pitch: 'A4' }]] });
S('TUSEN.', 2, 'N', 'R', IM.whip, { sfx: 'thud', pre: [-0.35, 'whip', 'whoosh', { len: 0.5, from: 500, to: 2600 }] });
S('TUSEN.', 1, 'B', 'slam', IM.planClose, { sfx: 'click', shake: 9 });
S('BILER.', 2, 'A', 'D', IM.plan, { sfx: 'thud', shake: 10 });
S('BILER.', 2, 'N', 'letters', IM.planClose, { sfx: 'click', shake: 6 });
['SALG.', 'LAGER.', 'VERKSTED.', 'DELER.', 'GARANTI.', 'SERVICE.'].forEach((w, i) =>
  S(w, 1, i % 2 ? 'N' : 'A', ['L', 'R', 'slam', 'L', 'R', 'slam'][i], IM.dept(i), { sfx: 'click', shake: 7, dy: 0 }));
S('ALT.', 2, 'A', 'slam', IM.fieldAll, { sfx: 'thud', shake: 12 });
S('HENGER.', 2, 'B', 'R', IM.fieldPulse, { sfx: 'thud', shake: 8 });
S('SAMMEN.', 3, 'N', 'slam', IM.net, { sfx: 'thud', shake: 10, hits: [0.15, 0.3, 0.45, 0.6, 0.8, 1.0, 1.2].map((h, i) => [h, 'link', 'tick', { pitch: 1200 + i * 110 }]) });
S('DATA.', 2, 'B', 'slam', IM.explode, { sfx: 'impact', shake: 20 });
S('DATA.', 1, 'N', 'slam', IM.fly, { sfx: 'thud', shake: 10 });
S('FRA.', 1, 'A', 'L', IM.zoomCluster, { sfx: 'click', shake: 6 });
S('ALLE.', 1, 'N', 'R', IM.labels, { sfx: 'click', shake: 6 });
S('SYSTEMER.', 2, 'A', 'letters', IM.streams, { sfx: 'thud', shake: 8, pre: [-0.3, 'stream', 'whoosh', { len: 0.4, from: 800, to: 3000 }] });
S('SYSTEMER.', 1, 'B', 'slam', IM.streams, { sfx: 'click', shake: 6 });
S('ÉN.', 1, 'N', 'slam', IM.converge, { sfx: 'sub', shake: 8, dy: 160, pre: [-0.5, 'draw in', 'riser', { len: 0.45 }] });
S('PLATTFORM.', 4, 'R', 'slam', IM.lattice, { sfx: 'impact', shake: 28, hits: [[0.5, 'layer lock', 'sub', { len: 0.8 }]] });
S('BRANSJE-', 2, 'A', 'L', IM.rows, { sfx: 'thud', shake: 8 });
S('KUNNSKAP.', 2, 'N', 'R', IM.rearwheel, { sfx: 'thud', shake: 10 });
const KICK = 'BDO Bilbransjeanalysen 2025';
S('1,7 %', 3, 'B', 'slam', IM.base, { ...num((c) => `${lerp(1.5, 1.7, E.outCubic(prog(c.ls, 0, 0.3))).toFixed(1).replace('.', ',')} %`), fitStr: '1,7 %', label: 'Driftsmargin merkeforhandlere 2024', kick: KICK, sfx: 'thud', shake: 10, hits: [[0.8, 'count', 'tick']], dy: -40 });
S('2,5 %', 3, 'N', 'slam', IM.base, { ...num((c) => `${lerp(1.9, 2.5, E.outCubic(prog(c.ls, 0, 0.3))).toFixed(1).replace('.', ',')} %`), fitStr: '2,5 %', label: 'Bruktbilforhandlere 2024', kick: KICK, sfx: 'thud', shake: 10, hits: [[0.8, 'count', 'tick']], dy: -40 });
S('1 300+', 3, 'A', 'slam', IM.numDots, { ...num((c) => NF(`${Math.round(lerp(880, 1300, E.outCubic(prog(c.ls, 0, 0.4)))).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ' ')}${c.ls > 0.4 ? '+' : ''}`)), fitStr: NF('1 300+'), numDots: true, label: 'Bedrifter analysert', kick: KICK, sfx: 'thud', shake: 10, hits: [[0.8, 'count', 'tick']], dy: -40 });
S('MØNSTRE.', 2, 'N', 'letters', IM.chartUp, { sfx: 'thud', shake: 8, hits: [[0.3, 'chart', 'whoosh', { len: 0.4, from: 400, to: 2400 }]] });
S('MØNSTRE.', 1, 'A', 'slam', IM.chartZoom, { sfx: 'click', shake: 8 });
S('SAMMENHENGER.', 2, 'B', 'L', (ctx, c) => IM.links(ctx, c), { sfx: 'thud', shake: 8 });
S('SAMMENHENGER.', 1, 'N', 'R', (ctx, c) => IM.links(ctx, { ...c, mode: 1 }), { sfx: 'click', shake: 6 });
S('SE.', 1, 'B', 'slam', IM.cross, { sfx: 'thud', shake: 12 });
S('HER.', 3, 'N', 'slam', (ctx, c) => IM.herField(ctx, c), { align: 'l', dy: -90, sfx: 'thud', shake: 12, hits: [[0.6, 'ring snap', 'impact']] });
S('REVISOR.', 1, 'B', 'U', IM.dial, { sfx: 'click', shake: 4 });
S('VURDERER.', 4, 'A', 'D', (ctx, c) => IM.herField(ctx, { ...c, mode: 1 }), { sfx: 'sub', shake: 3, hold: true });
S('MÅLRETTET.', 2, 'N', 'slam', IM.target, { sfx: 'thud', shake: 10 });
S('REVISJON.', 2, 'A', 'R', IM.scan, { sfx: 'thud', shake: 8 });
S('RELEVANT.', 2, 'B', 'letters', IM.picks, { sfx: 'thud', shake: 6 });
S('INNSIKT.', 2, 'N', 'slam', IM.mirror, { sfx: 'thud', shake: 10, hits: [[0.3, 'insight', 'pop', { pitch: 'E6' }]] });
S('MER.', 1, 'A', 'L', IM.bars, { sfx: 'click', shake: 6 });
S('VERDI.', 3, 'R', 'slam', IM.chartSteep, { sfx: 'impact', shake: 28, hits: [[0.6, 'rise', 'sub', { len: 0.8 }]] });
S('FOR.', 1, 'N', 'R', IM.siteFor, { sfx: 'click', shake: 6 });
S('DEG.', 4, 'A', 'slam', IM.siteBack, { sfx: 'thud', shake: 12, hits: [[2.5, 'whole', 'swell', { len: 1.2 }]] });
if (cur !== END) throw new Error('shot list ends at beat ' + cur + ', expected ' + END);

const STARTS = SH.map((s) => s.b);
const shotAt = (u) => { let i = 0; for (let k = 0; k < SH.length; k++) if (STARTS[k] <= u) i = k; else break; return SH[i]; };

// ------------------------------------------------------------------ the word
const SIZE = new Map();
function wordGeom(ctx, sh) {
  const str = sh.fitStr || sh.word, key = str;
  if (!SIZE.has(key)) {
    const size = sh.size || M.fitSize(ctx, str, 800, DISPLAY, sh.maxW || 1728, 420, -0.03);
    const f = M.font(size, 800, DISPLAY), lay = M.layout(ctx, str, f, -0.03 * size);
    SIZE.set(key, { size, f, lay });
  }
  return SIZE.get(key);
}
function drawWord(ctx, sh, c) {
  const L = c.L, { size, f, lay } = wordGeom(ctx, sh), width = lay.width, tr = -0.03 * size;
  const x0 = (sh.align === 'l' ? 96 : sh.align === 'r' ? W - 96 - width : (W - width) / 2) + (sh.dx || 0);
  const base = H / 2 + 0.36 * size + (sh.dy || 0), pivX = x0 + width / 2, pivY = base - 0.36 * size;
  let ox = 0, oy = 0, sc = 1;
  const ls = c.ls;
  if (sh.anim === 'slam') { const q = E.outBack(prog(ls, 0, 0.14), 2.6); sc = 1 + 0.25 * (1 - q); }
  else if (sh.anim === 'L') ox = -(1 - E.outExpo(prog(ls, 0, 0.16))) * 130;
  else if (sh.anim === 'R') ox = (1 - E.outExpo(prog(ls, 0, 0.16))) * 130;
  else if (sh.anim === 'U') oy = (1 - E.outExpo(prog(ls, 0, 0.16))) * 90;
  else if (sh.anim === 'D') oy = -(1 - E.outExpo(prog(ls, 0, 0.16))) * 90;
  ctx.save();
  ctx.translate(pivX + ox, pivY + oy); ctx.scale(sc, sc); ctx.translate(-pivX, -pivY);
  if (sh.num) {
    ctx.font = f; ctx.letterSpacing = `${tr}px`; ctx.fillStyle = L.fg; ctx.textAlign = 'left'; ctx.fillText(sh.num(c), x0, base); ctx.letterSpacing = '0px';
  } else {
    lay.glyphs.forEach((g, i) => {
      let gx = x0 + g.cx, gy = base, gs = 1, d = 0;
      if (sh.anim === 'letters') {
        d = 0.012 * i;
        const q = E.outBack(prog(ls - d, 0, 0.14), 2.4); gs = 1 + 0.35 * (1 - q); gy = base - (1 - q) * 50;
      }
      if (g.ch === '.' && i === lay.glyphs.length - 1) { const r = size * 0.088; disc(ctx, gx, gy - r, r * gs, L.dot); }
      else M.glyph(ctx, g.ch, gx, gy, f, L.fg, gs, gs);
    });
  }
  ctx.restore();
  return { x0, base, size, width };
}

// ------------------------------------------------------------------ frame
function shot(ctx, u, t, sh) {
  const lt = u - sh.b, ls = lt * P, p = clamp(lt / sh.n), L = sh.L;
  const c = { ls, lt, p, n: sh.n, L };
  ctx.save();
  // camera: always moving. micro push (log zoom to 1.04) + K-style shake on the hit
  const z = Math.exp(Math.log(1.04) * p * (sh.hold ? 1 : 1));
  const punch = 1 + 0.025 * (1 - E.outExpo(prog(ls, 0, 0.25))) * (sh.shake > 12 ? 1.6 : 1);
  ctx.translate(W / 2, H / 2); ctx.scale(z * punch, z * punch); ctx.translate(-W / 2, -H / 2);
  M.shake(ctx, lt, 0, sh.shake ?? 6, (sh.b * 7 + 3) | 0, 0.12);
  paintBg(ctx, L, u);
  if (sh.numDots) c.num01 = E.outCubic(prog(ls, 0.05, 0.5));
  sh.im(ctx, c);
  const wg = drawWord(ctx, sh, c);
  if (sh.label) {
    ctx.font = M.font(40, 500, MONO); ctx.letterSpacing = '5px'; ctx.textAlign = 'center'; ctx.fillStyle = L.fg;
    const q = E.outCubic(prog(ls, 0.15, 0.35));
    ctx.save(); ctx.beginPath(); ctx.rect(0, wg.base + 30, W, 90); ctx.clip();
    ctx.fillText(sh.label.toUpperCase(), W / 2 + 2, wg.base + 94 + (1 - q) * 70); ctx.restore(); ctx.letterSpacing = '0px';
  }
  if (sh.kick) K.kicker(ctx, sh.kick, 96, 120, ls, { color: L.fg, t0: 0, size: 28 });
  ctx.restore();
}

function endCard(ctx, u, t) {
  const ls = (u - END) * P;
  ctx.fillStyle = C.white; ctx.fillRect(0, 0, W, H);
  K.drawLogo(ctx, W / 2, 400, 160);
  const f = M.font(44, 300, DISPLAY); ctx.font = f; ctx.letterSpacing = '0px';
  const tag = 'Vi ser detaljene. Vi forstår helheten.', tw = ctx.measureText(tag).width;
  K.line(ctx, tag, W / 2, 640, ls, { size: 44, weight: 300, color: C.graphite, t0: 0.12, align: 'center', dur: 0.45 });
  const q = E.outQuint(prog(ls, 0.45, 0.75));
  if (q > 0) {
    ctx.save(); ctx.font = M.font(24, 500, MONO); ctx.letterSpacing = '5px'; ctx.textAlign = 'center'; ctx.fillStyle = C.blue;
    ctx.beginPath(); ctx.rect(0, 690, W, 60); ctx.clip(); ctx.fillText('TETT PÅ BILBRANSJEN', W / 2 + 2, 732 + (1 - q) * 40); ctx.restore();
  }
  // the red dot lands beside the tagline
  const sp = M.spring(ls - 0.55, { stiffness: 420, damping: 15 });
  if (ls > 0.55) disc(ctx, W / 2 + tw / 2 + 22, 631 - (1 - sp) * 280, 8, C.red);
}

// ------------------------------------------------------------------ hits
const HITS = [];
SH.forEach((s) => {
  HITS.push([s.b, s.word, s.sfx || 'thud', s.sfx === 'impact' ? { gain: 1 } : undefined].filter((x) => x !== undefined));
  if (s.pre) HITS.push([s.b + s.pre[0], s.pre[1], s.pre[2], s.pre[3]]);
  (s.hits || []).forEach((h) => HITS.push(Array.isArray(h) ? [s.b + h[0], h[1], h[2], h[3]].filter((x) => x !== undefined) : [s.b + h, 'hit', 'tick']));
});
HITS.push([END, 'logo', 'impact'], [END + 1.15, 'red dot lands', 'tick', { pitch: 'A5' }]);
HITS.sort((a, b) => a[0] - b[0]);

M.film({
  fonts: [M.font(100, 800, DISPLAY), M.font(100, 300, DISPLAY), M.font(40, 500, MONO)],
  hits: HITS,
  samples: 6,
  async init() { await K.loadLogo(); },
  draw(ctx, u, t) {
    if (u >= END) { endCard(ctx, u, t); return; }
    shot(ctx, Math.max(0, u), t, shotAt(u));
  },
  post(ctx, u, t) { if (u < END && shotAt(u).L.id === 'N') K.grain(ctx, t, 0); },
});
