// kit.js: the shared drawing kit. Every scene module uses ONLY these colours, faces and text styles,
// so the film reads as one hand: a precise architect's line drawing (night sheets for the physical
// world, bright sheets for the digital one) with a single red point as the signature.
import * as M from './motion.js';
export const W = 1920, H = 1080;
// BDO's official colour profile (Pantone/CMYK/RGB from the BDO brand portal) + two neutral working tones.
// red 185, blue 287, charcoal Cool Grey 11, slate 174+10, burgundy 208, pale charcoal.
export const C = { red: '#E81A3B', blue: '#22409A', charcoal: '#333333', slate: '#5B6E7F', burgundy: '#98002E', pale: '#E7E7E7',
  white: '#FFFFFF', arctic: '#F7F7F7', night: '#17191C',
  // role aliases used by the scenes (mapped onto the official colours)
  cloud: '#E7E7E7', graphite: '#333333', light: '#5B6E7F', teal: '#5B6E7F' };
export const DISPLAY = 'Display', MONO = 'Mono';
export const { clamp, lerp, prog, E, mix, hash01, noise1, mulberry32, spring, SPRING, kf } = M;
export const ss = (a, b, x) => { const t = clamp((x - a) / (b - a)); return t * t * (3 - 2 * t); };
export const rgba = (hex, a) => { const [r, g, b] = M.hexRgb(hex); return `rgba(${r},${g},${b},${clamp(a)})`; };
// window: 0 outside [a,b], eases in over `i` seconds and out over `o`
export const win = (t, a, b, i = 0.6, o = 0.6) => ss(a, a + i, t) * (1 - ss(b - o, b, t));

// ---------------------------------------------------------------- sheets (backgrounds)
// k = 0 night sheet, 1 bright sheet; in between is the dawn the film uses to move from physical to digital
export function sheet(ctx, k = 0) {
  ctx.fillStyle = mix(C.night, C.arctic, k); ctx.fillRect(0, 0, W, H);
  const g = ctx.createRadialGradient(W / 2, H * 0.46, 100, W / 2, H * 0.5, 1250);
  if (k < 0.5) { g.addColorStop(0, rgba(C.graphite, 0.55 * (1 - 2 * k))); g.addColorStop(1, rgba(C.graphite, 0)); }
  else { g.addColorStop(0, rgba(C.white, (k - 0.5) * 2)); g.addColorStop(1, rgba(C.cloud, 0.55 * (k - 0.5) * 2)); }
  ctx.fillStyle = g; ctx.fillRect(0, 0, W, H);
}
// faint drafting grid, in screen space (parallax: pass an offset)
export function grid(ctx, k, alpha = 1, ox = 0, oy = 0, step = 120) {
  ctx.save(); ctx.lineWidth = 1; ctx.strokeStyle = k < 0.5 ? rgba(C.light, 0.07 * alpha) : rgba(C.blue, 0.06 * alpha);
  ctx.beginPath();
  for (let x = ((ox % step) + step) % step; x < W; x += step) { ctx.moveTo(x, 0); ctx.lineTo(x, H); }
  for (let y = ((oy % step) + step) % step; y < H; y += step) { ctx.moveTo(0, y); ctx.lineTo(W, y); }
  ctx.stroke(); ctx.restore();
}

// ---------------------------------------------------------------- lines
// a polyline drawn on from its start: p in 0..1 (pts = [[x,y],...])
export function drawOn(ctx, pts, p = 1, close = false) {
  if (p <= 0 || pts.length < 2) return;
  const n = (pts.length - 1) * clamp(p), k = Math.floor(n), f = n - k;
  ctx.beginPath(); ctx.moveTo(pts[0][0], pts[0][1]);
  for (let i = 1; i <= k; i++) ctx.lineTo(pts[i][0], pts[i][1]);
  if (k < pts.length - 1 && f > 0) ctx.lineTo(lerp(pts[k][0], pts[k + 1][0], f), lerp(pts[k][1], pts[k + 1][1], f));
  if (close && p >= 1) ctx.closePath();
  ctx.stroke();
}
// sample a cubic bezier chain into a polyline: segs = [[x0,y0], [c1x,c1y,c2x,c2y,x,y], ...]
export function bez(segs, n = 14) {
  const out = [segs[0].slice(0, 2)]; let [x0, y0] = segs[0];
  for (let s = 1; s < segs.length; s++) {
    const q = segs[s];
    if (q.length === 2) { out.push([q[0], q[1]]); [x0, y0] = q; continue; }
    for (let i = 1; i <= n; i++) { const t = i / n, u = 1 - t; out.push([u * u * u * x0 + 3 * u * u * t * q[0] + 3 * u * t * t * q[2] + t * t * t * q[4], u * u * u * y0 + 3 * u * u * t * q[1] + 3 * u * t * t * q[3] + t * t * t * q[5]]); }
    x0 = q[4]; y0 = q[5];
  }
  return out;
}

// ---------------------------------------------------------------- the red signature
export function redDot(ctx, x, y, r = 7, a = 1) {
  if (a <= 0 || r <= 0) return;
  const g = ctx.createRadialGradient(x, y, 0, x, y, r * 4.5);
  g.addColorStop(0, rgba(C.red, 0.35 * a)); g.addColorStop(1, rgba(C.red, 0));
  ctx.fillStyle = g; ctx.beginPath(); ctx.arc(x, y, r * 4.5, 0, 7); ctx.fill();
  ctx.fillStyle = rgba(C.red, a); ctx.beginPath(); ctx.arc(x, y, r, 0, 7); ctx.fill();
}

// ---------------------------------------------------------------- type
// One line of type that rises out of its own baseline mask (never a plain fade).
// t0 = when it enters, out = when it leaves (optional). Returns the measured width.
export function line(ctx, str, x, y, t, { size = 64, weight = 600, color = C.white, t0 = 0, out = 1e9, align = 'left', track = 0, family = DISPLAY, dur = 0.7 } = {}) {
  const pi = E.outQuint(prog(t, t0, t0 + dur)), po = E.inCubic(prog(t, out - 0.45, out));
  if (pi <= 0 || po >= 1) return 0;
  ctx.save();
  ctx.font = M.font(size, weight, family); ctx.letterSpacing = `${track * size}px`; ctx.textAlign = align; ctx.textBaseline = 'alphabetic';
  const w = ctx.measureText(str).width, x0 = align === 'center' ? x - w / 2 : align === 'right' ? x - w : x;
  ctx.beginPath(); ctx.rect(x0 - 20, y - size * 1.05, w + 40, size * 1.32); ctx.clip();
  ctx.fillStyle = color; ctx.globalAlpha *= (1 - po);
  ctx.fillText(str, x, y + (1 - pi) * size * 1.15 - po * size * 0.5);
  ctx.restore();
  return w;
}
// Small mono label with the red tick in front: department names, stage names, sources.
export function kicker(ctx, str, x, y, t, { color = C.cloud, t0 = 0, out = 1e9, size = 22, tick = true } = {}) {
  const pi = E.outQuint(prog(t, t0, t0 + 0.5)), po = prog(t, out - 0.35, out);
  if (pi <= 0 || po >= 1) return;
  ctx.save(); ctx.globalAlpha *= pi * (1 - po);
  if (tick) { ctx.fillStyle = C.red; ctx.fillRect(x, y - size * 0.36, 26 * pi, 3); }
  ctx.font = M.font(size, 500, MONO); ctx.letterSpacing = `${size * 0.12}px`; ctx.fillStyle = color; ctx.textAlign = 'left';
  ctx.fillText(str.toUpperCase(), x + (tick ? 40 : 0), y);
  ctx.restore();
}
// The film's caption position: big statement lines, lower left, on every sheet.
export const CAP = { x: 120, y: 930 };
export const ink = (k) => (k < 0.5 ? C.white : C.graphite);         // text colour for a sheet
export const sub = (k) => (k < 0.5 ? C.cloud : C.blue);             // secondary line/text colour

// ---------------------------------------------------------------- the data field
// N points with a seeded home in the "network" layout: 7 clusters (six departments + the car park)
// joined by hub links. Index 0 is THE red point. Screen px, frame 1920x1080.
export const CLUSTERS = [
  { name: 'Salg', x: 420, y: 330 }, { name: 'Lager', x: 760, y: 250 }, { name: 'Verksted', x: 1120, y: 300 },
  { name: 'Deler', x: 1480, y: 380 }, { name: 'Garanti', x: 1360, y: 700 }, { name: 'Service', x: 980, y: 760 }, { name: 'Bilpark', x: 560, y: 690 },
];
export const NPTS = 2400;
export const PTS = (() => {
  const r = mulberry32(77), out = [];
  for (let i = 0; i < NPTS; i++) {
    const c = i === 0 ? 6 : Math.floor(r() * 7), a = r() * Math.PI * 2, d = 150 * Math.sqrt(-Math.log(1 - r() * 0.96)) * 0.62;
    out.push({ i, c, x: CLUSTERS[c].x + Math.cos(a) * d * 1.15, y: CLUSTERS[c].y + Math.sin(a) * d * 0.8, r: 1.3 + r() * r() * 3.2, seed: r(), z: r() });
  }
  out[0].x = 560; out[0].y = 690; out[0].r = 7;
  return out;
})();
export const LINKS = (() => {       // short links inside clusters + the hub ring
  const r = mulberry32(5), out = [];
  for (let k = 0; k < 520; k++) {
    const a = PTS[1 + Math.floor(r() * (NPTS - 1))]; let best = null, bd = 1e9;
    for (let q = 0; q < 14; q++) { const b = PTS[1 + Math.floor(r() * (NPTS - 1))]; const d = Math.hypot(a.x - b.x, a.y - b.y); if (b !== a && d < bd && (b.c === a.c || r() < 0.08)) { bd = d; best = b; } }
    if (best && bd < 260) out.push([a.i, best.i]);
  }
  return out;
})();
// draw the field in its network layout (k = sheet value: light points on night, blue on bright).
// pos(p) may return [x, y, alpha, radiusScale] to move points; default = home positions.
export function field(ctx, k, { alpha = 1, pos = null, links = 1, red = 1 } = {}) {
  const col = k < 0.5 ? C.cloud : C.blue, P = pos ? PTS.map(pos) : null;
  const at = (i) => (P ? P[i] : [PTS[i].x, PTS[i].y, 1, 1]);
  if (links > 0) {
    ctx.lineWidth = 1; ctx.strokeStyle = rgba(k < 0.5 ? C.light : C.blue, 0.22 * alpha * links); ctx.beginPath();
    for (const [a, b] of LINKS) { const A = at(a), B = at(b); if ((A[2] ?? 1) * (B[2] ?? 1) < 0.5) continue; ctx.moveTo(A[0], A[1]); ctx.lineTo(B[0], B[1]); }
    ctx.stroke();
  }
  for (let i = 1; i < NPTS; i++) {
    const p = PTS[i], q = at(i), a = (q[2] ?? 1) * alpha * (0.35 + 0.65 * p.z);
    if (a <= 0.01) continue;
    ctx.fillStyle = rgba(col, a); ctx.beginPath(); ctx.arc(q[0], q[1], p.r * (q[3] ?? 1), 0, 7); ctx.fill();
  }
  const q = at(0); redDot(ctx, q[0], q[1], 7 * (q[3] ?? 1), red * (q[2] ?? 1));
}
// film grain + vignette: call from post() only
export function grain(ctx, t, k) {
  const r = mulberry32(Math.floor(t * 30) + 1);
  ctx.fillStyle = k < 0.5 ? 'rgba(255,255,255,0.035)' : 'rgba(38,52,69,0.03)';
  for (let i = 0; i < 1400; i++) ctx.fillRect(r() * W, r() * H, 1.5, 1.5);
}

// ---------------------------------------------------------------- the official BDO logo
// The logo file lives in <repo>/brand/bdo-logo.svg (or .png). The packed HTML inlines it as
// window.BDO_LOGO; the dev server/renderer loads it from ../brand/. Until the file is there,
// drawLogo() sets a plain typographic stand-in so the end card still composes.
let LOGO = null, LOGO_TRIED = false;
export async function loadLogo() {
  if (LOGO_TRIED) return LOGO;
  LOGO_TRIED = true;
  const srcs = typeof window !== 'undefined' && window.BDO_LOGO ? [window.BDO_LOGO] : ['../brand/bdo-logo.svg', '../brand/bdo-logo.png'];
  for (const src of srcs) {
    try { const im = new Image(); im.src = src; await im.decode(); if (im.naturalWidth) { LOGO = im; break; } } catch { /* try next */ }
  }
  return LOGO;
}
export const hasLogo = () => !!LOGO;
// Draw the logo with its visual centre at (cx, cy), height h px. wipe 0..1 reveals it left->right.
// Returns the drawn box {x, y, w, h}.
export function drawLogo(ctx, cx, cy, h, { alpha = 1, wipe = 1 } = {}) {
  let w;
  if (LOGO) w = (h * LOGO.naturalWidth) / LOGO.naturalHeight;
  else { ctx.save(); ctx.font = M.font(h * 1.18, 800, DISPLAY); w = ctx.measureText('BDO').width; ctx.restore(); }
  const x = cx - w / 2, y = cy - h / 2;
  if (alpha <= 0 || wipe <= 0) return { x, y, w, h };
  ctx.save(); ctx.globalAlpha *= alpha;
  ctx.beginPath(); ctx.rect(x - 4, y - h, (w + 8) * clamp(wipe), h * 3); ctx.clip();
  if (LOGO) ctx.drawImage(LOGO, x, y, w, h);
  else { ctx.font = M.font(h * 1.18, 800, DISPLAY); ctx.fillStyle = C.blue; ctx.textBaseline = 'middle'; ctx.fillText('BDO', x, cy + h * 0.04); }
  ctx.restore();
  return { x, y, w, h };
}
