// value.js: the auditor's judgement (120-155), value for the customer (155-185), the end card (185-195).
import * as K from '../lib/kit.js';
import { site, toScreen } from './world.js';

const { C, E, ss, win, prog, lerp, clamp, rgba, mix } = K;
const W = K.W, H = K.H;

export const HITS = [
  [120.6, 'rows draw on', 'whoosh'], [128.5, 'cursor', 'pop'], [132.0, 'caption', 'tick'], [136.6, 'status', 'click'],
  [140.6, 'status', 'click'], [145.2, 'rows slide away', 'whoosh'], [147.2, 'caption', 'tick'], [153.4, 'dots merge', 'pop'],
  [155.4, 'verdi revisor', 'swell'], [158.9, 'line', 'tick'], [163.2, 'red line leaves', 'whoosh'], [166.5, 'lands', 'pop'],
  [168.4, 'kunde 1', 'tick'], [171.2, 'kunde 2', 'tick'], [174.0, 'kunde 3', 'tick'], [176.8, 'kunde 4', 'tick'],
  [179.6, 'effektiv', 'click'], [181.3, 'relevant', 'click'], [182.9, 'mer verdi', 'bell'], [185.0, 'impact', 'impact'],
  [185.6, 'datamengder', 'tick'], [188.0, 'sammenhenger', 'sub'], [189.2, 'points fly', 'whoosh'], [191.1, 'wordmark', 'swell'],
  [191.8, 'red bar', 'pop'], [192.5, 'tagline', 'blip'],
];
export const sheetAt = (t) => 1;

// ------------------------------------------------------------------ helpers
const f = K.rgba;
const setFont = (ctx, size, weight, fam, track = 0) => { ctx.font = `${weight} ${size}px ${fam}`; ctx.letterSpacing = `${track * size}px`; };
const measure = (ctx, str, size, weight, fam = K.DISPLAY, track = 0) => { ctx.save(); setFont(ctx, size, weight, fam, track); const w = ctx.measureText(str).width; ctx.restore(); return w; };
function mono(ctx, str, x, y, t, t0, { size = 15, color = C.blue, track = 0.1, weight = 500, align = 'left', out = 1e9 } = {}) {
  return K.line(ctx, str, x, y, t, { size, weight, color, t0, out, align, track, family: K.MONO, dur: 0.55 });
}
function line2(ctx, x0, y0, x1, y1, p, color, w = 1.5) {
  if (p <= 0) return;
  ctx.save(); ctx.strokeStyle = color; ctx.lineWidth = w; K.drawOn(ctx, [[x0, y0], [x1, y1]], p); ctx.restore();
}
function ptAt(pts, fr) {
  const n = (pts.length - 1) * clamp(fr), k = Math.min(pts.length - 2, Math.floor(n)), q = n - k;
  return [lerp(pts[k][0], pts[k + 1][0], q), lerp(pts[k][1], pts[k + 1][1], q)];
}
function pathLen(pts) { let L = 0; for (let i = 1; i < pts.length; i++) L += Math.hypot(pts[i][0] - pts[i - 1][0], pts[i][1] - pts[i - 1][1]); return L; }
function slice(ctx, pts, d0, d1) {
  if (d1 <= d0) return;
  ctx.beginPath(); let acc = 0, started = false;
  for (let i = 1; i < pts.length; i++) {
    const a = pts[i - 1], b = pts[i], L = Math.hypot(b[0] - a[0], b[1] - a[1]);
    if (L > 0 && acc + L > d0 && acc < d1) {
      const u0 = Math.max(0, (d0 - acc) / L), u1 = Math.min(1, (d1 - acc) / L);
      if (!started) { ctx.moveTo(lerp(a[0], b[0], u0), lerp(a[1], b[1], u0)); started = true; }
      ctx.lineTo(lerp(a[0], b[0], u1), lerp(a[1], b[1], u1));
    }
    acc += L;
  }
  ctx.stroke();
}
const drift = (ctx, t, k = 1) => K.grid(ctx, 1, k, -t * 7, -t * 3.5);

// ------------------------------------------------------------------ ACT 5 data
const ROWS = [
  ['Garantikrav uten arbeidsordre', 'Garantiportal ↔ Verkstedsystem'], ['Kreditnota nær periodeslutt', 'DMS · salg'],
  ['Bruktbil · lang lagertid', 'Lagerstyring'], ['Delelager · telledifferanse', 'Delelager'],
  ['Manuell rabatt over fullmakt', 'DMS · salg'], ['Innbytte · verdivurdering', 'Kjøretøyregister'],
  ['Arbeidsordre uten faktura', 'Verkstedsystem'], ['Periodisering nybilsalg', 'Økonomisystem'],
];
const FU = [0, 2, 3];                       // rows that stay "FØLGES OPP"
const FU_Y = [250, 350, 450];
const VISIT = (i) => 128.7 + 2.05 * i;      // cursor arrives
const STAT = (i) => VISIT(i) + 0.7;         // status is set
const rowY = (i) => 250 + 80 * i;

function cursorY(t) {
  let y = 170, tp = 128.0;
  for (let i = 0; i < 8; i++) {
    const ta = VISIT(i) - 0.55;
    if (t < ta) break;
    y = lerp(i ? rowY(i - 1) : 170, rowY(i), K.spring(t - ta, { stiffness: 140, damping: 16 }));
    if (t > ta + 3) y = rowY(i);
  }
  return y;
}

function drawRow(ctx, t, i) {
  const ts = 120.4 + 0.34 * i, fi = FU.indexOf(i), fu = fi >= 0;
  const y0 = rowY(i);
  let y = y0, sc = 1, xo = 0, ga = 1, vis = 1, dotMerge = 0;
  if (fu) {
    const m = K.spring(t - (146.0 + 0.25 * fi), { stiffness: 90, damping: 15 });
    y = lerp(y0, FU_Y[fi], m); sc = 1 + 0.35 * clamp(m, 0, 1.08);
    const rv = ss(153.0 + 0.15 * fi, 154.0 + 0.15 * fi, t); vis = 1 - rv;
    dotMerge = ss(153.4, 154.4, t);
  } else {
    const ord = i - (i > 3 ? 2 : i > 2 ? 1 : i > 0 ? 0 : 0) , sl = E.inCubic(prog(t, 145.0 + 0.16 * (i % 5), 146.6 + 0.16 * (i % 5)));
    xo = -sl * 900; ga = 1 - ss(0.4, 1, sl);
    if (sl >= 1) return;
  }
  const X = 300 + xo, vt = STAT(i), seen = t >= vt;
  const expl = !fu, m = expl ? ss(vt, vt + 0.5, t) : 0;
  ctx.save(); ctx.globalAlpha *= ga;
  ctx.translate(X, y); ctx.scale(sc, sc);
  // rule that draws right from the dot
  const lp = E.outCubic(prog(t, ts, ts + 1.1));
  line2(ctx, 0, 26, 940 * vis, 26, lp, rgba(C.blue, expl && m > 0 ? 0.14 : 0.24), 1);
  if (vis > 0.01) {
    ctx.save(); ctx.beginPath(); ctx.rect(-4, -44, 1100 * vis + 8, 100); ctx.clip();
    const col = mix(C.graphite, C.light, m * 0.85);
    K.line(ctx, ROWS[i][0], 40, 10, t, { size: 30, weight: 600, color: col, t0: ts + 0.35 });
    mono(ctx, ROWS[i][1], 560, 8, t, ts + 0.7, { size: 13, color: C.blue, track: 0.04 });
    if (seen) mono(ctx, expl ? 'FORKLART' : 'FØLGES OPP', 850, 8, t, vt, { size: 15, color: expl ? C.light : C.red, weight: 600, track: 0.1 });
    if (fu && seen) {
      const w = measure(ctx, ROWS[i][0], 30, 600);
      line2(ctx, 40, 19, 40 + w, 19, E.outCubic(prog(t, vt + 0.1, vt + 0.8)), C.red, 2);
    }
    ctx.restore();
  }
  // the dot: red, or blue once explained
  const dr = lp > 0.02 || i < 8;
  if (!(i === 0 && t >= 154.4)) {
    const da = 1 - dotMerge * (i === 0 ? 0 : 1);
    if (expl) { K.redDot(ctx, 0, 0, 8, (1 - m) * da); if (m > 0) { ctx.fillStyle = rgba(C.blue, m * da); ctx.beginPath(); ctx.arc(0, 0, 8, 0, 7); ctx.fill(); } }
    else if (i === 0) K.redDot(ctx, 0, 0, 8, 1);
    else {
      ctx.save(); ctx.setTransform(1, 0, 0, 1, 0, 0);
      K.redDot(ctx, lerp(X, 300, dotMerge), lerp(y, 250, dotMerge), 8 * (1 - 0.0 * dotMerge), da); ctx.restore();
    }
  }
  ctx.restore();
}

function drawPanel(ctx, t) {
  const t0 = 128.7, out = 145.0;
  if (t < t0 - 0.1 || t > 147.6) return;
  const slide = E.inCubic(prog(t, out, 147.4)) * 520;
  ctx.save(); ctx.translate(slide, 0);
  const X = 1380, ink = C.graphite;
  line2(ctx, 1350, 200, 1350, 830, E.outCubic(prog(t, t0, t0 + 1.4)), rgba(C.blue, 0.35), 1);
  K.kicker(ctx, 'Bransjekunnskap', X, 252, t, { color: C.blue, t0: t0 + 0.2, size: 17 });
  mono(ctx, 'BDO Bilbransjeanalysen 2025', X, 286, t, t0 + 0.5, { size: 15, color: ink, weight: 400, track: 0.06 });
  const bar = (label, y, v, tt, strong, num) => {
    mono(ctx, label, X + 0, y - 18, t, tt, { size: 14, color: C.light, weight: 500, track: 0.08 });
    const len = v * 110, p = E.outCubic(prog(t, tt + 0.2, tt + 1.1));
    line2(ctx, X + 60, y, X + 60 + len, y, p, strong ? C.blue : C.light, strong ? 3 : 2);
    if (p > 0.98) line2(ctx, X + 60 + len, y - 7, X + 60 + len, y + 7, 1, strong ? C.blue : C.light, 1.5);
    K.line(ctx, num, X + 60 + len + 16, y + 10, t, { size: 30, weight: 600, color: ink, t0: tt + 0.5 });
  };
  mono(ctx, 'Driftsmargin merkeforhandlere', X, 360, t, t0 + 1.0, { size: 15, color: ink, weight: 500, track: 0.04 });
  const yr = (s, y, tt) => mono(ctx, s, X, y + 5, t, tt, { size: 14, color: C.light, track: 0.08 });
  yr('2023', 410, t0 + 1.3); line2(ctx, X + 60, 410, X + 60 + 1.5 * 110, 410, E.outCubic(prog(t, t0 + 1.4, t0 + 2.3)), C.light, 2);
  K.line(ctx, '1,5 %', X + 60 + 165 + 16, 421, t, { size: 30, weight: 600, color: ink, t0: t0 + 1.9 });
  yr('2024', 462, t0 + 1.7); line2(ctx, X + 60, 462, X + 60 + 1.7 * 110, 462, E.outCubic(prog(t, t0 + 1.8, t0 + 2.7)), C.blue, 3);
  K.line(ctx, '1,7 %', X + 60 + 187 + 16, 473, t, { size: 30, weight: 600, color: ink, t0: t0 + 2.3 });
  mono(ctx, 'Bruktbilforhandlere', X, 570, t, t0 + 3.0, { size: 15, color: ink, weight: 500, track: 0.04 });
  yr('2024', 620, t0 + 3.3); line2(ctx, X + 60, 620, X + 60 + 2.5 * 110, 620, E.outCubic(prog(t, t0 + 3.4, t0 + 4.3)), C.blue, 3);
  K.line(ctx, '2,5 %', X + 60 + 275 + 16, 631, t, { size: 30, weight: 600, color: ink, t0: t0 + 3.9 });
  K.line(ctx, 'Over 1 300', X, 750, t, { size: 56, weight: 300, color: ink, t0: t0 + 5.0 });
  mono(ctx, 'bilbransjebedrifter analysert', X, 790, t, t0 + 5.5, { size: 15, color: C.blue, weight: 500, track: 0.06 });
  ctx.restore();
}

function act5(ctx, t) {
  drift(ctx, t, ss(120.2, 122.0, t));
  K.kicker(ctx, 'Revisors vurdering', 300, 150, t, { color: C.blue, t0: 120.5, out: 154 });
  for (let i = 0; i < 8; i++) drawRow(ctx, t, i);
  drawPanel(ctx, t);
  // cursor ring
  const ca = ss(128.0, 128.5, t) * (1 - ss(144.6, 145.2, t));
  if (ca > 0) {
    const y = cursorY(t); ctx.save(); ctx.globalAlpha = ca; ctx.strokeStyle = C.graphite; ctx.lineWidth = 1.5;
    ctx.beginPath(); ctx.arc(300, y, 17 - 2 * ss(0, 1, Math.sin(t * 5) * 0.5 + 0.5), 0, 7); ctx.stroke(); ctx.restore();
  }
  if (t >= 154.4) K.redDot(ctx, 300, 250, 8, 1);
  // captions
  const cs = { size: 64, weight: 600, color: C.graphite };
  K.line(ctx, 'Teknologien viser hvor vi bør se.', K.CAP.x, 930, t, { ...cs, t0: 132, out: 138.2 });
  K.line(ctx, 'Revisor vurderer hva det betyr.', K.CAP.x, 930, t, { ...cs, t0: 138.5, out: 144.7 });
  K.line(ctx, 'FRA FORSTÅELSE', K.CAP.x, 860, t, { ...cs, t0: 147, out: 153.6 });
  K.line(ctx, 'TIL MÅLRETTET OPPFØLGING.', K.CAP.x, 930, t, { ...cs, color: C.blue, t0: 147.5, out: 153.6 });
}

// ------------------------------------------------------------------ ACT 6/7 camera + route
const WP = [[-40, -35], [-9.5, -35], [28.5, -35], [47.5, -35]];
const REV = ['Bedre oversikt over datagrunnlaget', 'Mer målrettede analyser og kontroller', 'Bedre grunnlag for risikovurdering', 'Mer tid til faglig vurdering'];
const CUST = [
  { s: 'Bedre oversikt over egen virksomhet', x: 120, y: 700, al: 'left', a: 0, t0: 168.2 },
  { s: 'Mer relevant oppfølging', x: 740, y: 500, al: 'left', a: 1, t0: 171.0 },
  { s: 'Tydeligere dialog om funn', x: 1420, y: 580, al: 'right', a: 2, t0: 173.8 },
  { s: 'Innsikt i utvikling og sammenhenger', x: 1150, y: 800, al: 'left', a: 3, t0: 176.6 },
];
function cam(t) {
  const pb = E.inOutCubic(prog(t, 185, 188.8));
  const sPush = 9 + 3.4 * E.inOutSine(prog(t, 163, 185));
  const s = lerp(sPush, 10.5, pb);
  const slide = 1 - E.outCubic(prog(t, 163.2, 167.2));
  const dr = Math.sin((t - 163) * 0.16) * 2.5 * (1 - pb) + (t > 189 ? Math.sin(t * 0.5) * 0.25 : 0);
  const cy = lerp(-8 + 7 * slide, 0, pb);
  const alpha = t < 179 ? 0.9 * ss(164.6, 167.4, t) : t < 185 ? lerp(0.9, 0.14, ss(179, 180.6, t)) : lerp(0.14, 0.3, ss(185, 187, t)) * (1 - ss(189, 191, t));
  return { s, cx: dr, cy, alpha };
}

// wordmark geometry + sampled letter points (lazy: fonts are loaded by the time we draw)
let WM = null;
function wordmark(ctx) {
  if (WM) return WM;
  const size = 300, base = 520;
  ctx.save(); setFont(ctx, size, 800, K.DISPLAY); const w = ctx.measureText('BDO').width; const mt = ctx.measureText('BDO'); ctx.restore();
  const asc = mt.actualBoundingBoxAscent || size * 0.72, x0 = W / 2 - w / 2;
  const pts = [];
  try {
    const cv = document.createElement('canvas'); cv.width = W; cv.height = 700;
    const c = cv.getContext('2d'); c.font = `800 ${size}px ${K.DISPLAY}`; c.fillStyle = '#000'; c.textBaseline = 'alphabetic'; c.fillText('BDO', x0, base);
    const d = c.getImageData(0, 0, W, 700).data;
    for (let y = 0; y < 700; y += 3) for (let x = Math.floor(x0) - 5; x < x0 + w + 5; x += 3) if (d[(y * W + x) * 4 + 3] > 128) pts.push([x, y]);
  } catch (e) { /* fall back to a block */ }
  if (!pts.length) for (let i = 0; i < 4000; i++) pts.push([x0 + K.hash01(i, 3) * w, base - K.hash01(i, 4) * asc]);
  const r = K.mulberry32(901), tg = [];
  for (let i = 0; i < K.NPTS; i++) { const p = pts[Math.floor(r() * pts.length)]; tg.push([p[0] + (r() - 0.5) * 3, p[1] + (r() - 0.5) * 3]); }
  WM = { size, base, w, asc, x0, tg };
  return WM;
}

function backing(ctx, x, y, w, size, t, t0, out) {
  const p = E.outCubic(prog(t, t0 - 0.05, t0 + 0.55)), po = ss(out - 0.3, out, t);
  if (p <= 0 || po >= 1) return;
  ctx.save(); ctx.globalAlpha *= 0.86 * (1 - po); ctx.fillStyle = C.arctic;
  ctx.fillRect(x - 26, y - size * 0.98 - 10, (w + 52) * p, size * 1.5 + 14); ctx.restore();
}

function act6(ctx, t) {
  const cm = cam(t), ink = C.graphite;
  // ---- revisor value (155-166)
  if (t < 167) {
    K.kicker(ctx, 'Verdi for revisor', 340, 258, t, { color: C.blue, t0: 155.0, out: 163.4, tick: false });
    for (let i = 0; i < 4; i++) {
      const t0 = 155.9 + 1.5 * i, sl = E.inCubic(prog(t, 163.3 + 0.13 * i, 165.0 + 0.13 * i));
      if (sl >= 1) continue;
      ctx.save(); ctx.translate(-sl * 1300, 0);
      const y = 390 + 100 * i;
      line2(ctx, 300, y - 14, 340, y - 14, E.outCubic(prog(t, t0 - 0.1, t0 + 0.5)), C.blue, 2);
      K.line(ctx, REV[i], 360, y, t, { size: 40, weight: 600, color: ink, t0: t0 + 0.2 });
      ctx.restore();
    }
  }
  // ---- the customer's business
  if (t < 185 && cm.alpha > 0.002) site(ctx, { k: 1, cx: cm.cx, cy: cm.cy, s: cm.s, alpha: cm.alpha });
}

function route(t) {
  const cm = cam(t), P = (x, y) => toScreen(cm.cx, cm.cy, cm.s, x, y);
  const S = [300, 250], E0 = P(...WP[0]);
  const arc = K.bez([S, [1500, 90, 1760, 860, E0[0], E0[1]]], 44);
  const pieces = [arc];
  for (let i = 1; i < 4; i++) pieces.push([P(...WP[i - 1]), P(...WP[i])]);
  const m = WM, bx = m ? m.x0 - 40 : 560, top = m ? m.base - m.asc : 304;
  const last = P(...WP[3]);
  pieces.push(K.bez([last, [last[0] + 300, last[1] + 500, bx - 220, 140, bx, top - 70]], 34).concat([[bx, top]]));
  return { pieces, bx, top, bot: m ? m.base : 520 };
}

function redLine(ctx, t) {
  if (t < 163) return;
  const R = route(t), n = R.pieces.length;
  const fr = [E.inOutCubic(prog(t, 163.1, 166.4)), E.inOutCubic(prog(t, 170.0, 171.0)), E.inOutCubic(prog(t, 172.8, 173.8)), E.inOutCubic(prog(t, 175.6, 176.6)), E.inOutCubic(prog(t, 189.0, 190.8))];
  // assemble polyline and head distance
  const all = []; let head = 0, tot = 0;
  R.pieces.forEach((p, i) => { const L = pathLen(p); if (i) { all.push(...p); } else all.push(...p); head += L * fr[i]; tot += L; });
  const barLen = R.bot - R.top, la = pathLen(R.pieces[0]), tail = la * ss(167.2, 168.9, t) + (tot - la) * ss(189.5, 191.5, t);
  const lw = 2;
  ctx.save(); ctx.strokeStyle = C.red; ctx.lineWidth = lw; ctx.lineJoin = 'round'; ctx.lineCap = 'round';
  ctx.globalAlpha = t < 179 ? 1 : lerp(1, 0.7, ss(179, 181, t));
  slice(ctx, all, tail, head); ctx.restore();
  // head position
  let d = head, hp = all[0]; for (let i = 1; i < all.length; i++) { const a = all[i - 1], b = all[i], L = Math.hypot(b[0] - a[0], b[1] - a[1]); if (d <= L) { hp = [lerp(a[0], b[0], d / (L || 1)), lerp(a[1], b[1], d / (L || 1))]; break; } d -= L; hp = b; }
  const fb = E.inOutCubic(prog(t, 190.8, 191.5));
  if (fb > 0) { ctx.fillStyle = C.red; ctx.fillRect(R.bx - 3, R.top, 6, barLen * fb); hp = [R.bx, lerp(R.top, R.bot, fb)]; }
  const dotA = 1 - ss(191.4, 191.8, t);
  if (dotA > 0) K.redDot(ctx, hp[0], hp[1], 8, dotA);
  if (t >= 191.5) { ctx.fillStyle = C.red; ctx.fillRect(R.bx - 3, R.top, 6, barLen); }
  return R;
}

function customer(ctx, t) {
  const cm = cam(t), P = (x, y) => toScreen(cm.cx, cm.cy, cm.s, x, y), ink = C.graphite;
  if (t < 166.6 || t > 179.4) return;
  K.kicker(ctx, 'Verdi for kunden', 120, 1000, t, { color: C.blue, t0: 167.0, out: 179.0 });
  const kw = 330; backing(ctx, 120, 1000, kw + 40, 22, t, 167.0, 179.0);
  K.kicker(ctx, 'Verdi for kunden', 120, 1000, t, { color: C.blue, t0: 167.0, out: 179.0 });
  for (const c of CUST) {
    const w = measure(ctx, c.s, 40, 600), x = c.al === 'right' ? c.x - w : c.x;
    const out = 179.2, an = P(...WP[c.a]);
    backing(ctx, x, c.y, w, 40, t, c.t0, out);
    const top = c.y - 40 * 0.98 - 10, ex = clamp(an[0], x + 12, x + w - 12);
    const lp = E.outCubic(prog(t, c.t0 - 0.2, c.t0 + 0.6)) * (1 - ss(out - 0.3, out, t));
    line2(ctx, an[0], an[1], ex, top, lp, C.blue, 1.2);
    if (lp > 0) { ctx.strokeStyle = C.blue; ctx.lineWidth = 1.2; ctx.beginPath(); ctx.arc(an[0], an[1], 5, 0, 7); ctx.stroke(); }
    K.line(ctx, c.s, x, c.y, t, { size: 40, weight: 600, color: ink, t0: c.t0 + 0.1, out });
  }
}

function statements(ctx, t) {
  K.line(ctx, 'Effektiv revisjon.', 120, 620, t, { size: 84, weight: 600, color: C.graphite, t0: 179.6, out: 184.6 });
  K.line(ctx, 'Relevant innsikt.', 120, 740, t, { size: 84, weight: 600, color: C.graphite, t0: 181.3, out: 184.6 });
  K.line(ctx, 'Mer verdi.', 120, 860, t, { size: 84, weight: 600, color: C.blue, t0: 182.9, out: 184.6 });
}

function act7(ctx, t) {
  const m = wordmark(ctx);
  // field: appears in network layout, then flies into the wordmark
  const fl = (p) => {
    const d = 185.1 + p.seed * 2.4, e = E.outCubic(prog(t, d, d + 1.1));
    let x = p.x, y = p.y + (1 - e) * 70, a = e, r = e;
    if (t >= 189) {
      const dd = 189 + p.seed * 1.4, u = K.spring(t - dd, { stiffness: 55, damping: 11 }), g = m.tg[p.i];
      const sw = Math.sin(Math.PI * clamp((t - dd) / 1.6)) * 60 * (p.z - 0.5);
      x = lerp(x, g[0], u) + sw; y = lerp(y, g[1], u) - sw * 0.5; r = lerp(1, 0.9, ss(0, 1, u));
      a = 1 - ss(192.3, 193.3, t);
    }
    return [x, y, a, r];
  };
  if (t > 185.0) K.field(ctx, 1, { alpha: 0.95, pos: fl, links: (1 - ss(189, 190.4, t)) * ss(185.2, 186.5, t), red: 0 });
  // three lines
  const L = [['STORE DATAMENGDER.', 185.7], ['TYDELIGE SAMMENHENGER.', 186.8], ['RELEVANT INNSIKT.', 187.9]];
  L.forEach(([s, t0], i) => {
    const y = 840 + 74 * i, w = measure(ctx, s, 44, 600, K.DISPLAY, 0.06);
    backing(ctx, 120, y, w, 44, t, t0, 189.1);
    K.line(ctx, s, 120, y, t, { size: 44, weight: 600, color: C.graphite, t0, out: 189.1, track: 0.06 });
  });
  // solid wordmark wipe
  if (t >= 191) {
    const dx = t > 193 ? Math.sin(t * 0.9) * 1.5 : 0, dy = t > 193 ? Math.cos(t * 0.7) * 1 : 0;
    const wp = E.inOutCubic(prog(t, 191.0, 192.7));
    ctx.save(); ctx.translate(dx, dy); ctx.beginPath(); ctx.rect(m.x0 - 40, m.base - m.asc - 30, (m.w + 80) * wp, m.asc + 80); ctx.clip();
    setFont(ctx, m.size, 800, K.DISPLAY); ctx.fillStyle = C.blue; ctx.textBaseline = 'alphabetic'; ctx.fillText('BDO', m.x0, m.base); ctx.restore();
  }
  K.line(ctx, 'Vi ser detaljene. Vi forstår helheten.', W / 2, 650, t, { size: 44, weight: 300, color: C.graphite, t0: 191.6, align: 'center' });
  K.line(ctx, 'TETT PÅ BILBRANSJEN', W / 2, 712, t, { size: 22, weight: 500, color: C.blue, t0: 192.4, align: 'center', track: 0.28, family: K.MONO });
}

export function draw(ctx, t) {
  K.sheet(ctx, 1);
  ctx.save(); ctx.lineCap = 'butt';
  if (t < 155) act5(ctx, t);
  else { drift(ctx, t); act6(ctx, t); }
  if (t >= 155) {
    if (t < 185) { customer(ctx, t); statements(ctx, t); }
  }
  if (t >= 185) { drift(ctx, t, 0); }
  if (t >= 185) { const cm = cam(t); if (cm.alpha > 0.002) site(ctx, { k: 1, cx: cm.cx, cy: cm.cy, s: cm.s, alpha: cm.alpha }); }
  if (t >= 155 && t < 163) K.redDot(ctx, 300, 250, 8, 1);
  if (t >= 163) {
    if (t >= 185) wordmark(ctx);
    if (t >= 185) act7(ctx, t);
    redLine(ctx, t);
  }
  ctx.restore();
}
