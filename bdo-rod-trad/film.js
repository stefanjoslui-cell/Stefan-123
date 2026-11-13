// BDO Automotive – Den røde tråden. A pure function of time: window.seek(t) paints frame t.
// ONE continuous shot: a single red line is drawn by hand across one enormous bright drawing sheet.
// The camera tracks the tip of the line (look-ahead, log-space zoom); everything the line passes is
// drawn on just before the tip arrives. All chapters live in one world coordinate system ("u";
// 1 u = 1 m of the dealership site plan). Screen = centre + (world - cam) * s.
import * as M from './lib/motion.js';
import * as K from './lib/kit.js';
import * as world from './scenes/world.js';

const { W, H, C, clamp, lerp, prog, E, ss, rgba, win } = K;
const PERIOD = 0.625;
const G = C.graphite, BL = C.blue, LB = C.light, CL = C.cloud, RED = C.red;
const MONO = K.MONO, DISPLAY = K.DISPLAY;

// ------------------------------------------------------------------ monotone cubic (PCHIP) for keyed curves
function pchip(keys, startSecant = false) {
  const n = keys.length, h = [], d = [], m = new Array(n).fill(0);
  for (let i = 0; i < n - 1; i++) { h[i] = keys[i + 1][0] - keys[i][0]; d[i] = (keys[i + 1][1] - keys[i][1]) / h[i]; }
  for (let i = 1; i < n - 1; i++) {
    if (d[i - 1] * d[i] <= 0) m[i] = 0;
    else { const w1 = 2 * h[i] + h[i - 1], w2 = h[i] + 2 * h[i - 1]; m[i] = (w1 + w2) / (w1 / d[i - 1] + w2 / d[i]); }
  }
  if (startSecant) m[0] = d[0];
  return (x) => {
    if (x <= keys[0][0]) return keys[0][1];
    if (x >= keys[n - 1][0]) return keys[n - 1][1];
    let i = 0; while (x > keys[i + 1][0]) i++;
    const t = (x - keys[i][0]) / h[i], t2 = t * t, t3 = t2 * t;
    return (2 * t3 - 3 * t2 + 1) * keys[i][1] + (t3 - 2 * t2 + t) * h[i] * m[i] + (-2 * t3 + 3 * t2) * keys[i + 1][1] + (t3 - t2) * h[i] * m[i + 1];
  };
}

// ------------------------------------------------------------------ the thread: one polyline, arc-length parametrised
const OX = 145;                               // the site plan sits at world x = OX + site x
const WPS = [], MK = {};
const wp = (x, y, name) => { if (name) MK[name] = WPS.length; WPS.push([x, y]); };
const sp = (x, y, name) => wp(x + OX, y, name);
// chapter 1-2: ledger line, then the car's history line
wp(-0.3, 0.14, 'start'); wp(16.2, 0.14, 'ledgerEnd'); wp(23.2, 0.14, 'carIn'); wp(28.7, 0.14, 'carOut'); wp(33, 0.14); wp(42, 3); wp(50, 11);
// chapter 3: between the rows of the car park, into the building, through six departments
sp(-90, 13.35, 'siteIn'); sp(-84, 13.35); sp(-70, 13.35); sp(-63, 13.3); sp(-58.2, 11.6); sp(-55.6, 7.5); sp(-55.15, 4, 'lane0');
sp(-56.3, -1); sp(-54, -8); sp(-56, -14); sp(-54.5, -19.5); sp(-52.5, -24.2, 'door'); sp(-52, -29); sp(-50.5, -34); sp(-47.5, -36.8, 'd0');
for (let i = 1; i <= 5; i++) { sp(-57 + 19 * i, -27.3, 'x' + i); sp(-47.5 + 19 * i, -36.8, 'd' + i); }
sp(57, -27.3, 'exitB'); sp(63, -26.8);
// chapter 4: braid, platform (three plates), industry-knowledge column
const PX = 380, PY = -15;
wp(215, -26); wp(238, -21); wp(262, -11); wp(284, -3); wp(300, 4, 'pEnter'); wp(322, 5); wp(345, 4, 'pc0'); wp(363, -5); wp(380, -15, 'pc1'); wp(397, -25); wp(415, -34, 'pc2'); wp(433, -33); wp(452, -33, 'pExit');
const KX = 496;
wp(472, -48); wp(486, -60); wp(493, -65, 'kTop'); wp(496, -62); wp(496, -56, 'rule1'); wp(496, -29, 'rule2'); wp(496, -2, 'rule3'); wp(496, 22, 'kEnd'); wp(499, 30); wp(510, 37); wp(528, 38.5, 'c0');
// chapter 5: along the timeline chart, looping round three points
function loop(cx, cy, r, name) { for (let k = 0; k <= 12; k++) { const a = Math.PI / 2 - (k * 2 * Math.PI) / 12; wp(cx + r * Math.cos(a), cy + r * Math.sin(a), k === 0 ? name : undefined); } }
const CH = { x0: 540, mw: 200 / 12, lanes: [20, 28.5, 37, 45.5, 54], names: ['SALG', 'LAGER', 'VERKSTED', 'GARANTI', 'KREDITNOTA'] };
const FLAGS = [
  { name: 'p1', x: CH.x0 + 4.5 * CH.mw, y: CH.lanes[3], r: 3.2, text: 'Garantikrav uten arbeidsordre' },
  { name: 'p2', x: CH.x0 + 8.6 * CH.mw, y: CH.lanes[1], r: 3.2, text: 'Bruktbil · lang lagertid' },
  { name: 'p3', x: CH.x0 + 11.3 * CH.mw, y: CH.lanes[4], r: 3.2, text: 'Kreditnota nær periodeslutt' },
];
wp(552, 34); wp(575, 40); wp(595, 46.5); wp(606, 48.7);
loop(FLAGS[0].x, FLAGS[0].y, FLAGS[0].r, 'p1');
wp(632, 47.5); wp(646, 40); wp(660, 34); wp(672, 32); wp(678.5, FLAGS[1].y + 3.2);
loop(FLAGS[1].x, FLAGS[1].y, FLAGS[1].r, 'p2');
wp(700, 34); wp(708, 46); wp(713, 56); wp(717, 57.2); wp(722, 57.2);
loop(FLAGS[2].x, FLAGS[2].y, FLAGS[2].r, 'p3');
wp(740, 57.2); wp(752, 55); wp(768, 48); wp(779, 40); wp(786, 31, 'cEnd'); wp(796, 24); wp(812, 20.5); wp(824, 20.8); wp(830, 21.8, 'conv0'); wp(833, 26); wp(832.4, 31.5); wp(829.6, 36.6);
// chapter 6: stitching between the two columns
const MX = 830, ROWS = [40, 47, 54];
wp(826, ROWS[0], 'sL0'); wp(830, ROWS[0] + 0.3); wp(834, ROWS[0], 'sR0'); wp(831.5, 43.5); wp(827.5, 45); wp(826, ROWS[1], 'sL1'); wp(830, ROWS[1] + 0.3); wp(834, ROWS[1], 'sR1');
wp(831.5, 50.5); wp(827.5, 52); wp(826, ROWS[2], 'sL2'); wp(830, ROWS[2] + 0.3); wp(834, ROWS[2], 'sR2');
// chapter 7: out to the red point beside the logo
wp(838, 57.5); wp(850, 63, 'exitC'); wp(872, 64.5); wp(898, 57); wp(922, 50.5); const END = [945, 48]; wp(END[0], END[1], 'end');

// centripetal Catmull-Rom through waypoints; returns the polyline and the polyline index of every waypoint
function crPoly(W_, step = 0.3) {
  const n = W_.length, out = [W_[0].slice()], idx = [0];
  const refl = (a, b) => [2 * a[0] - b[0], 2 * a[1] - b[1]];
  const dd = (a, b) => Math.max(1e-4, Math.hypot(b[0] - a[0], b[1] - a[1]) ** 0.5);
  for (let i = 0; i < n - 1; i++) {
    const p1 = W_[i], p2 = W_[i + 1], p0 = i === 0 ? refl(p1, p2) : W_[i - 1], p3 = i + 2 >= n ? refl(p2, p1) : W_[i + 2];
    const len = Math.hypot(p2[0] - p1[0], p2[1] - p1[1]), cnt = Math.max(1, Math.ceil(len / step));
    const t0 = 0, t1 = t0 + dd(p0, p1), t2 = t1 + dd(p1, p2), t3 = t2 + dd(p2, p3);
    for (let q = 1; q <= cnt; q++) {
      const t = t1 + ((t2 - t1) * q) / cnt, o = [0, 0];
      for (let c = 0; c < 2; c++) {
        const a1 = ((t1 - t) / (t1 - t0)) * p0[c] + ((t - t0) / (t1 - t0)) * p1[c];
        const a2 = ((t2 - t) / (t2 - t1)) * p1[c] + ((t - t1) / (t2 - t1)) * p2[c];
        const a3 = ((t3 - t) / (t3 - t2)) * p2[c] + ((t - t2) / (t3 - t2)) * p3[c];
        const b1 = ((t2 - t) / (t2 - t0)) * a1 + ((t - t0) / (t2 - t0)) * a2;
        const b2 = ((t3 - t) / (t3 - t1)) * a2 + ((t - t1) / (t3 - t1)) * a3;
        o[c] = ((t2 - t) / (t2 - t1)) * b1 + ((t - t1) / (t2 - t1)) * b2;
      }
      out.push(o);
    }
    idx.push(out.length - 1);
  }
  return { pts: out, idx };
}
const CRT = crPoly(WPS, 0.3), TH = CRT.pts, WPI = CRT.idx, CL_ = [0];
for (let i = 1; i < TH.length; i++) CL_.push(CL_[i - 1] + Math.hypot(TH[i][0] - TH[i - 1][0], TH[i][1] - TH[i - 1][1]));
const TOTAL = CL_[CL_.length - 1];
const ML = (name) => CL_[WPI[MK[name]]];                      // arclength of a named waypoint
function idxAt(L) { let lo = 0, hi = CL_.length - 1; while (hi - lo > 1) { const m = (lo + hi) >> 1; if (CL_[m] <= L) lo = m; else hi = m; } return lo; }
function posAt(L, off = 0) {
  L = clamp(L, 0, TOTAL); const i = idxAt(L), f = (L - CL_[i]) / Math.max(1e-9, CL_[i + 1] - CL_[i]);
  const a = TH[i], b = TH[Math.min(i + 1, TH.length - 1)];
  let x = lerp(a[0], b[0], f), y = lerp(a[1], b[1], f);
  if (off) { const dx = b[0] - a[0], dy = b[1] - a[1], l = Math.hypot(dx, dy) || 1; x += (-dy / l) * off; y += (dx / l) * off; }
  return [x, y];
}

// tip arclength as a function of time: keyed at named waypoints
const TIMES = [['ledgerEnd', 5.0], ['carIn', 8.4], ['carOut', 14.8], ['siteIn', 16.4], ['lane0', 18.2], ['door', 20.4], ['d0', 21.5], ['d1', 22.75], ['d2', 24.0], ['d3', 25.25], ['d4', 26.5], ['d5', 27.75],
  ['exitB', 29.0], ['pEnter', 32.6], ['pc0', 33.8], ['pc1', 34.9], ['pc2', 36.0], ['pExit', 37.2], ['kTop', 38.3], ['rule1', 39.2], ['rule2', 41.0], ['rule3', 42.8], ['kEnd', 44.0], ['c0', 45.8],
  ['p1', 48.8], ['p2', 52.0], ['p3', 55.0], ['cEnd', 57.7], ['conv0', 59.2], ['sL0', 60.0], ['sR0', 60.7], ['sL1', 61.6], ['sR1', 62.3], ['sL2', 63.2], ['sR2', 63.9], ['exitC', 65.2], ['end', 70.3]];
const LKEYS = [[0, ML('ledgerEnd') - 10.7], ...TIMES.map(([n, t]) => [t, ML(n)])];
const Lof = pchip(LKEYS, true);
const tipPos = (t) => posAt(Lof(t));
const tipT = (L) => { let lo = 0, hi = 80; for (let i = 0; i < 40; i++) { const m = (lo + hi) / 2; if (Lof(m) < L) lo = m; else hi = m; } return hi; };
const markT = (name) => tipT(ML(name));
// time at which the tip first reaches world x (running maximum), for left-to-right reveals
const XT = (() => { const o = []; let mx = -1e9; for (let t = 0; t <= 75; t += 0.05) { mx = Math.max(mx, tipPos(t)[0]); o.push([t, mx]); } return o; })();
const timeAtX = (x) => { for (let i = 0; i < XT.length; i++) if (XT[i][1] >= x) return XT[i][0]; return 75; };

// ------------------------------------------------------------------ camera: follows the tip, log-space zoom
const lnS = pchip([[0, 100], [4, 106], [7, 140], [9, 230], [14.4, 240], [18.4, 15], [20, 15], [21.8, 30], [28.2, 30], [31.4, 10.3], [33, 10.8], [35, 12.8], [37, 13], [38.5, 14], [42.5, 14],
  [44, 12.5], [56.5, 12.5], [58.4, 17], [63, 17.4], [66, 17]].map(([t, s]) => [t, Math.log(s)]));
const txK = pchip([[0, 960], [28, 960], [29.5, 1380], [32, 1420], [34, 960], [60, 960]]);
const tyK = pchip([[0, 430], [8, 380], [10, 300], [14, 300], [16.5, 520], [19.8, 520], [21.5, 600], [28, 600], [30, 520], [34, 440], [38, 440], [41, 520], [46, 380], [56, 380], [58, 400], [62, 380], [66, 400]]);
function follow(t) {
  let fx = 0, fy = 0; const N = 9;
  for (let i = 0; i < N; i++) { const p = tipPos(t + 0.15 + (0.85 * i) / (N - 1)); fx += p[0]; fy += p[1]; }
  fx /= N; fy /= N;
  const s = Math.exp(lnS(t));
  return { cx: fx - (txK(t) - 960) / s, cy: fy - (tyK(t) - 540) / s, s };
}
function zoomBlend(A, B, w) {
  const s = Math.exp(lerp(Math.log(A.s), Math.log(B.s), w)), dn = 1 / B.s - 1 / A.s, k = Math.abs(dn) < 1e-9 ? w : (1 / s - 1 / A.s) / dn;
  return { cx: lerp(A.cx, B.cx, k), cy: lerp(A.cy, B.cy, k), s };
}
const OVER = { cx: 485, cy: -4, s: 1.82 };
const SHEET = { x0: -20, y0: -92, x1: 990, y1: 100 };
let cam = null, LOGOW = 400;
const finalCam = () => { const s = 40, dotX = 960 - LOGOW / 2 - 64; return { s, cx: END[0] - (dotX - 960) / s, cy: END[1] - (430 - 540) / s }; };
const MC = { x: pchip([[27.6, 150], [29.5, 190], [31.2, 215], [33.8, 365], [37.5, 395]]), y: pchip([[27.6, -30], [31.2, -30], [33.8, -5.5], [37.5, -5.5]]),
  s: pchip([[29, 10.3], [31.4, 10.3], [33.8, 11.6], [37.5, 11.6]].map(([t, s]) => [t, Math.log(s)])) };
function camera(t) {
  let c = follow(t);
  const wm = ss(27.6, 30.0, t) * (1 - ss(37.0, 38.8, t));
  if (wm > 0) c = zoomBlend(c, { cx: MC.x(t), cy: MC.y(t), s: Math.exp(MC.s(t)) }, wm);
  const wk = ss(37.6, 38.8, t) * (1 - ss(43.6, 44.8, t)), wy = ss(43.6, 45, t) * (1 - ss(57.6, 58.8, t)), wc = ss(58.0, 59.6, t) * (1 - ss(64.6, 65.8, t));
  c = { ...c, cx: lerp(lerp(c.cx, 524, wk), 830, wc), cy: lerp(lerp(c.cy, 46, wy), 51.5, wc) };
  const w1 = E.inOutCubic(prog(t, 66.0, 68.4)), w2 = E.inOutCubic(prog(t, 69.3, 72.3));
  if (w1 > 0) c = zoomBlend(c, OVER, w1);
  if (w2 > 0) c = zoomBlend(OVER, finalCam(), w2);
  return c;
}
const kz = (ref) => clamp(cam.s / ref);
const P = (x, y) => [W / 2 + (x - cam.cx) * cam.s, H / 2 + (y - cam.cy) * cam.s];
const visible = (x0, y0, x1, y1, m = 0) => !(P(x1, y1)[0] < -m || P(x0, y0)[0] > W + m || P(x1, y1)[1] < -m || P(x0, y0)[1] > H + m);

// ------------------------------------------------------------------ text helpers (screen px)
function mono(ctx, str, x, y, size, color, { align = 'left', track = 0.1, weight = 500, halo = 0 } = {}) {
  ctx.save(); ctx.font = M.font(size, weight, MONO); ctx.letterSpacing = `${track * size}px`; ctx.textAlign = align; ctx.textBaseline = 'alphabetic';
  if (halo) { ctx.lineWidth = halo; ctx.strokeStyle = rgba(C.arctic, 0.95); ctx.lineJoin = 'round'; ctx.strokeText(str, x, y); }
  ctx.fillStyle = color; ctx.fillText(str, x, y); ctx.restore();
}
// a label that rises out of its own baseline mask (with a paper-coloured halo for legibility on drawings)
function tag(ctx, str, x, y, t, t0, out, { size = 24, color = G, weight = 500, family = MONO, halo = 7 } = {}) {
  const pi = E.outQuint(prog(t, t0, t0 + 0.55)), po = E.inCubic(prog(t, out - 0.4, out));
  if (pi <= 0 || po >= 1) return 0;
  ctx.save(); ctx.font = M.font(size, weight, family); ctx.letterSpacing = family === MONO ? `${0.04 * size}px` : '0px'; ctx.textBaseline = 'alphabetic';
  const w = ctx.measureText(str).width;
  ctx.beginPath(); ctx.rect(x - 8, y - size * 1.0, w + 16, size * 1.35); ctx.clip();
  ctx.globalAlpha *= 1 - po; const yy = y + (1 - pi) * size * 1.1 - po * size * 0.4;
  ctx.lineWidth = halo; ctx.strokeStyle = rgba(C.arctic, 0.96); ctx.lineJoin = 'round'; ctx.strokeText(str, x, yy);
  ctx.fillStyle = color; ctx.fillText(str, x, yy); ctx.restore();
  return w;
}

// ------------------------------------------------------------------ paper, grid, sheet
function paper(ctx) {
  ctx.fillStyle = M.mix(C.arctic, CL, 0.55); ctx.fillRect(0, 0, W, H);
  const [x0, y0] = P(SHEET.x0, SHEET.y0), [x1, y1] = P(SHEET.x1, SHEET.y1);
  ctx.save(); ctx.shadowColor = rgba(G, 0.14); ctx.shadowBlur = 24; ctx.shadowOffsetY = 6;
  ctx.fillStyle = C.arctic; ctx.fillRect(x0, y0, x1 - x0, y1 - y0); ctx.restore();
  ctx.fillStyle = C.arctic; ctx.fillRect(x0, y0, x1 - x0, y1 - y0);
  const ov = clamp(1 - cam.s / 6);
  if (ov > 0) {
    ctx.strokeStyle = rgba(G, 0.35 * ov); ctx.lineWidth = 1; ctx.strokeRect(x0 + 10, y0 + 10, x1 - x0 - 20, y1 - y0 - 20);
    ctx.strokeStyle = rgba(G, 0.25 * ov); ctx.strokeRect(x0, y0, x1 - x0, y1 - y0);
  }
}
function worldGrid(ctx) {
  const vx0 = cam.cx - W / 2 / cam.s, vx1 = cam.cx + W / 2 / cam.s, vy0 = cam.cy - H / 2 / cam.s, vy1 = cam.cy + H / 2 / cam.s;
  ctx.lineWidth = 1;
  for (const step of [1, 10, 100]) {
    const px = step * cam.s, a = ss(30, 80, px) * (1 - ss(300, 900, px)); if (a < 0.01) continue;
    ctx.strokeStyle = rgba(BL, 0.07 * a); ctx.beginPath();
    const ax = Math.max(vx0, SHEET.x0), bx = Math.min(vx1, SHEET.x1), ay = Math.max(vy0, SHEET.y0), by = Math.min(vy1, SHEET.y1);
    for (let x = Math.ceil(ax / step) * step; x <= bx; x += step) { const [sx, sy0] = P(x, ay), sy1 = P(x, by)[1]; ctx.moveTo(Math.round(sx) + 0.5, sy0); ctx.lineTo(Math.round(sx) + 0.5, sy1); }
    for (let y = Math.ceil(ay / step) * step; y <= by; y += step) { const [sx0, sy] = P(ax, y), sx1 = P(bx, y)[0]; ctx.moveTo(sx0, Math.round(sy) + 0.5); ctx.lineTo(sx1, Math.round(sy) + 0.5); }
    ctx.stroke();
  }
}

// ------------------------------------------------------------------ chapter 1: the ledger line
const LEDGER = [['BILAG', 'DATO', 'TEKST', 'BELØP'], ['118212', '12.01.2024', 'INNKJØP NYBIL', '412 300,00'], ['118240', '02.02.2024', 'KLARGJØRING', '8 450,00'],
  ['118432', '14.03.2024', 'SALG NYBIL', '489 900,00'], ['118433', '14.03.2024', 'INNBYTTE', '−96 000,00'], ['121874', '20.09.2024', 'SERVICE 30 000 KM', '6 180,00']];
const row = (r) => r[0].padEnd(9) + r[1].padEnd(13) + r[2].padEnd(20) + r[3].padStart(12);
function drawLedger(ctx, t) {
  if (!visible(-3, -4, 20, 4, 60)) return;
  const fs = 0.5, S = cam.s;
  ctx.save(); ctx.font = M.font(fs * S, 500, MONO); ctx.textBaseline = 'alphabetic';
  LEDGER.forEach((r, i) => {
    const [x, y] = P(0, (i - 3) * 0.95);
    ctx.fillStyle = i === 3 ? G : i === 0 ? rgba(LB, 0.9) : rgba(G, 0.26);
    ctx.fillText(row(r), x, y);
  });
  ctx.fillStyle = rgba(LB, 0.5); ctx.fillRect(P(0, 0)[0], P(0, -2.4)[1], 16.2 * S, 1.2);
  ctx.restore();
}

// ------------------------------------------------------------------ chapter 2: the car, side elevation, bright sheet
const CXE = 26, GY = 2.09, CT0 = 4.2;
const HX = [24.1, 25.05, 26.2, 27.1, 27.95];
function drawCar(ctx, t) {
  if (t < CT0 || !visible(CXE - 4.5, 0, CXE + 4.5, 4, 80)) return;
  const S = cam.s, Q = ([x, y]) => P(CXE + x, GY - y), path = (pts) => pts.map(Q), EL = world.EL;
  const p = (t0, d = 1.6) => E.inOutCubic(prog(t, CT0 + t0, CT0 + t0 + d));
  ctx.save(); ctx.lineJoin = 'round'; ctx.lineCap = 'round';
  const bp = path(EL.body), fillA = ss(CT0 + 3.4, CT0 + 5.4, t);
  if (fillA > 0) {
    const yT = Q([0, 1.5])[1], yB = Q([0, 0.25])[1], g = ctx.createLinearGradient(0, yT, 0, yB);
    g.addColorStop(0, rgba(C.white, fillA)); g.addColorStop(1, rgba(CL, 0.7 * fillA));
    ctx.fillStyle = g; ctx.beginPath(); bp.forEach(([x, y], i) => (i ? ctx.lineTo(x, y) : ctx.moveTo(x, y))); ctx.closePath(); ctx.fill();
  }
  const gp = path(EL.glass); const ga = ss(CT0 + 1.2, CT0 + 2.6, t);
  ctx.fillStyle = rgba(LB, 0.2 * ga); ctx.beginPath(); gp.forEach(([x, y], i) => (i ? ctx.lineTo(x, y) : ctx.moveTo(x, y))); ctx.fill();
  const lw = clamp(S / 110, 1.4, 2.6);
  ctx.strokeStyle = rgba(G, 0.95); ctx.lineWidth = lw; K.drawOn(ctx, gp, p(0, 1.4));
  K.drawOn(ctx, bp, p(0.2, 2.6), true);
  ctx.lineWidth = clamp(S / 180, 1, 1.8); ctx.strokeStyle = rgba(G, 0.55);
  EL.lines.forEach((l, i) => K.drawOn(ctx, path(l), p(1 + i * 0.16, 1.2)));
  for (const w of EL.wheels) {
    const wp_ = p(1.4, 1.4);
    ctx.fillStyle = rgba(CL, 0.95 * wp_); ctx.beginPath(); path(w.rings[0]).forEach(([x, y], i) => (i ? ctx.lineTo(x, y) : ctx.moveTo(x, y))); ctx.fill();
    ctx.strokeStyle = rgba(G, 0.95); ctx.lineWidth = lw; K.drawOn(ctx, path(w.rings[0]), wp_);
    ctx.strokeStyle = rgba(G, 0.6); ctx.lineWidth = clamp(S / 180, 1, 1.8); K.drawOn(ctx, path(w.rings[1]), p(1.7, 1.2)); K.drawOn(ctx, path(w.rings[2]), p(2, 0.8));
    w.spokes.forEach((sp_, i) => K.drawOn(ctx, path(sp_), p(2.1 + i * 0.08, 0.6)));
  }
  ctx.strokeStyle = rgba(LB, 0.8); ctx.lineWidth = 1;
  K.drawOn(ctx, path([[-3.6, 0], [3.6, 0]]), p(0.6, 2));
  ctx.setLineDash([8, 6]);
  for (const [x] of [EL.FA, EL.RA]) K.drawOn(ctx, path([[x, -0.42], [x, 1.72]]), p(2.4, 1));
  ctx.setLineDash([]);
  const dp = p(2.8, 1); K.drawOn(ctx, path([[EL.RA[0], -0.3], [EL.FA[0], -0.3]]), dp);
  if (dp > 0.6 && S > 120) { const [tx, ty] = Q([0.015, -0.42]); mono(ctx, '2 870', tx, ty, 22, rgba(LB, ss(0.6, 1, dp)), { align: 'center' }); }
  // the history line under the thread, with its ticks; each pops as the tip passes it
  const hp = E.inOutCubic(prog(t, CT0 + 2, CT0 + 4.2));
  ctx.strokeStyle = rgba(LB, 0.75); ctx.lineWidth = 1; K.drawOn(ctx, path([[-2.3, 1.95], [2.35, 1.95]]), hp);
  const lab = clamp(0.105 * S, 14, 26), la = ss(70, 140, S);
  EL.history = EL.history || world.HISTORY;
  world.HISTORY.forEach(([x, date, label], i) => {
    const tt = tipT(ML('carIn') + (CXE + x - 23.2)) - 0.25, q = E.outCubic(prog(t, tt, tt + 0.5)); if (q <= 0) return;
    const main = i === 2, [px, py] = Q([x, 1.95]);
    ctx.strokeStyle = main ? RED : rgba(LB, 0.95); ctx.lineWidth = main ? 2.2 : 1.4;
    ctx.beginPath(); ctx.moveTo(px, py - 0.075 * S * q); ctx.lineTo(px, py + 0.075 * S * q); ctx.stroke();
    ctx.save(); ctx.globalAlpha = la * q;
    mono(ctx, date, px - 4, py - 0.34 * S * q - 6, lab * 0.9, main ? G : rgba(G, 0.6), { track: 0.04 });
    mono(ctx, label, px - 4, py - 0.2 * S * q - 4, lab, main ? G : rgba(BL, 1), { track: 0.04, weight: main ? 600 : 500 });
    ctx.restore();
    if (main) {
      const [dx, dy] = Q([world.DOT.x, world.DOT.y]); ctx.strokeStyle = rgba(RED, 0.9); ctx.lineWidth = 1.5;
      K.drawOn(ctx, [[px, py + 0.075 * S], [dx, dy]], q);
      ctx.fillStyle = RED; ctx.beginPath(); ctx.arc(dx, dy, Math.max(3, 0.026 * S) * q, 0, 7); ctx.fill();
    }
  });
  ctx.restore();
}

// ------------------------------------------------------------------ chapter 3: the site, six departments
const DEPT_X = (i) => -47.5 + 19 * i;
const LABELS = [['kjøpekontrakt', 'innbytte', 'faktura 118432'], ['VIN', 'lagerdager 64', 'beholdning'], ['arbeidsordre', 'timer 3,5', 'deler'], ['delenummer', 'uttak', 'varetelling'], ['garantikrav', 'refusjon', 'VIN'], ['servicehistorikk', 'booking', 'kunde']];
const TD = [0, 1, 2, 3, 4, 5].map((i) => markT('d' + i));
const edgeK = pchip([[14.6, -100], [16.4, -58], [18.5, -10], [21, 45], [24, 96], [24.5, 100]]);
function drawSite(ctx, t) {
  if (t < 14.5 || !visible(OX - 86, -52, OX + 86, 52, 20)) return;
  const edge = edgeK(t), reveal = t < 24.5;
  const call = (c) => world.site(c, { k: 1, cx: cam.cx - OX, cy: cam.cy, s: cam.s });
  if (reveal) {
    const ex = P(OX + edge, 0)[0], fe = Math.max(30, 7 * cam.s);
    if (ex > -10) {
      ctx.save(); ctx.beginPath(); ctx.rect(0, 0, Math.min(W, ex), H); ctx.clip();
      call(ctx);
      if (ex < W + fe) { const g = ctx.createLinearGradient(ex - fe, 0, ex, 0); g.addColorStop(0, rgba(C.arctic, 0)); g.addColorStop(1, rgba(C.arctic, 1)); ctx.fillStyle = g; ctx.fillRect(ex - fe, 0, fe + 2, H); }
      ctx.restore();
    }
    if (edge < 99) { ctx.strokeStyle = rgba(LB, 0.8); ctx.lineWidth = 1; ctx.beginPath(); const [a, b] = [P(OX + edge, -52), P(OX + edge, 52)]; ctx.moveTo(a[0], a[1]); ctx.lineTo(b[0], b[1]); ctx.stroke(); }
  } else call(ctx);
  // departments light up as the thread passes
  const S = cam.s;
  for (let i = 0; i < 6; i++) {
    const t0 = TD[i], a1 = E.outCubic(prog(t, t0 - 0.9, t0 + 0.1)), a = a1 * (1 - 0.7 * ss(t0 + 1.7, t0 + 2.9, t));
    if (a <= 0.01) continue;
    const [x, y] = P(OX - 57 + 19 * i, -46), w = 19 * S, h = 22 * S;
    ctx.fillStyle = rgba(BL, 0.07 * a); ctx.fillRect(x, y, w, h);
    ctx.strokeStyle = rgba(BL, 0.95 * a); ctx.lineWidth = 1.6; ctx.strokeRect(x, y, w, h);
  }
  // kicker + data labels popping out of each department
  const kS = clamp((cam.s - 12) / 12); ctx.save(); ctx.globalAlpha *= kS;
  for (let i = 0; i < 6 && kS > 0.02; i++) {
    const t0 = TD[i], out = t0 + 3.1; if (t < t0 - 0.8 || t > out) continue;
    const [kx, ky] = P(OX - 57 + 19 * i + 1.3, -49.2);
    K.kicker(ctx, world.DEPTS[i], kx, ky, t, { color: BL, t0: t0 - 0.7, out, size: 24 });
    const La = ML('d' + i), offs = [-4.5, 0, 4.5], ys = [-41.4, -43.9, -46.4];
    LABELS[i].forEach((txt, j) => {
      const tj = t0 - 0.2 + j * 0.28, q = E.outCubic(prog(t, tj, tj + 0.45)) * (1 - E.inCubic(prog(t, out - 0.4, out)));
      if (q <= 0) return;
      const [px, py] = posAt(La + offs[j]), a = P(px, py - 0.5), b = P(px, ys[j] + 0.3 + (1 - q) * 4);
      ctx.strokeStyle = rgba(LB, 0.95); ctx.lineWidth = 1.2; ctx.beginPath(); ctx.moveTo(a[0], a[1]); ctx.lineTo(b[0], b[1]); ctx.stroke();
      const d = P(px, py); ctx.fillStyle = BL; ctx.beginPath(); ctx.arc(d[0], d[1], 3.5 * q, 0, 7); ctx.fill();
      const [lx, ly] = P(px + 0.45, ys[j]);
      tag(ctx, txt, lx, ly, t, tj + 0.12, out, { size: 24, color: G });
    });
  }
  ctx.restore();
}

// ------------------------------------------------------------------ chapter 4: braid + data platform
const TRACK_Y = (j) => -52.5 - (5 - j) * 1.7;
const STRANDS = (() => {
  const out = [], Lex = ML('exitB'), Lpe = ML('pEnter');
  for (let j = 0; j < 6; j++) {
    const cx = DEPT_X(j), xJ = 236 + 4 * j, xs = cx + OX + 12.5;
    // join arclength on the red path at world x = xJ
    let Lj = Lex; for (let i = idxAt(Lex); i < TH.length; i++) if (TH[i][0] >= xJ) { Lj = CL_[i]; break; }
    let st0 = TH[idxAt(ML('d' + j))]; for (let i = idxAt(ML('d' + j)); i < TH.length; i++) if (TH[i][0] >= xs) { st0 = TH[i]; break; }
    const jp = posAt(Lj), ty = TRACK_Y(j), A = [[st0[0], st0[1]], [xs, -34], [xs, -44], [xs, -48.5], [xs + 2.6, ty + 0.3], [xs + 6, ty], [OX + 76 + (j > 3 ? 4 : 0), ty], [xJ - 16, ty], [xJ - 7, (ty + jp[1]) / 2 - 1], [jp[0], jp[1]]];
    const pts = crPoly(A, 0.4).pts;
    const lenA = pts.length;
    // braid along the red path, then fan into the first plate
    const lam = 13, amp = 0.95;
    for (let L = Lj; L < Lpe; L += 0.5) { const env = ss(0, 9, L - Lj), ph = (j / 6) * Math.PI * 2; pts.push(posAt(L, amp * env * Math.sin(((L - Lj) / lam) * Math.PI * 2 + ph))); }
    const rowN = [0, 1, 2, 3, 4, 2][j], node = [PX - 48 + 1.5 * rowN + (j === 5 ? 6 : 0), PY + 25 - 3 * rowN];
    const last = pts[pts.length - 1], mid = [lerp(last[0], node[0], 0.5), lerp(last[1], node[1], 0.5)];
    for (let q = 1; q <= 14; q++) { const f = q / 14; pts.push([lerp(lerp(last[0], mid[0], f), lerp(mid[0], node[0], f), f), lerp(lerp(last[1], mid[1], f), lerp(mid[1], node[1], f), f)]); }
    out.push({ pts, node, fa: lenA / pts.length, tj: tipT(Lj) });
  }
  return out;
})();
const PLATES = ['KJØRETØYDATA', 'TRANSAKSJONER', 'ØKONOMISKE DATA'];
const PLATE_T0 = [32.0, 32.5, 33.0];
function plateGeom(j) {
  const cy = PY + 19 - 19 * j, hh = 6, lean = 6;
  return { cy, tl: [PX - 48 + lean, cy - hh], tr: [PX + 48 + lean, cy - hh], br: [PX + 48, cy + hh], bl: [PX - 48, cy + hh] };
}
function drawPlatform(ctx, t) {
  if (t < 29 || !visible(PX - 60, PY - 52, PX + 62, PY + 40, 30)) return;
  const S = cam.s, kp = kz(11), ka = ss(0.4, 0.75, kp);
  for (let j = 0; j < 3; j++) {
    const g = plateGeom(j), rp = E.inOutCubic(prog(t, PLATE_T0[j], PLATE_T0[j] + 1.8)); if (rp <= 0) continue;
    const pierceT = markT(['pc0', 'pc1', 'pc2'][j]), lit = E.outCubic(prog(t, pierceT - 0.5, pierceT + 0.3));
    const pts = [g.bl, g.br, g.tr, g.tl].map(([x, y]) => P(x, y));
    ctx.save();
    ctx.fillStyle = rgba(C.white, 0.9 * ss(0.15, 0.7, rp)); ctx.beginPath(); pts.forEach(([x, y], i) => (i ? ctx.lineTo(x, y) : ctx.moveTo(x, y))); ctx.closePath(); ctx.fill();
    ctx.fillStyle = rgba(BL, (0.04 + 0.08 * lit) * ss(0.15, 0.7, rp)); ctx.fill();
    // slab thickness
    ctx.strokeStyle = rgba(BL, 0.55 * rp); ctx.lineWidth = 1;
    const th = 0.9 * S, bl = pts[0], br = pts[1]; ctx.beginPath(); ctx.moveTo(bl[0], bl[1]); ctx.lineTo(bl[0], bl[1] + th); ctx.lineTo(br[0], br[1] + th); ctx.lineTo(br[0], br[1]); ctx.stroke();
    // lattice
    const la = ss(0.3, 0.95, rp); ctx.beginPath();
    for (let r = 1; r < 4; r++) { const f = r / 4, a = [lerp(g.bl[0], g.tl[0], f), lerp(g.bl[1], g.tl[1], f)], b = [lerp(g.br[0], g.tr[0], f), lerp(g.br[1], g.tr[1], f)], A = P(...a), B = P(...b); ctx.moveTo(A[0], A[1]); ctx.lineTo(lerp(A[0], B[0], rp), lerp(A[1], B[1], rp)); }
    for (let c = 1; c < 16; c++) { const f = c / 16, a = [lerp(g.bl[0], g.br[0], f), g.bl[1]], b = [lerp(g.tl[0], g.tr[0], f), g.tl[1]], A = P(...a), B = P(...b); ctx.moveTo(A[0], A[1]); ctx.lineTo(B[0], B[1]); }
    ctx.strokeStyle = rgba(BL, (0.3 + 0.2 * lit) * la); ctx.stroke();
    ctx.fillStyle = rgba(BL, 0.85 * la); ctx.beginPath();
    for (let r = 0; r <= 4; r++) for (let c = 0; c <= 16; c++) {
      if (c / 16 > rp + 0.02) break;
      const fr = r / 4, fc = c / 16, x = lerp(lerp(g.bl[0], g.tl[0], fr), lerp(g.br[0], g.tr[0], fr), fc), y = lerp(g.bl[1], g.tl[1], fr), [sx, sy] = P(x, y);
      ctx.moveTo(sx + 2.2, sy); ctx.arc(sx, sy, 2.2, 0, 7);
    }
    ctx.fill();
    ctx.strokeStyle = rgba(BL, 0.95); ctx.lineWidth = 1.8; K.drawOn(ctx, [...pts, pts[0]], rp);
    ctx.restore();
    // the label above the plate's top-left corner
    const [lx, ly] = P(g.tl[0] - 0.4, g.tl[1] - 1.4);
    if (ka > 0) { ctx.save(); ctx.globalAlpha *= ka; tag(ctx, PLATES[j], lx, ly, t, pierceT - 0.3, 1e9, { size: 24 * kp, color: BL, weight: 600, halo: 8 }); ctx.restore(); }
  }
  // vertical links between the plates, dashed
  const la = ss(33.6, 35, t);
  if (la > 0) {
    ctx.save(); ctx.setLineDash([6, 6]); ctx.strokeStyle = rgba(BL, 0.5 * la); ctx.lineWidth = 1; ctx.beginPath();
    for (let j = 0; j < 2; j++) for (const f of [0.19, 0.5, 0.81]) {
      const a = plateGeom(j), b = plateGeom(j + 1), A = P(lerp(a.tl[0], a.tr[0], f), a.tl[1]), B = P(lerp(b.bl[0], b.br[0], f), b.bl[1]);
      ctx.moveTo(A[0], A[1]); ctx.lineTo(B[0], B[1]);
    }
    ctx.stroke(); ctx.restore();
  }
  if (ka > 0) { ctx.save(); ctx.globalAlpha *= ka; K.kicker(ctx, 'BDO dataplattform', P(PX - 42, PY - 45.5)[0], P(PX - 42, PY - 45.5)[1], t, { color: BL, t0: markT('pEnter') - 0.5, out: markT('kTop') + 0.8, size: 26 * kp }); ctx.restore(); }
}
function drawStrands(ctx, t) {
  if (t < 22) return;
  ctx.save(); ctx.lineJoin = 'round'; ctx.lineCap = 'round';
  const lw = clamp(cam.s / 8, 1, 1.8);
  ctx.lineWidth = lw; ctx.strokeStyle = rgba(BL, 0.95);
  STRANDS.forEach((st, j) => {
    const p = st.fa * E.inOutSine(prog(t, TD[j] + 1.2, st.tj)) + (1 - st.fa) * E.inOutCubic(prog(t, st.tj - 0.1, st.tj + 2.6)); if (p <= 0) return;
    const pts = st.pts.map(([x, y]) => P(x, y)); K.drawOn(ctx, pts, p);
    if (p >= 1) { const [x, y] = P(...st.node); ctx.fillStyle = BL; ctx.beginPath(); ctx.arc(x, y, 3, 0, 7); ctx.fill(); }
  });
  ctx.restore();
}

// ------------------------------------------------------------------ chapter 4b: industry knowledge column
const KNOW = [
  { n: '1,7 %', l: 'Driftsmargin, merkeforhandlere 2024', s: 'opp fra 1,5 % i 2023' },
  { n: '2,5 %', l: 'Driftsmargin, bruktbilforhandlere 2024', s: 'BDO Bilbransjeanalysen 2025' },
  { n: '1 300+', l: 'bilbransjebedrifter analysert', s: 'BDO Bilbransjeanalysen 2025' },
];
function drawKnow(ctx, t) {
  if (t < 35 || !visible(KX - 4, -70, KX + 60, 30, 20)) return;
  const S = cam.s, RY = [-56, -29, -2], kk = kz(13), ka = ss(0.4, 0.75, kk);
  const hp = E.outCubic(prog(t, markT('kTop') - 0.2, markT('kTop') + 0.6));
  ctx.save(); ctx.globalAlpha *= ka;
  K.kicker(ctx, 'Bransjekunnskap', P(KX + 4, -62.2)[0], P(KX + 4, -62.2)[1], t, { color: BL, t0: markT('kTop') - 0.3, out: markT('kEnd') + 1.6, size: 24 * kk });
  ctx.restore();
  KNOW.forEach((k, i) => {
    const t0 = markT('rule' + (i + 1)) - 0.15;
    const rp = E.outCubic(prog(t, t0, t0 + 0.9));
    if (rp > 0) { ctx.strokeStyle = rgba(LB, 0.9); ctx.lineWidth = 1; ctx.beginPath(); const a = P(KX, RY[i]), b = P(KX + 56 * rp, RY[i]); ctx.moveTo(a[0], a[1]); ctx.lineTo(b[0], b[1]); ctx.stroke(); }
    const [nx, ny] = P(KX + 3, RY[i] + 12.2);
    K.line(ctx, k.n, nx, ny, t, { size: clamp(9.5 * S, 60 * kk, 150), weight: 300, color: BL, t0: t0 + 0.1, dur: 0.9 });
    ctx.save(); ctx.globalAlpha *= ka;
    const [lx, ly] = P(KX + 3, RY[i] + 17.4);
    K.line(ctx, k.l, lx, ly, t, { size: 30 * kk, weight: 500, color: G, t0: t0 + 0.35 });
    const [sx, sy] = P(KX + 3, RY[i] + 21.4);
    K.line(ctx, k.s.toUpperCase(), sx, sy, t, { size: 21 * kk, weight: 500, color: LB, t0: t0 + 0.55, family: MONO, track: 0.08 });
    ctx.restore();
  });
  void hp;
}

// ------------------------------------------------------------------ chapter 5: timeline chart
const MONTHS = ['JAN', 'FEB', 'MAR', 'APR', 'MAI', 'JUN', 'JUL', 'AUG', 'SEP', 'OKT', 'NOV', 'DES'];
const LANE = (() => {
  const r = K.mulberry32(404), lanes = [];
  for (let j = 0; j < 5; j++) {
    const n = [72, 49, 66, 44, 36][j], items = [];
    for (let i = 0; i < n; i++) {
      const x = CH.x0 + ((i + 0.5) / n) * 200, m = (x - CH.x0) / CH.mw, seas = 0.5 + 0.35 * Math.sin(m * 0.9 + j), v = clamp(seas * 0.7 + r() * 0.5, 0.12, 1);
      let h = v * 3.3;
      if (j === 3) h = (r() < 0.55 ? 0.3 + r() * 0.9 : 0.1) + (Math.abs(m - 4.5) < 0.8 ? 1.5 : 0);
      if (j === 4) h = 0.25 + r() * 0.7 + (m > 10.6 ? 1.8 * r() : 0);
      if (j === 2) h = 0.4 + r() * 2.6;
      items.push({ x, h, tr: timeAtX(x) - 1.5 });
    }
    lanes.push(items);
  }
  return lanes;
})();
function drawChart(ctx, t) {
  if (t < 40 || !visible(CH.x0 - 40, 10, CH.x0 + 270, 64, 40)) return;
  const S = cam.s, x1 = CH.x0 + 200, kc = kz(11), ka = ss(0.4, 0.75, kc);
  ctx.save(); ctx.lineJoin = 'round'; ctx.lineCap = 'butt';
  const t0 = markT('c0') - 3.2, ap = E.outCubic(prog(t, t0, t0 + 2.2));
  // month gridlines + axis
  ctx.strokeStyle = rgba(CL, 1); ctx.lineWidth = 1; ctx.beginPath();
  for (let m = 0; m <= 12; m++) { const x = CH.x0 + m * CH.mw, a = P(x, 16.5), b = P(x, 60.5 - (1 - ap) * 44); ctx.moveTo(a[0], a[1]); ctx.lineTo(a[0], lerp(a[1], b[1], ap)); }
  ctx.stroke();
  if (ka > 0) for (let m = 0; m < 12; m++) { const [mx, my] = P(CH.x0 + (m + 0.5) * CH.mw, 60); mono(ctx, MONTHS[m], mx, my + 4, 20 * kc, rgba(LB, ap * ka), { align: 'center', track: 0.12 }); }
  if (ka > 0) { ctx.save(); ctx.globalAlpha *= ka; K.kicker(ctx, 'Analyse · tolv måneder', P(CH.x0, 12.2)[0], P(CH.x0, 12.2)[1], t, { color: BL, t0: t0 + 0.2, out: 56.8, size: 24 * kc }); ctx.restore(); }
  for (let j = 0; j < 5; j++) {
    const y = CH.lanes[j], a = P(CH.x0, y + 1.9), b = P(x1, y + 1.9);
    ctx.strokeStyle = rgba(LB, 0.55); ctx.lineWidth = 1; ctx.beginPath(); ctx.moveTo(a[0], a[1]); ctx.lineTo(lerp(a[0], b[0], ap), a[1]); ctx.stroke();
    const items = LANE[j];
    ctx.strokeStyle = rgba(BL, 0.85); ctx.lineWidth = clamp(S / 8, 1.2, 2); ctx.beginPath();
    if (j === 1) {
      let started = false;
      items.forEach((it) => { const g = E.outCubic(prog(t, it.tr, it.tr + 0.8)); if (g <= 0) return; const [x, yy] = P(it.x, y + 1.9 - it.h * g * 1.0); if (!started) { ctx.moveTo(x, yy); started = true; } else ctx.lineTo(x, yy); });
      ctx.stroke();
    } else if (j === 2) {
      ctx.fillStyle = rgba(BL, 0.8);
      items.forEach((it) => { const g = E.outCubic(prog(t, it.tr, it.tr + 0.8)); if (g <= 0) return; const [x, yy] = P(it.x, y + 1.9 - it.h * g); ctx.moveTo(x + 2.4, yy); ctx.arc(x, yy, 2.4, 0, 7); });
      ctx.fill();
    } else {
      items.forEach((it) => { const g = E.outCubic(prog(t, it.tr, it.tr + 0.8)); if (g <= 0) return; const [x, yy] = P(it.x, y + 1.9), [, y2] = P(it.x, y + 1.9 - it.h * g); ctx.moveTo(x, yy); ctx.lineTo(x, y2); });
      ctx.stroke();
    }
  }
  // lane names, sticky at the left of the screen while the chart is under the camera
  const la = win(t, 44.2, 57.0, 0.8, 1.0) * ka;
  if (la > 0) {
    ctx.save(); ctx.globalAlpha *= la;
    for (let j = 0; j < 5; j++) { const [, yy] = P(0, CH.lanes[j] + 1.9); mono(ctx, CH.names[j], 104, yy - 8, 22, rgba(G, 0.85), { track: 0.1, halo: 9 }); }
    ctx.restore();
  }
  // the three flagged points
  for (const f of FLAGS) {
    const tp = markT(f.name), [px, py] = P(f.x, f.y - 0.2), q = spr(t - tp - 0.3) * Math.max(kc, 0.4);
    ctx.fillStyle = rgba(BL, 0.9 * (1 - ss(tp, tp + 0.5, t))); ctx.beginPath(); ctx.arc(px, py, 4 * Math.max(kc, 0.4), 0, 7); ctx.fill();
    if (q > 0) { ctx.fillStyle = RED; ctx.beginPath(); ctx.arc(px, py, 6.5 * q, 0, 7); ctx.fill(); }
    const tt = tp + 1.1, tq = E.outCubic(prog(t, tt, tt + 0.5));
    if (tq > 0 && ka > 0.05) {
      const lx = px - 12, by = P(0, 15.0)[1];
      ctx.save(); ctx.font = M.font(24, 500, DISPLAY); const w = ctx.measureText(f.text).width; ctx.restore();
      ctx.strokeStyle = rgba(RED, 0.85); ctx.lineWidth = 1.2; ctx.beginPath(); ctx.moveTo(px, py - 48 * kc); ctx.lineTo(px, lerp(py - 48 * kc, by, tq)); ctx.stroke();
      const bw = (w + 30) * tq, bh = 42;
      ctx.fillStyle = rgba(C.white, 0.97); ctx.fillRect(lx, by - bh, bw, bh);
      ctx.strokeStyle = rgba(G, 0.18); ctx.lineWidth = 1; ctx.strokeRect(lx, by - bh, bw, bh);
      ctx.fillStyle = RED; ctx.fillRect(lx, by - bh, 4, bh);
      ctx.save(); ctx.beginPath(); ctx.rect(lx + 4, by - bh, bw - 4, bh); ctx.clip();
      ctx.font = M.font(24, 500, DISPLAY); ctx.fillStyle = G; ctx.fillText(f.text, lx + 16, by - 14); ctx.restore();
    }
  }
  ctx.restore();
}
const spr = (s) => M.spring(s, M.SPRING.bouncy);

// ------------------------------------------------------------------ chapter 6: the conversation
const VAL_R = ['Bedre oversikt over datagrunnlaget', 'Bedre grunnlag for risikovurdering', 'Mer tid til faglig vurdering'];
const VAL_K = ['Bedre oversikt over egen virksomhet', 'Mer relevant oppfølging', 'Tydeligere dialog om funn'];
function drawConv(ctx, t) {
  if (t < 54 || !visible(MX - 60, 26, MX + 60, 62, 20)) return;
  const S = cam.s, tc = markT('cEnd') - 0.6, tag_ = (t_) => t_ - 0.0, kv = kz(16), ka = ss(0.4, 0.75, kv);
  ctx.save(); ctx.globalAlpha *= ka;
  const [lx0, hy] = P(822.5, 33), [rx0] = P(837.5, 33);
  K.line(ctx, 'VERDI FOR REVISOR', lx0, hy, t, { size: 24, weight: 500, color: BL, t0: tc, family: MONO, align: 'right', track: 0.12 });
  K.line(ctx, 'VERDI FOR KUNDEN', rx0, hy, t, { size: 24, weight: 500, color: BL, t0: tc + 0.15, family: MONO, track: 0.12 });
  const rp = E.outCubic(prog(t, tc + 0.1, tc + 1.2)); ctx.strokeStyle = rgba(LB, 0.9); ctx.lineWidth = 1;
  if (rp > 0) { ctx.beginPath(); let a = P(822.5, 34.4), b = P(822.5 - 30 * rp, 34.4); ctx.moveTo(a[0], a[1]); ctx.lineTo(b[0], b[1]); a = P(837.5, 34.4); b = P(837.5 + 42 * rp, 34.4); ctx.moveTo(a[0], a[1]); ctx.lineTo(b[0], b[1]); ctx.stroke(); }
  for (let i = 0; i < 3; i++) {
    const tl = markT('sL' + i) - 0.1, tr = markT('sR' + i) - 0.1, [ax, ay] = P(822.5, ROWS[i] + 0.7), [bx, by] = P(837.5, ROWS[i] + 0.7);
    K.line(ctx, VAL_R[i], ax, ay, t, { size: 34, weight: 500, color: G, t0: tl, align: 'right' });
    K.line(ctx, VAL_K[i], bx, by, t, { size: 34, weight: 500, color: G, t0: tr });
    // stitches: short hairlines from the thread to the text
    for (const [x0, x1, q0] of [[826, 823.6, tl], [834, 836.4, tr]]) {
      const q = E.outCubic(prog(t, q0, q0 + 0.4)); if (q <= 0) continue;
      const a = P(x0, ROWS[i] - 0.1), b = P(lerp(x0, x1, q), ROWS[i] - 0.1); ctx.strokeStyle = rgba(RED, 0.9); ctx.lineWidth = 1.6; ctx.beginPath(); ctx.moveTo(a[0], a[1]); ctx.lineTo(b[0], b[1]); ctx.stroke();
    }
  }
  ctx.restore(); void tag_; void S;
}

// ------------------------------------------------------------------ overview labels (the payoff view)
const OVL = [[-6, 'Transaksjon · bil'], [60, 'Virksomheten'], [330, 'Dataplattform'], [496, 'Bransjekunnskap'], [600, 'Analyse'], [800, 'Verdi']];
function drawOverview(ctx, t) {
  const a = ss(2.6, 1.9, cam.s) * ss(66.4, 68.2, t) * (1 - ss(70.4, 71.6, t));
  if (a <= 0.01) return;
  OVL.forEach(([x, str], i) => {
    const [sx, sy0] = P(x, 70), sy = sy0 + (i % 2) * 30; ctx.save(); ctx.globalAlpha *= a;
    ctx.fillStyle = rgba(G, 0.4); ctx.fillRect(sx, sy0 - 24, 1, 24 + (i % 2) * 30 + 6);
    mono(ctx, String(i + 1).padStart(2, '0') + '  ' + str.toUpperCase(), sx + 8, sy, 18, G, { track: 0.1 });
    ctx.restore();
  });
  const [tx, ty] = P(SHEET.x1 - 12, SHEET.y1 - 14);
  mono(ctx, 'BDO  ·  DEN RØDE TRÅDEN  ·  ARK 1 AV 1', tx, ty, 18, rgba(G, 0.7 * a), { align: 'right', track: 0.12 });
}

// ------------------------------------------------------------------ the thread itself
function drawThread(ctx, t) {
  const L = Lof(t), i = idxAt(L), f = (L - CL_[i]) / Math.max(1e-9, CL_[i + 1] - CL_[i]), a = TH[i], b = TH[Math.min(i + 1, TH.length - 1)];
  const tip = [lerp(a[0], b[0], f), lerp(a[1], b[1], f)];
  const w = 3.2 + 1.0 * clamp(1 - cam.s / 8);
  const sx = W / 2 - cam.cx * cam.s, sy = H / 2 - cam.cy * cam.s, s = cam.s;
  ctx.save(); ctx.lineJoin = 'round'; ctx.lineCap = 'round';
  const trace = () => { ctx.beginPath(); ctx.moveTo(sx + TH[0][0] * s, sy + TH[0][1] * s); for (let k = 1; k <= i; k++) ctx.lineTo(sx + TH[k][0] * s, sy + TH[k][1] * s); ctx.lineTo(sx + tip[0] * s, sy + tip[1] * s); };
  trace(); ctx.strokeStyle = rgba(C.white, 0.9); ctx.lineWidth = w + 4; ctx.stroke();
  trace(); ctx.strokeStyle = RED; ctx.lineWidth = w; ctx.stroke();
  const [tx, ty] = P(tip[0], tip[1]), pulse = 1 + 0.25 * Math.sin(t * 5);
  ctx.strokeStyle = rgba(RED, 0.35); ctx.lineWidth = 1.5; ctx.beginPath(); ctx.arc(tx, ty, 12 * pulse, 0, 7); ctx.stroke();
  ctx.fillStyle = RED; ctx.beginPath(); ctx.arc(tx, ty, 6.5, 0, 7); ctx.fill();
  ctx.restore();
}

// ------------------------------------------------------------------ captions
const CAPS = [
  { s: ['Én transaksjon.'], a: 0.4, b: 7.6 },
  { s: ['Én bil. En hel historikk.'], a: 8.6, b: 15.6 },
  { s: ['Hele virksomheten.'], a: 17.6, b: 28.2 },
  { s: ['Data fra alle systemer. Samlet på ett sted.'], a: 30.0, b: 36.4, size: 66 },
  { s: ['Bransjekunnskap gir forståelse.'], a: 37.8, b: 43.8 },
  { s: ['Teknologien viser hvor vi bør se.'], a: 46.0, b: 52.6 },
  { s: ['Revisor vurderer hva det betyr.'], a: 53.3, b: 58.8 },
  { s: ['Effektiv revisjon. Relevant innsikt.'], a: 61.2, b: 66.2 },
  { s: ['Hele bildet.'], a: 66.9, b: 69.6 },
];
function drawCaptions(ctx, t) {
  let sc = 0; for (const c of CAPS) sc = Math.max(sc, win(t, c.a, c.b, 0.5, 0.45));
  if (sc > 0.01) {
    const g = ctx.createLinearGradient(0, 760, 0, 880); g.addColorStop(0, rgba(C.arctic, 0)); g.addColorStop(1, rgba(C.arctic, 0.94 * sc));
    ctx.fillStyle = g; ctx.fillRect(0, 760, W, 120); ctx.fillStyle = rgba(C.arctic, 0.94 * sc); ctx.fillRect(0, 880, W, 200);
  }
  for (const c of CAPS) {
    if (t < c.a - 0.05 || t > c.b) continue;
    const size = c.size || (c.s.length > 1 ? 76 : 80), two = c.s.length > 1;
    c.s.forEach((str, k) => K.line(ctx, str, K.CAP.x, two ? 860 + k * 70 : 930, t, { size, weight: 600, color: G, t0: c.a + k * 0.22, out: c.b }));
  }
}

// ------------------------------------------------------------------ hits: one per visual accent
const HITS = [];
const hit = (t, label, sfx, o) => HITS.push([t / PERIOD, label, sfx, o]);
hit(0.05, 'thread starts', 'tick'); hit(0.4, 'caption 1', 'whoosh'); hit(5.0, 'ledger done', 'blip');
HX.forEach((x, i) => hit(tipT(ML('carIn') + (CXE + (world.HISTORY[i][0]) - 23.2)) - 0.25, 'history tick ' + (i + 1), i === 2 ? 'click' : 'tick'));
hit(8.6, 'caption 2', 'whoosh'); hit(14.8, 'car left', 'swell'); hit(16.4, 'site', 'whoosh'); hit(17.6, 'caption 3', 'whoosh');
TD.forEach((t0, i) => hit(t0 - 0.2, 'dept ' + world.DEPTS[i], 'pop')); hit(29.0, 'exit building', 'swell'); hit(29.5, 'blue threads', 'riser');
hit(30.0, 'caption 4', 'whoosh'); ['pc0', 'pc1', 'pc2'].forEach((m, i) => hit(markT(m), 'plate ' + (i + 1), i === 2 ? 'thud' : 'blip'));
hit(37.8, 'caption 5', 'whoosh'); ['rule1', 'rule2', 'rule3'].forEach((m, i) => hit(markT(m), 'figure ' + (i + 1), 'pop'));
hit(46.0, 'caption 6', 'whoosh'); FLAGS.forEach((f, i) => { hit(markT(f.name) + 0.3, 'flag ' + (i + 1), 'bell'); hit(markT(f.name) + 1.1, 'tag ' + (i + 1), 'pop'); });
hit(53.3, 'caption 7', 'whoosh'); ['sL0', 'sR0', 'sL1', 'sR1', 'sL2', 'sR2'].forEach((m) => hit(markT(m), m, 'tick'));
hit(61.2, 'caption 8', 'whoosh'); hit(66.0, 'pull back', 'riser'); hit(68.4, 'whole picture', 'swell'); hit(70.3, 'thread ends', 'blip'); hit(71.4, 'logo', 'impact'); hit(72.4, 'tagline', 'bell');

// ------------------------------------------------------------------ the film
function drawAll(ctx, t) {
  LOGOW = K.drawLogo(ctx, 0, 0, 150, { alpha: 0 }).w;
  cam = camera(t);
  paper(ctx); worldGrid(ctx);
  drawLedger(ctx, t); drawCar(ctx, t); drawSite(ctx, t); drawPlatform(ctx, t); drawKnow(ctx, t); drawChart(ctx, t); drawConv(ctx, t);
  drawStrands(ctx, t); drawOverview(ctx, t); drawThread(ctx, t);
  // the logo, set into the sheet beside the end of the thread
  const lw = clamp(prog(t, 71.3, 72.5)); if (lw > 0) { const fc = finalCam(), lx = END[0] + (64 + LOGOW / 2) / fc.s, [sx, sy] = P(lx, END[1]); K.drawLogo(ctx, sx, sy, (150 / fc.s) * cam.s, { wipe: E.outQuint(lw) }); }
  drawCaptions(ctx, t);
  if (t > 71.5) {
    const q = E.outQuint(prog(t, 71.9, 72.9)); void q;
    K.line(ctx, 'Vi ser detaljene. Vi forstår helheten.', W / 2, 640, t, { size: 48, weight: 300, color: G, t0: 71.6, align: 'center' });
    K.line(ctx, 'TETT PÅ BILBRANSJEN', W / 2, 720, t, { size: 30, weight: 500, color: BL, t0: 72.0, family: MONO, track: 0.2, align: 'center' });
  }
}

M.film({
  fonts: [M.font(100, 600, DISPLAY), M.font(100, 300, DISPLAY), M.font(100, 500, DISPLAY), M.font(40, 500, MONO)],
  hits: HITS,
  samples: (t) => (t < 14 || t > 38 ? 4 : 5),
  async init() { await K.loadLogo(); },
  draw(ctx, u, t) { drawAll(ctx, t); },
  post(ctx, u, t) { K.grain(ctx, t, 1); },
});
