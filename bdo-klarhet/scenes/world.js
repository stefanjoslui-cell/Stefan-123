// world.js  0–50 s. The physical world as one architect's drawing, seen in one continuous pull-out:
// a ledger line -> the car it belongs to (side elevation) -> the car park (plan) -> the whole dealership
// -> a glide through the six departments -> the data lifts off the drawing into a network.
// Also exports site(): the site plan, reused by the ending on a bright sheet.
import * as K from '../lib/kit.js';
import * as M from '../lib/motion.js';
const { W, H, C, clamp, lerp, prog, E, ss, rgba, win, PTS } = K;

// ------------------------------------------------------------------ small helpers
function curve(keys) {       // C1 curve through [t, v] keys
  const n = keys.length, m = keys.map((_, i) => { const a = keys[Math.max(0, i - 1)], b = keys[Math.min(n - 1, i + 1)]; return (b[1] - a[1]) / (b[0] - a[0]); });
  m[0] = 0; m[n - 1] = 0;
  return (x) => {
    if (x <= keys[0][0]) return keys[0][1];
    if (x >= keys[n - 1][0]) return keys[n - 1][1];
    let i = 0; while (x > keys[i + 1][0]) i++;
    const h = keys[i + 1][0] - keys[i][0], t = (x - keys[i][0]) / h, t2 = t * t, t3 = t2 * t;
    return (2 * t3 - 3 * t2 + 1) * keys[i][1] + (t3 - 2 * t2 + t) * h * m[i] + (-2 * t3 + 3 * t2) * keys[i + 1][1] + (t3 - t2) * h * m[i + 1];
  };
}
export const toScreen = (cx, cy, s, x, y) => [W / 2 + (x - cx) * s, H / 2 + (y - cy) * s];
const arc = (cx, cy, r, a0, a1, n = 28) => Array.from({ length: n + 1 }, (_, i) => { const a = lerp(a0, a1, i / n); return [cx + r * Math.cos(a), cy + r * Math.sin(a)]; });
const lineCol = (k, a = 1) => (k < 0.5 ? rgba(C.cloud, a) : rgba(C.graphite, a));
const blueCol = (k, a = 1) => (k < 0.5 ? rgba(C.light, a) : rgba(C.blue, a));

// ------------------------------------------------------------------ the car, in plan (metres, nose +x)
const CAR = (() => {
  const body = new Path2D('M-2.35,-0.55 C-2.35,-0.85 -2.2,-0.92 -1.9,-0.93 L1.2,-0.93 C1.9,-0.93 2.28,-0.8 2.34,-0.42 C2.38,-0.15 2.38,0.15 2.34,0.42 C2.28,0.8 1.9,0.93 1.2,0.93 L-1.9,0.93 C-2.2,0.92 -2.35,0.85 -2.35,0.55 Z');
  const glass = new Path2D('M-1.52,-0.6 C-1.5,-0.72 -1.3,-0.76 -1.0,-0.76 L0.5,-0.76 C0.8,-0.76 1.02,-0.7 1.1,-0.56 C1.17,-0.2 1.17,0.2 1.1,0.56 C1.02,0.7 0.8,0.76 0.5,0.76 L-1.0,0.76 C-1.3,0.76 -1.5,0.72 -1.52,0.6 C-1.57,0.2 -1.57,-0.2 -1.52,-0.6 Z');
  const roof = new Path2D('M-1.02,-0.6 L0.42,-0.62 C0.5,-0.2 0.5,0.2 0.42,0.62 L-1.02,0.6 C-1.08,0.2 -1.08,-0.2 -1.02,-0.6 Z');
  const detail = new Path2D('M1.12,-0.56 C1.6,-0.6 2.0,-0.56 2.3,-0.4 M1.12,0.56 C1.6,0.6 2.0,0.56 2.3,0.4 M0.62,-0.93 L0.7,-1.08 L0.86,-1.08 L0.84,-0.93 M0.62,0.93 L0.7,1.08 L0.86,1.08 L0.84,0.93');
  return { body, glass, roof, detail };
})();
const TONE_N = [0.34, 0.34, 0.24, 0.16, 0.1, 0.06, 0.06, 0.2];           // night: paint as cloud-white wash
const TONE_D = ['#FFFFFF', '#FFFFFF', '#F7F7F7', '#EFEFEF', '#E7E7E7', '#D6D6D6', '#C6C6C6', '#BFC6CE'];
function carPlan(ctx, cam, x, y, rot, tone, k, a = 1, detail = true) {
  const [px, py] = toScreen(cam.cx, cam.cy, cam.s, x, y);
  if (px < -200 || px > W + 200 || py < -200 || py > H + 200) return;
  const c = Math.cos(rot) * cam.s, s = Math.sin(rot) * cam.s;
  ctx.save(); ctx.transform(c, s, -s, c, px, py);
  ctx.lineWidth = 1.25 / cam.s;
  ctx.fillStyle = k < 0.5 ? rgba(tone === 7 ? C.light : C.cloud, TONE_N[tone] * a) : TONE_D[tone];
  if (k >= 0.5) ctx.globalAlpha *= a;
  ctx.fill(CAR.body);
  ctx.strokeStyle = k < 0.5 ? rgba(C.cloud, 0.8 * a) : rgba(C.light, 0.85); ctx.stroke(CAR.body);
  ctx.fillStyle = k < 0.5 ? rgba(C.night, 0.75 * a) : rgba(C.light, 0.32); ctx.fill(CAR.glass);
  if (detail && cam.s > 9) {
    ctx.fillStyle = k < 0.5 ? rgba(C.cloud, TONE_N[tone] * 0.8 * a) : TONE_D[tone]; ctx.fill(CAR.roof);
    if (cam.s > 16) { ctx.lineWidth = 1 / cam.s; ctx.strokeStyle = k < 0.5 ? rgba(C.cloud, 0.45 * a) : rgba(C.graphite, 0.5); ctx.stroke(CAR.detail); ctx.stroke(CAR.roof); }
  }
  ctx.restore();
}

// ------------------------------------------------------------------ the site (world metres, y down)
const HERO = { x: 0, y: 12 };
const DEPTS = ['Salg', 'Lager', 'Verksted', 'Deler', 'Garanti', 'Service'];
const DEPT_DATA = ['ordre · kontrakter · innbytte', 'beholdning · VIN · lagerdager', 'arbeidsordre · timer · deler', 'delenummer · uttak · telling', 'garantikrav · refusjon', 'booking · servicehistorikk'];
const B = { x0: -57, x1: 57, y0: -46, y1: -24, w: 19 };
const LOT = (() => {
  const r = K.mulberry32(11), cars = [];
  for (let n = -4; n <= 4; n++) for (let m = -9; m <= 11; m++) for (const side of [0, 1]) {
    const x = 16.5 * n + (side ? 5.2 : 0), y = 12 + 2.7 * m, hero = n === 0 && m === 0 && !side;
    if (!hero && r() < 0.07) continue;
    cars.push({ x, y, rot: side ? Math.PI : 0, tone: hero ? 2 : Math.floor(r() * 8), hero, d: Math.hypot(x, y - 12) });
  }
  return cars;
})();
const LAMPS = []; for (let n = -4; n <= 5; n++) for (const y of [-8, 16, 40]) LAMPS.push([16.5 * n - 5.65, y]);
const TREES = (() => { const r = K.mulberry32(3), o = []; for (let x = -82; x <= 82; x += 6.5) { o.push([x + r() * 2, 48.2 + r(), 1.8 + r()]); } for (let y = -44; y < 46; y += 7) { o.push([-83 + r(), y + r() * 2, 1.8 + r()]); o.push([83 - r(), y + r() * 2, 1.8 + r()]); } return o; })();
// cars and fixtures inside the building: [dept, x, y, rot, tone]
const INDOOR = (() => {
  const o = [], x0 = (i) => B.x0 + B.w * i;
  [[4.5, -40, 0.5], [12, -41, -0.35], [6, -32.5, 0.2], [13.5, -33, -0.6]].forEach(([x, y, r], j) => o.push([0, x0(0) + x, y, r, [0, 7, 2, 5][j]]));
  for (let c = 0; c < 6; c++) for (const y of [-41, -34.5]) o.push([1, x0(1) + 2.6 + c * 2.75, y, Math.PI / 2, (c * 3 + (y < -40 ? 1 : 4)) % 8]);
  for (let c = 0; c < 4; c++) o.push([2, x0(2) + 3.2 + c * 4.2, -39.5, -Math.PI / 2, [3, 0, 6, 2][c]]);
  for (let c = 0; c < 2; c++) o.push([4, x0(4) + 5.5 + c * 7.5, -38.5, -Math.PI / 2, [1, 4][c]]);
  o.push([5, x0(5) + 6, -30.2, 0, 0], [5, x0(5) + 12.5, -30.2, 0, 5], [5, x0(5) + 6, -34.4, 0, 3]);
  return o;
})();

function rooms(ctx, cam, k, a, reveal) {
  const S = cam.s, T = (x, y) => toScreen(cam.cx, cam.cy, S, x, y);
  const rect = (x, y, w, h) => { const [px, py] = T(x, y); ctx.rect(px, py, w * S, h * S); };
  const seg = (x0, y0, x1, y1) => { const a0 = T(x0, y0), a1 = T(x1, y1); ctx.moveTo(a0[0], a0[1]); ctx.lineTo(a1[0], a1[1]); };
  // floor wash + walls
  ctx.fillStyle = k < 0.5 ? rgba(C.light, 0.06 * a) : rgba(C.white, 0.9 * a);
  ctx.beginPath(); rect(B.x0, B.y0, B.x1 - B.x0, B.y1 - B.y0); ctx.fill();
  ctx.strokeStyle = lineCol(k, 0.95 * a); ctx.lineWidth = clamp(0.3 * S, 1.5, 5);
  ctx.beginPath(); seg(B.x0, B.y1, B.x0, B.y0); seg(B.x0, B.y0, B.x1, B.y0); seg(B.x1, B.y0, B.x1, B.y1); ctx.stroke();
  // glazed front: double line + mullions
  ctx.lineWidth = 1.2; ctx.strokeStyle = blueCol(k, 0.9 * a); ctx.beginPath();
  seg(B.x0, B.y1, B.x1, B.y1); seg(B.x0, B.y1 - 0.35, B.x1, B.y1 - 0.35);
  if (S > 6) for (let x = B.x0; x <= B.x1 + 0.01; x += 1.9) seg(x, B.y1 - 0.35, x, B.y1);
  ctx.stroke();
  for (let i = 0; i < 6; i++) {
    const x0 = B.x0 + B.w * i, p = reveal ? reveal(i) : 1; if (p <= 0) continue;
    ctx.save(); ctx.globalAlpha *= p;
    // partition with a doorway on the corridor side
    if (i > 0) { ctx.strokeStyle = lineCol(k, 0.7 * a); ctx.lineWidth = clamp(0.2 * S, 1.2, 3); ctx.beginPath(); seg(x0, B.y0, x0, -31.5); ctx.stroke(); ctx.lineWidth = 1; ctx.strokeStyle = lineCol(k, 0.4 * a); ctx.beginPath(); const [dx, dy] = T(x0, -31.5); ctx.arc(dx, dy, 3.2 * S, 0, Math.PI / 2); ctx.stroke(); }
    ctx.strokeStyle = lineCol(k, 0.6 * a); ctx.lineWidth = 1.1; ctx.fillStyle = k < 0.5 ? rgba(C.cloud, 0.1 * a) : rgba(C.cloud, 0.8 * a);
    ctx.beginPath();
    if (i === 0) { rect(x0 + 1.2, -29.5, 4.2, 1.3); rect(x0 + 8.5, -28.2, 3.2, 1.2); const [tx, ty] = T(x0 + 15.5, -28.2); ctx.moveTo(tx + 1.1 * S, ty); ctx.arc(tx, ty, 1.1 * S, 0, 7); }
    if (i === 1) for (let c = 0; c <= 6; c++) { seg(x0 + 1.2 + c * 2.75, -44.2, x0 + 1.2 + c * 2.75, -38.4); seg(x0 + 1.2 + c * 2.75, -37.4, x0 + 1.2 + c * 2.75, -31.8); }
    if (i === 2) { for (let c = 0; c < 4; c++) { const bx = x0 + 3.2 + c * 4.2; rect(bx - 1.5, -40.6, 0.32, 0.5); rect(bx + 1.18, -40.6, 0.32, 0.5); rect(bx - 1.5, -38.9, 0.32, 0.5); rect(bx + 1.18, -38.9, 0.32, 0.5); seg(bx - 2.1, -43.5, bx - 2.1, -36); } seg(x0 + 17.6, -43.5, x0 + 17.6, -36); rect(x0 + 1.2, -45.3, 13, 0.9); for (let c = 0; c < 3; c++) { const [tx, ty] = T(x0 + 15.6 + c * 1.0, -44.8); ctx.moveTo(tx + 0.4 * S, ty); ctx.arc(tx, ty, 0.4 * S, 0, 7); } }
    if (i === 3) { for (let c = 0; c < 7; c++) { const sx = x0 + 2 + c * 2.35; rect(sx, -44.6, 0.9, 11.2); if (S > 20) for (let q = 1; q < 9; q++) seg(sx, -44.6 + q * 1.245, sx + 0.9, -44.6 + q * 1.245); } rect(x0 + 3, -29.8, 9, 1.1); }
    if (i === 4) { for (let c = 0; c < 2; c++) { const bx = x0 + 5.5 + c * 7.5; rect(bx - 1.9, -42.3, 3.8, 0.35); rect(bx - 1.9, -35.1, 3.8, 0.35); } rect(x0 + 13.2, -29.8, 4.2, 1.2); rect(x0 + 1.4, -45.2, 5, 0.8); }
    if (i === 5) { seg(x0 + 1.5, -32.3, x0 + 17, -32.3); seg(x0 + 1.5, -36.5, x0 + 17, -36.5); rect(x0 + 2, -44.6, 7.5, 1.2); for (let c = 0; c < 5; c++) rect(x0 + 11.5 + c * 1.1, -44.4, 0.8, 0.8); }
    ctx.fill(); ctx.stroke();
    if (i === 4 && S > 12) { ctx.setLineDash([5, 5]); ctx.strokeStyle = blueCol(k, 0.8 * a); ctx.beginPath(); for (let c = 0; c < 2; c++) rect(x0 + 5.5 + c * 7.5 - 1.5, -41.6, 3, 6.2); ctx.stroke(); ctx.setLineDash([]); }
    for (const [d, x, y, r, tone] of INDOOR) if (d === i) carPlan(ctx, cam, x, y, r, tone, k, a);
    // room name, set into the drawing
    if (S > 7 && S < 30) { ctx.font = M.font(Math.max(9, 1.15 * S), 500, K.MONO); ctx.letterSpacing = `${0.16 * S}px`; ctx.fillStyle = k < 0.5 ? rgba(C.cloud, 0.7 * a) : rgba(C.blue, 0.9 * a); const [tx, ty] = T(x0 + 1.3, B.y0 - 1.2); ctx.fillText(DEPTS[i].toUpperCase(), tx, ty); ctx.letterSpacing = '0px'; }
    ctx.restore();
  }
}

function lot(ctx, cam, k, a, heroA = 1) {
  const S = cam.s, T = (x, y) => toScreen(cam.cx, cam.cy, S, x, y);
  // lamp pools (night only)
  if (k < 0.5) for (const [x, y] of LAMPS) { const [px, py] = T(x, y), r = 13 * S; if (px < -r || px > W + r || py < -r || py > H + r) continue; const g = ctx.createRadialGradient(px, py, 0, px, py, r); g.addColorStop(0, rgba(C.cloud, 0.13 * a)); g.addColorStop(1, rgba(C.cloud, 0)); ctx.fillStyle = g; ctx.fillRect(px - r, py - r, 2 * r, 2 * r); }
  // site boundary, road, bay ticks
  ctx.lineWidth = 1; ctx.strokeStyle = lineCol(k, 0.3 * a); ctx.beginPath();
  { const [x0, y0] = T(-85, -50); ctx.rect(x0, y0, 170 * S, 100 * S); }
  for (let n = -4; n <= 4; n++) { const cxm = 16.5 * n + 2.6; const a0 = T(cxm, -13.5), a1 = T(cxm, 43.2); ctx.moveTo(a0[0], a0[1]); ctx.lineTo(a1[0], a1[1]); if (S > 5) for (let m = -9; m <= 12; m++) { const b0 = T(cxm - 5.1, 10.65 + 2.7 * m), b1 = T(cxm + 5.1, 10.65 + 2.7 * m); ctx.moveTo(b0[0], b0[1]); ctx.lineTo(b1[0], b1[1]); } }
  ctx.stroke();
  ctx.setLineDash([S * 2, S * 2]); ctx.strokeStyle = lineCol(k, 0.22 * a); ctx.beginPath(); { const a0 = T(-85, -19), a1 = T(85, -19); ctx.moveTo(a0[0], a0[1]); ctx.lineTo(a1[0], a1[1]); } ctx.stroke(); ctx.setLineDash([]);
  // trees: two loose rings, like a landscape plan
  ctx.strokeStyle = k < 0.5 ? rgba(C.teal, 0.4 * a) : rgba(C.teal, 0.8 * a); ctx.beginPath();
  for (const [x, y, r] of TREES) { const [px, py] = T(x, y); ctx.moveTo(px + r * S, py); ctx.arc(px, py, r * S, 0, 7); ctx.moveTo(px + r * S * 0.45, py); ctx.arc(px, py, r * S * 0.45, 0, 7); }
  ctx.stroke();
  for (const c of LOT) carPlan(ctx, cam, c.x, c.y, c.rot, c.tone, k, a * (c.hero ? heroA : 1));
  ctx.fillStyle = lineCol(k, 0.8 * a); ctx.beginPath(); for (const [x, y] of LAMPS) { const [px, py] = T(x, y); ctx.moveTo(px, py); ctx.arc(px, py, Math.max(1.2, 0.22 * S), 0, 7); } ctx.fill();
}
// the public site drawing (used by the ending on a bright sheet)
export function site(ctx, { k = 0, cx = 0, cy = 0, s = 10, alpha = 1 } = {}) {
  if (alpha <= 0) return;
  const cam = { cx, cy, s };
  ctx.save(); ctx.globalAlpha *= alpha; lot(ctx, cam, k, 1); rooms(ctx, cam, k, 1, null); ctx.restore();
}

// ------------------------------------------------------------------ the hero car in side elevation
// elevation coords: x forward, y up from the ground; world = (x, GROUND - y)
export const GROUND = 12.72;
export const EL = (() => {
  const FA = [1.45, 0.345], RA = [-1.42, 0.345], AR = 0.41;
  const body = [
    ...K.bez([[-1.83, 0.3], [-2.22, 0.3], [-2.36, 0.34, -2.37, 0.5, -2.36, 0.62], [-2.35, 0.75, -2.33, 0.82, -2.3, 0.9], [-2.26, 0.99, -2.18, 1.04, -2.05, 1.08],
      [-1.75, 1.24, -1.45, 1.4, -0.85, 1.455], [-0.45, 1.49, 0.05, 1.47, 0.38, 1.42], [0.62, 1.33, 0.95, 1.13, 1.16, 1.02], [1.5, 0.97, 1.9, 0.93, 2.14, 0.86],
      [2.28, 0.82, 2.34, 0.74, 2.35, 0.64], [2.37, 0.5, 2.36, 0.36, 2.26, 0.3], [1.86, 0.3]]),
    ...arc(FA[0], FA[1], AR, -0.11, Math.PI + 0.11), [-1.01, 0.3], ...arc(RA[0], RA[1], AR, -0.11, Math.PI + 0.11),
  ];
  const glass = K.bez([[1.02, 1.0], [0.85, 1.1, 0.6, 1.3, 0.36, 1.365], [0.0, 1.41, -0.5, 1.42, -0.85, 1.39], [-1.2, 1.36, -1.55, 1.2, -1.72, 1.06], [-1.74, 1.04, -1.7, 1.03, -1.66, 1.03], [1.02, 1.0]]);
  const lines = [
    K.bez([[2.2, 0.8], [1.2, 0.9, -1.0, 0.93, -2.28, 0.9]], 30),                                   // shoulder line
    K.bez([[1.0, 1.0], [1.03, 0.7, 0.99, 0.45, 0.93, 0.31]]), [[-0.15, 1.015], [-0.15, 0.31]],       // door cuts
    K.bez([[-1.3, 1.035], [-1.28, 0.85, -1.12, 0.8, -1.0, 0.62]]), [[0.93, 0.34], [-0.97, 0.34]],
    [[-0.1, 1.02], [-0.1, 1.405]], [[-0.2, 1.02], [-0.2, 1.41]],                                      // B pillar
    [[2.12, 0.84], [2.3, 0.765], [2.33, 0.7], [2.05, 0.78], [2.12, 0.84]],                            // head lamp
    [[-2.3, 0.9], [-2.345, 0.8], [-2.1, 0.84], [-2.08, 0.93], [-2.3, 0.9]],                           // tail lamp
    [[0.55, 0.93], [0.72, 0.93]], [[-0.62, 0.945], [-0.45, 0.945]],                                   // handles
    arc(0.95, 1.045, 0.075, 0, Math.PI * 2, 16),                                                      // mirror
  ];
  const wheels = [FA, RA].map(([x, y]) => ({ x, y, rings: [arc(x, y, 0.345, 0, Math.PI * 2, 40), arc(x, y, 0.25, 0, Math.PI * 2, 32), arc(x, y, 0.05, 0, Math.PI * 2, 10)],
    spokes: Array.from({ length: 5 }, (_, i) => { const a = (i / 5) * Math.PI * 2 - 0.3; return [[x + 0.05 * Math.cos(a), y + 0.05 * Math.sin(a)], [x + 0.25 * Math.cos(a + 0.12), y + 0.25 * Math.sin(a + 0.12)]]; }) }));
  return { body, glass, lines, wheels, FA, RA };
})();
const DOT = { x: 0.2, y: 1.215 };                // the red point sits on the front side glass
const HISTORY = [[-1.9, '12.01', 'INNKJØP'], [-0.95, '02.02', 'KLARGJØRING'], [0.2, '14.03', 'SALG'], [1.1, '20.09', 'SERVICE'], [1.95, '03.11', 'GARANTI']];

function elevation(ctx, cam, t, a) {
  if (a <= 0) return;
  const S = cam.s, P = ([x, y]) => toScreen(cam.cx, cam.cy, S, x, GROUND - y), path = (pts) => pts.map(P);
  const p = (t0, d = 1.6) => E.inOutCubic(prog(t, t0, t0 + d));
  ctx.save(); ctx.globalAlpha *= a; ctx.lineJoin = 'round'; ctx.lineCap = 'round';
  // body tone
  const bp = path(EL.body), fillA = ss(4.2, 6.2, t);
  if (fillA > 0) {
    const [, yT] = P([0, 1.5]), [, yB] = P([0, 0.25]); const g = ctx.createLinearGradient(0, yT, 0, yB);
    g.addColorStop(0, rgba(C.cloud, 0.2 * fillA)); g.addColorStop(0.55, rgba(C.cloud, 0.09 * fillA)); g.addColorStop(1, rgba(C.cloud, 0.02 * fillA));
    ctx.fillStyle = g; ctx.beginPath(); bp.forEach(([x, y], i) => (i ? ctx.lineTo(x, y) : ctx.moveTo(x, y))); ctx.closePath(); ctx.fill();
  }
  // glass is there from the first frame: the ledger is written on it
  const gp = path(EL.glass); ctx.fillStyle = rgba(C.light, 0.13); ctx.beginPath(); gp.forEach(([x, y], i) => (i ? ctx.lineTo(x, y) : ctx.moveTo(x, y))); ctx.fill();
  ctx.strokeStyle = rgba(C.cloud, 0.95); ctx.lineWidth = clamp(S / 150, 1.4, 2.6);
  K.drawOn(ctx, gp, p(2.6, 1.4));
  K.drawOn(ctx, bp, p(3.0, 2.6), true);
  ctx.lineWidth = clamp(S / 220, 1, 1.8); ctx.strokeStyle = rgba(C.cloud, 0.6);
  EL.lines.forEach((l, i) => K.drawOn(ctx, path(l), p(3.6 + i * 0.16, 1.2)));
  for (const w of EL.wheels) {
    const wp = p(4.0, 1.4);
    ctx.fillStyle = rgba(C.night, 0.9 * wp); ctx.beginPath(); path(w.rings[0]).forEach(([x, y], i) => (i ? ctx.lineTo(x, y) : ctx.moveTo(x, y))); ctx.fill();
    ctx.strokeStyle = rgba(C.cloud, 0.95); ctx.lineWidth = clamp(S / 150, 1.4, 2.4); K.drawOn(ctx, path(w.rings[0]), wp);
    ctx.strokeStyle = rgba(C.cloud, 0.6); ctx.lineWidth = clamp(S / 220, 1, 1.8); K.drawOn(ctx, path(w.rings[1]), p(4.3, 1.2)); K.drawOn(ctx, path(w.rings[2]), p(4.6, 0.8));
    w.spokes.forEach((sp, i) => K.drawOn(ctx, path(sp), p(4.7 + i * 0.08, 0.6)));
  }
  // drafting: ground line, wheel centre lines, wheelbase dimension
  ctx.strokeStyle = rgba(C.light, 0.55); ctx.lineWidth = 1;
  K.drawOn(ctx, path([[-3.6, 0], [3.6, 0]]), p(3.2, 2));
  ctx.setLineDash([8, 6]);
  for (const [x] of [EL.FA, EL.RA]) K.drawOn(ctx, path([[x, -0.42], [x, 1.72]]), p(5, 1));
  ctx.setLineDash([]);
  const dp = p(5.4, 1); K.drawOn(ctx, path([[EL.RA[0], -0.3], [EL.FA[0], -0.3]]), dp);
  if (dp > 0.6 && S > 90) { ctx.fillStyle = rgba(C.light, 0.9 * ss(0.6, 1, dp)); ctx.font = M.font(0.085 * S, 500, K.MONO); ctx.textAlign = 'center'; const [tx, ty] = P([0.015, -0.42]); ctx.fillText('2 870', tx, ty); ctx.textAlign = 'left'; }
  // the car's history, set out above the roof
  const hp = p(5.0, 1.8) * ss(125, 200, S);
  if (hp > 0) {
    ctx.strokeStyle = rgba(C.light, 0.7); K.drawOn(ctx, path([[-2.3, 1.95], [2.35, 1.95]]), hp);
    ctx.font = M.font(0.088 * S, 500, K.MONO);
    HISTORY.forEach(([x, date, label], i) => {
      const q = E.outCubic(prog(hp, 0.15 + i * 0.14, 0.45 + i * 0.14)); if (q <= 0) return;
      const main = i === 2, [px, py] = P([x, 1.95]);
      ctx.strokeStyle = main ? C.red : rgba(C.light, 0.8); ctx.lineWidth = main ? 2 : 1;
      ctx.beginPath(); ctx.moveTo(px, py - 0.07 * S * q); ctx.lineTo(px, py + 0.07 * S * q); ctx.stroke();
      ctx.globalAlpha = a * q; ctx.fillStyle = main ? C.white : rgba(C.cloud, 0.75);
      ctx.fillText(date, px - 0.02 * S, py - 0.3 * S); ctx.fillStyle = main ? C.white : rgba(C.light, 0.95); ctx.fillText(label, px - 0.02 * S, py - 0.16 * S);
      ctx.globalAlpha = a;
      if (main) { ctx.strokeStyle = rgba(C.red, 0.9); ctx.lineWidth = 1.5; K.drawOn(ctx, [P([x, 1.88]), P([DOT.x, DOT.y + 0.03])], q); }
    });
  }
  ctx.restore();
}

// the ledger: six lines of this car's bookkeeping, written small on the glass next to the red point
const LEDGER = [['BILAG', 'DATO', 'TEKST', 'BELØP'], ['118212', '12.01.2024', 'INNKJØP NYBIL', '412 300,00'], ['118240', '02.02.2024', 'KLARGJØRING', '8 450,00'],
  ['118432', '14.03.2024', 'SALG NYBIL', '489 900,00'], ['118433', '14.03.2024', 'INNBYTTE', '−96 000,00'], ['121874', '20.09.2024', 'SERVICE 30 000 KM', '6 180,00']];
function ledger(ctx, cam, t, a) {
  if (a <= 0) return;
  const S = cam.s, fs = 0.0107, pitch = 0.032, [ox, oy] = toScreen(cam.cx, cam.cy, S, DOT.x, GROUND - DOT.y);
  ctx.save(); ctx.globalAlpha *= a; ctx.font = M.font(fs * S, 500, K.MONO); ctx.textBaseline = 'middle';
  const cols = [0.022, 0.066, 0.142, 0.33].map((c) => c * S);
  LEDGER.forEach((row, i) => {
    const main = i === 3, y = oy + (i - 3) * pitch * S, typed = main ? 1 : ss(0.3 + i * 0.25, 0.9 + i * 0.25, t);
    ctx.fillStyle = main ? C.white : i === 0 ? rgba(C.light, 0.75 * typed) : rgba(C.cloud, 0.42 * typed);
    row.forEach((cell, c) => { ctx.textAlign = c === 3 ? 'right' : 'left'; ctx.fillText(cell, ox + cols[c], y - (1 - typed) * 0.004 * S); });
    if (i === 0) { ctx.fillStyle = rgba(C.light, 0.4 * typed); ctx.fillRect(ox + cols[0], y + pitch * 0.5 * S, (cols[3] - cols[0]) * typed, 1.2); }
  });
  // a hairline under the one transaction, drawn on with the first beat
  ctx.fillStyle = C.red; ctx.fillRect(ox + cols[0], oy + 0.0095 * S, (cols[3] - cols[0]) * E.outQuint(prog(t, 0.15, 1.3)), Math.max(1.5, 0.0006 * S));
  ctx.restore();
}

// ------------------------------------------------------------------ where each data point lives in the world
const SRC = PTS.map((p) => {
  if (p.i === 0) return [HERO.x + DOT.x, GROUND - DOT.y, 0];
  if (p.c === 6) { const c = LOT[Math.floor(p.seed * LOT.length)]; return [c.x + (p.z - 0.5) * 2.2, c.y + (K.hash01(p.i, 4) - 0.5) * 1.2, c.d]; }
  const gx = Math.floor(p.seed * 24), gy = Math.floor(K.hash01(p.i, 9) * 15); return [B.x0 + B.w * p.c + 1.45 + gx * 0.7, B.y0 + 1.6 + gy * 1.3 + (gx % 2) * 0.65, 0];
});
const LOTLINKS = (() => { const ids = PTS.filter((p) => p.c === 6 && p.i).map((p) => p.i), r = K.mulberry32(21), out = [];
  for (let n = 0; n < 230; n++) { const a = ids[Math.floor(r() * ids.length)]; let best = -1, bd = 1e9; for (let q = 0; q < 40; q++) { const b = ids[Math.floor(r() * ids.length)], d = Math.hypot(SRC[a][0] - SRC[b][0], SRC[a][1] - SRC[b][1]); if (b !== a && d > 6 && d < bd) { bd = d; best = b; } } if (best >= 0 && bd < 26) out.push([a, best]); }
  return out; })();

// ------------------------------------------------------------------ camera
const LN = Math.log;
const zoom1 = curve([[0, LN(5000)], [2.4, LN(4400)], [4.6, LN(1300)], [7.2, LN(330)], [9.2, LN(240)], [11.5, LN(60)], [13.6, LN(24)], [17.4, LN(10.6)], [20, LN(10)]]);
const A0 = [DOT.x + (W / 2 - 150) / 5000, GROUND - DOT.y + 100 / 5000], A1 = [0, 11.86], A2 = [0, -1];
const T_IN = 20, T_GL0 = 22.6, T_GL1 = 38.6, T_OUT = 41.8;
const deptT = (i) => T_GL0 + ((i + 0.5) / 6) * (T_GL1 - T_GL0);
const camX2 = curve([[T_IN, 0], [T_GL0, -47.5], [T_GL1, 47.5], [T_OUT, 39.3], [46, 39.3], [49.6, 0]]);
const camY2 = curve([[T_IN, -1], [T_GL0, -35.5], [T_GL1, -35.5], [T_OUT, -1], [46, -1], [49.6, 0]]);
const camS2 = curve([[T_IN, LN(10)], [T_GL0, LN(56)], [30.6, LN(60)], [T_GL1, LN(56)], [T_OUT, LN(7.4)], [46, LN(7.3)], [49.6, LN(10.5)]]);
function camera(t) {
  if (t < T_IN) {
    const ls = zoom1(t), s = Math.exp(ls), w1 = ss(LN(4300), LN(420), ls), w2 = ss(LN(210), LN(11), ls);
    return { s, cx: lerp(lerp(A0[0], A1[0], w1), A2[0], w2) + 0.0012 * Math.sin(t * 0.9) * (1 - w1), cy: lerp(lerp(A0[1], A1[1], w1), A2[1], w2) };
  }
  return { s: Math.exp(camS2(t)), cx: camX2(t), cy: camY2(t) + 0.25 * Math.sin(t * 0.7) };
}
// the red point's world position after it leaves the hero car
const RED_PATH = K.bez([[HERO.x + DOT.x, GROUND - DOT.y], [-6, -6, -30, -14, -44, -22.5], [-47, -26, -47.5, -28, -45, -29.6]], 24);
function redWorld(t) {
  if (t < T_IN + 0.2) return RED_PATH[0];
  if (t < T_GL0) { const f = E.inOutCubic(prog(t, T_IN + 0.2, T_GL0)) * (RED_PATH.length - 1), i = Math.min(RED_PATH.length - 2, Math.floor(f)); return [lerp(RED_PATH[i][0], RED_PATH[i + 1][0], f - i), lerp(RED_PATH[i][1], RED_PATH[i + 1][1], f - i)]; }
  return [Math.min(52, camX2(Math.min(t, T_GL1)) + 2.5 + 2.5 * prog(t, T_GL0, T_GL0 + 2)), -29.6];
}

// ------------------------------------------------------------------ BDO Bilbransjeanalysen 2025 (real figures only)
function bransje(ctx, t) {
  const a = win(t, 40.2, 48.4, 0.5, 0.7); if (a <= 0) return;
  const x = 1352, out = 48.2;
  K.kicker(ctx, 'BDO Bilbransjeanalysen 2025', x, 232, t, { t0: 40.2, out, color: C.cloud });
  const item = (y, t0, big, label, small) => {
    K.line(ctx, big, x, y, t, { size: 104, weight: 300, t0, out, color: C.white, track: -0.02 });
    const q = ss(t0 + 0.3, t0 + 0.9, t) * (1 - ss(out - 0.4, out, t)); if (q <= 0) return;
    ctx.save(); ctx.globalAlpha = q; ctx.font = M.font(17, 500, K.MONO); ctx.letterSpacing = '2px'; ctx.fillStyle = C.light; ctx.fillText(label, x, y + 40);
    if (small) { ctx.fillStyle = rgba(C.cloud, 0.75); ctx.fillText(small, x, y + 70); }
    ctx.restore();
  };
  item(378, 40.7, '1,7 %', 'DRIFTSMARGIN · MERKEFORHANDLERE 2024', 'OPP FRA 1,5 % I 2023');
  item(596, 41.9, '2,5 %', 'DRIFTSMARGIN · BRUKTBILFORHANDLERE 2024');
  item(790, 43.1, '1 300+', 'BILBRANSJEBEDRIFTER ANALYSERT');
  // the 2023 -> 2024 step as two drawn bars beside the first figure
  const bp = E.outCubic(prog(t, 41.1, 42.2)) * (1 - ss(out - 0.4, out, t));
  if (bp > 0) { ctx.fillStyle = rgba(C.light, 0.7); ctx.fillRect(x + 300, 378 - 75 * bp, 10, 75 * bp); ctx.fillStyle = C.white; ctx.fillRect(x + 322, 378 - 85 * bp, 10, 85 * bp); }
  ctx.strokeStyle = rgba(C.light, 0.5 * a); ctx.lineWidth = 1; K.drawOn(ctx, [[x - 34, 200], [x - 34, 830]], E.outCubic(prog(t, 40.2, 41.6)));
}

// ------------------------------------------------------------------ the scene
export const sheetAt = () => 0;
export const HITS = [
  [0, 'one transaction', 'blip', { pitch: 'D5' }], [0.6, 'caption', 'tick'], [2.6, 'pull out', 'whoosh', { len: 2.2, from: 300, to: 1400 }],
  [6, 'one car', 'pop', { pitch: 'A4' }], [9.2, 'pull out 2', 'whoosh', { len: 2.4, from: 400, to: 1800 }], [10.5, 'many connections', 'blip', { pitch: 'F#5' }],
  [14, 'the whole business', 'sub', { len: 1.2 }], [17.2, 'quote', 'tick'], [20, 'into the building', 'whoosh', { len: 2.2, from: 1600, to: 500 }],
  ...DEPTS.map((d, i) => [+(deptT(i) - 0.95).toFixed(2), d, 'blip', { pitch: ['D5', 'E5', 'F#5', 'A5', 'B5', 'D6'][i] }]),
  [38.6, 'pull out', 'whoosh', { len: 2.6, from: 500, to: 1800 }], [40.7, 'bransje 1', 'tick'], [41.9, 'bransje 2', 'tick'], [43.1, 'bransje 3', 'tick'],
  [46.2, 'data lifts', 'swell', { len: 2.5 }],
];

export function draw(ctx, t) {
  const cam = camera(t), S = cam.s;
  K.sheet(ctx, 0);
  K.grid(ctx, 0, 1, -cam.cx * S * 0.12, -cam.cy * S * 0.12);
  const rise = (p) => E.inOutCubic(prog(t, 46.3 + p.seed * 1.3, 48.3 + p.seed * 1.3));
  const siteA = (t < T_IN ? ss(LN(175), LN(80), LN(Math.min(S, 5000))) : 1) * (1 - ss(46.4, 49.4, t)) * (t > 39.5 ? lerp(1, 0.55, ss(39.5, 42, t)) : 1);

  // the plan drawing
  if (siteA > 0) {
    ctx.save(); ctx.globalAlpha = siteA;
    lot(ctx, cam, 0, t < T_IN ? 1 : lerp(1, 0.5, ss(T_IN, T_GL0, t) * (1 - ss(T_GL1, T_OUT, t))), t < T_IN ? 1 - ss(LN(72), LN(150), LN(S)) : 1);
    // each department wakes as the camera reaches it
    const reveal = (i) => (t < T_IN + 1 || t > T_GL1 ? 1 : lerp(0.3, 1, ss(B.x0 + B.w * i - 13, B.x0 + B.w * i - 3, cam.cx + 16) ));
    rooms(ctx, cam, 0, 1, reveal);
    ctx.restore();
  }
  // the red trail through the business
  if (t > T_IN && siteA > 0) {
    ctx.save(); ctx.globalAlpha = siteA; ctx.strokeStyle = C.red; ctx.lineWidth = 2; ctx.lineJoin = 'round';
    const rw = redWorld(t), pts = [...RED_PATH.slice(0, Math.max(2, Math.ceil(E.inOutCubic(prog(t, T_IN + 0.2, T_GL0)) * RED_PATH.length))), ...(t >= T_GL0 ? [rw] : [])];
    K.drawOn(ctx, pts.map(([x, y]) => toScreen(cam.cx, cam.cy, S, x, y)), 1);
    ctx.restore();
  }
  // the hero car as a side elevation, and its ledger
  const elA = ss(LN(72), LN(150), LN(S)) * (t < T_IN ? 1 : 0);
  if (elA > 0) {
    // clear the plan behind the elevation while both are on the sheet
    elevation(ctx, cam, t, elA);
  }
  ledger(ctx, cam, t, ss(LN(520), LN(1700), LN(S)));

  // data points: over the cars, in the departments, then lifting into the network
  if (t > 9.8) {
    const k1 = ss(46.3, 49.6, t);
    K.field(ctx, 0, {
      links: 0, red: 0, alpha: 1,
      pos: (p) => {
        if (p.i === 0) return [0, 0, 0, 0];
        const [wx, wy, d] = SRC[p.i], on = p.c === 6 ? ss(10.2 + d / 34, 10.9 + d / 34, t) : (t < T_GL0 ? 0 : ss(-12, -2, cam.cx + 14 - (B.x0 + B.w * p.c + 9.5) + (p.seed - 0.5) * 8) + ss(T_GL1 - 0.5, T_GL1 + 1.5, t));
        if (on <= 0) return [0, 0, 0, 0];
        const [sx, sy] = toScreen(cam.cx, cam.cy, S, wx, wy - 0.9 * Math.min(1, on) * (0.4 + p.z)), r = rise(p);
        const bob = Math.sin(t * (0.7 + p.z) + p.seed * 40) * 2.5;
        return [lerp(sx, p.x, r), lerp(sy + bob, p.y, r), Math.min(1, on) * lerp(p.c === 6 && t < T_IN ? 0.95 : 0.8, 1, r), lerp(p.c === 6 ? clamp(S / 11, 0.9, 1.5) : clamp(S / 40, 0.7, 1.25), 1, r)];
      },
    });
    // links between related cars (act 1), then the network's own links at the hand-off
    const la = ss(11.2, 13, t) * (1 - ss(19, 21.5, t));
    if (la > 0) {
      ctx.strokeStyle = rgba(C.light, 0.6 * la); ctx.lineWidth = 1; ctx.beginPath();
      for (const [a, b] of LOTLINKS) {
        if (ss(10.9 + SRC[a][2] / 34, 11.6 + SRC[a][2] / 34, t) < 1) continue;
        const A = toScreen(cam.cx, cam.cy, S, SRC[a][0], SRC[a][1] - 0.9), Bp = toScreen(cam.cx, cam.cy, S, SRC[b][0], SRC[b][1] - 0.9);
        ctx.moveTo(A[0], A[1]); ctx.quadraticCurveTo((A[0] + Bp[0]) / 2, (A[1] + Bp[1]) / 2 - Math.hypot(A[0] - Bp[0], A[1] - Bp[1]) * 0.22, Bp[0], Bp[1]);
      }
      ctx.stroke();
    }
    if (k1 > 0) K.field(ctx, 0, { alpha: 0, links: ss(48.6, 49.9, t), red: 0, pos: (p) => [p.x, p.y, ss(48.3, 49.6, t), 1] });
  }
  // the red point
  {
    const rw = redWorld(t), [rx, ry] = toScreen(cam.cx, cam.cy, S, rw[0], rw[1]), r = E.inOutCubic(prog(t, 46.4, 49.5));
    const pulse = 1 + 0.18 * Math.max(0, Math.sin(t * Math.PI * 3.2)) * (t < 3 ? 1 : 0);
    K.redDot(ctx, lerp(rx, PTS[0].x, r), lerp(ry, PTS[0].y, r), lerp(t < 5 ? clamp(S * 0.004, 7, 22) : 7, 7, r) * pulse);
  }

  // department names while we travel through the building
  if (t > T_GL0 - 0.5 && t < T_GL1 + 0.6) {
    const span = (T_GL1 - T_GL0) / 6;
    DEPTS.forEach((d, i) => {
      const t0 = T_GL0 + i * span - 0.25, out = t0 + span + 0.05;
      K.line(ctx, d.toUpperCase(), K.CAP.x, K.CAP.y, t, { size: 120, weight: 600, t0, out, dur: 0.55, track: 0.01 });
      K.kicker(ctx, DEPT_DATA[i], K.CAP.x, K.CAP.y - 135, t, { t0: t0 + 0.15, out: out - 0.1, color: C.cloud });
    });
    // travel indicator: six ticks, the current one red
    const prg = prog(t, T_GL0, T_GL1) * 6;
    for (let i = 0; i < 6; i++) { ctx.fillStyle = Math.floor(prg) === i ? C.red : rgba(C.cloud, 0.4); ctx.fillRect(W - 120 - (6 - i) * 34, 928, 22, 3); }
  }
  // a quiet scrim so captions stay readable over the drawing
  { const g = ctx.createLinearGradient(0, H - 330, 0, H); g.addColorStop(0, rgba(C.night, 0)); g.addColorStop(0.55, rgba(C.night, 0.8)); g.addColorStop(1, rgba(C.night, 0.92)); const sa = Math.max(siteA, 0) * (t > 9.5 ? 1 : 0); if (sa > 0) { ctx.globalAlpha = sa; ctx.fillStyle = g; ctx.fillRect(0, H - 330, W, 330); ctx.globalAlpha = 1; } }
  // captions
  const cap = (str, t0, out, y = K.CAP.y, o = {}) => K.line(ctx, str, K.CAP.x, y, t, { size: 76, weight: 600, t0, out, track: 0.01, ...o });
  cap('ÉN TRANSAKSJON.', 0.5, 4.4);
  cap('ÉN BIL.', 6, 9.4);
  cap('MANGE SAMMENHENGER.', 10.6, 13.6);
  cap('EN HEL VIRKSOMHET.', 14.1, 17);
  cap('Hver bil har en historikk.', 17.3, 21.2, 856, { weight: 300, size: 62 });
  cap('Hver transaksjon etterlater spor.', 17.9, 21.4, 930, { weight: 300, size: 62 });
  cap('Data gir oss informasjon.', 40.9, 48.7, 856, { weight: 300, size: 62 });
  cap('Bransjekunnskap gir oss forståelse.', 43.9, 48.9, 930, { weight: 600, size: 62 });
  bransje(ctx, t);
}
