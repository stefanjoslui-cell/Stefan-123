// BDO Automotive – Tallene forteller. Editorial data story, 75 s, 96 bpm. Pure function of time.
//  1  0–7.5   one number -> a ledger wall -> collapses to a horizon line
//  2  7.5–17  one car's life as a timeline -> 300 timelines -> collapses to a horizon
//  3  17–30   the dealership at night (dark interlude), six departments, illustrative record counts
//  4  30–42.5 BDO Bilbransjeanalysen 2025: bars, then 1 300 dots
//  5  42.5–55 dots -> scatter of transactions with bands, three points flagged
//  6  55–67.5 two-column spread, then the statement
//  7  67.5–75 end card
import * as M from './lib/motion.js';
import * as K from './lib/kit.js';
import { site } from './scenes/world.js';

const { W, H, C, clamp, lerp, prog, E, ss, rgba, hash01 } = K;
const FD = K.DISPLAY, FM = K.MONO, PERIOD = 0.625;
const INK = C.graphite;

// ------------------------------------------------------------------ helpers
function tabText(ctx, str, x, y, size, weight, col, align = 'left', fam = FD) {
  ctx.font = M.font(size, weight, fam); ctx.fillStyle = col; ctx.textAlign = 'left';
  const dw = ctx.measureText('0').width, chars = [...str];
  const adv = chars.map((c) => (c >= '0' && c <= '9' ? dw : ctx.measureText(c).width));
  const tot = adv.reduce((a, b) => a + b, 0);
  let px = align === 'right' ? x - tot : align === 'center' ? x - tot / 2 : x;
  chars.forEach((c, i) => {
    if (c >= '0' && c <= '9') { ctx.textAlign = 'center'; ctx.fillText(c, px + dw / 2, y); ctx.textAlign = 'left'; } else ctx.fillText(c, px, y);
    px += adv[i];
  });
  return tot;
}
function tabW(ctx, str, size, weight, fam = FD) {
  ctx.font = M.font(size, weight, fam); const dw = ctx.measureText('0').width;
  return [...str].reduce((a, c) => a + (c >= '0' && c <= '9' ? dw : ctx.measureText(c).width), 0);
}
const fmt = (n, dec = 0) => { const s = Math.abs(n).toFixed(dec); let [a, b] = s.split('.'); a = a.replace(/\B(?=(\d{3})+(?!\d))/g, ' '); return b ? `${a},${b}` : a; };
// generic rise-from-baseline mask for custom drawing
function masked(ctx, t, t0, out, y, size, fn, dur = 0.7) {
  const pi = E.outQuint(prog(t, t0, t0 + dur)), po = E.inCubic(prog(t, out - 0.45, out));
  if (pi <= 0 || po >= 1) return;
  ctx.save(); ctx.beginPath(); ctx.rect(0, y - size * 1.05, W, size * 1.4); ctx.clip();
  ctx.globalAlpha *= 1 - po; ctx.translate(0, (1 - pi) * size * 1.15 - po * size * 0.5); fn(); ctx.restore();
}
function txt(ctx, s, x, y, size, weight, col, fam = FM, align = 'left', track = 0) {
  ctx.save(); ctx.font = M.font(size, weight, fam); ctx.fillStyle = col; ctx.textAlign = align; ctx.letterSpacing = `${track * size}px`; ctx.fillText(s, x, y); ctx.restore();
}
const cap = (ctx, t, s, t0, out, col = INK) => K.line(ctx, s, 160, 930, t, { size: 56, weight: 600, color: col, t0, out });
function bright(ctx, t) { K.sheet(ctx, 1); K.grid(ctx, 1, 0.9, t * 5, t * 2.5); }
const hl = (ctx, y, x0, x1, col, lw = 1) => { ctx.strokeStyle = col; ctx.lineWidth = lw; ctx.beginPath(); ctx.moveTo(x0, y); ctx.lineTo(x1, y); ctx.stroke(); };

// ================================================================== 1. THE NUMBER
const LROWS = 40, RH = 36, BX = [110, 540, 970, 1400];
const TXTS = ['Bilsalg', 'Deler uttak', 'Arbeidsordre', 'Garantikrav', 'Innbytte', 'Lagerføring', 'Serviceordre', 'Leasingfaktura'];
const SPECIAL = { b: 1, R: 11, y0: 440 };   // block 1, absolute row 11, at y0 when scroll is 0 (set so it sits at ~y 270 at t=4.4)
const scroll1 = (t) => 70 * Math.max(0, t - 2.2);
function rowStr(b, R) {
  const h = (k) => hash01(R * 7 + b * 131, k);
  const no = String(10000 + Math.floor(h(1) * 89999)), d = String(1 + Math.floor(h(2) * 28)).padStart(2, '0') + '.' + String(1 + Math.floor(h(3) * 12)).padStart(2, '0') + '.';
  const tx = TXTS[Math.floor(h(4) * TXTS.length)].padEnd(14, ' ');
  const amt = Math.floor(Math.pow(10, 3 + h(5) * 2.9)); return { head: `${no} ${d} ${tx}`, amt: fmt(amt) + ',00' };
}
function s1(ctx, t) {
  bright(ctx, t);
  const q = E.inOutCubic(prog(t, 5.8, 7.3)), sc = scroll1(t), n = 489900;
  // ledger wall
  if (t > 2.3) {
    ctx.save();
    ctx.translate(0, 540); ctx.scale(1, 1 - q * 0.99); ctx.translate(0, -540);
    ctx.font = M.font(17, 400, FM); ctx.textAlign = 'left';
    const first = Math.floor(sc / RH) - 1;
    const sy = 440 + 0; // special row screen y (t-dependent below)
    for (let r = 0; r < LROWS + 3; r++) {
      const R = first + r, y = 40 + R * RH - sc + 14;
      if (y < -20 || y > 1000) continue;
      for (let b = 0; b < 4; b++) {
        const dist = Math.hypot((BX[b] + 220) - (BX[SPECIAL.b] + 220), y - 300) * 0.9;
        const a = clamp((t - 2.5) * 1.3 - dist / 700) * (1 - 0.7 * q);
        if (a <= 0.01) continue;
        const rs = rowStr(b, R), isSp = b === SPECIAL.b && R === SPECIAL.R;
        ctx.fillStyle = rgba(isSp ? C.blue : C.light, (isSp ? 1 : 0.62) * a);
        ctx.fillText(rs.head, BX[b], y);
        if (!isSp) { ctx.textAlign = 'right'; ctx.fillText(rs.amt, BX[b] + 410, y); ctx.textAlign = 'left'; }
      }
    }
    ctx.restore();
  }
  // special row position (screen) in the ledger
  const spY = 40 + SPECIAL.R * RH - sc + 14, finalSize = 24, bigSize = 240;
  const m = E.inOutCubic(prog(t, 2.7, 4.5));
  const cnt = n * E.outExpo(prog(t, 0.0, 2.4)) * 0.84 + n * 0.16 * E.outCubic(prog(t, 0.0, 2.4));
  const val = t > 2.4 ? n : Math.round(cnt * 100) / 100;
  const str = fmt(val, 2);
  const size = lerp(bigSize, finalSize, m), wt = lerp(300, 500, m);
  const wFin = tabW(ctx, '489 900,00', finalSize, 500);
  const xFin = BX[SPECIAL.b] + 410 - wFin, yFin = spY;
  const x = lerp(160, xFin, m), y = lerp(560, yFin, m);
  const col = mix2(INK, C.blue, m);
  if (t < 7.0) {
    ctx.save(); ctx.globalAlpha *= 1 - E.inCubic(prog(t, 6.2, 7.0));
    // underline in red, glued to the number
    const w = tabW(ctx, '489 900,00', size, wt);
    tabText(ctx, str, x, y, size, wt, col, 'left');
    const hy = y + lerp(52, 5, m), hp = E.outQuint(prog(t, 0.5, 1.6));
    ctx.fillStyle = C.red; ctx.fillRect(x, hy, w * hp, lerp(3, 1.6, m));
    ctx.restore();
  }
  K.kicker(ctx, 'Én transaksjon', 160, 350, t, { color: C.blue, t0: 0.15, out: 3.0 });
  // fades at the ledger edges so type stays clean
  const g = ctx.createLinearGradient(0, 760, 0, 900); g.addColorStop(0, rgba(C.arctic, 0)); g.addColorStop(1, rgba(C.arctic, 1));
  if (t > 2.3 && q < 1) { ctx.save(); ctx.globalAlpha = 1 - q; ctx.fillStyle = g; ctx.fillRect(0, 760, W, 320); ctx.restore(); }
  cap(ctx, t, 'Én av mange tusen.', 4.3, 6.9);
  // horizon line the ledger becomes
  hl(ctx, 540, 0, W, rgba(C.light, 0.9 * ss(5.6, 6.6, t)), 1.2);
}
function mix2(a, b, t) { return M.mix(a, b, t); }

// ================================================================== 2. THE CAR'S LIFE
const AX = { x0: 200, x1: 1700, y: 540 };
const D0 = Date.UTC(2024, 0, 5), DAY = 86400000, dayX = (ms) => lerp(AX.x0, AX.x1, (ms - D0) / (Date.UTC(2025, 1, 18) - D0));
const EVENTS = [
  { n: 'Innkjøp', d: '05.01.2024', a: 'kr 412 000', x: dayX(Date.UTC(2024, 0, 5)), up: true },
  { n: 'Klargjøring', d: '29.01.2024', a: 'kr 9 800', x: dayX(Date.UTC(2024, 0, 29)), up: false },
  { n: 'Salg', d: '12.03.2024', a: 'kr 489 900', x: dayX(Date.UTC(2024, 2, 12)), up: true, hot: true, tall: true },
  { n: 'Service', d: '10.09.2024', a: 'kr 3 450', x: dayX(Date.UTC(2024, 8, 10)), up: true },
  { n: 'Garanti', d: '18.02.2025', a: 'kr 6 200', x: dayX(Date.UTC(2025, 1, 18)), up: false, right: true },
];
const NL = 300, HERO = 150;
const LN = Array.from({ length: NL }, (_, j) => {
  const x0 = 160 + hash01(j, 1) * 380, len = 500 + hash01(j, 2) * 1000, x1 = Math.min(1780, x0 + len);
  return { j, y: 540 + (j - HERO) * 2.3, x0, x1, ev: [0, 1, 2, 3, 4].map((k) => lerp(x0, x1, k === 0 ? 0 : hash01(j, 10 + k) * 0.95 + 0.05)), d: Math.abs(j - HERO) / HERO, h: hash01(j, 99) };
});
function s2(ctx, t) {
  bright(ctx, t);
  hl(ctx, 540, 0, W, rgba(C.light, 0.9), 1.2);
  const lx = lerp(AX.x0, AX.x1, E.inOutSine(prog(t, 8.0, 12.4)));
  const multi = prog(t, 13.3, 15.9), coll = E.inOutCubic(prog(t, 16.2, 17.1));
  // month ticks
  const tk = ctx.save() || 0;
  ctx.strokeStyle = rgba(C.light, 0.9 * (1 - ss(13.2, 13.7, t))); ctx.lineWidth = 1; ctx.beginPath();
  const labs = [];
  for (let k = 0; k < 14; k++) {
    const ms = Date.UTC(2024, k, 1); if (ms < D0 || ms > Date.UTC(2025, 1, 18)) continue;
    const x = dayX(ms); ctx.moveTo(x, 540 - 7); ctx.lineTo(x, 540 + 7);
    if (k % 3 === 0) labs.push([x, ['JAN', 'APR', 'JUL', 'OKT', 'JAN', 'APR'][k / 3] + (k === 12 ? ' 25' : '')]);
  }
  ctx.stroke();
  const ta = (1 - ss(13.2, 13.7, t)) * ss(8.0, 8.6, t);
  for (const [x, s] of labs) txt(ctx, s, x, 566, 18, 500, rgba(C.light, ta), FM, 'center', 0.1);
  ctx.restore();
  // dense field of other cars
  if (multi > 0 && coll < 1) {
    const camX = -34 * (t - 13.3);
    ctx.save(); ctx.translate(camX * (1 - coll), 0);
    ctx.lineWidth = 1;
    for (let pass = 0; pass < 2; pass++) {
      ctx.beginPath();
      for (const L of LN) {
        if (L.j === HERO) continue;
        const p = E.outCubic(prog(t, 13.3 + L.d * 0.9, 13.3 + L.d * 0.9 + 1.3)); if (p <= 0) continue;
        const y = lerp(540, L.y, p) , y2 = lerp(y, 540, coll);
        const grp = L.h < 0.5 ? 0 : 1; if (grp !== pass) continue;
        const xe = lerp(L.x0, L.x1, p);
        ctx.moveTo(L.x0, y2); ctx.lineTo(xe, y2);
        if (coll < 0.35) for (const ex of L.ev) { if (ex > xe) continue; ctx.moveTo(ex, y2 - 4); ctx.lineTo(ex, y2 + 4); }
      }
      ctx.strokeStyle = rgba(pass ? C.blue : C.light, (pass ? 0.5 : 0.7) * (1 - coll * 0.5)); ctx.stroke();
    }
    ctx.restore();
  }
  // the one car
  ctx.strokeStyle = C.blue; ctx.lineWidth = 3;
  ctx.beginPath(); ctx.moveTo(AX.x0, 540); ctx.lineTo(lx, 540); ctx.stroke();
  const fade = 1 - ss(13.3, 13.9, t);
  for (let i = 0; i < EVENTS.length; i++) {
    const e = EVENTS[i], p = E.outQuint(clamp((lx - e.x + 4) / 70)); if (p <= 0) continue;
    const dir = e.up ? -1 : 1, stemL = (e.tall ? 240 : e.up ? 100 : 95) * p;
    ctx.save(); ctx.globalAlpha *= fade;
    ctx.strokeStyle = e.hot ? C.red : C.blue; ctx.lineWidth = 1.5; ctx.beginPath(); ctx.moveTo(e.x, 540); ctx.lineTo(e.x, 540 + dir * stemL); ctx.stroke();
    ctx.fillStyle = e.hot ? C.red : C.white; ctx.strokeStyle = e.hot ? C.red : C.blue; ctx.lineWidth = 3;
    ctx.beginPath(); ctx.arc(e.x, 540, (e.hot ? 11 : 9) * (0.4 + 0.6 * p), 0, 7); ctx.fill(); ctx.stroke();
    const al = e.right ? 'right' : 'left', tx = e.right ? e.x + 0 : e.x - 0;
    const ys = e.tall ? [540 - 358, 540 - 288, 540 - 252] : e.up ? [540 - 218, 540 - 148, 540 - 112] : [540 + 150, 540 + 218, 540 + 254];
    ctx.globalAlpha *= p;
    ctx.save(); ctx.translate((e.right ? -1 : 1) * (1 - p) * 20, 0);
    txt(ctx, e.n.toUpperCase(), tx, ys[0], 24, 500, C.blue, FM, al, 0.12);
    tabText(ctx, e.a, tx, ys[1], 56, 300, e.hot ? C.red : INK, al);
    txt(ctx, e.d, tx, ys[2], 22, 400, C.light, FM, al, 0.04);
    ctx.restore(); ctx.restore();
  }
  // hero red accent persists into the dense field
  if (t > 13.3) { const sx = EVENTS[2].x + (-34 * (t - 13.3)) * (1 - coll); K.redDot(ctx, sx, 540, 6, (1 - coll) * 0.9 + 0.1 * 0); }
  K.kicker(ctx, 'Én bils livsløp · 2024–2025', 160, 130, t, { color: C.blue, t0: 8.0, out: 13.4 });
  K.kicker(ctx, 'Bilparken som data', 160, 130, t, { color: C.blue, t0: 14.0, out: 16.3 });
  cap(ctx, t, 'Hver bil har en historikk.', 11.9, 16.6);
}

// ================================================================== 3. THE DEALERSHIP AT NIGHT
const DEPTS = [
  ['Salg', 12400, 'ordrelinjer'], ['Lager', 9600, 'lagerposteringer'], ['Verksted', 38000, 'arbeidstimer'],
  ['Deler', 54000, 'delebevegelser'], ['Garanti', 2700, 'garantikrav'], ['Service', 21000, 'serviceordrer'],
];
const dT = (i) => 18.4 + i * 1.75;
function s3(ctx, t) {
  K.sheet(ctx, 0);
  const p = (t - 17) / 13, s = lerp(11.6, 13.4, E.inOutSine(clamp(p)));
  const cx = lerp(-34, 26, E.inOutSine(clamp(p))), cy = -35 - 150 / s;
  K.grid(ctx, 0, 0.8, -cx * 4, 0);
  site(ctx, { k: 0, cx, cy, s, alpha: 1 });
  // shade for type
  const g = ctx.createLinearGradient(0, 0, 0, 520); g.addColorStop(0, rgba(C.night, 0.75)); g.addColorStop(1, rgba(C.night, 0));
  ctx.fillStyle = g; ctx.fillRect(0, 0, W, 520);
  const g2 = ctx.createLinearGradient(0, 860, 0, 1080); g2.addColorStop(0, rgba(C.night, 0)); g2.addColorStop(1, rgba(C.night, 0.85));
  ctx.fillStyle = g2; ctx.fillRect(0, 860, W, 220);
  const T = (x, y) => [W / 2 + (x - cx) * s, H / 2 + (y - cy) * s];
  for (let i = 0; i < 6; i++) {
    const t0 = dT(i), out = i === 5 ? 29.2 : dT(i + 1) - 0.05;
    const a = ss(t0, t0 + 0.4, t) * (1 - ss(out - 0.35, out, t));
    if (a > 0) {
      const x0 = -57 + 19 * i, [px, py] = T(x0, -46), [qx, qy] = T(x0 + 19, -24);
      ctx.save(); ctx.globalAlpha *= a; ctx.fillStyle = rgba(C.light, 0.13); ctx.fillRect(px, py, qx - px, qy - py);
      ctx.strokeStyle = rgba(C.cloud, 0.95); ctx.lineWidth = 2; ctx.strokeRect(px, py, qx - px, qy - py);
      // leader from room to the figure
      const lxp = clamp((px + qx) / 2, 400, 1500); ctx.strokeStyle = rgba(C.cloud, 0.6); ctx.lineWidth = 1.2; ctx.beginPath(); ctx.moveTo((px + qx) / 2, py); ctx.lineTo((px + qx) / 2, 560); ctx.stroke();
      ctx.restore();
    }
    const [name, v, lab] = DEPTS[i];
    K.kicker(ctx, `${name} · illustrativt tall`, 160, 188, t, { color: C.cloud, t0, out });
    const val = Math.round(v * E.outQuint(prog(t, t0 + 0.2, t0 + 1.25)));
    masked(ctx, t, t0 + 0.1, out, 392, 190, () => tabText(ctx, fmt(val), 160, 392, 190, 800, C.white, 'left'));
    K.line(ctx, lab, 164, 470, t, { size: 52, weight: 300, color: C.cloud, t0: t0 + 0.3, out });
  }
  cap(ctx, t, 'Seks avdelinger. Tusenvis av datapunkter.', 24.0, 29.3, C.white);
  K.kicker(ctx, 'Anlegget · plan', 1500, 120, t, { color: C.light, t0: 17.9, out: 29.0, tick: false });
  // the axis that becomes the next sheet's y-axis
  if (t > 29.0) { ctx.save(); ctx.strokeStyle = rgba(C.light, 0.9); ctx.lineWidth = 2; K.drawOn(ctx, [[240, 790], [240, 250]], E.outCubic(prog(t, 29.0, 29.45))); ctx.restore(); }
}

// ================================================================== 4. BRANSJEKUNNSKAP
const MX = (i) => 830 + (i % 45) * 20, MY = (i) => 230 + Math.floor(i / 45) * 20;
const NDOT = 1300, REDI = 1259;
const BAR = [{ v: 1.5, x: 300, lab: '2023', col: C.light, t0: 32.0 }, { v: 1.7, x: 480, lab: '2024', col: C.blue, t0: 33.3 }, { v: 2.5, x: 660, lab: '2024', col: C.teal, t0: 35.5 }];
const BASE = 790, SC = 200;
function s4(ctx, t) {
  bright(ctx, t);
  K.kicker(ctx, 'BDO Bilbransjeanalysen 2025', 160, 120, t, { color: C.blue, t0: 30.6 });
  const exitB = ss(37.7, 38.5, t);
  // axis + gridlines
  const axA = E.outCubic(prog(t, 30.5, 31.3)) * (1 - exitB);
  if (axA > 0) {
    ctx.save(); ctx.globalAlpha *= 1; ctx.lineWidth = 1;
    for (let v = 1; v <= 3; v++) { const y = BASE - v * SC; hl(ctx, y, 240, 240 + 640 * axA, rgba(C.light, 0.35 * (1 - exitB)), 1); }
    ctx.restore();
    ctx.save(); ctx.globalAlpha *= 1 - exitB;
    hl(ctx, BASE, 240, 240 + 640 * axA, C.blue, 2); ctx.strokeStyle = C.light; ctx.lineWidth = 2;
    ctx.beginPath(); ctx.moveTo(240, BASE); ctx.lineTo(240, 250); ctx.stroke();
    for (let v = 0; v <= 3; v++) txt(ctx, `${v} %`, 218, BASE - v * SC + 7, 20, 400, C.light, FM, 'right');
    ctx.restore();
  }
  for (const b of BAR) {
    const p = E.outCubic(prog(t, b.t0, b.t0 + 0.9)) * (1 - E.inCubic(exitB)); if (p <= 0.001) continue;
    const h = b.v * SC * p; ctx.fillStyle = b.col; ctx.fillRect(b.x, BASE - h, 140, h);
    if (p > 0.9 && exitB < 0.2) txt(ctx, fmt(b.v, 1) + ' %', b.x + 70, BASE - h - 14, 28, 500, INK, FM, 'center');
    ctx.save(); ctx.globalAlpha *= (1 - exitB) * ss(b.t0, b.t0 + 0.6, t); txt(ctx, b.lab, b.x + 70, BASE + 36, 22, 500, C.light, FM, 'center'); ctx.restore();
  }
  if (exitB < 1) {
    ctx.save(); ctx.globalAlpha *= (1 - exitB) * ss(32.4, 33.0, t);
    hl(ctx, BASE + 62, 300, 620, C.blue, 1.5); txt(ctx, 'MERKEFORHANDLERE', 300, BASE + 92, 20, 500, C.blue, FM, 'left', 0.1); ctx.restore();
    ctx.save(); ctx.globalAlpha *= (1 - exitB) * ss(35.5, 36.1, t);
    hl(ctx, BASE + 62, 660, 800, C.teal, 1.5); txt(ctx, 'BRUKTBIL', 660, BASE + 92, 20, 500, C.teal, FM, 'left', 0.1); ctx.restore();
  }
  // right column
  const ex = 37.9;
  K.line(ctx, 'Driftsmargin', 1000, 300, t, { size: 84, weight: 800, color: INK, t0: 31.1, out: ex });
  K.line(ctx, 'merkeforhandlere', 1000, 392, t, { size: 84, weight: 300, color: C.blue, t0: 31.3, out: ex });
  K.line(ctx, '1,5 %', 1000, 590, t, { size: 140, weight: 800, color: INK, t0: 32.2, out: ex });
  K.kicker(ctx, '2023', 1000, 478, t, { color: C.light, t0: 32.3, out: ex, tick: false });
  { const p = E.outCubic(prog(t, 33.0, 33.6)) * (1 - ss(ex - 0.45, ex, t)); if (p > 0) { ctx.save(); ctx.strokeStyle = C.blue; ctx.lineWidth = 2; K.drawOn(ctx, [[1330, 548], [1400, 548]], p); ctx.beginPath(); ctx.moveTo(1388, 538); ctx.lineTo(1400, 548); ctx.lineTo(1388, 558); ctx.stroke(); ctx.restore(); } }
  K.line(ctx, '1,7 %', 1430, 590, t, { size: 140, weight: 800, color: C.blue, t0: 33.4, out: ex });
  K.kicker(ctx, '2024', 1430, 478, t, { color: C.light, t0: 33.5, out: ex, tick: false });
  K.kicker(ctx, 'Bruktbilforhandlere · 2024', 1000, 706, t, { color: C.teal, t0: 35.6, out: ex });
  K.line(ctx, '2,5 %', 1000, 836, t, { size: 140, weight: 800, color: C.teal, t0: 35.8, out: ex });
  // matrix
  const cnt = Math.round(1300 * E.outCubic(prog(t, 38.9, 41.0)));
  const mt = 38.9;
  for (let i = 0; i < NDOT; i++) {
    const col = i % 45, row = Math.floor(i / 45), td = mt + col * 0.045 + row * 0.003, p = E.outBack(prog(t, td, td + 0.4));
    if (p <= 0) continue;
    const r = 5.2 * p, isR = i === REDI;
    ctx.fillStyle = isR ? C.red : rgba(C.blue, 0.45 + 0.55 * hash01(i, 3)); ctx.beginPath(); ctx.arc(MX(i), MY(i), r, 0, 7); ctx.fill();
  }
  if (t > 40.3) { const p = prog(t, 40.3, 41.0); K.redDot(ctx, MX(REDI), MY(REDI), 5.2, p); }
  masked(ctx, t, 38.8, 42.6, 560, 170, () => { tabText(ctx, fmt(cnt) + '+', 160, 560, 170, 800, INK, 'left'); });
  K.line(ctx, 'bilbransjebedrifter', 164, 626, t, { size: 44, weight: 300, color: C.blue, t0: 39.1, out: 42.6 });
  K.line(ctx, 'analysert.', 164, 678, t, { size: 44, weight: 300, color: C.blue, t0: 39.25, out: 42.6 });
  cap(ctx, t, 'Vi kjenner bransjen tallene kommer fra.', 38.9, 42.5);
}

// ================================================================== 5. THE ANALYSIS
const PX0 = 200, PX1 = 1720, PY1 = 800;
const yc = (x) => 560 + 35 * Math.sin(((x - PX0) / 1520) * Math.PI * 2.2 + 0.6);
const hw = (x) => 95 + 14 * Math.sin(x / 230);
const gauss = (i) => Math.sqrt(-2 * Math.log(Math.max(1e-6, hash01(i, 6)))) * Math.cos(6.2832 * hash01(i, 7));
const OUT = [
  { i: REDI, x: 610, dy: -hw(610) - 140, label: 'Garantikrav uten arbeidsordre', side: 'r', t0: 46.9 },
  { i: 402, x: 1100, dy: hw(1100) + 120, label: 'Manuell rabatt over fullmakt', side: 'r', t0: 47.9 },
  { i: 777, x: 1480, dy: -hw(1480) - 160, label: 'Bruktbil · lang lagertid', side: 'l', t0: 48.9 },
];
const SCAT = Array.from({ length: NDOT }, (_, i) => {
  const o = OUT.find((q) => q.i === i);
  const x = o ? o.x : PX0 + 30 + hash01(i, 5) * (1520 - 60);
  const dev = o ? o.dy : clamp(gauss(i) * 38, -hw(x) * 0.93, hw(x) * 0.93);
  return { x, y: yc(x) + dev };
});
function s5(ctx, t) {
  bright(ctx, t);
  const bandP = E.outCubic(prog(t, 44.6, 46.0)), exitP = ss(53.0, 53.6, t);
  const ax = E.outCubic(prog(t, 43.0, 44.2));
  ctx.save(); ctx.globalAlpha *= 1 - exitP;
  hl(ctx, PY1, PX0, PX0 + (PX1 - PX0) * ax, C.light, 1.5);
  ctx.strokeStyle = C.light; ctx.lineWidth = 1.5; ctx.beginPath(); ctx.moveTo(PX0, PY1); ctx.lineTo(PX0, PY1 - 560 * ax); ctx.stroke();
  const MON = ['JAN', 'FEB', 'MAR', 'APR', 'MAI', 'JUN', 'JUL', 'AUG', 'SEP', 'OKT', 'NOV', 'DES'];
  ctx.beginPath(); for (let m = 0; m < 12; m++) { const x = PX0 + 1520 * (m + 0.5) / 12; ctx.moveTo(x, PY1); ctx.lineTo(x, PY1 + 8); } ctx.strokeStyle = C.light; ctx.stroke();
  for (let m = 0; m < 12; m++) txt(ctx, MON[m], PX0 + 1520 * (m + 0.5) / 12, PY1 + 38, 18, 500, rgba(C.light, ax), FM, 'center', 0.08);
  txt(ctx, 'TRANSAKSJONER · 12 MÅNEDER', PX0 + 20, 232, 18, 500, rgba(C.light, ax), FM, 'left', 0.12);
  // confidence band
  if (bandP > 0) {
    const n = 60, xe = PX0 + 1520 * bandP;
    ctx.beginPath();
    for (let k = 0; k <= n; k++) { const x = PX0 + (xe - PX0) * k / n; ctx.lineTo(x, yc(x) - hw(x)); }
    for (let k = n; k >= 0; k--) { const x = PX0 + (xe - PX0) * k / n; ctx.lineTo(x, yc(x) + hw(x)); }
    ctx.closePath(); ctx.fillStyle = rgba(C.cloud, 0.65); ctx.fill();
    ctx.strokeStyle = rgba(C.light, 0.9); ctx.lineWidth = 1.2; ctx.setLineDash([2, 5]);
    for (const sg of [-1, 1]) { ctx.beginPath(); for (let k = 0; k <= n; k++) { const x = PX0 + (xe - PX0) * k / n; ctx.lineTo(x, yc(x) + sg * hw(x)); } ctx.stroke(); }
    ctx.setLineDash([]);
    ctx.strokeStyle = C.blue; ctx.lineWidth = 2; ctx.beginPath(); for (let k = 0; k <= n; k++) { const x = PX0 + (xe - PX0) * k / n; ctx.lineTo(x, yc(x)); } ctx.stroke();
  }
  ctx.restore();
  // dots
  const sweep = prog(t, 50.8, 53.0), sx = PX0 + 1520 * sweep;
  const flag = (q) => E.outBack(prog(t, q.t0, q.t0 + 0.5));
  for (let i = 0; i < NDOT; i++) {
    const col = i % 45, row = Math.floor(i / 45);
    const e = E.inOutCubic(prog(t, 42.7 + col * 0.025, 42.7 + col * 0.025 + 1.1));
    const col2 = E.inCubic(prog(t, 53.6 + hash01(i, 8) * 0.5, 54.6 + hash01(i, 8) * 0.4));
    let x = lerp(MX(i), SCAT[i].x, e), y = lerp(MY(i), SCAT[i].y, e) - Math.sin(Math.PI * e) * 36 * (hash01(i, 9) - 0.5);
    x = lerp(x, 960, col2);
    let r = lerp(5.2, 3.3, e), a = 0.45 + 0.55 * hash01(i, 3);
    if (sweep > 0 && sweep < 1) a = Math.min(1, a + 0.5 * Math.exp(-(((x - sx) / 70) ** 2)));
    const o = OUT.find((q) => q.i === i);
    if (o && e > 0.99) { const f = flag(o); r = 3.3 + 4.2 * f; ctx.fillStyle = i === REDI ? C.red : mix2(C.blue, C.red, clamp(f)); }
    else ctx.fillStyle = i === REDI ? C.red : rgba(C.blue, a);
    ctx.beginPath(); ctx.arc(x, y, r, 0, 7); ctx.fill();
  }
  // flags: ring, leader, label
  for (const o of OUT) {
    const f = prog(t, o.t0, o.t0 + 0.5), ex2 = 1 - ss(53.0, 53.6, t); if (f <= 0 || ex2 <= 0) continue;
    const X = o.x, Y = SCAT[o.i].y;
    ctx.save(); ctx.globalAlpha *= ex2;
    ctx.strokeStyle = C.red; ctx.lineWidth = 1.8; ctx.beginPath(); ctx.arc(X, Y, 8 + 12 * E.outCubic(f), 0, 7); ctx.globalAlpha *= (1 - 0.0); ctx.stroke();
    const dir = o.side === 'r' ? 1 : -1, ll = E.outCubic(prog(t, o.t0 + 0.2, o.t0 + 0.8));
    ctx.lineWidth = 1.5; K.drawOn(ctx, [[X + dir * 22, Y], [X + dir * 66, Y]], ll);
    ctx.restore();
    K.line(ctx, o.label, X + dir * 80, Y + 11, t, { size: 34, weight: 600, color: INK, t0: o.t0 + 0.35, out: 53.0, align: o.side === 'r' ? 'left' : 'right' });
  }
  K.kicker(ctx, 'Krever nærmere undersøkelse', 160, 120, t, { color: C.red, t0: 46.4, out: 53.2 });
  K.kicker(ctx, 'Bilbransjen · 1 300 punkter', 160, 120, t, { color: C.blue, t0: 43.0, out: 46.0 });
  // numeral from scene 4 leaves
  { const p = E.inCubic(prog(t, 42.5, 43.2)); if (p < 1) { ctx.save(); ctx.globalAlpha *= 1 - p; ctx.translate(-220 * p, 0); tabText(ctx, '1 300+', 160, 560, 170, 800, INK, 'left'); txt(ctx, 'bilbransjebedrifter', 164, 626, 44, 300, C.blue, FD); txt(ctx, 'analysert.', 164, 678, 44, 300, C.blue, FD); ctx.restore(); } }
  cap(ctx, t, 'Fra store datamengder til målrettet revisjon.', 44.8, 50.7);
  cap(ctx, t, 'Revisor vurderer. Teknologien viser vei.', 51.3, 54.7);
}

// ================================================================== 6. VALUE
const COLS = [
  { x: 160, head: 'For revisor', pts: ['Målrettet revisjon', 'Mer tid til vurdering', 'Sporbar dokumentasjon'], dir: -1, d: 0 },
  { x: 1040, head: 'For kunden', pts: ['Innsikt i egen drift', 'Bransjesammenligning', 'Færre overraskelser'], dir: 1, d: 0.45 },
];
const ROWY = [320, 500, 680];
function statement(ctx) {
  const size = M.fitSize(ctx, 'egen virksomhet', 800, FD, 1500, 190);
  ctx.font = M.font(size, 800, FD);
  const w2 = ctx.measureText('egen virksomhet').width, wm = ctx.measureText('Mer ').width, wi = ctx.measureText('innsikt').width;
  return { size, y1: 460, y2: 460 + size * 1.04, w2, wm, wi, px: 160 + w2 + size * 0.08 + size * 0.06, r: size * 0.062 };
}
function s6(ctx, t) {
  bright(ctx, t);
  const ex = 61.2;
  // centre rule
  const ce = E.outCubic(prog(t, 55.0, 56.2)), ceo = 1 - E.inOutCubic(prog(t, 60.5, 61.5));
  if (ceo > 0) { const c = 515, h = 340 * ce * ceo; hl2(ctx, 960, c - h, c + h, C.blue, 2); }
  for (const col of COLS) {
    const t0 = 55.8 + col.d;
    K.line(ctx, col.head.toUpperCase(), col.x, 220, t, { size: 64, weight: 800, color: INK, t0, out: ex });
    K.kicker(ctx, col.dir < 0 ? 'Revisjonsteamet' : 'Virksomheten', col.x, 148, t, { color: C.blue, t0: t0 - 0.1, out: ex });
    col.pts.forEach((p, k) => {
      const ty = t0 + 0.4 + k * 0.7, y = ROWY[k];
      const rp = E.outCubic(prog(t, ty - 0.2, ty + 0.7)) * (1 - E.inOutCubic(prog(t, ex - 0.5, ex)));
      if (rp > 0) {
        const x0 = col.dir < 0 ? 880 : 1040, x1 = col.dir < 0 ? 160 : 1760;
        hl(ctx, y, x0, lerp(x0, x1, rp), C.blue, 1.8);
      }
      txt2(ctx, t, `0${k + 1}`, col.x, y + 40, ty + 0.1, ex);
      K.line(ctx, p, col.x, y + 120, t, { size: 52, weight: 600, color: INK, t0: ty + 0.15, out: ex });
    });
  }
  // the statement
  const S = statement(ctx), t0 = 61.9, out = 67.35;
  if (t > 61.5) {
    K.line(ctx, 'Mer ', 160, S.y1, t, { size: S.size, weight: 800, color: INK, t0, out });
    K.line(ctx, 'innsikt', 160 + S.wm, S.y1, t, { size: S.size, weight: 800, color: C.blue, t0: t0 + 0.1, out });
    K.line(ctx, ' i', 160 + S.wm + S.wi, S.y1, t, { size: S.size, weight: 800, color: INK, t0: t0 + 0.15, out });
    K.line(ctx, 'egen virksomhet', 160, S.y2, t, { size: S.size, weight: 800, color: INK, t0: t0 + 0.35, out });
    const up = E.outQuint(prog(t, 63.6, 64.8)) * (1 - E.inCubic(prog(t, out - 0.4, out)));
    ctx.fillStyle = C.blue; ctx.fillRect(160 + S.wm, S.y1 + 18, S.wi * up, 6);
    const pd = E.outBack(prog(t, t0 + 0.8, t0 + 1.3));
    if (t < 67.0 && pd > 0) K.redDot(ctx, S.px, S.y2 - S.r, S.r * pd, 1);
  }
}
function hl2(ctx, x, y0, y1, col, lw) { ctx.strokeStyle = col; ctx.lineWidth = lw; ctx.beginPath(); ctx.moveTo(x, y0); ctx.lineTo(x, y1); ctx.stroke(); }
function txt2(ctx, t, s, x, y, t0, out) {
  const p = E.outQuint(prog(t, t0, t0 + 0.5)) * (1 - ss(out - 0.4, out, t)); if (p <= 0) return;
  txt(ctx, s, x, y, 22, 500, rgba(C.light, p), FM, 'left', 0.1);
}

// ================================================================== 7. END CARD
const TAG = 'Vi ser detaljene. Vi forstår helheten.';
function landing(ctx) { ctx.font = M.font(56, 300, FD); const w = ctx.measureText(TAG).width; return [960 + w / 2 + 40, 600 - 18]; }
function s7(ctx, t) {
  bright(ctx, t);
  const wipe = E.outCubic(prog(t, 67.9, 68.9));
  K.drawLogo(ctx, 960, 390, 150, { alpha: 1, wipe });
  K.line(ctx, TAG, 960, 600, t, { size: 56, weight: 300, color: INK, t0: 69.0, align: 'center' });
  K.line(ctx, 'TETT PÅ BILBRANSJEN', 960, 690, t, { size: 24, weight: 500, color: C.blue, t0: 69.7, align: 'center', family: FM, track: 0.3 });
}
function flight(ctx, t) {
  if (t < 67.0) return;
  const S = statement(ctx), P0 = [S.px, S.y2 - S.r], P1 = landing(ctx), a = clamp((t - 67.0) / 4.0);
  if (a >= 1) { const p = prog(t, 71.0, 71.9); K.redDot(ctx, P1[0], P1[1], 8, 1); if (p > 0 && p < 1) { ctx.strokeStyle = rgba(C.red, 1 - p); ctx.lineWidth = 1.5; ctx.beginPath(); ctx.arc(P1[0], P1[1], 10 + 50 * E.outCubic(p), 0, 7); ctx.stroke(); } return; }
  const e = E.inOutCubic(a), c1 = [1860, 380], c2 = [1590, 330], u = 1 - e;
  const x = u * u * u * P0[0] + 3 * u * u * e * c1[0] + 3 * u * e * e * c2[0] + e * e * e * P1[0];
  const y = u * u * u * P0[1] + 3 * u * u * e * c1[1] + 3 * u * e * e * c2[1] + e * e * e * P1[1];
  K.redDot(ctx, x, y, lerp(S.r, 8, e), 1);
}

// ================================================================== film
const HITS = [
  [0, 'tallet', 'blip', { pitch: 'D5' }], [2.7, 'krymper', 'whoosh', { len: 1.8, from: 300, to: 1600 }], [4.3, 'bildetekst', 'tick'], [6.0, 'raden blir linje', 'riser', { len: 1.5 }],
  [7.5, 'tidslinje', 'pop', { pitch: 'A4' }], [9.0, 'innkjøp', 'tick'], [10.2, 'salg', 'click'], [11.6, 'service', 'tick'], [13.3, 'multipliseres', 'whoosh', { len: 2.4, from: 400, to: 2000 }],
  [17.1, 'horisont åpner', 'whoosh', { len: 1.6, from: 1800, to: 500 }],
  ...[0, 1, 2, 3, 4, 5].map((i) => [+dT(i).toFixed(2), DEPTS[i][0], 'blip', { pitch: ['D5', 'E5', 'F#5', 'A5', 'B5', 'D6'][i] }]),
  [29.4, 'akse', 'click'], [32.0, 'søyle', 'tick'], [33.3, 'søyle 2', 'tick'], [35.5, 'bruktbil', 'pop', { pitch: 'F#5' }], [38.9, '1300 prikker', 'swell', { len: 2.3 }],
  [42.7, 'spredning', 'whoosh', { len: 2.0, from: 500, to: 1800 }], [46.9, 'punkt 1', 'pop', { pitch: 'A5' }], [47.9, 'punkt 2', 'pop', { pitch: 'B5' }], [48.9, 'punkt 3', 'pop', { pitch: 'D6' }],
  [51.3, 'revisor vurderer', 'tick'], [53.6, 'prikker til linje', 'riser', { len: 1.4 }], [55.0, 'midtlinje', 'click'],
  [56.2, 'punkt', 'tick'], [57.6, 'punkt', 'tick'], [59.0, 'punkt', 'tick'], [61.9, 'utsagn', 'sub', { len: 1.2 }], [67.0, 'prikken løsner', 'whoosh', { len: 3.5, from: 600, to: 2400 }],
  [67.9, 'logo', 'swell', { len: 1.6 }], [71.0, 'siste datapunkt', 'bell', { pitch: 'D6' }],
].map(([s, ...r]) => [s / PERIOD, ...r]);

function route(ctx, t) {
  if (t < 7.5) s1(ctx, t);
  else if (t < 17.1) s2(ctx, t);
  else if (t < 18.5) {
    s3(ctx, t); const d = 540 * E.inOutCubic(prog(t, 17.1, 18.5));
    ctx.save(); ctx.beginPath(); ctx.rect(0, 0, W, 540 - d); ctx.rect(0, 540 + d, W, 540 - d); ctx.clip(); s2(ctx, t); ctx.restore();
    hl(ctx, 540 - d, 0, W, rgba(C.light, 0.8), 1.2); hl(ctx, 540 + d, 0, W, rgba(C.light, 0.8), 1.2);
  } else if (t < 29.4) s3(ctx, t);
  else if (t < 30.6) {
    s3(ctx, t); const d = 1750 * E.inOutCubic(prog(t, 29.45, 30.6));
    ctx.save(); ctx.beginPath(); ctx.rect(240 - d, 0, 2 * d, H); ctx.clip(); s4(ctx, t); ctx.restore();
    hl2(ctx, 240 - d, 0, H, rgba(C.light, 0.8), 1.2); hl2(ctx, 240 + d, 0, H, rgba(C.light, 0.8), 1.2);
  } else if (t < 42.5) s4(ctx, t);
  else if (t < 55.0) s5(ctx, t);
  else if (t < 67.5) s6(ctx, t);
  else s7(ctx, t);
  flight(ctx, t);
}
const sheetK = (t) => ((t > 18.5 && t < 29.8) || (t > 17.1 && t <= 18.5 && false) ? 0 : 1);

M.film({
  fonts: [M.font(100, 600, FD), M.font(100, 300, FD), M.font(100, 800, FD), M.font(40, 500, FM), M.font(40, 400, FM)],
  hits: HITS,
  samples: 4,
  async init() { await K.loadLogo(); },
  draw(ctx, u, t) { route(ctx, t); },
  post(ctx, u, t) { K.grain(ctx, t, sheetK(t)); },
});
