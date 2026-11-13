// BDO Automotive – Action B «TYPOGRAFI». 60 s trailer at 128 bpm, 32 bars.
// A pure kinetic-typography film in the Swiss poster tradition: type, rules, flat colour fields, the red dot,
// flat bars and a dot matrix. A pure function of time: window.seek(t) paints frame t. No state, no Math.random.
// Script: a manifesto on three pillars told as an equation (BRANSJEKUNNSKAP + DATA + TEKNOLOGI = ...).
import * as M from './lib/motion.js';
import * as K from './lib/kit.js';

const { W, H, C, clamp, lerp, prog, E, hash01, mulberry32 } = K;
const P = 0.46875;                       // one beat in seconds
const END = 116;                         // end card: bar 30
const DISPLAY = K.DISPLAY, MONO = K.MONO;
const TAN35 = Math.tan((35 * Math.PI) / 180);

// ------------------------------------------------------------------ fields (flat full-bleed colour)
const F = {
  N: { id: 'N', bg: C.night, fg: C.white, dim: 'rgba(255,255,255,0.22)', hair: 'rgba(255,255,255,0.10)', acc: C.red, mute: 'rgba(255,255,255,0.55)' },
  A: { id: 'A', bg: C.arctic, fg: C.blue, dim: 'rgba(34,64,154,0.12)', hair: 'rgba(34,64,154,0.10)', acc: C.red, mute: C.slate },
  B: { id: 'B', bg: C.blue, fg: C.white, dim: 'rgba(255,255,255,0.26)', hair: 'rgba(255,255,255,0.14)', acc: C.red, mute: 'rgba(255,255,255,0.62)' },
  R: { id: 'R', bg: C.red, fg: C.white, dim: C.burgundy, hair: 'rgba(255,255,255,0.18)', acc: C.white, mute: 'rgba(255,255,255,0.75)' },
  X: { id: 'X', bg: C.white, fg: C.charcoal, dim: C.pale, hair: C.pale, acc: C.red, mute: C.slate },
};

// ------------------------------------------------------------------ drawing helpers
let DOT = null, WIN = null;              // last full stop / first O drawn this frame (transitions use them)
const disc = (ctx, x, y, r, col) => { ctx.fillStyle = col; ctx.beginPath(); ctx.arc(x, y, Math.max(0.1, r), 0, 7); ctx.fill(); };
const base = (sz) => H / 2 + 0.36 * sz;
const fit = (ctx, s, maxW = 1728, max = 1100, w = 800) => M.fitSize(ctx, s, w, DISPLAY, maxW, max, -0.03);
const ein = (c, a, d) => E.outExpo(prog(c.lt, a, a + d));

// One line of display type, per-glyph control. The full stop is always a geometric disc.
function word(ctx, str, x, bs, size, o = {}) {
  const wgt = o.weight || 800, tr = (o.track ?? -0.03) * size;
  const f = M.font(size, wgt, DISPLAY), lay = M.layout(ctx, str, f, tr);
  const x0 = (o.align === 'c' ? x - lay.width / 2 : o.align === 'r' ? x - lay.width : x) + (o.dx || 0);
  const by = bs + (o.dy || 0), n = lay.glyphs.length, col = o.color || '#fff';
  if (o.mask) { ctx.save(); ctx.beginPath(); ctx.rect(-300, by - size * 1.25, W + 600, size * 1.25 + size * 0.07); ctx.clip(); }
  for (let i = 0; i < n; i++) {
    const g = lay.glyphs[i], p = o.per ? o.per(i, n, g) : null;
    if (p && p.hide) continue;
    const gx = x0 + g.cx + (p ? p.dx || 0 : 0), gy = by + (p ? p.dy || 0 : 0), sc = p && p.sc != null ? p.sc : 1;
    if (g.ch === '.' && i === n - 1 && o.dot !== false) {
      const r = size * (wgt < 500 ? 0.055 : 0.088) * sc; disc(ctx, gx, gy - r, r, o.dotColor || C.red); DOT = { x: gx, y: gy - r, r }; continue;
    }
    if (g.ch.trim() === '') continue;
    const cc = (p && p.col) || (o.colors ? o.colors(i, g) : col);
    if (o.outline) {
      if (g.ch === '.' && o.dot === false && o.noDot) continue;
      ctx.save(); ctx.translate(gx, gy); if (p && p.rot) ctx.rotate(p.rot); ctx.scale(sc, sc);
      ctx.font = f; ctx.textAlign = 'center'; ctx.lineWidth = o.lw || 3; ctx.lineJoin = 'round'; ctx.strokeStyle = cc; ctx.strokeText(g.ch, 0, 0); ctx.restore();
    } else M.glyph(ctx, g.ch, gx, gy, f, cc, sc, sc, (p && p.rot) || 0);
    if (o.win && g.ch === 'O' && !WIN) WIN = { x: gx, y: gy - 0.36 * size, s: size };
  }
  if (o.mask) ctx.restore();
  return { x0, w: lay.width, lay, size };
}
// endless ticker line (tiles `unit` across the frame, scrolled by x px)
const UW = new Map();
function ticker(ctx, unit, y, size, col, x, weight = 800) {
  const f = M.font(size, weight, DISPLAY), tr = -0.03 * size, key = f + unit;
  ctx.font = f; ctx.letterSpacing = `${tr}px`; ctx.textAlign = 'left'; ctx.fillStyle = col;
  if (!UW.has(key)) UW.set(key, ctx.measureText(unit).width);
  const w = UW.get(key); let xx = ((x % w) + w) % w - w;
  for (; xx < W; xx += w) ctx.fillText(unit, xx, y);
  ctx.letterSpacing = '0px';
}
// MONO label, optionally wiped in (p 0..1) with the red tick
function label(ctx, str, x, y, col, { size = 34, p = 1, align = 'left', tick = false, track = 0.12 } = {}) {
  if (p <= 0) return 0;
  ctx.save(); ctx.font = M.font(size, 500, MONO); ctx.letterSpacing = `${size * track}px`; ctx.textAlign = 'left';
  const s = str.toUpperCase(), w = ctx.measureText(s).width, x0 = align === 'right' ? x - w : x;
  ctx.beginPath(); ctx.rect(x0 - 50, y - size * 1.2, (w + 70 + (tick ? 42 : 0)) * clamp(p), size * 1.8); ctx.clip();
  if (tick) { ctx.fillStyle = C.red; ctx.fillRect(x0, y - size * 0.36, 26, 4); }
  ctx.fillStyle = col; ctx.fillText(s, x0 + (tick ? 42 : 0), y); ctx.restore(); return w;
}
const rect = (ctx, x, y, w, h, col) => { ctx.fillStyle = col; ctx.fillRect(x, y, w, h); };
function ledger(ctx, c, x, y, w, rows, pitch, col, t0 = 0.3) {
  for (let r = 0; r < rows; r++) { const q = E.outCubic(prog(c.lt, t0 + 0.06 * r, t0 + 0.55 + 0.06 * r)); rect(ctx, x, y + r * pitch, w * q, 3, col); }
}
function halfPath(ctx, cx, cy, ang) {      // half-plane right/below a line through (cx,cy), line direction rising at `ang` degrees
  const a = (ang * Math.PI) / 180, dx = Math.cos(a), dy = -Math.sin(a), nx = Math.sin(a), ny = Math.cos(a), R = 4000;
  ctx_path(ctx, cx - dx * R, cy - dy * R, cx + dx * R, cy + dy * R, cx + dx * R + nx * R, cy + dy * R + ny * R, cx - dx * R + nx * R, cy - dy * R + ny * R);
}
function ctx_path(ctx, ...p) { ctx.beginPath(); ctx.moveTo(p[0], p[1]); for (let i = 2; i < p.length; i += 2) ctx.lineTo(p[i], p[i + 1]); ctx.closePath(); }
function edge(ctx, cx, cy, ang, col, lw) {
  const a = (ang * Math.PI) / 180, dx = Math.cos(a), dy = -Math.sin(a), R = 4000;
  ctx.beginPath(); ctx.moveTo(cx - dx * R, cy - dy * R); ctx.lineTo(cx + dx * R, cy + dy * R); ctx.lineWidth = lw; ctx.strokeStyle = col; ctx.stroke();
}
// two colour fields divided by a line; `draw(colour, field)` paints the same word once per field
function split(ctx, cx, cy, ang, F1, F2, draw) {
  rect(ctx, -100, -100, W + 200, H + 200, F1.bg); draw(F1.fg, F1);
  ctx.save(); halfPath(ctx, cx, cy, ang); ctx.clip(); rect(ctx, -100, -100, W + 200, H + 200, F2.bg); draw(F2.fg, F2); ctx.restore();
  edge(ctx, cx, cy, ang, C.red, 6);
}
const SIZE = (ctx, s, max) => fit(ctx, s, 1728, max);

// ------------------------------------------------------------------ transitions (pre-roll on the shot being left)
const T = {
  diag: { type: 'diag', d: 0.8 }, ruleX: { type: 'rule', dir: 'x', d: 0.7 }, ruleY: { type: 'rule', dir: 'y', d: 0.7 },
  grow: { type: 'grow', d: 0.8 }, win: { type: 'win', d: 0.9 }, slow: { type: 'rule', dir: 'x', d: 1.3 },
};
const RM = Math.hypot(W, H) + 200;
function overlay(ctx, u, nx, ni) {
  const tin = nx.tin; if (!tin) return;
  const q = (u - (nx.b - tin.d)) / tin.d; if (q <= 0 || q >= 1) return;
  const bg = nx.L.bg, edgeCol = bg === C.red || bg === C.white ? C.blue : C.red;
  if (tin.type === 'grow') {
    if (!DOT) return;
    const R = (k) => DOT.r + (RM - DOT.r) * Math.pow(clamp(k), 3);
    if (bg === C.red) disc(ctx, DOT.x, DOT.y, R(q), C.red);
    else { disc(ctx, DOT.x, DOT.y, R(q), C.red); disc(ctx, DOT.x, DOT.y, R((q - 0.16) / 0.84), bg); }
  } else if (tin.type === 'diag') {
    const a = (35 * Math.PI) / 180, nxv = Math.sin(a), nyv = Math.cos(a), smax = W * nxv + H * nyv + 60, smin = -60;
    const s = lerp(smax, smin, E.inOutCubic(q)), cx = nxv * s, cy = nyv * s;
    halfPath(ctx, cx, cy, 35); ctx.fillStyle = bg; ctx.fill(); edge(ctx, cx, cy, 35, edgeCol, 14);
  } else if (tin.type === 'rule') {
    const e = E.inOutCubic(q);
    if (tin.dir === 'x') { const X = lerp(-60, W + 60, e); rect(ctx, -100, -100, X + 100, H + 200, bg); rect(ctx, X - 2, -100, 16, H + 200, edgeCol); }
    else { const Y = lerp(H + 60, -60, e); rect(ctx, -100, Y, W + 200, H + 200, bg); rect(ctx, -100, Y - 14, W + 200, 16, edgeCol); }
  } else if (tin.type === 'win') {
    if (!WIN) return;
    const e = E.inOutCubic(q), rx = lerp(WIN.s * 0.11, RM, e * e), ry = lerp(WIN.s * 0.24, RM, e * e), wx = WIN.x, wy = WIN.y;
    ctx.save(); ctx.beginPath(); ctx.ellipse(wx, wy, rx, ry, 0, 0, 7); ctx.clip();
    ctx.setTransform(1, 0, 0, 1, 0, 0); renderShot(ctx, nx.b, nx, ni, false); ctx.restore();
    ctx.beginPath(); ctx.ellipse(wx, wy, rx, ry, 0, 0, 7); ctx.lineWidth = 8; ctx.strokeStyle = edgeCol; ctx.stroke();
  }
}

// ------------------------------------------------------------------ the shots
const S = [];
const add = (b, n, L, tin, name, sfx, fn, o = {}) => S.push({ b, n, L, tin, name, sfx, fn, ...o });

// ===== bars 1–4: the claim, word by word =====
add(0, 3, F.N, null, 'REVISJON', 'impact', (ctx, c) => {          // gigantic, cropped, snapping into place
  const sz = SIZE(ctx, 'REVISJON.', 420), e = ein(c, 0, 0.38), s = 1 + 1.9 * (1 - e);
  ctx.save(); ctx.translate(W / 2, 560); ctx.scale(s, s); ctx.translate(-W / 2, -560);
  word(ctx, 'REVISJON.', W / 2, base(sz) + 40, sz, { align: 'c', color: c.L.fg, win: true }); ctx.restore();
  ledger(ctx, c, 96, 230, 1728, 3, 14, c.L.fg, 0.5);
  label(ctx, 'Revisjon', 96, 930, c.L.mute, { p: prog(c.lt, 0.8, 1.4), tick: true });
}, { shake: 16 });
add(3, 2, F.A, T.win, 'KAN', 'thud', (ctx, c) => {                  // a 35-degree field split through the word
  const e = ein(c, 0, 0.42), cx = lerp(W + 700, 1000, e), sz = SIZE(ctx, 'KAN.', 800);
  split(ctx, cx, H / 2, 35, F.N, F.A, (col, Fd) => word(ctx, 'KAN.', W / 2, base(sz), sz, { align: 'c', color: col, dotColor: Fd.acc }));
}, { shake: 12 });
add(5, 2, F.B, T.diag, 'VÆRE', 'thud', (ctx, c) => {               // letters arrive from the frame edges
  const sz = SIZE(ctx, 'VÆRE.', 600), D = [[-480, -300], [420, 380], [-300, 420], [500, -380], [-420, 260]];
  word(ctx, 'VÆRE.', W / 2, base(sz), sz, { align: 'c', color: c.L.fg, per: (i) => { const e = E.outExpo(prog(c.lt, 0.07 * i, 0.07 * i + 0.5)); const d = D[i % 5]; return { dx: d[0] * (1 - e), dy: d[1] * (1 - e), rot: (1 - e) * 0.4 * (i % 2 ? 1 : -1) }; } });
});
add(7, 2, F.N, T.grow, 'MER', 'impact', (ctx, c) => {              // stacked repetition: outline, outline, solid
  const sz = 330, rows = [340, 630, 920];
  rows.forEach((by, k) => {
    const e = ein(c, 0.12 * k, 0.4);
    word(ctx, 'MER.', 96, by, sz, { color: c.L.fg, outline: k < 2, lw: 3, dx: -(1 - e) * 600, dot: k === 2 });
  });
  ledger(ctx, c, 1250, 300, 574, 12, 52, c.L.hair, 0.2);
}, { shake: 12 });
add(9, 3, F.A, T.ruleX, 'ENN', 'thud', (ctx, c) => {               // outline echo behind a solid word
  const sz = 780, e = ein(c, 0, 0.32), bs = base(sz);
  word(ctx, 'ENN.', 1824 - 170, bs, sz, { align: 'r', outline: true, lw: 3, color: c.L.fg, dot: false });
  word(ctx, 'ENN.', 1824, bs, sz, { align: 'r', color: c.L.fg, dx: (1 - e) * 700 });
  label(ctx, 'Mer enn', 96, 960, c.L.mute, { p: prog(c.lt, 0.5, 1.1), tick: true });
}, { shake: 8 });
add(12, 4, F.B, null, 'KONTROLL', 'impact', (ctx, c) => {          // tracking collapses in front of a ticker
  const L = c.L, sz = SIZE(ctx, 'KONTROLL.', 400), e = ein(c, 0, 0.6);
  ticker(ctx, 'KONTROLL.  ', 300, 340, L.dim, -c.ls * 240); ticker(ctx, 'KONTROLL.  ', 900, 340, L.dim, c.ls * 240 - 700);
  word(ctx, 'KONTROLL.', W / 2, base(sz), sz, { align: 'c', color: L.fg, win: true, per: (i, n) => ({ dx: (i - (n - 1) / 2) * 90 * (1 - e) }) });
}, { shake: 18 });

// ===== bars 5–12: pillar 1 BRANSJEKUNNSKAP =====
add(16, 2, F.N, T.win, 'VI', 'thud', (ctx, c) => {
  const sz = 900, e = E.outQuint(prog(c.lt, 0, 0.9)), bs = base(sz) + 20;
  word(ctx, 'VI.', 96, bs, sz, { color: c.L.fg, mask: true, dy: (1 - e) * sz * 0.5 });
  label(ctx, 'Pilar 1', 1300, 200, c.L.mute, { p: prog(c.lt, 0.4, 1), tick: true });
  label(ctx, 'Bransjekunnskap', 1300, 250, c.L.fg, { p: prog(c.lt, 0.6, 1.2), size: 40 });
}, { shake: 6, zoom: 1.035 });
add(18, 2, F.A, T.ruleY, 'KJENNER', 'tick', (ctx, c) => {
  const sz = SIZE(ctx, 'KJENNER.', 400), e = E.outQuint(prog(c.lt, 0, 0.8));
  word(ctx, 'KJENNER.', 1824, base(sz), sz, { align: 'r', color: c.L.fg, mask: true, dy: -(1 - e) * sz * 0.55 });
  rect(ctx, 96, base(sz) + 50, 1728 * E.outCubic(prog(c.lt, 0.3, 1.3)), 4, c.L.fg);
});
add(20, 3, F.B, null, 'BILBRANSJEN', 'thud', (ctx, c) => {
  const sz = SIZE(ctx, 'BILBRANSJEN.', 340);
  word(ctx, 'BILBRANSJEN.', 96, base(sz), sz, { color: c.L.fg, per: (i) => ({ dx: -(1 - E.outExpo(prog(c.lt, 0.03 * i, 0.03 * i + 0.4))) * 260 }) });
  rect(ctx, 96, 760, 1728 * E.outCubic(prog(c.lt, 0.4, 1.4)), 6, c.L.fg);
  label(ctx, 'Vi kjenner bransjen', 96, 850, c.L.mute, { p: prog(c.lt, 0.8, 1.5), tick: true });
});
add(23, 4, F.A, T.grow, 'BDO BILBRANSJEANALYSEN 2025', 'impact', (ctx, c) => {   // the title card
  const L = c.L, ln = SIZE(ctx, 'BILBRANSJEANALYSEN', 170), e = ein(c, 0, 0.6), e2 = E.inOutCubic(prog(c.lt, 0.3, 1.6));
  label(ctx, 'BDO', 96, 150, L.fg, { size: 46, tick: true });
  word(ctx, 'BILBRANSJEANALYSEN', 96, 340, ln, { color: L.fg, dot: false, per: (i, n) => ({ dx: (i - (n - 1) / 2) * 26 * (1 - e) }) });
  rect(ctx, 96, 385, 1728, 4, C.red);
  const s2 = 640, by = 920;
  word(ctx, '2025.', 96, by, s2, { outline: true, lw: 4, color: L.fg, dot: false, noDot: true });
  ctx.save(); ctx.beginPath(); ctx.rect(0, by - s2 * 0.8 * (1 - 0) + s2 * 0.8 * (1 - e2), W, H); ctx.clip();
  word(ctx, '2025.', 96, by, s2, { color: L.fg }); ctx.restore();
  rect(ctx, 96, 450 + 0, 1728 * (1 - e2) , 0, L.fg);
}, { shake: 10 });
add(27, 4, F.N, T.diag, '1 300+', 'thud', (ctx, c) => {             // 100 x 13 = 1300 dots
  const L = c.L, cols = 100, rows = 13, px = 1728 / (cols - 1), py = px, n = Math.floor(1300 * E.outCubic(prog(c.lt, 0, 2.8))) + 90;
  label(ctx, 'Bilbransjebedrifter analysert', 96, 150, L.fg, { p: prog(c.lt, 0.1, 0.7), tick: true, size: 38 });
  const sz = 540, e = ein(c, 0, 0.6);
  word(ctx, '1 300+', 96, 600, sz, { color: L.fg, dot: false, per: (i, nn) => ({ dx: (i - (nn - 1) / 2) * 50 * (1 - e) }) });
  ctx.fillStyle = L.dim; ctx.beginPath();
  for (let i = 0; i < 1300; i++) { if (i < n) continue; const x = 96 + Math.floor(i / rows) * px, y = 700 + (i % rows) * py; ctx.moveTo(x + 2, y); ctx.arc(x, y, 2, 0, 7); }
  ctx.fill(); ctx.fillStyle = C.white; ctx.beginPath();
  for (let i = 0; i < Math.min(n, 1300); i++) { const x = 96 + Math.floor(i / rows) * px, y = 700 + (i % rows) * py; ctx.moveTo(x + 5.5, y); ctx.arc(x, y, 5.5, 0, 7); }
  ctx.fill();
  if (n < 1300) disc(ctx, 96 + Math.floor(n / rows) * px, 700 + (n % rows) * py, 8, C.red);
}, { shake: 8 });
add(31, 4, F.A, T.ruleX, '1,5 → 1,7', 'tick', (ctx, c) => {         // two flat bars in true proportion 1,5 : 1,7
  const L = c.L, u = 300, h1 = 1.5 * u, h2 = 1.7 * u, e1 = 0.4 + 0.6 * ein(c, 0, 0.4), e2 = ein(c, 1, 0.5);
  label(ctx, 'Driftsmargin merkeforhandlere', 96, 150, L.fg, { p: prog(c.lt, 0, 0.6), size: 38 });
  label(ctx, '2023 → 2024', 96, 210, L.mute, { p: prog(c.lt, 0.2, 0.8), size: 34 });
  label(ctx, 'Etterlengtet økning', 1824, 150, L.fg, { p: prog(c.lt, 1.6, 2.3), size: 38, align: 'right', tick: true });
  const y1 = H + 40 - (h1 + 40) * e1, y2 = H + 40 - (h2 + 40) * e2;
  rect(ctx, 96, y1, 784, H + 60 - y1, C.slate);
  rect(ctx, 960, y2, 784, H + 60 - y2, C.blue);
  word(ctx, '1,5 %', 136, y1 + 60 + 190, 260, { color: C.white, dot: false });
  if (e2 > 0.5) word(ctx, '1,7 %', 1000, y2 + 60 + 190, 260, { color: C.white, dot: false });
  label(ctx, '2023', 96, y1 - 20, C.slate, { size: 32 }); label(ctx, '2024', 960, y2 - 20, C.blue, { size: 32, p: e2 });
  if (e2 > 0.3) { ctx.save(); ctx.setLineDash([14, 12]); ctx.strokeStyle = C.white; ctx.lineWidth = 3; ctx.beginPath(); ctx.moveTo(960, y1); ctx.lineTo(960 + 784 * e2, y1); ctx.stroke(); ctx.restore(); rect(ctx, 1744, y2, 28, y1 - y2, C.red); }
});
add(35, 3, F.B, null, '2,5 %', 'thud', (ctx, c) => {
  const L = c.L, u = 300, h = 2.5 * u * (0.3 + 0.7 * ein(c, 0, 0.5));
  label(ctx, 'Driftsmargin', 96, 150, L.fg, { size: 38, tick: true, p: prog(c.lt, 0, 0.5) });
  label(ctx, 'Bruktbilforhandlere 2024', 96, 202, L.mute, { size: 34, p: prog(c.lt, 0.15, 0.7) });
  word(ctx, '2,5 %', 96, 500, 330, { color: L.fg, dot: false, dx: -(1 - ein(c, 0, 0.35)) * 400 });
  rect(ctx, 1250, H + 40 - h, 574, h + 60, C.white);
  const yg = H - 1.7 * u; ctx.save(); ctx.setLineDash([14, 12]); ctx.strokeStyle = C.white; ctx.lineWidth = 3; ctx.beginPath(); ctx.moveTo(900, yg); ctx.lineTo(1250, yg); ctx.stroke(); ctx.restore();
  label(ctx, 'Merkeforhandlere 1,7 %', 1230, yg - 20, L.mute, { size: 28, align: 'right', p: prog(c.lt, 0.8, 1.4) });
}, { shake: 8 });
add(38, 3, F.A, T.diag, 'SMÅ MELLOMSTORE STORE', 'tick', (ctx, c) => {   // size shown as type size
  const L = c.L;
  const e0 = ein(c, 0, 0.3), e1 = ein(c, 1, 0.3), e2 = ein(c, 2, 0.3);
  word(ctx, 'SMÅ.', 96, 244, 160, { color: L.fg, dx: -(1 - e0) * 500 });
  word(ctx, 'MELLOMSTORE.', 96, 482, 233, { color: L.fg, dx: -(1 - e1) * 900, dot: c.lt >= 1 });
  word(ctx, 'STORE.', 96, 840, 400, { color: L.fg, dx: -(1 - e2) * 1200, dot: c.lt >= 2 });
  label(ctx, 'Forhandlere sammenlignet etter størrelse', 96, 960, L.mute, { size: 30, p: prog(c.lt, 0.1, 0.9) });
}, { shake: 8 });
add(41, 3, F.B, T.ruleX, '10 ÅR', 'thud', (ctx, c) => {            // two rows of ten strokes, no values
  const L = c.L, e = ein(c, 0, 0.4);
  word(ctx, '10 ÅR.', 96, 560, 470, { color: L.fg, dx: (1 - e) * 600 });
  [['Omsetning', 660], ['Driftsmargin', 810]].forEach(([s, y], r) => {
    label(ctx, s, 96, y, L.fg, { size: 32, p: prog(c.lt, 0.1 + r * 0.1, 0.6) });
    for (let k = 0; k < 10; k++) { const q = E.outCubic(prog(c.lt, 0.3 + 0.1 * k + r * 0.3, 0.6 + 0.1 * k + r * 0.3)); rect(ctx, 96 + k * 192, y + 20 + (1 - q) * 70, 14, 70 * q, L.fg); }
  });
  label(ctx, 'Omsetning og driftsmargin i et 10-års perspektiv', 96, 970, L.mute, { size: 28, p: prog(c.lt, 1, 1.8) });
}, { shake: 8 });
add(44, 1, F.R, T.grow, 'TETT', 'impact', (ctx, c) => {            // red hit, bar 12
  const sz = SIZE(ctx, 'TETT.', 900), e = ein(c, 0, 0.18), s = 1 + 0.9 * (1 - e);
  ctx.save(); ctx.translate(W / 2, 540); ctx.scale(s, s); ctx.translate(-W / 2, -540);
  word(ctx, 'TETT.', W / 2, base(sz), sz, { align: 'c', color: c.L.fg, dotColor: C.white }); ctx.restore();
}, { shake: 22 });
add(45, 1, F.R, T.ruleX, 'PÅ', 'thud', (ctx, c) => {
  const sz = SIZE(ctx, 'PÅ.', 1000);
  word(ctx, 'PÅ.', W / 2, base(sz) + 40, sz, { align: 'c', color: c.L.fg, dotColor: C.white, per: (i) => ({ dy: -(1 - E.outBounce(prog(c.lt, 0.05 * i, 0.05 * i + 0.4))) * 260 }) });
}, { shake: 14 });
add(46, 2, F.R, T.ruleY, 'BRANSJEN', 'impact', (ctx, c) => {
  const sz = SIZE(ctx, 'BRANSJEN.', 400), e = ein(c, 0, 0.6);
  word(ctx, 'BRANSJEN.', W / 2, base(sz), sz, { align: 'c', color: c.L.fg, dotColor: C.white, per: (i, n) => ({ dx: (i - (n - 1) / 2) * 70 * (1 - e) }) });
  rect(ctx, 96, base(sz) + 60, 1728 * E.outCubic(prog(c.lt, 0.2, 1)), 14, C.white);
}, { shake: 20 });

// ===== bars 13–20: pillar 2 DATA =====
const DEPT = ['SALG.', 'LAGER.', 'VERKSTED.', 'DELER.', 'GARANTI.', 'SERVICE.'];
add(48, 10, F.N, T.diag, 'LISTE', 'thud', (ctx, c) => {            // the list wall with a magnifying band stepping down
  const L = c.L, rowc = (k) => 190 + k * 140 - 47;
  DEPT.forEach((w, k) => ticker(ctx, `${w}  `, 190 + k * 140, 130, L.dim, (k % 2 ? 1 : -1) * c.ls * 50 - k * 90));
  const idx = Math.min(5, Math.floor(c.lt / 1.4)), yp = rowc(Math.max(0, idx - 1)), yc = lerp(yp, rowc(idx), E.outExpo(prog(c.lt - idx * 1.4, 0, 0.35)));
  rect(ctx, -100, yc - 125, W + 200, 250, L.bg); rect(ctx, 0, yc - 125, W, 3, L.fg); rect(ctx, 0, yc + 122, W, 3, L.fg);
  word(ctx, DEPT[idx], 96, yc + 90, 250, { color: L.fg });
  label(ctx, 'Pilar 2 — Data', 1824, yc - 140, L.mute, { size: 26, align: 'right', tick: false });
}, { shake: 7, zoom: 1.03 });
add(58, 2, F.A, T.diag, 'HVER', 'impact', (ctx, c) => {
  const sz = SIZE(ctx, 'HVER.', 700), e = ein(c, 0, 0.25);
  word(ctx, 'HVER.', W / 2, base(sz), sz, { align: 'c', color: c.L.fg, dx: (1 - e) * 800 });
}, { shake: 14 });
add(60, 3, F.N, T.diag, 'TRANSAKSJON', 'thud', (ctx, c) => {
  const L = c.L, sz = SIZE(ctx, 'TRANSAKSJON.', 400), e = ein(c, 0, 0.55);
  ticker(ctx, 'TRANSAKSJON.  ', 300, 380, L.dim, -c.ls * 240); ticker(ctx, 'TRANSAKSJON.  ', 900, 380, L.dim, c.ls * 240 - 900);
  word(ctx, 'TRANSAKSJON.', W / 2, base(sz), sz, { align: 'c', color: L.fg, win: true, per: (i, n) => ({ dx: (i - (n - 1) / 2) * 80 * (1 - e) }) });
});
add(63, 2, F.B, T.win, 'HVER', 'impact', (ctx, c) => {
  const e = ein(c, 0, 0.4), cx = lerp(W + 700, 1000, e), sz = SIZE(ctx, 'HVER.', 700);
  split(ctx, cx, H / 2, 35, F.B, F.N, (col, Fd) => word(ctx, 'HVER.', W / 2, base(sz), sz, { align: 'c', color: col, dotColor: Fd.acc }));
}, { shake: 12 });
add(65, 3, F.A, T.grow, 'BIL', 'thud', (ctx, c) => {
  const sz = 840, e = ein(c, 0, 0.3), bs = base(sz);
  word(ctx, 'BIL.', 1824 - 180, bs, sz, { align: 'r', outline: true, lw: 3, color: c.L.fg, dot: false });
  word(ctx, 'BIL.', 1824, bs, sz, { align: 'r', color: c.L.fg, dx: (1 - e) * 700 });
}, { shake: 10 });
add(68, 2, F.N, T.ruleX, 'ALT', 'impact', (ctx, c) => {
  const e = ein(c, 0, 0.2), cx = lerp(W + 400, 960, e), sz = SIZE(ctx, 'ALT.', 900), s = 1 + 0.5 * (1 - e);
  split(ctx, cx, H / 2, 35, F.N, F.B, (col, Fd) => { ctx.save(); ctx.translate(W / 2, 540); ctx.scale(s, s); ctx.translate(-W / 2, -540); word(ctx, 'ALT.', W / 2, base(sz), sz, { align: 'c', color: col, dotColor: C.red }); ctx.restore(); });
}, { shake: 14 });
const SCAT = (() => { const r = mulberry32(31), o = []; for (let i = 0; i < 1012; i++) o.push([r() * W, r() * H, 0.3 + 4.2 * r()]); return o; })();
add(70, 6, F.A, null, 'SAMLET', 'thud', (ctx, c) => {              // 1012 dots collect into a strict grid
  const L = c.L, cols = 44, px = 1728 / 43, py = 888 / 22;
  ctx.fillStyle = 'rgba(34,64,154,0.55)'; ctx.beginPath();
  for (let i = 0; i < 1012; i++) {
    const s = SCAT[i], e = E.outCubic(prog(c.lt, 0.15 + s[2] * 0.9, 0.15 + s[2] * 0.9 + 0.9));
    const x = lerp(s[0], 96 + (i % cols) * px, e), y = lerp(s[1], 96 + Math.floor(i / cols) * py, e);
    ctx.moveTo(x + 4, y); ctx.arc(x, y, 4, 0, 7);
  }
  ctx.fill();
  const sz = SIZE(ctx, 'SAMLET.', 440), e = ein(c, 0, 0.5);
  const w = word(ctx, 'SAMLET.', W / 2, base(sz), sz, { align: 'c', color: L.fg, per: (i, n) => ({ dx: (i < 3 ? -1 : 1) * (1 - e) * 400 }) });
}, { shake: 8 });
add(76, 1.5, F.R, T.grow, 'ÉN', 'impact', (ctx, c) => {            // red hit, bar 20
  const sz = 880, e = ein(c, 0, 0.35), s = 1 + 1.2 * (1 - e);
  ctx.save(); ctx.translate(760, 560); ctx.scale(s, s); ctx.translate(-760, -560);
  word(ctx, 'ÉN.', 96, 900, sz, { color: c.L.fg, dotColor: C.white }); ctx.restore();
  ledger(ctx, c, 1570, 330, 254, 10, 60, c.L.fg, 0.3);
}, { shake: 22 });
add(77.5, 2.5, F.R, T.ruleX, 'PLATTFORM', 'thud', (ctx, c) => {
  const L = c.L, sz = SIZE(ctx, 'PLATTFORM.', 400), bs = 520;
  word(ctx, 'PLATTFORM.', 96, bs, sz, { color: L.fg, dotColor: C.white, win: true, mask: true, per: (i) => ({ dy: (1 - E.outQuint(prog(c.lt, 0.04 * i, 0.04 * i + 0.5))) * sz * 0.55 }) });
  label(ctx, 'BDO Dataplattform', 96, bs + 100, L.fg, { size: 38, p: prog(c.lt, 0.3, 0.9), tick: true });
  [800, 850, 900].forEach((y, k) => rect(ctx, 96, y, 1728 * E.outCubic(prog(c.lt, 0.4 + 0.2 * k, 1 + 0.2 * k)), 20, L.fg));
}, { shake: 16 });

// ===== bars 21–24: pillar 3 TEKNOLOGI =====
add(80, 1, F.N, T.win, 'SER', 'impact', (ctx, c) => {              // crosshair rules converge on the full stop
  const L = c.L, sz = SIZE(ctx, 'SER.', 850), e = ein(c, 0, 0.18), s = 1 + 0.7 * (1 - e);
  ctx.save(); ctx.translate(W / 2, 540); ctx.scale(s, s); ctx.translate(-W / 2, -540);
  word(ctx, 'SER.', W / 2, base(sz), sz, { align: 'c', color: L.fg }); ctx.restore();
  if (DOT) { const q = E.outCubic(prog(c.lt, 0.1, 0.9)); ctx.fillStyle = 'rgba(255,255,255,0.7)'; ctx.fillRect(0, DOT.y - 1, W * q, 3); ctx.fillRect(DOT.x - 1, H * (1 - q), 3, H * q); }
}, { shake: 14 });
add(81, 2, F.A, T.diag, 'MØNSTRE', 'tick', (ctx, c) => {           // stripes at 35 degrees
  const L = c.L, sz = SIZE(ctx, 'MØNSTRE.', 400);
  ctx.save(); ctx.translate(W / 2, H / 2); ctx.rotate(-(35 * Math.PI) / 180); ctx.strokeStyle = 'rgba(34,64,154,0.16)'; ctx.lineWidth = 9; ctx.beginPath();
  const off = (c.ls * 70) % 60; for (let y = -1500; y < 1500; y += 60) { ctx.moveTo(-2500, y + off); ctx.lineTo(2500, y + off); } ctx.stroke(); ctx.restore();
  word(ctx, 'MØNSTRE.', W / 2, base(sz), sz, { align: 'c', color: L.fg, per: (i) => ({ dy: (i % 2 ? 1 : -1) * (1 - ein(c, 0.03 * i, 0.4)) * 320 }) });
});
add(83, 1.5, F.B, T.ruleY, 'FINNER', 'tick', (ctx, c) => {         // a rule scans: outline becomes solid
  const L = c.L, sz = SIZE(ctx, 'FINNER.', 500), bs = base(sz), x = lerp(96, 1824, E.inOutCubic(prog(c.lt, 0, 1.2)));
  word(ctx, 'FINNER.', W / 2, bs, sz, { align: 'c', outline: true, lw: 3, color: L.fg, dot: false });
  ctx.save(); ctx.beginPath(); ctx.rect(0, 0, x, H); ctx.clip(); word(ctx, 'FINNER.', W / 2, bs, sz, { align: 'c', color: L.fg }); ctx.restore();
  rect(ctx, x - 3, 96, 6, H - 192, L.fg);
}, { shake: 6 });
add(84.5, 2, F.N, T.diag, 'SAMMENHENGER', 'blip', (ctx, c) => {    // zigzag rules join letter to letter
  const L = c.L, sz = 330, f = M.font(sz, 800, DISPLAY);
  const a = word(ctx, 'SAMMEN', 96, 340, sz, { color: L.fg, per: (i) => ({ dx: -(1 - ein(c, 0.03 * i, 0.4)) * 300 }) });
  const b = word(ctx, 'HENGER.', 96, 698 + 120, sz, { color: L.fg, per: (i) => ({ dx: (1 - ein(c, 0.03 * i, 0.4)) * 300 }) });
  const pts = []; for (let i = 0; i < 6; i++) { pts.push([a.x0 + a.lay.glyphs[i].cx, 352]); pts.push([b.x0 + b.lay.glyphs[i].cx, 818 - 238 - 12]); }
  ctx.lineWidth = 3; ctx.strokeStyle = C.red; K.drawOn(ctx, pts, E.outCubic(prog(c.lt, 0.4, 1.6)));
  ctx.fillStyle = C.white; pts.forEach((p, i) => { if (i / 11 <= prog(c.lt, 0.4, 1.6)) { ctx.beginPath(); ctx.arc(p[0], p[1], 7, 0, 7); ctx.fill(); } });
}, { shake: 6 });
add(86.5, 1.5, F.B, T.ruleX, 'VISER VEI', 'tick', (ctx, c) => {
  const L = c.L, e0 = ein(c, 0, 0.3), e1 = ein(c, 0.6, 0.3);
  word(ctx, 'VISER.', 96, 420, 460, { color: L.fg, dx: -(1 - e0) * 700 });
  word(ctx, 'VEI.', 1824, 860, 460, { align: 'r', color: L.fg, dx: (1 - e1) * 800 });
  rect(ctx, 96, 920, 1728 * E.outCubic(prog(c.lt, 0.7, 1.4)), 8, L.fg);
}, { shake: 10 });
add(88, 4, F.A, T.slow, 'REVISOR', 'bell', (ctx, c) => {           // the breath: slow, near still
  const L = c.L, sz = SIZE(ctx, 'REVISOR.', 420), e = E.outQuint(prog(c.lt, 0, 1.6));
  word(ctx, 'REVISOR.', 96, base(sz) + 20, sz, { color: L.fg, mask: true, dy: (1 - e) * sz * 0.5 });
  rect(ctx, 96, 800, 1728 * E.outCubic(prog(c.lt, 0.5, 3)), 3, L.fg);
  label(ctx, 'Revisors faglige vurdering', 96, 880, L.mute, { size: 34, p: prog(c.lt, 1.5, 2.8), tick: true });
}, { shake: 0, zoom: 1.015 });
add(92, 4, F.N, T.slow, 'VURDERER', 'bell', (ctx, c) => {
  const L = c.L, sz = SIZE(ctx, 'VURDERER.', 420), e = E.outQuint(prog(c.lt, 0, 1.6));
  word(ctx, 'VURDERER.', 1824, base(sz) + 20, sz, { align: 'r', color: L.fg, mask: true, dy: -(1 - e) * sz * 0.5 });
  rect(ctx, 96, 800, 1728 * E.outCubic(prog(c.lt, 0.5, 3)), 3, L.fg);
}, { shake: 0, zoom: 1.015 });

// ===== bars 25–28: the sum, set as an equation =====
add(96, 5, F.N, T.ruleX, 'LIGNINGEN', 'thud', (ctx, c) => {
  const L = c.L, mk = (t0) => E.outQuint(prog(c.lt, t0, t0 + 0.55));
  const s1 = SIZE(ctx, 'BRANSJEKUNNSKAP', 1728, 400);
  word(ctx, 'BRANSJEKUNNSKAP', 96, 282, s1, { color: L.fg, dot: false, mask: true, dy: (1 - Math.max(mk(-0.1), 0)) * s1 * 0.5 });
  word(ctx, '+ DATA', 96, 596, 360, { color: L.fg, dot: false, mask: true, dy: (1 - mk(1)) * 360 * 0.5, colors: (i, g) => (g.ch === '+' ? C.red : L.fg) });
  word(ctx, '+ TEKNOLOGI', 96, 824, 240, { color: L.fg, dot: false, mask: true, dy: (1 - mk(2)) * 240 * 0.5, colors: (i, g) => (g.ch === '+' ? C.red : L.fg) });
  const q = E.outBack(prog(c.lt, 3, 3.3), 2.5);
  if (c.lt >= 3) word(ctx, '=', 1824, 824, 240, { align: 'r', color: C.red, dot: false, per: () => ({ sc: q }) });
}, { shake: 7 });
add(101, 4, F.B, T.diag, 'EFFEKTIV MÅLRETTET REVISJON', 'thud', (ctx, c) => {   // justified stack
  const L = c.L, ws = ['MÅLRETTET.', 'EFFEKTIV.', 'REVISJON.'], y = [], sz = ws.map((w) => SIZE(ctx, w, 1380, 330));
  let top = 100; ws.forEach((w, i) => { y.push(top + sz[i] * 0.72 + (i === 0 ? sz[i] * 0.2 : 0)); top = y[i] + 44; });
  ws.forEach((w, i) => { const e = E.outQuint(prog(c.lt, i - (i === 0 ? 0.2 : 0), i + 0.55)); word(ctx, w, 96, y[i], sz[i], { color: L.fg, mask: true, dy: (1 - e) * sz[i] * 0.5, dot: i === 2 }); });
  label(ctx, 'Effektiv revisjon', 96, 1010, L.mute, { size: 26, p: prog(c.lt, 2.4, 3.2) });
}, { shake: 8 });
add(105, 3, F.A, T.ruleY, 'OG INNSIKT', 'thud', (ctx, c) => {
  const L = c.L, e0 = ein(c, 0, 0.35), e1 = ein(c, 1, 0.5), s2 = SIZE(ctx, 'INNSIKT.', 460);
  word(ctx, 'OG.', 96, 340, 420, { color: L.fg, dx: -(1 - e0) * 700 });
  word(ctx, 'INNSIKT.', 96, 840, s2, { color: L.fg, per: (i, n) => ({ dx: (i - (n - 1) / 2) * 60 * (1 - e1) }) });
}, { shake: 8 });
add(108, 3, F.R, T.grow, 'SOM GIR VERDI', 'impact', (ctx, c) => {  // red hit, bar 28
  const L = c.L, e = (t0) => ein(c, t0, 0.28);
  [['SOM.', 338, 330, 0], ['GIR.', 616, 330, 1], ['VERDI.', 944, 400, 2]].forEach(([w, by, sz, t0]) => {
    const s = 1 + 0.35 * (1 - e(t0)); ctx.save(); ctx.translate(96, by - sz * 0.36); ctx.scale(s, s); ctx.translate(-96, -(by - sz * 0.36));
    if (c.lt >= t0 - 0.01) word(ctx, w, 96, by, sz, { color: L.fg, dotColor: C.white }); ctx.restore();
  });
}, { shake: 22 });
add(111, 2, F.A, T.ruleX, 'FOR KUNDEN', 'thud', (ctx, c) => {
  const L = c.L, e0 = ein(c, 0, 0.3), e1 = ein(c, 1, 0.3), s2 = SIZE(ctx, 'KUNDEN.', 440);
  word(ctx, 'FOR.', 96, 391, 410, { color: L.fg, dx: -(1 - e0) * 700 });
  word(ctx, 'KUNDEN.', 1824, 800, s2, { align: 'r', color: L.fg, dx: (1 - e1) * 1000 });
}, { shake: 10 });
add(113, 3, F.N, T.diag, 'BDO AUTOMOTIVE', 'impact', (ctx, c) => {  // three hard accents
  const L = c.L, e0 = ein(c, 0, 0.2), s2 = SIZE(ctx, 'AUTOMOTIVE', 300, 300);
  word(ctx, 'BDO', 96, 528, 600, { color: L.fg, weight: 300, dot: false, dx: -(1 - e0) * 900 });
  if (c.lt >= 1) word(ctx, 'AUTOMOTIVE', 96, 816, s2, { color: L.fg, weight: 300, dot: false, dx: (1 - ein(c, 1, 0.2)) * 1200 });
  if (c.lt >= 2) { const q = E.outBack(prog(c.lt, 2, 2.25), 3); disc(ctx, 1824 - 100, 528 - 60, 60 * q, C.red); rect(ctx, 96, 880, 1728 * E.outExpo(prog(c.lt, 2, 2.4)), 8, L.fg); }
}, { shake: 12 });

// ===== end card =====
add(END, 12, F.X, T.diag, 'END', 'impact', (ctx, c) => {
  const L = c.L, e0 = E.outQuint(prog(c.lt, 0.3, 1.5)), e1 = E.outQuint(prog(c.lt, 0.7, 1.9));
  word(ctx, 'Vi ser detaljene.', 96, 450, 150, { color: C.charcoal, weight: 300, track: -0.02, mask: true, dy: (1 - e0) * 150 * 0.9, dotColor: C.red });
  word(ctx, 'Vi forstår helheten.', 96, 630, 150, { color: C.charcoal, weight: 300, track: -0.02, mask: true, dy: (1 - e1) * 150 * 0.9, dotColor: C.red });
  rect(ctx, 96, 740, 140 * E.outCubic(prog(c.lt, 1.2, 2.2)), 4, C.red);
  label(ctx, 'BDO AUTOMOTIVE', 96, 984, C.slate, { size: 30, p: prog(c.lt, 1.4, 2.4) });
  K.cornerstone(ctx, E.outQuint(prog(c.lt, 2.2, 4.2)));
}, { still: true });

// ------------------------------------------------------------------ compositor
const starts = S.map((s) => s.b);
const shotIdx = (u) => { let i = 0; while (i + 1 < S.length && u >= starts[i + 1]) i++; return i; };
function renderShot(ctx, u, sh, idx, withOverlay) {
  const lt = Math.max(0, u - sh.b), ls = lt * P, p = clamp(lt / sh.n), L = sh.L;
  const c = { u, lt, ls, p, n: sh.n, L };
  ctx.save();
  if (!sh.still) {
    const z = Math.exp(Math.log(sh.zoom || 1.045) * p);
    ctx.translate(W / 2, H / 2); ctx.scale(z, z); ctx.translate(-W / 2, -H / 2);
    // parallax: the grid hairlines drift against the camera
    M.shake(ctx, lt, 0, sh.shake ?? 6, (idx * 7 + 3) | 0, 0.12);
  }
  DOT = null; WIN = null;
  ctx.fillStyle = L.bg; ctx.fillRect(-120, -120, W + 240, H + 240);
  if (!sh.still) {
    ctx.save(); ctx.translate(-lt * 9, 0); ctx.fillStyle = L.hair;
    for (let k = 0; k <= 12; k++) ctx.fillRect(96 + k * 144 - 1, -120, 2, H + 240);
    ctx.restore();
  }
  sh.fn(ctx, c);
  if (!sh.still) {
    ctx.save(); ctx.font = M.font(22, 500, MONO); ctx.letterSpacing = '3px'; ctx.fillStyle = L.mute; ctx.textAlign = 'left';
    ctx.fillText('BDO AUTOMOTIVE', 96, 66); ctx.textAlign = 'right'; ctx.fillText(`${String(idx + 1).padStart(2, '0')} / ${S.length}`, W - 96 + 3, 66); ctx.restore();
  }
  if (withOverlay !== false) { const nx = S[idx + 1]; if (nx && nx.tin) overlay(ctx, u, nx, idx + 1); }
  ctx.restore();
}

// ------------------------------------------------------------------ hits
const HITS = [];
S.forEach((s, i) => {
  HITS.push([s.b, s.name, s.sfx]);
  if (s.tin && s.tin.type !== 'grow' && i > 0) HITS.push([s.b - s.tin.d, 'sweep', s.tin.type === 'win' ? 'blip' : 'whoosh']);
});
HITS.push([96 + 3, 'equals', 'click'], [96 + 1, 'plus', 'tick'], [96 + 2, 'plus', 'tick'], [101 + 1, 'line', 'tick'], [101 + 2, 'line', 'tick'],
  [108 + 1, 'GIR', 'thud'], [108 + 2, 'VERDI', 'impact'], [105 + 1, 'INNSIKT', 'tick'], [113 + 1, 'AUTOMOTIVE', 'impact'], [113 + 2, 'dot', 'impact'],
  [END + 4.5, 'cornerstone', 'impact']);
HITS.sort((a, b) => a[0] - b[0]);

M.film({
  fonts: [M.font(100, 800, DISPLAY), M.font(100, 300, DISPLAY), M.font(40, 500, MONO)],
  hits: HITS,
  samples: (t) => { const u = t / P; return u >= END + 5 ? 1 : 4; },
  async init() { await K.loadLogos(); },
  draw(ctx, u) { const i = shotIdx(Math.max(0, u)); renderShot(ctx, Math.max(0, u), S[i], i); },
});
