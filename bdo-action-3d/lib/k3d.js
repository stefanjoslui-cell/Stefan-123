// k3d.js: a tiny perspective engine on the 2-D canvas (no WebGL, so it plays everywhere and every
// frame stays a pure function of time). World: x right, y UP, z towards the viewer is negative
// (the camera looks along its own forward vector). All helpers take a cam made by camera().
import * as M from './motion.js';
const W = 1920, H = 1080;
const sub = (a, b) => [a[0] - b[0], a[1] - b[1], a[2] - b[2]];
const cross = (a, b) => [a[1] * b[2] - a[2] * b[1], a[2] * b[0] - a[0] * b[2], a[0] * b[1] - a[1] * b[0]];
const dot = (a, b) => a[0] * b[0] + a[1] * b[1] + a[2] * b[2];
const norm = (a) => { const l = Math.hypot(a[0], a[1], a[2]) || 1; return [a[0] / l, a[1] / l, a[2] / l]; };
export const v = { sub, cross, dot, norm, add: (a, b) => [a[0] + b[0], a[1] + b[1], a[2] + b[2]], mul: (a, s) => [a[0] * s, a[1] * s, a[2] * s], lerp: (a, b, t) => [a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t, a[2] + (b[2] - a[2]) * t] };

// camera({ pos, look, fov (vertical, degrees), roll (radians) })
export function camera({ pos = [0, 0, 10], look = [0, 0, 0], fov = 40, roll = 0 } = {}) {
  const f = norm(sub(look, pos)); let r = norm(cross(f, [0, 1, 0])); let u = cross(r, f);
  if (roll) { const c = Math.cos(roll), s = Math.sin(roll); const r2 = [r[0] * c + u[0] * s, r[1] * c + u[1] * s, r[2] * c + u[2] * s]; u = [u[0] * c - r[0] * s, u[1] * c - r[1] * s, u[2] * c - r[2] * s]; r = r2; }
  const focal = H / 2 / Math.tan((fov * Math.PI) / 360), NEAR = 0.05;
  const cam = { pos, f, r, u, focal, NEAR,
    // view space: [x right, y up, depth]
    view(p) { const d = sub(p, pos); return [dot(d, r), dot(d, u), dot(d, f)]; },
    // screen: [sx, sy, px-per-unit at that depth, depth]; depth <= NEAR means behind the camera
    project(p) { const q = cam.view(p), k = focal / Math.max(q[2], 1e-4); return [W / 2 + q[0] * k, H / 2 - q[1] * k, k, q[2]]; },
  };
  return cam;
}
// clip a segment against the near plane (view space), returns screen points or null
function clipSeg(cam, a, b) {
  let A = cam.view(a), B = cam.view(b);
  if (A[2] < cam.NEAR && B[2] < cam.NEAR) return null;
  if (A[2] < cam.NEAR) { const t = (cam.NEAR - A[2]) / (B[2] - A[2]); A = [A[0] + (B[0] - A[0]) * t, A[1] + (B[1] - A[1]) * t, cam.NEAR]; }
  if (B[2] < cam.NEAR) { const t = (cam.NEAR - B[2]) / (A[2] - B[2]); B = [B[0] + (A[0] - B[0]) * t, B[1] + (A[1] - B[1]) * t, cam.NEAR]; }
  const ka = cam.focal / A[2], kb = cam.focal / B[2];
  return [W / 2 + A[0] * ka, H / 2 - A[1] * ka, W / 2 + B[0] * kb, H / 2 - B[1] * kb];
}
// add a 3-D segment to the current path (call ctx.beginPath() before, ctx.stroke() after: batch!)
export function seg(ctx, cam, a, b) { const s = clipSeg(cam, a, b); if (s) { ctx.moveTo(s[0], s[1]); ctx.lineTo(s[2], s[3]); } }
// a 3-D polyline, stroked (p = draw-on 0..1)
export function path3(ctx, cam, pts, p = 1) { const n = Math.max(1, Math.floor((pts.length - 1) * p)); ctx.beginPath(); for (let i = 0; i < n; i++) seg(ctx, cam, pts[i], pts[i + 1]); ctx.stroke(); }
// a filled convex polygon (skipped if any vertex is behind the camera)
export function poly(ctx, cam, pts, fill, stroke) {
  const s = pts.map((p) => cam.project(p)); if (s.some((q) => q[3] < cam.NEAR)) return false;
  ctx.beginPath(); s.forEach((q, i) => (i ? ctx.lineTo(q[0], q[1]) : ctx.moveTo(q[0], q[1]))); ctx.closePath();
  if (fill) { ctx.fillStyle = fill; ctx.fill(); } if (stroke) { ctx.strokeStyle = stroke; ctx.stroke(); }
  return true;
}
// an axis-aligned box sitting at (x, y, z) = min corner, size (w, h, d). cols = { top, side, front } (CSS colours);
// only the faces that face the camera are painted (correct for a single convex box; sort boxes far -> near yourself).
export function box(ctx, cam, x, y, z, w, h, d, cols, stroke) {
  const P = (i, j, k) => [x + i * w, y + j * h, z + k * d], c = cam.pos;
  if (c[1] > y + h) poly(ctx, cam, [P(0, 1, 0), P(1, 1, 0), P(1, 1, 1), P(0, 1, 1)], cols.top, stroke);
  if (c[1] < y) poly(ctx, cam, [P(0, 0, 0), P(1, 0, 0), P(1, 0, 1), P(0, 0, 1)], cols.side, stroke);
  if (c[0] < x) poly(ctx, cam, [P(0, 0, 0), P(0, 1, 0), P(0, 1, 1), P(0, 0, 1)], cols.side, stroke);
  if (c[0] > x + w) poly(ctx, cam, [P(1, 0, 0), P(1, 1, 0), P(1, 1, 1), P(1, 0, 1)], cols.side, stroke);
  if (c[2] > z + d) poly(ctx, cam, [P(0, 0, 1), P(1, 0, 1), P(1, 1, 1), P(0, 1, 1)], cols.front, stroke);
  if (c[2] < z) poly(ctx, cam, [P(0, 0, 0), P(1, 0, 0), P(1, 1, 0), P(0, 1, 0)], cols.front, stroke);
}
// a ground grid on the plane y = y0 (lines every `step`, half-extent `ext`), one batched stroke
export function floor(ctx, cam, { y = 0, step = 2, ext = 40, cx = 0, cz = 0 } = {}) {
  ctx.beginPath();
  const x0 = Math.round(cx / step) * step, z0 = Math.round(cz / step) * step;
  for (let i = -ext; i <= ext; i += step) { seg(ctx, cam, [x0 + i, y, z0 - ext], [x0 + i, y, z0 + ext]); seg(ctx, cam, [x0 - ext, y, z0 + i], [x0 + ext, y, z0 + i]); }
  ctx.stroke();
}
// many points in one fill: pts = [[x,y,z,r?], ...], r in world units (default 0.05); rect sprites (fast)
export function points(ctx, cam, pts, color, r0 = 0.05, minPx = 0.8) {
  ctx.fillStyle = color; ctx.beginPath();
  for (const p of pts) { const q = cam.project(p); if (q[3] < cam.NEAR) continue; const r = Math.max(minPx, (p[3] || r0) * q[2]); if (q[0] < -r || q[0] > W + r || q[1] < -r || q[1] > H + r) continue; if (r < 2.2) ctx.rect(q[0] - r, q[1] - r, 2 * r, 2 * r); else { ctx.moveTo(q[0] + r, q[1]); ctx.arc(q[0], q[1], r, 0, 7); } }
  ctx.fill();
}

// ---------------------------------------------------------------- type in space
const TXT = new Map();
function textCanvas(str, weight, family, color, track) {
  const key = `${str}|${weight}|${family}|${color}|${track}`; if (TXT.has(key)) return TXT.get(key);
  const size = 220, cv = document.createElement('canvas'), c = cv.getContext('2d');
  c.font = M.font(size, weight, family); c.letterSpacing = `${track * size}px`;
  const m = c.measureText(str), w = Math.ceil(m.width) + 8, asc = Math.ceil(size * 0.78), h = Math.ceil(size * 1.02);
  cv.width = w; cv.height = h; c.font = M.font(size, weight, family); c.letterSpacing = `${track * size}px`; c.fillStyle = color; c.textBaseline = 'alphabetic'; c.fillText(str, 4, asc);
  const o = { cv, w, h, aspect: w / h }; TXT.set(key, o); return o;
}
// A flat line of type standing in 3-D: centre `at`, unit vectors `right` and `up` give its plane,
// `height` = cap-box height in world units. Drawn as perspective-correct vertical strips, so it can
// swing like a door, lie on the floor, or fly past the camera. Returns false if it is behind the camera.
export function text3(ctx, cam, str, { at = [0, 0, 0], right = [1, 0, 0], up = [0, 1, 0], height = 1, weight = 800, family = 'Display', color = '#fff', track = -0.03, alpha = 1, strips = 0 } = {}) {
  const T = textCanvas(str, weight, family, color, track), wWorld = height * T.aspect;
  const L = v.add(at, v.mul(right, -wWorld / 2)), half = v.mul(up, height / 2);
  const cornerAt = (s, top) => v.add(v.add(L, v.mul(right, wWorld * s)), top ? half : v.mul(half, -1));
  const c0 = cam.view(cornerAt(0, 1)), c1 = cam.view(cornerAt(1, 1));
  if (c0[2] < cam.NEAR * 4 || c1[2] < cam.NEAR * 4) return false;
  // facing the camera head-on: one drawImage is exact enough and much cheaper
  const n = strips || (Math.abs(c0[2] - c1[2]) < 0.02 * Math.min(c0[2], c1[2]) ? 1 : Math.min(96, Math.max(24, Math.round(T.w / 14))));
  ctx.save(); ctx.globalAlpha *= alpha; ctx.imageSmoothingQuality = 'high';
  const iw0 = 1 / c0[2], iw1 = 1 / c1[2];
  let prevS = 0, pt = cam.project(cornerAt(0, 1)), pb = cam.project(cornerAt(0, 0));
  for (let i = 1; i <= n; i++) {
    const f = i / n, s = (f * iw1) / ((1 - f) * iw0 + f * iw1);       // perspective-correct source position
    const qt = cam.project(cornerAt(s, 1)), qb = cam.project(cornerAt(s, 0));
    // affine map of the source strip onto the (near-parallelogram) screen strip
    const sx0 = prevS * T.w, sw = (s - prevS) * T.w, dx = qt[0] - pt[0] || 0.01;
    ctx.save();
    ctx.transform(dx / sw, (qt[1] - pt[1]) / sw, (pb[0] - pt[0]) / T.h, (pb[1] - pt[1]) / T.h, pt[0], pt[1]);
    ctx.drawImage(T.cv, sx0, 0, sw + 0.6, T.h, 0, 0, sw + 0.6, T.h);
    ctx.restore();
    prevS = s; pt = qt; pb = qb;
  }
  ctx.restore();
  return true;
}
// Extruded block type on the 2-D frame (screen space): the face plus a stack of side layers running
// towards a vanishing direction. depth in px, dir = [dx, dy] unit-ish. Bold poster-3-D for the big words.
export function slab(ctx, str, x, y, size, { depth = 40, dir = [0.6, 0.8], face = '#fff', side = '#22409A', weight = 800, family = 'Display', track = -0.03, align = 'center', steps = 0 } = {}) {
  ctx.save(); ctx.font = M.font(size, weight, family); ctx.letterSpacing = `${track * size}px`; ctx.textAlign = align; ctx.textBaseline = 'alphabetic';
  const n = steps || Math.max(6, Math.round(depth / 2));
  ctx.fillStyle = side; for (let i = n; i >= 1; i--) ctx.fillText(str, x + (dir[0] * depth * i) / n, y + (dir[1] * depth * i) / n);
  ctx.fillStyle = face; ctx.fillText(str, x, y);
  ctx.restore();
}
