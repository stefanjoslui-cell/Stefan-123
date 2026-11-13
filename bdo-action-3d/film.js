import * as M from './lib/motion.js';
import * as K from './lib/kit.js';
import * as D from './lib/k3d.js';
M.film({ fonts: [M.font(100, 800, K.DISPLAY), M.font(40, 500, K.MONO)], async init() { await K.loadLogos(); },
  draw(ctx, u, t) {
    K.sheet(ctx, 0);
    const a = t * 0.6, cam = D.camera({ pos: [Math.sin(a) * 9, 3.2, Math.cos(a) * 9], look: [0, 1.2, 0], fov: 42 });
    ctx.strokeStyle = K.rgba(K.C.slate, 0.5); ctx.lineWidth = 1; D.floor(ctx, cam, { step: 1, ext: 14 });
    const bars = [[-3, 1.5, 'slate'], [-1.5, 1.7, 'blue'], [0, 2.5, 'red']];
    bars.map(([x, h, c]) => ({ x, h, c, d: Math.hypot(x - cam.pos[0], cam.pos[2]) })).sort((p, q) => q.d - p.d).forEach(({ x, h, c }) => D.box(ctx, cam, x, 0, -0.5, 1, h, 1, { top: K.mix(K.C[c], '#ffffff', 0.25), side: K.mix(K.C[c], '#000000', 0.3), front: K.C[c] }));
    D.text3(ctx, cam, 'DATA.', { at: [0, 4, 0], height: 1.6, right: [1, 0, 0] });
    D.text3(ctx, cam, 'BILBRANSJEN', { at: [3, 0.01, 3], height: 1, right: [1, 0, 0], up: [0, 0, -1], color: K.C.pale });
    D.slab(ctx, 'VERDI.', 500, 980, 150, { depth: 36, face: K.C.white, side: K.C.red });
    K.cornerstone(ctx, 1);
  } });
