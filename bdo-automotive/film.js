// BDO Automotive – Hele bildet. A pure function of time: window.seek(t) paints frame t.
// Three scene modules share lib/kit.js; each paints the whole frame for its span.
//   scenes/world.js  0–50 s    one transaction -> one car -> the car park -> the business -> the network
//   scenes/data.js   50–120 s  network -> BDO's data platform -> analysis -> the selected points
//   scenes/value.js  120–195 s the auditor's judgement -> value for the customer -> end card
import * as M from './lib/motion.js';
import * as K from './lib/kit.js';
import * as world from './scenes/world.js';
import * as data from './scenes/data.js';
import * as value from './scenes/value.js';

const PERIOD = 0.625;   // 96 bpm: every act boundary (20, 50, 80, 120, 155, 185, 195 s) is a bar line
const scene = (t) => (t < 50 ? world : t < 120 ? data : value);
const HITS = [...world.HITS, ...data.HITS, ...value.HITS].sort((a, b) => a[0] - b[0]).map(([s, ...rest]) => [s / PERIOD, ...rest]);

M.film({
  fonts: [M.font(100, 600, K.DISPLAY), M.font(100, 300, K.DISPLAY), M.font(40, 500, K.MONO)],
  hits: HITS,
  async init() { await K.loadLogo(); },
  samples: 6,
  draw(ctx, u, t) { scene(t).draw(ctx, t); },
  post(ctx, u, t) { K.grain(ctx, t, scene(t).sheetAt ? scene(t).sheetAt(t) : 0); },
});
