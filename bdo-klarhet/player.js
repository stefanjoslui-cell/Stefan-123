// player.js: a small, dependency-free player for the packed film (inlined by build.mjs as a plain script).
// The film is window.seek(t). Time comes from the inlined <audio id="track"> when there is audio,
// otherwise from performance.now(). Nothing is fetched from the network.
(() => {
  'use strict';

  const FALLBACK_DURATION = 195;
  const SYNC_S = 0.08;     // hard resync to the audio clock when the two drift further apart than this
  const PULL = 0.05;       // per-frame pull toward the audio clock inside the dead band
  const IDLE_MS = 2500;    // controls hide after this long without pointer movement while playing
  const SKIP_S = 5;

  const css = `
    :root { --bg:#17191C; --ink:#F7F7F7; --cloud:#E7E7E7; --blue:#22409A; --red:#E81A3B; --slate:#333333; }
    html, body { margin: 0; height: 100%; background: var(--bg); overflow: hidden; }
    #stage { position: fixed; inset: 0; display: grid; place-items: center; background: var(--bg);
      color: var(--ink); font-family: Display, sans-serif; }
    #stage.idle { cursor: none; }
    #film { display: block; width: min(100vw, calc(100vh * 16 / 9)); height: auto; aspect-ratio: 16 / 9;
      max-width: 100vw; max-height: 100vh; cursor: pointer; }
    .bar { position: fixed; left: 0; right: 0; bottom: 0; z-index: 2; display: flex; align-items: center; gap: 12px;
      padding: 44px clamp(14px, 3vw, 36px) 18px;
      background: linear-gradient(rgba(11,17,26,0), rgba(11,17,26,0.9)); transition: opacity 0.4s ease; }
    #stage.idle .bar { opacity: 0; pointer-events: none; }
    .btn { flex: none; width: 40px; height: 40px; border: 0; border-radius: 50%; padding: 0; display: grid;
      place-items: center; background: none; color: var(--ink); cursor: pointer; transition: background 0.2s; }
    .btn:hover { background: rgba(220,229,240,0.14); }
    .btn:focus-visible, .big:focus-visible, .scrub:focus-visible { outline: 2px solid var(--cloud); outline-offset: 2px; }
    .btn svg, .big svg { width: 22px; height: 22px; fill: currentColor; }
    .scrub { position: relative; flex: 1; height: 32px; cursor: pointer; touch-action: none; outline: none; }
    .cloud { position: absolute; left: 0; right: 0; top: 50%; height: 6px; margin-top: -3px; border-radius: 3px;
      background: rgba(220,229,240,0.22); box-shadow: inset 0 0 0 1px rgba(220,229,240,0.12); overflow: hidden; }
    .fill { position: absolute; inset: 0; background: var(--blue); transform-origin: 0 50%; transform: scaleX(0); }
    .dot { position: absolute; top: 50%; left: 0; width: 10px; height: 10px; margin: -5px 0 0 -5px; border-radius: 50%;
      background: var(--red); box-shadow: 0 0 0 2px rgba(11,17,26,0.6); }
    .time { flex: none; font-family: Mono, monospace; font-size: 13px; color: var(--cloud); white-space: nowrap;
      font-variant-numeric: tabular-nums; letter-spacing: 0.02em; }
    .big { position: fixed; left: 50%; top: 50%; z-index: 1; width: 96px; height: 96px; margin: -48px 0 0 -48px;
      padding: 0; border-radius: 50%; border: 2px solid var(--ink); background: rgba(11,17,26,0.4); color: var(--ink);
      display: grid; place-items: center; cursor: pointer; transition: background 0.25s, transform 0.25s; }
    .big:hover { background: rgba(11,17,26,0.7); transform: scale(1.04); }
    .big svg { width: 34px; height: 34px; }
    .big .lbl { display: none; position: absolute; top: calc(100% + 16px); left: 50%; transform: translateX(-50%);
      font-family: Display, sans-serif; font-size: 15px; letter-spacing: 0.22em; color: var(--cloud); white-space: nowrap; }
    .big.ended .lbl { display: block; }
    .off { display: none !important; }
  `;
  const style = document.createElement('style');
  style.textContent = css;
  document.head.appendChild(style);

  const svg = (inner) => `<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">${inner}</svg>`;
  const stroke = 'fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"';
  const ICON = {
    play: svg('<path d="M8 5.2v13.6L19 12z"/>'),
    pause: svg('<path d="M6.5 5h4v14h-4zM13.5 5h4v14h-4z"/>'),
    replay: svg(`<path d="M4.5 12a7.5 7.5 0 1 0 2.2-5.3" ${stroke}/><path d="M3.5 3.5V8.5h5z"/>`),
    vol: svg(`<path d="M3.5 9.5h3.7L12 5.8v12.4L7.2 14.5H3.5z"/><path d="M15.5 9a4.2 4.2 0 0 1 0 6M18 6.6a7.8 7.8 0 0 1 0 10.8" ${stroke}/>`),
    mute: svg(`<path d="M3.5 9.5h3.7L12 5.8v12.4L7.2 14.5H3.5z"/><path d="M15.5 9.5l5 5M20.5 9.5l-5 5" ${stroke}/>`),
    fs: svg(`<path d="M4 9V4h5M15 4h5v5M20 15v5h-5M9 20H4v-5" ${stroke}/>`),
    unfs: svg(`<path d="M9 4v5H4M15 4v5h5M15 20v-5h5M9 20v-5H4" ${stroke}/>`),
  };

  const $ = (id) => document.getElementById(id);
  const clamp = (x, a, b) => Math.min(b, Math.max(a, x));
  const fmt = (s) => { const n = Math.floor(Math.max(0, s)); return `${Math.floor(n / 60)}:${String(n % 60).padStart(2, '0')}`; };
  const paint = (t) => { try { window.seek(t); } catch (e) { console.error('player: seek failed', e); } };

  const canvas = $('film');
  const stage = $('stage');
  const track = $('track');          // <audio>, absent when the film has no audio

  const bar = document.createElement('div');
  bar.className = 'bar';
  bar.innerHTML = `
    <button class="btn" id="pp" aria-label="Play"></button>
    <div class="scrub" id="scrub" role="slider" tabindex="0" aria-label="Seek" aria-valuemin="0" aria-valuenow="0">
      <div class="cloud"><div class="fill" id="fill"></div></div>
      <div class="dot" id="dot"></div>
    </div>
    <div class="time" id="time">0:00 / 0:00</div>
    <button class="btn" id="mute" aria-label="Mute"></button>
    <button class="btn" id="fsb" aria-label="Full screen"></button>`;
  const big = document.createElement('button');
  big.className = 'big off';
  big.id = 'big';
  big.setAttribute('aria-label', 'Play');
  big.innerHTML = '<span class="ico"></span><span class="lbl">replay</span>';
  stage.append(big, bar);

  const el = {
    pp: $('pp'), scrub: $('scrub'), fill: $('fill'), dot: $('dot'), time: $('time'),
    mute: $('mute'), fsb: $('fsb'), bigIco: big.querySelector('.ico'),
  };

  const S = {
    ready: false, playing: false, started: false, ended: false,
    t: 0, duration: FALLBACK_DURATION, last: 0, raf: 0,
    audio: false,      // true while the <audio> element is the clock
    token: 0,          // invalidates stale play() promises
    muted: false, drag: false, idleTimer: 0,
  };

  function setIcon(node, key, label) {
    if (node.dataset.k !== key) { node.dataset.k = key; node.innerHTML = ICON[key]; }
    if (label) node.setAttribute('aria-label', label);
  }

  function ui() {
    const p = S.duration ? clamp(S.t / S.duration, 0, 1) : 0;
    el.fill.style.transform = `scaleX(${p})`;
    el.dot.style.left = `${p * 100}%`;
    el.scrub.setAttribute('aria-valuemax', String(S.duration));
    el.scrub.setAttribute('aria-valuenow', S.t.toFixed(1));
    el.time.textContent = `${fmt(S.t)} / ${fmt(S.duration)}`;
    setIcon(el.pp, S.playing ? 'pause' : 'play', S.playing ? 'Pause' : 'Play');
    setIcon(el.mute, S.muted ? 'mute' : 'vol', S.muted ? 'Unmute' : 'Mute');
    setIcon(el.fsb, document.fullscreenElement ? 'unfs' : 'fs', document.fullscreenElement ? 'Exit full screen' : 'Full screen');
    const showBig = S.ready && (!S.started || S.ended);
    big.classList.toggle('off', !showBig);
    big.classList.toggle('ended', S.ended);
    setIcon(el.bigIco, S.ended ? 'replay' : 'play');
    big.setAttribute('aria-label', S.ended ? 'Replay' : 'Play');
    if (!S.playing) { stage.classList.remove('idle'); clearTimeout(S.idleTimer); }
  }

  function wake() {
    stage.classList.remove('idle');
    clearTimeout(S.idleTimer);
    if (S.playing) {
      S.idleTimer = setTimeout(() => { if (S.playing && !S.drag) stage.classList.add('idle'); }, IDLE_MS);
    }
  }

  function failAudio() {
    // play() was refused (sandbox, autoplay policy, decode error): keep the film going on the clock, muted.
    S.audio = false;
    S.muted = true;
    if (track) { try { track.pause(); } catch (e) { /* ignore */ } track.muted = true; }
  }

  function startAudio() {
    if (!track) return;
    const token = ++S.token;
    try {
      track.currentTime = S.t;
      track.muted = S.muted;
      const p = track.play();
      if (p && typeof p.then === 'function') {
        p.then(() => { if (token === S.token && S.playing) S.audio = true; },
          () => { if (token === S.token && S.playing) failAudio(); });
      } else {
        S.audio = true;
      }
    } catch (e) {
      failAudio();
    }
  }

  function stopAudio() {
    S.audio = false;
    if (track) { try { track.pause(); } catch (e) { /* ignore */ } }
  }

  function seekTo(x) {
    S.t = clamp(x, 0, S.duration);
    S.ended = S.t >= S.duration;
    S.last = performance.now();
    if (track) { try { track.currentTime = S.t; } catch (e) { /* not seekable yet */ } }
    paint(S.t);
    ui();
  }

  function finish() {
    S.t = S.duration;
    S.playing = false;
    S.ended = true;
    cancelAnimationFrame(S.raf);
    S.raf = 0;
    stopAudio();
    paint(S.duration);
    ui();
  }

  function tick(now) {
    S.raf = 0;
    if (!S.playing) return;
    const dt = clamp((now - S.last) / 1000, 0, 0.25);   // a hidden tab must not jump the film
    S.last = now;
    let t = S.t + dt;
    if (S.audio && track) {
      if (track.ended || track.paused) {
        // the audio stopped on its own (shorter than the film): finish the film on the clock
        S.audio = false;
      } else if (!track.seeking) {
        const a = track.currentTime;
        if (Math.abs(t - a) > SYNC_S) t = a;             // hard resync
        else t += (a - t) * PULL;                        // smooth pull inside the dead band
      }
    }
    if (t >= S.duration) { finish(); return; }
    S.t = t;
    paint(t);
    ui();
    S.raf = requestAnimationFrame(tick);
  }

  function play() {
    if (!S.ready || S.playing) return;
    if (S.t >= S.duration) seekTo(0);                    // replay
    S.playing = true;
    S.started = true;
    S.ended = false;
    S.audio = false;
    S.last = performance.now();
    startAudio();
    S.raf = requestAnimationFrame(tick);
    ui();
    wake();
  }

  function pause() {
    if (!S.playing) return;
    S.playing = false;
    cancelAnimationFrame(S.raf);
    S.raf = 0;
    S.token++;
    stopAudio();
    ui();
  }

  const togglePlay = () => (S.playing ? pause() : play());

  function toggleMute() {
    S.muted = !S.muted;
    if (track) track.muted = S.muted;
    ui();
  }

  function toggleFs() {
    try {
      const p = document.fullscreenElement ? document.exitFullscreen() : document.documentElement.requestFullscreen();
      if (p && typeof p.catch === 'function') p.catch(() => { /* refused, e.g. in a sandboxed frame */ });
    } catch (e) { /* no fullscreen API here */ }
  }

  // ------------------------------------------------------------ input
  const timeAt = (e) => {
    const r = el.scrub.getBoundingClientRect();
    return (r.width ? clamp((e.clientX - r.left) / r.width, 0, 1) : 0) * S.duration;
  };
  el.scrub.addEventListener('pointerdown', (e) => {
    if (!S.ready) return;
    S.drag = true;
    try { el.scrub.setPointerCapture(e.pointerId); } catch (err) { /* ignore */ }
    seekTo(timeAt(e));
  });
  el.scrub.addEventListener('pointermove', (e) => { if (S.drag) seekTo(timeAt(e)); });
  const endDrag = () => { S.drag = false; };
  el.scrub.addEventListener('pointerup', endDrag);
  el.scrub.addEventListener('pointercancel', endDrag);

  el.pp.addEventListener('click', togglePlay);
  big.addEventListener('click', togglePlay);
  canvas.addEventListener('click', () => { if (S.ready) togglePlay(); });
  el.mute.addEventListener('click', toggleMute);
  el.fsb.addEventListener('click', toggleFs);
  document.addEventListener('fullscreenchange', ui);

  addEventListener('keydown', (e) => {
    if (!S.ready || e.metaKey || e.ctrlKey || e.altKey) return;
    switch (e.key) {
      case ' ':
      case 'Spacebar':
        e.preventDefault();
        // a focused button would also take the space as a click
        if (document.activeElement && document.activeElement.tagName === 'BUTTON') document.activeElement.blur();
        if (!e.repeat) togglePlay();
        break;
      case 'ArrowRight': e.preventDefault(); seekTo(S.t + SKIP_S); break;
      case 'ArrowLeft': e.preventDefault(); seekTo(S.t - SKIP_S); break;
      case 'f': case 'F': toggleFs(); break;
      case 'm': case 'M': toggleMute(); break;
      default: return;
    }
    wake();
  });
  addEventListener('pointermove', wake);
  addEventListener('pointerdown', wake);

  // ------------------------------------------------------------ start
  function init() {
    const d = Number(window.DURATION);
    S.duration = d > 0 ? d : FALLBACK_DURATION;
    if (track) track.preload = 'auto';
    S.ready = true;
    ui();
  }

  if (window.filmReady) {
    init();
  } else {
    const id = setInterval(() => {
      if (window.filmReady) { clearInterval(id); init(); }
    }, 40);
  }
})();
