# /// script
# requires-python = ">=3.10"
# dependencies = ["librosa>=0.10", "numpy", "scipy", "soundfile"]
# ///
"""BDO action/trailer: 45 s, 128 bpm (beat 0.46875 s, bar 1.875 s, 24 bars), D minor -> D major.

    cd <repo> && MOTION_REEL_SKILL=<skill> uv run bdo-action/audio/score.py

  bars 1-4    tension: drone, ticking, low taiko, bass ostinato creeping in
  risers into bars 5, 9, 13, 19; braams on bars 12 and 21
  bars 5-17   driving 16th bass, kick, toms, claps, brass stabs
  bar 18      breath: near silence + accelerating ticks (riser from nothing)
  bars 19-22  full power; hard stop at 41.25 s (end of bar 22)
  bars 23-24  single soft resolving D major chord, rings out by 45 s
"""
import os
import sys
import zlib
from pathlib import Path

import numpy as np
import soundfile as sf
from scipy import signal


def _skill_dir():
    if os.environ.get("MOTION_REEL_SKILL"):
        return Path(os.environ["MOTION_REEL_SKILL"])
    user = Path.home() / ".claude/skills/motion-reel"
    if (user / "scripts/dsp.py").exists():
        return user
    found = sorted((Path.home() / ".claude/plugins").glob("**/skills/motion-reel/scripts/dsp.py"))
    if found:
        return found[-1].parent.parent
    sys.exit("can't find the motion-reel skill: set MOTION_REEL_SKILL to the folder that holds its SKILL.md")


SKILL = _skill_dir()
sys.path.insert(0, str(SKILL / "scripts"))
from dsp import *  # noqa: E402,F403

HERE = Path(__file__).resolve().parent
_C = {}


def mem(fn, *a):
    k = (fn.__name__,) + a
    if k not in _C:
        _C[k] = fn(*a)
    return _C[k]


def tr(n, s):
    return note_name(midi(n) + s)


def env_ar(n, a, r):
    t = np.arange(n) / SR
    return np.sin(np.clip(t / a, 0, 1) * np.pi / 2) ** 2 * np.sin(np.clip((n / SR - t) / r, 0, 1) * np.pi / 2) ** 2


def mkbus():
    b = Bus(DUR, BPM)
    b.x = b.x.astype(np.float32)
    return b


def bar_t(bar, beat=0.0):
    return (bar * 4 + beat) * BEAT


def mkcurve():
    tb = np.arange(N) / SR / BAR

    def curve(*pts):
        return np.interp(tb, [p[0] for p in pts], [p[1] for p in pts]).astype(np.float32)
    return curve


def chord(name):
    minor = name.endswith("m")
    root = name[:-1] if minor else name
    b = midi(root + "2")
    t = 3 if minor else 4
    nn = note_name
    ar = b + 24 if b < 41 else b + 12
    return dict(sub=nn(b - 12), bass=nn(b), oct=nn(b + 12), third=nn(b + t),
                arp=[nn(ar), nn(ar + t), nn(ar + 7), nn(ar + 12)],
                pad=[nn(b), nn(b + 7), nn(b + 12 + t), nn(b + 26)],
                padO=[nn(b), nn(b + 7), nn(b + 12 + t), nn(b + 19), nn(b + 26)],
                key=[nn(b), nn(b + 7), nn(b + 12 + t), nn(b + 19), nn(b + 24 + t)])


# ---------------------------------------------------------------- instruments

def pad(notes, length, cutoff, attack=0.7, release=1.2, voices=4, detune=0.10, wide=False):
    notes, n = list(notes), int(length * SR)
    pr = np.random.default_rng(zlib.crc32(",".join(notes).encode()))
    L, R = np.zeros(n), np.zeros(n)
    for ni, nt in enumerate(notes):
        f = hz(nt)
        for v in range(voices):
            c = (v - (voices - 1) / 2) / ((voices - 1) / 2) * detune * 100
            x = saw(f * 2 ** (c / 1200), n, phase0=pr.uniform())
            p = 0.5 + (0.46 if wide else 0.3) * np.sin(ni * 1.7 + v * 2.1)
            L += x * np.sqrt(1 - p)
            R += x * np.sqrt(p)
    out = np.stack([filt(L, "lowpass", cutoff, 2), filt(R, "lowpass", cutoff, 2)])
    return (out * env_ar(n, attack, release) / np.sqrt(len(notes) * voices)).astype(np.float32)


def subpad(note, length):
    n = int(length * SR)
    t = tvec(n)
    f = hz(note)
    x = np.sin(2 * np.pi * f * t) + 0.18 * np.sin(2 * np.pi * 2 * f * t)
    return (x * env_ar(n, 1.2, 1.6)).astype(np.float32)


def piano(note, length=3.5):
    f = hz(note)
    n = int(length * SR)
    t = tvec(n)
    pr = np.random.default_rng(int(f * 10))
    x = np.zeros(n)
    tau0 = 2.4 * (262 / f) ** 0.35
    for k, a in enumerate([1, .6, .32, .2, .12, .08, .05], 1):
        fk = f * k * np.sqrt(1 + 3e-4 * k * k)
        if fk > 9000:
            break
        x += a * np.sin(2 * np.pi * fk * t + pr.uniform(0, 6)) * np.exp(-t / (tau0 / k ** 0.55))
    x *= 1 - np.exp(-t / 0.004)
    x = filt(x, "lowpass", 2300, 2)
    x += filt(pr.standard_normal(n) * np.exp(-t / 0.012), "lowpass", 450) * 0.12
    return x * np.clip((length - t) / 0.25, 0, 1)


def bell(note, length=3.0):
    f = hz(note)
    n = int(length * SR)
    t = tvec(n)
    x = np.sin(2 * np.pi * f * t + 1.8 * np.exp(-t / 0.25) * np.sin(2 * np.pi * 3 * f * t)) * np.exp(-t / 1.3)
    x += 0.25 * np.sin(2 * np.pi * 2 * f * t) * np.exp(-t / 0.5)
    x *= (1 - np.exp(-t / 0.003)) * np.clip((length - t) / 0.3, 0, 1)
    return filt(x, "lowpass", 6500, 1)


def lead(note, length):
    n = int(length * SR)
    t = tvec(n)
    f = hz(note) * (1 + 0.003 * np.sin(2 * np.pi * 5.2 * t) * np.clip(t / 0.6, 0, 1))
    x = saw(f, n) * 0.45 + np.sin(2 * np.pi * np.cumsum(f) / SR) * 0.7
    return filt(x, "lowpass", 2300, 2) * adsr(n, 0.015, 0.7, 0.55, 0.25)


def hat_soft():
    return filt(hat(False), "lowpass", 9500, 2)


def click():
    n = int(0.035 * SR)
    t = tvec(n)
    return filt(rng.standard_normal(n), "highpass", 6500, 3) * np.exp(-t / 0.005) * (1 - np.exp(-t / 0.0003))


def tick():
    n = int(0.03 * SR)
    t = tvec(n)
    return filt(rng.standard_normal(n), "bandpass", [2200, 4800], 2) * np.exp(-t / 0.006) * (1 - np.exp(-t / 0.0004))


def duck_curve(hits, n, depth=0.22, rel=0.28):
    g = np.ones(n, dtype=np.float32)
    w = int(rel * 5 * SR)
    shape = (1 - depth * np.exp(-np.arange(w) / SR / rel)).astype(np.float32)
    for h in hits:
        s = int(h * SR)
        e = min(n, s + w)
        if e > s:
            g[s:e] = np.minimum(g[s:e], shape[: e - s])
    return g


def verb(send, length=2.8, decay=0.85, seed_=5, lp=5500, hp=180):
    ir = reverb_ir(length, decay, seed_)
    wet = np.stack([signal.fftconvolve(send[c], ir[c])[:N] for c in range(2)]).astype(np.float32)
    return np.stack([filt(filt(ch, "lowpass", lp, 2), "highpass", hp, 1) for ch in wet])


def roll(bus, notes, bar, beat, g, spread=0.045, length=3.5):
    for i, nt in enumerate(notes):
        bus.add(mem(piano, nt, length), 0, g * (1 - 0.05 * i), at=bar_t(bar, beat) + i * spread)


def hats(bus, bars, mode, g, fn=None):
    fn = fn or hat_soft
    for bar in bars:
        pos = {"off8": [.5, 1.5, 2.5, 3.5], "all8": [i * .5 for i in range(8)],
               "16": [i * .25 for i in range(16)]}[mode]
        for p in pos:
            if mode == "16":
                w = 1.0 if p % 1 == 0.5 else 0.7 if p % 1 == 0 else 0.4
            elif mode == "all8":
                w = 0.7 if p % 1 == 0 else 1.0
            else:
                w = 1.0
            bus.add(mem(fn), bar * 4 + p, g * w, pan=0.25 if p % 1 else -0.2)


MOTIF = [(0, 0, "A4", 1), (0, 1, "D5", 1), (0, 2, "F#5", 2),
         (1, 0, "E5", 1.5), (1, 1.5, "F#5", 0.5), (1, 2, "E5", 2),
         (2, 0, "D5", 1), (2, 1, "F#5", 1), (2, 2, "A5", 1.5), (2, 3.5, "F#5", 0.5),
         (3, 0, "G5", 1), (3, 1, "F#5", 1), (3, 2, "D5", 2)]
PAT8 = [0, 1, 2, 3, 2, 1, 2, 3]
PAT16 = [0, 2, 1, 3, 2, 1, 3, 2, 0, 2, 1, 3, 2, 3, 1, 2]
BP_SOFT = [(0, "r", 2.4, .8), (2.5, "r", 1.2, .5), (3.5, "o", .4, .3)]
BP_8 = [(0, "r", .45, .8), (.5, "r", .45, .5), (1, "r", .45, .6), (1.5, "o", .45, .5),
        (2, "r", .45, .7), (2.5, "r", .45, .5), (3, "r", .45, .6), (3.5, "t", .45, .5)]


def bass_bar(bus, bar, ch, pat, g=1.0, cutoff=700):
    for beat, w, dur, gg in pat:
        note = {"r": ch["bass"], "o": ch["oct"], "t": ch["third"]}[w]
        bus.add(mem(bass_note, note, round(dur * BEAT, 3), cutoff), bar * 4 + beat, g * gg)


def master(mix, secs, target, ceil, lp=9500, soft=False, fade=None):
    mix = np.stack([filt(filt(ch, "highpass", 30, 2), "lowpass", lp, 2) for ch in mix[:, :N]])
    if fade:
        t = np.arange(N) / SR
        a, b = fade
        mix = mix * np.where(t < a, 1.0, np.cos(np.clip((t - a) / (b - a), 0, 1) * np.pi / 2) ** 2).astype(np.float32)
    path = HERE / "music.wav"
    g = 0.5 / np.abs(mix).max()
    c = 10 ** (ceil / 20)
    for _ in range(8):
        x = mix * g
        if soft:
            x = c * np.tanh(x / c)
        y = limit(x, ceil)[:, :N]
        sf.write(path, y.T, SR, subtype="PCM_24")
        lufs, pk = ebur128(path)
        print(f"gain {g:.3f}: {lufs:.2f} LUFS, true-peak {pk:.2f} dBFS")
        if abs(target - lufs) < 0.25:
            break
        g *= 10 ** ((target - lufs) / 20)
    print("section RMS (dBFS):")
    mono = y.mean(axis=0)
    for nm, a, b in secs:
        seg = mono[int(a * BAR * SR):int(b * BAR * SR)]
        print(f"  {nm:18s} bars {a:4g}-{b:4g} ({a * BAR:5.2f}-{b * BAR:5.2f}s)  "
              f"{20 * np.log10(np.sqrt(np.mean(seg ** 2)) + 1e-12):6.1f} dB")
    print(f"  samples {y.shape[1]} = {y.shape[1] / SR:.4f} s; last 0.3 s peak {np.abs(mono[-int(.3 * SR):]).max():.6f}")

BPM = 128
BEAT = 60 / BPM          # 0.46875 exact
BAR = 4 * BEAT           # 1.875
DUR = 45.0
N = int(SR * DUR)
seed(45004)
prng = np.random.default_rng(13)
PROG = ["Dm", "Dm", "Dm", "Dm", "Dm", "Dm", "Bb", "A", "Dm", "Bb", "Gm", "Dm", "Bb", "C", "Dm", "Bb",
        "Gm", "A", "Dm", "Bb", "Dm", "A"]
STOP = 22                # bars; hard stop at 41.25 s


def tom(f0, length=0.7):
    n = int(length * SR)
    t = tvec(n)
    f = f0 * (1 + 1.2 * np.exp(-t / 0.03))
    body = np.sin(2 * np.pi * np.cumsum(f) / SR) * np.exp(-t / 0.28)
    thud = filt(rng.standard_normal(n) * np.exp(-t / 0.012), "lowpass", 600) * 0.8
    skin = np.sin(2 * np.pi * f0 * 2.3 * t) * np.exp(-t / 0.06) * 0.25
    return np.tanh(1.7 * (body + thud + skin)) * (1 - np.exp(-t / 0.0015)) * np.clip((length - t) / 0.05, 0, 1)


def impact():
    return layer(sub_drop(2.0, 75, 28) * 0.9, kick() * 0.8, filt(crash(2.2), "lowpass", 9000) * 0.25)


def braam(note, length=3.4):
    n = int(length * SR)
    t = tvec(n)
    f = hz(note)
    fc = 220 + 2500 * (1 - np.exp(-t / 0.12)) * np.exp(-t / 1.1)
    env = (1 - np.exp(-t / 0.02)) * np.exp(-t / 1.4) * np.clip((length - t) / 0.4, 0, 1)
    out = []
    for ch in range(2):
        x = np.zeros(n)
        for fm in (1.0, 1.4983):                       # root + fifth
            for k in range(7):
                x += saw(f * fm * 2 ** ((k - 3) * 12 / 1200), n, phase0=float(prng.uniform()))
        x += 2 * saw(f * 0.5, n, phase0=float(prng.uniform()))
        x = sweep_filter(x, fc, "lowpass", 0.9, 128)
        out.append(np.tanh(1.4 * x / 8) * env)
    return np.stack(out).astype(np.float32)


def brass(notes, length=0.4):
    n = int(length * SR)
    t = tvec(n)
    x = sum(saw(hz(nt) * 2 ** (c / 1200), n, phase0=float(prng.uniform())) for nt in notes for c in (-9, 0, 9))
    x = sweep_filter(x, 400 + 2600 * np.exp(-t / 0.12), "lowpass", 1.0, 128)
    return np.tanh(x / len(notes) * 0.7) * adsr(n, 0.01, 0.15, 0.7, 0.08)


def main():
    B = {k: mkbus() for k in ("pad", "sub", "bass", "kick", "tom", "clap", "hat", "tick", "brass", "braam", "fx", "imp",
                              "res", "rfx")}
    kick_t = []
    curve = mkcurve()

    def kk(bar, beats, g):
        for bt in beats:
            B["kick"].add(mem(kick, False), bar * 4 + bt, g)
            kick_t.append(bar_t(bar, bt))

    TOMS = {"lo": 62.0, "mid": 82.0, "hi": 110.0}

    def tm(bar, beat, kind, g):
        B["tom"].add(mem(tom, TOMS[kind]), bar * 4 + beat, g, pan={"lo": 0.0, "mid": -0.25, "hi": 0.25}[kind])

    # pads / sub: tension drone, then power
    for bar in range(STOP):
        ch = chord(PROG[bar])
        at = bar_t(bar) - 0.1
        L = BAR + 0.8
        B["pad"].add(mem(pad, tuple(ch["pad"]), L, 900 if bar < 18 else 2200, 0.4, 0.5, 5, 0.15), 0, 1.0, at=at)
        B["sub"].add(mem(subpad, ch["sub"], L), 0, 1.0, at=at)

    # bass 16th ostinato
    ACC = [1.0, .5, .7, .55, .85, .5, .7, .55, .9, .5, .75, .55, .85, .5, .8, .6]
    OCT = {3, 6, 10, 14}
    for bar in range(2, STOP):
        ch = chord(PROG[bar])
        cut = 350 if bar < 4 else 550 if bar < 8 else 850 if bar < 12 else 1300 if bar < 18 else 1900
        for s in range(16):
            note = ch["oct"] if s in OCT else ch["bass"]
            B["bass"].add(mem(bass_note, note, 0.11, cut), bar * 4 + s * 0.25, ACC[s])

    # ticking
    for bar in range(0, 17):
        step = 0.5 if bar < 8 else 0.25
        g = 0.25 + 0.03 * bar
        for i in range(int(4 / step)):
            B["tick"].add(mem(tick), bar * 4 + i * step, g * (1.0 if i % 2 == 0 else 0.6), pan=0.3 if i % 2 else -0.3)
    # breath: ticks accelerate through bar 18 (index 17)
    pos, step = 17 * 4.0, 0.5
    while pos < 18 * 4 - 0.01:
        B["tick"].add(mem(tick), pos, 0.45 + 0.4 * (pos - 68) / 4, pan=float(prng.uniform(-.4, .4)))
        step = max(0.0625, step * 0.88)
        pos += step

    # drums
    for bar in range(0, 4):
        tm(bar, 0, "lo", 0.9 if bar else 1.0)
        tm(bar, 2.5, "lo", 0.4)
        B["fx"].add(mem(sub_drop, 1.6, 62, 30), bar * 4, 0.35 + 0.05 * bar)
    for bar in range(4, 8):
        kk(bar, (0, 2), 0.7)
        tm(bar, 0, "lo", 0.7)
        tm(bar, 1.5, "mid", 0.4)
        tm(bar, 3.5, "mid", 0.4)
        hats(B["hat"], [bar], "off8", 0.10, hat)
    for bar in range(8, 17):
        kk(bar, (0, 1, 2, 3), 0.8)
        tm(bar, 0, "lo", 0.7)
        tm(bar, 1.5, "mid", 0.4)
        tm(bar, 2.5, "hi", 0.4)
        tm(bar, 3.5, "mid", 0.45)
        for bt in (1, 3):
            B["clap"].add(mem(clap, True), bar * 4 + bt, 0.35)
        hats(B["hat"], [bar], "16" if bar >= 12 else "all8", 0.09, hat)
    for bar in range(18, STOP):
        kk(bar, (0, 1, 2, 3), 0.9)
        if bar == STOP - 1:
            kk(bar, (3.5,), 0.7)
        for i in range(8):
            tm(bar, i * 0.5, ("lo", "mid", "hi", "mid")[i % 4] if i else "lo", 0.85 if i % 2 == 0 else 0.5)
        for bt in (1, 3):
            B["clap"].add(mem(clap, True), bar * 4 + bt, 0.45)
        hats(B["hat"], [bar], "16", 0.11, hat)
    # tom rolls into 9, 13
    for bar in (7, 11):
        for i in range(8):
            tm(bar, 2 + i * 0.25, "mid" if i < 6 else "hi", 0.25 + 0.07 * i)
    # impacts / risers / braams
    for bar in (4, 8, 12, 18):
        B["imp"].add(mem(impact), bar * 4, 0.9 if bar == 18 else 0.7)
    for bar in (3, 7, 11, 17):
        B["fx"].add(filt(riser(BAR), "lowpass", 7000 if bar < 17 else 9000, 2), bar * 4, 0.12 if bar < 17 else 0.3)
    B["braam"].add(mem(braam, "D2"), 11 * 4, 0.85)
    B["braam"].add(mem(braam, "D2"), 20 * 4, 1.0)
    B["imp"].add(mem(impact), 20 * 4, 0.9)
    kick_t.extend([bar_t(11), bar_t(20)])

    # brass stabs
    for bar in list(range(12, 17)) + list(range(18, STOP)):
        ch = chord(PROG[bar])
        b0 = midi(ch["bass"]) + 12
        notes = [note_name(b0 + x) for x in (0, 7, 12, 15 if ch["third"] == note_name(midi(ch["bass"]) + 3) else 16)]
        pat = [0, 0.75, 1.5, 2.0, 3.0] if bar >= 18 else [0, 1.5, 3.0]
        for p in pat:
            B["brass"].add(mem(brass, tuple(notes), 0.35), bar * 4 + p, (0.55 if bar >= 18 else 0.4) * (1 if p == 0 else .7),
                           pan=0.1)

    # resolve: single soft D major chord under the logo (bars 23-24)
    R0 = bar_t(STOP)
    B["res"].add(mem(pad, ("D3", "A3", "F#4", "A4", "E5"), 5.0, 1800, 0.18, 2.6, 5, 0.12), 0, 1.0, at=R0)
    B["res"].add(mem(subpad, "D2", 4.5), 0, 0.8, at=R0)
    for n_, g in (("D5", .30), ("F#5", .2), ("A5", .16)):
        B["res"].add(mem(bell, n_, 3.5), 0, g, at=R0 + 0.02, pan=0.15)

    c_pad = curve((0, .5), (4, .55), (12, .8), (17, .8), (18, 1.0), (22, 1.0))
    c_sub = curve((0, .6), (4, .7), (18, 1.0))
    c_bass = curve((2, .35), (4, .6), (8, .8), (12, .95), (18, 1.1), (22, 1.1))
    duck = duck_curve(kick_t, N, 0.3, 0.12)
    gate = curve((16.995, 1), (17.0, 0.02), (17.99, 0.02), (17.995, 1), (21.998, 1), (22.0, 0))

    def sc(name, g=1.0, cv=None, dk=False):
        x = B[name].x[:, :N] * np.float32(g)
        if cv is not None:
            x = x * cv
        if dk:
            x = x * duck
        return x

    pd, sb, bs = sc("pad", 0.45, c_pad, True), sc("sub", 0.30, c_sub, True), sc("bass", 0.62, c_bass, True)
    kkk, tms, cl = sc("kick", 0.85), sc("tom", 0.85), sc("clap", 0.6)
    ht, br, brm = sc("hat", 0.5), sc("brass", 0.8), sc("braam", 0.75)
    imp = sc("imp", 0.7)
    tk = sc("tick", 0.9)
    fx = sc("fx", 0.8)
    main_dry = pd + sb + bs + kkk + tms + cl + ht + br + brm + imp
    send = pd * 0.3 + tms * 0.45 + br * 0.5 + brm * 0.25 + imp * 0.3 + cl * 0.5
    main = (main_dry + 0.35 * verb(send, 1.8, 0.45, 3, 7000, 200)) * gate
    mix = main + (tk + fx) * gate.clip(0.02, 1) * 0 + tk + fx           # ticks and riser live through the breath
    mix = mix * np.where(np.arange(N) / SR < STOP * BAR, 1.0, 0.0).astype(np.float32)
    r = sc("res", 0.7)
    mix = mix + r + 0.5 * verb(r, 3.0, 1.0, 9, 6000, 200)
    master(mix, [("tension", 0, 4), ("drive", 4, 8), ("build", 8, 12), ("heavy", 12, 17), ("breath", 17, 18),
                 ("full power", 18, 22), ("resolve", 22, 24)], -15.5, -3.3, 12000, True, (43.0, 44.8))


if __name__ == "__main__":
    main()
