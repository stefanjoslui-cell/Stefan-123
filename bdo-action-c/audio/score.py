# /// script
# requires-python = ">=3.10"
# dependencies = ["librosa>=0.10", "numpy", "scipy", "soundfile"]
# ///
"""BDO action/trailer: 60 s, 128 bpm (32 bars), D minor -> D major. 3D trailer score.

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
DUR = 60.0
N = int(SR * DUR)
seed(60003)
prng = np.random.default_rng(13)
PROG = ["Dm"] * 4 + ["Dm", "Dm", "Bb", "A", "Dm", "Bb", "Gm", "Dm", "Dm", "Bb",
        "F", "Gm", "Bb", "C", "A", "Dm", "Dm", "Bb", "Gm", "A", "D", "A", "G", "D", "D", "D", "D", "D"]
STOP = 29                # bars; hard stop at bar 29 beat 3.5


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


def pluck(note, length=0.2):
    n = int(length * SR)
    t = tvec(n)
    f = hz(note)
    x = saw(f * 2 ** (-7 / 1200), n, phase0=0.1) + saw(f * 2 ** (7 / 1200), n, phase0=0.6)
    x = sweep_filter(x, 500 + 3800 * np.exp(-t / 0.07), "lowpass", 1.0, 128)
    return (x * 0.5 * np.exp(-t / 0.09) * (1 - np.exp(-t / 0.002)) * np.clip((length - t) / 0.02, 0, 1)).astype(np.float32)


def stab_notes(ch):
    b0 = midi(ch["bass"]) + 12
    return tuple(note_name(b0 + x) for x in (0, 7, 12, 15 if ch["third"] == note_name(midi(ch["bass"]) + 3) else 16))


def main():
    B = {k: mkbus() for k in ("pad", "sub", "bass", "kick", "tom", "clap", "hat", "tick", "brass", "braam", "fx", "imp",
                              "res", "arp", "lead", "pn")}
    kick_t = []
    curve = mkcurve()

    def kk(bar, beats, g):
        for bt in beats:
            B["kick"].add(mem(kick, False), bar * 4 + bt, g)
            kick_t.append(bar_t(bar, bt))

    TOMS = {"lo": 62.0, "mid": 82.0, "hi": 110.0}

    def tm(bar, beat, kind, g):
        B["tom"].add(mem(tom, TOMS[kind]), bar * 4 + beat, g, pan={"lo": 0.0, "mid": -0.25, "hi": 0.25}[kind])

    def troll(bar, b0=2.0):
        for i in range(int((4 - b0) / 0.25)):
            tm(bar, b0 + i * 0.25, "mid" if i < 6 else "hi", 0.25 + 0.07 * i)

    def clapb(bar, bts, g):
        for bt in bts:
            B["clap"].add(mem(clap, True), bar * 4 + bt, g)

    def stabs(bar, pat, g):
        notes = stab_notes(chord(PROG[bar]))
        for p in pat:
            B["brass"].add(mem(brass, notes, 0.35), bar * 4 + p, g * (1 if p == 0 else .7), pan=0.1)

    def arp(bar, pat, g, up=0):
        ch = chord(PROG[bar])
        for s, i in enumerate(pat):
            nt = note_name(midi(ch["arp"][i]) + up)
            B["arp"].add(mem(pluck, nt, 0.2), bar * 4 + s * (4 / len(pat)), g * (1 if s % 4 == 0 else .6),
                         pan=0.35 if s % 2 else -0.35)

    # pads / sub
    for bar in range(STOP):
        ch = chord(PROG[bar])
        at = bar_t(bar) - 0.1
        cut = 900 if bar < 8 else 1400 if bar < 14 else 1800 + 160 * (bar - 14) if bar < 19 else 2400 if bar < 24 else 3800
        B["pad"].add(mem(pad, tuple(ch["pad"]), BAR + 0.8, cut, 0.4, 0.5, 5, 0.15, bar >= 24), 0, 1.0, at=at)
        B["sub"].add(mem(subpad, ch["sub"], BAR + 0.8), 0, 1.0, at=at)

    # bass
    ACC = [1.0, .5, .7, .55, .85, .5, .7, .55, .9, .5, .75, .55, .85, .5, .8, .6]
    OCT = {3, 6, 10, 14}
    for bar in range(4, STOP):
        ch = chord(PROG[bar])
        if bar in (12, 13):
            bass_bar(B["bass"], bar, ch, BP_SOFT, 1.0, 500)
        elif bar == 14:
            bass_bar(B["bass"], bar, ch, BP_8, 1.0, 700)
        else:
            cut = 550 if bar < 8 else 850 if bar < 12 else 1000 + 220 * (bar - 15) if bar < 19 else 2200 if bar < 24 else 1900
            for s in range(16):
                note = ch["oct"] if s in OCT else ch["bass"]
                B["bass"].add(mem(bass_note, note, 0.11, cut), bar * 4 + s * 0.25, ACC[s])

    # ticking / clicks
    for bar in range(0, 22):
        step = 0.5 if bar < 8 or bar in (12, 13) else 0.25
        g = (0.25 + 0.025 * bar) * (0.6 if bar in (12, 13) else 1.0)
        for i in range(int(4 / step)):
            B["tick"].add(mem(tick), bar * 4 + i * step, g * (1.0 if i % 2 == 0 else 0.6), pan=0.3 if i % 2 else -0.3)
    for bar in range(19, 23):                       # clicky precision, bars 20-23
        for i in range(16 if bar < 22 else 8):
            B["tick"].add(mem(click), bar * 4 + i * 0.25 + 0.125 * (bar > 19), 0.25, pan=-0.5 if i % 2 else 0.5)
    pos, step = 90.0, 0.9                              # breath: ticks accelerate (bar 23 beat 3 -> bar 24 beat 3)
    while pos < 94 - 0.01:
        B["tick"].add(mem(tick), pos, 0.3 + 0.35 * (pos - 90) / 4, pan=float(prng.uniform(-.4, .4)))
        step = max(0.125, step * 0.86)
        pos += step
    for i in range(8):                                 # pickup ticks bar 24 beats 3-4
        B["tick"].add(mem(tick), 94 + i * 0.25, 0.5 + 0.05 * i, pan=0.3 if i % 2 else -0.3)
    for bar in range(24, 28):
        for i in range(8):
            B["tick"].add(mem(tick), bar * 4 + i * 0.5, 0.2 if i % 2 == 0 else 0.12)

    # drums
    for bar in (0, 2):
        tm(bar, 0, "lo", 1.0)
        B["fx"].add(mem(sub_drop, 1.6, 62, 30), bar * 4, 0.4)
    for bar in range(4, 8):
        kk(bar, (0, 2), 0.7)
        tm(bar, 0, "lo", 0.7)
        tm(bar, 1.5, "mid", 0.4)
        tm(bar, 3.5, "mid", 0.4)
        hats(B["hat"], [bar], "off8", 0.10, hat)
    for bar in range(8, 12):
        kk(bar, (0, 1, 2, 3), 0.8)
        tm(bar, 0, "lo", 0.7)
        tm(bar, 1.5, "mid", 0.4)
        tm(bar, 2.5, "hi", 0.4)
        tm(bar, 3.5, "mid", 0.45)
        clapb(bar, (1, 3), 0.35)
        hats(B["hat"], [bar], "all8", 0.09, hat)
    for bar in (12, 13):                               # half-time: space
        kk(bar, (0, 2.5) if bar == 12 else (0,), 0.8)
        tm(bar, 0, "lo", 0.6)
        if bar == 13:
            tm(bar, 2, "mid", 0.35)
        clapb(bar, (2,), 0.4)
        hats(B["hat"], [bar], "off8", 0.05, hat)
        for bt in (0, 2):
            roll(B["pn"], chord(PROG[bar])["key"], bar, bt, 0.5 if bt == 0 else 0.35, 0.06, 3.2)
    for bar in range(14, 19):                          # build
        k = bar - 14
        kk(bar, (0, 2) if k < 2 else (0, 1, 2, 3), 0.65 + 0.05 * k)
        tm(bar, 0, "lo", 0.6 + 0.05 * k)
        if k >= 1:
            tm(bar, 1.5, "mid", 0.4)
            tm(bar, 2.5, "hi", 0.4)
        if k >= 2:
            tm(bar, 3.5, "mid", 0.45)
        clapb(bar, (1, 3) if k >= 1 else (2,), 0.3 + 0.03 * k)
        hats(B["hat"], [bar], "all8" if k < 1 else "16", 0.07 + 0.01 * k, hat)
        roll(B["pn"], chord(PROG[bar])["key"], bar, 0, 0.45, 0.05, 3.0)
        if k >= 1:
            stabs(bar, [0, 1.5, 3.0] if k < 3 else [0, 0.75, 1.5, 3.0], 0.3 + 0.05 * k)
        arp(bar, PAT8, 0.14 + 0.02 * k)
    for bar in range(19, 24):                          # full / audit
        kk(bar, (0, 1, 2, 3), 0.9 if bar == 19 else 0.8)
        if bar == 19:
            for i in range(8):
                tm(bar, i * 0.5, ("lo", "mid", "hi", "mid")[i % 4] if i else "lo", 0.85 if i % 2 == 0 else 0.5)
        else:
            tm(bar, 0, "lo", 0.5)
            tm(bar, 3.5, "mid", 0.3)
        clapb(bar, (1, 3), 0.4)
        hats(B["hat"], [bar], "16", 0.10, hat)
        stabs(bar, [0, 0.75, 1.5, 2.0, 3.0] if bar == 19 else [0, 0.75, 1.5, 2.5, 3.0], 0.55 if bar == 19 else 0.4)
        arp(bar, PAT16, 0.2)
    for bar in range(24, 28):                          # customer insight: lift
        kk(bar, (0, 1, 2, 3), 0.95)
        if bar == 27:
            kk(bar, (3.5,), 0.7)
        for i in range(8):
            tm(bar, i * 0.5, ("lo", "mid", "hi", "mid")[i % 4] if i else "lo", 0.8 if i % 2 == 0 else 0.45)
        clapb(bar, (1, 3), 0.5)
        hats(B["hat"], [bar], "16", 0.12, hat)
        stabs(bar, [0, 0.75, 1.5, 2.0, 3.0], 0.6)
        arp(bar, PAT8 * 2, 0.22, 12)
        roll(B["pn"], chord(PROG[bar])["key"], bar, 0, 0.5, 0.05, 3.0)
        roll(B["pn"], chord(PROG[bar])["key"], bar, 2, 0.4, 0.05, 3.0)
    kk(23, (2, 2.5, 3, 3.5), 0.7)                      # pickup into bar 25
    for i in range(8):
        pass
    for bar in (7, 10, 18, 23, 26):
        troll(bar, 2.0)
    # lead motif (bars 25-28)
    for bar, bt, nt, d in MOTIF:
        B["lead"].add(mem(lead, nt, round(d * BEAT * 0.97, 3)), (24 + bar) * 4 + bt, 0.55, pan=0.1)
        B["lead"].add(mem(lead, tr(nt, -12), round(d * BEAT * 0.97, 3)), (24 + bar) * 4 + bt, 0.25, pan=-0.1)
    # accents bar 29 beats 1-3
    for i, g in enumerate((1.0, 0.85, 0.9)):
        B["imp"].add(mem(impact), 28 * 4 + i, g * 0.8)
        B["brass"].add(mem(brass, stab_notes(chord("D")), 0.45), 28 * 4 + i, 0.8)
        tm(28, i, "lo", 0.9)
        kk(28, (i,), 0.9)

    # impacts / risers / braams
    for bar, g in ((4, .7), (8, .7), (19, .9), (24, .8)):
        B["imp"].add(mem(impact), bar * 4, g)
    B["fx"].add(mem(sub_drop, 1.8, 62, 30), 12 * 4, 0.35)
    B["imp"].add(mem(impact), 27 * 4, 1.0)
    for bar, L_, g in ((0, 4 * BAR, .10), (3, BAR, .12), (7, BAR, .14), (10, BAR, .14), (14, 5 * BAR, .20), (26, BAR, .3)):
        B["fx"].add(filt(mem(riser, L_), "lowpass", 8000, 2), bar * 4, g)
    B["fx"].add(filt(mem(riser, 2 * BEAT), "lowpass", 9000, 2), 94, 0.35)
    for bar, g in ((11, .85), (19, 1.0), (27, 1.0)):
        B["braam"].add(mem(braam, "D2"), bar * 4, g)
        kick_t.append(bar_t(bar))

    # resolve (bars 30-32)
    R0 = bar_t(STOP)
    B["res"].add(mem(pad, ("D3", "A3", "F#4", "A4", "E5"), 5.0, 1800, 0.18, 2.6, 5, 0.12), 0, 1.0, at=R0)
    B["res"].add(mem(subpad, "D2", 4.5), 0, 0.8, at=R0)
    for n_, g in (("D5", .30), ("F#5", .2), ("A5", .16)):
        B["res"].add(mem(bell, n_, 3.5), 0, g, at=R0 + 0.02, pan=0.15)

    c_pad = curve((0, .45), (4, .5), (8, .7), (12, .9), (14, .6), (19, 1.0), (24, 1.2))
    c_sub = curve((0, .6), (4, .7), (12, .9), (19, 1.0), (24, 1.1))
    c_bass = curve((4, .45), (8, .7), (12, .95), (14, .9), (19, 1.1), (24, 1.15))
    duck = duck_curve(kick_t, N, 0.3, 0.12)
    gate = curve((22.495, 1), (22.5, 0.02), (23.495, 0.02), (23.5, 1), (28.86, 1), (28.875, 0))

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
    imp, ar_, ld, pn = sc("imp", 0.7), sc("arp", 0.5), sc("lead", 0.5), sc("pn", 0.5)
    tk, fx = sc("tick", 0.9), sc("fx", 0.8)
    main_dry = pd + sb + bs + kkk + tms + cl + ht + br + brm + imp + ar_ + ld + pn
    send = pd * 0.3 + tms * 0.45 + br * 0.5 + brm * 0.25 + imp * 0.3 + cl * 0.5 + ar_ * 0.35 + ld * 0.5 + pn * 0.5
    main = (main_dry + 0.35 * verb(send, 1.8, 0.45, 3, 7000, 200)) * gate
    mix = main + tk + fx
    mix = mix * np.where(np.arange(N) / SR < STOP * BAR, 1.0, 0.0).astype(np.float32)
    r = sc("res", 0.7)
    mix = mix + r + 0.5 * verb(r, 3.0, 1.0, 9, 6000, 200)
    master(mix, [("1-4 cold open", 0, 4), ("5-8 business", 4, 8), ("9-12 tech", 8, 12), ("13-14 half-time", 12, 14),
                 ("15-19 build", 14, 19), ("20-22 audit", 19, 22.5), ("breath", 22.5, 23.5), ("pickup", 23.5, 24),
                 ("25-28 insight", 24, 28), ("29 accents", 28, 29), ("30-32 resolve", 29, 32)],
           -16.0, -3.3, 12000, True, (56.0, 59.5))


if __name__ == "__main__":
    main()
