# /// script
# requires-python = ">=3.10"
# dependencies = ["librosa>=0.10", "numpy", "scipy", "soundfile"]
# ///
"""BDO "Tallene forteller": 75 s, 96 bpm (bar 2.5 s, 30 bars), B minor -> D major.

    cd <repo> && MOTION_REEL_SKILL=<skill> uv run bdo-tallene/audio/score.py

  bars 1-7    muted pulse, clicky hats, round bass, piano stabs (B minor)
  bars 8-12   (17.5-30 s) night interlude: Em/Gm/F#m colour, dark pad, sparse pulse
  bars 13-26  (30-65 s) bright and confident, D major, motif at 55 s
  bars 27-30  (65-75 s) warm end, long D chord rings out
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

BPM = 96
BEAT = 60 / BPM
BAR = 4 * BEAT
DUR = 75.0
N = int(SR * DUR)
seed(75002)
prng = np.random.default_rng(5)
PROG = ["Bm", "Bm", "G", "G", "Bm", "Bm", "G", "Em", "Gm", "Em", "F#m", "A",
        "D", "A", "Bm", "G", "D", "A", "Bm", "G", "Em", "A", "D", "A", "Bm", "G", "D", "G", "D", "D"]
HATP = [0.0, 0.35, 1.0, 0.35, 0.0, 0.35, 1.0, 0.5, 0.0, 0.35, 1.0, 0.35, 0.0, 0.5, 1.0, 0.7]
BP_ROUND = [(0, "r", 1.0, .9), (1.5, "r", .4, .55), (2.5, "r", .8, .7), (3.5, "t", .4, .4)]


def main():
    B = {k: mkbus() for k in ("padD", "padB", "padF", "sub", "pn", "bell", "arp", "lead", "bass", "kick", "hat", "fx")}
    kick_t = []

    def kk(bar, beats, g):
        for bt in beats:
            B["kick"].add(mem(kick, True), bar * 4 + bt, g)
            kick_t.append(bar_t(bar, bt))

    for bar in range(30):
        ch = chord(PROG[bar])
        at = bar_t(bar) - 0.15
        L = BAR + 1.4
        if bar < 26:
            B["padD"].add(mem(pad, tuple(ch["pad"]), L, 650, 0.7, 1.2), 0, 1.0, at=at)
            B["sub"].add(mem(subpad, ch["sub"], L), 0, 1.0, at=at)
            if bar >= 12:
                B["padB"].add(mem(pad, tuple(ch["padO"]), L, 3000, 0.5, 1.2), 0, 1.0, at=at)
        elif bar in (26, 27):
            B["padF"].add(mem(pad, tuple(ch["padO"]), L + 0.4, 1700, 0.8, 1.6), 0, 1.0, at=at)
            B["sub"].add(mem(subpad, ch["sub"], L), 0, 1.0, at=at)
    B["padF"].add(mem(pad, tuple(chord("D")["padO"]), 8.5, 2000, 0.9, 4.5), 0, 1.0, at=bar_t(28) - 0.15)
    B["sub"].add(mem(subpad, "D1", 8.0), 0, 1.0, at=bar_t(28) - 0.15)

    # muted pulse
    for bar in range(0, 7):
        kk(bar, (0, 1, 2, 3), 0.5)
    for bar in range(7, 12):
        kk(bar, (0, 2) if bar % 2 == 0 else (0,), 0.4)
    for bar in range(12, 26):
        kk(bar, (0, 1, 2, 3), 0.62 if bar < 22 else 0.7)
    kk(26, (0, 2), 0.45)
    kk(27, (0,), 0.4)
    # clicky hats (16th grid, sparse)
    for bar in range(2, 28):
        if bar < 7:
            g, sel = 0.06 if bar < 4 else 0.09, 0.5
        elif bar < 12:
            g, sel = 0.045, 0.3
        elif bar < 26:
            g, sel = 0.11 if bar < 22 else 0.13, 1.0
        else:
            g, sel = 0.05, 0.4
        for s in range(16):
            w = HATP[s]
            if w == 0 or w < 1 - sel * 1.2 and s % 4 != 2:
                continue
            if bar >= 22 and bar < 26 and s % 2 == 1:
                w *= 1.0
            B["hat"].add(mem(click), bar * 4 + s * 0.25, g * w * (1 + 0.1 * prng.standard_normal()), pan=0.3 if s % 2 else -0.3)
    # round bass
    for bar in range(4, 28):
        ch = chord(PROG[bar])
        cut = 380 if 7 <= bar < 12 else 480
        bass_bar(B["bass"], bar, ch, BP_ROUND, 1.0, cut)
    # piano-like stabs, syncopated 3-3-2
    for bar in range(4, 28):
        k = chord(PROG[bar])["key"]
        if bar < 7:
            pos, g, voc, ln = ([1.5] if bar % 2 == 0 else [1.5, 3.0]), 0.30, k[1:4], 0.9
        elif bar < 12:
            pos, g, voc, ln = [2.5], 0.30, [tr(x, -12) for x in k[1:4]], 1.6
        elif bar < 26:
            pos, g, voc, ln = [0, 1.5, 3.0], 0.34, [tr(x, 0) for x in k[1:5]], 1.0
        else:
            pos, g, voc, ln = [0, 1.5], 0.30, k[1:5], 1.4
        for p in pos:
            for i, nt in enumerate(voc):
                B["pn"].add(mem(piano, nt, ln), 0, g * (1 - 0.06 * i) * (1.0 if p == 0 else 0.8),
                            at=bar_t(bar, p) + i * 0.008, pan=-0.15 + 0.1 * i)
    # night colour: low sparse bell notes
    for bar in range(7, 12):
        B["bell"].add(mem(bell, tr(chord(PROG[bar])["arp"][1], 0), 3.0), bar * 4 + 0.0, 0.2, pan=-0.2)
        B["bell"].add(mem(bell, tr(chord(PROG[bar])["arp"][2], 12), 3.0), bar * 4 + 2.0, 0.12, pan=0.2)
    # precise 8th pluck arp in the bright section
    for bar in range(12, 27):
        notes = [tr(n, 0 if bar < 22 else 12) for n in chord(PROG[bar])["arp"]]
        for s in range(8):
            idx = PAT8[s]
            B["arp"].add(mem(pluck, (notes[idx],), 0.26, 3000), bar * 4 + s * 0.5,
                         (0.9 if s % 2 == 0 else 0.55) * (1 + 0.05 * prng.standard_normal()), pan=[-.3, -.1, .1, .3][idx])
    # bright arrival
    for n_ in ("D5", "F#5", "A5"):
        B["bell"].add(mem(bell, n_, 3.0), 12 * 4, 0.28, pan=0.2)
    B["fx"].add(filt(riser(BAR * 1.0), "lowpass", 6000, 2), 11 * 4, 0.07)
    B["fx"].add(mem(sub_drop, 1.4, 70, 34), 12 * 4, 0.35)
    # motif in the confident section (D at bar 22) and the end
    for bo, beat, note, dur in MOTIF:
        pos = (22 + bo) * 4 + beat
        B["bell"].add(mem(bell, note, round(min(max(dur * BEAT * 1.8, 1.4), 3.5), 2)), pos, 0.38, pan=0.1)
        B["pn"].add(mem(piano, note, 3.0), pos, 0.42, pan=-0.1)
        B["lead"].add(mem(lead, note, round(dur * BEAT * 1.05, 3)), pos, 0.22, pan=0.05)
    B["fx"].add(filt(riser(BAR), "lowpass", 5500, 2), 21 * 4, 0.07)
    B["fx"].add(mem(sub_drop, 1.4, 70, 34), 22 * 4, 0.35)
    B["fx"].add(mem(sub_drop, 1.8, 70, 34), 26 * 4, 0.40)
    roll(B["pn"], ["D3", "A3", "E4", "F#4", "A4", "D5"], 26, 0, 0.45, 0.05, 6.0)
    for bo, beat, note in [(26, 0, "A4"), (26, 1, "D5"), (26, 2, "F#5"), (27, 0, "E5"), (27, 1.5, "F#5"), (27, 2, "E5"),
                           (28, 0, "D5")]:
        pos = bo * 4 + beat
        B["bell"].add(mem(bell, note, 3.5), pos, 0.36, pan=0.1)
        if pos > 26 * 4:
            B["pn"].add(mem(piano, note, 4.5), pos, 0.36, pan=-0.1)
    B["bell"].add(mem(bell, "D6", 3.5), 28 * 4, 0.16, pan=0.3)

    curve = mkcurve()
    c_padD = curve((0, .45), (4, .6), (7, .75), (12, .5), (22, .4), (26, 0))
    c_padB = curve((12, .9), (22, 1.0), (26, .8), (26.5, 0))
    c_sub = curve((0, .4), (7, .8), (12, .6), (26, .7), (30, .3))
    c_arp = curve((12, .9), (22, 1.0), (26, .6), (27, 0))
    c_hat = curve((0, 1), (30, 1))
    c_kick = curve((0, .8), (7, .7), (12, 1.0), (26, .9), (28, 0))
    duck = duck_curve(kick_t, N, 0.25, 0.2)

    def sc(name, g=1.0, cv=None, dk=False):
        x = B[name].x[:, :N] * np.float32(g)
        if cv is not None:
            x = x * cv
        if dk:
            x = x * duck
        return x

    ax = sc("arp", 0.7, c_arp)
    d = int(BEAT * 0.75 * SR)
    a0, a1 = ax[0].copy(), ax[1].copy()
    ax[1, d:] += 0.30 * a0[:-d]
    ax[0, 2 * d:] += 0.14 * a1[:-2 * d]
    padD = sc("padD", 0.55, c_padD, True)
    padB = sc("padB", 0.48, c_padB, True)
    padF = sc("padF", 0.55)
    sub = sc("sub", 0.28, c_sub, True)
    pn, bl, ld = sc("pn", 0.95), sc("bell", 0.85), sc("lead", 0.85)
    bs = sc("bass", 0.62, None, True)
    kkk, ht, fx = sc("kick", 0.80, c_kick), sc("hat", 0.55), sc("fx", 0.60)
    dry = padD + padB + padF + sub + pn + bl + ax + ld + bs + kkk + ht + fx
    send = (padD + padB + padF) * 0.28 + pn * 0.55 + bl * 0.9 + ax * 0.4 + ld * 0.5 + fx * 0.5
    mix = dry + 0.40 * verb(send, 2.4, 0.7, 7)
    master(mix, [("intro (Bm)", 0, 7), ("night interlude", 7, 12), ("bright", 12, 22), ("confident+motif", 22, 26),
                 ("warm end", 26, 30)], -17.0, -3.3, 9500, False, (71.5, 74.5))


if __name__ == "__main__":
    main()
