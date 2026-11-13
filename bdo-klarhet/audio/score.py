# /// script
# requires-python = ">=3.10"
# dependencies = ["librosa>=0.10", "numpy", "scipy", "soundfile"]
# ///
"""BDO "Fra stoey til klarhet": 60 s, 96 bpm (bar 2.5 s, 24 bars), D major.

    cd <repo> && MOTION_REEL_SKILL=<skill> uv run bdo-klarhet/audio/score.py

  bars 1-2    (0-5 s)      filtered noise haze + granular grains, no pitch
  bars 3-9    (5-22.5 s)   a pitch emerges, dark harmony Bm-G-Em-A
  bar 10      (22.5 s)     decisive brightening: D major, filter opens fully by 24 s
  bars 16-20  (37.5-50 s)  wide climax, motif
  bars 21-24  (50-60 s)    single sustained chord + bell
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
DUR = 60.0
N = int(SR * DUR)
seed(60003)
prng = np.random.default_rng(9)
PROG = [None, None, "Bm", "Bm", "G", "G", "Em", "Em", "A", "D", "A", "Bm", "G", "D", "A",
        "D", "A", "Bm", "G", "A", "D", "D", "D", "D"]


def grain(dur, cf):
    n = int(dur * SR)
    t = tvec(n)
    x = filt(rng.standard_normal(n), "bandpass", [cf * 0.8, cf * 1.25], 2)
    return x * (1 - np.exp(-t / 0.004)) * np.exp(-t / (dur / 3)) * np.clip((dur - t) / 0.01, 0, 1)


def haze(n, lo, hi, rate, ph):
    t = np.arange(n) / SR
    fc = np.exp(np.log(lo) + (np.log(hi) - np.log(lo)) * (0.5 + 0.5 * np.sin(2 * np.pi * rate * t + ph)))
    return sweep_filter(rng.standard_normal(n), fc, "bandpass", 0.8, 256)


def main():
    B = {k: mkbus() for k in ("padD", "padB", "padW", "padF", "sub", "tone", "pn", "bell", "arp", "lead", "bass",
                              "kick", "hat", "grain", "fx")}
    kick_t = []
    n_h = int(30 * SR)
    hz_ = np.stack([haze(n_h, 350, 4200, 0.07, 0.0), haze(n_h, 500, 5200, 0.05, 2.0)]).astype(np.float32)
    air = np.stack([filt(rng.standard_normal(int(32 * SR)), "highpass", 5500, 2) for _ in range(2)]).astype(np.float32)

    # grains: dense and pitchless at first, thinning as harmony arrives; a soft tick on every beat from t = 0
    t_ = 0.0
    while t_ < 9 * BAR:
        dens = 12 if t_ < 2 * BAR else 6 if t_ < 6 * BAR else 3
        cf = float(np.exp(prng.uniform(np.log(500), np.log(6500))))
        B["grain"].add(mem(grain, round(float(prng.uniform(0.04, 0.16)), 3), round(cf, -1)), 0, float(prng.uniform(0.3, 1.0)),
                       pan=float(prng.uniform(-0.9, 0.9)), at=t_)
        t_ += float(prng.exponential(1 / dens))
    for b in range(0, 36):
        B["grain"].add(mem(tick), b, 0.9 if b % 4 == 0 else 0.5, pan=0.2 if b % 2 else -0.2)

    # pads
    for bar in range(2, 20):
        ch = chord(PROG[bar])
        at = bar_t(bar) - 0.15
        L = BAR + 1.4
        if bar < 9:
            B["padD"].add(mem(pad, tuple(ch["pad"][:3]), L, 600, 1.0, 1.4), 0, 1.0, at=at)
            B["sub"].add(mem(subpad, ch["sub"], L), 0, 1.0, at=at)
        else:
            B["padD"].add(mem(pad, tuple(ch["pad"]), L, 700, 0.5, 1.2), 0, 1.0, at=at)
            B["padB"].add(mem(pad, tuple(ch["padO"]), L, 3200, 0.12 if bar == 9 else 0.5, 1.2), 0, 1.0, at=at)
            B["sub"].add(mem(subpad, ch["sub"], L), 0, 1.0, at=at)
        if bar >= 15:
            wn = tuple(ch["padO"] + [tr(ch["padO"][2], 12), tr(ch["padO"][3], 12)])
            B["padW"].add(mem(pad, wn, L, 4000, 0.5, 1.2, 6, 0.2, True), 0, 1.0, at=at)
    D = chord("D")
    B["padF"].add(mem(pad, tuple(D["padO"]), 11.0, 2200, 1.0, 5.5), 0, 1.0, at=bar_t(20) - 0.15)
    B["sub"].add(mem(subpad, "D1", 9.5), 0, 1.0, at=bar_t(20) - 0.15)

    # the pitch emerges: a pure tone fades up out of the haze at bar 3
    for nt, g in (("B3", 1.0), ("F#4", 0.5)):
        n = int(7.5 * SR)
        x = (np.sin(2 * np.pi * hz(nt) * tvec(n)) * env_ar(n, 2.0, 1.5)).astype(np.float32)
        B["tone"].add(x, 0, g, at=bar_t(2))
    B["bell"].add(mem(bell, "B4", 3.5), 2 * 4, 0.35, pan=0.1)
    B["bell"].add(mem(bell, "F#5", 3.5), 4 * 4 + 2, 0.2, pan=0.2)

    # bass, arps, piano, drums
    for bar in range(4, 20):
        bass_bar(B["bass"], bar, chord(PROG[bar]), BP_SOFT, 1.0, 450 if bar < 9 else 650)
    for bar in range(6, 20):
        notes = [tr(n, 0 if bar < 15 else 12) for n in chord(PROG[bar])["arp"]]
        cut = 1400 if bar < 9 else 3000
        for s in range(8):
            idx = PAT8[s]
            B["arp"].add(mem(pluck, (notes[idx],), 0.28, cut), bar * 4 + s * 0.5,
                         (1.0 if s % 2 == 0 else 0.65) * (1 + 0.05 * prng.standard_normal()), pan=[-.3, -.1, .1, .3][idx])
    for bar in range(9, 15, 2):
        k = chord(PROG[bar])["key"]
        roll(B["pn"], k[1:], bar, 0, 0.30, 0.05)
    for bar in range(15, 20):
        roll(B["pn"], chord(PROG[bar])["key"], bar, 0, 0.30, 0.04)
    for bar in range(12, 20):
        for bt in ((0, 1, 2, 3) if bar >= 15 else (0, 2)):
            B["kick"].add(mem(kick, True), bar * 4 + bt, 0.55 if bar >= 15 else 0.4)
            kick_t.append(bar_t(bar, bt))
    hats(B["hat"], range(12, 20), "off8", 0.10)
    # decisive brightening at bar 10
    B["fx"].add(filt(riser(BAR), "lowpass", 8000, 2), 8 * 4, 0.12)
    B["fx"].add(mem(sub_drop, 1.6, 70, 34), 9 * 4, 0.5)
    for n_ in ("D5", "F#5", "A5", "D6"):
        B["bell"].add(mem(bell, n_, 3.5), 9 * 4, 0.28, pan=0.2)
    B["fx"].add(filt(riser(BAR), "lowpass", 6000, 2), 14 * 4, 0.08)
    B["fx"].add(mem(sub_drop, 1.6, 70, 34), 15 * 4, 0.5)
    # climax motif (D at bar 15)
    for bo, beat, note, dur in MOTIF:
        pos = (15 + bo) * 4 + beat
        B["bell"].add(mem(bell, note, round(min(max(dur * BEAT * 1.8, 1.4), 3.5), 2)), pos, 0.42, pan=0.1)
        B["pn"].add(mem(piano, note, 3.5), pos, 0.46, pan=-0.1)
        B["lead"].add(mem(lead, note, round(dur * BEAT * 1.05, 3)), pos, 0.34, pan=0.05)
        B["bell"].add(mem(bell, tr(note, 12), 2.0), pos, 0.15, pan=0.25)
    for n_ in [tr(x, 12) for x in D["arp"][:3]]:
        B["bell"].add(mem(bell, n_, 3.0), 15 * 4, 0.3, pan=0.2)
    # ending: one sustained chord + one bell
    B["fx"].add(mem(sub_drop, 1.8, 70, 34), 20 * 4, 0.35)
    B["bell"].add(mem(bell, "D6", 4.0), 20 * 4, 0.34, pan=0.15)
    B["bell"].add(mem(bell, "A5", 4.0), 22 * 4 + 2, 0.14, pan=-0.15)

    curve = mkcurve()
    c_haze = curve((0, .9), (2, .9), (5, .45), (8, .22), (9, .12), (9.5, 0), (30, 0))
    c_air = curve((0, 0), (9, 0), (10.5, .05), (15, .07), (20, .03), (21, 0))
    c_grain = curve((0, 1), (2, 1), (9, .6), (9.1, 0))
    c_padD = curve((2, 0), (3.5, .5), (6, .8), (8.9, .9), (9, .25), (15, .35), (20, 0))
    c_padB = curve((9, 1.0), (15, 1.0), (20, .9), (20.3, 0))
    c_padW = curve((15, 1.1), (20, 1.1), (20.3, 0))
    c_sub = curve((3, .1), (9, .7), (15, .9), (20, .8), (24, .3))
    c_tone = curve((2, 0), (3, .7), (9, .7), (9.8, 0))
    c_bass = curve((4, .45), (9, .6), (9.01, 1.0), (20, 1.0), (20.3, 0))
    c_arp = curve((6, .4), (9, .5), (9.01, .8), (15, .9), (20, .8), (20.2, 0))
    duck = duck_curve(kick_t, N, 0.22, 0.25)

    def sc(name, g=1.0, cv=None, dk=False):
        x = B[name].x[:, :N] * np.float32(g)
        if cv is not None:
            x = x * cv
        if dk:
            x = x * duck
        return x

    hzv = hz_ * c_haze[:n_h][None] * 0.5
    haze_full = np.zeros((2, N), dtype=np.float32)
    haze_full[:, :n_h] = hzv
    air_full = np.zeros((2, N), dtype=np.float32)
    air_full[:, :air.shape[1]] = air
    air_full *= c_air
    ax = sc("arp", 0.7, c_arp)
    d = int(BEAT * 0.75 * SR)
    a0, a1 = ax[0].copy(), ax[1].copy()
    ax[1, d:] += 0.35 * a0[:-d]
    ax[0, 2 * d:] += 0.16 * a1[:-2 * d]
    padD = sc("padD", 0.55, c_padD, True)
    padB = sc("padB", 0.48, c_padB, True)
    padW = sc("padW", 0.50, c_padW, True)
    padF = sc("padF", 0.55)
    sub = sc("sub", 0.28, c_sub, True)
    tone = sc("tone", 0.22, c_tone)
    gr = sc("grain", 0.55, c_grain)
    pn, bl, ld = sc("pn", 0.95), sc("bell", 0.85), sc("lead", 0.85)
    bs = sc("bass", 0.60, c_bass, True)
    kkk, ht, fx = sc("kick", 0.80), sc("hat", 0.55), sc("fx", 0.60)
    dry = (padD + padB + padW + padF + sub + tone + pn + bl + ax + ld + bs + kkk + ht + fx + gr
           + haze_full * 0.9 + air_full * 0.35)
    send = ((padD + padB + padW + padF) * 0.30 + pn * 0.75 + bl * 0.9 + ax * 0.45 + ld * 0.5 + fx * 0.5
            + tone * 0.6 + gr * 0.5 + haze_full * 0.2)
    mix = dry + 0.5 * verb(send, 3.0, 0.9, 11)
    master(mix, [("noise haze", 0, 2), ("pitch + harmony", 2, 9), ("brightening", 9, 15), ("wide climax", 15, 20),
                 ("quiet pure end", 20, 24)], -17.0, -3.3, 9500, False, (56.0, 59.0))


if __name__ == "__main__":
    main()
