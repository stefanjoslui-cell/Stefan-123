# /// script
# requires-python = ">=3.10"
# dependencies = ["librosa>=0.10", "numpy", "scipy", "soundfile"]
# ///
"""BDO film score: 195 s, 96 bpm, 4/4 (312 beats = 78 bars), D major / B minor.

    cd <repo> && MOTION_REEL_SKILL=<skill> uv run bdo-automotive/audio/score.py

Writes <film>/audio/music.wav and <film>/beats.json. The grid is synthesized exactly
(beat = 0.625 s, beat 0 at t = 0), so beats.json stores the exact values.

Sections (bars, 1 bar = 2.5 s):
  0-8   one data point      blip on every beat, sub pad swells in, felt-piano motif appears in pieces
  8-20  the business        8th pluck arpeggio, soft kick on 1 and 3, warm bass, lift at bar 16
  20-32 data platform       harmony opens (D-A-Bm-G), pad filter opens over 4 bars, arp an octave up, hats
  32-48 analysis            four-on-the-floor, 16th pluck, moving bass; breakdown 44-48 (pad + pulse)
  48-62 the auditor         felt-piano chords, kick from bar 56, slow build, riser into 62
  62-74 value               full arrangement, 4-bar motif x3, widest pad
  74-78 ending              final chord on 74, motif once more in bells/keys, ring out
"""
import json
import os
import sys
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
FILM = HERE.parent
BPM = 96
BEAT = 60 / BPM          # 0.625 s, exact in binary
BAR = 4 * BEAT           # 2.5 s
DUR = 195.0
N = int(SR * DUR)
NBARS = 78
TARGET_LUFS = -17.0
CEIL_DB = -3.3
seed(20261113)
prng = np.random.default_rng(11)

# ---------------------------------------------------------------- helpers

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


# ---------------------------------------------------------------- instruments

def pad(notes, length, cutoff, attack=0.7, release=1.2, voices=4, detune=0.10, wide=False):
    """Detuned-saw pad, stereo; static lowpass, so each chord is built once."""
    notes, n = list(notes), int(length * SR)
    pr = np.random.default_rng(abs(hash(tuple(notes))) % 2 ** 31)
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
    """Felt-piano-ish: inharmonic decaying partials, dull top, soft hammer thump."""
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


def blip(f=440.0):
    n = int(0.4 * SR)
    t = tvec(n)
    x = np.sin(2 * np.pi * f * t) * np.exp(-t / 0.11) + 0.3 * np.sin(2 * np.pi * 2 * f * t) * np.exp(-t / 0.05)
    return x * (1 - np.exp(-t / 0.004))


def hat_soft():
    return filt(hat(False), "lowpass", 9500, 2)


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


# ---------------------------------------------------------------- harmony

CH = {
    "Bm": dict(sub="B1", bass="B2", oct="B3", third="D3", arp=["B3", "D4", "F#4", "A4"],
               pad=["B2", "F#3", "D4", "A4"], padO=["B2", "F#3", "A3", "C#4", "D4"],
               key=["B2", "F#3", "A3", "D4", "F#4"]),
    "G": dict(sub="G1", bass="G2", oct="G3", third="B2", arp=["G3", "B3", "D4", "F#4"],
              pad=["G2", "D3", "B3", "F#4"], padO=["G2", "D3", "B3", "F#4", "A4"],
              key=["G2", "D3", "F#3", "B3", "D4"]),
    "D": dict(sub="D2", bass="D2", oct="D3", third="F#2", arp=["D4", "F#4", "A4", "D5"],
              pad=["D3", "A3", "E4", "F#4"], padO=["D3", "A3", "C#4", "E4", "F#4"],
              key=["D3", "A3", "E4", "F#4", "A4"]),
    "A": dict(sub="A1", bass="A2", oct="A3", third="C#3", arp=["A3", "C#4", "E4", "A4"],
              pad=["A2", "E3", "B3", "C#4"], padO=["A2", "E3", "B3", "C#4", "E4"],
              key=["A2", "E3", "B3", "C#4", "E4"]),
}
PROG = ([(0, 2, c) for c in []]
        + [(2 * i, 2, c) for i, c in enumerate(["Bm", "G", "D", "A", "Bm", "G", "D", "A", "Bm", "G"])]
        + [(20 + 2 * i, 2, c) for i, c in enumerate(["D", "A", "Bm", "G", "D", "A"])]
        + [(32 + 2 * i, 2, c) for i, c in enumerate(["Bm", "G", "D", "A", "Bm", "G", "D", "A"])]
        + [(48 + 2 * i, 2, c) for i, c in enumerate(["Bm", "G", "D", "A", "Bm", "G", "A"])]
        + [(62 + i, 1, c) for i, c in enumerate(["D", "A", "Bm", "G"] * 3)]
        + [(74, 4, "D")])


def chord_at(bar):
    for s, d, c in PROG:
        if s <= bar < s + d:
            return c


# the 4-bar motif (bar offset, beat, note, beats): D / A / Bm / G, one chord per bar in the climax
MOTIF = [(0, 0, "A4", 1), (0, 1, "D5", 1), (0, 2, "F#5", 2),
         (1, 0, "E5", 1.5), (1, 1.5, "F#5", 0.5), (1, 2, "E5", 2),
         (2, 0, "D5", 1), (2, 1, "F#5", 1), (2, 2, "A5", 1.5), (2, 3.5, "F#5", 0.5),
         (3, 0, "G5", 1), (3, 1, "F#5", 1), (3, 2, "D5", 2)]
ENDM = [(75, 0, "A4", 1), (75, 1, "D5", 1), (75, 2, "F#5", 2),
        (76, 0, "E5", 1.5), (76, 1.5, "F#5", 0.5), (76, 2, "E5", 2), (77, 0, "D5", 3.5)]

PAT8 = [0, 1, 2, 3, 2, 1, 2, 3]
PAT16 = [0, 2, 1, 3, 2, 1, 3, 2, 0, 2, 1, 3, 2, 3, 1, 2]
BP_SOFT = [(0, "r", 2.4, .8), (2.5, "r", 1.2, .5), (3.5, "o", .4, .3)]
BP_8 = [(0, "r", .45, .8), (.5, "r", .45, .5), (1, "r", .45, .6), (1.5, "o", .45, .5),
        (2, "r", .45, .7), (2.5, "r", .45, .5), (3, "r", .45, .6), (3.5, "t", .45, .5)]
BP_CL = [(0, "r", 1.4, .85), (1.5, "o", .4, .5), (2, "r", 1.4, .75), (3.5, "t", .4, .45)]

# ---------------------------------------------------------------- score


def main():
    B = {k: mkbus() for k in ("padD", "padB", "padW", "sub", "blip", "pn", "bell", "arp", "lead", "bass", "kick",
                              "hat", "fx")}
    kick_t = []

    def bar_t(bar, beat=0.0):
        return (bar * 4 + beat) * BEAT

    # ---- pads + sub on the chord grid
    for s, d, c in PROG:
        ch = CH[c]
        L = d * BAR + 1.4 if d < 4 else 9.8
        rel = 1.2 if d < 4 else 3.0
        at = bar_t(s) - 0.15
        B["padD"].add(mem(pad, tuple(ch["pad"]), L, 700, 0.7, rel), 0, 1.0, at=at)
        B["sub"].add(mem(subpad, ch["sub"], L), 0, 1.0, at=at)
        if s >= 20:
            B["padB"].add(mem(pad, tuple(ch["padO"]), L, 2800, 0.6, rel), 0, 1.0, at=at)
        if s >= 62:
            wn = tuple(ch["padO"] + [tr(ch["padO"][2], 12), tr(ch["padO"][3], 12)])
            B["padW"].add(mem(pad, wn, L, 3400, 0.5, rel, 6, 0.2, True), 0, 1.0, at=at)

    # ---- helpers that place events
    def blips(b0, b1):
        for b in range(b0, b1):
            B["blip"].add(mem(blip, 440.0), b, 1.0 if b % 4 == 0 else 0.75)

    def kicks(bars, beats, g):
        for bar in bars:
            for bt in beats:
                B["kick"].add(mem(kick, True), bar * 4 + bt, g)
                kick_t.append(bar_t(bar, bt))

    def bass(bar, pat, g=1.0, cutoff=700):
        ch = CH[chord_at(bar)]
        for beat, w, dur, gg in pat:
            note = {"r": ch["bass"], "o": ch["oct"], "t": ch["third"]}[w]
            B["bass"].add(mem(bass_note, note, round(dur * BEAT, 3), cutoff), bar * 4 + beat, g * gg)

    def arp(bar, kind, octv, g, cutoff):
        notes = [tr(n, octv) for n in CH[chord_at(bar)]["arp"]]
        steps, pat = (8, PAT8) if kind == 8 else (16, PAT16)
        for s in range(steps):
            idx = pat[s]
            if kind == 8:
                w = 1.0 if s % 2 == 0 else 0.65
            else:
                w = 1.0 if s % 4 == 0 else 0.7 if s % 2 == 0 else 0.5
            w *= 1 + 0.07 * prng.standard_normal()
            B["arp"].add(mem(pluck, (notes[idx],), 0.3, cutoff), bar * 4 + s * (4 / steps), g * w,
                         pan=[-.3, -.1, .1, .3][idx])

    def hats(bars, mode, g):
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
                B["hat"].add(mem(hat_soft), bar * 4 + p, g * w, pan=0.25 if p % 1 else -0.2)

    def roll(notes, bar, beat, g, spread=0.045, length=3.5):
        for i, nt in enumerate(notes):
            B["pn"].add(mem(piano, nt, length), 0, g * (1 - 0.05 * i), at=bar_t(bar, beat) + i * spread)

    def melody(bar0, offs, bg, pg, lg=0.0, up=0.0):
        for bo, beat, note, dur in MOTIF:
            if bo not in offs:
                continue
            pos = (bar0 + bo) * 4 + beat
            if bg:
                B["bell"].add(mem(bell, note, round(min(max(dur * BEAT * 1.8, 1.4), 3.5), 2)), pos, bg, pan=0.1)
            if pg:
                B["pn"].add(mem(piano, note, 3.5), pos, pg, pan=-0.1)
            if lg:
                B["lead"].add(mem(lead, note, round(dur * BEAT * 1.05, 3)), pos, lg, pan=0.05)
            if up:
                B["bell"].add(mem(bell, tr(note, 12), 2.0), pos, up, pan=0.25)

    # ---- 1. one data point (bars 0-8)
    blips(0, 96)
    blips(176, 192)
    melody(4, (0,), 0, 0.55)           # motif arrives in pieces
    melody(5, (1,), 0, 0.55)           # bar 6

    # ---- 2. the business (bars 8-20)
    for bar in range(8, 20):
        arp(bar, 8, 0, 0.55, 2000)
        bass(bar, BP_SOFT, 0.9 if bar < 16 else 1.0)
    kicks(range(8, 20), (0, 2), 0.55)
    for bar in range(8, 20, 2):
        a = CH[chord_at(bar)]["arp"]
        B["pn"].add(mem(piano, a[3], 3.5), bar * 4, 0.40)
        B["pn"].add(mem(piano, a[2], 3.5), bar * 4 + 1.5, 0.32, pan=-0.1)
        B["pn"].add(mem(piano, tr(a[1], 12), 3.5), bar * 4 + 6, 0.28, pan=0.1)
    for n_ in [tr(x, 12) for x in CH["Bm"]["arp"][:3]]:               # lift at bar 16
        B["bell"].add(mem(bell, n_, 3.0), 16 * 4, 0.35, pan=0.2)
    B["fx"].add(filt(riser(BAR), "lowpass", 5000, 2), 15 * 4, 0.05)
    hats(range(16, 20), "off8", 0.06)

    # ---- 3. dawn / data platform (bars 20-32)
    for bar in range(20, 32):
        arp(bar, 8, 12, 0.5, 2600)
        bass(bar, BP_SOFT)
    kicks(range(20, 32), (0, 2), 0.55)
    hats(range(20, 32), "off8", 0.12)
    for bar in range(20, 32, 2):
        a = CH[chord_at(bar)]["arp"]
        B["bell"].add(mem(bell, tr(a[3], 12), 3.0), bar * 4, 0.20, pan=0.2)
        B["pn"].add(mem(piano, a[2], 3.5), bar * 4 + 2.5, 0.3, pan=-0.1)
    melody(28, (0, 1), 0.30, 0.40)
    melody(28, (2, 3), 0.30, 0.40)
    B["fx"].add(filt(riser(BAR), "lowpass", 5000, 2), 19 * 4, 0.05)

    # ---- 4. analysis (bars 32-48)
    for bar in range(32, 44):
        arp(bar, 16, 0, 0.42, 2700)
        bass(bar, BP_8, 1.0, 650)
    kicks(range(32, 44), (0, 1, 2, 3), 0.75)
    hats(range(32, 44), "16", 0.14)
    for bar in range(32, 44):
        a = CH[chord_at(bar)]["arp"]
        if bar % 2 == 0:
            B["bell"].add(mem(bell, tr(a[3], 12), 3.0), bar * 4 + 2.5, 0.16, pan=0.3)
    for bar in (44, 46):                                              # breakdown: pad + pulse, one soft chord
        roll(CH[chord_at(bar)]["key"][1:], bar, 0, 0.30, 0.06, 4.0)
    B["fx"].add(filt(riser(BAR), "lowpass", 5000, 2), 47 * 4, 0.04)

    # ---- 5. the auditor's judgement (bars 48-62)
    for bar in range(48, 62):
        k = CH[chord_at(bar)]["key"]
        roll(k, bar, 0, 0.24 + 0.008 * (bar - 48))
        roll(k[2:], bar, 2, 0.14, 0.04)
        B["pn"].add(mem(piano, tr(k[4], 12), 3.5), bar * 4 + 3.5, 0.11, pan=0.15)
    for bar in range(52, 62):
        bass(bar, BP_SOFT, 0.45 + 0.04 * (bar - 52))
    kicks(range(56, 60), (0, 2), 0.38)
    kicks(range(60, 62), (0, 1, 2, 3), 0.48)
    hats(range(58, 62), "off8", 0.08)
    for bar in range(58, 62):
        arp(bar, 8, 0, 0.30, 2200)
    melody(58, (0, 1), 0.0, 0.38)                                     # foreshadow
    melody(58, (2, 3), 0.0, 0.38)
    B["fx"].add(filt(riser(BAR), "lowpass", 6000, 2), 61 * 4, 0.10)

    # ---- 6. value for the customer (bars 62-74)
    B["fx"].add(mem(sub_drop, 1.4, 70, 34), 62 * 4, 0.45)
    for bar in range(62, 74):
        arp(bar, 16, 12, 0.40, 3000)
        bass(bar, BP_CL, 1.0, 800)
        roll(CH[chord_at(bar)]["key"], bar, 0, 0.30)
    kicks(range(62, 74), (0, 1, 2, 3), 0.62)
    hats(range(62, 74), "16", 0.11)
    melody(62, (0, 1, 2, 3), 0.42, 0.50)
    melody(66, (0, 1, 2, 3), 0.42, 0.45, 0.30)
    melody(70, (0, 1, 2, 3), 0.42, 0.45, 0.40, 0.18)
    for bar in (62, 66, 70):
        for n_ in [tr(x, 12) for x in CH[chord_at(bar)]["arp"][:3]]:
            B["bell"].add(mem(bell, n_, 3.0), bar * 4, 0.12, pan=0.2)

    # ---- 7. ending (bars 74-78)
    B["kick"].add(mem(kick, True), 74 * 4, 0.8)
    kick_t.append(bar_t(74))
    B["fx"].add(mem(sub_drop, 1.8, 70, 34), 74 * 4, 0.55)
    B["bass"].add(mem(bass_note, "D2", 6.0, 500), 74 * 4, 0.7)
    roll(["D3", "A3", "E4", "F#4", "A4", "D5"], 74, 0, 0.55, 0.05, 7.0)
    for n_ in ("D5", "F#5", "A5", "D6"):
        B["bell"].add(mem(bell, n_, 3.5), 74 * 4, 0.28, pan=0.2)
    for bar, beat, note, dur in ENDM:
        pos = bar * 4 + beat
        B["bell"].add(mem(bell, note, 3.5 if bar == 77 else round(min(max(dur * BEAT * 1.8, 1.4), 3.5), 2)),
                      pos, 0.42, pan=0.1)
        B["pn"].add(mem(piano, note, 4.5 if bar == 77 else 3.5), pos, 0.42, pan=-0.1)
    B["bell"].add(mem(bell, "D6", 3.5), 77 * 4, 0.15, pan=0.3)

    # ---- bar-based gain curves
    tb = np.arange(B["padD"].x.shape[1]) / SR / BAR

    def curve(*pts):
        return np.interp(tb, [p[0] for p in pts], [p[1] for p in pts]).astype(np.float32)

    c_padD = curve((0, .10), (4, .30), (8, .70), (14, .75), (16, .95), (20, .95), (24, 0), (47, 0), (48.5, .5),
                   (56, .6), (62, .35), (74, .45), (78, .45))
    c_padB = curve((20, 0), (24, .85), (32, .90), (44, 1.0), (47.5, .8), (48.5, 0), (62, 0), (62.5, .5), (78, .5))
    c_padW = curve((61, 0), (62, 1.0), (74, 1.1), (78, 1.0))
    c_sub = curve((0, .05), (8, .6), (20, .7), (44, .5), (48, .35), (62, .8), (74, .9), (78, .5))
    c_blip = curve((0, .10), (4, .25), (8, .42), (19, .30), (24, 0), (43.9, 0), (44, .40), (47.9, .45), (48, 0))
    c_arp = curve((8, .35), (16, .35), (16.05, .5), (32, .85), (48, .85), (58, .5), (62, .95), (74, 1.0))
    c_bass = curve((8, .75), (20, .8), (32, .9), (52, .7), (62, 1.0), (74, 1.0))
    duck = duck_curve(kick_t, B["padD"].x.shape[1])

    def sc(name, g=1.0, cv=None, dk=False):
        x = B[name].x[:, :N] * np.float32(g)
        if cv is not None:
            x = x * cv[:N]
        if dk:
            x = x * duck[:N]
        return x

    # arp with dotted-8th ping-pong echo
    ax = sc("arp", 0.7, c_arp)
    d = int(BEAT * 0.75 * SR)
    a0 = ax[0].copy()
    a1 = ax[1].copy()
    ax[1, d:] += 0.35 * a0[:-d]
    ax[0, 2 * d:] += 0.16 * a1[:-2 * d]
    ax[0, d:] += 0.10 * a1[:-d]

    padD = sc("padD", 0.55, c_padD, True)
    padB = sc("padB", 0.48, c_padB, True)
    padW = sc("padW", 0.50, c_padW, True)
    sub = sc("sub", 0.28, c_sub, True)
    blp = sc("blip", 0.55, c_blip)
    pn = sc("pn", 0.95)
    bl = sc("bell", 0.85)
    ld = sc("lead", 0.85)
    bs = sc("bass", 0.60, c_bass, True)
    kk = sc("kick", 0.80)
    ht = sc("hat", 0.55)
    fx = sc("fx", 0.60)

    dry = padD + padB + padW + sub + blp + pn + bl + ax + ld + bs + kk + ht + fx
    send = (padD + padB + padW) * 0.30 + pn * 0.75 + bl * 0.9 + ax * 0.45 + ld * 0.5 + blp * 0.7 + fx * 0.5
    ir = reverb_ir(2.8, 0.85, 5)
    wet = np.stack([signal.fftconvolve(send[c], ir[c])[:N] for c in range(2)]).astype(np.float32)
    wet = np.stack([filt(filt(ch, "lowpass", 5500, 2), "highpass", 180, 1) for ch in wet])
    mix = dry + 0.45 * wet
    del dry, send, wet, padD, padB, padW, sub, blp, pn, bl, ax, ld, bs, kk, ht, fx
    grid_src = (B["kick"].x[:, :N] + B["hat"].x[:, :N] * 0.5 + B["blip"].x[:, :N] * 0.5).mean(axis=0)
    del B

    # ---- master: tame the top, ring out
    mix = np.stack([filt(filt(ch, "highpass", 30, 2), "lowpass", 9500, 2) for ch in mix])
    t = np.arange(N) / SR
    mix *= np.where(t < 188.5, 1.0, np.cos(np.clip((t - 188.5) / 6.0, 0, 1) * np.pi / 2) ** 2).astype(np.float32)
    mix[:, int(194.5 * SR):] = 0

    path = HERE / "music.wav"
    g = 0.5 / np.abs(mix).max()
    for _ in range(5):
        y = limit(mix * g, CEIL_DB)
        sf.write(path, y.T, SR, subtype="PCM_24")
        lufs, pk = ebur128(path)
        print(f"gain {g:.3f}: {lufs:.2f} LUFS, peak {pk:.2f} dBFS")
        if abs(TARGET_LUFS - lufs) < 0.3:
            break
        g *= 10 ** ((TARGET_LUFS - lufs) / 20)

    # ---- section report
    print("section RMS (dBFS), spectral centroid, share of energy above 6 kHz:")
    names = [("one data point", 0, 8), ("the business", 8, 20), ("data platform", 20, 32), ("analysis", 32, 44),
             ("breakdown", 44, 48), ("auditor", 48, 62), ("value", 62, 74), ("ending", 74, 78)]
    mono = y.mean(axis=0)
    for nm, a, b in names:
        seg = mono[int(a * BAR * SR):int(b * BAR * SR)]
        f, p = signal.welch(seg, SR, nperseg=8192)
        print(f"  {nm:16s} bars {a:2d}-{b:2d}  {20 * np.log10(np.sqrt(np.mean(seg ** 2)) + 1e-12):6.1f} dB  "
              f"centroid {np.sum(f * p) / np.sum(p):6.0f} Hz  >6k {100 * p[f > 6000].sum() / p.sum():.2f}%")
    tail = mono[int(194.5 * SR):]
    print(f"  tail after 194.5 s: peak {np.abs(tail).max():.6f}; samples {y.shape[1]} = {y.shape[1] / SR:.3f} s")

    # ---- beats.json: the grid is synthesized exactly
    try:
        m = measure_beats(grid_src, SR, BPM, n_beats=313)
        print(f"measure_beats sanity: {m['bpm']} bpm, offset {m['offset'] * 1000:.2f} ms, "
              f"fit rms {m['fit_residual_ms']} ms, {m['fit_inliers']} inliers")
    except Exception as e:  # sanity check only
        print("measure_beats skipped:", e)
    beats = [round(i * BEAT, 6) for i in range(NBARS * 4 + 1)]
    secs = [("one data point", 0, 8), ("the business", 8, 20), ("dawn / data platform", 20, 32),
            ("analysis", 32, 48), ("the auditor's judgement", 48, 62), ("value for the customer", 62, 74),
            ("ending", 74, 78)]
    out = {"bpm": 96, "period": 0.625, "offset": 0.0, "beats": beats, "downbeats": beats[::4],
           "sections": [{"name": n, "bar0": a, "bar1": b, "t0": a * BAR, "t1": b * BAR} for n, a, b in secs],
           "duration": DUR, "key": "D major / B minor",
           "source": "bdo-automotive/audio/score.py (synthesized on an exact 96 bpm grid)"}
    (FILM / "beats.json").write_text(json.dumps(out, indent=1))
    print(f"beats.json: {len(beats)} beats, last {beats[-1]} s")


if __name__ == "__main__":
    main()
