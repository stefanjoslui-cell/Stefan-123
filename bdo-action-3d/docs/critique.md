# Critique log: BDO Automotive Action (flat 2-D v2, 60 s, 128 bpm)

Scores 1-5: hook / readability / motion / variety / brand.

## Round 1 (contact sheet, every beat)  4 / 3 / 4 / 4 / 4
Three worst:
1. 1,7 % shot: the mono label ran into the bars and was cut off. Fixed: bars moved right (x 1290+), label 34 px, left aligned under the number.
2. 'HER.' tag 'Garantikrav uten arbeidsordre' was clipped on the left by its own mask. Fixed: mask widened.
3. Word-to-word 'slide' showed the old and new letters on top of each other at the cut frame ('DDANTEA'). Fixed: old letters leave in 0.2 s (stagger 20 ms), new word starts 0.12 s later.
Also: dot matrix too faint on night (raised alpha), end-card tagline too small (132 -> 150 px).

## Round 2 (performance + strips of every transition type)  4 / 4 / 4 / 5 / 4
Three worst:
1. Frame cost: full-frame gradient sheet, 1 400 grain rects, 1 300 dot arcs and a vector site drawn per frame. Fixed: flat bg fill, 260 grain rects, batched fillRect dots, cached field layers (links+points, one drawImage), own cached site raster (siteC), transitions composite B through one clipped blit of an off-screen canvas. Average 6.7 ms/frame; steady frames well under 30 ms.
2. Push-through from the wheel: the iris was smaller than the wheel rim, so the match cut did not match. Fixed (r0 440).
3. TUSEN whip image redrew the big site raster up to 3x at 2.6x scale (52 ms). Fixed: cached raster at 12 px/m, low-quality scaling, scale 26 (22 ms).

## Round 3 (re-render, strips: iris, stop-dot, push-through, whip, diag35, line wipe, O-counter push, odometer)  5 / 4 / 4 / 5 / 4
Three remaining worst, status:
1. A few transition frames (HENGER line wipe, DEG->stack) still cost 50-60 ms in this software canvas (worst single frame ~64 ms); steady frames are < 25 ms. Not fully fixed.
2. Stack lines on the closing shot are 168 px (16-character words cannot be 240 px inside 96 px margins); title card lines are 290 px.
3. Logo files are not in brand/ yet, so the cornerstone carries the typographic stand-in 'BDO' (kit fallback); no logo is drawn by the film itself.
Verified: every transition strip shows both shots on screen at the cut frame, no blank frames, no [page] errors, end card holds still from beat ~119.5 (about 4 s), nothing in its lower-right third.
