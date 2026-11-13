# Critique: Action C «DATA» (60 s, 68 shots)

Script angle: told from the customer's side. "Du har tallene; vi viser hva de betyr." Only real figures (1 300+, 1,5 % -> 1,7 %, 2,5 %, SMÅ/MELLOMSTORE/STORE, ti streker). The customer's own position is a red dot with no number.

## Round 1 (first contact sheets)
- Hook 4/5: DU. as a dot-word on night, first frame already alive. Legibility 4/5. Movement 3/5: most shots read as "word on top, dots below". Variation 3/5. Brand 4/5.
- Found: camera push on SE. HER./DU? zoomed about the red dot and exposed the sheet edge and pushed the word off frame; odometer strip overlapped the label; end card first frame blank white and text small; label type too small for a phone; 2025 sat on the histogram.
- Fixed: removed the off-centre push (1.09 centred), prerendered the odometer digits and cropped them with drawImage, end card text starts mid-rise at 140 px, labels 40 px, title moved up.

## Round 2
- Hook 4, legibility 4, movement 4, variation 4, brand 4.
- Moved six shots to word at the bottom with the data above, so composition alternates. TALL/TRENGER share one formation so dots do not cross the word. Comma descender no longer hits the label (label +112 px for numeral words). 2,5 % set smaller so the bar value does not collide.
- Strips checked: ledger row -> blue bar fill; 10 ÅR chart line edge pushing the frame; DU? red dot -> red ring -> red flat -> dots reform VI.; slam rows; dot morphs between formations. No [page] errors.

## Round 3
- Hook 4, legibility 5 (every word >= 230 px except the title line and numerals), movement 4 (log push on every shot, drift on dots, shake on hits, calm 4-beat breath on REVISOR./VURDERER.), variation 4, brand 5 (red only for signature, ring, three flats at bars 12/20/28).
- Open: columns (cols6/sizes) read a little noisy at dot scale; one-beat words (PÅ., ÉN., I.) are fast by design; title card line is 150 px because the real name is one long word.
- Perf (packed build, flushed with getImageData): worst 26 ms/frame, typically 1-15 ms. samples 3.
