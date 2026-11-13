# Action-filmen, 60 s – felles oppdrag for alle eksempler (flat 2D, INGEN 3D)

Utgangspunkt: den originale klippen `bdo-action/film.js` (ett stort flatt ord per klipp, harde klipp på takten, vekslende
flater natt / lys / BDO Blue / sjelden rød, enkle streektegnede bilder). Kunden likte den, men vil ha bedre overganger,
effekter og innhold. Ikke bruk `lib/k3d.js` (ingen perspektiv, ingen 3D-bokser, ingen ekstrudert tekst).

Les først: `BRIEF.md` (kontrakt, BDOs fargeprofil, ekte tall, verifisering), `bdo-action-3d/docs/shotlist.md`
(SEKSJONER, ORD og taktinndeling gjelder – ignorer kolonnen «3-D image»), `lib/kit.js`, `lib/motion.js`, `scenes/world.js`.

Fast struktur (128 bpm, 128 beats, 32 takter, 60 s – musikken følger denne):
1–4 cold open · 5–8 virksomheten · 9–12 data/teknologi (rød full flate takt 12) · 13–20 BRANSJEKUNNSKAP (rød takt 20)
· 21–24 effektiv revisjon (pust i takt 23–24) · 25–28 innsikt for kunden (rød takt 28) · 29–32 slutt.

Innhold som MÅ fram:
1. BDO kjenner bransjen: «BDO BILBRANSJEANALYSEN 2025» som stort tittelkort, og KUN disse ekte punktene:
   1 300+ bilbransjebedrifter analysert · driftsmargin merkeforhandlere 1,5 % (2023) → 1,7 % (2024) («ETTERLENGTET ØKNING»)
   · driftsmargin bruktbilforhandlere 2,5 % (2024) · forhandlere sammenlignet etter størrelse (SMÅ / MELLOMSTORE / STORE)
   · omsetning og driftsmargin i et 10-års perspektiv (ti streker, ingen verdier). Ingen andre tall som statistikk.
2. Teknologi behandler dataene og gir effektiv, målrettet revisjon. Teknologien viser hvor vi bør se; revisor vurderer.
   Aldri «AI bestemmer», «feil», «avvik funnet».
3. Innsikt for kunden i egne data (DINE. DATA. DIN. INNSIKT. … MER. VERDI. FOR. DEG.).

Sluttkort (takt 30–32): hvitt kort, «Vi ser detaljene.» / «Vi forstår helheten.» til venstre, MONO «BDO AUTOMOTIVE»,
`K.cornerstone(ctx, p)` nede til høyre tegnet SIST (rød, 35°, heldekkende, ingenting over eller i nærheten), stille ≥ 3 s.
`await K.loadLogos()` i init(). Ikke tegn BDO-logo selv.

Håndverk: bildet PÅ klippet viser alltid det nye bildet (ingen tomme bilder, ingen fading, ingen glød, ingen gradienter).
Overganger skal være motiverte og varierte (maks ca. halvparten rene klipp). Kamera aldri stille (log-push, shake på treff,
parallakse). Hvert ord lesbart i sin takt, aldri klippet i hvile (96 px marg). Ren funksjon av tid, `samples` ≤ 4,
< 30 ms per bilde i pakket bygg (mål med
`node <film>/build.mjs && node /tmp/claude-0/-home-user-Stefan-123/0121f7c5-2d95-598a-a73c-a124db5d916c/scratchpad/fperf.mjs /home/user/Stefan-123/<film>/dist/<output>.html 1 5 10 15 20 25 30 35 40 45 50 55 58`).
HITS for hvert klipp/aksent (70–90). Verifiser: kontaktark `--contact --every 1 --noblur` (se ALLE sider + phone-sidene),
`--strip a:b:8 --name x` over minst seks ulike overganger, fiks alle `[page]`-feil, 3 kritikkrunder i `docs/critique.md`.
Rør bare din egen filmmappe; ikke `lib/`, ikke `audio/`.
