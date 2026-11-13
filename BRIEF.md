# BDO Automotive – felles brief for alle versjoner

Promofilm for BDO (revisjon/rådgivning) om revisjon av bilforhandlere. Kjernebudskap:
**Med bransjekunnskap, data og teknologi skaper BDO en mer effektiv og målrettet revisjon – og innsikt som gir verdi for kunden.**
Historiebuen: én transaksjon → én bil → hel bilpark → hele virksomheten (salg, lager, verksted, deler, garanti, service)
→ tusenvis av datapunkter → BDOs dataplattform → analyse på tvers → områder som krever nærmere undersøkelse
→ revisors faglige vurdering → verdi for revisor og for kunden → BDO-logo.
Tone: teknologien hjelper revisor å se mer og prioritere riktig; den tar ALDRI beslutninger. Ikke skriv «feil», «avvik funnet», «AI».

## Teknisk kontrakt (alle filmer)
- Hver film er en mappe `<film>/` med `film.js` (ES-modul) som kaller `M.film({...})` fra `lib/motion.js`. Se `bdo-automotive/film.js` for mønsteret.
  `draw(ctx, u, t)` maler HELE 1920×1080-bildet for tid t (sekunder). Ren funksjon av t: ingen tilstand mellom kall, ingen Math.random
  (bruk `K.mulberry32`/`K.hash01`), ingen timere. Forhåndsberegn layout én gang på modulnivå.
- Bruk `lib/kit.js` (les den!): farger `K.C`, skrifter `K.DISPLAY` (Hanken Grotesk, vekt 100–900) og `K.MONO` (IBM Plex Mono), `K.sheet`, `K.grid`,
  `K.field` (2400-punkts datafelt, punkt 0 er den røde signaturen), `K.line` (tekst som reiser seg fra en maske), `K.kicker`, `K.redDot`,
  `K.drawOn`, `K.bez`, `K.win`, `K.ss`, `K.E`, `K.kf`, `K.spring`, `K.grain` (kun i `post`).
- `scenes/world.js` (kopi) eksporterer `site(ctx, {k, cx, cy, s, alpha})` – forhandleranlegget som arkitekttegning i plan
  (verdensmeter, x −85..85, y −50..50, bygning øverst y −46..−24 med seks avdelinger fra venstre: Salg, Lager, Verksted, Deler, Garanti, Service;
  bilpark under; s≈10.5 viser hele anlegget, s≈60 noen få biler) og `toScreen(cx,cy,s,x,y)`. Du kan også gjenbruke/tilpasse tegnefunksjonene
  (bil i oppriss `elevation`, bil i plan `carPlan`) ved å eksportere dem fra DIN kopi av world.js. Ikke rør `bdo-automotive/`.
- **Logo:** sluttkortet SKAL bruke `await K.loadLogo()` (i `init()`) og `K.drawLogo(ctx, cx, cy, h, {alpha, wipe})` – den offisielle logoen
  legges senere i `brand/`; til da tegnes en enkel typografisk plassholder. Ikke tegn din egen BDO-logo. Gi logoen luft (hvit/lys flate,
  ingen effekter på selve logoen, ingen rotasjon/forvrengning).
- `HITS = [[beat, 'label', sfx, opts?], ...]` i beats (`t / period`), én per visuell aksent. sfx: tick, click, pop, whoosh, swell, riser, bell, blip, sub, thud, impact, crash.
- `samples` (bevegelsesuskarphet) ≤ 6. Hver frame må tegnes på < 30 ms i Chromium (sjekk!).

## Visuell identitet – BDOs offisielle fargeprofil (streng)
BDO Red #E81A3B (Pantone 185) · BDO Blue #22409A (Pantone 287) · BDO Charcoal #333333 (Cool Grey 11) · BDO Slate #5B6E7F (174+10)
· BDO Burgundy #98002E (208) · BDO Pale Charcoal #E7E7E7 · hvit. Arbeidstoner: lyst ark #F7F7F7, mørkt ark #17191C.
BDO Blue er dominerende digital farge; rødt KUN for signaturpunkt/-linje og markeringer; Charcoal for tekst på lyst; Slate for sekundære linjer/etiketter;
Burgundy sjelden. I kit.js: K.C.red/blue/charcoal/slate/burgundy/pale (+ alias cloud→pale, graphite→charcoal, light/teal→slate).
Skandinavisk, rent, presist: tynne linjer, luft. Ingen neon, ingen glød på UI, ingen fargerike gradienter, ingen ikoner,
ingen PowerPoint-bokser, ingen tekst som bare toner inn. INGEN 3D-bilmodeller.

## Ekte tall (kun disse – ikke finn på andre)
BDO Bilbransjeanalysen 2025: driftsmargin merkeforhandlere 1,5 % (2023) → 1,7 % (2024); bruktbilforhandlere 2,5 % (2024);
over 1 300 bilbransjebedrifter analysert. Fiktive regnskapslinjer (bilagsnr., beløp) er greit som illustrasjon.

## Kvalitetskrav og verifisering (obligatorisk)
Render stillbilder fra prosjektroten:
`node /root/.claude/plugins/synced/7399c4b9-c02f-4b5c-97b1-35fff68278af_6517a19b-b444-43eb-950e-b55efe9c8ecd/d92346f7-5140-4ddb-b284-a3b42995e5b0/skills/motion-reel/scripts/render.mjs --film <film> --at 1,5,9,... --noblur`
(→ `out/<film>/stills/`) og kontaktark `--contact --every 4 --noblur` (→ `out/<film>/contact_p1.png`). SE på dem (Read) og iterer minst 3 runder
som en streng art director: ingen tekst som overlapper/klippes (96 px sikker marg), alt lesbart på telefon (`phone_p1.png`), ingen tomme/statiske bilder
(noe nytt hvert 2–4 s, langsom kontinuerlig kamerabevegelse ellers), streng palett, første bilde er allerede interessant (ingen fade fra svart).
Fiks alle `[page]`-feil. Skriv `<film>/docs/critique.md` med poeng (hook, lesbarhet, bevegelse, variasjon, merkevare) per runde.
Ikke render MP4 og ikke rør andre mapper enn din egen film.
