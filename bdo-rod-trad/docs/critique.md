# Kritikk, Den røde tråden (poeng 1-5: hook / lesbarhet / bevegelse / variasjon / merkevare)

Runde 1: 4 / 3 / 4 / 4 / 4. Funn: tekst i oversiktsbildet ble rot (faste px-størrelser), tråden krysset overskriften i kapittel 6, plattformen kolliderte med bildetekst, kamera for høyt på avdelingene, blå tråder med krøller (uniform Catmull-Rom).
Tiltak: tekst skaleres med kamera og skjules under terskel, ny rute inn til kolonnene, plattform komprimert og manuelt kamera, sentripetal spline for alle linjer, kamera-ty justert.

Runde 2: 4 / 4 / 4 / 5 / 4. Funn: etikettene i analysekapittelet ble krysset av tråden, oversiktsetiketter overlappet, siste kort hadde to tekster for tett, utjevningskant på tegneark kostet 30 ms/frame.
Tiltak: flaggetiketter flyttet til bånd over diagrammet med linje, etiketter forskjøvet i to rader, MONO-linje sentrert under logo, kant lagd med klipp + gradient (under 5 ms).

Runde 3: 5 / 4 / 4 / 4 / 4. Restpunkter: høyre halvdel er tom et par sekunder ved overgangen diagram til samtale; offisiell logo mangler (plassholder). Bygg OK.
