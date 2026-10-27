# MVA-oppslagsverk for Bauda

**Åpne `MVA-oppslagsverk.html`.** Filen er selvstendig, virker uten nett og kan sendes som vedlegg. Åpne den i en vanlig nettleser (Chrome, Edge). Forhåndsvisning i OneDrive eller Outlook kjører ikke skript, så last den ned først.

## Slik bruker du det

1. Svar på de tre spørsmålene (bruk, kjøretøy, varighet). Tallet ved hvert svar viser hvor mange scenarioer som blir igjen.
2. «Utleie» og «Andre forhold» over brikkene filtrerer på gruppe (scenario 1–5 og 6–8).
3. Scenarioer som ikke passer, glir ut. «Vis dem» viser dem nedtonet, med årsak.
4. Klikk et scenario for å se hele konklusjonen. Når bare ett gjenstår, åpnes det av seg selv.
5. **Treet i menyen** vokser og beskjæres etter valgene dine. Hold musen over et svar for å se hva som blir igjen før du velger. Trykk på treet eller «Nullstill valg» for å fjerne alle filtre.
6. «Yrkesmessig utleie» står som egen seksjon øverst, med de tre punktene fra PP-en.

## PDF-eksport

- **Eksporter alt til PDF:** forside, innholdsfortegnelse og alle 8 scenarioer med yrkesmessig utleie (6 sider).
- **Eksporter valgte til PDF:** bare scenarioene som passer og er avkrysset. Kompakt topp på første side, uten forside. Scenarioene flyter etter hverandre.
- **Kun denne til PDF:** eksporterer ett scenario, eller bare «Yrkesmessig utleie» (1 side).
- Knappene finnes både i toppbildet og i menyen. Holder du musen over dem, lyser de tre yrkesmessig-boksene i menyen svakt hvis den delen tas med. Velg «Lagre som PDF» som skriver i utskriftsdialogen.
- Eksempler: `eksempel-pdf/`.

## Innhold

- `innhold.json` er eneste kilde til tekst. Teksten er gjengitt ordrett fra «Bistand bauda (1) (003).pptx» (lysbilde 2, 3 og 5). Fotnotene («*» og «**») står i boksene til scenario 1 og 7. Interne notater («– Richard» og «Hør med Richard…») vises ikke og legges ikke inn i HTML-filen.
- Koblingen mellom spørsmål og scenario er utledet av ordlyden i overskriftene (`ATTR` i `kilde/app.del`). Et scenario uten vilkår på ett spørsmål, for eksempel varighet for varebil, blir stående uansett svar.
- De korte tipsene ved hover og navnene på sluttalternativene i treet er egne korte formuleringer. De erstatter ikke teksten i scenarioene.
- Bildet er hentet fra HTML-rapporten (`bilder/bil-header.jpg`), og BDO-logoen er tegnet fra vektorlogoen i PowerPoint-filen (`bilder/bdo-logo.svg`) i BDO-farger: rød strek og blå bokstaver.

## Redigere tekst (uten programmering)

1. Åpne `Rediger.html` i nettleseren.
2. Brett ut et scenario, og endre tekstfeltene (farge og fet per tekstbit kan endres, avsnitt og tekstbiter kan legges til og slettes).
3. Trykk **Last ned ny MVA-oppslagsverk.html** og erstatt den gamle filen. «Lagre redigering» / «Åpne lagret redigering» lar deg fortsette senere.

Ikke redigerbart i verktøyet: hvilke spørsmål/valg et scenario hører til, og de korte hover-tipsene (`SAMMENDRAG` i `kilde/app.del`).

## Bygge på nytt

Etter rettelser i `innhold.json` eller i `kilde/`:

```bash
python -I kilde/bygg.py
```

## Mappen

- `MVA-oppslagsverk.html`: verktøyet.
- `Rediger.html`: redigeringsverktøy for tekst (bygges av `bygg.py`).
- `innhold.json`: tekst og data.
- `bauda-tema.css`, `tema-palett.html`: Baudas farger og skrift.
- `kilde/`: byggekode (`app.del`, `common.js`, `print.css`, `bygg.py`).
- `bilder/`: bilde og logo som bygges inn.
- `eksempel-pdf/`: to PDF-er laget fra verktøyet.
