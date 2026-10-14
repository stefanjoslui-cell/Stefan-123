# Handoff: OFV-kontroller (VBA + Databricks/Python)

> Skrevet for en fremtidig Claude Code-sesjon uten kontekst fra denne
> samtalen. Les hele dokumentet før du gjør endringer - spesielt
> "Konvensjoner og arbeidsflyt" og "Apne punkter" nederst.

## 1. Hva dette prosjektet er

BDO-prosjekt ("Python program for bilregistrering") som kontrollerer en
bilforhandlers bokføring mot Opplysningsrådet for Veitrafikken (OFV) sitt
Transactions API - dvs. det offisielle registeret over eierskifter på
kjøretøy i Norge. Det finnes **tre uavhengige kontroller**:

1. **Kontroll solgte biler** - er bokførte bilsalg faktisk registrert som
   eierskifte i OFV, på riktig dato?
2. **Varekjøp bruktbil** - er bruktbiler kjøpt inn i en periode faktisk
   OFV-bekreftet og bokført, og kan UB (biler fortsatt på lager)
   forklares?
3. **Kontroll Demobil** - står en bil forhandleren sier er demobil
   fortsatt registrert på forhandlerens eget navn i OFV (ikke videreselgt
   uten at det er fanget opp)?

Løsningen finnes i to parallelle former:

- **`OFV_RefreshInfo.bas`** - en produksjonsklar VBA-makro som kjøres
  direkte i en Excel-arbeidsbok (knapper + Input/Resultat/Kontroll-faner
  i samme bok). Dette er **sannheten** - all kontrollogikk er designet og
  herdet her først.
- **`ofv_kontroller_databricks.py`** - en Python-port av samme logikk,
  til en Databricks-notebook (serverless compute). Leser tre separate
  Excel-filer, skriver EN samlet output-arbeidsbok + en HTML-rapport.

**VBA-filen er master.** Når kontrollogikken endres, skal den endres i
VBA FØRST (det er der den er herdet mot ekte data), og deretter porteres
til Python. De to filene representerer samme logikk om omtrent samme
tidspunkt, men er ikke automatisk synkronisert - sjekk alltid om en
VBA-endring også må speiles i `ofv_kontroller_databricks.py`.

## 2. Filer i repoet

| Fil | Status | Beskrivelse |
|---|---|---|
| `OFV_RefreshInfo.bas` | **Aktiv, master** | All VBA-logikk for de tre kontrollene. ~5300 linjer. |
| `ofv_kontroller_databricks.py` | **Aktiv** | Databricks/Python-port. ~1850 linjer, 5 "celler" (`# COMMAND ----------`). |
| `HentOFVTransaksjoner.bas` | **Forlatt prototype** | Tidlig, enkel testmakro (ett regnr pr rad i kolonne B, ett API-kall pr rad). Bruker en ANNEN (eldre/feil) URL (`https://data.ofv.no/transactions/v1/query` i stedet for `https://api.ofv.no/transactions/v1/`). Ikke rør denne med mindre brukeren spesifikt ber om det - den er ikke i aktiv bruk. |
| `.claude/settings.json` | Config | Minimal, sjekk innhold selv ved behov. |

Det finnes ingen `main`/`master`-gren i dette repoet slik det er klonet -
all historikk (46 commits) ligger på `claude/kind-gates-rezvpl`, som er
branchen denne sesjonen jobber på. Hvis du starter en ny sesjon og får en
annen branch oppgitt, sjekk `git log --oneline` for å orientere deg -
denne filen kan være foreldet hvis mye har skjedd siden.

## 3. Arkitektur: `OFV_RefreshInfo.bas`

### 3.1 Input-arket ("Input"-fanen)

Alle tre kontroller leser fra SAMME Excel-ark, navngitt "Input", men fra
egne kolonne-seksjoner side ved side. Input-arket leses KUN fra - makroen
skriver aldri til det.

```
B1 = OFV API-nøkkel (evt. navngitt område OFV_API)
B2 = Statens Vegvesen API-nøkkel (evt. navngitt område SVV_API) - valgfri

Kontroll 1 - "Kontroll solgte biler" (kolonne A-D):
  B8         = Juridisk enhet (Selger) - org.nr (valgfri - styrer "Selvhandel"-sjekken)
  A12 og ned = Regnr, B12 og ned = VIN (TO kolonner, bare en fylles ut pr rad)
  D12 og ned = Bokført dato

Kontroll 2 - "Varekjøp bruktbil" (kolonne F-H):
  G8         = Juridisk enhet (Org Nr) - kjøper
  G9 / G10   = Dato fra / Dato til (kontrollperioden)
  F12 og ned = Innkjøp i perioden (regnr/VIN BLANDET i samme kolonne)
  G12 og ned = IB - kjøpt forrige periode (regnr/VIN blandet)
  H12 og ned = UB - fortsatt på lager (regnr/VIN blandet)

Kontroll 3 - "Kontroll Demobil" (kolonne J-L):
  K8         = Juridisk enhet - org.nr
  K9 / K10   = Dato fra / Dato til (kun til visning, påvirker ikke logikken)
  J12 og ned = Regnr/VIN BLANDET i samme kolonne (K er IKKE i bruk som liste)
  L12 og ned = Bokført inn dato
```

Merk at Kontroll 1 har regnr/VIN i **to parallelle kolonner**
(`ReadInputIdentifier`), mens Kontroll 2 og 3 har dem **blandet i én
kolonne** (`ReadIdentifikatorListe` / direkte `NormalizeIdentifier` +
`ErVIN`). Dette er historisk (Kontroll 1 ble bygget først, før
én-kolonne-mønsteret ble innført) - ikke en feil, men vær klar over
forskjellen hvis du skal endre inputlesing.

VIN kjennes fra regnr automatisk på lengde: `ErVIN(tekst) = (Len(tekst) = 17)`
(ISO 3779-standarden). Dette brukes konsekvent over hele kodebasen,
inkludert i Python-porten.

### 3.2 Output pr kontroll

Hver kontroll har sin egen Resultat-fane (full, rå
transaksjonshistorikk, én rad pr OFV-transaksjon) og sin egen
Kontroll-fane (den bearbeidede analysen - KPI-bokser, statuskoder,
fargelegging):

| Kontroll | Resultat-fane | Kontroll-fane |
|---|---|---|
| 1 | `Resultat` | `Kontroll solgte biler` |
| 2 | `Resultat Varekjop` | `Kontroll Varekjop Bruktbil` |
| 3 | `Resultat Demobil` | `Kontroll Demobil` |

Resultat-fanene bygges av DELTE funksjoner (`GetFieldMap`,
`GetOrCreateResultTable`, `ResizeResultTable`, `FormatResultTable`,
`FormatResultTableGrouping`, `ApplyCalculatedColumns`) - identisk
24-kolonne skjema og formatering for alle tre. Kontroll-fanene bygges av
hver sin dedikerte `Update*ControlSheet`-sub, og ser forskjellige ut
(se under).

Hvert ark bygges **fra scratch hver kjøring** (`ws.Cells.Clear` først) -
det er ingen inkrementell oppdatering, og ingen levende Excel-formler
(alt regnes ut i VBA og skrives som faste verdier) - bortsett fra
`ApplyCalculatedColumns` i Resultat-fanene, som setter
`SelgerType`/`KjøperType` som Excel-formler (Python-porten gjør dette
enklere: regner det direkte i Python, se §4).

### 3.3 Kontroll 1: Kontroll solgte biler

**Nøkkelfunksjon:** `BuildKontrollRow` (kalt fra `KjorKontrollSolgteBiler`).

Logikk: for hver bil, hent HELE transaksjonshistorikken (ett API-kall,
ufiltrert). Finn transaksjonen som ligger **nærmest bokført dato i tid**
(uansett om før eller etter) - IKKE bare siste/nyeste transaksjon. Dette
"nærmeste, ikke siste"-prinsippet er gjennomgående i hele kodebasen (også
i `FinnLagerperiode` i Kontroll 2, se §3.4).

- Har OFV **ingen** transaksjoner i det hele tatt for kjøretøyet, brukes
  førstegangsregistreringsdato fra Statens Vegvesen (SVV) som reserve
  (`FetchVehicleInfoFromSVV`) - både som kontrollgrunnlag og i kolonnen
  "Forstegangsregistrert". Mangler SVV-nøkkelen, hoppes SVV bare over
  (ingen hard feil).
- Valgfri "Selvhandel"-sjekk: er `B8` (Selger-org.nr) fylt ut, og er
  SAMME org.nr både selger og kjøper i den matchede transaksjonen,
  flagges raden alltid rød (uavhengig av faktisk dagers avvik).
- Fargekoding av "Dager avvik": 0-2 dager grønt, 3-14 gult, 15+ rødt
  (konstantene `AVVIK_GRONN_MAX`/`AVVIK_GUL_MAX`). Selvhandel overstyrer
  alltid til rødt.
- Kontroll-arket (`UpdateControlSheet`) viser **én rad pr bil** (ikke all
  historikk - det ligger i Resultat-arket), sortert med størst avvik
  først (`SorterKontrollRadPaAvvik`).

### 3.4 Kontroll 2: Varekjøp bruktbil

Dette er den klart mest ombygde kontrollen i prosjektet - se git-loggen
for hele utviklingen (fra en enkel "kjøpt i perioden"-sjekk, via flere
runder med feilrettinger og presiseringer, til dagens delte
overlapp-motor). **Les denne historien i git-loggen hvis du skal endre
logikken** - flere tidligere design ble forkastet for gode grunner.

**Tre lister** på Input-arket (F/G/H), alle blandet regnr/VIN:

- **Innkjøp** (Seksjon A) - biler forhandleren sier var på lager i
  PERIODEN (IKKE strengt "kjøpt nøyaktig i denne perioden" - dette er en
  viktig presisering fra brukeren etter en forvirringsrunde, se §5).
- **IB** (Seksjon B) - biler kjøpt FØR perioden startet (forrige
  periodes UB).
- **UB** (Seksjon C) - biler fortsatt på lager ved kontrolltidspunktet.

Alle tre er valgfrie - kontrollen fungerer helt fint med BARE
Innkjøp-listen utfylt (ingen IB/UB).

**Delt motor** (brukt av Seksjon A og B identisk):

- `FinnLagerperiode(vehicleTxRows, buyerOrgNo, kjopGrense, ByRef kjoptTxRow, ByRef solgtTxRow, ...)` -
  en `Sub` (ikke `Function` - bruker `ByRef`-utparametre). Finner:
  - **Kjøpt**: forhandlerens SISTE kjøp av bilen (`ToOwnerOrgNo = dealer`)
    på eller FØR `kjopGrense`.
  - **Solgt**: forhandlerens FØRSTE videresalg (`FromOwnerOrgNo = dealer`)
    ETTER det kjøpet - helt uten øvre tidsbegrensning (kan skje langt
    utenfor kontrollperioden).
  - `kjopGrense` er det som skiller Seksjon A fra B: for A = `Dato til`
    (kjøpet kan ha skjedd når som helst frem til periodeslutt), for B =
    `Dato fra - 1 dag` (en IB-bil må per definisjon være kjøpt FØR
    perioden startet).
- `LagerStatusTekst(kjoptTxRow, solgtTxRow, dateFra)` - avgjør
  OK/Avvik: **OK** så lenge bilens lagerperiode overlapper
  kontrollperioden (ikke solgt, ELLER solgt PÅ/ETTER periodestart).
  **Avvik** hvis solgt FØR periodestart, eller hvis det ikke finnes noe
  opprinnelig kjøp hos forhandleren å sjekke mot.

Dette er en **dato-intervall-overlapp-test** ("var bilen på lager i
perioden"), ikke en eksistens-sjekk - en bil kjøpt forrige periode og
solgt i inneværende periode er OK, ikke avvik (se §5 for hvorfor dette
tok flere runder å lande på riktig).

**"I OFV-liste"-flagget** (Seksjon A) / **"På UB-liste"** (Seksjon B) er
RENT INFORMATIVT - de påvirker ikke OK/Avvik-status. "I OFV-liste" viser
om selve kjøpet skjedde NØYAKTIG i perioden (mot org+periode-bulklisten
fra steg 1, `FetchOFVTransactionsByBuyerOrg`), mens hovedstatusen bruker
den bredere overlapp-logikken.

**Biler OFV sier er kjøpt, men som IKKE finnes i Innkjøp-listen**
("mangler i bokføring") hentes OGSÅ med full transaksjonshistorikk (ikke
bare den ene org-bulk-raden) slik at samme `FinnLagerperiode`-motor kan
brukes på dem - se `BuildManglendeBokforingRow`.

**Kontroll-arket** (`UpdateVarekjopControlSheet`) viser **én rad pr bil**
i en kompakt tabell (ikke "bilkort" - det ble prøvd og reversert, se
git-commit `75e1939` → `5fcd136`). Kolonneorden (etter siste endring,
commit `2b8ce2f`): Regnr/VIN, Chassisnummer, Modell, [flagg], **På
lager fra, På lager til** (gruppert sammen, lagerperioden FØRST), **Kjøpt
fra, Solgt til** (transaksjonspartene, gruppert sammen ETTER), Status. En
egen `Forhandler: <navn>`-linje står over hver tabell (siden forhandleren
er lik for alle rader - unngår en redundant kolonne).

Helper-funksjoner delt mellom seksjonene: `SkrivForhandlerLinje`,
`SkrivBilTabellHeader`, `SkrivBilRad`.

**Viktig, nå SLETTET logikk** (ikke gjenbruk denne hvis du ser den i
gamle commits): `RegistreringsType = "Juridisk eierskifte (ikke i
bestand)"` ble tidligere brukt som signal for "bilen forlot bestanden",
men ble funnet UPÅLITELIG ved analyse av ekte data (se §5) og er fjernet
helt fra Seksjon B sin salg-sjekk.

### 3.5 Kontroll 3: Kontroll Demobil

**Nøkkelfunksjon:** `BuildDemobilRow`.

Enkel logikk: hent hele transaksjonshistorikken for bilen, finn den
NYESTE registrerte transaksjonen (uansett dato), og sjekk om kjøperen der
(`ToOwnerOrgNo`) er lik org.nr i `K8`. **OK** hvis ja (bilen står fortsatt
hos selskapet), **Avvik** hvis nei (sannsynligvis videreselgt uten at det
ble fanget opp). Input er regnr/VIN blandet i én kolonne (`J`, konstant
`COL_IDENT_3` - kolonne `K` er IKKE i bruk til liste-data, bare `K8`/`K9`/`K10`
over er konfigurasjonsceller).

Denne kontrollen er den ENKLESTE av de tre og har endret seg minst - kun
inputkolonne-layouten ble endret underveis (fra to kolonner J/K til én
blandet kolonne J), selve logikken var allerede korrekt fra starten.

### 3.6 Delte hjelpefunksjoner (brukes av flere/alle kontroller)

| Funksjon | Hva den gjør |
|---|---|
| `BuildVehicleKey(vin, regNo)` | `"VIN\|" & vin` eller `"REG\|" & regNo` - kanonisk dictionary-nøkkel for en bil, brukt over ALT (kryss-liste-oppslag, API-cache-nøkler). |
| `NormalizeIdentifier(value)` | `UCase(Trim(Replace(value, " ", "")))` - brukt på ALLE org.nr/regnr/VIN-sammenligninger. |
| `ErVIN(tekst)` | `Len(tekst) = 17`. |
| `ComputeOwnerLabel(ownerType, companyName)` | Privat vinner over firmanavn, ellers firmanavn hvis det finnes, ellers rå eiertype-tekst. Brukt for Selger/Kjøper-visning OG tilsvarer `SelgerType`/`KjøperType`-Excel-formelen i Resultat-tabellene. |
| `TolkBokfortDato(raw)` | Tolker brukerinntastet dato EKSPLISITT som dag.måned.år (norsk), aldri via `CDate` (unngår regionsinnstillings-avhengighet). Ekte Excel-datoer brukes direkte. |
| `DateFromISO(isoDate)` | Parser OFV/SVV sine ISO-datoer via faste tegnposisjoner - helt locale-uavhengig. |
| `FetchOFVTransactions` | Ett kall pr kjøretøy, ufiltrert, cursor-paginert (1000/side), sortert nyeste først. |
| `FetchOFVTransactionsByBuyerOrg` | Org+periode-scoped bulk-liste (kun Kontroll 2). |
| `FetchVehicleInfoFromSVV` | SVV-reserve (kun Kontroll 1). |
| `PostOFVWithRetries` / `GetSVVResponse` | Retry på 429/5xx, umiddelbar feil på 401/403/404. VBA bruker **lineær** backoff (`RETRY_WAIT_MS * forsøksnummer`) - Python-porten bruker ekte eksponentiell backoff, se §4. |
| `JSON_ExtractValue` / `JSON_ExtractObject` / `JSON_ArrayAllElements` / `JSON_FindMatchingBrace` | Håndrullet tegn-for-tegn JSON-parser (VBA har ikke innebygd JSON). `JSON_ExtractValue` gjør en FLAT `InStr`-søk etter `"key":` HVOR SOM HELST i JSON-strengen, uavhengig av nesting - en forgivende "finn dette navnet uansett hvor det ligger"-strategi. Python-porten bruker ekte `json`-modulen for strukturert tilgang, men replikerer den samme forgivende søke-strategien (`find_any`) for SVV-responsen spesifikt, siden dens nesting er ukjent. |
| `FargeleggStatusCelle(celle, statusText)` | Grønn hvis status starter med "OK", rød hvis "Avvik" - brukt konsekvent over alle kontroll-ark. |

### 3.7 Konvensjoner å kjenne til før du endrer VBA-koden

- **Dynamisk rad-cursor**: ark bygges med en `r`-variabel som økes
  fortløpende (`r = r + 1`, `r = r + 2` osv.), IKKE faste rad-konstanter -
  fordi seksjoner kan hoppes over (IB/UB er valgfrie). Følg dette
  mønsteret hvis du legger til nye seksjoner.
- **Dictionary-nøkkel-disiplin**: hver `Build*Row`-funksjon returnerer en
  `Scripting.Dictionary` hvis nøkler MÅ matche nøyaktig det
  render-loopen leser via `row("NokkelNavn")` - en Dictionary returnerer
  `Empty` for en manglende nøkkel i stedet for å kaste feil, så en
  skrivefeil i nøkkelnavn feiler HELT STILLE. Kryss-sjekk alltid med
  `grep` etter en endring (se eksempel-arbeidsflyt i §6).
- **VBA linje-fortsettelse** (`" _"` på slutten av en fysisk linje): hvis
  du skriver et script for å verifisere strukturell balanse
  (If/EndIf/Sub/Function/parenteser), må du FØRST slå sammen
  fortsettelseslinjer til "logiske linjer" - en naiv linje-for-linje
  regex gir falske negative-dybde-feil på multi-linje `If`-betingelser.
  Se eksempel i §6.
- **Hele arket bygges på nytt hver kjøring** (`ws.Cells.Clear` først) -
  ikke prøv å gjøre inkrementelle oppdateringer av eksisterende celler.
- **Fargekonstanter**: `COLOR_GREEN_FILL`/`FONT`, `COLOR_YELLOW_FILL`/`FONT`,
  `COLOR_RED_FILL`/`FONT` er Excels innebygde Bra/Nøytral/Dårlig-farger
  (RGB 198,239,206 / 255,235,156 / 255,199,206). Header-blå er
  `RGB(31,78,120)`, lys seksjons-bakgrunn `RGB(221,235,247)`.

## 4. Arkitektur: `ofv_kontroller_databricks.py`

Python-port for en Databricks-notebook (serverless compute), strukturert
som 5 "celler" via `# COMMAND ----------`-markører (gjenkjennes av
Databricks som en notebook når filen importeres):

1. **Konfigurasjon** - stier, API-endepunkter, retry-konstanter, secrets,
   BDO-fargepalett, `FIELD_MAP` (24-kolonne skjema).
2. **Hjelpefunksjoner** - tekst/dato/identifikator-helpers, HTTP med
   retry, OFV/SVV-kall, transaksjonsrad-bygging, input-lesere
   (`read_konfig`/`read_biler`), Excel-stil-helpers (BDO-farget).
3. **Hovedlogikk** - `run_kontroll_solgte_biler`,
   `run_kontroll_varekjop_bruktbil`, `run_kontroll_demobil` - hver en
   selvstendig pipeline som leser sin inputfil, kaller API, returnerer en
   "bundle"-dict med rader + stats.
4. **Output** - bygger ÉN samlet `openpyxl.Workbook` (Input-manifest +
   Resultat/Kontroll for alle tre) + en Jinja2-generert HTML-rapport med
   samme innhold.
5. **Oppsummering** - printer stats pr kontroll + filstier.

### 4.1 Input - tre separate filer (IKKE samme Input-fane-mønster som VBA)

Dette var en EKSPLISITT designbeslutning i denne sesjonen (brukeren ville
ha en annen inputstruktur for Databricks-pipelinen enn VBA sin
delt-Input-ark-modell): hver kontroll har sin EGEN Excel-fil, med fast
filnavn i `INPUT_DIR`:

```python
INPUT_FILE_SOLGTE_BILER = INPUT_DIR / "kontroll_solgte_biler_input.xlsx"
INPUT_FILE_VAREKJOP     = INPUT_DIR / "kontroll_varekjop_bruktbil_input.xlsx"
INPUT_FILE_DEMOBIL      = INPUT_DIR / "kontroll_demobil_input.xlsx"
```

Hver fil har to faner:

- **"Konfig"** - to kolonner `Felt`/`Verdi`, lest til en dict
  (`read_konfig`). Feltnavn pr kontroll: `OrgNrSelger` (valgfri, K1);
  `OrgNrKjoper`+`DatoFra`+`DatoTil` (K2); `OrgNr`+`DatoFra`+`DatoTil`
  (K3, datoene er valgfrie/kun visning).
- **"Biler"** - kolonnen `Identifikator` (regnr ELLER VIN, auto-oppdaget
  på lengde) + en kontroll-spesifikk ekstrakolonne: `BokfortDato` (K1),
  `Liste` med verdiene `Innkjop`/`IB`/`UB` (K2), `BokfortInnDato` (K3).

Kolonnenavnene er ASCII-trygge med vilje (ingen `ø`/`æ`/`å`/`/`) for å
være enkle for en annen pipeline (Databricks-jobb, ETL) å generere
programmatisk.

### 4.2 Output - én samlet arbeidsbok + HTML

`OUTPUT_DIR/OFV_Kontroller_<tidsstempel>.xlsx` med fanene:

1. **Input** - manifest: hvilken kildefil/orgnr/periode ble brukt for
   hver kontroll, OG en flat liste over ALLE bilene som ble lest, tagget
   med hvilken kontroll de hører til (`skriv_input_fane`).
2. `Resultat solgte biler` / `Kontroll solgte biler`
3. `Resultat Varekjop` / `Kontroll Varekjop Bruktbil`
4. `Resultat Demobil` / `Kontroll Demobil`

Samme innhold speiles i `OUTPUT_DIR/OFV_Kontroller_<tidsstempel>.html`
(Jinja2-template, samme BDO-farger).

### 4.3 BDO-fargepalett (ikke Excels innebygde grønt/gult/rødt)

Brukeren ba spesifikt om BDOs egen profilpalett i stedet for Excels
innebygde farger. Kilde:
`cloud.brandmaster.com/point/no/bdobc/component/default/113577`.

```python
BDO = {
    "red": "E81A3B", "charcoal": "333333", "slate": "5B6E7F",
    "burgundy": "98002E", "pale_charcoal": "E7E7E7",
    "jade": "009966", "ocean": "008FD2", "gold": "D67900", "white": "FFFFFF",
}
```

**VIKTIG**: BDO Blå (`#22409A`) er per BDOs egen merkevareguide
FORBEHOLDT LOGOEN og skal IKKE brukes på andre flater - derfor er den
IKKE med i `FARGE_*`-semantikk-mappingen. Rollene er: Charcoal =
header/tittel-bakgrunn, Pale Charcoal = seksjons/KPI-bakgrunn, Jade = OK,
Red = Avvik, Gold = Varsel/mellomnivå.

### 4.4 Secrets og stier

```python
INPUT_DIR = "/Workspace/Users/stefan.luidold@bdo.no/OFV API/Input"
OUTPUT_DIR = "/Workspace/Users/stefan.luidold@bdo.no/OFV API/Output"
SECRET_SCOPE = "ofv-api"
OFV_API_KEY = dbutils.secrets.get(scope=SECRET_SCOPE, key="ofv-api-key")  # pakrevd
SVV_API_KEY = dbutils.secrets.get(scope=SECRET_SCOPE, key="svv-api-key")  # valgfri
```

Disse stiene/secret-navnene er brukerens FAKTISKE Databricks-oppsett
(oppgitt direkte av brukeren) - ikke placeholder-verdier. Hvis brukeren
sier stiene har endret seg, oppdater dem i Celle 1.

### 4.5 Forbedringer i forhold til VBA (gjort med vilje, ikke drift)

- **Ekte eksponentiell backoff** (`RETRY_BASE_S * 2**(attempt-1)` + jitter)
  i stedet for VBA sin lineære `RETRY_WAIT_MS * attempt`. Brukeren ba
  eksplisitt om dette i sin opprinnelige spesifikasjon for
  Databricks-scriptet.
- **Ekte JSON** (`requests`/`json`-modulen) i stedet for en håndrullet
  parser - strukturert tilgang via `.get()`-kjeder som følger samme
  objektpath som VBA (`from.owner.companyInfo.name` osv.), UNNTATT for
  SVV-responsen der nestingen er ukjent (der brukes `find_any`, en
  rekursiv "søk etter denne nøkkelen hvor som helst"-funksjon som
  replikerer VBAs forgivende strategi).
- **SelgerType/KjøperType regnes direkte i Python** når transaksjonsraden
  bygges (`compute_owner_label`), ikke som en etterskuddsvis Excel-formel
  (`ApplyCalculatedColumns` i VBA).
- API-nøkler ALDRI hardkodet - alltid via `dbutils.secrets.get`.

### 4.6 Testing - ingen ekte API-tilgang i denne sesjonen

Scriptet er IKKE kjørt mot det ekte OFV/SVV-API-et (ingen nettverkstilgang
til disse fra denne økten). Det er verifisert med en **mocket dry-run**:
falsk `dbutils` (secrets+widgets), monkey-patchet
`requests.Session.post`/`.get` som returnerer syntetiske
JSON-responser formet som ekte OFV-transaksjoner, og tre syntetiske
input-arbeidsbøker bygget med `openpyxl`. Testharnessen selv ble SLETTET
etter bruk (lå i scratchpad, ikke i repoet) - hvis du vil gjenskape den,
se mønsteret:

1. Bygg `Konfig`/`Biler`-arbeidsbøker med `openpyxl.Workbook()`.
2. Lag en `fake_dbutils = types.SimpleNamespace(secrets=..., widgets=...)`.
3. Monkey-patch `requests.Session.post`/`requests.Session.get` til
   funksjoner som returnerer en `FakeResponse`-klasse med `.status_code`
   og `.json()`.
4. Les scriptets kildekode, bytt ut de hardkodede `/Workspace/...`-stiene
   med scratch-stier via streng-erstatning, og `exec()` den resulterende
   kildekoden med `dbutils`/`display` injisert i namespace.

Dette fanget flere reelle feil under utvikling (bl.a. en dødkode-rest i
`skriv_kontroll_varekjop`) og bør gjentas etter enhver større endring i
scriptet, siden det ikke finnes noen automatisert testsuite.

**Ikke bekreftet mot ekte API-respons**: feltnavnet
`organizationNumber` inni `companyInfo`-objektet (antatt, basert på
samme mønster som filter-navnene `fromOrganizationNumber`/
`toOrganizationNumber`, men ALDRI verifisert mot en faktisk OFV-respons -
se `build_transaction_row` i Python og `BuildTransactionRow` i VBA).
Samme antagelse gjelder SVV-feltnavnene
`registrertForstegangNorgeDato`/`registrertForstegangDato`. Hvis
selvhandel-/orgnr-sjekkene eller SVV-datoene ikke treffer riktig mot ekte
data, sjekk disse feltnavnene først.

## 5. Viktige funn og beslutninger underveis (domenekunnskap / "memory")

Disse er ikke åpenbare fra koden alene - de kom fram gjennom flere
runder med brukeren og reell datadrøfting. Verdt å kjenne til før du
endrer logikk:

- **`RegistreringsType`-feltet er UPÅLITELIG som signal.** En analyse av
  et ekte OFV-dataeksport (`Bok1.xlsx`, 450 rader/129 biler, sentrert på
  Toyota Romerike AS) viste at `RegistrationType = "Juridisk eierskifte
  (ikke i bestand)"` manglet på ca. 9% av reelle salg (Næring→Næring og
  Næring→Privat-transaksjoner) og dukket opp uventet på et par
  Privat→Næring-kjøp. Derfor ble dette feltet FJERNET fra Seksjon B sin
  salgs-sjekk (commit `43cce6e`) - stol IKKE på det som et pålitelig
  "forlot bestanden"-signal noe sted i kodebasen.
- **"Innkjøp i perioden" (kolonne F) betyr IKKE strengt "kjøpt nøyaktig i
  denne perioden"** - det betyr "biler forhandleren selv sier var PÅ
  LAGER i perioden" (samme spørsmål som IB-listen svarer på). Dette var
  en reell misforståelse tidlig i prosjektet som så ut som en bug
  (brukeren rapporterte 65% avvik-rate og trodde noe var feil), men viste
  seg å være en begrepsavklaring - se commits rundt `639145d`. Seksjon A
  og B bruker derfor i dag IDENTISK overlapp-logikk
  (`FinnLagerperiode`/`LagerStatusTekst`), bare med ulik `kjopGrense`.
- **Matching skal alltid finne transaksjonen som involverer DEN OPPGITTE
  forhandleren** - aldri en senere, urelatert transaksjon mellom andre
  parter, selv om bilen har skiftet eier flere ganger etter at
  forhandleren solgte den videre. Dette var en reell bug (fantes en
  fallback som matchet "en hvilken som helst transaksjon i perioden") som
  ble fikset i commit `dec5fa3`.
- **"Nærmeste, ikke siste"** er et gjennomgående prinsipp: både Kontroll
  1 (nærmest bokført dato) og Kontroll 2 (siste kjøp PÅ ELLER FØR en
  grense, første salg ETTER det kjøpet) bruker dette - ALDRI "bare siste
  transaksjon i historikken" uten kontekst.
- **BDO Blå er reservert til logoen** - dette kom direkte fra BDOs egen
  merkevareside og er fulgt strengt i Python-portens fargevalg (se §4.3).
- **Et uttrykkelig ønske om "en rad pr bil"** i Kontroll 2 sitt
  kontroll-ark - et "bilkort"-design (flere rader pr bil, med en egen
  "Forhandler"/"På lager fra-til"-seksjon) ble bygget, vist til brukeren,
  og deretter EKSPLISITT reversert til en kompakt tabell (commits `75e1939`
  → `5fcd136`). Ikke gjeninnfør kort-formatet uten at brukeren spør om
  det igjen.

## 6. Konvensjoner og arbeidsflyt brukt i denne og forrige sesjon

- **Alle samtaler, kommentarer og forklaringer er på norsk** (bokmål,
  ASCII-vennlig i kode - VBA-kommentarer unngår konsekvent æ/ø/å i
  identifikator-/kommentar-tekst av historiske kompatibilitetsgrunner,
  mens Python-koden bruker ekte UTF-8 fritt).
- **Spør FØR store, tvetydige endringer** - brukeren har gjentatte ganger
  bedt om at Claude stiller avklarende spørsmål (`AskUserQuestion`) før
  en antagelse tas for strukturelle/design-beslutninger, spesielt når
  flere rimelige tolkninger finnes. Ikke gjett på vesentlige
  arkitekturvalg.
- **Etter HVER kodeendring i `OFV_RefreshInfo.bas`**, kjør en strukturell
  sanity-sjekk før commit - et Python-engangsskript (ikke lagret i
  repoet, skriv det på nytt ved behov) som:
  1. Slår sammen VBA-linjefortsettelser (`" _"`) til logiske linjer.
  2. Teller `Sub`/`End Sub`, `Function`/`End Function`, `With`/`End With`,
     `For`/`Next` - skal være like par.
  3. Teller åpne/lukk-parenteser - skal være like.
  4. Sporer `If`/`End If`-dybde linje for linje (på de SAMMENSLÅTTE
     logiske linjene) - skal aldri gå negativt, og slutte på 0.
  5. Kryss-sjekk med `grep` at hver `Build*Row`-funksjons
     dictionary-nøkler (`result("X") = ...`) matcher nøyaktig det
     tilhørende render-loop leser (`row("X")`).
- **Etter hver verifisert endring**: `git add` + `git commit` (beskrivende
  norsk commit-melding, flere avsnitt for større endringer) + `git push -u
  origin claude/kind-gates-rezvpl` + `SendUserFile` av den endrede filen
  + en kort norsk oppsummering til brukeren i chat.
- **Bekreft alltid at Kontroll 1 og 3 forblir urørt** når du jobber på
  Kontroll 2 (eller omvendt) - brukeren har bedt om dette eksplisitt flere
  ganger, og det er lett å sjekke med `git diff` gruppert på
  funksjonsnavn (`git diff -U0 | grep -oE "@@.*@@ .*"`).
- Ved større Python-endringer: skriv en mocket dry-run-test (se §4.6)
  FØR commit, siden det ikke finnes ekte API-tilgang i denne typen økt.

## 7. Relevante skills (Claude Code)

- **`anthropic-skills:xlsx`** - brukt for å analysere opplastede
  Excel-eksporter fra brukeren (ekte OFV-transaksjonsdata) for å
  validere/forbedre kontrollogikken mot virkeligheten (se §5, funnet om
  `RegistreringsType`). Bruk denne igjen hvis brukeren laster opp nye
  dataeksempler å verifisere logikk mot.
- **`AskUserQuestion`**-mønsteret er sentralt i hvordan dette prosjektet
  har blitt bygget - design- og arkitekturbeslutninger har konsekvent
  blitt avklart via strukturerte spørsmål med 2-4 konkrete alternativer
  FØR koden skrives, ikke i etterkant. Fortsett dette mønsteret.
- Ingen andre skills (dataviz, docs, pptx osv.) er relevante for dette
  prosjektet per nå - det er ren VBA/Python-kodeutvikling pluss
  periodisk xlsx-dataanalyse.

## 8. Åpne punkter / ting å verifisere videre

1. **OFV API-feltnavn ikke bekreftet**: `organizationNumber` i
   `companyInfo`-objektet er en antagelse (se §4.6). Verifiser mot en
   ekte respons første gang noen har tilgang.
2. **SVV-responsens nøyaktige struktur er ukjent** - `find_any`/
   `JSON_ExtractValue` sin forgivende søkestrategi er et bevisst
   arbeid-rundt, ikke en permanent løsning. Hvis SVV-fallbacken
   (Kontroll 1) gir feil datoer mot ekte data, se her først.
3. **Databricks-scriptet er ALDRI kjørt mot ekte Databricks/OFV/SVV** -
   kun mocket lokalt (se §4.6). Første ekte kjøring bør overvåkes nøye,
   spesielt pagineringslogikken (cursor-basert, 1000 rader/side) og
   secrets-oppsettet (`ofv-api`-scope med `ofv-api-key`/`svv-api-key`).
4. **Ingen eksempel-input-filer er lagt i repoet** - brukeren har bedt om
   en forklaring av formatet (gitt i chat, se §4.1), men ingen faktiske
   `.xlsx`-maler er committet. Vurder å lage disse hvis brukeren ber om
   det, eller hvis en Databricks-pipeline skal generere dem automatisk.
5. **`HentOFVTransaksjoner.bas` er ikke ryddet bort** - den er en forlatt
   prototype med en potensielt FEIL URL. Vurder å slette den eller merke
   den tydelig som utdatert hvis brukeren spør om opprydning.
6. **Ingen CLAUDE.md finnes i repoet** - brukeren valgte eksplisitt bort
   dette til fordel for denne frittstående `HANDOFF.md`-filen (en
   Claude Code-sesjon laster IKKE denne filen automatisk - du må lese den
   selv, f.eks. ved sesjonsstart eller når brukeren henviser til den).
