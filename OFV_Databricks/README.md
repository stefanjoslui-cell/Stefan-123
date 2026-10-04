# OFV API Workspace

Denne mappen inneholder en Databricks-basert løsning for å:

* lese rå Excel-filer per selskap
* transformere dem til standardiserte inputfiler for OFV-kontroller
* hente data fra OFV API og eventuelt Statens vegvesen
* skrive resultater til Excel- og HTML-rapporter

README-filen skal holdes oppdatert når notebooker, mapper, filstruktur eller arbeidsflyt endres.

## Struktur i workspacet

Rotmappe:
`/Workspace/Users/stefan.luidold@bdo.no/OFV API`

Innhold:

* [01_Verifisering_Input](#notebook-1535288192161841)
  * Leser råfiler fra `RåData/<Selskap>/`
  * Validerer og normaliserer innhold
  * Skriver 3 standardiserte inputfiler til `Input/<Selskap>/`
* [02_OFV_API_Kontroller](#notebook-1535288192161840)
  * Leser standardiserte inputfiler fra `Input/<Selskap>/`
  * Kaller OFV API og eventuelt SVV API
  * Skriver Excel- og HTML-resultater til `Output/<Selskap>/`
* [RåData](#folder-1535288192161839)
  * Undermapper per selskap
  * Hver undermappe skal hete det aktuelle selskapet
* [Input](#folder-1055969311049368)
  * Standardiserte inputfiler per selskap
* [Output](#folder-1055969311049369)
  * Resultatfiler per selskap
* [README.md](#file-1535288192161842)
  * Denne dokumentasjonen
* `Ressurser/`
  * Excel-leser (SheetJS) og fonter som bygges inn i HTML-rapporten, slik at den virker uten nett
  * Last opp filene fra repoets `OFV_Databricks/ressurser/` hit (mangler de, forsøker 02-notebooken å laste dem ned og lagre dem her)

## Mappestruktur per selskap

For hvert selskap brukes samme navn som undermappe i alle relevante mapper:

* `RåData/<Selskap>/`
* `Input/<Selskap>/`
* `Output/<Selskap>/`

Eksempel:

* `RåData/Toyota Bilia Oslo AS/`
* `Input/Toyota Bilia Oslo AS/`
* `Output/Toyota Bilia Oslo AS/`

## Arbeidsflyt

### Steg 1: Legg inn råfiler
Opprett en undermappe i `RåData/` med selskapsnavnet og legg rå Excel-filer i denne mappen.

### Steg 2: Kjør verifiseringsnotebooken
Kjør [01_Verifisering_Input](#notebook-1535288192161841).

Denne notebooken:

* scanner `RåData/` etter undermapper
* viser tilgjengelige selskaper i en dropdown-widget `SELSKAP`
* scanner `RåData/<Selskap>/` etter Excel-filer
* matcher filer automatisk til riktig kontroll basert på filnavn
* bygger standardiserte inputfiler i `Input/<Selskap>/`

### Steg 3: Kjør OFV-kontrollnotebooken
Kjør [02_OFV_API_Kontroller](#notebook-1535288192161840) med samme `SELSKAP`-verdi.

Denne notebooken:

* leser de tre standardiserte inputfilene fra `Input/<Selskap>/`
* hopper over kontroller der inputfilen mangler (i stedet for å stoppe)
* kjører kontrollene
* genererer samlet Excel-arbeidsbok og HTML-rapport i `Output/<Selskap>/`

Celler i 02_OFV_API_Kontroller:

* Celle 1: Konfigurasjon (widget, stier, secrets, farger)
* Celle 2: Hjelpefunksjoner (API-kall med buffer, orgnr-/datotolking, input-lesere)
* Celle 3: Hovedlogikk (de tre kontrollene)
* Celle 4: Excel-arbeidsbok
* Celle 5: HTML-rapport
* Celle 6: Oppsummering (inkl. tidsbruk og antall API-kall)

Under kjøringen viser Celle 3 et **fremdriftsvindu** (som frmFremdrift i VBA): samlet fremdrift med tid brukt og anslått gjenstående tid, én linje per kontroll (Venter / Kjører / Ferdig) og gjeldende OFV-/SVV-oppslag. Styres av `FREMDRIFT_VISNING` i Celle 1: `"auto"` (standard), `"widget"`, `"html"` eller `"tekst"` (én tekstlinje per bil som før). Viser `"auto"` et vindu som ikke oppdaterer seg, sett den til `"html"`. Med levende vindu skrives fremdriften som tekst hver 10. prosent (synlig i jobbloggen).

## HTML-rapporten

`OFV_Kontroller_<tidsstempel>.html` er én selvstendig fil i BDO-presentasjonsstil (Work Sans, rød aksent) som virker uten nett. Tilbake/Neste og prikkene nederst blar gjennom hele filen; menyen øverst på rapportsidene hopper direkte.

1. **Presentasjonen** (fra `kontroll_solgte_biler_8.html`, uendret): forside, før/nå og «Slik henter vi dataene» 1–3 med faste, anonymiserte eksempelbiler. «Til rapporten» hopper rett til rapportdelen.
2. **Rapporten**:
   * **Oversikt** – kjøreinfo og ett kort per kontroll
   * **Per forhandler** – fordeling av dagers avvik per datasett i Kontroll solgte biler; klikk en rad for å filtrere Solgte biler
   * **Solgte biler**, **Varekjøp bruktbil**, **Demobil** – KPI-bokser (klikk for å filtrere), kontrollregel, kontrolltabell og full OFV-historikk
   * **Analyse** – nøkkeltall og grafer (status per kontroll, avviksfordeling, per forhandler, lagertid, hvem kjøper/selger, merke, drivstoff, eierskifter per måned)
   * alle tabeller kan sorteres, søkes i og filtreres per kolonne (tekst, verdiliste, min/maks, fra/til-dato) og på status, og filtrert utvalg kan lastes ned som CSV
3. **Spørsmål**-siden til slutt.

**Tannhjulet** (nede til høyre, eller `#rediger` bak filnavnet) åpner panelet for bilder (forside, spørsmål-side, logo i Hovedboken-boksen) og datasett. Opptil 20 forhandlere fra Kontroll solgte biler kan limes inn (kopier tabellen inkl. overskriftsrad fra Excel) eller lastes opp som .xlsx – også Excel-arbeidsboken fra 02-notebooken. Databricks-kjøringen er alltid første datasett (navn = `SELSKAP`), og de innlimte vises sammen med den på Per forhandler, Solgte biler og Analyse. Alt skjer lokalt i nettleseren; innlimte data lagres i nettleseren per selskap.

## Automatisk filgjenkjenning i 01_Verifisering_Input

Notebooken bruker filnavn til å avgjøre hvilken kontroll en råfil tilhører.

Gjeldende regler:

* Filnavn som inneholder `solgt` eller `salg` → Kontroll solgte biler
* Filnavn som inneholder `varekj` eller `brukt` → Varekjøp bruktbil
* Filnavn som inneholder `demo` → Kontroll Demobil

Viktige regler:

* Matchingen er ikke sensitiv for store/små bokstaver
* Hvis flere filer matcher samme kontroll, brukes den nyest endrede filen
* Filer som ikke matcher noen regel listes som ukjente
* Kontroller uten matchende fil blir hoppet over og rapportert som dette

## Standardiserte inputfiler som produseres

01_Verifisering_Input skriver disse filene til `Input/<Selskap>/`:

* `kontroll_solgte_biler_input.xlsx`
* `kontroll_varekjop_bruktbil_input.xlsx`
* `kontroll_demobil_input.xlsx`

Hver fil inneholder to faner:

* `Konfig`
  * kolonner: `Felt`, `Verdi`
* `Biler`
  * kontrollspesifikke kolonner i riktig format

## Mapping og tilpasning

Selv om filvalg skjer automatisk, må kolonnemapping fortsatt vedlikeholdes i Celle 1 i [01_Verifisering_Input](#notebook-1535288192161841).

Der defineres blant annet:

* `source_sheet`
* `identifier_col`
* `regno_col`
* `vin_col`
* `rename_map`
* `required_biler_columns`
* `konfig`

Hvis råfilene endrer struktur eller kolonnenavn, må mappingen oppdateres.

## Widgets

### 01_Verifisering_Input
* `SELSKAP` — dropdown med undermapper funnet i `RåData/`
* `OrgNrSelger` — org.nr. selger, valgfri (brukes til selvhandel-sjekk i solgte biler)
* `OrgNrKjoper` — org.nr. kjøper, påkrevd for varekjøp bruktbil
* `OrgNrDemobil` — org.nr. demobil-eier, påkrevd for demobil
* `DatoFra` — periodens startdato (dd.mm.åååå), påkrevd for varekjøp, valgfri for demobil
* `DatoTil` — periodens sluttdato (dd.mm.åååå), påkrevd for varekjøp, valgfri for demobil

Verdiene skrives automatisk inn i Konfig-fanen i de tre outputfilene.

### 02_OFV_API_Kontroller
* `SELSKAP` — tekst/widget brukt til å peke notebooken til riktig `Input/<Selskap>/` og `Output/<Selskap>/`

## Job: OFV Kontroller Pipeline

[OFV Kontroller Pipeline](#job-1106976622734411) (job ID `1106976622734411`) kjører begge notebookene som én samlet prosess.

Konfigurasjon:

* **Task 1: `Verifisering_Input`**
  * Notebook: `/Workspace/Users/stefan.luidold@bdo.no/OFV API/01_Verifisering_Input`
  * Compute: serverless
* **Task 2: `OFV_API_Kontroller`** (avhenger av Task 1)
  * Notebook: `/Workspace/Users/stefan.luidold@bdo.no/OFV API/02_OFV_API_Kontroller`
  * Compute: serverless
* **Job-parameter:** `SELSKAP` (tom default — fylles inn ved kjøring)
  * Verdien sendes videre til begge notebooks via `{{job.parameters.SELSKAP}}`
* **Schedule:** Ingen — manuell trigger
* **E-postvarsling ved feil:** stefan.luidold@bdo.no

For å kjøre: åpne Joben, klikk "Run now", og fyll inn selskapsnavnet i `SELSKAP`-feltet.

## Secrets og API-nøkler

[02_OFV_API_Kontroller](#notebook-1535288192161840) forventer Databricks secrets i scope:

* `stefan-api-keys`

Nøkler:

* `ofv-api-key` — OFV API (påkrevd)
* `vegvesen-api-key` — Statens vegvesen (valgfri, brukes som fallback for førstegangsregistrering)

## Lokal test

`tests/kjor_notebook_lokalt.py` kjører 02_OFV_API_Kontroller utenfor Databricks med falske OFV/SVV-svar og testdata (inkl. Demobil-tilfellet med tomme datoer). Brukes til å verifisere endringer før notebooken importeres til Databricks.

## Endringslogg

### 04.10.2026 – 02_OFV_API_Kontroller

* **Feilretting Kontroll Demobil:** orgnr i Konfig-fanen ble lest som desimaltall (`933749312.0`) når DatoFra/DatoTil var tomme, så ingen biler matchet OFV sitt `933749312` og alle ble flagget som avvik. Orgnr normaliseres nå (`normalize_orgnr`) i alle tre kontrollene. Tom `OrgNrSelger` gir nå tom selvhandel-kolonne i stedet for «Nei».
* Samme regnr/VIN slås opp bare én gang per kjøring (resultatet er identisk).
* Kontroller uten inputfil hoppes over med tydelig melding.
* Datoer i ISO-format (`aaaa-mm-dd`) som tekst tolkes også.
* Excel: autofilter og frosne overskrifter i kontroll- og resultatfanene; ellers likt innhold.
* Ny, interaktiv HTML-rapport (se over). Demobil viser i tillegg «Eid av enhet» (Ja/Nei) i HTML.
* `jinja2` er ikke lenger nødvendig.

### 04.10.2026 – HTML-rapporten slått sammen med presentasjonen

* Presentasjonen `kontroll_solgte_biler_8.html` (eksempelsidene, Per forhandler, Resultat, Spørsmål og tannhjul-panelet) er bygd inn i HTML-rapporten. Eksempelsidene er kopiert uendret.
* Resultat-siden er slått sammen med Solgte biler; innlimte datasett vises sammen med Databricks-kjøringen.
* Innliming finner nå overskriftsraden selv om den ikke er første rad (f.eks. Excel-arket fra Databricks), og kjenner igjen kolonnene `Selvhandel` og `Førstegangsreg.`.
* Felles BDO-tema for hele filen. Excel-leseren og fontene bygges inn fra `Ressurser/`, så filen virker uten nett.

### 04.10.2026 – Fremdriftsvindu i 02_OFV_API_Kontroller

* Celle 3 viser et levende fremdriftsvindu med gjenstående tid, fremdrift per kontroll og gjeldende OFV/SVV-oppslag. Ny innstilling `FREMDRIFT_VISNING` i Celle 1.

## Viktig ved videre endringer

Denne README-filen skal oppdateres når noe av følgende endres:

* notebooknavn eller notebook-ID-er
* mappestruktur
* widgets eller parametere
* regler for filgjenkjenning
* input-/outputfilnavn
* mappinglogikk
* arbeidsflyt mellom notebookene
* krav til secrets eller API-tilkoblinger

## Vedlikeholdsregel

Når det gjøres endringer i dette OFV API-workspacet, skal README oppdateres i samme arbeidsøkt slik at dokumentasjonen alltid reflekterer faktisk løsning.
