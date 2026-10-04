# Databricks notebook source
# MAGIC %md
# MAGIC # OFV-kontroller - samlet Databricks-notebook
# MAGIC
# MAGIC Python/Databricks-port av `OFV_RefreshInfo.bas`. Kjorer alle tre
# MAGIC kontrollene og samler input og alle resultater i EN Excel-arbeidsbok
# MAGIC (+ en HTML-rapport), i stedet for tre separate VBA-ark i samme
# MAGIC arbeidsbok.
# MAGIC
# MAGIC ## Hva kontrollene gjor
# MAGIC - **Kontroll solgte biler**: for hver bil hentes hele OFV-
# MAGIC   transaksjonshistorikken. Den transaksjonen som ligger **naermest**
# MAGIC   bokfort dato (for eller etter, likegyldig) brukes som
# MAGIC   kontrollgrunnlag. Har OFV ingen transaksjoner for bilen i det hele
# MAGIC   tatt, brukes forstegangsregistreringsdato fra Statens vegvesen
# MAGIC   (SVV) i stedet. Er "OrgNrSelger" fylt ut i input, flagges det
# MAGIC   ekstra ("Selvhandel") hvis samme selskap star som bade selger og
# MAGIC   kjoper i den matchede transaksjonen.
# MAGIC - **Varekjop bruktbil**: tre lister - Innkjop (i perioden), IB (kjopt
# MAGIC   forrige periode) og UB (fortsatt pa lager). For hver bil i Innkjop
# MAGIC   og IB beregnes "lagerperioden" (siste kjop av oppgitt
# MAGIC   juridisk enhet, og forste videresalg etter det kjopet) - OK sa
# MAGIC   lenge lagerperioden overlapper kontrollperioden. I tillegg hentes
# MAGIC   full historikk for biler OFV sier er kjopt av selskapet i perioden,
# MAGIC   men som IKKE star i Innkjop-listen ("mangler i bokforing").
# MAGIC - **Kontroll Demobil**: for hver bil sjekkes om bilens NYESTE
# MAGIC   registrerte OFV-transaksjon fortsatt har oppgitt juridisk enhet som
# MAGIC   kjoper (OK), eller om bilen er videreselgt (Avvik).
# MAGIC
# MAGIC ## Input - tre separate Excel-filer
# MAGIC Hver kontroll har sin EGEN inputfil, med faste filnavn i `INPUT_DIR`
# MAGIC (se Celle 1). Hver fil har to faner:
# MAGIC
# MAGIC - **"Konfig"**: to kolonner, `Felt` og `Verdi` - en rad per
# MAGIC   innstilling (orgnr/datoer).
# MAGIC - **"Biler"**: en rad per bil/kjoretoy, med kolonnen `Identifikator`
# MAGIC   (regnr ELLER VIN - bare en av dem per rad, koden kjenner dem
# MAGIC   automatisk fra hverandre pa lengde: VIN er alltid 17 tegn) pluss en
# MAGIC   kontroll-spesifikk kolonne.
# MAGIC
# MAGIC | Fil | Konfig-felt | Biler-kolonner |
# MAGIC |---|---|---|
# MAGIC | `kontroll_solgte_biler_input.xlsx` | `OrgNrSelger` (valgfri) | `Identifikator`, `BokfortDato` |
# MAGIC | `kontroll_varekjop_bruktbil_input.xlsx` | `OrgNrKjoper`, `DatoFra`, `DatoTil` | `Identifikator`, `Liste` (`Innkjop`/`IB`/`UB`) |
# MAGIC | `kontroll_demobil_input.xlsx` | `OrgNr`, `DatoFra` (valgfri), `DatoTil` (valgfri) | `Identifikator`, `BokfortInnDato` |
# MAGIC
# MAGIC Datoer kan skrives som ekte Excel-datoer eller som tekst `dd.mm.aaaa`.
# MAGIC
# MAGIC ## Output
# MAGIC EN samlet arbeidsbok `OFV_Kontroller_<tidsstempel>.xlsx` i
# MAGIC `OUTPUT_DIR`, med fanene:
# MAGIC
# MAGIC 1. **Input** - oversikt over hvilken kildefil/konfigurasjon/liste med
# MAGIC    biler som ble brukt for hver av de tre kontrollene.
# MAGIC 2. **Resultat solgte biler** / **Kontroll solgte biler**
# MAGIC 3. **Resultat Varekjop** / **Kontroll Varekjop Bruktbil**
# MAGIC 4. **Resultat Demobil** / **Kontroll Demobil**
# MAGIC
# MAGIC Samme innhold skrives ogsa som en selvstendig HTML-rapport
# MAGIC (`OFV_Kontroller_<tidsstempel>.html`).
# MAGIC
# MAGIC Fargene er BDOs egen profilpalett (se Celle 1) i stedet for Excels
# MAGIC innebygde gront/gult/rodt - BDO bla er forbeholdt logoen og brukes
# MAGIC derfor ikke her.
# MAGIC
# MAGIC ## Forbedringer i forhold til VBA-versjonen
# MAGIC - JSON bygges/parses med `json`/`requests` i stedet for en
# MAGIC   handrullet tegn-for-tegn-parser.
# MAGIC - Retry bruker ekte eksponentiell backoff (+ jitter) i stedet for
# MAGIC   VBA sin lineaere "ventetid * forsoksnummer".
# MAGIC - SelgerType/KjoperType regnes direkte i Python (samme regel som
# MAGIC   Excel-formelen i VBA-versjonen), ikke som en live Excel-formel.
# MAGIC - API-nokler hentes fra Databricks secrets, aldri hardkodet.

# COMMAND ----------

# ==================================================================
# CELLE 1: KONFIGURASJON
# ==================================================================

import json
import os
import random
import time
from datetime import date, datetime

import pandas as pd
import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ------------------------------------------------------------------
# Workspace-stier
# ------------------------------------------------------------------

INPUT_DIR = "/Workspace/Users/stefan.luidold@bdo.no/OFV API/Input"
OUTPUT_DIR = "/Workspace/Users/stefan.luidold@bdo.no/OFV API/Output"

INPUT_FILE_SOLGTE_BILER = os.path.join(INPUT_DIR, "kontroll_solgte_biler_input.xlsx")
INPUT_FILE_VAREKJOP = os.path.join(INPUT_DIR, "kontroll_varekjop_bruktbil_input.xlsx")
INPUT_FILE_DEMOBIL = os.path.join(INPUT_DIR, "kontroll_demobil_input.xlsx")

os.makedirs(OUTPUT_DIR, exist_ok=True)

_TIDSSTEMPEL = datetime.now().strftime("%Y%m%d_%H%M%S")
OUTPUT_XLSX = os.path.join(OUTPUT_DIR, f"OFV_Kontroller_{_TIDSSTEMPEL}.xlsx")
OUTPUT_HTML = os.path.join(OUTPUT_DIR, f"OFV_Kontroller_{_TIDSSTEMPEL}.html")

# ------------------------------------------------------------------
# API-endepunkter og retry-oppsett
# ------------------------------------------------------------------

OFV_URL = "https://api.ofv.no/transactions/v1/"
SVV_URL = "https://akfell-datautlevering.atlas.vegvesen.no/enkeltoppslag/kjoretoydata"

MAX_RETRIES = 4
RETRY_BASE_S = 2.0       # eksponentiell backoff: ca 2, 4, 8, 16 sek (+ jitter)
RETRY_JITTER_S = 0.5
API_PAUSE_S = 0.15       # pause mellom hvert kall, for a vaere en god API-borger

RETRYABLE_STATUS_CODES = {429, 500, 502, 503, 504}

# ------------------------------------------------------------------
# Secrets - API-nokler hentes ALDRI hardkodet
# ------------------------------------------------------------------

SECRET_SCOPE = "ofv-api"

try:
    OFV_API_KEY = dbutils.secrets.get(scope=SECRET_SCOPE, key="ofv-api-key")
except Exception as exc:
    raise RuntimeError(
        f"Fant ikke OFV API-nokkel i secret scope '{SECRET_SCOPE}' "
        f"(key 'ofv-api-key'). Opprett secreten med "
        f"'databricks secrets put --scope {SECRET_SCOPE} --key ofv-api-key'. "
        f"({exc})"
    )

try:
    SVV_API_KEY = dbutils.secrets.get(scope=SECRET_SCOPE, key="svv-api-key")
except Exception:
    # SVV er kun en reserve for forstegangsregistrering (Kontroll solgte
    # biler) nar OFV ikke har noen transaksjoner for kjoretoyet - mangler
    # nokkelen, hoppes SVV-oppslaget bare over, akkurat som i VBA.
    SVV_API_KEY = None

# ------------------------------------------------------------------
# BDOs profilpalett (se cloud.brandmaster.com/point/no/bdobc) - BDO
# Bla (#22409A) er forbeholdt logoen og brukes IKKE i rapporten.
# ------------------------------------------------------------------

BDO = {
    "red": "E81A3B",
    "charcoal": "333333",
    "slate": "5B6E7F",
    "burgundy": "98002E",
    "pale_charcoal": "E7E7E7",
    "jade": "009966",
    "ocean": "008FD2",
    "gold": "D67900",
    "white": "FFFFFF",
}

# Semantisk bruk - samme roller som fargekodene i VBA-versjonen
# (morkebla header / lysebla KPI-bakgrunn / gronn-gul-rod status), bare
# med BDOs egne farger i stedet for Excels innebygde palett.
FARGE_HEADER_BG = BDO["charcoal"]
FARGE_HEADER_FONT = BDO["white"]
FARGE_SEKSJON_BG = BDO["pale_charcoal"]
FARGE_SEKSJON_FONT = BDO["charcoal"]
FARGE_OK_BG = BDO["jade"]
FARGE_OK_FONT = BDO["white"]
FARGE_AVVIK_BG = BDO["red"]
FARGE_AVVIK_FONT = BDO["white"]
FARGE_VARSEL_BG = BDO["gold"]
FARGE_VARSEL_FONT = BDO["white"]
FARGE_GRUPPE_BG = "F4F4F4"
FARGE_KANT = BDO["slate"]

# Grenser for fargekoding av "Dager avvik" i Kontroll solgte biler:
# 0-2 dager = OK (Jade), 3-14 dager = varsel (Gold), 15+ dager = avvik (Red).
AVVIK_GRONN_MAX = 2
AVVIK_GUL_MAX = 14

# Felt som regnes som "bil-nivaa" (samme bil i flere transaksjonsrader
# viser disse bare pa den FORSTE/nyeste raden, for lesbarhet i
# Resultat-tabellene) - resten av feltene er transaksjonsspesifikke og
# vises pa hver rad.
CAR_LEVEL_FIELDS = {
    "ChassisNumber", "MakeName", "ModelName", "FuelGroup",
    "IsLeased", "IsUsedImported", "FirstRegistrationDate",
}

# 24-kolonners feltkart for Resultat-tabellene: (nokkel, kolonneoverskrift, er_dato).
FIELD_MAP = [
    ("Input", "Input", False),
    ("Kilde", "Kilde", False),
    ("RegNo", "RegNo", False),
    ("ChassisNumber", "Chassisnummer", False),
    ("MakeName", "Merke", False),
    ("ModelName", "Modell", False),
    ("RegistrationType", "RegistreringsType", False),
    ("FuelGroup", "Drivstoffgruppe", False),
    ("IsLeased", "Leaset", False),
    ("IsUsedImported", "Bruktimportert", False),
    ("FirstRegistrationDate", "ForstegangsRegistrering", True),
    ("TransactionNumber", "TransaksjonsNummer", False),
    ("TransactionDate", "Eierskiftedato", True),
    ("SelgerType", "SelgerType", False),
    ("KjoperType", "KjoperType", False),
    ("FromOwnerType", "SelgerEierType", False),
    ("FromOwnerCompanyName", "SelgerEierFirma", False),
    ("FromOwnerOrgNo", "SelgerOrgNr", False),
    ("FromOwnerCounty", "SelgerEierFylke", False),
    ("ToOwnerType", "KjoperEierType", False),
    ("ToOwnerCompanyName", "KjoperEierFirma", False),
    ("ToOwnerOrgNo", "KjoperOrgNr", False),
    ("ToOwnerCounty", "KjoperEierFylke", False),
    ("Status", "Status", False),
]

# COMMAND ----------

# ==================================================================
# CELLE 2: HJELPEFUNKSJONER
# ==================================================================

# ------------------------------------------------------------------
# Tekst, dato og identifikatorer
# ------------------------------------------------------------------


def normalize_identifier(value) -> str:
    return str(value or "").strip().upper().replace(" ", "")


def is_vin(text: str) -> bool:
    """VIN (chassisnummer) er alltid noyaktig 17 tegn per ISO 3779 - en
    fast, internasjonal standard. Norske regnr kan variere i lengde, sa
    alt som IKKE er 17 tegn regnes som regnr."""
    return len(text) == 17


def split_identifikator(raw) -> tuple:
    """Deler en raw Identifikator-celle inn i (regno, vin) - akkurat en
    av de to er fylt ut, avhengig av lengden."""
    verdi = normalize_identifier(raw)
    if not verdi:
        return "", ""
    if is_vin(verdi):
        return "", verdi
    return verdi, ""


def build_vehicle_key(vin: str, regno: str) -> str:
    if vin:
        return f"VIN|{vin}"
    if regno:
        return f"REG|{regno}"
    return ""


def parse_bokfort_dato(raw):
    """Tolker en dato skrevet inn av en bruker (Excel-celle). Ekte
    datetime/Timestamp/date-objekter (fra pandas/openpyxl som har lest en
    ekte Excel-datocelle) brukes direkte. Tekst tolkes EKSPLISITT som
    dag.maned.ar (norsk standard) - ikke via locale-avhengig parsing."""

    if raw is None:
        return None
    if isinstance(raw, float) and pd.isna(raw):
        return None
    if isinstance(raw, pd.Timestamp):
        return raw.date()
    if isinstance(raw, datetime):
        return raw.date()
    if isinstance(raw, date):
        return raw

    text = str(raw).strip()
    if not text:
        return None

    text = text.replace("/", ".").replace("-", ".")
    parts = text.split(".")
    if len(parts) != 3:
        return None

    try:
        dag, maned, ar = int(parts[0]), int(parts[1]), int(parts[2])
    except ValueError:
        return None

    if ar < 100:
        ar += 2000

    if not (1 <= maned <= 12 and 1 <= dag <= 31 and 1900 <= ar <= 2100):
        return None

    try:
        return date(ar, maned, dag)
    except ValueError:
        return None


def date_from_iso(iso_value):
    """Parser OFV/SVV sine ISO-formatterte datoer via faste
    tegnposisjoner - locale-uavhengig."""
    if not iso_value:
        return None
    s = str(iso_value)
    if len(s) < 10:
        return None
    try:
        return date(int(s[0:4]), int(s[5:7]), int(s[8:10]))
    except ValueError:
        return None


def find_any(obj, key):
    """Sok rekursivt etter forste forekomst av `key` hvor som helst i et
    nestet dict/list-tre. Brukes for SVV-responsen, der nestingsniva er
    ukjent/kan variere - samme forgivende "finn dette navnet uansett hvor
    det ligger"-prinsipp som den opprinnelige VBA JSON-parseren brukte."""

    if isinstance(obj, dict):
        if key in obj and obj[key] not in (None, ""):
            return obj[key]
        for v in obj.values():
            found = find_any(v, key)
            if found is not None:
                return found
    elif isinstance(obj, list):
        for item in obj:
            found = find_any(item, key)
            if found is not None:
                return found
    return None


def compute_owner_label(owner_type, company_name) -> str:
    """Privat vinner over firmanavn, ellers firmanavn hvis det finnes,
    ellers den ra eiertype-teksten."""
    owner_type = owner_type or ""
    company_name = company_name or ""
    if not owner_type and not company_name:
        return ""
    if owner_type == "Privat":
        return "Privat"
    return company_name or owner_type


# ------------------------------------------------------------------
# HTTP - OFV og SVV, med eksponentiell backoff
# ------------------------------------------------------------------

_SESSION = requests.Session()


def _backoff_sleep(attempt: int) -> None:
    wait = RETRY_BASE_S * (2 ** (attempt - 1)) + random.uniform(0, RETRY_JITTER_S)
    time.sleep(wait)


def post_ofv_with_retries(api_key: str, body: dict):
    """POST mot OFV med eksponentiell backoff-retry pa 429/5xx. 401/403/404
    er ikke-forbigaende feil og gir umiddelbart opp."""

    last_error = ""

    for attempt in range(1, MAX_RETRIES + 1):

        try:
            resp = _SESSION.post(
                OFV_URL,
                headers={
                    "Content-Type": "application/json",
                    "Cache-Control": "no-cache",
                    "Ocp-Apim-Subscription-Key": api_key,
                },
                json=body,
                timeout=30,
            )
        except requests.RequestException as exc:
            last_error = str(exc)
            _backoff_sleep(attempt)
            continue

        if resp.status_code == 200:
            return resp.json(), "OK"
        if resp.status_code == 401:
            return None, "Feil: OFV 401 - kontroller API-nokkelen"
        if resp.status_code == 403:
            return None, "Feil: OFV 403 - tilgang eller kvote"
        if resp.status_code == 404:
            return None, "Feil: OFV-endepunktet ble ikke funnet"
        if resp.status_code in RETRYABLE_STATUS_CODES:
            last_error = f"{resp.status_code}: {resp.text[:200]}"
            _backoff_sleep(attempt)
            continue

        return None, f"Feil: OFV HTTP {resp.status_code}"

    return None, f"Feil: OFV - {last_error}"


def get_svv_response(api_key: str, filter_name: str, identifier: str):

    last_error = ""

    for attempt in range(1, MAX_RETRIES + 1):

        try:
            resp = _SESSION.get(
                SVV_URL,
                params={filter_name: identifier},
                headers={
                    "SVV-Authorization": f"Apikey {api_key}",
                    "Accept": "application/json",
                },
                timeout=30,
            )
        except requests.RequestException as exc:
            last_error = str(exc)
            _backoff_sleep(attempt)
            continue

        if resp.status_code == 200:
            return resp.json(), "Statens vegvesen"
        if resp.status_code == 401:
            return None, "Feil: Vegvesenet 401 - kontroller API-nokkel"
        if resp.status_code == 403:
            return None, "Feil: Vegvesenet 403 - tilgang eller kvote"
        if resp.status_code == 404:
            return None, "Statens vegvesen - ingen treff"
        if resp.status_code in RETRYABLE_STATUS_CODES:
            last_error = f"{resp.status_code}: {resp.text[:200]}"
            _backoff_sleep(attempt)
            continue

        return None, f"Feil: Vegvesenet HTTP {resp.status_code}"

    return None, f"Feil: Vegvesenet - {last_error}"


def fetch_vehicle_info_from_svv(api_key: str, regno: str, vin: str) -> dict:
    """Kalles kun nar OFV ikke har noen transaksjoner for kjoretoyet.
    Prover regnr forst, deretter VIN hvis regnr ikke gir treff."""

    result = {"Status": "Statens vegvesen - ingen dato"}

    data = None
    status = None

    if regno:
        data, status = get_svv_response(api_key, "kjennemerke", regno)
    if data is None and vin:
        data, status = get_svv_response(api_key, "understellsnummer", vin)

    if data is None:
        result["Status"] = status or "Statens vegvesen - ingen dato"
        return result

    iso_date = (
        find_any(data, "registrertForstegangNorgeDato")
        or find_any(data, "registrertForstegangDato")
    )

    d = date_from_iso(iso_date) if iso_date else None
    if d:
        result["FirstRegistrationDate"] = d
        result["Status"] = "Statens vegvesen"

    return result


# ------------------------------------------------------------------
# OFV-transaksjoner - bygging av rader + de to API-kallene
# ------------------------------------------------------------------


def build_transaction_row(identifier, tx: dict, use_vin: bool,
                           original_regno: str, original_vin: str) -> dict:

    from_side = tx.get("from") or {}
    to_side = tx.get("to") or {}
    from_owner = from_side.get("owner") or {}
    to_owner = to_side.get("owner") or {}
    from_company = from_owner.get("companyInfo") or {}
    to_company = to_owner.get("companyInfo") or {}

    from_type = from_owner.get("type") or ""
    from_company_name = from_company.get("name") or ""
    to_type = to_owner.get("type") or ""
    to_company_name = to_company.get("name") or ""

    return {
        "Input": identifier,
        "Kilde": "VIN" if use_vin else "Regnr",
        "RegNo": tx.get("regNo") or original_regno,
        "ChassisNumber": tx.get("chassisNumber") or original_vin,
        "MakeName": tx.get("makeName"),
        "ModelName": tx.get("modelName"),
        "RegistrationType": tx.get("registrationType"),
        "FuelGroup": tx.get("fuelGroup"),
        "IsLeased": tx.get("isLeased"),
        "IsUsedImported": tx.get("isUsedImported"),
        "FirstRegistrationDate": date_from_iso(tx.get("firstRegistrationDate")),
        "TransactionNumber": tx.get("transactionNumber"),
        "TransactionDate": date_from_iso(tx.get("transactionDate")),
        "SelgerType": compute_owner_label(from_type, from_company_name),
        "KjoperType": compute_owner_label(to_type, to_company_name),
        "FromOwnerType": from_type,
        "FromOwnerCompanyName": from_company_name,
        "FromOwnerOrgNo": from_company.get("organizationNumber"),
        "FromOwnerCounty": from_owner.get("countyName"),
        "ToOwnerType": to_type,
        "ToOwnerCompanyName": to_company_name,
        "ToOwnerOrgNo": to_company.get("organizationNumber"),
        "ToOwnerCounty": to_owner.get("countyName"),
        "Status": "OK",
    }


def build_empty_transaction_row(identifier, use_vin: bool,
                                 original_regno: str, original_vin: str,
                                 status_text: str) -> dict:
    return {
        "Input": identifier,
        "Kilde": "VIN" if use_vin else "Regnr",
        "RegNo": original_regno,
        "ChassisNumber": original_vin,
        "Status": status_text,
    }


def fetch_ofv_transactions(api_key: str, identifier: str, use_vin: bool,
                            original_regno: str, original_vin: str) -> list:
    """Henter ALLE transaksjoner for kjoretoyet - ett kall per kjoretoy,
    uten datofilter, sortert nyeste forst (cursor-paginert, 1000 per side)."""

    filter_key = "chassisNumber" if use_vin else "regNo"
    rows = []
    cursor = None

    while True:

        pagination = {"first": 1000}
        if cursor:
            pagination["cursor"] = cursor

        body = {
            "filters": {filter_key: identifier},
            "pagination": pagination,
            "sorting": {"orderBy": "transactionDate", "orderDirection": "DESC"},
        }

        data, status = post_ofv_with_retries(api_key, body)

        if status != "OK":
            rows.append(build_empty_transaction_row(
                identifier, use_vin, original_regno, original_vin, status))
            return rows

        for tx in (data or {}).get("transactions", []):
            rows.append(build_transaction_row(
                identifier, tx, use_vin, original_regno, original_vin))

        pg = (data or {}).get("pagination", {})

        if pg.get("hasNextPage"):
            cursor = pg.get("endCursor")
            if not cursor:
                break
            time.sleep(API_PAUSE_S)
        else:
            break

    if not rows:
        rows.append(build_empty_transaction_row(
            identifier, use_vin, original_regno, original_vin,
            "Ingen eierskifter funnet"))

    return rows


def fetch_ofv_transactions_by_buyer_org(api_key: str, org_no: str,
                                         date_fra: date, date_til: date) -> list:
    """Henter ALLE OFV-transaksjoner der gitt orgnr star som KJOPER
    (toOrganizationNumber), innenfor et datointervall - ett samlekall
    uavhengig av regnr/VIN."""

    rows = []
    cursor = None

    while True:

        pagination = {"first": 1000}
        if cursor:
            pagination["cursor"] = cursor

        body = {
            "filters": {
                "toOrganizationNumber": org_no,
                "transactionDateFrom": date_fra.strftime("%Y-%m-%d"),
                "transactionDateTo": date_til.strftime("%Y-%m-%d"),
            },
            "pagination": pagination,
            "sorting": {"orderBy": "transactionDate", "orderDirection": "DESC"},
        }

        data, status = post_ofv_with_retries(api_key, body)

        if status != "OK":
            raise RuntimeError(f"FetchOFVTransactionsByBuyerOrg: {status}")

        for tx in (data or {}).get("transactions", []):
            rows.append(build_transaction_row("", tx, False, "", ""))

        pg = (data or {}).get("pagination", {})

        if pg.get("hasNextPage"):
            cursor = pg.get("endCursor")
            if not cursor:
                break
            time.sleep(API_PAUSE_S)
        else:
            break

    return rows


# ------------------------------------------------------------------
# Input-lesere - "Konfig" (Felt/Verdi) og "Biler" (en rad per kjoretoy)
# ------------------------------------------------------------------


def read_konfig(path: str) -> dict:
    """Leser fanen 'Konfig' (kolonner Felt/Verdi) til en dict med
    streng-trimmede nokler. Verdier returneres rait (kan vaere en
    pandas.Timestamp for datofelt - tolkes av parse_bokfort_dato)."""

    df = pd.read_excel(path, sheet_name="Konfig")
    konfig = {}
    for _, row in df.iterrows():
        felt = str(row.get("Felt", "")).strip()
        if felt and felt.lower() != "nan":
            konfig[felt] = row.get("Verdi")
    return konfig


def read_biler(path: str, extra_col: str) -> list:
    """Leser fanen 'Biler' (kolonner Identifikator + en kontroll-
    spesifikk ekstrakolonne) til en liste med dicts
    {regno, vin, identifikator, <extra_col>}."""

    df = pd.read_excel(path, sheet_name="Biler")
    out = []

    for _, row in df.iterrows():

        raw_ident = row.get("Identifikator")
        if raw_ident is None or (isinstance(raw_ident, float) and pd.isna(raw_ident)):
            continue

        regno, vin = split_identifikator(raw_ident)
        if not regno and not vin:
            continue

        out.append({
            "regno": regno,
            "vin": vin,
            "identifikator": vin or regno,
            extra_col: row.get(extra_col) if extra_col in df.columns else None,
        })

    return out


# ------------------------------------------------------------------
# Resultat-tabell - delt av alle tre kontroller
# ------------------------------------------------------------------


def build_resultat_rows(all_rows: list) -> list:
    """Full transaksjonshistorikk, gruppert pr bil (samme 'Input'-verdi
    pa flere rader i rad). Kun den FORSTE raden i hver gruppe viser
    bil-nivaa-felt (CAR_LEVEL_FIELDS) - resten viser bare
    transaksjonsspesifikk info, for lesbarhet."""

    records = []
    forrige_input = None

    for row in all_rows:

        denne_input = row.get("Input")
        er_forste = (denne_input != forrige_input)
        forrige_input = denne_input

        record = {"_ErForsteIGruppe": er_forste, "_Input": denne_input}

        for key, header, _is_date in FIELD_MAP:

            if key in ("SelgerType", "KjoperType"):
                record[header] = row.get(key)
            elif not er_forste and key in CAR_LEVEL_FIELDS:
                record[header] = None
            else:
                record[header] = row.get(key)

        records.append(record)

    return records


# ------------------------------------------------------------------
# Excel-stiler (BDO-palett) - brukt av output-cellen
# ------------------------------------------------------------------

_THIN_BORDER_SIDE = Side(style="thin", color=FARGE_KANT)
THIN_BORDER = Border(
    left=_THIN_BORDER_SIDE, right=_THIN_BORDER_SIDE,
    top=_THIN_BORDER_SIDE, bottom=_THIN_BORDER_SIDE,
)


def set_title_bar(ws, row: int, col_span: int, text: str) -> None:
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=col_span)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = Font(bold=True, size=14, color=FARGE_HEADER_FONT)
    cell.fill = PatternFill("solid", fgColor=FARGE_HEADER_BG)
    cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[row].height = 26


def set_explanation_row(ws, row: int, col_span: int, text: str) -> None:
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=col_span)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = Font(italic=True)
    cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[row].height = 30


def set_section_bar(ws, row: int, col_span: int, text: str) -> None:
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=col_span)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = Font(bold=True, size=12, color=FARGE_SEKSJON_FONT)
    cell.fill = PatternFill("solid", fgColor=FARGE_SEKSJON_BG)
    cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[row].height = 20


def set_kpi_box(ws, row: int, col_from: int, col_to: int, label: str,
                 value, bg_hex: str = None, font_hex: str = None,
                 big: bool = False) -> None:
    """Skriver en KPI-boks over to rader: `row` = etikett, `row+1` =
    tallverdi (eller omvendt, se bruk). Slar sammen col_from:col_to."""

    ws.merge_cells(start_row=row, start_column=col_from, end_row=row, end_column=col_to)
    cell = ws.cell(row=row, column=col_from, value=label)
    cell.font = Font(bold=True)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    if bg_hex:
        cell.fill = PatternFill("solid", fgColor=bg_hex)
        cell.font = Font(bold=True, color=font_hex or "000000")

    ws.merge_cells(start_row=row + 1, start_column=col_from, end_row=row + 1, end_column=col_to)
    vcell = ws.cell(row=row + 1, column=col_from, value=value)
    vcell.font = Font(bold=True, size=16 if big else 12, color=font_hex or "000000")
    vcell.alignment = Alignment(horizontal="center", vertical="center")
    if bg_hex:
        vcell.fill = PatternFill("solid", fgColor=bg_hex)


def style_table_header(ws, row: int, headers: list) -> None:
    for i, header in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=i, value=header)
        cell.font = Font(bold=True, color=FARGE_HEADER_FONT)
        cell.fill = PatternFill("solid", fgColor=FARGE_HEADER_BG)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 30


def color_status_cell(cell, status_text: str) -> None:
    status_text = status_text or ""
    if status_text.startswith("OK"):
        cell.fill = PatternFill("solid", fgColor=FARGE_OK_BG)
        cell.font = Font(color=FARGE_OK_FONT, bold=True)
    elif status_text.startswith("Avvik"):
        cell.fill = PatternFill("solid", fgColor=FARGE_AVVIK_BG)
        cell.font = Font(color=FARGE_AVVIK_FONT, bold=True)


def set_column_widths(ws, widths: dict) -> None:
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width


def set_date_format(ws, col_letter: str, first_row: int, last_row: int) -> None:
    if last_row < first_row:
        return
    for row in range(first_row, last_row + 1):
        ws[f"{col_letter}{row}"].number_format = "dd.mm.yyyy"

# COMMAND ----------

# ==================================================================
# CELLE 3: HOVEDLOGIKK
# ==================================================================

# ------------------------------------------------------------------
# Kontroll 1: Kontroll solgte biler
# ------------------------------------------------------------------


def build_kontroll_row(regno: str, vin: str, bokfort_raw, tx_rows: list,
                        svv_info, seller_org_no: str) -> dict:
    """Bokfort dato sjekkes mot HELE bilens OFV-transaksjonshistorikk -
    den transaksjonen som ligger NAERMEST bokfort dato brukes, uansett om
    den er for eller etter. Har OFV ingen transaksjoner i det hele tatt,
    brukes forstegangsregistreringsdato fra SVV i stedet - bade som
    kontrollgrunnlag og i kolonnen Forstegangsregistrert.

    Valgfri tilleggskontroll (kun nar seller_org_no er fylt ut): er OGSA
    kjoperen samme org som selgerOrgNo (samme selskap pa begge sider),
    flagges Selvhandel = "Ja" - raden tvinges da rod uansett dagers
    avvik nar den skrives til Excel/HTML."""

    bokfort_date = parse_bokfort_dato(bokfort_raw)
    seller_norm = normalize_identifier(seller_org_no)

    model_name = ""
    chassis_no = vin
    regno_resolved = regno
    first_reg_date = None
    has_any_ok = False
    error_status = ""

    nearest_row = None
    nearest_diff = None

    for r in tx_rows:

        status = r.get("Status")

        if status == "OK":

            has_any_ok = True

            if not model_name:
                model_name = r.get("ModelName") or ""
            if r.get("ChassisNumber"):
                chassis_no = r["ChassisNumber"]
            if r.get("RegNo"):
                regno_resolved = r["RegNo"]
            if first_reg_date is None and r.get("FirstRegistrationDate"):
                first_reg_date = r["FirstRegistrationDate"]

            tx_date = r.get("TransactionDate")
            if tx_date and bokfort_date:
                diff = abs((tx_date - bokfort_date).days)
                if nearest_diff is None or diff < nearest_diff:
                    nearest_diff = diff
                    nearest_row = r

        elif isinstance(status, str) and status.startswith("Feil:"):
            if not error_status:
                error_status = status

    if not has_any_ok and svv_info and svv_info.get("FirstRegistrationDate"):
        first_reg_date = svv_info["FirstRegistrationDate"]

    result = {
        "RegnrInput": regno_resolved,
        "Chassisnummer": chassis_no,
        "Modell": model_name,
        "BokfortDato": bokfort_date,
        "Forstegangsregistrert": first_reg_date,
        "KontrollTransaksjonDato": None,
        "RegistreringsType": "",
        "DagerAvvik": None,
        "Selger": "",
        "Kjoper": "",
        "Kilde": "Ingen",
        "Selvhandel": "",
    }

    if bokfort_date is None:

        if has_any_ok or first_reg_date is not None:
            result["ApiTreff"] = "Mangler bokfort dato"
        elif error_status:
            result["ApiTreff"] = error_status
        else:
            result["ApiTreff"] = "Ingen treff"

        result["Kontrollert"] = "Nei"

    elif nearest_row is not None:

        result["KontrollTransaksjonDato"] = nearest_row["TransactionDate"]
        result["RegistreringsType"] = nearest_row.get("RegistrationType") or ""
        result["DagerAvvik"] = nearest_diff
        result["ApiTreff"] = "Treff OFV eierskifte"
        result["Kontrollert"] = "Ja"
        result["Kilde"] = "OFV"

        result["Selger"] = compute_owner_label(
            nearest_row.get("FromOwnerType"), nearest_row.get("FromOwnerCompanyName"))
        result["Kjoper"] = compute_owner_label(
            nearest_row.get("ToOwnerType"), nearest_row.get("ToOwnerCompanyName"))

        if seller_norm:
            fra_org = normalize_identifier(nearest_row.get("FromOwnerOrgNo"))
            til_org = normalize_identifier(nearest_row.get("ToOwnerOrgNo"))
            result["Selvhandel"] = "Ja" if (
                fra_org == seller_norm and til_org == seller_norm) else "Nei"

    elif first_reg_date is not None:

        result["KontrollTransaksjonDato"] = first_reg_date
        result["RegistreringsType"] = "Forstegangsregistrering (SVV)"
        result["DagerAvvik"] = abs((first_reg_date - bokfort_date).days)
        result["ApiTreff"] = "Treff SVV forstegangsregistrering"
        result["Kontrollert"] = "Ja"
        result["Kilde"] = "SVV"

    else:

        if svv_info is None:
            result["ApiTreff"] = error_status or "Ingen treff"
        else:
            result["ApiTreff"] = svv_info.get("Status", "Ingen treff")

        result["Kontrollert"] = "Nei"

    return result


def run_kontroll_solgte_biler() -> dict:

    print("=== Kontroll solgte biler ===")

    konfig = read_konfig(INPUT_FILE_SOLGTE_BILER)
    seller_org_no = str(konfig.get("OrgNrSelger") or "").strip()
    biler = read_biler(INPUT_FILE_SOLGTE_BILER, "BokfortDato")

    all_rows = []
    kontroll_rows = []
    total = len(biler)
    svv_treff = 0
    svv_feil = 0
    ofv_feil = 0
    ofv_uten_treff = 0

    for i, bil in enumerate(biler, start=1):

        regno, vin = bil["regno"], bil["vin"]
        ident = bil["identifikator"]
        use_vin = bool(vin)

        print(f"  OFV: Eierskifter {ident}  ({i}/{total})")

        tx_rows = fetch_ofv_transactions(OFV_API_KEY, ident, use_vin, regno, vin)
        all_rows.extend(tx_rows)

        has_any_ok = any(r.get("Status") == "OK" for r in tx_rows)
        if not has_any_ok:
            if any(isinstance(r.get("Status"), str) and r["Status"].startswith("Feil:")
                   for r in tx_rows):
                ofv_feil += 1
            else:
                ofv_uten_treff += 1

        svv_info = None
        if not has_any_ok and SVV_API_KEY:
            print(f"  SVV: Forstegangsregistrering {ident}  ({i}/{total})")
            svv_info = fetch_vehicle_info_from_svv(SVV_API_KEY, regno, vin)
            if svv_info.get("FirstRegistrationDate"):
                svv_treff += 1
            if isinstance(svv_info.get("Status"), str) and svv_info["Status"].startswith("Feil:"):
                svv_feil += 1

        kontroll_rows.append(
            build_kontroll_row(regno, vin, bil["BokfortDato"], tx_rows, svv_info, seller_org_no))

        time.sleep(API_PAUSE_S)

    kontroll_rows_sortert = sorted(
        kontroll_rows,
        key=lambda r: r["DagerAvvik"] if isinstance(r["DagerAvvik"], (int, float)) else -1,
        reverse=True,
    )

    return {
        "kontroll": "Kontroll solgte biler",
        "kildefil": INPUT_FILE_SOLGTE_BILER,
        "orgnr": seller_org_no,
        "dato_fra": None,
        "dato_til": None,
        "antall_biler": total,
        "biler_manifest": biler,
        "all_rows": all_rows,
        "kontroll_rows": kontroll_rows_sortert,
        "stats": {
            "Antall biler": total,
            "OFV-treff": sum(1 for r in kontroll_rows if r["Kilde"] == "OFV"),
            "Uten OFV-treff": ofv_uten_treff,
            "OFV-feil": ofv_feil,
            "SVV-forstegangsregistreringer": svv_treff,
            "SVV-feil": svv_feil,
        },
    }


# ------------------------------------------------------------------
# Kontroll 2: Varekjop bruktbil - delt lagerperiode-motor
# ------------------------------------------------------------------


def finn_lagerperiode(tx_rows: list, buyer_org_no: str, kjop_grense: date):
    """Delt motor for "var bilen pa lager i perioden": finner Fra-
    transaksjonen (siste kjop av oppgitt forhandler PA ELLER FOR
    kjop_grense) og Til-transaksjonen (forste salg FRA forhandleren
    ETTER det kjopet, uansett nar det falt). Brukes av bade Seksjon A
    (Innkjop - kjop_grense = Dato til) og Seksjon B (IB - kjop_grense =
    Dato fra minus en dag, siden en IB-bil per definisjon ble kjopt FOR
    perioden startet).

    Returnerer (kjopt_row, solgt_row, chassis_no, model_name, regno)."""

    buyer_norm = normalize_identifier(buyer_org_no)

    chassis_no = ""
    model_name = ""
    regno_resolved = ""
    kjopt_row = None
    solgt_row = None

    for r in tx_rows:

        if r.get("Status") != "OK":
            continue

        if not model_name:
            model_name = r.get("ModelName") or ""
        if r.get("ChassisNumber"):
            chassis_no = r["ChassisNumber"]
        if r.get("RegNo"):
            regno_resolved = r["RegNo"]

        tx_date = r.get("TransactionDate")
        if not tx_date:
            continue

        if (normalize_identifier(r.get("ToOwnerOrgNo")) == buyer_norm
                and tx_date <= kjop_grense):
            if kjopt_row is None or tx_date > kjopt_row["TransactionDate"]:
                kjopt_row = r

    if kjopt_row is not None:

        for r in tx_rows:

            if r.get("Status") != "OK":
                continue

            tx_date = r.get("TransactionDate")
            if not tx_date:
                continue

            if (normalize_identifier(r.get("FromOwnerOrgNo")) == buyer_norm
                    and tx_date > kjopt_row["TransactionDate"]):
                if solgt_row is None or tx_date < solgt_row["TransactionDate"]:
                    solgt_row = r

    return kjopt_row, solgt_row, chassis_no, model_name, regno_resolved


def lager_status_tekst(kjopt_row, solgt_row, dato_fra: date) -> str:
    """OK sa lenge bilens lagerperiode overlapper kontrollperioden (solgt
    pa/etter Dato fra, eller fortsatt ikke solgt). Avvik nar bilen var
    solgt FOR perioden startet, eller nar det ikke finnes noe opprinnelig
    kjop a sjekke mot i det hele tatt."""

    if kjopt_row is None:
        return "Avvik: fant ikke opprinnelig kjop hos forhandleren"
    if solgt_row is None:
        return "OK - bilen var pa lager i perioden"
    if solgt_row["TransactionDate"] >= dato_fra:
        return "OK - bilen var pa lager i perioden"
    return "Avvik: solgt for perioden startet"


def build_innkjop_eller_ib_row(regno: str, vin: str, tx_rows: list, buyer_org_no: str,
                                dato_fra: date, dato_til: date, kjop_grense: date,
                                flagg_etikett: str, flagg_verdi: str) -> dict:
    """Delt radbygger for Seksjon A (Innkjop) og Seksjon B (IB) - de to
    bruker nokyaktig samme felt-sett og samme motor, bare med forskjellig
    kjop_grense (se finn_lagerperiode)."""

    kjopt_row, solgt_row, chassis_no, model_name, regno_resolved = finn_lagerperiode(
        tx_rows, buyer_org_no, kjop_grense)

    result = {
        "RegnrInput": regno_resolved or regno,
        "Chassisnummer": chassis_no or vin,
        "Modell": model_name,
        "FlaggEtikett": flagg_etikett,
        "FlaggVerdi": flagg_verdi,
        "FraDato": kjopt_row["TransactionDate"] if kjopt_row else None,
        "KjoptFra": compute_owner_label(
            kjopt_row.get("FromOwnerType"), kjopt_row.get("FromOwnerCompanyName")
        ) if kjopt_row else "",
        "TilDato": solgt_row["TransactionDate"] if solgt_row else None,
        "Kjoper": compute_owner_label(
            solgt_row.get("ToOwnerType"), solgt_row.get("ToOwnerCompanyName")
        ) if solgt_row else "",
    }

    result["Status"] = lager_status_tekst(kjopt_row, solgt_row, dato_fra)

    return result


def build_manglende_bokforing_row(ofv_row: dict, tx_rows: list, buyer_org_no: str,
                                   dato_til: date) -> dict:
    """Seksjon A, tillegg - en bil OFV sier er kjopt av selskapet i
    perioden, men som IKKE star i Innkjop-listen (fullstendighet pa
    bokforingen). tx_rows er full historikk for bilen (hentet av
    kalleren pa samme mate som Innkjop/IB-bilene), sa Fra/Til dato
    finnes med samme finn_lagerperiode-motor - ofv_row brukes bare som
    sikkert fallback om historikk-kallet av noen grunn ikke gir treff."""

    kjopt_row, solgt_row, chassis_no, model_name, regno_resolved = finn_lagerperiode(
        tx_rows, buyer_org_no, dato_til)

    if kjopt_row is None:
        kjopt_row = ofv_row

    result = {
        "RegnrInput": regno_resolved or ofv_row.get("RegNo") or "",
        "Chassisnummer": chassis_no or ofv_row.get("ChassisNumber") or "",
        "Modell": model_name or ofv_row.get("ModelName") or "",
        "FlaggEtikett": "I OFV-liste",
        "FlaggVerdi": "Ja",
        "FraDato": kjopt_row.get("TransactionDate"),
        "KjoptFra": compute_owner_label(
            kjopt_row.get("FromOwnerType"), kjopt_row.get("FromOwnerCompanyName")),
        "TilDato": solgt_row["TransactionDate"] if solgt_row else None,
        "Kjoper": compute_owner_label(
            solgt_row.get("ToOwnerType"), solgt_row.get("ToOwnerCompanyName")
        ) if solgt_row else "",
        "Status": "Avvik: OFV viser kjop, mangler i bokforing",
    }

    return result


def build_ub_row(regno: str, vin: str, tx_rows: list, er_i_ib: bool,
                  er_i_innkjop: bool) -> dict:
    """Seksjon C - UB (fortsatt pa lager ved kontrolltidspunkt): finnes
    bilen i IB eller Innkjop-listen (forklart), eller er den uten kjent
    opprinnelse (uforklart)?"""

    chassis_no, model_name, regno_resolved = vin, "", regno

    for r in tx_rows:
        if r.get("Status") != "OK":
            continue
        if not model_name:
            model_name = r.get("ModelName") or ""
        if r.get("ChassisNumber"):
            chassis_no = r["ChassisNumber"]
        if r.get("RegNo"):
            regno_resolved = r["RegNo"]

    return {
        "RegnrInput": regno_resolved,
        "Chassisnummer": chassis_no,
        "Modell": model_name,
        "FunnetIIB": "Ja" if er_i_ib else "Nei",
        "FunnetIInnkjop": "Ja" if er_i_innkjop else "Nei",
        "Status": "OK - forklart" if (er_i_ib or er_i_innkjop) else "Avvik: uforklart lagerbeholdning",
    }


def run_kontroll_varekjop_bruktbil() -> dict:

    print("=== Varekjop bruktbil ===")

    konfig = read_konfig(INPUT_FILE_VAREKJOP)
    buyer_org_no = str(konfig.get("OrgNrKjoper") or "").strip()
    dato_fra = parse_bokfort_dato(konfig.get("DatoFra"))
    dato_til = parse_bokfort_dato(konfig.get("DatoTil"))

    if not buyer_org_no or not dato_fra or not dato_til:
        raise ValueError(
            "Varekjop bruktbil: OrgNrKjoper/DatoFra/DatoTil ma vaere fylt ut "
            "i Konfig-fanen."
        )

    biler = read_biler(INPUT_FILE_VAREKJOP, "Liste")

    queue_innkjop, queue_ib, queue_ub = {}, {}, {}

    for bil in biler:
        key = build_vehicle_key(bil["vin"], bil["regno"])
        if not key:
            continue
        liste = str(bil.get("Liste") or "").strip().lower()
        if liste == "innkjop":
            queue_innkjop[key] = bil
        elif liste == "ib":
            queue_ib[key] = bil
        elif liste == "ub":
            queue_ub[key] = bil

    print(f"  Innkjop: {len(queue_innkjop)}, IB: {len(queue_ib)}, UB: {len(queue_ub)}")

    # Steg 1: alle biler OFV sier er kjopt av orgnr i perioden.
    print(f"  OFV: henter kjopsliste for org {buyer_org_no} ({dato_fra} - {dato_til}) ...")
    ofv_liste = fetch_ofv_transactions_by_buyer_org(OFV_API_KEY, buyer_org_no, dato_fra, dato_til)
    print(f"  OFV: {len(ofv_liste)} transaksjoner funnet i perioden.")

    ofv_regnr = {normalize_identifier(r["RegNo"]) for r in ofv_liste if r.get("RegNo")}
    ofv_vin = {normalize_identifier(r["ChassisNumber"]) for r in ofv_liste if r.get("ChassisNumber")}

    # Steg 2: full transaksjonshistorikk for hver bil i Innkjop + IB.
    queue_api = {**queue_innkjop, **queue_ib}
    vehicle_rows_by_key = {}
    all_rows = []
    total = len(queue_api)

    for i, (key, bil) in enumerate(queue_api.items(), start=1):
        regno, vin = bil["regno"], bil["vin"]
        ident = bil["identifikator"]
        print(f"  OFV: Eierskifter {ident}  ({i}/{total})")
        tx_rows = fetch_ofv_transactions(OFV_API_KEY, ident, bool(vin), regno, vin)
        vehicle_rows_by_key[key] = tx_rows
        all_rows.extend(tx_rows)
        time.sleep(API_PAUSE_S)

    # Seksjon A: Innkjop i perioden - samme overlapp-logikk som Seksjon B.
    seksjon_a = []

    for key, bil in queue_innkjop.items():
        regno, vin = bil["regno"], bil["vin"]
        tx_rows = vehicle_rows_by_key.get(key, [])
        er_i_ofv_liste = (
            (regno and normalize_identifier(regno) in ofv_regnr)
            or (vin and normalize_identifier(vin) in ofv_vin)
        )
        seksjon_a.append(build_innkjop_eller_ib_row(
            regno, vin, tx_rows, buyer_org_no, dato_fra, dato_til, dato_til,
            "I OFV-liste", "Ja" if er_i_ofv_liste else "Nei",
        ))

    # Biler OFV sier er kjopt i perioden, men IKKE i Innkjop-listen -
    # hent full historikk for disse ogsa (bade for lagerperiode-
    # beregningen her og for Resultat Varekjop).
    for ofv_row in ofv_liste:

        reg_n = normalize_identifier(ofv_row.get("RegNo"))
        vin_n = normalize_identifier(ofv_row.get("ChassisNumber"))

        funnet = any(
            (reg_n and reg_n == normalize_identifier(b["regno"]))
            or (vin_n and vin_n == normalize_identifier(b["vin"]))
            for b in queue_innkjop.values()
        )
        if funnet:
            continue

        key = build_vehicle_key(vin_n, reg_n)
        tx_rows = vehicle_rows_by_key.get(key)

        if tx_rows is None:
            ident = vin_n or reg_n
            print(f"  OFV: Eierskifter (mangler i bokforing) {ident}")
            tx_rows = fetch_ofv_transactions(OFV_API_KEY, ident, bool(vin_n), reg_n, vin_n)
            vehicle_rows_by_key[key] = tx_rows
            all_rows.extend(tx_rows)
            time.sleep(API_PAUSE_S)

        seksjon_a.append(build_manglende_bokforing_row(ofv_row, tx_rows, buyer_org_no, dato_til))

    # Seksjon B: IB (kjopt forrige periode) - kjop_grense = Dato fra - 1 dag.
    seksjon_b = []
    kjop_grense_ib = dato_fra - pd.Timedelta(days=1)
    kjop_grense_ib = date(kjop_grense_ib.year, kjop_grense_ib.month, kjop_grense_ib.day)

    for key, bil in queue_ib.items():
        regno, vin = bil["regno"], bil["vin"]
        tx_rows = vehicle_rows_by_key.get(key, [])
        er_pa_ub = key in queue_ub
        seksjon_b.append(build_innkjop_eller_ib_row(
            regno, vin, tx_rows, buyer_org_no, dato_fra, dato_til, kjop_grense_ib,
            "Pa UB-liste", "Ja" if er_pa_ub else "Nei",
        ))

    # Seksjon C: UB (fortsatt pa lager) - forklart av IB/Innkjop, eller uforklart?
    seksjon_c = []

    for key, bil in queue_ub.items():
        regno, vin = bil["regno"], bil["vin"]
        tx_rows = vehicle_rows_by_key.get(key, [])
        seksjon_c.append(build_ub_row(
            regno, vin, tx_rows, key in queue_ib, key in queue_innkjop))

    return {
        "kontroll": "Varekjop bruktbil",
        "kildefil": INPUT_FILE_VAREKJOP,
        "orgnr": buyer_org_no,
        "dato_fra": dato_fra,
        "dato_til": dato_til,
        "antall_biler": len(biler),
        "biler_manifest": biler,
        "all_rows": all_rows,
        "seksjon_a": seksjon_a,
        "seksjon_b": seksjon_b,
        "seksjon_c": seksjon_c,
        "stats": {
            "Innkjop (antall)": len(queue_innkjop),
            "IB (antall)": len(queue_ib),
            "UB (antall)": len(queue_ub),
            "Innkjop - OK": sum(1 for r in seksjon_a if r["Status"].startswith("OK")),
            "Innkjop - Avvik": sum(1 for r in seksjon_a if r["Status"].startswith("Avvik")
                                   and "mangler i bokforing" not in r["Status"]),
            "OFV-kjop uten bokforing": sum(
                1 for r in seksjon_a if "mangler i bokforing" in r["Status"]),
        },
    }


# ------------------------------------------------------------------
# Kontroll 3: Kontroll Demobil
# ------------------------------------------------------------------


def build_demobil_row(regno: str, vin: str, bokfort_inn_raw, tx_rows: list,
                       company_org_no: str) -> dict:
    """En demobil skal vaere registrert pa selskapets juridiske navn:
    sjekker om bilens NYESTE registrerte OFV-transaksjon (uansett dato)
    fortsatt har company_org_no som kjoper."""

    bokfort_date = parse_bokfort_dato(bokfort_inn_raw)
    company_norm = normalize_identifier(company_org_no)

    model_name = ""
    chassis_no = vin
    regno_resolved = regno
    nyeste_row = None

    for r in tx_rows:

        if r.get("Status") != "OK":
            continue

        if not model_name:
            model_name = r.get("ModelName") or ""
        if r.get("ChassisNumber"):
            chassis_no = r["ChassisNumber"]
        if r.get("RegNo"):
            regno_resolved = r["RegNo"]

        tx_date = r.get("TransactionDate")
        if tx_date and (nyeste_row is None or tx_date > nyeste_row["TransactionDate"]):
            nyeste_row = r

    result = {
        "RegnrInput": regno_resolved,
        "Chassisnummer": chassis_no,
        "Modell": model_name,
        "BokfortInnDato": bokfort_date,
        "SisteTransaksjonsDato": None,
        "SisteKjoper": "",
    }

    if nyeste_row is None:
        result["Status"] = "Ingen treff - ingen OFV-transaksjoner funnet"
    else:
        result["SisteTransaksjonsDato"] = nyeste_row["TransactionDate"]
        result["SisteKjoper"] = compute_owner_label(
            nyeste_row.get("ToOwnerType"), nyeste_row.get("ToOwnerCompanyName"))

        if normalize_identifier(nyeste_row.get("ToOwnerOrgNo")) == company_norm:
            navn = nyeste_row.get("ToOwnerCompanyName") or company_org_no
            result["Status"] = f"OK - fortsatt eid av {navn}"
        else:
            result["Status"] = f"Avvik - siste eierskifte er til {result['SisteKjoper']}"

    return result


def run_kontroll_demobil() -> dict:

    print("=== Kontroll Demobil ===")

    konfig = read_konfig(INPUT_FILE_DEMOBIL)
    company_org_no = str(konfig.get("OrgNr") or "").strip()
    dato_fra = parse_bokfort_dato(konfig.get("DatoFra"))
    dato_til = parse_bokfort_dato(konfig.get("DatoTil"))

    if not company_org_no:
        raise ValueError("Kontroll Demobil: OrgNr ma vaere fylt ut i Konfig-fanen.")

    biler = read_biler(INPUT_FILE_DEMOBIL, "BokfortInnDato")

    all_rows = []
    kontroll_rows = []
    total = len(biler)

    for i, bil in enumerate(biler, start=1):

        regno, vin = bil["regno"], bil["vin"]
        ident = bil["identifikator"]

        print(f"  OFV: Eierskifter {ident}  ({i}/{total})")

        tx_rows = fetch_ofv_transactions(OFV_API_KEY, ident, bool(vin), regno, vin)
        all_rows.extend(tx_rows)

        kontroll_rows.append(
            build_demobil_row(regno, vin, bil["BokfortInnDato"], tx_rows, company_org_no))

        time.sleep(API_PAUSE_S)

    return {
        "kontroll": "Kontroll Demobil",
        "kildefil": INPUT_FILE_DEMOBIL,
        "orgnr": company_org_no,
        "dato_fra": dato_fra,
        "dato_til": dato_til,
        "antall_biler": total,
        "biler_manifest": biler,
        "all_rows": all_rows,
        "kontroll_rows": kontroll_rows,
        "stats": {
            "Antall biler": total,
            "Fortsatt hos enhet": sum(
                1 for r in kontroll_rows if r["Status"].startswith("OK")),
            "Avvik (videresolgt)": sum(
                1 for r in kontroll_rows if r["Status"].startswith("Avvik")),
        },
    }


# ------------------------------------------------------------------
# Kjor alle tre kontrollene
# ------------------------------------------------------------------

_t0 = time.time()

resultat_solgte_biler = run_kontroll_solgte_biler()
resultat_varekjop = run_kontroll_varekjop_bruktbil()
resultat_demobil = run_kontroll_demobil()

_elapsed_s = time.time() - _t0

# COMMAND ----------

# ==================================================================
# CELLE 4: OUTPUT - samlet Excel-arbeidsbok + HTML-rapport
# ==================================================================

# ------------------------------------------------------------------
# "Input"-fane - viser hvor input kommer fra og hva den tilhorer
# ------------------------------------------------------------------


def skriv_input_fane(wb: Workbook, bundles: list) -> None:

    ws = wb.create_sheet("Input")

    set_title_bar(ws, 1, 6, "Input - oversikt over kildefiler og konfigurasjon")
    ws.row_dimensions[1].height = 26

    set_section_bar(ws, 3, 6, "Konfigurasjon brukt per kontroll")

    headers = ["Kontroll", "Kildefil", "Orgnr", "Dato fra", "Dato til", "Antall biler"]
    style_table_header(ws, 4, headers)

    r = 5
    for b in bundles:
        ws.cell(row=r, column=1, value=b["kontroll"])
        ws.cell(row=r, column=2, value=b["kildefil"])
        ws.cell(row=r, column=3, value=b["orgnr"])
        if b["dato_fra"]:
            ws.cell(row=r, column=4, value=b["dato_fra"]).number_format = "dd.mm.yyyy"
        if b["dato_til"]:
            ws.cell(row=r, column=5, value=b["dato_til"]).number_format = "dd.mm.yyyy"
        ws.cell(row=r, column=6, value=b["antall_biler"])
        r += 1

    r += 2
    set_section_bar(ws, r, 6, "Biler lest per kontroll")
    r += 1

    headers2 = ["Kontroll", "Regnr/VIN", "Bokfort dato", "Liste-type", "Bokfort inn dato"]
    style_table_header(ws, r, headers2)
    r += 1

    for b in bundles:
        for bil in b["biler_manifest"]:
            ws.cell(row=r, column=1, value=b["kontroll"])
            ws.cell(row=r, column=2, value=bil["identifikator"])
            if "BokfortDato" in bil and bil["BokfortDato"] is not None:
                d = parse_bokfort_dato(bil["BokfortDato"])
                if d:
                    ws.cell(row=r, column=3, value=d).number_format = "dd.mm.yyyy"
            if "Liste" in bil and bil["Liste"]:
                ws.cell(row=r, column=4, value=str(bil["Liste"]))
            if "BokfortInnDato" in bil and bil["BokfortInnDato"] is not None:
                d = parse_bokfort_dato(bil["BokfortInnDato"])
                if d:
                    ws.cell(row=r, column=5, value=d).number_format = "dd.mm.yyyy"
            r += 1

    set_column_widths(ws, {"A": 26, "B": 45, "C": 16, "D": 18, "E": 16, "F": 14})


# ------------------------------------------------------------------
# "Resultat ..." - full transaksjonshistorikk, delt mellom alle tre
# ------------------------------------------------------------------


def skriv_resultat_fane(wb: Workbook, sheet_name: str, all_rows: list) -> None:

    ws = wb.create_sheet(sheet_name[:31])
    records = build_resultat_rows(all_rows)

    headers = [header for _key, header, _is_date in FIELD_MAP]
    style_table_header(ws, 1, headers)

    for r_idx, rec in enumerate(records, start=2):

        for c_idx, (_key, header, is_date) in enumerate(FIELD_MAP, start=1):
            value = rec.get(header)
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            if is_date and value:
                cell.number_format = "dd.mm.yyyy"

        if rec["_ErForsteIGruppe"]:
            for c_idx in range(1, len(headers) + 1):
                ws.cell(row=r_idx, column=c_idx).fill = PatternFill(
                    "solid", fgColor=FARGE_GRUPPE_BG)
                ws.cell(row=r_idx, column=c_idx).font = Font(bold=True)

    widths = {get_column_letter(i): 18 for i in range(1, len(headers) + 1)}
    widths.update({"A": 14, "D": 20, "P": 25, "Q": 25, "U": 25})
    set_column_widths(ws, widths)
    ws.freeze_panes = "A2"


# ------------------------------------------------------------------
# "Kontroll solgte biler"
# ------------------------------------------------------------------


def skriv_kontroll_solgte_biler(wb: Workbook, bundle: dict) -> None:

    ws = wb.create_sheet("Kontroll solgte biler"[:31])
    rows = bundle["kontroll_rows"]

    set_title_bar(ws, 1, 12, "Kontroll solgte biler")
    set_explanation_row(
        ws, 2, 12,
        "Kontrollregel: Bokfort dato sjekkes mot den OFV-transaksjonen som "
        "ligger NAERMEST bokfort dato i tid (for eller etter). Dager avvik "
        "er antall dager mellom denne datoen og bokfort dato. Har OFV ingen "
        "transaksjoner for kjoretoyet, brukes forstegangsregistrering fra "
        "Statens vegvesen (SVV) i stedet. Er OrgNrSelger fylt ut i input, "
        "flagges Selvhandel - det tvinger Dager avvik rod uansett faktisk "
        "avvik. Sortert med storst avvik forst."
    )

    bucket0 = sum(1 for r in rows if r["Kontrollert"] == "Ja"
                  and isinstance(r["DagerAvvik"], (int, float))
                  and r["Selvhandel"] != "Ja" and r["DagerAvvik"] <= AVVIK_GRONN_MAX)
    bucket1 = sum(1 for r in rows if r["Kontrollert"] == "Ja"
                  and isinstance(r["DagerAvvik"], (int, float))
                  and r["Selvhandel"] != "Ja"
                  and AVVIK_GRONN_MAX < r["DagerAvvik"] <= AVVIK_GUL_MAX)
    bucket2 = sum(1 for r in rows if r["Kontrollert"] == "Ja"
                  and (r["Selvhandel"] == "Ja"
                       or not isinstance(r["DagerAvvik"], (int, float))
                       or r["DagerAvvik"] > AVVIK_GUL_MAX))

    set_kpi_box(ws, 4, 1, 2, "Inputbiler", len(rows), big=True)
    set_kpi_box(ws, 4, 3, 4, "0-2 dager", bucket0, FARGE_OK_BG, FARGE_OK_FONT, big=True)
    set_kpi_box(ws, 4, 5, 6, "3-14 dager", bucket1, FARGE_VARSEL_BG, FARGE_VARSEL_FONT, big=True)
    set_kpi_box(ws, 4, 7, 8, "15+ dager / selvhandel", bucket2, FARGE_AVVIK_BG, FARGE_AVVIK_FONT, big=True)

    header_row = 7
    headers = ["Kilde", "API-treff", "Regnr", "Chassisnummer", "Modell",
               "Forstegangsregistrert", "Bokfort dato", "Transaksjonsdato",
               "RegistreringsType", "Dager avvik", "Selger", "Kjoper"]
    style_table_header(ws, header_row, headers)

    r = header_row + 1
    for row in rows:

        ws.cell(row=r, column=1, value=row["Kilde"])
        ws.cell(row=r, column=2, value=row["ApiTreff"])
        ws.cell(row=r, column=3, value=row["RegnrInput"])
        ws.cell(row=r, column=4, value=row["Chassisnummer"])
        ws.cell(row=r, column=5, value=row["Modell"])
        if row["Forstegangsregistrert"]:
            ws.cell(row=r, column=6, value=row["Forstegangsregistrert"]).number_format = "dd.mm.yyyy"
        if row["BokfortDato"]:
            ws.cell(row=r, column=7, value=row["BokfortDato"]).number_format = "dd.mm.yyyy"
        if row["KontrollTransaksjonDato"]:
            ws.cell(row=r, column=8, value=row["KontrollTransaksjonDato"]).number_format = "dd.mm.yyyy"
        ws.cell(row=r, column=9, value=row["RegistreringsType"])

        dager = row["DagerAvvik"]
        if isinstance(dager, (int, float)):
            avvik_cell = ws.cell(row=r, column=10, value=dager)
            if row["Selvhandel"] == "Ja" or dager > AVVIK_GUL_MAX:
                avvik_cell.fill = PatternFill("solid", fgColor=FARGE_AVVIK_BG)
                avvik_cell.font = Font(color=FARGE_AVVIK_FONT, bold=True)
            elif dager <= AVVIK_GRONN_MAX:
                avvik_cell.fill = PatternFill("solid", fgColor=FARGE_OK_BG)
                avvik_cell.font = Font(color=FARGE_OK_FONT, bold=True)
            else:
                avvik_cell.fill = PatternFill("solid", fgColor=FARGE_VARSEL_BG)
                avvik_cell.font = Font(color=FARGE_VARSEL_FONT, bold=True)

        ws.cell(row=r, column=11, value=row["Selger"])
        ws.cell(row=r, column=12, value=row["Kjoper"])

        r += 1

    set_column_widths(ws, {
        "A": 10, "B": 24, "C": 14, "D": 22, "E": 18, "F": 16, "G": 16,
        "H": 16, "I": 24, "J": 12, "K": 25, "L": 25,
    })


# ------------------------------------------------------------------
# "Kontroll Varekjop Bruktbil"
# ------------------------------------------------------------------


def _skriv_bil_rad(ws, r: int, row: dict) -> None:
    ws.cell(row=r, column=1, value=row["RegnrInput"])
    ws.cell(row=r, column=2, value=row["Chassisnummer"])
    ws.cell(row=r, column=3, value=row["Modell"])
    ws.cell(row=r, column=4, value=row["FlaggVerdi"])
    if row["FraDato"]:
        ws.cell(row=r, column=5, value=row["FraDato"]).number_format = "dd.mm.yyyy"
    if row["TilDato"]:
        ws.cell(row=r, column=6, value=row["TilDato"]).number_format = "dd.mm.yyyy"
    ws.cell(row=r, column=7, value=row["KjoptFra"])
    if row["TilDato"]:
        ws.cell(row=r, column=8, value=row["Kjoper"])
    else:
        ws.cell(row=r, column=8, value="(fortsatt pa lager)")
    status_cell = ws.cell(row=r, column=9, value=row["Status"])
    color_status_cell(status_cell, row["Status"])


def skriv_kontroll_varekjop(wb: Workbook, bundle: dict) -> None:

    ws = wb.create_sheet("Kontroll Varekjop Bruktbil"[:31])
    seksjon_a, seksjon_b, seksjon_c = bundle["seksjon_a"], bundle["seksjon_b"], bundle["seksjon_c"]
    buyer_org_no = bundle["orgnr"]

    # Forhandlernavn hentes fra en transaksjon der buyerOrgNo star som kjoper.
    buyer_org_name = ""
    buyer_org_norm = normalize_identifier(buyer_org_no)
    for bundle_row in bundle["all_rows"]:
        if normalize_identifier(bundle_row.get("ToOwnerOrgNo")) == buyer_org_norm:
            buyer_org_name = bundle_row.get("ToOwnerCompanyName") or ""
            if buyer_org_name:
                break

    set_title_bar(ws, 1, 9, "Kontroll Varekjop Bruktbil")
    set_explanation_row(
        ws, 2, 9,
        f"Kontrollregel: Biler i Innkjop-listen ({bundle['dato_fra']:%d.%m.%Y} - "
        f"{bundle['dato_til']:%d.%m.%Y}) kontrolleres mot OFV sin "
        f"transaksjonshistorikk for {buyer_org_name} (orgnr {buyer_org_no}): var "
        "bilen pa lager hos forhandleren pa et tidspunkt i perioden (kjopt av "
        "forhandleren, og eventuelt ikke solgt videre for perioden startet), og "
        "er alt OFV sier er kjopt i perioden faktisk bokfort."
    )
    set_explanation_row(
        ws, 3, 9,
        "Seksjon B og C er tillegg og vises bare hvis IB- og/eller UB-listen "
        "er fylt ut: biler kjopt forrige periode (IB) kontrolleres mot om de "
        "er solgt i denne perioden, og biler fortsatt pa lager (UB) "
        "kontrolleres mot om de kan forklares av IB eller innkjop i perioden."
    )

    r = 5

    # Seksjon A
    bekreftet = sum(1 for x in seksjon_a if x["Status"].startswith("OK"))
    avvik = sum(1 for x in seksjon_a if x["Status"].startswith("Avvik")
                and "mangler i bokforing" not in x["Status"])
    mangler = sum(1 for x in seksjon_a if "mangler i bokforing" in x["Status"])

    set_section_bar(ws, r, 9, "A - Innkjop i perioden")
    r += 2
    set_kpi_box(ws, r - 1, 1, 2, "Innkjop", len(seksjon_a), big=True)
    set_kpi_box(ws, r - 1, 3, 4, "OK - pa lager i perioden", bekreftet, FARGE_OK_BG, FARGE_OK_FONT, big=True)
    set_kpi_box(ws, r - 1, 5, 6, "Avvik", avvik, FARGE_AVVIK_BG, FARGE_AVVIK_FONT, big=True)
    set_kpi_box(ws, r - 1, 7, 8, "OFV-kjop uten bokforing", mangler, FARGE_VARSEL_BG, FARGE_VARSEL_FONT, big=True)
    r += 1

    set_explanation_row(ws, r, 9, f"Forhandler: {buyer_org_name}")
    r += 1

    style_table_header(ws, r, ["Regnr/VIN", "Chassisnummer", "Modell", "I OFV-liste",
                                "Pa lager fra", "Pa lager til", "Kjopt fra", "Solgt til", "Status"])
    r += 1
    for row in seksjon_a:
        _skriv_bil_rad(ws, r, row)
        r += 1

    r += 1

    # Seksjon B
    if seksjon_b:

        ok_ib = sum(1 for x in seksjon_b if x["Status"].startswith("OK"))
        solgt_for = sum(1 for x in seksjon_b if x["Status"] == "Avvik: solgt for perioden startet")
        ikke_funnet = sum(1 for x in seksjon_b
                           if x["Status"] == "Avvik: fant ikke opprinnelig kjop hos forhandleren")

        set_section_bar(ws, r, 9, "B - IB (kjopt forrige periode)")
        r += 2
        set_kpi_box(ws, r - 1, 1, 2, "IB-biler", len(seksjon_b), big=True)
        set_kpi_box(ws, r - 1, 3, 4, "OK - pa lager i perioden", ok_ib, FARGE_OK_BG, FARGE_OK_FONT, big=True)
        set_kpi_box(ws, r - 1, 5, 6, "Avvik: solgt for perioden", solgt_for, FARGE_AVVIK_BG, FARGE_AVVIK_FONT, big=True)
        set_kpi_box(ws, r - 1, 7, 8, "Avvik: fant ikke kjop", ikke_funnet, FARGE_AVVIK_BG, FARGE_AVVIK_FONT, big=True)
        r += 1

        set_explanation_row(ws, r, 9, f"Forhandler: {buyer_org_name}")
        r += 1

        style_table_header(ws, r, ["Regnr/VIN", "Chassisnummer", "Modell", "Pa UB-liste",
                                    "Pa lager fra", "Pa lager til", "Kjopt fra", "Solgt til", "Status"])
        r += 1
        for row in seksjon_b:
            _skriv_bil_rad(ws, r, row)
            r += 1

        r += 1

    # Seksjon C
    if seksjon_c:

        forklart = sum(1 for x in seksjon_c if x["Status"] == "OK - forklart")
        uforklart = len(seksjon_c) - forklart

        set_section_bar(ws, r, 9, "C - UB (fortsatt pa lager)")
        r += 2
        set_kpi_box(ws, r - 1, 1, 3, "UB-biler", len(seksjon_c), big=True)
        set_kpi_box(ws, r - 1, 4, 5, "Forklart", forklart, FARGE_OK_BG, FARGE_OK_FONT, big=True)
        set_kpi_box(ws, r - 1, 6, 8, "Uforklart", uforklart, FARGE_AVVIK_BG, FARGE_AVVIK_FONT, big=True)
        r += 1

        style_table_header(ws, r, ["Regnr/VIN", "Chassisnummer", "Modell",
                                    "Funnet i IB", "Funnet i innkjop", "Status"])
        r += 1
        for row in seksjon_c:
            ws.cell(row=r, column=1, value=row["RegnrInput"])
            ws.cell(row=r, column=2, value=row["Chassisnummer"])
            ws.cell(row=r, column=3, value=row["Modell"])
            ws.cell(row=r, column=4, value=row["FunnetIIB"])
            ws.cell(row=r, column=5, value=row["FunnetIInnkjop"])
            status_cell = ws.cell(row=r, column=6, value=row["Status"])
            color_status_cell(status_cell, row["Status"])
            r += 1

    set_column_widths(ws, {
        "A": 14, "B": 22, "C": 18, "D": 16, "E": 18, "F": 18, "G": 25, "H": 25, "I": 34,
    })


# ------------------------------------------------------------------
# "Kontroll Demobil"
# ------------------------------------------------------------------


def skriv_kontroll_demobil(wb: Workbook, bundle: dict) -> None:

    ws = wb.create_sheet("Kontroll Demobil"[:31])
    rows = bundle["kontroll_rows"]

    set_title_bar(ws, 1, 7, "Kontroll Demobil")
    set_explanation_row(
        ws, 2, 7,
        "Kontrollregel: For hver bil sjekkes bilens NYESTE registrerte "
        f"OFV-transaksjon (uansett dato) - kjoperen der skal vaere juridisk "
        f"enhet (orgnr {bundle['orgnr']}). Er kjoperen et annet selskap "
        "eller en privatperson, er bilen sannsynligvis solgt videre og "
        "flagges som avvik."
    )

    ok = sum(1 for r in rows if r["Status"].startswith("OK"))
    avvik = sum(1 for r in rows if r["Status"].startswith("Avvik"))

    set_kpi_box(ws, 4, 1, 2, "Inputbiler", len(rows), big=True)
    set_kpi_box(ws, 4, 3, 4, "Fortsatt hos enhet", ok, FARGE_OK_BG, FARGE_OK_FONT, big=True)
    set_kpi_box(ws, 4, 5, 6, "Avvik (videresolgt)", avvik, FARGE_AVVIK_BG, FARGE_AVVIK_FONT, big=True)

    header_row = 7
    style_table_header(ws, header_row, ["Regnr", "Chassisnummer", "Modell",
                                         "Bokfort inn dato", "Siste transaksjonsdato",
                                         "Siste kjoper", "Status"])

    r = header_row + 1
    for row in rows:
        ws.cell(row=r, column=1, value=row["RegnrInput"])
        ws.cell(row=r, column=2, value=row["Chassisnummer"])
        ws.cell(row=r, column=3, value=row["Modell"])
        if row["BokfortInnDato"]:
            ws.cell(row=r, column=4, value=row["BokfortInnDato"]).number_format = "dd.mm.yyyy"
        if row["SisteTransaksjonsDato"]:
            ws.cell(row=r, column=5, value=row["SisteTransaksjonsDato"]).number_format = "dd.mm.yyyy"
        ws.cell(row=r, column=6, value=row["SisteKjoper"])
        status_cell = ws.cell(row=r, column=7, value=row["Status"])
        color_status_cell(status_cell, row["Status"])
        r += 1

    set_column_widths(ws, {"A": 14, "B": 22, "C": 18, "D": 18, "E": 18, "F": 25, "G": 38})


# ------------------------------------------------------------------
# Bygg hele arbeidsboken
# ------------------------------------------------------------------

_wb = Workbook()
_wb.remove(_wb.active)

_bundles = [resultat_solgte_biler, resultat_varekjop, resultat_demobil]

skriv_input_fane(_wb, _bundles)
skriv_resultat_fane(_wb, "Resultat solgte biler", resultat_solgte_biler["all_rows"])
skriv_kontroll_solgte_biler(_wb, resultat_solgte_biler)
skriv_resultat_fane(_wb, "Resultat Varekjop", resultat_varekjop["all_rows"])
skriv_kontroll_varekjop(_wb, resultat_varekjop)
skriv_resultat_fane(_wb, "Resultat Demobil", resultat_demobil["all_rows"])
skriv_kontroll_demobil(_wb, resultat_demobil)

_wb.save(OUTPUT_XLSX)
print(f"Excel-arbeidsbok lagret: {OUTPUT_XLSX}")

# ------------------------------------------------------------------
# HTML-rapport (samme innhold, BDO-farger) via Jinja2
# ------------------------------------------------------------------

from jinja2 import Template  # noqa: E402  (importert her, nar den trengs)

_HTML_TEMPLATE = Template("""
<!DOCTYPE html>
<html lang="no">
<head>
<meta charset="utf-8">
<title>OFV-kontroller</title>
<style>
  body { font-family: Arial, Helvetica, sans-serif; color: #{{ bdo.charcoal }};
         background: #ffffff; margin: 24px; }
  h1 { background: #{{ bdo.charcoal }}; color: #ffffff; padding: 12px 16px;
       border-left: 6px solid #{{ bdo.red }}; }
  h2 { background: #{{ bdo.pale_charcoal }}; padding: 8px 12px; margin-top: 32px; }
  .forklaring { font-style: italic; color: #{{ bdo.slate }}; margin-bottom: 8px; }
  .kpi-row { display: flex; gap: 12px; margin-bottom: 12px; flex-wrap: wrap; }
  .kpi { border: 1px solid #{{ bdo.pale_charcoal }}; border-radius: 6px;
         padding: 10px 18px; text-align: center; min-width: 140px; }
  .kpi .label { font-size: 12px; font-weight: bold; }
  .kpi .value { font-size: 24px; font-weight: bold; }
  .kpi.ok { background: #{{ bdo.jade }}; color: #fff; }
  .kpi.avvik { background: #{{ bdo.red }}; color: #fff; }
  .kpi.varsel { background: #{{ bdo.gold }}; color: #fff; }
  table { border-collapse: collapse; width: 100%; margin-bottom: 24px; font-size: 13px; }
  th { background: #{{ bdo.charcoal }}; color: #fff; padding: 6px 8px; text-align: left; }
  td { padding: 5px 8px; border-bottom: 1px solid #{{ bdo.pale_charcoal }}; }
  tr.group-first { background: #F4F4F4; font-weight: bold; }
  .status-ok { background: #{{ bdo.jade }}; color: #fff; font-weight: bold; }
  .status-avvik { background: #{{ bdo.red }}; color: #fff; font-weight: bold; }
</style>
</head>
<body>

<h1>OFV-kontroller - samlet rapport</h1>

<h2>Input - kildefiler og konfigurasjon</h2>
<table>
<tr><th>Kontroll</th><th>Kildefil</th><th>Orgnr</th><th>Dato fra</th><th>Dato til</th><th>Antall biler</th></tr>
{% for b in bundles %}
<tr>
  <td>{{ b.kontroll }}</td><td>{{ b.kildefil }}</td><td>{{ b.orgnr }}</td>
  <td>{{ b.dato_fra or "" }}</td><td>{{ b.dato_til or "" }}</td><td>{{ b.antall_biler }}</td>
</tr>
{% endfor %}
</table>

{% for seksjon in seksjoner %}
<h2>{{ seksjon.titel }}</h2>
{% if seksjon.forklaring %}<div class="forklaring">{{ seksjon.forklaring }}</div>{% endif %}
<div class="kpi-row">
{% for kpi in seksjon.kpier %}
  <div class="kpi {{ kpi.klasse }}"><div class="label">{{ kpi.label }}</div><div class="value">{{ kpi.verdi }}</div></div>
{% endfor %}
</div>
<table>
<tr>{% for h in seksjon.kolonner %}<th>{{ h }}</th>{% endfor %}</tr>
{% for rad in seksjon.rader %}
<tr class="{{ rad.radklasse }}">
  {% for verdi, celleklasse in rad.celler %}<td class="{{ celleklasse }}">{{ verdi if verdi is not none else "" }}</td>{% endfor %}
</tr>
{% endfor %}
</table>
{% endfor %}

</body>
</html>
""")


def _status_klasse(status: str) -> str:
    status = status or ""
    if status.startswith("OK"):
        return "status-ok"
    if status.startswith("Avvik"):
        return "status-avvik"
    return ""


def _bygg_html_seksjoner(bundles: list) -> list:

    seksjoner = []

    # Kontroll solgte biler
    b1 = bundles[0]
    seksjoner.append({
        "titel": "Kontroll solgte biler",
        "forklaring": "Bokfort dato sjekkes mot naermeste OFV-transaksjon i tid.",
        "kpier": [
            {"label": "Inputbiler", "verdi": b1["stats"]["Antall biler"], "klasse": ""},
            {"label": "OFV-treff", "verdi": b1["stats"]["OFV-treff"], "klasse": "ok"},
            {"label": "Uten treff", "verdi": b1["stats"]["Uten OFV-treff"], "klasse": "varsel"},
            {"label": "OFV-feil", "verdi": b1["stats"]["OFV-feil"], "klasse": "avvik"},
        ],
        "kolonner": ["Kilde", "API-treff", "Regnr", "Chassisnummer", "Modell",
                     "Bokfort dato", "Transaksjonsdato", "Dager avvik", "Selger", "Kjoper"],
        "rader": [
            {
                "radklasse": "",
                "celler": [
                    (r["Kilde"], ""), (r["ApiTreff"], ""), (r["RegnrInput"], ""),
                    (r["Chassisnummer"], ""), (r["Modell"], ""), (r["BokfortDato"], ""),
                    (r["KontrollTransaksjonDato"], ""), (r["DagerAvvik"], ""),
                    (r["Selger"], ""), (r["Kjoper"], ""),
                ],
            }
            for r in b1["kontroll_rows"]
        ],
    })

    # Varekjop bruktbil (Seksjon A)
    b2 = bundles[1]
    seksjoner.append({
        "titel": "Varekjop Bruktbil - Seksjon A (Innkjop i perioden)",
        "forklaring": f"Orgnr {b2['orgnr']}, periode {b2['dato_fra']} - {b2['dato_til']}.",
        "kpier": [
            {"label": "Innkjop", "verdi": len(b2["seksjon_a"]), "klasse": ""},
            {"label": "OK", "verdi": b2["stats"]["Innkjop - OK"], "klasse": "ok"},
            {"label": "Avvik", "verdi": b2["stats"]["Innkjop - Avvik"], "klasse": "avvik"},
            {"label": "Mangler bokforing", "verdi": b2["stats"]["OFV-kjop uten bokforing"], "klasse": "varsel"},
        ],
        "kolonner": ["Regnr/VIN", "Chassisnummer", "Modell", "Pa lager fra",
                     "Pa lager til", "Kjopt fra", "Solgt til", "Status"],
        "rader": [
            {
                "radklasse": "",
                "celler": [
                    (r["RegnrInput"], ""), (r["Chassisnummer"], ""), (r["Modell"], ""),
                    (r["FraDato"], ""), (r["TilDato"], ""), (r["KjoptFra"], ""),
                    (r["Kjoper"] or "(fortsatt pa lager)", ""),
                    (r["Status"], _status_klasse(r["Status"])),
                ],
            }
            for r in b2["seksjon_a"]
        ],
    })

    # Kontroll Demobil
    b3 = bundles[2]
    seksjoner.append({
        "titel": "Kontroll Demobil",
        "forklaring": f"Orgnr {b3['orgnr']} - siste OFV-transaksjon skal ha dette orgnr som kjoper.",
        "kpier": [
            {"label": "Inputbiler", "verdi": b3["stats"]["Antall biler"], "klasse": ""},
            {"label": "Fortsatt hos enhet", "verdi": b3["stats"]["Fortsatt hos enhet"], "klasse": "ok"},
            {"label": "Avvik", "verdi": b3["stats"]["Avvik (videresolgt)"], "klasse": "avvik"},
        ],
        "kolonner": ["Regnr", "Chassisnummer", "Modell", "Bokfort inn dato",
                     "Siste transaksjonsdato", "Siste kjoper", "Status"],
        "rader": [
            {
                "radklasse": "",
                "celler": [
                    (r["RegnrInput"], ""), (r["Chassisnummer"], ""), (r["Modell"], ""),
                    (r["BokfortInnDato"], ""), (r["SisteTransaksjonsDato"], ""),
                    (r["SisteKjoper"], ""), (r["Status"], _status_klasse(r["Status"])),
                ],
            }
            for r in b3["kontroll_rows"]
        ],
    })

    return seksjoner


_html = _HTML_TEMPLATE.render(bdo=BDO, bundles=_bundles, seksjoner=_bygg_html_seksjoner(_bundles))

with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
    f.write(_html)

print(f"HTML-rapport lagret: {OUTPUT_HTML}")

# COMMAND ----------

# ==================================================================
# CELLE 5: OPPSUMMERING
# ==================================================================

print("=" * 70)
print("OFV-KONTROLLER - OPPSUMMERING")
print("=" * 70)
print(f"Tidsbruk: {_elapsed_s:.1f} sekunder\n")

for bundle in _bundles:
    print(f"--- {bundle['kontroll']} ---")
    print(f"  Kildefil: {bundle['kildefil']}")
    print(f"  Orgnr: {bundle['orgnr']}")
    if bundle["dato_fra"] and bundle["dato_til"]:
        print(f"  Periode: {bundle['dato_fra']} - {bundle['dato_til']}")
    for key, value in bundle["stats"].items():
        print(f"  {key}: {value}")
    print()

print(f"Excel-arbeidsbok: {OUTPUT_XLSX}")
print(f"HTML-rapport: {OUTPUT_HTML}")
