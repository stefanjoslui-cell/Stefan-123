# Databricks notebook source
# MAGIC %md
# MAGIC # OFV-kontroller (Python/Databricks-versjon av OFV_RefreshInfo.bas)
# MAGIC
# MAGIC Gjor noyaktig det samme som VBA-makroen, portert til en Databricks notebook:
# MAGIC
# MAGIC - **Kontroll 1 - Kontroll solgte biler**: for hver bil hentes hele
# MAGIC   OFV-transaksjonshistorikken, og den transaksjonen som ligger **naermest**
# MAGIC   bokfort dato (uansett om for eller etter) brukes som kontrollgrunnlag.
# MAGIC   Dager avvik farges gronn (0-2), gul (3-14) eller rod (15+). Er
# MAGIC   "Juridisk enhet (Selger)"-orgnr fylt ut, sjekkes i tillegg om samme
# MAGIC   selskap star som bade kjoper og selger (selvhandel) - da tvinges raden
# MAGIC   rod uansett dagers avvik.
# MAGIC - **Kontroll 2 - Varekjop bruktbil**: henter forst ALLE biler OFV sier er
# MAGIC   kjopt av et gitt orgnr i en periode (ett samlekall, uavhengig av regnr),
# MAGIC   sammenligner mot bokforingslisten, og sjekker om hver bil senere er
# MAGIC   avregistrert (RegistreringsType = "Juridisk eierskifte (ikke i
# MAGIC   bestand)") av samme selskap.
# MAGIC - **Kontroll 3 - Kontroll Demobil**: sjekker om bilens NYESTE registrerte
# MAGIC   transaksjon (uansett dato) fortsatt har juridisk enhet som kjoper.
# MAGIC
# MAGIC ## Antagelser du bor verifisere mot en ekte OFV-respons
# MAGIC - Organisasjonsnummer i transaksjonsresponsen antas a ligge i
# MAGIC   `from.owner.companyInfo.organizationNumber` / `to.owner.companyInfo.organizationNumber`
# MAGIC   (samme monster som filternavnene `fromOrganizationNumber`/
# MAGIC   `toOrganizationNumber`, men IKKE bekreftet i swagger-dokumentasjonen).
# MAGIC   Sjekk `build_transaction_row()` hvis selvhandel-/orgnr-sjekkene ikke
# MAGIC   treffer riktig.
# MAGIC - SVV-responsens feltnavn for forstegangsregistrering er ukjent i dybde/
# MAGIC   nesting, sa `find_any()` soker rekursivt etter feltnavnet i hele
# MAGIC   responsen (samme "finn dette navnet uansett hvor det ligger"-prinsipp
# MAGIC   som den opprinnelige VBA-koden brukte).
# MAGIC
# MAGIC ## Input
# MAGIC Last opp EN Excel-fil (sti satt i widget-en "input_fil") med tre faner,
# MAGIC navngitt eksakt slik:
# MAGIC - `KontrollSolgteBiler` - kolonner: Regnr, VIN, BokfortDato
# MAGIC - `VarekjopBruktbil` - kolonner: Regnr, VIN, BokfortDato
# MAGIC - `KontrollDemobil` - kolonner: Regnr, VIN, BokfortInnDato
# MAGIC
# MAGIC (Kun en av Regnr/VIN fylles ut per rad - koden kjenner dem automatisk fra
# MAGIC hverandre pa lengde, siden VIN alltid er noyaktig 17 tegn.)
# MAGIC
# MAGIC Organisasjonsnumre og datointervaller settes som egne widgets (de er
# MAGIC enkeltverdier, ikke kolonner i tabellen).
# MAGIC
# MAGIC ## Output
# MAGIC En ny Excel-arbeidsbok (sti satt i widget-en "output_fil") med en fane per
# MAGIC Resultat-/Kontroll-tabell, inkludert samme fargekoding som i VBA-versjonen.

# COMMAND ----------

# MAGIC %pip install openpyxl requests

# COMMAND ----------

import time
from datetime import date, datetime

import pandas as pd
import requests
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# COMMAND ----------

# ==============================================================
# WIDGETS - dette er "vinduet" du bruker i stedet for
# frmVelgKontroll-popup-vinduet / InputBox-en i VBA-versjonen.
# Databricks kan ikke vise et tkinter-vindu (ingen skjerm pa
# klyngen), sa valget gjores her ovenst i notebooken i stedet.
# ==============================================================

dbutils.widgets.dropdown(
    "kontroll", "4 - Alle",
    ["1 - Kontroll solgte biler", "2 - Varekjop bruktbil",
     "3 - Kontroll Demobil", "4 - Alle"],
    "Hvilken kontroll skal kjores?",
)

dbutils.widgets.text("input_fil", "/dbfs/FileStore/ofv_input.xlsx",
                      "Sti til input-arbeidsbok (.xlsx)")
dbutils.widgets.text("output_fil", "/dbfs/FileStore/ofv_resultater.xlsx",
                      "Sti til output-arbeidsbok (.xlsx)")
dbutils.widgets.text("secret_scope", "ofv-secrets",
                      "Databricks secret scope for API-nokler")

dbutils.widgets.text("selger_org_no", "",
                      "Kontroll 1: Juridisk enhet (Selger) - orgnr (valgfritt)")

dbutils.widgets.text("kjoper_org_no", "",
                      "Kontroll 2: Juridisk enhet (Org Nr) - orgnr")
dbutils.widgets.text("kjoper_dato_fra", "01.01.2025",
                      "Kontroll 2: Dato fra (DD.MM.AAAA)")
dbutils.widgets.text("kjoper_dato_til", "31.12.2025",
                      "Kontroll 2: Dato til (DD.MM.AAAA)")

dbutils.widgets.text("demobil_org_no", "",
                      "Kontroll 3: Juridisk enhet - orgnr")

# COMMAND ----------

# ==============================================================
# KONSTANTER
# ==============================================================

OFV_URL = "https://api.ofv.no/transactions/v1/"
SVV_URL = "https://akfell-datautlevering.atlas.vegvesen.no/enkeltoppslag/kjoretoydata"

MAX_RETRIES = 4
RETRY_WAIT_S = 3
API_PAUSE_S = 0.15

# RegistreringsType-verdi som (etter avtale) betyr at kjoretoyet forlot
# den registrerte eierens aktive bestand i denne transaksjonen - brukes
# til "avregistrert"-sjekken i Varekjop bruktbil.
REGTYPE_IKKE_BESTAND = "Juridisk eierskifte (ikke i bestand)"

AVVIK_GRONN_MAX = 2
AVVIK_GUL_MAX = 14

FILL_GRONN = PatternFill("solid", fgColor="C6EFCE")
FILL_GUL = PatternFill("solid", fgColor="FFEB9C")
FILL_ROD = PatternFill("solid", fgColor="FFC7CE")

CAR_LEVEL_FIELDS = {
    "RegNo", "ChassisNumber", "MakeName", "ModelName", "FuelGroup",
    "IsLeased", "IsUsedImported", "FirstRegistrationDate",
}

# COMMAND ----------

# ==============================================================
# GENERISKE HJELPEFUNKSJONER (tekst, dato, JSON-oppslag)
# ==============================================================


def normalize_identifier(value) -> str:
    return str(value or "").strip().upper().replace(" ", "")


def is_vin(text: str) -> bool:
    """VIN (chassisnummer) er alltid noyaktig 17 tegn per ISO 3779 - en
    fast, internasjonal standard. Norske regnr kan derimot variere i
    lengde, sa vi sjekker IKKE et fast bokstav/tall-monster for regnr -
    alt som ikke er 17 tegn regnes som regnr."""
    return len(text) == 17


def parse_bokfort_dato(raw):
    """Tolker en dato skrevet inn av en bruker (Excel-celle eller
    widget-tekst). Ekte datetime/Timestamp-objekter (fra pandas som har
    lest en Excel-datocelle) brukes direkte. Tekst tolkes EKSPLISITT som
    dag.maned.ar - ikke via pandas/locale-avhengig parsing."""

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
    tegnposisjoner - locale-uavhengig, akkurat som DateFromISO i VBA."""

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
    nestet dict/list-tre - mirrorer den flate "finn dette navnet uansett
    hvor det ligger"-tilnaermingen den opprinnelige VBA JSON-parseren
    brukte for SVV-responsen (hvis nesting-nivaet er ukjent/varierer)."""

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
    """Samme utledning som SelgerType/KjoperType-formlene i VBA-versjonen:
    Privat vinner over firmanavn, ellers firmanavn hvis det finnes,
    ellers den ra eiertype-teksten."""

    owner_type = owner_type or ""
    company_name = company_name or ""
    if not owner_type and not company_name:
        return ""
    if owner_type == "Privat":
        return "Privat"
    return company_name or owner_type


def build_vehicle_key(vin: str, regno: str) -> str:
    if vin:
        return f"VIN|{vin}"
    if regno:
        return f"REG|{regno}"
    return ""

# COMMAND ----------

# ==============================================================
# HTTP - OFV OG SVV
# ==============================================================


def post_ofv(api_key: str, body: dict):
    """POST med retry-logikk mot OFV, samme statuskode-handtering som
    PostOFVWithRetries i VBA-versjonen."""

    last_error = ""

    for attempt in range(1, MAX_RETRIES + 1):

        try:
            resp = requests.post(
                OFV_URL,
                headers={
                    "Content-Type": "application/json",
                    "Cache-Control": "no-cache",
                    "Ocp-Apim-Subscription-Key": api_key,
                },
                json=body,
                timeout=30,
            )
        except requests.RequestException as e:
            last_error = str(e)
            time.sleep(RETRY_WAIT_S * attempt)
            continue

        if resp.status_code == 200:
            return resp.json(), "OK"
        if resp.status_code == 401:
            return None, "Feil: OFV 401 - kontroller API-nokkelen"
        if resp.status_code == 403:
            return None, "Feil: OFV 403 - tilgang eller kvote"
        if resp.status_code == 404:
            return None, "Feil: OFV-endepunktet ble ikke funnet"
        if resp.status_code in (429, 500, 502, 503, 504):
            last_error = f"{resp.status_code}: {resp.text}"
            time.sleep(RETRY_WAIT_S * attempt)
            continue

        return None, f"Feil: OFV HTTP {resp.status_code}"

    return None, f"Feil: OFV - {last_error}"


def get_svv_response(api_key: str, filter_name: str, identifier: str):

    last_error = ""

    for attempt in range(1, MAX_RETRIES + 1):

        try:
            resp = requests.get(
                SVV_URL,
                params={filter_name: identifier},
                headers={
                    "SVV-Authorization": f"Apikey {api_key}",
                    "Accept": "application/json",
                },
                timeout=30,
            )
        except requests.RequestException as e:
            last_error = str(e)
            time.sleep(RETRY_WAIT_S * attempt)
            continue

        if resp.status_code == 200:
            return resp.json(), "Statens vegvesen"
        if resp.status_code == 401:
            return None, "Feil: Vegvesenet 401 - kontroller API-nokkel"
        if resp.status_code == 403:
            return None, "Feil: Vegvesenet 403 - tilgang eller kvote"
        if resp.status_code == 404:
            return None, "Statens vegvesen - ingen treff"
        if resp.status_code in (429, 500, 502, 503, 504):
            last_error = f"{resp.status_code}: {resp.text}"
            time.sleep(RETRY_WAIT_S * attempt)
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

# COMMAND ----------

# ==============================================================
# OFV-TRANSAKSJONER
# ==============================================================


def build_transaction_row(identifier, tx: dict, use_vin: bool,
                           original_regno: str, original_vin: str) -> dict:

    from_side = tx.get("from") or {}
    to_side = tx.get("to") or {}
    from_owner = from_side.get("owner") or {}
    to_owner = to_side.get("owner") or {}
    from_company = from_owner.get("companyInfo") or {}
    to_company = to_owner.get("companyInfo") or {}

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
        "FromOwnerType": from_owner.get("type"),
        "FromOwnerCompanyName": from_company.get("name"),
        # Antatt feltnavn - se merknad om antagelser ovenst i filen.
        "FromOwnerOrgNo": from_company.get("organizationNumber"),
        "FromOwnerCounty": from_owner.get("countyName"),
        "ToOwnerType": to_owner.get("type"),
        "ToOwnerCompanyName": to_company.get("name"),
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
    uten datofilter, sortert nyeste forst (cursor-paginert)."""

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

        data, status = post_ofv(api_key, body)

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
            "Ingen eierskifter i perioden"))

    return rows


def fetch_ofv_transactions_by_buyer_org(api_key: str, org_no: str,
                                         date_fra: date, date_til: date) -> list:
    """Henter ALLE OFV-transaksjoner der gitt orgnr star som KJOPER
    (toOrganizationNumber), innenfor et datointervall - ett samlekall
    uavhengig av regnr/VIN. Feil her stopper hele Varekjop bruktbil-
    kjoringen (kastes videre), i motsetning til per-kjoretoy-kallene som
    heller skriver en feilrad og fortsetter."""

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

        data, status = post_ofv(api_key, body)

        if status != "OK":
            raise RuntimeError(status)

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

# COMMAND ----------

# ==============================================================
# RESULTAT-TABELL (full transaksjonshistorikk, master/detalj)
# ==============================================================


def build_resultat_dataframe(all_rows: list) -> pd.DataFrame:
    """Full transaksjonshistorikk, nyeste transaksjon forst per bil (slik
    OFV allerede returnerer den). Kun den FORSTE (nyeste) raden for hver
    bil viser bilinfo (merke/modell/chassisnr osv.) - eldre transaksjoner
    for samme bil viser bare transaksjonsspesifikk info."""

    records = []
    forrige_input = None

    for row in all_rows:

        denne_input = row.get("Input")
        er_forste = denne_input != forrige_input
        forrige_input = denne_input

        def car(field, _er_forste=er_forste, _row=row):
            return _row.get(field) if _er_forste else None

        records.append({
            "Input": row.get("Input"),
            "Kilde": row.get("Kilde"),
            "RegNo": car("RegNo"),
            "Chassisnummer": car("ChassisNumber"),
            "Merke": car("MakeName"),
            "Modell": car("ModelName"),
            "RegistreringsType": row.get("RegistrationType"),
            "Drivstoffgruppe": car("FuelGroup"),
            "Leaset": car("IsLeased"),
            "Bruktimportert": car("IsUsedImported"),
            "ForstegangsRegistrering": car("FirstRegistrationDate"),
            "TransaksjonsNummer": row.get("TransactionNumber"),
            "Eierskiftedato": row.get("TransactionDate"),
            "SelgerType": compute_owner_label(
                row.get("FromOwnerType"), row.get("FromOwnerCompanyName")),
            "KjoperType": compute_owner_label(
                row.get("ToOwnerType"), row.get("ToOwnerCompanyName")),
            "SelgerEierType": row.get("FromOwnerType"),
            "SelgerEierFirma": row.get("FromOwnerCompanyName"),
            "SelgerOrgNr": row.get("FromOwnerOrgNo"),
            "SelgerEierFylke": row.get("FromOwnerCounty"),
            "KjoperEierType": row.get("ToOwnerType"),
            "KjoperEierFirma": row.get("ToOwnerCompanyName"),
            "KjoperOrgNr": row.get("ToOwnerOrgNo"),
            "KjoperEierFylke": row.get("ToOwnerCounty"),
            "Status": row.get("Status"),
        })

    return pd.DataFrame(records)

# COMMAND ----------

# ==============================================================
# KONTROLL 1: KONTROLL SOLGTE BILER
# ==============================================================


def build_kontroll_row(regno: str, vin: str, bokfort_raw, tx_rows: list,
                        svv_info, seller_org_no: str) -> dict:
    """bokfort dato mot HELE bilens OFV-transaksjonshistorikk - den
    transaksjonen som ligger NAERMEST bokfort dato brukes, uansett om
    den er for eller etter. Har OFV ingen transaksjoner i det hele tatt,
    brukes forstegangsregistreringsdato fra SVV i stedet.

    Valgfri tilleggskontroll (kun nar seller_org_no er fylt ut): er OGSA
    kjoperen samme org som selgerOrgNo (samme selskap pa begge sider),
    flagges Selvhandel="Ja" - raden skal da tvinges rod uansett dagers
    avvik nar den skrives til Excel."""

    bokfort_date = parse_bokfort_dato(bokfort_raw)
    seller_norm = normalize_identifier(seller_org_no)

    model_name = ""
    chassis_no = vin
    regno_resolved = regno
    first_reg_date = None
    has_any_ok = False

    nearest_row = None
    nearest_diff = None

    alle_transaksjoner = []

    for r in tx_rows:

        if r.get("Status") != "OK":
            continue

        has_any_ok = True
        alle_transaksjoner.append(r)

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
        "AlleTransaksjoner": sorted(
            alle_transaksjoner,
            key=lambda r: r.get("TransactionDate") or date.min,
            reverse=True,
        ),
        "MatchetTransaksjon": None,
    }

    if bokfort_date is None:
        if has_any_ok or first_reg_date is not None:
            result["ApiTreff"] = "Mangler bokfort dato"
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
        result["MatchetTransaksjon"] = nearest_row

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
        result["ApiTreff"] = (svv_info or {}).get("Status", "Ingen treff")
        result["Kontrollert"] = "Nei"

    return result


def process_kontroll_solgte_biler(df_input: pd.DataFrame, seller_org_no: str,
                                   ofv_key: str, svv_key: str):

    df_input = df_input.fillna("")
    all_rows = []
    kontroll_rows = []
    total = len(df_input)

    for i, row in enumerate(df_input.itertuples(index=False), start=1):

        regnr_raw = str(getattr(row, "Regnr", "") or "").strip()
        vin_raw = str(getattr(row, "VIN", "") or "").strip()
        bokfort_raw = getattr(row, "BokfortDato", None)

        ident = normalize_identifier(regnr_raw) if regnr_raw else normalize_identifier(vin_raw)
        use_vin = is_vin(ident)
        regno = "" if use_vin else ident
        vin = ident if use_vin else ""

        print(f"OFV: Eierskifter {ident}  ({i}/{total})  {round(i / total * 100)}%")

        tx_rows = fetch_ofv_transactions(ofv_key, ident, use_vin, regno, vin)
        all_rows.extend(tx_rows)

        has_any_ok = any(r.get("Status") == "OK" for r in tx_rows)

        svv_info = None
        if not has_any_ok and svv_key:
            print(f"SVV: Forstegangsregistrering {ident}  ({i}/{total})")
            svv_info = fetch_vehicle_info_from_svv(svv_key, regno, vin)

        kontroll_rows.append(
            build_kontroll_row(regno, vin, bokfort_raw, tx_rows, svv_info, seller_org_no))

        time.sleep(API_PAUSE_S)

    resultat_df = build_resultat_dataframe(all_rows)

    kontroll_df = pd.DataFrame([
        {
            "Kilde": r["Kilde"],
            "ApiTreff": r["ApiTreff"],
            "Regnr": r["RegnrInput"],
            "Chassisnummer": r["Chassisnummer"],
            "Modell": r["Modell"],
            "Forstegangsregistrert": r["Forstegangsregistrert"],
            "BokfortDato": r["BokfortDato"],
            "TransaksjonsDato": r["KontrollTransaksjonDato"],
            "RegistreringsType": r["RegistreringsType"],
            "DagerAvvik": r["DagerAvvik"],
            "Selger": r["Selger"],
            "Kjoper": r["Kjoper"],
            "Selvhandel": r["Selvhandel"],
        }
        for r in kontroll_rows
    ]).sort_values(
        by="DagerAvvik", ascending=False,
        key=lambda s: s.fillna(-1),
    ).reset_index(drop=True)

    detaljer_rows = []
    for r in kontroll_rows:
        matchet = r["MatchetTransaksjon"]
        for tx in r["AlleTransaksjoner"]:
            if matchet is not None and tx is matchet:
                continue
            detaljer_rows.append({
                "Regnr": r["RegnrInput"],
                "TransaksjonsDato": tx.get("TransactionDate"),
                "RegistreringsType": tx.get("RegistrationType"),
                "Selger": compute_owner_label(
                    tx.get("FromOwnerType"), tx.get("FromOwnerCompanyName")),
                "Kjoper": compute_owner_label(
                    tx.get("ToOwnerType"), tx.get("ToOwnerCompanyName")),
            })
    detaljer_df = pd.DataFrame(detaljer_rows)

    return resultat_df, kontroll_df, detaljer_df

# COMMAND ----------

# ==============================================================
# KONTROLL 2: VAREKJOP BRUKTBIL
# ==============================================================


def build_varekjop_row(regno: str, vin: str, bokfort_raw, tx_rows: list,
                        er_i_ofv_liste: bool, buyer_org_no: str) -> dict:

    bokfort_date = parse_bokfort_dato(bokfort_raw)
    buyer_norm = normalize_identifier(buyer_org_no)

    model_name = ""
    chassis_no = vin
    regno_resolved = regno
    kjopt_row = None
    avreg_row = None

    for r in tx_rows:

        if r.get("Status") != "OK":
            continue

        if not model_name:
            model_name = r.get("ModelName") or ""
        if r.get("ChassisNumber"):
            chassis_no = r["ChassisNumber"]
        if r.get("RegNo"):
            regno_resolved = r["RegNo"]

        # Kjopt av juridisk enhet: transaksjon der selskapet star som
        # kjoper (til-siden) - bruker den seneste hvis flere.
        if normalize_identifier(r.get("ToOwnerOrgNo")) == buyer_norm:
            if (kjopt_row is None
                    or (r.get("TransactionDate") and kjopt_row.get("TransactionDate")
                        and r["TransactionDate"] > kjopt_row["TransactionDate"])):
                kjopt_row = r

        # Avregistrert av juridisk enhet: transaksjon der selskapet star
        # som selger (fra-siden) OG registreringstypen viser at bilen
        # forlot bestanden - bruker den tidligste (forste avregistrering
        # etter kjop).
        if (normalize_identifier(r.get("FromOwnerOrgNo")) == buyer_norm
                and r.get("RegistrationType") == REGTYPE_IKKE_BESTAND):
            if (avreg_row is None
                    or (r.get("TransactionDate") and avreg_row.get("TransactionDate")
                        and r["TransactionDate"] < avreg_row["TransactionDate"])):
                avreg_row = r

    result = {
        "RegnrInput": regno_resolved,
        "Chassisnummer": chassis_no,
        "Modell": model_name,
        "BokfortDato": bokfort_date,
        "IOFVListe": "Ja" if er_i_ofv_liste else "Nei",
        "KjoptDato": kjopt_row["TransactionDate"] if kjopt_row else None,
        "AvregistrertDato": avreg_row["TransactionDate"] if avreg_row else None,
    }

    if avreg_row is not None:
        result["Avregistrert"] = "Ja"
    elif kjopt_row is None:
        result["Avregistrert"] = ""
    else:
        result["Avregistrert"] = "Nei"

    if er_i_ofv_liste and bokfort_date is None:
        result["Status"] = "Avvik: OFV viser kjop, mangler i bokforing"
    elif not er_i_ofv_liste and bokfort_date is not None:
        result["Status"] = "Avvik: Bokfort, ikke bekreftet kjopt av OFV i perioden"
    elif er_i_ofv_liste and bokfort_date is not None and result["Avregistrert"] == "Ja":
        result["Status"] = "OK - kjopt og avregistrert"
    elif er_i_ofv_liste and bokfort_date is not None:
        result["Status"] = "OK - kjopt, ikke avregistrert enna"
    else:
        result["Status"] = "Ingen treff"

    return result


def process_varekjop_bruktbil(df_input: pd.DataFrame, buyer_org_no: str,
                               date_fra: date, date_til: date, ofv_key: str):

    df_input = df_input.fillna("")

    print(f"OFV: henter kjopsliste for org {buyer_org_no} "
          f"({date_fra} - {date_til}) ...")
    ofv_liste = fetch_ofv_transactions_by_buyer_org(ofv_key, buyer_org_no, date_fra, date_til)
    print(f"OFV: {len(ofv_liste)} transaksjoner funnet for org {buyer_org_no} i perioden.")

    ofv_regnr = {normalize_identifier(r["RegNo"]) for r in ofv_liste if r.get("RegNo")}
    ofv_vin = {normalize_identifier(r["ChassisNumber"]) for r in ofv_liste if r.get("ChassisNumber")}

    all_rows = []
    kontroll_rows = []
    input_par = []
    total = len(df_input)

    for i, row in enumerate(df_input.itertuples(index=False), start=1):

        regnr_raw = str(getattr(row, "Regnr", "") or "").strip()
        vin_raw = str(getattr(row, "VIN", "") or "").strip()
        bokfort_raw = getattr(row, "BokfortDato", None)

        ident = normalize_identifier(regnr_raw) if regnr_raw else normalize_identifier(vin_raw)
        use_vin = is_vin(ident)
        regno = "" if use_vin else ident
        vin = ident if use_vin else ""
        input_par.append((normalize_identifier(regno), normalize_identifier(vin)))

        print(f"OFV: Eierskifter {ident}  ({i}/{total})  {round(i / total * 100)}%")

        tx_rows = fetch_ofv_transactions(ofv_key, ident, use_vin, regno, vin)
        all_rows.extend(tx_rows)

        er_i_ofv_liste = (
            (normalize_identifier(regno) in ofv_regnr and regno)
            or (normalize_identifier(vin) in ofv_vin and vin)
        )

        kontroll_rows.append(
            build_varekjop_row(regno, vin, bokfort_raw, tx_rows, er_i_ofv_liste, buyer_org_no))

        time.sleep(API_PAUSE_S)

    manglende = []
    for r in ofv_liste:
        reg_n = normalize_identifier(r.get("RegNo"))
        vin_n = normalize_identifier(r.get("ChassisNumber"))
        funnet = any(
            (reg_n and reg_n == ir) or (vin_n and vin_n == iv)
            for ir, iv in input_par
        )
        if not funnet:
            manglende.append({
                "Regnr": r.get("RegNo"),
                "Chassisnummer": r.get("ChassisNumber"),
                "KjoptDato": r.get("TransactionDate"),
                "Selger": compute_owner_label(
                    r.get("FromOwnerType"), r.get("FromOwnerCompanyName")),
            })

    resultat_df = build_resultat_dataframe(all_rows)
    kontroll_df = pd.DataFrame(kontroll_rows)
    manglende_df = pd.DataFrame(manglende)

    return resultat_df, kontroll_df, manglende_df

# COMMAND ----------

# ==============================================================
# KONTROLL 3: KONTROLL DEMOBIL
# ==============================================================


def build_demobil_row(regno: str, vin: str, bokfort_inn_raw, tx_rows: list,
                       company_org_no: str) -> dict:

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

        if r.get("TransactionDate") and (
                nyeste_row is None
                or r["TransactionDate"] > nyeste_row["TransactionDate"]):
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


def process_kontroll_demobil(df_input: pd.DataFrame, company_org_no: str, ofv_key: str):

    df_input = df_input.fillna("")
    all_rows = []
    kontroll_rows = []
    total = len(df_input)

    for i, row in enumerate(df_input.itertuples(index=False), start=1):

        regnr_raw = str(getattr(row, "Regnr", "") or "").strip()
        vin_raw = str(getattr(row, "VIN", "") or "").strip()
        bokfort_raw = getattr(row, "BokfortInnDato", None)

        ident = normalize_identifier(regnr_raw) if regnr_raw else normalize_identifier(vin_raw)
        use_vin = is_vin(ident)
        regno = "" if use_vin else ident
        vin = ident if use_vin else ""

        print(f"OFV: Eierskifter {ident}  ({i}/{total})  {round(i / total * 100)}%")

        tx_rows = fetch_ofv_transactions(ofv_key, ident, use_vin, regno, vin)
        all_rows.extend(tx_rows)

        kontroll_rows.append(build_demobil_row(regno, vin, bokfort_raw, tx_rows, company_org_no))

        time.sleep(API_PAUSE_S)

    resultat_df = build_resultat_dataframe(all_rows)
    kontroll_df = pd.DataFrame(kontroll_rows)

    return resultat_df, kontroll_df

# COMMAND ----------

# ==============================================================
# HOVEDKJORING - leser widgets, dispatcher til riktig kontroll(er)
# ==============================================================

kontroll_valg = dbutils.widgets.get("kontroll")[0]  # "1"/"2"/"3"/"4"
input_fil = dbutils.widgets.get("input_fil")
output_fil = dbutils.widgets.get("output_fil")
secret_scope = dbutils.widgets.get("secret_scope")

ofv_key = dbutils.secrets.get(scope=secret_scope, key="ofv-api-key")

try:
    svv_key = dbutils.secrets.get(scope=secret_scope, key="svv-api-key")
except Exception:
    svv_key = None  # SVV er kun reserve - mangler nokkelen, hoppes den bare over.

resultater = {}

if kontroll_valg in ("1", "4"):
    print("=== Kontroll solgte biler ===")
    df1 = pd.read_excel(input_fil, sheet_name="KontrollSolgteBiler")
    selger_org_no = dbutils.widgets.get("selger_org_no").strip()
    resultat1, kontroll1, detaljer1 = process_kontroll_solgte_biler(
        df1, selger_org_no, ofv_key, svv_key)
    resultater["Resultat"] = resultat1
    resultater["KontrollSolgteBiler"] = kontroll1
    resultater["KSB_OvrigeTransaksjoner"] = detaljer1

if kontroll_valg in ("2", "4"):
    print("=== Varekjop bruktbil ===")
    df2 = pd.read_excel(input_fil, sheet_name="VarekjopBruktbil")
    kjoper_org_no = dbutils.widgets.get("kjoper_org_no").strip()
    dato_fra = parse_bokfort_dato(dbutils.widgets.get("kjoper_dato_fra"))
    dato_til = parse_bokfort_dato(dbutils.widgets.get("kjoper_dato_til"))
    resultat2, kontroll2, manglende2 = process_varekjop_bruktbil(
        df2, kjoper_org_no, dato_fra, dato_til, ofv_key)
    resultater["ResultatVarekjop"] = resultat2
    resultater["KontrollVarekjopBruktbil"] = kontroll2
    resultater["Varekjop_ManglerIBokforing"] = manglende2

if kontroll_valg in ("3", "4"):
    print("=== Kontroll Demobil ===")
    df3 = pd.read_excel(input_fil, sheet_name="KontrollDemobil")
    demobil_org_no = dbutils.widgets.get("demobil_org_no").strip()
    resultat3, kontroll3 = process_kontroll_demobil(df3, demobil_org_no, ofv_key)
    resultater["ResultatDemobil"] = resultat3
    resultater["KontrollDemobil"] = kontroll3

for navn, df in resultater.items():
    print(f"\n=== {navn} ({len(df)} rader) ===")
    display(df)

# COMMAND ----------

# ==============================================================
# SKRIV RESULTATER TIL EN EXCEL-ARBEIDSBOK (med samme fargekoding
# som i VBA-versjonen)
# ==============================================================

with pd.ExcelWriter(output_fil, engine="openpyxl") as writer:
    for navn, df in resultater.items():
        df.to_excel(writer, sheet_name=navn[:31], index=False)

wb = load_workbook(output_fil)


def fargelegg_avvik_kolonne(ws, header_navn="DagerAvvik", selvhandel_navn="Selvhandel"):
    header = [c.value for c in ws[1]]
    if header_navn not in header:
        return
    col_avvik = header.index(header_navn) + 1
    col_selv = header.index(selvhandel_navn) + 1 if selvhandel_navn in header else None

    for row in ws.iter_rows(min_row=2):
        avvik_cell = row[col_avvik - 1]
        selv_cell = row[col_selv - 1] if col_selv else None

        if selv_cell is not None and selv_cell.value == "Ja":
            avvik_cell.fill = FILL_ROD
            selv_cell.fill = FILL_ROD
        elif isinstance(avvik_cell.value, (int, float)):
            if avvik_cell.value <= AVVIK_GRONN_MAX:
                avvik_cell.fill = FILL_GRONN
            elif avvik_cell.value <= AVVIK_GUL_MAX:
                avvik_cell.fill = FILL_GUL
            else:
                avvik_cell.fill = FILL_ROD


def fargelegg_status_kolonne(ws, header_navn="Status"):
    header = [c.value for c in ws[1]]
    if header_navn not in header:
        return
    col_status = header.index(header_navn) + 1

    for row in ws.iter_rows(min_row=2):
        cell = row[col_status - 1]
        tekst = str(cell.value or "")
        if tekst.startswith("OK"):
            cell.fill = FILL_GRONN
        elif tekst.startswith("Avvik"):
            cell.fill = FILL_ROD


if "KontrollSolgteBiler" in wb.sheetnames:
    fargelegg_avvik_kolonne(wb["KontrollSolgteBiler"])

for navn in ("KontrollVarekjopBruktbil", "KontrollDemobil"):
    if navn in wb.sheetnames:
        fargelegg_status_kolonne(wb[navn])

wb.save(output_fil)
print(f"\nFerdig. Resultater skrevet til: {output_fil}")
