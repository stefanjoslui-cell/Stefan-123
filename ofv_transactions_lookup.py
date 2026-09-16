"""
OFV Transactions API - oppslag av registreringsdatoer for kjoretoy
====================================================================

Hva gjor dette programmet?
---------------------------
Du laster opp en CSV- eller Excel-fil med en kolonne som inneholder
VIN-numre (understellsnummer) og/eller norske registreringsnummer
(regnr) - de kan gjerne vaere blandet i samme kolonne. (Eller du kan
skrive dem rett inn i TEST_IDENTIFIERS nedenfor, for testing.)

For hver bil slar programmet opp mot OFV sitt Transactions-API
(https://data.ofv.no/api-details#api=transactions-api-v1) og henter
kjoretoydata, eierskiftedatoer og geografi/eierinfo for selger (from)
og kjoper (to). Se kolonneoversikt i build_empty_result() nedenfor.

Resultatet skrives til en Excel-fil (.xlsx).

Hvordan bruke det?
-------------------
1. Fyll inn API-nokkelen din i feltet API_KEY nedenfor.
2. Juster INPUT_FILE/TEST_IDENTIFIERS, PERIOD_FROM/PERIOD_TO og
   OUTPUT_FOLDER/OUTPUT_FILENAME etter behov.
3. Installer avhengigheter en gang:  pip install -r requirements.txt
4. Kjor:  python ofv_transactions_lookup.py
"""

import os
import sys
import time
from datetime import datetime

import openpyxl
import pandas as pd
import requests

# =====================================================================
# 1) API-NOKKEL - SKRIV INN DIN OFV API-NOKKEL HER:
# =====================================================================
API_KEY = "SKRIV_INN_DIN_API_NOKKEL_HER"
# =====================================================================

# ---------------------------------------------------------------------
# 2) Inn-data
# ---------------------------------------------------------------------
# FOR TEST: skriv inn regnr/VIN direkte her, med komma mellom hvert.
# F.eks: TEST_IDENTIFIERS = "EE96644, AB12345, YV1UZA8VCN1234567"
# Sa lenge denne IKKE er tom, brukes den i stedet for INPUT_FILE nedenfor.
TEST_IDENTIFIERS = ""

# Filen du vil laste opp. Kan vaere .csv, .xlsx eller .xls.
# (Brukes bare nar TEST_IDENTIFIERS ovenfor er tom.)
INPUT_FILE = "input.xlsx"

# Navnet pa kolonnen som inneholder VIN/regnr. Sett til None for at
# programmet skal forsoke a finne riktig kolonne automatisk.
INPUT_COLUMN = None

# ---------------------------------------------------------------------
# 3) Utfil
# ---------------------------------------------------------------------
OUTPUT_FOLDER = r"C:\Users\cn6971\OneDrive - BDO AS\100. Utvikling\OFV API\Output py. test"
OUTPUT_FILENAME = "resultat.xlsx"

# ---------------------------------------------------------------------
# 4) Periode for "siste eierskifte i gitt periode" (valgfritt)
#    Format: DDMMAAAA, f.eks. 20122026 for 20.12.2026.
#    La begge sta tomme ("") for a IKKE bruke periodefilteret.
# ---------------------------------------------------------------------
PERIOD_FROM = ""
PERIOD_TO = ""

# ---------------------------------------------------------------------
# Teknisk konfigurasjon - trenger normalt ikke endres
# ---------------------------------------------------------------------
BASE_URL = "https://api.ofv.no/transactions/v1/"
REQUEST_TIMEOUT = 30
MAX_RETRIES = 4
RETRY_BACKOFF_SECONDS = 3
PAUSE_BETWEEN_VEHICLES = 0.15  # unngar a hamre APIet

CANDIDATE_COLUMN_NAMES = [
    "vin", "vinnr", "vin-nr", "vinnummer",
    "understellsnummer", "understellsnr", "chassisnumber", "chassis",
    "regnr", "reg.nr", "reg nr", "regno", "reg_no",
    "registreringsnummer", "kjennemerke", "regnummer",
]


def ddmmyyyy_to_iso(date_str: str) -> str:
    """Konverterer DDMMAAAA (f.eks. 20122026) til YYYY-MM-DD for APIet."""
    parsed = datetime.strptime(date_str.strip(), "%d%m%Y")
    return parsed.strftime("%Y-%m-%d")


def iso_to_ddmmyyyy(date_str):
    """Konverterer en API-dato (YYYY-MM-DD...) til DDMMAAAA for output."""
    if not date_str:
        return None
    try:
        return datetime.strptime(str(date_str)[:10], "%Y-%m-%d").strftime("%d%m%Y")
    except ValueError:
        return date_str


def looks_like_vin(value: str) -> bool:
    """VIN/understellsnummer er 17 alfanumeriske tegn."""
    return len(value) == 17 and value.isalnum()


def detect_input_column(df: pd.DataFrame) -> str:
    if INPUT_COLUMN:
        if INPUT_COLUMN not in df.columns:
            raise ValueError(
                f"Fant ikke kolonnen '{INPUT_COLUMN}' i {INPUT_FILE}. "
                f"Tilgjengelige kolonner: {list(df.columns)}"
            )
        return INPUT_COLUMN

    lowered = {str(col).strip().lower(): col for col in df.columns}
    for candidate in CANDIDATE_COLUMN_NAMES:
        if candidate in lowered:
            return lowered[candidate]

    # Fant ingen kjent kolonnenavn - bruk forste kolonne som fallback.
    first_col = df.columns[0]
    print(
        f"Fant ingen kolonne med kjent navn (vin/regnr/...). "
        f"Bruker forste kolonne '{first_col}'. "
        f"Sett INPUT_COLUMN i toppen av scriptet hvis dette er feil."
    )
    return first_col


def read_input_file() -> pd.DataFrame:
    if INPUT_FILE.lower().endswith(".csv"):
        return pd.read_csv(INPUT_FILE, dtype=str)
    return pd.read_excel(INPUT_FILE, dtype=str)


def write_results_to_excel(results: list, path: str) -> None:
    """Skriver rett med openpyxl (ikke pandas.to_excel), fordi pandas sin
    Excel-skriver konverterer tallaktige tekststrenger (f.eks. datoer i
    DDMMAAAA-format) til tall og fjerner ledende nuller."""
    columns = list(results[0].keys()) if results else []
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(columns)
    for row in results:
        sheet.append([row.get(col) for col in columns])
    workbook.save(path)


def post_with_retries(session: requests.Session, payload: dict) -> dict:
    """POST mot Transactions-API med retry pa 429/5xx."""
    last_error = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = session.post(BASE_URL, json=payload, timeout=REQUEST_TIMEOUT)
        except requests.RequestException as exc:
            last_error = str(exc)
            time.sleep(RETRY_BACKOFF_SECONDS * attempt)
            continue

        if response.status_code == 200:
            return response.json()

        if response.status_code == 401:
            raise RuntimeError(
                "401 Unauthorized - API-nokkelen er ugyldig eller mangler. "
                "Sjekk API_KEY i toppen av scriptet."
            )

        if response.status_code == 403:
            raise RuntimeError(
                f"403 Forbidden - kvote overskredet. Svar fra API: {response.text}"
            )

        if response.status_code in (429, 500, 502, 503, 504):
            last_error = f"{response.status_code}: {response.text}"
            time.sleep(RETRY_BACKOFF_SECONDS * attempt)
            continue

        # Andre feil (400 osv.) - gi opp umiddelbart med detaljer.
        raise RuntimeError(f"{response.status_code}: {response.text}")

    raise RuntimeError(f"Ga opp etter {MAX_RETRIES} forsok. Siste feil: {last_error}")


def build_identifier_filter(identifier: str) -> dict:
    if looks_like_vin(identifier):
        return {"chassisNumber": identifier}
    return {"regNo": identifier}


def build_empty_result(identifier: str, id_filter: dict, period_active: bool) -> dict:
    """Alle kolonner i output, i den rekkefolgen de skal vises i Excel."""
    result = {
        # --- identifikasjon ---
        "input": identifier,
        "identifiertype": "VIN" if "chassisNumber" in id_filter else "Regnr",
        "regNo": None,
        "chassisNumber": None,
        "makeName": None,
        "modelName": None,
        # --- kjoretoyattributter ---
        "registrationType": None,
        "fuelGroup": None,
        "isLeased": None,
        "isUsedImported": None,
        # --- datoer (format DDMMAAAA) ---
        "forstegangsRegistreringsdato": None,
        "sisteEierskifteDato": None,
        # --- selger (from) ---
        "from_owner_type": None,
        "from_owner_companyName": None,
        "from_owner_countyName": None,
        "from_owner_municipalityName": None,
        "from_user_countyName": None,
        "from_user_municipalityName": None,
        "from_leaseHolder_countyName": None,
        "from_leaseHolder_municipalityName": None,
        # --- kjoper (to) ---
        "to_owner_type": None,
        "to_owner_companyName": None,
        # --- status ---
        "status": "OK",
    }
    if period_active:
        result["sisteEierskifteIPeriode"] = None
    return result


def party_county(party):
    return (party or {}).get("countyName")


def party_municipality(party):
    return (party or {}).get("municipalityName")


def company_name(party):
    return ((party or {}).get("companyInfo") or {}).get("name")


def lookup_vehicle(
    session: requests.Session,
    identifier: str,
    period_active: bool,
    period_from_iso: str,
    period_to_iso: str,
) -> dict:
    id_filter = build_identifier_filter(identifier)
    result = build_empty_result(identifier, id_filter, period_active)

    try:
        overall_payload = {
            "filters": dict(id_filter),
            "pagination": {"first": 1},
            "sorting": {"orderBy": "transactionDate", "orderDirection": "DESC"},
        }
        overall_data = post_with_retries(session, overall_payload)
        transactions = overall_data.get("transactions", [])

        if not transactions:
            result["status"] = "Ingen treff"
            return result

        latest = transactions[0]
        from_side = latest.get("from") or {}
        to_side = latest.get("to") or {}
        from_owner = from_side.get("owner") or {}
        from_user = from_side.get("user") or {}
        from_lease = from_side.get("leaseHolder") or {}
        to_owner = to_side.get("owner") or {}

        result["regNo"] = latest.get("regNo")
        result["chassisNumber"] = latest.get("chassisNumber")
        result["makeName"] = latest.get("makeName")
        result["modelName"] = latest.get("modelName")

        result["registrationType"] = latest.get("registrationType")
        result["fuelGroup"] = latest.get("fuelGroup")
        result["isLeased"] = latest.get("isLeased")
        result["isUsedImported"] = latest.get("isUsedImported")

        result["forstegangsRegistreringsdato"] = iso_to_ddmmyyyy(latest.get("firstRegistrationDate"))
        result["sisteEierskifteDato"] = iso_to_ddmmyyyy(latest.get("transactionDate"))

        result["from_owner_type"] = from_owner.get("type")
        result["from_owner_companyName"] = company_name(from_owner)
        result["from_owner_countyName"] = party_county(from_owner)
        result["from_owner_municipalityName"] = party_municipality(from_owner)
        result["from_user_countyName"] = party_county(from_user)
        result["from_user_municipalityName"] = party_municipality(from_user)
        result["from_leaseHolder_countyName"] = party_county(from_lease)
        result["from_leaseHolder_municipalityName"] = party_municipality(from_lease)

        result["to_owner_type"] = to_owner.get("type")
        result["to_owner_companyName"] = company_name(to_owner)

        if period_active:
            period_payload = {
                "filters": {
                    **id_filter,
                    "transactionDateFrom": period_from_iso,
                    "transactionDateTo": period_to_iso,
                },
                "pagination": {"first": 1},
                "sorting": {"orderBy": "transactionDate", "orderDirection": "DESC"},
            }
            period_data = post_with_retries(session, period_payload)
            period_transactions = period_data.get("transactions", [])
            if period_transactions:
                result["sisteEierskifteIPeriode"] = iso_to_ddmmyyyy(
                    period_transactions[0].get("transactionDate")
                )
            else:
                result["sisteEierskifteIPeriode"] = "Ingen eierskifte i perioden"

    except RuntimeError as exc:
        result["status"] = f"Feil: {exc}"

    return result


def main() -> None:
    if not API_KEY or API_KEY == "SKRIV_INN_DIN_API_NOKKEL_HER":
        print("Du ma fylle inn API_KEY i toppen av scriptet for a kunne kjore det.")
        sys.exit(1)

    period_active = bool(PERIOD_FROM.strip()) and bool(PERIOD_TO.strip())
    period_from_iso = period_to_iso = None
    if period_active:
        try:
            period_from_iso = ddmmyyyy_to_iso(PERIOD_FROM)
            period_to_iso = ddmmyyyy_to_iso(PERIOD_TO)
        except ValueError:
            print("PERIOD_FROM/PERIOD_TO ma vaere pa formatet DDMMAAAA, f.eks. 20122026.")
            sys.exit(1)
        print(f"Periodefilter aktivt: {PERIOD_FROM} - {PERIOD_TO}")
    else:
        print("Periodefilter er ikke i bruk (PERIOD_FROM/PERIOD_TO er tomme).")

    if TEST_IDENTIFIERS.strip():
        identifiers = [v.strip() for v in TEST_IDENTIFIERS.split(",") if v.strip()]
        if not identifiers:
            print("TEST_IDENTIFIERS er fylt ut, men inneholder ingen gyldige verdier.")
            sys.exit(1)
        print(f"Bruker {len(identifiers)} regnr/VIN fra TEST_IDENTIFIERS. Starter oppslag...")
    else:
        try:
            df = read_input_file()
        except FileNotFoundError:
            print(f"Fant ikke inputfilen '{INPUT_FILE}'. Sjekk INPUT_FILE i toppen av scriptet.")
            sys.exit(1)

        if df.empty:
            print(f"Inputfilen '{INPUT_FILE}' er tom.")
            sys.exit(1)

        id_column = detect_input_column(df)
        identifiers = [str(v).strip() for v in df[id_column].dropna().tolist() if str(v).strip()]

        if not identifiers:
            print(f"Fant ingen verdier i kolonnen '{id_column}'.")
            sys.exit(1)

        print(f"Fant {len(identifiers)} kjoretoy i kolonnen '{id_column}'. Starter oppslag...")

    session = requests.Session()
    session.headers.update({
        "Ocp-Apim-Subscription-Key": API_KEY,
        "Content-Type": "application/json",
    })

    results = []
    for index, identifier in enumerate(identifiers, start=1):
        print(f"[{index}/{len(identifiers)}] Slar opp {identifier} ...")
        results.append(
            lookup_vehicle(session, identifier, period_active, period_from_iso, period_to_iso)
        )
        time.sleep(PAUSE_BETWEEN_VEHICLES)

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    output_path = os.path.join(OUTPUT_FOLDER, OUTPUT_FILENAME)

    write_results_to_excel(results, output_path)
    print(f"\nFerdig. Resultatet er lagret i '{output_path}'.")


if __name__ == "__main__":
    main()
