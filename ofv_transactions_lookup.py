"""
OFV Transactions API - alle registreringer for kjoretoy i et datointervall
====================================================================

Hva gjor dette programmet?
---------------------------
Du laster opp en CSV- eller Excel-fil med en kolonne som inneholder
VIN-numre (understellsnummer) og/eller norske registreringsnummer
(regnr) - de kan gjerne vaere blandet i samme kolonne. (Eller du kan
skrive dem rett inn i TEST_IDENTIFIERS nedenfor, for testing.)

For hver bil slar programmet opp mot OFV sitt Transactions-API
(https://data.ofv.no/api-details#api=transactions-api-v1) og henter
ALLE registreringer/eierskifter som faller innenfor datointervallet du
angir i DATE_FROM/DATE_TO nedenfor. Har en bil flere registreringer i
intervallet, far den en egen rad per registrering i resultatet.

Resultatet skrives til en Excel-fil (.xlsx).

Hvordan bruke det?
-------------------
1. Fyll inn API-nokkelen din i feltet API_KEY nedenfor.
2. Fyll inn DATE_FROM og DATE_TO (format DDMMAAAA, f.eks. 20122026).
3. Juster INPUT_FILE/TEST_IDENTIFIERS og OUTPUT_FOLDER/OUTPUT_FILENAME
   etter behov.
4. Installer avhengigheter en gang:  pip install -r requirements.txt
5. Kjor:  python ofv_transactions_lookup.py
"""

import os
import sys
import time
from datetime import date, datetime

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
# 4) Datointervall (obligatorisk)
#    Alle registreringer for hver bil som faller innenfor dette
#    intervallet tas med, en rad per registrering.
#    Format: DDMMAAAA, f.eks. 20122026 for 20.12.2026.
# ---------------------------------------------------------------------
DATE_FROM = ""
DATE_TO = ""

# ---------------------------------------------------------------------
# Teknisk konfigurasjon - trenger normalt ikke endres
# ---------------------------------------------------------------------
BASE_URL = "https://api.ofv.no/transactions/v1/"
REQUEST_TIMEOUT = 30
MAX_RETRIES = 4
RETRY_BACKOFF_SECONDS = 3
PAUSE_BETWEEN_REQUESTS = 0.15  # unngar a hamre APIet
SORT_DIRECTION = "ASC"  # ASC = eldste registrering forst, DESC = nyeste forst

CANDIDATE_COLUMN_NAMES = [
    "vin", "vinnr", "vin-nr", "vinnummer",
    "understellsnummer", "understellsnr", "chassisnumber", "chassis",
    "regnr", "reg.nr", "reg nr", "regno", "reg_no",
    "registreringsnummer", "kjennemerke", "regnummer",
]

# Kolonner i resultatet, i den rekkefolgen de skal vises i Excel.
ROW_TEMPLATE_KEYS = [
    "input", "identifiertype", "regNo", "chassisNumber", "makeName", "modelName",
    "registrationType", "fuelGroup", "isLeased", "isUsedImported",
    "forstegangsRegistreringsdato", "transactionNumber", "eierskifteDato",
    "from_owner_type", "from_owner_companyName", "from_owner_countyName",
    "from_owner_municipalityName", "from_user_countyName",
    "to_owner_type", "to_owner_companyName", "status",
]


def ddmmyyyy_to_iso(date_str: str) -> str:
    """Konverterer DDMMAAAA (f.eks. 20122026) til YYYY-MM-DD for APIet."""
    parsed = datetime.strptime(date_str.strip(), "%d%m%Y")
    return parsed.strftime("%Y-%m-%d")


def parse_api_date(date_str):
    """Konverterer en API-dato (YYYY-MM-DD...) til et ekte dato-objekt, slik
    at Excel viser/sorterer den som en dato (formatert DD.MM.AAAA) i stedet
    for en tekststreng eller et tall uten ledende nuller."""
    if not date_str:
        return None
    try:
        return datetime.strptime(str(date_str)[:10], "%Y-%m-%d").date()
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
    Excel-skriver konverterer tallaktige tekststrenger til tall og fjerner
    ledende nuller. Dato-celler far et ekte dato-objekt med format
    DD.MM.AAAA, slik at Excel kan sortere/filtrere pa dem som datoer."""
    columns = list(results[0].keys()) if results else ROW_TEMPLATE_KEYS
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(columns)
    for row_index, row in enumerate(results, start=2):
        sheet.append([row.get(col) for col in columns])
        for col_index, col in enumerate(columns, start=1):
            if isinstance(row.get(col), date):
                sheet.cell(row=row_index, column=col_index).number_format = "DD.MM.YYYY"
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


def empty_row(identifier: str, id_filter: dict, status: str) -> dict:
    row = dict.fromkeys(ROW_TEMPLATE_KEYS)
    row["input"] = identifier
    row["identifiertype"] = "VIN" if "chassisNumber" in id_filter else "Regnr"
    row["status"] = status
    return row


def party_county(party):
    return (party or {}).get("countyName")


def party_municipality(party):
    return (party or {}).get("municipalityName")


def company_name(party):
    return ((party or {}).get("companyInfo") or {}).get("name")


def build_row_from_transaction(identifier: str, id_filter: dict, transaction: dict) -> dict:
    from_side = transaction.get("from") or {}
    to_side = transaction.get("to") or {}
    from_owner = from_side.get("owner") or {}
    from_user = from_side.get("user") or {}
    to_owner = to_side.get("owner") or {}

    row = empty_row(identifier, id_filter, "OK")
    row["regNo"] = transaction.get("regNo")
    row["chassisNumber"] = transaction.get("chassisNumber")
    row["makeName"] = transaction.get("makeName")
    row["modelName"] = transaction.get("modelName")
    row["registrationType"] = transaction.get("registrationType")
    row["fuelGroup"] = transaction.get("fuelGroup")
    row["isLeased"] = transaction.get("isLeased")
    row["isUsedImported"] = transaction.get("isUsedImported")
    row["forstegangsRegistreringsdato"] = parse_api_date(transaction.get("firstRegistrationDate"))
    row["transactionNumber"] = transaction.get("transactionNumber")
    row["eierskifteDato"] = parse_api_date(transaction.get("transactionDate"))
    row["from_owner_type"] = from_owner.get("type")
    row["from_owner_companyName"] = company_name(from_owner)
    row["from_owner_countyName"] = party_county(from_owner)
    row["from_owner_municipalityName"] = party_municipality(from_owner)
    row["from_user_countyName"] = party_county(from_user)
    row["to_owner_type"] = to_owner.get("type")
    row["to_owner_companyName"] = company_name(to_owner)
    return row


def fetch_all_transactions(session: requests.Session, id_filter: dict, date_from_iso: str, date_to_iso: str) -> list:
    """Henter ALLE transaksjoner for et kjoretoy innenfor datointervallet,
    med paginering (en bil har normalt fa treff, men vi handterer det uansett)."""
    transactions = []
    cursor = None
    while True:
        pagination = {"first": 1000}
        if cursor:
            pagination["cursor"] = cursor

        payload = {
            "filters": {
                **id_filter,
                "transactionDateFrom": date_from_iso,
                "transactionDateTo": date_to_iso,
            },
            "pagination": pagination,
            "sorting": {"orderBy": "transactionDate", "orderDirection": SORT_DIRECTION},
        }
        data = post_with_retries(session, payload)
        transactions.extend(data.get("transactions", []))

        page_info = data.get("pagination", {})
        if page_info.get("hasNextPage") and page_info.get("endCursor"):
            cursor = page_info["endCursor"]
            time.sleep(PAUSE_BETWEEN_REQUESTS)
        else:
            break

    return transactions


def lookup_vehicle_transactions(
    session: requests.Session,
    identifier: str,
    date_from_iso: str,
    date_to_iso: str,
) -> list:
    id_filter = build_identifier_filter(identifier)

    try:
        transactions = fetch_all_transactions(session, id_filter, date_from_iso, date_to_iso)
    except RuntimeError as exc:
        return [empty_row(identifier, id_filter, f"Feil: {exc}")]

    if not transactions:
        return [empty_row(identifier, id_filter, "Ingen registreringer i perioden")]

    return [build_row_from_transaction(identifier, id_filter, txn) for txn in transactions]


def main() -> None:
    if not API_KEY or API_KEY == "SKRIV_INN_DIN_API_NOKKEL_HER":
        print("Du ma fylle inn API_KEY i toppen av scriptet for a kunne kjore det.")
        sys.exit(1)

    if not DATE_FROM.strip() or not DATE_TO.strip():
        print("Du ma fylle inn DATE_FROM og DATE_TO (format DDMMAAAA, f.eks. 20122026).")
        sys.exit(1)

    try:
        date_from_iso = ddmmyyyy_to_iso(DATE_FROM)
        date_to_iso = ddmmyyyy_to_iso(DATE_TO)
    except ValueError:
        print("DATE_FROM/DATE_TO ma vaere pa formatet DDMMAAAA, f.eks. 20122026.")
        sys.exit(1)

    print(f"Datointervall: {DATE_FROM} - {DATE_TO}")

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
        rows = lookup_vehicle_transactions(session, identifier, date_from_iso, date_to_iso)
        results.extend(rows)
        print(f"    -> {len(rows)} rad(er)")
        time.sleep(PAUSE_BETWEEN_REQUESTS)

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    output_path = os.path.join(OUTPUT_FOLDER, OUTPUT_FILENAME)

    write_results_to_excel(results, output_path)
    print(f"\nFerdig. {len(results)} rad(er) for {len(identifiers)} kjoretoy lagret i '{output_path}'.")


if __name__ == "__main__":
    main()
