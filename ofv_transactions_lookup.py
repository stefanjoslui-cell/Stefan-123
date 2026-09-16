"""
OFV Transactions API - oppslag av registreringsdatoer for kjoretoy
====================================================================

Hva gjor dette programmet?
---------------------------
Du laster opp en CSV- eller Excel-fil med en kolonne som inneholder
VIN-numre (understellsnummer) og/eller norske registreringsnummer
(regnr) - de kan gjerne vaere blandet i samme kolonne.

For hver bil slar programmet opp mot OFV sitt Transactions-API
(https://data.ofv.no/api-details#api=transactions-api-v1) og henter:

  - forstegangsRegistreringsdato : nar bilen ble registrert forste gang
  - sisteEierskifteDato          : dato for siste eierskifte/omregistrering
  - sisteEierskifteIPeriode      : siste eierskifte innenfor en periode du
                                    selv angir nedenfor (f.eks. mellom
                                    20.12.2026 og 10.01.2027)

Resultatet skrives til en Excel-fil (.xlsx).

Hvordan bruke det?
-------------------
1. Fyll inn API-nokkelen din i feltet API_KEY nedenfor.
2. Juster INPUT_FILE, INPUT_COLUMN, PERIOD_FROM/PERIOD_TO og OUTPUT_FILE
   etter behov.
3. Installer avhengigheter en gang:  pip install -r requirements.txt
4. Kjor:  python ofv_transactions_lookup.py
"""

import sys
import time
from datetime import datetime

import pandas as pd
import requests

# =====================================================================
# 1) API-NOKKEL - SKRIV INN DIN OFV API-NOKKEL HER:
# =====================================================================
API_KEY = "SKRIV_INN_DIN_API_NOKKEL_HER"
# =====================================================================

# ---------------------------------------------------------------------
# 2) Inn- og utfil
# ---------------------------------------------------------------------
# Filen du vil laste opp. Kan vaere .csv, .xlsx eller .xls.
INPUT_FILE = "input.xlsx"

# Navnet pa kolonnen som inneholder VIN/regnr. Sett til None for at
# programmet skal forsoke a finne riktig kolonne automatisk.
INPUT_COLUMN = None

# Fila resultatet skrives til (Excel).
OUTPUT_FILE = "resultat.xlsx"

# ---------------------------------------------------------------------
# 3) Periode for "siste eierskifte i gitt periode"
#    Skriv datoene pa formatet DD.MM.YYYY
# ---------------------------------------------------------------------
PERIOD_FROM = "20.12.2026"
PERIOD_TO = "10.01.2027"

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


def norwegian_date_to_iso(date_str: str) -> str:
    """Konverterer DD.MM.YYYY til YYYY-MM-DD (ISO 8601) for APIet."""
    parsed = datetime.strptime(date_str.strip(), "%d.%m.%Y")
    return parsed.strftime("%Y-%m-%d")


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


def lookup_vehicle(session: requests.Session, identifier: str, period_from_iso: str, period_to_iso: str) -> dict:
    id_filter = build_identifier_filter(identifier)

    result = {
        "input": identifier,
        "identifiertype": "VIN" if "chassisNumber" in id_filter else "Regnr",
        "regNo": None,
        "chassisNumber": None,
        "makeName": None,
        "modelName": None,
        "forstegangsRegistreringsdato": None,
        "sisteEierskifteDato": None,
        f"sisteEierskifteIPeriode ({PERIOD_FROM}-{PERIOD_TO})": None,
        "status": "OK",
    }

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
        result["regNo"] = latest.get("regNo")
        result["chassisNumber"] = latest.get("chassisNumber")
        result["makeName"] = latest.get("makeName")
        result["modelName"] = latest.get("modelName")
        result["forstegangsRegistreringsdato"] = latest.get("firstRegistrationDate")
        result["sisteEierskifteDato"] = latest.get("transactionDate")

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
            result[f"sisteEierskifteIPeriode ({PERIOD_FROM}-{PERIOD_TO})"] = (
                period_transactions[0].get("transactionDate")
            )
        else:
            result[f"sisteEierskifteIPeriode ({PERIOD_FROM}-{PERIOD_TO})"] = "Ingen eierskifte i perioden"

    except RuntimeError as exc:
        result["status"] = f"Feil: {exc}"

    return result


def main() -> None:
    if not API_KEY or API_KEY == "SKRIV_INN_DIN_API_NOKKEL_HER":
        print("Du ma fylle inn API_KEY i toppen av scriptet for a kunne kjore det.")
        sys.exit(1)

    try:
        period_from_iso = norwegian_date_to_iso(PERIOD_FROM)
        period_to_iso = norwegian_date_to_iso(PERIOD_TO)
    except ValueError:
        print("PERIOD_FROM/PERIOD_TO ma vaere pa formatet DD.MM.YYYY, f.eks. 20.12.2026.")
        sys.exit(1)

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
        results.append(lookup_vehicle(session, identifier, period_from_iso, period_to_iso))
        time.sleep(PAUSE_BETWEEN_VEHICLES)

    result_df = pd.DataFrame(results)
    result_df.to_excel(OUTPUT_FILE, index=False)
    print(f"\nFerdig. Resultatet er lagret i '{OUTPUT_FILE}'.")


if __name__ == "__main__":
    main()
