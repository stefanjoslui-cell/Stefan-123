"""
OFV Transactions API - kun forstegangsregistreringsdato
====================================================================
Enkelt script: for hver bil (regnr eller VIN) hentes KUN
forstegangsregistreringsdato fra OFV sitt Transactions-API.

Viktig a vite om dette: Transactions-API har ingen egen "forstegangs-
registrering"-transaksjon - APIet dekker kun eierskifter. Men feltet
firstRegistrationDate folger med som en egenskap ved kjoretoyet pa
ENHVER transaksjon som finnes. Derfor henter scriptet bare siste
transaksjon (uansett dato) og leser av det feltet.

Konsekvens: har bilen ALDRI byttet eier (ingen transaksjon i det hele
tatt), finnes det ingen post a hente datoen fra, og resultatet blir
"Ingen treff" - det er en begrensning i APIet, ikke i scriptet.

Hvordan bruke det?
-------------------
1. Fyll inn API-nokkelen din i feltet API_KEY nedenfor.
2. Fyll inn regnr/VIN i TEST_IDENTIFIERS, kommaseparert.
3. Installer avhengigheter en gang:  pip install requests openpyxl
4. Kjor:  python ofv_forstegangsregistrering.py
"""

import sys
import time
from datetime import date, datetime

import openpyxl
import requests

# =====================================================================
# 1) API-NOKKEL - SKRIV INN DIN OFV API-NOKKEL HER:
# =====================================================================
API_KEY = "SKRIV_INN_DIN_API_NOKKEL_HER"
# =====================================================================

# ---------------------------------------------------------------------
# 2) Regnr/VIN du vil sla opp, kommaseparert.
# ---------------------------------------------------------------------
TEST_IDENTIFIERS = "EE96644, YV1UZA8VCN1234567"

# ---------------------------------------------------------------------
# 3) Utfil
# ---------------------------------------------------------------------
OUTPUT_FILE = "forstegangsregistrering.xlsx"

# ---------------------------------------------------------------------
# Teknisk konfigurasjon - trenger normalt ikke endres
# ---------------------------------------------------------------------
BASE_URL = "https://api.ofv.no/transactions/v1/"
REQUEST_TIMEOUT = 30
MAX_RETRIES = 4
RETRY_BACKOFF_SECONDS = 3
PAUSE_BETWEEN_REQUESTS = 0.15


def looks_like_vin(value: str) -> bool:
    """VIN/understellsnummer er 17 alfanumeriske tegn."""
    return len(value) == 17 and value.isalnum()


def parse_date(value):
    """Konverterer en API-dato (YYYY-MM-DD...) til et ekte dato-objekt."""
    if not value:
        return None
    try:
        return datetime.strptime(str(value)[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


def post_with_retries(session: requests.Session, payload: dict) -> dict:
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
            raise RuntimeError(f"403 Forbidden - kvote overskredet. Svar fra API: {response.text}")

        if response.status_code in (429, 500, 502, 503, 504):
            last_error = f"{response.status_code}: {response.text}"
            time.sleep(RETRY_BACKOFF_SECONDS * attempt)
            continue

        raise RuntimeError(f"{response.status_code}: {response.text}")

    raise RuntimeError(f"Ga opp etter {MAX_RETRIES} forsok. Siste feil: {last_error}")


def fetch_first_registration(session: requests.Session, identifier: str) -> dict:
    id_filter = {"chassisNumber": identifier} if looks_like_vin(identifier) else {"regNo": identifier}

    payload = {
        "filters": id_filter,
        "pagination": {"first": 1},
        "sorting": {"orderBy": "transactionDate", "orderDirection": "DESC"},
    }

    try:
        data = post_with_retries(session, payload)
    except RuntimeError as exc:
        return {"input": identifier, "forstegangsregistrering": None, "status": f"Feil: {exc}"}

    transactions = data.get("transactions", [])
    if not transactions:
        return {"input": identifier, "forstegangsregistrering": None, "status": "Ingen treff"}

    return {
        "input": identifier,
        "forstegangsregistrering": parse_date(transactions[0].get("firstRegistrationDate")),
        "status": "OK",
    }


def write_results_to_excel(results: list, path: str) -> None:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["Input", "ForstegangsRegistrering", "Status"])
    for row_index, result in enumerate(results, start=2):
        sheet.append([result["input"], result["forstegangsregistrering"], result["status"]])
        if isinstance(result["forstegangsregistrering"], date):
            sheet.cell(row=row_index, column=2).number_format = "DD.MM.YYYY"
    workbook.save(path)


def main() -> None:
    if not API_KEY or API_KEY == "SKRIV_INN_DIN_API_NOKKEL_HER":
        print("Du ma fylle inn API_KEY i toppen av scriptet for a kunne kjore det.")
        sys.exit(1)

    identifiers = [v.strip() for v in TEST_IDENTIFIERS.split(",") if v.strip()]
    if not identifiers:
        print("TEST_IDENTIFIERS er tom. Fyll inn minst ett regnr eller VIN.")
        sys.exit(1)

    session = requests.Session()
    session.headers.update({
        "Ocp-Apim-Subscription-Key": API_KEY,
        "Content-Type": "application/json",
    })

    results = []
    for index, identifier in enumerate(identifiers, start=1):
        print(f"[{index}/{len(identifiers)}] Slar opp {identifier} ...")
        results.append(fetch_first_registration(session, identifier))
        time.sleep(PAUSE_BETWEEN_REQUESTS)

    write_results_to_excel(results, OUTPUT_FILE)
    print(f"\nFerdig. Resultatet er lagret i '{OUTPUT_FILE}'.")


if __name__ == "__main__":
    main()
