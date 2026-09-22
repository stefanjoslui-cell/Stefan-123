#!/usr/bin/env python3
"""
Henter en TILFELDIG registrert bil fra siste uke via OFV Sandbox Registrations
API, og lagrer ALL informasjon API et tilbyr om denne ene bilen til en Excel-fil.

Gjenbruker flatting/Excel-logikken fra ofv_registration_api_test.py, som må
ligge i samme mappe som denne filen.

Viktig om sandbox-abonnementet (server-side, kan ikke overstyres):
  - Returnerer kun elektriske personbiler registrert i Oslo, uansett hva vi ber om.
  - Maks 2 resultater per kall, maks 10 kall per dag.
  - "Siste uke" her er en innsnevring av datofilteret INNENFOR det tvungne
    2-ukers-vinduet (tillatt - datofiltre kan gjøres strammere, ikke videre).

Hvert kall til / -endepunktet bruker en av de 10 daglige kallene. Skriptet gjør
som standard ETT kall (PAGES = 1) og trekker tilfeldig blant de (maks 2) bilene
det kallet returnerer. Sett PAGES høyere for å trekke blant flere kandidater -
det bruker da flere av dagens 10 kall (én side = ett kall).

Usage (enklest):
    1. Fyll inn API_KEY i feltet merket "FYLL INN HER" like under.
    2. Kjør: python ofv_random_car_test.py

Usage (alternativt, uten å redigere filen):
    python ofv_random_car_test.py --api-key <key> --days 7 --pages 1
"""

import argparse
import getpass
import random
import sys
from datetime import date, timedelta
from pathlib import Path

import requests

from ofv_registration_api_test import API_URL, build_workbook

# ============================================================
#  FYLL INN HER — API-nøkkel du har fått fra OFV:
#
API_KEY = ""    # din "Ocp-Apim-Subscription-Key"
#
# ============================================================

DEFAULT_OUTPUT_NAME = "OFV random car test.xlsx"
DEFAULT_DAYS_BACK = 7
DEFAULT_PAGES = 1

# Event-typer som regnes som "førstegangsregistrering" (se RegistrationEventType).
FIRST_REGISTRATION_EVENT_TYPES = [
    "Førstegangsregistrerte nye",
    "Førstegangsregistrerte nye (Bruktimportert)",
    "Førstegangsregistrerte nye (Gjenoppbygget)",
]


def fetch_candidates(api_key, days_back, pages, new_only):
    """Spør inntil `pages` sider (maks 2 biler per side) om registreringer de
    siste `days_back` dagene, og returner alle raw registration-dicts funnet,
    pluss siste payload (for filters/pagination-metadata)."""
    date_from = (date.today() - timedelta(days=days_back)).isoformat()
    date_to = date.today().isoformat()

    filters = {
        "registrationDateFrom": date_from,
        "registrationDateTo": date_to,
    }
    if new_only:
        filters["historyEventType"] = FIRST_REGISTRATION_EVENT_TYPES

    headers = {
        "Ocp-Apim-Subscription-Key": api_key,
        "Content-Type": "application/json",
    }

    candidates = []
    cursor = None
    last_payload = None

    for page_num in range(1, pages + 1):
        pagination = {"first": 2}
        if cursor:
            pagination["cursor"] = cursor

        body = {
            "filters": filters,
            "pagination": pagination,
            "sorting": {"orderBy": "regNo", "orderDirection": "ASC"},
            "include": {
                "variants": True,
                "equipment": True,
                "history": True,
                "demographics": True,
                "lien": True,
            },
        }
        response = requests.post(API_URL, headers=headers, json=body, timeout=30)
        if response.status_code != 200:
            try:
                error = response.json()
            except ValueError:
                error = {"statusCode": response.status_code, "message": response.text}
            print(
                f"Feil ({response.status_code}) på side {page_num}: {error.get('message', error)}",
                file=sys.stderr,
            )
            break

        payload = response.json()
        last_payload = payload
        page_registrations = payload.get("registrations", [])
        candidates.extend(page_registrations)
        print(f"Side {page_num}: fikk {len(page_registrations)} bil(er).")

        page_info = payload.get("pagination", {})
        if not page_info.get("hasNextPage") or not page_info.get("endCursor"):
            break
        cursor = page_info["endCursor"]

    return candidates, last_payload


def main():
    parser = argparse.ArgumentParser(
        description="Hent en tilfeldig registrert bil fra siste uke (OFV Sandbox Registrations API)."
    )
    parser.add_argument("--api-key", help="Ocp-Apim-Subscription-Key")
    parser.add_argument(
        "--days", type=int, default=DEFAULT_DAYS_BACK,
        help="Antall dager tilbake å søke i (default: 7)",
    )
    parser.add_argument(
        "--pages", type=int, default=DEFAULT_PAGES,
        help="Antall sider/kall å hente, maks 2 biler per side (default: 1). "
             "Hvert kall bruker av dagskvoten på 10.",
    )
    parser.add_argument(
        "--new-only", action="store_true",
        help="Kun tell førstegangsregistreringer (ikke eierskifter/omregistreringer) som kandidater.",
    )
    parser.add_argument("--output", default=DEFAULT_OUTPUT_NAME, help="Filnavn for Excel-output")
    args = parser.parse_args()

    api_key = args.api_key or API_KEY or getpass.getpass("API-nøkkel (Ocp-Apim-Subscription-Key): ").strip()
    if not api_key:
        print("API-nøkkel er påkrevd.", file=sys.stderr)
        sys.exit(1)

    print(f"Søker etter biler registrert de siste {args.days} dagene (maks {args.pages} kall)...")
    candidates, last_payload = fetch_candidates(api_key, args.days, args.pages, args.new_only)

    if not candidates:
        print(
            "Fant ingen kandidater i perioden. Husk at sandbox-abonnementet uansett kun "
            "returnerer elektriske personbiler registrert i Oslo de siste 2 ukene - "
            "utenfor det vinduet vil alltid gi tomt resultat."
        )
        sys.exit(1)

    chosen = random.choice(candidates)
    print(f"Trakk tilfeldig blant {len(candidates)} kandidat(er): regNo = {chosen.get('currentRegNo')}")

    payload = {
        "filters": (last_payload or {}).get("filters", {}),
        "pagination": (last_payload or {}).get("pagination", {}),
        "registrations": [chosen],
    }

    wb = build_workbook(payload)
    output_path = Path(args.output)
    wb.save(output_path)
    print(f"Lagret: {output_path.resolve()}")


if __name__ == "__main__":
    main()
