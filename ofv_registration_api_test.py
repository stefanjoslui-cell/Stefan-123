#!/usr/bin/env python3
"""
Test program for OFV Sandbox Registrations API (v1).

Sandbox restrictions in effect (server-side, cannot be overridden):
  - Only electric passenger cars registered in the last 2 weeks in Oslo are returned.
  - Max 2 results per call.
  - Max 10 calls per day.

Usage (enklest):
    1. Fyll inn REGNR og API_KEY i feltet merket "FYLL INN HER" like under.
    2. Kjør: python ofv_registration_api_test.py

Usage (alternativt, uten å redigere filen):
    python ofv_registration_api_test.py --regnr AB12345 --api-key <key>
    python ofv_registration_api_test.py               # spør deg interaktivt

Output:
    An .xlsx workbook with one sheet per data section (Registration, History,
    EuCheckHistory, Variants, Equipment, Liens, Lienholders, Info), containing
    every column the API returned for the queried vehicle.
"""

import argparse
import getpass
import json
import sys
from pathlib import Path

import requests
from openpyxl import Workbook

# ============================================================
#  FYLL INN HER — regnr og API-nøkkel du har fått fra OFV:
#
REGNR = ""      # f.eks. "AB12345"
API_KEY = ""    # din "Ocp-Apim-Subscription-Key"
#
# ============================================================

API_URL = "https://api.ofv.no/registrations/v1/"
DEFAULT_OUTPUT_NAME = "OFV registration API test.xlsx"

# Fields that are lists-of-objects and are broken out into their own sheets
# instead of being flattened inline.
LIST_OF_OBJECT_FIELDS = {"history", "variants"}


def flatten(obj, parent_key="", sep="."):
    """Recursively flatten a dict into {dotted.key: scalar_value}.

    Lists of scalars are joined into a comma-separated string.
    Lists of dicts and nested dicts explicitly handled elsewhere by the
    caller should be excluded from `obj` before calling this.
    """
    items = {}
    if isinstance(obj, dict):
        for key, value in obj.items():
            new_key = f"{parent_key}{sep}{key}" if parent_key else key
            items.update(flatten(value, new_key, sep))
    elif isinstance(obj, list):
        if not obj:
            items[parent_key] = None
        elif all(not isinstance(v, (dict, list)) for v in obj):
            items[parent_key] = ", ".join(str(v) for v in obj if v is not None)
        else:
            # Fallback: list of objects we didn't special-case -> keep as JSON
            items[parent_key] = json.dumps(obj, ensure_ascii=False)
    else:
        items[parent_key] = obj
    return items


def write_sheet(wb, title, rows):
    """Write a list of flat dicts to a new sheet, unioning all column names."""
    ws = wb.create_sheet(title=title[:31])  # Excel sheet name limit
    if not rows:
        ws.append(["(ingen data)"])
        return
    columns = []
    seen = set()
    for row in rows:
        for key in row.keys():
            if key not in seen:
                seen.add(key)
                columns.append(key)
    ws.append(columns)
    for row in rows:
        ws.append([row.get(col) for col in columns])


def query_registration(reg_no, api_key):
    body = {
        "filters": {"regNo": reg_no},
        "pagination": {"first": 2},
        "sorting": {"orderBy": "regNo", "orderDirection": "ASC"},
        "include": {
            "variants": True,
            "equipment": True,
            "history": True,
            "demographics": True,
            "lien": True,
        },
    }
    headers = {
        "Ocp-Apim-Subscription-Key": api_key,
        "Content-Type": "application/json",
    }
    response = requests.post(API_URL, headers=headers, json=body, timeout=30)
    return response


def build_workbook(payload):
    wb = Workbook()
    wb.remove(wb.active)

    registrations = payload.get("registrations", [])

    info_rows = [
        {
            "requested_filters": json.dumps(payload.get("filters", {}), ensure_ascii=False),
            "forcedFilters": json.dumps(
                payload.get("filters", {}).get("forcedFilters", {}), ensure_ascii=False
            ),
            "pagination": json.dumps(payload.get("pagination", {}), ensure_ascii=False),
            "resultCount": len(registrations),
        }
    ]
    write_sheet(wb, "Info", info_rows)

    reg_rows, history_rows, eucheck_rows = [], [], []
    variant_rows, equipment_rows, lien_rows, lienholder_rows = [], [], [], []

    for reg in registrations:
        reg_no = reg.get("currentRegNo")

        reg_copy = dict(reg)
        history = reg_copy.pop("history", None) or []
        variants = reg_copy.pop("variants", None) or []
        eu_check = reg_copy.pop("EuCheck", None) or {}
        eu_check_history = eu_check.pop("history", None) or []
        lien_group = reg_copy.pop("lienGroup", None) or {}
        liens = lien_group.pop("liens", None) or []

        flat_reg = flatten(reg_copy)
        flat_reg.update({f"EuCheck.{k}": v for k, v in flatten(eu_check).items()})
        reg_rows.append(flat_reg)

        for entry in history:
            flat_entry = {"currentRegNo": reg_no}
            flat_entry.update(flatten(entry))
            history_rows.append(flat_entry)

        for entry in eu_check_history:
            flat_entry = {"currentRegNo": reg_no}
            flat_entry.update(flatten(entry))
            eucheck_rows.append(flat_entry)

        for variant in variants:
            vehicle = dict(variant.get("vehicle", {}))
            equipment = vehicle.pop("equipment", None) or []
            flat_variant = {"currentRegNo": reg_no, "matchScore": variant.get("matchScore")}
            flat_variant.update({f"vehicle.{k}": v for k, v in flatten(vehicle).items()})
            variant_rows.append(flat_variant)

            for item in equipment:
                flat_item = {
                    "currentRegNo": reg_no,
                    "variantId": vehicle.get("variantId"),
                }
                flat_item.update(flatten(item) if isinstance(item, dict) else {"equipment": item})
                equipment_rows.append(flat_item)

        for lien in liens:
            lienholders = lien.pop("lienholders", None) or []
            flat_lien = {"currentRegNo": reg_no}
            flat_lien.update(flatten(lien))
            lien_rows.append(flat_lien)

            for holder in lienholders:
                flat_holder = {"currentRegNo": reg_no}
                flat_holder.update(flatten(holder))
                lienholder_rows.append(flat_holder)

    write_sheet(wb, "Registration", reg_rows)
    write_sheet(wb, "History_Transactions", history_rows)
    write_sheet(wb, "EuCheckHistory", eucheck_rows)
    write_sheet(wb, "Variants", variant_rows)
    write_sheet(wb, "Equipment", equipment_rows)
    write_sheet(wb, "Liens", lien_rows)
    write_sheet(wb, "Lienholders", lienholder_rows)

    return wb


def main():
    parser = argparse.ArgumentParser(description="Test the OFV Sandbox Registrations API for one vehicle.")
    parser.add_argument("--regnr", help="Registreringsnummer (f.eks. AB12345)")
    parser.add_argument("--api-key", help="Ocp-Apim-Subscription-Key")
    parser.add_argument("--output", default=DEFAULT_OUTPUT_NAME, help="Filnavn for Excel-output")
    args = parser.parse_args()

    # Prioritet: kommandolinje-argument > verdi fylt inn i toppen av filen > interaktivt spørsmål
    reg_no = args.regnr or REGNR or input("Registreringsnummer: ").strip()
    api_key = args.api_key or API_KEY or getpass.getpass("API-nøkkel (Ocp-Apim-Subscription-Key): ").strip()

    if not reg_no or not api_key:
        print("Regnr og API-nøkkel er påkrevd.", file=sys.stderr)
        sys.exit(1)

    print(f"Spør OFV Sandbox Registrations API om {reg_no} ...")
    response = query_registration(reg_no, api_key)

    if response.status_code != 200:
        try:
            error = response.json()
        except ValueError:
            error = {"statusCode": response.status_code, "message": response.text}
        print(f"Feil ({response.status_code}): {error.get('message', error)}", file=sys.stderr)
        sys.exit(1)

    payload = response.json()
    registrations = payload.get("registrations", [])
    if not registrations:
        print(
            "Ingen resultater. Husk at sandbox-abonnementet kun returnerer elbiler "
            "personbiler registrert i Oslo de siste 2 ukene - regnr som ikke matcher "
            "disse filtrene vil gi et tomt resultat selv om bilen finnes."
        )
    else:
        print(f"Fikk {len(registrations)} resultat(er).")

    wb = build_workbook(payload)
    output_path = Path(args.output)
    wb.save(output_path)
    print(f"Lagret: {output_path.resolve()}")


if __name__ == "__main__":
    main()
