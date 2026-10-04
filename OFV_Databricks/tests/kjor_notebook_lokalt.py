"""Kjorer 02_OFV_API_Kontroller lokalt med falske OFV/SVV-API-er og testdata.

Lager inputfiler slik 01_Verifisering_Input gjor (inkl. Demobil med tomme
DatoFra/DatoTil, som utloste orgnr-feilen), stubber dbutils og requests, og
kjorer alle kodecellene. Krever pandas, openpyxl og requests.

Bruk: python tests/kjor_notebook_lokalt.py <notebook.ipynb> <arbeidsmappe> [--uten-varekjop]
"""
import json
import os
import random
import sys
from datetime import date, timedelta

import pandas as pd
import requests

nb_path, work = sys.argv[1], sys.argv[2]  # f.eks. notebooks/02_OFV_API_Kontroller.ipynb /tmp/ofv_test
uten_varekjop = "--uten-varekjop" in sys.argv
SELSKAP = "Test 1"
base = os.path.join(work, "OFV API")
inp = os.path.join(base, "Input", SELSKAP)
os.makedirs(inp, exist_ok=True)

rnd = random.Random(42)
PCSO = ("933749312", "Premium Cars Stor-Oslo AS")
TOYOTA = ("922754136", "Toyota Romerike AS")
BAUDA = ("950273038", "Bauda AS")
ANDRE = [("911111111", "Bilhuset AS"), ("922222222", "Motorsenteret AS"),
         ("933333333", "Leasing Norge AS"), ("944444444", "Bruktbil Oslo AS")]
MERKER = [("Polestar", "Polestar 3", "Elektrisitet"), ("Toyota", "RAV4", "Hybrid"),
          ("Zeekr", "Zeekr X", "Elektrisitet"), ("Volvo", "XC60", "Diesel"),
          ("Toyota", "Corolla", "Bensin"), ("BMW", "i4", "Elektrisitet")]


def owner(org=None, privat=False, fylke="Oslo"):
    if privat:
        return {"owner": {"type": "Privat", "countyName": fylke}}
    return {"owner": {"type": "Firma", "countyName": fylke,
                      "companyInfo": {"name": org[1], "organizationNumber": org[0]}}}


DB = {}  # ident -> list of tx (newest first)


def lag_bil(ident, hendelser, merke=None):
    m = merke or rnd.choice(MERKER)
    vin = ident if len(ident) == 17 else "YSM" + ident.rjust(14, "0")
    reg = ident if len(ident) != 17 else "EJ" + str(rnd.randint(10000, 99999))
    txs = []
    for i, (d, fra, til) in enumerate(hendelser):
        txs.append({"regNo": reg, "chassisNumber": vin, "makeName": m[0], "modelName": m[1],
                    "registrationType": "Eierskifte", "fuelGroup": m[2], "isLeased": False,
                    "isUsedImported": False, "firstRegistrationDate": "2022-03-01",
                    "transactionNumber": f"T{abs(hash((ident, i))) % 10**8}",
                    "transactionDate": d.isoformat(), "from": fra, "to": til})
    txs.sort(key=lambda t: t["transactionDate"], reverse=True)
    DB[reg] = txs
    DB[vin] = txs
    return reg, vin


# ---- Solgte biler ----
solgte = []
for i in range(40):
    bok = date(2025, 10, 1) + timedelta(days=rnd.randint(0, 60))
    if i < 5:
        solgte.append({"Regnr": f"ER{26000 + i}", "Vin": None, "Bokfort": pd.Timestamp(bok)})
        continue  # ingen OFV-treff -> SVV
    avvik = rnd.choice([0, 1, 2, 5, 9, 20, 45, 200])
    reg, _ = lag_bil(f"DS{22000 + i}", [
        (bok - timedelta(days=400), owner(rnd.choice(ANDRE)), owner(BAUDA)),
        (bok + timedelta(days=avvik), owner(BAUDA), owner(BAUDA) if i == 7 else owner(privat=True)),
    ])
    solgte.append({"Regnr": reg, "Vin": None, "Bokfort": pd.Timestamp(bok)})
solgte.append(dict(solgte[6]))  # duplikat-regnr (skal hentes fra buffer)

# ---- Varekjop ----
varekjop = []
for i in range(30):
    kjop = date(2025, 8, 1) + timedelta(days=rnd.randint(-60, 30))
    hend = [(kjop, owner(rnd.choice(ANDRE)) if i % 3 else owner(privat=True), owner(TOYOTA))]
    if i % 4 == 0:
        hend.append((kjop + timedelta(days=rnd.randint(10, 200)), owner(TOYOTA), owner(privat=True)))
    if i == 3:
        hend = [(date(2025, 5, 1), owner(privat=True), owner(TOYOTA)),
                (date(2025, 6, 1), owner(TOYOTA), owner(privat=True))]
    reg, _ = lag_bil(f"CV{91300 + i}", hend)
    varekjop.append({"Reg.nr.": reg})
# Biler OFV sier er kjopt i perioden, men som ikke er bokfort
for i in range(4):
    lag_bil(f"MB{10000 + i}", [(date(2025, 8, 10 + i), owner(privat=True), owner(TOYOTA))])

# ---- Demobil ----
demo = []
for i in range(25):
    inn = date(2024, 1, 1) + timedelta(days=rnd.randint(0, 600))
    hend = [(inn, owner(("999999999", "Importor AS")), owner(PCSO))]
    if i % 3 == 0:
        hend.append((inn + timedelta(days=rnd.randint(30, 300)), owner(PCSO), owner(privat=True)))
    if i == 1:
        hend.append((inn + timedelta(days=90), owner(PCSO), owner(rnd.choice(ANDRE))))
    ident = f"YSMYKEAE0RB{i:06d}" if i % 2 else f"EJ{44800 + i}"
    lag_bil(ident, hend, MERKER[0] if i == 0 else None)
    demo.append({"Regnr / VIN-nr": ident, "Avsk.hittil": pd.Timestamp(inn)})
demo.append({"Regnr / VIN-nr": "XX99999", "Avsk.hittil": pd.Timestamp(date(2024, 5, 1))})


# ---- Skriv inputfiler akkurat som 01_Verifisering_Input (Felt/Verdi som str) ----
def skriv(path, konfig, biler):
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        pd.DataFrame([{"Felt": k, "Verdi": v} for k, v in konfig.items()]).to_excel(w, sheet_name="Konfig", index=False)
        pd.DataFrame(biler).to_excel(w, sheet_name="Biler", index=False)


skriv(os.path.join(inp, "kontroll_solgte_biler_input.xlsx"), {"OrgNrSelger": "950273038"},
      [{"Identifikator": r["Regnr"], "BokfortDato": r["Bokfort"].date()} for r in solgte])
if not uten_varekjop:
    skriv(os.path.join(inp, "kontroll_varekjop_bruktbil_input.xlsx"),
          {"OrgNrKjoper": "922754136", "DatoFra": "01.08.2025", "DatoTil": "31.08.2025"},
          [{"Identifikator": r["Reg.nr."], "Liste": "Innkjop"} for r in varekjop]
          + [{"Identifikator": varekjop[0]["Reg.nr."], "Liste": "IB"},
             {"Identifikator": varekjop[1]["Reg.nr."], "Liste": "IB"},
             {"Identifikator": varekjop[1]["Reg.nr."], "Liste": "UB"},
             {"Identifikator": "ZZ12345", "Liste": "UB"}])
# Demobil: tomme DatoFra/DatoTil -> pandas leser Verdi som float (feilen)
skriv(os.path.join(inp, "kontroll_demobil_input.xlsx"), {"OrgNr": "933749312", "DatoFra": "", "DatoTil": ""},
      [{"Identifikator": r["Regnr / VIN-nr"], "BokfortInnDato": r["Avsk.hittil"].date()} for r in demo])


# ---- Falske HTTP-svar ----
class Resp:
    def __init__(self, code, data):
        self.status_code, self._d, self.text = code, data, json.dumps(data)

    def json(self):
        return self._d


class FakeSession:
    def post(self, url, headers=None, json=None, timeout=None):
        f = json["filters"]
        if "toOrganizationNumber" in f:
            fra, til = f["transactionDateFrom"], f["transactionDateTo"]
            seen, out = set(), []
            for txs in DB.values():
                for t in txs:
                    if id(t) in seen:
                        continue
                    seen.add(id(t))
                    org = ((t["to"].get("owner") or {}).get("companyInfo") or {}).get("organizationNumber")
                    if org == f["toOrganizationNumber"] and fra <= t["transactionDate"] <= til:
                        out.append(t)
            return Resp(200, {"transactions": out, "pagination": {"hasNextPage": False}})
        ident = f.get("regNo") or f.get("chassisNumber")
        return Resp(200, {"transactions": DB.get(ident, []), "pagination": {"hasNextPage": False}})

    def get(self, url, params=None, headers=None, timeout=None):
        return Resp(200, {"kjoretoydataListe": [{"forstegangsregistrering": {"registrertForstegangNorgeDato": "2025-09-28"}}]})


class Widgets:
    def text(self, *a, **k):
        pass

    def get(self, name):
        return SELSKAP


class Secrets:
    def get(self, scope, key):
        return "fake-" + key


class DBUtils:
    widgets, secrets = Widgets(), Secrets()


requests.Session = FakeSession
import time as _t
_t.sleep = lambda s: None

nb = json.load(open(nb_path, encoding="utf-8"))
ns = {"dbutils": DBUtils(), "__name__": "__main__"}
for c in nb["cells"]:
    if c["cell_type"] != "code":
        continue
    src = "".join(c["source"])
    if src.lstrip().startswith("%pip"):
        continue
    src = src.replace('"/Workspace/Users/stefan.luidold@bdo.no/OFV API"', repr(base))
    exec(compile(src, "<cell>", "exec"), ns)

print("HTML:", ns["OUTPUT_HTML"])
print("Demobil orgnr:", repr(ns["resultat_demobil"]["orgnr"]))
