import csv
import io
import os

import paramiko
import streamlit as st

NIVEAUER = ["Enhed", "Kontor"]

LEDELSESLAG_PER_NIVEAU = {
    "Kontor": "Kontorleder",
}

ROOT_ID = "KU"
ROOT_NAVN = "Københavns Universitet"

INSTITUT_KONTOR_FIL = "institut_kontor.csv"
FORKORTELSER_FIL = os.path.join(os.path.dirname(__file__), "navne_til_forkortelse.csv")
ENCODING = "utf-8-sig"
# Kun disse tre ledes af en campusdirektør - resten (koncernenheder,
# Rektoratets Stab, Tilskud m.fl.) ledes af en vicedirektør. Bruges også i
# administrativt_omraade() til at afgøre, om et kontor-navnepræfiks skal
# tolkes som et CA-kontor (fx "HR Nørre").
CAMPUSDIREKTOER_ENHEDER = {
    "Campusadministration Frederiksberg+",
    "Campusadministration Nørre",
    "Campusadministration Søndre",
}


def administrativt_omraade(institut: str, kontor: str):
    """
    Samme logik som administrativt_omraade() i agg_data.py (som igen er en
    Python-oversættelse af det oprindelige R case_when) - holdes i sync
    manuelt, da scripts og app kører hver for sig. Dækker bevidst kun disse
    fem områder; alt andet får None (og udelades som valgmulighed i appen).
    """
    kontor = kontor or ""

    if institut == "KU HR":
        return "HR"
    if institut in CAMPUSDIREKTOER_ENHEDER and kontor.startswith("HR "):
        return "HR"

    if institut == "KU Bygninger":
        return "Bygninger"
    if institut in CAMPUSDIREKTOER_ENHEDER and kontor.startswith("Bygninger "):
        return "Bygninger"

    if institut == "KU Uddannelse":
        return "Uddannelse"
    if institut in CAMPUSDIREKTOER_ENHEDER and kontor.startswith("Uddannelse "):
        return "Uddannelse"

    if institut == "KU Økonomi":
        return "Økonomi"
    if institut in CAMPUSDIREKTOER_ENHEDER and kontor.startswith("Økonomi "):
        return "Økonomi"

    if institut == "KU IT":
        return "IT"
    if institut in CAMPUSDIREKTOER_ENHEDER and kontor == "IT-support":
        return "IT"

    return None

@st.cache_resource
def _get_sftp_client():
    """Genbruger samme forbindelsesmønster som publikationsappen."""
    creds = st.secrets["erda"]
    transport = paramiko.Transport((creds["host"], creds.get("port", 22)))
    transport.connect(username=creds["username"], password=creds["password"])
    return paramiko.SFTPClient.from_transport(transport)


@st.cache_data
def _load_csv_from_erda(filename: str) -> str:
    sftp = _get_sftp_client()
    path = f"{st.secrets['erda']['data_path']}/{filename}"
    with sftp.open(path) as f:
        raw = f.read()
    return raw.decode(ENCODING)

def _load_forkortelser(filename: str = FORKORTELSER_FIL):
    """
    Indlæser (Type, Navn) -> Forkortet navn fra navne_til_forkortelse.csv,
    som ligger lokalt i GitHub-repoet (samme mappe som denne fil) - IKKE
    på ERDA, da den ikke indeholder følsomme data, kun navne/forkortelser.

    Nøglen inkluderer Type, fordi samme navn kan optræde som både Enhed og
    Kontor (fx "KU Bygninger" er begge dele) med hver sin forkortelse.

    "NA" (eller tom) i Forkortet navn betyder, at enheden/kontoret skal
    UDELADES HELT fra data (ikke bare vise det fulde navn) - markeres her
    med None, og load_real_units() dropper så den række. Navne der slet
    ikke findes i filen, beholder deres fulde, oprindelige navn.

    Findes filen slet ikke (endnu ikke committet til repoet), returneres
    en tom mapping - appen virker stadig, bare uden forkortelser/udeladelser.
    """
    forkortelser = {}
    try:
        with open(filename, encoding=ENCODING, newline="") as f:
            tekst = f.read().lstrip("\ufeff")  # fjern ALLE indledende BOM'er, ikke kun én
        reader = csv.DictReader(io.StringIO(tekst), delimiter=";")
        for row in reader:
            type_ = (row.get("Type") or "").strip()
            navn = (row.get("Navn") or "").strip()
            kort = (row.get("Forkortet navn") or "").strip()
            if not navn:
                continue
            forkortelser[(type_, navn)] = None if (not kort or kort.upper() == "NA") else kort
    except FileNotFoundError:
        pass
    return forkortelser

def load_real_units(filename: str = INSTITUT_KONTOR_FIL):
    """
    Bygger den flade enheds-liste (samme form som den tidligere
    generate_dummy_units()) direkte ud fra institut_kontor.csv - Enhed
    (Institut) og Kontor kommer fra de institutter/kontorer, der reelt
    findes i data, ikke en hardkodet liste.

    Hver enhed har: id, navn, niveau, parent_id, ledelseslag, aarsvaerk,
    lonomkostninger. aarsvaerk/lonomkostninger er kun sat på Kontor-niveau
    (allerede aggregeret i institut_kontor.csv) - Enhed-niveauet summeres
    op i app.py (rollup()). Kontor-enheder har desuden et "omraade"-felt,
    beregnet med administrativt_omraade() ovenfor.
    """
    units = []
    enh_id_for_institut = {}
    forkortelser = _load_forkortelser()

    csv_tekst = _load_csv_from_erda(filename)
    reader = csv.DictReader(io.StringIO(csv_tekst), delimiter=";")
    for row in reader:
        institut = row["Institut"].strip()
        if institut == "Tilskud":
            continue
        if forkortelser.get(("Enhed", institut)) is None and ("Enhed", institut) in forkortelser:
            continue  # hele enheden er markeret NA - udelades fra data

        kontor = row["Kontor"].strip()
        kontor_kort = forkortelser.get(("Kontor", kontor), kontor)
        if kontor_kort is None:
            continue  # dette kontor er markeret NA - udelades fra data

        if institut not in enh_id_for_institut:
            enh_id = institut
            ledelseslag = "Campusdirektør" if institut in CAMPUSDIREKTOER_ENHEDER else "Vicedirektør"
            units.append({
                "id": enh_id,
                "navn": forkortelser.get(("Enhed", institut), institut),
                "niveau": "Enhed",
                "parent_id": ROOT_ID,
                "ledelseslag": ledelseslag,   
                "aarsvaerk": None,
                "lonomkostninger": None,
                "medarbejdere": None,
            })
            enh_id_for_institut[institut] = enh_id

        enh_id = enh_id_for_institut[institut]
        kontor_id = f"{enh_id}::{kontor}"
        omraade = administrativt_omraade(institut, kontor)

        units.append({
            "id": kontor_id,
            "navn": kontor_kort,
            "niveau": "Kontor",
            "parent_id": enh_id,
            "ledelseslag": LEDELSESLAG_PER_NIVEAU["Kontor"],
            "omraade": omraade,
            "aarsvaerk": float(row["antal_aarsvaerk"]),
            "lonomkostninger": float(row["lonomkostninger"]),
            "medarbejdere": float(row["antal_medarbejdere"]),
        })

    return units

