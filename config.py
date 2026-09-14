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

ADM_OMRAADER = [
    #"Økonomi", "Udd.adm.", "Kommunikation", "IT", 
    "Økonomi", "Uddannelsesadministration", "Kommunikation", "IT", 
    "Innovationsadministration", "HR", "Forskningsadministration",
    "Bygningsservice", "Stabe",
]

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
    Indlæser (Type, Navn) -> {"forkortet": ..., "omraade": ...} fra
    navne_til_forkortelse.csv, som ligger lokalt i GitHub-repoet (samme
    mappe som denne fil) - IKKE på ERDA, da den ikke indeholder følsomme
    data, kun navne/forkortelser/områder.

    Nøglen inkluderer Type, fordi samme navn kan optræde som både Enhed og
    Kontor (fx "KU Bygninger" er begge dele) med hver sin forkortelse.

    "NA" (eller tom) i "Forkortet navn" betyder, at enheden/kontoret skal
    UDELADES HELT fra data - markeres her med forkortet=None, og
    load_real_units() dropper så den række.

    "NA" (eller tom) i "Adm. område" betyder, at enheden/kontoret ikke har
    noget administrativt område - markeres her med omraade=None, men
    UDELUKKER IKKE rækken fra data (i modsætning til "Forkortet navn").

    Navne der slet ikke findes i filen, beholder deres fulde, oprindelige
    navn og får omraade=None.

    Findes filen slet ikke (endnu ikke committet til repoet), returneres
    en tom mapping - appen virker stadig, bare uden forkortelser/områder.
    """
    forkortelser = {}
    try:
        with open(filename, encoding=ENCODING, newline="") as f:
            tekst = f.read().lstrip("\ufeff")
        reader = csv.DictReader(io.StringIO(tekst), delimiter=";")
        for row in reader:
            type_ = (row.get("Type") or "").strip()
            navn = (row.get("Navn") or "").strip()
            kort = (row.get("Forkortet navn") or "").strip()
            omraade = (row.get("Adm. område") or "").strip()
            if not navn:
                continue
            forkortelser[(type_, navn)] = {
                "forkortet": None if (not kort or kort.upper() == "NA") else kort,
                "omraade": None if (not omraade or omraade.upper() == "NA") else omraade,
            }
    except FileNotFoundError:
        pass
    return forkortelser

def load_forkortelser_raw(filename: str = FORKORTELSER_FIL):
    """
    Læser navne_til_forkortelse.csv råt, som en liste af rækker - til
    visning i appens forklaring på, hvordan administrative områder er
    inddelt. IKKE til selve institut/kontor-opslaget (det er
    _load_forkortelser()'s job) - her vises filens indhold, som den er,
    inkl. evt. "NA"-tekst.
    """
    raekker = []
    try:
        with open(filename, encoding=ENCODING, newline="") as f:
            tekst = f.read().lstrip("\ufeff")
        reader = csv.DictReader(io.StringIO(tekst), delimiter=";")
        for row in reader:
            raekker.append({
                "Navn": (row.get("Navn") or "").strip(),
                "Forkortet navn": (row.get("Forkortet navn") or "").strip(),
                "Administrativt område": (row.get("Adm. område") or "").strip(),
            })
    except FileNotFoundError:
        pass
    return raekker

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
        enhed_info = forkortelser.get(("Enhed", institut))
        if enhed_info is not None and enhed_info["forkortet"] is None:
            continue  # hele enheden er markeret NA - udelades fra data

        kontor = row["Kontor"].strip()
        kontor_info = forkortelser.get(("Kontor", kontor))
        if kontor_info is not None and kontor_info["forkortet"] is None:
            continue  # dette kontor er markeret NA - udelades fra data
        kontor_kort = kontor_info["forkortet"] if kontor_info is not None else kontor
        omraade = kontor_info["omraade"] if kontor_info is not None else None

        if institut not in enh_id_for_institut:
            enh_id = institut
            ledelseslag = "Campusdirektør" if institut in CAMPUSDIREKTOER_ENHEDER else "Vicedirektør"
            enhed_navn = enhed_info["forkortet"] if enhed_info is not None else institut
            units.append({
                "id": enh_id,
                "navn": enhed_navn,
                "fuldt_navn": institut,
                "niveau": "Enhed",
                "parent_id": ROOT_ID,
                "ledelseslag": ledelseslag,   
                "aarsvaerk": None,
                "medarbejdere": None,
            })
            enh_id_for_institut[institut] = enh_id

        enh_id = enh_id_for_institut[institut]
        kontor_id = f"{enh_id}::{kontor}"

        units.append({
            "id": kontor_id,
            "navn": kontor_kort,
            "fuldt_navn": kontor,
            "niveau": "Kontor",
            "parent_id": enh_id,
            "ledelseslag": LEDELSESLAG_PER_NIVEAU["Kontor"],
            "omraade": omraade,
            "aarsvaerk": float(row["antal_aarsvaerk"]),
            "medarbejdere": float(row["antal_medarbejdere"]),
            "er_selvnavngivet": kontor == institut,
        })

    return units

