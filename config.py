import random

NIVEAUER = ["Enhed", "Kontor"]

LEDELSESLAG_PER_NIVEAU = {
    "Kontor": "Kontorleder",
}

ROOT_ID = "KU"
ROOT_NAVN = "Københavns Universitet"

# De 3 campusadministrationer (CA) + 8 koncernenheder (KE) + Rektoratets
# Stab, jf. https://om.ku.dk/organisation/administration/ - (navn,
# ledelseslag)-par, da campusadministrationerne ledes af en campusdirektør,
# mens koncernenhederne og Rektoratets Stab hver ledes af en vicedirektør.
ENHEDER = [
    ("Campusadministration Frederiksberg+", "Campusdirektør"),
    ("Campusadministration Nørre", "Campusdirektør"),
    ("Campusadministration Søndre", "Campusdirektør"),
    ("KU Bygninger", "Vicedirektør"),
    ("KU Forskning og Informationssikkerhed", "Vicedirektør"),
    ("KU HR", "Vicedirektør"),
    ("KU Innovation og Erhvervssamarbejde", "Vicedirektør"),
    ("KU IT", "Vicedirektør"),
    ("KU Kommunikation", "Vicedirektør"),
    ("KU Uddannelse", "Vicedirektør"),
    ("KU Økonomi", "Vicedirektør"),
    ("Rektoratets Stab", "Vicedirektør"),
]

# Kendte, rigtige kontornavne pr. enhed, hentet fra hver enheds egen
# underside på om.ku.dk/organisation/administration/. Enheder der ikke er
# nævnt her (hvis flere kommer til senere) får dummy-genererede
# kontornavne i stedet, se generate_dummy_units.
KENDTE_KONTORER = {
    "Campusadministration Frederiksberg+": [
        "Bygninger", "Forskningsfinansiering", "HR", "IT-support",
        "Kommunikation", "Uddannelse", "Økonomi",
    ],
    "Campusadministration Nørre": [
        "Bygninger", "Forskningsfinansiering", "HR", "Kommunikation",
        "Ph.d.-administration", "Uddannelse", "Økonomi",
    ],
    "Campusadministration Søndre": [
        "Bygninger", "Forskningsfinansiering", "HR", "Kommunikation",
        "Uddannelse", "Økonomi",
    ],
    "KU Bygninger": [
        "Byggeri", "Plan", "Strategi og Styring", "Drift",
    ],
    "KU Forskning og Informationssikkerhed": [
        "Forskningscompliance", "Forskningsservice og Udvikling",
    ],
    "KU HR": [
        "HR Administration", "HR Udvikling og Strategi",
    ],
    "KU Innovation og Erhvervssamarbejde": [
        "Eksterne Samarbejder og Ledelsesbetjening", "Forsknings- og IP-jura",
        "KU Lighthouse",
    ],
    "KU IT": [
        "Digital Transformation", "Digitale Løsninger Administration",
        "Digitale Løsninger Forskning, Undervisning og Produktion",
        "Digitale Løsninger Uddannelse", "Infrastruktur og Platforme",
        "IT-sikkerhed og Support", "Rammer og Styring",
    ],
    "KU Kommunikation": [
        "Engagement", "Markedsføring af Uddannelser", "Medier",
        "Organisatorisk Kommunikation", "Presse",
        "Rådgivning og Medietræning", "Web og Visuel Identitet",
    ],
    "KU Uddannelse": [
        "Digitalisering, Eksamen og Lokaleplanlægning",
        "International Uddannelse", "Studieliv og Studiestøtte",
        "Uddannelsesstrategi og Analyse", "Uddannelsesvalg og Optagelse",
        "Videreuddannelse og Livslang Læring",
    ],
    "KU Økonomi": [
        "Koncern-Bygningsøkonomi", "Koncernregnskab", "Koncernindkøb",
        "Koncernøkonomi",
    ],
    # https://om.ku.dk/organisation/administration/rektoratets-stab/
    "Rektoratets Stab": [
        "Analyse og Business Intelligence",
        "Fora, Policy og Internationale Samarbejder",
        "Jura og Forkontor",
        "Strategi, Udvikling og Styring",
    ],
}


def generate_dummy_units(
    seed: int = 42,
    kontorer_per_enhed=(3, 8),
):
    """
    Genererer en flad liste af dict'e for hele hierarkiet under roden (KU).

    To niveauer under roden: Enhed (CA/KE) -> Kontor, direkte - ingen
    mellemliggende Afdeling-niveau. Enheder nævnt i KENDTE_KONTORER får
    deres rigtige kontornavne; resten får dummy-genererede placeholder-
    kontorer (mellem kontorer_per_enhed[0] og kontorer_per_enhed[1] stk.),
    som bør erstattes efterhånden som I finder de rigtige navne.

    Hver enhed har: id, navn, niveau, parent_id, ledelseslag, aarsvaerk,
    lonomkostninger. aarsvaerk/lonomkostninger er kun sat på leaf-enheder
    (Kontor) - Enhed-niveauet summeres op i app.py (rollup()).
    """
    rng = random.Random(seed)
    units = []

    for i, (enh_navn, enh_ledelseslag) in enumerate(ENHEDER, start=1):
        enh_id = f"ENH{i}"
        units.append({
            "id": enh_id,
            "navn": enh_navn,
            "niveau": "Enhed",
            "parent_id": ROOT_ID,
            "ledelseslag": enh_ledelseslag,
        })

        if enh_navn in KENDTE_KONTORER:
            kontor_navne = KENDTE_KONTORER[enh_navn]
        else:
            n_kontor = rng.randint(*kontorer_per_enhed)
            kontor_navne = [f"{enh_navn} - Kontor {k}" for k in range(1, n_kontor + 1)]

        for k, kontor_navn in enumerate(kontor_navne, start=1):
            kontor_id = f"{enh_id}-KONTOR{k}"
            units.append({
                "id": kontor_id,
                "navn": kontor_navn,
                "niveau": "Kontor",
                "parent_id": enh_id,
                "ledelseslag": LEDELSESLAG_PER_NIVEAU["Kontor"],
            })

    # Find leaf-enheder (dem uden børn) - kun de får tildelt tal direkte.
    children_of = {}
    for u in units:
        children_of.setdefault(u["parent_id"], []).append(u["id"])
    is_leaf = {u["id"]: (u["id"] not in children_of) for u in units}

    for u in units:
        if is_leaf[u["id"]]:
            aarsvaerk = rng.uniform(3, 35)
            gns_loen = rng.gauss(555_000, 45_000)
            u["aarsvaerk"] = round(aarsvaerk, 1)
            u["lonomkostninger"] = round(aarsvaerk * gns_loen)
        else:
            u["aarsvaerk"] = None
            u["lonomkostninger"] = None

    return units