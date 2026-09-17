import io
import os

import streamlit as st
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from pptx import Presentation
from pptx.util import Inches
import matplotlib
#matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker

from config import ROOT_ID, ROOT_NAVN, NIVEAUER, load_real_units, ADM_OMRAADER, load_forkortelser_raw
from data.loader import load_logo, logo_base64

PPTX_SKABELON = os.path.join(os.path.dirname(__file__), "ku_skabelon.pptx")
PPTX_LAYOUT_NAVN = "1_Title and Content"

@st.cache_data(show_spinner="Henter data...")
def load_units():
    return load_real_units()

def build_lookup_and_rollup(units):
    """
    by_id: {id: enhed-dict}, inkl. en tilføjet rod-enhed (ROOT_ID).
    children_of: {parent_id: [child_id, ...]}
    Aarsvaerk/medarbejdere rulles op, så ALLE enheder (ikke kun
    leaf-enheder) har summerede tal.
    """
    by_id = {u["id"]: dict(u) for u in units}
    by_id[ROOT_ID] = {
        "id": ROOT_ID, "navn": ROOT_NAVN, "niveau": "Rod",
        "parent_id": None, "ledelseslag": "Rektorat/direktion",
        "aarsvaerk": None, "medarbejdere": None,
    }
 
    children_of = {}
    for u in by_id.values():
        if u["parent_id"] is not None:
            children_of.setdefault(u["parent_id"], []).append(u["id"])
 
    def rollup(unit_id):
        u = by_id[unit_id]
        kids = children_of.get(unit_id, [])
        if not kids:
            return u["aarsvaerk"] or 0.0, u.get("medarbejdere") or 0
        total_av, total_med = 0.0, 0
        for k in kids:
            av, med = rollup(k)
            total_av += av
            total_med += med
        u["aarsvaerk"] = round(total_av, 1)
        u["medarbejdere"] = total_med
        return total_av, total_med
 
    rollup(ROOT_ID)
    return by_id, children_of
 
def path_to_root(by_id, unit_id):
    """Brødkrumme fra roden ned til unit_id, som liste af id'er."""
    path = [unit_id]
    while by_id[path[-1]]["parent_id"] is not None:
        path.append(by_id[path[-1]]["parent_id"])
    return list(reversed(path))
 
 
def leaves_under(children_of, unit_id):
    """
    Alle leaf-enheder (det yderste niveau, fx Kontor) under unit_id,
    fladtgjort på tværs af varierende dybde. Har unit_id selv ingen børn,
    returneres en tom liste.
    """
    kids = children_of.get(unit_id, [])
    if not kids:
        return []
    leaves = []
    for k in kids:
        if children_of.get(k):
            leaves.extend(leaves_under(children_of, k))
        else:
            leaves.append(k)
    return leaves

def _split_by_omraade(by_id, children_of, enh_uid, omraade_valgt, metric):
    """
    Deler en enheds kontorer i to grupper - dem der hører til omraade_valgt,
    og resten - og returnerer (omraade_vaerdi, rest_vaerdi) for den valgte
    metric. Begge nuværende metrics (Antal medarbejdere, Antal årsværk) er
    additive, så de bare summeres pr. gruppe.
    """
    kontor_ids = children_of.get(enh_uid, [])
    om_ids = [k for k in kontor_ids if by_id[k]["omraade"] == omraade_valgt]
    rest_ids = [k for k in kontor_ids if by_id[k]["omraade"] != omraade_valgt]

    if metric == "Antal medarbejdere":
        om_v = sum(by_id[k]["medarbejdere"] for k in om_ids)
        rest_v = sum(by_id[k]["medarbejdere"] for k in rest_ids)
    else:  # "Antal årsværk"
        om_v = sum(by_id[k]["aarsvaerk"] for k in om_ids)
        rest_v = sum(by_id[k]["aarsvaerk"] for k in rest_ids)
    return om_v, rest_v

def _bar_chart_png(navne, values, farve_hex, metric_label, enhed_tekst, width_in, height_in):
    """Bygger et vandret søjlediagram med matplotlib - samme stil som appens
    egne Plotly-diagrammer (KU-farver, størst øverst) - og returnerer det
    som PNG-bytes i en BytesIO. Kræver ikke Chrome."""
    farver = [farve_hex] * len(navne) if isinstance(farve_hex, str) else farve_hex
    colors = [f"#{f}" for f in farver]
    
    fig, ax = plt.subplots(figsize=(width_in, height_in), dpi=200)
    y_pos = range(len(navne))
    ax.barh(y_pos, values, color=colors)
    ax.set_yticks(list(y_pos))
    ax.set_yticklabels(navne, fontsize=9)
    ax.invert_yaxis()
    ax.set_xlabel(metric_label, fontsize=9)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    decimaler = 1 if enhed_tekst == "årsværk" else 0
    for i, v in enumerate(values):
        label = f"{v:,.{decimaler}f} {enhed_tekst}".replace(",", ".")
        ax.text(v, i, " " + label, va="center", fontsize=8)

    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png")
    plt.close(fig)
    buf.seek(0)
    return buf

def render_overblik(by_id, children_of, niveau1_ids, metric, key_prefix):
    """
    Den delte gengivelseslogik for "Organisatoriske enheder"-fanen - viser
    enhederne (Niveau 3) og evt. deres afdelinger (Niveau 4).
    """
    def metric_value(uid):
        u = by_id[uid]
        if metric == "Antal medarbejdere":
            return u["medarbejdere"]
        else:
            return u["aarsvaerk"]

    value_fmt = "%{x:,.1f} årsværk" if metric == "Antal årsværk" else "%{x:,.0f} medarbejdere"

    def _unikt_navn(navn, brugte_navne):
        unikt = navn
        while unikt in brugte_navne:
            unikt += " "
        brugte_navne.add(unikt)
        return unikt

    def _akse_label(metric):
        return f"{metric}"

    def _format_tal(v):
        if not v:
            return ""
        if metric == "Antal årsværk":
            return f"{v:,.1f}"
        return f"{v:,.0f}"

    overblik_niveau = st.radio(
        "**Vælg, hvilket niveau figurene skal vise:**",
        ["Niveau 3 (KE/CA)", "Niveau 3+4"],
        horizontal=True,
        key=f"{key_prefix}_overblik_niveau",
    )
    vis_kontor = overblik_niveau == "Niveau 3+4"

    alle_vaerdier = [metric_value(uid) for uid in niveau1_ids]
    if vis_kontor:
        for uid in niveau1_ids:
            for kid in children_of.get(uid, []):
                alle_vaerdier.append(metric_value(kid))
    x_maks = max(alle_vaerdier) * 1.05 if alle_vaerdier else 1

    if overblik_niveau == "Niveau 3 (KE/CA)":
        udvidet_key = f"{key_prefix}_niveau3_udvidet"
        if udvidet_key not in st.session_state:
            st.session_state[udvidet_key] = set()

        navne, vaerdier, farver, fuldnavne, klik_uid = [], [], [], [], []

        for uid in niveau1_ids:
            navne.append(f"<b>{by_id[uid]['navn']}</b>")
            vaerdier.append(metric_value(uid))
            farver.append("#901A1E")
            fuldnavne.append(by_id[uid].get("fuldt_navn", by_id[uid]["navn"]))
            klik_uid.append(uid)

            if uid in st.session_state[udvidet_key]:
                kontor_ids = sorted(
                    (kid for kid in children_of.get(uid, []) if not by_id[kid].get("er_selvnavngivet")),
                    key=metric_value, reverse=True,
                )
                for kid in kontor_ids:
                    navne.append(by_id[kid]["navn"])
                    vaerdier.append(metric_value(kid))
                    farver.append("#7992b5")
                    fuldnavne.append(by_id[kid].get("fuldt_navn", by_id[kid]["navn"]))
                    klik_uid.append(None)

        brugte_navne_n3 = set()
        navne = [_unikt_navn(n, brugte_navne_n3) for n in navne]

        fig_niveau3 = go.Figure(go.Bar(
            x=vaerdier,
            y=navne,
            orientation="h",
            marker_color=farver,
            marker_line_color="white",
            marker_line_width=1,
            text=[_format_tal(v) for v in vaerdier],
            textposition="outside",
            customdata=fuldnavne,
            hovertemplate="<b>%{customdata}</b><br>" + value_fmt + "<extra></extra>",
        ))
        fig_niveau3.update_layout(
            title=f"{metric} for Niveau 3",
            margin=dict(t=60, l=10, r=10, b=10),
            height=max(400, 30 * len(navne)),
            xaxis=dict(title=_akse_label(metric)),
            yaxis=dict(autorange="reversed"),
            bargap=0,
        )

        event_n3 = st.plotly_chart(
            fig_niveau3,
            key=f"{key_prefix}_overblik_niveau3_samlet",
            on_select="rerun",
            selection_mode=["points"],
            width="stretch",
        )

        if event_n3 and event_n3.selection and event_n3.selection["points"]:
            idx = event_n3.selection["points"][0].get("point_index")
            if idx is not None and idx < len(klik_uid) and klik_uid[idx] is not None:
                klikket_uid = klik_uid[idx]
                if klikket_uid in st.session_state[udvidet_key]:
                    st.session_state[udvidet_key].discard(klikket_uid)
                else:
                    st.session_state[udvidet_key].add(klikket_uid)
                st.rerun()

        st.caption("Klik på en søjle ovenfor for at folde dens afdelinger ud.")

    else:
        KOLONNE_GRUPPERING = [
            [
                "Campusadministration Frederiksberg+",
                "KU Bygninger",
                "KU IT",
                "KU HR",
            ],
            [
                "Campusadministration Nørre",
                "KU Forskning og Informationssikkerhed",
                "KU Økonomi",
                "KU Kommunikation",
            ],
            [
                "Campusadministration Søndre",
                "Rektoratets stab",
                "KU Uddannelse",
                "KU Innovation og Erhvervssamarbejde",
            ],
        ]
        grupper = [
            [uid for uid in gruppe_navne if uid in niveau1_ids]
            for gruppe_navne in KOLONNE_GRUPPERING
        ]

        fig_overblik = make_subplots(rows=1, cols=3, horizontal_spacing=0.10)
        hoejeste_raekkeantal = 0

        for g_idx, gruppe in enumerate(grupper):
            navne, vaerdier, farver, fuldnavne = [], [], [], []
            brugte_navne = set()

            for uid in gruppe:
                kontor_ids = sorted(
                    (kid for kid in children_of.get(uid, []) if not by_id[kid].get("er_selvnavngivet")),
                    key=metric_value, reverse=True,
                )

                if navne:
                    navne.append(_unikt_navn(" ", brugte_navne))
                    vaerdier.append(0)
                    farver.append("rgba(0,0,0,0)")
                    fuldnavne.append("")
                navne.append(_unikt_navn(f"<b>{by_id[uid]['navn']}</b>", brugte_navne))
                vaerdier.append(metric_value(uid))
                farver.append("#901A1E")
                fuldnavne.append(by_id[uid].get("fuldt_navn", by_id[uid]["navn"]))

                for kid in kontor_ids:
                    navne.append(_unikt_navn(by_id[kid]["navn"], brugte_navne))
                    vaerdier.append(metric_value(kid))
                    farver.append("#7992b5")
                    fuldnavne.append(by_id[kid].get("fuldt_navn", by_id[kid]["navn"]))

            if not navne:
                continue

            hoejeste_raekkeantal = max(hoejeste_raekkeantal, len(navne))
            kol = g_idx + 1

            fig_overblik.add_trace(
                go.Bar(
                    x=vaerdier,
                    y=navne,
                    orientation="h",
                    marker_color=farver,
                    marker_line_color="white",
                    marker_line_width=1,
                    text=[_format_tal(v) for v in vaerdier],
                    textposition="outside",
                    customdata=fuldnavne,
                    hovertemplate="<b>%{customdata}</b><br>" + value_fmt + "<extra></extra>",
                    showlegend=False,
                ),
                row=1, col=kol,
            )
            fig_overblik.update_yaxes(autorange="reversed", row=1, col=kol)

        fig_overblik.update_xaxes(range=[0, x_maks], title=_akse_label(metric))
        fig_overblik.update_layout(
            barmode="stack",
            title=f"{metric} for begge niveauer",
            margin=dict(t=60, l=10, r=10, b=70),
            height=max(160, 20 * hoejeste_raekkeantal + 60),
            bargap=0,
        )
        st.plotly_chart(fig_overblik, key=f"{key_prefix}_overblik_samlet", width="stretch")

def render_omraader(by_id, metric, key_prefix):
    """
    "Administrative områder"-fanen - Niveau 3 er her de administrative
    områder selv (HR, IT, Bygninger, osv.), og Niveau 4 er de kontorer,
    der hører til hvert område, uanset hvilken organisatorisk enhed de
    sidder under. Samme niveau-valg, klik-udfoldning og
    Niveau 3+4-overblik som render_overblik(), blot grupperet efter
    administrativt område i stedet for institut.
    """
    def metric_value(uid):
        u = by_id[uid]
        if metric == "Antal medarbejdere":
            return u["medarbejdere"]
        else:
            return u["aarsvaerk"]

    value_fmt = "%{x:,.1f} årsværk" if metric == "Antal årsværk" else "%{x:,.0f} medarbejdere"

    def _unikt_navn(navn, brugte_navne):
        unikt = navn
        while unikt in brugte_navne:
            unikt += " "
        brugte_navne.add(unikt)
        return unikt

    def _akse_label(metric):
        return f"{metric}"

    def _format_tal(v):
        if not v:
            return ""
        if metric == "Antal årsværk":
            return f"{v:,.1f}"
        return f"{v:,.0f}"

    # Omraade -> liste af kontor-id'er, på tværs af ALLE institutter.
    omraade_kontorer = {}
    for u in by_id.values():
        if u.get("niveau") != "Kontor":
            continue
        omraade = u.get("omraade")
        if omraade is None:
            continue
        omraade_kontorer.setdefault(omraade, []).append(u["id"])

    def omraade_total(omraade):
        return sum(metric_value(kid) for kid in omraade_kontorer[omraade])

    omraade_navne = [o for o in ADM_OMRAADER if o in omraade_kontorer]
    
    CA_RAEKKEFOELGE = [
        "Campusadministration Frederiksberg+",
        "Campusadministration Nørre",
        "Campusadministration Søndre",
    ]

    def institutter_i_omraade(omraade):
        """
        Grupperer et områdes kontorer efter deres institut (Niveau 3/KE-CA).
        Campusadministrationerne kommer altid først, i fast rækkefølge
        (Frederiksberg+, Nørre, Søndre) - resten derefter, sorteret efter
        faldende subtotal. Returnerer en liste af (institut_id, kontor_ids).
        """
        pr_institut = {}
        for kid in omraade_kontorer[omraade]:
            institut_id = by_id[kid]["parent_id"]
            pr_institut.setdefault(institut_id, []).append(kid)

        def sortnoegle(iid):
            if iid in CA_RAEKKEFOELGE:
                return (0, CA_RAEKKEFOELGE.index(iid), 0)
            return (1, 0, -sum(metric_value(k) for k in pr_institut[iid]))

        institut_ids_sorteret = sorted(pr_institut, key=sortnoegle)
        return [(iid, pr_institut[iid]) for iid in institut_ids_sorteret]

    overblik_niveau = st.radio(
        "**Vælg, hvilket niveau figurene skal vise:**",
        ["Niveau 3 (KE/CA)", "Niveau 3+4"],
        horizontal=True,
        key=f"{key_prefix}_overblik_niveau",
    )

    if overblik_niveau == "Niveau 3 (KE/CA)":
        udvidet_key = f"{key_prefix}_niveau3_udvidet"
        if udvidet_key not in st.session_state:
            st.session_state[udvidet_key] = set()

        navne, vaerdier, farver, fuldnavne, klik_omraade = [], [], [], [], []

        for omraade in omraade_navne:
            navne.append(f"<b>{omraade}</b>")
            vaerdier.append(omraade_total(omraade))
            farver.append("#901A1E")
            fuldnavne.append(omraade)
            klik_omraade.append(omraade)

            if omraade in st.session_state[udvidet_key]:
                for institut_id, kontor_ids in institutter_i_omraade(omraade):
                    navne.append(by_id[institut_id]["navn"])
                    vaerdier.append(sum(metric_value(k) for k in kontor_ids))
                    farver.append("#7992b5")
                    fuldnavne.append(by_id[institut_id].get("fuldt_navn", by_id[institut_id]["navn"]))
                    klik_omraade.append(None)

        brugte_navne = set()
        navne = [_unikt_navn(n, brugte_navne) for n in navne]

        fig = go.Figure(go.Bar(
            x=vaerdier,
            y=navne,
            orientation="h",
            marker_color=farver,
            marker_line_color="white",
            marker_line_width=1,
            text=[_format_tal(v) for v in vaerdier],
            textposition="outside",
            customdata=fuldnavne,
            hovertemplate="<b>%{customdata}</b><br>" + value_fmt + "<extra></extra>",
        ))
        fig.update_layout(
            title=f"{metric} pr. administrativt område",
            margin=dict(t=60, l=10, r=10, b=10),
            height=max(330, 33 * len(navne)),
            xaxis=dict(title=_akse_label(metric)),
            yaxis=dict(autorange="reversed"),
            bargap=0,
        )

        event = st.plotly_chart(
            fig,
            key=f"{key_prefix}_omraader_niveau3",
            on_select="rerun",
            selection_mode=["points"],
            width="stretch",
        )

        if event and event.selection and event.selection["points"]:
            idx = event.selection["points"][0].get("point_index")
            if idx is not None and idx < len(klik_omraade) and klik_omraade[idx] is not None:
                klikket = klik_omraade[idx]
                if klikket in st.session_state[udvidet_key]:
                    st.session_state[udvidet_key].discard(klikket)
                else:
                    st.session_state[udvidet_key].add(klikket)
                st.rerun()

        st.caption("Klik på en søjle ovenfor for at folde det administrative områdes Niveau 3-KE/CA ud")

    else:
        # Niveau 3+4: samme struktur som Organisatoriske enheder-plottet -
        # 1 række, 3 kolonner (IKKE et 3x3-grid). De 9 områder fordeles 3
        # pr. kolonne, og inden for hver kolonne stables områderne oven på
        # hinanden med luft imellem, så hver kolonnes højde naturligt
        # afspejler dens eget indhold.
        OMRAADE_GRUPPERING = [
            #["Udd.adm.", "Økonomi", "Innov.adm."],
            ["Uddannelsesadministration", "Økonomi", "Innovationsadministration"],
            ["IT", "HR", "Forskningsadministration"],
            ["Bygningsservice", "Kommunikation", "Stabe"],
        ]

        grupper_omraader = [
            [o for o in gruppe_navne if o in omraade_kontorer]
            for gruppe_navne in OMRAADE_GRUPPERING
        ]

        fig = make_subplots(rows=1, cols=3, horizontal_spacing=0.16)
        hoejeste_raekkeantal = 0
        hoejeste_vaerdi = 0

        for kol_idx, gruppe in enumerate(grupper_omraader):
            navne, vaerdier, farver, fuldnavne = [], [], [], []
            brugte_navne = set()

            for omraade in gruppe:
                if navne:
                    navne.append(_unikt_navn(" ", brugte_navne))
                    vaerdier.append(0)
                    farver.append("rgba(0,0,0,0)")
                    fuldnavne.append("")

                navne.append(_unikt_navn(f"<b>{omraade}</b>", brugte_navne))
                vaerdier.append(omraade_total(omraade))
                farver.append("#901A1E")
                fuldnavne.append(omraade)

                for institut_id, kontor_ids in institutter_i_omraade(omraade):
                    kontor_ids = sorted(kontor_ids, key=metric_value, reverse=True)
                    institut_fuldt = by_id[institut_id].get("fuldt_navn", by_id[institut_id]["navn"])
                    for kid in kontor_ids:
                        if by_id[kid].get("er_selvnavngivet"):
                            kontor_label = by_id[institut_id]["navn"]
                            kontor_fuldt_label = institut_fuldt
                        else:
                            kontor_label = f"{by_id[institut_id]['navn']} | {by_id[kid]['navn']}"
                            kontor_fuldt = by_id[kid].get("fuldt_navn", by_id[kid]["navn"])
                            kontor_fuldt_label = f"{institut_fuldt} | {kontor_fuldt}"
                        navne.append(_unikt_navn(kontor_label, brugte_navne))
                        vaerdier.append(metric_value(kid))
                        farver.append("#7992b5")
                        fuldnavne.append(kontor_fuldt_label)

            if not navne:
                continue

            hoejeste_raekkeantal = max(hoejeste_raekkeantal, len(navne))
            hoejeste_vaerdi = max(hoejeste_vaerdi, max(vaerdier))
            kol = kol_idx + 1

            fig.add_trace(
                go.Bar(
                    x=vaerdier,
                    y=navne,
                    orientation="h",
                    marker_color=farver,
                    marker_line_color="white",
                    marker_line_width=1,
                    text=[_format_tal(v) for v in vaerdier],
                    textposition="outside",
                    customdata=fuldnavne,
                    hovertemplate="<b>%{customdata}</b><br>" + value_fmt + "<extra></extra>",
                    showlegend=False,
                ),
                row=1, col=kol,
            )
            fig.update_yaxes(autorange="reversed", row=1, col=kol)

        x_maks = hoejeste_vaerdi * 1.15 if hoejeste_vaerdi else 1
        fig.update_xaxes(range=[0, x_maks], title=_akse_label(metric))
        fig.update_layout(
            title=f"{metric} for alle administrative områder",
            margin=dict(t=60, l=10, r=10, b=10),
            height=max(160, 20 * hoejeste_raekkeantal + 60),
            bargap=0,
        )
        st.plotly_chart(fig, key=f"{key_prefix}_omraader_niveau34", width="stretch")


def main():
    st.set_page_config(
        page_title="KU ledelseslag",
        page_icon=load_logo(),
        layout="wide",
    )

    col_logo, col_title, col_download = st.columns([1, 4, 1])

    with col_logo:
        st.markdown(
            f'<img src="data:image/png;base64,{logo_base64()}" '
            f'style="max-width:180px; width:100%;">',
            unsafe_allow_html=True
        )

    with col_title:
        st.title("Personaleoverblik")
        st.markdown("#### Materiale til A-DIR-seminar, 21. oktober 2026")

    st.markdown(
"""
Dette værktøj viser årsværk for KU's administrative enheder på Niveau 3 
(koncernenheder og campusadministrationer) og deres afdelinger (Niveau 4). 

**Sådan bruger du værktøjet:**
- **Organisatoriske enheder / Administrative områder**: Vælg fanen nedenfor - 
"Organisatoriske enheder" viser den organisatoriske opdeling, mens "Administrative 
områder" i stedet viser de administrative områder (f.eks. HR eller IT) på tværs 
af de organisatoriske enheder.
- **Fuldt overblik**: I hver fane kan du vælge, om figurene skal vise Niveau 3 alene eller et fuldt overblik med
både Niveau 3 og 4. I Niveau 3-visningen kan du klikke på en søjle for at folde dens underliggende enheder ud; klik
igen for at folde sammen. Du kan folde flere søjler ud samtidig. 

**Bemærk**: Niveau 4 er det mest detaljerede niveau, værktøjet viser. Eventuelle 
underliggende Niveau 5- og 6-sektioner indgår i tallene for den Niveau 4-afdeling, 
de hører under, men er ikke brudt særskilt ned. 
""")

    units = load_units()
    by_id, children_of = build_lookup_and_rollup(units)

    niveau1_navn = NIVEAUER[0]
    niveau1_ids = sorted(
        (uid for uid, u in by_id.items() if u["niveau"] == niveau1_navn),
        key=lambda uid: by_id[uid]["navn"],
    )

    metric = "Antal årsværk"

    fane_enheder, fane_omraader = st.tabs(["Organisatoriske enheder", "Administrative områder"])

    with fane_enheder:
        render_overblik(by_id, children_of, niveau1_ids, metric, key_prefix="enh")

    with fane_omraader:
        render_omraader(by_id, metric, key_prefix="omr")

    st.markdown("##### Dokumentation")

    with st.expander("Datagrundlag"):
        st.markdown(
"""
Optællingsmetrikken årsværk er opgjort som 'beregnet personaleforbrug' for august måned 2026 fra 
Personalesammensætning på Tableauserveren. 

KUorg anvendes til at placere de enkelte ansatte på Niveau 4-afdelinger. Mellem de to datakilder er der 
uoverensstemmelse for to personer, hvor data fra Personalesammensætning ikke stemmer overens med den 
organisatoriske placering, KUorg angiver for personen. Begge tilfælde er blevet ekskluderet. 

Derudover har to personer årsværk fordelt på flere KE/CA (Niveau 3) i Personalesammensætning - deres årsværk er fordelt
ud på de KE/CA og afdelinger, de reelt er tilknyttet. 
""" 
    )

    with st.expander("Hvordan er de administrative områder blevet inddelt?"):
        
        st.table(
            load_forkortelser_raw(),
            hide_index=True
        )

    st.markdown(f"""
<hr style="margin-top: 50px;">
<div style="text-align:center; color:#666; font-size: 0.9em;">
  REKSTAB Analyse · Amanda Schramm Petersen · <a href="mailto:ascp@adm.ku.dk">ascp@adm.ku.dk</a>
  · opdateret 14. september 2026
</div>
""", unsafe_allow_html=True)



if __name__ == "__main__":
    main()