import io

import streamlit as st
import plotly.graph_objects as go
from pptx import Presentation
from pptx.util import Inches
import matplotlib
#matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker

from config import ROOT_ID, ROOT_NAVN, NIVEAUER, load_real_units
from data.loader import load_logo, logo_base64

# Ret disse to til jeres faktiske skabelon og layoutnavn (samme princip som
# i create_pptx.py - I finder layoutnavne med:
# [l.name for l in Presentation(PPTX_SKABELON).slide_layouts])
PPTX_SKABELON = r"C:\Users\rjp530\Desktop\kontormøde_pptx\ku_skabelon.pptx"
PPTX_LAYOUT_NAVN = "1_Title and Content"

@st.cache_data
def load_units():
    return load_real_units()


def build_lookup_and_rollup(units):
    """
    by_id: {id: enhed-dict}, inkl. en tilføjet rod-enhed (ROOT_ID).
    children_of: {parent_id: [child_id, ...]}
    Aarsvaerk/lonomkostninger rulles op, så ALLE enheder (ikke kun
    leaf-enheder) har summerede tal.
    """
    by_id = {u["id"]: dict(u) for u in units}
    by_id[ROOT_ID] = {
        "id": ROOT_ID, "navn": ROOT_NAVN, "niveau": "Rod",
        "parent_id": None, "ledelseslag": "Rektorat/direktion",
        "aarsvaerk": None, "lonomkostninger": None,
    }
 
    children_of = {}
    for u in by_id.values():
        if u["parent_id"] is not None:
            children_of.setdefault(u["parent_id"], []).append(u["id"])
 
    def rollup(unit_id):
        u = by_id[unit_id]
        kids = children_of.get(unit_id, [])
        if not kids:
            return u["aarsvaerk"] or 0.0, u["lonomkostninger"] or 0
        total_av, total_lon = 0.0, 0
        for k in kids:
            av, lon = rollup(k)
            total_av += av
            total_lon += lon
        u["aarsvaerk"] = round(total_av, 1)
        u["lonomkostninger"] = total_lon
        return total_av, total_lon
 
    rollup(ROOT_ID)
    return by_id, children_of
 
 
def gns_loen(by_id, unit_id):
    u = by_id[unit_id]
    if not u["aarsvaerk"]:
        return 0
    return u["lonomkostninger"] / u["aarsvaerk"]
 
 
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
    metric. For "Gns. lønomkostning pr. årsværk" beregnes et separat
    gennemsnit pr. gruppe (kan ikke stakkes, da et gennemsnit ikke er
    additivt) - for de to andre metrics summeres der (additive, kan stakkes).
    """
    kontor_ids = children_of.get(enh_uid, [])
    om_ids = [k for k in kontor_ids if by_id[k]["omraade"] == omraade_valgt]
    rest_ids = [k for k in kontor_ids if by_id[k]["omraade"] != omraade_valgt]

    if metric == "Gns. lønomkostning pr. årsværk":
        om_av = sum(by_id[k]["aarsvaerk"] for k in om_ids)
        rest_av = sum(by_id[k]["aarsvaerk"] for k in rest_ids)
        om_v = (sum(by_id[k]["lonomkostninger"] for k in om_ids) / om_av) if om_av else 0
        rest_v = (sum(by_id[k]["lonomkostninger"] for k in rest_ids) / rest_av) if rest_av else 0
    elif metric == "Samlede lønomkostninger":
        om_v = sum(by_id[k]["lonomkostninger"] for k in om_ids)
        rest_v = sum(by_id[k]["lonomkostninger"] for k in rest_ids)
    else:
        om_v = sum(by_id[k]["aarsvaerk"] for k in om_ids)
        rest_v = sum(by_id[k]["aarsvaerk"] for k in rest_ids)
    return om_v, rest_v

def _bar_chart_png(navne, values, farve_hex, metric_label, vaerdi_er_kr, width_in, height_in):
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
    if vaerdi_er_kr:
        ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{x:,.0f}".replace(",", ".")))

    for i, v in enumerate(values):
        label = f"{v:,.0f} kr.".replace(",", ".") if vaerdi_er_kr else f"{v:.1f} årsværk"
        ax.text(v, i, " " + label, va="center", fontsize=8)

    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png")
    plt.close(fig)
    buf.seek(0)
    return buf

def build_full_pptx(by_id, children_of, niveau1_ids, metric, metric_value):
    """
    Bygger en pptx ud fra KU-skabelonen med native PowerPoint-diagrammer
    (kræver ikke Chrome/kaleido, i modsætning til billedeksport af plotly-
    figurer). Slide 1: oversigt over alle CA/KE-enheder. Slide 2-N: én
    slide PR. ENHED med to diagrammer side om side - overblikket til
    venstre, enhedens kontorer til højre - dvs. alle enheder, ikke kun
    den der aktuelt er valgt/vist i appen.
    """
    prs = Presentation(PPTX_SKABELON)
    layout = next((l for l in prs.slide_layouts if l.name == PPTX_LAYOUT_NAVN), None)
    if layout is None:
        layout = prs.slide_layouts[min(1, len(prs.slide_layouts) - 1)]

    vaerdi_er_kr = metric != "Antal årsværk"

    def set_title(slide, title):
        title_ph = next((p for p in slide.placeholders if p.placeholder_format.idx == 0), None)
        if title_ph is not None:
            title_ph.text_frame.text = title
        else:
            box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.3), Inches(0.7))
            box.text_frame.text = title

    # Slide 1: alle CA/KE-enheder
    navne = [by_id[uid]["navn"] for uid in niveau1_ids]
    values = [metric_value(uid) for uid in niveau1_ids]

    slide = prs.slides.add_slide(layout)
    set_title(slide, "Campusadministrationer og koncernenheder")
    def add_chart_image(slide, left, top, width, height, navne, values, farve_hex):
        png_buf = _bar_chart_png(
            navne, values, farve_hex, metric, vaerdi_er_kr,
            width_in=width.inches, height_in=height.inches,
        )
        slide.shapes.add_picture(png_buf, left, top, width=width, height=height)
    add_chart_image(
        slide, Inches(0.6), Inches(1.7), Inches(12.1), Inches(5.4),
        navne, values, "901A1E",
    )

    # Slide 2..N: én slide pr. enhed, med to diagrammer side om side
    for uid in niveau1_ids:
        leaf_ids = leaves_under(children_of, uid)
        if not leaf_ids:
            continue
        leaf_ids = sorted(leaf_ids, key=lambda lid: by_id[lid]["navn"])
        leaf_navne = [by_id[lid]["navn"] for lid in leaf_ids]
        leaf_values = [metric_value(lid) for lid in leaf_ids]

        slide = prs.slides.add_slide(layout)
        set_title(slide, by_id[uid]["navn"])
        bar_farver = ["901A1E" if uid2 == uid else "E6C9CC" for uid2 in niveau1_ids]
        add_chart_image(
            slide, Inches(0.4), Inches(1.7), Inches(6.1), Inches(5.4),
            navne, values, bar_farver,
        )
        add_chart_image(
            slide, Inches(6.8), Inches(1.7), Inches(6.1), Inches(5.4),
            leaf_navne, leaf_values, "BAC7D9",
        )

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf

def main():
    # --- Page config ---
    st.set_page_config(
        page_title="KU ledelseslag",
        page_icon=load_logo(),
        layout="wide",
    )
 
    col_logo, col_title = st.columns([1, 4])
    with col_logo:
        st.markdown(
            f'<img src="data:image/png;base64,{logo_base64()}" '
            f'style="max-width:180px; width:100%;">',
            unsafe_allow_html=True
        )
 
    with col_title:
        st.title("Københavns Universitets ledelseslag")
 
    # --- Data: indlæs og rul årsværk/lønomkostninger op gennem hierarkiet ---
    units = load_units()
    by_id, children_of = build_lookup_and_rollup(units)
 
    # Niveau 4 = det yderste niveau (KU=1, Enhed=2, Afdeling=3, Kontor=4)
    niveau1_navn = NIVEAUER[0]
    niveau1_ids = sorted(
        (uid for uid, u in by_id.items() if u["niveau"] == niveau1_navn),
        key=lambda uid: by_id[uid]["navn"],
    )
 
    metric = st.radio(
        " ",
        ["Gns. lønomkostning pr. årsværk", "Samlede lønomkostninger", "Antal årsværk"],
        horizontal=True,
        key="metric_valg",
    )
 
    def metric_value(uid):
        if metric == "Gns. lønomkostning pr. årsværk":
            return gns_loen(by_id, uid)
        elif metric == "Samlede lønomkostninger":
            return by_id[uid]["lonomkostninger"]
        else:
            return by_id[uid]["aarsvaerk"]
 
    y_fmt = "%{y:,.1f} årsværk" if metric == "Antal årsværk" else "%{y:,.0f} kr."

    def _nulstil_valg():
        st.session_state.pop("valgt_niveau1", None)

    visning = st.radio(
        "Visning",
        ["Afdelinger", "Administrative områder"],
        horizontal=True,
        key="visning_valg",
        on_change=_nulstil_valg,
    )

    omraade_valgt = None
    if visning == "Administrative områder":
        omraader = sorted(set(
            u["omraade"] for u in by_id.values()
            if u.get("niveau") == "Kontor" and u.get("omraade") is not None
        ))
        omraade_valgt = st.selectbox("Vælg område", omraader, key="omraade_valg")
 
    col_bar1, col_bar2 = st.columns(2)
 
    # -----------------------------------------------------------------
    # Venstre: søjlediagram 1 - alle niveau 4-ledere (Kontor)
    # -----------------------------------------------------------------
    with col_bar1:
        st.subheader("Campusadministrationer og koncernenheder")

        navne = [by_id[uid]["navn"] for uid in niveau1_ids]
        value_fmt = "%{x:,.1f} årsværk" if metric == "Antal årsværk" else "%{x:,.0f} kr."

        if visning == "Afdelinger":
            y = [metric_value(uid) for uid in niveau1_ids]

            fig1 = go.Figure(go.Bar(
                x=y,
                y=navne,
                orientation="h",
                marker_color="#901A1E",
                hovertemplate="<b>%{y}</b><br>" + value_fmt + "<extra></extra>",
            ))
            fig1.update_layout(
                barmode="stack",
                margin=dict(t=40, l=10, r=10, b=50),
                height=max(420, 28 * len(navne)),
                xaxis_title=metric,
                yaxis=dict(autorange="reversed"),
                legend=dict(orientation="h", yanchor="top", y=-0.22, xanchor="left", x=0),
            )
        else:
            omraade_serie, rest_serie = [], []
            for uid in niveau1_ids:
                o, r = _split_by_omraade(by_id, children_of, uid, omraade_valgt, metric)
                omraade_serie.append(o)
                rest_serie.append(r)

            fig1 = go.Figure()
            fig1.add_trace(go.Bar(
                x=omraade_serie, y=navne, orientation="h", name=omraade_valgt,
                marker=dict(color="#901A1E"),
                hovertemplate="<b>%{y}</b><br>" + omraade_valgt + ": " + value_fmt + "<extra></extra>",
            ))
            fig1.add_trace(go.Bar(
                x=rest_serie, y=navne, orientation="h", name="Øvrige",
                marker=dict(color="#E6C9CC"),
                hovertemplate="<b>%{y}</b><br>Øvrige: " + value_fmt + "<extra></extra>",
            ))
            fig1.update_layout(
                barmode="stack",
                margin=dict(t=40, l=10, r=10, b=10),
                height=max(420, 28 * len(navne)),
                xaxis_title=metric,
                yaxis=dict(autorange="reversed"),
                legend=dict(orientation="h", yanchor="top", y=-0.22, xanchor="left", x=0),
            )

        event = st.plotly_chart(
            fig1,
            key="bar_niveau1",
            on_select="rerun",
            selection_mode=["points"],
            width="stretch",
        )

        if event and event.selection and event.selection["points"]:
            point = event.selection["points"][0]
            idx = point.get("point_index")
            curve = point.get("curve_number")
            # I "Områder"-visning er trace 0 den KU-røde område-del - kun
            # klik dér skal opdatere højre diagram. I "Kontorer"-visning er
            # der kun én trace (curve altid 0), så alle klik tæller.
            if idx is not None and curve == 0:
                st.session_state.valgt_niveau1 = niveau1_ids[idx]
 
    # -----------------------------------------------------------------
    # Højre: søjlediagram 2 - kontorerne (yderste niveau) under den
    # valgte campusadministration/koncernenhed
    # -----------------------------------------------------------------
    with col_bar2:
        valgt = st.session_state.get("valgt_niveau1")
 
        if valgt is None or valgt not in by_id:
            st.header(" \n \n ")
            st.header(" \n \n ")
            st.error("Klik på en søjle til venstre for at se kontorerne under den enhed.")
        else:
            leaf_ids = leaves_under(children_of, valgt)
            if visning == "Administrative områder":
                st.subheader(f"{omraade_valgt}-andel pr. kontor under: {by_id[valgt]['navn']}")
            else:
                st.subheader(f"Kontorer under: {by_id[valgt]['navn']}")

            if not leaf_ids:
                st.info("Denne enhed har ingen underliggende kontorer i dummy-dataen.")
            else:
                leaf_ids = sorted(leaf_ids, key=lambda uid: by_id[uid]["navn"])
                leaf_navne = [by_id[uid]["navn"] for uid in leaf_ids]

                if visning == "Administrative områder":
                    leaf_y = [
                        metric_value(uid) if by_id[uid]["omraade"] == omraade_valgt else 0
                        for uid in leaf_ids
                    ]
                    leaf_farver = [
                        "#901A1E" if by_id[uid]["omraade"] == omraade_valgt else "#E6C9CC"
                        for uid in leaf_ids
                    ]
                else:
                    leaf_y = [metric_value(uid) for uid in leaf_ids]
                    leaf_farver = "#BAC7D9"

                fig2 = go.Figure(go.Bar(
                    x=leaf_y,
                    y=leaf_navne,
                    orientation="h",
                    marker_color=leaf_farver,
                    hovertemplate="<b>%{y}</b><br>" + value_fmt + "<extra></extra>",
                ))
                fig2.update_layout(
                    margin=dict(t=40, l=10, r=10, b=10),
                    height=max(420, 28 * len(leaf_navne)),
                    xaxis_title=metric,
                    yaxis=dict(autorange="reversed"),
                )
                st.plotly_chart(fig2, key="bar_kontorer", width="stretch")
    
    st.divider()
    st.subheader("Fuldt overblik: alle enheder og kontorer")

    overblik_niveau = st.radio(
        "Vis niveau",
        ["Niveau 3 (enheder)", "Niveau 4 (kontorer)", "Begge niveauer"],
        horizontal=True,
        key="overblik_niveau",
    )
    vis_enhed = overblik_niveau in ("Niveau 3 (enheder)", "Begge niveauer")
    vis_kontor = overblik_niveau in ("Niveau 4 (kontorer)", "Begge niveauer")

    # Fælles x-akse-grænse på tværs af ALLE tre plots, så de er sammenlignelige.
    alle_vaerdier = []
    for uid in niveau1_ids:
        if vis_enhed:
            alle_vaerdier.append(metric_value(uid))
        if vis_kontor:
            for kid in children_of.get(uid, []):
                alle_vaerdier.append(metric_value(kid))
    x_maks = max(alle_vaerdier) * 1.05 if alle_vaerdier else 1

    def _unikt_navn(navn, brugte_navne):
        """
        Tilføjer et usynligt mellemrum, hvis navnet allerede optræder i
        samme plot (fx et kontor der hedder det samme som sin enhed) -
        ellers slår Plotly de to søjler sammen til én, da den kategoriske
        y-akse matcher på selve teksten, ikke listeposition.
        """
        unikt = navn
        while unikt in brugte_navne:
            unikt += " "
        brugte_navne.add(unikt)
        return unikt

    if overblik_niveau == "Niveau 3 (enheder)":
        # Ét samlet plot, ligesom fig1 foroven - ingen gruppe/kolonne-opdeling
        # nødvendig, da der kun er 12 søjler i alt.
        navne = [by_id[uid]["navn"] for uid in niveau1_ids]
        vaerdier = [metric_value(uid) for uid in niveau1_ids]

        fig_niveau3 = go.Figure(go.Bar(
            x=vaerdier,
            y=navne,
            orientation="h",
            marker_color="#901A1E",
            marker_line_color="white",
            marker_line_width=1,
            hovertemplate="<b>%{y}</b><br>" + value_fmt + "<extra></extra>",
        ))
        fig_niveau3.update_layout(
            margin=dict(t=40, l=10, r=10, b=10),
            height=max(420, 28 * len(navne)),
            xaxis=dict(range=[0, x_maks]),
            yaxis=dict(autorange="reversed"),
        )
        st.plotly_chart(fig_niveau3, key="overblik_niveau3_samlet", width="stretch")

    elif overblik_niveau == "Niveau 4 (kontorer)":
        # Alle kontorer på tværs af enheder, sorteret efter størrelse og delt
        # i tre nogenlunde lige store grupper - hver figur starter derfor
        # med den største i netop den gruppe.
        alle_kontor_ids = [kid for uid in niveau1_ids for kid in children_of.get(uid, [])]
        alle_kontor_ids.sort(key=metric_value, reverse=True)

        chunk_n4 = -(-len(alle_kontor_ids) // 3)  # oprund
        grupper_n4 = [alle_kontor_ids[i:i + chunk_n4] for i in range(0, len(alle_kontor_ids), chunk_n4)]
        while len(grupper_n4) < 3:
            grupper_n4.append([])

        kolonner_n4 = st.columns(3)
        for g_idx, gruppe_kontor_ids in enumerate(grupper_n4):
            if not gruppe_kontor_ids:
                continue
            brugte_navne_n4 = set()
            navne = [_unikt_navn(by_id[kid]["navn"], brugte_navne_n4) for kid in gruppe_kontor_ids]
            vaerdier = [metric_value(kid) for kid in gruppe_kontor_ids]

            fig_niveau4 = go.Figure(go.Bar(
                x=vaerdier,
                y=navne,
                orientation="h",
                marker_color="#BAC7D9",
                marker_line_color="white",
                marker_line_width=1,
                hovertemplate="<b>%{y}</b><br>" + value_fmt + "<extra></extra>",
            ))
            fig_niveau4.update_layout(
                margin=dict(t=40, l=10, r=10, b=10),
                height=max(160, 28 * len(navne) + 60),
                xaxis=dict(range=[0, x_maks]),
                yaxis=dict(autorange="reversed"),
            )
            with kolonner_n4[g_idx]:
                st.plotly_chart(fig_niveau4, key=f"overblik_niveau4_plot_{g_idx}", width="stretch")


    else:

        # Del de 13 enheder i tre nogenlunde lige store, sammenhængende grupper -
        # én gruppe pr. kolonne/plot.
        CA_RAEKKEFOELGE = [
            "Campusadministration Frederiksberg+",
            "Campusadministration Nørre",
            "Campusadministration Søndre",
        ]
        ca_ids = sorted(
            (uid for uid in niveau1_ids if by_id[uid]["navn"] in CA_RAEKKEFOELGE),
            key=lambda uid: CA_RAEKKEFOELGE.index(by_id[uid]["navn"]),
        )
        ovrige_ids = [uid for uid in niveau1_ids if by_id[uid]["navn"] not in CA_RAEKKEFOELGE]

        chunk = -(-len(ovrige_ids) // 3)  # oprund
        ovrige_grupper = [ovrige_ids[i:i + chunk] for i in range(0, len(ovrige_ids), chunk)]
        while len(ovrige_grupper) < 3:
            ovrige_grupper.append([])
        while len(ca_ids) < 3:
            ca_ids.append(None)

        grupper = [
            ([ca] if ca is not None else []) + ovrige
            for ca, ovrige in zip(ca_ids, ovrige_grupper)
        ]

        kolonner = st.columns(3)
        for g_idx, gruppe in enumerate(grupper):
            navne, vaerdier, farver = [], [], []
            brugte_navne = set()

            for uid in gruppe:
                if vis_enhed:
                    if navne and vis_kontor:  # luft foer alle enheds-soejler undtagen den allerførste i plottet
                        navne.append(_unikt_navn(" ", brugte_navne))
                        vaerdier.append(0)
                        farver.append("rgba(0,0,0,0)")
                    navne.append(_unikt_navn(by_id[uid]["navn"], brugte_navne))
                    vaerdier.append(metric_value(uid))
                    farver.append("#901A1E")

                if vis_kontor:
                    kontor_ids = sorted(children_of.get(uid, []), key=metric_value, reverse=True)
                    for kid in kontor_ids:
                        navne.append(_unikt_navn(by_id[kid]["navn"], brugte_navne))
                        vaerdier.append(metric_value(kid))
                        farver.append("#BAC7D9")

            if not navne:
                continue

            fig_overblik = go.Figure(go.Bar(
                x=vaerdier,
                y=navne,
                orientation="h",
                marker_color=farver,
                marker_line_color="white",
                marker_line_width=1,
                hovertemplate="<b>%{y}</b><br>" + value_fmt + "<extra></extra>",
            ))
            fig_overblik.update_layout(
                margin=dict(t=40, l=10, r=10, b=10),
                height=max(160, 20 * len(navne) + 60),
                xaxis=dict(range=[0, x_maks]),
                yaxis=dict(autorange="reversed"),
                bargap=0,
            )

            with kolonner[g_idx]:
                st.plotly_chart(fig_overblik, key=f"overblik_plot_{g_idx}", width="stretch")

    st.divider()
    if st.button("Generér PowerPoint med alle enheder"):
        with st.spinner("Bygger PowerPoint..."):
            pptx_buf = build_full_pptx(by_id, children_of, niveau1_ids, metric, metric_value)
        st.download_button(
            "Download PowerPoint",
            data=pptx_buf,
            file_name="ledelseslag_alle_enheder.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )

 
if __name__ == "__main__":
    main()