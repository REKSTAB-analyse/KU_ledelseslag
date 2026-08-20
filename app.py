import streamlit as st
import plotly.graph_objects as go
 
from config import ROOT_ID, ROOT_NAVN, NIVEAUER, generate_dummy_units
from data.loader import load_logo, logo_base64

@st.cache_data
def load_units():
    return generate_dummy_units()


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
 
    col_bar1, col_bar2 = st.columns(2)
 
    # -----------------------------------------------------------------
    # Venstre: søjlediagram 1 - alle niveau 4-ledere (Kontor)
    # -----------------------------------------------------------------
    with col_bar1:
        st.subheader("Campusadministrationer og koncernenheder")
 
        navne = [by_id[uid]["navn"] for uid in niveau1_ids]
        y = [metric_value(uid) for uid in niveau1_ids]
 
        value_fmt = "%{x:,.1f} årsværk" if metric == "Antal årsværk" else "%{x:,.0f} kr."
 
        fig1 = go.Figure(go.Bar(
            x=y,
            y=navne,
            orientation="h",
            marker_color="#901A1E",
            hovertemplate="<b>%{y}</b><br>" + value_fmt + "<extra></extra>",
        ))
        fig1.update_layout(
            margin=dict(t=10, l=10, r=10, b=10),
            height=max(420, 28 * len(navne)),
            xaxis_title=metric,
            yaxis=dict(autorange="reversed"),
        )
 
        event = st.plotly_chart(
            fig1,
            key="bar_niveau1",
            on_select="rerun",
            selection_mode=["points"],
            width="stretch",
        )
 
        if event and event.selection and event.selection["points"]:
            idx = event.selection["points"][0].get("point_index")
            if idx is not None:
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
            st.subheader(f"Kontorer under: {by_id[valgt]['navn']}")
 
            if not leaf_ids:
                st.info("Denne enhed har ingen underliggende kontorer i dummy-dataen.")
            else:
                leaf_ids = sorted(leaf_ids, key=lambda uid: by_id[uid]["navn"])
                leaf_navne = [by_id[uid]["navn"] for uid in leaf_ids]
                leaf_y = [metric_value(uid) for uid in leaf_ids]
 
                fig2 = go.Figure(go.Bar(
                    x=leaf_y,
                    y=leaf_navne,
                    orientation="h",
                    marker_color="#BAC7D9",
                    hovertemplate="<b>%{y}</b><br>" + value_fmt + "<extra></extra>",
                ))
                fig2.update_layout(
                    margin=dict(t=10, l=10, r=10, b=10),
                    height=max(420, 28 * len(leaf_navne)),
                    xaxis_title=metric,
                    yaxis=dict(autorange="reversed"),
                )
                st.plotly_chart(fig2, key="bar_kontorer", width="stretch")
 
if __name__ == "__main__":
    main()