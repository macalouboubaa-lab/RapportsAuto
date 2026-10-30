"""Point d'entrée Streamlit — Rapport QoS T\u00e9l\u00e9com.

Ce fichier ne contient que la logique d'interface (UI). Toute la logique
métier est importée depuis le package ``src/`` :
- src.data_processing : nettoyage, filtrage, calculs
- src.visualizations  : graphiques Matplotlib
- src.ui_components    : cartes KPI et tableaux HTML
"""

from __future__ import annotations

import streamlit as st

from src import data_processing as dp
from src import ui_components as ui
from src import visualizations as viz

SEUIL_EXPORT_ERREURS = 8  # au-del\u00e0, on propose l'export Excel

st.set_page_config(
    page_title="Rapport QoS T\u00e9l\u00e9com",
    page_icon="\U0001F4F6",
    layout="wide",
)

# Injection du CSS global (une seule fois)
st.markdown(ui.global_css(), unsafe_allow_html=True)

st.title("\U0001F4F6 Rapport de Qualit\u00e9 de Service (QoS)")
st.caption("Monitoring t\u00e9l\u00e9com — g\u00e9n\u00e9ration de rapports \u00e0 partir d'un export Excel.")

# --------------------------------------------------------------------------
# 1. Barre lat\u00e9rale : upload + filtres
# --------------------------------------------------------------------------
with st.sidebar:
    st.header("\u2699\ufe0f Param\u00e8tres")
    fichier = st.file_uploader("Fichier Excel (.xlsx)", type=["xlsx"])
    st.markdown("---")
    st.caption("Les donn\u00e9es sont trait\u00e9es en m\u00e9moire. Aucun fichier n'est stock\u00e9 sur le serveur.")

if fichier is None:
    st.info("\u2b06\ufe0f Importez un fichier Excel pour g\u00e9n\u00e9rer le rapport.")
    st.stop()

# --------------------------------------------------------------------------
# 2. Chargement + pr\u00e9paration des donn\u00e9es
# --------------------------------------------------------------------------
try:
    df_brut = dp.load_excel(fichier)
except Exception as exc:  # noqa: BLE001
    st.error(f"Impossible de lire le fichier Excel : {exc}")
    st.stop()

df = dp.clean_dataframe(df_brut)
df = dp.drop_current_hour(df)

if df.empty:
    st.warning("Aucune transaction exploitable apr\u00e8s nettoyage (heure en cours exclue).")
    st.stop()

# Filtre filiale
filiales = dp.list_filiales(df)
with st.sidebar:
    filiale = st.selectbox("Filiale", options=filiales, index=0)

df = dp.filter_data(df, filiale)

if df.empty:
    st.warning("Aucune donn\u00e9e pour la filiale s\u00e9lectionn\u00e9e.")
    st.stop()

# --------------------------------------------------------------------------
# 3. KPI
# --------------------------------------------------------------------------
kpis = dp.compute_kpis(df)
st.subheader("Indicateurs cl\u00e9s")
st.markdown(ui.kpi_cards(kpis), unsafe_allow_html=True)

# --------------------------------------------------------------------------
# 4. Graphiques c\u00f4te \u00e0 c\u00f4te
# --------------------------------------------------------------------------
hourly = dp.hourly_stats(df)
st.subheader("Analyse horaire")
col_g, col_d = st.columns(2)
with col_g:
    st.pyplot(viz.plot_volume_qos(hourly), use_container_width=True)
with col_d:
    st.pyplot(viz.plot_errors(hourly), use_container_width=True)

# --------------------------------------------------------------------------
# 5. Synth\u00e8se par produit
# --------------------------------------------------------------------------
st.subheader("Synth\u00e8se par produit")
df_produits = dp.product_synthesis(df)
st.markdown(ui.product_table(df_produits), unsafe_allow_html=True)

# --------------------------------------------------------------------------
# 6. Messages d'erreur + export conditionnel
# --------------------------------------------------------------------------
st.subheader("Messages d'erreur")
df_erreurs = dp.error_messages(df)
st.markdown(ui.error_table(df_erreurs), unsafe_allow_html=True)

if len(df_erreurs) > SEUIL_EXPORT_ERREURS:
    st.download_button(
        label="\u2b07\ufe0f T\u00e9l\u00e9charger les erreurs (Excel)",
        data=dp.to_excel_bytes(df_erreurs, sheet_name="Erreurs"),
        file_name="rapport_qos_erreurs.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
