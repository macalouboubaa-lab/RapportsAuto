"""Couche de traitement des données QoS.

Toutes les fonctions sont pures (pas d'appel Streamlit, pas d'I/O disque).
Elles prennent / renvoient des objets Pandas ou des structures simples afin
de rester testables et réutilisables.

Adaptez au besoin les constantes de noms de colonnes ci-dessous pour coller
à l'export de votre outil interne.
"""

from __future__ import annotations

import io
from datetime import datetime

import pandas as pd

# ---------------------------------------------------------------------------
# Noms de colonnes attendus dans le fichier Excel exporté
# Centralisés ici : si l'export change, on ne modifie qu'un seul endroit.
# ---------------------------------------------------------------------------
COL_PAYS = "CODE PAYS ISO"
COL_FILIALE = "FILIALE"
COL_DATE = "DATE TRANSACTION"
COL_TRANSACTION = "NUM\u00c9RO TRANSACTION GU"
COL_STATUT = "STATUT"
COL_PRODUIT = "PRODUIT"
COL_SERVICE = "SERVICE"
COL_MESSAGE = "MESSAGE"

# Valeurs de statut normalisées (en MAJUSCULES)
STATUT_SUCCESS = "SUCCESSFUL"
STATUT_FAILED = "FAILED"
STATUT_PENDING = "PENDING"
STATUT_INITIATED = "INITIATED"

# Statuts considérés comme « réussis » pour le calcul de QoS
STATUTS_OK = {STATUT_SUCCESS}

TOUTES_FILIALES = "TOUTES LES FILIALES"


def load_excel(uploaded_file) -> pd.DataFrame:
    """Charge un fichier .xlsx (objet UploadedFile Streamlit ou chemin).

    La lecture se fait en mémoire : aucun écriture sur le disque.
    """
    data = uploaded_file.read() if hasattr(uploaded_file, "read") else uploaded_file
    buffer = io.BytesIO(data) if isinstance(data, (bytes, bytearray)) else uploaded_file
    df = pd.read_excel(buffer, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    return df


def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Nettoyage de base commun à toute la chaîne de traitement.

    - retire les apostrophes de la colonne NUMÉRO TRANSACTION GU
    - normalise le statut en MAJUSCULES
    - convertit la colonne date en datetime
    """
    df = df.copy()

    if COL_TRANSACTION in df.columns:
        df[COL_TRANSACTION] = (
            df[COL_TRANSACTION].astype(str).str.replace("'", "", regex=False).str.strip()
        )

    if COL_STATUT in df.columns:
        df[COL_STATUT] = df[COL_STATUT].astype(str).str.strip().str.upper()

    if COL_DATE in df.columns:
        df[COL_DATE] = pd.to_datetime(df[COL_DATE], errors="coerce")

    return df


def drop_current_hour(df: pd.DataFrame, now: datetime | None = None) -> pd.DataFrame:
    """Supprime les transactions de l'heure en cours.

    On ne conserve que les lignes strictement antérieures au début de
    l'heure courante (ex : à 14h37, on garde tout ce qui est < 14h00).
    """
    if COL_DATE not in df.columns:
        return df
    now = now or datetime.now()
    debut_heure = now.replace(minute=0, second=0, microsecond=0)
    return df[df[COL_DATE] < debut_heure].copy()


def list_filiales(df: pd.DataFrame) -> list[str]:
    """Liste des filiales disponibles pour le menu déroulant."""
    if COL_FILIALE not in df.columns:
        return [TOUTES_FILIALES]
    valeurs = sorted(str(v) for v in df[COL_FILIALE].dropna().unique())
    return [TOUTES_FILIALES] + valeurs


def filter_data(df: pd.DataFrame, filiale: str | None = None) -> pd.DataFrame:
    """Filtre le DataFrame par filiale (None ou TOUTES = pas de filtre)."""
    if not filiale or filiale == TOUTES_FILIALES or COL_FILIALE not in df.columns:
        return df.copy()
    return df[df[COL_FILIALE].astype(str) == str(filiale)].copy()


def _is_success(serie: pd.Series) -> pd.Series:
    return serie.isin(STATUTS_OK)


def compute_kpis(df: pd.DataFrame) -> dict:
    """Calcule les indicateurs globaux pour les cartes KPI."""
    total = int(len(df))
    if total == 0 or COL_STATUT not in df.columns:
        return {
            "total": total,
            "reussies": 0,
            "qos_globale": 0.0,
            "echecs_attente": 0,
            "echecs_attente_pct": 0.0,
            "creneau_critique": "\u2014",
            "creneau_critique_qos": 0.0,
        }

    reussies = int(_is_success(df[COL_STATUT]).sum())
    echecs_attente = total - reussies
    qos_globale = round(100.0 * reussies / total, 1)
    echecs_pct = round(100.0 * echecs_attente / total, 1)

    creneau, creneau_qos = _creneau_le_plus_critique(df)

    return {
        "total": total,
        "reussies": reussies,
        "qos_globale": qos_globale,
        "echecs_attente": echecs_attente,
        "echecs_attente_pct": echecs_pct,
        "creneau_critique": creneau,
        "creneau_critique_qos": creneau_qos,
    }


def _creneau_le_plus_critique(df: pd.DataFrame) -> tuple[str, float]:
    """Renvoie (libellé heure, QoS %) du créneau horaire le plus bas."""
    if COL_DATE not in df.columns or df[COL_DATE].isna().all():
        return "\u2014", 0.0
    tmp = df.dropna(subset=[COL_DATE]).copy()
    tmp["__heure"] = tmp[COL_DATE].dt.hour
    grp = tmp.groupby("__heure")[COL_STATUT]
    stats = grp.apply(lambda s: 100.0 * _is_success(s).sum() / len(s))
    if stats.empty:
        return "\u2014", 0.0
    heure = int(stats.idxmin())
    return f"{heure:02d}h\u2013{(heure + 1) % 24:02d}h", round(float(stats.min()), 1)


def hourly_stats(df: pd.DataFrame) -> pd.DataFrame:
    """Agrégation horaire : volume total, volume en échec et QoS %.

    Renvoie un DataFrame indexé par heure (0..23 présentes dans les données)
    avec les colonnes : heure, volume, erreurs, qos.
    """
    cols = ["heure", "volume", "erreurs", "qos"]
    if COL_DATE not in df.columns or COL_STATUT not in df.columns or df.empty:
        return pd.DataFrame(columns=cols)

    tmp = df.dropna(subset=[COL_DATE]).copy()
    if tmp.empty:
        return pd.DataFrame(columns=cols)

    tmp["heure"] = tmp[COL_DATE].dt.hour
    tmp["__ok"] = _is_success(tmp[COL_STATUT])

    agg = tmp.groupby("heure").agg(
        volume=("__ok", "size"),
        reussies=("__ok", "sum"),
    )
    agg["erreurs"] = agg["volume"] - agg["reussies"]
    agg["qos"] = (100.0 * agg["reussies"] / agg["volume"]).round(1)
    agg = agg.reset_index().sort_values("heure")
    return agg[cols].reset_index(drop=True)


def product_synthesis(df: pd.DataFrame) -> pd.DataFrame:
    """Synthèse par produit.

    Colonnes : PRODUIT, TOTAL, SUCCESSFUL, FAILED, PENDING, INITIATED, QOS.
    """
    cols = ["PRODUIT", "TOTAL", "SUCCESSFUL", "FAILED", "PENDING", "INITIATED", "QOS"]
    if COL_PRODUIT not in df.columns or COL_STATUT not in df.columns or df.empty:
        return pd.DataFrame(columns=cols)

    tmp = df.copy()
    pivot = (
        pd.crosstab(tmp[COL_PRODUIT], tmp[COL_STATUT])
        .reset_index()
        .rename(columns={COL_PRODUIT: "PRODUIT"})
    )
    pivot.columns.name = None

    for statut in (STATUT_SUCCESS, STATUT_FAILED, STATUT_PENDING, STATUT_INITIATED):
        if statut not in pivot.columns:
            pivot[statut] = 0

    pivot["TOTAL"] = (
        pivot[[STATUT_SUCCESS, STATUT_FAILED, STATUT_PENDING, STATUT_INITIATED]]
        .sum(axis=1)
    )
    pivot["QOS"] = (
        (100.0 * pivot[STATUT_SUCCESS] / pivot["TOTAL"].where(pivot["TOTAL"] != 0, 1))
        .round(1)
    )

    pivot = pivot.rename(
        columns={
            STATUT_SUCCESS: "SUCCESSFUL",
            STATUT_FAILED: "FAILED",
            STATUT_PENDING: "PENDING",
            STATUT_INITIATED: "INITIATED",
        }
    )
    pivot = pivot.sort_values("TOTAL", ascending=False)
    return pivot[cols].reset_index(drop=True)


def error_messages(df: pd.DataFrame) -> pd.DataFrame:
    """Top des messages d'erreur (statuts non réussis).

    Colonnes : SERVICE, MESSAGE, OCCURRENCES, % DU TOTAL.
    Le pourcentage est calculé sur le total des transactions en erreur.
    """
    cols = ["SERVICE", "MESSAGE", "OCCURRENCES", "% DU TOTAL"]
    if COL_STATUT not in df.columns or df.empty:
        return pd.DataFrame(columns=cols)

    erreurs = df[~_is_success(df[COL_STATUT])].copy()
    if erreurs.empty:
        return pd.DataFrame(columns=cols)

    if COL_SERVICE not in erreurs.columns:
        erreurs[COL_SERVICE] = "\u2014"
    if COL_MESSAGE not in erreurs.columns:
        erreurs[COL_MESSAGE] = "\u2014"

    erreurs[COL_MESSAGE] = erreurs[COL_MESSAGE].fillna("\u2014").astype(str).str.strip()
    erreurs[COL_SERVICE] = erreurs[COL_SERVICE].fillna("\u2014").astype(str).str.strip()

    total_erreurs = len(erreurs)
    grp = (
        erreurs.groupby([COL_SERVICE, COL_MESSAGE])
        .size()
        .reset_index(name="OCCURRENCES")
    )
    grp["% DU TOTAL"] = (100.0 * grp["OCCURRENCES"] / total_erreurs).round(1)
    grp = grp.rename(columns={COL_SERVICE: "SERVICE", COL_MESSAGE: "MESSAGE"})
    grp = grp.sort_values("OCCURRENCES", ascending=False)
    return grp[cols].reset_index(drop=True)


def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Erreurs") -> bytes:
    """Sérialise un DataFrame en fichier Excel (.xlsx) en mémoire.

    Renvoie des octets prêts pour st.download_button (aucune écriture disque).
    """
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    buffer.seek(0)
    return buffer.getvalue()
