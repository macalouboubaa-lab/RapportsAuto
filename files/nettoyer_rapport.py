#!/usr/bin/env python3
r"""
Nettoyage automatique du fichier "ListTransactions" (export brut -> fichier prêt pour le dashboard).

Usage :
    python nettoyer_rapport.py fichier_brut.xlsx dossier_sortie
    python nettoyer_rapport.py fichier_brut.xlsx dossier_sortie --heure-actuelle "2026-07-31 12:00"

Le fichier nettoyé est enregistré dans un sous-dossier nommé par heure
(ex: dossier_sortie\09H\ListTransactions_nettoye.xlsx), pour ne pas
écraser les résultats des heures précédentes.

Règles appliquées (confirmées avec l'utilisateur) :
  1. On ne garde que les données jusqu'à l'heure PLEINE précédente
     (si on lance le script à 9h, on garde jusqu'à 8h59m59s inclus).
  2. Colonnes supprimées : ID CLIENT + toutes les colonnes après CODE FLUX
     (COMMISSION GLOBALE, FRAIS GLOBAUX, DATE DE MODIFICATION, ID PARTENAIRE,
     UTILISATEUR, ID PARTENAIRE DIST, BANQUE, COMPTE EMETTEUR, COMPTE ACQUÉREUR, TAG).
  3. TYPE SERVICE : on supprime les lignes A2M_TRANSFER, CLIENT_COMPENSATION,
     CLIENT_TOPUP, M2A_TRANSFER, A2A, ou toute valeur purement numérique.
  4. PRODUIT : on supprime les lignes AGENT_API et BOCORPORATE.
  5. CODE PAYS ISO vide -> déduit du préfixe de la colonne SERVICE (ex: SN_MPM_ORM -> SN).
  6. Services PI (CI et SN uniquement) : si SERVICE contient "CCT_TPT",
     on ajoute le suffixe "_PI" à ENTITÉ JURIDIQUE.
  7. MONTANT : converti en nombre entier.
  8. DATE : la vraie valeur (avec minutes/secondes) est conservée, seul
     l'affichage est changé en format m/jj/aaaa h:mm (secondes masquées).
"""

import argparse
import os
import re
import sys
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

COLONNES_A_GARDER = [
    "SERVICE", "CODE PAYS ISO", "ENTITÉ JURIDIQUE", "DATE", "MONTANT",
    "NUMÉRO TRANSACTION GU", "TYPE SERVICE", "STATUT", "PRODUIT",
    "MESSAGE", "CODE FLUX",
]

TYPE_SERVICE_EXCLUS = {
    "A2M_TRANSFER", "CLIENT_COMPENSATION", "CLIENT_TOPUP",
    "M2A_TRANSFER", "A2A",
}
PRODUIT_EXCLUS = {"AGENT_API", "BOCORPORATE"}


# Noms alternatifs de colonnes (ex: export mensuel en minuscules avec underscores) -> nom standard
ALIAS_COLONNES = {
    "service": "SERVICE",
    "code_pays": "CODE PAYS ISO",
    "code_pays_iso": "CODE PAYS ISO",
    "entite_juridique": "ENTITÉ JURIDIQUE",
    "date_creation": "DATE",
    "date": "DATE",
    "montant": "MONTANT",
    "numero_transaction_gu": "NUMÉRO TRANSACTION GU",
    "type_service": "TYPE SERVICE",
    "statut": "STATUT",
    "produit": "PRODUIT",
    "message": "MESSAGE",
    "code_flux": "CODE FLUX",
}


def normaliser_colonnes(df: pd.DataFrame) -> pd.DataFrame:
    renommage = {}
    for c in df.columns:
        propre = str(c).strip()
        cle = propre.lower().replace(" ", "_")
        renommage[c] = ALIAS_COLONNES.get(cle, propre)
    return df.rename(columns=renommage)


def nettoyer(chemin_brut: str, heure_actuelle: datetime | None = None, appliquer_filtre_horaire: bool = True) -> pd.DataFrame:
    df = pd.read_excel(chemin_brut)
    df = normaliser_colonnes(df)  # accepte aussi les noms en minuscules/underscores (export mensuel)

    if "DATE" not in df.columns:
        raise SystemExit(
            "ERREUR : colonne DATE introuvable dans le fichier brut.\n"
            f"Colonnes trouvées : {list(df.columns)}"
        )

    # DATE : parsing (le brut est en texte JJ/MM/AAAA HH:MM:SS)
    # errors="coerce" : les valeurs invalides (ex: données corrompues) deviennent NaT au lieu de planter
    df["DATE"] = pd.to_datetime(df["DATE"], dayfirst=True, errors="coerce")
    nb_dates_invalides = df["DATE"].isna().sum()
    if nb_dates_invalides:
        print(f"ATTENTION : {nb_dates_invalides} ligne(s) avec une DATE invalide ignorée(s)", file=sys.stderr)
        df = df[df["DATE"].notna()].copy()

    # 1. Filtre temporel : on garde jusqu'à l'heure pleine précédente (rapport horaire uniquement)
    # (comparaison sur l'heure de la journée uniquement, peu importe la date de la ligne)
    if appliquer_filtre_horaire:
        if heure_actuelle is None:
            heure_actuelle = datetime.now()
        cutoff_time = heure_actuelle.replace(minute=0, second=0, microsecond=0).time()
        df = df[df["DATE"].dt.time < cutoff_time].copy()

    # 2. Colonnes à garder uniquement
    colonnes_presentes = [c for c in COLONNES_A_GARDER if c in df.columns]
    df = df[colonnes_presentes]

    # 3. Filtre TYPE SERVICE (on garde les valeurs vides, elles sont légitimes)
    def type_service_a_exclure(val):
        if pd.isna(val):
            return False
        val = str(val).strip()
        if val in TYPE_SERVICE_EXCLUS:
            return True
        if re.fullmatch(r"\d+", val):  # valeur purement numérique
            return True
        return False

    if "TYPE SERVICE" in df.columns:
        df = df[~df["TYPE SERVICE"].apply(type_service_a_exclure)]

    # 4. Filtre PRODUIT
    if "PRODUIT" in df.columns:
        df = df[~df["PRODUIT"].isin(PRODUIT_EXCLUS)]

    # 5. CODE PAYS ISO vide -> déduit du préfixe de SERVICE
    if "SERVICE" in df.columns and "CODE PAYS ISO" in df.columns:
        prefixe_service = df["SERVICE"].str.split("_").str[0]
        df["CODE PAYS ISO"] = df["CODE PAYS ISO"].fillna(prefixe_service)

    # 5bis. ENTITÉ JURIDIQUE vide (sécurité, pas rencontré dans les fichiers de référence)
    # -> on prend l'entité la plus fréquente déjà associée à ce CODE PAYS ISO dans le fichier
    if "ENTITÉ JURIDIQUE" in df.columns and "CODE PAYS ISO" in df.columns and df["ENTITÉ JURIDIQUE"].isna().any():
        entite_par_pays = (
            df.dropna(subset=["ENTITÉ JURIDIQUE"])
            .groupby("CODE PAYS ISO")["ENTITÉ JURIDIQUE"]
            .agg(lambda s: s.value_counts().idxmax())
        )
        df["ENTITÉ JURIDIQUE"] = df["ENTITÉ JURIDIQUE"].fillna(df["CODE PAYS ISO"].map(entite_par_pays))

    # 6. Services PI (CCT_TPT) -> suffixe _PI sur ENTITÉ JURIDIQUE
    if "SERVICE" in df.columns and "ENTITÉ JURIDIQUE" in df.columns:
        est_pi = df["SERVICE"].str.contains("CCT_TPT", na=False)
        deja_pi = df["ENTITÉ JURIDIQUE"].str.endswith("_PI", na=False)
        df.loc[est_pi & ~deja_pi, "ENTITÉ JURIDIQUE"] = df.loc[est_pi & ~deja_pi, "ENTITÉ JURIDIQUE"] + "_PI"

    # 7. MONTANT en entier
    if "MONTANT" in df.columns:
        df["MONTANT"] = pd.to_numeric(df["MONTANT"], errors="coerce").round(0).fillna(0).astype("int64")

    return df.reset_index(drop=True)


def sauvegarder(df: pd.DataFrame, chemin_sortie: str) -> None:
    with pd.ExcelWriter(chemin_sortie, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Transactions")

    # Mise en forme : police Arial + format de date sans les secondes
    wb = load_workbook(chemin_sortie)
    ws = wb["Transactions"]

    date_col_idx = df.columns.get_loc("DATE") + 1  # openpyxl est en base 1
    date_col_letter = get_column_letter(date_col_idx)

    for row in ws.iter_rows():
        for cell in row:
            cell.font = Font(name="Arial", size=10)

    for cell in ws[date_col_letter][1:]:  # on saute l'en-tête
        cell.number_format = "m/dd/yyyy h:mm"

    # Largeur des colonnes
    for col_idx, col_name in enumerate(df.columns, start=1):
        largeur = max(12, min(30, len(col_name) + 2))
        ws.column_dimensions[get_column_letter(col_idx)].width = largeur

    wb.save(chemin_sortie)


def main():
    parser = argparse.ArgumentParser(description="Nettoyage automatique du fichier ListTransactions")
    parser.add_argument("fichier_brut", help="Chemin du fichier Excel brut (export)")
    parser.add_argument("dossier_sortie", help="Dossier où créer le sous-dossier par heure (ex: RapportsAuto)")
    parser.add_argument(
        "--heure-actuelle",
        help="Pour test uniquement : force l'heure de référence (ex: '2026-07-31 12:00')",
        default=None,
    )
    args = parser.parse_args()

    heure_actuelle = None
    if args.heure_actuelle:
        heure_actuelle = datetime.strptime(args.heure_actuelle, "%Y-%m-%d %H:%M")

    heure_reference = heure_actuelle or datetime.now()
    df = nettoyer(args.fichier_brut, heure_actuelle)

    # sous-dossier nommé par l'heure de la coupure (ex: 09H, 16H)
    nom_dossier_heure = f"{heure_reference.hour:02d}H"
    dossier_final = os.path.join(args.dossier_sortie, nom_dossier_heure)
    os.makedirs(dossier_final, exist_ok=True)
    chemin_sortie = os.path.join(dossier_final, "ListTransactions_nettoye.xlsx")

    sauvegarder(df, chemin_sortie)
    print(f"OK : {len(df)} lignes écrites dans {chemin_sortie}")


if __name__ == "__main__":
    sys.exit(main())
