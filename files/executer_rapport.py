#!/usr/bin/env python3
r"""
Exécute le processus complet, une fois par heure :
  1. Nettoie le fichier brut ListTransactions.
  2. Dépose le résultat dans le dashboard et capture chaque filiale.

Usage :
    python executer_rapport.py fichier_brut.xlsx dashboard.html dossier_sortie
    python executer_rapport.py fichier_brut.xlsx dashboard.html dossier_sortie --heure-actuelle "2026-07-31 12:00"
"""

import argparse
import os
import sys
from datetime import datetime

from nettoyer_rapport import nettoyer, sauvegarder
from capturer_dashboard import capturer


def main():
    parser = argparse.ArgumentParser(description="Nettoyage + captures du dashboard, en une seule exécution")
    parser.add_argument("fichier_brut", help="Chemin du fichier Excel brut (export)")
    parser.add_argument("dashboard", help="Chemin du fichier dashboard.html")
    parser.add_argument("dossier_sortie", help="Dossier racine (sous-dossier par heure créé automatiquement)")
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

    # 1. Nettoyage
    df = nettoyer(args.fichier_brut, heure_actuelle)
    nom_dossier_heure = f"{heure_reference.hour:02d}H"
    dossier_heure = os.path.join(args.dossier_sortie, nom_dossier_heure)
    os.makedirs(dossier_heure, exist_ok=True)
    chemin_excel_nettoye = os.path.join(dossier_heure, "ListTransactions_nettoye.xlsx")
    sauvegarder(df, chemin_excel_nettoye)
    print(f"OK NETTOYAGE : {len(df)} lignes écrites dans {chemin_excel_nettoye}")

    # 2. Captures du dashboard, rangées dans le même dossier horaire
    dossier_captures = os.path.join(dossier_heure, "captures")
    rapport = capturer(args.dashboard, chemin_excel_nettoye, dossier_captures, heure_label=nom_dossier_heure)
    print("\n".join(rapport))


if __name__ == "__main__":
    sys.exit(main())
