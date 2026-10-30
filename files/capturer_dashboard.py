#!/usr/bin/env python3
r"""
Capture automatique du dashboard Supervision QoS Mobile Money, filiale par filiale,
au format prêt à copier-coller dans Teams (texte + images + éventuelle pièce jointe Excel).

Usage :
    python capturer_dashboard.py dashboard.html fichier_nettoye.xlsx dossier_sortie [--heure-actuelle "09h"]

Pour chaque filiale, dans dossier_sortie\<groupe>\<entite>\ :
  - message.txt              : le texte à coller dans Teams
  - kpis_graphiques.png       : les 4 indicateurs + les 2 graphiques
  - synthese_produits.png     : le tableau par produit (lignes QoS < 50% encadrées en rouge)
  - erreurs.png                OU  messages_erreur.xlsx  (selon le volume de messages distincts)
"""

import argparse
import os
import re
import sys
from datetime import datetime

from playwright.sync_api import sync_playwright

SEUIL_QOS_ALERTE = 50          # % — en dessous, produit encadré en rouge + cité dans le message
SEUIL_MIN_TRANSACTIONS = 5     # produits avec 5 transactions ou moins ignorés (pas assez de volume)
SEUIL_BEAUCOUP_ERREURS = 25    # messages distincts — au-delà, on joint le fichier Excel au lieu de capturer


def est_groupe_sn(entite_normalisee: str) -> bool:
    """SN et SN PI utilisent le suffixe historique _SA (pas un code pays) -> groupe 1."""
    return bool(re.fullmatch(r"INTOUCH_SA(_PI)?", entite_normalisee))


def libelle_filiale(entite: str) -> str:
    """GU3 SN / GU3 SN PI / GU3 CI / GU3 CI PI / GU3 BF / GU3 GN / GU3 ML / GU3 TG / GU3 BJ ..."""
    e = entite.upper()
    pi = e.endswith("_PI")
    base = e[:-3] if pi else e
    code = "SN" if base.endswith("_SA") else base.replace("INTOUCH_", "")
    return f"GU3 {code}" + (" PI" if pi else "")


def capturer(chemin_dashboard: str, chemin_excel: str, dossier_sortie: str, heure_label: str | None = None) -> list[str]:
    chemin_dashboard = os.path.abspath(chemin_dashboard)
    chemin_excel = os.path.abspath(chemin_excel)

    if not os.path.isfile(chemin_dashboard) or not chemin_dashboard.lower().endswith((".html", ".htm")):
        raise SystemExit(
            f"ERREUR : le fichier dashboard est introuvable ou invalide : {chemin_dashboard}\n"
            "Vérifiez qu'un fichier nommé exactement 'dashboard.html' se trouve bien dans le même "
            "dossier que les scripts, et que l'ordre des arguments est : fichier_brut, dashboard, dossier_sortie."
        )

    heure_label = heure_label or f"{datetime.now().hour:02d}H"
    heure_msg = heure_label.rstrip("Hh")

    rapport: list[str] = []

    with sync_playwright() as pw:
        browser = pw.chromium.launch()
        page = browser.new_page(viewport={"width": 1400, "height": 1000})
        page.goto(f"file:///{chemin_dashboard.replace(os.sep, '/')}")

        page.set_input_files("#fileInput", chemin_excel)
        page.wait_for_selector("#uploadStatus.ok", timeout=15000)

        # Liste réelle et complète des filiales présentes dans le menu déroulant
        # (on exclut l'option "Tous", et on répartit automatiquement SN/SN PI vs le reste,
        # donc toute nouvelle filiale — ex: BJ — est prise en compte sans modifier le script)
        options = page.eval_on_selector_all(
            "#f_entite option",
            "els => els.map(e => ({value: e.value, texte: e.textContent})).filter(o => o.value && o.value !== '__ALL__')",
        )
        entites_groupe1 = [o["value"] for o in options if est_groupe_sn(o["value"].strip().upper())]
        entites_groupe2 = [o["value"] for o in options if not est_groupe_sn(o["value"].strip().upper())]

        def traiter(valeur_reelle: str, dossier_groupe: str):
            page.select_option("#f_entite", valeur_reelle)
            page.wait_for_timeout(600)

            # remonter en haut et agrandir la fenêtre à la hauteur exacte de la page :
            # évite tout rognage d'en-tête lié au défilement, quelle que soit la taille du contenu
            page.evaluate("window.scrollTo(0, 0)")
            hauteur_page = page.evaluate("document.body.scrollHeight")
            page.set_viewport_size({"width": 1400, "height": int(hauteur_page) + 50})
            page.wait_for_timeout(150)

            dossier_entite = os.path.join(dossier_groupe, valeur_reelle)
            os.makedirs(dossier_entite, exist_ok=True)

            # --- Produits en QoS impactée (pour le message + l'encadré rouge) ---
            # On ignore les produits avec trop peu de transactions (pas assez de volume pour être significatif)
            lignes_produits = page.eval_on_selector_all(
                "#produitTableBody tr",
                """els => els.map(tr => {
                    const tds = tr.querySelectorAll('td');
                    if (tds.length < 7) return null;
                    const produit = tds[0].textContent.trim();
                    const totalText = tds[1].textContent.trim().replace(/\\s/g,'').replace(',','.');
                    const total = parseFloat(totalText);
                    const qosText = tds[6].textContent.trim().replace('%','').replace(',','.');
                    const qos = parseFloat(qosText);
                    return (isNaN(qos) || isNaN(total)) ? null : {produit, qos, total};
                }).filter(Boolean)""",
            )
            produits_impactes = [
                l["produit"] for l in lignes_produits
                if l["qos"] < SEUIL_QOS_ALERTE and l["total"] > SEUIL_MIN_TRANSACTIONS
            ]

            # encadrer en rouge les lignes impactées avant la capture (hors produits à faible volume)
            page.eval_on_selector_all(
                "#produitTableBody tr",
                f"""(els, params) => els.forEach(tr => {{
                    const tds = tr.querySelectorAll('td');
                    if (tds.length < 7) return;
                    const total = parseFloat(tds[1].textContent.trim().replace(/\\s/g,'').replace(',','.'));
                    const qos = parseFloat(tds[6].textContent.trim().replace('%','').replace(',','.'));
                    const impacte = !isNaN(qos) && !isNaN(total) && qos < params.seuil && total > params.seuilMin;
                    tr.style.outline = impacte ? '2px solid #ef4444' : '';
                    tr.style.outlineOffset = '-2px';
                }})""",
                {"seuil": SEUIL_QOS_ALERTE, "seuilMin": SEUIL_MIN_TRANSACTIONS},
            )

            # --- Image 1 : KPIs + graphiques ---
            box_kpis = page.locator(".kpis").bounding_box()
            box_charts = page.locator(".grid2.grid2-equal").bounding_box()
            if box_kpis and box_charts:
                x0 = min(box_kpis["x"], box_charts["x"])
                y0 = box_kpis["y"]
                x1 = max(box_kpis["x"] + box_kpis["width"], box_charts["x"] + box_charts["width"])
                y1 = box_charts["y"] + box_charts["height"]
                page.screenshot(
                    path=os.path.join(dossier_entite, "kpis_graphiques.png"),
                    clip={"x": x0, "y": y0, "width": x1 - x0, "height": y1 - y0},
                )

            # --- Image 2 : synthèse par produit (avec encadrés rouges) ---
            panel_produits = page.locator("h2:has-text('Synthèse par produit')").locator("xpath=..")
            panel_produits.screenshot(path=os.path.join(dossier_entite, "synthese_produits.png"))

            # remet le tableau propre pour ne pas fausser la prochaine capture
            page.eval_on_selector_all(
                "#produitTableBody tr", "els => els.forEach(tr => { tr.style.outline=''; })"
            )

            # --- Détail des erreurs : capture, PDF (si tronqué), ou export Excel selon le volume ---
            sous_texte = page.locator("#tableSub").text_content() or ""
            m = re.search(r"(\d+)\s+messages distincts", sous_texte)
            nb_messages = int(m.group(1)) if m else 0

            fichier_erreur_joint = None
            chemin_erreurs = os.path.join(dossier_entite, "erreurs.png")
            if nb_messages > SEUIL_BEAUCOUP_ERREURS:
                with page.expect_download() as dl_info:
                    page.click("#btnExportXlsx")
                download = dl_info.value
                fichier_erreur_joint = os.path.join(dossier_entite, "messages_erreur.xlsx")
                download.save_as(fichier_erreur_joint)
            else:
                # 0 à 20 messages : on capture toujours (même à 0, pour montrer qu'il n'y a rien à signaler)
                tronque = page.eval_on_selector(
                    "#errTableBody",
                    "el => { const s = el.closest('.table-scroll'); return s.scrollHeight > s.clientHeight + 2; }",
                )
                if not tronque:
                    panel_erreurs = page.locator("h2:has-text(\"Détail des messages d'erreur par service\")").locator("xpath=../..")
                    panel_erreurs.screenshot(path=chemin_erreurs)
                else:
                    # tableau tronqué -> on capture la vue complète depuis l'onglet PDF (pas de scroll là-bas)
                    with page.expect_popup() as popup_info:
                        page.click("#btnExportPdf")
                    popup = popup_info.value
                    popup.evaluate("window.print = () => {}")  # empêche le dialogue d'impression natif
                    popup.wait_for_load_state("domcontentloaded")
                    popup.locator("table").screenshot(path=chemin_erreurs)
                    popup.close()

            # --- Texte du message (heure, code filiale et produits impactés en gras) ---
            libelle = libelle_filiale(valeur_reelle)
            lignes_texte = [
                "Hello Team,",
                f"Veuillez trouver ci-joint le rapport QoS par produit **{libelle}** de **{heure_msg}H**.",
            ]
            lignes_html = [
                "Hello Team,",
                f"Veuillez trouver ci-joint le rapport QoS par produit <b>{libelle}</b> de <b>{heure_msg}H</b>.",
            ]
            if fichier_erreur_joint:
                phrase = (
                    "Le fichier joint contient également l'<b>ensemble des SMS d'erreur par service</b>, "
                    "afin de faciliter des analyses des incidents et le suivi des dysfonctionnements."
                )
                lignes_texte.append(phrase)
                lignes_html.append(phrase)
            if produits_impactes:
                lignes_texte.append(
                    f"⚠️ Attention aux produits {', '.join(f'**{p}**' for p in produits_impactes)} QoS impactée"
                )
                lignes_html.append(
                    f"⚠️ Attention aux produits {', '.join(f'<b>{p}</b>' for p in produits_impactes)} QoS impactée"
                )

            with open(os.path.join(dossier_entite, "message.txt"), "w", encoding="utf-8") as f:
                f.write("\n".join(lignes_texte) + "\n")

            # message.html : à OUVRIR DANS LE NAVIGATEUR, puis Ctrl+A / Ctrl+C / Ctrl+V dans Teams
            # (copier depuis un navigateur conserve le gras ; copier depuis le Bloc-notes ne le peut pas)
            corps_html = "<br>".join(lignes_html)
            with open(os.path.join(dossier_entite, "message.html"), "w", encoding="utf-8") as f:
                f.write(
                    "<html><body style=\"font-family:Segoe UI, Arial, sans-serif; font-size:14px;\">"
                    f"{corps_html}</body></html>"
                )

            rapport.append(
                f"OK : {valeur_reelle} -> {dossier_entite}"
                + (f" (+ {os.path.basename(fichier_erreur_joint)})" if fichier_erreur_joint else "")
            )

        dossier_g1 = os.path.join(dossier_sortie, "groupe1_SN")
        dossier_g2 = os.path.join(dossier_sortie, "groupe2_autres")

        for entite in entites_groupe1:
            traiter(entite, dossier_g1)
        for entite in entites_groupe2:
            traiter(entite, dossier_g2)

        browser.close()

    return rapport


def main():
    parser = argparse.ArgumentParser(description="Capture du dashboard QoS par filiale, prêt pour Teams")
    parser.add_argument("dashboard", help="Chemin du fichier dashboard.html")
    parser.add_argument("fichier_excel", help="Chemin du fichier Excel nettoyé à déposer")
    parser.add_argument("dossier_sortie", help="Dossier où ranger les captures")
    parser.add_argument("--heure-actuelle", dest="heure_label", default=None, help="Ex: 09H (sinon heure système)")
    args = parser.parse_args()

    rapport = capturer(args.dashboard, args.fichier_excel, args.dossier_sortie, args.heure_label)
    print("\n".join(rapport))


if __name__ == "__main__":
    sys.exit(main())