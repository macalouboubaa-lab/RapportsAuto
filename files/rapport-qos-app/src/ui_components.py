"""Couche présentation : génération des blocs HTML/CSS.

Chaque fonction renvoie une chaîne HTML destinée à ``st.markdown(...,
unsafe_allow_html=True)``. Tout le contenu provenant des données est
échappé via ``html.escape`` pour éviter toute injection.
"""

from __future__ import annotations

from html import escape

import pandas as pd

VERT = "#22C55E"
ROUGE = "#EF4444"
FOND_CARTE = "#1F2937"
FOND_PAGE = "#111827"
TEXTE = "#E5E7EB"
TEXTE_DOUX = "#9CA3AF"
BORDURE = "#374151"


def global_css() -> str:
    """Feuille de style globale injectée une seule fois au démarrage."""
    return f"""
<style>
.kpi-grid {{
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 16px;
    margin: 8px 0 24px 0;
}}
@media (max-width: 900px) {{ .kpi-grid {{ grid-template-columns: repeat(2, 1fr); }} }}
.kpi-card {{
    background: {FOND_CARTE};
    border: 1px solid {BORDURE};
    border-radius: 14px;
    padding: 18px 20px;
    box-shadow: 0 2px 10px rgba(0,0,0,.35);
}}
.kpi-card .kpi-label {{
    color: {TEXTE_DOUX};
    font-size: .78rem;
    text-transform: uppercase;
    letter-spacing: .06em;
    margin-bottom: 8px;
}}
.kpi-card .kpi-value {{ color: {TEXTE}; font-size: 1.9rem; font-weight: 700; line-height: 1.1; }}
.kpi-card .kpi-sub {{ color: {TEXTE_DOUX}; font-size: .82rem; margin-top: 6px; }}
.kpi-accent-left {{ border-left: 4px solid {{accent}}; }}

.qos-table {{
    width: 100%;
    border-collapse: collapse;
    background: {FOND_CARTE};
    border-radius: 12px;
    overflow: hidden;
    font-size: .9rem;
    margin: 8px 0 20px 0;
}}
.qos-table th {{
    background: #0F172A;
    color: {TEXTE};
    text-align: left;
    padding: 10px 14px;
    font-weight: 600;
    border-bottom: 1px solid {BORDURE};
}}
.qos-table td {{ color: {TEXTE}; padding: 9px 14px; border-bottom: 1px solid {BORDURE}; }}
.qos-table tr:last-child td {{ border-bottom: none; }}
.qos-ok {{ color: {VERT}; font-weight: 700; }}
.qos-ko {{ color: {ROUGE}; font-weight: 700; }}
.pastille {{
    display: inline-block;
    min-width: 34px;
    text-align: center;
    background: {ROUGE};
    color: #fff;
    font-weight: 700;
    padding: 2px 10px;
    border-radius: 999px;
    font-size: .82rem;
}}
</style>
"""


def _card(label: str, value: str, sub: str, accent: str) -> str:
    return (
        f'<div class="kpi-card" style="border-left:4px solid {accent};">'
        f'<div class="kpi-label">{escape(label)}</div>'
        f'<div class="kpi-value">{escape(value)}</div>'
        f'<div class="kpi-sub">{escape(sub)}</div>'
        f"</div>"
    )


def kpi_cards(kpis: dict) -> str:
    """Génère les 4 cartes KPI."""
    accent_qos = VERT if kpis.get("qos_globale", 0) >= 50 else ROUGE

    c1 = _card(
        "Op\u00e9rations",
        f"{kpis['total']:,}".replace(",", "\u202f"),
        f"{kpis['reussies']:,} r\u00e9ussies".replace(",", "\u202f"),
        "#3B82F6",
    )
    c2 = _card(
        "QoS globale",
        f"{kpis['qos_globale']:.1f}\u202f%",
        "Taux de r\u00e9ussite",
        accent_qos,
    )
    c3 = _card(
        "\u00c9checs / attente",
        f"{kpis['echecs_attente']:,}".replace(",", "\u202f"),
        f"{kpis['echecs_attente_pct']:.1f}\u202f% du total",
        ROUGE,
    )
    c4 = _card(
        "Cr\u00e9neau le plus critique",
        str(kpis["creneau_critique"]),
        f"QoS {kpis['creneau_critique_qos']:.1f}\u202f%",
        "#F59E0B",
    )
    return f'<div class="kpi-grid">{c1}{c2}{c3}{c4}</div>'


def product_table(df: pd.DataFrame) -> str:
    """Tableau HTML de synthèse par produit (QoS colorée)."""
    headers = ["PRODUIT", "TOTAL", "SUCCESSFUL", "FAILED", "PENDING", "INITIATED", "QOS"]
    if df.empty:
        return "<p style='color:#9CA3AF'>Aucune donn\u00e9e produit.</p>"

    ths = "".join(f"<th>{escape(h)}</th>" for h in headers)
    rows = []
    for _, r in df.iterrows():
        qos = float(r["QOS"])
        cls = "qos-ok" if qos >= 50 else "qos-ko"
        tds = (
            f"<td>{escape(str(r['PRODUIT']))}</td>"
            f"<td>{int(r['TOTAL'])}</td>"
            f"<td>{int(r['SUCCESSFUL'])}</td>"
            f"<td>{int(r['FAILED'])}</td>"
            f"<td>{int(r['PENDING'])}</td>"
            f"<td>{int(r['INITIATED'])}</td>"
            f'<td class="{cls}">{qos:.1f}\u202f%</td>'
        )
        rows.append(f"<tr>{tds}</tr>")
    return f'<table class="qos-table"><thead><tr>{ths}</tr></thead><tbody>{"".join(rows)}</tbody></table>'


def error_table(df: pd.DataFrame) -> str:
    """Tableau HTML des messages d'erreur (occurrences en pastille rouge)."""
    headers = ["SERVICE", "MESSAGE", "OCCURRENCES", "% DU TOTAL"]
    if df.empty:
        return "<p style='color:#9CA3AF'>Aucun message d'erreur. \U0001F389</p>"

    ths = "".join(f"<th>{escape(h)}</th>" for h in headers)
    rows = []
    for _, r in df.iterrows():
        tds = (
            f"<td>{escape(str(r['SERVICE']))}</td>"
            f"<td>{escape(str(r['MESSAGE']))}</td>"
            f'<td><span class="pastille">{int(r["OCCURRENCES"])}</span></td>'
            f"<td>{float(r['% DU TOTAL']):.1f}\u202f%</td>"
        )
        rows.append(f"<tr>{tds}</tr>")
    return f'<table class="qos-table"><thead><tr>{ths}</tr></thead><tbody>{"".join(rows)}</tbody></table>'
