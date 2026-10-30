"""Couche visualisation : graphiques Matplotlib sur thème sombre.

Chaque fonction renvoie un objet ``matplotlib.figure.Figure`` que la couche
UI (app.py) affichera via ``st.pyplot``. Aucune dépendance à Streamlit ici.
"""

from __future__ import annotations

import matplotlib

matplotlib.use("Agg")  # backend non interactif, adapté au serveur

import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.lines import Line2D
from matplotlib.patches import Patch
from matplotlib.ticker import MaxNLocator

# Palette du thème sombre
FOND = "#111827"
TEXTE = "#E5E7EB"
GRILLE = "#374151"
BLEU = "#3B82F6"
VERT = "#22C55E"
ROUGE = "#EF4444"


def _apply_dark_theme(fig, ax) -> None:
    """Applique le fond sombre et la couleur du texte à une figure/axe."""
    fig.patch.set_facecolor(FOND)
    ax.set_facecolor(FOND)
    for spine in ax.spines.values():
        spine.set_color(GRILLE)
    ax.tick_params(colors=TEXTE)
    ax.yaxis.label.set_color(TEXTE)
    ax.xaxis.label.set_color(TEXTE)
    ax.title.set_color(TEXTE)


def _labels_heures(heures) -> list[str]:
    return [f"{int(h):02d}h" for h in heures]


def plot_volume_qos(hourly: pd.DataFrame):
    """Graphique 1 : barres bleues (volume) + courbe verte (QoS %).

    Double axe Y, entiers forcés sur l'axe des volumes, légende
    personnalisée placée sous le graphique.
    """
    fig, ax_vol = plt.subplots(figsize=(6.4, 4.6), dpi=120)
    _apply_dark_theme(fig, ax_vol)

    if hourly.empty:
        ax_vol.text(0.5, 0.5, "Aucune donn\u00e9e", ha="center", va="center",
                    color=TEXTE, transform=ax_vol.transAxes)
        return fig

    heures = hourly["heure"].tolist()
    x = range(len(heures))

    ax_vol.bar(x, hourly["volume"], color=BLEU, alpha=0.85, label="Volume")
    ax_vol.set_ylabel("Volume d'op\u00e9rations")
    ax_vol.set_xticks(list(x))
    ax_vol.set_xticklabels(_labels_heures(heures), rotation=45, ha="right")
    ax_vol.yaxis.set_major_locator(MaxNLocator(integer=True))  # entiers forcés
    ax_vol.grid(axis="y", color=GRILLE, linestyle="--", linewidth=0.5, alpha=0.5)

    ax_qos = ax_vol.twinx()
    ax_qos.plot(x, hourly["qos"], color=VERT, marker="o", linewidth=2, label="QoS %")
    ax_qos.set_ylabel("QoS (%)")
    ax_qos.set_ylim(0, 100)
    ax_qos.tick_params(colors=TEXTE)
    ax_qos.yaxis.label.set_color(TEXTE)
    for spine in ax_qos.spines.values():
        spine.set_color(GRILLE)

    ax_vol.set_title("Volume horaire & QoS", pad=12, fontweight="bold")

    legende = [
        Patch(facecolor=BLEU, label="Volume d'op\u00e9rations"),
        Line2D([0], [0], color=VERT, marker="o", label="QoS (%)"),
    ]
    leg = ax_vol.legend(
        handles=legende,
        loc="upper center",
        bbox_to_anchor=(0.5, -0.22),
        ncol=2,
        frameon=False,
    )
    for text in leg.get_texts():
        text.set_color(TEXTE)

    fig.tight_layout()
    return fig


def plot_errors(hourly: pd.DataFrame):
    """Graphique 2 : barres rouges du volume d'erreurs par heure."""
    fig, ax = plt.subplots(figsize=(6.4, 4.6), dpi=120)
    _apply_dark_theme(fig, ax)

    if hourly.empty:
        ax.text(0.5, 0.5, "Aucune donn\u00e9e", ha="center", va="center",
                color=TEXTE, transform=ax.transAxes)
        return fig

    heures = hourly["heure"].tolist()
    x = range(len(heures))

    ax.bar(x, hourly["erreurs"], color=ROUGE, alpha=0.9, label="Erreurs")
    ax.set_ylabel("Volume d'erreurs")
    ax.set_xticks(list(x))
    ax.set_xticklabels(_labels_heures(heures), rotation=45, ha="right")
    ax.yaxis.set_major_locator(MaxNLocator(integer=True))
    ax.grid(axis="y", color=GRILLE, linestyle="--", linewidth=0.5, alpha=0.5)
    ax.set_title("Volume d'erreurs par heure", pad=12, fontweight="bold")

    leg = ax.legend(
        loc="upper center", bbox_to_anchor=(0.5, -0.22), ncol=1, frameon=False
    )
    for text in leg.get_texts():
        text.set_color(TEXTE)

    fig.tight_layout()
    return fig
