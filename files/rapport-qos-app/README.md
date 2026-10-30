# 📶 Rapport QoS Télécom — Application Streamlit

Application web interne permettant à l'équipe de monitoring télécom de générer
des **rapports de Qualité de Service (QoS)** à partir d'un export Excel.

Le code est **modulaire** (interface / traitement des données / visualisations)
et le détail d'implémentation reste invisible pour l'utilisateur final : seule
l'interface Streamlit est exposée.

## 🗂️ Structure du projet

```
rapport-qos-app/
├── app.py                  # Point d'entrée Streamlit (UI uniquement)
├── requirements.txt        # Dépendances
├── README.md               # Ce fichier
├── .gitignore              # Fichiers ignorés par Git
└── src/
    ├── __init__.py
    ├── data_processing.py  # Nettoyage, filtrage et calculs Pandas
    ├── visualizations.py   # Graphiques Matplotlib (thème sombre)
    └── ui_components.py    # Blocs HTML (cartes KPI, tableaux)
```

## ⚙️ Colonnes attendues dans le fichier Excel

Les noms de colonnes sont centralisés en haut de `src/data_processing.py`.
Adaptez-les si votre export diffère :

| Constante         | Nom de colonne par défaut   | Rôle                                    |
|-------------------|-----------------------------|-----------------------------------------|
| `COL_PAYS`        | `CODE PAYS ISO`             | Code pays (SN, CI, BF…)                 |
| `COL_FILIALE`     | `FILIALE`                   | Filiale (menu déroulant)                |
| `COL_DATE`        | `DATE TRANSACTION`          | Horodatage de la transaction            |
| `COL_TRANSACTION` | `NUMÉRO TRANSACTION GU`     | Numéro de transaction (apostrophes ôtées) |
| `COL_STATUT`      | `STATUT`                    | SUCCESSFUL / FAILED / PENDING / INITIATED |
| `COL_PRODUIT`     | `PRODUIT`                   | Produit                                 |
| `COL_SERVICE`     | `SERVICE`                   | Service concerné (messages d'erreur)    |
| `COL_MESSAGE`     | `MESSAGE`                   | Libellé du message d'erreur             |

## 📊 Ce que fait l'application

1. **Import & préparation**
   - Upload d'un fichier `.xlsx` (traité en mémoire via `io.BytesIO`, aucune écriture disque).
   - Suppression des transactions de **l'heure en cours**.
   - Nettoyage de la colonne `NUMÉRO TRANSACTION GU` (retrait des apostrophes).
   - Filtre par filiale (ou « TOUTES LES FILIALES »).
2. **4 cartes KPI** : Opérations, QoS globale, Échecs/attente, Créneau le plus critique.
3. **2 graphiques Matplotlib** (thème sombre, côte à côte) :
   - Volume (barres bleues) + QoS % (courbe verte) sur double axe Y.
   - Volume d'erreurs par heure (barres rouges).
4. **2 tableaux HTML** : synthèse par produit (QoS colorée) et messages d'erreur (pastilles rouges).
5. **Export Excel conditionnel** : si plus de 8 messages d'erreur, un bouton de téléchargement apparaît.

## 🚀 Installation locale

```bash
git clone <URL_DE_VOTRE_REPO>
cd rapport-qos-app

python -m venv .venv
source .venv/bin/activate      # Windows : .venv\Scripts\activate

pip install -r requirements.txt
streamlit run app.py
```

L'application s'ouvre sur <http://localhost:8501>.

## ☁️ Déploiement sur Streamlit Community Cloud

1. Poussez ce repository sur **GitHub** (dépôt public ou privé).
2. Connectez-vous sur <https://share.streamlit.io> avec votre compte GitHub.
3. Cliquez sur **« New app »**, sélectionnez le repo, la branche et le fichier
   principal `app.py`.
4. Cliquez sur **« Deploy »**. Les dépendances de `requirements.txt` sont
   installées automatiquement.
5. À chaque `git push`, l'application se redéploie toute seule.

> 🔒 **Confidentialité** : ne committez jamais de données QoS réelles. Le
> `.gitignore` exclut déjà les fichiers `*.xlsx` / `*.csv`. Pour un accès
> restreint, utilisez le paramètre de visibilité privée de Streamlit Cloud.

## 🧱 Maintenance / évolutions

- **Changer les noms de colonnes** → `src/data_processing.py` (constantes en tête).
- **Modifier le calcul de QoS** → `compute_kpis`, `product_synthesis`, `hourly_stats`.
- **Retoucher les graphiques** → `src/visualizations.py` (palette en tête).
- **Changer le design des cartes/tableaux** → `src/ui_components.py` (`global_css`).
- `app.py` reste volontairement minimal : il n'orchestre que l'UI.
