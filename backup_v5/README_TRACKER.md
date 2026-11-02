# Tradifoot Delivery Tracker & Intelligence Engine

Bienvenue dans l'outil de suivi automatisé et de veille logistique pour les **Programmes de Livraison Tradifoot** (`Programme de livraison.xlsx` & `Partages`).

---

## 🚀 Utilisation Rapide en 1 Clic

| Fichier | Action |
| :--- | :--- |
| **[`Launch_Dashboard.bat`](file:///c:/Users/ThinkPad%20E14/Desktop/Nero%20AI/OUTLOOK%20(CLASSIC)/Launch_Dashboard.bat)** | Double-cliquez pour **ouvrir instantanément le tableau de bord interactif** dans votre navigateur. |
| **[`Sync_and_View.bat`](file:///c:/Users/ThinkPad%20E14/Desktop/Nero%20AI/OUTLOOK%20(CLASSIC)/Sync_and_View.bat)** | Double-cliquez pour **interroger Outlook Classic, télécharger les nouveaux emails/programmes**, mettre à jour la base et ouvrir le dashboard. |

---

## 📊 Fonctionnalités du Tableau de Bord (`dashboard.html`)

1. **Indicateurs Clés (KPIs en direct) :**
   * Nombre de vagues de livraison tracées.
   * Total des articles expédiés (paires, chaussettes, sacs, etc.).
   * Nombre de boutiques et dépôts desservis.
   * Chauffeurs et véhicules mobilisés.

2. **Graphiques Dynamiques :**
   * Évolution mensuelle des volumes (2025–2026).
   * Part des marques (Skechers, Ecco, Clarks, Ambitious, Flyflot, Hi-Tec, Xtep, etc.).
   * Top 10 des magasins receveurs.
   * Répartition des tournées par chauffeur.

3. **Explorateur & Inspecteur de Programmes :**
   * Recherche en temps réel (par nom de magasin, date, chauffeur, fichier).
   * Filtre par année/période.
   * Détection automatique des **versions modifiées** et des **écarts retirés** (*« A enlever de System »*).
   * **Vue détaillée modale :** Cliquez sur n'importe quel programme pour afficher la tournée arrêt par arrêt (Priorité, Magasin, Wilaya, Quantité, Chauffeur, Véhicule, Retours).

---

## 🛠️ Commandes en Ligne de Commande (Optionnel)

Vous pouvez également lancer le script Python directement depuis PowerShell ou le terminal :

```bash
# Mettre à jour depuis Outlook et ouvrir le dashboard
python tradifoot_tracker.py --sync --open

# Effectuer un scan historique complet de la boîte mail
python tradifoot_tracker.py --full --open

# Afficher les statistiques rapides dans le terminal
python tradifoot_tracker.py --stats
```

---

## 📂 Architecture des Fichiers

* [`tradifoot_tracker.py`](file:///c:/Users/ThinkPad%20E14/Desktop/Nero%20AI/OUTLOOK%20(CLASSIC)/tradifoot_tracker.py) : Moteur d'ingestion Outlook (COM MAPI), parsing Excel et génération du dashboard.
* [`dashboard.html`](file:///c:/Users/ThinkPad%20E14/Desktop/Nero%20AI/OUTLOOK%20(CLASSIC)/dashboard.html) : Interface utilisateur moderne et interactive.
* [`data/delivery_db.json`](file:///c:/Users/ThinkPad%20E14/Desktop/Nero%20AI/OUTLOOK%20(CLASSIC)/data/delivery_db.json) : Base de données JSON normalisée.
* [`data/programs/`](file:///c:/Users/ThinkPad%20E14/Desktop/Nero%20AI/OUTLOOK%20(CLASSIC)/data/programs/) : Répertoire d'archivage des classeurs Excel originaux.
