# 🚚 Tradifoot — Programme de Livraison & Dashboard Calendrier

Interface moderne et réactive pour le suivi, la visualisation et la synchronisation en temps réel des tournées de livraison de Tradifoot.

---

## 🏗️ Architecture du Projet

Le projet est conçu selon une architecture légère, résiliente et sans dépendances lourdes :

1. **Frontend (Single-File Application) :**
   - **`index.html`** : Application web complète en fichier unique combinant structure HTML5 sémantique, Tailwind CSS (via CDN) pour l'esthétique Apple Dark Mode / Glassmorphism, et Vanilla JavaScript.
   - **LocalStorage Persistence** : Mémorisation locale et instantanée de l'état des livraisons cochées (`completedDeliveryIds` et dates de validation) persistante entre les rafraîchissements de page.
   - **Gamification & UX** : Anneau de progression SVG animé (iOS Blue `#0A84FF`), boutons à bascule tactiles, filtrage dynamique (*Tous*, *Livré*, *Non Livré*), et grille calendaire mensuelle adaptative (375px mobile jusqu'au grand écran).

2. **Backend Serverless (Google Apps Script Web App) :**
   - **`backend/Code.js`** : Endpoint serverless déployé sur Google Cloud / Google Workspace recevant les requêtes de mutation d'état en mode `write-only` (`mode: 'no-cors'`).
   - Verrou de concurrence (`LockService`) pour garantir l'intégrité des écritures lors des mises à jour rapides.

3. **Automatisation Locale (Python & Outlook Classic) :**
   - **`tradifoot_tracker.py`** & **`run_daily_routine.py`** : Moteur d'ingestion automatique connecté à Microsoft Outlook Classic (MAPI `win32com.client`) vérifiant chaque soir à 19h00 les nouveaux programmes de livraison Excel pour actualiser la base locale et l'application.

---

## 📊 Schéma des Données Google Sheet

Les données transmises vers la feuille de calcul Google Sheets suivent le schéma suivant :

| Colonne | En-tête | Type | Description |
| :---: | :--- | :--- | :--- |
| **A** | `ID` | String | Identifiant unique de la tournée (ex: `DEL-0001`) |
| **B** | `Date` | String (YYYY-MM-DD) | Date programmée de la livraison (ex: `2026-10-25`) |
| **C** | `Magasin` | String | Destination / Magasin / Dépôt (ex: `SKECHERS BAB EZZOUAR`) |
| **D** | `Chauffeur` | String | Nom du chauffeur assigné (ex: `MOUNIR`, `SAADI BILLEL`) |
| **E** | `Matricule` | String | Véhicule de transport (ex: `DOBLO`, `MERCEDES`) |
| **F** | `Quantite` | Number | Nombre de paires / unités expédiées |
| **G** | `Status` | String | Statut actuel (`Non Livré` ou `Livré`) |
| **H** | `Date Mise à Jour` | Timestamp | Horodatage de la validation de livraison |

---

## 🌐 Endpoint Actif (Web App URL)

* **Google Apps Script Web App URL :**
  ```text
  https://script.google.com/macros/s/AKfycbwM5y2SPfnT1lMg4_ingmt_ZWzfHFb7l-qTwTow3-eThQ52Rf1FfLRrcBertuvxUY6d/exec
  ```
* **Méthode supportée :** `POST` (Payload JSON) et `GET` (Health-check)
* **Mode de transmission client :** `fetch(GOOGLE_APP_URL, { method: 'POST', mode: 'no-cors', body: JSON.stringify(...) })`

---

## 🚀 Utilisation & Déploiement

* **En local :** Ouvrez simplement `index.html` dans n'importe quel navigateur (Edge, Chrome, Safari, Firefox).
* **Déploiement GitHub Pages :**
  * Dépôt officiel : [https://github.com/Neronero369/Delivery-Moh](https://github.com/Neronero369/Delivery-Moh)
  * Branche : `main` / Répertoire racine `/(root)`
