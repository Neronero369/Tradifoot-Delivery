# 🚚 Tradifoot — Programme de Livraison & Dashboard Calendrier (v2.0 Stable)

Application web interactive de gestion, suivi et synchronisation en temps réel des tournées de livraison de Tradifoot avec Authentification Sécurisée (RBAC), Thème Corporate Tradifoot Navy & Ice-White, et Synchronisation Google Sheets.

---

## 🏛️ Architecture du Projet

Le projet repose sur une architecture sans dépendances locales lourdes, ultra-rapide et optimisée pour mobile :

1. **Frontend (Single-File Application) :**
   - **`index.html`** : Application web complète en fichier unique combinant HTML5 sémantique, Tailwind CSS (via CDN) avec la palette officielle Tradifoot Navy (`#0B1F3F`), Royal Blue (`#2563EB`) et Ice-White (`bg-slate-100`, cartes dépolies `bg-white/95`), et Vanilla JavaScript pur.
   - **Système d'Authentification Sécurisé (Zero-Trust Gate)** :
     - Écran de connexion épuré sans distractions ni texte de placeholder.
     - Validation stricte des comptes pré-autorisés avec correspondance insensible à la casse.
     - Session persistante dans `localStorage` (`auth_session`).
     - Alerte d'erreur dynamique intégrée en cas de saisie incorrecte.
     - Bouton de **Déconnexion** réinitialisant immédiatement la session et la vue.
   - **Base des Comptes Pré-autorisés :**
     - **Admins** (`role: 'admin'`, accès global à l'ensemble des tournées) :
       - `mounir` (PIN: `5290`)
       - `BENAMAR NACIM` (PIN: `3021`)
       - `MEKHLOUFI MOHAMMED` (PIN: `1234`)
     - **Chauffeurs** (`role: 'chauffeur'`, accès exclusif à leurs tournées respectives) :
       - `SAADI BILLEL` (PIN: `9951`)
       - `SNOUCI SEIF-EDDINE` (PIN: `6427`)
       - `BOUKHALFA ADEL` (PIN: `6165`)
       - `BALLA YOUCEF` (PIN: `5763`)
       - `BAHLOULI ADEL` (PIN: `3484`)
   - **Vues Spécifiques par Rôle** :
     - **Chauffeur** :
       - Badge d'en-tête personnalisé : *Chauffeur: [Nom du Chauffeur]*.
       - Affichage strict et exclusif des livraisons assignées au chauffeur connecté (gestion des tournées conjointes et normalisation orthographique).
       - Bouton tactile de validation rapide (*Non Livré* / *Livré*).
       - Synchronisation immédiate vers Google Sheets (`action: 'update_status'`).
     - **Administrateur** :
       - Badge d'en-tête : *Admin Console*.
       - Visualisation globale de toutes les tournées de l'entreprise (870 livraisons).
       - Barre d'outils administrative avec KPI en temps réel et sélecteur de chauffeur.
   - **Contrôle Segmenté à Deux Onglets** :
     - Onglet par défaut : **Non Livré** (permet aux chauffeurs de voir immédiatement les colis en attente).
     - Onglet **Livré** avec compteurs dynamiques en temps réel.
   - **Indicateur de Progression Circulaire (SVG Ring)** :
     - Anneau SVG animé avec dégradé royal blue (`#2563EB`) vers émeraude (`#059669`).
   - **Persistance Locale & Notifications Toast** :
     - Mémorisation instantanée dans `localStorage` (`delivered_ids`).
     - Notifications élégantes style iOS.

2. **Backend Serverless (Google Apps Script Web App) :**
   - **`backend/Code.js`** : Endpoint serverless déployé sur Google Cloud / Google Workspace.
   - Gère le routage par paramètre `action` :
     - `action: 'get_deliveries'` : Récupère les livraisons filtrées selon le rôle (`admin` ou `chauffeur`) et le nom du chauffeur.
     - `action: 'update_status'` : Met à jour le statut dans la feuille Google Sheet avec verrouillage anti-collision (`LockService`).

---

## 📊 Schéma des Données Google Sheet

Les données transmises vers la feuille de calcul Google Sheets suivent le schéma suivant :

| Colonne | En-tête | Type | Description |
| :---: | :--- | :--- | :--- |
| **A** | `ID` | String | Identifiant unique de la tournée (ex: `12`, `145`) |
| **B** | `Date` | String (YYYY-MM-DD) | Date programmée de la livraison (ex: `2026-10-26`) |
| **C** | `Magasin` | String | Destination / Magasin / Dépôt |
| **D** | `Chauffeur` | String | Nom du chauffeur assigné |
| **E** | `Matricule` | String | Véhicule de transport |
| **F** | `Quantite` | Number | Nombre de paires / unités expédiées |
| **G** | `Status` | String | Statut actuel (`Non Livré` ou `Livré`) |
| **H** | `Timestamp` | Timestamp | Horodatage ISO de la synchronisation |

---

## 🔗 Endpoint Actif (Web App URL)

* **Google Apps Script Web App URL :**
  ```text
  https://script.google.com/macros/s/AKfycbwqhYV_Xm_k5qg2K2-QzGZgC7W76i1YyWjV-8o9_BqD7g/exec
  ```
* **Actions supportées (POST JSON) :**
  1. `get_deliveries` :
     ```json
     { "action": "get_deliveries", "role": "chauffeur", "chauffeurName": "SAADI BILLEL" }
     ```
  2. `update_status` :
     ```json
     { "action": "update_status", "id": "42", "status": "Livré", "date": "2026-10-26" }
     ```

---

## 📦 Sauvegardes & Snapshots

Les snapshots stables sont archivés dans le dossier `backups/` :
- `backups/index_v2_auth_stable.html` : Snapshot stable v2.0 avec authentification pré-autorisée et thème Tradifoot Navy.
- `backups/delivery_dashboard_v2_backup.zip` : Archive compressée complète du code source v2.0.
- `index.backup.html` : Miroir de secours direct à la racine.

---

## 🚀 Dépôt GitHub

- **Dépôt officiel :** [https://github.com/Neronero369/Tradifoot-Delivery](https://github.com/Neronero369/Tradifoot-Delivery)
- **Branche principale :** `main`
