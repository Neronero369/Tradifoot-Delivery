# 🚚 Tradifoot — Programme de Livraison & Dashboard Calendrier (RBAC Enabled)

Interface web réactive avec Authentification et Contrôle d'Accès Basé sur les Rôles (RBAC) pour le suivi, la visualisation et la synchronisation en temps réel des tournées de livraison de Tradifoot.

---

## 🏗️ Architecture du Projet

Le projet repose sur une architecture sans dépendances externes lourdes :

1. **Frontend (Single-File Application) :**
   - **`index.html`** : Application web complète en fichier unique combinant HTML5 sémantique, Tailwind CSS (via CDN) pour l'esthétique Apple Dark Mode / Glassmorphism (`#0b0d11`, `#0A84FF`), et Vanilla JavaScript.
   - **Système d'Authentification & Rôles (RBAC)** :
     - Modal d'authentification au design Apple Glassmorphic avec onglets **Se connecter** et **S'inscrire**.
     - Deux rôles : **Administrateur** (`admin`) et **Chauffeur** (`chauffeur`).
     - Session persistante dans `localStorage` (`auth_session`).
     - Comptes de démonstration pré-configurés :
       - Admin : `admin` / `123`
       - Chauffeur : `mounir` / `123` (Nom de chauffeur : `MOUNIR`)
     - Bouton de **Déconnexion** dans l'en-tête de l'application réinitialisant la session.
   - **Vues Spécifiques par Rôle** :
     - **Chauffeur** :
       - Badge d'en-tête personnalisé : *Chauffeur: [Nom du Chauffeur]*.
       - Affichage strict et exclusif des livraisons assignées au chauffeur connecté.
       - Case à cocher tactile permettant de basculer l'état (*Livré* / *Non Livré*).
       - Synchronisation immédiate vers Google Sheets (`action: 'update_status'`).
     - **Administrateur** :
       - Badge d'en-tête : *Admin Console*.
       - Visualisation globale de toutes les livraisons de l'entreprise.
       - Barre d'outils administrative avec indicateurs KPI cumulatifs (Total expédié, Nombre de colis livrés, Taux d'achèvement global).
       - Filtrage interactif par nom de chauffeur ou par magasin / destination.
   - **Persistance Locale & Notifications Toast** :
     - Mémorisation instantanée dans `localStorage` (`delivered_ids`).
     - Notifications élégantes style iOS (succès, avertissement, synchronisation hors ligne).

2. **Backend Serverless (Google Apps Script Web App) :**
   - **`backend/Code.js`** : Endpoint serverless déployé sur Google Cloud / Google Workspace.
   - Gère le routage par paramètre `action` :
     - `action: 'get_deliveries'` : Récupère les livraisons filtrées selon le rôle (`admin` ou `chauffeur`) et le nom du chauffeur.
     - `action: 'update_status'` : Met à jour le statut dans la feuille Google Sheet avec verrouillage anti-collision (`LockService`).
   - Requêtes `GET` de diagnostic et de vérification de schéma.

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
* **Actions supportées (POST JSON) :**
  1. `get_deliveries` :
     ```json
     { "action": "get_deliveries", "role": "chauffeur", "chauffeurName": "MOUNIR" }
     ```
  2. `update_status` :
     ```json
     { "action": "update_status", "id": "DEL-0042", "status": "Livré", "date": "2026-10-26" }
     ```

---

## 🚀 Utilisation & Déploiement

* **En local :** Ouvrez simplement `index.html` dans n'importe quel navigateur (Safari, Chrome, Edge, Firefox).
* **Comptes de test immédiats :**
  - **Administrateur** : `admin` / `123`
  - **Chauffeur** : `mounir` / `123`
* **Dépôt GitHub :**
  - Dépôt officiel : [https://github.com/Neronero369/Tradifoot-Delivery](https://github.com/Neronero369/Tradifoot-Delivery)
  - Branche : `main`
