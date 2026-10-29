# 🚚 Tradifoot — Programme de Livraison & Dashboard Calendrier (v2.0 Stable)

Application web interactive de gestion, suivi et synchronisation en temps réel des tournées de livraison de Tradifoot avec Authentification Sécurisée (RBAC), Thème Corporate Tradifoot Navy & Ice-White, et Synchronisation Supabase en temps réel.

---

## 🏛️ Architecture du Projet

Le projet repose sur une architecture moderne sans dépendances locales lourdes, ultra-rapide et optimisée pour mobile et desktop :

1. **Frontend (Single-File Application) :**
   - **`index.html`** : Application web complète en fichier unique combinant HTML5 sémantique, Tailwind CSS (via CDN) avec la palette officielle Tradifoot Navy (`#0B1F3F`), Royal Blue (`#2563EB`) et Ice-White (`bg-slate-100`, cartes dépolies `bg-white/95`), et Vanilla JavaScript pur.
   - **Système d'Authentification Sécurisé (Zero-Trust Gate)** :
     - Écran de connexion épuré sans distractions.
     - Validation stricte des comptes pré-autorisés avec correspondance insensible à la casse.
     - Session persistante dans `localStorage` (`auth_session`).
     - Alerte d'erreur dynamique intégrée en cas de saisie incorrecte.
     - Bouton de **Déconnexion** réinitialisant immédiatement la session, la vue et coupant les canaux WebSocket.
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
       - Badge d'en-tête personnalisé : *Espace Chauffeur • [Nom du Chauffeur]*.
       - Affichage strict et exclusif des livraisons assignées au chauffeur connecté.
       - Bouton tactile de validation rapide (*Non Livré* / *Livré*).
       - Synchronisation immédiate vers Supabase (`syncDeliveryStatusToSupabase`).
     - **Administrateur** :
       - Badge d'en-tête : *ADMIN CONSOLE*.
       - Visualisation globale de toutes les tournées de l'entreprise.
       - Mode consultation et supervision en lecture seule.
       - Barre d'outils administrative avec KPIs en temps réel et sélecteur de chauffeur.
   - **Contrôle Segmenté à Deux Onglets** :
     - Onglet par défaut : **Non Livré** (permet aux chauffeurs de voir immédiatement les colis en attente).
     - Onglet **Livré** avec compteurs dynamiques en temps réel.
   - **Indicateur de Progression Circulaire (SVG Ring)** :
     - Anneau SVG animé avec dégradé royal blue (`#2563EB`) vers émeraude (`#059669`).
   - **Persistance Locale & Notifications Toast** :
     - Mémorisation instantanée dans `localStorage` (`delivered_delivery_ids`).
     - Notifications élégantes style iOS.

2. **Backend Base de Données (Supabase Cloud PostgreSQL) :**
   - **PostgreSQL Database** hébergée sur Supabase Cloud.
   - **Connexion Directe SDK Client** (`@supabase/supabase-js` v2).
   - **WebSocket Real-Time Engine** (`supabaseClient.channel('public:deliveries')`) :
     - Écoute en continu les mutations `INSERT`, `UPDATE` et `DELETE` sur la table `deliveries`.
     - Répercute instantanément les changements validés sur smartphone vers les écrans de supervision PC sans rafraîchissement manuel.
   - **Résilience Réseau & File d'Attente Hors-Ligne** :
     - Stockage local des mutations en attente (`pending_supabase_sync_queue`).
     - Flush automatique dès le rétablissement de la connexion internet.

---

## 📊 Schéma des Données Supabase (`deliveries`)

Les données de livraison dans Supabase suivent le schéma relationnel suivant :

| Colonne | Type | Description |
| :--- | :--- | :--- |
| `id` | `bigint` (Identity) | Identifiant interne séquentiel de la table |
| `delivery_id` | `text` (Unique) | Identifiant unique de la tournée (ex: `DEL-0001`, `1001`) |
| `date` | `date` (YYYY-MM-DD) | Date programmée de la livraison (ex: `2026-10-28`) |
| `magasin` | `text` | Destination / Magasin / Dépôt |
| `chauffeur` | `text` | Nom du chauffeur assigné |
| `matricule` | `text` | Véhicule de transport |
| `quantite` | `integer` | Nombre de paires / unités expédiées |
| `status` | `text` | Statut actuel (`Non Livré` ou `Livré`) |
| `updated_at` | `timestamptz` | Horodatage ISO de la dernière mise à jour |

---

## 🔗 Configuration Supabase

* **Supabase Project URL :**
  ```text
  https://gmdgnzokgqyivlfeosdy.supabase.co
  ```
* **Supabase Anon / Publishable Key :**
  ```text
  sb_publishable_JPqzQ4b11U9BU8HzG16k2g_O2cHl5ur
  ```
* **REST API Endpoint :**
  ```text
  https://gmdgnzokgqyivlfeosdy.supabase.co/rest/v1/deliveries
  ```
* **Canal Realtime WebSocket :**
  ```text
  public:deliveries
  ```

---

## 🐍 Ingestion Python Automatisée (`parse_programme.py`)

Le script Python permet de valider et d'ingérer directement les programmes de livraison Excel dans Supabase :

```bash
# Ingestion directe dans Supabase
python parse_programme.py
```

Fonction principale :
```python
from parse_programme import push_manifest_to_supabase

# Ingestion d'un fichier Excel de livraison
result = push_manifest_to_supabase("template_programme_de_livraison.xlsx")
print(result)
```
