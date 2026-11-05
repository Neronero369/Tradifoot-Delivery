# Supabase Data Architecture & Dictionary
**Tradifoot Delivery & Fleet Management System**  
*Document Version: 1.0.0 — Production Reference*  
*Last Updated: 2026-11-05*

---

## 1. Executive Architecture Overview

The Tradifoot application relies on **Supabase (PostgreSQL + PostgREST + Storage)** as its cloud persistence layer. The data architecture unites two interdependent operational domains:

1. **Logistics & Delivery Operations (`deliveries`)**: Ingestion of daily Excel delivery schedules (*Programmes de Livraison*), driver tour assignment, real-time Proof of Delivery (PoD) capture, GPS coordinates, cash collection (*recouvrement*), and merchandise dispute/reserve tracking.
2. **Fleet Management (*Parc Auto*) (`vehicles` & `vehicle_handovers`)**: Fleet lifecycle tracking for 4 heavy/commercial trucks (Mercedes, Hyundai, Nissan), multi-point vehicle handover inspections (odometer, fuel, 5-point photo inspection), session-aware driver custody locks, and administrative maintenance quarantine.
3. **Blob Storage (`delivery-proofs` Storage Bucket)**: Immutable storage of photographic evidence (stamped waybills, cash receipts, vehicle inspection photos).

```
 +-----------------------------------------------------------------------------------+
 |                             Tradifoot Architecture                                 |
 +-----------------------------------------------------------------------------------+
                                          |
         +--------------------------------+--------------------------------+
         |                                                                 |
         v                                                                 v
+-----------------------+                                        +-------------------+
|  Excel / Outlook Sync |                                        | Driver & ADV Web  |
|  (parse_programme.py) |                                        |    (index.html)   |
+-----------------------+                                        +-------------------+
         |                                                                 |
         | (Insert-Only Sync)                                              | (Real-time updates)
         v                                                                 v
+------------------------------------------------------------------------------------+
|                               Supabase PostgreSQL Layer                            |
|                                                                                    |
|  +--------------------+        +--------------------+        +-------------------+ |
|  |     deliveries     |        |      vehicles      | 1    N | vehicle_handovers | |
|  |--------------------|        |--------------------|--------|-------------------| |
|  | PK: delivery_key   |        | PK: id (UUID)      |        | PK: id (TEXT/UUID)| |
|  | UK: delivery_id    |        | UK: plate_number   |        | FK: vehicle_id    | |
|  +--------------------+        +--------------------+        +-------------------+ |
+------------------------------------------------------------------------------------+
                                          |
                                          v
                         +---------------------------------+
                         |   Supabase Storage Bucket       |
                         |   ("delivery-proofs")           |
                         +---------------------------------+
```

---

## 2. Table Catalog & Column-Level Mapping

### 2.1 Table: `deliveries`
* **Business Description**: Stores daily store delivery drops, scheduled quantities, driver assignments, proof-of-delivery validation (dual photos, stamps), cash/cheque collection amounts, discrepancy dispute tracking, and GPS geocoding.
* **Sync Strategy**: **Insert-Only**. Automated scripts (`parse_programme.py`) query existing `delivery_id` / `delivery_key` hashes and skip existing rows to prevent wiping driver-submitted operational data.

| Column Name | Data Type | Constraints & Defaults | Source / Trigger / Ingestion Path | Description & Business Rules |
| :--- | :--- | :--- | :--- | :--- |
| `delivery_key` | `TEXT` | `PRIMARY KEY`, `NOT NULL` | Calculated automatically: `YYYY-MM-DD\|MAGASIN\|QUANTITE[\|OCCn]` | Unique deterministic hash used as the deduplication key (`onConflict: 'delivery_key'`). |
| `delivery_id` | `TEXT` | `UNIQUE`, `NOT NULL` | Ingested from Excel / slug generator (`DEL-YYYYMMDD-XXXX`) | Human-readable business identifier for manifest lookup and barcode matching. |
| `date` | `DATE` / `TEXT` | `NOT NULL` | Ingested from Excel column `Date` / `Programme` | Scheduled delivery date formatted as `YYYY-MM-DD`. Used for daily calendar partitioning. |
| `magasin` | `TEXT` | `NOT NULL` | Ingested from Excel column `Boutique` / `Magasin` | Uppercased store destination name (e.g., `"BAB EZZOUAR"`, `"ORAN SENIA"`). |
| `chauffeur` | `TEXT` | Nullable (`DEFAULT NULL`) | Ingested from Excel column `Chauffeur` | Assigned chauffeur name (e.g., `"SAADI BILLEL"`). Special routing: `"RETRAIT DÉPÔT"` for customer depot pickups. |
| `delivered_by` | `TEXT` | Nullable (`DEFAULT NULL`) | Updated by driver/ADV UI upon delivery | Actual authenticated user (`currentUser.chauffeurName`) who finalized and stamped the delivery. |
| `matricule` | `TEXT` | Nullable (`DEFAULT NULL`) | Ingested from Excel column `Matricule` / `Véhicule` | Assigned truck plate identifier or equipment type (e.g., `"MERCEDES-01"`, `"Camion Porteur"`). |
| `materiel` | `TEXT` | Nullable (`DEFAULT NULL`) | Ingested from Excel column `Matériel` / `Equipement` | Physical equipment code; filtered to `NULL` if blank or contains `nan`, `none`, `-`. |
| `quantite` | `INTEGER` | `NOT NULL`, `DEFAULT 0` | Ingested from Excel column `Globale` / `Quantité` | Total scheduled units (pairs of shoes/cartons) expected for this destination drop. |
| `quantite_recue` | `INTEGER` | Nullable (`DEFAULT NULL`) | Entered by driver in PoD modal UI | Actual units accepted by the store manager upon physical reception. |
| `quantite_litige` | `INTEGER` | Nullable (`DEFAULT 0`) | Calculated automatically in UI | Absolute quantity discrepancy: `abs(quantite - quantite_recue)`. |
| `status` | `TEXT` | `NOT NULL`, `DEFAULT 'Non Livré'` | Ingested as `'Non Livré'`, updated via UI PoD modal | Current operational lifecycle stage: `'Non Livré'`, `'Livré'`. |
| `statut_reserve` | `TEXT` | Nullable (`DEFAULT NULL`) | Updated by driver in PoD modal | Conformity flag: `'Conforme'` (if quantities match) or `'Avec Réserve'` (if discrepancy/damage). |
| `motif_reserve` | `TEXT` | Nullable (`DEFAULT NULL`) | Selected by driver via quick chips in PoD modal | Categorized incident reason: `"Colis Manquant"`, `"Colis Endommagé"`, `"Refus Boutique"`, etc. |
| `remarque_chauffeur`| `TEXT` | Nullable (`DEFAULT NULL`) | Entered by driver in PoD modal | Free-form explanation detailing merchandise discrepancies, store manager feedback, or damages. |
| `photo_url` | `TEXT` | Nullable (`DEFAULT NULL`) | Uploaded by driver via PoD camera/file input | Public HTTPS URL to the signed and stamped delivery slip (Bon de Livraison) in Storage. |
| `recouvrement_url` | `TEXT` | Nullable (`DEFAULT NULL`) | Uploaded by driver via PoD camera/file input | Public HTTPS URL to the cash receipt / cheque voucher (*Preuve de Recette*) in Storage. |
| `a_recouvrement` | `BOOLEAN` | `DEFAULT FALSE` | Calculated from driver inputs in PoD modal | Set to `TRUE` if cash collection occurred (`montant_recouvrement > 0` or receipt photo uploaded). |
| `montant_recouvrement`| `NUMERIC` | `DEFAULT 0` | Entered by driver in PoD modal | Cash or cheque collection amount collected by the driver (in Algerian Dinars DZD). |
| `latitude` | `DOUBLE PRECISION`| Nullable (`DEFAULT NULL`) | Browser Geolocation API (`getCurrentPosition`) | GPS Latitude captured at the exact moment the driver taps "Valider la Livraison". |
| `longitude` | `DOUBLE PRECISION`| Nullable (`DEFAULT NULL`) | Browser Geolocation API (`getCurrentPosition`) | GPS Longitude captured at the exact moment the driver taps "Valider la Livraison". |
| `updated_at` | `TIMESTAMPTZ` | `DEFAULT NOW()` | Supabase DB Trigger or UI payload | ISO timestamp recording the last modification time. |

---

### 2.2 Table: `vehicles`
* **Business Description**: Represents the fleet of 4 delivery trucks managed in the "Parc Auto" dashboard. Tracks operational availability, active chauffeur custody locks, latest odometer reading, and fuel level.
* **Sync Strategy**: Persistent master catalog. Mutated dynamically during driver check-in/out and maintenance toggles.

| Column Name | Data Type | Constraints & Defaults | Source / Trigger / Ingestion Path | Description & Business Rules |
| :--- | :--- | :--- | :--- | :--- |
| `id` | `UUID` | `PRIMARY KEY`, `DEFAULT gen_random_uuid()` | Database generated or seeded | Unique immutable UUID assigned to each commercial vehicle. |
| `brand` | `TEXT` | `NOT NULL` | Master catalog seed / Admin setup | Vehicle make/brand: `'MERCEDES'`, `'HYUNDAI'`, `'NISSAN'`. |
| `plate_number` | `TEXT` | `UNIQUE`, `NOT NULL` | Master catalog seed / Admin setup | Vehicle license plate / internal fleet identifier: `'MERCEDES-01'`, `'MERCEDES-02'`, `'HYUNDAI-01'`, `'NISSAN-01'`. |
| `status` | `TEXT` | `NOT NULL`, `DEFAULT 'Disponible'` | Mutated by driver actions or Admin toggle | Operational status: `'Disponible'`, `'En Tournée'`, `'Maintenance'`. |
| `current_mileage` | `INTEGER` | `NOT NULL`, `DEFAULT 0` | Updated by driver during handover inspection | Current verified odometer reading in kilometers (km). Updated on every check-in/out. |
| `fuel_level` | `TEXT` | `DEFAULT '100%'` | Updated by driver during handover inspection | Current tank level: `'25%'`, `'50%'`, `'75%'`, `'100%'`. |
| `current_driver` | `TEXT` | Nullable (`DEFAULT NULL`) | Set on `'Prise'`, cleared on `'Retour'`/`Maintenance` | Chauffeur currently in possession of the truck. Enforces strict session-aware vehicle lock. |
| `created_at` | `TIMESTAMPTZ` | `DEFAULT NOW()` | Database auto-timestamp | Vehicle registration timestamp. |

---

### 2.3 Table: `vehicle_handovers`
* **Business Description**: Immutable chronological audit log recording every vehicle check-in (*Prise*) and check-out (*Retour*). Stores driver identity, odometer reading, fuel level, optional remarks, and an array of 5 inspection photos.
* **Sync Strategy**: **Append-Only / Insert-Only**. Records are never updated or deleted once committed.

| Column Name | Data Type | Constraints & Defaults | Source / Trigger / Ingestion Path | Description & Business Rules |
| :--- | :--- | :--- | :--- | :--- |
| `id` | `TEXT` / `UUID` | `PRIMARY KEY` | Client generated (`ho_{timestamp}`) or `UUID` | Unique handover inspection audit event identifier. |
| `vehicle_id` | `UUID` | `NOT NULL`, `FOREIGN KEY REFERENCES vehicles(id)` | Captured from active vehicle card | Relational link pointing to the truck being inspected. |
| `driver_name` | `TEXT` | `NOT NULL` | Session variable (`localStorage.getItem('chauffeur')`) | Full name of the authenticated driver taking or returning the vehicle. |
| `action_type` | `TEXT` | `NOT NULL` | Modal handover context | Type of handover event: `'Prise'` (vehicle departure) or `'Retour'` (vehicle restitution). |
| `mileage` | `INTEGER` | `NOT NULL` | Input by driver in handover modal | Odometer reading recorded at the time of inspection (validates truck mileage advancement). |
| `fuel_level` | `TEXT` | `NOT NULL`, `DEFAULT '100%'` | Dropdown selection in handover modal | Fuel level recorded at the time of inspection (`'25%'`, `'50%'`, `'75%'`, `'100%'`). |
| `remarque` | `TEXT` | Nullable (`DEFAULT ''`) | Text input by driver in handover modal | Driver observations regarding vehicle bodywork, cleanliness, or mechanical noises. |
| `photo_urls` | `TEXT[]` / `JSONB` | Nullable (`DEFAULT '[]'`) | Uploaded via handover modal 5-point photo slots | Array of public URLs for the 5 inspection photos stored in bucket `delivery-proofs`. |
| `created_at` | `TIMESTAMPTZ` | `DEFAULT NOW()`, `NOT NULL` | Database / client ISO timestamp | Timestamp marking when the handover inspection was finalized. |

---

## 3. Storage Layer: Bucket `delivery-proofs`

The Supabase Storage bucket `delivery-proofs` stores compressed JPEG images uploaded from mobile devices.

```
delivery-proofs/
│
├── proofs/
│   ├── DEL-20261105-0001_1762345678.jpg       <-- Primary Proof of Delivery (Stamped BL)
│   └── REC-20261105-0001_1762345679.jpg       <-- Cash Collection Receipt (Recette)
│
└── handovers/
    ├── mercedes/
    │   ├── 1762345680000_slot1.jpg            <-- Slot 1: Face Avant (Front)
    │   ├── 1762345680000_slot2.jpg            <-- Slot 2: Côté Droit (Right Side)
    │   ├── 1762345680000_slot3.jpg            <-- Slot 3: Face Arrière (Rear)
    │   ├── 1762345680000_slot4.jpg            <-- Slot 4: Côté Gauche (Left Side)
    │   └── 1762345680000_slot5.jpg            <-- Slot 5: Tableau de Bord / Intérieur (Cabin/Dash)
    └── hyundai/
        └── ...
```

* **Client Compression**: Images are resized to a maximum width of `1280px` at `0.7` JPEG quality before transmission to conserve bandwidth and storage.
* **Access Policy**: Public read enabled; authenticated upload.

---

## 4. Entity Relationships & Foreign Keys

```
+---------------------------------------+
|              vehicles                 |
+---------------------------------------+
| PK  id               UUID             |<-------------+
|     brand            TEXT             |              |
|     plate_number     TEXT             |              | 1:N
|     status           TEXT             |              | (ON DELETE CASCADE)
|     current_mileage  INTEGER          |              |
|     fuel_level       TEXT             |              |
|     current_driver   TEXT             |              |
+---------------------------------------+              |
                                                       |
+---------------------------------------+              |
|          vehicle_handovers            |              |
+---------------------------------------+              |
| PK  id               TEXT / UUID      |              |
| FK  vehicle_id       UUID             |--------------+
|     driver_name      TEXT             |
|     action_type      TEXT             |
|     mileage          INTEGER          |
|     fuel_level       TEXT             |
|     remarque         TEXT             |
|     photo_urls       JSONB / TEXT[]   |
|     created_at       TIMESTAMPTZ      |
+---------------------------------------+
```

### Relational Dynamics:
1. **`vehicle_handovers.vehicle_id` -> `vehicles.id`**:
   - Foreign key relationship linking inspection logs to physical vehicles.
   - Used in the UI modal `"Voir Historique"`:
     ```javascript
     supabaseClient
       .from('vehicle_handovers')
       .select('*')
       .eq('vehicle_id', vehicleId)
       .order('created_at', { ascending: false });
     ```
2. **`deliveries.matricule` <-> `vehicles.plate_number`**:
   - Logical operational connection between planned delivery manifests and physical fleet vehicles.
3. **`deliveries.chauffeur` & `vehicle_handovers.driver_name`**:
   - Master list of authorized drivers:
     * `SAADI BILLEL`
     * `SNOUCI SEIF-EDDINE`
     * `BOUKHALFA ADEL`
     * `BALLA YOUCEF`
     * Special administrative route: `RETRAIT DÉPÔT` (handled by ADV).

---

## 5. State Machine & Business Logic Flags

### 5.1 Delivery Status Lifecycle (`deliveries.status`)

```
 [ Excel Schedule Ingested ]
              |
              v
       ( Non Livré ) <-----------------------------+
              |                                    |
              | Driver validates PoD in UI         | Reset / Admin Intervention
              v                                    |
          ( Livré ) -------------------------------+
              |
              +---> [ If quantite_recue != quantite ] ---> Statut Réserve: "Avec Réserve"
              +---> [ If quantite_recue == quantite ] ---> Statut Réserve: "Conforme"
```

* **`'Non Livré'`** *(Default)*:
  - Delivery is pending, assigned to a driver route, and visible in the active manifest.
  - Driver action button: `"Valider la Livraison"`.
* **`'Livré'`**:
  - Delivery has been completed by the driver.
  - Requires either a primary photo upload or explicit dual-photo confirmation.
  - UI shows green badge: `"Livré ✓"` (or amber badge `"Livré avec réserve ⚠️"`).
* **`'Conforme'` vs `'Avec Réserve'` (`statut_reserve`)**:
  - **Conforme**: Received quantity strictly equals planned quantity (`quantite_recue == quantite`).
  - **Avec Réserve**: Quantities mismatch or goods were damaged. Requires selecting a `motif_reserve` chip and recording driver remarks.

---

### 5.2 Vehicle Fleet State Machine (`vehicles.status`)

```
                   +----------------------------------+
                   |                                  |
                   v                                  |
            [ Disponible ]                            |
             (Green Badge)                            |
              /          \                            |
   Driver    /            \   Admin Sets              | Admin Sets
   Taps     /              \  Maintenance             | Available
  "Prendre"/                \                         |
          v                  v                        |
   [ En Tournée ]       [ Maintenance ] --------------+
    (Blue Badge)         (Red Badge)
          |                   |
          | Driver Taps       | Hard Block: Chauffeurs
          | "Restituer"       | cannot take or assign truck
          v                   |
   (Returns to Disponible) ---+
```

| Vehicle Status | Chauffeur UI Action | Admin / ADV UI Action | Custody Lock Rule |
| :--- | :--- | :--- | :--- |
| **`'Disponible'`** | Button: `"Prendre ce véhicule"` (Green). Opens departure inspection modal (`'Prise'`). | Button: `"Mettre en Maintenance"` (Red). | `current_driver = NULL`. Any authorized chauffeur may claim the truck. |
| **`'En Tournée'`** | **If current user == `current_driver`**: Button `"Restituer le véhicule"` (Blue).<br>**If current user != `current_driver`**: Button hidden; displays locked badge: `🔒 En utilisation par [Driver]`. | Button: `Bloqué (En Tournée)` (`disabled`, greyed out). Administrative maintenance is strictly blocked while a mission is active. | `current_driver` holds exclusive custody. Other drivers are locked out. |
| **`'Maintenance'`**| All action buttons completely hidden. Displays: `🚫 Véhicule Indisponible`. | Button: `"Remettre Disponible"` (Emerald). Restores truck to active service. | `current_driver = NULL`. Truck is physically in workshop/garage. |

---

### 5.3 Handover Action Types (`vehicle_handovers.action_type`)
* **`'Prise'` (Check-out)**:
  - Executed when a chauffeur takes custody of a vehicle before starting their delivery tour.
  - Updates `vehicles`:
    ```json
    {
      "status": "En Tournée",
      "current_driver": "NOM DU CHAUFFEUR",
      "current_mileage": 12450,
      "fuel_level": "75%"
    }
    ```
* **`'Retour'` (Check-in / Restitution)**:
  - Executed when the chauffeur returns the truck to the depot at the end of the day.
  - Updates `vehicles`:
    ```json
    {
      "status": "Disponible",
      "current_driver": null,
      "current_mileage": 12620,
      "fuel_level": "50%"
    }
    ```

---

## 6. SQL DDL Reference Script (Production Initialization)

For reference or disaster recovery, the tables documented above correspond to the following PostgreSQL schema definitions:

```sql
-- ========================================================
-- 1. DELIVERIES TABLE
-- ========================================================
CREATE TABLE IF NOT EXISTS public.deliveries (
    delivery_key TEXT PRIMARY KEY,
    delivery_id TEXT NOT NULL UNIQUE,
    date DATE NOT NULL,
    magasin TEXT NOT NULL,
    chauffeur TEXT,
    delivered_by TEXT,
    matricule TEXT,
    materiel TEXT,
    quantite INTEGER NOT NULL DEFAULT 0,
    quantite_recue INTEGER,
    quantite_litige INTEGER DEFAULT 0,
    status TEXT NOT NULL DEFAULT 'Non Livré',
    statut_reserve TEXT,
    motif_reserve TEXT,
    remarque_chauffeur TEXT,
    photo_url TEXT,
    recouvrement_url TEXT,
    a_recouvrement BOOLEAN DEFAULT FALSE,
    montant_recouvrement NUMERIC(12, 2) DEFAULT 0,
    latitude DOUBLE PRECISION,
    longitude DOUBLE PRECISION,
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_deliveries_date ON public.deliveries(date);
CREATE INDEX IF NOT EXISTS idx_deliveries_chauffeur ON public.deliveries(chauffeur);
CREATE INDEX IF NOT EXISTS idx_deliveries_status ON public.deliveries(status);

-- ========================================================
-- 2. VEHICLES TABLE
-- ========================================================
CREATE TABLE IF NOT EXISTS public.vehicles (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    brand TEXT NOT NULL,
    plate_number TEXT NOT NULL UNIQUE,
    status TEXT NOT NULL DEFAULT 'Disponible',
    current_mileage INTEGER NOT NULL DEFAULT 0,
    fuel_level TEXT DEFAULT '100%',
    current_driver TEXT,
    created_at TIMESTAMPTZ DEFAULT NOW()
);

-- ========================================================
-- 3. VEHICLE HANDOVERS TABLE
-- ========================================================
CREATE TABLE IF NOT EXISTS public.vehicle_handovers (
    id TEXT PRIMARY KEY,
    vehicle_id UUID NOT NULL REFERENCES public.vehicles(id) ON DELETE CASCADE,
    driver_name TEXT NOT NULL,
    action_type TEXT NOT NULL CHECK (action_type IN ('Prise', 'Retour')),
    mileage INTEGER NOT NULL,
    fuel_level TEXT NOT NULL DEFAULT '100%',
    remarque TEXT DEFAULT '',
    photo_urls JSONB DEFAULT '[]'::jsonb,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_handovers_vehicle_id ON public.vehicle_handovers(vehicle_id);
CREATE INDEX IF NOT EXISTS idx_handovers_created_at ON public.vehicle_handovers(created_at DESC);
```

---
*End of Supabase Data Architecture & Schema Documentation.*
