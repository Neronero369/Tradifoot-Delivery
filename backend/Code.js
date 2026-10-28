/**
 * =========================================================================
 * TRADIFOOT LOGISTICS — PROGRAMME DE LIVRAISON (WEB APP & INGESTION ENGINE)
 * =========================================================================
 * 
 * Serverless Google Apps Script Web App providing bidirectional synchronization,
 * Role-Based Access Control (RBAC) data filtering, and an Enterprise Upsert
 * (Update & Insert) Ingestion Engine with Status Preservation.
 * 
 * Active Deployment URL:
 * https://script.google.com/macros/s/AKfycbwM5y2SPfnT1lMg4_ingmt_ZWzfHFb7l-qTwTow3-eThQ52Rf1FfLRrcBertuvxUY6d/exec
 * 
 * Expected Column Schema (8 Columns):
 * [ID, Date, Magasin, Chauffeur, Matricule, Quantite, Status, Date Mise à Jour]
 * =========================================================================
 */

const SHEET_NAME = "Feuille 1";
const SCHEMA_HEADERS = ["ID", "Date", "Magasin", "Chauffeur", "Matricule", "Quantite", "Status", "Date Mise à Jour"];

// =========================================================================
// 1. DATA SANITIZATION & COMPOSITE KEY HELPERS
// =========================================================================

/**
 * Normalizes any date input (Date instance, timestamp, French DD/MM/YYYY, ISO)
 * into standard YYYY-MM-DD format.
 * 
 * @param {*} val - Raw date value
 * @return {string} Canonical YYYY-MM-DD string
 */
function normalizeDate(val) {
  if (!val) return "";
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return "";
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  const s = String(val).trim();
  // Match YYYY-MM-DD or YYYY/MM/DD
  const mIso = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (mIso) {
    return `${mIso[1]}-${String(mIso[2]).padStart(2, '0')}-${String(mIso[3]).padStart(2, '0')}`;
  }
  // Match French/European DD/MM/YYYY or DD-MM-YYYY
  const mFr = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
  if (mFr) {
    return `${mFr[3]}-${String(mFr[2]).padStart(2, '0')}-${String(mFr[1]).padStart(2, '0')}`;
  }
  return s;
}

/**
 * Normalizes text by trimming, converting to uppercase, and collapsing multiple spaces.
 * 
 * @param {*} val - Raw text
 * @return {string} Sanitized uppercase text
 */
function normalizeText(val) {
  if (val === null || val === undefined) return "";
  return String(val).trim().toUpperCase().replace(/\s+/g, ' ');
}

/**
 * Checks if a row's assigned driver matches the given target chauffeur name,
 * supporting exact match, joint delivery substring match (e.g. "ADEL + BILLEL"),
 * and common phonetic / spelling variations (e.g. "SEIF-EDINE" vs "SEIF-EDDINE").
 * 
 * @param {string} rowDriver - Driver string from sheet row
 * @param {string} targetChauffeur - Target chauffeur name to match
 * @return {boolean}
 */
function isChauffeurMatch(rowDriver, targetChauffeur) {
  if (!rowDriver || !targetChauffeur) return false;
  const a = normalizeText(rowDriver);
  const b = normalizeText(targetChauffeur);
  if (a === b) return true;
  if (a.includes(b) || b.includes(a)) return true;

  const cleanA = a.replace(/[^A-Z0-9]/g, "").replace(/DD/g, "D");
  const cleanB = b.replace(/[^A-Z0-9]/g, "").replace(/DD/g, "D");
  if (cleanA === cleanB) return true;
  if (cleanA.includes(cleanB) || cleanB.includes(cleanA)) return true;

  return false;
}

/**
 * Builds an immutable composite key from [Date + Magasin + Chauffeur].
 * Guarantees row identity even when source Excel files lack unique delivery IDs.
 * 
 * @param {*} dateVal - Delivery date
 * @param {*} storeVal - Store / Magasin name
 * @param {*} driverVal - Assigned Chauffeur name
 * @return {string} Normalized composite key
 */
function generateCompositeKey(dateVal, storeVal, driverVal) {
  const normDate = normalizeDate(dateVal);
  const normStore = normalizeText(storeVal);
  const normDriver = normalizeText(driverVal);
  return `${normDate}|||${normStore}|||${normDriver}`;
}

/**
 * Normalizes delivery status strictly into "Livré" or "Non Livré".
 * 
 * @param {*} val - Raw status string
 * @return {string} "Livré" or "Non Livré"
 */
function normalizeStatus(val) {
  if (!val) return "Non Livré";
  const s = String(val).trim().toLowerCase();
  if (s.includes("livr") && !s.includes("non")) {
    return "Livré";
  }
  return "Non Livré";
}

// =========================================================================
// 2. ENTERPRISE UPSERT INGESTION ENGINE (STATUS PRESERVATION)
// =========================================================================

/**
 * Ingests incoming delivery manifests (from Mohammed's Excel files) using
 * non-destructive Upsert (Update & Insert) logic with strict Status Preservation.
 * 
 * STATUS PRESERVATION CONTRACT:
 * 1. Read existing rows into memory once.
 * 2. If row exists and status is 'Livré' -> PRESERVE 'Livré' (never downgrade).
 * 3. If row exists and status is 'Non Livré' -> update mutable details (Quantité, Matricule)
 *    and keep 'Non Livré' (unless incoming explicitly marks 'Livré').
 * 4. If row is brand new -> insert as new row with status 'Non Livré'.
 * 5. NEVER call sheet.clear() or clearContents(). Historical days remain untouched.
 * 
 * @param {Sheet} sheet - Google Sheet instance
 * @param {Array<Object|Array>} incomingData - List of delivery objects or 2D array rows
 * @return {Object} Audit stats report
 */
function upsertDeliveryManifest(sheet, incomingData) {
  const lock = LockService.getScriptLock();
  // Acquire lock for up to 30s to prevent concurrent collisions with driver updates
  lock.waitLock(30000);

  const stats = {
    totalIncoming: 0,
    preservedLivre: 0,
    upgradedToLivre: 0,
    updatedNonLivre: 0,
    insertedNew: 0,
    skipped: 0,
    errors: []
  };

  try {
    // 1. Ensure header row exists
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(SCHEMA_HEADERS);
    }

    const lastRow = sheet.getLastRow();
    const idMap = new Map();
    const compositeMap = new Map();
    let maxNumericId = 1000;

    // 2. Fetch all existing rows in memory (Single batch read)
    if (lastRow > 1) {
      const existingValues = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
      for (let i = 0; i < existingValues.length; i++) {
        const row = existingValues[i];
        const rowIndex = i + 2; // 1-based index (Header is row 1)
        const rowId = String(row[0] || "").trim();
        const rowDate = normalizeDate(row[1]);
        const rowStore = normalizeText(row[2]);
        const rowDriver = normalizeText(row[3]);
        const rowMatricule = normalizeText(row[4]);
        const rowQty = Number(row[5]) || 0;
        const rowStatus = normalizeStatus(row[6]);

        // Track max numeric ID for auto-incrementing new rows
        const numMatch = rowId.match(/\d+/);
        if (numMatch) {
          const n = parseInt(numMatch[0], 10);
          if (n > maxNumericId) maxNumericId = n;
        }

        const record = {
          rowIndex: rowIndex,
          id: rowId,
          date: rowDate,
          magasin: rowStore,
          chauffeur: rowDriver,
          matricule: rowMatricule,
          quantite: rowQty,
          status: rowStatus,
          timestamp: row[7]
        };

        if (rowId) {
          idMap.set(rowId.toUpperCase(), record);
        }

        const compKey = generateCompositeKey(rowDate, rowStore, rowDriver);
        if (rowDate && rowStore) {
          compositeMap.set(compKey, record);
        }
      }
    }

    // 3. Process each incoming manifest row
    const rowsToAppend = [];
    const rowsToUpdate = [];
    const timestamp = new Date();

    if (!Array.isArray(incomingData)) {
      throw new Error("Invalid incoming data: expected an array.");
    }

    for (let idx = 0; idx < incomingData.length; idx++) {
      const raw = incomingData[idx];
      stats.totalIncoming++;

      // Convert 2D array row to structured object if needed
      let item = {};
      if (Array.isArray(raw)) {
        // Skip header row if present in incoming array
        if (idx === 0 && (String(raw[0]).toUpperCase() === "ID" || String(raw[1]).toUpperCase() === "DATE")) {
          stats.skipped++;
          continue;
        }
        item = {
          id: raw[0],
          date: raw[1],
          magasin: raw[2],
          chauffeur: raw[3],
          matricule: raw[4],
          quantite: raw[5],
          status: raw[6],
          remarques: raw[7]
        };
      } else if (raw && typeof raw === "object") {
        item = raw;
      } else {
        stats.skipped++;
        continue;
      }

      const inId = String(item.id || "").trim();
      const inDate = normalizeDate(item.date);
      const inStore = normalizeText(item.magasin || item.destination);
      const inDriver = normalizeText(item.chauffeur || item.driver);
      const inMatricule = normalizeText(item.matricule || item.vehicle || "Camion Porteur");
      const inQty = Number(item.quantite !== undefined ? item.quantite : item.quantity) || 0;
      const inStatus = normalizeStatus(item.status);

      // Must have at least a date or store to be a valid delivery record
      if (!inDate && !inStore) {
        stats.skipped++;
        continue;
      }

      const compKey = generateCompositeKey(inDate, inStore, inDriver);

      // Attempt matching: First by explicit ID, second by composite key [Date + Magasin + Chauffeur]
      let existingRecord = null;
      if (inId && idMap.has(inId.toUpperCase())) {
        existingRecord = idMap.get(inId.toUpperCase());
      } else if (compKey && compositeMap.has(compKey)) {
        existingRecord = compositeMap.get(compKey);
      }

      if (existingRecord) {
        // =================================================================
        // UPSERT LOGIC: UPDATE EXISTING RECORD WITH STATUS PRESERVATION
        // =================================================================
        const existingStatus = existingRecord.status;
        let finalStatus = "Non Livré";

        if (existingStatus === "Livré") {
          // CRITICAL: NEVER DOWNGRADE 'Livré' TO 'Non Livré' OR EMPTY
          finalStatus = "Livré";
          stats.preservedLivre++;
        } else {
          // Current status is 'Non Livré'
          if (inStatus === "Livré") {
            finalStatus = "Livré";
            stats.upgradedToLivre++;
          } else {
            finalStatus = "Non Livré";
            stats.updatedNonLivre++;
          }
        }

        // Keep original stable ID and creation keys; update mutable quantity and matricule
        const stableId = existingRecord.id || inId;
        const finalQty = inQty > 0 ? inQty : existingRecord.quantite;
        const finalMat = inMatricule || existingRecord.matricule;

        const updatedRowValues = [
          stableId,
          existingRecord.date || inDate,
          existingRecord.magasin || inStore,
          existingRecord.chauffeur || inDriver,
          finalMat,
          finalQty,
          finalStatus,
          timestamp
        ];

        rowsToUpdate.push({
          rowIndex: existingRecord.rowIndex,
          values: updatedRowValues
        });

        // Update in-memory cache
        existingRecord.status = finalStatus;
        existingRecord.quantite = finalQty;
        existingRecord.matricule = finalMat;
      } else {
        // =================================================================
        // UPSERT LOGIC: INSERT BRAND NEW RECORD
        // =================================================================
        maxNumericId++;
        const newId = inId && !inId.startsWith("AUTO-") ? inId : `DEL-${String(maxNumericId).padStart(4, '0')}`;
        // New record status: defaults to 'Non Livré' unless manifest explicitly provides 'Livré'
        const finalStatus = inStatus === "Livré" ? "Livré" : "Non Livré";

        const newRowValues = [
          newId,
          inDate,
          inStore,
          inDriver,
          inMatricule,
          inQty,
          finalStatus,
          timestamp
        ];

        rowsToAppend.push(newRowValues);
        stats.insertedNew++;

        // Add to in-memory maps so subsequent rows in same batch merge correctly
        const newRecord = {
          rowIndex: lastRow + rowsToAppend.length,
          id: newId,
          date: inDate,
          magasin: inStore,
          chauffeur: inDriver,
          matricule: inMatricule,
          quantite: inQty,
          status: finalStatus,
          timestamp: timestamp
        };
        idMap.set(newId.toUpperCase(), newRecord);
        compositeMap.set(compKey, newRecord);
      }
    }

    // 4. Non-Destructive In-Place Updates (Never clearing sheet)
    for (let u = 0; u < rowsToUpdate.length; u++) {
      const up = rowsToUpdate[u];
      sheet.getRange(up.rowIndex, 1, 1, 8).setValues([up.values]);
    }

    // 5. Batch Append New Rows
    if (rowsToAppend.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, 8).setValues(rowsToAppend);
    }

    // Spreadsheet flush to guarantee persistence
    SpreadsheetApp.flush();

    return stats;

  } finally {
    lock.releaseLock();
  }
}

// =========================================================================
// 3. HTTP POST CONTROLLER (API DISPATCHER)
// =========================================================================

/**
 * Handles HTTP POST requests sent by the client dashboard or automated Python pipelines.
 * Supported actions:
 *   - 'get_deliveries' -> retrieves deliveries with RBAC filtering
 *   - 'update_status'  -> single-row atomic status toggle from driver
 *   - 'batch_upsert'   -> multi-row manifest ingestion with status preservation
 *   - 'import_manifest'-> alias for batch_upsert
 * 
 * @param {Object} e - Event object containing postData.contents or parameter
 * @return {ContentService.TextOutput} JSON response
 */
function doPost(e) {
  try {
    let data = {};
    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (parseErr) {
        data = e.parameter || {};
      }
    } else if (e && e.parameter) {
      data = e.parameter;
    }

    const action = String(data.action || "update_status").trim().toLowerCase();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.getActiveSheet();

    // Ensure header row exists
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(SCHEMA_HEADERS);
    }

    // -------------------------------------------------------------
    // ACTION: get_deliveries (RBAC Read)
    // -------------------------------------------------------------
    if (action === "get_deliveries") {
      const role = data.role || "admin";
      const chauffeurName = data.chauffeurName || data.chauffeur || data.driver || "";
      const deliveries = fetchDeliveriesFromSheet(sheet, role, chauffeurName);
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        action: "get_deliveries",
        role: role,
        chauffeurName: chauffeurName,
        total: deliveries.length,
        deliveries: deliveries
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // -------------------------------------------------------------
    // ACTION: batch_upsert / import_manifest (Ingestion with Status Preservation)
    // -------------------------------------------------------------
    if (action === "batch_upsert" || action === "import_manifest") {
      const incomingList = data.deliveries || data.records || data.rows || [];
      if (!Array.isArray(incomingList) || incomingList.length === 0) {
        return ContentService.createTextOutput(JSON.stringify({
          status: "error",
          message: "No deliveries provided in payload."
        })).setMimeType(ContentService.MimeType.JSON);
      }

      const audit = upsertDeliveryManifest(sheet, incomingList);
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        action: "batch_upsert",
        audit: audit,
        timestamp: new Date().toISOString()
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // -------------------------------------------------------------
    // ACTION: update_status (Single-row mutation from driver)
    // -------------------------------------------------------------
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    const id = String(data.id || "").trim();
    const date = normalizeDate(data.date);
    const magasin = normalizeText(data.magasin || data.destination);
    const chauffeur = normalizeText(data.chauffeur || data.driver);
    const matricule = normalizeText(data.matricule || data.vehicle);
    const quantite = data.quantite !== undefined ? data.quantite : (data.quantity !== undefined ? data.quantity : "");
    const status = normalizeStatus(data.status || "Livré");
    const timestamp = new Date();

    let updated = false;
    const lastRow = sheet.getLastRow();

    // Check by ID or composite key
    if (lastRow > 1) {
      const values = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
      const targetCompKey = generateCompositeKey(date, magasin, chauffeur);

      for (let i = 0; i < values.length; i++) {
        const rowId = String(values[i][0] || "").trim();
        const rowCompKey = generateCompositeKey(values[i][1], values[i][2], values[i][3]);

        if ((id && rowId.toUpperCase() === id.toUpperCase()) || (targetCompKey && rowCompKey === targetCompKey)) {
          const targetRow = i + 2;
          const finalId = rowId || id;
          const finalDate = date || normalizeDate(values[i][1]);
          const finalMagasin = magasin || values[i][2];
          const finalChauffeur = chauffeur || values[i][3];
          const finalMatricule = matricule || values[i][4];
          const finalQty = quantite !== "" ? quantite : values[i][5];

          sheet.getRange(targetRow, 1, 1, 8).setValues([[
            finalId, finalDate, finalMagasin, finalChauffeur, finalMatricule, finalQty, status, timestamp
          ]]);
          updated = true;
          break;
        }
      }
    }

    // Append new row if record does not exist yet
    if (!updated) {
      const newId = id || `DEL-${String(lastRow + 1).padStart(4, '0')}`;
      sheet.appendRow([newId, date, magasin, chauffeur, matricule, quantite, status, timestamp]);
    }

    lock.releaseLock();

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      action: updated ? "updated" : "inserted",
      id: id,
      newStatus: status,
      timestamp: timestamp.toISOString()
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// =========================================================================
// 4. HTTP GET CONTROLLER (READ & HEALTHCHECK)
// =========================================================================

/**
 * Handles HTTP GET requests for diagnostics and read-only delivery fetching.
 * 
 * @param {Object} e - Event object
 * @return {ContentService.TextOutput} JSON response
 */
function doGet(e) {
  try {
    const action = e && e.parameter && e.parameter.action ? e.parameter.action.toLowerCase() : "";
    if (action === "get_deliveries") {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      let sheet = spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.getActiveSheet();
      const role = e.parameter.role || "admin";
      const chauffeurName = e.parameter.chauffeurName || e.parameter.chauffeur || "";
      const deliveries = fetchDeliveriesFromSheet(sheet, role, chauffeurName);

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        action: "get_deliveries",
        role: role,
        chauffeurName: chauffeurName,
        total: deliveries.length,
        deliveries: deliveries
      })).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: "online",
      service: "Tradifoot Logistics Delivery Web App API (Upsert & Status Preservation Enabled)",
      schema: SCHEMA_HEADERS,
      supportedActions: ["get_deliveries", "update_status", "batch_upsert", "import_manifest"],
      timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// =========================================================================
// 5. HELPER: FETCH DELIVERIES WITH FIELD ALIASES
// =========================================================================

/**
 * Retrieves deliveries from the sheet, supporting RBAC and providing dual
 * field aliases (e.g. magasin/destination, chauffeur/driver) for 100% frontend compatibility.
 * 
 * @param {Sheet} sheet - Google Sheet
 * @param {string} role - 'admin' or 'chauffeur'
 * @param {string} chauffeurName - Filter driver name
 * @return {Array<Object>} List of deliveries
 */
function fetchDeliveriesFromSheet(sheet, role, chauffeurName) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  const normalizedRole = String(role || "admin").toLowerCase();
  const targetDriver = chauffeurName ? normalizeText(chauffeurName) : "";
  const deliveries = [];

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowId = String(row[0] || "").trim();
    if (!rowId) continue;

    const rowDriver = String(row[3] || "").trim();

    // RBAC Filter:
    // Chauffeur gets ONLY the delivery assigned to him
    if (normalizedRole === 'chauffeur') {
      if (!targetDriver || !isChauffeurMatch(rowDriver, targetDriver)) {
        continue;
      }
    } else if (targetDriver && targetDriver !== 'ALL') {
      // Optional driver filtering for admin queries
      if (!isChauffeurMatch(rowDriver, targetDriver)) {
        continue;
      }
    }

    const rowStatus = normalizeStatus(row[6]);
    const store = String(row[2] || "").trim();
    const matricule = String(row[4] || "").trim();
    const qty = Number(row[5]) || 0;

    deliveries.push({
      id: rowId,
      date: normalizeDate(row[1]),
      magasin: store,
      destination: store, // Dual alias for frontend compatibility
      chauffeur: rowDriver,
      driver: rowDriver,   // Dual alias for frontend compatibility
      matricule: matricule,
      vehicle: matricule, // Dual alias for frontend compatibility
      quantite: qty,
      quantity: qty,      // Dual alias for frontend compatibility
      status: rowStatus,
      completed: (rowStatus === "Livré"),
      lastUpdated: row[7] ? new Date(row[7]).toISOString() : null
    });
  }

  return deliveries;
}

// =========================================================================
// 6. GOOGLE SHEETS UI MENU INTEGRATION & MANUAL RUNNERS
// =========================================================================

/**
 * Creates custom menu in Google Sheets UI upon opening the spreadsheet.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🚚 Tradifoot Logistics")
    .addItem("📥 Importer Programme de Livraison (Upsert Sécurisé)", "menuImportProgramme")
    .addItem("🔍 Tester la Préservation des Statuts (Auto-Test)", "testUpsertPreservation")
    .addToUi();
}

/**
 * Menu action to import an Excel sheet tab into the master database
 * with status preservation.
 */
function menuImportProgramme() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const response = ui.prompt(
    "Importer Programme de Livraison",
    "Entrez le nom de la feuille contenant le nouveau fichier Excel de Mohammed (ex: 'Import' ou 'Feuille 2') :",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;
  const sourceSheetName = response.getResponseText().trim();
  const sourceSheet = ss.getSheetByName(sourceSheetName);

  if (!sourceSheet) {
    ui.alert("Erreur", `La feuille '${sourceSheetName}' est introuvable dans ce classeur.`, ui.ButtonSet.OK);
    return;
  }

  const destSheet = ss.getSheetByName(SHEET_NAME) || ss.getActiveSheet();
  const dataRange = sourceSheet.getDataRange();
  const incomingRows = dataRange.getValues();

  const stats = upsertDeliveryManifest(destSheet, incomingRows);

  ui.alert(
    "Ingestion Réussie !",
    `Rapport d'Upsert Tradifoot:\n` +
    `- Lignes traitées : ${stats.totalIncoming}\n` +
    `- Statuts 'Livré' préservés : ${stats.preservedLivre}\n` +
    `- 'Non Livré' actualisés : ${stats.updatedNonLivre}\n` +
    `- Nouvelles livraisons créées : ${stats.insertedNew}\n` +
    `- Lignes ignorées : ${stats.skipped}\n\nAucune donnée historique n'a été écrasée.`,
    ui.ButtonSet.OK
  );
}

/**
 * Self-test suite to verify Status Preservation and Non-Destructive behavior.
 */
function testUpsertPreservation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const testSheetName = "__TEST_UPSERT_SANDBOX__";
  let testSheet = ss.getSheetByName(testSheetName);
  if (testSheet) ss.deleteSheet(testSheet);
  testSheet = ss.insertSheet(testSheetName);

  Logger.log("--- STARTING UPSERT PRESERVATION SELF-TEST ---");

  // Step 1: Initial load
  const initialData = [
    SCHEMA_HEADERS,
    ["DEL-101", "2026-10-28", "SKECHERS BAB EZZOUAR", "SAADI BILLEL", "MERCEDES", 150, "Livré", new Date()],
    ["DEL-102", "2026-10-28", "CLARKS GARDEN CITY", "SNOUCI SEIF-EDDINE", "MERCEDES", 200, "Non Livré", new Date()]
  ];
  testSheet.getRange(1, 1, initialData.length, 8).setValues(initialData);

  // Step 2: Ingest manifest from Mohammed (DEL-101 has status 'Non Livré' in Excel, DEL-102 has updated qty, DEL-103 is new)
  const incomingManifest = [
    { id: "DEL-101", date: "2026-10-28", magasin: "SKECHERS BAB EZZOUAR", chauffeur: "SAADI BILLEL", matricule: "MERCEDES-NEW", quantite: 180, status: "Non Livré" },
    { id: "DEL-102", date: "2026-10-28", magasin: "CLARKS GARDEN CITY", chauffeur: "SNOUCI SEIF-EDDINE", matricule: "MERCEDES", quantite: 250, status: "Non Livré" },
    { id: "", date: "2026-10-28", magasin: "GEOX DIDOUCHE", chauffeur: "BOUKHALFA ADEL", matricule: "DOBLO", quantite: 80, status: "Non Livré" }
  ];

  const report = upsertDeliveryManifest(testSheet, incomingManifest);
  Logger.log("Upsert Report: " + JSON.stringify(report));

  const resultRows = testSheet.getDataRange().getValues();

  // Assertions
  const row101 = resultRows.find(r => r[0] === "DEL-101");
  const row102 = resultRows.find(r => r[0] === "DEL-102");
  const row103 = resultRows.find(r => r[2] === "GEOX DIDOUCHE");

  let passed = true;

  // 1. DEL-101 MUST remain 'Livré'
  if (!row101 || row101[6] !== "Livré") {
    Logger.log("FAIL: DEL-101 status was downgraded to " + (row101 ? row101[6] : "missing"));
    passed = false;
  } else if (row101[4] !== "MERCEDES-NEW" || row101[5] !== 180) {
    Logger.log("FAIL: DEL-101 details were not updated");
    passed = false;
  } else {
    Logger.log("PASS: DEL-101 preserved 'Livré' status and updated details.");
  }

  // 2. DEL-102 must remain 'Non Livré' and have updated quantity 250
  if (!row102 || row102[6] !== "Non Livré" || row102[5] !== 250) {
    Logger.log("FAIL: DEL-102 failed update");
    passed = false;
  } else {
    Logger.log("PASS: DEL-102 remained 'Non Livré' and updated quantity to 250.");
  }

  // 3. Brand new delivery DEL-103 inserted with 'Non Livré'
  if (!row103 || row103[6] !== "Non Livré") {
    Logger.log("FAIL: Brand new delivery failed insertion");
    passed = false;
  } else {
    Logger.log("PASS: Brand new delivery correctly inserted as 'Non Livré'.");
  }

  // Cleanup test sheet
  ss.deleteSheet(testSheet);
  Logger.log("--- TEST COMPLETE: " + (passed ? "ALL PASSED" : "FAILED") + " ---");
  return passed;
}
