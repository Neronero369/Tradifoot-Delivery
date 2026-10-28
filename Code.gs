/**
 * =========================================================================
 * TRADIFOOT LOGISTICS — PROGRAMME DE LIVRAISON (WEB APP & INGESTION ENGINE)
 * =========================================================================
 * 
 * Serverless Google Apps Script Web App providing bidirectional synchronization,
 * Role-Based Access Control (RBAC) data filtering, and an Enterprise Upsert
 * (Update & Insert) Ingestion Engine with Status and Chauffeur Preservation.
 * 
 * Active Deployment URL:
 * https://script.google.com/macros/s/AKfycbwqLs9WYjdjESCnX9avFBABv2bZdTXh1a43rYDC6N51l4uIO6em0wS51OaTyTuMAOY/exec
 * 
 * Expected Column Schema (8 Columns):
 * [ID, Date, Magasin, Chauffeur, Matricule, Quantite, Status, Date Mise à Jour]
 *   A     B      C         D          E         F        G            H
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
  const mIso = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (mIso) {
    return `${mIso[1]}-${String(mIso[2]).padStart(2, '0')}-${String(mIso[3]).padStart(2, '0')}`;
  }
  const mFr = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
  if (mFr) {
    return `${mFr[3]}-${String(mFr[2]).padStart(2, '0')}-${String(mFr[1]).padStart(2, '0')}`;
  }
  return s;
}

/**
 * Normalizes text by trimming, converting to uppercase, and collapsing spaces.
 */
function normalizeText(val) {
  if (val === null || val === undefined) return "";
  return String(val).trim().toUpperCase().replace(/\s+/g, ' ');
}

/**
 * Checks if a row's assigned driver matches target chauffeur name.
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
 */
function generateCompositeKey(dateVal, storeVal, driverVal) {
  const normDate = normalizeDate(dateVal);
  const normStore = normalizeText(storeVal);
  const normDriver = normalizeText(driverVal);
  return `${normDate}|||${normStore}|||${normDriver}`;
}

/**
 * Normalizes delivery status strictly into "Livré" or "Non Livré".
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
// 2. HTTP GET HANDLER (RBAC READ & HEALTH CHECK)
// =========================================================================

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
// 3. HTTP POST HANDLER (MUTATIONS & SYNCHRONIZATION)
// =========================================================================

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
    // ACTION: get_deliveries (RBAC Read via POST)
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
    // ACTION: update_status (Single-row mutation from driver or admin)
    // -------------------------------------------------------------
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    const id = String(data.id || "").trim();
    const rawDate = data.date || "";
    const date = normalizeDate(rawDate);
    const magasin = normalizeText(data.magasin || data.destination || "");
    
    // Extract chauffeur value from the incoming JSON payload
    const incomingChauffeur = normalizeText(data.chauffeur || data.driver || data.chauffeurName || "");
    const matricule = normalizeText(data.matricule || data.vehicle || "");
    const quantite = data.quantite !== undefined ? data.quantite : (data.quantity !== undefined ? data.quantity : "");
    const status = normalizeStatus(data.status || "Livré");
    const timestamp = new Date();

    let updated = false;
    let finalChauffeurName = incomingChauffeur;
    const lastRow = sheet.getLastRow();

    // Column mapping:
    // A (1): ID
    // B (2): Date
    // C (3): Magasin
    // D (4): Chauffeur  <-- COLUMN D
    // E (5): Matricule
    // F (6): Quantite
    // G (7): Status     <-- COLUMN G
    // H (8): Date Mise à Jour
    if (lastRow > 1) {
      const values = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
      const targetCompKey = generateCompositeKey(date, magasin, incomingChauffeur);

      for (let i = 0; i < values.length; i++) {
        const rowId = String(values[i][0] || "").trim();
        const rowDate = normalizeDate(values[i][1]);
        const rowMagasin = normalizeText(values[i][2]);
        const rowChauffeur = normalizeText(values[i][3]);
        const rowCompKey = `${rowDate}|||${rowMagasin}|||${rowChauffeur}`;

        // Match by ID or composite key
        if ((id && rowId.toUpperCase() === id.toUpperCase()) || (targetCompKey && rowCompKey === targetCompKey)) {
          const targetRow = i + 2;
          const finalId = rowId || id;
          const finalDate = date || rowDate;
          const finalMagasin = magasin || values[i][2];

          // Priority for Chauffeur column (Column D):
          // 1. Incoming payload chauffeur (e.g. SAADI BILLEL)
          // 2. Existing sheet row chauffeur (if not empty / not "NON SPÉCIFIÉ")
          // 3. Fallback to "NON SPÉCIFIÉ" only if both are completely empty
          const existingChauffeur = String(values[i][3] || "").trim();
          finalChauffeurName = incomingChauffeur || (existingChauffeur && existingChauffeur.toUpperCase() !== "NON SPÉCIFIÉ" ? existingChauffeur : "") || "NON SPÉCIFIÉ";

          const finalMatricule = matricule || values[i][4];
          const finalQty = quantite !== "" ? quantite : values[i][5];

          // Write updated values: Column D = finalChauffeurName, Column G = status ("Livré")
          sheet.getRange(targetRow, 1, 1, 8).setValues([[
            finalId,
            finalDate,
            finalMagasin,
            finalChauffeurName,
            finalMatricule,
            finalQty,
            status,
            timestamp
          ]]);
          updated = true;
          break;
        }
      }
    }

    // Append new row if delivery record does not exist yet
    if (!updated) {
      const newId = id || `DEL-${String(lastRow + 1).padStart(4, '0')}`;
      finalChauffeurName = incomingChauffeur || "NON SPÉCIFIÉ";
      sheet.appendRow([
        newId,
        date,
        magasin,
        finalChauffeurName,
        matricule,
        quantite,
        status,
        timestamp
      ]);
      updated = true;
    }

    lock.releaseLock();

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      action: "update_status",
      id: id,
      updatedStatus: status,
      chauffeur: finalChauffeurName,
      timestamp: timestamp.toISOString()
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// =========================================================================
// 4. HELPER: FETCH DELIVERIES WITH FIELD ALIASES
// =========================================================================

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
    if (normalizedRole === 'chauffeur') {
      if (!targetDriver || !isChauffeurMatch(rowDriver, targetDriver)) {
        continue;
      }
    } else if (targetDriver && targetDriver !== 'ALL') {
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
      destination: store,
      chauffeur: rowDriver,
      driver: rowDriver,
      matricule: matricule,
      vehicle: matricule,
      quantite: qty,
      quantity: qty,
      status: rowStatus,
      completed: (rowStatus === "Livré"),
      lastUpdated: row[7] ? new Date(row[7]).toISOString() : null
    });
  }

  return deliveries;
}

// =========================================================================
// 5. ENTERPRISE UPSERT INGESTION ENGINE (STATUS PRESERVATION)
// =========================================================================

function upsertDeliveryManifest(sheet, incomingRecords) {
  const lastRow = sheet.getLastRow();
  const existingMap = new Map();

  if (lastRow > 1) {
    const dataRange = sheet.getRange(2, 1, lastRow - 1, 8);
    const existingValues = dataRange.getValues();

    for (let i = 0; i < existingValues.length; i++) {
      const r = existingValues[i];
      const rId = String(r[0] || "").trim();
      const compKey = generateCompositeKey(r[1], r[2], r[3]);

      const record = {
        rowIndex: i + 2,
        id: rId,
        date: normalizeDate(r[1]),
        magasin: r[2],
        chauffeur: r[3],
        matricule: r[4],
        quantite: r[5],
        status: normalizeStatus(r[6]),
        lastUpdated: r[7]
      };

      if (rId) existingMap.set(`ID:${rId.toUpperCase()}`, record);
      if (compKey) existingMap.set(`KEY:${compKey}`, record);
    }
  }

  const rowsToUpdate = [];
  const rowsToInsert = [];
  const now = new Date();

  let preservedLivreCount = 0;
  let updatedNonLivreCount = 0;
  let insertedNewCount = 0;
  let skippedCount = 0;

  for (let idx = 0; idx < incomingRecords.length; idx++) {
    const raw = incomingRecords[idx];
    if (!raw) { skippedCount++; continue; }

    let incId = "", incDate = "", incStore = "", incDriver = "", incMat = "", incQty = 0, incStatus = "Non Livré";

    if (Array.isArray(raw)) {
      if (idx === 0 && (String(raw[0]).toUpperCase() === "ID" || String(raw[1]).toUpperCase() === "DATE")) {
        skippedCount++;
        continue;
      }
      incId = String(raw[0] || "").trim();
      incDate = normalizeDate(raw[1]);
      incStore = String(raw[2] || "").trim();
      incDriver = String(raw[3] || "").trim();
      incMat = String(raw[4] || "").trim();
      incQty = Number(raw[5]) || 0;
      incStatus = normalizeStatus(raw[6]);
    } else if (typeof raw === 'object') {
      incId = String(raw.id || "").trim();
      incDate = normalizeDate(raw.date);
      incStore = String(raw.magasin || raw.destination || "").trim();
      incDriver = String(raw.chauffeur || raw.driver || "").trim();
      incMat = String(raw.matricule || raw.vehicle || "").trim();
      incQty = Number(raw.quantite !== undefined ? raw.quantite : raw.quantity) || 0;
      incStatus = normalizeStatus(raw.status);
    }

    if (!incDate && !incStore) { skippedCount++; continue; }

    const compKey = generateCompositeKey(incDate, incStore, incDriver);
    let match = null;
    if (incId && existingMap.has(`ID:${incId.toUpperCase()}`)) {
      match = existingMap.get(`ID:${incId.toUpperCase()}`);
    } else if (compKey && existingMap.has(`KEY:${compKey}`)) {
      match = existingMap.get(`KEY:${compKey}`);
    }

    if (match) {
      const finalStatus = (match.status === "Livré") ? "Livré" : incStatus;
      if (match.status === "Livré") preservedLivreCount++;
      else updatedNonLivreCount++;

      rowsToUpdate.push({
        targetRow: match.rowIndex,
        values: [
          match.id || incId,
          incDate || match.date,
          incStore || match.magasin,
          incDriver || match.chauffeur || "NON SPÉCIFIÉ",
          incMat || match.matricule,
          incQty || match.quantite,
          finalStatus,
          now
        ]
      });
    } else {
      const assignedId = incId || `DEL-${String(lastRow + rowsToInsert.length + 1).padStart(4, '0')}`;
      rowsToInsert.push([
        assignedId,
        incDate,
        incStore,
        incDriver || "NON SPÉCIFIÉ",
        incMat,
        incQty,
        incStatus,
        now
      ]);
      insertedNewCount++;
    }
  }

  for (const item of rowsToUpdate) {
    sheet.getRange(item.targetRow, 1, 1, 8).setValues([item.values]);
  }

  if (rowsToInsert.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToInsert.length, 8).setValues(rowsToInsert);
  }

  return {
    totalIncoming: incomingRecords.length,
    preservedLivre: preservedLivreCount,
    updatedNonLivre: updatedNonLivreCount,
    insertedNew: insertedNewCount,
    skipped: skippedCount
  };
}

// =========================================================================
// 6. GOOGLE SHEETS UI MENU INTEGRATION
// =========================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🚚 Tradifoot Logistics")
    .addItem("📥 Importer Programme de Livraison (Upsert Sécurisé)", "menuImportProgramme")
    .addToUi();
}

function menuImportProgramme() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const response = ui.prompt(
    "Importer Programme de Livraison",
    "Entrez le nom de la feuille contenant le fichier Excel (ex: 'Import') :",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;
  const sourceSheetName = response.getResponseText().trim();
  const sourceSheet = ss.getSheetByName(sourceSheetName);

  if (!sourceSheet) {
    ui.alert("Erreur", `La feuille '${sourceSheetName}' est introuvable.`, ui.ButtonSet.OK);
    return;
  }

  const destSheet = ss.getSheetByName(SHEET_NAME) || ss.getActiveSheet();
  const incomingRows = sourceSheet.getDataRange().getValues();
  const stats = upsertDeliveryManifest(destSheet, incomingRows);

  ui.alert("Ingestion Réussie !", JSON.stringify(stats, null, 2), ui.ButtonSet.OK);
}
