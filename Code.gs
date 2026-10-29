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

const ADMIN_NAMES = [
  "MOUNIR",
  "MEKHLOUFI MOHAMMED",
  "BENAMAR NACIM",
  "ADMIN",
  "ALL"
];

function isAdminName(val) {
  if (!val) return false;
  const norm = normalizeText(val);
  return ADMIN_NAMES.some(function(admin) {
    return norm === admin || norm.replace(/\s+/g, '') === admin.replace(/\s+/g, '');
  });
}

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
/**
 * Builds a normalized primary composite key from [Date + Magasin + Quantite + Chauffeur].
 * Includes Number(qtyVal) || 0 so multiple deliveries on the same date/store with different quantities are uniquely identified.
 */
function generateCompositeKey(dateVal, storeVal, driverVal, qtyVal) {
  const normDate = normalizeDate(dateVal);
  const normStore = normalizeText(storeVal);
  const normDriver = normalizeText(driverVal);
  const normQty = Number(qtyVal) || 0;
  return normDate + "|" + normStore + "|" + normQty + "|" + normDriver;
}

/**
 * Builds a fallback composite key from [Date + Magasin + Quantite]
 * Used when Chauffeur is 'NON SPÉCIFIÉ' or an admin name.
 */
function generateFallbackKey(dateVal, storeVal, qtyVal) {
  const normDate = normalizeDate(dateVal);
  const normStore = normalizeText(storeVal);
  const normQty = Number(qtyVal) || 0;
  return normDate + "|" + normStore + "|" + normQty;
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
    // ACTION: update_status (Single-row mutation from chauffeur)
    // -------------------------------------------------------------
    const userRole = String(data.role || "").trim().toLowerCase();
    if (userRole === "admin") {
      return ContentService.createTextOutput(JSON.stringify({
        status: "error",
        message: "Admin accounts are strictly read-only and cannot alter delivery status."
      })).setMimeType(ContentService.MimeType.JSON);
    }

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

      for (let i = 0; i < values.length; i++) {
        const rowId = String(values[i][0] || "").trim();
        const rowDate = normalizeDate(values[i][1]);
        const rowMagasin = normalizeText(values[i][2]);
        const rowQty = Number(values[i][5]) || 0;

        // Match existing row by ID, Primary Composite Fingerprint, or Fallback Fingerprint
        const isIdMatch = Boolean(id && rowId && id.toUpperCase() === rowId.toUpperCase());
        const rowChauffeur = normalizeText(values[i][3]);
        const targetCompKey = generateCompositeKey(date, magasin, incomingChauffeur, quantite);
        const rowCompKey = generateCompositeKey(rowDate, rowMagasin, rowChauffeur, rowQty);
        const targetFbKey = generateFallbackKey(date, magasin, quantite);
        const rowFbKey = generateFallbackKey(rowDate, rowMagasin, rowQty);

        const isPrimaryMatch = Boolean(date && magasin && targetCompKey && rowCompKey && targetCompKey === rowCompKey);
        const isFallbackMatch = Boolean(date && magasin && targetFbKey && rowFbKey && targetFbKey === rowFbKey);

        if (isIdMatch || isPrimaryMatch || isFallbackMatch) {
          const targetRow = i + 2;
          const finalId = rowId || id || `LIV-${String(targetRow - 1).padStart(4, '0')}`;
          const finalDate = date || rowDate;
          const finalMagasin = values[i][2] || (data.Magasin || data.magasin || magasin);

          // Preserve original Chauffeur name (never write "NON SPÉCIFIÉ" if legitimate name exists)
          const existingChauffeur = String(values[i][3] || "").trim();
          if (existingChauffeur && existingChauffeur.toUpperCase() !== "NON SPÉCIFIÉ" && !isAdminName(existingChauffeur)) {
            finalChauffeurName = existingChauffeur;
          } else if (incomingChauffeur && incomingChauffeur.toUpperCase() !== "NON SPÉCIFIÉ" && !isAdminName(incomingChauffeur)) {
            finalChauffeurName = incomingChauffeur;
          } else if (existingChauffeur && existingChauffeur.toUpperCase() !== "NON SPÉCIFIÉ") {
            finalChauffeurName = existingChauffeur;
          } else {
            finalChauffeurName = incomingChauffeur || "NON SPÉCIFIÉ";
          }

          const finalMatricule = values[i][4] || matricule;
          const finalQty = rowQty || quantite;

          sheet.getRange(targetRow, 1, 1, 8).setValues([[
            finalId,
            finalDate,
            finalMagasin,
            finalChauffeurName,
            finalMatricule,
            finalQty,
            status, // "Livré"
            timestamp
          ]]);
          updated = true;
          break;
        }
      }
    }

    // Append complete row with all fields if not found yet (never blank / NaN columns)
    if (!updated) {
      const newId = id || `LIV-${String(lastRow + 1).padStart(4, '0')}`;
      finalChauffeurName = (incomingChauffeur && !isAdminName(incomingChauffeur)) ? incomingChauffeur : "NON SPÉCIFIÉ";
      const cleanStore = String(data.Magasin || data.magasin || magasin).trim();
      sheet.appendRow([
        newId,
        date,
        cleanStore,
        finalChauffeurName,
        matricule,
        quantite,
        status, // "Livré"
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
  const missingIdUpdates = [];

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowDate = normalizeDate(row[1]);
    const store = String(row[2] || "").trim();
    const qty = Number(row[5]) || 0;

    // Skip corrupted/empty rows where both Date and Magasin are empty
    if (!rowDate && !store) {
      continue;
    }

    let rowId = String(row[0] || "").trim();
    if (!rowId) {
      const rowIndex = i + 2;
      rowId = `LIV-${String(rowIndex - 1).padStart(4, '0')}`;
      missingIdUpdates.push({ row: rowIndex, id: rowId });
    }

    const rowDriver = String(row[3] || "").trim();

    // Role filter only if a specific chauffeur driver filter is explicitly provided
    if (normalizedRole === 'chauffeur' && targetDriver && targetDriver !== 'ALL') {
      if (!isChauffeurMatch(rowDriver, targetDriver)) {
        continue;
      }
    }

    const rowStatus = normalizeStatus(row[6]);
    const matricule = String(row[4] || "").trim();

    deliveries.push({
      id: rowId,
      ID: rowId,
      date: rowDate,
      Date: rowDate,
      magasin: store,
      Magasin: store,
      destination: store,
      chauffeur: rowDriver,
      Chauffeur: rowDriver,
      driver: rowDriver,
      matricule: matricule,
      Matricule: matricule,
      vehicle: matricule,
      quantite: qty,
      Quantite: qty,
      quantity: qty,
      status: rowStatus,
      Status: rowStatus,
      completed: (rowStatus === "Livré"),
      lastUpdated: row[7] ? new Date(row[7]).toISOString() : null
    });
  }

  // Persist generated Unique IDs back into Google Sheet Column A
  if (missingIdUpdates.length > 0) {
    try {
      for (let k = 0; k < missingIdUpdates.length; k++) {
        sheet.getRange(missingIdUpdates[k].row, 1).setValue(missingIdUpdates[k].id);
      }
    } catch (wbErr) {
      console.warn("Could not backfill Unique IDs to sheet:", wbErr);
    }
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
      const compKey = generateCompositeKey(r[1], r[2], r[3], r[5]);
      const fbKey = generateFallbackKey(r[1], r[2], r[5]);

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

    const compKey = generateCompositeKey(incDate, incStore, incDriver, incQty);
    const fbKey = generateFallbackKey(incDate, incStore, incQty);
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
