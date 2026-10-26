/**
 * =========================================================================
 * TRADIFOOT LOGISTICS — PROGRAMME DE LIVRAISON (WEB APP BACKEND)
 * =========================================================================
 * 
 * Serverless Google Apps Script Web App providing bidirectional synchronization
 * and Role-Based Access Control (RBAC) data filtering between the client-side
 * delivery dashboard and the operational Google Sheet.
 * 
 * Active Deployment URL:
 * https://script.google.com/macros/s/AKfycbwM5y2SPfnT1lMg4_ingmt_ZWzfHFb7l-qTwTow3-eThQ52Rf1FfLRrcBertuvxUY6d/exec
 * 
 * Expected Column Schema (8 Columns):
 * [ID, Date, Magasin, Chauffeur, Matricule, Quantite, Status, Date Mise à Jour]
 * =========================================================================
 */

const SHEET_NAME = "Feuille 1";

/**
 * Handles HTTP POST requests sent by the Delivery Calendar Dashboard.
 * Supports action-based dispatching:
 *   - action: 'get_deliveries' -> retrieves deliveries filtered by role & chauffeur
 *   - action: 'update_status'  -> updates delivery status with concurrency locking
 * 
 * @param {Object} e - Event object containing postData.contents or parameter
 * @return {ContentService.TextOutput} JSON response
 */
function doPost(e) {
  try {
    // 1. Parse incoming JSON payload or form parameter
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
      sheet.appendRow(["ID", "Date", "Magasin", "Chauffeur", "Matricule", "Quantite", "Status", "Date Mise à Jour"]);
    }

    // -------------------------------------------------------------
    // ACTION: get_deliveries
    // -------------------------------------------------------------
    if (action === "get_deliveries") {
      const deliveries = fetchDeliveriesFromSheet(sheet, data.role, data.chauffeurName);
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        action: "get_deliveries",
        role: data.role || "unknown",
        total: deliveries.length,
        deliveries: deliveries
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // -------------------------------------------------------------
    // ACTION: update_status (Default write/mutation)
    // -------------------------------------------------------------
    // Acquire a mutex lock for up to 10 seconds to prevent concurrent write collisions
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    const id = String(data.id || "").trim();
    const date = String(data.date || "").trim();
    const magasin = String(data.magasin || "").trim();
    const chauffeur = String(data.chauffeur || "").trim();
    const matricule = String(data.matricule || "").trim();
    const quantite = data.quantite !== undefined ? data.quantite : "";
    const status = String(data.status || "Livré").trim();
    const timestamp = new Date();

    let updated = false;
    const lastRow = sheet.getLastRow();

    // Check if record with matching ID already exists
    if (lastRow > 1 && id) {
      const idRange = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < idRange.length; i++) {
        if (String(idRange[i][0]).trim() === id) {
          const targetRow = i + 2;
          sheet.getRange(targetRow, 1, 1, 8).setValues([[
            id, date, magasin, chauffeur, matricule, quantite, status, timestamp
          ]]);
          updated = true;
          break;
        }
      }
    }

    // Append new row if not found
    if (!updated && id) {
      sheet.appendRow([id, date, magasin, chauffeur, matricule, quantite, status, timestamp]);
    }

    // Release script lock
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

/**
 * Helper function to retrieve and format deliveries from the sheet.
 * Supports RBAC filtering for chauffeurs.
 * 
 * @param {Sheet} sheet - Active Google Sheet
 * @param {string} role - 'admin' or 'chauffeur'
 * @param {string} chauffeurName - Name of the chauffeur
 * @return {Array<Object>} List of deliveries
 */
function fetchDeliveriesFromSheet(sheet, role, chauffeurName) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  const normalizedDriver = chauffeurName ? chauffeurName.toString().toUpperCase().replace(/\s+/g, ' ').trim() : "";
  const deliveries = [];

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowId = String(row[0] || "").trim();
    if (!rowId) continue;

    const rowDriver = String(row[3] || "").trim();
    const rowDriverNorm = rowDriver.toUpperCase().replace(/\s+/g, ' ');

    // Filter by chauffeur if role is chauffeur
    if (role === 'chauffeur' && normalizedDriver) {
      if (!rowDriverNorm.includes(normalizedDriver) && !normalizedDriver.includes(rowDriverNorm)) {
        continue;
      }
    }

    deliveries.push({
      id: rowId,
      date: String(row[1] || ""),
      destination: String(row[2] || ""),
      driver: rowDriver,
      vehicle: String(row[3] || ""),
      quantity: Number(row[5]) || 0,
      status: String(row[6] || "Non Livré"),
      lastUpdated: row[7] ? new Date(row[7]).toISOString() : null
    });
  }

  return deliveries;
}

/**
 * Handles HTTP GET requests to provide health checks and schema verification.
 * Also supports read-only GET queries for deliveries (?action=get_deliveries&role=admin).
 * 
 * @param {Object} e - Event object
 * @return {ContentService.TextOutput} JSON health status or query response
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
        total: deliveries.length,
        deliveries: deliveries
      })).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: "online",
      service: "Tradifoot Logistics Delivery Web App API (RBAC Enabled)",
      schema: ["ID", "Date", "Magasin", "Chauffeur", "Matricule", "Quantite", "Status", "Date Mise à Jour"],
      supportedActions: ["get_deliveries", "update_status"],
      timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
