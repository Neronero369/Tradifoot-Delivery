/**
 * =========================================================================
 * TRADIFOOT LOGISTICS — PROGRAMME DE LIVRAISON (WEB APP BACKEND)
 * =========================================================================
 * 
 * Serverless Write-Only Google Apps Script Web App receiver for real-time 
 * delivery status synchronization between the client-side dashboard and 
 * the operational Google Sheet.
 * 
 * Active Deployment URL:
 * https://script.google.com/macros/s/AKfycbwM5y2SPfnT1lMg4_ingmt_ZWzfHFb7l-qTwTow3-eThQ52Rf1FfLRrcBertuvxUY6d/exec
 * 
 * Expected Column Schema (8 Columns):
 * [ID, Date, Magasin, Chauffeur, Matricule, Quantite, Status, Timestamp]
 * =========================================================================
 */

const SHEET_NAME = "Feuille 1";

/**
 * Handles HTTP POST requests sent by the Delivery Calendar Dashboard.
 * Parses JSON payload and updates or appends delivery records in Google Sheet.
 * 
 * @param {Object} e - Event object containing postData.contents
 * @return {ContentService.TextOutput} JSON response confirming mutation
 */
function doPost(e) {
  try {
    // Acquire a mutex lock for up to 10 seconds to prevent race conditions during concurrent toggles
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.getActiveSheet();

    // Parse incoming JSON payload
    let data = {};
    if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      data = e.parameter;
    }

    const id = String(data.id || "").trim();
    const date = String(data.date || "").trim();
    const magasin = String(data.magasin || "").trim();
    const chauffeur = String(data.chauffeur || "").trim();
    const matricule = String(data.matricule || "").trim();
    const quantite = data.quantite !== undefined ? data.quantite : "";
    const status = String(data.status || "Livré").trim();
    const timestamp = new Date();

    // Initialize Header Row if sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["ID", "Date", "Magasin", "Chauffeur", "Matricule", "Quantite", "Status", "Date Mise à Jour"]);
    }

    let updated = false;
    const lastRow = sheet.getLastRow();

    // Check if record with matching ID exists
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

    // Append new row if not found or no ID match
    if (!updated) {
      sheet.appendRow([id, date, magasin, chauffeur, matricule, quantite, status, timestamp]);
    }

    // Release lock
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
 * Handles HTTP GET requests to provide health checks and schema verification.
 * 
 * @param {Object} e - Event object
 * @return {ContentService.TextOutput} JSON health status and API metadata
 */
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: "online",
    service: "Tradifoot Logistics Delivery Web App API",
    schema: ["ID", "Date", "Magasin", "Chauffeur", "Matricule", "Quantite", "Status", "Date Mise à Jour"],
    mode: "write-only receiver",
    timestamp: new Date().toISOString()
  })).setMimeType(ContentService.MimeType.JSON);
}
