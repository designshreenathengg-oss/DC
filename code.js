/*******************************************************
 * DC FLOW - Delivery Challan Tracking Backend
 * Google Apps Script / Code.gs
 *******************************************************/

const APP = {
  VERSION: "2.5.0",
  SHEETS: {
    CHALLANS: "Challans",
    UNITS: "Units",
    SCANS: "ScanLog",
    DELIVERIES: "Deliveries",
    AUDIT: "Audit",
    ITEMS: "Items"
  }
};

function formatDateDisplay_(dateObj) {
  if (!dateObj) return "";
  const d = new Date(dateObj);
  if (isNaN(d.getTime())) return String(dateObj);
  const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const day = String(d.getDate()).padStart(2, '0');
  const month = months[d.getMonth()];
  const year = d.getFullYear();
  return day + "-" + month + "-" + year;
}

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  const callback = e && e.parameter && e.parameter.callback;

  if (action === "getDb") {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const db = loadFullDbFromSheets_(ss);
    
    if (callback) {
      return ContentService.createTextOutput(callback + "(" + JSON.stringify(db) + ")")
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(JSON.stringify(db))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    application: "DC Flow - Delivery Challan Tracking",
    version: APP.VERSION,
    status: "BACKEND ONLINE",
    serverTime: formatDateDisplay_(new Date()) + " " + new Date().toLocaleTimeString()
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    let request = {};
    if (e && e.parameter && e.parameter.payload) {
      request = JSON.parse(decodeURIComponent(e.parameter.payload));
    } else if (e && e.postData && e.postData.contents) {
      request = JSON.parse(e.postData.contents);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    setupSheets_(ss);

    if (request.action === "syncDb" && request.db) {
      saveFullDbToSheets_(ss, request.db);
      return ContentService.createTextOutput(JSON.stringify({ success: true, message: "Full Backup Saved" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ success: false, error: "INVALID_ACTION" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function setupSheets_(ss) {
  getOrCreateSheet_(ss, APP.SHEETS.CHALLANS, ["Challan No", "Date", "Customer", "Address", "PO No", "PO Date", "Project", "Vehicle", "Status"]);
  getOrCreateSheet_(ss, APP.SHEETS.ITEMS, ["Challan No", "Sr No", "Item Code", "Description", "Qty", "Unit", "Remark"]);
  getOrCreateSheet_(ss, APP.SHEETS.UNITS, ["QR Code ID", "Challan No", "Item Code", "Description", "Customer", "Location", "Status", "Dispatch Officer", "Dispatch Time", "Site Supervisor", "Site Time"]);
  getOrCreateSheet_(ss, APP.SHEETS.AUDIT, ["Timestamp", "Challan No", "QR Code ID", "Action", "Actor", "Dispatch Person", "Device / OS Details", "Details"]);
}

function getOrCreateSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  }
  return sh;
}

function saveFullDbToSheets_(ss, db) {
  // 1. Sync Challans
  const cSheet = ss.getSheetByName(APP.SHEETS.CHALLANS);
  cSheet.clearContents();
  cSheet.appendRow(["Challan No", "Date", "Customer", "Address", "PO No", "PO Date", "Project", "Vehicle", "Status"]);
  
  const iSheet = ss.getSheetByName(APP.SHEETS.ITEMS);
  iSheet.clearContents();
  iSheet.appendRow(["Challan No", "Sr No", "Item Code", "Description", "Qty", "Unit", "Remark"]);

  if (db.challans) {
    Object.values(db.challans).forEach(c => {
      cSheet.appendRow([
        c.challanNo,
        formatDateDisplay_(c.date),
        c.customer,
        c.address,
        c.poNo || '',
        formatDateDisplay_(c.poDate),
        c.projectName || '',
        c.vehicle || '',
        c.status
      ]);

      if (c.items && Array.isArray(c.items)) {
        c.items.forEach((it, idx) => {
          iSheet.appendRow([c.challanNo, idx + 1, it.code, it.desc, it.qty, it.unit, it.remark]);
        });
      }
    });
  }

  // 2. Sync Units / QR Codes
  const uSheet = ss.getSheetByName(APP.SHEETS.UNITS);
  uSheet.clearContents();
  uSheet.appendRow(["QR Code ID", "Challan No", "Item Code", "Description", "Customer", "Location", "Status", "Dispatch Officer", "Dispatch Time", "Site Supervisor", "Site Time"]);

  const qrs = db.qrcodes || db.barcodes || {};
  Object.values(qrs).forEach(q => {
    uSheet.appendRow([
      q.qrCode || q.barcode,
      q.challanNo,
      q.itemCode,
      q.desc,
      q.customer,
      q.location,
      q.status,
      q.dispatchScan ? q.dispatchScan.officer : '',
      q.dispatchScan ? formatDateDisplay_(q.dispatchScan.time) : '',
      q.siteScan ? q.siteScan.supervisor : '',
      q.siteScan ? formatDateDisplay_(q.siteScan.time) : ''
    ]);
  });

  // 3. Sync Audit Log
  const aSheet = ss.getSheetByName(APP.SHEETS.AUDIT);
  aSheet.clearContents();
  aSheet.appendRow(["Timestamp", "Challan No", "QR Code ID", "Action", "Actor", "Dispatch Person", "Device / OS Details", "Details"]);

  if (db.audit && Array.isArray(db.audit)) {
    db.audit.forEach(a => {
      aSheet.appendRow([
        formatDateDisplay_(a.timestamp),
        a.challanNo,
        a.qrCode || a.barcode,
        a.action,
        a.actor,
        a.dispatchPerson || '-',
        a.deviceInfo || '',
        a.details
      ]);
    });
  }
}

function loadFullDbFromSheets_(ss) {
  const db = { challans: {}, qrcodes: {}, audit: [], customers: {} };
  
  // Read Challans
  const cSheet = ss.getSheetByName(APP.SHEETS.CHALLANS);
  if (cSheet && cSheet.getLastRow() > 1) {
    const cData = cSheet.getDataRange().getValues();
    for (let i = 1; i < cData.length; i++) {
      const row = cData[i];
      const dcNo = row[0];
      if (dcNo) {
        db.challans[dcNo] = {
          challanNo: dcNo,
          date: row[1],
          customer: row[2],
          address: row[3],
          poNo: row[4],
          poDate: row[5],
          projectName: row[6],
          vehicle: row[7],
          status: row[8],
          items: [],
          qrcodes: []
        };
        if (!db.customers[row[2]]) db.customers[row[2]] = { addresses: [] };
        if (row[3] && !db.customers[row[2]].addresses.includes(row[3])) {
          db.customers[row[2]].addresses.push(row[3]);
        }
      }
    }
  }

  // Read Items
  const iSheet = ss.getSheetByName(APP.SHEETS.ITEMS);
  if (iSheet && iSheet.getLastRow() > 1) {
    const iData = iSheet.getDataRange().getValues();
    for (let i = 1; i < iData.length; i++) {
      const row = iData[i];
      const dcNo = row[0];
      if (dcNo && db.challans[dcNo]) {
        db.challans[dcNo].items.push({
          code: row[2], desc: row[3], qty: row[4], unit: row[5], remark: row[6]
        });
      }
    }
  }

  // Read Units / QR Codes
  const uSheet = ss.getSheetByName(APP.SHEETS.UNITS);
  if (uSheet && uSheet.getLastRow() > 1) {
    const uData = uSheet.getDataRange().getValues();
    for (let i = 1; i < uData.length; i++) {
      const row = uData[i];
      const qrId = row[0];
      const dcNo = row[1];
      if (qrId) {
        db.qrcodes[qrId] = {
          qrCode: qrId,
          challanNo: dcNo,
          itemCode: row[2],
          desc: row[3],
          customer: row[4],
          location: row[5],
          status: row[6],
          dispatchScan: row[7] ? { officer: row[7], time: row[8] } : null,
          siteScan: row[9] ? { supervisor: row[9], time: row[10] } : null
        };
        if (dcNo && db.challans[dcNo]) {
          db.challans[dcNo].qrcodes.push(qrId);
        }
      }
    }
  }

  // Read Audit Logs
  const aSheet = ss.getSheetByName(APP.SHEETS.AUDIT);
  if (aSheet && aSheet.getLastRow() > 1) {
    const aData = aSheet.getDataRange().getValues();
    for (let i = 1; i < aData.length; i++) {
      const row = aData[i];
      db.audit.push({
        timestamp: row[0],
        challanNo: row[1],
        qrCode: row[2],
        action: row[3],
        actor: row[4],
        dispatchPerson: row[5],
        deviceInfo: row[6],
        details: row[7]
      });
    }
  }

  return db;
}