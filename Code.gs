// ════════════════════════════════════════════════
//  HotelPay — Google Apps Script
//  הדבק קוד זה ב-Apps Script של הגיליון שלך
// ════════════════════════════════════════════════

const DRIVE_FOLDER_ID = "1Tg0JyEhEUWHCSCZVPjyzQiKaLDwdZj6h";
const SHEET_INVOICES  = "חשבוניות";
const SHEET_SUPPLIERS = "ספקים";
const SHEET_QUOTES    = "הצעות_מחיר";
const VERCEL_PROXY    = "https://hotelpay-tau.vercel.app/api/proxy";

// ─── HTTP GET ────────────────────────────────────
function doGet(e) {
  try {
    const action = e.parameter.action;
    if (action === "readAll") return respond(readAll());
    return respond({ error: "Unknown action: " + action });
  } catch (err) {
    return respond({ error: err.toString() });
  }
}

// ─── HTTP POST ───────────────────────────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === "add")    return respond(addRow(body.sheet, body.data));
    if (action === "update") return respond(updateRow(body.sheet, body.id, body.updates));
    if (action === "scan")   return respond(scanWithClaude(body.imageData, body.mediaType));

    return respond({ error: "Unknown action: " + action });
  } catch (err) {
    return respond({ error: err.toString() });
  }
}

// ─── RESPOND ─────────────────────────────────────
function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── READ ALL ────────────────────────────────────
function readAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    invoices:  sheetToJson(getOrCreate(ss, SHEET_INVOICES)),
    suppliers: sheetToJson(getOrCreate(ss, SHEET_SUPPLIERS)),
    quotes:    sheetToJson(getOrCreate(ss, SHEET_QUOTES)),
  };
}

// ─── SHEET HELPERS ───────────────────────────────
function getOrCreate(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function sheetToJson(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1)
    .filter(row => row.some(cell => cell !== ""))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
}

function addRow(sheetName, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreate(ss, sheetName);
  const lastRow = sheet.getLastRow();

  if (lastRow === 0) {
    // גיליון חדש — צור כותרות + שורה ראשונה
    const keys = Object.keys(data);
    sheet.appendRow(keys);
    sheet.appendRow(keys.map(k => serializeVal(data[k])));
  } else {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    // הוסף עמודות חסרות
    const existingKeys = new Set(headers.filter(Boolean));
    Object.keys(data).forEach(k => {
      if (!existingKeys.has(k)) {
        const newCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, newCol).setValue(k);
        headers.push(k);
      }
    });
    sheet.appendRow(headers.map(h => serializeVal(data[h])));
  }
  return { success: true };
}

function updateRow(sheetName, id, updates) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { error: "Sheet not found: " + sheetName };

  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol   = headers.indexOf("id");
  if (idCol === -1) return { error: "No id column" };

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(id)) {
      Object.keys(updates).forEach(key => {
        const col = headers.indexOf(key);
        if (col !== -1) sheet.getRange(i + 1, col + 1).setValue(serializeVal(updates[key]));
      });
      return { success: true };
    }
  }
  return { error: "Row not found: " + id };
}

function serializeVal(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "object") return JSON.stringify(v);
  return v;
}

// ─── CLAUDE AI SCAN ──────────────────────────────
function scanWithClaude(imageData, mediaType) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("CLAUDE_API_KEY");
  if (!apiKey) return { error: 'מפתח Claude API לא מוגדר. הגדר CLAUDE_API_KEY ב-Script Properties.' };

  const payload = {
    model: "claude-opus-4-5",
    max_tokens: 1024,
    messages: [{
      role: "user",
      content: [
        {
          type: "image",
          source: { type: "base64", media_type: mediaType || "image/jpeg", data: imageData }
        },
        {
          type: "text",
          text: "חלץ נתוני חשבונית מהתמונה. החזר JSON בלבד עם השדות: supplierName, invoiceNumber, invoiceDate (DD/MM/YYYY), totalWithVat (מספר), totalWithoutVat (מספר), items (מערך של {name, qty, unitPrice, total}). אל תוסיף טקסט נוסף מחוץ ל-JSON."
        }
      ]
    }]
  };

  const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const result = JSON.parse(response.getContentText());
  if (result.error) return { error: result.error.message || JSON.stringify(result.error) };

  try {
    const text = result.content[0].text.trim().replace(/```json\n?|\n?```/g, "");
    return JSON.parse(text);
  } catch (e) {
    return { error: "לא ניתן לפרסר תשובת AI: " + result.content[0].text.slice(0, 200) };
  }
}

// ════════════════════════════════════════════════
//  DRIVE FOLDER MONITORING
//  הגדר טריגר: checkNewDriveInvoices כל 5 דקות
// ════════════════════════════════════════════════
function checkNewDriveInvoices() {
  const props   = PropertiesService.getScriptProperties();
  const processed = JSON.parse(props.getProperty("processedFiles") || "[]");
  const apiKey  = props.getProperty("CLAUDE_API_KEY");

  let folder;
  try {
    folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  } catch (e) {
    Logger.log("Drive folder error: " + e);
    return;
  }

  const files = folder.getFiles();
  const newProcessed = [...processed];

  while (files.hasNext()) {
    const file = files.next();
    const fileId = file.getId();
    if (processed.includes(fileId)) continue;

    Logger.log("עיבוד קובץ חדש: " + file.getName());
    try {
      processInvoiceFile(file, apiKey);
    } catch (e) {
      Logger.log("שגיאה בעיבוד קובץ " + file.getName() + ": " + e);
    }
    newProcessed.push(fileId);
  }

  props.setProperty("processedFiles", JSON.stringify(newProcessed));
}

function processInvoiceFile(file, apiKey) {
  const mimeType = file.getMimeType();
  const supportedTypes = [
    MimeType.PNG, MimeType.JPEG, MimeType.BMP,
    MimeType.PDF, "image/jpg", "image/png", "image/jpeg"
  ];
  const isImage = mimeType === MimeType.PNG || mimeType === MimeType.JPEG || mimeType === "image/jpg" || mimeType === "image/png" || mimeType === "image/jpeg";
  const isPdf   = mimeType === MimeType.PDF;

  const invId  = "_" + Math.random().toString(36).slice(2, 10);
  const today  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const fileDate = Utilities.formatDate(file.getDateCreated(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  let invoice = {
    id:           invId,
    supplierId:   "",
    supplierName: file.getName().replace(/\.[^.]+$/, ""),
    invoiceNo:    "",
    amount:       0,
    date:         fileDate,
    status:       "pending",
    items:        "[]",
    alerts:       "[]",
    notes:        "הועלה אוטומטית מתיקיית Drive · " + file.getName(),
    driveFileId:  file.getId(),
    driveFileUrl: file.getUrl(),
    quoteId:      "",
    createdAt:    new Date().toISOString()
  };

  // סריקה דרך Vercel proxy (שם מוגדר CLAUDE_API_KEY)
  if (isImage || isPdf) {
    try {
      const b64 = Utilities.base64Encode(file.getBlob().getBytes());
      const mt  = isPdf ? "application/pdf" : (mimeType === MimeType.PNG ? "image/png" : "image/jpeg");
      const response = UrlFetchApp.fetch(VERCEL_PROXY, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        payload: JSON.stringify({ action: "scan", imageData: b64, mediaType: mt }),
        muteHttpExceptions: true,
      });
      const scanned = JSON.parse(response.getContentText());
      if (!scanned.error) {
        if (scanned.supplierName)  invoice.supplierName = scanned.supplierName;
        if (scanned.invoiceNumber) invoice.invoiceNo    = scanned.invoiceNumber;
        if (scanned.totalWithVat)  invoice.amount       = Number(scanned.totalWithVat);
        if (scanned.invoiceDate)   invoice.date         = scanned.invoiceDate.split("/").reverse().join("-");
        if (scanned.items)         invoice.items        = JSON.stringify(scanned.items);
        invoice.status = "review";
        Logger.log("סריקת AI הצליחה: " + invoice.supplierName + " · ₪" + invoice.amount);
      } else {
        Logger.log("שגיאת AI: " + scanned.error);
      }
    } catch (e) {
      Logger.log("שגיאה בסריקה: " + e);
    }
  }

  addRow(SHEET_INVOICES, invoice);
  Logger.log("חשבונית נוספה: " + invoice.supplierName);
}

// ─── SETUP TRIGGER ───────────────────────────────
// הרץ פונקציה זו פעם אחת ידנית מתפריט Apps Script
function setupDriveTrigger() {
  // מחק טריגרים קיימים
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "checkNewDriveInvoices") {
      ScriptApp.deleteTrigger(t);
    }
  });

  // צור טריגר חדש — כל 5 דקות
  ScriptApp.newTrigger("checkNewDriveInvoices")
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log("✓ טריגר הוגדר — checkNewDriveInvoices כל 5 דקות");
}

// ─── RESET PROCESSED FILES (לבדיקות) ────────────
function resetProcessedFiles() {
  PropertiesService.getScriptProperties().deleteProperty("processedFiles");
  Logger.log("✓ רשימת קבצים מעובדים אופסה");
}
