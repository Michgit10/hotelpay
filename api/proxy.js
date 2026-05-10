const crypto = require("crypto");

const SPREADSHEET_ID  = process.env.SPREADSHEET_ID;
const CLAUDE_API_KEY  = process.env.CLAUDE_API_KEY;
const SHEETS_BASE     = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}`;

const SHEET_INVOICES  = "חשבוניות";
const SHEET_SUPPLIERS = "ספקים";
const SHEET_QUOTES    = "הצעות_מחיר";

// ─── JWT / OAuth2 for Service Account ────────────
let _token = null, _tokenExp = 0;

async function getToken() {
  if (_token && Date.now() < _tokenExp) return _token;

  const creds = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT || "{}");
  if (!creds.client_email) throw new Error("GOOGLE_SERVICE_ACCOUNT env var missing or invalid");

  const now  = Math.floor(Date.now() / 1000);
  const enc  = obj => Buffer.from(JSON.stringify(obj)).toString("base64url");
  const hdr  = enc({ alg: "RS256", typ: "JWT" });
  const pay  = enc({
    iss: creds.client_email,
    scope: "https://www.googleapis.com/auth/spreadsheets",
    aud: "https://oauth2.googleapis.com/token",
    iat: now,
    exp: now + 3600,
  });
  const sig_input = `${hdr}.${pay}`;
  const signer = crypto.createSign("RSA-SHA256");
  signer.update(sig_input);
  const sig = signer.sign(creds.private_key, "base64url");

  const res  = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: `grant_type=urn:ietf:params:oauth:grant-type:jwt-bearer&assertion=${sig_input}.${sig}`,
  });
  const data = await res.json();
  if (!data.access_token) throw new Error("Token error: " + JSON.stringify(data));
  _token    = data.access_token;
  _tokenExp = Date.now() + 3500 * 1000;
  return _token;
}

// ─── Sheets helpers ───────────────────────────────
async function sheetsRead(sheet) {
  const token = await getToken();
  const res = await fetch(`${SHEETS_BASE}/values/${encodeURIComponent(sheet)}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  return res.json();
}

async function sheetsAppend(sheet, row) {
  const token = await getToken();
  const res = await fetch(
    `${SHEETS_BASE}/values/${encodeURIComponent(sheet)}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`,
    {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify({ values: [row] }),
    }
  );
  return res.json();
}

async function sheetsPut(range, row) {
  const token = await getToken();
  const res = await fetch(
    `${SHEETS_BASE}/values/${encodeURIComponent(range)}?valueInputOption=RAW`,
    {
      method: "PUT",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify({ values: [row] }),
    }
  );
  return res.json();
}

function toJson(data) {
  const rows = data.values || [];
  if (rows.length < 2) return [];
  const headers = rows[0];
  return rows.slice(1)
    .filter(r => r.some(c => c !== ""))
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = r[i] ?? ""; });
      return obj;
    });
}

function serialize(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "object") return JSON.stringify(v);
  return String(v);
}

// ─── Actions ─────────────────────────────────────
async function readAll() {
  const [inv, sup, quo] = await Promise.all([
    sheetsRead(SHEET_INVOICES),
    sheetsRead(SHEET_SUPPLIERS),
    sheetsRead(SHEET_QUOTES),
  ]);
  return {
    invoices:  toJson(inv),
    suppliers: toJson(sup),
    quotes:    toJson(quo),
  };
}

async function addRow(sheetName, data) {
  const existing = await sheetsRead(sheetName);
  const rows = existing.values || [];
  let headers;
  if (rows.length === 0) {
    headers = Object.keys(data);
    await sheetsAppend(sheetName, headers);
  } else {
    headers = rows[0];
    // add missing columns
    const known = new Set(headers);
    for (const k of Object.keys(data)) {
      if (!known.has(k)) {
        headers.push(k);
        // update header row
        const token = await getToken();
        await fetch(
          `${SHEETS_BASE}/values/${encodeURIComponent(sheetName + "!A1")}?valueInputOption=RAW`,
          {
            method: "PUT",
            headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
            body: JSON.stringify({ values: [headers] }),
          }
        );
      }
    }
  }
  await sheetsAppend(sheetName, headers.map(h => serialize(data[h])));
  return { success: true };
}

async function updateRow(sheetName, id, updates) {
  const existing = await sheetsRead(sheetName);
  const rows = existing.values || [];
  if (rows.length < 2) return { error: "Sheet empty" };
  const headers = rows[0];
  const idCol   = headers.indexOf("id");
  if (idCol === -1) return { error: "No id column" };

  const rowIdx = rows.findIndex((r, i) => i > 0 && String(r[idCol]) === String(id));
  if (rowIdx === -1) return { error: "Row not found: " + id };

  const current = [...rows[rowIdx]];
  for (const [key, val] of Object.entries(updates)) {
    const col = headers.indexOf(key);
    if (col !== -1) current[col] = serialize(val);
  }
  // pad to header length
  while (current.length < headers.length) current.push("");
  const range = `${sheetName}!A${rowIdx + 1}:${String.fromCharCode(64 + headers.length)}${rowIdx + 1}`;
  await sheetsPut(range, current);
  return { success: true };
}

async function scanInvoice(imageData, mediaType) {
  if (!CLAUDE_API_KEY) return { error: "CLAUDE_API_KEY env var not set in Vercel" };
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": CLAUDE_API_KEY,
      "anthropic-version": "2023-06-01",
    },
    body: JSON.stringify({
      model: "claude-opus-4-5",
      max_tokens: 1024,
      messages: [{
        role: "user",
        content: [
          { type: "image", source: { type: "base64", media_type: mediaType || "image/jpeg", data: imageData } },
          { type: "text", text: "חלץ נתוני חשבונית מהתמונה. החזר JSON בלבד עם השדות: supplierName, invoiceNumber, invoiceDate (DD/MM/YYYY), totalWithVat (מספר), totalWithoutVat (מספר), items (מערך של {name, qty, unitPrice, total}). אל תוסיף טקסט נוסף מחוץ ל-JSON." }
        ],
      }],
    }),
  });
  const result = await res.json();
  if (result.error) return { error: result.error.message || JSON.stringify(result.error) };
  const text = result.content[0].text.trim().replace(/```json\n?|\n?```/g, "");
  try { return JSON.parse(text); }
  catch { return { error: "AI parse error: " + text.slice(0, 200) }; }
}

// ─── Handler ──────────────────────────────────────
export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();

  if (!SPREADSHEET_ID) return res.status(500).json({ error: "SPREADSHEET_ID env var not set in Vercel" });

  try {
    // parse body (sent as text/plain from frontend)
    const body = typeof req.body === "string" ? JSON.parse(req.body) : (req.body || {});

    if (req.method === "GET") {
      const action = req.query.action;
      if (action === "readAll") return res.json(await readAll());
      return res.json({ error: "Unknown GET action: " + action });
    }

    if (req.method === "POST") {
      const action = body.action;
      if (action === "add")    return res.json(await addRow(body.sheet, body.data));
      if (action === "update") return res.json(await updateRow(body.sheet, body.id, body.updates));
      if (action === "scan")   return res.json(await scanInvoice(body.imageData, body.mediaType));
      return res.json({ error: "Unknown POST action: " + action });
    }

    res.json({ error: "Method not allowed" });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
}
