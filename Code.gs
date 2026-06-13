// ═══════════════════════════════════════════════════════════════════════════
// G.M. ENTERPRISES — Google Apps Script (SECURED VERSION)
// ═══════════════════════════════════════════════════════════════════════════
//
// HOW TO SET UP (do this once before deploying):
//
// 1. In Apps Script editor → Project Settings (gear icon) → Script Properties
//    Add these properties:
//
//    API_SECRET       →  any long random string, e.g. "gme_k9x2mPqR7vLtN4wZ"
//    USER_admin       →  your admin password, e.g. "Admin@2024"
//    USER_manager     →  another user's password  (add as many as needed)
//
//    USER_ prefix + username = the key. Value = the password.
//    Example: username "ravi" → key "USER_ravi", value "Ravi@pass123"
//
// 2. Deploy → New Deployment → Web App
//    Execute as: Me
//    Who has access: Anyone
//    Click Deploy and copy the URL into your HTML files.
//
// 3. Every time you change this script, click "Deploy → Manage Deployments
//    → Edit (pencil) → Version: New version → Deploy" to apply changes.
//
// ═══════════════════════════════════════════════════════════════════════════

// ── Persistent Sessions using Google CacheService ────────────────────────────
const SESSION_TTL_SEC = 14400; // 4 hours in seconds

// ── Sheet names ──────────────────────────────────────────────────────────────
const SHEET_INVOICES = 'Invoices';
const SHEET_CHALLANS = 'Challans';
const SHEET_BUYERS   = 'Buyers';

// ── Helpers ──────────────────────────────────────────────────────────────────
function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getProps() {
  return PropertiesService.getScriptProperties().getProperties();
}

function generateToken() {
  return Utilities.getUuid().replace(/-/g, '') + Date.now().toString(36);
}

// Validate that a request carries a live session token from Cache.
function validateSession(token) {
  if (!token) return null;
  const cache = CacheService.getScriptCache();
  const sessionStr = cache.get(token);
  
  if (!sessionStr) return null; // Token not found or expired
  
  // Slide the expiry on activity
  cache.put(token, sessionStr, SESSION_TTL_SEC);
  return JSON.parse(sessionStr);
}

// Validate the shared API_SECRET sent with every request.
function validateSecret(secret) {
  const expected = getProps()['API_SECRET'] || '';
  return expected !== '' && secret === expected;
}

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name);
}

// ═══════════════════════════════════════════════════════════════════════════
// doGet — handles read actions
// ═══════════════════════════════════════════════════════════════════════════
function doGet(e) {
  const p      = e.parameter || {};
  const action = p.action    || '';
  const secret = p.secret    || '';
  const token  = p.token     || '';

  try {
    // ── Login does NOT need a token, but needs the API secret ──────────────
    if (action === 'checkLogin') {
      if (!validateSecret(secret)) return jsonOut({ success: false, message: 'Unauthorized' });
      return handleLogin(p.username || '', p.password || '');
    }

    // ── All other GET actions require a valid session token ─────────────────
    const session = validateSession(token);
    if (!session) return jsonOut({ success: false, message: 'Session expired. Please log in again.', sessionExpired: true });

    if (action === 'getInvoiceNos')   return handleGetInvoiceNos();
    if (action === 'getInvoiceByNo')  return handleGetInvoiceByNo(p.invoiceNo || '');
    if (action === 'getDashboard')    return handleGetDashboard();
    if (action === 'getChallansData') return handleGetChallansData();
    if (action === 'getChallanNos')   return handleGetChallanNos();
    if (action === 'searchBuyer')     return handleSearchBuyer(p.q || '');
    if (action === 'getBuyers')       return handleGetBuyers();

    return jsonOut({ success: false, message: 'Unknown action' });

  } catch (err) {
    return jsonOut({ success: false, message: err.toString() });
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// doPost — handles write actions
// ═══════════════════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents || '{}');
    const action = body.action || '';
    const secret = body.secret || '';
    const token  = body.token  || '';

    // ── All POST actions require BOTH a valid secret AND a live session ──────
    if (!validateSecret(secret)) return jsonOut({ success: false, message: 'Unauthorized' });

    const session = validateSession(token);
    if (!session) return jsonOut({ success: false, message: 'Session expired. Please log in again.', sessionExpired: true });

    if (action === 'addInvoice') return handleAddInvoice(body.data, session.username);
    if (action === 'addBuyer')   return handleAddBuyer(body.data, session.username);
    if (action === 'saveBuyer')  return handleSaveBuyer(body, session.username);
    if (action === 'saveChallan') return handleSaveChallan(body, session.username);

    return jsonOut({ success: false, message: 'Unknown action' });
  } catch (err) {
    return jsonOut({ success: false, message: err.toString() });
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// AUTH
// ═══════════════════════════════════════════════════════════════════════════
function handleLogin(username, password) {
  username = (username || '').trim().toLowerCase();
  if (!username || !password) {
    return jsonOut({ success: false, message: 'Username and password are required.' });
  }

  const props    = getProps();
  const key      = 'USER_' + username;
  const stored   = props[key];

  if (!stored || stored !== password) {
    return jsonOut({ success: false, message: 'Invalid username or password.' });
  }

  // Create session using CacheService
  const token = generateToken();
  const sessionData = {
    username:    username,
    displayName: props['DISPLAY_' + username] || username
  };
  
  CacheService.getScriptCache().put(token, JSON.stringify(sessionData), SESSION_TTL_SEC);

  return jsonOut({
    success:     true,
    token:       token,
    username:    username,
    displayName: sessionData.displayName
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// INVOICE HANDLERS
// ═══════════════════════════════════════════════════════════════════════════
function handleGetInvoiceNos() {
  const sheet = getSheet(SHEET_INVOICES);
  if (!sheet) return jsonOut({ invoiceNos: [] });

  const data  = sheet.getDataRange().getValues();
  // Column A = Invoice No (skip header row 0)
  const nos = data.slice(1).map(r => r[0]).filter(v => v && v.toString().trim());

  return jsonOut({ invoiceNos: nos });
}

function handleGetInvoiceByNo(invoiceNo) {
  const sheet = getSheet(SHEET_INVOICES);
  if (!sheet) return jsonOut({ invoice: null });

  const data  = sheet.getDataRange().getValues();
  const headers = data[0];
  for (let i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString().trim().toLowerCase() === invoiceNo.toLowerCase()) {
      const inv = {};
      headers.forEach((h, j) => inv[h] = data[i][j]);
      return jsonOut({ invoice: inv });
    }
  }
  return jsonOut({ invoice: null });
}

function handleGetDashboard() {
  const sheet = getSheet(SHEET_INVOICES);
  if (!sheet) return jsonOut({ invoices: [] });

  const data    = sheet.getDataRange().getValues();
  const headers = data[0];

  const invoices = data.slice(1).map(row => {
    const inv = {};
    headers.forEach((h, j) => inv[h] = row[j]);
    return inv;
  });

  return jsonOut({ invoices });
}

function handleAddInvoice(data, username) {
  const sheet = getSheet(SHEET_INVOICES);
  if (!sheet) return jsonOut({ result: 'error', message: 'Invoices sheet not found' });

  const allData = sheet.getDataRange().getValues();

  // Duplicate check
  const exists = allData.slice(1).some(r =>
    (r[0] || '').toString().trim().toLowerCase() === (data.invoiceNo || '').toLowerCase()
  );

  if (exists) return jsonOut({ result: 'duplicate' });

  const now = new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' });
  const itemsJson = JSON.stringify(data.items || []);

  sheet.appendRow([
    data.invoiceNo    || '',
    data.date         || '',
    data.challanNo    || '',
    data.buyerName    || '',
    data.buyerPhone   || '',
    data.buyerEmail   || '',
    data.buyerAddress || '',
    data.buyerGstin   || '',
    data.subTotal     || 0,
    data.cgstRate     || 0,
    data.cgstAmount   || 0,
    data.sgstRate     || 0,
    data.sgstAmount   || 0,
    data.igstRate     || 0,
    data.igstAmount   || 0,
    data.taxAmount    || 0,
    data.grandTotal   || 0,
    itemsJson,
    username,
    now
  ]);

  return jsonOut({ result: 'success' });
}

// ═══════════════════════════════════════════════════════════════════════════
// BUYER HANDLERS
// ═══════════════════════════════════════════════════════════════════════════
function handleSearchBuyer(q) {
  const sheet = getSheet(SHEET_BUYERS);
  if (!sheet) return jsonOut({ buyers: [] });
  const data  = sheet.getDataRange().getValues();
  if (data.length < 2) return jsonOut({ buyers: [] });
  const headers = data[0];
  const ql = q.toLowerCase();

  const buyers = data.slice(1)
    .filter(r => (r[0] || '').toString().toLowerCase().includes(ql))
    .map(row => {
      const b = {};
      headers.forEach((h, j) => b[h] = row[j]);
      return b;
    });

  return jsonOut({ buyers });
}

function handleGetBuyers() {
  const sheet = getSheet(SHEET_BUYERS);
  if (!sheet) return jsonOut({ buyers: [] });

  const data  = sheet.getDataRange().getValues();
  if (data.length < 2) return jsonOut({ buyers: [] });
  const headers = data[0];

  const buyers = data.slice(1).map(row => {
    const b = {};
    headers.forEach((h, j) => b[h] = row[j]);
    return b;
  });

  return jsonOut({ buyers });
}

function handleAddBuyer(data, username) {
  const sheet = getSheet(SHEET_BUYERS);
  if (!sheet) return jsonOut({ result: 'error', message: 'Buyers sheet not found' });

  const allData = sheet.getDataRange().getValues();
  const nameLC  = (data.name || '').toLowerCase();
  const rowIdx  = allData.slice(1).findIndex(r => (r[0] || '').toString().toLowerCase() === nameLC);

  const now = new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' });
  const row = [
    data.name      || '',
    data.address   || '',
    data.gstin     || '',
    data.state     || '',
    data.stateCode || '',
    data.phone     || '',
    data.email     || '',
    username,
    now
  ];

  if (rowIdx >= 0) {
    const sheetRow = rowIdx + 2; // +1 for header, +1 for 1-based
    sheet.getRange(sheetRow, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }

  return jsonOut({ result: 'success' });
}

function handleSaveBuyer(body, username) {
  return handleAddBuyer({
    name:      body.name      || '',
    address:   body.address   || '',
    gstin:     body.gstin     || '',
    state:     body.state     || '',
    stateCode: body.stateCode || '',
    phone:     body.phone     || '',
    email:     body.email     || ''
  }, username);
}

// ═══════════════════════════════════════════════════════════════════════════
// CHALLAN HANDLERS
// ═══════════════════════════════════════════════════════════════════════════
function handleGetChallanNos() {
  const sheet = getSheet(SHEET_CHALLANS);
  if (!sheet) return jsonOut({ challanNos: [] });
  const data = sheet.getDataRange().getValues();

  const nos  = data.slice(1).map(r => r[0]).filter(v => v && v.toString().trim());
  return jsonOut({ challanNos: nos });
}

function handleGetChallansData() {
  const sheet = getSheet(SHEET_CHALLANS);
  if (!sheet) return jsonOut({ challans: [] });

  const data    = sheet.getDataRange().getValues();
  const headers = data[0];

  const challans = data.slice(1).map(row => {
    const c = {};
    headers.forEach((h, j) => c[h] = row[j]);
    return c;
  });

  return jsonOut({ challans });
}

function handleSaveChallan(body, username) {
  const sheet = getSheet(SHEET_CHALLANS);
  if (!sheet) return jsonOut({ success: false, message: 'Challans sheet not found' });

  const allData = sheet.getDataRange().getValues();

  const challanNo = (body.id || body.challanNo || '').toString().trim();
  const exists = allData.slice(1).some(r =>
    (r[0] || '').toString().trim().toLowerCase() === challanNo.toLowerCase()
  );

  if (exists) return jsonOut({ success: false, message: 'Duplicate challan number' });

  const now = new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' });

  sheet.appendRow([
    challanNo,
    body.date         || '',
    body.challanNo    || body.challanNo || '',
    body.buyerName    || '',
    body.buyerPhone   || '',
    body.buyerEmail   || '',
    body.buyerAddress || '',
    body.buyerGstin   || '',
    body.subTotal     || 0,
    body.tax          || 0,
    body.grandTotal   || 0,
    body.itemsJson    || '',
    username,
    now
  ]);

  return jsonOut({ success: true });
}