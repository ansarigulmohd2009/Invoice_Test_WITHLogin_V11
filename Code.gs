// ═══════════════════════════════════════════════════════════════════════════
// G.M. ENTERPRISES — Google Apps Script (SECURED VERSION V11)
// ═══════════════════════════════════════════════════════════════════════════
//
// SET UP INSTRUCTIONS:
// 1. Open your Google Sheet named "New Database_V11".
// 2. Go to Extensions → Apps Script. Paste this code entirely.
// 3. Click the Gear Icon (Project Settings) on the left sidebar.
// 4. Under "Script Properties", click "Add script property" and add:
//    - API_SECRET   →  Choose any secret word/phrase (must match HTML files).
//    - USER_admin   →  Set your desired password, e.g., admin123
// 5. Click Deploy → New Deployment → Select "Web App".
//    - Execute As: Me
//    - Who has access: Anyone
//
// ═══════════════════════════════════════════════════════════════════════════

const SESSION_TTL_SEC = 14400; // 4 Hours persistent session duration

// Sheet configuration constants
const SHEET_INVOICES = 'Invoices';
const SHEET_CHALLANS = 'Challans';
const SHEET_BUYERS   = 'Buyers';

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

// Session validation with sliding expiration
function validateSession(token) {
  if (!token) return null;
  const cache = CacheService.getScriptCache();
  const sessionStr = cache.get(token);
  if (!sessionStr) return null;
  cache.put(token, sessionStr, SESSION_TTL_SEC); // Refresh TTL
  return JSON.parse(sessionStr);
}

function validateSecret(secret) {
  const expected = getProps()['API_SECRET'] || '';
  return expected !== '' && secret === expected;
}

// Dynamic target sheet accessor with automated database structural initialization
function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === SHEET_INVOICES) {
      sheet.appendRow([
        'Invoice No', 'Date', 'Challan No', 'Buyer Name', 'Buyer Phone', 'Buyer Email', 
        'Buyer Address', 'Buyer GSTIN', 'Sub Total', 'CGST Rate', 'CGST Amount', 
        'SGST Rate', 'SGST Amount', 'IGST Rate', 'IGST Amount', 'Tax Amount', 
        'Grand Total', 'Items (JSON)', 'Saved By', 'Timestamp'
      ]);
    } else if (name === SHEET_CHALLANS) {
      sheet.appendRow([
        'Challan No', 'Date', 'Challan Display No', 'Buyer Name', 'Buyer Phone', 'Buyer Email', 
        'Buyer Address', 'Buyer GSTIN', 'Sub Total', 'Tax Amount', 'Grand Total', 
        'Items (JSON)', 'Saved By', 'Timestamp'
      ]);
    } else if (name === SHEET_BUYERS) {
      sheet.appendRow([
        'Name', 'Address', 'GSTIN', 'State', 'State Code', 'Phone', 'Email', 'Saved By', 'Timestamp'
      ]);
    }
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold').setBackground('#f1f5f9');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── GET ROUTER ──────────────────────────────────────────────────────────────
function doGet(e) {
  const p      = e.parameter || {};
  const action = p.action    || '';
  const secret = p.secret    || '';
  const token  = p.token     || '';

  try {
    if (action === 'checkLogin') {
      if (!validateSecret(secret)) return jsonOut({ success: false, message: 'Unauthorized Request' });
      return handleLogin(p.username || '', p.password || '');
    }

    const session = validateSession(token);
    if (!session) return jsonOut({ success: false, message: 'Session expired. Please log in again.', sessionExpired: true });

    if (action === 'getInvoiceNos')   return handleGetInvoiceNos();
    if (action === 'getInvoiceByNo')  return handleGetInvoiceByNo(p.invoiceNo || '');
    if (action === 'getDashboard')    return handleGetDashboard();
    if (action === 'getChallansData') return handleGetChallansData();
    if (action === 'getChallanNos')   return handleGetChallanNos();
    if (action === 'searchBuyer')     return handleSearchBuyer(p.q || '');
    if (action === 'getBuyers')       return handleGetBuyers();

    return jsonOut({ success: false, message: 'Unknown action parameter' });
  } catch (err) {
    return jsonOut({ success: false, message: err.toString() });
  }
}

// ── POST ROUTER ─────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents || '{}');
    const action = body.action || '';
    const secret = body.secret || '';
    const token  = body.token  || '';

    if (!validateSecret(secret)) return jsonOut({ success: false, message: 'Unauthorized Request' });

    const session = validateSession(token);
    if (!session) return jsonOut({ success: false, message: 'Session expired. Please log in again.', sessionExpired: true });

    if (action === 'addInvoice')  return handleAddInvoice(body.data, session.username);
    if (action === 'addBuyer')    return handleAddBuyer(body.data, session.username);
    if (action === 'saveBuyer')   return handleSaveBuyer(body, session.username);
    if (action === 'saveChallan') return handleSaveChallan(body, session.username);

    return jsonOut({ success: false, message: 'Unknown action parameter' });
  } catch (err) {
    return jsonOut({ success: false, message: err.toString() });
  }
}

// ── LOGIC HANDLERS ──────────────────────────────────────────────────────────
function handleLogin(username, password) {
  username = (username || '').trim().toLowerCase();
  if (!username || !password) {
    return jsonOut({ success: false, message: 'Username and password fields are mandatory.' });
  }

  const props  = getProps();
  const key    = 'USER_' + username;
  const stored = props[key];

  if (!stored || stored !== password) {
    return jsonOut({ success: false, message: 'Invalid credentials provided.' });
  }

  const token = generateToken();
  const sessionData = {
    username: username,
    displayName: props['DISPLAY_' + username] || username.toUpperCase()
  };
  
  CacheService.getScriptCache().put(token, JSON.stringify(sessionData), SESSION_TTL_SEC);

  return jsonOut({
    success: true,
    token: token,
    username: username,
    displayName: sessionData.displayName
  });
}

function handleGetInvoiceNos() {
  const sheet = getSheet(SHEET_INVOICES);
  const data  = sheet.getDataRange().getValues();
  const nos   = data.slice(1).map(r => r[0]).filter(v => v && v.toString().trim());
  return jsonOut({ invoiceNos: nos });
}

function handleGetInvoiceByNo(invoiceNo) {
  const sheet = getSheet(SHEET_INVOICES);
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
  const allData = sheet.getDataRange().getValues();
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

function handleSearchBuyer(q) {
  const sheet = getSheet(SHEET_BUYERS);
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
    const sheetRow = rowIdx + 2; 
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

function handleGetChallanNos() {
  const sheet = getSheet(SHEET_CHALLANS);
  const data = sheet.getDataRange().getValues();
  const nos  = data.slice(1).map(r => r[0]).filter(v => v && v.toString().trim());
  return jsonOut({ challanNos: nos });
}

function handleGetChallansData() {
  const sheet = getSheet(SHEET_CHALLANS);
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
  const allData = sheet.getDataRange().getValues();
  const challanNo = (body.id || body.challanNo || '').toString().trim();
  const exists = allData.slice(1).some(r =>
    (r[0] || '').toString().trim().toLowerCase() === challanNo.toLowerCase()
  );

  if (exists) return jsonOut({ success: false, message: 'Duplicate challan number detected' });

  const now = new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' });

  sheet.appendRow([
    challanNo,
    body.date         || '',
    body.challanNo    || '',
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