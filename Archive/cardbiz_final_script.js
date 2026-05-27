// ============================================================
// CardBiz Pro — Google Apps Script (serves HTML + API)
// ============================================================

const AUTHORIZED_EMAILS = [
  'lndbconsulting@gmail.com',
  'Milawnprop@gmail.com'
];

// ============================================================
// SERVE THE HTML APP
// ============================================================
function doGet(e) {
  const userEmail = Session.getActiveUser().getEmail();
  
  if (!AUTHORIZED_EMAILS.map(x => x.toLowerCase()).includes(userEmail.toLowerCase())) {
    return HtmlService.createHtmlOutput(`
      <html><body style="font-family:sans-serif;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0;background:#f5f5f3;">
      <div style="text-align:center;padding:2rem;background:white;border-radius:12px;border:1px solid #e0dedd;max-width:380px;">
        <div style="font-size:24px;font-weight:700;margin-bottom:.5rem;">Card<span style="color:#e94560;">Biz</span> Pro</div>
        <p style="color:#888;font-size:14px;margin-bottom:1rem;">Access denied.</p>
        <p style="color:#555;font-size:13px;">Your account (<strong>${userEmail}</strong>) is not authorized to access this app.</p>
        <p style="color:#888;font-size:12px;margin-top:1rem;">Contact lndbconsulting@gmail.com to request access.</p>
      </div></body></html>
    `).setTitle('CardBiz Pro — Access Denied');
  }

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('CardBiz Pro')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// ============================================================
// API — POST handler
// ============================================================
function doPost(e) {
  const userEmail = Session.getActiveUser().getEmail();
  if (!AUTHORIZED_EMAILS.map(x => x.toLowerCase()).includes(userEmail.toLowerCase())) {
    return jsonOut({ success: false, error: 'Unauthorized' });
  }

  let body;
  try { body = JSON.parse(e.postData.contents); } 
  catch(err) { return jsonOut({ success: false, error: 'Invalid JSON' }); }

  try {
    switch(body.action) {
      case 'getAll':          return jsonOut(getAllData());
      case 'addInventory':    return jsonOut(addRow('Inventory', body.data));
      case 'updateInventory': return jsonOut(updateRow('Inventory', body.id, body.data));
      case 'deleteInventory': return jsonOut(deleteRow('Inventory', body.id));
      case 'addShow':         return jsonOut(addRow('Shows', body.data));
      case 'updateShow':      return jsonOut(updateRow('Shows', body.id, body.data));
      case 'deleteShow':      return jsonOut(deleteRow('Shows', body.id));
      case 'addSale':         return jsonOut(addRow('Sales', body.data));
      case 'deleteSale':      return jsonOut(deleteRow('Sales', body.id));
      case 'addPrize':        return jsonOut(addRow('Prizes', body.data));
      case 'deletePrize':     return jsonOut(deleteRow('Prizes', body.id));
      case 'addTip':          return jsonOut(addRow('Tips', body.data));
      case 'logImport':       return jsonOut(addRow('Imports', body.data));
      case 'saveConfig':      return jsonOut(saveConfig(body.data));
      default:                return jsonOut({ success: false, error: 'Unknown action: ' + body.action });
    }
  } catch(err) {
    return jsonOut({ success: false, error: err.message });
  }
}

// ============================================================
// SHEET SETUP — run once manually
// ============================================================
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = {
    Inventory: ['id','player','year','brand','cardNum','variation','gradeCo','grade','cert','cost','clValue','source','status','notes','added'],
    Shows:     ['id','name','date','type','platform','wnId','spotsTotal','spotsSold','spotPrice','spotsRev','packCost','otherCost','tips','notes','created'],
    Sales:     ['id','item','date','platform','price','fees','cost','shipping','buyer','channel','wnShowId','wnShowTitle','showId','invId','imported','created'],
    Prizes:    ['id','showId','card','value','kept','created'],
    Tips:      ['id','showId','fromUser','amt','created'],
    Imports:   ['id','filename','type','records','date'],
    Config:    ['key','value']
  };
  Object.entries(sheets).forEach(([name, headers]) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      sheet.getRange(1,1,1,headers.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  });
  const cfg = ss.getSheetByName('Config');
  if (cfg.getLastRow() <= 1) {
    cfg.appendRow(['p1','Partner 1']);
    cfg.appendRow(['p2','Partner 2']);
    cfg.appendRow(['split','60']);
    cfg.appendRow(['wnfee','8']);
    cfg.appendRow(['ebfee','13.25']);
    cfg.appendRow(['othfee','5']);
  }
  SpreadsheetApp.getUi().alert('CardBiz sheets ready!');
}

// ============================================================
// HELPERS
// ============================================================
function jsonOut(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function uid() { return Utilities.getUuid(); }

function sheetToObjects(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h,i) => obj[h] = row[i] === '' ? '' : String(row[i]));
    return obj;
  });
}

function addRow(sheetName, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  if (!data.id) data.id = uid();
  if (!data.created) data.created = new Date().toISOString();
  const row = headers.map(h => data[h] !== undefined ? data[h] : '');
  sheet.appendRow(row);
  return { success: true, id: data.id };
}

function updateRow(sheetName, id, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];
  const idCol = headers.indexOf('id');
  for (let i = 1; i < allData.length; i++) {
    if (String(allData[i][idCol]) === String(id)) {
      headers.forEach((h, ci) => {
        if (data[h] !== undefined) sheet.getRange(i+1, ci+1).setValue(data[h]);
      });
      return { success: true };
    }
  }
  return { success: false, error: 'Row not found' };
}

function deleteRow(sheetName, id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const allData = sheet.getDataRange().getValues();
  const idCol = allData[0].indexOf('id');
  for (let i = 1; i < allData.length; i++) {
    if (String(allData[i][idCol]) === String(id)) {
      sheet.deleteRow(i+1);
      return { success: true };
    }
  }
  return { success: false, error: 'Row not found' };
}

function getAllData() {
  return {
    success: true,
    inventory: sheetToObjects('Inventory'),
    shows:     sheetToObjects('Shows'),
    sales:     sheetToObjects('Sales'),
    prizes:    sheetToObjects('Prizes'),
    tips:      sheetToObjects('Tips'),
    imports:   sheetToObjects('Imports'),
    config:    getConfigObj()
  };
}

function getConfigObj() {
  const rows = sheetToObjects('Config');
  const cfg = {};
  rows.forEach(r => cfg[r.key] = r.value);
  return cfg;
}

function saveConfig(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Config');
  const allData = sheet.getDataRange().getValues();
  Object.entries(data).forEach(([key, value]) => {
    let found = false;
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][0] === key) {
        sheet.getRange(i+1, 2).setValue(value);
        found = true; break;
      }
    }
    if (!found) sheet.appendRow([key, value]);
  });
  return { success: true };
}
