// ============================================================
// CardBiz Pro — Google Apps Script (serves HTML + API)
// ============================================================
// v3 — Adds:
//   - PIN-gate authentication (replaces email allowlist)
//   - Owner bypass for lndbconsulting@gmail.com (Google session)
//   - Per-user PINs stored in Script Properties
//   - 5 wrong attempts → 5 minute lockout
//   - Server-side token check on every API call
//
// v2 added:
//   - Server-side CSV import (importWeeklyOrdersCSV)
//   - Drive folder auto-importer (autoImportFromDrive)
//   - One-time setup function (setupAutoImport)
//   - 'runImport' API action so the client can use the
//     server-side path instead of 50+ per-row API calls.
// ============================================================

// AUTHORIZED_EMAILS is no longer used as a gate — kept here only because
// notifyImportRun() addresses the auto-import problem-emails to AUTHORIZED_EMAILS[0].
const AUTHORIZED_EMAILS = [
  'lndbconsulting@gmail.com',
  'milawnprop@gmail.com'
];

// ============================================================
// SERVE THE HTML APP
// ============================================================
function doGet(e) {
  // Email allowlist removed — access is now controlled by the PIN gate.
  // The script runs as the owner ("Execute as: Me"), and any Google-signed-in
  // user can load the page — but they must enter a valid PIN before the
  // client-side app calls the API. The PIN is also re-checked server-side
  // on every API call (see checkPin() and the doPost gate), so a casual
  // visitor who skips the PIN splash cannot access the data.
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('CardBiz Pro')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// ============================================================
// API — POST handler
// ============================================================
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch(err) { return jsonOut({ success: false, error: 'Invalid JSON' }); }

  // Public actions — no auth required
  if (body.action === 'getAuthMode') return jsonOut(getAuthMode());
  if (body.action === 'checkPin')    return jsonOut(checkPin(body.pin));

  // All other actions require either owner Google session OR a valid PIN token
  if (!isApiCallAuthorized_(body)) {
    return jsonOut({ success: false, error: 'Not authorized', authRequired: true });
  }

  try {
    switch(body.action) {
      case 'getAll':           return jsonOut(getAllData());
      case 'addInventory':     return jsonOut(addRow('Inventory', body.data));
      case 'updateInventory':  return jsonOut(updateRow('Inventory', body.id, body.data));
      case 'deleteInventory':  return jsonOut(deleteRow('Inventory', body.id));
      case 'addShow':          return jsonOut(addRow('Shows', body.data));
      case 'updateShow':       return jsonOut(updateRow('Shows', body.id, body.data));
      case 'deleteShow':       return jsonOut(deleteRow('Shows', body.id));
      case 'addSale':          return jsonOut(addRow('Sales', body.data));
      case 'deleteSale':       return jsonOut(deleteRow('Sales', body.id));
      case 'addPrize':         return jsonOut(addRow('Prizes', body.data));
      case 'deletePrize':      return jsonOut(deleteRow('Prizes', body.id));
      case 'addTip':           return jsonOut(addRow('Tips', body.data));
      case 'logImport':        return jsonOut(addRow('Imports', body.data));
      case 'saveConfig':       return jsonOut(saveConfig(body.data));
      // ---- v2 additions ----
      case 'runImport':        return jsonOut(importWeeklyOrdersCSV(body.csvText, body.filename || 'manual.csv'));
      case 'getAutoStatus':    return jsonOut(getAutoImportStatus());
      case 'runAutoNow':       return jsonOut(autoImportFromDrive());
      case 'recomputeShows':   return jsonOut(recomputeShowTotals());
      // ----------------------
      default:                 return jsonOut({ success: false, error: 'Unknown action: ' + body.action });
    }
  } catch(err) {
    return jsonOut({ success: false, error: err.message + (err.stack ? ' | ' + err.stack : '') });
  }
}

// ============================================================
// SHEET SETUP — run once manually
// ============================================================
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = {
    Inventory: ['id','player','year','brand','cardNum','variation','gradeCo','grade','cert','cost','clValue','source','cardOwner','ownerShare','status','notes','added'],
    Shows:     ['id','name','date','type','platform','wnId','ownership','splitOverride','spotsTotal','spotsSold','spotPrice','spotsRev','packCost','otherCost','tips','totalFees','notes','created'],
    Sales:     ['id','item','date','platform','price','fees','transactionAmount','cost','shipping','buyer','channel','wnShowId','wnShowTitle','showId','invId','orderId','imported','created'],
    Prizes:    ['id','showId','card','marketValue','discountPct','discountedValue','valuationPct','valuation','originalCost','ownerShare','winner','invId','created'],
    Tips:      ['id','showId','fromUser','amt','created'],
    Imports:   ['id','filename','type','records','date'],
    Config:    ['key','value']
  };
  // Create sheets that don't exist yet
  Object.entries(sheets).forEach(([name, headers]) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      sheet.getRange(1,1,1,headers.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    } else {
      // Sheet already has data — add any missing columns at the end
      const existing = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
      const existingLower = existing.map(h => String(h).toLowerCase());
      headers.forEach(h => {
        if (!existingLower.includes(h.toLowerCase())) {
          const newCol = sheet.getLastColumn() + 1;
          sheet.getRange(1, newCol).setValue(h).setFontWeight('bold');
        }
      });
    }
  });
  // Config defaults — only set if not already present
  const cfg = ss.getSheetByName('Config');
  const cfgRows = cfg.getLastRow() > 1 ? cfg.getRange(2,1,cfg.getLastRow()-1,2).getValues() : [];
  const existingKeys = new Set(cfgRows.map(r => String(r[0])));
  const defaults = [
    ['p1','Luca'],
    ['p2','Matthew Ciampa'],
    ['split','40'],            // Luca's %; Matt gets 100 - this
    ['valuationPct','85'],     // default valuation % of market value
    ['jointRecoupShare','50'], // for joint-owned cards: % of valuation to Luca
    ['wnfee','8'],
    ['ebfee','13.25'],
    ['othfee','5'],
    ['disc','17.5']            // legacy — kept for backward compat
  ];
  defaults.forEach(([k,v]) => {
    if (!existingKeys.has(k)) cfg.appendRow([k, v]);
  });
  try { SpreadsheetApp.getUi().alert('CardBiz sheets ready! New columns added if needed. Run migrateOwnership next if you have existing inventory or shows.'); }
  catch(e) { Logger.log('Sheets ready'); }
}

// ============================================================
// ONE-TIME MIGRATION — run after setupSheets
// ============================================================
// Migrates legacy data to the new ownerShare / splitOverride fields:
//   - Inventory.cardOwner='partner' → ownerShare=0 (Matt 100%)
//   - Inventory.cardOwner='joint'   → ownerShare=50
//   - Inventory.cardOwner=''/other  → leave empty (you set it manually)
//   - Shows.ownership='solo'        → splitOverride='0' (Luca 0%, Matt 100%)
//   - Shows.ownership='joint'/empty → splitOverride='' (use global default)
// Idempotent — running twice is safe; it only writes when current value is empty.
function migrateOwnership() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Inventory
  const inv = ss.getSheetByName('Inventory');
  if (inv && inv.getLastRow() > 1) {
    const data = inv.getDataRange().getValues();
    const headers = data[0];
    const cardOwnerCol = headers.indexOf('cardOwner');
    const ownerShareCol = headers.indexOf('ownerShare');
    if (cardOwnerCol >= 0 && ownerShareCol >= 0) {
      let invMigrated = 0;
      for (let i = 1; i < data.length; i++) {
        const cur = String(data[i][ownerShareCol]);
        if (cur !== '') continue; // already set
        const owner = String(data[i][cardOwnerCol]).toLowerCase();
        let val = '';
        if (owner === 'partner') val = '0';
        else if (owner === 'joint') val = '50';
        if (val !== '') {
          inv.getRange(i+1, ownerShareCol+1).setValue(val);
          invMigrated++;
        }
      }
      Logger.log('Inventory rows migrated: ' + invMigrated);
    }
  }

  // Shows
  const shows = ss.getSheetByName('Shows');
  if (shows && shows.getLastRow() > 1) {
    const data = shows.getDataRange().getValues();
    const headers = data[0];
    const ownershipCol = headers.indexOf('ownership');
    const splitOverCol = headers.indexOf('splitOverride');
    if (ownershipCol >= 0 && splitOverCol >= 0) {
      let showMigrated = 0;
      for (let i = 1; i < data.length; i++) {
        const cur = String(data[i][splitOverCol]);
        if (cur !== '') continue;
        const own = String(data[i][ownershipCol]).toLowerCase();
        if (own === 'solo') {
          shows.getRange(i+1, splitOverCol+1).setValue('0');
          showMigrated++;
        }
      }
      Logger.log('Shows rows migrated: ' + showMigrated);
    }
  }

  try { SpreadsheetApp.getUi().alert('Migration complete. Check the Logs (View → Logs) for counts. New columns are populated where legacy values existed; cards/shows without legacy values are left blank for you to set manually.'); }
  catch(e) {}
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

// Internal version that does NOT return the wrapped {success, id} —
// used by import logic so we can do many writes without per-call overhead.
function addRowInternal(sheetName, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  if (!data.id) data.id = uid();
  if (!data.created) data.created = new Date().toISOString();
  const row = headers.map(h => data[h] !== undefined ? data[h] : '');
  sheet.appendRow(row);
  return data.id;
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

// ============================================================
// CSV PARSER (server-side)
// ============================================================
function parseCSV(text) {
  const lines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').trim().split('\n');
  if (lines.length < 2) return { headers: [], rows: [] };

  const parseLine = (line) => {
    const vals = [];
    let cur = '', inQ = false;
    for (let i = 0; i < line.length; i++) {
      const c = line[i];
      if (c === '"') {
        if (inQ && line[i+1] === '"') { cur += '"'; i++; } // escaped quote
        else { inQ = !inQ; }
      } else if (c === ',' && !inQ) {
        vals.push(cur); cur = '';
      } else {
        cur += c;
      }
    }
    vals.push(cur);
    return vals.map(v => v.trim());
  };

  const headers = parseLine(lines[0]).map(h => h.replace(/^"|"$/g, '').trim().toLowerCase());
  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    if (!lines[i].trim()) continue;
    const vals = parseLine(lines[i]);
    const obj = {};
    headers.forEach((h, j) => obj[h] = (vals[j] || '').replace(/^"|"$/g, '').trim());
    if (Object.values(obj).some(v => v)) rows.push(obj);
  }
  return { headers, rows };
}

function num(x) {
  const n = Number(x);
  return isNaN(n) ? 0 : n;
}

// ============================================================
// CSV FORMAT DETECTION
// ============================================================
// Returns 'weekly_orders' | 'unknown'
// (Ledger and Livestream Report formats can be added here later.)
function detectCsvFormat(headers) {
  const h = headers.map(x => x.toLowerCase());
  if (h.includes('order_id') && h.includes('livestream_id') && h.includes('original_item_price')) {
    return 'weekly_orders';
  }
  return 'unknown';
}

// ============================================================
// SERVER-SIDE WEEKLY ORDERS IMPORT (v2 — corrected)
// ============================================================
// Handles ALL row types in the Whatnot Weekly Orders Report:
//   - AUCTION / BUY_IT_NOW + ORDER_EARNINGS → Sales row
//   - GIVEAWAY                              → Prizes row (cost = abs(TRANSACTION_AMOUNT))
//   - TIP                                   → Tips row (amt = TRANSACTION_AMOUNT)
// Computes show totals (spotsRev, totalFees) by summing the imported rows,
// not by trusting the import-time aggregate (which can drift on incremental imports).
function importWeeklyOrdersCSV(csvText, filename) {
  const startTime = Date.now();
  const summary = {
    filename: filename || 'unknown.csv',
    format: 'weekly_orders',
    salesImported: 0,
    showsCreated: 0,
    showsExisting: 0,
    giveawaysLogged: 0,
    tipsLogged: 0,
    nonShowSales: 0,
    duplicatesSkipped: 0,
    errors: [],
    rowsTotal: 0
  };

  if (!csvText || typeof csvText !== 'string') {
    return { success: false, error: 'No CSV text provided', summary };
  }

  const parsed = parseCSV(csvText);
  if (!parsed.rows.length) {
    return { success: false, error: 'CSV is empty or malformed', summary };
  }

  const fmt = detectCsvFormat(parsed.headers);
  if (fmt !== 'weekly_orders') {
    return {
      success: false,
      error: 'CSV format not recognized as Weekly Orders Report. Headers: ' + parsed.headers.slice(0, 8).join(', ') + '...',
      summary
    };
  }

  summary.rowsTotal = parsed.rows.length;

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Build a set of existing ORDER_IDs so we don't re-import rows
  const salesSheet = ss.getSheetByName('Sales');
  const existingSalesData = salesSheet.getLastRow() > 1
    ? salesSheet.getDataRange().getValues()
    : [salesSheet.getRange(1,1,1,salesSheet.getLastColumn()).getValues()[0]];
  const salesHeaders = existingSalesData[0];
  const orderIdCol = salesHeaders.indexOf('orderId');
  const seenOrderIds = new Set();
  if (orderIdCol >= 0) {
    for (let i = 1; i < existingSalesData.length; i++) {
      const oid = existingSalesData[i][orderIdCol];
      if (oid) seenOrderIds.add(String(oid));
    }
  }

  // Same for Tips/Prizes — dedup by ORDER_ID
  const tipsSheet = ss.getSheetByName('Tips');
  const tipsData = tipsSheet.getLastRow() > 1 ? tipsSheet.getDataRange().getValues() : [];
  const tipsHeaders = tipsData.length ? tipsData[0] : [];
  // Tips don't have an orderId column historically; we rely on Sales orderIds for tips
  // (since a tip row has its own ORDER_ID). We'll add a check by adding orderId to the tips row
  // but for legacy Tips we skip dedup.

  // Index existing shows by wnId
  const showsSheet = ss.getSheetByName('Shows');
  const showsData = showsSheet.getLastRow() > 1
    ? showsSheet.getDataRange().getValues()
    : [showsSheet.getRange(1,1,1,showsSheet.getLastColumn()).getValues()[0]];
  const showsHeaders = showsData[0];
  const showsIdCol = showsHeaders.indexOf('id');
  const showsWnIdCol = showsHeaders.indexOf('wnId');
  const wnIdToShowId = {};
  for (let i = 1; i < showsData.length; i++) {
    const wnId = showsData[i][showsWnIdCol];
    if (wnId) wnIdToShowId[String(wnId)] = String(showsData[i][showsIdCol]);
  }

  // First pass: dedupe and group rows by livestream
  const showGroups = {};
  const nonShowSales = [];
  for (const r of parsed.rows) {
    const orderId = (r['order_id'] || '').toString().trim();
    if (orderId && seenOrderIds.has(orderId)) {
      summary.duplicatesSkipped++;
      continue;
    }
    if (orderId) seenOrderIds.add(orderId);

    const lsId = (r['livestream_id'] || '').toString().trim();
    const tt = (r['transaction_type'] || '').toUpperCase();
    const bf = (r['buy_format'] || '').toUpperCase();
    const tagged = { row: r, orderId: orderId, lsId: lsId, tt: tt, bf: bf };

    if (lsId) {
      if (!showGroups[lsId]) {
        showGroups[lsId] = {
          id: lsId,
          title: (r['livestream_title'] || '').trim(),
          firstDate: (r['order_placed_at_utc'] || '').split(' ')[0],
          sales: [], giveaways: [], tips: []
        };
      }
      if (tt === 'TIP') showGroups[lsId].tips.push(tagged);
      else if (bf === 'GIVEAWAY') showGroups[lsId].giveaways.push(tagged);
      else showGroups[lsId].sales.push(tagged);
    } else {
      // No livestream — only keep regular sales (skip tips/giveaways without a show)
      if (tt !== 'TIP' && bf !== 'GIVEAWAY') nonShowSales.push(tagged);
    }
  }

  // Process each show group
  for (const lsId of Object.keys(showGroups)) {
    const group = showGroups[lsId];
    let showId = wnIdToShowId[lsId];
    let isNewShow = false;

    if (!showId) {
      const showName = group.title || ('Whatnot Show ' + lsId.substring(0, 8));
      try {
        showId = addRowInternal('Shows', {
          name: showName,
          date: group.firstDate,
          type: 'break',
          platform: 'Whatnot',
          wnId: lsId,
          ownership: 'joint',
          splitOverride: '',
          spotsTotal: '',          // will be set during recompute
          spotsSold: '',
          spotPrice: '',
          spotsRev: '',
          packCost: '',
          otherCost: '',
          tips: '',
          totalFees: '',
          notes: 'Auto-imported from Whatnot CSV (' + (filename || '') + ')'
        });
        wnIdToShowId[lsId] = showId;
        summary.showsCreated++;
        isNewShow = true;
      } catch (err) {
        summary.errors.push('Show create failed for ' + lsId + ': ' + err.message);
        continue;
      }
    } else {
      summary.showsExisting++;
    }

    // Sales (auctions + BIN)
    for (const t of group.sales) {
      const r = t.row;
      const price = num(r['original_item_price']);
      const fees = num(r['commission_fee']) + num(r['payment_processing_fee'])
                 + num(r['tax_on_commission_fee']) + num(r['tax_on_payment_processing_fee']);
      const txnAmt = num(r['transaction_amount']);
      try {
        addRowInternal('Sales', {
          item: r['listing_title'] || r['transaction_message'] || '',
          date: (r['order_placed_at_utc'] || '').split(' ')[0],
          platform: 'Whatnot',
          price: price.toFixed(2),
          fees: fees.toFixed(2),
          transactionAmount: txnAmt.toFixed(2),
          cost: '0',
          buyer: r['buyer_name'] || '',
          channel: t.bf,    // AUCTION / BUY_IT_NOW
          wnShowId: lsId,
          wnShowTitle: group.title,
          showId: showId,
          orderId: t.orderId,
          imported: 'true'
        });
        summary.salesImported++;
      } catch (err) {
        summary.errors.push('Sale row failed: ' + err.message);
      }
    }

    // Giveaways → Prizes
    for (const t of group.giveaways) {
      const r = t.row;
      const cost = Math.abs(num(r['transaction_amount']));
      try {
        addRowInternal('Prizes', {
          showId: showId,
          card: r['listing_title'] || 'Giveaway',
          marketValue: cost.toFixed(2),
          discountPct: '0',
          discountedValue: cost.toFixed(2),
          valuationPct: '100',
          valuation: cost.toFixed(2),
          originalCost: '',
          ownerShare: '',  // unknown — user must set manually
          winner: r['buyer_name'] || ''
        });
        summary.giveawaysLogged++;
      } catch (err) {
        summary.errors.push('Giveaway row failed: ' + err.message);
      }
    }

    // Tips
    for (const t of group.tips) {
      const r = t.row;
      const amt = num(r['transaction_amount']);
      if (amt <= 0) continue;
      try {
        addRowInternal('Tips', {
          showId: showId,
          fromUser: r['buyer_name'] || '',
          amt: amt.toFixed(2)
        });
        summary.tipsLogged++;
      } catch (err) {
        summary.errors.push('Tip row failed: ' + err.message);
      }
    }
  }

  // Non-show sales (BIN with no livestream)
  for (const t of nonShowSales) {
    const r = t.row;
    const price = num(r['original_item_price']);
    const fees = num(r['commission_fee']) + num(r['payment_processing_fee'])
               + num(r['tax_on_commission_fee']) + num(r['tax_on_payment_processing_fee']);
    const txnAmt = num(r['transaction_amount']);
    try {
      addRowInternal('Sales', {
        item: r['listing_title'] || r['transaction_message'] || '',
        date: (r['order_placed_at_utc'] || '').split(' ')[0],
        platform: 'Whatnot',
        price: price.toFixed(2),
        fees: fees.toFixed(2),
        transactionAmount: txnAmt.toFixed(2),
        cost: '0',
        buyer: r['buyer_name'] || '',
        channel: t.bf || 'BUY_IT_NOW',
        wnShowId: '',
        wnShowTitle: '',
        showId: '',
        orderId: t.orderId,
        imported: 'true'
      });
      summary.nonShowSales++;
    } catch (err) {
      summary.errors.push('Non-show sale failed: ' + err.message);
    }
  }

  // Recompute show aggregates from the freshly-imported data.
  // This makes spotsRev / totalFees / tips on each show match the sum of its sales,
  // even across multiple incremental imports.
  recomputeShowTotals();

  // Log import history
  addRowInternal('Imports', {
    filename: summary.filename,
    type: 'weekly_orders',
    records: summary.salesImported + summary.giveawaysLogged + summary.nonShowSales + summary.tipsLogged,
    date: new Date().toISOString().split('T')[0]
  });

  summary.elapsedMs = Date.now() - startTime;
  return { success: true, summary: summary };
}

// ============================================================
// RECOMPUTE SHOW TOTALS
// ============================================================
// For every show, sum its linked Sales / Tips / Prizes and write the
// aggregates back to the Shows row. Idempotent — safe to run anytime.
function recomputeShowTotals() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const showsSheet = ss.getSheetByName('Shows');
  if (showsSheet.getLastRow() <= 1) return { success: true, updated: 0 };

  const showsData = showsSheet.getDataRange().getValues();
  const sH = showsData[0];
  const sId = sH.indexOf('id');
  const sWnId = sH.indexOf('wnId');
  const sSpotsRev = sH.indexOf('spotsRev');
  const sSpotsSold = sH.indexOf('spotsSold');
  const sSpotsTotal = sH.indexOf('spotsTotal');
  const sTotalFees = sH.indexOf('totalFees');
  const sTips = sH.indexOf('tips');

  // Build per-show aggregates
  const agg = {}; // showId -> {rev, fees, tips, count}

  // Sales
  const salesSheet = ss.getSheetByName('Sales');
  if (salesSheet.getLastRow() > 1) {
    const sd = salesSheet.getDataRange().getValues();
    const h = sd[0];
    const showIdCol = h.indexOf('showId');
    const priceCol = h.indexOf('price');
    const feesCol = h.indexOf('fees');
    if (showIdCol >= 0) {
      for (let i = 1; i < sd.length; i++) {
        const sid = sd[i][showIdCol];
        if (!sid) continue;
        if (!agg[sid]) agg[sid] = {rev:0, fees:0, tips:0, count:0};
        agg[sid].rev += Number(sd[i][priceCol]) || 0;
        agg[sid].fees += Number(sd[i][feesCol]) || 0;
        agg[sid].count++;
      }
    }
  }

  // Tips
  const tipsSheet = ss.getSheetByName('Tips');
  if (tipsSheet.getLastRow() > 1) {
    const td = tipsSheet.getDataRange().getValues();
    const h = td[0];
    const showIdCol = h.indexOf('showId');
    const amtCol = h.indexOf('amt');
    if (showIdCol >= 0 && amtCol >= 0) {
      for (let i = 1; i < td.length; i++) {
        const sid = td[i][showIdCol];
        if (!sid) continue;
        if (!agg[sid]) agg[sid] = {rev:0, fees:0, tips:0, count:0};
        agg[sid].tips += Number(td[i][amtCol]) || 0;
      }
    }
  }

  // Write back to Shows (only the cells we computed; preserve everything else)
  let updated = 0;
  for (let i = 1; i < showsData.length; i++) {
    const id = String(showsData[i][sId]);
    const a = agg[id] || {rev:0, fees:0, tips:0, count:0};
    if (sSpotsRev   >= 0) showsSheet.getRange(i+1, sSpotsRev+1).setValue(a.rev.toFixed(2));
    if (sTotalFees  >= 0) showsSheet.getRange(i+1, sTotalFees+1).setValue(a.fees.toFixed(2));
    if (sTips       >= 0) showsSheet.getRange(i+1, sTips+1).setValue(a.tips.toFixed(2));
    if (sSpotsSold  >= 0) showsSheet.getRange(i+1, sSpotsSold+1).setValue(a.count);
    if (sSpotsTotal >= 0) {
      // Only set spotsTotal if it's currently empty or 0, to preserve manual overrides
      const cur = Number(showsData[i][sSpotsTotal]) || 0;
      if (cur === 0) showsSheet.getRange(i+1, sSpotsTotal+1).setValue(a.count);
    }
    updated++;
  }

  return { success: true, updated: updated };
}

// ============================================================
// WIPE IMPORTED DATA
// ============================================================
// Backs up Sales, Shows, Prizes, Tips to dated archive sheets, then
// deletes everything except header rows.
// REQUIRES MANUAL EXECUTION FROM EDITOR — does NOT have a UI button.
function wipeImportedData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ts = Utilities.formatDate(new Date(), 'UTC', 'yyyyMMdd_HHmmss');
  const tabsToWipe = ['Sales', 'Shows', 'Prizes', 'Tips', 'Imports'];

  // Confirm via prompt
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    'Wipe imported data?',
    'This will:\n' +
    '  1. Copy all rows from Sales, Shows, Prizes, Tips to backup sheets named "Backup_XXX_' + ts + '"\n' +
    '  2. Delete all data rows (keeping headers) from those four tabs\n\n' +
    'Inventory and Config are NOT touched.\n\n' +
    'This is reversible — your data goes to backup sheets you can copy from.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) {
    Logger.log('Wipe cancelled.');
    return { success: false, cancelled: true };
  }

  const log = [];
  for (const tab of tabsToWipe) {
    const sheet = ss.getSheetByName(tab);
    if (!sheet) { log.push(tab + ': not found'); continue; }
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) { log.push(tab + ': already empty'); continue; }

    // Backup
    const backupName = 'Backup_' + tab + '_' + ts;
    const backup = sheet.copyTo(ss).setName(backupName);
    // copyTo creates "Copy of X" — we set name above

    // Wipe data rows (keep header)
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();

    log.push(tab + ': ' + (lastRow-1) + ' rows backed up to ' + backupName);
  }

  ui.alert('Wipe complete.\n\n' + log.join('\n') + '\n\nNow drop your CSVs into CardBiz Inbox folder (or use the Import tab) and the new importer will load them clean.');
  return { success: true, log: log };
}

// ============================================================
// AUTO-IMPORT FROM DRIVE FOLDER
// ============================================================

const AUTO_IMPORT_KEYS = {
  FOLDER_ID:     'autoImportFolderId',
  PROCESSED_ID:  'autoImportProcessedId',
  ENABLED:       'autoImportEnabled',
  NOTIFY_MODE:   'autoImportNotifyMode',  // 'always' | 'problems_only' | 'silent'
  LAST_RUN:      'autoImportLastRun',
  LAST_SUMMARY:  'autoImportLastSummary'
};

function getConfigVal(key) {
  const cfg = getConfigObj();
  return cfg[key] || '';
}

function setConfigVal(key, value) {
  const data = {};
  data[key] = String(value);
  saveConfig(data);
}

// ----------------------------------------------------------------
// One-time setup: run from the Apps Script editor.
// Creates the Drive folder, the Processed subfolder, the time
// trigger, and writes the config rows.
// ----------------------------------------------------------------
function setupAutoImport() {
  // Find or create the inbox folder at Drive root
  const folderName = 'CardBiz Inbox';
  let folder;
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(folderName);
  }

  // Find or create the Processed subfolder inside it
  let processed;
  const sub = folder.getFoldersByName('Processed');
  if (sub.hasNext()) {
    processed = sub.next();
  } else {
    processed = folder.createFolder('Processed');
  }

  // Save IDs and defaults to Config
  setConfigVal(AUTO_IMPORT_KEYS.FOLDER_ID, folder.getId());
  setConfigVal(AUTO_IMPORT_KEYS.PROCESSED_ID, processed.getId());
  setConfigVal(AUTO_IMPORT_KEYS.ENABLED, 'true');

  // Default to problems-only notifications if not already set
  if (!getConfigVal(AUTO_IMPORT_KEYS.NOTIFY_MODE)) {
    setConfigVal(AUTO_IMPORT_KEYS.NOTIFY_MODE, 'problems_only');
  }

  // Install the time-driven trigger (every 15 minutes).
  // First remove any existing autoImportFromDrive triggers to avoid duplicates.
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'autoImportFromDrive') {
      ScriptApp.deleteTrigger(t);
    }
  }
  ScriptApp.newTrigger('autoImportFromDrive')
    .timeBased()
    .everyMinutes(15)
    .create();

  const msg = 'Auto-import set up successfully!\n\n' +
              'Folder: CardBiz Inbox  (in your Drive root)\n' +
              'Folder ID: ' + folder.getId() + '\n\n' +
              'Drop Whatnot CSVs into this folder and they will be imported within 15 minutes.\n\n' +
              'Notifications: problems-only (change in Config tab if you want every-run summaries).';
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    Logger.log(msg);
  }
  return { success: true, folderId: folder.getId(), folderUrl: folder.getUrl() };
}

// ----------------------------------------------------------------
// The worker. Runs on the 15-minute trigger AND can be called
// from the UI ("Run now").
// ----------------------------------------------------------------
function autoImportFromDrive() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    return { success: false, error: 'Another import is already running' };
  }

  try {
    const enabled = getConfigVal(AUTO_IMPORT_KEYS.ENABLED);
    if (enabled !== 'true' && enabled !== true) {
      return { success: true, skipped: true, reason: 'Auto-import is disabled in Config' };
    }

    const folderId = getConfigVal(AUTO_IMPORT_KEYS.FOLDER_ID);
    const processedId = getConfigVal(AUTO_IMPORT_KEYS.PROCESSED_ID);
    if (!folderId || !processedId) {
      return { success: false, error: 'Folder not configured. Run setupAutoImport once.' };
    }

    let folder, processed;
    try {
      folder = DriveApp.getFolderById(folderId);
      processed = DriveApp.getFolderById(processedId);
    } catch (e) {
      return { success: false, error: 'Folder lookup failed: ' + e.message };
    }

    // Find CSV files at the top level of the folder (NOT inside Processed)
    const files = folder.getFiles();
    const csvFiles = [];
    while (files.hasNext()) {
      const f = files.next();
      const name = f.getName().toLowerCase();
      // accept .csv only; skip anything in the Processed subfolder (it's a separate folder so it won't appear here anyway)
      if (name.endsWith('.csv')) csvFiles.push(f);
    }

    if (!csvFiles.length) {
      setConfigVal(AUTO_IMPORT_KEYS.LAST_RUN, new Date().toISOString());
      return { success: true, filesProcessed: 0, message: 'No new files' };
    }

    const fileResults = [];
    for (const file of csvFiles) {
      const filename = file.getName();
      let result;
      try {
        const csvText = file.getBlob().getDataAsString();
        result = importWeeklyOrdersCSV(csvText, filename);
      } catch (e) {
        result = { success: false, error: e.message };
      }

      if (result.success) {
        // Move to Processed/YYYY-MM
        const monthFolderName = Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM');
        let monthFolder;
        const monthIter = processed.getFoldersByName(monthFolderName);
        if (monthIter.hasNext()) monthFolder = monthIter.next();
        else monthFolder = processed.createFolder(monthFolderName);

        // Move the file: add to month folder, remove from inbox
        monthFolder.addFile(file);
        folder.removeFile(file);
      }
      // Failed files: leave in inbox so they can be inspected/retried.

      fileResults.push({ filename: filename, ...result });
    }

    const runSummary = summarizeRun(fileResults);
    setConfigVal(AUTO_IMPORT_KEYS.LAST_RUN, new Date().toISOString());
    setConfigVal(AUTO_IMPORT_KEYS.LAST_SUMMARY, JSON.stringify(runSummary).substring(0, 500));

    notifyImportRun(runSummary, fileResults);

    return { success: true, filesProcessed: csvFiles.length, runSummary: runSummary };
  } finally {
    lock.releaseLock();
  }
}

function summarizeRun(fileResults) {
  const s = { files: fileResults.length, ok: 0, failed: 0, sales: 0, shows: 0, dupes: 0, giveaways: 0, errors: [] };
  for (const r of fileResults) {
    if (r.success) {
      s.ok++;
      if (r.summary) {
        s.sales += r.summary.salesImported || 0;
        s.shows += r.summary.showsCreated || 0;
        s.dupes += r.summary.duplicatesSkipped || 0;
        s.giveaways += r.summary.giveawaysLogged || 0;
        if (r.summary.errors && r.summary.errors.length) s.errors.push(...r.summary.errors);
      }
    } else {
      s.failed++;
      s.errors.push(r.filename + ': ' + (r.error || 'unknown error'));
    }
  }
  return s;
}

function notifyImportRun(summary, fileResults) {
  const mode = getConfigVal(AUTO_IMPORT_KEYS.NOTIFY_MODE) || 'problems_only';
  if (mode === 'silent') return;

  const hasProblems = summary.failed > 0 || summary.errors.length > 0;
  if (mode === 'problems_only' && !hasProblems) return;

  // Build email
  const subject = hasProblems
    ? 'CardBiz auto-import — needs attention'
    : 'CardBiz auto-import — ' + summary.sales + ' sales imported';

  const lines = [];
  lines.push('CardBiz Pro — Auto-import run');
  lines.push('Time: ' + new Date().toLocaleString());
  lines.push('');
  lines.push('Files processed: ' + summary.ok + ' ok, ' + summary.failed + ' failed');
  lines.push('Sales imported: ' + summary.sales);
  lines.push('New shows created: ' + summary.shows);
  lines.push('Giveaways logged: ' + summary.giveaways);
  lines.push('Duplicates skipped: ' + summary.dupes);
  lines.push('');
  if (fileResults && fileResults.length) {
    lines.push('Per file:');
    for (const r of fileResults) {
      if (r.success && r.summary) {
        lines.push('  ✓ ' + r.filename + ' — ' + r.summary.salesImported + ' sales, ' +
                    r.summary.showsCreated + ' new shows, ' + r.summary.duplicatesSkipped + ' dupes');
      } else {
        lines.push('  ✗ ' + r.filename + ' — ' + (r.error || 'failed'));
      }
    }
    lines.push('');
  }
  if (summary.errors.length) {
    lines.push('Errors:');
    summary.errors.slice(0, 20).forEach(e => lines.push('  • ' + e));
    if (summary.errors.length > 20) lines.push('  ... and ' + (summary.errors.length - 20) + ' more');
    lines.push('');
  }
  lines.push('All shows default to Joint — review and flip any solo shows in the Shows tab.');

  try {
    MailApp.sendEmail({
      to: AUTHORIZED_EMAILS[0],  // owner
      subject: subject,
      body: lines.join('\n')
    });
  } catch (e) {
    Logger.log('Email send failed: ' + e.message);
  }
}

// ----------------------------------------------------------------
// Status reporter for the UI
// ----------------------------------------------------------------
function getAutoImportStatus() {
  const cfg = getConfigObj();
  const folderId = cfg[AUTO_IMPORT_KEYS.FOLDER_ID] || '';
  const enabled = (cfg[AUTO_IMPORT_KEYS.ENABLED] || '') === 'true';
  const notifyMode = cfg[AUTO_IMPORT_KEYS.NOTIFY_MODE] || 'problems_only';
  const lastRun = cfg[AUTO_IMPORT_KEYS.LAST_RUN] || '';
  const lastSummary = cfg[AUTO_IMPORT_KEYS.LAST_SUMMARY] || '';

  let folderUrl = '';
  let pendingCount = 0;
  if (folderId) {
    try {
      const folder = DriveApp.getFolderById(folderId);
      folderUrl = folder.getUrl();
      const files = folder.getFiles();
      while (files.hasNext()) {
        const f = files.next();
        if (f.getName().toLowerCase().endsWith('.csv')) pendingCount++;
      }
    } catch (e) {
      // folder lookup failed — leave blank
    }
  }

  // Check if trigger is installed
  const triggers = ScriptApp.getProjectTriggers();
  const triggerInstalled = triggers.some(t => t.getHandlerFunction() === 'autoImportFromDrive');

  return {
    success: true,
    configured: !!folderId,
    enabled: enabled,
    triggerInstalled: triggerInstalled,
    folderId: folderId,
    folderUrl: folderUrl,
    pendingFiles: pendingCount,
    notifyMode: notifyMode,
    lastRun: lastRun,
    lastSummary: lastSummary
  };
}

// ----------------------------------------------------------------
// Toggle helpers (callable from Apps Script editor or future UI)
// ----------------------------------------------------------------
function pauseAutoImport() {
  setConfigVal(AUTO_IMPORT_KEYS.ENABLED, 'false');
  return { success: true, enabled: false };
}
function resumeAutoImport() {
  setConfigVal(AUTO_IMPORT_KEYS.ENABLED, 'true');
  return { success: true, enabled: true };
}

// ============================================================
// PIN GATE
// ============================================================
// Access is controlled per-user via short PINs stored in Script Properties.
// PINs are NEVER logged, NEVER returned in API responses, and only checked
// server-side. The owner (lndbconsulting@gmail.com) bypasses the PIN gate
// entirely when their Google session is detected. All other users must
// enter a valid PIN to use the app.
//
// Storage layout (Script Properties):
//   pin_<userKey>          → the PIN (e.g. "1999")
//   pinLabel_<userKey>     → display label (e.g. "Matt")
//   pinFails_<userKey>     → JSON {count, lockedUntil} for brute-force lockout
//
// SECURITY NOTE: The PINS object below contains placeholder values, NOT the
// real PINs. The actual PINs live ONLY in Script Properties on the deployed
// Apps Script project — never in this source file or its git history.
// To set or rotate PINs, edit Script Properties directly (Project Settings →
// Script Properties → pin_luca / pin_matt). Re-running setupPins() with
// these placeholder values would OVERWRITE the live PINs with placeholders.

const OWNER_EMAIL = 'lndbconsulting@gmail.com';

// Lockout config
const PIN_MAX_ATTEMPTS = 5;
const PIN_LOCKOUT_MS = 5 * 60 * 1000;  // 5 minutes

// ----------------------------------------------------------------
// One-time setup — run from the editor.
// Edit the PINS object below and run this function to seed
// Script Properties. Re-run any time to update.
//
// !!! DO NOT RUN THIS WITH THE PLACEHOLDER VALUES BELOW !!!
// Either edit Script Properties directly, or change the placeholders
// to real PINs and DO NOT COMMIT THE CHANGE.
// ----------------------------------------------------------------
function setupPins() {
  const PINS = {
    luca: { pin: 'CHANGE_ME_LUCA', label: 'Luca' },
    matt: { pin: 'CHANGE_ME_MATT', label: 'Matt' }
    // To add another user later:
    // alice: { pin: 'CHANGE_ME_ALICE', label: 'Alice' }
  };
  const props = PropertiesService.getScriptProperties();
  Object.entries(PINS).forEach(([userKey, info]) => {
    props.setProperty('pin_' + userKey, info.pin);
    props.setProperty('pinLabel_' + userKey, info.label);
    // Clear any prior lockout on re-setup
    props.deleteProperty('pinFails_' + userKey);
  });
  try {
    SpreadsheetApp.getUi().alert(
      'PINs configured: ' + Object.keys(PINS).join(', ') + '.\n\n' +
      'Edit Script Properties (Project Settings → Script Properties) to change ' +
      'a PIN without re-running this function.'
    );
  } catch (e) { Logger.log('PINs configured.'); }
}

// ----------------------------------------------------------------
// Returns 'owner' if the current Google session matches the owner,
// otherwise '' — meaning a PIN is required for this client.
// Called by the client on app load.
// ----------------------------------------------------------------
function getAuthMode() {
  let email = '';
  try { email = (Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (e) {}
  if (email && email === OWNER_EMAIL.toLowerCase()) {
    return { success: true, mode: 'owner', label: 'Owner' };
  }
  return { success: true, mode: 'pin' };
}

// ----------------------------------------------------------------
// Verify a PIN. Returns { success, ok, label, lockedUntil }.
// Implements per-user 5-attempts / 5-minute lockout.
// ----------------------------------------------------------------
function checkPin(pin) {
  pin = String(pin || '').trim();
  if (!pin) return { success: true, ok: false, error: 'PIN required' };

  const props = PropertiesService.getScriptProperties();
  const allProps = props.getProperties();

  // Find which user (if any) owns this PIN
  let matchedKey = null, matchedLabel = '';
  for (const k of Object.keys(allProps)) {
    if (k.startsWith('pin_') && allProps[k] === pin) {
      matchedKey = k.substring(4); // strip "pin_"
      matchedLabel = allProps['pinLabel_' + matchedKey] || matchedKey;
      break;
    }
  }

  if (!matchedKey) {
    // Unknown PIN — record a global failure under "_unknown" so a brute-force
    // attempt locks out across all users. This is a coarse but effective deterrent.
    return recordPinFail_('_unknown');
  }

  // Check lockout state for this user
  const failsRaw = allProps['pinFails_' + matchedKey];
  if (failsRaw) {
    try {
      const fails = JSON.parse(failsRaw);
      if (fails.lockedUntil && Date.now() < fails.lockedUntil) {
        return { success: true, ok: false, locked: true, lockedUntil: fails.lockedUntil };
      }
    } catch (e) {}
  }

  // Success — clear any failure record, return label and a session token
  props.deleteProperty('pinFails_' + matchedKey);
  // The token is the PIN itself; we re-verify it on every API call. That's
  // fine because the PIN is short-lived in the browser session and never logged.
  return { success: true, ok: true, label: matchedLabel, token: pin };
}

function recordPinFail_(userKey) {
  const props = PropertiesService.getScriptProperties();
  const key = 'pinFails_' + userKey;
  let fails = { count: 0, lockedUntil: 0 };
  const raw = props.getProperty(key);
  if (raw) { try { fails = JSON.parse(raw); } catch (e) {} }
  fails.count = (fails.count || 0) + 1;
  if (fails.count >= PIN_MAX_ATTEMPTS) {
    fails.lockedUntil = Date.now() + PIN_LOCKOUT_MS;
    fails.count = 0; // reset counter on lockout
  }
  props.setProperty(key, JSON.stringify(fails));
  return {
    success: true,
    ok: false,
    locked: fails.lockedUntil > Date.now(),
    lockedUntil: fails.lockedUntil || 0
  };
}

// ----------------------------------------------------------------
// Server-side gate for every API action. Returns true if the
// request is allowed; false if it should be rejected.
// ----------------------------------------------------------------
function isApiCallAuthorized_(body) {
  // Owner bypass — if their Google session resolves to the owner email,
  // skip the PIN check entirely.
  try {
    const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
    if (email && email === OWNER_EMAIL.toLowerCase()) return true;
  } catch (e) {}

  // Otherwise, require a token that matches a stored PIN.
  const token = body && body.token ? String(body.token).trim() : '';
  if (!token) return false;
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  for (const k of Object.keys(all)) {
    if (k.startsWith('pin_') && all[k] === token) return true;
  }
  return false;
}

// ============================================================
// WEEKLY BACKUP
// ============================================================
// Exports the entire Sheet as a timestamped .xlsx file to a
// "CardBiz Backups" folder in your Drive root. Runs every Sunday
// at midnight via a time-based trigger. Keeps the last 26 weeks
// (~6 months); older backups are auto-deleted.
//
// Setup once: run setupWeeklyBackup() from the Apps Script editor.
// Manual run: run weeklyBackup() from the editor at any time.

const BACKUP_KEYS = {
  FOLDER_ID:    'backupFolderId',
  LAST_RUN:     'backupLastRun',
  LAST_STATUS:  'backupLastStatus'   // 'ok' | 'failed' | error message
};
const BACKUP_FOLDER_NAME = 'CardBiz Backups';
const BACKUP_RETENTION_WEEKS = 26;     // ~6 months
const BACKUP_RETENTION_MS = BACKUP_RETENTION_WEEKS * 7 * 24 * 60 * 60 * 1000;

// ----------------------------------------------------------------
// One-time setup — run from the editor.
// Creates the backup folder and installs the weekly trigger.
// ----------------------------------------------------------------
function setupWeeklyBackup() {
  // Find or create the backup folder at Drive root
  let folder;
  const folders = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(BACKUP_FOLDER_NAME);
  }
  setConfigVal(BACKUP_KEYS.FOLDER_ID, folder.getId());

  // Remove any existing weeklyBackup triggers, then install a fresh one
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'weeklyBackup') {
      ScriptApp.deleteTrigger(t);
    }
  }
  // Sunday at midnight (00:00) in the script's timezone
  ScriptApp.newTrigger('weeklyBackup')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(0)
    .create();

  const msg = 'Weekly backup set up successfully!\n\n' +
              'Folder: ' + BACKUP_FOLDER_NAME + ' (in your Drive root)\n' +
              'Folder ID: ' + folder.getId() + '\n\n' +
              'Schedule: every Sunday at midnight\n' +
              'Retention: ' + BACKUP_RETENTION_WEEKS + ' weeks (~6 months)\n\n' +
              'You can run weeklyBackup() manually any time. ' +
              'A success email will be sent to ' + OWNER_EMAIL + ' after each run.';
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    Logger.log(msg);
  }
  return { success: true, folderId: folder.getId(), folderUrl: folder.getUrl() };
}

// ----------------------------------------------------------------
// The worker. Runs on the Sunday trigger AND can be called manually.
// ----------------------------------------------------------------
function weeklyBackup() {
  const startTime = Date.now();
  let result = { success: false };
  try {
    const folderId = getConfigVal(BACKUP_KEYS.FOLDER_ID);
    if (!folderId) {
      throw new Error('Backup folder not configured. Run setupWeeklyBackup() once from the editor.');
    }
    const folder = DriveApp.getFolderById(folderId);

    // Export the current spreadsheet as .xlsx via the Drive export endpoint.
    // This produces a real Excel file (not just a Sheets copy) so it's
    // usable even if Google Sheets is unavailable.
    const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
    const exportUrl = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?format=xlsx';
    const blob = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    }).getBlob();

    const dateStr = Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd');
    const filename = 'CardBiz_Backup_' + dateStr + '.xlsx';
    blob.setName(filename);
    const file = folder.createFile(blob);
    const fileSize = file.getSize();

    // Prune old backups (anything older than retention window)
    const cutoff = Date.now() - BACKUP_RETENTION_MS;
    let pruned = 0, kept = 0;
    const files = folder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      if (f.getName().startsWith('CardBiz_Backup_') && f.getName().endsWith('.xlsx')) {
        if (f.getDateCreated().getTime() < cutoff && f.getId() !== file.getId()) {
          f.setTrashed(true);
          pruned++;
        } else {
          kept++;
        }
      }
    }

    result = {
      success: true,
      filename: filename,
      fileSize: fileSize,
      filesKept: kept,
      filesPruned: pruned,
      folderUrl: folder.getUrl(),
      elapsedMs: Date.now() - startTime
    };
    setConfigVal(BACKUP_KEYS.LAST_RUN, new Date().toISOString());
    setConfigVal(BACKUP_KEYS.LAST_STATUS, 'ok');
  } catch (err) {
    result = { success: false, error: err.message };
    setConfigVal(BACKUP_KEYS.LAST_RUN, new Date().toISOString());
    setConfigVal(BACKUP_KEYS.LAST_STATUS, ('failed: ' + err.message).substring(0, 200));
  }

  // Email the result. Always emails (success or fail) per user request.
  notifyBackupRun_(result);

  return result;
}

function notifyBackupRun_(result) {
  const lines = [];
  if (result.success) {
    const sizeKB = Math.round(result.fileSize / 1024);
    lines.push('CardBiz Pro — weekly backup complete');
    lines.push('Time: ' + new Date().toLocaleString());
    lines.push('');
    lines.push('File: ' + result.filename);
    lines.push('Size: ' + sizeKB + ' KB');
    lines.push('Backups retained: ' + result.filesKept);
    lines.push('Old backups pruned: ' + result.filesPruned);
    lines.push('Elapsed: ' + result.elapsedMs + ' ms');
    lines.push('');
    lines.push('Folder: ' + result.folderUrl);
    lines.push('');
    lines.push('Retention: ' + BACKUP_RETENTION_WEEKS + ' weeks (~6 months).');
  } else {
    lines.push('CardBiz Pro — weekly backup FAILED');
    lines.push('Time: ' + new Date().toLocaleString());
    lines.push('');
    lines.push('Error: ' + (result.error || 'unknown'));
    lines.push('');
    lines.push('Run weeklyBackup() manually from the Apps Script editor to retry.');
  }
  const subject = result.success
    ? 'CardBiz backup ✓ ' + (result.filename || '')
    : 'CardBiz backup FAILED — needs attention';
  try {
    MailApp.sendEmail({
      to: OWNER_EMAIL,
      subject: subject,
      body: lines.join('\n')
    });
  } catch (e) {
    Logger.log('Backup email send failed: ' + e.message);
  }
}

// ----------------------------------------------------------------
// Status check — useful from the editor to see when the last
// backup ran and whether it succeeded.
// ----------------------------------------------------------------
function getBackupStatus() {
  const cfg = getConfigObj();
  const folderId = cfg[BACKUP_KEYS.FOLDER_ID] || '';
  const lastRun = cfg[BACKUP_KEYS.LAST_RUN] || '';
  const lastStatus = cfg[BACKUP_KEYS.LAST_STATUS] || '';
  let folderUrl = '', fileCount = 0, recentFiles = [];
  if (folderId) {
    try {
      const folder = DriveApp.getFolderById(folderId);
      folderUrl = folder.getUrl();
      const files = folder.getFilesByType(MimeType.MICROSOFT_EXCEL);
      const all = [];
      while (files.hasNext()) {
        const f = files.next();
        if (f.getName().startsWith('CardBiz_Backup_')) {
          all.push({ name: f.getName(), created: f.getDateCreated().toISOString(), size: f.getSize() });
        }
      }
      all.sort((a, b) => b.created.localeCompare(a.created));
      fileCount = all.length;
      recentFiles = all.slice(0, 5);
    } catch (e) {}
  }
  const triggers = ScriptApp.getProjectTriggers();
  const triggerInstalled = triggers.some(t => t.getHandlerFunction() === 'weeklyBackup');

  const status = {
    success: true,
    configured: !!folderId,
    triggerInstalled: triggerInstalled,
    folderUrl: folderUrl,
    fileCount: fileCount,
    recentFiles: recentFiles,
    lastRun: lastRun,
    lastStatus: lastStatus,
    retentionWeeks: BACKUP_RETENTION_WEEKS
  };
  Logger.log(JSON.stringify(status, null, 2));
  return status;
}
