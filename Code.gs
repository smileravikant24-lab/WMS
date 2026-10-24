const WAREHOUSES = ["Khera Kalan", "Mukherjee Nagar", "By Truck"];
const SALES_PERSONS = ["Chandresh Sale", "Dilip Sale", "Mr. Mukesh", "Mr. Pawan", "Mr. Pranav", "Mr. Rishabh", "Mr. Sonu"];

function ciEquals(a, b) {
  return (a || '').toString().toLowerCase().replace(/\s+/g, ' ').trim() ===
         (b || '').toString().toLowerCase().replace(/\s+/g, ' ').trim();
}

function ciKey(mill, product, warehouse) {
  return [mill, product, warehouse].map(s => (s||'').toString().toLowerCase().replace(/\s+/g, ' ').trim()).join('||');
}

// ─────────────────────────────────────────────────────────────────────────────
// API HELPER (CORS-SAFE JSON OUTPUT)
// ─────────────────────────────────────────────────────────────────────────────
function jsonOutput(data) {
  var str = typeof data === 'string' ? data : JSON.stringify(data);
  return ContentService.createTextOutput(str)
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────────────────────────────────────
// WEB APP API: doGet (Read Requests from GitHub Pages)
// ─────────────────────────────────────────────────────────────────────────────
function doGet(e) {
  try {
    const p = (e && e.parameter) ? e.parameter : {};
    const action = p.action;

    if (!action) {
      return jsonOutput({ status: 'ok', message: 'Satija Paper WMS API is online.' });
    }

    switch (action) {
      case 'getProductMasterList':
        return jsonOutput(getProductMasterList());

      case 'getDashboardData':
        return jsonOutput(getDashboardData());

      case 'getPartyList':
        return jsonOutput(getPartyList());

      case 'getTodayLog':
        return jsonOutput(getTodayLog(p.sheet || 'SalesLog'));

      case 'getLogsByDateRange':
        return jsonOutput(getLogsByDateRange(p.sheet, p.fromDate, p.toDate));

      case 'searchRecords':
        return jsonOutput(searchRecords(p.sheet, p.party || '', p.product || '', p.fromDate || '', p.toDate || ''));

      case 'getProductsFromSheet':
        return jsonOutput(getProductsFromSheet(p.sheet));

      case 'getPartiesFromSheet':
        return jsonOutput(getPartiesFromSheet(p.sheet));

      case 'getTallyStockReport':
        return jsonOutput(getTallyStockReport(p.fromDate, p.toDate, p.warehouse || 'All', p.mill || '', p.product || ''));

      case 'getDailyStockLedger':
        return jsonOutput(getDailyStockLedger(p.fromDate, p.toDate, p.warehouse || 'All', p.mill || '', p.product || ''));

      case 'runHealthCheck':
        return jsonOutput(runHealthCheck());

      default:
        return jsonOutput({ status: 'error', message: 'Unknown action: ' + action });
    }
  } catch (err) {
    return jsonOutput({ status: 'error', message: err.toString() });
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// WEB APP API: doPost (Save / Update / Delete Requests from GitHub Pages)
// ─────────────────────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    const action = payload.action;
    let result;

    switch (action) {
      case 'STOCK_IN':
      case 'STOCK_OUT':
      case 'TRANSFER':
        result = processData(payload);
        break;

      case 'updateRecord':
        result = updateRecord(payload.sheetName, payload.rowNumber, payload.updatedValues);
        break;

      case 'deleteRecord':
        result = deleteRecord(payload.sheetName, payload.rowNumber);
        break;

      case 'rebuildAllMaster':
        result = { status: 'success', count: _doFullRebuild(SpreadsheetApp.getActiveSpreadsheet()) };
        break;

      case 'rebuildMukherjeeNagar':
        rebuildMukherjeeNagarOnly();
        result = { status: 'success', message: 'Mukherjee Nagar rebuilt' };
        break;

      case 'rebuildKheraKalan':
        rebuildKheraKalanOnly();
        result = { status: 'success', message: 'Khera Kalan rebuilt' };
        break;

      case 'rebuildByTruck':
        rebuildByTruckOnly();
        result = { status: 'success', message: 'By Truck rebuilt' };
        break;

      default:
        if (payload.action) {
          result = processData(payload);
        } else {
          result = { status: 'error', message: 'Action missing in POST payload' };
        }
    }
    return jsonOutput(result);
  } catch (err) {
    return jsonOutput({ status: 'error', message: err.toString() });
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Spreadsheet Menu
// ─────────────────────────────────────────────────────────────────────────────
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('🛠 WMS Tools')
      .addItem('✅ Health Check (Master vs Logs)', 'showHealthCheck')
      .addSeparator()
      .addItem('📍 Rebuild Mukherjee Nagar Only', 'rebuildMukherjeeNagarOnly')
      .addItem('📍 Rebuild Khera Kalan Only', 'rebuildKheraKalanOnly')
      .addItem('📍 Rebuild By Truck Only', 'rebuildByTruckOnly')
      .addSeparator()
      .addItem('🔁 Full Rebuild Master (All Warehouses)', 'rebuildAllMaster')
      .addToUi();
  } catch (e) {}
}

function showHealthCheck() {
  const res = JSON.parse(runHealthCheck());
  let msg = res.message + '\nCombos checked: ' + (res.checked || 0);
  if (res.mismatches && res.mismatches.length > 0) {
    msg += '\n\n' + res.mismatches.slice(0, 20).map(m =>
      `${m.mill} | ${m.product} | ${m.warehouse}: Master=${m.masterQty}, Logs=${m.logQty} (diff ${m.diff})`
    ).join('\n');
    if (res.mismatches.length > 20) msg += `\n...and ${res.mismatches.length - 20} more`;
    msg += '\n\nTip: Run "Full Rebuild Master" to fix.';
  }
  SpreadsheetApp.getUi().alert('WMS Health Check', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

function logAudit(action, sheetName, rowNumber, oldData, newData, notes) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const audit = getOrCreateSheet(ss, 'AuditLog', ['Timestamp', 'Action', 'Sheet', 'Row', 'Old Values', 'New Values', 'Changed Fields', 'Notes']);
    const changedFields = [];
    if (oldData && newData) {
      const allKeys = new Set([...Object.keys(oldData), ...Object.keys(newData)]);
      allKeys.forEach(k => {
        const o = (oldData[k] || '').toString().trim();
        const n = (newData[k] || '').toString().trim();
        if (o !== n) changedFields.push(`${k}: "${o}" → "${n}"`);
      });
    }
    audit.appendRow([new Date(), action, sheetName, rowNumber || '', oldData ? JSON.stringify(oldData) : '', newData ? JSON.stringify(newData) : '', changedFields.join(' | '), notes || '']);
  } catch (e) {
    Logger.log('Audit log error: ' + e.toString());
  }
}

function shiftAuditRowsAfterDelete(ss, sheetName, deletedRow) {
  try {
    const audit = ss.getSheetByName('AuditLog');
    if (!audit || audit.getLastRow() < 2) return;
    const d = audit.getDataRange().getValues();
    const h = d[0].map(x => String(x || '').trim().toLowerCase());
    const sI = h.indexOf('sheet'), rI = h.indexOf('row');
    if (sI === -1 || rI === -1) return;
    const updates = [];
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][sI] || '') !== sheetName) continue;
      const rn = Number(d[i][rI]) || 0;
      if (!rn) continue;
      if (rn === deletedRow) updates.push([i + 1, '']);
      else if (rn > deletedRow) updates.push([i + 1, rn - 1]);
    }
    updates.forEach(([r, v]) => audit.getRange(r, rI + 1).setValue(v));
  } catch (e) {
    Logger.log('shiftAuditRowsAfterDelete error: ' + e.toString());
  }
}

function runHealthCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName('MasterStock');
  if (!master) return JSON.stringify({ status: 'error', message: 'MasterStock not found!' });
  const data = master.getDataRange().getValues();
  const mismatches = [];
  const checked = {};
  for (let i = 1; i < data.length; i++) {
    const mill = (data[i][0] || '').toString().trim(), product = (data[i][1] || '').toString().trim(), warehouse = (data[i][2] || '').toString().trim();
    if (!mill || !product || !warehouse) continue;
    const key = ciKey(mill, product, warehouse);
    if (checked[key]) continue; checked[key] = true;
    const currentQty = Number(data[i][3]) || 0;
    const computed = computeQtyFromLogs(ss, mill, product, warehouse);
    if (Math.abs(currentQty - computed) > 0.001) mismatches.push({ row: i + 1, mill, product, warehouse, masterQty: currentQty, logQty: computed, diff: computed - currentQty });
  }
  return JSON.stringify({ status: mismatches.length === 0 ? 'ok' : 'mismatch', checked: Object.keys(checked).length, mismatches: mismatches, message: mismatches.length === 0 ? '✅ All matches' : `⚠️ ${mismatches.length} mismatches` });
}

function computeQtyFromLogs(ss, mill, product, warehouse) {
  let qty = 0;
  const pLog = ss.getSheetByName(`PurchaseLog_${warehouse}`);
  if (pLog) {
    const d = pLog.getDataRange().getValues(); const h = d[0].map(x => String(x).trim().toLowerCase());
    const mi = h.indexOf('mills') !== -1 ? h.indexOf('mills') : h.indexOf('mill'); const pi = h.indexOf('product'), qi = h.indexOf('qty');
    for (let i = 1; i < d.length; i++) { if (ciEquals(d[i][mi], mill) && ciEquals(d[i][pi], product)) qty += (Number(d[i][qi]) || 0); }
  }
  const sLog = ss.getSheetByName(`SalesLog_${warehouse}`);
  if (sLog) {
    const d = sLog.getDataRange().getValues(); const h = d[0].map(x => String(x).trim().toLowerCase());
    const mi = h.indexOf('mills') !== -1 ? h.indexOf('mills') : h.indexOf('mill'); const pi = h.indexOf('product'), qi = h.indexOf('qty');
    for (let i = 1; i < d.length; i++) { if (ciEquals(d[i][mi], mill) && ciEquals(d[i][pi], product)) qty -= (Number(d[i][qi]) || 0); }
  }
  const tLog = ss.getSheetByName('TransferLog');
  if (tLog) {
    const d = tLog.getDataRange().getValues(); const h = d[0].map(x => String(x).trim().toLowerCase());
    const mi = h.indexOf('mills') !== -1 ? h.indexOf('mills') : h.indexOf('mill'); const pi = h.indexOf('product'), qi = h.indexOf('qty');
    let fi = h.indexOf('from'); if(fi === -1) fi = h.indexOf('from warehouse');
    let ti = h.indexOf('to'); if(ti === -1) ti = h.indexOf('to warehouse');
    for (let i = 1; i < d.length; i++) {
      if (ciEquals(d[i][mi], mill) && ciEquals(d[i][pi], product)) {
        if (ciEquals(d[i][fi], warehouse)) qty -= (Number(d[i][qi]) || 0);
        if (ciEquals(d[i][ti], warehouse)) qty += (Number(d[i][qi]) || 0);
      }
    }
  }
  return Math.max(0, qty);
}

function rebuildAllMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const n = _doFullRebuild(ss);
  ss.toast(`✅ Rebuilt ${n} combos successfully`, 'Full Rebuild Complete', 6);
}

function rebuildAllMasterSilent(ss) { _doFullRebuild(ss); }

function _doFullRebuild(ss) {
  const combos = new Set();
  const extractCombos = (sheetName, whCol1, whCol2 = null, explicitWh = null) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const d = sheet.getDataRange().getValues(), h = d[0].map(x => String(x).trim().toLowerCase());
    const mi = h.indexOf('mills') !== -1 ? h.indexOf('mills') : h.indexOf('mill'), pi = h.indexOf('product');
    const wi1 = whCol1 ? h.indexOf(whCol1.toLowerCase()) : -1, wi2 = whCol2 ? h.indexOf(whCol2.toLowerCase()) : -1;
    for (let i = 1; i < d.length; i++) {
      const m = (d[i][mi]||'').toString().trim(), p = (d[i][pi]||'').toString().trim();
      if (!m || !p) continue;
      if (explicitWh) combos.add(`${m}||${p}||${explicitWh}`);
      if (wi1 >= 0 && d[i][wi1]) combos.add(`${m}||${p}||${d[i][wi1]}`);
      if (wi2 >= 0 && d[i][wi2]) combos.add(`${m}||${p}||${d[i][wi2]}`);
    }
  };

  WAREHOUSES.forEach(wh => {
    extractCombos(`PurchaseLog_${wh}`, null, null, wh);
    extractCombos(`SalesLog_${wh}`, null, null, wh);
  });
  extractCombos('TransferLog', 'from', 'to');

  const master = ss.getSheetByName('MasterStock');
  if (master) {
    const d = master.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      const m = (d[i][0]||'').toString().trim(), p = (d[i][1]||'').toString().trim(), w = (d[i][2]||'').toString().trim();
      if (m && p && w) combos.add(`${m}||${p}||${w}`);
    }
  }

  const seen = {}, uniqueCombos = [];
  combos.forEach(key => { const lk = key.toLowerCase(); if (!seen[lk]) { seen[lk] = true; uniqueCombos.push(key); } });
  uniqueCombos.forEach(key => { const [m, p, w] = key.split('||'); recomputeMasterFromLogs(ss, m, p, w); recalcSalesLogProfit(ss, m, p, w); });
  return uniqueCombos.length;
}

function recomputeMasterFromLogs(ss, mill, product, warehouse) {
  const wh = (warehouse || '').toString().toLowerCase().trim();
  if (wh.includes('khera')) {
    return recomputeMasterKheraKalan(ss, mill, product);
  } else if (wh.includes('mukherjee')) {
    return recomputeMasterMukherjeeNagar(ss, mill, product);
  } else if (wh.includes('truck')) {
    return recomputeMasterByTruck(ss, mill, product);
  } else {
    return recomputeMasterKheraKalan(ss, mill, product);
  }
}

function getAuditTimeMap(ss) {
  const auditTimeMap = {};
  const auditSheet = ss.getSheetByName('AuditLog');
  if (auditSheet && auditSheet.getLastRow() > 1) {
    const ad = auditSheet.getDataRange().getValues(), ah = ad[0].map(x => String(x||'').trim().toLowerCase());
    const aTsI = ah.indexOf('timestamp'), aActI = ah.indexOf('action'), aSheetI = ah.indexOf('sheet'), aRowI = ah.indexOf('row');
    for (let i = 1; i < ad.length; i++) {
      if (String(ad[i][aActI]||'').toUpperCase() !== 'CREATE') continue;
      const sheetName = String(ad[i][aSheetI]||'');
      const rowNum = Number(ad[i][aRowI]) || 0;
      if (!sheetName || !rowNum) continue;
      const ts = ad[i][aTsI] ? new Date(ad[i][aTsI]) : null;
      if (!ts) continue;
      const key = `${sheetName}||${rowNum}`;
      if (!auditTimeMap[key] || ts < auditTimeMap[key]) auditTimeMap[key] = ts;
    }
  }
  return auditTimeMap;
}

function recomputeMasterMukherjeeNagar(ss, mill, product) {
  const warehouse = "Mukherjee Nagar";
  const master = getOrCreateSheet(ss, 'MasterStock', ['Mill','Product','Warehouse','Total Qty','Avg_GST','Avg_NonGST','Total Value','Unit']);
  const auditTimeMap = getAuditTimeMap(ss);
  const events = [];
  let lastKnownUnit = 'Ream';
  let seq = 0;
  let rateSetGST = 0, rateSetNonGST = 0;

  const pLog = ss.getSheetByName('PurchaseLog_Mukherjee Nagar');
  if (pLog) {
    const d = pLog.getDataRange().getValues(), h = d[0].map(x => String(x || '').trim().toLowerCase());
    const mi = h.indexOf('mills') !== -1 ? h.indexOf('mills') : h.indexOf('mill');
    const pi = h.indexOf('product'), qi = h.indexOf('qty'), ri = h.indexOf('price'), ui = h.indexOf('unit type');
    const ti = h.indexOf('bill type') !== -1 ? h.indexOf('bill type') : h.indexOf('type');

    for (let i = 1; i < d.length; i++) {
      if (!ciEquals(d[i][mi], mill) || !ciEquals(d[i][pi], product)) continue;
      const qty = Number(d[i][qi]) || 0;
      const basePrice = Number(d[i][ri]) || 0;
      const type = String(d[i][ti] || '').toUpperCase().replace(/[-\s]/g, '');
      const entryDate = d[i][0] ? new Date(d[i][0]) : new Date(0);
      if (ui >= 0 && d[i][ui]) lastKnownUnit = d[i][ui];

      const isNon = type.includes('NON');
      if (qty <= 0 && basePrice > 0) {
        if (isNon) rateSetNonGST = basePrice;
        else rateSetGST = basePrice;
        continue;
      }
      if (qty <= 0) continue;

      const auditTs = auditTimeMap[`PurchaseLog_Mukherjee Nagar||${i + 1}`] || null;
      events.push({
        date: entryDate,
        auditTs,
        seq: seq++,
        type: 'IN',
        qty,
        gstPrice: isNon ? 0 : basePrice,
        nonGstPrice: isNon ? basePrice : 0
      });
    }
  }

  const tLog = ss.getSheetByName('TransferLog');
  if (tLog) {
    const d = tLog.getDataRange().getValues(), h = d[0].map(x => String(x || '').trim().toLowerCase());
    const mi = h.indexOf('mill') !== -1 ? h.indexOf('mill') : h.indexOf('mills');
    const pi = h.indexOf('product'), qi = h.indexOf('qty');
    let rateI = h.indexOf('rate');
    if (rateI === -1) rateI = h.indexOf('avg price');
    if (rateI === -1) rateI = h.indexOf('avg_price');
    let fi = h.indexOf('from');
    if (fi === -1) fi = h.indexOf('from warehouse');
    let ti = h.indexOf('to');
    if (ti === -1) ti = h.indexOf('to warehouse');

    for (let i = 1; i < d.length; i++) {
      if (!ciEquals(d[i][mi], mill) || !ciEquals(d[i][pi], product)) continue;
      const qty = Number(d[i][qi]) || 0;
      if (qty <= 0) continue;
      const entryDate = d[i][0] ? new Date(d[i][0]) : new Date(0);
      const auditTs = auditTimeMap[`TransferLog||${i + 1}`] || null;

      if (ciEquals(d[i][fi], warehouse)) {
        events.push({ date: entryDate, auditTs, seq: seq++, type: 'OUT', qty, gstPrice: 0, nonGstPrice: 0 });
      }
      if (ciEquals(d[i][ti], warehouse)) {
        let transferRate = rateI >= 0 ? (Number(d[i][rateI]) || 0) : 0;
        if (transferRate <= 0) {
          const fromWh = String(d[i][fi] || 'Khera Kalan').trim();
          const srcRow = findProductRowWithWarehouse(master, mill, product, fromWh);
          if (srcRow > 0) transferRate = Number(master.getRange(srcRow, 5).getValue()) || 0;
        }
        events.push({ date: entryDate, auditTs, seq: seq++, type: 'IN', qty, gstPrice: transferRate > 0 ? transferRate : 0, nonGstPrice: 0 });
      }
    }
  }

  const sLog = ss.getSheetByName('SalesLog_Mukherjee Nagar');
  if (sLog) {
    const d = sLog.getDataRange().getValues(), h = d[0].map(x => String(x || '').trim().toLowerCase());
    const mi = h.indexOf('mills') !== -1 ? h.indexOf('mills') : h.indexOf('mill');
    const pi = h.indexOf('product'), qi = h.indexOf('qty');
    for (let i = 1; i < d.length; i++) {
      if (!ciEquals(d[i][mi], mill) || !ciEquals(d[i][pi], product)) continue;
      const qty = Number(d[i][qi]) || 0;
      if (qty <= 0) continue;
      const entryDate = d[i][0] ? new Date(d[i][0]) : new Date(0);
      const auditTs = auditTimeMap[`SalesLog_Mukherjee Nagar||${i + 1}`] || null;
      events.push({ date: entryDate, auditTs, seq: seq++, type: 'OUT', qty, gstPrice: 0, nonGstPrice: 0 });
    }
  }

  _processAndWriteToMaster(master, mill, product, warehouse, events, lastKnownUnit, rateSetGST, rateSetNonGST);
}

function recomputeMasterKheraKalan(ss, mill, product) {
  const warehouse = "Khera Kalan";
  const master = getOrCreateSheet(ss, 'MasterStock', ['Mill','Product','Warehouse','Total Qty','Avg_GST','Avg_NonGST','Total Value','Unit']);
  const auditTimeMap = getAuditTimeMap(ss);
  const events = [];
  let lastKnownUnit = 'Ream';
  let seq = 0;
  let rateSetGST = 0, rateSetNonGST = 0;

  const pLog = ss.getSheetByName('PurchaseLog_Khera Kalan');
  if (pLog) {
    const d = pLog.getDataRange().getValues(), h = d[0].map(x => String(x || '').trim().toLowerCase());
    const mi = h.indexOf('mills') !== -1 ? h.indexOf('mills') : h.indexOf('mill');
    const pi = h.indexOf('product'), qi = h.indexOf('qty'), ri = h.indexOf('price'), ui = h.indexOf('unit type');
    const ti = h.indexOf('bill type') !== -1 ? h.indexOf('bill type') : h.indexOf('type');

    for (let i = 1; i < d.length; i++) {
      if (!ciEquals(d[i][mi], mill) || !ciEquals(d[i][pi], product)) continue;
      const qty = Number(d[i][qi]) || 0;
      const basePrice = Number(d[i][ri]) || 0;
      const type = String(d[i][ti] || '').toUpperCase().replace(/[-\s]/g, '');
      const entryDate = d[i][0] ? new Date(d[i][0]) : new Date(0);
      if (ui >= 0 && d[i][ui]) lastKnownUnit = d[i][ui];

      const isNon = type.includes('NON');
      if (qty <= 0 && basePrice > 0) {
        if (isNon) rateSetNonGST = basePrice;
        else rateSetGST = basePrice;
        continue;
      }
      if (qty <= 0) continue;

      const auditTs = auditTimeMap[`PurchaseLog_Khera Kalan||${i + 1}`] || null;
      events.push({
        date: entryDate,
        auditTs,
        seq: seq++,
        type: 'IN',
        qty,
        gstPrice: isNon ? 0 : basePrice,
        nonGstPrice: isNon ? basePrice : 0
      });
    }
  }

  const tLog = ss.getSheetByName('TransferLog');
  if (tLog) {
    const d = tLog.getDataRange().getValues(), h = d[0].map(x => String(x || '').trim().toLowerCase());
    const mi = h.indexOf('mill') !== -1 ? h.indexOf('mill') : h.indexOf('mills');
    const pi = h.indexOf('product'), qi = h.indexOf('qty');
    let rateI = h.indexOf('rate');
    if (rateI === -1) rateI = h.indexOf('avg price');
    if (rateI === -1) rateI = h.indexOf('avg_price');
    let fi = h.indexOf('from');
    if (fi === -1) fi = h.indexOf('from warehouse');
    let ti = h.indexOf('to');
    if (ti === -1) ti = h.indexOf('to warehouse');

    for (let i = 1; i < d.length; i++) {
      if (!ciEquals(d[i][mi], mill) || !ciEquals(d[i][pi], product)) continue;
      const qty = Number(d[i][qi]) || 0;
      if (qty <= 0) continue;
      const entryDate = d[i][0] ? new Date(d[i][0]) : new Date(0);
      const auditTs = auditTimeMap[`TransferLog||${i + 1}`] || null;

      if (ciEquals(d[i][fi], warehouse)) {
        events.push({ date: entryDate, auditTs, seq: seq++, type: 'OUT', qty, gstPrice: 0, nonGstPrice: 0 });
      }
      if (ciEquals(d[i][ti], warehouse)) {
        let transferRate = rateI >= 0 ? (Number(d[i][rateI]) || 0) : 0;
        if (transferRate <= 0) {
          const fromWh = String(d[i][fi] || 'Mukherjee Nagar').trim();
          const srcRow = findProductRowWithWarehouse(master, mill, product, fromWh);
          if (srcRow > 0) transferRate = Number(master.getRange(srcRow, 5).getValue()) || 0;
        }
        events.push({ date: entryDate, auditTs, seq: seq++, type: 'IN', qty, gstPrice: transferRate > 0 ? transferRate : 0, nonGstPrice: 0 });
      }
    }
  }

  const sLog = ss.getSheetByName('SalesLog_Khera Kalan');
  if (sLog) {
    const d = sLog.getDataRange().getValues(), h = d[0].map(x => String(x || '').trim().toLowerCase());
    const mi = h.indexOf('mills') !== -1 ? h.indexOf('mills') : h.indexOf('mill');
    const pi = h.indexOf('product'), qi = h.indexOf('qty');
    for (let i = 1; i < d.length; i++) {
      if (!ciEquals(d[i][mi], mill) || !ciEquals(d[i][pi], product)) continue;
      const qty = Number(d[i][qi]) || 0;
      if (qty <= 0) continue;
      const entryDate = d[i][0] ? new Date(d[i][0]) : new Date(0);
      const auditTs = auditTimeMap[`SalesLog_Khera Kalan||${i + 1}`] || null;
      events.push({ date: entryDate, auditTs, seq: seq++, type: 'OUT', qty, gstPrice: 0, nonGstPrice: 0 });
    }
  }

  _processAndWriteToMaster(master, mill, product, warehouse, events, lastKnownUnit, rateSetGST, rateSetNonGST);
}

function recomputeMasterByTruck(ss, mill, product) {
  const warehouse = "By Truck";
  const master = getOrCreateSheet(ss, 'MasterStock', ['Mill','Product','Warehouse','Total Qty','Avg_GST','Avg_NonGST','Total Value','Unit']);
  const auditTimeMap = getAuditTimeMap(ss);
  const events = [];
  let lastKnownUnit = 'Ream';
  let seq = 0;
  let rateSetGST = 0, rateSetNonGST = 0;

  const pLog = ss.getSheetByName('PurchaseLog_By Truck');
  if (pLog) {
    const d = pLog.getDataRange().getValues(), h = d[0].map(x => String(x || '').trim().toLowerCase());
    const mi = h.indexOf('mills') !== -1 ? h.indexOf('mills') : h.indexOf('mill');
    const pi = h.indexOf('product'), qi = h.indexOf('qty'), ri = h.indexOf('price'), ui = h.indexOf('unit type');
    const ti = h.indexOf('bill type') !== -1 ? h.indexOf('bill type') : h.indexOf('type');

    for (let i = 1; i < d.length; i++) {
      if (!ciEquals(d[i][mi], mill) || !ciEquals(d[i][pi], product)) continue;
      const qty = Number(d[i][qi]) || 0;
      const basePrice = Number(d[i][ri]) || 0;
      const type = String(d[i][ti] || '').toUpperCase().replace(/[-\s]/g, '');
      const entryDate = d[i][0] ? new Date(d[i][0]) : new Date(0);
      if (ui >= 0 && d[i][ui]) lastKnownUnit = d[i][ui];

      const isNon = type.includes('NON');
      if (qty <= 0 && basePrice > 0) {
        if (isNon) rateSetNonGST = basePrice;
        else rateSetGST = basePrice;
        continue;
      }
      if (qty <= 0) continue;

      const auditTs = auditTimeMap[`PurchaseLog_By Truck||${i + 1}`] || null;
      events.push({
        date: entryDate,
        auditTs,
        seq: seq++,
        type: 'IN',
        qty,
        gstPrice: isNon ? 0 : basePrice,
        nonGstPrice: isNon ? basePrice : 0
      });
    }
  }

  const tLog = ss.getSheetByName('TransferLog');
  if (tLog) {
    const d = tLog.getDataRange().getValues(), h = d[0].map(x => String(x || '').trim().toLowerCase());
    const mi = h.indexOf('mill') !== -1 ? h.indexOf('mill') : h.indexOf('mills');
    const pi = h.indexOf('product'), qi = h.indexOf('qty');
    let rateI = h.indexOf('rate');
    if (rateI === -1) rateI = h.indexOf('avg price');
    if (rateI === -1) rateI = h.indexOf('avg_price');
    let fi = h.indexOf('from');
    if (fi === -1) fi = h.indexOf('from warehouse');
    let ti = h.indexOf('to');
    if (ti === -1) ti = h.indexOf('to warehouse');

    for (let i = 1; i < d.length; i++) {
      if (!ciEquals(d[i][mi], mill) || !ciEquals(d[i][pi], product)) continue;
      const qty = Number(d[i][qi]) || 0;
      if (qty <= 0) continue;
      const entryDate = d[i][0] ? new Date(d[i][0]) : new Date(0);
      const auditTs = auditTimeMap[`TransferLog||${i + 1}`] || null;

      if (ciEquals(d[i][fi], warehouse)) {
        events.push({ date: entryDate, auditTs, seq: seq++, type: 'OUT', qty, gstPrice: 0, nonGstPrice: 0 });
      }
      if (ciEquals(d[i][ti], warehouse)) {
        const transferRate = rateI >= 0 ? (Number(d[i][rateI]) || 0) : 0;
        events.push({ date: entryDate, auditTs, seq: seq++, type: 'IN', qty, gstPrice: transferRate > 0 ? transferRate : 0, nonGstPrice: 0 });
      }
    }
  }

  const sLog = ss.getSheetByName('SalesLog_By Truck');
  if (sLog) {
    const d = sLog.getDataRange().getValues(), h = d[0].map(x => String(x || '').trim().toLowerCase());
    const mi = h.indexOf('mills') !== -1 ? h.indexOf('mills') : h.indexOf('mill');
    const pi = h.indexOf('product'), qi = h.indexOf('qty');
    for (let i = 1; i < d.length; i++) {
      if (!ciEquals(d[i][mi], mill) || !ciEquals(d[i][pi], product)) continue;
      const qty = Number(d[i][qi]) || 0;
      if (qty <= 0) continue;
      const entryDate = d[i][0] ? new Date(d[i][0]) : new Date(0);
      const auditTs = auditTimeMap[`SalesLog_By Truck||${i + 1}`] || null;
      events.push({ date: entryDate, auditTs, seq: seq++, type: 'OUT', qty, gstPrice: 0, nonGstPrice: 0 });
    }
  }

  _processAndWriteToMaster(master, mill, product, warehouse, events, lastKnownUnit, rateSetGST, rateSetNonGST);
}

function _processAndWriteToMaster(master, mill, product, warehouse, events, lastKnownUnit, rateSetGST, rateSetNonGST) {
  events.sort((a, b) => {
    const dt = a.date - b.date;
    if (dt !== 0) return dt;
    if (a.auditTs && b.auditTs) {
      const at = a.auditTs - b.auditTs;
      if (at !== 0) return at;
    }
    return a.seq - b.seq;
  });

  let runningQty = 0;
  let gstSum = 0, gstQty = 0;
  let nonGstSum = 0, nonGstQty = 0;
  let lastGstRate = 0, lastNonGstRate = 0;

  for (const ev of events) {
    if (ev.type === 'IN') {
      if (runningQty <= 0.001) {
        gstSum = 0;
        gstQty = 0;
        nonGstSum = 0;
        nonGstQty = 0;
      }
      runningQty += ev.qty;
      if (ev.gstPrice > 0) {
        gstSum += ev.gstPrice * ev.qty;
        gstQty += ev.qty;
        lastGstRate = ev.gstPrice;
      }
      if (ev.nonGstPrice > 0) {
        nonGstSum += ev.nonGstPrice * ev.qty;
        nonGstQty += ev.qty;
        lastNonGstRate = ev.nonGstPrice;
      }
    } else {
      runningQty -= ev.qty;
    }
  }

  let finalQty = Math.max(0, Number(runningQty.toFixed(2)));
  let avgGST = 0;
  let avgNonGST = 0;

  if (warehouse === 'By Truck') {
    if (finalQty > 0.001) {
      avgGST = gstQty > 0 ? Number((gstSum / gstQty).toFixed(2)) : (rateSetGST || lastGstRate);
      avgNonGST = nonGstQty > 0 ? Number((nonGstSum / nonGstQty).toFixed(2)) : (rateSetNonGST || lastNonGstRate);
    } else {
      finalQty = 0;
      avgGST = 0;
      avgNonGST = 0;
    }
  } else {
    if (finalQty > 0.001) {
      avgGST = gstQty > 0 ? Number((gstSum / gstQty).toFixed(2)) : (rateSetGST || lastGstRate);
      avgNonGST = nonGstQty > 0 ? Number((nonGstSum / nonGstQty).toFixed(2)) : (rateSetNonGST || lastNonGstRate);
    } else {
      finalQty = 0;
      avgGST = rateSetGST || lastGstRate || 0;
      avgNonGST = rateSetNonGST || lastNonGstRate || 0;
    }
  }

  const effectiveAvg = avgGST > 0 ? avgGST : avgNonGST;
  const totalVal = Number((finalQty * effectiveAvg).toFixed(2));

  const masterRow = findProductRowWithWarehouse(master, mill, product, warehouse);
  if (masterRow > 0) {
    master.getRange(masterRow, 4, 1, 4).setValues([[
      finalQty,
      avgGST,
      avgNonGST,
      totalVal
    ]]);
  } else {
    master.appendRow([mill, product, warehouse, finalQty, avgGST, avgNonGST, totalVal, lastKnownUnit || 'Ream']);
  }
}

function rebuildMukherjeeNagarOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _rebuildSingleWarehouse(ss, 'Mukherjee Nagar', recomputeMasterMukherjeeNagar);
  ss.toast('✅ Mukherjee Nagar stock successfully updated!', 'Done', 5);
}

function rebuildKheraKalanOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _rebuildSingleWarehouse(ss, 'Khera Kalan', recomputeMasterKheraKalan);
  ss.toast('✅ Khera Kalan stock successfully updated!', 'Done', 5);
}

function rebuildByTruckOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _rebuildSingleWarehouse(ss, 'By Truck', recomputeMasterByTruck);
  ss.toast('✅ By Truck stock successfully updated!', 'Done', 5);
}

function _rebuildSingleWarehouse(ss, warehouse, recomputeFn) {
  const combos = new Set();
  const pSheet = ss.getSheetByName(`PurchaseLog_${warehouse}`);
  if (pSheet) {
    const d = pSheet.getDataRange().getValues(), h = d[0].map(x => String(x).trim().toLowerCase());
    const mi = h.indexOf('mills') !== -1 ? h.indexOf('mills') : h.indexOf('mill'), pi = h.indexOf('product');
    for (let i = 1; i < d.length; i++) {
      if (d[i][mi] && d[i][pi]) combos.add(`${d[i][mi]}||${d[i][pi]}`);
    }
  }
  const sSheet = ss.getSheetByName(`SalesLog_${warehouse}`);
  if (sSheet) {
    const d = sSheet.getDataRange().getValues(), h = d[0].map(x => String(x).trim().toLowerCase());
    const mi = h.indexOf('mills') !== -1 ? h.indexOf('mills') : h.indexOf('mill'), pi = h.indexOf('product');
    for (let i = 1; i < d.length; i++) {
      if (d[i][mi] && d[i][pi]) combos.add(`${d[i][mi]}||${d[i][pi]}`);
    }
  }
  const tSheet = ss.getSheetByName('TransferLog');
  if (tSheet) {
    const d = tSheet.getDataRange().getValues(), h = d[0].map(x => String(x).trim().toLowerCase());
    const mi = h.indexOf('mill') !== -1 ? h.indexOf('mill') : h.indexOf('mills'), pi = h.indexOf('product');
    let fi = h.indexOf('from'), ti = h.indexOf('to');
    for (let i = 1; i < d.length; i++) {
      if (d[i][mi] && d[i][pi] && (ciEquals(d[i][fi], warehouse) || ciEquals(d[i][ti], warehouse))) {
        combos.add(`${d[i][mi]}||${d[i][pi]}`);
      }
    }
  }

  combos.forEach(key => {
    const [m, p] = key.split('||');
    recomputeFn(ss, m, p);
    recalcSalesLogProfit(ss, m, p, warehouse);
  });
}

function recalcSalesLogProfit(ss, mill, product, warehouse) {
  const sLog = ss.getSheetByName(`SalesLog_${warehouse}`), master = ss.getSheetByName('MasterStock');
  if (!sLog || !master) return;

  let baseGST = 0, baseNonGST = 0;
  const masterRow = findProductRowWithWarehouse(master, mill, product, warehouse);
  if (masterRow > 0) {
    const masterVals = master.getRange(masterRow, 4, 1, 3).getValues()[0];
    baseGST    = Number(masterVals[1]) || 0;
    baseNonGST = Number(masterVals[2]) || 0;
  }

  const d = sLog.getDataRange().getValues(), h = d[0].map(x => String(x || '').trim().toLowerCase());
  const mi = h.indexOf('mills') !== -1 ? h.indexOf('mills') : h.indexOf('mill');
  const pi = h.indexOf('product'), qi = h.indexOf('qty'), ai = h.indexOf('amount'), ti = h.indexOf('bill type');
  const ri = h.indexOf('rate'), di = h.indexOf('discount');
  const ci = h.indexOf('cost price'), fi = h.indexOf('profit'), frIdx = h.indexOf('freight'), recIdx = h.indexOf('recovered?');
  if (fi === -1) return;

  for (let i = 1; i < d.length; i++) {
    if (!ciEquals(d[i][mi], mill) || !ciEquals(d[i][pi], product)) continue;

    const existingCP = ci !== -1 ? (Number(d[i][ci]) || 0) : 0;
    if (existingCP > 0) continue;

    const isRecovered = recIdx !== -1 && String(d[i][recIdx] || '').trim().toLowerCase() === 'yes';
    const freight = frIdx !== -1 ? (Number(d[i][frIdx]) || 0) : 0;
    const freightDeduction = isRecovered ? 0 : freight;

    const qty = Number(d[i][qi]) || 0, rate = Number(d[i][ri]) || 0, discount = Number(d[i][di]) || 0;
    let saleAmt = Number(d[i][ai]) || 0;
    if (saleAmt === 0 && qty > 0 && rate > 0) saleAmt = (qty * rate) - discount;

    const billType = String(d[i][ti] || '').toUpperCase().replace(/[-\s]/g, '');
    const costPrice = billType.includes('NON') ? baseNonGST : baseGST;
    const profit = costPrice > 0
      ? Number((saleAmt - (qty * costPrice) - freightDeduction).toFixed(2))
      : Number((saleAmt - freightDeduction).toFixed(2));

    if (ci !== -1) sLog.getRange(i + 1, ci + 1).setValue(Number(costPrice.toFixed(2)));
    sLog.getRange(i + 1, fi + 1).setValue(profit);
  }
}

function processData(payload) {
  const data = typeof payload === 'string' ? JSON.parse(payload) : payload;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const entryDate = data.date ? new Date(data.date) : new Date();

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    return { status: 'error', message: '⏳ System busy hai — dusra save chal raha hai. 10 sec baad try karein.' };
  }

  try {
    if (!data.warehouse) throw new Error('Please Choose Warehouse!');

    if (data.action === 'STOCK_OUT') {
      const master = ss.getSheetByName('MasterStock');
      const sheetName = `SalesLog_${data.warehouse}`;
      const log = getOrCreateSheet(ss, sheetName, ['Date','Sales Person','Pad No.','Bill No.','Party Name','Mills','Product','Warehouse','Bill Type','Qty','Rate','Discount','Amount','Freight','Recovered?','Cost Price','Profit']);
      const partyNameUpper = (data.partyName || '').toString().toUpperCase().trim();
      const totalFreight = Number(data.freight) || 0;
      const isRecovered = (data.freightRecovered || 'No').toString().toLowerCase() === 'yes';
      const totalQtyAll = data.items.reduce((s, i) => s + (Number(i.qty) || 0), 0);

      function getFifoBatches(mill, product, warehouse, billType) {
        const pLog = ss.getSheetByName(`PurchaseLog_${warehouse}`);
        const sLog = ss.getSheetByName(`SalesLog_${warehouse}`);
        const incomingTypeIsNon = String(billType).toUpperCase().replace(/[-\s]/g,'').includes('NON');
        const purchases = [];

        if (pLog) {
          const pd = pLog.getDataRange().getValues(), ph = pd[0].map(x => String(x||'').trim().toLowerCase());
          const p_mi = ph.indexOf('mills') !== -1 ? ph.indexOf('mills') : ph.indexOf('mill');
          const p_pi = ph.indexOf('product'), p_qi = ph.indexOf('qty'), p_ri = ph.indexOf('price');
          const p_ti = ph.indexOf('bill type') !== -1 ? ph.indexOf('bill type') : ph.indexOf('type'), p_di = 0;
          const p_insI = ph.indexOf('insurance'), p_carI = ph.indexOf('cartage'), p_lodI = ph.indexOf('loading');
          const p_extrasClaimedI = ph.indexOf('extras claimed');

          for (let i = 1; i < pd.length; i++) {
            if (!ciEquals(pd[i][p_mi], mill) || !ciEquals(pd[i][p_pi], product)) continue;
            const bTypeIsNon = String(pd[i][p_ti]||'').toUpperCase().replace(/[-\s]/g,'').includes('NON');
            if (incomingTypeIsNon !== bTypeIsNon) continue;
            const qty = Number(pd[i][p_qi]) || 0;
            if (qty <= 0) continue;
            const price = Number(pd[i][p_ri]) || 0;
            let extras = 0, extrasClaimed = true;
            if (warehouse === 'By Truck') {
              extras = (p_insI >= 0 ? Number(pd[i][p_insI])||0 : 0)
                     + (p_carI >= 0 ? Number(pd[i][p_carI])||0 : 0)
                     + (p_lodI >= 0 ? Number(pd[i][p_lodI])||0 : 0);
              extrasClaimed = p_extrasClaimedI >= 0 ? String(pd[i][p_extrasClaimedI]||'').toLowerCase().trim() === 'yes' : true;
            }
            purchases.push({ date: pd[i][p_di] ? new Date(pd[i][p_di]) : new Date(0), qty, price, purchaseRow: i + 1, extras, extrasClaimed });
          }
        }

        if (!incomingTypeIsNon) {
          const tLog = ss.getSheetByName('TransferLog');
          if (tLog) {
            const td = tLog.getDataRange().getValues(), th = td[0].map(x => String(x||'').trim().toLowerCase());
            const t_mi = th.indexOf('mill') !== -1 ? h.indexOf('mill') : th.indexOf('mills');
            const t_pi = th.indexOf('product'), t_qi = th.indexOf('qty');
            let t_rateI = th.indexOf('rate');
            if (t_rateI === -1) t_rateI = th.indexOf('avg price');
            if (t_rateI === -1) t_rateI = th.indexOf('avg_price');
            let t_toI = th.indexOf('to'); if (t_toI === -1) t_toI = th.indexOf('to warehouse');
            const t_di = 0;

            for (let i = 1; i < td.length; i++) {
              if (!ciEquals(td[i][t_mi], mill) || !ciEquals(td[i][t_pi], product)) continue;
              if (!ciEquals(td[i][t_toI], warehouse)) continue;
              const qty = Number(td[i][t_qi]) || 0;
              if (qty <= 0) continue;
              const price = t_rateI >= 0 ? (Number(td[i][t_rateI]) || 0) : 0;
              if (price <= 0) continue;
              purchases.push({ date: td[i][t_di] ? new Date(td[i][t_di]) : new Date(0), qty, price, purchaseRow: null, extras: 0, extrasClaimed: true });
            }
          }
        }

        purchases.sort((a, b) => a.date - b.date);

        const soldQtyByCP = {};
        if (sLog && sLog.getLastRow() > 1) {
          const sd = sLog.getDataRange().getValues(), sh = sd[0].map(x => String(x||'').trim().toLowerCase());
          const s_mi = sh.indexOf('mills') !== -1 ? sh.indexOf('mills') : sh.indexOf('mill');
          const s_pi = sh.indexOf('product'), s_qi = sh.indexOf('qty'), s_ci = sh.indexOf('cost price'), s_ti = sh.indexOf('bill type');
          for (let i = 1; i < sd.length; i++) {
            if (!ciEquals(sd[i][s_mi], mill) || !ciEquals(sd[i][s_pi], product)) continue;
            if (s_ti >= 0) { if (incomingTypeIsNon !== String(sd[i][s_ti]||'').toUpperCase().replace(/[-\s]/g,'').includes('NON')) continue; }
            const cpKey = (Number(sd[i][s_ci]) || 0).toFixed(2);
            soldQtyByCP[cpKey] = (soldQtyByCP[cpKey] || 0) + (Number(sd[i][s_qi]) || 0);
          }
        }

        const purchasePool = purchases.map(p => ({ price: p.price, remaining: p.qty, purchaseRow: p.purchaseRow, extras: p.extras, extrasClaimed: p.extrasClaimed }));
        for (const cpKey of Object.keys(soldQtyByCP)) {
          let toConsume = soldQtyByCP[cpKey];
          for (const batch of purchasePool) {
            if (toConsume <= 0) break;
            if (Math.abs(batch.price - parseFloat(cpKey)) < 0.01) { const c = Math.min(batch.remaining, toConsume); batch.remaining -= c; toConsume -= c; }
          }
          if (toConsume > 0) {
            for (const batch of purchasePool) {
              if (toConsume <= 0) break;
              const c = Math.min(batch.remaining, toConsume); batch.remaining -= c; toConsume -= c;
            }
          }
        }

        const merged = [];
        for (const b of purchasePool) {
          if (b.remaining > 0.001) {
            const canMergeWithPrev = merged.length > 0
              && Math.abs(merged[merged.length-1].price - b.price) < 0.01
              && merged[merged.length-1].purchaseRow === b.purchaseRow;
            if (canMergeWithPrev) merged[merged.length-1].available += b.remaining;
            else merged.push({ price: b.price, available: b.remaining, purchaseRow: b.purchaseRow, extras: b.extras, extrasClaimed: b.extrasClaimed });
          }
        }
        return merged;
      }

      const validatedItems = [];
      data.items.forEach(item => {
        const row = findProductRowWithWarehouse(master, item.mill, item.product, data.warehouse);
        if (row === -1) throw new Error(`${item.product} not found in ${data.warehouse}!`);
        const vals = master.getRange(row, 4, 1, 3).getValues()[0];
        const currentQty = Number(vals[0]) || 0, itemQty = Number(item.qty);
        if (itemQty <= 0) throw new Error(`Invalid qty for ${item.product}`);
        if (currentQty < itemQty) throw new Error(`Stock kam hai: ${item.product} (available: ${currentQty})`);
        validatedItems.push({ item, row, currentQty, avgGST: Number(vals[1]) || 0, avgNonGST: Number(vals[2]) || 0, itemQty });
      });

      function markExtrasClaimed(warehouse, purchaseRow) {
        if (warehouse !== 'By Truck' || !purchaseRow) return;
        const pLog = ss.getSheetByName(`PurchaseLog_${warehouse}`);
        if (!pLog) return;
        const ph = pLog.getRange(1, 1, 1, pLog.getLastColumn()).getValues()[0].map(x => String(x||'').trim().toLowerCase());
        const claimedI = ph.indexOf('extras claimed');
        if (claimedI === -1) return;
        pLog.getRange(purchaseRow, claimedI + 1).setValue('Yes');
      }

      let freightRunning = 0;
      validatedItems.forEach(({ item, row, currentQty, avgGST, avgNonGST, itemQty }, idx) => {
        const billTypeRaw = String(item.billType || '').toUpperCase().replace(/[-\s]/g, '');
        const batches = getFifoBatches(item.mill, item.product, data.warehouse, item.billType);

        let remainingSaleQty = itemQty;
        const splitRows = [];
        for (const batch of batches) {
          if (remainingSaleQty <= 0.001) break;
          const useQty = Math.min(remainingSaleQty, batch.available);
          if (useQty > 0.001) {
            let extrasToClaim = 0;
            if (data.warehouse === 'By Truck' && !batch.extrasClaimed && batch.extras > 0) {
              extrasToClaim = batch.extras;
              markExtrasClaimed(data.warehouse, batch.purchaseRow);
            }
            splitRows.push({ qty: useQty, costPrice: batch.price, extras: extrasToClaim });
            remainingSaleQty -= useQty;
          }
        }

        if (remainingSaleQty > 0.001) {
          let fallbackCP = 0;
          const pLogFB = ss.getSheetByName(`PurchaseLog_${data.warehouse}`);
          if (pLogFB) {
            const pd = pLogFB.getDataRange().getValues(), ph = pd[0].map(x => String(x||'').trim().toLowerCase());
            const pMi = ph.indexOf('mills') !== -1 ? ph.indexOf('mills') : ph.indexOf('mill');
            const pPi = ph.indexOf('product'), pQi = ph.indexOf('qty'), pRi = ph.indexOf('price');
            const pTi = ph.indexOf('bill type') !== -1 ? ph.indexOf('bill type') : ph.indexOf('type');
            let sum = 0, qtyTotal = 0;
            for (let i = 1; i < pd.length; i++) {
              if (!ciEquals(pd[i][pMi], item.mill) || !ciEquals(pd[i][pPi], item.product)) continue;
              const q = Number(pd[i][pQi]) || 0;
              if (q <= 0) continue;
              const typeIsNon = String(pd[i][pTi]||'').toUpperCase().replace(/[-\s]/g,'').includes('NON');
              if (typeIsNon === billTypeRaw.includes('NON')) { sum += (Number(pd[i][pRi])||0) * q; qtyTotal += q; }
            }
            fallbackCP = qtyTotal > 0 ? sum / qtyTotal : (billTypeRaw.includes('NON') ? avgNonGST : avgGST);
          } else {
            fallbackCP = billTypeRaw.includes('NON') ? avgNonGST : avgGST;
          }
          splitRows.push({ qty: remainingSaleQty, costPrice: fallbackCP, extras: 0 });
        }
        if (splitRows.length === 0) {
          const fallback = billTypeRaw.includes('NON') ? avgNonGST : avgGST;
          splitRows.push({ qty: itemQty, costPrice: fallback, extras: 0 });
        }

        let itemFreight = 0;
        if (totalFreight > 0 && totalQtyAll > 0) {
          if (idx === validatedItems.length - 1) itemFreight = Math.round((totalFreight - freightRunning) * 100) / 100;
          else { itemFreight = Math.round((itemQty / totalQtyAll) * totalFreight * 100) / 100; freightRunning += itemFreight; }
        }

        let newQty = currentQty - itemQty;
        if (newQty <= 0.001) { newQty = 0; }
        let finalAvgGST = avgGST, finalAvgNonGST = avgNonGST;
        let finalEffectiveAvg = finalAvgGST > 0 ? finalAvgGST : finalAvgNonGST;
        master.getRange(row, 4, 1, 4).setValues([[ Math.max(0, newQty), Number(finalAvgGST.toFixed(2)), Number(finalAvgNonGST.toFixed(2)), Number((Math.max(0, newQty) * finalEffectiveAvg).toFixed(2)) ]]);

        splitRows.forEach((split, splitIdx) => {
          const splitDiscount = splitIdx === 0 ? (Number(item.discount) || 0) : 0;
          const splitSaleAmt = Number(((split.qty * Number(item.price)) - splitDiscount).toFixed(2));
          const splitFreight = splitIdx === 0 ? itemFreight : 0;
          const freightDeduction = isRecovered ? 0 : splitFreight;
          const splitExtras = split.extras || 0;
          const splitProfit = split.costPrice > 0
            ? Number((splitSaleAmt - (split.qty * split.costPrice) - freightDeduction - splitExtras).toFixed(2))
            : Number((splitSaleAmt - freightDeduction - splitExtras).toFixed(2));

          log.appendRow([ entryDate, data.salesPerson, data.padNo, data.billNo, partyNameUpper, item.mill, item.product, data.warehouse, item.billType, split.qty, item.price, splitDiscount, splitSaleAmt, splitFreight, data.freightRecovered, split.costPrice, splitProfit ]);
          logAudit('CREATE', sheetName, log.getLastRow(), null, { Mill: item.mill, Product: item.product, Qty: split.qty, Rate: item.price, BillType: item.billType }, 'Stock Out');
        });

        SpreadsheetApp.flush();
        recomputeMasterFromLogs(ss, item.mill, item.product, data.warehouse);
        recalcSalesLogProfit(ss, item.mill, item.product, data.warehouse);
      });
      updatePartyMaster(ss, partyNameUpper);
      return { status: 'success', message: 'Dispatch Successful!' };
    }

    if (data.action === 'TRANSFER') {
      const master = ss.getSheetByName('MasterStock');
      const log = getOrCreateSheet(ss, 'TransferLog', ['Date','Mill','Product','From','To','Qty','Rate']);

      if (!data.toWarehouse) throw new Error('Destination warehouse missing!');
      if (data.warehouse === data.toWarehouse) throw new Error("Source and destination can't be same!");

      data.items.forEach(item => {
        const tQty = Number(item.qty);
        if (tQty <= 0) throw new Error(`Invalid qty for ${item.product}`);
        const tRate = Number(item.rate) || Number(item.price) || 0;
        if (tRate <= 0) throw new Error(`Rate zaroori hai: ${item.product}`);

        const srcRow = findProductRowWithWarehouse(master, item.mill, item.product, data.warehouse);
        if (srcRow !== -1) {
          const availableStock = Number(master.getRange(srcRow, 4).getValue()) || 0;
          if (availableStock < tQty) throw new Error(`Stock kam hai: ${item.product} (available: ${availableStock})`);
        }

        log.appendRow([entryDate, item.mill, item.product, data.warehouse, data.toWarehouse, tQty, tRate]);
        logAudit('CREATE', 'TransferLog', log.getLastRow(), null, { Mill: item.mill, Product: item.product, From: data.warehouse, To: data.toWarehouse, Qty: tQty, Rate: tRate }, 'Transfer');

        SpreadsheetApp.flush();
        recomputeMasterFromLogs(ss, item.mill, item.product, data.warehouse);
        recomputeMasterFromLogs(ss, item.mill, item.product, data.toWarehouse);
      });
      return { status: 'success', message: 'Transfer Successful!' };
    }

    if (data.action === 'STOCK_IN') {
      const sheetName = `PurchaseLog_${data.warehouse}`;
      const headers = ['Date','Mill','Product','Warehouse','Unit Type','Qty','Price','Amount','Bill Type'];
      if (data.warehouse === 'By Truck') headers.push('Insurance', 'Cartage', 'Loading', 'Extras Claimed');
      const log = getOrCreateSheet(ss, sheetName, headers);

      if (data.warehouse === 'By Truck') {
        const existingHeaders = log.getRange(1, 1, 1, log.getLastColumn()).getValues()[0].map(x => String(x||'').trim());
        if (!existingHeaders.includes('Extras Claimed')) {
          log.getRange(1, existingHeaders.length + 1).setValue('Extras Claimed').setFontWeight('bold');
          const lastRow = log.getLastRow();
          if (lastRow > 1) {
            const col = existingHeaders.length + 1;
            const fill = [];
            for (let r = 2; r <= lastRow; r++) fill.push(['No']);
            log.getRange(2, col, lastRow - 1, 1).setValues(fill);
          }
        }
      }

      data.items.forEach(item => {
        const itemBillType = item.billType || data.billType;
        const addedQty = Number(item.qty) || 0, buyPrice = Number(item.price) || 0;
        if (buyPrice < 0) throw new Error(`Invalid price for ${item.product}`);
        if (addedQty < 0) throw new Error(`Invalid qty for ${item.product}`);
        const amt = addedQty * buyPrice;
        const rowData = [ entryDate, item.mill, item.product, data.warehouse, item.unit || 'Ream', addedQty, buyPrice, amt, itemBillType ];
        if (data.warehouse === 'By Truck') rowData.push(Number(item.insurance)||0, Number(item.cartage)||0, Number(item.loading)||0, 'No');

        log.appendRow(rowData);
        logAudit('CREATE', sheetName, log.getLastRow(), null, { Mill: item.mill, Product: item.product, Qty: addedQty, Price: buyPrice }, 'Stock In');
        SpreadsheetApp.flush();
        recomputeMasterFromLogs(ss, item.mill, item.product, data.warehouse);
      });
      return { status: 'success', message: 'Stock In Successful!' };
    }

    throw new Error('Unknown action: ' + data.action);
  } catch (err) {
    return { status: 'error', message: err.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function updateRecord(sheetName, rowNumber, updatedValues) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { status: 'error', message: 'Sheet not found!' };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    return { status: 'error', message: '⏳ System busy hai — 10 sec baad try karein.' };
  }

  try {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const currentRow = sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
    const parsed = typeof updatedValues === 'string' ? JSON.parse(updatedValues) : updatedValues;
    const oldData = {}; headers.forEach((h, i) => { oldData[h] = currentRow[i]; });

    const millIdx = headers.indexOf('Mills') !== -1 ? headers.indexOf('Mills') : headers.indexOf('Mill');
    const prodIdx = headers.indexOf('Product'), whIdx = headers.indexOf('Warehouse'), fromIdx = headers.indexOf('From'), toIdx = headers.indexOf('To');

    const oldMill = millIdx >= 0 ? (currentRow[millIdx]||'').toString().trim() : '', oldProd = prodIdx >= 0 ? (currentRow[prodIdx]||'').toString().trim() : '';
    const oldWh = whIdx >= 0 ? (currentRow[whIdx]||'').toString().trim() : '', oldFrom = fromIdx >= 0 ? (currentRow[fromIdx]||'').toString().trim() : '', oldTo = toIdx >= 0 ? (currentRow[toIdx]||'').toString().trim() : '';

    const newRow = headers.map((h, i) => {
      if (parsed.hasOwnProperty(h)) return h === 'Date' ? currentRow[i] : parsed[h];
      return currentRow[i];
    });

    const newMill = millIdx >= 0 ? (newRow[millIdx]||'').toString().trim() : '', newProd = prodIdx >= 0 ? (newRow[prodIdx]||'').toString().trim() : '';
    const newWh = whIdx >= 0 ? (newRow[whIdx]||'').toString().trim() : '', newFrom = fromIdx >= 0 ? (newRow[fromIdx]||'').toString().trim() : '', newTo = toIdx >= 0 ? (newRow[toIdx]||'').toString().trim() : '';

    if (sheetName.startsWith('SalesLog')) {
      const qi = headers.indexOf('Qty'), ri = headers.indexOf('Rate'), di = headers.indexOf('Discount'), ai = headers.indexOf('Amount');
      const cpi = headers.indexOf('Cost Price'), fi = headers.indexOf('Profit'), fri = headers.indexOf('Freight'), reci = headers.indexOf('Recovered?');

      let qty = 0, rate = 0, disc = 0, amt = 0, cp = 0, freight = 0, isRec = false;
      if (qi >= 0) qty = Number(newRow[qi]) || 0;
      if (ri >= 0) rate = Number(newRow[ri]) || 0;
      if (di >= 0) disc = Number(newRow[di]) || 0;

      amt = (qty * rate) - disc;
      if (ai >= 0) newRow[ai] = amt;

      if (fri >= 0) freight = Number(newRow[fri]) || 0;
      if (reci >= 0) isRec = String(newRow[reci]).toLowerCase().trim() === 'yes';
      if (cpi >= 0) cp = Number(newRow[cpi]) || 0;

      const fDed = isRec ? 0 : freight;
      const whForExtras = newWh || oldWh;
      const millForExtras = newMill || oldMill;
      const prodForExtras = newProd || oldProd;
      let extras = 0;
      if (whForExtras === 'By Truck') {
        const pLog = ss.getSheetByName(`PurchaseLog_${whForExtras}`);
        if (pLog) {
          const pd = pLog.getDataRange().getValues(), ph = pd[0].map(x => String(x||'').trim().toLowerCase());
          const pMi = ph.indexOf('mills') !== -1 ? ph.indexOf('mills') : ph.indexOf('mill');
          const pPi = ph.indexOf('product'), pQi = ph.indexOf('qty');
          const insI = ph.indexOf('insurance'), carI = ph.indexOf('cartage'), lodI = ph.indexOf('loading');
          for (let i = 1; i < pd.length; i++) {
            if (!ciEquals(pd[i][pMi], millForExtras) || !ciEquals(pd[i][pPi], prodForExtras)) continue;
            const q = Number(pd[i][pQi]) || 0;
            if (q <= 0) continue;
            extras += (insI >= 0 ? Number(pd[i][insI])||0 : 0)
                    + (carI >= 0 ? Number(pd[i][carI])||0 : 0)
                    + (lodI >= 0 ? Number(pd[i][lodI])||0 : 0);
          }
        }
      }
      const profit = amt - (qty * cp) - fDed - extras;
      if (fi >= 0) newRow[fi] = Number(profit.toFixed(2));
    }

    if (sheetName.startsWith('PurchaseLog')) {
      const qi = headers.indexOf('Qty'), pi2 = headers.indexOf('Price'), ai = headers.indexOf('Amount');
      if (qi >= 0 && pi2 >= 0 && ai >= 0) newRow[ai] = (Number(newRow[qi])||0) * (Number(newRow[pi2])||0);
    }

    sheet.getRange(rowNumber, 1, 1, newRow.length).setValues([newRow]);
    SpreadsheetApp.flush();

    const newData = {}; headers.forEach((h, i) => { newData[h] = newRow[i]; });
    logAudit('UPDATE', sheetName, rowNumber, oldData, newData, 'Manual edit via Modify tab');

    const combos = new Set();
    if (sheetName.startsWith('SalesLog') || sheetName.startsWith('PurchaseLog')) {
      if (oldMill && oldProd && oldWh) combos.add(`${oldMill}||${oldProd}||${oldWh}`);
      const fm = newMill || oldMill, fw = newWh || oldWh, fp = newProd || oldProd;
      if (fm && fp && fw) combos.add(`${fm}||${fp}||${fw}`);
    }
    if (sheetName === 'TransferLog') {
      if (oldMill && oldProd) { if (oldFrom) combos.add(`${oldMill}||${oldProd}||${oldFrom}`); if (oldTo) combos.add(`${oldMill}||${oldProd}||${oldTo}`); }
      const fm = newMill || oldMill, fp = newProd || oldProd;
      if (fm && fp) { if (newFrom) combos.add(`${fm}||${fp}||${newFrom}`); if (newTo) combos.add(`${fm}||${fp}||${newTo}`); }
    }

    combos.forEach(key => { const [m, p, w] = key.split('||'); recomputeMasterFromLogs(ss, m, p, w); recalcSalesLogProfit(ss, m, p, w); });
    if (sheetName.startsWith('SalesLog') || combos.size === 0) rebuildAllMasterSilent(ss);

    return { status: 'success', message: `✅ Updated! Sync successful` };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function deleteRecord(sheetName, rowNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { status: 'error', message: 'Sheet not found!' };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    return { status: 'error', message: '⏳ System busy hai — 10 sec baad try karein.' };
  }

  try {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowData = sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
    sheet.deleteRow(rowNumber);
    SpreadsheetApp.flush();

    const oldData = {}; headers.forEach((h, i) => { oldData[h] = rowData[i]; });
    logAudit('DELETE', sheetName, rowNumber, oldData, null, 'Row deleted via Modify tab');
    shiftAuditRowsAfterDelete(ss, sheetName, rowNumber);

    const millIdx = headers.indexOf('Mills') !== -1 ? headers.indexOf('Mills') : headers.indexOf('Mill');
    const prodIdx = headers.indexOf('Product'), whIdx = headers.indexOf('Warehouse'), fromIdx = headers.indexOf('From'), toIdx = headers.indexOf('To');
    const mill = millIdx >= 0 ? (rowData[millIdx]||'').toString().trim() : '', prod = prodIdx >= 0 ? (rowData[prodIdx]||'').toString().trim() : '';
    const wh = whIdx >= 0 ? (rowData[whIdx]||'').toString().trim() : '', fromWh = fromIdx >= 0 ? (rowData[fromIdx]||'').toString().trim() : '', toWh = toIdx >= 0 ? (rowData[toIdx]||'').toString().trim() : '';

    const combos = new Set();
    if (sheetName.startsWith('SalesLog') || sheetName.startsWith('PurchaseLog')) { if (mill && prod && wh) combos.add(`${mill}||${prod}||${wh}`); }
    if (sheetName === 'TransferLog') { if (mill && prod) { if (fromWh) combos.add(`${mill}||${prod}||${fromWh}`); if (toWh) combos.add(`${mill}||${prod}||${toWh}`); } }

    combos.forEach(key => { const [m, p, w] = key.split('||'); recomputeMasterFromLogs(ss, m, p, w); recalcSalesLogProfit(ss, m, p, w); });
    return { status: 'success', message: `✅ Deleted successfully.` };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function getLogsByDateRange(baseSheetName, fromDate, toDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let results = [];
  const sheetsToSearch = baseSheetName === 'TransferLog' ? ['TransferLog'] : WAREHOUSES.map(w => `${baseSheetName}_${w}`);
  const fromStr = Utilities.formatDate(new Date(fromDate), 'GMT+5:30', 'yyyy-MM-dd');
  const toStr = Utilities.formatDate(new Date(toDate), 'GMT+5:30', 'yyyy-MM-dd');

  sheetsToSearch.forEach(sName => {
    const sheet = ss.getSheetByName(sName);
    if(sheet) {
      const d = sheet.getDataRange().getValues();
      if(d.length > 1) {
        const h = d[0];
        for(let i=1; i<d.length; i++) {
          if(!d[i][0]) continue;
          try {
            const dt = Utilities.formatDate(new Date(d[i][0]), 'GMT+5:30', 'yyyy-MM-dd');
            if (dt >= fromStr && dt <= toStr) {
              const obj = { _sheetRow: i+1, _sheetName: sName };
              h.forEach((head, j) => { if(head) obj[head] = d[i][j] instanceof Date ? Utilities.formatDate(d[i][j], 'GMT+5:30', 'dd-MM-yyyy') : d[i][j]; });
              results.push(obj);
            }
          } catch(e) {}
        }
      }
    }
  });
  return JSON.stringify(results);
}

function getTodayLog(baseSheetName) {
  const t = Utilities.formatDate(new Date(), 'GMT+5:30', 'yyyy-MM-dd');
  return getLogsByDateRange(baseSheetName, t, t);
}

function searchRecords(baseSheetName, searchParty, searchProduct, fromDate, toDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let results = [];
  const sheetsToSearch = baseSheetName === 'TransferLog' ? ['TransferLog'] : WAREHOUSES.map(w => `${baseSheetName}_${w}`);
  const pLow = (searchParty || '').toLowerCase().trim(), prLow = (searchProduct || '').toLowerCase().trim();
  let fromStr = '', toStr = '';
  if (fromDate && toDate) { fromStr = Utilities.formatDate(new Date(fromDate), 'GMT+5:30', 'yyyy-MM-dd'); toStr = Utilities.formatDate(new Date(toDate), 'GMT+5:30', 'yyyy-MM-dd'); }

  sheetsToSearch.forEach(sName => {
    const sheet = ss.getSheetByName(sName);
    if(sheet) {
      const d = sheet.getDataRange().getValues();
      if(d.length > 1) {
        const h = d[0];
        for(let i=1; i<d.length; i++) {
          const rowStr = d[i].map(c => c.toString().toLowerCase()).join(' ');
          if ((!pLow || rowStr.includes(pLow)) && (!prLow || rowStr.includes(prLow))) {
            let inDate = true;
            if (fromStr && toStr && d[i][0]) {
              try { const dt = Utilities.formatDate(new Date(d[i][0]), 'GMT+5:30', 'yyyy-MM-dd'); if (dt < fromStr || dt > toStr) inDate = false; } catch(e) { inDate = false; }
            }
            if (inDate) {
              const obj = { _sheetRow: i+1, _sheetName: sName };
              h.forEach((head, j) => { if(head) obj[head] = d[i][j] instanceof Date ? Utilities.formatDate(d[i][j], 'GMT+5:30', 'dd-MM-yyyy') : d[i][j]; });
              results.push(obj);
            }
          }
        }
      }
    }
  });
  return JSON.stringify(results.slice(-100));
}

function getProductsFromSheet(baseSheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToSearch = baseSheetName === 'TransferLog' ? ['TransferLog'] : WAREHOUSES.map(w => `${baseSheetName}_${w}`);
  const map = new Map();
  sheetsToSearch.forEach(sName => {
    const sheet = ss.getSheetByName(sName);
    if (sheet) {
      const data = sheet.getDataRange().getValues(), ci = data[0].findIndex(h => h.toString().toLowerCase().includes('product'));
      if (ci !== -1) {
        for (let i = 1; i < data.length; i++) { if (data[i][ci]) { const v = data[i][ci].toString().trim(); map.set(v.toLowerCase().replace(/\s+/g, ' '), v); } }
      }
    }
  });
  return JSON.stringify(Array.from(map.values()));
}

function getPartiesFromSheet(baseSheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToSearch = baseSheetName === 'TransferLog' ? ['TransferLog'] : WAREHOUSES.map(w => `${baseSheetName}_${w}`);
  const map = new Map();
  sheetsToSearch.forEach(sName => {
    const sheet = ss.getSheetByName(sName);
    if (sheet) {
      const data = sheet.getDataRange().getValues(), headers = data[0];
      let ci = -1;
      if (sName.startsWith('SalesLog')) ci = headers.indexOf('Party Name');
      else if (sName.startsWith('PurchaseLog')) { ci = headers.indexOf('Mill'); if (ci === -1) ci = headers.indexOf('Mills'); }
      if (ci !== -1) {
        for (let i = 1; i < data.length; i++) { if (data[i][ci]) { const v = data[i][ci].toString().trim(); map.set(v.toLowerCase().replace(/\s+/g, ' '), v); } }
      }
    }
  });
  return JSON.stringify(Array.from(map.values()));
}

function getProductMasterList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ProductMaster');
  if (!sheet) return JSON.stringify([]);
  const data = sheet.getDataRange().getValues(), map = new Map();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] && data[i][2]) {
      const m = data[i][1].toString().trim(), p = data[i][2].toString().trim();
      const k = m.toLowerCase().replace(/\s+/g, ' ') + '||' + p.toLowerCase().replace(/\s+/g, ' ');
      if (!map.has(k)) map.set(k, { mill: m, name: p, baseUnit: data[i][3] || 'Ream', label: m + ' - ' + p });
    }
  }
  return JSON.stringify(Array.from(map.values()));
}

function getDashboardData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MasterStock');
  if (!sheet) return JSON.stringify({ items: [], totalValue: 0 });
  const data = sheet.getDataRange().getValues();
  let items = [], grandTotal = 0;
  for (let i = 1; i < data.length; i++) {
    grandTotal += Number(data[i][6]) || 0;
    items.push({ mill: data[i][0], product: data[i][1], warehouse: data[i][2], qty: data[i][3], avgGST: data[i][4], avgNonGST: data[i][5], totalVal: data[i][6], unit: data[i][7] });
  }
  return JSON.stringify({ items, totalValue: grandTotal });
}

function getTallyStockReport(fromDate, toDate, filterWh, filterMill, filterProduct) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fDate = new Date(fromDate), tDate = new Date(toDate);
  const report = {};

  function sanitize(str) { return String(str||'').toLowerCase().replace(/\s+/g, ' ').trim(); }
  function initItem(key, mill, prod) { if (!report[key]) report[key] = { mill: mill.trim(), product: prod.trim(), opQty: 0, inQty: 0, outQty: 0 }; }

  const fMill = sanitize(filterMill), fProd = sanitize(filterProduct);

  WAREHOUSES.forEach(wh => {
    if (filterWh && filterWh !== 'All' && !ciEquals(wh, filterWh)) return;
    const pSheet = ss.getSheetByName(`PurchaseLog_${wh}`);
    if (pSheet) {
      const d = pSheet.getDataRange().getValues(), pMi = d[0].indexOf('Mill') !== -1 ? d[0].indexOf('Mill') : d[0].indexOf('Mills'), pPi = d[0].indexOf('Product'), pQi = d[0].indexOf('Qty');
      if (pMi !== -1) {
        for (let i = 1; i < d.length; i++) {
          const date = d[i][0], mill = d[i][pMi], prod = d[i][pPi], qty = d[i][pQi];
          if (!date || (fMill && sanitize(mill) !== fMill) || (fProd && sanitize(prod) !== fProd)) continue;
          const key = sanitize(mill) + '||' + sanitize(prod), dt = new Date(date);
          initItem(key, mill, prod);
          if (dt < fDate) report[key].opQty += Number(qty)||0; else if (dt <= tDate) report[key].inQty += Number(qty)||0;
        }
      }
    }
    const sSheet = ss.getSheetByName(`SalesLog_${wh}`);
    if (sSheet) {
      const d = sSheet.getDataRange().getValues(), sMi = d[0].indexOf('Mill') !== -1 ? d[0].indexOf('Mill') : d[0].indexOf('Mills'), sPi = d[0].indexOf('Product'), sQi = d[0].indexOf('Qty');
      if (sMi !== -1) {
        for (let i = 1; i < d.length; i++) {
          const date = d[i][0], mill = d[i][sMi], prod = d[i][sPi], qty = d[i][sQi];
          if (!date || (fMill && sanitize(mill) !== fMill) || (fProd && sanitize(prod) !== fProd)) continue;
          const key = sanitize(mill) + '||' + sanitize(prod), dt = new Date(date);
          initItem(key, mill, prod);
          if (dt < fDate) report[key].opQty -= Number(qty)||0; else if (dt <= tDate) report[key].outQty += Number(qty)||0;
        }
      }
    }
  });

  const tSheet = ss.getSheetByName('TransferLog');
  if (tSheet && filterWh && filterWh !== 'All') {
    const d = tSheet.getDataRange().getValues(), tMi = d[0].indexOf('Mill') !== -1 ? d[0].indexOf('Mill') : d[0].indexOf('Mills');
    if (tMi !== -1) {
      const th = d[0].map(x=>String(x).toLowerCase()), fIdx = th.indexOf('from'), tIdx = th.indexOf('to'), qIdx = th.indexOf('qty'), pIdx = th.indexOf('product');
      for (let i = 1; i < d.length; i++) {
        const date = d[i][0], mill = d[i][tMi], prod = d[i][pIdx], fromWh = d[i][fIdx], toWh = d[i][tIdx], qty = d[i][qIdx];
        if (!date || (fMill && sanitize(mill) !== fMill) || (fProd && sanitize(prod) !== fProd)) continue;
        const key = sanitize(mill) + '||' + sanitize(prod), dt = new Date(date), q = Number(qty)||0;
        initItem(key, mill, prod);
        if (ciEquals(fromWh, filterWh)) { if (dt < fDate) report[key].opQty -= q; else if (dt <= tDate) report[key].outQty += q; }
        if (ciEquals(toWh,   filterWh)) { if (dt < fDate) report[key].opQty += q; else if (dt <= tDate) report[key].inQty  += q; }
      }
    }
  }

  return JSON.stringify(Object.values(report).map(item => {
    const clQty = item.opQty + item.inQty - item.outQty;
    return { mill: item.mill, product: item.product, opQty: item.opQty, opAmt: 0, inQty: item.inQty, inAmt: 0, outQty: item.outQty, outAmt: 0, clQty, clAmt: 0 };
  }));
}

function getDailyStockLedger(fromDate, toDate, filterWh, filterMill, filterProduct) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), start = new Date(fromDate), end = new Date(toDate);
  const masterData = (ss.getSheetByName('MasterStock') || { getDataRange: () => ({ getValues: () => [] }) }).getDataRange().getValues();
  const allWHS = [...new Set(masterData.slice(1).map(r => r[2]).filter(Boolean))];
  const TWH = (filterWh === 'All' || !filterWh) ? allWHS : [filterWh];
  const opWh = {}, daily = {}; TWH.forEach(w => { opWh[w] = 0; });
  function ensureDay(ds) { if (!daily[ds]) { daily[ds] = {}; TWH.forEach(w => { daily[ds][w] = { inQty: 0, outQty: 0 }; }); } }
  function findWh(w) { return TWH.find(x => ciEquals(x, w)); }

  TWH.forEach(wh => {
    const pSheet = ss.getSheetByName(`PurchaseLog_${wh}`);
    if (pSheet) {
      const d = pSheet.getDataRange().getValues(), pMi = d[0].indexOf('Mill') !== -1 ? d[0].indexOf('Mill') : d[0].indexOf('Mills');
      if (pMi !== -1) {
        d.slice(1).forEach(r => {
          if (!r[0] || !ciEquals(r[pMi], filterMill) || !ciEquals(r[d[0].indexOf('Product')], filterProduct)) return;
          const dt = new Date(r[0]), qty = Number(r[d[0].indexOf('Qty')])||0;
          if (dt < start) opWh[wh] += qty;
          else if (dt <= end) { const ds = Utilities.formatDate(dt,'GMT+5:30','yyyy-MM-dd'); ensureDay(ds); daily[ds][wh].inQty += qty; }
        });
      }
    }
    const sSheet = ss.getSheetByName(`SalesLog_${wh}`);
    if (sSheet) {
      const d = sSheet.getDataRange().getValues(), sMi = d[0].indexOf('Mill') !== -1 ? d[0].indexOf('Mill') : d[0].indexOf('Mills');
      if (sMi !== -1) {
        d.slice(1).forEach(r => {
          if (!r[0] || !ciEquals(r[sMi], filterMill) || !ciEquals(r[d[0].indexOf('Product')], filterProduct)) return;
          const dt = new Date(r[0]), qty = Number(r[d[0].indexOf('Qty')])||0;
          if (dt < start) opWh[wh] += qty;
          else if (dt <= end) { const ds = Utilities.formatDate(dt,'GMT+5:30','yyyy-MM-dd'); ensureDay(ds); daily[ds][wh].outQty += qty; }
        });
      }
    }
  });

  const tSheet = ss.getSheetByName('TransferLog');
  if(tSheet) {
    const d = tSheet.getDataRange().getValues(), tMi = d[0].indexOf('Mill') !== -1 ? d[0].indexOf('Mill') : d[0].indexOf('Mills');
    if (tMi !== -1) {
      const th = d[0].map(x=>String(x).toLowerCase()), fIdx = th.indexOf('from'), tIdx = th.indexOf('to'), qIdx = th.indexOf('qty'), pIdx = th.indexOf('product');
      d.slice(1).forEach(r => {
        if (!r[0] || !ciEquals(r[tMi], filterMill) || !ciEquals(r[pIdx], filterProduct)) return;
        const fwh = findWh(r[fIdx]), twh = findWh(r[tIdx]), q = Number(r[qIdx])||0, dt = new Date(r[0]);
        if (fwh) { if (dt < start) opWh[fwh] -= q; else if (dt <= end) { const ds = Utilities.formatDate(dt,'GMT+5:30','yyyy-MM-dd'); ensureDay(ds); daily[ds][fwh].outQty += q; } }
        if (twh) { if (dt < start) opWh[twh] += q; else if (dt <= end) { const ds = Utilities.formatDate(dt,'GMT+5:30','yyyy-MM-dd'); ensureDay(ds); daily[ds][twh].inQty  += q; } }
      });
    }
  }

  const running = Object.assign({}, opWh), result = [];
  Object.keys(daily).sort().forEach(ds => {
    TWH.forEach(wh => {
      const mv = daily[ds][wh];
      if (mv.inQty > 0 || mv.outQty > 0) {
        const opQty = running[wh] || 0;
        running[wh] = opQty + mv.inQty - mv.outQty;
        result.push({ date: Utilities.formatDate(new Date(ds),'GMT+5:30','dd-MM-yyyy'), dateSort: ds, warehouse: wh, opQty, inQty: mv.inQty, outQty: mv.outQty, clQty: running[wh] });
      }
    });
  });
  return JSON.stringify(result);
}

function findProductRowWithWarehouse(sheet, mill, product, warehouse) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { if (ciEquals(data[i][0], mill) && ciEquals(data[i][1], product) && ciEquals(data[i][2], warehouse)) return i + 1; }
  return -1;
}

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) { sheet = ss.insertSheet(name); sheet.appendRow(headers); sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold'); }
  return sheet;
}

function getPartyList() {
  const sheet = getOrCreateSheet(SpreadsheetApp.getActiveSpreadsheet(), 'PartyMaster', ['Party Name']);
  return JSON.stringify(sheet.getDataRange().getValues().slice(1).map(r => r[0]));
}

function updatePartyMaster(ss, name) {
  const upper = (name || '').toString().toUpperCase().trim();
  if (!upper) return;
  const sheet = getOrCreateSheet(ss, 'PartyMaster', ['Party Name']);
  const data = sheet.getDataRange().getValues().flat().map(v => v.toString().toUpperCase().trim());
  if (!data.includes(upper)) sheet.appendRow([upper]);
}
