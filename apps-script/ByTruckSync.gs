// ========================================================================================
// 1. PURCHASE SOURCE SHEET KA URL YA ID:
// (Jis Google Sheet ke andar "Orders (Readable)" tab hai)
// ========================================================================================
const PURCHASE_SOURCE_SHEET_URL_OR_ID = "https://docs.google.com/spreadsheets/d/1HuZABEjx7A68gEcJd_BmQhjTdCqX-UpVdvm51r3YhvA/edit?gid=1991998043#gid=1991998043";

// ========================================================================================
// 2. SALES (BY TRUCK) SOURCE SHEET KA URL YA ID:
// (Jis Google Sheet ke andar "Workflow" tab hai)
// ========================================================================================
const SALES_SOURCE_SHEET_URL_OR_ID = "https://docs.google.com/spreadsheets/d/1STKk27ovhhTZYEH9Fv0FdFE81u_06C5VI2K-mipH3Lo/edit?gid=668464145#gid=668464145";

// ========================================================================================
// 3. WMS SPREADSHEET ID:
// (Agar script WMS sheet me hi attach hai to ye auto-detect ho jata hai, warna ye ID use hogi)
// ========================================================================================
const WMS_SHEET_ID = "1J0NHuUWHj9ZtXkyVE9Jad8D1d6eMnTPMd9vtFahYjMo";

// Legacy fallback (agar kisi purane reference me ORDERS_SOURCE_SHEET_URL_OR_ID use ho)
const ORDERS_SOURCE_SHEET_URL_OR_ID = PURCHASE_SOURCE_SHEET_URL_OR_ID;

const BY_TRUCK_SYNC_CONFIG = {
  purchaseSourceProperty: 'BY_TRUCK_PURCHASE_SOURCE_ID',
  salesSourceProperty: 'BY_TRUCK_SALES_SOURCE_ID',
  sourceSpreadsheetProperty: 'BY_TRUCK_SOURCE_SPREADSHEET_ID',
  wmsSpreadsheetProperty: 'BY_TRUCK_WMS_SPREADSHEET_ID',
  purchaseSourceSheet: 'Orders (Readable)',
  purchaseTargetSheet: 'PurchaseLog_By Truck',
  salesSourceSheet: 'Workflow',
  salesTargetSheet: 'SalesLog_By Truck',
  productWeightsSheet: 'ProductWeights',
  purchaseStartRow: 273,
  salesStartRow: 171,
  sourceKeyHeader: '__WMS_SOURCE_KEY',
  purchaseSourceColumns: {
    po: 0,
    product: 3,
    ratePerKg: 6,
    qtyKg: 7,
    date: 22,
    status: 25,
    insurance: 28,
    freight: 29
  },
  salesSourceColumns: {
    po: 1,
    party: 18,
    salesPerson: 23,
    billNo: 34,
    date: 39,
    product: 56,
    qtyReam: 57,
    ratePerKg: 60
  }
};

function syncByTruckData() {
  let lock = null;
  if (typeof LockService !== 'undefined' && LockService.getScriptLock) {
    lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) {
      Logger.log('[syncByTruckData] Busy or locked, skipping tick.');
      return;
    }
  }
  try {
    try {
      fixBiltCopierPlainLabels();
    } catch (e) {
      Logger.log('fixBiltCopierPlainLabels notice: ' + e.message);
    }
    syncByTruckPurchases_();
    syncByTruckSales_();
    try {
      const books = getByTruckSpreadsheets_();
      dedupeByTruckSheet_(books.wms.getSheetByName(BY_TRUCK_SYNC_CONFIG.purchaseTargetSheet), 'purchase');
      dedupeByTruckSheet_(books.wms.getSheetByName(BY_TRUCK_SYNC_CONFIG.salesTargetSheet), 'sales');
    } catch (e) {
      Logger.log('dedupe notice: ' + e.message);
    }
    try {
      if (typeof rebuildByTruckOnly === 'function') {
        rebuildByTruckOnly();
      }
    } catch (e) {
      Logger.log('rebuildByTruckOnly after sync error: ' + e.message);
    }
  } catch (err) {
    Logger.log('[syncByTruckData Error]: ' + err.message);
  } finally {
    if (lock) {
      try { lock.releaseLock(); } catch (e) {}
    }
  }
}

function installByTruckSyncTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const hasSyncClockTrigger = triggers.some(function (trigger) {
    return trigger.getHandlerFunction() === 'syncByTruckData' &&
      trigger.getEventType() === ScriptApp.EventType.CLOCK;
  });
  if (!hasSyncClockTrigger) {
    ScriptApp.newTrigger('syncByTruckData')
      .timeBased()
      .everyMinutes(10)
      .create();
  }

  const legacyHandlers = {
    autoSyncPurchaseLogByTruck: true,
    autoSyncSalesAndKheraKalan: true
  };
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (legacyHandlers[trigger.getHandlerFunction()]) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  return !hasSyncClockTrigger;
}

function syncByTruckPurchases_() {
  const books = getByTruckSpreadsheets_();
  const source = getByTruckSheet_(books.purchaseSource || books.source, BY_TRUCK_SYNC_CONFIG.purchaseSourceSheet);
  const target = getByTruckSheet_(books.wms, BY_TRUCK_SYNC_CONFIG.purchaseTargetSheet);
  const weights = getProductWeights_(books.wms);
  const millMap = getProductMasterMills_(books.wms);
  const sourceRows = source.getDataRange().getValues();
  const purchases = buildByTruckPurchases_(sourceRows, weights, millMap);
  upsertByTruckRows_(target, purchases, 'purchase');
}

function syncByTruckSales_() {
  const books = getByTruckSpreadsheets_();
  const source = getByTruckSheet_(books.salesSource || books.source, BY_TRUCK_SYNC_CONFIG.salesSourceSheet);
  const purchaseSheet = getByTruckSheet_(books.wms, BY_TRUCK_SYNC_CONFIG.purchaseTargetSheet);
  const target = getByTruckSheet_(books.wms, BY_TRUCK_SYNC_CONFIG.salesTargetSheet);
  const weights = getProductWeights_(books.wms);
  const millMap = getProductMasterMills_(books.wms);
  const costs = getByTruckPurchaseCosts_(purchaseSheet, weights);
  const sourceRows = source.getDataRange().getValues();
  const sales = buildByTruckSales_(sourceRows, weights, costs, millMap);
  upsertByTruckRows_(target, sales, 'sales');
}

function extractSpreadsheetId_(input) {
  if (!input) return '';
  const str = String(input).trim();
  const match = str.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (match && match[1]) return match[1];
  return str;
}

function getByTruckSpreadsheets_() {
  const properties = (typeof PropertiesService !== 'undefined' && PropertiesService.getScriptProperties)
    ? PropertiesService.getScriptProperties()
    : null;

  // 1. Resolve WMS Spreadsheet (Active Spreadsheet has priority if container-bound)
  let wms = null;
  if (typeof SpreadsheetApp !== 'undefined' && SpreadsheetApp.getActiveSpreadsheet) {
    try { wms = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  }
  const propWms = (properties && properties.getProperty(BY_TRUCK_SYNC_CONFIG.wmsSpreadsheetProperty)) || WMS_SHEET_ID;
  const wmsId = extractSpreadsheetId_(propWms);
  if (!wms && wmsId && wmsId !== 'wms' && typeof SpreadsheetApp !== 'undefined') {
    try { wms = SpreadsheetApp.openById(wmsId); } catch (e) {}
  }

  // 2. Resolve Purchase Source Spreadsheet
  const propPurchase = (properties && (
    properties.getProperty(BY_TRUCK_SYNC_CONFIG.purchaseSourceProperty) ||
    properties.getProperty(BY_TRUCK_SYNC_CONFIG.sourceSpreadsheetProperty)
  )) || (typeof PURCHASE_SOURCE_SHEET_URL_OR_ID !== 'undefined' ? PURCHASE_SOURCE_SHEET_URL_OR_ID : '');

  const purchaseId = extractSpreadsheetId_(propPurchase);
  let purchaseSource = null;
  if (purchaseId && typeof SpreadsheetApp !== 'undefined') {
    try { purchaseSource = SpreadsheetApp.openById(purchaseId); } catch (e) {}
  }
  if (!purchaseSource && wms && wms.getSheetByName(BY_TRUCK_SYNC_CONFIG.purchaseSourceSheet)) {
    purchaseSource = wms;
  }

  // 3. Resolve Sales Source Spreadsheet
  const propSales = (properties && (
    properties.getProperty(BY_TRUCK_SYNC_CONFIG.salesSourceProperty) ||
    properties.getProperty(BY_TRUCK_SYNC_CONFIG.sourceSpreadsheetProperty)
  )) || (typeof SALES_SOURCE_SHEET_URL_OR_ID !== 'undefined' ? SALES_SOURCE_SHEET_URL_OR_ID : '');

  const salesId = extractSpreadsheetId_(propSales);
  let salesSource = null;
  if (salesId && typeof SpreadsheetApp !== 'undefined') {
    try { salesSource = SpreadsheetApp.openById(salesId); } catch (e) {}
  }
  if (!salesSource && purchaseSource && purchaseSource.getSheetByName(BY_TRUCK_SYNC_CONFIG.salesSourceSheet)) {
    salesSource = purchaseSource;
  }
  if (!salesSource && wms && wms.getSheetByName(BY_TRUCK_SYNC_CONFIG.salesSourceSheet)) {
    salesSource = wms;
  }

  if (!purchaseSource) {
    throw new Error(
      '⚠️ Purchase Source Sheet open nahi ho saki!\n' +
      'Kripya ByTruckSync.gs me PURCHASE_SOURCE_SHEET_URL_OR_ID verify karein.'
    );
  }
  if (!salesSource) {
    throw new Error(
      '⚠️ Sales (By Truck) Source Sheet open nahi ho saki!\n' +
      'Kripya ByTruckSync.gs me SALES_SOURCE_SHEET_URL_OR_ID verify karein.'
    );
  }
  if (!wms) {
    throw new Error('⚠️ WMS Spreadsheet open nahi ho saki. Kripya WMS_SHEET_ID verify karein.');
  }

  return {
    source: purchaseSource,
    purchaseSource: purchaseSource,
    salesSource: salesSource,
    wms: wms
  };
}

function getByTruckSheet_(spreadsheet, name) {
  let sheet = spreadsheet.getSheetByName(name);
  if (sheet) return sheet;

  if (typeof spreadsheet.getSheets === 'function') {
    const sheets = spreadsheet.getSheets();
    const targetNorm = normalizeText_(name);
    for (let i = 0; i < sheets.length; i++) {
      if (normalizeText_(sheets[i].getName()) === targetNorm) {
        return sheets[i];
      }
    }
  }

  throw new Error('Required sheet not found: ' + name);
}

function getProductWeights_(spreadsheet) {
  let sheet = null;
  try {
    sheet = getByTruckSheet_(spreadsheet, BY_TRUCK_SYNC_CONFIG.productWeightsSheet);
  } catch (e) {
    if (typeof spreadsheet.insertSheet === 'function') {
      try {
        sheet = spreadsheet.insertSheet(BY_TRUCK_SYNC_CONFIG.productWeightsSheet);
      } catch (err) {}
    }
  }

  const defaultSeeds = [
    ['DOUBLE A A4 70 GSM', 2.18],
    ['DOUBLE A A4 75 GSM', 2.34],
    ['DOUBLE A A4 80 GSM', 2.50],
    ['DOUBLE A A4 100 GSM', 3.12],
    ['N R COPIER A4 70 GSM', 2.18],
    ['N R COPIER A4 75 GSM', 2.34],
    ['N R COPIER A4 80 GSM', 2.50],
    ['N R COPIER A3 70 GSM', 4.37],
    ['N R COPIER A3 75 GSM', 4.68],
    ['KHANNA A4 KBOLT (S) 65 GSM', 2.03],
    ['KHANNA A4 EPRINT (S) 70 GSM', 2.18],
    ['KHANNA A4 ECOPY (S) 75 GSM', 2.34],
    ['KHANNA A3 KBOLT (S) 65 GSM', 4.05],
    ['SATIA A4 65 GSM', 2.03],
    ['SATIA A4 70 GSM', 2.18],
    ['SATIA A4 75 GSM', 2.34],
    ['SATIA A3 65 GSM', 4.05],
    ['SATIA A3 70 GSM', 4.37],
    ['SATIA FS 70 GSM', 2.60],
    ['ROZAANA A4 70 GSM', 2.18],
    ['ROZAANA A4 75 GSM', 2.34],
    ['I K COPY A4 70 GSM', 2.18],
    ['PP LITE A4 70 GSM', 2.18],
    ['PAPER ONE A4 70 GSM', 2.18],
    ['PAPER ONE A4 75 GSM', 2.34],
    ['PAPER ONE A4 100 GSM', 3.12],
    ['PAPER ONE A3 100 GSM', 6.24],
    ['MAPLE A4 70 GSM', 2.18],
    ['SPECTRA A4 75 GSM', 2.34],
    ['SPECTRA FS 75 GSM', 2.78],
    ['CENTURY STAR A4 75 GSM', 2.34],
    ['J K RED A4 75 GSM', 2.34]
  ];

  const weights = {};
  defaultSeeds.forEach(function (pair) {
    weights[normalizeText_(pair[0])] = pair[1];
  });

  if (!sheet) return weights;

  const values = sheet.getDataRange ? sheet.getDataRange().getValues() : [];
  if (!values || values.length === 0) {
    try {
      const rowsToInsert = [['Product', 'Kg per Ream']].concat(defaultSeeds);
      sheet.getRange(1, 1, rowsToInsert.length, 2).setValues(rowsToInsert);
    } catch (e) {}
    return weights;
  }

  const productAliases = [
    'product', 'productname', 'products', 'item', 'items', 'itemname',
    'description', 'itemdescription', 'particulars', 'particular', 'name', 'paper'
  ];
  const weightAliases = [
    'kgperream', 'weightperream', 'kgream', 'weight', 'weights', 'wt',
    'reamweight', 'weightkg', 'wtkg', 'kg', 'perream', 'weightream',
    'reamwt', 'weightinkg', 'weightkgream', 'weightperreamkg', 'kgperreamkg',
    'perreamweight', 'perreamwt', 'kgperbox', 'weightperbox'
  ];

  let headerRowIndex = -1;
  let productColumn = -1;
  let weightColumn = -1;

  for (let r = 0; r < Math.min(values.length, 15); r++) {
    const row = values[r];
    const pCol = findHeaderColumn_(row, productAliases);
    const wCol = findHeaderColumn_(row, weightAliases);
    if (pCol >= 0 && wCol >= 0) {
      headerRowIndex = r;
      productColumn = pCol;
      weightColumn = wCol;
      break;
    }
  }

  if (headerRowIndex < 0) {
    for (let r = 0; r < Math.min(values.length, 15); r++) {
      const row = values[r];
      const pCol = findHeaderColumn_(row, productAliases);
      if (pCol >= 0) {
        headerRowIndex = r;
        productColumn = pCol;
        weightColumn = (pCol === 0 && row.length > 1) ? 1 : (pCol > 0 ? 0 : 1);
        break;
      }
    }
  }

  let foundAny = false;
  if (headerRowIndex >= 0 && productColumn >= 0 && weightColumn >= 0) {
    for (let i = headerRowIndex + 1; i < values.length; i++) {
      const product = normalizeText_(values[i][productColumn]);
      const weight = toNumber_(values[i][weightColumn]);
      if (product && weight > 0) {
        weights[product] = weight;
        foundAny = true;
      }
    }
  }

  if (!foundAny && values.length <= 1) {
    try {
      const rowsToInsert = [['Product', 'Kg per Ream']].concat(defaultSeeds);
      sheet.getRange(1, 1, rowsToInsert.length, 2).setValues(rowsToInsert);
    } catch (e) {}
  }

  return weights;
}

function getProductMasterMills_(spreadsheet) {
  let sheet = null;
  if (!spreadsheet) return {};
  try {
    sheet = spreadsheet.getSheetByName('ProductMaster');
  } catch (e) {}
  const millMap = {};
  if (!sheet) return millMap;
  const values = sheet.getDataRange ? sheet.getDataRange().getValues() : [];
  for (let i = 1; i < values.length; i++) {
    const mill = cleanText_(values[i][1]);
    const prod = normalizeText_(values[i][2]);
    if (prod && mill) {
      millMap[prod] = mill;
    }
  }
  return millMap;
}

function getMillForProduct_(product, millMap) {
  if (!product) return '';
  const norm = normalizeText_(product);
  if (millMap && millMap[norm]) return millMap[norm];

  // Fallback inferences if not in ProductMaster:
  if (/^DOUBLE\s*A/i.test(product)) return 'Double A';
  if (/^N\s*R\s*COPIER/i.test(product) || /^NR\s*COPIER/i.test(product) || /^NR\b/i.test(product)) return 'NR';
  if (/^KHANNA/i.test(product)) return 'Khanna';
  if (/^SATIA/i.test(product)) return 'Satia';
  if (/^BILT/i.test(product) || /^BILL\s*MATRIX/i.test(product)) return 'BILT';
  if (/^REFLECTION/i.test(product) || /^ANDHRA/i.test(product)) return 'Andhra';
  if (/^SPECTRA/i.test(product)) return 'Spectra';
  if (/^PAPER\s*ONE/i.test(product)) return 'Paper One';
  if (/^ROZAANA/i.test(product) || /^ROZANA/i.test(product)) return 'Rozaana';
  if (/^CENTURY/i.test(product)) return 'Century';
  if (/^J\s*K/i.test(product) || /^JK/i.test(product)) return 'JK';

  const firstWord = product.trim().split(/\s+/)[0];
  return firstWord || '';
}

function normalizeProductDisplayName_(productName) {
  if (!productName) return '';
  const raw = cleanText_(productName);
  const norm = normalizeText_(raw);
  const aliases = {
    'biltcopierplaina475gsm': 'BILT WHITE COPIER A4 75 GSM',
    'biltcopierplaina470gsm': 'BILT WHITE COPIER A4 70 GSM',
    'biltcopierplaina480gsm': 'BILT WHITE COPIER A4 80 GSM',
    'biltcopierplaina375gsm': 'BILT WHITE COPIER A3 75 GSM',
    'biltcopierplaina370gsm': 'BILT WHITE COPIER A3 70 GSM',
    'biltplaincopiera475gsm': 'BILT WHITE COPIER A4 75 GSM',
    'biltplaincopiera470gsm': 'BILT WHITE COPIER A4 70 GSM',
    'biltwhitecopierplaina475gsm': 'BILT WHITE COPIER A4 75 GSM',
    'biltwhitecopierplaina470gsm': 'BILT WHITE COPIER A4 70 GSM'
  };
  if (aliases[norm]) {
    return aliases[norm];
  }
  if (/^bilt/i.test(raw) && /copier\s*plain|plain\s*copier/i.test(raw)) {
    return raw.replace(/copier\s*plain/i, 'WHITE COPIER').replace(/plain\s*copier/i, 'WHITE COPIER');
  }
  return raw;
}

function fixBiltCopierPlainLabels() {
  const books = getByTruckSpreadsheets_();
  let fixedInward = 0;
  let fixedMaster = 0;
  let fixedSource = 0;

  // 1. Fix PurchaseLog_By Truck in WMS
  try {
    const pSheet = books.wms.getSheetByName(BY_TRUCK_SYNC_CONFIG.purchaseTargetSheet);
    if (pSheet && pSheet.getLastRow() > 1) {
      const pHeaders = pSheet.getRange(1, 1, 1, pSheet.getLastColumn()).getValues()[0].map(normalizeText_);
      const pProdCol = pHeaders.indexOf('product');
      if (pProdCol >= 0) {
        const pRange = pSheet.getRange(2, pProdCol + 1, pSheet.getLastRow() - 1, 1);
        const pValues = pRange.getValues();
        let changed = false;
        for (let r = 0; r < pValues.length; r++) {
          const currentProd = cleanText_(pValues[r][0]);
          const normalized = normalizeProductDisplayName_(currentProd);
          if (normalized !== currentProd) {
            pValues[r][0] = normalized;
            fixedInward++;
            changed = true;
          }
        }
        if (changed) {
          pRange.setValues(pValues);
        }
      }
    }
  } catch (e) {
    Logger.log('fixBiltCopierPlainLabels purchase target notice: ' + e.message);
  }

  // 2. Fix MasterStock in WMS (delete stale 'BILT COPIER PLAIN' rows)
  try {
    const mSheet = books.wms.getSheetByName('MasterStock');
    if (mSheet && mSheet.getLastRow() > 1) {
      const mData = mSheet.getDataRange().getValues();
      for (let r = mData.length - 1; r >= 1; r--) {
        const prod = cleanText_(mData[r][1]);
        const wh = cleanText_(mData[r][2]);
        if (normalizeText_(wh) === 'bytruck' && /copier\s*plain|plain\s*copier/i.test(prod)) {
          if (typeof mSheet.deleteRow === 'function') {
            mSheet.deleteRow(r + 1);
            fixedMaster++;
          }
        }
      }
    }
  } catch (e) {
    Logger.log('fixBiltCopierPlainLabels master notice: ' + e.message);
  }

  // 3. Fix Orders (Readable) in Purchase Source Sheet if accessible
  try {
    const srcSheet = books.purchaseSource ? books.purchaseSource.getSheetByName(BY_TRUCK_SYNC_CONFIG.purchaseSourceSheet) : null;
    if (srcSheet && srcSheet.getLastRow() >= BY_TRUCK_SYNC_CONFIG.purchaseStartRow) {
      const pCol = BY_TRUCK_SYNC_CONFIG.purchaseSourceColumns.product;
      const numRows = srcSheet.getLastRow() - BY_TRUCK_SYNC_CONFIG.purchaseStartRow + 1;
      const sRange = srcSheet.getRange(BY_TRUCK_SYNC_CONFIG.purchaseStartRow, pCol + 1, numRows, 1);
      const sValues = sRange.getValues();
      let sChanged = false;
      for (let r = 0; r < sValues.length; r++) {
        const prodVal = cleanText_(sValues[r][0]);
        if (/copier\s*plain|plain\s*copier/i.test(prodVal)) {
          const norm = normalizeProductDisplayName_(prodVal);
          if (norm !== prodVal) {
            sValues[r][0] = norm;
            fixedSource++;
            sChanged = true;
          }
        }
      }
      if (sChanged) {
        sRange.setValues(sValues);
      }
    }
  } catch (e) {
    Logger.log('fixBiltCopierPlainLabels purchase source notice: ' + e.message);
  }

  return { fixedInward: fixedInward, fixedMaster: fixedMaster, fixedSource: fixedSource };
}

function getByTruckStartRows_() {
  const props = (typeof PropertiesService !== 'undefined' && PropertiesService.getScriptProperties)
    ? PropertiesService.getScriptProperties()
    : null;
  const pProp = props ? Number(props.getProperty('BY_TRUCK_PURCHASE_START_ROW')) : 0;
  const sProp = props ? Number(props.getProperty('BY_TRUCK_SALES_START_ROW')) : 0;
  return {
    purchaseStartRow: pProp > 0 ? pProp : BY_TRUCK_SYNC_CONFIG.purchaseStartRow,
    salesStartRow: sProp > 0 ? sProp : BY_TRUCK_SYNC_CONFIG.salesStartRow
  };
}

function freezeByTruckSyncToNewRowsOnly() {
  const books = getByTruckSpreadsheets_();
  const pSheet = getByTruckSheet_(books.purchaseSource || books.source, BY_TRUCK_SYNC_CONFIG.purchaseSourceSheet);
  const sSheet = getByTruckSheet_(books.salesSource || books.source, BY_TRUCK_SYNC_CONFIG.salesSourceSheet);
  
  const pNext = pSheet.getLastRow() + 1;
  const sNext = sSheet.getLastRow() + 1;
  
  const props = PropertiesService.getScriptProperties();
  props.setProperty('BY_TRUCK_PURCHASE_START_ROW', String(pNext));
  props.setProperty('BY_TRUCK_SALES_START_ROW', String(sNext));
  
  const msg = `✅ By Truck Purana Data Freeze Ho Gaya!\n\n` +
              `• Purchase Start Row set to: ${pNext}\n` +
              `• Sales Start Row set to: ${sNext}\n\n` +
              `Ab 2:00 PM se pehle ka koi bhi purana historical record inward/outward nahi hoga.\n` +
              `Kewal aage aane wale NAYE records hi sync honge!`;
  Logger.log(msg);
  try {
    SpreadsheetApp.getUi().alert('By Truck Sync Configured', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {}
  return { status: 'success', purchaseStartRow: pNext, salesStartRow: sNext, message: msg };
}

function cleanTodaySyncedByTruckRows() {
  const books = getByTruckSpreadsheets_();
  const pSheet = books.wms.getSheetByName(BY_TRUCK_SYNC_CONFIG.purchaseTargetSheet);
  const sSheet = books.wms.getSheetByName(BY_TRUCK_SYNC_CONFIG.salesTargetSheet);
  
  let pDeleted = 0, sDeleted = 0;
  
  if (pSheet && pSheet.getLastRow() > 1) {
    const pHeaders = pSheet.getRange(1, 1, 1, pSheet.getLastColumn()).getValues()[0].map(normalizeText_);
    const pKeyCol = pHeaders.indexOf(normalizeText_(BY_TRUCK_SYNC_CONFIG.sourceKeyHeader));
    if (pKeyCol >= 0) {
      const pData = pSheet.getDataRange().getValues();
      for (let r = pData.length - 1; r >= 1; r--) {
        const key = cleanText_(pData[r][pKeyCol]);
        if (key && key.startsWith('purchase:')) {
          pSheet.deleteRow(r + 1);
          pDeleted++;
        }
      }
    }
  }
  
  if (sSheet && sSheet.getLastRow() > 1) {
    const sHeaders = sSheet.getRange(1, 1, 1, sSheet.getLastColumn()).getValues()[0].map(normalizeText_);
    const sKeyCol = sHeaders.indexOf(normalizeText_(BY_TRUCK_SYNC_CONFIG.sourceKeyHeader));
    if (sKeyCol >= 0) {
      const sData = sSheet.getDataRange().getValues();
      for (let r = sData.length - 1; r >= 1; r--) {
        const key = cleanText_(sData[r][sKeyCol]);
        if (key && key.startsWith('sales:')) {
          sSheet.deleteRow(r + 1);
          sDeleted++;
        }
      }
    }
  }
  
  if (typeof rebuildByTruckOnly === 'function') {
    try { rebuildByTruckOnly(); } catch (e) {}
  }
  
  const msg = `✅ Today's Synced Rows Cleaned!\n` +
              `• PurchaseLog_By Truck: ${pDeleted} automated rows removed\n` +
              `• SalesLog_By Truck: ${sDeleted} automated rows removed\n\n` +
              `Manual rows intact hain aur By Truck MasterStock recalculate ho gaya!`;
  Logger.log(msg);
  try {
    SpreadsheetApp.getUi().alert('Clean Synced By Truck Rows', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {}
  return { status: 'success', pDeleted, sDeleted, message: msg };
}

function buildByTruckPurchases_(sourceRows, weights, millMap) {
  const startRows = getByTruckStartRows_();
  const start = startRows.purchaseStartRow - 1;
  const columns = BY_TRUCK_SYNC_CONFIG.purchaseSourceColumns;
  const rows = [];
  const totalsByPo = {};

  for (let i = start; i < sourceRows.length; i++) {
    const row = sourceRows[i];
    const po = cleanText_(row[columns.po]);
    const product = normalizeProductDisplayName_(cleanText_(row[columns.product]));
    const qtyKg = toNumber_(row[columns.qtyKg]);
    const invoiceDate = parseIndianInvoiceDate_(row[columns.date]);
    if (
      normalizeText_(row[columns.status]) !== 'yes' ||
      !invoiceDate ||
      !po ||
      !product ||
      qtyKg <= 0
    ) continue;

    const poKey = normalizeText_(po);
    if (!totalsByPo[poKey]) {
      totalsByPo[poKey] = { qtyKg: 0, insurance: {}, freight: {} };
    }
    totalsByPo[poKey].qtyKg += qtyKg;
    const insurance = toNumber_(row[columns.insurance]);
    const freight = toNumber_(row[columns.freight]);
    if (insurance) totalsByPo[poKey].insurance[String(insurance)] = insurance;
    if (freight) totalsByPo[poKey].freight[String(freight)] = freight;
    rows.push({ row, sourceRow: i + 1, po, product, qtyKg, poKey, invoiceDate });
  }

  return rows.map(function (entry) {
    const weight = getProductWeight_(weights, entry.product);
    const poTotals = totalsByPo[entry.poKey];
    const totalInsurance = sumUniqueValues_(poTotals.insurance);
    const totalFreight = sumUniqueValues_(poTotals.freight);
    const ratePerKg = toNumber_(entry.row[columns.ratePerKg]);
    const allocatedInsurance = roundMoney_(
      poTotals.qtyKg ? totalInsurance * entry.qtyKg / poTotals.qtyKg : 0
    );
    const allocatedFreight = roundMoney_(
      poTotals.qtyKg ? totalFreight * entry.qtyKg / poTotals.qtyKg : 0
    );
    const qtyReam = Math.round(entry.qtyKg / weight);
    if (qtyReam <= 0) {
      throw new Error('Kg-to-ream conversion produced zero quantity for product: ' + entry.product);
    }
    const pricePerReam = Number((ratePerKg * weight).toFixed(2));
    const amount = Number((qtyReam * pricePerReam).toFixed(2));
    const netBuyPerReam = (amount + allocatedInsurance + allocatedFreight) / qtyReam;
    const mill = getMillForProduct_(entry.product, millMap);
    return {
      sourceKey: 'purchase:' + entry.sourceRow,
      naturalKey: [entry.po, entry.product],
      values: {
        date: entry.invoiceDate,
        mill: mill,
        product: entry.product,
        warehouse: 'By Truck',
        unitType: 'Ream',
        qtyReam: qtyReam,
        pricePerReam: pricePerReam,
        amount: amount,
        billType: 'GST',
        insurance: allocatedInsurance,
        freight: allocatedFreight,
        loading: 0,
        extrasClaim: 0,
        netBuyPerReam: netBuyPerReam,
        netBuyPerKg: netBuyPerReam / weight,
        po: entry.po,
        qtyKg: entry.qtyKg,
        ratePerKg: ratePerKg
      }
    };
  });
}

function getByTruckPurchaseCosts_(purchaseSheet, weights) {
  const values = purchaseSheet.getDataRange().getValues();
  if (values.length < 2) return {};
  const columns = resolveColumns_(values[0], {
    po: ['PO No', 'PO Number', 'PO'],
    product: ['Product', 'Product Name', 'Item'],
    qtyReam: ['Qty', 'Quantity', 'Qty (Ream)', 'Qty Ream'],
    netBuyPerKg: ['Net Buy Kg', 'Net Buy/Kg', 'Net Buy per Kg', 'Net Rate per Kg'],
    netBuyPerReam: ['Net Buy Ream', 'Net Buy/Ream', 'Net Buy per Ream'],
    amount: ['Amount', 'Purchase Amount', 'Total Amount'],
    freight: ['Cartage', 'Freight'],
    insurance: ['Insurance']
  });
  if (columns.po < 0 || columns.product < 0 || columns.qtyReam < 0) {
    throw new Error('PurchaseLog_By Truck must contain PO No, Product, and Qty columns.');
  }
  const costs = {};
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const po = cleanText_(row[columns.po]);
    const rawProd = cleanText_(row[columns.product]);
    const normalizedProd = normalizeProductDisplayName_(rawProd);
    const product = normalizeText_(normalizedProd);
    if (!po || !product) continue;
    const weight = getProductWeight_(weights, normalizedProd);
    const qtyReam = toNumber_(row[columns.qtyReam]);
    const qtyKg = qtyReam * weight;
    let netPerKg = columns.netBuyPerKg >= 0 ? toNumber_(row[columns.netBuyPerKg]) : 0;
    if (!netPerKg && columns.netBuyPerReam >= 0) {
      netPerKg = toNumber_(row[columns.netBuyPerReam]) / weight;
    }
    if (!netPerKg && columns.amount >= 0 && qtyReam > 0) {
      const extras = (columns.freight >= 0 ? toNumber_(row[columns.freight]) : 0) +
        (columns.insurance >= 0 ? toNumber_(row[columns.insurance]) : 0);
      netPerKg = (toNumber_(row[columns.amount]) + extras) / qtyReam / weight;
    }
    const costKey = normalizeKey_([po, product]);
    if (!costs[costKey]) costs[costKey] = { total: 0, qtyKg: 0, fallback: netPerKg };
    if (qtyKg > 0) {
      costs[costKey].total += netPerKg * qtyKg;
      costs[costKey].qtyKg += qtyKg;
    }
  }
  Object.keys(costs).forEach(function (costKey) {
    costs[costKey] = costs[costKey].qtyKg
      ? costs[costKey].total / costs[costKey].qtyKg
      : costs[costKey].fallback;
  });
  return costs;
}

function buildByTruckSales_(sourceRows, weights, costs, millMap) {
  const startRows = getByTruckStartRows_();
  const start = startRows.salesStartRow - 1;
  const columns = BY_TRUCK_SYNC_CONFIG.salesSourceColumns;
  const sales = [];
  for (let i = start; i < sourceRows.length; i++) {
    const row = sourceRows[i];
    const po = cleanText_(row[columns.po]);
    const billNo = cleanText_(row[columns.billNo]);
    const product = normalizeProductDisplayName_(cleanText_(row[columns.product]));
    const qtyReam = toNumber_(row[columns.qtyReam]);
    const salesDate = parseIndianInvoiceDate_(row[columns.date]);
    if (!po || !billNo || !product || qtyReam <= 0 || !salesDate) continue;

    const weight = getProductWeight_(weights, product);
    const ratePerKg = toNumber_(row[columns.ratePerKg]);
    const ratePerReam = ratePerKg * weight;
    const amount = roundMoney_(qtyReam * ratePerReam);
    let costPerKg = costs[normalizeKey_([po, product])];
    if (costPerKg === undefined || isNaN(costPerKg)) {
      // Smart fallback 1: Check if PO has a purchase with matching normalized alias or GSM
      const poPrefix = normalizeText_(po) + '|';
      const poMatches = Object.keys(costs).filter(function (k) { return k.indexOf(poPrefix) === 0; });
      if (poMatches.length === 1) {
        costPerKg = costs[poMatches[0]];
      } else if (poMatches.length > 1) {
        const gsmMatch = normalizeText_(product).match(/\d{2,3}gsm/);
        if (gsmMatch) {
          const found = poMatches.find(function (k) { return k.indexOf(gsmMatch[0]) !== -1; });
          if (found) costPerKg = costs[found];
        }
      }
    }
    if (costPerKg === undefined || isNaN(costPerKg)) {
      Logger.log('⚠️ [ByTruckSync Warning] No purchase cost found for PO: ' + po + ', Product: ' + product);
      costPerKg = ratePerKg > 0 ? ratePerKg : 0;
    }
    const costPerReam = costPerKg * weight;
    const costAmount = roundMoney_(qtyReam * costPerReam);
    const mill = getMillForProduct_(product, millMap);

    sales.push({
      sourceKey: 'sales:' + (i + 1),
      naturalKey: [po, billNo, product],
      values: {
        date: salesDate,
        salesPerson: cleanText_(row[columns.salesPerson]),
        padNo: '',
        billNo: billNo,
        party: cleanText_(row[columns.party]),
        mill: mill,
        product: product,
        warehouse: 'By Truck',
        billType: 'GST',
        qtyReam: qtyReam,
        ratePerKg: ratePerKg,
        ratePerReam: ratePerReam,
        discount: 0,
        amount: amount,
        freight: 0,
        recovered: 'No',
        costPerReam: costPerReam,
        costAmount: costAmount,
        profit: roundMoney_(amount - costAmount),
        po: po
      }
    });
  }
  return sales;
}

function upsertByTruckRows_(sheet, records, kind) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) throw new Error('Target sheet has no header row: ' + sheet.getName());
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  let keyColumn = findHeaderColumn_(headers, [BY_TRUCK_SYNC_CONFIG.sourceKeyHeader]);
  if (keyColumn < 0) {
    keyColumn = headers.length;
    sheet.getRange(1, keyColumn + 1).setValue(BY_TRUCK_SYNC_CONFIG.sourceKeyHeader);
    sheet.hideColumns(keyColumn + 1);
    headers.push(BY_TRUCK_SYNC_CONFIG.sourceKeyHeader);
  }

  const fieldAliases = kind === 'purchase' ? {
    date: ['Date', 'Invoice Date', 'Bill Date'],
    mill: ['Mill', 'Mills', 'Category', 'Brand'],
    product: ['Product', 'Product Name', 'Item'],
    warehouse: ['Warehouse', 'Godown'],
    unitType: ['Unit Type', 'Unit', 'UOM'],
    qtyReam: ['Qty', 'Quantity', 'Qty (Ream)', 'Qty Ream'],
    pricePerReam: ['Price per Ream', 'Rate per Ream', 'Price', 'Rate'],
    amount: ['Amount', 'Purchase Amount', 'Total Amount'],
    billType: ['Bill Type', 'BillType', 'Type'],
    insurance: ['Insurance'],
    freight: ['Cartage', 'Freight'],
    loading: ['Loading'],
    extrasClaim: ['Extras Claimed', 'Extras Claim', 'Extras'],
    netBuyPerReam: ['Net Buy Ream', 'Net Buy/Ream', 'Net Buy per Ream'],
    netBuyPerKg: ['Net Buy Kg', 'Net Buy/Kg', 'Net Buy per Kg', 'Net Rate per Kg'],
    po: ['PO No', 'PO Number', 'PO']
  } : {
    date: ['Date', 'Invoice Date', 'Bill Date'],
    salesPerson: ['Sales Person', 'Salesperson'],
    padNo: ['Pad No.', 'Pad No', 'Pad'],
    billNo: ['Bill No', 'Bill Number', 'Invoice No', 'Invoice Number'],
    party: ['Party', 'Party Name', 'Customer'],
    mill: ['Mills', 'Mill', 'Category', 'Brand'],
    product: ['Product', 'Product Name', 'Item'],
    warehouse: ['Warehouse', 'Godown'],
    billType: ['Bill Type', 'BillType', 'Type'],
    qtyReam: ['Qty', 'Quantity', 'Qty (Ream)', 'Qty Ream'],
    ratePerReam: ['Rate', 'Rate per Ream', 'Price per Ream', 'Price'],
    ratePerKg: ['Rate per Kg', 'Rate/Kg', 'Price per Kg', 'Price/Kg'],
    discount: ['Discount', 'Disc.'],
    amount: ['Amount', 'Sales Amount', 'Total Amount'],
    freight: ['Freight', 'Cartage'],
    recovered: ['Recovered?', 'Recovered', 'Freight Recovered'],
    costPerReam: ['Cost Price', 'Cost per Ream', 'Purchase Price'],
    costAmount: ['Cost Amount', 'Total Cost'],
    profit: ['Profit', 'Gross Profit', 'GP'],
    po: ['PO No', 'PO Number', 'PO']
  };
  const columns = {};
  Object.keys(fieldAliases).forEach(function (field) {
    columns[field] = findHeaderColumn_(headers, fieldAliases[field]);
  });
  const poFieldColumn = columns.po >= 0 ? columns.po : (kind === 'purchase' ? 15 : 17);
  columns.po = poFieldColumn;

  const existingRows = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues()
    : [];
  const exactRows = {};
  const naturalRows = {};
  existingRows.forEach(function (row, index) {
    const sourceKey = cleanText_(row[keyColumn]);
    if (sourceKey) exactRows[sourceKey] = index + 2;
    else {
      const naturalKey = getNaturalTargetKey_(row, columns, kind);
      if (naturalKey) {
        if (!naturalRows[naturalKey]) naturalRows[naturalKey] = [];
        naturalRows[naturalKey].push(index + 2);
      }
    }
  });

  const claimedRows = {};
  let nextRow = Math.max(sheet.getLastRow() + 1, 2);
  const updates = [];
  records.forEach(function (record) {
    const cleanPo = cleanText_(record.values.po).replace(/^po[\s\-_]*/i, '');
    const naturalKey = normalizeKey_(
      kind === 'sales'
        ? [cleanPo, cleanText_(record.values.billNo), record.values.product]
        : [cleanPo, record.values.product]
    );
    let targetRow = exactRows[record.sourceKey];
    if (!targetRow && naturalRows[naturalKey]) {
      targetRow = naturalRows[naturalKey].find(function (candidate) {
        return !claimedRows[candidate];
      });
    }
    if (!targetRow && kind === 'sales' && record.values.billNo) {
      const billKey = normalizeKey_(['bill', cleanText_(record.values.billNo), record.values.product]);
      if (naturalRows[billKey]) {
        targetRow = naturalRows[billKey].find(function (candidate) {
          return !claimedRows[candidate];
        });
      }
    }
    if (!targetRow) targetRow = nextRow++;
    claimedRows[targetRow] = true;
    updates.push({ row: targetRow, sourceKey: record.sourceKey, values: record.values });
  });

  // Preserve manual sheet entries (e.g. Freight, Insurance) if source had 0/blank
  updates.forEach(function (update) {
    const rowIndex = update.row - 2;
    if (rowIndex >= 0 && rowIndex < existingRows.length) {
      const existingRow = existingRows[rowIndex];

      if (kind === 'purchase') {
        let changed = false;
        // 1. Preserve Freight / Cartage if manually entered in sheet and incoming from source is 0
        const existingFreight = columns.freight >= 0 ? toNumber_(existingRow[columns.freight]) : 0;
        const incomingFreight = toNumber_(update.values.freight);
        if (existingFreight > 0 && incomingFreight === 0) {
          update.values.freight = existingFreight;
          changed = true;
        }

        // 2. Preserve Insurance if manually entered in sheet and incoming from source is 0
        const existingInsurance = columns.insurance >= 0 ? toNumber_(existingRow[columns.insurance]) : 0;
        const incomingInsurance = toNumber_(update.values.insurance);
        if (existingInsurance > 0 && incomingInsurance === 0) {
          update.values.insurance = existingInsurance;
          changed = true;
        }

        // 3. Preserve Loading if manually entered
        const existingLoading = columns.loading >= 0 ? toNumber_(existingRow[columns.loading]) : 0;
        const incomingLoading = toNumber_(update.values.loading);
        if (existingLoading > 0 && incomingLoading === 0) {
          update.values.loading = existingLoading;
        }

        // 4. Preserve Extras Claimed if manually entered
        if (columns.extrasClaim >= 0) {
          const existingClaim = cleanText_(existingRow[columns.extrasClaim]);
          if (existingClaim && (!update.values.extrasClaim || update.values.extrasClaim === 'No' || update.values.extrasClaim === 0)) {
            update.values.extrasClaim = existingClaim;
          }
        }

        // Only recalculate Net Buy if freight or insurance was preserved from sheet
        if (changed && update.values.qtyReam > 0) {
          const effFreight = toNumber_(update.values.freight);
          const effInsurance = toNumber_(update.values.insurance);
          const weight = (update.values.ratePerKg && update.values.ratePerKg > 0)
            ? (update.values.pricePerReam / update.values.ratePerKg)
            : 2.34;
          update.values.netBuyPerReam = (update.values.amount + effFreight + effInsurance) / update.values.qtyReam;
          update.values.netBuyPerKg = update.values.netBuyPerReam / weight;
        }
      } else if (kind === 'sales') {
        let extraFreightDeduction = 0;
        // 1. Preserve Freight in Sales if manually entered in sheet
        const existingFreight = columns.freight >= 0 ? toNumber_(existingRow[columns.freight]) : 0;
        const incomingFreight = toNumber_(update.values.freight);
        if (existingFreight > 0 && incomingFreight === 0) {
          update.values.freight = existingFreight;
          const isRec = columns.recovered >= 0
            ? String(existingRow[columns.recovered] || '').toLowerCase().trim() === 'yes'
            : false;
          if (!isRec) extraFreightDeduction = existingFreight;
        }

        // 2. Preserve Recovered? in Sales if manually set
        if (columns.recovered >= 0) {
          const existingRec = cleanText_(existingRow[columns.recovered]);
          if (existingRec && (!update.values.recovered || update.values.recovered === 'No')) {
            update.values.recovered = existingRec;
          }
        }

        // 3. Preserve Discount in Sales if manually entered
        const existingDisc = columns.discount >= 0 ? toNumber_(existingRow[columns.discount]) : 0;
        const incomingDisc = toNumber_(update.values.discount);
        if (existingDisc > 0 && incomingDisc === 0) {
          update.values.discount = existingDisc;
        }

        // 4. Preserve Pad No if manually entered
        if (columns.padNo >= 0) {
          const existingPad = cleanText_(existingRow[columns.padNo]);
          if (existingPad && !update.values.padNo) {
            update.values.padNo = existingPad;
          }
        }

        if (extraFreightDeduction > 0) {
          update.values.profit = roundMoney_(update.values.profit - extraFreightDeduction);
        }
      }
    }
  });

  const lastRow = Math.max(sheet.getLastRow(), nextRow - 1);
  Object.keys(columns).forEach(function (field) {
    const column = columns[field] + 1;
    if (column < 1) return;
    const values = lastRow > 1
      ? sheet.getRange(2, column, lastRow - 1, 1).getValues()
      : [];
    updates.forEach(function (update) {
      if (update.values[field] !== undefined) values[update.row - 2] = [update.values[field]];
    });
    if (values.length) sheet.getRange(2, column, values.length, 1).setValues(values);
  });

  const keyValues = lastRow > 1
    ? sheet.getRange(2, keyColumn + 1, lastRow - 1, 1).getValues()
    : [];
  updates.forEach(function (update) {
    keyValues[update.row - 2] = [update.sourceKey];
  });
  if (keyValues.length) sheet.getRange(2, keyColumn + 1, keyValues.length, 1).setValues(keyValues);
}

function getNaturalTargetKey_(row, columns, kind) {
  const poRaw = cleanText_(row[columns.po]);
  const po = poRaw.replace(/^po[\s\-_]*/i, '');
  const rawProduct = cleanText_(row[columns.product]);
  const product = normalizeProductDisplayName_(rawProduct);
  if (!product) return '';

  if (kind === 'sales') {
    const billNo = cleanText_(row[columns.billNo]);
    if (po && billNo) {
      return normalizeKey_([po, billNo, product]);
    }
    if (billNo) {
      return normalizeKey_(['bill', billNo, product]);
    }
    return '';
  }

  if (!po) return '';
  return normalizeKey_([po, product]);
}

function findHeaderColumn_(headers, aliases) {
  const normalizedHeaders = headers.map(normalizeText_);
  for (let i = 0; i < aliases.length; i++) {
    const index = normalizedHeaders.indexOf(normalizeText_(aliases[i]));
    if (index >= 0) return index;
  }
  return -1;
}

function resolveColumns_(headers, aliases) {
  const columns = {};
  Object.keys(aliases).forEach(function (field) {
    columns[field] = findHeaderColumn_(headers, aliases[field]);
  });
  return columns;
}

function estimateWeightFromName_(productName) {
  const norm = String(productName || '').toUpperCase();
  let gsm = 70;
  const gsmMatch = norm.match(/(\d{2,3})\s*GSM/);
  if (gsmMatch && gsmMatch[1]) {
    gsm = Number(gsmMatch[1]);
  }

  if (norm.includes('A3')) {
    return roundMoney_((62.37 * gsm) / 1000);
  } else if (norm.includes('FS') || norm.includes('LEGAL') || norm.includes('FOOLSCAP')) {
    return roundMoney_((37.0875 * gsm) / 1000);
  } else {
    return roundMoney_((31.185 * gsm) / 1000);
  }
}

function getProductWeight_(weights, product) {
  if (!product) return 2.34;
  const key = normalizeText_(product);
  if (weights && weights[key] && weights[key] > 0) return weights[key];

  if (weights) {
    const keys = Object.keys(weights);
    for (let i = 0; i < keys.length; i++) {
      const k = keys[i];
      if (k && (key.includes(k) || k.includes(key))) {
        return weights[k];
      }
    }
  }

  const estimated = estimateWeightFromName_(product);
  if (estimated > 0) {
    if (weights) weights[key] = estimated;
    return estimated;
  }

  return 2.34;
}

function normalizeText_(value) {
  return String(value === null || value === undefined ? '' : value)
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

function normalizeKey_(parts) {
  return parts.map(function (part) { return normalizeText_(part); }).join('|');
}

function cleanText_(value) {
  return String(value === null || value === undefined ? '' : value).trim();
}

function toNumber_(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  const parsed = Number(String(value === null || value === undefined ? '' : value)
    .replace(/,/g, '').trim());
  return Number.isFinite(parsed) ? parsed : 0;
}

function parseIndianInvoiceDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Number.isNaN(value.getTime()) ? null : value;
  }
  const match = cleanText_(value).match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!match) return null;
  const day = Number(match[1]);
  const month = Number(match[2]);
  const year = Number(match[3]);
  if (day < 1 || month < 1 || month > 12 || year < 1) return null;
  const date = new Date(year, month - 1, day);
  if (
    date.getFullYear() !== year ||
    date.getMonth() !== month - 1 ||
    date.getDate() !== day
  ) return null;
  return date;
}

function roundMoney_(value) {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}

function sumUniqueValues_(values) {
  return Object.keys(values).reduce(function (total, key) {
    return total + values[key];
  }, 0);
}

function dedupeByTruckSheet_(sheet, kind) {
  if (!sheet || sheet.getLastRow() <= 1) return 0;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const normalizedHeaders = headers.map(normalizeText_);
  const keyColumn = findHeaderColumn_(headers, [BY_TRUCK_SYNC_CONFIG.sourceKeyHeader]);
  const prCol = normalizedHeaders.indexOf('product');
  const poCol = findHeaderColumn_(normalizedHeaders, ['pono', 'ponumber', 'po']);
  const billCol = kind === 'sales' ? findHeaderColumn_(normalizedHeaders, ['billno', 'billnumber', 'invoiceno']) : -1;
  const qtyCol = normalizedHeaders.indexOf('qty');
  if (qtyCol < 0 || prCol < 0) return 0;

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  const seen = {};
  const toDelete = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowNum = i + 2;
    const prod = normalizeProductDisplayName_(row[prCol]);
    const qty = toNumber_(row[qtyCol]);
    const po = poCol >= 0 ? cleanText_(row[poCol]).replace(/^po[\s\-_]*/i, '') : '';
    const bill = billCol >= 0 ? cleanText_(row[billCol]) : '';
    const sig = kind === 'sales'
      ? (bill ? `${bill}||${prod}||${qty}` : `${po}||${prod}||${qty}`)
      : `${po}||${prod}||${qty}`;
    const hasKey = keyColumn >= 0 && Boolean(cleanText_(row[keyColumn]));

    if (seen[sig]) {
      if (!seen[sig].hasKey && hasKey) {
        toDelete.push(seen[sig].row);
        seen[sig] = { row: rowNum, hasKey: true };
      } else {
        toDelete.push(rowNum);
      }
    } else {
      seen[sig] = { row: rowNum, hasKey: hasKey };
    }
  }

  toDelete.sort(function(a, b) { return b - a; }).forEach(function(r) {
    if (typeof sheet.deleteRow === 'function') {
      sheet.deleteRow(r);
    }
  });

  return toDelete.length;
}
