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
  syncByTruckPurchases_();
  syncByTruckSales_();
  try {
    if (typeof rebuildByTruckOnly === 'function') {
      rebuildByTruckOnly();
    }
  } catch (e) {
    Logger.log('rebuildByTruckOnly after sync error: ' + e.message);
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
  const sourceRows = source.getDataRange().getValues();
  const purchases = buildByTruckPurchases_(sourceRows, weights);
  upsertByTruckRows_(target, purchases, 'purchase');
}

function syncByTruckSales_() {
  const books = getByTruckSpreadsheets_();
  const source = getByTruckSheet_(books.salesSource || books.source, BY_TRUCK_SYNC_CONFIG.salesSourceSheet);
  const purchaseSheet = getByTruckSheet_(books.wms, BY_TRUCK_SYNC_CONFIG.purchaseTargetSheet);
  const target = getByTruckSheet_(books.wms, BY_TRUCK_SYNC_CONFIG.salesTargetSheet);
  const weights = getProductWeights_(books.wms);
  const costs = getByTruckPurchaseCosts_(purchaseSheet, weights);
  const sourceRows = source.getDataRange().getValues();
  const sales = buildByTruckSales_(sourceRows, weights, costs);
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
  const sheet = getByTruckSheet_(spreadsheet, BY_TRUCK_SYNC_CONFIG.productWeightsSheet);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) throw new Error('ProductWeights has no product rows.');

  const productColumn = findHeaderColumn_(values[0], ['Product', 'Product Name', 'Item']);
  const weightColumn = findHeaderColumn_(values[0], [
    'Kg per Ream', 'KG per Ream', 'Weight per Ream', 'Kg/Ream'
  ]);
  if (productColumn < 0 || weightColumn < 0) {
    throw new Error('ProductWeights must contain Product and Kg per Ream columns.');
  }

  const weights = {};
  for (let i = 1; i < values.length; i++) {
    const product = normalizeText_(values[i][productColumn]);
    const weight = toNumber_(values[i][weightColumn]);
    if (product && weight > 0) weights[product] = weight;
  }
  return weights;
}

function buildByTruckPurchases_(sourceRows, weights) {
  const start = BY_TRUCK_SYNC_CONFIG.purchaseStartRow - 1;
  const columns = BY_TRUCK_SYNC_CONFIG.purchaseSourceColumns;
  const rows = [];
  const totalsByPo = {};

  for (let i = start; i < sourceRows.length; i++) {
    const row = sourceRows[i];
    const po = cleanText_(row[columns.po]);
    const product = cleanText_(row[columns.product]);
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
    return {
      sourceKey: 'purchase:' + entry.sourceRow,
      naturalKey: [entry.po, entry.product],
      values: {
        date: entry.invoiceDate,
        po: entry.po,
        product: entry.product,
        qtyReam: qtyReam,
        pricePerReam: pricePerReam,
        amount: amount,
        qtyKg: entry.qtyKg,
        ratePerKg: ratePerKg,
        netBuyPerReam: netBuyPerReam,
        netBuyPerKg: netBuyPerReam / weight,
        insurance: allocatedInsurance,
        freight: allocatedFreight
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
    const product = normalizeText_(row[columns.product]);
    if (!po || !product) continue;
    const weight = getProductWeight_(weights, row[columns.product]);
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

function buildByTruckSales_(sourceRows, weights, costs) {
  const start = BY_TRUCK_SYNC_CONFIG.salesStartRow - 1;
  const columns = BY_TRUCK_SYNC_CONFIG.salesSourceColumns;
  const sales = [];
  for (let i = start; i < sourceRows.length; i++) {
    const row = sourceRows[i];
    const po = cleanText_(row[columns.po]);
    const billNo = cleanText_(row[columns.billNo]);
    const product = cleanText_(row[columns.product]);
    const qtyReam = toNumber_(row[columns.qtyReam]);
    const salesDate = parseIndianInvoiceDate_(row[columns.date]);
    if (!po || !billNo || !product || qtyReam <= 0 || !salesDate) continue;

    const weight = getProductWeight_(weights, product);
    const ratePerKg = toNumber_(row[columns.ratePerKg]);
    const ratePerReam = ratePerKg * weight;
    const amount = roundMoney_(qtyReam * ratePerReam);
    const costPerKg = costs[normalizeKey_([po, product])];
    if (costPerKg === undefined) {
      throw new Error('No By Truck purchase cost found for PO ' + po + ', product: ' + product);
    }
    const costPerReam = costPerKg * weight;
    const costAmount = roundMoney_(qtyReam * costPerReam);

    sales.push({
      sourceKey: 'sales:' + (i + 1),
      naturalKey: [po, billNo, product],
      values: {
        date: salesDate,
        po: po,
        party: row[columns.party],
        salesPerson: row[columns.salesPerson],
        billNo: billNo,
        product: product,
        qtyReam: qtyReam,
        ratePerKg: ratePerKg,
        ratePerReam: ratePerReam,
        amount: amount,
        costPerReam: costPerReam,
        costAmount: costAmount,
        profit: roundMoney_(amount - costAmount)
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
    po: ['PO No', 'PO Number', 'PO'],
    product: ['Product', 'Product Name', 'Item'],
    qtyReam: ['Qty', 'Quantity', 'Qty (Ream)', 'Qty Ream'],
    pricePerReam: ['Price per Ream', 'Rate per Ream', 'Price', 'Rate'],
    amount: ['Amount', 'Purchase Amount', 'Total Amount'],
    netBuyPerReam: ['Net Buy Ream', 'Net Buy/Ream', 'Net Buy per Ream'],
    netBuyPerKg: ['Net Buy Kg', 'Net Buy/Kg', 'Net Buy per Kg', 'Net Rate per Kg'],
    insurance: ['Insurance'],
    freight: ['Cartage', 'Freight']
  } : {
    date: ['Date', 'Invoice Date', 'Bill Date'],
    po: ['PO No', 'PO Number', 'PO'],
    party: ['Party', 'Party Name', 'Customer'],
    salesPerson: ['Sales Person', 'Salesperson'],
    billNo: ['Bill No', 'Bill Number', 'Invoice No', 'Invoice Number'],
    product: ['Product', 'Product Name', 'Item'],
    qtyReam: ['Qty', 'Quantity', 'Qty (Ream)', 'Qty Ream'],
    ratePerKg: ['Rate per Kg', 'Rate/Kg', 'Price per Kg', 'Price/Kg'],
    ratePerReam: ['Rate per Ream', 'Price per Ream', 'Rate', 'Price'],
    amount: ['Amount', 'Sales Amount', 'Total Amount'],
    costPerReam: ['Cost Price', 'Cost per Ream', 'Purchase Price'],
    costAmount: ['Cost Amount', 'Total Cost'],
    profit: ['Profit', 'Gross Profit', 'GP']
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
    const naturalKey = normalizeKey_(record.naturalKey);
    let targetRow = exactRows[record.sourceKey];
    if (!targetRow && naturalRows[naturalKey]) {
      targetRow = naturalRows[naturalKey].find(function (candidate) {
        return !claimedRows[candidate];
      });
    }
    if (!targetRow) targetRow = nextRow++;
    claimedRows[targetRow] = true;
    updates.push({ row: targetRow, sourceKey: record.sourceKey, values: record.values });
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
  const po = cleanText_(row[columns.po]);
  const product = cleanText_(row[columns.product]);
  if (!po || !product) return '';
  const parts = kind === 'sales'
    ? [po, cleanText_(row[columns.billNo]), product]
    : [po, product];
  return normalizeKey_(parts);
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

function getProductWeight_(weights, product) {
  const weight = weights[normalizeText_(product)];
  if (!weight) throw new Error('No Kg per Ream value found for product: ' + product);
  return weight;
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
