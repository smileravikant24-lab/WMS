const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

class FakeRange {
  constructor(sheet, row, column, rowCount, columnCount) {
    this.sheet = sheet;
    this.row = row;
    this.column = column;
    this.rowCount = rowCount;
    this.columnCount = columnCount;
  }

  getValues() {
    const result = [];
    for (let r = 0; r < this.rowCount; r++) {
      const values = [];
      for (let c = 0; c < this.columnCount; c++) {
        const sourceRow = this.sheet.data[this.row - 1 + r] || [];
        values.push(sourceRow[this.column - 1 + c] || '');
      }
      result.push(values);
    }
    return result;
  }

  setValues(values) {
    values.forEach((row, r) => row.forEach((value, c) => {
      this.sheet.ensureCell(this.row - 1 + r, this.column - 1 + c);
      this.sheet.data[this.row - 1 + r][this.column - 1 + c] = value;
    }));
  }

  setValue(value) {
    this.setValues([[value]]);
  }
}

class FakeSheet {
  constructor(name, data) {
    this.name = name;
    this.data = data;
  }

  ensureCell(row, column) {
    while (this.data.length <= row) this.data.push([]);
    while (this.data[row].length <= column) this.data[row].push('');
  }

  getName() {
    return this.name;
  }

  getLastRow() {
    for (let i = this.data.length - 1; i >= 0; i--) {
      if (this.data[i].some(value => value !== '' && value !== null && value !== undefined)) return i + 1;
    }
    return 0;
  }

  getLastColumn() {
    return this.data.reduce((max, row) => Math.max(max, row.length), 0);
  }

  getDataRange() {
    return this.getRange(1, 1, this.getLastRow(), this.getLastColumn());
  }

  getRange(row, column, rowCount = 1, columnCount = 1) {
    return new FakeRange(this, row, column, rowCount, columnCount);
  }

  deleteRow(rowIndex) {
    this.data.splice(rowIndex - 1, 1);
  }

  hideColumns() {}
}

class FakeSpreadsheet {
  constructor(sheets) {
    this.sheets = sheets;
  }

  getSheetByName(name) {
    return this.sheets[name] || null;
  }
}

const purchaseHeaders = [
  'Date', 'Mill', 'Product', 'Warehouse', 'Unit Type', 'Qty', 'Price', 'Amount',
  'Bill Type', 'Insurance', 'Cartage', 'Loading', 'Extras Claimed',
  'Net Buy Ream', 'Net Buy Kg', 'PO No'
];

const salesHeaders = [
  'Date', 'Sales Person', 'Pad No.', 'Bill No', 'Party Name', 'Mills', 'Product',
  'Warehouse', 'Bill Type', 'Qty', 'Rate', 'Discount', 'Amount', 'Freight',
  'Recovered?', 'Cost Price', 'Profit', 'PO No'
];

function paddedRow(size, entries) {
  const row = Array(size).fill('');
  Object.keys(entries).forEach(key => { row[Number(key)] = entries[key]; });
  return row;
}

function createHarness() {
  const purchaseSource = Array.from({ length: 272 }, () => []);
  purchaseSource.push(['PO', '', '', 'Product', '', '', 'Rate/Kg', 'Qty Kg']);
  const salesSource = Array.from({ length: 170 }, () => []);
  salesSource.push([]);
  const source = new FakeSpreadsheet({
    'Orders (Readable)': new FakeSheet('Orders (Readable)', purchaseSource),
    Workflow: new FakeSheet('Workflow', salesSource)
  });
  const wms = new FakeSpreadsheet({
    'PurchaseLog_By Truck': new FakeSheet('PurchaseLog_By Truck', [purchaseHeaders]),
    'SalesLog_By Truck': new FakeSheet('SalesLog_By Truck', [salesHeaders]),
    ProductWeights: new FakeSheet('ProductWeights', [
      ['Product', 'Kg per Ream'],
      ['A4', 22]
    ])
  });
  const triggers = [];
  const createdTriggers = [];
  const context = {
    PropertiesService: {
      getScriptProperties: () => ({ getProperty: key => key.includes('SOURCE') ? 'source' : 'wms' })
    },
    SpreadsheetApp: {
      openById: id => (id === 'source' || String(id).includes('HuZABEjx') || String(id).includes('STKk27')) ? source : wms,
      getActiveSpreadsheet: () => wms,
      getUi: () => ({ ButtonSet: { OK: 'OK' }, alert() {} })
    },
    ScriptApp: {
      EventType: { CLOCK: 'CLOCK', ON_EDIT: 'ON_EDIT' },
      getProjectTriggers: () => triggers,
      deleteTrigger: trigger => triggers.splice(triggers.indexOf(trigger), 1),
      newTrigger: handler => ({
        timeBased() { return this; },
        everyMinutes(minutes) {
          this.minutes = minutes;
          return this;
        },
        create() {
          const trigger = {
            handler,
            minutes: this.minutes,
            getHandlerFunction() { return handler; },
            getEventType() { return 'CLOCK'; }
          };
          triggers.push(trigger);
          createdTriggers.push(trigger);
          return trigger;
        }
      })
    },
    Number,
    Math,
    String,
    Error,
    Object,
    Logger: { log() {} }
  };
  const script = fs.readFileSync(path.join(__dirname, '..', 'apps-script', 'ByTruckSync.gs'), 'utf8');
  vm.runInNewContext(script, context);
  return { context, source, wms, triggers, createdTriggers };
}

function sourcePurchase(
  po,
  ratePerKg,
  qtyKg,
  freight,
  insurance,
  status = 'yes',
  date = '20/11/2026'
) {
  return paddedRow(30, {
    0: po, 3: 'A4', 6: ratePerKg, 7: qtyKg, 22: date,
    25: status, 28: insurance, 29: freight
  });
}

function sourceSale(po, bill, qtyReam, ratePerKg, date = '21/11/2026') {
  return paddedRow(61, {
    1: po, 18: 'Customer', 23: 'Seller', 34: bill, 39: date,
    56: 'A4', 57: qtyReam, 60: ratePerKg
  });
}

function appendSource(sheet, row) {
  sheet.data.push(row);
}

test('refreshes blank freight and rate, then applies changed purchase quantity', () => {
  const { context, source, wms } = createHarness();
  const purchaseSource = source.getSheetByName('Orders (Readable)');
  const purchaseTarget = wms.getSheetByName('PurchaseLog_By Truck');
  purchaseTarget.data.push(paddedRow(16, { 2: 'A4', 15: 'PO-1' }));
  appendSource(purchaseSource, sourcePurchase('PO-1', '', 100, '', ''));
  context.syncByTruckData();

  purchaseSource.data[273][6] = 2.337;
  purchaseSource.data[273][28] = 25;
  purchaseSource.data[273][29] = 50;
  context.syncByTruckData();

  const purchaseSheet = purchaseTarget;
  const headers = purchaseSheet.data[0];
  const row = purchaseSheet.data[1];
  const value = name => row[headers.indexOf(name)];
  assert.equal(purchaseSheet.getLastRow(), 2);
  assert.equal(value('Cartage'), 50);
  assert.equal(value('Insurance'), 25);
  assert.equal(value('Qty'), 5);
  assert.equal(value('Price'), 51.41);
  assert.equal(value('Amount'), 257.05);
  assert.equal(value('Net Buy Ream'), 66.41);
  assert.equal(value('Net Buy Kg'), 66.41 / 22);

  purchaseSource.data[273][7] = 200;
  context.syncByTruckData();
  assert.equal(purchaseSheet.getLastRow(), 2);
  assert.equal(row[headers.indexOf('Qty')], 9);
  assert.equal(row[headers.indexOf('Amount')], 462.69);
  assert.equal(row[headers.indexOf('Net Buy Ream')], (462.69 + 75) / 9);
  assert.equal(row[headers.indexOf('Net Buy Kg')], ((462.69 + 75) / 9) / 22);
});

test('skips purchases without yes status or a strict valid Indian invoice date', () => {
  const { context, source, wms } = createHarness();
  const orders = source.getSheetByName('Orders (Readable)');
  appendSource(orders, sourcePurchase('PO-NO', 2, 100, 0, 0, 'no'));
  appendSource(orders, sourcePurchase('PO-CONFIRMED', 2, 100, 0, 0, 'Confirmed'));
  appendSource(orders, sourcePurchase('PO-INVALID-DAY', 2, 100, 0, 0, 'yes', '31/02/2026'));
  appendSource(orders, sourcePurchase('PO-ISO-DATE', 2, 100, 0, 0, 'yes', '2026-11-20'));
  context.syncByTruckData();

  assert.equal(wms.getSheetByName('PurchaseLog_By Truck').getLastRow(), 1);
});

test('repeated sync is idempotent and repeated products remain separate', () => {
  const { context, source, wms } = createHarness();
  const orders = source.getSheetByName('Orders (Readable)');
  appendSource(orders, sourcePurchase('PO-2', 1, 40, 10, 0));
  appendSource(orders, sourcePurchase('PO-2', 1, 60, 10, 0));
  context.syncByTruckData();
  context.syncByTruckData();

  const purchaseSheet = wms.getSheetByName('PurchaseLog_By Truck');
  assert.equal(purchaseSheet.getLastRow(), 3);
  assert.notEqual(purchaseSheet.data[1].at(-1), purchaseSheet.data[2].at(-1));
  assert.equal(purchaseSheet.data[1][purchaseSheet.data[0].indexOf('Qty')], 2);
  assert.equal(purchaseSheet.data[2][purchaseSheet.data[0].indexOf('Qty')], 3);
  assert.equal(purchaseSheet.data[1][purchaseSheet.data[0].indexOf('Cartage')], 4);
  assert.equal(purchaseSheet.data[2][purchaseSheet.data[0].indexOf('Cartage')], 6);
});

test('sales rate and quantity changes recalculate amount, cost, and profit', () => {
  const { context, source, wms } = createHarness();
  appendSource(source.getSheetByName('Orders (Readable)'), sourcePurchase('PO-3', 2, 100, 50, 25));
  appendSource(source.getSheetByName('Workflow'), sourceSale('PO-3', 'B-1', 2, 3));
  context.syncByTruckData();

  const salesSheet = wms.getSheetByName('SalesLog_By Truck');
  const headers = salesSheet.data[0];
  const row = salesSheet.data[1];
  const value = name => row[headers.indexOf(name)];
  assert.equal(value('Rate'), 66);
  assert.equal(value('Amount'), 132);
  assert.equal(value('Cost Price'), 59);
  assert.equal(value('Profit'), 14);

  source.getSheetByName('Workflow').data[171][57] = 3;
  source.getSheetByName('Workflow').data[171][60] = 4;
  context.syncByTruckData();
  assert.equal(salesSheet.getLastRow(), 2);
  assert.equal(value('Rate'), 88);
  assert.equal(value('Amount'), 264);
  assert.equal(value('Profit'), 87);
});

test('sales cost uses the purchase matching both PO and product', () => {
  const { context, source, wms } = createHarness();
  const orders = source.getSheetByName('Orders (Readable)');
  appendSource(orders, sourcePurchase('PO-LOW', 2, 100, 0, 0));
  appendSource(orders, sourcePurchase('PO-HIGH', 5, 100, 0, 0));
  appendSource(source.getSheetByName('Workflow'), sourceSale('PO-HIGH', 'B-HIGH', 2, 6));
  context.syncByTruckData();

  const sheet = wms.getSheetByName('SalesLog_By Truck');
  const headers = sheet.data[0];
  const row = sheet.data[1];
  const value = name => row[headers.indexOf(name)];
  assert.equal(value('Cost Price'), 110);
  assert.equal(value('Amount'), 264);
  assert.equal(value('Profit'), 44);
});

test('sales rows with invalid Indian dates are not imported', () => {
  const { context, source, wms } = createHarness();
  appendSource(source.getSheetByName('Orders (Readable)'), sourcePurchase('PO-DATE', 2, 100, 0, 0));
  appendSource(source.getSheetByName('Workflow'), sourceSale('PO-DATE', 'B-DATE', 2, 3, '31/02/2026'));
  context.syncByTruckData();

  assert.equal(wms.getSheetByName('SalesLog_By Truck').getLastRow(), 1);
});

test('installer adds one 10-minute sync trigger and preserves other triggers', () => {
  const { context, triggers, createdTriggers } = createHarness();
  const unrelated = {
    getHandlerFunction: () => 'existingImporter',
    getEventType: () => 'CLOCK'
  };
  const legacyPurchase = {
    getHandlerFunction: () => 'autoSyncPurchaseLogByTruck',
    getEventType: () => 'CLOCK'
  };
  const legacySales = {
    getHandlerFunction: () => 'autoSyncSalesAndKheraKalan',
    getEventType: () => 'CLOCK'
  };
  triggers.push(unrelated);
  triggers.push(legacyPurchase, legacySales);

  assert.equal(context.installByTruckSyncTrigger(), true);
  assert.equal(context.installByTruckSyncTrigger(), false);
  assert.equal(createdTriggers.length, 1);
  assert.equal(createdTriggers[0].handler, 'syncByTruckData');
  assert.equal(createdTriggers[0].minutes, 10);
  assert.deepEqual(triggers.map(trigger => trigger.getHandlerFunction()), [
    'existingImporter',
    'syncByTruckData'
  ]);
  assert.equal(triggers[0], unrelated);
});

test('existing trigger setup installs the dynamic sync and keeps unrelated triggers', () => {
  const { context, triggers } = createHarness();
  const unrelated = {
    getHandlerFunction: () => 'unrelatedDailyReport',
    getEventType: () => 'CLOCK'
  };
  triggers.push(unrelated, {
    getHandlerFunction: () => 'autoSyncPurchaseLogByTruck',
    getEventType: () => 'CLOCK'
  }, {
    getHandlerFunction: () => 'autoSyncSalesAndKheraKalan',
    getEventType: () => 'CLOCK'
  });

  vm.runInNewContext(fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8'), context);
  context.setupAllSyncTriggers();

  const handlers = triggers.map(trigger => trigger.getHandlerFunction());
  assert.ok(handlers.includes('syncByTruckData'));
  assert.ok(handlers.includes('syncKheraKalanSales'));
  assert.ok(handlers.includes('autoSyncPurchaseLogKheraKalan'));
  assert.ok(handlers.includes('unrelatedDailyReport'));
  assert.ok(!handlers.includes('autoSyncPurchaseLogByTruck'));
  assert.ok(!handlers.includes('autoSyncSalesAndKheraKalan'));
});

test('syncs across two distinct external spreadsheets for Purchase and Sales without Script Properties', () => {
  const purchaseSourceData = Array.from({ length: 272 }, () => []);
  purchaseSourceData.push(['PO', '', '', 'Product', '', '', 'Rate/Kg', 'Qty Kg']);
  const salesSourceData = Array.from({ length: 170 }, () => []);
  salesSourceData.push([]);

  const purchaseSpreadsheet = new FakeSpreadsheet({
    'Orders (Readable)': new FakeSheet('Orders (Readable)', purchaseSourceData)
  });
  const salesSpreadsheet = new FakeSpreadsheet({
    Workflow: new FakeSheet('Workflow', salesSourceData)
  });
  const wmsSpreadsheet = new FakeSpreadsheet({
    'PurchaseLog_By Truck': new FakeSheet('PurchaseLog_By Truck', [purchaseHeaders]),
    'SalesLog_By Truck': new FakeSheet('SalesLog_By Truck', [salesHeaders]),
    ProductWeights: new FakeSheet('ProductWeights', [
      ['Product', 'Kg per Ream'],
      ['A4', 22]
    ])
  });

  const context = {
    PropertiesService: {
      getScriptProperties: () => ({ getProperty: () => null })
    },
    SpreadsheetApp: {
      openById: id => {
        if (String(id).includes('1HuZABEjx7A68gEcJd_BmQhjTdCqX-UpVdvm51r3YhvA')) return purchaseSpreadsheet;
        if (String(id).includes('1STKk27ovhhTZYEH9Fv0FdFE81u_06C5VI2K-mipH3Lo')) return salesSpreadsheet;
        return wmsSpreadsheet;
      },
      getActiveSpreadsheet: () => wmsSpreadsheet,
      getUi: () => ({ ButtonSet: { OK: 'OK' }, alert() {} })
    },
    ScriptApp: {
      EventType: { CLOCK: 'CLOCK', ON_EDIT: 'ON_EDIT' },
      getProjectTriggers: () => [],
      deleteTrigger: () => {},
      newTrigger: () => ({ timeBased: () => ({ everyMinutes: () => ({ create: () => ({}) }) }) })
    },
    Number,
    Math,
    String,
    Error,
    Object,
    Logger: { log() {} }
  };

  const script = fs.readFileSync(path.join(__dirname, '..', 'apps-script', 'ByTruckSync.gs'), 'utf8');
  vm.runInNewContext(script, context);

  // 1. Initial purchase with 0 freight and 0 insurance
  appendSource(purchaseSpreadsheet.getSheetByName('Orders (Readable)'), sourcePurchase('PO-DISTINCT-1', 2, 100, 0, 0));
  appendSource(salesSpreadsheet.getSheetByName('Workflow'), sourceSale('PO-DISTINCT-1', 'B-DIST-1', 2, 4));

  context.syncByTruckData();

  const pSheet = wmsSpreadsheet.getSheetByName('PurchaseLog_By Truck');
  const sSheet = wmsSpreadsheet.getSheetByName('SalesLog_By Truck');
  const pHeaders = pSheet.data[0];
  const sHeaders = sSheet.data[0];

  assert.equal(pSheet.getLastRow(), 2);
  assert.equal(sSheet.getLastRow(), 2);
  assert.equal(pSheet.data[1][pHeaders.indexOf('Net Buy Ream')], 44);
  assert.equal(sSheet.data[1][sHeaders.indexOf('Cost Price')], 44);
  assert.equal(sSheet.data[1][sHeaders.indexOf('Profit')], 88); // 176 - (2 * 44) = 88

  // 2. Later, Insurance and Freight are added to the Purchase sheet
  purchaseSpreadsheet.getSheetByName('Orders (Readable)').data[273][28] = 30; // Insurance
  purchaseSpreadsheet.getSheetByName('Orders (Readable)').data[273][29] = 70; // Freight

  context.syncByTruckData();

  // 3. Purchase landed cost and Sales profit update dynamically!
  assert.equal(pSheet.getLastRow(), 2);
  assert.equal(pSheet.data[1][pHeaders.indexOf('Insurance')], 30);
  assert.equal(pSheet.data[1][pHeaders.indexOf('Cartage')], 70);
  // Net buy ream = (220 + 100) / 5 = 64
  assert.equal(pSheet.data[1][pHeaders.indexOf('Net Buy Ream')], 64);
  // Sales cost price updates to 64, profit updates to 176 - (2 * 64) = 176 - 128 = 48
  assert.equal(sSheet.data[1][sHeaders.indexOf('Cost Price')], 64);
  assert.equal(sSheet.data[1][sHeaders.indexOf('Profit')], 48);
});

test('handles unformatted or empty ProductWeights gracefully without throwing header errors', () => {
  const { context, source, wms } = createHarness();
  // Clear ProductWeights to simulate unformatted or newly created blank sheet
  const pw = wms.getSheetByName('ProductWeights');
  pw.data = [['Some Title Banner', ''], ['', '']];

  appendSource(source.getSheetByName('Orders (Readable)'), sourcePurchase('PO-PW-TEST', 2, 100, 0, 0));
  assert.doesNotThrow(() => {
    context.syncByTruckData();
  });

  const pSheet = wms.getSheetByName('PurchaseLog_By Truck');
  assert.equal(pSheet.getLastRow(), 2);
});

test('populates all columns including Mill, Warehouse, Unit Type, Bill Type, and Recovered', () => {
  const { context, source, wms } = createHarness();
  const purchaseSource = source.getSheetByName('Orders (Readable)');
  const salesSource = source.getSheetByName('Workflow');

  appendSource(purchaseSource, sourcePurchase('PO-FULL-1', 10, 110, 50, 20, 'yes', '20/11/2026'));
  appendSource(salesSource, sourceSale('PO-FULL-1', 'BILL-100', 5, 20, '21/11/2026'));

  context.syncByTruckData();

  const pSheet = wms.getSheetByName('PurchaseLog_By Truck');
  const sSheet = wms.getSheetByName('SalesLog_By Truck');
  const pHeaders = pSheet.data[0];
  const sHeaders = sSheet.data[0];

  // Verify Purchase row has no blank core fields
  assert.equal(pSheet.data[1][pHeaders.indexOf('Warehouse')], 'By Truck');
  assert.equal(pSheet.data[1][pHeaders.indexOf('Unit Type')], 'Ream');
  assert.equal(pSheet.data[1][pHeaders.indexOf('Bill Type')], 'GST');
  assert.ok(pSheet.data[1][pHeaders.indexOf('Mill')]);

  // Verify Sales row has no blank core fields
  assert.equal(sSheet.data[1][sHeaders.indexOf('Warehouse')], 'By Truck');
  assert.equal(sSheet.data[1][sHeaders.indexOf('Bill Type')], 'GST');
  assert.equal(sSheet.data[1][sHeaders.indexOf('Recovered?')], 'No');
  assert.ok(sSheet.data[1][sHeaders.indexOf('Mills')]);
});

test('preserves manually entered freight in PurchaseLog_By Truck when source has zero freight', () => {
  const { context, source, wms } = createHarness();
  const purchaseSource = source.getSheetByName('Orders (Readable)');
  const pSheet = wms.getSheetByName('PurchaseLog_By Truck');

  // Initial purchase with 0 insurance and 0 freight from form/source
  appendSource(purchaseSource, sourcePurchase('PO-MANUAL-FREIGHT', 10, 110, 0, 0, 'yes', '20/11/2026'));
  context.syncByTruckData();

  const pHeaders = pSheet.data[0];
  const cartageCol = pHeaders.indexOf('Cartage');
  const netBuyCol = pHeaders.indexOf('Net Buy Ream');
  assert.equal(pSheet.data[1][cartageCol], 0);

  // User manually fills Freight (Cartage = 500) directly in the sheet
  pSheet.data[1][cartageCol] = 500;

  // Next sync cycle runs (source still has 0 freight)
  context.syncByTruckData();

  // Verify that the manually entered freight is PRESERVED and not wiped to 0!
  assert.equal(pSheet.data[1][cartageCol], 500);
  // Verify Net Buy Ream is recalculated to include the 500 freight
  const expectedNetBuy = (pSheet.data[1][pHeaders.indexOf('Amount')] + 500) / pSheet.data[1][pHeaders.indexOf('Qty')];
  assert.equal(pSheet.data[1][netBuyCol], expectedNetBuy);
});

test('normalizes BILT COPIER PLAIN to BILT WHITE COPIER and matches sales cost across aliases', () => {
  const { context, source, wms } = createHarness();
  const purchaseSource = source.getSheetByName('Orders (Readable)');
  const salesSource = source.getSheetByName('Workflow');
  const pSheet = wms.getSheetByName('PurchaseLog_By Truck');
  const sSheet = wms.getSheetByName('SalesLog_By Truck');

  // Purchase entry entered as BILT COPIER PLAIN A4 75 GSM
  const pRow = sourcePurchase('PO-147', 10, 234, 0, 0, 'yes', '26/10/2026');
  pRow[3] = 'BILT COPIER PLAIN A4 75 GSM';
  appendSource(purchaseSource, pRow);

  // Sales entry entered as BILT WHITE COPIER A4 75 GSM
  const sRow = sourceSale('PO-147', 'BILL-147', 100, 15, '26/10/2026');
  sRow[56] = 'BILT WHITE COPIER A4 75 GSM';
  appendSource(salesSource, sRow);

  context.syncByTruckData();

  const pHeaders = pSheet.data[0];
  const sHeaders = sSheet.data[0];

  // Inward product should be normalized to BILT WHITE COPIER A4 75 GSM
  assert.equal(pSheet.data[1][pHeaders.indexOf('Product')], 'BILT WHITE COPIER A4 75 GSM');
  assert.equal(pSheet.data[1][pHeaders.indexOf('Mill')], 'BILT');

  // Sales product should be BILT WHITE COPIER A4 75 GSM
  assert.equal(sSheet.data[1][sHeaders.indexOf('Product')], 'BILT WHITE COPIER A4 75 GSM');
  assert.equal(sSheet.data[1][sHeaders.indexOf('Mills')], 'BILT');

  // Sales Cost should match the inward purchase net rate (10/kg), not fall back to rate (15/kg)
  const costPrice = sSheet.data[1][sHeaders.indexOf('Cost Price')];
  const saleRate = sSheet.data[1][sHeaders.indexOf('Rate')];
  const profit = sSheet.data[1][sHeaders.indexOf('Profit')];

  assert.ok(costPrice < saleRate, `Cost Price ${costPrice} should be lower than Sale Rate ${saleRate}`);
  assert.ok(profit > 0, `Profit ${profit} should be greater than 0`);
});

test('fixBiltCopierPlainLabels migrates existing rows in PurchaseLog_By Truck', () => {
  const { context, wms } = createHarness();
  const pSheet = wms.getSheetByName('PurchaseLog_By Truck');
  const pHeaders = pSheet.data[0];

  // Pre-existing row with BILT COPIER PLAIN A4 75 GSM in PurchaseLog
  const row = paddedRow(pHeaders.length, {
    [pHeaders.indexOf('Product')]: 'BILT COPIER PLAIN A4 75 GSM',
    [pHeaders.indexOf('PO No')]: '147',
    [pHeaders.indexOf('Mill')]: 'BILT'
  });
  pSheet.data.push(row);

  const result = context.fixBiltCopierPlainLabels();
  assert.equal(result.fixedInward, 1);
  assert.equal(pSheet.data[1][pHeaders.indexOf('Product')], 'BILT WHITE COPIER A4 75 GSM');
});

test('dedupeByTruckSheet_ removes duplicate rows when sales or purchase repeat', () => {
  const { context, wms } = createHarness();
  const pSheet = wms.getSheetByName('PurchaseLog_By Truck');
  const pHeaders = pSheet.data[0];

  // Add duplicate purchase rows for PO 147
  const row1 = paddedRow(pHeaders.length, {
    [pHeaders.indexOf('Product')]: 'BILT WHITE COPIER A4 75 GSM',
    [pHeaders.indexOf('PO No')]: '147',
    [pHeaders.indexOf('Qty')]: 100
  });
  const row2 = paddedRow(pHeaders.length, {
    [pHeaders.indexOf('Product')]: 'BILT WHITE COPIER A4 75 GSM',
    [pHeaders.indexOf('PO No')]: '147',
    [pHeaders.indexOf('Qty')]: 100
  });
  pSheet.data.push(row1, row2);
  assert.equal(pSheet.getLastRow(), 3);

  const deleted = context.dedupeByTruckSheet_(pSheet, 'purchase');
  assert.equal(deleted, 1);
  assert.equal(pSheet.getLastRow(), 2);
});

