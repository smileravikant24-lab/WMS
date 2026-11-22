# By Truck Apps Script sync

`ByTruckSync.gs` adds a scheduled refresh for the By Truck purchase and sales
logs alongside the existing WMS backend in `Code.gs`. The backend's hard-coded
spreadsheet fallback now reads `BY_TRUCK_WMS_SPREADSHEET_ID` from Script
Properties.

## Setup

1. Add both `Code.gs` and `apps-script/ByTruckSync.gs` to the existing Apps
   Script project.
2. In **Project Settings → Script properties**, set:
   - `BY_TRUCK_SOURCE_SPREADSHEET_ID`: spreadsheet containing `Orders (Readable)`
     and `Workflow`.
   - `BY_TRUCK_WMS_SPREADSHEET_ID`: spreadsheet containing the By Truck logs
     and `ProductWeights`.
3. Ensure the source tabs begin at row 273 (`Orders (Readable)`) and row 171
   (`Workflow`), and that `ProductWeights` has `Product` and `Kg per Ream`
   columns.
4. Run `installByTruckSyncTrigger()` once from the Apps Script editor and
   authorize it. It creates a dedicated 10-minute clock trigger for
   `syncByTruckData()` and is safe to rerun. It replaces only the legacy
   By Truck append-only clock triggers and leaves unrelated triggers in place.
   The existing `setupAllSyncTriggers()` also invokes the new By Truck
   installer.

The target tabs must have a header row. The sync locates destination fields by
header (including common spelling variants); it falls back to column P for
purchase PO and column R for sales PO, matching the existing layout. It adds
and hides a `__WMS_SOURCE_KEY` helper column on each target tab. The key tracks
source row identity, so changed quantities or rates update their existing log
row, including repeated products on the same PO. Older rows are matched by
PO/product (and bill for sales) in row order on their first refresh.

Purchases sync only when Transfer Status is `yes` and the invoice date is a
valid Indian date (or a valid spreadsheet Date value); sales rows with invalid
dates are skipped too. Purchase fields use the existing A:P schema, including
`Cartage`, `Net Buy Ream`, and `Net Buy Kg`. Freight/cartage and insurance are
split across a PO's product rows by kg. Conversion and amount follow the
importer formulas: quantity is rounded to a whole ream, the per-ream price is
rounded to two decimals, amount is rounded ream quantity times per-ream price,
Net Buy Ream adds allocated insurance/cartage and divides by ream quantity,
and Net Buy Kg is Net Buy Ream divided by product weight. Sales costs are
looked up by the exact PO and product. The sync updates and appends rows but
does not delete target rows when a source row is removed.
