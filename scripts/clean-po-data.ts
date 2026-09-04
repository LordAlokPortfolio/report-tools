// Office Script - run from Excel's "Automate" tab (Excel Online or desktop M365).
// Cleans the PO history sheet IN PLACE -- overwrites the sheet's own cells
// with cleaned, reordered values. No separate output sheet, no report.
// Full rule list is documented in CLEANING-LOG.md; keep that file in sync
// with this script.
//
// Rules applied:
// 1. PO Date, PO DateReceived: strip time-of-day, keep date only.
// 2. PO DateReceived = "NULL":
//      POLineClosed = -1 (closed) -> fill from PO DateRequired (same row).
//                                     PO DateRequired is confirmed unreliable in
//                                     general and used NOWHERE else in this
//                                     script or the reporting tool -- this one
//                                     narrow fallback is intentional (least-bad
//                                     stand-in for a genuinely missing date).
//      POLineClosed = 0  (open)   -> left as "NULL". Never guessed -- a fabricated
//                                     date written into a still-open order would
//                                     freeze into permanent data and later get
//                                     read as a real receipt date if the order
//                                     closes without this cell being updated,
//                                     silently corrupting downstream lead-time
//                                     figures.
// 3. QtyReceived = 0:
//      POLineClosed = 0  (open)   -> leave as 0 (order genuinely not received yet).
//      POLineClosed = -1 (closed) -> replace with this row's Quantity value.
// 4. QtyReceived is non-numeric (not 0, just not a number) -> replace with
//    this row's Quantity value.
// 5. INVENTORY ID: trim leading/trailing whitespace.
// 6. Duplicate PO No + ITEM NO: ignored, no rule.
// 7. Negative UnitCost: ignored, no rule.
// 8. PO DateRevised: untouched, no rule.
// 9. Columns not used by any report view are moved to the right side of the
//    table (still in the file, just reordered).
// 10. ID: looked up from the "Da Vinci File" sheet by matching this row's
//     INVENTORY ID against that sheet's INVENTORY ID column, and replaced
//     with that sheet's ID value. No match found -> left unchanged.
// 11. Supplier ID: looked up from the "Supplier File" sheet by matching this
//     row's Supplier ID against that sheet's SUPPLIER ID column, and
//     REPLACED with that sheet's SUPPLIER NAME (the code is overwritten,
//     not kept alongside). No match found -> left unchanged.
//
// Limitation: "working days" here means Mon-Fri only. No company holiday
// calendar is available to this script.
//
// How to run:
// 1. BACK UP THE FILE FIRST -- this script overwrites the sheet's own data.
// 2. Open the workbook with the raw PO history sheet active (or edit SHEET_NAME below).
//    The workbook must also contain the "Da Vinci File" and "Supplier File" sheets.
// 3. Automate tab -> open this script -> Run.
// 4. If you hit the payload-limit error, lower CHUNK_SIZE and re-run.

function main(workbook: ExcelScript.Workbook) {
  const SHEET_NAME = ""; // leave blank to use the active sheet, or set e.g. "PO History"
  const DA_VINCI_SHEET_NAME = "Da Vinci File";
  const SUPPLIER_SHEET_NAME = "SUPPLIER FILE";
  const CHUNK_SIZE = 2000; // rows per read/write; lower this if you hit the payload limit

  const sheet = SHEET_NAME
    ? workbook.getWorksheet(SHEET_NAME)
    : workbook.getActiveWorksheet();

  const usedRange = sheet.getUsedRange();
  const totalRows = usedRange.getRowCount();
  const totalCols = usedRange.getColumnCount();

  const headerRow = sheet.getRangeByIndexes(0, 0, 1, totalCols).getValues()[0];
  const headers = headerRow.map(h => String(h).trim());
  const headerIndex: { [key: string]: number } = {};
  headers.forEach((h, i) => (headerIndex[h] = i));

  const NECESSARY_COLS = [
    "PO No", "Supplier ID", "PO Date", "ITEM NO", "Quantity", "INVENTORY ID",
    "PO DateReceived", "PO DateRevised", "QtyReceived",
    "POLineClosed", "UnitCost", "Category",
  ];
  // PO DateRequired: confirmed unreliable (purchaser-entered) and not used by
  // any report view. Still read internally by one narrow cleaning fallback
  // below (closed order, no real PO DateReceived) -- kept in the file, just
  // moved out of the report-facing columns.
  const SIDE_COLS = ["ID", "PO DateRequired", "SpecialRequest", "QtyThisShip", "Scanned", "Tag", "ShipLocalle", "ShipBy"];
  const outputCols = [...NECESSARY_COLS, ...SIDE_COLS].filter(c => headerIndex[c] !== undefined);

  function isNullCell(v: string | number | boolean): boolean {
    return typeof v === "string" && v.trim().toUpperCase() === "NULL";
  }
  const supplierIdx = headerIndex["Supplier ID"];
  const poDateIdx = headerIndex["PO Date"];
  const poDateReceivedIdx = headerIndex["PO DateReceived"];
  const poDateRequiredIdx = headerIndex["PO DateRequired"];
  const poLineClosedIdx = headerIndex["POLineClosed"];
  const quantityIdx = headerIndex["Quantity"];
  const qtyReceivedIdx = headerIndex["QtyReceived"];
  const inventoryIdIdx = headerIndex["INVENTORY ID"];
  const idIdx = headerIndex["ID"];

  // Build lookup: INVENTORY ID -> ID, from the Da Vinci File sheet.
  const daVinciSheet = workbook.getWorksheet(DA_VINCI_SHEET_NAME);
  const daVinciRange = daVinciSheet.getUsedRange();
  const daVinciValues = daVinciRange.getValues();
  const daVinciHeaders = daVinciValues[0].map(h => String(h).trim());
  const dvInventoryIdCol = daVinciHeaders.indexOf("INVENTORY ID");
  const dvIdCol = daVinciHeaders.indexOf("ID");
  const idByInventoryId = new Map<string, string | number>();
  for (let i = 1; i < daVinciValues.length; i++) {
    const key = String(daVinciValues[i][dvInventoryIdCol]).trim();
    idByInventoryId.set(key, daVinciValues[i][dvIdCol] as string | number);
  }

  // Build lookup: SUPPLIER ID -> SUPPLIER NAME, from the Supplier File sheet.
  const supplierSheet = workbook.getWorksheet(SUPPLIER_SHEET_NAME);
  const supplierRange = supplierSheet.getUsedRange();
  const supplierValues = supplierRange.getValues();
  const supplierHeaders = supplierValues[0].map(h => String(h).trim());
  const sfSupplierIdCol = supplierHeaders.indexOf("SUPPLIER ID");
  const sfSupplierNameCol = supplierHeaders.indexOf("SUPPLIER NAME");
  const nameBySupplierId = new Map<string, string | number>();
  for (let i = 1; i < supplierValues.length; i++) {
    const key = String(supplierValues[i][sfSupplierIdCol]).trim();
    nameBySupplierId.set(key, supplierValues[i][sfSupplierNameCol] as string | number);
  }

  // Clean each row and write it back to the SAME sheet, reordered.
  sheet.getRangeByIndexes(0, 0, 1, outputCols.length).setValues([outputCols]);

  for (let start = 1; start < totalRows; start += CHUNK_SIZE) {
    const rowsInChunk = Math.min(CHUNK_SIZE, totalRows - start);
    const chunk = sheet.getRangeByIndexes(start, 0, rowsInChunk, totalCols).getValues();
    const cleanedChunk: (string | number | boolean)[][] = [];

    for (const row of chunk) {
      const cleaned = row.slice();
      const closed = row[poLineClosedIdx] === -1;

      // Rule: PO DateReceived = NULL
      // PO DateRequired is confirmed unreliable in general, but this one narrow
      // fallback is intentional -- for a closed order with no real receipt date
      // at all, it's still the least-bad stand-in available. Not used anywhere
      // else in this script or the reporting tool.
      // Open orders: left as "NULL", never guessed. A guessed date written into a
      // still-open order's real cell would freeze into permanent data -- if the
      // order later actually closes, this script only fills a NULL cell, so a
      // stale guess would sit there forever masquerading as a real receipt date
      // and silently corrupt every lead-time figure that reads it.
      if (isNullCell(cleaned[poDateReceivedIdx]) && closed) {
        cleaned[poDateReceivedIdx] = cleaned[poDateRequiredIdx];
      }

      // Rule: PO Date / PO DateReceived -- strip time-of-day
      if (typeof cleaned[poDateIdx] === "number") {
        cleaned[poDateIdx] = Math.floor(cleaned[poDateIdx] as number);
      }
      if (typeof cleaned[poDateReceivedIdx] === "number") {
        cleaned[poDateReceivedIdx] = Math.floor(cleaned[poDateReceivedIdx] as number);
      }

      // Rule: QtyReceived = 0 (closed -> Quantity, open -> leave as 0)
      if (cleaned[qtyReceivedIdx] === 0) {
        if (closed) {
          cleaned[qtyReceivedIdx] = cleaned[quantityIdx];
        }
      } else if (typeof cleaned[qtyReceivedIdx] !== "number") {
        // Rule: QtyReceived non-numeric -> Quantity
        cleaned[qtyReceivedIdx] = cleaned[quantityIdx];
      }

      // Rule: INVENTORY ID -- trim whitespace
      if (typeof cleaned[inventoryIdIdx] === "string") {
        cleaned[inventoryIdIdx] = (cleaned[inventoryIdIdx] as string).trim();
      }

      // Rule: ID -- looked up from Da Vinci File by INVENTORY ID, replaced. No match -> unchanged.
      const invIdKey = String(cleaned[inventoryIdIdx]).trim();
      if (idByInventoryId.has(invIdKey)) {
        cleaned[idIdx] = idByInventoryId.get(invIdKey);
      }

      // Rule: Supplier ID -- looked up from Supplier File by Supplier ID, REPLACED with the
      // supplier name. No match -> unchanged (still the original code).
      const supplierKey = String(row[supplierIdx]).trim();
      if (nameBySupplierId.has(supplierKey)) {
        cleaned[supplierIdx] = nameBySupplierId.get(supplierKey);
      }

      // Duplicate PO No + ITEM NO: ignored, no rule.
      // Negative UnitCost: ignored, no rule.
      // PO DateRevised: untouched, no rule.

      cleanedChunk.push(outputCols.map(col => cleaned[headerIndex[col]]));
    }

    sheet.getRangeByIndexes(start, 0, rowsInChunk, outputCols.length).setValues(cleanedChunk);
  }

  sheet.getUsedRange().getFormat().autofitColumns();
}
