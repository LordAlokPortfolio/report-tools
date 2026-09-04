// Office Script — run from Excel's "Automate" tab (Excel Online or desktop M365).
// Cleans the PO history sheet IN PLACE -- overwrites the sheet's own cells
// with cleaned, reordered values. No separate output sheet, no report.
// Full rule list is documented in CLEANING-LOG.md; keep that file in sync
// with this script.
//
// Rules applied:
// 1. PO Date, PO DateReceived: strip time-of-day, keep date only.
// 2. PO DateReceived = "NULL":
//      POLineClosed = -1 (closed) -> fill from PO DateRequired (same row).
//      POLineClosed = 0  (open)   -> fill from PO Date + this supplier's
//                                     median working-day lead time, computed
//                                     from that supplier's other closed
//                                     orders (real PO Date -> PO DateReceived
//                                     pairs). No historical data for that
//                                     supplier -> leave as "NULL".
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
    "PO DateReceived", "PO DateRequired", "PO DateRevised", "QtyReceived",
    "POLineClosed", "UnitCost", "Category",
  ];
  const SIDE_COLS = ["ID", "SpecialRequest", "QtyThisShip", "Scanned", "Tag", "ShipLocalle", "ShipBy"];
  const outputCols = [...NECESSARY_COLS, ...SIDE_COLS].filter(c => headerIndex[c] !== undefined);

  function isNullCell(v: string | number | boolean): boolean {
    return typeof v === "string" && v.trim().toUpperCase() === "NULL";
  }
  function excelSerialToUTCDate(serial: number): Date {
    return new Date(Date.UTC(1899, 11, 30) + Math.round(serial) * 86400000);
  }
  function isWeekendSerial(daySerial: number): boolean {
    const dow = excelSerialToUTCDate(daySerial).getUTCDay();
    return dow === 0 || dow === 6;
  }
  function workingDaysBetween(startSerial: number, endSerial: number): number {
    let count = 0;
    const s = Math.floor(startSerial);
    const e = Math.floor(endSerial);
    for (let day = s + 1; day <= e; day++) {
      if (!isWeekendSerial(day)) count++;
    }
    return count;
  }
  function median(nums: number[]): number {
    const sorted = [...nums].sort((a, b) => a - b);
    const mid = Math.floor(sorted.length / 2);
    return sorted.length % 2 === 0 ? (sorted[mid - 1] + sorted[mid]) / 2 : sorted[mid];
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

  // Pass 1: collect each supplier's working-day lead time from their closed
  // orders with real (non-NULL) PO Date / PO DateReceived pairs.
  const leadTimesBySupplier = new Map<string, number[]>();

  for (let start = 1; start < totalRows; start += CHUNK_SIZE) {
    const rowsInChunk = Math.min(CHUNK_SIZE, totalRows - start);
    const chunk = sheet.getRangeByIndexes(start, 0, rowsInChunk, totalCols).getValues();

    for (const row of chunk) {
      const closed = row[poLineClosedIdx] === -1;
      const poDate = row[poDateIdx];
      const poDateReceived = row[poDateReceivedIdx];
      if (closed && typeof poDate === "number" && typeof poDateReceived === "number") {
        const supplier = String(row[supplierIdx]);
        const days = workingDaysBetween(poDate, poDateReceived);
        if (!leadTimesBySupplier.has(supplier)) leadTimesBySupplier.set(supplier, []);
        leadTimesBySupplier.get(supplier).push(days);
      }
    }
  }

  const medianLeadTimeBySupplier = new Map<string, number>();
  leadTimesBySupplier.forEach((days, supplier) => {
    medianLeadTimeBySupplier.set(supplier, median(days));
  });

  // Pass 2: clean each row and write it back to the SAME sheet, reordered.
  sheet.getRangeByIndexes(0, 0, 1, outputCols.length).setValues([outputCols]);

  for (let start = 1; start < totalRows; start += CHUNK_SIZE) {
    const rowsInChunk = Math.min(CHUNK_SIZE, totalRows - start);
    const chunk = sheet.getRangeByIndexes(start, 0, rowsInChunk, totalCols).getValues();
    const cleanedChunk: (string | number | boolean)[][] = [];

    for (const row of chunk) {
      const cleaned = row.slice();
      const closed = row[poLineClosedIdx] === -1;
      const open = row[poLineClosedIdx] === 0;

      // Rule: PO DateReceived = NULL
      if (isNullCell(cleaned[poDateReceivedIdx])) {
        if (closed) {
          cleaned[poDateReceivedIdx] = cleaned[poDateRequiredIdx];
        } else if (open) {
          const supplier = String(row[supplierIdx]);
          const poDate = row[poDateIdx];
          const supplierMedian = medianLeadTimeBySupplier.get(supplier);
          if (typeof poDate === "number" && supplierMedian !== undefined) {
            cleaned[poDateReceivedIdx] = Math.floor(poDate) + Math.ceil(supplierMedian);
          } else {
            cleaned[poDateReceivedIdx] = "NULL";
          }
        }
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
