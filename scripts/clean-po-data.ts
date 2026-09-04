// Office Script — run from Excel's "Automate" tab (Excel Online or desktop M365).
// Cleans the raw PO history sheet according to the rules documented in
// CLEANING-LOG.md. Does NOT modify the source sheet -- writes a cleaned,
// reordered copy to a new "Clean Data" sheet (replaced if this is re-run).
//
// Rules applied (see CLEANING-LOG.md for the full writeup):
// 1. PO Date, PO DateReceived: strip time-of-day, keep date only.
// 2. PO DateReceived = "NULL":
//      POLineClosed = -1 (closed) -> fill from PO DateRequired (same row).
//      POLineClosed = 0  (open)   -> fill from PO Date + this supplier's
//                                     median working-day lead time, computed
//                                     from that supplier's other closed
//                                     orders (real PO Date -> PO DateReceived
//                                     pairs). No historical data for that
//                                     supplier -> leave as "NULL" (never
//                                     invent a number with nothing behind it).
// 3. QtyReceived = 0:
//      POLineClosed = 0  (open)   -> leave as 0 (order genuinely not received yet).
//      POLineClosed = -1 (closed) -> replace with this row's Quantity value
//                                     (closed order implies fully received).
// 4. PO DateRevised: untouched. No rule applied.
// 5. Columns not used by any report view are moved to the right side of the
//    output table, not deleted, not hidden from the file -- just out of the
//    way of the columns the reports actually read.
//
// Limitation: "working days" here means Mon-Fri only. No company holiday
// calendar is available to this script, so statutory holidays are not
// excluded from the day count. Document this if it matters to a report.
//
// How to run:
// 1. Open the workbook with the raw PO history sheet active (or edit SHEET_NAME below).
// 2. Automate tab -> open this script -> Run.
// 3. If you hit the payload-limit error, lower CHUNK_SIZE and re-run.
// 4. Read the new "Clean Data" sheet. The original sheet is untouched.

function main(workbook: ExcelScript.Workbook) {
  const SHEET_NAME = ""; // leave blank to use the active sheet, or set e.g. "PO History"
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

  // Columns the report views actually use, kept on the left of the output table.
  const NECESSARY_COLS = [
    "PO No", "Supplier ID", "PO Date", "ITEM NO", "Quantity", "INVENTORY ID",
    "PO DateReceived", "PO DateRequired", "PO DateRevised", "QtyReceived",
    "POLineClosed", "UnitCost", "Category",
  ];
  // Not used by any report view -- kept in the output, just moved to the right.
  const SIDE_COLS = ["ID", "SpecialRequest", "QtyThisShip", "Scanned", "Tag", "ShipLocalle", "ShipBy"];

  const outputCols = [...NECESSARY_COLS, ...SIDE_COLS].filter(c => headerIndex[c] !== undefined);

  function isNullCell(v: string | number | boolean): boolean {
    return typeof v === "string" && v.trim().toUpperCase() === "NULL";
  }

  // Excel serial date (with or without a time fraction) -> whole calendar day, as an integer serial.
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

  const poNoIdx = headerIndex["PO No"];
  const supplierIdx = headerIndex["Supplier ID"];
  const poDateIdx = headerIndex["PO Date"];
  const poDateReceivedIdx = headerIndex["PO DateReceived"];
  const poDateRequiredIdx = headerIndex["PO DateRequired"];
  const poLineClosedIdx = headerIndex["POLineClosed"];
  const quantityIdx = headerIndex["Quantity"];
  const qtyReceivedIdx = headerIndex["QtyReceived"];

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

  // Pass 2: build the cleaned, reordered output.
  const outHeader = outputCols;
  const outRows: (string | number | boolean)[][] = [outHeader];

  for (let start = 1; start < totalRows; start += CHUNK_SIZE) {
    const rowsInChunk = Math.min(CHUNK_SIZE, totalRows - start);
    const chunk = sheet.getRangeByIndexes(start, 0, rowsInChunk, totalCols).getValues();

    for (const row of chunk) {
      const cleaned = row.slice(); // copy; only touch the specific cells the rules cover
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
            cleaned[poDateReceivedIdx] = "NULL"; // uncalculable -- no supplier history, leave as NULL
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

      // Rule: QtyReceived = 0
      if (cleaned[qtyReceivedIdx] === 0) {
        if (closed) {
          cleaned[qtyReceivedIdx] = cleaned[quantityIdx];
        }
        // open -> leave as 0, that's correct (not received yet)
      }

      // PO DateRevised: untouched, no rule.

      outRows.push(outputCols.map(col => cleaned[headerIndex[col]]));
    }
  }

  const existing = workbook.getWorksheet("Clean Data");
  if (existing) existing.delete();
  const outSheet = workbook.addWorksheet("Clean Data");

  for (let start = 0; start < outRows.length; start += CHUNK_SIZE) {
    const rowsInChunk = Math.min(CHUNK_SIZE, outRows.length - start);
    const slice = outRows.slice(start, start + rowsInChunk);
    outSheet.getRangeByIndexes(start, 0, rowsInChunk, outputCols.length).setValues(slice);
  }

  outSheet.getUsedRange().getFormat().autofitColumns();
  outSheet.activate();
}
