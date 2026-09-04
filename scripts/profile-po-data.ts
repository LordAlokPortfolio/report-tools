// Office Script — run from Excel's "Automate" tab (Excel Online or desktop M365).
// Does NOT modify your data. Reads the PO history sheet in row batches (to stay
// under the Office Scripts payload limit) and writes a summary of issues to a
// new "Profile Report" sheet. Six columns (ID, PO DateRevised, SpecialRequest,
// QtyThisShip, Scanned, ShipLocalle) are skipped during analysis — left alone
// in your file, just not checked here since they don't matter for cleaning.
//
// How to run:
// 1. Open the workbook with the PO history sheet active (or edit SHEET_NAME below).
// 2. Automate tab -> New Script -> delete the placeholder code -> paste this file's contents.
// 3. Run. If you still hit the payload-limit error, lower CHUNK_SIZE and re-run.
// 4. Read the "Profile Report" sheet it creates (or replaces if run again).

function main(workbook: ExcelScript.Workbook) {
  const SHEET_NAME = ""; // leave blank to use the active sheet, or set e.g. "PO History"
  const CHUNK_SIZE = 2000; // rows per read; lower this if you still hit the payload limit

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

  // Columns excluded from analysis (kept in the file, just not checked/reported here).
  const EXCLUDED_COLS = ["ID", "PO DateRevised", "SpecialRequest", "QtyThisShip", "Scanned", "ShipLocalle"];

  const requiredCols = ["PO No", "Supplier ID", "PO Date", "ITEM NO", "INVENTORY ID"].filter(c => !EXCLUDED_COLS.includes(c));
  const dateCols = ["PO Date", "PO DateReceived", "PO DateRequired", "PO DateRevised", "ShipBy"].filter(c => !EXCLUDED_COLS.includes(c));
  const numCols = ["Quantity", "QtyReceived", "QtyThisShip", "UnitCost"].filter(c => !EXCLUDED_COLS.includes(c));
  const textCols = ["Supplier ID", "ITEM NO", "INVENTORY ID", "SpecialRequest", "Tag", "ShipLocalle", "Category", "PO No"].filter(c => !EXCLUDED_COLS.includes(c));
  const distinctCols = ["POLineClosed", "Scanned", "Tag", "Category", "ShipLocalle"].filter(c => !EXCLUDED_COLS.includes(c));

  const issueMap = new Map<string, number[]>();
  function flag(type: string, column: string, rowNum: number) {
    const key = type + "||" + column;
    if (!issueMap.has(key)) issueMap.set(key, []);
    issueMap.get(key).push(rowNum);
  }

  const dupKeyRows = new Map<string, number[]>();
  const distinctValues = new Map<string, Map<string, number>>();
  for (const col of distinctCols) distinctValues.set(col, new Map<string, number>());

  let total = 0;

  for (let start = 1; start < totalRows; start += CHUNK_SIZE) {
    const rowsInChunk = Math.min(CHUNK_SIZE, totalRows - start);
    const chunk = sheet.getRangeByIndexes(start, 0, rowsInChunk, totalCols).getValues();
    total += chunk.length;

    for (let i = 0; i < chunk.length; i++) {
      const row = chunk[i];
      const rowNum = start + i + 1; // 1-based sheet row number

      for (const col of requiredCols) {
        const idx = headerIndex[col];
        if (idx === undefined) continue;
        const v = row[idx];
        if (v === "" || v === null || v === undefined) {
          flag("Blank required field", col, rowNum);
        }
      }

      for (const col of dateCols) {
        const idx = headerIndex[col];
        if (idx === undefined) continue;
        const v = row[idx];
        if (v === "" || v === null || v === undefined) continue;
        if (typeof v !== "number") {
          flag("Unparseable date", col, rowNum);
        }
      }

      for (const col of numCols) {
        const idx = headerIndex[col];
        if (idx === undefined) continue;
        const v = row[idx];
        if (v === "" || v === null || v === undefined) continue;
        if (typeof v !== "number") {
          flag("Non-numeric value", col, rowNum);
        } else {
          if (v < 0) flag("Negative value", col, rowNum);
          if (col === "UnitCost" && v === 0) flag("Zero unit cost", col, rowNum);
        }
      }

      for (const col of textCols) {
        const idx = headerIndex[col];
        if (idx === undefined) continue;
        const v = row[idx];
        if (typeof v === "string" && v !== v.trim()) {
          flag("Leading/trailing whitespace", col, rowNum);
        }
      }

      const poIdx = headerIndex["PO No"];
      const itemIdx = headerIndex["ITEM NO"];
      if (poIdx !== undefined && itemIdx !== undefined) {
        const key = String(row[poIdx]) + "||" + String(row[itemIdx]);
        if (!dupKeyRows.has(key)) dupKeyRows.set(key, []);
        dupKeyRows.get(key).push(rowNum);
      }

      for (const col of distinctCols) {
        const idx = headerIndex[col];
        if (idx === undefined) continue;
        const v = row[idx];
        const key = v === null || v === undefined || v === "" ? "(blank)" : String(v);
        const m = distinctValues.get(col);
        m.set(key, (m.get(key) || 0) + 1);
      }
    }
  }

  dupKeyRows.forEach(rowNums => {
    if (rowNums.length > 1) {
      rowNums.forEach(r => flag("Duplicate PO No + ITEM NO", "PO No / ITEM NO", r));
    }
  });

  const existing = workbook.getWorksheet("Profile Report");
  if (existing) existing.delete();
  const reportSheet = workbook.addWorksheet("Profile Report");

  const outRows: (string | number)[][] = [];
  outRows.push(["Issue Type", "Column", "Count", "Sample Rows (up to 5)"]);

  const entries = Array.from(issueMap.entries())
    .map(([key, rowNums]) => {
      const [type, column] = key.split("||");
      return { type, column, count: rowNums.length, samples: rowNums.slice(0, 5) };
    })
    .sort((a, b) => b.count - a.count);

  for (const e of entries) {
    outRows.push([e.type, e.column, e.count, e.samples.join(", ")]);
  }

  outRows.push(["", "", "", ""]);
  outRows.push(["--- Distinct value counts (categorical columns) ---", "", "", ""]);
  distinctValues.forEach((m, col) => {
    outRows.push([`Column: ${col}`, "", "", ""]);
    const sorted = Array.from(m.entries()).sort((a, b) => b[1] - a[1]);
    for (const [val, count] of sorted) {
      outRows.push(["", val, count, ""]);
    }
  });

  outRows.push(["", "", "", ""]);
  outRows.push([`Total data rows scanned: ${total}`, "", "", ""]);

  const outRange = reportSheet.getRangeByIndexes(0, 0, outRows.length, 4);
  outRange.setValues(outRows);
  reportSheet.getUsedRange().getFormat().autofitColumns();
  reportSheet.activate();
}
