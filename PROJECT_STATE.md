# Project State — paint-report-reader.html

## 1. What Exists

**paint-report-reader.html** is a single-file, fully local (no server, no internet after page load) HTML tool that reads Da Vinci 360 "PaintReport Production" PDFs and exports the data to a formatted Excel workbook.

**Parsing strategy (ExType-anchor):** the parser does not try to interpret the report top-down. It scans every normalized line for one regex anchor:

```
^ExType (.+?) Total Length\(ft\) ([\d,]*\.?\d+) Lengths (\d+)$
```

Every match is a data line. The parser then pairs it with the line directly above it (the "colour line," format `{ExType}-{Brand} {Colour}~{Finish} | {Secondary} ... Profile {code}`), recovering the ExType by testing each `-{brand} ` boundary in the colour line against the anchor's combined ExType+Description string. This avoids assuming ExType is a fixed-format 4-digit code (real ExTypes include `1118-REMOVE Fin`, `2343Ext`, `1033A`) and avoids assuming colours are single words (`Comm Brown`, `Windswept Smoke`).

**Mandatory checksum validation** (`validateReport`) — a partial or misread parse is never allowed to look like a success:
- Report's own `Summary for 'Production' = {batch} ({N} detail records)` line must be found; batch here must match the `Production {batch}` header.
- Report's own `Sum {total}` line must be found.
- Structural pairing check: count of colour lines (regex `-\S+\s+[^~]*~`) must equal count of parsed records 1:1 — any orphan colour or data line fails loudly.
- Parsed sum vs. report's `Sum` line must be within tolerance `0.05 × rowCount + 0.005` (accounts for the report's per-row footage being rounded to 1 decimal while `Sum` carries 2).
- Any failure disables the results table and Excel export; the page shows a red status listing every specific mismatch.

**Output fields per record:** `exType, colour, finish, secondaryColour, description, totalLengthFt, lengthsCount, bothSides` (true when secondary colour is absent — same paint both sides). Report-level: `batch, reportDate, detailRecordsN`. Excel export adds `Batch#`, `ReportDate`, and `Gallons` (= display footage ÷ `FT_PER_GALLON`, doubled only for `bothSides` rows — doubling applies to gallons/footage-for-paint-consumption, never to linear feet or piece counts in the raw `TotalLengthFt`/`LengthsCount` fields).

Multi-file batches consolidate into one workbook: sheet "Paint Report" (plain data + TOTAL row) and a separate sheet "PaintSummary" (bordered, colour+finish rollup with a dark title bar and blue header row, styled via the `xlsx-js-style` library since the stock `xlsx` library silently drops all cell styling).

## 2. Test Results (actually observed)

**PaintReport070901.pdf** — parser run against it directly (source extracted verbatim from the current HTML) via Node/pdf.js:

- Checksum: **PASS**
- Batch: 48973 — matches expected 48973
- Detail records N (from report's own summary line): 107 — matches expected 107
- Printed/parsed rows (aggregated records, not raw detail rows): 19
- Parsed sum: 742.60 ft
- Report's own Sum line: 742.63 ft — matches expected 742.63
- Diff 0.03 ft, well inside computed tolerance 0.955

**PaintReport063005.pdf** (expected batch 48801, 157 records, 1,105.37 ft) — **NOT TESTED.** This file does not exist anywhere in the dev environment or the repo. No `/test-samples/` directory exists in the repo either. This result is UNVERIFIED, not PASS, not FAIL — the sample was never available to test against.

Note: "107 detail records" in the report's summary line refers to underlying database detail rows, not printed/output rows — the parser correctly reports 19 aggregated rows while independently matching the 107 count from the report's own summary text (used only as a checksum, not as the row count of the output table).

## 3. Known Limitations

- **OCR path (Tesseract.js) is implemented but untested.** It's wired to trigger automatically when a PDF page has no extractable text layer (scanned/image PDF) or when an image file is uploaded directly. No scanned sample was available to verify OCR output against the checksum validator — if OCR text doesn't match the parser's exact line format, it will fail the checksum loudly (by design) rather than silently produce wrong numbers, but this fallback path has not been exercised end-to-end.
- **`FT_PER_GALLON = 400` is a placeholder**, not a real paint-coverage constant. It's a single flat number used for every colour/finish/profile combination. Real gallon estimates need per-surface, per-profile coverage rates, which do not exist yet in this tool.
- **Fragile:** the ExType/colour-line pairing depends on strict line adjacency (data line directly follows its colour line with zero lines between). Any PDF export variant that inserts a blank line, a wrapped description, or reorders these two lines will break pairing and fail the checksum (intentional — no silent partial success — but it means any new report layout needs a fresh sample run through the checksum before trusting it).
- Only one real sample PDF has been checksum-verified end-to-end. The tool has not been run against multiple different brand/report format variants beyond what's noted in git history (KV, GENTEK prefixes).

## 4. Strategic Note

PIVOT UNDER EVALUATION: The PDF is generated by Da Vinci 360 (ERP). If Da Vinci can export this report directly as Excel/CSV, parsing becomes unnecessary for the primary path. Owner is checking the Da Vinci export dialog manually. If Excel export exists, next build = extend the existing Selenium downloader (Dv360_downloader_v2.py, separate codebase, not in this repo) to fetch Excel directly. This HTML parser then becomes the fallback for when only a PDF or a photo of the report is available (e.g., manager forwards a printout). DO NOT delete or deprecate the parser.
