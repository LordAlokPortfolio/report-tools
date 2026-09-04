# PO History — Cleaning Decision Log

Working log of cleaning rules for the raw PO history file (KV Custom Windows
and Doors), built from a column-by-column profiling pass. Companion to
`CLEANING-PO-INVENTORY-TABLE-AND-VENDOR-ANALYSIS-SPEC.md` — that file
describes the reporting tool; this one describes how the raw file gets from
"messy" to "the clean file the tool assumes is sitting there."

**Rule for this log: rules only, never data.** No real row values, no sample
values, no counts derived from the live file get written here. Only the
decision and the reasoning behind it.

These rules repeat every time a new PO export gets cleaned, so they're
written down precisely enough to re-apply without re-deriving them from
scratch. `scripts/clean-po-data.ts` is the current implementation.

---

## Columns excluded from cleaning entirely

Left untouched in the source file, moved to the right side of the output
table (not deleted, not hidden from the file) by `clean-po-data.ts` because
no report view uses them.

| Column | Reason |
|---|---|
| `SpecialRequest` | Free text, not used by any report view. |
| `QtyThisShip` | Not used by any report view (yet). |
| `Scanned` | Not used by any report view. |
| `ShipLocalle` | Not used by any report view. |
| `Tag` | Not used by any report view. |
| `ShipBy` | Unparseable as a date in 100% of rows — not a usable date column, and out of scope for cleaning. |

`PO DateRevised` is kept in the output table (it's part of the necessary
columns) but no cleaning rule is applied to it — see "Explicitly not
cleaned" below.

## Resolved rules

### UnitCost = 0
**Decision: legitimate value, not an error.** Zero-cost lines are real
business cases (samples, internal transfers, etc.). The cleaning pipeline
must NOT drop, blank, or flag these rows as bad data. Report math that
divides or averages by UnitCost needs to handle 0 without treating it as
missing.

### Category = "0"
**Decision: real category code, leave alone.** Even though it's the single
most common value (~55% of rows), it is not being treated as "uncategorized"
or blanked out. No transformation applied.

### PO Date, PO DateReceived — strip time-of-day
**Decision: keep date only.** Both columns carry a time component that
isn't needed. Rule: if the cell is a number (a real Excel date/time serial),
floor it to drop the time fraction. Applied after the NULL-fill rules below,
so a filled-in PO DateReceived also gets floored.

### PO DateReceived = "NULL"
Two-case rule based on `POLineClosed` (`-1` = closed, `0` = open):

- **Closed (`POLineClosed = -1`)** → fill from that row's `PO DateRequired`
  value.
- **Open (`POLineClosed = 0`)** → fill from `PO Date` + this supplier's
  median working-day lead time. The median is computed from that same
  supplier's *other* closed orders where both `PO Date` and `PO DateReceived`
  are real (non-NULL) dates — working days = Mon–Fri, counted between the
  two dates. **If that supplier has no such historical orders, leave the
  cell as `"NULL"`** — never invent a lead time with nothing behind it, per
  the "only what the data can prove" rule in the spec.
- No minimum sample-size threshold is enforced (unlike the 6–8 order
  minimum described for the *reporting* tool's Speed/Pattern views in the
  spec) — even a single historical closed order for a supplier is used as
  its median. If this turns out to produce noisy fills, revisit.

### QtyReceived = 0
Two-case rule based on `POLineClosed`:

- **Open (`POLineClosed = 0`)** → leave as `0`. This is correct, not an
  error — the order genuinely hasn't been received yet.
- **Closed (`POLineClosed = -1`)** → replace `0` with that row's `Quantity`
  value. A closed order implies it was fully received, so `QtyReceived`
  should equal `Quantity`.

### INVENTORY ID — leading/trailing whitespace
**Decision: trim.** Uncontroversial — whitespace-only difference, trimming
can't lose information.

### Non-numeric QtyReceived
**Decision: replace with that row's `Quantity` value.** Same resolution as
the `QtyReceived = 0` + closed case above — if `QtyReceived` isn't a usable
number at all, treat it the same way: fall back to `Quantity`.

### ID — resolved via lookup against the "Da Vinci File" sheet
**Decision: no longer excluded.** The workbook now also holds a "Da Vinci
File" reference sheet (inventory master: `ID, INVENTORY ID, INVENTORY NAME,
DESCRIPTION, ...`) whose `ID` column is the actual unique inventory-code ID
used elsewhere (the user's history table). Rule: match this sheet's
`INVENTORY ID` (after whitespace trim) against Da Vinci File's `INVENTORY
ID`, and replace this sheet's `ID` cell with Da Vinci File's `ID` value for
that match. **No match found → left unchanged**, not blanked or guessed.

### Supplier ID — resolved via lookup against the "Supplier File" sheet
**Decision: `Supplier ID` cells are overwritten with the supplier's name.**
The workbook also holds a "Supplier File" reference sheet (`SUPPLIER ID,
SUPPLIER NAME, ADDRESS, ...`). Rule: match `Supplier ID` (after whitespace
trim) against Supplier File's `SUPPLIER ID`, and replace the cell with
Supplier File's `SUPPLIER NAME`. **No match found → left unchanged** (still
the original code, not blanked).

**Explicitly considered and rejected:** keeping the Supplier ID code
alongside a new Supplier Name column, so the code (a stable join key) isn't
lost and the lookup can be safely re-run later. User confirmed the
overwrite is intentional anyway — noting the tradeoff here so it isn't
forgotten if this ever needs re-joining to Supplier ID elsewhere.

### Explicitly NOT cleaned / no rule

- **`PO DateRevised`** — an earlier draft of this log had a rule for this
  column (NULL + closed → copy the `Quantity` value into it). That rule was
  flagged as suspicious (it would put a quantity number into a date column)
  and was **dropped** when the rules were restated. No cleaning is applied
  to `PO DateRevised`.
- **Duplicate PO No + ITEM NO** — decision: ignore, no rule. (Earlier
  hypothesis was partial shipments vs. true duplicates; not pursued further.)
- **Negative UnitCost** — decision: ignore, no rule. (Earlier hypothesis was
  returns/credits; not pursued further.)

---

## Known limitation

**No company holiday calendar.** "Working days" in the lead-time
calculation means Monday–Friday only — no Ontario statutory holidays are
excluded, unlike the business-day math described for the earlier
PowerShell/Python vendor tools in the spec. This can overstate a computed
lead time slightly around holidays. Add a holiday list if this needs to be
exact.

## Process

1. `scripts/clean-po-data.ts` runs in Excel (Automate tab) against the raw
   PO history sheet. **It edits the sheet in place** — overwrites the
   sheet's own cells with cleaned, reordered values. No separate output
   sheet, no report. Back up the file before running.
2. Any new issue type found gets discussed here before a rule is written —
   business meaning first, code second.
3. Record the resolved rule here (no data, ever).
4. All currently known issue types are resolved and implemented as of this
   writing. Re-profile after a new raw export lands, since a different file
   may surface new issue types.
