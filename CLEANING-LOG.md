# PO History — Cleaning Decision Log

Working log of cleaning rules for the raw PO history file (KV Custom Windows
and Doors), built from a column-by-column profiling pass. Companion to
`MATERIAL-TOOL-SPEC.md` — that file describes the reporting tool; this one
describes how the raw file gets from "messy" to "the clean file the tool
assumes is sitting there."

**Rule for this log: rules only, never data.** No real row values, no sample
values, no counts derived from the live file get written here. Only the
decision and the reasoning behind it.

Source: `scripts/profile-po-data.ts` — an Office Script that scans the raw
sheet in Excel and reports issue counts per column without ever leaving
Excel or this conversation's chat context.

---

## Columns excluded from cleaning entirely

These are left untouched in the source file and skipped by every check in
the profiler. Not "cleaned," not "checked" — out of scope.

| Column | Reason |
|---|---|
| `ID` | Not used by any report view. |
| `PO DateRevised` | Untouched by every earlier tool too (per spec); not in scope. |
| `SpecialRequest` | Free text, not used by any report view. |
| `QtyThisShip` | Not used by any report view (yet). |
| `Scanned` | Not used by any report view. |
| `ShipLocalle` | Not used by any report view. |
| `Tag` | Not used by any report view. |
| `ShipBy` | Unparseable as a date in 100% of rows — not a usable date column, and out of scope for cleaning. |

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

### INVENTORY ID — leading/trailing whitespace
**Decision: trim.** Uncontroversial — whitespace-only difference, trimming
can't lose information. (Small count relative to file size; applies to the
cleaning script when it's built, not yet implemented.)

## Open — needs a decision

### Duplicate PO No + ITEM NO (24 rows across 24 combos)
Per `MATERIAL-TOOL-SPEC.md`, a PO line can legitimately arrive in partial
shipments, which would make the same PO No + ITEM NO appear more than once
with different `QtyThisShip` / `PO DateReceived`. Not yet determined whether
the flagged rows are partial shipments (expected, not an error) or true
duplicate rows (an error to dedupe).

**Next step:** check the flagged rows directly in Excel. For each duplicate
PO No + ITEM NO pair, compare `QtyThisShip` and `PO DateReceived` across the
occurrences:
- Different `QtyThisShip` / `PO DateReceived` per occurrence → partial
  shipment, expected, not an error.
- Identical across all fields → true duplicate row, needs dedup rule.

### PO DateReceived — unparseable (2,657 rows)
Not blank — something is in these cells that isn't a valid date. Need to
know what (text placeholder like "N/A"/"Pending"? wrong format? stray
characters?) before a parsing/fallback rule can be written.

### PO DateRequired — unparseable (24 rows)
Same open question as above, much smaller count.

### Negative UnitCost (5 rows)
Not yet asked — likely returns/credits, but unconfirmed. Needs a decision:
keep as-is (if returns are real), or flag as a data-entry error.

### Non-numeric QtyReceived (3 rows)
Not yet asked. Needs a decision on what's actually in these 3 cells.

---

## Process

1. Run `scripts/profile-po-data.ts` in Excel (Automate tab) against the raw
   PO history sheet. It never modifies the sheet — only adds a read-only
   "Profile Report" sheet.
2. Go through each issue type together; user confirms the business meaning
   and the resolution rule.
3. Record the resolved rule here (no data, ever).
4. Once every open issue above is resolved, turn this log into an actual
   cleaning script/tool.
