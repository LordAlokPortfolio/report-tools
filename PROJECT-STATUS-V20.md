# Project Status - PO Cleaning & Vendor Analysis (through v20)

Snapshot of what this branch (`alok-idea/cleaning-po-inventory-table-and-vendor-analysis`)
has actually built, as of the vendor-analysis.html "v20" tag. Written so a
later session (or a later you) can pick this up without re-deriving it.

---

## What this branch is

Two connected pieces, cleaning the raw PO/inventory workbook and then
reporting on it:

1. **Cleaning** (`scripts/clean-po-data.ts`) - an Office Script that edits
   the live workbook's `PO-INVENTORY-TABLE` sheet in place: fixes date
   formatting, fills specific NULL patterns using explicit business rules,
   resolves `ID` and `Supplier ID` via lookups against two reference sheets,
   and reorders columns. Every rule is documented in `CLEANING-LOG.md`.
2. **Reporting** (`vendor-analysis.html`) - a standalone web page, not yet
   linked into the main toolkit's `index.html`, implementing six analysis
   views against the cleaned data.

## The workbook (live, in OneDrive)

`tblINVENTORY_PO.xlsx`, found by filename search (not a fixed path) via
Microsoft Graph. Sheets in use:

- `PO-INVENTORY-TABLE` (trailing space in the real name) - the main PO
  history, cleaned by the Office Script.
- `Da Vinci File` - inventory master; source of truth for `ID`, matched by
  `INVENTORY ID`.
- `SUPPLIER FILE` - supplier master; source of truth for `SUPPLIER NAME`,
  matched by `Supplier ID`.
- `BOTTLENECK` (trailing space in the real name) - BOM/component list,
  still being built by the user; not yet populated with the
  `ParentSKU, ComboID, ComponentSKU, QtyPerUnit` columns the Bottleneck view
  expects.
- `HISTORY SHEET` (table `tbl_HISTORY`) - inventory movement ledger:
  `ID, INVENTORY NAME, DATE, QTY, TYPE, BATCH_ID, BATCH_TIME, SORT_DATETIME`.
  `TYPE` semantics (confirmed): `Count` = snapshot/reset of on-hand, `Receivings`
  = adds to on-hand, `Transfers` = always subtracts (usage). QTY is unsigned.

## How vendor-analysis.html connects to the workbook

No file upload in normal use - MSAL.js signs the user into the KV Custom
Windows & Doors Microsoft 365 tenant, then reads the workbook live via
Microsoft Graph (`Files.Read` scope, not `Files.Read.Selected` - the
originally-planned OneDrive file picker turned out to depend on a picker
SDK version that no longer exists publicly, so the simpler path-independent
approach is: find the file by filename search, cache its item ID, then read
each sheet's `usedRange` by worksheet ID - not by name string, since sheet
names on this workbook carry inconsistent trailing whitespace). A manual
file-upload fallback panel still exists in case the OneDrive path breaks.

Nothing is cached beyond the browser session except the workbook's item ID
(in `localStorage`, not sensitive). Every "Refresh" re-reads live.

## The six views - current state

| View | Status | Notes |
|---|---|---|
| **Speed** | Working | Headline is the median of the vendor's last 3 closed orders ("current pace"), not a multi-year blended median - a vendor's lead time can genuinely shift (e.g. 10 days -> 35 days) and a long-window median hides that. Full-timeline median shown as context, with an explicit callout when the two diverge. Per-PO table, newest first. |
| **Creep** | **Broken as a pricing signal - do not trust `UnitCost` from this data.** See "Pricing is unreliable" below. |
| **Pattern** | Reworked away from raw correlation into an order-to-order "did cost move the same direction as lead time" tendency - but its cost side inherits the same pricing problem as Creep. The lead-time-only parts of the underlying logic are sound; the cost-linked half is not. |
| **Habits** | Buckets orders by quarter x order-size, reporting median lead time (trustworthy) and median cost (not trustworthy) per bucket. |
| **Bottleneck** | Needs the `BOTTLENECK` sheet populated with real BOM columns - not yet usable. Its lead-time rollup logic is sound; its cost rollup is not. |
| **Runout** | Working, price-independent. On-hand and usage rate computed from `tbl_HISTORY` (Count/Receivings/Transfers), lead time from the PO table's closed orders, matched via `ID`. |

## Pricing is unreliable - the reason this status doc exists

`UnitCost` in the PO history is contaminated: purchasers send custom POs
with **$0 as a placeholder price**, not a real cost. A partial fix was built
(`isStockItem()` - a row's `ID` is numeric for a real stock item vs. free
text like "CUSTOM ORDER OF: ..." for a custom order, so custom-order rows
could be excluded from cost math) but a live diagnostic (v20) showed
several ITEM NOs still producing identical, implausible cost jumps even
after that filter - meaning either the filter's assumption about which
column carries the custom-order marker is wrong, or the pricing data has
deeper problems than one filter can fix.

**Decision (this conversation): stop trying to fix `UnitCost`-based
analysis. Treat every dollar figure in this dataset as untrustworthy until
proven otherwise, and design the reporting tool around what the data
*can* prove without price.** Creep in particular is fundamentally a
price-based view and should be considered retired unless the pricing data
gets a real fix at the source (not a filter in this tool).

## Columns confirmed trustworthy (no known contamination)

`PO No`, `Supplier ID`/`SUPPLIER NAME`, `PO Date`, `PO DateReceived`,
`PO DateRequired`, `ITEM NO`, `INVENTORY ID`, `ID` (post Da Vinci lookup),
`Quantity`, `QtyReceived`, `POLineClosed`, `Category` (per earlier decision,
even the literal value `"0"` is a real code, not blank). `tbl_HISTORY`'s
`DATE`, `QTY`, `TYPE` for stock items.

## Not yet done

- Bottleneck needs real BOM data.
- The `ID`-column diagnostic in Creep (v20) was never resolved - abandoned
  in favor of dropping price-based analysis entirely per this session's
  decision, rather than continuing to debug it.
- Nothing on this branch is linked into `index.html` / the main toolkit yet
  - explicit standing instruction is to keep `main` untouched until the
  whole project is ready.
- No holiday calendar for business-day math (documented limitation,
  `CLEANING-LOG.md`).

## Version history (vendor-analysis.html)

v1-v12 (not individually tagged): initial 6-view build, OneDrive/Graph
integration built and debugged from scratch (broken CDN library versions,
a picker SDK that turned out not to exist publicly, sheet names with
invisible trailing whitespace, path-based Graph lookups replaced by
filename search and then by worksheet-ID matching).

- **v13** - Speed/Creep reworked around recency (median of last 3 orders,
  not a multi-year blend).
- **v14** - Pattern reworked from a raw correlation number into an
  order-to-order "did cost follow lead time" tendency with a plain-English
  "so what."
- **v15** - Speed/Creep simplified to single shareable sentences instead of
  multi-paragraph stat dumps.
- **v16** - Speed's per-PO table sorted newest-first.
- **v17** - Pattern collapses items with no tendency into a count instead of
  repeating the same sentence for every item.
- **v18** - Fixed a real bug: Creep's percent-change math divided by zero
  when `First cost` was $0, silently showing "-" even when an item clearly
  jumped from $0 to a real price.
- **v19** - Attempted fix: exclude custom-order placeholder pricing
  (`isStockItem()`) from Creep, Pattern, Habits, and Bottleneck's cost math.
  Also: all em/en dashes replaced with plain hyphens across every tracked
  file on this branch, per explicit request.
- **v20** - Added a live diagnostic to Creep to see what the `ID` column
  actually contains, after the v19 fix didn't hold up against real data.
  Superseded by the decision to drop price-based analysis (this document).
