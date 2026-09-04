# Project Status - PO Cleaning & Vendor Analysis (through v22)

Snapshot of what this branch (`alok-idea/cleaning-po-inventory-table-and-vendor-analysis`)
has actually built, as of the vendor-analysis.html "v22" tag. Written so a
later session (or a later you) can pick this up without re-deriving it.
Supersedes the file's original "through v20" version - kept as one file,
not a new one per version, so it doesn't fork into stale copies.

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
   linked into the main toolkit's `index.html`, implementing seven analysis
   views against the cleaned data - all of them price-free (see "Pricing is
   unreliable" below for why).

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
  Real values are `Count`, `RECEIVING`, `TRANSFER`, `REJECTION` (singular, all
  caps in practice - matched case-insensitively, singular or plural). `TYPE`
  semantics (confirmed): `Count` = snapshot/reset of on-hand. `Receiving` =
  adds to on-hand. `Transfer` = subtracts from on-hand AND counts as usage
  demand for Runout's usage-rate calc. `Rejection` = subtracts from on-hand
  but does NOT count as usage demand (defective stock leaving inventory isn't
  the same signal as real consumption). QTY is unsigned.

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

## The seven views - current state

Creep and Pattern are **retired**, not just fixed - their core premise was
price itself ("is price rising," "does cost move with lead time"), so once
`UnitCost` is untrustworthy there's no non-cost version of either question
to fall back to. Replaced with three new views built only from data nobody
has flagged as unreliable.

| View | Status | Notes |
|---|---|---|
| **Speed** | Working | Headline is the median of the vendor's last 3 closed orders ("current pace"), not a multi-year blended median - a vendor's lead time can genuinely shift (e.g. 10 days -> 35 days) and a long-window median hides that. Full-timeline median shown as context, with an explicit callout when the two diverge. Per-PO table, newest first. |
| **Short-shipping** (new, replaces Creep's slot) | Working | `Quantity` ordered vs `QtyReceived` actually received, per item, for the selected vendor/timeline. Which items has this vendor delivered less of than was ordered. |
| **Reorder Cadence** (new, replaces Pattern's slot) | Working | Gap between consecutive `PO Date`s for the same item. Flags items now being ordered noticeably more often than their own history - says the rhythm changed, not why. |
| **Vendor Concentration** (new) | Working | Whole-building view (ignores vendor/timeline picker, like Runout). Which vendors are the sole source for at least one item, based on every stock item's supplier history. |
| **Habits** | Working, cost column dropped | Buckets orders by quarter x order-size, reporting median lead time only. |
| **Bottleneck** | Needs the `BOTTLENECK` sheet populated with real BOM columns - not yet usable. Cost-to-build column dropped; reports lead-time-only bottleneck. |
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

**Decision: stop trying to fix `UnitCost`-based analysis entirely.** Every
dollar figure in this dataset is untrustworthy until proven otherwise. As
of v22, no view in this tool depends on price at all. `isStockItem()`
survived, repurposed: it now excludes custom-order rows (which aren't
comparable stock items by definition) from Vendor Concentration, not from
cost math.

**Also confirmed unreliable, separately: `PO DateRequired`** (purchaser-
entered, not trustworthy). Used in exactly one narrow, intentional
exception (`clean-po-data.ts`'s closed-order `PO DateReceived` NULL-fill,
the least-bad option when no real date exists at all) - nowhere else in
cleaning or reporting.

## Columns confirmed trustworthy (no known contamination)

`PO No`, `Supplier ID`/`SUPPLIER NAME`, `PO Date`, `PO DateReceived`,
`ITEM NO`, `INVENTORY ID`, `ID` (post Da Vinci lookup), `Quantity`,
`QtyReceived`, `POLineClosed`, `Category` (per earlier decision, even the
literal value `"0"` is a real code, not blank). `tbl_HISTORY`'s `DATE`,
`QTY`, `TYPE` for stock items.

**Not on this list, deliberately: `UnitCost`, `PO DateRequired`.** Both
confirmed unreliable - see above.

**Not yet checked either way**: every other column not named here.
Absence from this list isn't a claim of trustworthiness - it's what hasn't
been verified or contradicted yet. Both `UnitCost` and `PO DateRequired`
were assumed fine until someone with direct knowledge said otherwise; that
same caution should apply to everything not yet stress-tested.

## Not yet done

- Bottleneck needs real BOM data.
- Nothing on this branch is linked into `index.html` / the main toolkit yet
  - explicit standing instruction is to keep `main` untouched until the
  whole project is ready.
- Company-specific holiday closures beyond the standard 9 Ontario
  statutory holidays aren't in the business-day math (`vendor-analysis.html`,
  `ontarioHolidaysForYear()`).

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
- **v21** - Stopped using `PO DateRequired` (confirmed unreliable by the
  purchaser), then restored it as one narrow, intentional exception in
  `clean-po-data.ts`'s closed-order NULL-fill rule once scope was clarified.
- **v22** - Retired Creep and Pattern entirely (their core premise was
  price, not just decorated with it - no non-cost version exists). Added
  Short-shipping, Reorder Cadence, and Vendor Concentration. Dropped the
  cost columns from Habits and Bottleneck. As of this version, no view in
  the tool depends on price at all.
