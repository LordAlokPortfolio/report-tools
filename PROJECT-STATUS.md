# Project Status - PO Cleaning & Vendor Analysis (through v43)

Snapshot of what this branch (`alok-idea/cleaning-po-inventory-table-and-vendor-analysis`)
has actually built, as of the vendor-analysis.html "v43" tag. Written so a
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
   linked into the main toolkit's `index.html`, implementing nine analysis
   tabs plus an About tab against the cleaned data - all of them price-free
   (see "Pricing is unreliable" below for why). The tab bar is split into
   two visually divided groups: vendor-scoped (Speed, Shortfall, Frequency,
   Backbone, Sensitivity, Actions) and whole-inventory (Bottleneck, Concentration,
   Runout), followed by About. Bottleneck moved to the whole-inventory group
   in v44 - its computation was already whole-building (ignores the vendor
   picker), it was just visually misplaced.

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

## The nine tabs (plus About) - current state

Creep and Pattern are **retired**, not just fixed - their core premise was
price itself ("is price rising," "does cost move with lead time"), so once
`UnitCost` is untrustworthy there's no non-cost version of either question
to fall back to. Replaced with new views built only from data nobody has
flagged as unreliable.

| Tab | Status | Notes |
|---|---|---|
| **Speed** | Working | Headline is the median of the vendor's last 3 closed orders ("current pace"), not a multi-year blended median - a vendor's lead time can genuinely shift (e.g. 10 days -> 35 days) and a long-window median hides that. Full-timeline median shown as context, with an explicit callout when the two diverge. Per-PO table, newest first. |
| **Shortfall** (replaces Creep's slot) | Working | `Quantity` ordered vs `QtyReceived` actually received, per item, for the selected vendor/timeline. Which items has this vendor delivered less of than was ordered. |
| **Frequency** (replaces Pattern's slot, was "Reorder Cadence") | Working | Gap between consecutive `PO Date`s for the same item, restricted to stock items (`isStockItem()`). Flags items now being ordered noticeably more often than their own history - says the rhythm changed, not why. Sorted by total order count descending (most-ordered item first), not by recency. |
| **Backbone** (renamed from Mix in v48-v49, briefly "Staple" in between) | Working | Per selected vendor: score is units received per week (a rate, not a total), computed from each item's own first-to-last order span, gated on `MIN_HISTORY` (6) orders spread over more than one day. Replaces the v46/v47 share-based MVP score, which couldn't tell a real recurring need apart from a single bulk order - a screw ordered constantly scored the same as a one-time million-unit purchase, since both are just "big totals." |
| **Sensitivity** (renamed from "Habits" in v40) | Working, cost column dropped | Buckets orders by quarter x order-size, reporting median lead time only. Heading (v39): "does a bigger order, or a different time of year, get delivered faster or slower?" |
| **Bottleneck** | Needs the `BOTTLENECK` sheet populated with real BOM columns - not yet usable. Cost-to-build column dropped; reports lead-time-only bottleneck. |
| **Actions** (new, v43) | Working | Vendor-scoped: for the selected vendor, one card per SKU with a computable runout date, sorted soonest-due first. Reuses `computeRunoutRows()`, the same computation Runout uses, filtered to `r.vendor === selected vendor`. Relies on the confirmed fact that no two suppliers share an `ID` - no separate relationship-key lookup needed. |
| **Concentration** (was "Vendor Concentration") | Working | Whole-building view (ignores vendor/timeline picker, like Runout). Columns: Vendor, Codes ordered, Total codes by vendor, % of assigned catalog ordered (Da Vinci master list), Most-ordered code. Custom-order rows excluded via `isStockItem()`. |
| **Runout** | Working, price-independent. On-hand and usage rate computed from `tbl_HISTORY` (`Count`/`RECEIVING`/`TRANSFER`/`REJECTION`, singular and plural both matched), lead time from the PO table's closed orders, matched via `ID`. `REJECTION` subtracts on-hand but is excluded from usage-rate demand. |
| **About** (new) | Working | In-page tab describing the tool; restructured (v39) to lead directly with what each tab does, instruction-manual style, instead of opening with narrative. |

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
- **v23-v37** (not individually logged here) - Added the Mix tab (MVP by
  order-count share, corrected from an initial quantity/volume-share
  implementation). Renamed "Reorder Cadence" to "Frequency" and "Vendor
  Concentration" to "Concentration." Fixed Runout's TYPE matching to accept
  singular `RECEIVING`/`TRANSFER`/`REJECTION` (real ledger values), not just
  plural. Confirmed `REJECTION` subtracts on-hand but is excluded from
  usage-rate demand. Fixed Short-shipping and Frequency to group by `ID`
  (canonical resolved identifier) instead of `ITEM NO` (vendor-specific,
  unreliable), with `isStockItem()` filtering added to both. Reworked
  Concentration's columns: replaced "Only they supply" (which coincidentally
  matched "Codes ordered") with "Total codes by vendor," reordered columns,
  converted "Codes assigned to them" to a percentage instead of a raw
  fraction. Added the About tab. Rewrote
  `CLEANING-PO-INVENTORY-TABLE-AND-VENDOR-ANALYSIS-SPEC.md` from stale
  pre-build planning notes into continuous human prose. Restored
  `PO DateRequired` as the one narrow exception in the closed-order
  `PO DateReceived` NULL-fill rule after briefly dropping it entirely.
  Security incident: a passphrase value was leaked into a commit message,
  rotated to a new passphrase, offending commit reworded via interactive
  rebase and force-push.
- **v38** - New About tab content; spec rewritten in plain human prose.
- **v39** - Tab bar split into two visually divided groups (Speed,
  Shortfall, Frequency, Mix, Habits, Bottleneck | Concentration, Runout),
  with About after. Frequency's sort fixed to rank by total order count
  descending instead of recency. Habits' heading rewritten into a plain
  English sentence. About tab and the spec doc both restructured to open
  directly with what each tab does, instruction-manual style, instead of
  starting with unprompted narrative.
- **v40** - "Habits" tab renamed to "Sensitivity" - the old name didn't
  match what the tab measures (does order size/season change vendor
  response), which reads more like a vendor-behavior test than a record of
  the user's own habits.
- **v41** - `itemLabel()` (the "ID - Name" resolver used by Mix, Frequency,
  and Shortfall) now prefers a "Name" column the user added directly to
  `PO-INVENTORY-TABLE` (column U as of this writing - a user-maintained,
  more-readable name per `ID`), falling back to Da Vinci File / ledger
  names only for an `ID` that column doesn't cover. This is a display-only
  change - no cleaning-script logic or column position assumption changed,
  since the tool reads by header name (`"Name"`), not by column letter.
- **v42** - Added a stockout-count KPI to Runout: walks the same ledger
  movements Runout already reads and counts every time an ID's running
  on-hand crosses from above zero to zero-or-below, in date order. Reported
  as a total plus a per-ID breakdown table, sorted worst-first. This is a
  historical count, not a forecast, and doesn't invent a dollar figure -
  it's the one number that directly answers "did we run out," meant to be
  tracked over time as the "never run out" mandate's actual scorecard.
  Confirmed with the user: every `ID` is single-sourced (no two suppliers
  share an `ID`), so this count needs no per-vendor disambiguation.
- **v43** - Added an Actions tab: vendor-scoped runout cards. Refactored
  Runout's per-ID computation out into a shared `computeRunoutRows()`
  function (on-hand, weekly usage, weeks left, lead time, order-by,
  stockout count) so Runout (whole-building table) and Actions (cards
  filtered to the selected vendor, sorted soonest-due first) read the
  exact same numbers instead of two independent calculations.
- **v44** - Fixed a real bug: `vendorBySku` was built only from closed PO
  rows, so any SKU without a closed order yet (still open, or never
  reordered) got no vendor tag at all and could never appear in Actions
  for any vendor - even though it was sitting right there in the ledger
  and visible in the whole-building Runout table. Vendor assignment is now
  built from every PO row, open or closed; lead-time calculation stays
  closed-only (a real delivery is still required to measure lead time).
  Also moved Bottleneck from the vendor-scoped group to the whole-inventory
  group - its own computation already reads across all vendors/history and
  ignores the vendor picker, so it was categorized wrong, not just placed
  wrong.
- **v45** - About tab: each tab's explanation cut to one sentence, grouped
  under Category 1 (vendor-scoped) / Category 2 (whole-inventory), matching
  the tab bar's own grouping. Added the missing Actions entry.
- **v46** - Fixed a real bug in Mix: MVP was ranked purely by order-count
  share, so a code ordered 19 times for a total of 33 units outranked one
  ordered 8 times for 800 units. MVP score is now a 50/50 blend of
  order-count share and received-quantity share (`QtyReceived`, not the
  ordered `Quantity`, since that's what actually arrived). Table now shows
  both shares plus the blended score, not just one number pretending to be
  the whole picture.
- **v47** - Trimmed Mix's table to Item / Orders / Qty received / MVP score
  (dropped the two share-percentage columns), sorted highest score first,
  and added a small (i) tooltip on the MVP score header showing the
  formula in plain terms.
- **v48** - Reworked Mix into Staple (placeholder name - not yet finalized):
  the share-based MVP score (order-count share + qty-received share) could
  not tell a real recurring need apart from a single bulk order - a screw
  ordered constantly and a one-time million-unit purchase of anything else
  scored the same, since both are just "big totals." Score is now units
  received per week, computed from each item's own first-to-last order
  span (the same rate concept Runout already uses), gated on `MIN_HISTORY`
  (6) orders spread over more than one day - an item with less history
  than that has no repeated pattern to measure a rate from, and none is
  guessed. Table columns unchanged in shape (Item / Orders / Qty received /
  rate), tooltip updated to the new formula.
- **v49** - Tab named "Backbone" (was Mix, briefly "Staple" as a
  placeholder) - final name for the rate-based recurring-demand view
  introduced in v48.
