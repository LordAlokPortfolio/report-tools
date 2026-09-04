# Material & Vendor Tool - design notes

Working spec for a new tool to be added to the Operations Toolkit
(`index.html`). Nothing here is built yet. This file exists so a new
Claude session can pick up the thinking without starting over.

Company: KV Custom Windows and Doors, Brampton ON.
Repo: `LordAlokPortfolio/report-tools`, public, served by GitHub Pages.

---

## The one rule

**The tool only says what the data can prove.**

This came out of a long back-and-forth and it is the most important line
in this file. Every earlier draft of this idea drifted into inventing
things - a dollar figure for what a delay "cost," a recommendation to
drop a vendor, an explanation for why a price moved. None of that is in
the data. All of it was made up.

So:

- A number the tool prints must trace to a column in a file, or to
  arithmetic on those columns, or to a setting the user typed in.
- No dollar impact unless the user supplies the cost assumption
  themselves, and then the math is shown openly.
- No "switch vendors" / "renegotiate" verdicts. The tool shows the
  pattern; the human decides what to do about it.
- No causal claims. "Cost and lead time moved together" is allowed.
  "Cost rose because lead time rose" is not.
- Where there isn't enough history to be meaningful, say "not enough
  history yet" instead of printing a confident-looking number off three
  data points.

If a future session finds itself writing a sentence the CSV can't back,
that sentence is wrong, no matter how good it sounds.

---

## Data it reads

### 1. Purchase order history (main file)

One CSV, roughly 8.6 MB, January 2025 to present. Real columns:

```
ID, PO No, Supplier ID, PO Date, ITEM NO, Quantity, INVENTORY ID,
PO DateReceived, PO DateRequired, PO DateRevised, QtyReceived, ShipBy,
SpecialRequest, QtyThisShip, POLineClosed, Scanned, Tag, ShipLocalle,
UnitCost, Category
```

Notes on the columns that matter:

- `PO Date` → `PO DateReceived` is the real lead time. This is the
  ground truth the whole tool is built on.
- `PO DateRequired` is what was asked for / promised. Comparing it to
  `PO DateReceived` tells you how a vendor does against their own word,
  which is a different question from how long they take.
- `PO DateRevised` is a vendor moving their own goalposts mid-order.
  Nothing built on this yet - it's untouched by every earlier tool too.
- `Quantity` vs `QtyReceived` vs `QtyThisShip` - a PO line can arrive in
  pieces. Untouched so far.
- `ITEM NO` looks like each vendor's own part number, **not** a shared
  code across vendors. Two vendors selling the same physical thing will
  not match on it.
- `Supplier ID` is a code, not a name. No name column exists in this
  file. Earlier tools solved this by splitting one Excel file per vendor
  upstream, so identity came from the filename. If vendor names are
  wanted here, a small `Supplier ID → name` lookup file has to be built.
- `Tag` was used in an earlier project to split Cardinal's orders into
  Stock (`Tag` == "-") vs Custom (anything else).

Cleaning this file is a separate side project, handled outside this
tool. The tool assumes a clean file is sitting there when it runs.

### 2. Inventory ledger (for idea 6)

Stock movement per SKU, March to present, ledger format - rows of
movement, not snapshots. Gives on-hand and usage rate.

### 3. Component list / BOM (for idea 5 only)

A second small CSV the user maintains by hand, one row per component per
combination:

| ParentSKU | ComboID | ComponentSKU | QtyPerUnit |
|---|---|---|---|
| SKU1 | A | sku-a | 1 |
| SKU1 | A | sku-b | 2 |
| SKU1 | B | sku-a | 1 |
| SKU1 | B | sku-x | 2 |

Same parent SKU can have several combos where one component swaps out.
`ComponentSKU` joins to `ITEM NO` / `INVENTORY ID` in the PO file.
`QtyPerUnit` is needed for cost rollup - add it when the list is built
out, or the per-unit cost can't be exact.

---

## Shared rules for all the date math

- Lead time is counted in **business days** - weekends and the Ontario
  holiday list excluded. Both earlier projects (the JS forecasting tool
  and the PowerShell/Python vendor scripts) already do it this way.
- When converting a lead time into a suggested date, **round up**. Never
  hand someone a date sooner than the vendor has actually proven. Label
  it as rounded up so nobody thinks it's the raw median.
- Use the **median**, not the average, for typical lead time. One
  disaster order shouldn't drag the number.
- Minimum history before showing a number: roughly 6-8 orders. Below
  that, say so instead of guessing.
- Every screen shows the date of the data file it read. The ledger and
  the PO file both keep moving; nobody should act on stale numbers
  without knowing they're stale.

---

## The six views

Names are deliberately one plain word each.

### 1. Speed

How fast a vendor actually is, and when to place the next order.

Needs: `PO Date`, `PO DateReceived`.
Median business-day lead time, rounded up, added to today.

> "Vendor Y has actually taken 7 weeks lately, not the 5 they quote.
> Order today, and place the next one by March 12 if you don't want a
> gap."

### 2. Creep

Price rising quietly, a few cents at a time, across many orders.

Needs: `ITEM NO`, `UnitCost`, `PO Date`.
First vs latest unit cost, percent change, biggest single jump.

> "Item #4471 from Vendor Y: 11 orders since January, cost up from $4.20
> to $5.35. Twenty-seven percent, no single jump bigger than fifteen
> cents."

### 3. Pattern

Whether cost and lead time move together for an item. Correlation only -
not a claim about cause.

Needs: `ITEM NO`, `UnitCost`, `PO Date`, `PO DateReceived`.
Correlation between the per-order cost series and the per-order lead
time series, both in PO Date order. Needs 6-8+ orders or it's noise.

> "Same item, same vendor, four orders this year: price and lead time
> went up together every time. Correlation 0.81 - strong."

### 4. Habits

Whether *when* and *how big* you buy changes what you get - from the
vendor you already have. Not about switching vendors.

Needs: `PO Date` (→ season), `Quantity` (→ size bucket), `UnitCost`,
lead time. Group this vendor's orders by season × size, compare.

> "Vendor Y in Q3: 5 weeks. Same vendor, same item, Q1: 8 weeks. Big
> orders ran $4.60 a unit, small ones $5.10."

Caveat: history starts January 2025, so some season/size buckets will be
too thin to mean anything. Say so rather than averaging two orders.

### 5. Bottleneck

A build moves at the speed of its slowest part. Everything else is just
waiting.

Needs: the BOM file + lead time and cost per component from the PO file.
Per parent SKU + combo: pull each component's lead time, take the
longest - that's the earliest the build can start. Sum component costs
(× QtyPerUnit) for a cost-to-build.

> "SKU1 needs three parts. Two arrive in under two weeks. The third -
> sku-b, Vendor 2 - takes five. That's the date that matters. Parts
> cost: $47.20 a unit."

SAP calls this the material availability date. It doesn't care which
suppliers were fast; it only reports the one that decides anything.

### 6. Runout

What's going to run dry, when, and what to do about it today.

This one is different from the other five: it's **not** filtered by
vendor or timeline. "Am I about to run out" is a today question about
the whole building. It lives behind its own tab.

Needs: inventory ledger (on hand + usage rate) + `PO Date` /
`PO DateReceived` (lead time) + open POs (`POLineClosed` = false, with
`Quantity` and expected date).

The BOM is **not** used here. The ledger movement carries the signal on
its own.

> "You have 40 units of sku-b on hand, using 6 a week. At that rate you
> run out in 6.7 weeks. Vendor 2 takes 5 weeks to deliver it. Order now
> and you're fine - order two weeks from now and you'll run dry before
> the truck arrives."

**Output: an action list, grouped by vendor.** Nine columns, nothing
spare - each one either states the problem or states the move:

| SKU | Vendor | On hand | Weekly usage | Weeks left | Open PO (qty + date) | Real lead time | Order by | Suggested qty |

Splitting into multiple POs: the tool can group by vendor and break a
large order into several under a ceiling - but the ceiling is a number
the user types in. The tool has no idea what the approval limits are and
must not invent one.

---

## What the app looks like

New page in `index.html` alongside the NCR generator, paint reader,
material transfer form and tariff dashboard.

1. **Landing** - pick a vendor, pick a timeline.
2. **Views 1-5** - the vendor/timeline choice drives all of them.
3. **Runout tab** - separate, ignores the vendor/timeline choice
   entirely. Click it and get the whole-building picture.

Every screen shows "inventory as of [date]" and "purchase orders as of
[date]" from the files it read.

---

## Public site, private data

Decision: the **tool** is public, the **data** is not.

- Sample CSVs ship in the repo so anyone opening the public link sees a
  working tool with fake numbers.
- The real PO history and inventory ledger are gitignored and never
  pushed. Same pattern already used for `data/company-email.json`.
- Someone in the KV building points the tool at the real file on their
  own machine.

**Be honest about the passphrase.** On a public static site, a
passphrase in the page is not security - anyone can read the page source
and find it, and anything actually published at a public URL can be
fetched directly. It is a doorbell, not a lock: fine for "don't wander
in here by accident," useless against anyone who looks. The real
protection is that the sensitive files are never published at all. A
future session should not describe it as more than that, and should not
be talked into putting real cost or vendor data into the repo to make
the passphrase "work."

---

## Things a static page genuinely cannot do

Worth writing down because these came up and the honest answer matters
more than the enthusiastic one.

- **Send email.** It can open the user's mail app with a draft already
  filled in, or hand over text to copy. Long lists get truncated in a
  mail draft, so a copy/download path is safer.
- **Create a purchase order in the purchasing system.** It can produce a
  PO-shaped document per vendor (PDF or Excel), ready to send or key in.
  That's what the existing NCR and transfer tools already do. Generating
  a document is not the same as creating a PO, and shouldn't be
  described as if it were.
- **Share data between people.** Every existing tool in this repo stores
  per-browser only. There is no backend and no database.

---

## Not decided yet

- Name of the tool.
- Whether to build the `Supplier ID → vendor name` lookup file.
- `QtyPerUnit` in the BOM - needed for exact cost rollup, may be filled
  in later once more combos are logged.
- Whether the cross-vendor comparison (same physical item, two vendors)
  is worth the manual crosswalk file it would require. Parked.
- Nothing built. The CSV cleaning is still in progress.

---

## Earlier versions of this idea

Three tools already exist outside this repo. Worth reading before
rebuilding anything, mostly so the same logic isn't reinvented
differently.

- **`purchasing.py`** - ETL. Walks a folder of per-vendor Excel exports,
  forces them into a fixed schema, rebuilds a SQLite DB per run.
- **`vendor_analysis.ps1`** - PowerShell wrapper that writes out two
  Python scripts, installs deps, scans an input folder, prompts for each
  vendor's standard lead time, and produces Excel + PDF + charts + a
  PowerPoint per vendor plus a master scorecard. Business-day math with
  an Ontario holiday list. Splits Cardinal by `Tag` into Stock/Custom.
- **`vendor_analysis_gui.py`** - the most finished one. PyQt6 desktop
  app over three SQLite DBs, hardcoded vendor list, per-vendor and
  cross-vendor prompts, clickable matplotlib charts, exports one
  annotated Excel workbook with plain-English explanations of each flag.
- **`main.js`** (Universal Forecasting Tool, v1) - different problem:
  reorder timing, not vendor scoring. Browser-only, reads two CSVs,
  infers consumption by diffing cycle-count columns, does adaptive ABC
  classification, and produces PLACE ORDER / REVIEW / NO ACTION per SKU
  against lead-time demand. Has a locked 24-month horizon and the same
  Ontario holiday list. Contains an unfinished "door equivalent" model
  that nothing calls yet.

The user's own summary: the logic in these is outdated, and the point is
to do the same thing a different way, inside the toolkit, reading a
wired-in CSV rather than uploads.
