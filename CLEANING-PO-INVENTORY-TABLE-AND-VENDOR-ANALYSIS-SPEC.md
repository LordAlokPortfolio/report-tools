# Vendor & Material Analysis — what this tool actually is

KV Custom Windows and Doors, Brampton ON. This document explains the
tool as it exists today — what each tab does first, then how it works
underneath.

## What each tab does

**Speed** — how long a vendor actually takes to deliver right now, not
what they quoted once. Weighted toward their most recent orders,
because a vendor's pace can genuinely shift and blending that shift
into years of older history would hide it instead of catching it.

**Shortfall** — what was ordered vs. what actually arrived, per item,
for the vendor selected. Catches a vendor who quietly ships less than
the PO said, order after order — invisible unless someone is counting
boxes against the PO.

**Frequency** — which items are being reordered more often than their
own history suggests, ranked by how many times each has been ordered.
Doesn't say why the rhythm changed; just flags it.

**Staple** (placeholder name, still being decided) — for the vendor
selected, which item has real recurring demand: units received per
week, measured across that item's own order history, not a total. A
total can't tell a genuine ongoing need apart from a single bulk
purchase — a million units bought once and a million units bought
steadily over two years look identical if you only sum them. Requires
at least a handful of orders spread over more than one day before it
will compute a rate at all; otherwise there's no repeated pattern to
measure a rate from, and none is guessed.

**Sensitivity** — does a bigger order, or a different time of year, get
delivered faster or slower by this vendor. About this vendor's
behavior toward you specifically, not a comparison between vendors.

**Bottleneck** — for anything built from multiple purchased parts: a
build only moves as fast as its slowest component. Pulls the real lead
time for each part in a bill of materials and reports the date that
actually governs the build.

**Actions** — for the vendor selected, one card per SKU that vendor
supplies with enough history to compute a runout date, sorted soonest-
due first. Same on-hand, usage, and lead-time figures Runout computes
for the whole building, just filtered down to this vendor's own SKUs
and shown as something to act on today rather than a table to scan.

**Concentration** — ignores the vendor selected; a whole-building
view. Which vendors would hurt the most if they disappeared: how many
inventory codes each one is the only source for, and what share of
everything assigned to them has actually been ordered.

**Runout** — also ignores the vendor selected, since running out of
something is a whole-building question, not a per-vendor one. Reads
the inventory ledger directly — on hand, usage rate, lead time — and
says plainly whether ordering today is soon enough.

## What problem this solves

Every purchasing decision at KV depends on knowing things that live
scattered across a messy spreadsheet: how long a vendor actually takes
to deliver, whether they're shipping less than what was ordered,
whether the business depends too heavily on one supplier for certain
parts, and what's about to run out on the shop floor. Nobody wants to
open a 40,000-row workbook and figure that out by hand every time a
question comes up. This tool reads that workbook directly, does the
arithmetic, and shows the answer as one sentence a person can act on.

The one rule that shaped every decision in building this: the tool
only says what the data can prove. Early drafts of this idea kept
drifting toward inventing things — a dollar figure for what a delay
cost, a recommendation to drop a vendor, an explanation for why a
price moved. None of that lives in the data. It would have been made
up, and a tool that quietly guesses is worse than a tool that says "I
don't have enough here to tell you." So every number this tool prints
traces back to an actual column in the workbook, or to arithmetic on
those columns, or to something the user typed in themselves. If there
isn't enough history to say something meaningful, the tool says that
plainly instead of printing a confident-looking number built from
three data points.

## How it connects to your data

The tool lives at a public web address, but the data it reads is
private. You sign in with your Microsoft account, and from then on the
tool reads your PO workbook live from OneDrive — nothing is uploaded,
nothing is copied anywhere else, and nobody who isn't signed in as you
can see it. Because this page is public, anyone who opens the link
without knowing the passphrase sees the tool running on a made-up
sample dataset instead — fake vendors, fake numbers, but every tab
fully working, so the tool can be shown off publicly without exposing
anything real. The passphrase itself isn't real security — it's
readable in the page's own source code by anyone who looks for it.
It's a doorbell, not a lock: enough to keep a casual visitor from
stumbling into the live-data mode, not something that would stop
anyone determined. The actual protection is that your real data never
gets published anywhere; it only ever gets read live, directly from
your own OneDrive, while you're signed in.

## Why the data needed cleaning first

The raw workbook, as purchasers actually use it day to day, has real
problems. Custom orders get entered with a placeholder cost of zero
dollars because nobody's typed in a real price yet — and until that
was noticed, the tool was quietly treating those zeroes as real
prices, which made every cost-based number wrong. The column meant to
hold a vendor's promised delivery date turned out to be unreliable —
purchasers don't fill it in carefully, so it can't be trusted for
anything except one narrow fallback. Dates come in with time-of-day
attached where only the date matters. Some rows are missing a receipt
date entirely. None of this is a flaw in the tool; it's the normal
mess of a spreadsheet that real people type into under time pressure.

A separate script (`scripts/clean-po-data.ts`, run from Excel's own
Automate tab) fixes what can be fixed without guessing, and leaves
alone what can't be trusted. It strips time-of-day from dates, fills a
handful of specific gaps using rules that were worked out and
confirmed one at a time — never invented — and looks up the proper
inventory ID and vendor name from the workbook's own reference sheets,
since the raw codes on their own aren't readable to a person. Every
one of those rules, and the reasoning behind it, is written down in
`CLEANING-LOG.md` so the next person working on this doesn't have to
re-derive the same decisions from scratch.

One thing worth being direct about, because it changed the shape of
this tool significantly: pricing data turned out to be unreliable
enough, in ways that kept surfacing new problems even after being
patched, that every dollar-based feature was eventually dropped rather
than kept as something half-trustworthy. What's left only uses columns
that have actually held up under scrutiny — dates, quantities, order
counts, and inventory movement.

## What this tool deliberately does not do

It never prints a dollar figure unless a human supplies the cost
assumption themselves, and even then the arithmetic is shown openly,
not hidden behind a single number. It never tells you to switch
vendors or renegotiate a contract — it shows the pattern, and the
decision stays a human one. It never claims one thing caused another;
"cost and lead time moved together" is a fact the data can support,
"cost rose because lead time rose" is not, and the tool won't blur
that line even when the second sentence sounds more useful.

It also can't do a few things a bigger system could: it can't send
email on its own, it can't create an actual purchase order in your
purchasing system, and it doesn't share data between different
people's browsers — each session only ever holds what it just read
from OneDrive.

## Where this still has open edges

Not every part of this is finished. The Bottleneck view needs a real
bill-of-materials sheet to be useful, and that sheet is still being
built. The business-day math accounts for the standard Ontario
statutory holidays, but not for any company-specific closures beyond
those. And the honest caveat that applies to the whole tool: it's only
as good as the workbook it reads. When a new kind of bad data turns
up — and pricing already proved that it can — the right response is
the same one that's been used every time so far: stop, ask what's
actually true, and only build the next piece on ground that's been
checked.
