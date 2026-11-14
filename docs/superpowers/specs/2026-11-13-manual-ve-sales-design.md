# Manual VE Sales Entry + Sale Edit/Delete — Design

**Date:** 2026-11-13
**Status:** Approved in brainstorming (visual mockup v3), pending written-design review
**Mockup:** `.superpowers/brainstorm/1855-1794637990/content/cohesive-v3.html`

## Why

The company now sells at tradeshows through a second point-of-sale app as well as the usual one. Those sales never reach the Excel sheets that the VE Sales tab imports, and the new app has no export. Its only output is an on-screen "Transaction History": one collapsible receipt per sale, showing buyer ("Everbloom"), line items (`1× Small Fall Arrangement`, `1 × $309.99`), subtotal, tax (7.25%), total, funding/event ("Trade Show Credit · OC Trade Show"), date/time and a UUID transaction ID. Copy-pasting is impractical because every receipt is a collapsed dropdown.

These sales happen on the **same day as a normal tradeshow** that is imported from the sheet. So the usual job is "add these to OC Trade Show." Occasionally a new event is needed.

Volume is low, so the user counts how many of each product sold and types the counts in. Item names and prices in the new app are the same as in the sheet. (The user isn't sure of the exact formatting, which is why entry goes through product links and not typed names.)

Second need from the same conversation: **there's no way to fix a sale once it's in.** Both manual and imported sales need Edit and Delete.

## Goals

1. Enter sales for an event from a count table: pick products, set quantities.
2. Each unit in the count table becomes its **own sale** (qty 2 = 2 transactions), matching the one-item receipts.
3. **Group sales:** several products in a single transaction, for receipts with more than one item.
4. Attach to an **existing event** or create a **new event** from the same window.
5. Manual sales use **product linking**, so they count in Product Breakdown, best sellers, margins, break-even and dashboards exactly like imported sales.
6. **Re-imports never remove manual sales.**
7. **Edit and Delete** for every sale in the Transactions list, manual and imported alike.

## Non-goals

- Bulk import of the new app's data (it has no export). Future ideas are listed at the end; none are part of this build.
- Keeping edits to *imported* sales across re-imports. The user decided that **a re-import overwrites edits to imported sales**, which is fine.
- Visual changes to existing UI beyond the new button and the new Edit/Delete column (standing user preference).
- Storing the new app's UUID transaction IDs.

## Anti-goals (would count as failure even if it "works")

- A re-import of the tradeshow sheet silently deletes manual sales.
- Manual sales landing in a "manual" bucket that drops out of Online vs Trade Show splits, event totals or product analytics.
- The event's journal entry silently drifting out of sync after sales are added, edited or deleted.
- A design that looks unlike the VE Sales tab.

## Constraints

- Vanilla JS, sql.js in the browser, existing `Database` / `UI` / `App` structure (see CLAUDE.md map).
- Must survive the existing migration pattern (`try SELECT col … catch ALTER TABLE`) in `database.js`.
- Visual language must match VE Sales: `.btn-primary`, `.ve-subtabs` style segmented control, `.ve-table` (primary-colored header, DM Mono numbers), `.ve-source-badge` pills, `.ve-source-card`, `.pvm-search`, `.modal-content`. New CSS goes in a new block and must not restyle existing selectors. Dark mode must work (`[data-theme="dark"]` patterns already used for `.ve-table`).

## Decisions

### Data model

- New column **`ve_sales.is_manual INTEGER DEFAULT 0`** (migration). Manual sales have `is_manual = 1`.
- Manual sales use **`source = 'tradeshow'`** so every existing Online / Trade Show filter, split and journal flow includes them unchanged. If the user picks an existing event whose sales are online, use `'online'`. Rule: use the source of the event's existing sales, and default to `'tradeshow'` for new or empty events.
- Transaction numbers: **`MAN-0001`, `MAN-0002`, …**, the next number after the highest existing `MAN-` number. This never collides with imported numbers. `UNIQUE(transaction_no, source)` still holds.
- Line items go in `ve_sale_items` as usual (`source` = the sale's source, `inferred = 0`, `taxable = 1` when tax rate > 0).
- `event_id` is set directly on insert.
- Sale fields: `date` = sale date from the form; `billing_name` = group-sale buyer, or blank for single sales; `subtotal` = Σ price×qty; `tax` = subtotal × rate, rounded to cents per sale; `total` = subtotal + tax; `shipping`/`discount` = 0; `description` = `"Manual entry"`.

### Protecting manual sales from re-import

- `Database.clearVESales(source)` and `clearVESales()` (no source) must **skip `is_manual = 1` sales and their items**. This covers every import path: Excel per source (`app.js` `veHandleExcelImport`), JSON import, and server sync.
- The **"Clear All Data"** button is the exception: it says "all", so it removes manual sales too. Its confirm text should mention that manual sales are included (e.g. "Clear all VE Sales data, including N manual sales?").
- `upsertVESales` uses `INSERT OR REPLACE`. Imported numbers can't collide with `MAN-` numbers, so nothing more is needed.

### Product linking

- The count table lists all **non-discontinued products** (`Database.getProducts()`), searchable by name (same behavior as the Sales Data link picker).
- For each product, the name+price written to the sale item comes from its links (`getMappingsForProduct`):
  - one link → use it;
  - several links → default to the one whose price equals `product.price`, otherwise the first; the row shows a small selector to choose another *(builder's call how compact)*;
  - no link → use `product.name` / `product.price`, and on save call `addProductVeMapping(product.id, product.name, product.price)` so the new item is linked automatically.
- The price shown is that link's price. It isn't editable in the count table (edit the sale afterwards if needed).

### The Add Sales window

Entry point: **"+ Add Sales"** `.btn-primary .btn-sm` at the right end of `.ve-controls` on the VE Sales tab. Hidden in view-only mode, like other actions.

Layout, top to bottom (see mockup):
1. **Event:** segmented control `Existing | + New`.
   - Existing: dropdown of events (newest first, "Name · date"). If the VE controls bar's Event filter is set, preselect that event.
   - New: name (required), start date (defaults to sale date), end date (defaults to start), type `tradeshow`. Created on save via the existing add-event DB call.
2. **Sale date:** defaults to the selected event's start date, or today for a new event.
3. **Tax rate %:** defaults to the last-used manual tax rate (stored in `app_meta`), first time 7.25.
4. **Single Sales:** search box plus `.ve-table` of Product | Price | Qty (−/+ stepper, also typeable) | Line Total. Rows with qty > 0 get the light-blue highlight. Hint: "Each item is its own sale. Qty 2 = 2 sales."
5. **Group Sales:** zero or more `.ve-source-card` cards titled "Group Sale N", each with optional Buyer, product rows with qty steppers, "+ Add product" (opens a product search like the link picker) and Remove. A dashed "+ Add Group Sale" button sits below.
6. **Footer:** summary ("Creates **4 sales** · 3 single + 1 group · 6 items"), Subtotal, Tax (red), Total (green), all live.
7. **Journal notice:** if the target event has `journal_added = 1`, show an amber notice that saving will ask about updating its journal entry.
8. Buttons: Cancel and "Add N Sales" (disabled at 0).

Save: insert the sales and items in one go, then call `veRebuildIgnoredKeys`/`veRenderAll` (whatever refreshes the tab today) and show a success toast ("Added 4 sales to OC Trade Show"). Push an undo entry if the undo system supports it cheaply *(builder's call)*.

### Edit / Delete in the Transactions list

- Add a final actions column to the VE Transactions table (`index.html` around line 496; `veRenderTransactions` in `app.js`) with **Edit** and **Delete**. Fix the empty-state `colspan`.
- Manual sales show a second badge **"Manual"** (`.ve-source-badge.manual` with `--amber-bg` / `--amber`) next to the source badge.
- **Delete:** confirm, then remove the sale and its items. Works for imported sales too. The deleted imported sale comes back on the next re-import, which is expected.
- **Edit:** opens the same Add Sales window in **single-sale mode** for that one sale:
  - title "Edit Sale MAN-0004" (or the imported number), one item list (the group-sale card layout) prefilled with its items, plus event, date, tax rate (derived as tax/subtotal) and buyer;
  - items can be added, removed or have quantities changed; the event can be reassigned;
  - imported sales show the note: "Changes to imported sales are replaced the next time you re-import that sheet."
  - Imported items that aren't linked to any product show by their raw name and price, and stay as they are unless removed.
- Edit and Delete are hidden in view-only mode.

#### Edit rules (settled by the designer after review; marked as designer's calls)

- **Totals on edit:** if no item, quantity or tax rate changed (for example, only the event or date moved), keep the stored `subtotal`/`tax`/`total`/`shipping`/`discount` exactly. If items changed, recalculate `subtotal` from items, **keep `shipping` and `discount` as stored**, compute tax as rate × (sum of *taxable* items only), and `total = subtotal − discount + shipping + tax`. The rate is prefilled as `tax ÷ taxable subtotal` (or the last-used rate if that's 0). Newly added items are taxable when the rate is > 0. Manual sales follow the same rules; their shipping and discount are just 0.
- **Source on move:** imported sales **always keep their sheet's source**. Manual sales take the new event's source: the source of that event's existing non-manual sales, otherwise the event's `type` (`online_event` → `'online'`, everything else → `'tradeshow'`). The same rule replaces the "default to `'tradeshow'`" rule for inserts into empty events. Changing source re-keys the sale and its items' `source` together.
- **No event:** Add mode requires an event. Edit mode adds a **"No event"** choice to the dropdown, so a sale (imported or manual) can stay or become unassigned. Journal sync then only fires for the side that has a journaled event.
- **Clear All** gets its own DB path (e.g. `clearVESales(null, { includeManual: true })`), because today it shares `clearVESales()` with JSON import and server sync, which must keep manual sales.

### Event journal sync

If the event's `journal_added = 1` and its sales change (add, edit, delete, or a sale moved between events), prompt: "This event's journal entry is out of date. Update it now?" On Yes, call the existing `veUpdateJournalForEvent(eventId, evt)` and refresh. On No, leave it. For moves between two journaled events, handle both.

## Left to the builder

- Code organization (new methods in `app.js` near the VE section, DB helpers in `database.js` VE section).
- Exact markup and CSS class names for new pieces, as long as they reuse the existing look.
- Selector UI for products with several links.
- Keyboard handling (Enter moves to the next qty is nice to have).
- Tests: add cases to `tests/ve-sales.test.html` for clearVESales keeping manual sales, MAN numbering, sale totals and tax rounding, and edit/delete.
- Update CLAUDE.md's map with the new locations.

## Future ideas (ideation pass, not in this build)

| Idea | Why | Effort | Impact |
|---|---|---|---|
| **Reconcile against the POS total.** Type the new app's day total; the window shows the difference from what you entered | Catches a miscounted product before saving | S | High |
| **Duplicate warning.** If the event already has manual sales on that date, show "N manual sales already on this event" | Stops entering the same day twice | S | Med |
| **CSV template.** Download a Product/Qty CSV, fill it, upload it to prefill the count table | The "import later" path; reuses the same window | M | Med |
| **Paste the collapsed list.** The collapsed rows show "1× Item · total · event", so a paste parser could cover the single-item receipts | Bulk without opening each dropdown | M | Med (only if volume grows) |
| **Screenshot import.** Read receipts from a screenshot | Most hands-off, but needs an AI service and doesn't suit a local-first app | L | Low for now |
