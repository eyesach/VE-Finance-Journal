# Manual VE Sales Entry + Sale Edit/Delete Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Let the user add VE sales by product count (single and group sales) to an existing or new event, protected from re-import, and edit/delete any VE sale.

**Architecture:** One new column (`ve_sales.is_manual`) keeps manual sales out of every import-time delete. The maths lives in `utils.js` as pure functions so it can be tested, persistence lives in new `Database` VE methods, and one modal (`#veAddSalesModal`) serves both add mode and single-sale edit mode. The Transactions table gets a Manual badge and Edit/Delete actions.

**Tech Stack:** Vanilla JS, sql.js, existing `App` / `Database` / `UI` / `Utils` objects, browser-run HTML tests.

**Spec:** `docs/superpowers/specs/2026-11-13-manual-ve-sales-design.md` (mockup: `.superpowers/brainstorm/1855-1794637990/content/cohesive-v3.html`)

## Global Constraints

- Manual sales: `is_manual = 1`, `transaction_no` `MAN-0001` style (4-digit zero pad, next after highest), `description = "Manual entry"`, `shipping = 0`, `discount = 0`, `billing_name` = group buyer or `''`.
- Source rule: source of the event's existing non-manual sales (most frequent), else event `type === 'online_event'` → `'online'`, else `'tradeshow'`. Imported sales never change source on edit.
- Tax: `rate` is a percent number (7.25). Tax = round-to-cents(rate/100 × sum of taxable item amounts), per sale. `total = subtotal − discount + shipping + tax`.
- Default tax rate: last used, stored with `Database.setMetaValue('ve_manual_tax_rate', rate)`; first time `7.25`.
- Every import path (`clearVESales(source)`, `clearVESales()`) keeps manual sales. Only "Clear All Data" removes them, and its confirm reads: `Clear all VE Sales data, including N manual sales? This cannot be undone.` (omit the "including…" clause when N = 0).
- Edit note for imported sales, verbatim: `Changes to imported sales are replaced the next time you re-import that sheet.`
- UI reuses the existing look: `.btn .btn-primary`, `.btn-secondary`, `.ve-subtabs`/`.ve-subtab` styling, `.ve-table`, `.ve-source-badge`, `.ve-source-card`, `.pvm-search`, `.modal`/`.modal-content`. New CSS goes in a new `/* ===== VE MANUAL SALES ===== */` block and must not restyle existing selectors. Dark mode via existing `[data-theme="dark"]` patterns.
- Do not restyle existing UI. Only add the buttons, the Manual badge and the Edit/Delete column.
- All user-supplied or product strings rendered with `Utils.escapeHtml`.
- Hidden or disabled when `App.isViewOnly`.

## Review Focus

1. **Empty VE tab (nothing imported yet):** `#veControls` is hidden, so Add Sales must also be reachable from `#veEmptyState`. After saving, the dashboard appears. → Task 4 manual check.
2. **Bad quantity input** (blank, negative, decimal, letters) → treated as 0 or floored to an integer ≥ 0. "Add N Sales" stays disabled while N = 0. → Task 2 test `sanitizeQty`.
3. **Re-import after manual sales were assigned to an event** → manual sales keep `event_id` and their items. → Task 1 test.
4. **Same transaction number in both sources** (`X1` online and `X1` tradeshow): edit/delete of one must not touch the other. → Task 3 test.
5. **Saving the same unlinked product twice** → exactly one mapping row (no duplicates), and Product Breakdown attributes both sales to the product. → Task 3 test.

## How to run tests

New file `tests/manual-sales.test.html`. It uses the same bootstrap as `tests/sales-data.test.html` (sql.js CDN, `../js/utils.js`, `window.CompanyManager = {}`, `../js/database.js`, `freshDb()`, `suite/test/assert/assertEqual/assertClose`, summary text `N/M tests passed, F failed`).

Run (Git Bash):
```bash
"/c/Program Files/Google/Chrome/Application/chrome.exe" --headless=new --disable-gpu --allow-file-access-from-files --virtual-time-budget=20000 --dump-dom "file:///C:/Users/mrisa/Claude%20Projects/Accounting%20Journal%20Calculator/tests/manual-sales.test.html" | grep -oE '[0-9]+/[0-9]+ tests passed, [0-9]+ failed|CRASH:[^<]*|✘[^<]*'
```
Pass = `N/N tests passed, 0 failed` and no `✘` lines. Also re-run `tests/sales-data.test.html` and `tests/ve-sales.test.html` the same way at the end of Tasks 1 and 3 to check nothing regressed.

---

### Task 1: `is_manual` column and import protection

**Files:**
- Modify: `js/database.js` (`migrateSchema`, near the `ve_sale_items.source` migration ~line 758; `clearVESales` ~line 4450)
- Modify: `js/app.js` (`#veClearBtn` handler ~line 14104)
- Create: `tests/manual-sales.test.html`

**Interfaces:**
- Produces: column `ve_sales.is_manual INTEGER DEFAULT 0`; `Database.clearVESales(source = null, { includeManual = false } = {})`; `Database.getManualSaleCount() -> number`.

- [ ] **Step 1: Create the test file with bootstrap and failing tests** (suite "Import protection"):
  - `fresh schema has is_manual`: `PRAGMA table_info(ve_sales)` includes `is_manual`.
  - `migration adds is_manual to old db`: drop/recreate `ve_sales` without the column, run `migrateSchema()`, column exists.
  - `clearVESales('tradeshow') keeps manual tradeshow sales and items`: seed imported `T1` (tradeshow, 1 item) and manual `MAN-0001` (tradeshow, `is_manual=1`, `event_id=5`, 1 item tagged `tradeshow`). After the clear: `ve_sales` rows = `['MAN-0001']`, its `event_id === 5`, `ve_sale_items` count = 1.
  - `clearVESales() keeps manual sales`: the same seed plus online `O1`. After the clear, only `MAN-0001` and its item remain.
  - `clearVESales(null, {includeManual:true}) removes everything`: 0 sales, 0 items.
  - `getManualSaleCount`: returns 1 for the seed.
- [ ] **Step 2: Run tests** and confirm they fail (`is_manual` missing / `getManualSaleCount is not a function`).
- [ ] **Step 3: Implement**
  - Migration: `try SELECT is_manual FROM ve_sales LIMIT 1 catch ALTER TABLE ve_sales ADD COLUMN is_manual INTEGER DEFAULT 0`.
  - `clearVESales`: unless `includeManual`, add `AND COALESCE(is_manual,0) = 0` to the sale delete. Item deletes must exclude items whose `(transaction_no, source)` belongs to a manual sale. `getManualSaleCount` = `SELECT COUNT(*) FROM ve_sales WHERE is_manual = 1`.
- [ ] **Step 4: Update `#veClearBtn`** to build the confirm text from `getManualSaleCount()` (exact copy in Global Constraints) and call `Database.clearVESales(null, { includeManual: true })`.
- [ ] **Step 5: Run the new tests plus `sales-data` and `ve-sales` tests.** All pass.
- [ ] **Step 6: Commit** `feat(ve-sales): is_manual flag; imports keep manual sales`

---

### Task 2: Pure maths helpers in utils.js

**Files:**
- Modify: `js/utils.js` (new section `// ===== VE manual sales =====` at the end of the `Utils` object)
- Test: `tests/manual-sales.test.html` (suite "Manual sale maths")

**Interfaces:**
- Produces:
  - `Utils.sanitizeQty(v) -> integer ≥ 0`
  - `Utils.nextManualTxNo(existingTxNos: string[]) -> string`
  - `Utils.pickProductLink(product, mappings: {name, price}[]) -> { name, price, linked: boolean, options: {name, price}[] }`
  - `Utils.computeVESaleTotals(items: {price, quantity, taxable}[], ratePct, { shipping = 0, discount = 0 } = {}) -> { subtotal, tax, total, items: [...with amount] }`
  - `Utils.buildManualSales({ singles: {name, price, quantity}[], groups: {buyer, items: {name, price, quantity}[]}[], ratePct, date, startTxNo }) -> Array<{ transactionNo, date, billingName, subtotal, tax, total, items }>`. `startTxNo` is the first number to use, e.g. `'MAN-0007'`; numbers increment per sale.

- [ ] **Step 1: Write failing tests**
  - `sanitizeQty`: `''→0`, `'-3'→0`, `'2.7'→2`, `'abc'→0`, `5→5`.
  - `nextManualTxNo`: `[]→'MAN-0001'`; `['MAN-0002','T9','MAN-0010']→'MAN-0011'`.
  - `pickProductLink`, product `{name:'Small Fall', price:309.99}`:
    - no mappings → `{name:'Small Fall', price:309.99, linked:false}`;
    - one mapping `{name:'Small Fall Arrangement', price:309.99}` → that one, `linked:true`;
    - two mappings with prices 300 and 309.99 → the 309.99 one, and `options.length === 2`;
    - two mappings with neither price matching → the first.
  - `computeVESaleTotals([{price:309.99,quantity:1,taxable:true}], 7.25)` → subtotal 309.99, tax 22.47, total 332.46 (the receipt in the spec).
  - `computeVESaleTotals` mixed: items `{100,1,taxable}` and `{50,1,untaxed}`, rate 10, shipping 5, discount 20 → subtotal 150, tax 10, total 145.
  - `buildManualSales`: singles `[{Large,850,2},{Medium,410,1}]` plus one group `{buyer:'Booth 14', items:[{Small,309.99,1},{Tulip,95,2}]}`, rate 7.25, `startTxNo 'MAN-0001'` →
    - length 4, tx nos `MAN-0001..MAN-0004`;
    - first sale has 1 item with quantity 1 and total 911.63;
    - the group sale has `billingName 'Booth 14'`, 2 items and subtotal 499.99;
    - every item is `taxable: true` when rate > 0, and `false` when rate is 0.
- [ ] **Step 2: Run tests.** They fail (functions undefined).
- [ ] **Step 3: Implement the five functions.** Round to cents with `Math.round(x*100)/100` on every sum. Singles expand to `quantity` separate one-unit sales, in input order, before groups. Group items with quantity 0 are dropped, and a group left with no items produces no sale.
- [ ] **Step 4: Run tests.** All pass.
- [ ] **Step 5: Commit** `feat(ve-sales): manual sale maths helpers`

---

### Task 3: Database writes (insert, update, delete, source rule)

**Files:**
- Modify: `js/database.js` (VE SALES OPERATIONS section, after `getAllVESaleItems`)
- Test: `tests/manual-sales.test.html` (suite "Manual sale persistence")

**Interfaces:**
- Consumes: `Utils.buildManualSales` output shape (Task 2).
- Produces:
  - `Database.getAllVETxNos() -> string[]` (all `transaction_no`, for `nextManualTxNo`)
  - `Database.getEventSaleSource(eventId|null) -> 'online'|'tradeshow'` (source rule from Global Constraints; `null` → `'tradeshow'`)
  - `Database.insertManualSales(sales, { eventId, source }) -> string[]`: inserts with `is_manual=1` and `description='Manual entry'`, items tagged with `source`, `inferred=0`, item `amount = price*quantity`; returns tx nos.
  - `Database.updateVESale(origTxNo, origSource, sale: {transactionNo, date, billingName, subtotal, tax, shipping, discount, total, source, eventId}, items) -> void`: updates the row identified by `(origTxNo, origSource)`, keeping `is_manual` and `description`. Replaces that sale's items (by txNo + `origSource`) with `items` tagged with the new `source`.
  - `Database.deleteVESale(txNo, source) -> void`: removes the row and its items (tagged `source`, or untagged when no other source shares `txNo`).
  - `Database.ensureProductLink(productId, name, price) -> void`: thin wrapper over `addProductVeMapping` (INSERT OR IGNORE).

- [ ] **Step 1: Write failing tests**
  - `insertManualSales stores flag, event, source, items`: insert 2 built sales with `{eventId: 3, source: 'tradeshow'}` → both rows `is_manual=1`, `event_id=3`, `description='Manual entry'`; `getAllVESaleItems().get('MAN-0001')[0].source === 'tradeshow'`.
  - `getEventSaleSource`:
    - an event with imported online sales and one manual tradeshow sale → `'online'`;
    - an empty event of type `online_event` → `'online'`;
    - an empty event of type `tradeshow` → `'tradeshow'`;
    - `null` → `'tradeshow'`.
  - `updateVESale changes items and totals`: update `MAN-0001` to 2 items → item count 2, `total` stored.
  - `updateVESale re-keys source`: manual tradeshow → online. The row has source `online`, and its items are tagged `online`.
  - `updateVESale can unassign event`: `eventId: null` → `event_id IS NULL`.
  - **Review Focus 4:**
    - seed `X1` online and `X1` tradeshow with one item each;
    - `deleteVESale('X1','tradeshow')` leaves the online `X1` row and its item;
    - `updateVESale('X1','online', …)` leaves the tradeshow `X1` untouched.
  - **Review Focus 5:**
    - `ensureProductLink(p, 'Small Fall', 309.99)` twice → 1 row in `product_ve_mappings`;
    - after inserting two manual sales of that item, `getLinkedProductAnalytics(null, null, 'all').byProduct` shows `units_sold` 2 for `p`.
  - `manual sale survives re-import end to end`: insert a manual sale, then `clearVESales('tradeshow')` + `upsertVESales([...])` → the manual row still present.
- [ ] **Step 2: Run tests.** They fail.
- [ ] **Step 3: Implement the six methods** with prepared statements, each ending in `this.autoSave()`.
- [ ] **Step 4: Run the new tests plus `sales-data` and `ve-sales` tests.** All pass.
- [ ] **Step 5: Commit** `feat(ve-sales): manual sale insert/update/delete persistence`

---

### Task 4: Add Sales modal (add mode)

**Files:**
- Modify: `index.html`. Add `+ Add Sales` button (`id="veAddSalesBtn"`, `.btn .btn-primary .btn-sm`) at the end of `#veControls` after the Sort by group; a second button (`id="veAddSalesEmptyBtn"`) inside `#veEmptyState`; new `#veAddSalesModal` next to `#veLinkPickerModal` (~line 1632).
- Modify: `js/app.js`. New methods after `sdBindListeners`; bind listeners in the VE listener setup next to `#veClearBtn`.
- Modify: `css/styles.css`. New `/* ===== VE MANUAL SALES ===== */` block after the `.pvm-*` rules (~line 3148).

**Interfaces:**
- Consumes: Tasks 1–3 functions; existing `Database.getProducts`, `getMappingsForProduct`, `getAllVEEvents`, `createVEEvent`, `getMetaValue/setMetaValue`, `App.refreshVESales`, `App.veUpdateJournalForEvent(eventId, evt)`, `UI.showModal/hideModal/showNotification`.
- Produces (used by Task 5):
  - `App.veOpenAddSales({ editSale = null } = {})`
  - `App._veMs` state object: `{ mode: 'add'|'edit', eventMode: 'existing'|'new', eventId, date, ratePct, singles: Map<productId, qty>, groups: [{buyer, rows: [{productId|null, name, price, qty, taxable}]}], orig }`
  - `App.veMsRender()`
  - `App.veMsSave()`
  - `App.veMsSyncJournals(eventIds: number[])`: refreshes VE data, then for each event with `journal_added === 1`, `confirm('This event's journal entry is out of date. Update it now?')`, and on yes calls `veUpdateJournalForEvent` followed by `refreshAll()`.

Modal content, top to bottom (match the mockup):
1. Title `Add Sales`, plus hint `For sales that aren't in your import sheets. Re-imports never remove these.`
2. Event segmented control `Existing | + New`:
   - Existing: dropdown `"Name · Mon D, YYYY"`, newest first, preselecting `ve-filterEvent` if set.
   - New: Name (required), Start date (defaults to sale date) and End date (defaults to start).
3. Sale date: defaults to the event start, or today for a new event.
4. Tax rate %.
5. **Single Sales** `.ve-table`:
   - `.pvm-search` filters rows by name;
   - columns Product (with a `Linked` pill when `pickProductLink(...).linked`, plus a compact `<select>` when `options.length > 1`), Price, Qty stepper (− input +), Line Total;
   - rows with qty > 0 get a `--blue-bg` highlight;
   - hint `Each item is its own sale. Qty 2 = 2 sales.`
6. **Group Sales:** cards with Buyer input, product rows with steppers, `+ Add product` (inline `.pvm-search` dropdown of products) and Remove; dashed `+ Add Group Sale` below.
7. Footer: `Creates N sales · S single + G group · I items`, Subtotal, Tax (red), Total (green).
8. Amber notice when the selected event has `journal_added === 1`: `This event's journal entry has already been created. Saving will ask whether to update it.`
9. Buttons: Cancel, and `Add N Sales` (disabled when N = 0).

- [ ] **Step 1: Add markup, CSS and entry buttons.** Both buttons open the modal and are hidden when `isViewOnly`.
- [ ] **Step 2: Implement `veOpenAddSales` and `veMsRender`.**
  - Products come from `getProducts().filter(p => !p.is_discontinued)`.
  - Quantities pass through `Utils.sanitizeQty`.
  - Totals come from `Utils.buildManualSales` on every input change, which also gives N, S, G and I.
  - Re-render only the footer and changed row on qty input, so focus stays in the field.
- [ ] **Step 3: Implement `veMsSave` (add mode).**
  - Validate: a new event needs a name (else `UI.showNotification('Enter an event name.', 'error')`), and N must be > 0.
  - Create the event if new (`type: 'tradeshow'`).
  - Set `source = Database.getEventSaleSource(eventId)`.
  - Set `startTxNo = Utils.nextManualTxNo(Database.getAllVETxNos())`.
  - Call `insertManualSales`, `ensureProductLink` for every unlinked product used, and `setMetaValue('ve_manual_tax_rate', ratePct)`.
  - Hide the modal, call `refreshVESales()`, toast `Added N sales to <event name>`, then `veMsSyncJournals([eventId])`.
- [ ] **Step 4: Manual check in the browser** (open `index.html` via the existing dev setup):
  - (a) On a company with no VE data, the empty-state button opens the modal. Saving 2 singles creates the event, and the dashboard appears with 2 transactions (Review Focus 1).
  - (b) Entering 1 × $309.99 at 7.25% shows a total of $332.46.
  - (c) Re-importing the tradeshow sheet keeps the manual sales.
  - (d) Product Breakdown counts the manual units.
  - (e) Dark mode looks right.
  - (f) In view-only mode, the buttons are hidden.
- [ ] **Step 5: Commit** `feat(ve-sales): Add Sales modal for manual single and group sales`

---

### Task 5: Transactions list Edit/Delete and edit mode

**Files:**
- Modify: `index.html` (`#ve-txTable` header ~line 1357: add `<th></th>` after Items)
- Modify: `js/app.js` (`veRenderTransactions` ~line 12388: Manual badge, actions cell, empty-state `colspan="10"`; delegated click listener on `#ve-txBody`; edit-mode branches in `veOpenAddSales`/`veMsRender`/`veMsSave`)
- Modify: `css/styles.css` (`.ve-source-badge.manual { background: var(--amber-bg); color: var(--amber); }` and action button styles, in the Task 4 block)
- Modify: `CLAUDE.md` (VE Sales map entries for the new methods, table column and test file); `app.js` changelog list (~line 15029) with one `feature` entry.

**Interfaces:**
- Consumes: Task 3 `updateVESale`, `deleteVESale`, `getEventSaleSource`; Task 2 `computeVESaleTotals`; Task 4 `veOpenAddSales`, `_veMs`, `veMsSyncJournals`.

Behavior:
- Row actions: `✎ Edit` and `Delete` buttons carrying `data-txno` and `data-source`, hidden when `isViewOnly`. Manual rows add `<span class="ve-source-badge manual">Manual</span>` after the source badge.
- **Delete:** `confirm('Delete sale <txno>? ' + (imported ? 'It will come back if you re-import that sheet.' : 'This cannot be undone.'))`, then `deleteVESale`, `refreshVESales()`, toast, and `veMsSyncJournals([event_id])` when the sale had an event.
- **Edit mode:**
  - Title `Edit Sale <txno>`.
  - Shows one group-style card prefilled with the sale's items (`veItemsForSale`), mapped to products where a link matches name+price, else raw name/price rows. Buyer = `billing_name`. No Single Sales table.
  - Event dropdown gains a `No event` option. Date is prefilled.
  - Rate prefilled as `tax ÷ taxable subtotal × 100` (rounded to 2 decimals), or the last-used rate when that is 0.
  - Imported sales show the verbatim note from Global Constraints.
  - Button `Save Changes`.
- **Edit save:**
  - If the items (name, price, qty, taxable) and rate are unchanged, keep the stored `subtotal`/`tax`/`total`/`shipping`/`discount`. Otherwise use `Utils.computeVESaleTotals(items, rate, { shipping, discount })` with the stored shipping/discount. Newly added items are `taxable = rate > 0`.
  - Source: imported keeps `orig.source`; manual uses `getEventSaleSource(newEventId)`.
  - Call `updateVESale(orig.transaction_no, orig.source, …)` and `refreshVESales()`, then `veMsSyncJournals` with the old and new event ids (deduped, non-null).

- [ ] **Step 1: Write a failing test** (suite "Edit rules") for the unchanged-detection helper `Utils.veSaleItemsChanged(origItems, newItems, origRate, newRate) -> boolean`:
  - same items in a different order and the same rate → `false`;
  - qty changed → `true`;
  - rate changed by more than 0.005 → `true`.
- [ ] **Step 2: Run the test.** It fails.
- [ ] **Step 3: Implement `Utils.veSaleItemsChanged`**, then the table column, badge, actions, delete flow and edit mode as described.
- [ ] **Step 4: Run all three test files.** All pass.
- [ ] **Step 5: Manual check:**
  - (a) Edit a manual sale's qty: totals update and the journal prompt appears for a journaled event.
  - (b) Move a manual sale from a tradeshow event to an online event: the Source badge flips to Online.
  - (c) Change only the event of an imported sale with shipping: its total is unchanged.
  - (d) Set an imported sale to No event: the Event column shows "—".
  - (e) Delete an imported sale, then re-import: it returns.
  - (f) Clear All Data: the confirm mentions the manual count, and everything is cleared.
- [ ] **Step 6: Update CLAUDE.md and the changelog, then commit** `feat(ve-sales): edit and delete VE sales; manual badge`
