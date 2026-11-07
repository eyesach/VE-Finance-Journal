# Sales Data Subtab Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a "Sales Data" subtab to VE Sales. It shows unit-weighted average price and COGS (with a "Use in Break-even" fill-in), product margins and mix, best sellers by month, shipping, a % linked indicator with an Excluded Sales modal, and reversible "test item" exclusion.

**Architecture:** The data layer returns flat, filtered line rows from SQLite (`Database.getSalesDataLines`). A pure function (`Utils.computeSalesData`) turns those rows into every number the UI shows. That keeps the math unit-testable without the DOM. The UI renders from that result using existing VE Sales classes. Ignored items live in a new `ve_ignored_items` table keyed by (name, price), the same key `product_ve_mappings` uses.

**Tech Stack:** Vanilla JS, sql.js (SQLite in browser), existing browser-based test pages in `tests/*.test.html` (open in a browser; summary line shows pass/fail counts).

**Spec:** `docs/superpowers/specs/2026-11-07-sales-data-subtab-design.md` (read it; mockup at `.superpowers/brainstorm/153-1794081169/content/sales-data-layout-v2.html`)

## Global Constraints

- All figures are **pre-tax**, with **discount allocated proportionally**: `vs.discount * (vi.amount / vs.subtotal)` when `vs.subtotal > 0`, else 0.
- Weighting is **by units sold** (`SUM(vi.quantity)`).
- Shipping is never part of price, COGS, or margin.
- Break-even values change **only** when the user clicks "Use in Break-even".
- Nothing is deleted from `ve_sales` or `ve_sale_items`. Ignoring is reversible.
- An item that is both linked and ignored is treated as **ignored**.
- A sale line linked to several products counts **once** and is attributed to the **lowest `product_id`**.
- Match key for mappings and ignored items: `ve_item_name = vi.name AND ROUND(ve_item_price, 2) = ROUND(vi.price, 2)`.
- Filter values come from the existing bar: source `'both'|'online'|'tradeshow'`, from/to `YYYY-MM-DD` or empty, event `''|'__unassigned__'|'<id>'`.
- Reuse existing classes (`.ve-subtab`, `.ve-card*`, `.ve-table*`, `.ve-source-card`, `.ve-source-stat`, `.btn*`, `.modal`). New CSS uses existing tokens only and must work in dark mode.
- Order-level Sales subtab figures (summary cards, source breakdown, Transactions, journal creation) are **unchanged** by ignoring.
- In view-only mode, action buttons are hidden through `applyViewOnlyRestrictions` (`js/app.js:11774`).
- Call `this.autoSave()` after every DB write.

## Review Focus

1. **Order with `subtotal = 0`** (fully comped or bad import): discount allocation must yield 0, not NaN/Infinity. Covered in the Task 2 test `lines: zero-subtotal order allocates no discount`.
2. **The same item linked to two products with different COGS:** counts once, attributed to the lowest id, conflict reported. Covered in the Task 3 test `multi-link: conflicting cogs flagged`.
3. **Window with only unlinked or ignored items:** weighted values are `null` (the UI shows "—" and the button is disabled), not NaN. `linkedPct` is `null` when the denominator is 0. Covered in the Task 3 test `empty linked set returns nulls`.
4. **"Last month" in January:** must roll back to December of the previous year. Covered in the Task 3 test `getLastMonthsRange: january rollover`.
5. **Item names with HTML characters** (`<b>Mug</b>`, quotes) in the modal, table, and Test Items section: they must render as text. All new renderers use `Utils.escapeHtml` on names. Checked in the Task 6 manual verification step.

---

## File Map

| File | Change |
|---|---|
| `js/database.js` | `ve_ignored_items` table (schema + migration); ignored-item CRUD; `addProductVeMapping`; `getSalesDataLines`; `getShippingSummary` |
| `js/utils.js` | `computeSalesData`, `getLastMonthsRange` (pure) |
| `js/app.js` | Presets; subtab switch; `veRenderSalesData`; Use in Break-even; Excluded Sales modal; Product Breakdown exclusion; Test Items section; view-only ids; changelog |
| `index.html` | Third subtab button + `#veSubtabData` content; `#veExcludedModal`; `#veLinkPickerModal`; Test Items section in `#veSubtabSales` |
| `css/styles.css` | `.sd-*` additions next to the VE block (~line 2905) |
| `tests/sales-data.test.html` | New test page: DB tests (sql.js + `../js/database.js`) and pure-function tests (`../js/utils.js`) |
| `CLAUDE.md` | Code-map entry for Sales Data |

---

### Task 1: Ignored-items table, CRUD, and single-mapping add

**Files:**
- Modify: `js/database.js`: `createSchema()` (after `product_ve_mappings`, ~line 345), `migrateSchema()` (after the mapping-table creation, ~line 726), and new methods next to `setMappingsForProduct` (~line 2884)
- Create: `tests/sales-data.test.html`

**Interfaces:**
- Produces:
  - `Database.getIgnoredItems() -> Array<{id:number, name:string, price:number}>`, ordered by name
  - `Database.addIgnoredItem(name:string, price:number) -> void` (INSERT OR IGNORE, price rounded to 2dp)
  - `Database.removeIgnoredItem(name:string, price:number) -> void`
  - `Database.addProductVeMapping(productId:number, name:string, price:number) -> void` (INSERT OR IGNORE; does not touch other mappings)
  - Test harness globals in `tests/sales-data.test.html`: `suite`, `test`, `assert`, `assertEqual`, `assertClose`, and an async `freshDb()` that returns `Database` with a new in-memory DB and schema

- [ ] **Step 1: Create the test page with the harness and failing tests**

Copy the harness block (styles, `suite`/`test`/`assert*`, summary) from `tests/notes-analysis.test.html`. Load `sql-wasm.js` (1.8.0 cdnjs), `../js/utils.js`, then `../js/database.js`. Put a stub `window.CompanyManager = {}` before `database.js`. `freshDb()`:

```js
async function freshDb() {
  if (!Database.SQL) Database.SQL = await initSqlJs({ locateFile: f => `https://cdnjs.cloudflare.com/ajax/libs/sql.js/1.8.0/${f}` });
  Database.autoSave = () => {};
  Database.db = new Database.SQL.Database();
  Database.createSchema();
  return Database;
}
```

Wrap the tests in an async IIFE (tests are sync once the DB is ready). Tests:
- `ignored: add then list` → `addIgnoredItem('TEST ITEM', 1)` → `getIgnoredItems()` has length 1, `name === 'TEST ITEM'`, `price === 1`
- `ignored: duplicate add is ignored` → add twice → length 1
- `ignored: price rounding matches` → `addIgnoredItem('A', 2.005)` then `removeIgnoredItem('A', 2.01)` → length 0
- `ignored: remove` → add then remove → length 0
- `mapping: add single keeps existing` → `addProduct('P','',10,0,4,null)`, `setMappingsForProduct(1,[{name:'X',price:10}])`, `addProductVeMapping(1,'Y',10)` → `SELECT COUNT(*) FROM product_ve_mappings WHERE product_id=1` is 2
- `migration: table exists after migrateSchema on old db` → new DB, `createSchema()`, `DROP TABLE ve_ignored_items`, `migrateSchema()` → `SELECT name FROM sqlite_master WHERE name='ve_ignored_items'` returns 1 row

- [ ] **Step 2: Open `tests/sales-data.test.html` in a browser.** Expected: the new tests FAIL (`addIgnoredItem is not a function`).

- [ ] **Step 3: Implement.** Table DDL, in both `createSchema` and `migrateSchema` (`CREATE TABLE IF NOT EXISTS`):

```sql
CREATE TABLE IF NOT EXISTS ve_ignored_items (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  ve_item_name TEXT NOT NULL,
  ve_item_price REAL NOT NULL DEFAULT 0,
  created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
  UNIQUE(ve_item_name, ve_item_price)
)
```

Store `Math.round(price*100)/100`. `removeIgnoredItem` matches on `ROUND(ve_item_price,2) = ROUND(?,2)`. Each write calls `this.autoSave()`.

- [ ] **Step 4: Reload the test page.** Expected: all Task 1 tests PASS. Then run `tests/ve-sales.test.html` and `tests/notes-analysis.test.html` and confirm there are no new failures.

- [ ] **Step 5: Commit**

```bash
git add js/database.js tests/sales-data.test.html
git commit -m "feat(sales-data): ignored VE items table and single-mapping add"
```

---

### Task 2: Filtered line query and shipping summary

**Files:**
- Modify: `js/database.js` (new methods after `getLinkedProductAnalytics`, ~line 2970)
- Test: `tests/sales-data.test.html`

**Interfaces:**
- Consumes: Task 1 table
- Produces:
  - `Database.getSalesDataLines(filters:{source, from, to, event}) -> Array<SalesLine>`, where
    `SalesLine = { transaction_no:string, date:string, month:'YYYY-MM', name:string, price:number, quantity:number, amount:number, discount:number /*allocated*/, product_ids:number[] /*ascending, may be empty*/, ignored:boolean }`
  - `Database.getShippingSummary(filters) -> { total:number, orders:number /*count where shipping>0*/ }`

- [ ] **Step 1: Write failing tests.** Seed helper in the test page: `seedSales(db)` inserts these rows.
  - Products: `(1,'Bundle v1',price 25,cogs 13.75)`, `(2,'Sticker',6,1.2)`, `(3,'Sticker Rebrand',6,1.2)`
  - Mappings: product 1 ↔ (`'Bundle'`, 25); product 2 ↔ (`'Sticker'`, 6); product 3 ↔ (`'Sticker'`, 6)
  - Sales:
    - `T1`, 2026-08-03, online, subtotal 56, discount 5.60, shipping 6, event null
    - `T2`, 2026-09-10, tradeshow, subtotal 12, discount 0, shipping 0, event 7
    - `T3`, 2026-09-12, online, subtotal 0, discount 2, shipping 4
  - Items:
    - `T1`: (`'Bundle'`, 25, qty 2, amount 50) and (`'Sticker'`, 6, qty 1, amount 6)
    - `T2`: (`'Sticker'`, 6, qty 1, amount 6) and (`'TEST ITEM'`, 1, qty 6, amount 6)
    - `T3`: (`'Mystery'`, 0, qty 1, amount 0)
  - Ignored: (`'TEST ITEM'`, 1)

  Tests:
  - `lines: all filters empty returns 5 lines`
  - `lines: discount allocated proportionally` → T1 Bundle line `discount` ≈ 5.00 (5.60 × 50/56), Sticker ≈ 0.60
  - `lines: zero-subtotal order allocates no discount` → T3 line `discount === 0`
  - `lines: multi-link product_ids ascending` → T1 Sticker line `product_ids` deep-equals `[2,3]`
  - `lines: unlinked has empty product_ids` → Mystery `product_ids.length === 0`
  - `lines: ignored flag` → TEST ITEM `ignored === true`, others `false`
  - `lines: source filter` → `{source:'online'}` returns 3 lines
  - `lines: date filter` → `{from:'2026-09-01', to:'2026-09-30'}` returns 3 lines
  - `lines: event filter id` → `{event:'7'}` returns 2 lines; `{event:'__unassigned__'}` returns 3
  - `lines: month field` → T2 lines have `month === '2026-09'`
  - `shipping: summary` → all: `{total:10, orders:2}`; online + Aug: `{total:6, orders:1}`

- [ ] **Step 2: Reload the test page.** Expected: the new tests FAIL.

- [ ] **Step 3: Implement.** One query over `ve_sale_items vi JOIN ve_sales vs ON vs.transaction_no = vi.transaction_no`. Build the WHERE clause from the filters (`source` is skipped when `'both'` or empty; for event, `'__unassigned__'` → `vs.event_id IS NULL`, an id → `vs.event_id = ?`).
  - `product_ids`: a correlated subquery, `(SELECT GROUP_CONCAT(product_id) FROM (SELECT m.product_id FROM product_ve_mappings m WHERE <match key> ORDER BY m.product_id))`, split and parsed in JS.
  - `ignored`: `EXISTS (SELECT 1 FROM ve_ignored_items g WHERE <match key>)`.
  - Allocated discount uses the Global Constraints formula.
  - Shipping is a separate `SELECT COALESCE(SUM(shipping),0), SUM(CASE WHEN shipping>0 THEN 1 ELSE 0 END) FROM ve_sales` with the same filters.

- [ ] **Step 4: Reload.** Expected: all Task 1–2 tests PASS.

- [ ] **Step 5: Commit**: `feat(sales-data): filtered sales line query and shipping summary`

---

### Task 3: Pure computation: `computeSalesData` and `getLastMonthsRange`

**Files:**
- Modify: `js/utils.js` (append inside the `Utils` object, near the break-even math ~line 692)
- Test: `tests/sales-data.test.html`

**Interfaces:**
- Consumes: the `SalesLine[]` shape from Task 2; `products: Array<{id, name, cogs}>` (from `Database.getProducts()`)
- Produces:
  - `Utils.getLastMonthsRange(n:number, today:Date) -> {from:'YYYY-MM-DD', to:'YYYY-MM-DD'}`, covering the previous `n` full calendar months
  - `Utils.computeSalesData(lines, products) -> SalesData`, where:

```
SalesData = {
  weighted: { units, avgPrice, avgCogs, cm, cmPct, productCount } | null,  // null when 0 linked units
  byProduct: [{ productId, name, alsoNames:string[], units, netRevenue, avgPrice, cogs,
                margin, marginPct, profit, mix, missingCogs:boolean,
                cogsConflict: number[] | null }],   // sorted by units desc
  linked:   { amount, units },          // non-ignored linked lines (amount = pre-tax vi.amount)
  eligible: { amount, units },          // all non-ignored lines
  linkedPct: number | null,             // linked.amount / eligible.amount * 100, null if eligible.amount == 0
  unlinked: [ItemGroup], ignored: [ItemGroup],
  bestSellers: { [month:'YYYY-MM']: { units:[Rank], revenue:[Rank], margin:[Rank] } } // top 3 each, months ascending
}
ItemGroup = { name, price, qty, orders, firstDate, lastDate }   // grouped by name + price.toFixed(2), sorted by qty desc
Rank = { productId, name, value }
```

All money values are rounded to cents at output. `cmPct` and `marginPct` are percentages (61.5, not 0.615). Revenue = `amount − discount`. A product's margin $ for a month = `units × (monthNetRevenue/units − cogs)`.

- [ ] **Step 1: Write failing tests** (pure; build `lines` arrays inline):
  - `getLastMonthsRange: last 3` → `(3, new Date(2026,10,7))` → `{from:'2026-08-01', to:'2026-10-31'}`
  - `getLastMonthsRange: january rollover` → `(1, new Date(2027,0,15))` → `{from:'2026-12-01', to:'2026-12-31'}`
  - `weighted: two products` → product A (cogs 9.5): 168 units, revenue 4032. Product B (cogs 1.2): 121 units, revenue 726. Expect `units 289`, `avgPrice` 16.46 (4758/289), `avgCogs` 6.03 ((168×9.5+121×1.2)/289 = 1741.2/289 = 6.0249 → 6.02). Use `assertClose(…, 0.01)` against the computed literals 16.46 and 6.02, and `cm` = 10.44.
  - `discount reduces revenue` → one line `amount 50, discount 5, qty 2`, cogs 10 → `avgPrice 22.5`, `margin 12.5`
  - `multi-link: counted once, lowest id` → line with `product_ids [3,2]` (unsorted on purpose), qty 4 → byProduct has one row, `productId 2`, `units 4`, `alsoNames ['Sticker Rebrand']`
  - `multi-link: conflicting cogs flagged` → products 2 (cogs 1.2) and 3 (cogs 1.5) both linked → row uses cogs 1.2, `cogsConflict` deep-equals `[1.2, 1.5]`
  - `ignored wins over linked` → a linked line with `ignored:true` is not in byProduct, is in `ignored`, and is not in `eligible`
  - `linkedPct` → linked amount 90, unlinked amount 10, ignored amount 50 → `linkedPct === 90`
  - `empty linked set returns nulls` → only unlinked lines → `weighted === null`; all-ignored → `linkedPct === null`
  - `missing cogs flagged` → product cogs 0 → `missingCogs === true`, still counted
  - `item groups` → two lines `'Mystery'`, $0, in 2 orders on 2026-09-01 and 2026-09-12 → `unlinked[0]` = `{qty:2, orders:2, firstDate:'2026-09-01', lastDate:'2026-09-12'}`
  - `bestSellers top 3 per metric` → 4 products in 2026-09 with distinct units → `bestSellers['2026-09'].units.length === 3`, ordered desc. The revenue ranking differs from the units ranking in the fixture.

- [ ] **Step 2: Reload.** Expected: the Task 3 tests FAIL.
- [ ] **Step 3: Implement both functions** in `js/utils.js`. Attribution: `Math.min(...product_ids)`. `alsoNames` are the names of the other ids, deduped across all of that product's lines.
- [ ] **Step 4: Reload.** Expected: all tests PASS.
- [ ] **Step 5: Commit**: `feat(sales-data): weighted sales mix computation`

---

### Task 4: Date presets "Last month" / "Last 3 months"

**Files:**
- Modify: `js/app.js` `vePopulatePresetDropdown` (~12039) and `veApplyPreset` (~13113)

**Interfaces:**
- Consumes: `Utils.getLastMonthsRange`
- Produces: preset values `'last1'` and `'last3'`

- [ ] **Step 1: Add options.** Right after `<option value="all">All time</option>`, add `<option value="last1">Last month</option><option value="last3">Last 3 months</option>`. Add them **before** the early `return` so they always exist.
- [ ] **Step 2: Handle in `veApplyPreset`:** `/^last(\d)$/` → `Utils.getLastMonthsRange(n, new Date())` → set From/To → `veApplyFilters()`.
- [ ] **Step 3: Verify manually.** Run `npx serve .` (or open `index.html`), go to VE Sales, and pick "Last 3 months". Expected: From = `2026-08-01`, To = `2026-10-31` (on 2026-11-07), and the Sales subtab filters.
- [ ] **Step 4: Commit**: `feat(ve-sales): last month / last 3 months presets`

---

### Task 5: Sales Data subtab: layout and rendering

**Files:**
- Modify: `index.html`: add the button in `#veSubtabs` (line ~1262) and a `#veSubtabData` block after `#veSubtabEvents` (~1383), inside `#veDashboard`
- Modify: `js/app.js`: `veSwitchSubtab` (~13133), `veApplyFilters`/`veRenderAll` (~12131), new `veRenderSalesData()`, `veGetFilters()`
- Modify: `css/styles.css`: new `.sd-*` rules after `.ve-subtab` (~2909)

**Interfaces:**
- Consumes: `Database.getSalesDataLines`, `Database.getShippingSummary`, `Database.getProducts`, `Utils.computeSalesData`
- Produces:
  - `App.veGetFilters() -> {source, from, to, event}` (reads `#ve-filterSource/From/To/Event`)
  - `App.veRenderSalesData() -> void`, which caches its result on `this._salesData` (a `SalesData`)
  - DOM ids: `#sdLinkedPill`, `#sdUseInBreakevenBtn`, `#sdCardPrice`, `#sdCardCogs`, `#sdCardCm`, `#sdCardCmSub`, `#sdCardUnits`, `#sdCardUnitsSub`, `#sdMarginBody`, `#sdShipTotal`, `#sdShipAvg`, `#sdShipSub`, `#sdBestSellers`, `.sd-seg[data-sd-metric]`

- [ ] **Step 1: Markup.** Follow mockup v2's structure top to bottom:
  - fill-in panel `.sd-be-panel`: header with title and `#sdLinkedPill` (a button); four `.ve-card`s; footer with `#sdUseInBreakevenBtn.btn.btn-primary` and a hint
  - `.sd-grid2`: Product Margins & Mix `.ve-table-section` > `table.ve-table` with columns Product, Units, Mix, Avg Price, COGS, Margin, Margin %, Profit, using `data-sdsort` headers, plus a Shipping column holding two `.ve-card`s with value class `.ve-shipping`
  - Best Sellers `.ve-table-section`: the h3 holds a `.sd-seg` with Units/Revenue/Margin $ buttons; body `#sdBestSellers.ve-source-breakdown.sd-months`

  Card value classes: price `.ve-pretax`, COGS `.ve-tax`, CM `.ve-total`.
- [ ] **Step 2: CSS:**
  - `.sd-be-panel` (background `var(--color-accent-bg)`, `--radius`, padding 12px 14px, margin-bottom 12px)
  - `.sd-be-head` / `.sd-be-foot` (flex, space-between / gap 10px)
  - `.sd-pill` (amber: `--amber-bg`/`--amber`; `.sd-pill.ok` green: `--green-bg`/`--green`; radius 10px, `--font-xs`, cursor pointer)
  - `.sd-mixbar` (80×6px track `rgba(var(--color-primary-rgb),.14)`, inner fill `var(--c2,var(--primary))`)
  - `.sd-seg` (copy the `.ve-subtabs` look at `--font-xs`)
  - `.sd-grid2` (grid `1fr 220px`, gap 12px; single column at `max-width:768px`)
  - `.sd-months` (`grid-template-columns: repeat(auto-fill, minmax(200px,1fr))`)
  - `.sd-warn` (amber text), `.sd-also` (muted `--font-xs`)
- [ ] **Step 3: Logic.**
  - `veSwitchSubtab` also toggles `#veSubtabData` for `'data'` and calls `veRenderSalesData()`.
  - `veRenderAll` calls `veRenderSalesData()` when `#veSubtabData` is visible.
  - Rendering:
    - With `weighted === null`, cards show `—`, the button is `disabled`, and the units sub reads "Link products to see averages".
    - Pill text is `${linkedPct.toFixed(0)}% of sales linked · view excluded ›`. It gets `.ok` at 100 and reads "No sales in window" when `null`. Its `title` shows linked/eligible units.
    - Table rows:
      - Margin % below `weighted.cmPct` gets `.sd-warn`.
      - `missingCogs` shows COGS as `—` with `.ve-muted`.
      - `alsoNames` adds `<span class="sd-also">also: …</span>` under the name.
      - `cogsConflict` adds a ⚠ with `title="Linked products have different COGS: $1.20, $1.50"`.
      - The last row is `.ve-total-row` "Weighted total".
    - Sorting via `data-sdsort` (default units desc), stored on `this._sdSort`.
    - Best sellers: one `.ve-source-card` per month with `h4` "September 2026" (`toLocaleDateString` month long + year), up to 3 `.ve-source-stat` rows `1. Name` / value. Values are units as integers, revenue and margin via `veFmt`. If a month has none: "No linked sales". Metric stored on `this._sdMetric` (default `'units'`).
    - Shipping avg = `total/orders` or `—` when `orders === 0`.
  - Escape every name with `Utils.escapeHtml`.
- [ ] **Step 4: Verify manually** with `example_data` or the user's DB. Check:
  - The subtab switches.
  - Changing Source/From/To/Event re-renders.
  - The weighted-total row equals the fill-in cards.
  - Toggling the metric re-ranks.
  - Dark mode (gear → dark) and one other theme preset look right.
  - Width 768px stacks the margins and shipping.
  - The console is clean.
- [ ] **Step 5: Commit**: `feat(sales-data): Sales Data subtab with fill-in, margins, best sellers, shipping`

---

### Task 6: Use in Break-even and the Excluded Sales modal

**Files:**
- Modify: `index.html`: add `#veExcludedModal` and `#veLinkPickerModal` near the other VE/product modals (follow the `.modal` markup at `index.html:355`)
- Modify: `js/app.js`: handlers, event wiring in `setupEventListeners`, and the `hideIds` list in `applyViewOnlyRestrictions` (`js/app.js:11775`)

**Interfaces:**
- Consumes: `this._salesData`, `Database.getBreakevenConfig/setBreakevenConfig`, `this.pushUndo({type,label,undo,redo})`, `Database.addIgnoredItem/removeIgnoredItem/addProductVeMapping`, `Database.getProducts()`
- Produces: `App.sdUseInBreakeven()`, `App.sdOpenExcluded(view:'unlinked'|'ignored')`, `App.sdRenderExcluded()`, `App.sdOpenLinkPicker(name, price)`, `App.sdAfterExclusionChange()` (refreshes the modal if open, `veRenderSalesData()`, and `veRenderProducts()`)

- [ ] **Step 1: Use in Break-even.**
  - Snapshot the old `consumer` object, set `avgPrice`/`avgCogs` (rounded to 2dp), set `enabled = true`, then `setBreakevenConfig`.
  - Call `pushUndo` with `label: 'Use sales averages in Break-even'`, where `undo` restores the snapshot and `redo` reapplies.
  - Toast: `Break-even updated: price $18.42, COGS $7.10` (`UI.showNotification(…, 'success')`).
  - If the break-even tab is visible, call `refreshBreakeven()`.
- [ ] **Step 2: Excluded modal markup.**
  - Title "Excluded Sales".
  - A `.sd-seg` with `Unlinked (n)` / `Ignored (n)`.
  - `table.ve-table` columns: Item, Price, Qty, Orders, First, Last, plus an actions column.
  - Unlinked row actions: `Link…` and `Mark as test` (`.btn.btn-secondary.btn-sm`).
  - Ignored row action: `Restore`.
  - Empty state for Unlinked: "Every sale in this window is linked to a product."
  - Rows come from `this._salesData.unlinked` / `.ignored`.
- [ ] **Step 3: Link picker modal.**
  - Header "Link "<name>" ($price) to…".
  - A search input plus a list of non-discontinued products (name, price, COGS). The product whose price equals the item price is shown first, with a `.ve-linked-badge` "price match".
  - Click → `addProductVeMapping` → toast "Linked to <product>" → close → `sdAfterExclusionChange()`.
- [ ] **Step 4: Wire it up.**
  - `#sdLinkedPill` → `sdOpenExcluded('unlinked')`.
  - Mark as test → `addIgnoredItem` → toast "Moved to Ignored. Restore anytime from Ignored." → `sdAfterExclusionChange()`.
  - Restore → `removeIgnoredItem` → same refresh.
  - Use event delegation on the modal tbody with `data-name`/`data-price` attributes (escaped).
  - Add `sdUseInBreakevenBtn` to `hideIds`. In view-only mode, the modal renders with no actions column (check the existing view-only flag the app uses, e.g. `this._viewOnly` or `this.isViewOnly`, and reuse it).
- [ ] **Step 5: Verify manually.**
  - Pill opens the modal.
  - Mark an item as test: it moves to Ignored, the pill % goes up, and the averages change.
  - Restore it: everything is back as before.
  - Link an item: it leaves Unlinked and appears in the margins table.
  - Use in Break-even: the Break-even config modal shows the new values; Ctrl+Z (or the undo toast) restores the old ones.
  - An item named `<b>x</b>"` (insert one via the console: `Database.db.run(...)`) renders as literal text everywhere.
- [ ] **Step 6: Commit**: `feat(sales-data): use-in-break-even, excluded sales modal, link picker`

---

### Task 7: Sales subtab: Test Items section, Product Breakdown exclusion, docs

**Files:**
- Modify: `index.html`: new `.ve-table-section#veTestItemsSection` after the Transactions section inside `#veSubtabSales` (~line 1362)
- Modify: `js/app.js`: `veRenderProducts` (~12227), `veRenderTransactions` (~12339), new `veRenderTestItems()`, `veRenderAll`, `changelogData`
- Modify: `CLAUDE.md`

**Interfaces:**
- Consumes: `Database.getIgnoredItems()`, `Database.removeIgnoredItem`, `this._veFiltered`, `this._veItemsCache`, `App.sdAfterExclusionChange` (extend it to also call `veRenderTestItems()`)
- Produces: `App._veIgnoredKeys: Set<string>` with keys `name + '|' + price.toFixed(2)` (exact name, not lowercased, to match the SQL key), rebuilt at the start of `veRenderAll`

- [ ] **Step 1: Product Breakdown.** In `veRenderProducts`, skip items whose key is in `_veIgnoredKeys`. Summary cards, source breakdown, and journal creation stay untouched.
- [ ] **Step 2: Transactions.** In the Items cell, append `<span class="ve-inferred-badge sd-test-badge">test</span>` after an item that is ignored (reuse the amber badge style).
- [ ] **Step 3: Test Items section.**
  - Collapsible: the h3 "Test Items (n)" toggles the body, collapsed by default and remembered in `localStorage` key `veTestItemsOpen` (wrapped in try/catch).
  - Body: `table.ve-table` with Item, Price, Qty, Orders, First, Last, and Restore, built from the ignored lines within `_veFiltered`.
  - Hidden entirely when there are no ignored items in the window.
  - Restore button hidden in view-only mode.
- [ ] **Step 4: Docs.**
  - Add a "### Sales Data (VE Sales subtab)" section to `CLAUDE.md`'s Feature map, listing the new HTML ids, the app.js functions, the DB methods, and `Utils.computeSalesData`.
  - Add a changelog entry at the top of `changelogData`: `{ type:'feature', text:'Sales Data subtab: weighted avg price/COGS with "Use in Break-even", product margins & mix, best sellers by month, excluded/test item management' }`. Follow the existing release object shape.
- [ ] **Step 5: Verify manually.**
  - Mark an item as test on Sales Data, then switch to Sales: it's gone from Product Breakdown, listed under Test Items, and tagged "test" in Transactions.
  - Pretax and post-tax cards are unchanged.
  - Restore from Test Items brings it back in both subtabs.
  - Run all three `tests/*.test.html` pages: all PASS.
- [ ] **Step 6: Commit**: `feat(ve-sales): test items section and docs for Sales Data`
