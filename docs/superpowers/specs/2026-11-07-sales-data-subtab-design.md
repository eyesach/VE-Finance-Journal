# Sales Data Subtab — Design

**Date:** 2026-11-07
**Status:** Reviewed. Awaiting user approval.
**Mockup:** `.superpowers/brainstorm/153-1794081169/content/sales-data-layout-v2.html`

## Intent

The user runs a virtual (simulated) business. Their last break-even analysis assumed every product sold equally and relied on hand-picked estimates. They want break-even inputs grounded in **actual sales performance**: a unit-weighted average selling price and COGS derived from imported VE sales and the existing product ↔ VE-item links.

This is the standard multi-product break-even method (weighted-average contribution margin by sales mix). Break-even's existing consumer channel already takes one average price and one average COGS. Feeding it unit-weighted values from real sales makes the existing calculation a correct multi-product break-even.

The user also wants to see the sales picture at a glance: best sellers each month and side-by-side product margins.

## Goals

1. Add a third subtab, **Sales Data**, to the VE Sales tab, next to Sales and Events.
2. Show a **break-even fill-in** panel with the unit-weighted average selling price, average COGS, contribution margin ($ and %), and linked unit count for the current filter window.
3. Add a **"Use in Break-even"** button that copies avg price and avg COGS into the Break-even consumer channel. This is a suggestion, never automatic.
4. Show a **"% of sales linked"** indicator. Clicking it opens an **Excluded Sales** modal listing unlinked line items. From each row the user can link the item to a product or mark it as test data.
5. Show a **Product Margins & Mix** table: units, mix %, avg price, COGS, margin $, margin %, total profit, plus a weighted-total row.
6. Show **Best Sellers by Month**: the top 3 products per month in the window, switchable between Units, Revenue, and Margin $.
7. Show **Shipping** on the side (collected total and average per shipped order), excluded from all price and margin math.

## Non-goals (this build)

- Mix-scenario / seasonal what-if editor, sensitivity analysis, and margin of safety. These are **Phase 2**, designed separately.
- Payment/processing fees. The user explicitly declined them, since the business is simulated.
- COGS history. A cost change is handled by creating a new product version (e.g. "Bundle v1", "Bundle v2"), per the user.
- Changing Break-even's calculation or UI beyond receiving the copied values.
- Any change to journal entries or the P&L.

## Anti-goals

- Break-even values must **never** change without the user clicking "Use in Break-even".
- Unlinked sales must not silently count as zero-cost units. They are excluded from the averages, and the % indicator makes the exclusion visible.
- Test-data exclusion must be reversible. Nothing is deleted from `ve_sales` or `ve_sale_items`.
- Must not look like a new design language. It must reuse the existing VE Sales styling (see Visual design).

## Definitions and math

All figures come from linked line items only, **pre-tax**, with **discounts allocated proportionally** per order. This matches the existing `Database.getLinkedProductAnalytics` approach: `vs.discount * (vi.amount / vs.subtotal)`.

The filter window is the existing VE filter bar: Source, From, To, and Event.

Per linked product *p* in the window:
- `units_p = Σ vi.quantity`
- `net_revenue_p = Σ (vi.amount − allocated_discount)`
- `avg_price_p = net_revenue_p / units_p`
- `cogs_p = products.cogs` (single static value)
- `margin_p = avg_price_p − cogs_p`; `margin_pct_p = margin_p / avg_price_p`
- `profit_p = units_p × margin_p`
- `mix_p = units_p / Σ units`

Weighted (break-even fill-in and the table's total row):
- `avg_price = Σ net_revenue_p / Σ units_p`
- `avg_cogs = Σ (units_p × cogs_p) / Σ units_p`
- `cm = avg_price − avg_cogs`; `cm_pct = cm / avg_price`

**% of sales linked** = pre-tax line-item amount of linked items ÷ pre-tax line-item amount of all **non-ignored** items in the window. Ignored (test) items are removed from both numerator and denominator, so 100% means every real sale is linked. *(Builder's call: revenue-based rather than unit-based, so a few $1 test units can't distort it; the tooltip shows the unit counts too.)*

Products with `cogs = 0` are still included. The table shows their COGS as "—" in a muted color, so a missing cost is visible.

## Behavior

### Subtab and filters
- A third `.ve-subtab` button, `data-ve-subtab="data"`, labelled **Sales Data**, plus a `#veSubtabData` content block. `veSwitchSubtab` handles the new value and renders the panel when activated.
- The Sales Data subtab reuses the shared `#veControls` filter bar. When filters change while Sales Data is active, it re-renders (hook into `veApplyFilters`).
- Add **"Last month"** and **"Last 3 months"** to `#ve-filterPreset`, under "All time". They mean the previous full calendar month(s) relative to today. For example, on 2026-11-07, "Last 3 months" is Aug 1 – Oct 31. `veApplyPreset` sets From/To accordingly. A custom range is just the existing From/To inputs.

### Break-even fill-in panel
- Four `.ve-card`s: Avg Selling Price, Avg COGS, Contribution Margin (sub: "% of price"), and Linked Units (sub: "across N products").
- The linked-sales pill sits in the panel header. It is amber below 100% and green at 100%.
- **Use in Break-even** reads `Database.getBreakevenConfig()`, sets `consumer.avgPrice` and `consumer.avgCogs` to the weighted values rounded to cents, and sets `consumer.enabled = true`. It saves with `setBreakevenConfig`, registers an undo via the existing `pushUndo` mechanism, and shows a success toast naming the values. If Break-even is rendered, it refreshes.
- With no linked sales in the window, the cards show "—", the button is disabled, and a hint explains that products need to be linked.

### Excluded Sales modal
- Two views, switched by a segmented control: **Unlinked** and **Ignored**.
- **Unlinked** rows are grouped by (item name, price) within the window. Columns: item name, price, total qty, # orders, first/last date. Actions:
  - **Link…** opens a compact product picker (active products from `Database.getProducts()`). Choosing one **adds** a `product_ve_mappings` row (product_id, name, price) without touching the product's other links. *Note:* `setMappingsForProduct` replaces all of a product's mappings, so a new add-single-mapping DB method is needed.
  - **Mark as test** adds (name, price) to the ignored list.
- **Ignored** rows show the same columns with a **Restore** action.
- After any action, the modal list, the pill, and the Sales Data panel all refresh.

### Ignored (test) items: storage
- New table `ve_ignored_items (id INTEGER PK, ve_item_name TEXT NOT NULL, ve_item_price REAL NOT NULL DEFAULT 0, created_at DATETIME DEFAULT CURRENT_TIMESTAMP, UNIQUE(ve_item_name, ve_item_price))`. It is created in both the schema and the migrations sections of `database.js`.
- Matching uses item name plus `ROUND(price, 2)`, the same key `product_ve_mappings` uses. *(Builder's call: ignoring is per name+price, not per individual line, consistent with how linking works.)*
- **Scope (user decision):** ignored items must never just disappear. They are kept visible in their own section, apart from the other sales.
  - **Sales Data subtab:** excluded from all averages, the table, best sellers, and the % denominator. They are listed under the modal's **Ignored** view.
  - **Sales subtab:** removed from the **Product Breakdown** table and shown in a new collapsible **Test Items** section below the Transactions table. That section lists name, price, qty, dates, and a **Restore** action, styled as a `.ve-table-section`.
  - *(Builder's call, flagged for user check:)* order-level figures stay unchanged: summary cards, source breakdown, the Transactions table, and journal-entry creation. That keeps the non-goal of not touching journal entries or the P&L. In the Transactions table, a line item marked as test gets a small amber "test" badge in the Items column.
- An item that is both linked and ignored is treated as **ignored**.

### Items linked to more than one product
- **User decision:** one VE item may legitimately be linked to several products. That happens when a product is rebranded; the products share the same price and COGS.
- Each sale line counts **once**. It must never be double-counted in units, revenue, mix, or the weighted averages.
- *(Builder's call:)* the line is attributed to the linked product with the **lowest `product_id`** (the original). In the margins table, that row gets a muted "also: <other names>" note.
- If the linked products' COGS differ (contrary to the rebrand assumption), the attributed product's COGS is used. The row also gets an amber warning icon with a tooltip naming the conflicting COGS values.
- Apply the same once-per-line rule anywhere the new queries join through `product_ve_mappings`. Changing the existing Products-tab analytics is out of scope.

### Product Margins & Mix table
- A `.ve-table` with blue header and sortable columns, following the existing `data-vesort` pattern. Default sort: units, descending.
- The Mix column shows a thin primary-tinted bar plus the %.
- Margin cells are green. Margin % below the weighted `cm_pct` is shown in amber.
- The last row is `.ve-total-row`, "Weighted total", with the weighted values.

### Best Sellers by Month
- One `.ve-source-card` per calendar month that overlaps the window, in chronological order. The grid wraps; long ranges just add rows.
- Each card shows the top 3 linked products for that month. If fewer were sold, it shows fewer. If none, it shows "No linked sales".
- A segmented toggle in the section header switches the ranking metric: Units (default), Revenue (net of allocated discount), or Margin $ (`units × (avg_price − cogs)` for that month).

### Shipping card
- Two purple `.ve-card`s: **Shipping Collected** = `Σ ve_sales.shipping` in the window, and **Avg per Order** = that total ÷ count of orders with `shipping > 0`.
- They respect the filter window and are reference only.

## Visual design

The design follows the approved mockup (v2) and reuses existing classes and tokens wherever possible.
- Subtabs: `.ve-subtabs` / `.ve-subtab`.
- Cards: `.ve-cards` / `.ve-card` / `.ve-card-label` / `.ve-card-value`, with color modifiers `.ve-pretax` (price), `.ve-tax` (COGS), `.ve-total` (margin), and `.ve-shipping`.
- Sections: `.ve-table-section` with an `h3` underlined in blue, and `.ve-count`.
- Tables: `.ve-table`, `.num` cells, `.ve-total-row`.
- Month cards: `.ve-source-breakdown` / `.ve-source-card` / `.ve-source-stat`. The grid uses `repeat(auto-fill, minmax(200px, 1fr))` instead of the fixed 2 columns.
- The fill-in panel follows the `.ve-journal-preview` look (`--color-accent-bg`, `--radius`).
- Buttons: `.btn .btn-primary`, and `.btn .btn-secondary .btn-sm` for row actions.
- Modal: the existing generic modal pattern (`UI.showModal` / `hideModal`).
- New CSS is limited to the fill-in panel header/footer, the linked pill, the mix bar, and a small segmented toggle. It uses only existing tokens (`--c2/--primary`, `--amber`, `--amber-bg`, `--green`, `--radius`, `--font-xs`, `--mono`) and must work in dark mode and all theme presets.
- Responsive: cards and month grids collapse via auto-fit/auto-fill. The margins/shipping two-column row stacks below 768px.

## Constraints

- Vanilla JS, matching the existing structure. SQL helpers live in `js/database.js` next to `getLinkedProductAnalytics`, which should be reused or extended rather than duplicated. Rendering goes in `js/ui.js` or alongside the existing VE renderers in `js/app.js`, following local convention. Event wiring goes in `setupEventListeners`.
- The new table must be part of the SQLite file, so it flows through save/load, multi-company, and sync automatically. Call `autoSave()` after writes.
- View-only (shared) mode: the Use in Break-even, Link, Mark as test, and Restore actions are hidden or disabled, consistent with `applyViewOnlyRestrictions`.
- Update `CLAUDE.md`'s code map and add a changelog entry (`renderChangelog`).

## Left to the builder

- Exact function names and file placement within the conventions above.
- Whether to extend `getLinkedProductAnalytics` with new filters (event, ignored items) or add a sibling query.
- Tests: follow whatever pattern the repo uses for the anomaly rule unit tests on this branch.
- Empty states and loading details, within the existing VE empty-state style.

## Phase 2 (not this build)

A mix-scenario editor that starts from the actual mix, lets the user adjust shares or add a seasonal product at an expected share (others scale proportionally), and shows the resulting WACM and break-even. It would also cover per-month mix (e.g. same months last year), sensitivity (±price, ±COGS, mix shift), and margin of safety.
