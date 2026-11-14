# Deposits & Other Assets Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a Deposits & Other Assets section (refundable security deposits) with date-aware, auto-classified (current vs. long-term) Balance Sheet presentation, a refund lifecycle, and one shared asset-totals calculation.

**Architecture:** New `deposits` table and CRUD in `database.js`, which auto-creates journal entries tagged with `source_type` `deposit_payment` / `deposit_refund` / `deposit_withheld`. A pure classifier lives in `utils.js`. One new `Database.getAssetTotalsAsOf(month)` replaces every hand-rolled cash + AR + fixed-asset sum in `app.js` / `ui.js`. The UI is a new section on the Fixed Assets tab plus two modals.

**Tech Stack:** Vanilla JS, sql.js (SQLite in browser), HTML test harness pages under `tests/`.

**Spec:** `docs/superpowers/specs/2026-11-14-deposits-other-assets-design.md`. Read it first. Its "Journal entries", "Balance Sheet rules", and "Migration" sections are binding.

## Global Constraints

- Category names exactly: `Security Deposits` (hidden from P&L: `show_on_pl = 1`, type `payable`) and `Deposit Forfeitures` (P&L-visible expense, `show_on_pl = 0`, type `payable`).
- source_type values exactly: `deposit_payment`, `deposit_refund`, `deposit_withheld`.
- Months are `YYYY-MM` strings and dates are `YYYY-MM-DD`. Compare months as strings, as the existing code does.
- Current means `expected_refund_date` month ≤ as-of month + 12 months. A missing expected date means long-term.
- Money is rounded to 2 decimals with the `round2` idiom already used in `app.js`.
- Every DB mutation calls `this.autoSave()`. Every UI mutation calls `this._guardViewOnly()` first.
- Reuse existing modal, button, pill, and form classes. Don't add a new design language.
- Test pages load sql.js 1.8.0 from cdnjs, the same as `tests/sales-data.test.html`.

## Review Focus

1. **Refund dated in the same month as payment.** The deposit is never "held" at month end, and the BS must still balance with cash net = refunded amount − withheld. → test in Task 3.
2. **Editing the amount of a deposit that's already refunded.** The refund entry must track the new amount, or else block the edit. Decision: block it. Editing a refunded deposit's amount or paid date shows an error saying "Undo refund first". → test in Task 1.
3. **Deposit saved with "Create journal entry" unchecked, then refunded and refund undone.** No crash, and no stray `deposit_payment` gets created. → test in Task 1.
4. **Expected refund date in the past, not refunded.** It stays current and shows an overdue marker, and never vanishes from the BS. → test in Task 2.
5. **Fixed asset purchased after the as-of month.** It's excluded from cost and accumulated depreciation, and Total Assets drops accordingly. → test in Task 3.

---

### Task 1: `deposits` table, CRUD, journal-entry lifecycle, snapshot/restore

**Files:**
- Modify: `js/database.js`: `createSchema()` (~line 70), `migrateSchema()` (~line 385), new `// ==================== DEPOSITS ====================` block after the fixed-asset block (after `getAssetDepreciationByMonth`)
- Create: `tests/deposits.test.html` (copy the harness header, `suite` / `test` / `assert*` helpers, and `freshDb()` from `tests/sales-data.test.html`)

**Interfaces:**
- Produces (all on `Database`):
  - `getDeposits() -> Array<DepositRow>`, ordered by `paid_date ASC, name ASC`
  - `getDepositById(id) -> DepositRow|null`
  - `addDeposit({name, amount, paid_date, expected_refund_date=null, notes=null, createEntry=true}) -> number` (new id)
  - `updateDeposit(id, {name, amount, paid_date, expected_refund_date, notes}) -> void`. Throws `Error('Undo refund first')` if the deposit is refunded and `amount` or `paid_date` changed. Updates the linked `deposit_payment` row (amount, entry_date, month_due, month_paid, description) if one exists.
  - `deleteDeposit(id) -> void`. Deletes the row and all transactions with the three deposit source types and `source_id = id`.
  - `markDepositRefunded(id, {refund_date, refunded_amount, withheld_category_id=null}) -> void`. Throws if `refund_date < paid_date` or `refunded_amount` is outside `[0, amount]`. Creates the `deposit_refund` receivable for the **full amount**, status `received`. If `amount − refunded_amount > 0`, also creates a `deposit_withheld` payable, status `paid`, in `withheld_category_id` (or `getOrCreateDepositForfeituresCategory()` when null). Sets `refunded_date` / `refunded_amount`.
  - `undoDepositRefund(id) -> void`. Deletes the `deposit_refund` and `deposit_withheld` rows and nulls `refunded_date` / `refunded_amount`.
  - `getOrCreateSecurityDepositsCategory() -> number`, `getOrCreateDepositForfeituresCategory() -> number`
  - `snapshotDeposit(id) -> {deposit: DepositRow|null, transactions: Array<Object>}`
  - `restoreDeposit(snapshot) -> void`. Deletes any current row and deposit-linked transactions for that id, then re-inserts the snapshot rows with their original ids (reuse `restoreTransactionRows`).
- `DepositRow` = `{id, name, amount, paid_date, expected_refund_date, refunded_date, refunded_amount, notes, created_at}`. The columns are given in the spec's Data table.
- Descriptions: `Deposit: <name>`, `Deposit refund: <name>`, `Deposit withheld: <name>`.

- [ ] **Step 1: Write failing tests in `tests/deposits.test.html`, suite "Deposits CRUD & entries"**
  - `addDeposit creates paid deposit_payment in hidden Security Deposits category`: add `{name:'Lease', amount:2000, paid_date:'2026-03-01'}` and assert one transaction exists with `source_type='deposit_payment'`, `transaction_type='payable'`, `status='paid'`, `amount=2000`, `month_paid='2026-03'`, and a category named `Security Deposits` with `show_on_pl=1`.
  - `addDeposit with createEntry=false creates no transaction`
  - `updateDeposit syncs payment entry`: change amount to 2500 and assert the payment row amount is 2500.
  - `full refund creates received receivable for full amount, no withheld`: refund on `'2027-06-15'` for 2000 and assert a `deposit_refund` row (amount 2000, `status='received'`, `month_paid='2027-06'`), no `deposit_withheld` row, and `refunded_amount=2000`.
  - `partial refund creates withheld expense in Deposit Forfeitures`: refund 1700 of 2000 and assert a `deposit_withheld` row with amount 300, `status='paid'`, in a category named `Deposit Forfeitures` with `show_on_pl=0`.
  - `refund rejects date before paid_date and amount out of range`: each call throws.
  - `updateDeposit on refunded deposit with amount change throws 'Undo refund first'`, but changing only the name is allowed. *(Review Focus 2)*
  - `undoDepositRefund removes refund rows and clears fields`
  - `refund + undo on deposit without payment entry works, never creates deposit_payment` *(Review Focus 3)*
  - `deleteDeposit removes row and all three entry types`
  - `snapshot/restore round-trips a refunded deposit`: snapshot, delete, restore, and assert the deposit row and both transaction rows are back with the same ids.

- [ ] **Step 2: Open `tests/deposits.test.html` in a browser and confirm these tests fail**
  - Expected: every test in the suite is red with "is not a function".

- [ ] **Step 3: Implement the table in `createSchema()` and an identical `CREATE TABLE IF NOT EXISTS deposits` in `migrateSchema()`, then the functions above**
  - The category helpers copy the find-or-create shape of `getOrCreateShippingCategory()` (`database.js:~1129`) and call `addCategory(name, false, null, 'payable', null, showOnPl)`.
  - Transactions are inserted via `this.addTransaction({...source_type, source_id: depositId})`, the same as `_autoCreateAssetTransaction`.

- [ ] **Step 4: Reload the test page and confirm the "Deposits CRUD & entries" suite is all green**

- [ ] **Step 5: Commit**

```bash
git add js/database.js tests/deposits.test.html
git commit -m "feat(deposits): deposits table, CRUD and refund journal lifecycle"
```

---

### Task 2: Deposit classification (pure)

**Files:**
- Modify: `js/utils.js`, adding a function near the depreciation helpers
- Test: `tests/deposits.test.html`, new suite "Deposit classification"

**Interfaces:**
- Produces: `Utils.classifyDeposit(deposit, asOfMonth) -> {held: boolean, bucket: 'current'|'longterm'|null, overdue: boolean}`
  - `held` = `paid_date` month ≤ `asOfMonth` AND (`refunded_date` null OR `refunded_date` month > `asOfMonth`)
  - When not held, `bucket` = null and `overdue` = false.
  - `bucket` = `'current'` when `expected_refund_date` is set and its month ≤ (`asOfMonth` + 12 months). Otherwise `'longterm'`.
  - `overdue` = held AND `expected_refund_date` month < `asOfMonth`
  - Month arithmetic: add 12 to the year of the `YYYY-MM` string. No Date objects, to avoid timezone drift.

- [ ] **Step 1: Write failing tests**. Deposit `{paid_date:'2026-03-01', expected_refund_date:'2028-02-28'}`:
  - as of `'2026-02'` → `{held:false, bucket:null, overdue:false}`
  - as of `'2026-03'` → held, `'longterm'`
  - as of `'2027-02'` → held, `'current'` (exactly 12 months out)
  - as of `'2027-01'` → held, `'longterm'` (13 months out)
  - as of `'2028-05'`, not refunded → held, `'current'`, `overdue:true` *(Review Focus 4)*
  - with `refunded_date:'2028-03-10'`: as of `'2028-02'` → held, and as of `'2028-03'` → not held
  - `expected_refund_date:null`, as of `'2026-03'` → held, `'longterm'`

- [ ] **Step 2: Reload and confirm the new suite fails with "classifyDeposit is not a function"**
- [ ] **Step 3: Implement `Utils.classifyDeposit`**
- [ ] **Step 4: Reload and confirm green**
- [ ] **Step 5: Commit** `feat(deposits): current vs long-term classifier`

---

### Task 3: Shared `getAssetTotalsAsOf` with a date-aware fixed-asset filter

**Files:**
- Modify: `js/database.js`, a new function in the DEPOSITS block (it reads cash, AR, fixed assets, and deposits)
- Test: `tests/deposits.test.html`, new suite "Asset totals & balance"

**Interfaces:**
- Consumes: `Utils.classifyDeposit` (Task 2), `getDeposits` (Task 1), the existing `getCashAsOf`, `getAccountsReceivableAsOf`, `getFixedAssets`, and `Utils.computeDepreciationSchedule`
- Produces: `Database.getAssetTotalsAsOf(asOfMonth) -> AssetTotals`, shaped exactly as in the spec's "Shared calculation" block. Detail arrays:
  - `fixedAssetDetails[]`: the asset row plus `accum_depreciation` and `net_book_value` (the same shape `refreshBalanceSheet` builds today), **only assets with `purchase_date` month ≤ asOfMonth**
  - `currentDepositDetails[]` / `longTermDepositDetails[]`: the deposit row plus `overdue`

- [ ] **Step 1: Write failing tests**. Use a helper `bsDiff(m) = totalAssets − (getAccountsPayableAsOf(m) + getSalesTaxPayableAsOf(m) + getRetainedEarningsAsOf(m, 'passthrough'))` with tax mode set to `'passthrough'` and no overrides. Each test seeds a fresh DB with one received $10,000 receivable in a P&L-visible "Sales" category in `2026-01`, so cash and RE are non-zero.
  - `deposit paid 2026-03: absent from totals in 2026-02, present in 2026-03, bsDiff 0 in both`
  - `classification moves deposit between currentDeposits and longTermDeposits, totalAssets unchanged`: compare `'2027-01'` and `'2027-02'` with expected refund `2028-02-28`.
  - `full refund 2028-03: deposit gone, cash +2000 vs 2028-02, bsDiff 0`
  - `partial refund 1700/2000: cash +1700, bsDiff 0`
  - `refund in same month as payment: bsDiff 0, cash net change −300 for 1700/2000` *(Review Focus 1)*
  - `fixed asset purchased 2026-06 excluded in 2026-05 (cost and accum depr), included in 2026-06` *(Review Focus 5)*
  - `totalCurrentAssets = cash + ar + currentDeposits; totalAssets = totalCurrentAssets + netFixedAssets + totalOtherAssets`

- [ ] **Step 2: Reload and confirm the new suite fails**
- [ ] **Step 3: Implement `getAssetTotalsAsOf`**, moving the fixed-asset loop from `app.js` `refreshBalanceSheet` (~7594-7616) into it and adding the purchase-month filter
- [ ] **Step 4: Reload and confirm all three suites are green**
- [ ] **Step 5: Commit** `feat(deposits): shared date-aware asset totals`

---

### Task 4: Route every asset consumer through `getAssetTotalsAsOf`, and the Balance Sheet layout

**Files:**
- Modify: `js/app.js`:
  - `refreshBalanceSheet` (~7584-7680): replace the cash / ar / fixed-asset block and pass the totals into `bsData`
  - dashboard BS block (~3285-3330): `totalAssets`
  - ~3364: `currentAssets` for the current ratio
  - ~2537: `ca = cash + ar`
  - ~3071: dashboard BS chart
  - ~1874: working-capital trend
  - ~1409: cash runway context
- Modify: `js/ui.js`:
  - `renderBalanceSheet` (~1560-1600): new layout from the spec's "Balance Sheet layout" section
  - `renderFinancialRatios` (~1681): use `data.totalCurrentAssets`

**Interfaces:**
- Consumes: `Database.getAssetTotalsAsOf` (Task 3)
- Produces: `bsData` keeps every existing key and adds `currentDeposits`, `currentDepositDetails`, `longTermDeposits`, `longTermDepositDetails`, `totalCurrentAssets`, and `totalOtherAssets`. The existing `data-bs-key` attributes stay. Add `data-bs-key="totalOtherAssets"`.

- [ ] **Step 1: Before changing anything, run `grep -n "cash + ar\|cash + receivables\|netFixedAssets" js/app.js js/ui.js`**
  - Record every hit. After this task, each one must read from the totals object or be inside `getAssetTotalsAsOf`.
- [ ] **Step 2: Replace each site. Keep each site's existing variable names by destructuring from the totals object, so the code around it stays untouched.**
- [ ] **Step 3: Update `renderBalanceSheet` to the spec layout.**
  - Detail rows use the existing `bs-detail-indent` class, the same as the AR breakdown.
  - Overdue deposits get `<span class="bs-projected-label">Overdue</span>`, reusing the existing pill.
- [ ] **Step 4: Re-run the grep from Step 1.**
  - Expected: no `cash + ar` / `cash + receivables` outside `getAssetTotalsAsOf`.
- [ ] **Step 5: Manual check in the app (`index.html`):**
  - Add a fixed asset dated in a future month. The BS for the current month excludes it, and the dashboard working capital and BS current ratio match the BS Total Current Assets.
  - Reload `tests/*.test.html` and confirm all green.
- [ ] **Step 6: Commit** `refactor(bs): single asset-totals source; deposits on balance sheet`

---

### Task 5: Deposits panel, modals, controllers, undo, view-only

**Files:**
- Modify: `index.html`:
  - new `<section class="deposits-section">` between the assets section and the equity section (~line 790)
  - `#depositModal` and `#depositRefundModal` next to `fixedAssetModal`
  - delete confirmation via the existing confirm pattern used by `deleteAssetModal`
- Modify: `js/ui.js`: `renderDepositsSection(deposits, todayMonth)` next to `renderFixedAssetsTab`
- Modify: `js/app.js`:
  - `refreshFixedAssets()` also calls `UI.renderDepositsSection(Database.getDeposits(), Utils.getCurrentMonth())`
  - new controllers `openDepositModal(id=null)`, `handleSaveDeposit()`, `openDepositRefundModal(id)`, `handleSaveDepositRefund()`, `handleUndoDepositRefund(id)`, `confirmDeleteDeposit(id)`
  - listeners wired in `setupEventListeners`
- Modify: `css/styles.css`: only what the new section needs, scoped under `.deposits-section`

**Interfaces:**
- Consumes: the Task 1 functions and `Utils.classifyDeposit` (Task 2)
- Element ids: `addDepositBtn`, `depositsListPanel`, `depositModal`, `depositName`, `depositAmount`, `depositPaidDate`, `depositExpectedRefundDate`, `depositNotes`, `depositAutoTransaction`, `depositRefundModal`, `refundDate`, `refundAmount`, `refundWithheldCategory`, `refundWithheldRow`
- Row actions use `data-deposit-action="edit|refund|undo-refund|delete"` and `data-deposit-id`, handled by one delegated listener on `depositsListPanel`.

- [ ] **Step 1: Add the markup.**
  - The `depositAutoTransaction` checkbox is checked by default on add and hidden on edit. Under it, the hint text reads exactly: `Uncheck if this payment is already in your journal.`
  - `refundWithheldRow` is hidden unless `refundAmount` < the deposit amount.
- [ ] **Step 2: Implement `renderDepositsSection`.**
  - Each row shows name, amount, paid date, expected refund date, and a status pill: `Held · Current` / `Held · Long-term` / `Refunded` / `Overdue`.
  - Status comes from `classifyDeposit(d, todayMonth)` plus `refunded_date`.
  - The empty state reads: `No deposits yet. Add a rent or utility deposit you expect to get back.`
  - In view-only mode, actions and the add button are hidden.
- [ ] **Step 3: Implement the controllers.**
  - Validation matches the spec and the Task 1 function guards. Errors are shown with `UI.showNotification(msg, 'error')`.
  - The refund modal defaults the date to `Utils.getTodayDate()` and the amount to the full deposit amount. The category select lists P&L-visible payable categories, with `Deposit Forfeitures` preselected.
  - The delete confirm text says it also removes the deposit's journal entries.
- [ ] **Step 4: Add undo.** Each mutating controller takes `before = Database.snapshotDeposit(id)` (an empty snapshot for add), performs the action, takes `after`, then calls:

```js
this.pushUndo({ type: 'deposit-<action>', label: '<Verb> deposit "<name>"',
  undo: () => Database.restoreDeposit(before), redo: () => Database.restoreDeposit(after) });
```

  - For add, `undo` must delete the new deposit. `restoreDeposit` with `{deposit:null, transactions:[]}` plus the id must delete it, so give the snapshot an `id` field.
- [ ] **Step 5: Manual check:**
  - Add a deposit, then edit it.
  - Mark it refunded partially and check the journal shows the refund and withheld rows. Check the P&L shows only the withheld amount.
  - Undo refund, then use Ctrl+Z / redo across each action.
  - Delete the deposit and confirm its journal rows are gone.
  - Turn on view-only and confirm the controls are hidden.
  - The BS updates after each step and shows no imbalance banner in pass-through mode.
- [ ] **Step 6: Commit** `feat(deposits): deposits panel, refund modal, undo`

---

### Task 6: Company copy, section reset, docs

**Files:**
- Modify: `js/companies.js` copySection `'assets'` (~line 254): also `DELETE FROM deposits` and copy `name, amount, paid_date, expected_refund_date, refunded_date, refunded_amount, notes`
- Modify: `js/database.js` `resetTabData` `case 'assets'` (~line 3448): also delete `deposits` and transactions with `source_type IN ('deposit_payment','deposit_refund','deposit_withheld')`
- Modify: `CLAUDE.md`: add a "Deposits & Other Assets" entry under the Feature map (HTML / Logic / DB / UI / Tests), add `deposits` to the schema table list (19 → 20 tables), and note `getAssetTotalsAsOf` under Balance Sheet

- [ ] **Step 1: Add the test `resetTabData('assets') clears deposits and their entries` to `tests/deposits.test.html`, then confirm it fails**
- [ ] **Step 2: Implement the reset and copy changes, then confirm the test is green**
- [ ] **Step 3: Manual check:** copy the Assets section from one company to another and confirm the deposits appear
- [ ] **Step 4: Update `CLAUDE.md`**
- [ ] **Step 5: Commit** `chore(deposits): company copy, section reset, code map`
