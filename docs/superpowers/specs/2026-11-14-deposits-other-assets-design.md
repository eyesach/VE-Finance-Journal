# Deposits & Other Assets — Design

**Date:** 2026-11-14
**Status:** Draft for approval

## Intent & why

The user has a refundable rent security deposit. In their real spreadsheet it's an asset. In the app it's currently a **non-depreciable fixed asset**, which is the wrong classification. A refundable deposit isn't physical, doesn't depreciate, and is money the business gets back when the lease ends.

Standard practice: a refundable security deposit is not a lease payment and is not depreciated. It is shown as its own asset, as a **current asset** when it's expected back within 12 months, otherwise under **"Other Assets"** (non-current). Any portion the landlord keeps is an expense. Sources: AccountingCoach "Security deposit asset", Crunchafi/LeaseCrunch KB "Where should the security deposit be posted", Deloitte DART ASC 842 Roadmap §6.1.

The refund won't happen within this journal's current timeline. The user still wants the full lifecycle (pay, then refund) so the feature is correct when it does.

## Goals

1. A separate **Deposits & Other Assets** section where deposits live, apart from depreciating fixed assets.
2. **Date-aware Balance Sheet:** a deposit counts only from the month it was paid until the month it's refunded. The same date rule is applied to regular fixed assets, which currently count in every month regardless of purchase date.
3. **Automatic current vs. long-term classification** for each Balance Sheet month, based on the expected refund date.
4. **One shared total-assets / current-assets calculation** used by the Balance Sheet, dashboard KPIs, financial ratios, and analyze charts, so they never disagree.
5. **Mark refunded** action: records the cash coming back (full or partial) and expenses any withheld portion, keeping the Balance Sheet balanced.

## Non-goals

- Real intangibles (trademarks, software) and amortization. The section name leaves room for them later, but they aren't built now.
- A one-click "Move to Deposits" migration. The user will re-enter their deposit and delete the old fixed-asset row (see Migration).
- Dashboard reminders for upcoming refunds.
- Removing the non-depreciable option from Fixed Assets. Land legitimately uses it.
- A formal cash flow statement split into investing and operating sections. The app's Cash Flow tab is category-based and stays that way.

## Anti-goals (failure even if it "works")

- The Balance Sheet stops balancing in any month: before payment, while held, after refund, or after a partial refund.
- The Balance Sheet, dashboard working capital / current ratio, and BS ratios show different asset totals for the same month.
- Deposits show up as expenses on the P&L. Only a withheld portion may.
- Deleting or refunding a deposit silently rewrites past months' cash.

## Current state (verified in code)

- `balance_sheet_assets` table (`database.js:158`). Saving an asset calls `_autoCreateAssetTransaction` (`app.js:7964`). That creates a **paid payable** in an "Asset Purchases" category hidden from P&L (`show_on_pl = 1` means hidden; the semantics are inverted), with `source_type = 'asset_purchase'`.
- `deleteFixedAsset` deletes the linked transaction too (`database.js:2362`).
- Cash = received receivables − paid payables with `month_paid <= asOf` (`getCashAsOf`, `database.js:3885`).
- Fixed assets are summed at full cost for **every** month regardless of `purchase_date`: `refreshBalanceSheet` `app.js:7594-7617` and the dashboard `app.js:3317-3326`. Today that unbalances the BS for months before a purchase, because cash hasn't dropped yet but the asset is already counted.
- Total assets are computed at `app.js:3326` and `app.js:7617`. Current assets (`cash + ar`) are computed at `app.js:2537`, `app.js:3364`, `ui.js:1582`, and `ui.js:1681`.
- Section copy between companies (`companies.js:255`) and section reset (`database.js:3448`) handle `'assets'` by touching `balance_sheet_assets` only.
- Fixed Assets tab markup: `index.html:772-800` (assets section, then equity section).

## Design

### Data

New table `deposits` (created in schema plus a migration guard, following the existing `CREATE TABLE IF NOT EXISTS` pattern):

| column | type | notes |
|---|---|---|
| id | INTEGER PK | |
| name | TEXT NOT NULL | e.g. "Office lease deposit" |
| amount | DECIMAL(10,2) NOT NULL | amount paid |
| paid_date | DATE NOT NULL | |
| expected_refund_date | DATE NULL | usually the lease end; null means treated as long-term |
| refunded_date | DATE NULL | set by Mark Refunded |
| refunded_amount | DECIMAL(10,2) NULL | cash actually returned |
| notes | TEXT NULL | |
| created_at | DATETIME | default now |

*Builder's call (decided in design):* use a separate table rather than an `asset_class` column on `balance_sheet_assets`. That avoids carrying depreciation columns and keeps the depreciation code paths untouched.

### Journal entries (how cash and the P&L stay right)

All deposit entries use a **"Security Deposits"** category, found or created like "Asset Purchases", with `show_on_pl = 1` (hidden from P&L). Exception: the withheld entry.

| Event | Entry | source_type |
|---|---|---|
| Deposit saved (if "Create journal entry" is checked — default on, mirroring `assetAutoTransaction`) | payable, status `paid`, amount = `amount`, month_due = month_paid = paid month | `deposit_payment` |
| Mark refunded | receivable, status `received`, amount = **full deposit amount**, month_due = month_paid = refund month | `deposit_refund` |
| Mark refunded with withheld > 0 | payable, status `paid`, amount = `amount − refunded_amount`, refund month, in a **P&L-visible expense category** | `deposit_withheld` |

How the partial refund balances: the asset drops by the full amount. Cash goes up by the full amount, then down by the withheld portion, for a net change equal to `refunded_amount`. The withheld portion hits the P&L as an expense, which reduces retained earnings. Assets fall by the withheld amount and so does equity.

- The withheld category is picked in the Mark Refunded modal. It defaults to a found-or-created **"Deposit Forfeitures"** expense category. *(Builder's call; the user can rename or choose another.)*
- *(Builder's call)* The deposit modal has a **"Create journal entry"** checkbox, default checked on add and hidden on edit, the same as fixed assets. It exists for deposits whose cash already left through a manual journal entry. Mark Refunded and Undo refund must work whether or not a `deposit_payment` entry exists.
- Editing a deposit updates its `deposit_payment` entry (amount and dates) if one exists.
- Deleting a deposit deletes all of its linked entries (all three source types), matching `deleteFixedAsset`. The delete confirmation must say this removes its journal entries.
- **Undo refund:** a refunded deposit offers "Undo refund". It deletes the `deposit_refund` and `deposit_withheld` entries and clears `refunded_date` / `refunded_amount`.

### Balance Sheet rules (per as-of month `M`)

- **Fixed asset counts** when `purchase_date` month ≤ `M`. This is a new filter. Accumulated depreciation already respects months.
- **Deposit is held** when `paid_date` month ≤ `M` AND (`refunded_date` is null OR `refunded_date` month > `M`).
- **A held deposit is current** if `expected_refund_date` is set AND its month ≤ `M` + 12 months. Otherwise it's long-term. A deposit whose expected refund date has passed but isn't marked refunded stays **current**, shown with a small "overdue" marker.

### Shared calculation

New `Database.getAssetTotalsAsOf(asOfMonth)` returns:

```
{ cash, ar,
  currentDeposits, currentDepositDetails[],
  totalCurrentAssets,              // cash + ar + currentDeposits
  fixedAssetDetails[], totalFixedAssetCost, totalAccumDepr, netFixedAssets,
  longTermDeposits, longTermDepositDetails[],
  totalOtherAssets,                // longTermDeposits
  totalAssets }                    // totalCurrentAssets + netFixedAssets + totalOtherAssets
```

Every site listed under "Current state" (`app.js:2537, 3317-3326, 3364, 7594-7617`; `ui.js:1582, 1681`) and the additional current-asset sites found in review (`app.js:1409` cash runway context, `app.js:1874` working-capital trend, `app.js:3071` dashboard BS chart) switches to this function or to its returned fields. No asset math stays duplicated.

### Balance Sheet layout (`ui.js` renderBalanceSheet)

```
Current Assets
  Cash
  Accounts Receivable (+ category breakdown, unchanged)
  Security Deposits (current)      ← only if > 0, with per-deposit detail rows
  Total Current Assets
Fixed Assets                        ← only assets purchased by M
  ... Net Fixed Assets
Other Assets                        ← only if > 0
  Security Deposits (long-term) + detail rows
  Total Other Assets
Total Assets
```

### UI: Deposits & Other Assets panel

- New `<section class="deposits-section">` on the Fixed Assets tab, between the Fixed Assets section and the Stockholders' Equity section, using the same header pattern: an `<h3>` and a "+ Add Deposit" button.
- List rows: name, amount, paid date, expected refund date, status pill (Held · Current / Held · Long-term / Refunded / Overdue). Status is computed as of today.
- Row actions: Edit, Mark Refunded (or Undo refund), Delete.
- **Deposit modal:** name, amount, paid date, expected refund date (optional), notes. Validation: amount > 0, paid date required, and the expected refund date can't be before the paid date.
- **Mark Refunded modal:** refund date (default today, must be ≥ paid date), amount returned (default = full amount, 0 ≤ x ≤ amount), withheld category (shown only when amount returned < amount).
- Visual quality should match the existing assets section and the user's preference for polished UI. Reuse existing modal, form, button, and pill styles. No new design language.

### Housekeeping

- `companies.js` copySection `'assets'` also copies `deposits` rows. Linked journal entries are not copied, consistent with how fixed assets are copied today.
- `resetTabData('assets')` also deletes `deposits` and transactions with the three deposit source types.
- Undo/redo: fixed-asset saves do **not** push undo today. Deposits get new undo/redo actions for add, edit, delete, refund, and undo refund. They're written from scratch on the existing `pushUndo` + `snapshotTransactionWithChildren` / `restoreTransactionRows` pattern. Fixed assets stay as they are (out of scope).
- View-only mode: hide the add, edit, refund, and delete controls, like the other tabs.
- Update `CLAUDE.md` code map with the new table, functions, and section.

### Migration of the user's existing deposit

No automated migration. The user confirmed the payment shows in the journal under **Asset Purchases** and wasn't typed into the journal by hand. They aren't sure whether it was auto-linked to the asset.

Steps, in this order:
1. Delete the old non-depreciable fixed asset.
2. Check the journal for the "Purchase: …" entry under Asset Purchases.
   - **Gone** (it was auto-linked): add the deposit with **"Create journal entry" checked**.
   - **Still there** (not linked): either delete it and add the deposit with the box checked (preferred, because it moves the payment into "Security Deposits"), or keep it and add the deposit with the box **unchecked**.
3. Spot-check the Balance Sheet for the paid month and the month before. Cash and total assets should match their pre-migration values in the paid month and later. Months before the paid month will now be correct, because the old fixed asset had been counted there too early.

*(Builder's call)* The deposit modal shows a one-line hint under the checkbox: "Uncheck if this payment is already in your journal."

## Acceptance checks

*(Builder's call)* "Balances" means **this feature adds no imbalance**. The BS already has gaps from causes this feature doesn't touch: corporate-mode tax is deducted from retained earnings with no tax-payable line, P&L overrides diverge from journal entries, and fixed assets saved without a payment entry. Fixing those is out of scope. Balance checks run in pass-through tax mode with no P&L overrides. In other modes, the check is that the BS difference is the same before and after each deposit action, apart from the tax effect of a withheld expense.

1. A deposit paid in month P doesn't appear on the BS for months before P, and the BS balances in those months.
2. Viewing a month more than 12 months before the expected refund shows it under Other Assets. Viewing a month within 12 months shows it under Current Assets. Total Assets is the same either way.
3. Full refund in month R: the deposit is gone from the BS from month R, cash is up by the full amount, the P&L is unchanged, and the BS balances.
4. Partial refund (e.g. $2,000 deposit, $1,700 returned): cash nets +$1,700, the P&L shows a $300 expense in month R, and the BS balances.
5. Undo refund restores the pre-refund state exactly.
6. A fixed asset purchased in month P no longer counts in months before P.
7. Dashboard working capital / current ratio, BS ratios, and the BS itself show identical current and total asset figures for the same month.
8. Copy-section and reset for "assets" include deposits.

## Left to the builder

- Exact CSS, icons, and empty-state copy for the panel.
- Whether the detail rows on the BS sit under one "Security Deposits" line or are listed directly (keep it consistent with the AR breakdown style).
- Test approach. The repo has an HTML test harness pattern (`tests/sales-data.test.html`). A similar `tests/deposits.test.html` covering the classification rules and the balancing scenarios above is expected.
