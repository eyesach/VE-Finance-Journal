# Share Report Month — Design Spec

## Problem

When sharing a view-only snapshot, it reflects the current state of the data including the current month's actuals. Users want to share monthly reports (e.g., "March report") where the snapshot reflects the projected state as of a specific month, without April actuals contaminating the view.

## Solution

Add a "Report As Of" month picker to the share modal. When set, the exported database includes a `report_month` metadata value. The view-only recipient's experience is adjusted so that the data appears as if "today" were the last day of that month.

## Approach: Metadata + Display-Time Filtering

The report month is stored in `app_meta` inside the exported database. All data is preserved — no destructive modifications. Filtering happens at display and query time.

### Why this approach

- Non-destructive: full data preserved in the snapshot DB
- Existing tab month pickers (Dashboard, Balance Sheet, Break-Even) remain functional for exploration
- Simple implementation: one metadata value, targeted rendering checks
- Backwards compatible: shares without `report_month` behave exactly as before

## Share Modal UX

### Two-phase modal structure

The current share modal immediately uploads on open. This changes to a two-phase flow:

1. **Phase 1 — Configuration**: Modal opens showing the "Report As Of" dropdown and a "Generate Share Link" button. No upload yet.
2. **Phase 2 — Result**: After clicking "Generate," the modal shows QR code, share link, and report month confirmation.

This requires restructuring the share modal in `index.html` (add dropdown + generate button) and splitting `shareJournal()` in `app.js` into `openShareModal()` (show picker) and `generateShare()` (clone + upload).

### Month picker

- "Report As Of" dropdown appears at the top of phase 1
- Options: "Current (no filter)" (default) + past months from the timeline range
- Same month source as the existing Dashboard snapshot picker (`_getTimelineMonths()` filtered to `<= current month`)
- Helper text: "Transactions paid after this month appear as pending."

### Generate flow

1. User selects report month (or leaves on "Current")
2. Clicks "Generate Share Link"
3. If a report month is selected, app writes `report_month` to `app_meta` in a cloned DB
4. Cloned DB is exported, uploaded to Supabase, share URL + QR generated together
5. Modal switches to phase 2: QR code, link, and "Report: Mar 2026" confirmation in the expiry line

### "Current (no filter)" behavior

No `report_month` is written to `app_meta`. The share behaves exactly as it does today — full current-state snapshot.

## View-Only Side Behavior

### Banner

The view-only banner shows a pill badge: "Report: Mar 2026" next to the existing share info. Only appears when `report_month` is set.

### Per-tab behavior

| Tab | Report Month Effect |
|-----|-------------------|
| **Journal Entries** | Transactions where `entry_date` is after the report month are **hidden**. Transactions dated within the report month but with `month_paid` after the report month show as **"Pending"**. |
| **Cash Flow** | Post-process spreadsheet data: for months after report month, treat rows with `month_paid > report_month` as unpaid (projections only). See "Cash Flow / P&L rollback" section. |
| **P&L** | Same approach as Cash Flow — post-process to roll back actuals after report month. |
| **Balance Sheet** | Defaults to "As of {report month}" on load. Month picker remains functional for exploration. |
| **Dashboard KPIs** | Snapshot month picker defaults to report month on load. Picker remains functional. |
| **Break-Even** | Snapshot month picker defaults to report month on load. Picker remains functional. |
| **Projected Sales** | Unaffected — already projection-based data. |
| **Fixed Assets** | Unaffected — depreciation schedules are computed, not actuals. |
| **Loans** | Unaffected — amortization schedules are computed, not actuals. |
| **Budget** | Unaffected — budget tab shows planned schedules, not payment status. Budget actuals surface through the Journal tab rollback on `transactions` rows with `source_type = 'budget'`. |
| **VE Sales** | Unaffected — VE Sales tab renders `ve_sales` rows which have no payment status. Any linked journal entries are subject to the Journal tab rollback. |
| **B2B Contracts** | Unaffected — B2B contract tab has no payment status display. Linked journal entries are subject to the Journal tab rollback. |
| **Products** | Unaffected. |

### Payment status rollback rule

The `transactions` table stores payment status as `month_paid TEXT` (YYYY-MM format) and `status TEXT` ('pending'/'paid'). At display time, for any transaction: if `report_month` is set and `month_paid > report_month`, render the status as "Pending" instead of "Paid." The underlying `month_paid` value in the DB is not modified.

The comparison is simple string comparison: `"2026-04" > "2026-03"` → rolled back. No end-of-month date calculation needed for this check.

### Journal hiding rule

At query/filter time: if `report_month` is set, exclude transactions where `substr(entry_date, 1, 7) > report_month`. This is the only tab that hides entries — all other tabs show entries with status rollback only.

### Cash Flow rollback

Cash Flow data is pre-aggregated in SQL queries (`getCashFlowSpreadsheet`). The query filters on `status != 'pending'` and `month_paid IS NOT NULL`.

**Recommended: Pass `reportMonth` to DB query method.** Add an optional `reportMonth` parameter to `getCashFlowSpreadsheet()`. When provided, add `AND (month_paid <= ?)` to the WHERE clause for actual rows. This ensures rows paid after the report month are excluded from actuals aggregation and fall back to their projected/pending amounts.

### P&L rollback

P&L uses **accrual basis** — it queries on `month_due`, not `month_paid`. Revenue, COGS, and OpEx are placed in the month they're due regardless of payment status. This means P&L doesn't need `month_paid` filtering.

For P&L with a report month set: months through the report month display as-is (accrual data is correct regardless of payment timing). Months after the report month should still show projected amounts. Since P&L already shows projections for future months, the main effect is that the timeline/column display should treat the report month as the "current" month boundary — handled by the existing timeline filtering in `refreshPnL()` when `_reportMonth` is set.

## Data Layer

### Storage

- Key: `report_month` in `app_meta` table
- Value: `YYYY-MM` string (e.g., `2026-03`), or absent for "Current"
- DB methods: `getReportMonth()` and `setReportMonth(month)` in `database.js`

### Share-time flow

1. User selects month in share modal
2. Clone the current database:
   - `const bytes = Database.db.export()`
   - `const clone = new Database.SQL.Database(bytes)`
3. Write `report_month` to `app_meta` in the clone
4. Export the clone: `const blob = new Uint8Array(clone.export())`
5. **Close the clone to free WASM memory**: `clone.close()`
6. Upload blob to Supabase, generate share URL + QR
7. Original database is untouched

### View-only load flow

**Critical ordering**: `_reportMonth` and tab defaults must be set BEFORE `refreshAll()` is called.

1. `initViewOnlyMode()` loads the shared DB
2. **Before `refreshAll()`**: read `report_month` from `app_meta`
3. If present, set `this._reportMonth = value`
4. Set tab defaults:
   - `this._dashSnapshotMonth = reportMonth`
   - Balance Sheet month/year selects set to report month
   - Break-Even snapshot month set to report month
5. Call `refreshAll()` — all renders now check `this._reportMonth` for status rollback and journal filtering

### Helper function

```javascript
// Returns true if a month_paid should be treated as "not yet paid"
// given the active report month. Both values are YYYY-MM strings.
isPaymentRolledBack(monthPaid, reportMonth) {
    if (!reportMonth || !monthPaid) return false;
    return monthPaid > reportMonth;
}
```

For the Journal hiding rule, a separate check uses `entry_date` (YYYY-MM-DD):
```javascript
// Returns true if a transaction should be hidden from Journal tab
isEntryAfterReportMonth(entryDate, reportMonth) {
    if (!reportMonth || !entryDate) return false;
    return entryDate.substring(0, 7) > reportMonth;
}
```

## Files Modified

| File | Changes |
|------|---------|
| `index.html` | Restructure share modal: add "Report As Of" dropdown, add "Generate Share Link" button, two-phase layout (config → result) |
| `js/database.js` | Add `getReportMonth()`, `setReportMonth()`. Add optional `reportMonth` param to `getCashFlowSpreadsheet()` to exclude rows with `month_paid > reportMonth` from actuals. P&L is accrual-based (`month_due`) and does not need this param. |
| `js/app.js` | Split `shareJournal()` into `openShareModal()` + `generateShare()`. Clone DB + write report_month before export. View-only load: read report_month BEFORE `refreshAll()`, set tab defaults. Journal filter: hide entries where `entry_date` month > report_month. |
| `js/ui.js` | Transaction rendering: check `month_paid > reportMonth` for payment status rollback display |

## Backwards Compatibility

- Shares created before this feature have no `report_month` in `app_meta` → `getReportMonth()` returns null → all behavior unchanged
- "Current (no filter)" option produces the same result — no metadata written
- DB query methods with optional `reportMonth` param: when null/omitted, queries behave identically to current code
- No schema migration needed — `app_meta` is a key-value table, new keys are simply absent in old snapshots
