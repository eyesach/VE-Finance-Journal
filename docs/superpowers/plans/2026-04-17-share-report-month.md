# Share Report Month Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a "Report As Of" month picker to the share modal so view-only snapshots show data as if today were the last day of the selected month.

**Architecture:** Store `report_month` in `app_meta` inside a cloned DB at share time. On the view-only side, read it before `refreshAll()` and use it to: hide future journal entries, roll back payment statuses, default tab pickers, and filter Cash Flow actuals via a DB query parameter.

**Tech Stack:** Vanilla HTML/CSS/JS, sql.js (SQLite in browser), Supabase for share storage.

**Spec:** `docs/superpowers/specs/2026-04-16-share-report-month-design.md`

---

### Task 1: Add `getReportMonth()` and `setReportMonth()` to database.js

**Files:**
- Modify: `js/database.js:3185-3224` (after `getTimeline`/`setTimeline` methods)

- [ ] **Step 1: Add the two methods after `setTimelineEnd()`**

In `js/database.js`, after the `setTimelineEnd` method (around line 3224), add:

```javascript
/**
 * Get the report month for view-only snapshots
 * @returns {string|null} 'YYYY-MM' or null
 */
getReportMonth() {
    const result = this.db.exec("SELECT value FROM app_meta WHERE key = 'report_month'");
    if (result.length > 0 && result[0].values.length > 0) return result[0].values[0][0];
    return null;
},

/**
 * Set the report month for view-only snapshots
 * @param {string|null} month - 'YYYY-MM' or null to clear
 */
setReportMonth(month) {
    if (month) {
        this.db.run("INSERT OR REPLACE INTO app_meta (key, value) VALUES ('report_month', ?)", [month]);
    } else {
        this.db.run("DELETE FROM app_meta WHERE key = 'report_month'");
    }
    this.autoSave();
},
```

- [ ] **Step 2: Verify by searching for the method names**

Run: `grep -n "getReportMonth\|setReportMonth" js/database.js`
Expected: Two method definitions found.

- [ ] **Step 3: Commit**

```bash
git add js/database.js
git commit -m "feat: add getReportMonth/setReportMonth to database.js"
```

---

### Task 2: Add `reportMonth` parameter to `getCashFlowSpreadsheet()`

**Files:**
- Modify: `js/database.js:1639-1663` (`getCashFlowSpreadsheet` method)

- [ ] **Step 1: Add optional `reportMonth` parameter and filter to both queries**

Change the method signature and add `AND (month_paid <= ?)` when `reportMonth` is provided. Replace the current `getCashFlowSpreadsheet()` method:

```javascript
getCashFlowSpreadsheet(reportMonth = null) {
    // Get all distinct months from month_paid (sorted ASC)
    const monthsParams = [];
    let monthsFilter = 'WHERE month_paid IS NOT NULL AND status != \'pending\'';
    if (reportMonth) {
        monthsFilter += ' AND month_paid <= ?';
        monthsParams.push(reportMonth);
    }
    const monthsResult = this.db.exec(
        `SELECT DISTINCT month_paid as month FROM transactions ${monthsFilter} ORDER BY month ASC`,
        monthsParams
    );
    const months = monthsResult.length > 0 ? monthsResult[0].values.map(r => r[0]) : [];

    // Get per-category, per-month totals for completed transactions
    const dataParams = [];
    let dataFilter = 'WHERE t.status != \'pending\' AND t.month_paid IS NOT NULL';
    if (reportMonth) {
        dataFilter += ' AND t.month_paid <= ?';
        dataParams.push(reportMonth);
    }
    const dataResult = this.db.exec(`
        SELECT c.name as category_name, c.id as category_id,
               c.is_b2b, c.is_cogs,
               t.transaction_type, t.month_paid as month,
               SUM(t.amount) as total
        FROM transactions t
        JOIN categories c ON t.category_id = c.id
        ${dataFilter}
        GROUP BY c.id, t.month_paid, t.transaction_type
        ORDER BY c.cashflow_sort_order ASC, c.name ASC
    `, dataParams);
    const data = dataResult.length > 0 ? this.rowsToObjects(dataResult[0]) : [];

    return { months, data };
},
```

- [ ] **Step 2: Verify the existing call in `app.js` still works (no args = same behavior)**

Run: `grep -n "getCashFlowSpreadsheet" js/app.js`
Expected: Line ~285 shows `Database.getCashFlowSpreadsheet()` — no args, so `reportMonth` defaults to `null`, no filter applied. Behavior unchanged.

- [ ] **Step 3: Commit**

```bash
git add js/database.js
git commit -m "feat: add optional reportMonth filter to getCashFlowSpreadsheet"
```

---

### Task 3: Restructure share modal HTML (two-phase layout)

**Files:**
- Modify: `index.html:2712-2731` (shareModal)

- [ ] **Step 1: Replace the share modal with two-phase layout**

Replace the current share modal (lines 2712-2731) with:

```html
<!-- Share View-Only Link Modal -->
<div id="shareModal" class="modal">
    <div class="modal-content">
        <h3>Share View-Only Link</h3>
        <!-- Phase 1: Configuration -->
        <div id="shareConfigSection">
            <div class="form-group">
                <label for="shareReportMonth">Report As Of</label>
                <select id="shareReportMonth" class="form-control">
                    <option value="">Current (no filter)</option>
                </select>
                <p class="share-report-hint">Transactions paid after this month appear as pending.</p>
            </div>
            <div class="form-actions">
                <button type="button" id="closeShareBtn" class="btn btn-secondary">Cancel</button>
                <button type="button" id="generateShareBtn" class="btn btn-primary">Generate Share Link</button>
            </div>
        </div>
        <!-- Phase 2: Result -->
        <div id="shareResultSection" style="display:none;">
            <p id="shareStatus" class="share-status"></p>
            <div id="shareUrlSection" style="display:none;">
                <div id="shareReportBadge" class="share-report-badge" style="display:none;"></div>
                <div class="form-group">
                    <label>Share URL</label>
                    <div class="share-url-row">
                        <input type="text" id="shareUrlDisplay" class="form-control share-url-input" readonly>
                        <button type="button" id="copyShareUrlBtn" class="btn btn-small">Copy</button>
                    </div>
                </div>
                <div id="shareQrCode" class="share-qr-container"></div>
                <p id="shareExpiry" class="share-expiry"></p>
            </div>
            <div class="form-actions">
                <button type="button" id="closeShareResultBtn" class="btn btn-secondary">Close</button>
            </div>
        </div>
    </div>
</div>
```

- [ ] **Step 2: Add CSS for the new share modal elements**

In `css/styles.css`, find the `.share-expiry` rule (line ~3175) and add after it:

```css
.share-report-hint { font-size: var(--font-xs); color: var(--text-muted); margin-top: 2px; }
.share-report-badge { text-align: center; margin-bottom: 12px; font-size: var(--font-sm); font-weight: 600; color: var(--c2, var(--primary)); background: var(--color-accent-bg); padding: 6px 14px; border-radius: var(--radius); }
```

- [ ] **Step 3: Commit**

```bash
git add index.html css/styles.css
git commit -m "feat: restructure share modal with two-phase layout and report month picker"
```

---

### Task 4: Implement `openShareModal()` and `generateShare()` in app.js

**Files:**
- Modify: `js/app.js:10383-10427` (replace `shareJournal()`)
- Modify: `js/app.js:5258-5268` (share button event listeners)

- [ ] **Step 1: Replace `shareJournal()` with `openShareModal()` and `generateShare()`**

Replace the `shareJournal()` method (lines 10383-10427) with:

```javascript
openShareModal() {
    const supaConfig = Database.getSupabaseConfig();
    if (!supaConfig || !supaConfig.url || !supaConfig.anonKey) {
        UI.showNotification('Sharing requires Supabase. Set up Group Sync first.', 'error');
        return;
    }

    // Reset to phase 1
    document.getElementById('shareConfigSection').style.display = '';
    document.getElementById('shareResultSection').style.display = 'none';
    document.getElementById('shareUrlSection').style.display = 'none';
    document.getElementById('shareStatus').textContent = '';

    // Populate report month dropdown
    const select = document.getElementById('shareReportMonth');
    const realCurrent = Utils.getCurrentMonth();
    const allMonths = this._getTimelineMonths().filter(m => m <= realCurrent);
    select.innerHTML = '<option value="">Current (no filter)</option>' +
        allMonths.map(m => `<option value="${m}">${Utils.formatMonthShort(m)}</option>`).join('');
    select.value = '';

    UI.showModal('shareModal');
},

async generateShare() {
    const supaConfig = Database.getSupabaseConfig();
    if (!SupabaseAdapter.isInitialized()) {
        SupabaseAdapter.init(supaConfig.url, supaConfig.anonKey);
    }

    const reportMonth = document.getElementById('shareReportMonth').value || null;

    // Switch to phase 2 with loading state
    document.getElementById('shareConfigSection').style.display = 'none';
    document.getElementById('shareResultSection').style.display = '';
    document.getElementById('shareStatus').textContent = 'Uploading snapshot...';
    document.getElementById('shareUrlSection').style.display = 'none';

    try {
        // Clone DB and inject report_month if set
        let blob;
        if (reportMonth) {
            const bytes = Database.db.export();
            const clone = new Database.SQL.Database(new Uint8Array(bytes));
            clone.run("INSERT OR REPLACE INTO app_meta (key, value) VALUES ('report_month', ?)", [reportMonth]);
            blob = new Uint8Array(clone.export());
            clone.close();
        } else {
            blob = new Uint8Array(Database.db.export());
        }

        const journalName = Database.getJournalOwner() || 'Accounting Journal';
        const syncConfig = Database.getSyncConfig();
        const createdBy = (syncConfig && syncConfig.userName) || journalName;

        const result = await SupabaseAdapter.createShare(blob, createdBy, journalName);

        // Build share URL
        const shareToken = btoa(JSON.stringify({
            s: result.shareId,
            u: supaConfig.url,
            k: supaConfig.anonKey
        }));
        const shareUrl = window.location.origin + window.location.pathname + '#share=' + shareToken;

        // Display results
        document.getElementById('shareUrlDisplay').value = shareUrl;
        document.getElementById('shareStatus').textContent = '';
        document.getElementById('shareUrlSection').style.display = '';

        // Show report month badge if set
        const badge = document.getElementById('shareReportBadge');
        if (reportMonth) {
            badge.textContent = 'Report: ' + Utils.formatMonthShort(reportMonth);
            badge.style.display = '';
        } else {
            badge.style.display = 'none';
        }

        document.getElementById('shareExpiry').textContent =
            'Expires: ' + new Date(result.expiresAt).toLocaleDateString();

        this.generateShareQR(shareUrl);
    } catch (err) {
        console.error('Share failed:', err);
        document.getElementById('shareStatus').textContent = 'Share failed: ' + err.message;
    }
},
```

- [ ] **Step 2: Update the event listeners for the share button**

Find the share button event listeners (around line 5258-5268) and replace:

```javascript
// Share view-only link
document.getElementById('shareBtn').addEventListener('click', () => {
    this.openShareModal();
});
document.getElementById('generateShareBtn').addEventListener('click', () => {
    this.generateShare();
});
document.getElementById('closeShareBtn').addEventListener('click', () => {
    UI.hideModal('shareModal');
});
document.getElementById('closeShareResultBtn').addEventListener('click', () => {
    UI.hideModal('shareModal');
});
document.getElementById('copyShareUrlBtn').addEventListener('click', () => {
    const url = document.getElementById('shareUrlDisplay').value;
    if (url) this.copyToClipboard(url);
});
```

- [ ] **Step 3: Verify no remaining references to `shareJournal`**

Run: `grep -rn "shareJournal" js/`
Expected: No matches (all references replaced).

- [ ] **Step 4: Commit**

```bash
git add js/app.js
git commit -m "feat: split shareJournal into openShareModal + generateShare with DB clone"
```

---

### Task 5: Read `report_month` in `initViewOnlyMode()` and set tab defaults

**Files:**
- Modify: `js/app.js:10484-10509` (inside `initViewOnlyMode`, between DB load and `refreshAll()`)

- [ ] **Step 1: Add report month reading and tab defaults before `refreshAll()`**

In `initViewOnlyMode()`, find these lines (around line 10507-10509):

```javascript
            this.loadAndApplyTimeline();
            this.initBalanceSheetDate();
            this.refreshAll();
```

Replace them with:

```javascript
            this.loadAndApplyTimeline();

            // Read report month from snapshot and set tab defaults BEFORE refreshAll
            const reportMonth = Database.getReportMonth();
            if (reportMonth) {
                this._reportMonth = reportMonth;
                this._dashSnapshotMonth = reportMonth;
                // Balance Sheet defaults will be set after initBalanceSheetDate
            }

            this.initBalanceSheetDate();

            // Override Balance Sheet date if report month is set
            if (reportMonth) {
                const [rmYear, rmMonth] = reportMonth.split('-');
                document.getElementById('bsMonthMonth').value = rmMonth;
                document.getElementById('bsMonthYear').value = rmYear;
            }

            this.refreshAll();
```

- [ ] **Step 2: Commit**

```bash
git add js/app.js
git commit -m "feat: read report_month in view-only mode and set tab defaults before refreshAll"
```

---

### Task 6: Update view-only banner to show report month pill

**Files:**
- Modify: `js/app.js:10559-10576` (`showViewOnlyBanner` method)
- Modify: `css/styles.css:3164-3165` (banner styles)

- [ ] **Step 1: Add report month pill to the banner**

Replace the `showViewOnlyBanner` method with:

```javascript
showViewOnlyBanner(shareMeta) {
    const banner = document.createElement('div');
    banner.className = 'view-only-banner';

    const createdDate = new Date(shareMeta.created_at).toLocaleDateString();
    const createdBy = shareMeta.created_by || 'Unknown';
    const journalName = shareMeta.journal_name || 'Journal';

    let reportPill = '';
    if (this._reportMonth) {
        reportPill = '<span class="view-only-report-pill">Report: ' +
            Utils.formatMonthShort(this._reportMonth) + '</span>';
    }

    banner.innerHTML =
        '<div class="view-only-banner-content">' +
        '<span>&#128274;</span> ' +
        '<span><strong>View-Only Snapshot</strong> &mdash; ' +
        this._escapeHtml(journalName) + '</span>' +
        reportPill +
        '<span class="view-only-banner-meta">Shared by ' +
        this._escapeHtml(createdBy) + ' on ' + createdDate + '</span>' +
        '</div>';

    document.body.insertBefore(banner, document.body.firstChild);
},
```

- [ ] **Step 2: Add CSS for the report pill**

In `css/styles.css`, after the `.view-only-banner-content` rule (line 3165), add:

```css
.view-only-report-pill { background: rgba(var(--color-primary-rgb), 0.15); color: var(--c2, var(--primary)); padding: 2px 10px; border-radius: 12px; font-weight: 600; font-size: var(--font-xs); }
.view-only-banner-meta { color: var(--text-muted); font-size: var(--font-xs); }
```

- [ ] **Step 3: Commit**

```bash
git add js/app.js css/styles.css
git commit -m "feat: show report month pill badge in view-only banner"
```

---

### Task 7: Hide journal entries after report month

**Files:**
- Modify: `js/app.js:262-271` (`refreshTransactions` method)

- [ ] **Step 1: Filter out transactions after report month**

In `refreshTransactions()`, after the line `const transactions = Database.getTransactions(filters);` (line 264), add the report month filter:

```javascript
    refreshTransactions() {
        const filters = this.getActiveFilters();
        let transactions = Database.getTransactions(filters);

        // In view-only mode with report month, hide entries dated after the report month
        if (this._reportMonth) {
            transactions = transactions.filter(t =>
                t.entry_date.substring(0, 7) <= this._reportMonth
            );
        }

        UI.renderTransactions(transactions, this.currentSortMode);

        const allTransactions = Database.getTransactions();
        const months = Utils.getUniqueMonths(allTransactions);
        UI.populateFilterMonths(months);
    },
```

Note: Change `const transactions` to `let transactions` since we may reassign it.

- [ ] **Step 2: Commit**

```bash
git add js/app.js
git commit -m "feat: hide journal entries after report month in view-only mode"
```

---

### Task 8: Roll back payment status in transaction rendering

**Files:**
- Modify: `js/ui.js:385-404` (`renderTransactionRow` method)

- [ ] **Step 1: Add payment status rollback at the top of `renderTransactionRow`**

In `renderTransactionRow(t)`, right after the opening of the method (line 385), add rollback logic before the existing status checks. Find these lines:

```javascript
renderTransactionRow(t) {
    const isOverdue = Utils.isOverdue(t.month_due, t.status);
    const statusClass = isOverdue ? 'status-overdue' : `status-${t.status}`;
```

Replace with:

```javascript
renderTransactionRow(t) {
    // Report month rollback: treat payments after report month as pending
    let effectiveStatus = t.status;
    if (App._reportMonth && t.month_paid && t.month_paid > App._reportMonth) {
        effectiveStatus = 'pending';
    }

    const isOverdue = Utils.isOverdue(t.month_due, effectiveStatus);
    const statusClass = isOverdue ? 'status-overdue' : `status-${effectiveStatus}`;
```

- [ ] **Step 2: Update the status dropdown to use `effectiveStatus`**

Find the status dropdown section (around line 396-404):

```javascript
    const statusOptions = t.transaction_type === 'receivable'
        ? ['pending', 'received']
        : ['pending', 'paid'];

    const statusDropdown = `
        <select class="status-select ${statusClass}" data-id="${t.id}">
            ${statusOptions.map(s => `
                <option value="${s}" ${t.status === s ? 'selected' : ''}>
                    ${this.capitalizeFirst(s)}
                </option>
            `).join('')}
        </select>
    `;
```

Replace `${t.status === s ? 'selected' : ''}` with `${effectiveStatus === s ? 'selected' : ''}`:

```javascript
    const statusOptions = t.transaction_type === 'receivable'
        ? ['pending', 'received']
        : ['pending', 'paid'];

    const statusDropdown = `
        <select class="status-select ${statusClass}" data-id="${t.id}">
            ${statusOptions.map(s => `
                <option value="${s}" ${effectiveStatus === s ? 'selected' : ''}>
                    ${this.capitalizeFirst(s)}
                </option>
            `).join('')}
        </select>
    `;
```

- [ ] **Step 3: Update the late payment info to use `effectiveStatus`**

Find the late payment check (around line 407):

```javascript
    const isPaidLate = Utils.isPaidLate(t.month_due, t.month_paid);
    const lateInfo = isPaidLate
        ? `<span class="late-info">in ${Utils.formatMonthShort(t.month_paid)}</span>`
        : '';
```

Replace with:

```javascript
    const isPaidLate = (effectiveStatus !== 'pending') && Utils.isPaidLate(t.month_due, t.month_paid);
    const lateInfo = isPaidLate
        ? `<span class="late-info">in ${Utils.formatMonthShort(t.month_paid)}</span>`
        : '';
```

- [ ] **Step 4: Commit**

```bash
git add js/ui.js
git commit -m "feat: roll back payment status display when report month is active"
```

---

### Task 9: Pass `reportMonth` to Cash Flow and adjust P&L current month

**Files:**
- Modify: `js/app.js:284-346` (`refreshCashFlow` method)
- Modify: `js/app.js:879-946` (`refreshPnL` method)

- [ ] **Step 1: Pass `_reportMonth` to `getCashFlowSpreadsheet`**

In `refreshCashFlow()` (line 285), change:

```javascript
const data = Database.getCashFlowSpreadsheet();
```

to:

```javascript
const data = Database.getCashFlowSpreadsheet(this._reportMonth || null);
```

- [ ] **Step 2: Adjust P&L current month boundary when report month is set**

In `refreshPnL()` (around line 884), change:

```javascript
const currentMonth = Utils.getCurrentMonth();
```

to:

```javascript
const currentMonth = this._reportMonth || Utils.getCurrentMonth();
```

This makes P&L treat the report month as the "current" boundary — months through report month show actuals, months after show projections. The existing projection logic already uses `currentMonth` for this boundary.

- [ ] **Step 3: Do the same for Cash Flow's current month reference**

In `refreshCashFlow()` (around line 287), change:

```javascript
const currentMonth = Utils.getCurrentMonth();
```

to:

```javascript
const currentMonth = this._reportMonth || Utils.getCurrentMonth();
```

- [ ] **Step 4: Commit**

```bash
git add js/app.js
git commit -m "feat: pass reportMonth to Cash Flow query and adjust P&L current month boundary"
```

---

### Task 10: Default Break-Even snapshot month to report month

**Files:**
- Modify: `js/app.js:8764-8782` (Break-Even snapshot dropdown population in `refreshBreakeven`)

- [ ] **Step 1: Set Break-Even snapshot default to report month**

In `refreshBreakeven()`, find the snapshot population code (around line 8774-8781):

```javascript
            const prevVal = beSnapshotSelect.value;
            beSnapshotSelect.innerHTML = '<option value="">Current Month</option>' +
                availableMonths.map(m => `<option value="${m}">${Utils.formatMonthShort(m)}</option>`).join('');
            if (prevVal && availableMonths.includes(prevVal)) {
                beSnapshotSelect.value = prevVal;
            } else {
                beSnapshotSelect.value = '';
            }
```

Replace with:

```javascript
            const prevVal = beSnapshotSelect.value;
            beSnapshotSelect.innerHTML = '<option value="">Current Month</option>' +
                availableMonths.map(m => `<option value="${m}">${Utils.formatMonthShort(m)}</option>`).join('');
            if (prevVal && availableMonths.includes(prevVal)) {
                beSnapshotSelect.value = prevVal;
            } else if (this._reportMonth && availableMonths.includes(this._reportMonth)) {
                beSnapshotSelect.value = this._reportMonth;
            } else {
                beSnapshotSelect.value = '';
            }
```

- [ ] **Step 2: Commit**

```bash
git add js/app.js
git commit -m "feat: default Break-Even snapshot month to report month in view-only mode"
```

---

### Task 11: Manual end-to-end testing

**Files:** None (testing only)

- [ ] **Step 1: Test the share modal flow (non-view-only)**

1. Open the app normally
2. Ensure Group Sync / Supabase is configured
3. Click the Share button
4. Verify the modal shows Phase 1: "Report As Of" dropdown + "Generate Share Link" button
5. Verify the dropdown lists "Current (no filter)" and past months from the timeline
6. Select a month (e.g., "March 2026")
7. Click "Generate Share Link"
8. Verify Phase 2 shows: QR code, share URL, "Report: Mar 2026" badge, expiry date
9. Copy the share URL

- [ ] **Step 2: Test the "Current (no filter)" option**

1. Open share modal
2. Leave dropdown on "Current (no filter)"
3. Generate the link
4. Open the link in a new tab
5. Verify normal view-only behavior — no report month badge, all data shown as-is

- [ ] **Step 3: Test view-only mode with a report month set**

1. Open the share URL generated in Step 1 (with March selected)
2. Verify the banner shows "Report: Mar 2026" pill
3. **Journal tab**: Verify April entries are hidden; March entries with April `month_paid` show as "Pending"
4. **Dashboard**: Verify snapshot month picker defaults to March
5. **Cash Flow**: Verify months through March show actuals; April+ shows projections only
6. **P&L**: Verify months through March show actuals; April+ shows projections
7. **Balance Sheet**: Verify defaults to "As of March 2026"
8. **Break-Even**: Verify snapshot picker defaults to March
9. **Dashboard/BS/BE pickers**: Verify they're still functional — can change to other months

- [ ] **Step 4: Test backwards compatibility**

1. If you have a previously shared link (without report month), open it
2. Verify it works exactly as before — no report pill, no filtering

- [ ] **Step 5: Commit if any fixes were needed**

```bash
git add -A
git commit -m "fix: address issues found during manual testing"
```
