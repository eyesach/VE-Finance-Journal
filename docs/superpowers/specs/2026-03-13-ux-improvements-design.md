# UX Improvements Design Spec

**Date:** 2026-03-13
**Status:** Approved
**Scope:** Workflow speed, data visibility, and analytics enhancements

---

## Overview

Five features to improve the Accounting Journal Calculator's user experience, focused on faster data entry and richer financial insights. The central architectural change is splitting the app into two modes — Work (data entry) and Analyze (visualization) — accessed via a sub-tab toggle.

---

## Feature 1: Work / Analyze View Toggle

### What it does
A sub-tab toggle at the top of the main content area (styled like the existing VE Sales "Sales | Events" toggle) that switches between Work and Analyze modes. The sidebar tab list updates to show the relevant tabs for each mode.

### Tab assignments

**Work mode:** Journal, Budget, Products, VE Sales, Assets & Equity, Loans, Projected Sales

**Analyze mode:** Dashboard (new), Cash Flow, P&L, Balance Sheet, Break-Even

**Always accessible:** Change Log (visible in both modes via a settings/gear menu link, not a sidebar tab)

### Tab management interaction
- The existing tab hide/show system (gear modal, drag-to-reorder) operates within the active mode only. Hidden tab state and custom ordering are stored per-mode.
- The gear modal shows only the current mode's tabs for reorder/hide operations.
- `restoreTabOrder()` and `applyHiddenTabs()` become mode-aware, filtering by the active mode's tab list before applying.

### Behavior
- Switching modes remembers which tab you were on in each mode
- Summary cards (Cash Balance, AR, AP) stay visible in both modes, positioned above the mode toggle so they are unaffected by mode switching
- The toggle lives in the header area, above the tab content but below the summary cards
- Default mode on load: Work
- Dashboard tab requires a new `<div id="dashboardTab">` in index.html and a `refreshDashboard()` function wired into the tab switch logic

---

## Feature 2: Multi-Line Quick Entry Modal

### What it does
A spreadsheet-style modal for entering multiple transactions or budget items at once. Accessed via a "Quick Add" button on the Journal and Budget tabs.

### Journal version — columns
Date, Category, Amount, Type (Receivable/Payable), Status, Month Due, Notes

### Smart expand behavior
- When you select a category, if that category type needs extra fields (e.g., a sales category reveals Pretax Amount, Sale Period, Inventory Cost), an expand arrow appears on that row
- Clicking it reveals a second detail line below the row with the conditional fields: Date Processed, Month Paid, Pretax Amount, Inventory Cost, Sale Period (start/end), Payment For Month
- Rows that don't need extra fields stay compact

### Budget version — columns
Expense Name, Amount, Budget Group, Frequency (monthly/quarterly/annual/one-time)

### Budget expand behavior
Budget rows have no expand behavior — all fields fit in the main row. The Budget quick entry is a simpler grid than the Journal version.

### Validation on Save All
When "Save All" is clicked, rows with missing required fields (no date, no category, no amount for Journal; no name, no amount for Budget) are highlighted with a red border. Save proceeds only for valid rows, or the user can fix and retry. A count shows "2 of 5 rows have errors."

### Interaction
- Tab between cells to move across, Enter or Tab from last cell adds a new empty row
- Category dropdown uses type-ahead search (start typing to filter)
- "Save All (N entries)" button at the bottom right
- Each row has a delete icon to remove it before saving
- Cancel discards everything

---

## Feature 3: Analyze Dashboard (New Tab)

### What it does
The default landing tab when switching to Analyze mode. A hybrid layout — KPI summary at the top, then expandable chart sections below for deeper dives.

### Top section — KPI Cards with Sparklines
- **Cash Position** — current balance + 6-month sparkline trend
- **Monthly Burn Rate** — average expenses + sparkline
- **Revenue Trend** — this month vs. last month + sparkline
- **Overdue Receivables** — total overdue amount + count
- **Break-Even Progress** — percentage toward break-even this month, shown as a small progress bar

These are enhanced versions of the existing summary cards (Cash Balance, AR, AP) plus new derived metrics.

### Below the KPIs — Collapsible sections
Each section has a header you can click to expand/collapse. Expanded by default on first load, then remembers your preference.

1. **Cash Flow Overview** — Waterfall or horizontal bar chart showing money in vs. money out by month, color-coded by category
2. **P&L Trends** — Line chart with revenue and expense lines over time, profit margin overlay
3. **Month-over-Month Comparison** — Two-month side-by-side with category-level increases/decreases highlighted in green/red
4. **Balance Sheet Snapshot** — Stacked bar showing assets vs. liabilities, equity gap visual
5. **Break-Even Tracker** — Thermometer/progress bar showing cumulative progress toward break-even, updated from real transaction data

### Behavior
- All charts pull from the same database the existing report tabs use — no duplicate data
- Charts refresh every time an Analyze tab is activated (same pattern as existing `switchMainTab` calling each tab's refresh function). No dirty-flag tracking needed.
- Sections are collapsible so you can focus on what matters
- Collapse state persisted in localStorage (keyed per section, not per company — this is a UI preference)
- "Overdue Receivables" defined as: transactions where `status = 'pending'` AND `type = 'receivable'` AND `month_due < current YYYY-MM`. Reuses existing `checkLatePayments()` logic.
- Sparklines rendered as tiny Chart.js line charts (no axes, no labels, no tooltips, ~80x30px canvas) — minimal config to avoid over-engineering

---

## Feature 4: Chart-Enhanced Analyze Report Tabs

### What it does
The Cash Flow, P&L, Balance Sheet, and Break-Even tabs in Analyze mode show the same data as their Work mode counterparts, but led by charts with the data table secondary.

### Cash Flow (Analyze)
- Top: Waterfall-style chart implemented as Chart.js floating bar chart using `[min, max]` data format (Chart.js 4.x has no native waterfall type). Green bars for positive months, red for negative.
- Below: The existing cash flow spreadsheet table, read-only (editing stays in Work mode)

### P&L (Analyze)
- Top: Dual line chart — revenue vs. expenses over time, with a shaded profit/loss area between them
- Optional toggle: monthly vs. quarterly view
- Below: Read-only P&L table

### Balance Sheet (Analyze)
- Top: Stacked bar — assets breakdown on left, liabilities + equity on right
- Financial ratios displayed as gauge-style indicators (current ratio, debt-to-equity) — implemented as CSS-only gauges using conic-gradient (no additional library needed)
- Below: Read-only balance sheet table

### Break-Even (Analyze)
- Top: Progress thermometer showing how close to break-even this period
- Channel breakdown as a grouped bar chart (if tracking multiple channels)
- Below: Read-only break-even data table

### Key principle
Analyze tabs are read-only views. Editing overrides or adjusting numbers happens in Work mode. This keeps the Analyze view clean and focused on insights.

---

## Feature 5: Undo/Redo System

### What it does
A history stack that lets you reverse recent actions with Ctrl+Z (undo) and Ctrl+Y (redo). Covers destructive or significant actions — not every keystroke.

### Actions tracked
- Transaction create, edit, delete
- Bulk status changes
- Budget expense create, edit, delete
- Category/folder create, edit, delete
- Fixed asset and loan changes

### How it works
- Each tracked action stores a "before" and "after" snapshot (just the affected row(s), not the whole database)
- Bulk operations (e.g., bulk status change on 30 rows) are stored as a single history entry containing all affected rows, not one entry per row
- Undo reverts the database change and refreshes the current tab
- A small toast notification appears at the bottom: "Undid: Delete transaction 'Office supplies'" with a "Redo" button
- Toast auto-dismisses after 4 seconds
- History stack holds ~50 actions, clears on page reload or company switch
- Company switch integration: the undo stack (in-memory array on App) is cleared inside the company switch handler, before the new company's database is loaded

### What it doesn't track
- Navigation (tab switches, filter changes)
- Report override edits (deliberate and rare)
- Settings changes

### Keyboard shortcuts
- Ctrl+Z — Undo
- Ctrl+Y or Ctrl+Shift+Z — Redo

---

## Appendix: Future Ideas — Navigation & Organization (C)

These are undesigned ideas saved for a future ideation session.

1. **Full-text search bar** — Search across all transaction descriptions, notes, and categories from anywhere in the app. Results show the matching entry with a link to jump to it.
2. **Saved filters / filter presets** — Save commonly used filter combinations (e.g., "Overdue payables this quarter") as named presets you can apply with one click.
3. **Date range picker** — Replace the two separate date inputs with a single date range picker component. Include presets like "This month", "Last quarter", "YTD".
4. **Breadcrumb trail** — Show current context (Company > Mode > Tab) as a breadcrumb in the header so you always know where you are.
5. **Recent items / quick jump** — A keyboard shortcut (Ctrl+K) that opens a command palette to jump to any tab, recent transaction, or category by typing.

## Appendix: Future Ideas — Data Management (D)

1. **Drag-and-drop CSV import** — Drop a CSV file anywhere on the journal to start an import flow with column mapping and preview.
2. **Bulk edit modal** — Select multiple transactions with checkboxes, then bulk-change category, status, or month due in one action.
3. **Database backup/restore UI** — A settings panel to manually export/import the full SQLite database as a file, with timestamped backups.
4. **Sync status dashboard** — A dedicated panel showing last sync time, conflict history, and manual push/pull buttons for Supabase sync.
5. **Auto-categorization rules** — Define rules like "Description contains 'rent' → Category: Rent, Type: Payable" to auto-fill fields during entry.
