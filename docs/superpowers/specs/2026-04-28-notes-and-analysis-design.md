# Notes & Analysis Layer — Design Spec

**Date:** 2026-04-28
**Status:** Draft
**Scope:** Five tightly-related features that add a "meaning layer" to the existing accounting data: written commentary, automated insight surfacing, and one place to see all of it.

## Problem

The app tracks rich financial data (P&L, Cash Flow, Balance Sheet, KPIs, break-even) and supports view-only sharing — but there is no way for the user to:

1. Attach a written explanation to a specific month's column ("March revenue spiked because of the Westbrook order").
2. Be told *which* months/cells are unusual without manually scanning.
3. Aggregate observations across the app into a single browsable log.
4. Pin context to specific points on charts.
5. Carry any of this written context forward into shared view-only reports.

Per-record `notes` fields already exist on transactions, fixed assets, loans, and budget expenses, but they're row-scoped free-text — they cannot describe a month, a column, or a chart point. The result is that shared reports are pure data without narrative, and the user has to re-derive past reasoning by reading raw numbers.

## Solution Summary

Five features, designed to share infrastructure where possible:

| # | Feature | One-liner |
|---|---------|-----------|
| 1 | **Period Notes** | Per-tab, per-month written commentary. Foundation for everything else. |
| 3 | **Variance Auto-Flagging** | Cells with >20% change get an amber tint + corner dot; click for delta math + "Explain this" → Period Note. |
| 5 | **Anomaly Dashboard Widget** | Collapsible "Worth Looking At" panel on Dashboard with top 3 ranked callouts; "Capture explanation" → Period Note. |
| 6 | **Insights Hub Tab** | New tab aggregating Period Notes + Variance annotations + Anomaly explanations + Chart pins + Transaction notes, filterable, searchable. |
| 7 | **Chart Annotations** | Click a point on a Dashboard analyze-mode trend chart to drop a labeled callout marker. |

All five share a common backbone: a generic `notes` table keyed by `(target_type, target_key)` with a single store, single sync path, and unified surfacing in the Insights Hub.

---

## Architecture

### Shared backbone: unified `notes` table

A single new table powers all written content from these features (Period Notes, Variance annotations, Anomaly explanations, Chart annotations). Transaction notes remain on the `transactions.notes` column — they're aggregated into the Insights Hub via a UNION query, not migrated.

```sql
CREATE TABLE IF NOT EXISTS notes (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  target_type TEXT NOT NULL,    -- 'period' | 'variance' | 'anomaly' | 'chart'
  target_key TEXT NOT NULL,     -- composite key, format depends on target_type
  body TEXT NOT NULL,           -- markdown-ish: bold/italic/bullets only
  pinned_to_share INTEGER NOT NULL DEFAULT 0,
  created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
  updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
  UNIQUE(target_type, target_key)
);
CREATE INDEX IF NOT EXISTS idx_notes_target ON notes(target_type, target_key);
```

`target_key` formats:
- `period`: `<tab>:<YYYY-MM>` (e.g., `pnl:2026-03`)
- `variance`: `<tab>:<category_id>:<YYYY-MM>` (e.g., `pnl:42:2026-03`)
- `anomaly`: `<anomaly_signature>:<YYYY-MM>` (signature is a stable hash of the rule + target — see Anomaly section)
- `chart`: `<chart_id>:<x_axis_key>` (e.g., `dashChartPnL:2026-03`)

### Multi-company scoping

The `notes` table lives inside each company's SQLite database (same as `transactions`, `pl_overrides`, etc.). Multi-company isolation is automatic — switching companies switches databases, so notes don't leak across companies. No `company_id` column needed.

### Sync integration

The `notes` table is part of the SQLite file that's exported during sync push. Existing sync logic in `js/sync.js` already serializes the entire DB on push and replaces it on pull, so notes flow with the rest of the data without code changes.

### Settings additions to `app_meta`

| Key | Default | Purpose |
|-----|---------|---------|
| `variance_threshold_pct` | `20` | % change above which cells get flagged. Single global value across all tabs. |
| `pinned_notes_share_cap` | `5` | Max number of pinned notes shown in view-only share. Soft warning if exceeded. |
| `dismissed_variances` | JSON array | `[{"target_key": "pnl:42:2026-03", "value": 5800}]` — re-fires only if value changes. |
| `dismissed_anomalies` | JSON array | `[{"signature": "runway_drop", "month": "2026-04", "target_id": null}]` — `target_id` distinguishes per-target rules (e.g., `category_spike` for category 42 vs category 17 in the same month); `null` for rules without a per-target dimension. |

---

## Feature 1: Period Notes

### Behavior

- A small file/note SVG icon appears in each month column header on **P&L**, **Cash Flow**, and **Balance Sheet** tabs, **on hover only** (the column header gains a subtle hover state).
- Filled amber icon = note exists for that (tab, month). Outline icon = empty.
- Clicking opens an **anchored popover** below the column header. Popover is **draggable** (drag handle in header) and **resizable** (corner grip), with size/position **not persisted** between sessions — it just lets the user adjust during a writing session if they need more room.
- Editor supports **bold, italics, bullets** via simple markdown (`**bold**`, `*italic*`, `- list`). Renders as formatted HTML in view-only mode.
- Auto-saves on blur. "Last edited" timestamp shown in popover footer.
- Per-tab scope: a note on "P&L · March 2026" is distinct from "Cash Flow · March 2026."
- Each note has a **"📌 Pin to shared report"** toggle (drives view-only behavior — see below).

### View-only behavior

- Notes render as inline **footnotes** below the spreadsheet, grouped by month, ordered chronologically.
- Only **pinned** notes appear in shares.
- If pinned count exceeds `pinned_notes_share_cap` (default 5), a warning shows in the share modal: *"You have 7 pinned notes — only the most recent 5 will appear in this share."* Sender can adjust the cap or unpin notes before generating.
- "Most recent" = `notes.updated_at DESC` (most recently edited wins). Ties broken by `created_at DESC`, then `id DESC`.
- The note body renders as formatted markdown (bold/italic/bullets only — no links, no images).

### Files to modify / create

| File | Change |
|------|--------|
| `js/database.js` | Migration to create `notes` table; CRUD methods `getNote(targetType, targetKey)`, `setNote(targetType, targetKey, body)`, `deleteNote(targetType, targetKey)`, `setNotePinned(targetType, targetKey, pinned)`, `getPinnedShareNotes(reportMonth)` |
| `js/ui.js` | `renderProfitLossSpreadsheet`, `renderCashFlowSpreadsheet`, `renderBalanceSheet` — add note icon to month column headers; helper `renderNoteIcon(targetKey)` |
| `js/app.js` | New controller block: `openPeriodNote(tab, month)` (opens popover), `savePeriodNote(...)`, `togglePinNote(...)`, drag/resize handlers |
| `index.html` | New `periodNotePopover` element (single reusable popover, repositioned per click) |
| `css/styles.css` | New section "Period Notes" — icon hover, popover, drag handle, resize grip, footnote rendering |
| View-only logic in `app.js` | `applyViewOnlyRestrictions` — hook a render-footnotes call after each spreadsheet render |

### Markdown rendering

Use a tiny inline parser (no dependency) supporting only:
- `**text**` → `<strong>`
- `*text*` → `<em>`
- Lines starting with `- ` → `<ul><li>`

**Render order is critical for safety:**

1. HTML-escape the entire raw input first (`<` → `&lt;`, `>` → `&gt;`, `&` → `&amp;`, `"` → `&quot;`).
2. Then run the bold/italic/list transforms on the escaped string.
3. Inject result via `innerHTML`.

This ordering guarantees that user-typed `<script>` or other HTML in a note body cannot execute. The transforms only operate on the literal characters `*` and `- `, which survive escaping. Same parser is used everywhere notes render.

---

## Feature 3: Variance Auto-Flagging

### Behavior

- For every value cell in **P&L** and **Cash Flow** spreadsheets, compute two deltas:
  - `delta_prior = (current - prior_month) / |prior_month|`
  - `delta_avg3 = (current - avg_of_prior_3) / |avg_of_prior_3|`
- If `|delta_prior| >= variance_threshold_pct` OR `|delta_avg3| >= variance_threshold_pct` (default 20%) → **flag the cell**.
- Visual: amber background tint (`#fffbeb`) + small corner dot (top-right, 6px, amber). Same flag works in dark mode (use existing variant tokens).
- Hovering shows a quick tooltip: *"+142% vs prior month."* Clicking opens the **Variance Popover**.

### Variance Popover

Anchored to the cell. Contents:
- Title: `Marketing · March 2026`
- Stats:
  - This month: `$5,800`
  - vs prior month: `+$3,400 · +142%`
  - vs 3-mo avg: `+$3,500 · +152%`
  - YoY (same month last year): `+$3,200 · +123%` or `no data`
- Buttons: `Dismiss flag` · `Explain this →`
- "Explain this" closes the variance popover and opens a Period Note popover for the same `(tab, month)` — pre-focused, with a placeholder hint: *"Why was Marketing high in March?"*

### Dismissal

- "Dismiss flag" stores `{target_key: "pnl:42:2026-03", value: 5800}` in `dismissed_variances`.
- Flag does not appear unless the cell's value changes. If it changes back to a flaggable delta, the entry is invalidated and the flag reappears.

### View-only behavior

- Variance flags **do not render** in shares (clean visual presentation).
- However, any Period Notes captured *via* "Explain this →" are normal Period Notes and DO appear (subject to pin status).

### Settings

A single setting `variance_threshold_pct` (default 20) in `app_meta`. Exposed in the gear/settings popover. Single global value across both P&L and Cash Flow.

### Files to modify / create

| File | Change |
|------|--------|
| `js/utils.js` | New `computeVariance(values, currentIdx, threshold)` returning `{flagged, deltaPrior, deltaAvg3, deltaYoY}` |
| `js/ui.js` | `renderProfitLossSpreadsheet`, `renderCashFlowSpreadsheet` — call `computeVariance` per cell; apply `.cell-flagged` class + corner dot. Add `varianceTooltip` hover handler. |
| `js/app.js` | `openVariancePopover(tab, categoryId, month, value)`, `dismissVariance(...)`, integration with `openPeriodNote` |
| `index.html` | New `variancePopover` element |
| `css/styles.css` | `.cell-flagged`, `.cell-flagged::after` corner dot, `.variance-popover` |
| Settings UI | Add `variance_threshold_pct` input to gear/settings popover (`index.html` + `app.js` handler) |

---

## Feature 5: Anomaly Dashboard Widget

### Behavior

- A **collapsible** card titled "⚡ Worth Looking At — {Month}" appears on the Dashboard, above existing KPI sections.
- Default state: **collapsed** with a count badge ("3 items"). User clicks to expand.
- Shows top 3 anomalies for the current dashboard snapshot month, ranked by severity.
- "Show more" expander reveals up to 7 more if available.
- Each anomaly card: severity icon · title · description · action buttons.

### Severity tiers

| Tier | Background | Border | Use |
|------|------------|--------|-----|
| **Critical (red)** | `--red-bg` | `--red` | Things that hurt: cash runway dropped, profitability flipped negative, DSCR < 1, AR aging climbing 3+ months |
| **Notable (amber)** | `--amber-bg` | `--amber` | Big variances: category 3×+ avg, runway shifts >10%, unusual one-time expenses |
| **Positive (blue)** | `--blue-bg` | `--primary` | Good news to acknowledge: first profitable month, milestone hits, streak achievements |

### Anomaly rules (v1 set)

Each rule produces a stable `signature` for dedup/dismissal.

| Signature | Tier | Rule | Description format |
|-----------|------|------|--------------------|
| `runway_drop` | critical | Cash runway shorter by >10 days vs prior month | "Cash runway shortened by {N} days" |
| `category_spike` | notable | Any category > 3× its 6-month avg in current month | "{Category} spend {N}× the 6-month avg" |
| `ar_aging_trend` | critical | AR Days Outstanding climbing 3 months in a row | "AR Days Outstanding climbing 3 months in a row" |
| `profitability_flip` | positive (if positive) / critical (if negative) | Net income sign change vs prior month | "First profitable month / first unprofitable month" |
| `dscr_breach` | critical | DSCR < 1.0 in current month, was ≥ 1.0 prior | "DSCR dropped below 1.0 — covenant risk" |
| `recurring_disappeared` | notable | A category with 6 consecutive months of activity has $0 this month | "{Category} (recurring) has no entries this month" |

Rules are computed live on dashboard load from existing DB queries (no caching for v1 — sql.js queries are cheap).

### Action buttons per anomaly

- **Capture explanation** → opens a Period Note popover for the relevant `(tab, month)` based on which tab the anomaly stems from. Pre-fills the textarea with a starter prompt like *"This anomaly was: {anomaly title}. Why?"* (user can delete).
- **View [tab] →** jumps to the relevant tab + scrolls to the anomaly's source area.
- **Dismiss** → stores `{signature, month}` in `dismissed_anomalies`. Anomaly will not appear again *for that month* but will reappear next month if still triggered.

### Files to modify / create

| File | Change |
|------|--------|
| `js/utils.js` | Anomaly rule functions, one per signature, returning `{signature, severity, title, description, sourceTab, targetMonth, explanationStarter} \| null` |
| `js/app.js` | New `_computeAnomalies(snapshotMonth)`, `_renderAnomalyWidget()` invoked from `_renderDashboardSections`; collapse/expand state; dismissal handlers |
| `js/ui.js` | `renderAnomalyCard(anomaly)` helper |
| `index.html` | Anomaly widget shell in Dashboard tab |
| `css/styles.css` | New section "Anomaly Widget" — card, collapsed/expanded states, severity tiers |

---

## Feature 6: Insights Hub Tab

### Tab placement

A new tab labeled **"Insights"** added to the main tab strip. Default position: between "Balance Sheet" and "Fixed Assets" (or wherever Tab Manager places it — respects user's tab reordering).

### Layout

Two-column layout inside the tab:

- **Left sidebar (220px):** filter pills grouped by Type, Source tab, Time range. Each pill shows a count.
- **Right content area:** search box on top, then a chronological list of note cards grouped by month (newest first).

### Data sources unified into the feed

All entries are aggregated via UNION query at render time. Note: there is **no `target_type='variance'` row source** — the "Explain this →" flow creates `target_type='period'` notes, not variance-typed ones. Variance flags themselves are computed live, not stored. The four real sources are:

| Source | Tag | Source label |
|--------|-----|--------------|
| `notes` where `target_type='period'` | 📝 Period Note | `{Tab} · {Month}` |
| `notes` where `target_type='anomaly'` | 🎯 Anomaly | `Dashboard · {Month}` |
| `notes` where `target_type='chart'` | 📊 Chart Pin | `{Chart name} · {Month}` |
| `transactions` where `notes IS NOT NULL AND notes != ''` | 💬 Transaction | `Journal · #{id} · {date}` |

### Unified column projection (UNION shape)

`getInsightsFeed(filters)` returns rows with these columns regardless of source:

| Column | Type | Description |
|--------|------|-------------|
| `source_type` | TEXT | `'period'` \| `'anomaly'` \| `'chart'` \| `'transaction'` |
| `source_id` | INTEGER | Original row id (`notes.id` or `transactions.id`) |
| `target_key` | TEXT | For notes: original `target_key`. For transactions: `'txn:' \|\| id` |
| `body` | TEXT | `notes.body` or `transactions.notes` |
| `source_month` | TEXT | YYYY-MM extracted: from `target_key` for notes, from `transactions.entry_date` for txns |
| `source_label` | TEXT | Human-readable, computed at SQL or JS level (table above) |
| `updated_at` | DATETIME | `notes.updated_at` or `transactions.updated_at` |
| `pinned_to_share` | INTEGER | `notes.pinned_to_share` or `0` for txns (txns are not pinnable to shares in v1) |

### Filters

- **Type:** All / Period / Anomaly / Chart / Transaction (multi-select) — no separate "Variance" type, since variance-driven notes are stored as Period Notes
- **Source tab:** Cash Flow / P&L / Balance Sheet / Dashboard / Journal (multi-select)
- **Time range:** All time / This quarter / This year (single-select)
- **Search:** full-text across `body` (or `transactions.notes`)

### Card interactions

- Click a card → cross-tab navigation: switches to source tab, scrolls to source location, opens the relevant popover/modal pre-focused.
  - Period note → P&L/CF/BS tab + scroll to month + open Period Note popover
  - Variance → P&L/CF tab + scroll to row + open Variance popover
  - Anomaly → Dashboard + expand widget + flash the anomaly card
  - Chart pin → Dashboard + scroll to chart + flash the annotation
  - Transaction → Journal tab + scroll to row + highlight

### Sort

Default: by source month, newest first. Within a month, by `updated_at` descending.

### Files to modify / create

| File | Change |
|------|--------|
| `js/database.js` | `getInsightsFeed(filters)` — UNION query across notes + transactions |
| `js/app.js` | New tab controller `refreshInsights()`, filter state, search debounce, card click → cross-tab nav router |
| `js/ui.js` | `renderInsightsTab(feed, filters)`, `renderInsightCard(item)` |
| `index.html` | New `insightsTab` div with sidebar + content shell |
| `css/styles.css` | New section "Insights Hub" — sidebar pills, search, timeline groups, cards |
| `index.html` (tabs) | Add Insights tab to the main tab strip |
| Tab manager logic | Ensure new tab is included in default order, can be hidden/reordered like others |

### Export (deferred)

Out of scope for v1. Future: a "Download as PDF/markdown" button that bundles current filter state. Re-evaluate after v1 is shipped and user has actual content to export.

---

## Feature 7: Chart Annotations

### Scope (v1)

Only the three Dashboard analyze-mode trend charts:

- `dashChartCashFlow` — Cash Flow trend
- `dashChartPnL` — P&L trend
- `dashChartBS` — Balance Sheet trend

The Revenue Concentration **pie chart** and per-KPI **sparklines** are **excluded** — annotations don't make sense for pies, and sparklines are too small.

### Behavior

- Each chart shows a small **"📌 Add annotation"** button in its header.
- Clicking enters "annotation mode" for that chart: the next click on a data point on that chart opens an inline editor anchored to the point.
- User types annotation text, hits Save.
- Annotation renders as:
  - A larger amber dot at the data point
  - A dashed vertical line down to the X axis
  - A callout box near the point with `📍 {month} — {body excerpt}`
- Hovering the point shows the full body in a tooltip.
- Multiple annotations on one chart **auto-stack** vertically near their points; user can also **drag the callout box** to reposition it (X/Y offset persisted in the note's metadata — see schema note below).

### Drag positioning persistence

The `notes.body` field stores user text, but chart annotations also need an offset. Add a JSON metadata column or use a sidecar key:

```
notes.body   = "Westbrook Co. wholesale order — flag as one-time when forecasting."
notes.target_key = "dashChartPnL:2026-03"
```

Position offset stored separately: append `position_x`/`position_y` to the body using a delimited prefix? **No** — cleaner: add a `metadata TEXT` column to the `notes` table (nullable JSON blob) that anything can use without schema migration. Chart annotations use it for `{"dx": 40, "dy": -20}`. Other note types ignore it.

### View-only behavior

- Annotations **render on hover** in shared reports (not always visible). Hovering the data point shows the callout. Marker dot is always visible.
- Rationale: cleaner default view of the chart, but the explanation is one hover away. (User explicitly chose this.)

### Files to modify / create

| File | Change |
|------|--------|
| `js/database.js` | Migration: add `metadata TEXT` column to `notes` table; helper `setChartAnnotation(chartId, xKey, body, position)` |
| `js/app.js` | Annotation mode toggle per chart; click-on-point handler (Chart.js's `onClick`); inline editor; drag handlers writing back to `metadata` |
| `js/ui.js` | `renderChartAnnotations(chartId, annotations)` — overlays positioned absolutely over the canvas |
| `index.html` | "Add annotation" button injected into chart headers |
| `css/styles.css` | New section "Chart Annotations" — pin marker, callout box, dashed line, dragging cursor states |

### Edge cases

- Chart x-axis range narrows so the annotation's `xKey` is no longer in view → annotation hidden (not deleted; reappears when range expands).
- Annotation's `xKey` no longer maps to a data point at all (rare; e.g., user changes timeline or category) → annotation is orphaned, hidden by default. A footer pill shows "1 hidden annotation" with a button to view/delete orphans.

---

## Cross-cutting concerns

### Build phasing

Build in this order, each phase independently shippable:

1. **Phase 1: Period Notes (foundation)** — schema, popover, view-only footnotes, pin-to-share. Establishes the `notes` table and markdown parser used by everything else.
2. **Phase 2: Variance Auto-Flagging** — adds variance computation + popover + "Explain this" handoff to Period Notes. Reuses the popover pattern.
3. **Phase 3: Insights Hub** — pure aggregation feature, no new data writes. Validates that the unified-feed approach works.
4. **Phase 4: Anomaly Widget** — depends on Period Notes for "Capture explanation" handoff. Anomaly rules can land incrementally.
5. **Phase 5: Chart Annotations** — last because it's the most isolated feature and uses the `notes.metadata` column added late.

### Sync

The `notes` table sits inside the SQLite DB that's already serialized on sync push and replaced on pull. No changes to `js/sync.js` needed. Verify in testing that two users editing the same `(target_type, target_key)` triggers the existing last-write-wins behavior cleanly.

### Settings UI changes

Add to the existing gear/settings popover:

- Variance threshold % (number input, default 20)
- Pinned notes share cap (number input, default 5)

Both writable to `app_meta` via existing meta-set helpers.

### Tab management

Insights tab participates in the existing tab reorder/hide system (`getTabOrder`, `getHiddenTabs` in `js/database.js`). Default position is between Balance Sheet and Fixed Assets. Default visibility: shown.

### Multi-company

All notes live in the company's SQLite DB. Switching companies via `CompanyManager.switchTo` switches DBs, so notes are automatically scoped. No additional code needed.

### View-only mode

Pinned Period Notes render as footnotes after each spreadsheet. Variance flags are stripped. Chart annotations hover-only. Anomaly widget is hidden in view-only mode (it's a workflow tool, not part of the report).

Insights Hub tab is **always hidden in view-only mode**, enforced by a hardcoded check inside `applyViewOnlyRestrictions` — *not* via the user's `hidden_tabs` preference. Rationale: `hidden_tabs` is per-company sender configuration, but view-only is recipient-side; we don't want the sender's hide preference to leak into shared reports, and we want this hide to be unconditional regardless of how the sender configured their tab visibility.

### Testing approach

- **Period Notes:** create/edit/delete; verify markdown rendering; verify per-tab scope (P&L Mar ≠ CF Mar); verify sync round-trip.
- **Variance:** verify threshold behavior at edges (19.9% → no flag, 20.0% → flag); verify dismissal persistence; verify re-fire on value change.
- **Anomaly:** unit-test each rule function with crafted inputs; verify dismissal scope (per-month).
- **Insights Hub:** verify UNION query returns all 5 source types; verify search hits all body fields; verify cross-tab nav lands on correct row + opens correct popover.
- **Chart annotations:** verify drag persists; verify orphan handling when timeline shifts.

---

## Open questions / future work

- **Export from Insights Hub** (PDF/markdown bundle) — defer to v2.
- **Custom anomaly rules** — let users write their own thresholds. v2.
- **Annotation collision avoidance** — auto-stack works for v1; if charts get >5 annotations, may need a "list view" alternative.
- **Notes on KPI snapshots** — separate from Period Notes; could be added as a 6th `target_type` later if useful.
- **Search highlighting** — show matched terms with `<mark>` in Insights Hub. v1.5.
- **Note history / versioning** — currently last-write-wins; v2 could keep an edit history.

---

## Out of scope

- Threaded comments / multi-user discussion on shares (would require auth on the share recipient side).
- Rich-text WYSIWYG editor (markdown subset is enough for the audience).
- Notifications when someone in your sync group adds a note (no notification system exists today).
- Automated narrative generation ("AI-written period notes") — explicitly user-driven only.
- Visual edits to existing UI components — per user's directive (2026-04-28), the new features adopt the polished mockup style but **no existing UI is restyled** as part of this work.
