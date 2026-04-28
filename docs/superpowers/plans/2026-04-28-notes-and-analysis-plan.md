# Notes & Analysis Layer — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a written-commentary + auto-insight layer on top of the existing accounting data — five interconnected features (Period Notes, Variance Auto-Flagging, Anomaly Widget, Insights Hub, Chart Annotations) that turn raw numbers into a narrative.

**Architecture:** A single new `notes` table keyed by `(target_type, target_key)` powers four of the five features (transactions keep their existing `notes` column). Variance flags are computed live and not stored. Anomalies are computed live; their dismissals live in `app_meta`. Each phase is independently shippable; later phases reuse the popover/markdown/handoff patterns from Phase 1.

**Tech Stack:** Vanilla HTML/JS, sql.js (in-browser SQLite), Chart.js (for chart annotations), DM Sans, no build pipeline beyond `node build-frontend.js` (Tauri). No new dependencies introduced.

**Spec reference:** `docs/superpowers/specs/2026-04-28-notes-and-analysis-design.md`

**Testing approach:** This codebase has no automated test runner; the only existing test file is `tests/ve-sales.test.html` (browser-loaded). For this work:
- **Pure logic** (markdown parser, variance math, anomaly rule functions, dismissal matching) → a new `tests/notes-analysis.test.html` with the same suite/test/assert pattern as `ve-sales.test.html`. Engineer opens it in browser, expects all green.
- **UI behavior** → smoke-tested manually (open `index.html`, do X, verify Y). Each task lists exact click steps and expected state.
- **Persistence** → verified by saving, refreshing the page, confirming state restored.

**Conventions to follow:**
- Match existing `database.js` migration pattern (`try { exec SELECT col } catch { ALTER TABLE ADD COLUMN }` for new columns; `CREATE TABLE IF NOT EXISTS` for new tables).
- Match existing `app.js` controller method naming (camelCase, `_privateHelper`, `openXModal`/`closeX`/`saveX`).
- All new CSS goes in a new section in `css/styles.css` with a header comment block.
- Per user directive (2026-04-28): **do NOT restyle existing UI**. Build new components in the higher-fidelity mockup style; leave older sections untouched.
- Frequent commits — one per task. Use messages like `feat(notes): add period notes schema migration`.

---

## File Structure

| File | Role |
|------|------|
| `js/database.js` | Add `notes` table migration; CRUD: `getNote`, `setNote`, `deleteNote`, `setNotePinned`, `getPeriodNotesForMonth`, `getPinnedShareNotes`, `getInsightsFeed` |
| `js/utils.js` | `renderMarkdownSafe(body)`, `computeVariance(values, idx, threshold)`, `computeAnomalies(db, snapshotMonth, dismissed)` and 6 individual rule functions |
| `js/ui.js` | `renderNoteIcon`, `renderPeriodNoteFootnotes`, `renderInsightsTab`, `renderInsightCard`, `renderAnomalyCard`, `renderChartAnnotations` |
| `js/app.js` | Period note popover controller (open/save/close/drag/resize/pin); variance popover controller; anomaly widget controller; insights tab controller; chart annotation mode |
| `index.html` | New `periodNotePopover`, `variancePopover` elements; `insightsTab` shell; anomaly widget shell on dashboard; "Insights" tab in main strip; settings inputs for thresholds |
| `css/styles.css` | New section "Notes & Analysis Layer" — icon, popovers, footnotes, variance flags, insights hub, anomaly cards, chart annotations |
| `tests/notes-analysis.test.html` | Browser tests for pure logic units |

---

## Phase 1: Period Notes (foundation)

This phase establishes the `notes` table, the markdown parser, the popover UX, and the view-only footnote rendering. Everything in later phases reuses these.

### Task 1.1: Schema migration for `notes` table

**Files:**
- Modify: `js/database.js` (add to `migrateSchema()` near line 770, before final closing brace)

- [ ] **Step 1: Add migration block**

Find the `migrateSchema()` method in `js/database.js` and append at the end (just before the closing `}` of the method):

```javascript
        // === Notes & Analysis Layer (2026-04-28) ===
        this.db.run(`
            CREATE TABLE IF NOT EXISTS notes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                target_type TEXT NOT NULL,
                target_key TEXT NOT NULL,
                body TEXT NOT NULL,
                pinned_to_share INTEGER NOT NULL DEFAULT 0,
                metadata TEXT,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(target_type, target_key)
            )
        `);
        try { this.db.exec('SELECT metadata FROM notes LIMIT 1'); }
        catch (e) { this.db.run('ALTER TABLE notes ADD COLUMN metadata TEXT'); }
        this.db.run('CREATE INDEX IF NOT EXISTS idx_notes_target ON notes(target_type, target_key)');
```

- [ ] **Step 2: Smoke-test the migration**

Open `index.html` in the browser. Open DevTools console. Run:

```javascript
DB.db.exec("SELECT sql FROM sqlite_master WHERE name='notes'")
```

Expected: returns one row with the CREATE TABLE statement including `metadata TEXT`.

- [ ] **Step 3: Commit**

```bash
git add js/database.js
git commit -m "feat(notes): add notes table schema migration"
```

---

### Task 1.2: Markdown parser with HTML safety

**Files:**
- Modify: `js/utils.js` (append new function near end, before module export if any)
- Create: `tests/notes-analysis.test.html`

- [ ] **Step 1: Create test file scaffold (copy from ve-sales pattern)**

Create `tests/notes-analysis.test.html`:

```html
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Notes & Analysis — Unit Tests</title>
<style>
  body { font-family: 'DM Sans', system-ui, sans-serif; padding: 24px; background: #f3f4f6; }
  h1 { font-size: 1.4rem; margin-bottom: 16px; }
  .suite { margin-bottom: 20px; }
  .suite h2 { font-size: 1rem; margin-bottom: 8px; border-bottom: 2px solid #3b82f6; padding-bottom: 4px; }
  .test { padding: 6px 12px; margin: 2px 0; border-radius: 4px; font-size: 0.85rem; }
  .test.pass { background: #d1fae5; color: #065f46; }
  .test.fail { background: #fee2e2; color: #991b1b; }
  .summary { margin-top: 20px; padding: 16px; background: #fff; border-radius: 8px; font-weight: 600; }
  .error-detail { font-size: 0.78rem; color: #991b1b; margin-top: 2px; font-family: monospace; white-space: pre-wrap; }
</style>
</head>
<body>
<h1>Notes & Analysis — Unit Tests</h1>
<div id="results"></div>
<div id="summary" class="summary"></div>

<script src="https://cdnjs.cloudflare.com/ajax/libs/sql.js/1.8.0/sql-wasm.js"></script>
<script src="../js/utils.js"></script>

<script>
let passed = 0, failed = 0;
const resultsEl = document.getElementById('results');

function suite(name) {
  const div = document.createElement('div');
  div.className = 'suite';
  div.innerHTML = `<h2>${name}</h2>`;
  resultsEl.appendChild(div);
  return div;
}
function test(suiteEl, name, fn) {
  const div = document.createElement('div');
  div.className = 'test';
  try { fn(); div.className = 'test pass'; div.textContent = '✔ ' + name; passed++; }
  catch (e) { div.className = 'test fail'; div.innerHTML = '✘ ' + name + `<div class="error-detail">${e.message || e}</div>`; failed++; }
  suiteEl.appendChild(div);
}
function assert(cond, msg) { if (!cond) throw new Error(msg || 'Assertion failed'); }
function assertEq(a, b, msg) { if (a !== b) throw new Error((msg || 'Not equal') + `: got ${JSON.stringify(a)}, expected ${JSON.stringify(b)}`); }

// === Markdown parser tests ===
const s1 = suite('renderMarkdownSafe');
test(s1, 'escapes < > & " before transforms', () => {
  const out = Utils.renderMarkdownSafe('<script>alert(1)</script>');
  assert(!out.includes('<script>'), 'must not contain raw <script>');
  assert(out.includes('&lt;script&gt;'), 'must contain escaped tag');
});
test(s1, 'bold transform applies after escaping', () => {
  assertEq(Utils.renderMarkdownSafe('**hello**'), '<strong>hello</strong>');
});
test(s1, 'italic transform applies', () => {
  assertEq(Utils.renderMarkdownSafe('*hi*'), '<em>hi</em>');
});
test(s1, 'list lines transform', () => {
  const out = Utils.renderMarkdownSafe('- one\n- two');
  assert(out.includes('<ul>'), 'has ul');
  assert(out.includes('<li>one</li>'), 'has first li');
  assert(out.includes('<li>two</li>'), 'has second li');
});
test(s1, 'mixed: escape + bold + list', () => {
  const out = Utils.renderMarkdownSafe('**bold** with <b>raw</b>\n- item');
  assert(out.includes('<strong>bold</strong>'), 'bold transformed');
  assert(out.includes('&lt;b&gt;raw&lt;/b&gt;'), 'raw HTML escaped');
  assert(out.includes('<li>item</li>'), 'list rendered');
});
test(s1, 'plain text passes through unchanged', () => {
  assertEq(Utils.renderMarkdownSafe('hello world'), 'hello world');
});

// Display summary
document.getElementById('summary').textContent = `${passed} passed, ${failed} failed`;
document.getElementById('summary').style.background = failed === 0 ? '#d1fae5' : '#fee2e2';
</script>
</body>
</html>
```

- [ ] **Step 2: Run the failing tests**

Open `tests/notes-analysis.test.html` in the browser.
Expected: all 6 tests fail with "Utils.renderMarkdownSafe is not a function" or similar.

- [ ] **Step 3: Implement `renderMarkdownSafe` in `js/utils.js`**

Append to `js/utils.js` inside the `Utils` object (before the closing `}`):

```javascript
    /**
     * Render minimal markdown (bold, italic, bullets) safely.
     * Order: escape HTML first, THEN transform. Prevents <script> injection.
     * @param {string} body Raw user input
     * @returns {string} HTML string safe for innerHTML
     */
    renderMarkdownSafe(body) {
        if (!body) return '';
        // 1. Escape HTML
        let s = String(body)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
        // 2. Bold (must run before italic since ** contains *)
        s = s.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
        // 3. Italic
        s = s.replace(/\*([^*]+)\*/g, '<em>$1</em>');
        // 4. Lists: contiguous lines starting with "- " become a <ul>
        const lines = s.split('\n');
        const out = [];
        let inList = false;
        for (const line of lines) {
            if (line.startsWith('- ')) {
                if (!inList) { out.push('<ul>'); inList = true; }
                out.push('<li>' + line.slice(2) + '</li>');
            } else {
                if (inList) { out.push('</ul>'); inList = false; }
                out.push(line);
            }
        }
        if (inList) out.push('</ul>');
        return out.join('\n').replace(/\n+/g, function(m) {
            // Preserve plain newlines outside of structural tags
            return m;
        });
    },
```

- [ ] **Step 4: Re-run tests**

Reload `tests/notes-analysis.test.html`.
Expected: 6 / 6 passing, summary card green.

- [ ] **Step 5: Commit**

```bash
git add js/utils.js tests/notes-analysis.test.html
git commit -m "feat(notes): add safe markdown parser with tests"
```

---

### Task 1.3: Notes CRUD methods in database.js

**Files:**
- Modify: `js/database.js` (add new methods near the bottom, e.g. after the existing `setCashFlowOverride` block)

- [ ] **Step 1: Add CRUD methods**

Append to the `DB` object (find a logical location near other meta/override methods):

```javascript
    // ============== NOTES & ANALYSIS LAYER ==============

    /**
     * Get a single note by composite key.
     * @returns {Object|null} { id, target_type, target_key, body, pinned_to_share, metadata, created_at, updated_at }
     */
    getNote(targetType, targetKey) {
        const result = this.db.exec(
            'SELECT id, target_type, target_key, body, pinned_to_share, metadata, created_at, updated_at ' +
            'FROM notes WHERE target_type = ? AND target_key = ?',
            [targetType, targetKey]
        );
        if (!result.length || !result[0].values.length) return null;
        const r = result[0].values[0];
        return {
            id: r[0], target_type: r[1], target_key: r[2], body: r[3],
            pinned_to_share: r[4], metadata: r[5], created_at: r[6], updated_at: r[7]
        };
    },

    /**
     * Upsert a note. Empty body deletes the row.
     */
    setNote(targetType, targetKey, body, metadata = null) {
        const trimmed = (body || '').trim();
        if (!trimmed) {
            this.deleteNote(targetType, targetKey);
            return;
        }
        const existing = this.getNote(targetType, targetKey);
        if (existing) {
            this.db.run(
                'UPDATE notes SET body = ?, metadata = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?',
                [trimmed, metadata, existing.id]
            );
        } else {
            this.db.run(
                'INSERT INTO notes (target_type, target_key, body, metadata) VALUES (?, ?, ?, ?)',
                [targetType, targetKey, trimmed, metadata]
            );
        }
        this.autoSave();
    },

    deleteNote(targetType, targetKey) {
        this.db.run(
            'DELETE FROM notes WHERE target_type = ? AND target_key = ?',
            [targetType, targetKey]
        );
        this.autoSave();
    },

    setNotePinned(targetType, targetKey, pinned) {
        this.db.run(
            'UPDATE notes SET pinned_to_share = ?, updated_at = CURRENT_TIMESTAMP WHERE target_type = ? AND target_key = ?',
            [pinned ? 1 : 0, targetType, targetKey]
        );
        this.autoSave();
    },

    /**
     * Return all period notes for a given (tab, month).
     * @param {string} tab 'pnl' | 'cashflow' | 'balancesheet'
     * @param {string} month 'YYYY-MM'
     * @returns {Object|null}
     */
    getPeriodNote(tab, month) {
        return this.getNote('period', `${tab}:${month}`);
    },

    /**
     * Get pinned period notes for a view-only share, capped and sorted by recency.
     * @param {number} cap Max count (default 5)
     * @returns {Array<Object>} note objects with parsed tab/month from target_key
     */
    getPinnedShareNotes(cap = 5) {
        const result = this.db.exec(
            "SELECT id, target_key, body, updated_at, created_at " +
            "FROM notes WHERE target_type = 'period' AND pinned_to_share = 1 " +
            "ORDER BY updated_at DESC, created_at DESC, id DESC LIMIT ?",
            [cap]
        );
        if (!result.length) return [];
        return result[0].values.map(r => {
            const [tab, month] = r[1].split(':');
            return { id: r[0], target_key: r[1], body: r[2], updated_at: r[3], created_at: r[4], tab, month };
        });
    },
```

- [ ] **Step 2: Smoke-test in console**

Reload `index.html`. In DevTools console:

```javascript
DB.setNote('period', 'pnl:2026-03', 'Test note **bold**');
DB.getPeriodNote('pnl', '2026-03');
// Expected: object with body 'Test note **bold**', pinned_to_share=0
DB.setNotePinned('period', 'pnl:2026-03', true);
DB.getPinnedShareNotes(5);
// Expected: array with 1 item, body matches
DB.deleteNote('period', 'pnl:2026-03');
DB.getPeriodNote('pnl', '2026-03');
// Expected: null
```

- [ ] **Step 3: Commit**

```bash
git add js/database.js
git commit -m "feat(notes): add CRUD methods for notes table"
```

---

### Task 1.4: Period note icon + popover HTML/CSS

**Files:**
- Modify: `index.html` (add popover element near other modals)
- Modify: `css/styles.css` (new section at end of file)

- [ ] **Step 1: Add popover HTML**

In `index.html`, find the modals section (near line ~2000, where other modals like `quickEntryModal` live). Add:

```html
        <!-- Period Note Popover (Notes & Analysis Layer 2026-04-28) -->
        <div id="periodNotePopover" class="period-note-popover" hidden>
            <div class="pn-pop-header" data-drag-handle>
                <div class="pn-pop-titles">
                    <div class="pn-pop-eyebrow" id="periodNotePopEyebrow">P&amp;L Note</div>
                    <div class="pn-pop-month" id="periodNotePopMonth">March 2026</div>
                </div>
                <div class="pn-pop-actions-top">
                    <button type="button" class="pn-pop-pin" id="periodNotePopPin" title="Pin to shared report">
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                            <path d="M12 2v6M5 8l7-6 7 6M19 14H5a2 2 0 0 0-2 2v4h18v-4a2 2 0 0 0-2-2z"/>
                        </svg>
                    </button>
                    <button type="button" class="pn-pop-close" id="periodNotePopClose" aria-label="Close">×</button>
                </div>
            </div>
            <textarea
                id="periodNotePopTextarea"
                class="pn-pop-textarea"
                placeholder="What's worth noting about this month? Supports **bold**, *italic*, and - bullet lists."
            ></textarea>
            <div class="pn-pop-footer">
                <span class="pn-pop-meta" id="periodNotePopMeta">Auto-saves on close</span>
                <button type="button" class="btn btn-primary btn-sm" id="periodNotePopDone">Done</button>
            </div>
            <div class="pn-pop-resize-grip" data-resize-handle></div>
        </div>
```

- [ ] **Step 2: Add CSS section**

Append to `css/styles.css`:

```css
/* ================================================================
   NOTES & ANALYSIS LAYER (2026-04-28)
   ================================================================ */

/* === Period Note icon in column headers === */
.pn-icon-wrap {
    display: inline-flex;
    align-items: center;
    margin-left: 6px;
    cursor: pointer;
    opacity: 0;
    transition: opacity 150ms ease, color 150ms ease;
    color: var(--text-secondary);
    vertical-align: middle;
}
th:hover .pn-icon-wrap,
.pn-icon-wrap.has-note,
.pn-icon-wrap:focus-visible { opacity: 1; }
.pn-icon-wrap.has-note { color: var(--amber); }
.pn-icon-wrap:hover { color: var(--primary); }
.pn-icon-wrap svg { width: 14px; height: 14px; }

/* === Period Note popover === */
.period-note-popover {
    position: fixed;
    width: 360px;
    min-width: 280px;
    min-height: 200px;
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    box-shadow: 0 8px 24px rgba(0,0,0,0.12);
    z-index: 9000;
    padding: 14px 16px 12px;
    display: flex;
    flex-direction: column;
    resize: none;
}
.period-note-popover[hidden] { display: none; }
.pn-pop-header {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    margin-bottom: 8px;
    cursor: move;
    user-select: none;
}
.pn-pop-titles { line-height: 1.2; }
.pn-pop-eyebrow {
    font-size: 10px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    color: var(--text-secondary);
    font-weight: 600;
}
.pn-pop-month { font-size: 14px; font-weight: 600; color: var(--text); }
.pn-pop-actions-top { display: flex; gap: 4px; align-items: center; }
.pn-pop-pin, .pn-pop-close {
    background: transparent; border: none; cursor: pointer;
    color: var(--text-secondary); padding: 4px;
    border-radius: 4px;
}
.pn-pop-pin:hover, .pn-pop-close:hover { background: var(--surface-alt); color: var(--text); }
.pn-pop-pin.active { color: var(--amber); }
.pn-pop-close { font-size: 18px; line-height: 1; padding: 0 6px; }
.pn-pop-textarea {
    flex: 1;
    width: 100%;
    min-height: 90px;
    padding: 10px;
    border: 1px solid var(--border);
    border-radius: var(--radius-sm);
    font-family: inherit;
    font-size: 13px;
    resize: none;
    box-sizing: border-box;
}
.pn-pop-textarea:focus { outline: none; box-shadow: var(--shadow-focus); }
.pn-pop-footer {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-top: 10px;
}
.pn-pop-meta { font-size: 11px; color: var(--text-muted); }
.pn-pop-resize-grip {
    position: absolute;
    bottom: 0; right: 0;
    width: 14px; height: 14px;
    cursor: nwse-resize;
    background: linear-gradient(135deg, transparent 50%, var(--border) 50%, var(--border) 60%, transparent 60%, transparent 70%, var(--border) 70%, var(--border) 80%, transparent 80%);
}

/* === Period Note footnote in view-only === */
.pn-footnote {
    background: var(--amber-bg);
    border-left: 3px solid var(--amber);
    padding: 10px 14px;
    margin: 8px 16px;
    font-size: 12px;
    color: #92400e;
    border-radius: 0 6px 6px 0;
    line-height: 1.5;
}
[data-theme="dark"] .pn-footnote { color: #fde68a; background: rgba(217,119,6,0.12); }
.pn-footnote strong { color: #78350f; font-weight: 600; }
[data-theme="dark"] .pn-footnote strong { color: #fef3c7; }
.pn-footnote ul { margin: 4px 0 0 0; padding-left: 18px; }
.pn-footnote em { font-style: italic; }

.btn-sm { padding: 4px 10px !important; font-size: 12px !important; }
```

- [ ] **Step 3: Smoke-test the popover renders**

Reload `index.html`. In DevTools console:

```javascript
const pop = document.getElementById('periodNotePopover');
pop.removeAttribute('hidden');
pop.style.top = '120px'; pop.style.left = '120px';
```

Expected: popover appears at top-left with header, textarea, pin button, close button, footer.

- [ ] **Step 4: Hide it again**

```javascript
pop.setAttribute('hidden', '');
```

- [ ] **Step 5: Commit**

```bash
git add index.html css/styles.css
git commit -m "feat(notes): add period note popover HTML and CSS"
```

---

### Task 1.5: Period note popover controller (open/save/close)

**Files:**
- Modify: `js/app.js` (add new controller block)

- [ ] **Step 1: Add controller methods**

Add to the App object in `js/app.js`. Place near other modal controllers (search for `openCategoryModal` to find a good neighbor):

```javascript
    // ============== PERIOD NOTES (Notes & Analysis Layer) ==============

    _periodNoteState: {
        tab: null,
        month: null,
        currentBody: '',
        anchorEl: null
    },

    /**
     * Open the period note popover anchored under a column header.
     * @param {string} tab 'pnl' | 'cashflow' | 'balancesheet'
     * @param {string} month 'YYYY-MM'
     * @param {HTMLElement} anchorEl The column header element to anchor to
     */
    openPeriodNote(tab, month, anchorEl) {
        const pop = document.getElementById('periodNotePopover');
        const eyebrow = document.getElementById('periodNotePopEyebrow');
        const monthEl = document.getElementById('periodNotePopMonth');
        const textarea = document.getElementById('periodNotePopTextarea');
        const pin = document.getElementById('periodNotePopPin');
        const meta = document.getElementById('periodNotePopMeta');

        const tabLabels = { pnl: 'P&L Note', cashflow: 'Cash Flow Note', balancesheet: 'Balance Sheet Note' };
        eyebrow.textContent = tabLabels[tab] || 'Note';
        monthEl.textContent = Utils.formatMonthDisplay(month);

        const existing = DB.getPeriodNote(tab, month);
        textarea.value = existing ? existing.body : '';
        pin.classList.toggle('active', !!(existing && existing.pinned_to_share));
        meta.textContent = existing
            ? `Last edited ${Utils.formatDate(existing.updated_at)} · auto-saves on close`
            : 'New note — auto-saves on close';

        this._periodNoteState = { tab, month, currentBody: textarea.value, anchorEl };

        // Position: below anchor if provided, else center
        pop.removeAttribute('hidden');
        pop.style.width = ''; pop.style.height = '';  // reset any prior drag/resize
        if (anchorEl) {
            const rect = anchorEl.getBoundingClientRect();
            const popW = 360;
            let left = rect.left + (rect.width / 2) - (popW / 2);
            const top = rect.bottom + 8;
            // Clamp to viewport
            left = Math.max(8, Math.min(left, window.innerWidth - popW - 8));
            pop.style.left = left + 'px';
            pop.style.top = top + 'px';
        } else {
            pop.style.left = (window.innerWidth / 2 - 180) + 'px';
            pop.style.top = '120px';
        }
        textarea.focus();
    },

    closePeriodNote() {
        const pop = document.getElementById('periodNotePopover');
        const textarea = document.getElementById('periodNotePopTextarea');
        const { tab, month, currentBody } = this._periodNoteState;
        if (tab && month) {
            const newBody = textarea.value;
            if (newBody !== currentBody) {
                DB.setNote('period', `${tab}:${month}`, newBody);
                // Refresh whichever spreadsheet is showing so icon flips state
                if (this.currentTab === 'pnl') this.refreshPnL();
                else if (this.currentTab === 'cashflow') this.refreshCashFlow();
                else if (this.currentTab === 'balancesheet') this.refreshBalanceSheet();
            }
        }
        pop.setAttribute('hidden', '');
        this._periodNoteState = { tab: null, month: null, currentBody: '', anchorEl: null };
    },

    togglePeriodNotePin() {
        const { tab, month } = this._periodNoteState;
        if (!tab || !month) return;
        const textarea = document.getElementById('periodNotePopTextarea');
        const pin = document.getElementById('periodNotePopPin');
        // Save body first if changed (so the row exists)
        DB.setNote('period', `${tab}:${month}`, textarea.value);
        const existing = DB.getPeriodNote(tab, month);
        if (!existing) return; // empty body → no pinning
        const newPinned = !existing.pinned_to_share;
        DB.setNotePinned('period', `${tab}:${month}`, newPinned);
        pin.classList.toggle('active', newPinned);
        this._periodNoteState.currentBody = textarea.value;  // prevent re-save on close
    },
```

- [ ] **Step 2: Wire up event listeners**

Find `setupEventListeners()` in `js/app.js` (around line 2206). Add at the end:

```javascript
        // Period Note popover
        document.getElementById('periodNotePopClose').addEventListener('click', () => this.closePeriodNote());
        document.getElementById('periodNotePopDone').addEventListener('click', () => this.closePeriodNote());
        document.getElementById('periodNotePopPin').addEventListener('click', () => this.togglePeriodNotePin());
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape' && !document.getElementById('periodNotePopover').hidden) {
                this.closePeriodNote();
            }
        });
```

- [ ] **Step 3: Smoke-test open/save/close**

Reload `index.html`. In DevTools console:

```javascript
App.openPeriodNote('pnl', '2026-03', null);
// Popover should appear at center; type "Test note" in textarea
// Click "Done" or press Escape → popover closes
DB.getPeriodNote('pnl', '2026-03');
// Expected: object with body "Test note"
DB.deleteNote('period', 'pnl:2026-03');  // cleanup
```

- [ ] **Step 4: Commit**

```bash
git add js/app.js
git commit -m "feat(notes): add period note popover open/save/close controller"
```

---

### Task 1.6: Drag and resize handlers for popover

**Files:**
- Modify: `js/app.js` (add to controller block from Task 1.5)

- [ ] **Step 1: Add drag and resize logic**

Add to the period-note controller block in `js/app.js`:

```javascript
    _setupPeriodNoteDragResize() {
        const pop = document.getElementById('periodNotePopover');
        const handle = pop.querySelector('[data-drag-handle]');
        const grip = pop.querySelector('[data-resize-handle]');

        // === Drag ===
        let dragStart = null;
        handle.addEventListener('mousedown', (e) => {
            // Don't start drag if clicking buttons inside header
            if (e.target.closest('button')) return;
            const rect = pop.getBoundingClientRect();
            dragStart = { x: e.clientX, y: e.clientY, left: rect.left, top: rect.top };
            e.preventDefault();
        });
        document.addEventListener('mousemove', (e) => {
            if (!dragStart) return;
            const dx = e.clientX - dragStart.x;
            const dy = e.clientY - dragStart.y;
            const newLeft = Math.max(0, Math.min(window.innerWidth - 100, dragStart.left + dx));
            const newTop = Math.max(0, Math.min(window.innerHeight - 50, dragStart.top + dy));
            pop.style.left = newLeft + 'px';
            pop.style.top = newTop + 'px';
        });
        document.addEventListener('mouseup', () => { dragStart = null; });

        // === Resize ===
        let resizeStart = null;
        grip.addEventListener('mousedown', (e) => {
            const rect = pop.getBoundingClientRect();
            resizeStart = { x: e.clientX, y: e.clientY, w: rect.width, h: rect.height };
            e.preventDefault();
            e.stopPropagation();
        });
        document.addEventListener('mousemove', (e) => {
            if (!resizeStart) return;
            const dw = e.clientX - resizeStart.x;
            const dh = e.clientY - resizeStart.y;
            pop.style.width = Math.max(280, resizeStart.w + dw) + 'px';
            pop.style.height = Math.max(200, resizeStart.h + dh) + 'px';
        });
        document.addEventListener('mouseup', () => { resizeStart = null; });
    },
```

- [ ] **Step 2: Call setup once at init**

Find `init()` in `js/app.js` (around line 200-300, where other one-time setups live). Add inside it:

```javascript
        this._setupPeriodNoteDragResize();
```

- [ ] **Step 3: Smoke-test drag and resize**

Reload `index.html`. Open popover via console: `App.openPeriodNote('pnl','2026-03',null)`. Drag the header → popover follows mouse. Drag the bottom-right grip → popover resizes. Close. Cleanup: `DB.deleteNote('period','pnl:2026-03')`.

- [ ] **Step 4: Commit**

```bash
git add js/app.js
git commit -m "feat(notes): add drag and resize handlers for period note popover"
```

---

### Task 1.7: Render note icon in P&L spreadsheet column headers

**Files:**
- Modify: `js/ui.js` (modify `renderProfitLossSpreadsheet`)

- [ ] **Step 1: Add a helper function for the note icon**

Add to the `UI` object in `js/ui.js`, near the top:

```javascript
    /**
     * Render the period-note icon HTML for a spreadsheet column header.
     * @param {string} tab 'pnl' | 'cashflow' | 'balancesheet'
     * @param {string} month 'YYYY-MM'
     * @returns {string} HTML
     */
    renderNoteIcon(tab, month) {
        const note = DB.getPeriodNote(tab, month);
        const has = note && note.body ? 'has-note' : '';
        const filled = has
            ? '<line x1="9" y1="13" x2="15" y2="13"/><line x1="9" y1="17" x2="13" y2="17"/>'
            : '';
        return `<span class="pn-icon-wrap ${has}" data-pn-tab="${tab}" data-pn-month="${month}" title="${has ? 'Edit note' : 'Add note for this month'}">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                <polyline points="14 2 14 8 20 8"/>
                ${filled}
            </svg>
        </span>`;
    },
```

- [ ] **Step 2: Modify `renderProfitLossSpreadsheet` to inject the icon**

Find `renderProfitLossSpreadsheet` in `js/ui.js` (around line 1055). Locate where month column headers are built — typically a loop building `<th>` cells per month. Modify the month-header cell to include the icon. Example transform (the exact existing code may differ — preserve the structure, only add the icon HTML inside the `<th>`):

Find the existing pattern that looks roughly like:

```javascript
html += `<th>${Utils.formatMonthDisplay(month)}</th>`;
```

Replace with:

```javascript
html += `<th>${Utils.formatMonthDisplay(month)} ${this.renderNoteIcon('pnl', month)}</th>`;
```

If the existing code uses a different escape helper or wrapper, preserve it but inject `${this.renderNoteIcon('pnl', month)}` after the formatted month label, inside the `<th>`.

- [ ] **Step 3: Wire click handler**

In `js/app.js`, find `setupEventListeners()` and add a delegated click handler (after the existing tab-switch listeners):

```javascript
        // Period note icon clicks (delegated; works for any rendered icon)
        document.addEventListener('click', (e) => {
            const wrap = e.target.closest('.pn-icon-wrap');
            if (!wrap) return;
            e.stopPropagation();
            const tab = wrap.getAttribute('data-pn-tab');
            const month = wrap.getAttribute('data-pn-month');
            if (tab && month) this.openPeriodNote(tab, month, wrap);
        });
```

- [ ] **Step 4: Smoke-test**

Reload `index.html`. Switch to P&L tab. Hover any month column header → note icon fades in. Click the icon → popover opens anchored below header. Type "Hello world" → click Done → popover closes, icon turns amber (filled). Click the icon again → popover reopens with "Hello world" in textarea. Clear textarea, click Done → icon returns to outline.

- [ ] **Step 5: Commit**

```bash
git add js/ui.js js/app.js
git commit -m "feat(notes): wire period note icons in P&L column headers"
```

---

### Task 1.8: Same wiring for Cash Flow and Balance Sheet

**Files:**
- Modify: `js/ui.js` (`renderCashFlowSpreadsheet`, `renderBalanceSheet`)

- [ ] **Step 1: Inject icon in Cash Flow headers**

Find `renderCashFlowSpreadsheet` in `js/ui.js` (around line 674). Locate month column header rendering. Inject:

```javascript
${this.renderNoteIcon('cashflow', month)}
```

inside the appropriate `<th>` element next to the formatted month label.

- [ ] **Step 2: Inject icon in Balance Sheet headers**

Find `renderBalanceSheet` in `js/ui.js` (around line 1443). Balance Sheet typically has a single "as of" month indicator rather than multiple columns. Inject the icon next to that "as of" label using:

```javascript
${this.renderNoteIcon('balancesheet', asOfMonth)}
```

If Balance Sheet has multiple month columns, treat each like P&L.

- [ ] **Step 3: Smoke-test all three tabs**

Reload `index.html`. Test each tab:
- **P&L:** add a note "PnL March test" → icon turns amber
- **Cash Flow:** add a note "CF March test" → icon turns amber, **PnL note still says only "PnL March test"** (per-tab scope verified)
- **Balance Sheet:** add a note "BS test" → icon turns amber

Verify in console:
```javascript
DB.getPeriodNote('pnl','2026-03').body         // 'PnL March test'
DB.getPeriodNote('cashflow','2026-03').body    // 'CF March test'
```

Cleanup:
```javascript
DB.deleteNote('period','pnl:2026-03');
DB.deleteNote('period','cashflow:2026-03');
DB.deleteNote('period','balancesheet:2026-03');
```

- [ ] **Step 4: Commit**

```bash
git add js/ui.js
git commit -m "feat(notes): wire period note icons in Cash Flow and Balance Sheet headers"
```

---

### Task 1.9: Render footnotes in view-only mode

**Files:**
- Modify: `js/ui.js` (modify the three spreadsheet renderers to append footnotes block)
- Modify: `js/app.js` (add footnote-render call to view-only restrictions)

- [ ] **Step 1: Add footnote rendering helper to UI**

Add to `js/ui.js`:

```javascript
    /**
     * Build view-only footnote HTML for pinned period notes.
     * @param {string} tab Optionally filter to one tab; if null, render all
     * @param {number} cap
     * @returns {string} HTML block (empty string if no notes)
     */
    renderPeriodNoteFootnotes(tab = null, cap = 5) {
        const cap_ = parseInt(DB.getMetaValue && DB.getMetaValue('pinned_notes_share_cap')) || cap;
        let notes = DB.getPinnedShareNotes(cap_);
        if (tab) notes = notes.filter(n => n.tab === tab);
        if (!notes.length) return '';
        const items = notes.map(n => {
            const tabLabel = { pnl: 'P&L', cashflow: 'Cash Flow', balancesheet: 'Balance Sheet' }[n.tab] || n.tab;
            const monthLabel = Utils.formatMonthDisplay(n.month);
            const bodyHtml = Utils.renderMarkdownSafe(n.body);
            return `<div class="pn-footnote">
                <strong>📝 ${tabLabel} · ${monthLabel}:</strong> ${bodyHtml}
            </div>`;
        });
        return `<div class="pn-footnotes-block">${items.join('')}</div>`;
    },
```

- [ ] **Step 2: Add `getMetaValue` helper to `database.js` if not already present**

Check if `DB.getMetaValue` exists. If not, add to `js/database.js`:

```javascript
    getMetaValue(key) {
        const r = this.db.exec('SELECT value FROM app_meta WHERE key = ?', [key]);
        if (!r.length || !r[0].values.length) return null;
        return r[0].values[0][0];
    },
    setMetaValue(key, value) {
        this.db.run('INSERT OR REPLACE INTO app_meta (key, value) VALUES (?, ?)', [key, String(value)]);
        this.autoSave();
    },
```

- [ ] **Step 3: Append footnotes after each spreadsheet in view-only mode**

In each of the three renderers (`renderProfitLossSpreadsheet`, `renderCashFlowSpreadsheet`, `renderBalanceSheet`), find the end of the rendered HTML where the spreadsheet table is closed. Append (only when in view-only mode):

```javascript
if (App.isViewOnly) {
    const tabKey = 'pnl'; // or 'cashflow' / 'balancesheet' per renderer
    html += this.renderPeriodNoteFootnotes(tabKey);
}
```

- [ ] **Step 4: Smoke-test view-only mode**

Add a pinned note: in console, run

```javascript
DB.setNote('period','pnl:2026-03','Test footnote — **bold** text');
DB.setNotePinned('period','pnl:2026-03', true);
App.isViewOnly = true; App.refreshPnL();
```

Expected: P&L renders with an amber footnote at the bottom: "📝 P&L · March 2026: Test footnote — **bold** text" (with **bold** rendered as `<strong>`).

Reset: `App.isViewOnly = false; DB.deleteNote('period','pnl:2026-03'); App.refreshPnL();`

- [ ] **Step 5: Commit**

```bash
git add js/ui.js js/database.js js/app.js
git commit -m "feat(notes): render pinned period notes as footnotes in view-only mode"
```

---

### Task 1.10: Pinned-notes cap warning in share modal

**Files:**
- Modify: `js/app.js` (modify `generateShare` or equivalent share-build function)
- Modify: `index.html` (add warning element to share modal if not present)

- [ ] **Step 1: Add warning span to share modal**

Find `shareModal` in `index.html` (around line 2604). Add inside the modal body, above the action buttons:

```html
<div id="shareNotesWarning" class="share-notes-warning" hidden></div>
```

- [ ] **Step 2: CSS for warning**

Append to `css/styles.css` Notes section:

```css
.share-notes-warning {
    background: var(--amber-bg);
    border: 1px solid var(--amber);
    color: #92400e;
    padding: 8px 12px;
    border-radius: var(--radius-sm);
    font-size: 12px;
    margin: 8px 0;
}
```

- [ ] **Step 3: Wire warning when share modal opens**

Find the share-modal-opening function in `js/app.js` (search for `shareModal` or `openShare`). Add inside it:

```javascript
        // Pinned notes cap warning
        const cap = parseInt(DB.getMetaValue('pinned_notes_share_cap')) || 5;
        const r = DB.db.exec("SELECT COUNT(*) FROM notes WHERE target_type='period' AND pinned_to_share=1");
        const pinnedCount = r.length ? r[0].values[0][0] : 0;
        const warn = document.getElementById('shareNotesWarning');
        if (pinnedCount > cap) {
            warn.textContent = `You have ${pinnedCount} pinned notes — only the most recent ${cap} will appear in this share. Edit notes or raise the cap in Settings to include more.`;
            warn.removeAttribute('hidden');
        } else {
            warn.setAttribute('hidden', '');
        }
```

- [ ] **Step 4: Smoke-test**

Pin 6 notes:
```javascript
for (let m=1; m<=6; m++) {
    const mm = String(m).padStart(2,'0');
    DB.setNote('period', `pnl:2026-${mm}`, `Note ${m}`);
    DB.setNotePinned('period', `pnl:2026-${mm}`, true);
}
```
Open share modal → warning appears: "You have 6 pinned notes — only the most recent 5 will appear..." Cleanup pinned notes after.

- [ ] **Step 5: Commit**

```bash
git add index.html css/styles.css js/app.js
git commit -m "feat(notes): warn in share modal when pinned notes exceed cap"
```

---

### Task 1.11: Settings input for pinned-notes cap and variance threshold (placeholder)

**Files:**
- Modify: `index.html` (gear/settings popover)
- Modify: `js/app.js`

- [ ] **Step 1: Add settings inputs**

Find the gear/settings popover in `index.html` (search for `gear` or `themePreset`). Add a new section:

```html
<div class="settings-group">
    <label class="settings-label">Pinned notes cap (view-only shares)</label>
    <input type="number" id="settingPinnedCap" min="1" max="20" value="5" class="settings-input">
</div>
<div class="settings-group">
    <label class="settings-label">Variance flag threshold (%)</label>
    <input type="number" id="settingVarianceThreshold" min="5" max="200" value="20" class="settings-input">
</div>
```

- [ ] **Step 2: CSS**

Append to `css/styles.css` Notes section:

```css
.settings-group { display: flex; flex-direction: column; gap: 4px; padding: 6px 0; }
.settings-label { font-size: 12px; color: var(--text-secondary); }
.settings-input { padding: 6px 8px; border: 1px solid var(--border); border-radius: 6px; font-size: 13px; max-width: 100px; }
```

- [ ] **Step 3: Load + save handlers**

In `js/app.js` `init()` (or wherever theme/timeline is loaded), add:

```javascript
        // Notes & analysis settings
        const pinnedCap = parseInt(DB.getMetaValue('pinned_notes_share_cap')) || 5;
        const varThreshold = parseInt(DB.getMetaValue('variance_threshold_pct')) || 20;
        document.getElementById('settingPinnedCap').value = pinnedCap;
        document.getElementById('settingVarianceThreshold').value = varThreshold;
        document.getElementById('settingPinnedCap').addEventListener('change', (e) => {
            DB.setMetaValue('pinned_notes_share_cap', e.target.value);
        });
        document.getElementById('settingVarianceThreshold').addEventListener('change', (e) => {
            DB.setMetaValue('variance_threshold_pct', e.target.value);
            // Re-render flagged tabs if open
            if (this.currentTab === 'pnl') this.refreshPnL();
            else if (this.currentTab === 'cashflow') this.refreshCashFlow();
        });
```

- [ ] **Step 4: Smoke-test persistence**

Reload `index.html`. Open settings, change cap to 7 → close + reopen settings → input still shows 7. Refresh page → input still shows 7.

- [ ] **Step 5: Commit**

```bash
git add index.html css/styles.css js/app.js
git commit -m "feat(notes): add settings inputs for pinned cap and variance threshold"
```

---

### Phase 1 Complete — Manual Acceptance Test

- [ ] **Verify Phase 1 end-to-end:**
1. Reload app, switch to P&L
2. Hover March column header → icon fades in
3. Click icon → popover opens below header, anchored
4. Type `**Big** spike — *one-time* deal\n- Westbrook Co.\n- $18K wholesale` → click Done
5. Icon turns amber/filled
6. Click pin icon → it activates
7. Switch to view-only mode (`App.isViewOnly=true; App.refreshPnL()`) → footnote appears with bold/italic/list rendered correctly
8. Switch back (`App.isViewOnly=false; App.refreshPnL()`) → footnote gone
9. Reload page → note persists, pin persists
10. Drag popover header → moves; resize grip → resizes
11. Same note for Cash Flow March is independent (per-tab scope)

---

## Phase 2: Variance Auto-Flagging

### Task 2.1: `computeVariance` helper + tests

**Files:**
- Modify: `js/utils.js`
- Modify: `tests/notes-analysis.test.html`

- [ ] **Step 1: Add tests for variance**

In `tests/notes-analysis.test.html` (before the summary line), append:

```javascript
const s2 = suite('computeVariance');
test(s2, '142% jump vs prior month flags', () => {
  const v = Utils.computeVariance([2100, 2400, 5800], 2, 20);
  assert(v.flagged, 'should flag');
  assert(Math.abs(v.deltaPrior - 1.4166666) < 0.01, 'deltaPrior ≈ 142%');
});
test(s2, '5% change does not flag at 20% threshold', () => {
  const v = Utils.computeVariance([1000, 1050], 1, 20);
  assert(!v.flagged, 'should not flag');
});
test(s2, 'first month never flags (no prior)', () => {
  const v = Utils.computeVariance([5000], 0, 20);
  assert(!v.flagged, 'first month no flag');
  assertEq(v.deltaPrior, null);
});
test(s2, 'flags when 3-mo avg delta exceeds even if prior is close', () => {
  const v = Utils.computeVariance([100, 100, 100, 130], 3, 20);
  assert(v.flagged, '+30% vs avg flags');
});
test(s2, 'zero prior month with non-zero current flags', () => {
  const v = Utils.computeVariance([0, 500], 1, 20);
  assert(v.flagged, 'going from 0 to 500 should flag');
});
test(s2, 'all zero values do not flag', () => {
  const v = Utils.computeVariance([0, 0, 0], 2, 20);
  assert(!v.flagged, 'no flag for all zero');
});
```

- [ ] **Step 2: Run failing tests**

Open `tests/notes-analysis.test.html` in browser. Expected: 6 new tests fail with "Utils.computeVariance is not a function."

- [ ] **Step 3: Implement `computeVariance`**

Append to `Utils` in `js/utils.js`:

```javascript
    /**
     * Compute variance flag for a value at index in a series of monthly values.
     * @param {Array<number>} values Monthly values (indexed in order)
     * @param {number} idx Current index
     * @param {number} thresholdPct Threshold (e.g. 20 for 20%)
     * @returns {Object} { flagged, deltaPrior, deltaAvg3, current }
     */
    computeVariance(values, idx, thresholdPct) {
        const out = { flagged: false, deltaPrior: null, deltaAvg3: null, current: values[idx] };
        if (idx < 1) return out;
        const cur = Number(values[idx]) || 0;
        const prior = Number(values[idx - 1]) || 0;
        // delta vs prior
        if (prior !== 0) {
            out.deltaPrior = (cur - prior) / Math.abs(prior);
        } else if (cur !== 0) {
            out.deltaPrior = Infinity;
        }
        // delta vs 3-mo avg (uses prior 3 months — exclude current)
        if (idx >= 3) {
            const slice = values.slice(idx - 3, idx).map(Number);
            const avg = slice.reduce((a, b) => a + b, 0) / 3;
            if (avg !== 0) {
                out.deltaAvg3 = (cur - avg) / Math.abs(avg);
            } else if (cur !== 0) {
                out.deltaAvg3 = Infinity;
            }
        }
        const t = thresholdPct / 100;
        if ((out.deltaPrior !== null && Math.abs(out.deltaPrior) >= t) ||
            (out.deltaAvg3 !== null && Math.abs(out.deltaAvg3) >= t)) {
            out.flagged = true;
        }
        return out;
    },
```

- [ ] **Step 4: Re-run tests**

Reload `tests/notes-analysis.test.html`. Expected: all 12 tests passing (6 markdown + 6 variance).

- [ ] **Step 5: Commit**

```bash
git add js/utils.js tests/notes-analysis.test.html
git commit -m "feat(variance): add computeVariance helper with tests"
```

---

### Task 2.2: Variance popover HTML/CSS

**Files:**
- Modify: `index.html`
- Modify: `css/styles.css`

- [ ] **Step 1: Add popover HTML**

In `index.html`, near the `periodNotePopover`, add:

```html
        <!-- Variance Popover (Notes & Analysis Layer 2026-04-28) -->
        <div id="variancePopover" class="variance-popover" hidden>
            <div class="vr-pop-header">
                <div class="vr-pop-titles">
                    <div class="vr-pop-eyebrow">⚡ Variance Flag</div>
                    <div class="vr-pop-target" id="variancePopTarget">—</div>
                </div>
                <button type="button" class="pn-pop-close" id="variancePopClose" aria-label="Close">×</button>
            </div>
            <div class="vr-pop-stats" id="variancePopStats"></div>
            <div class="vr-pop-actions">
                <button type="button" class="btn btn-text" id="variancePopDismiss">Dismiss flag</button>
                <button type="button" class="btn btn-primary btn-sm" id="variancePopExplain">Explain this →</button>
            </div>
        </div>
```

- [ ] **Step 2: CSS**

Append to `css/styles.css` Notes section:

```css
/* === Variance flag === */
.cell-flagged {
    background: var(--amber-bg) !important;
    position: relative;
}
[data-theme="dark"] .cell-flagged { background: rgba(217,119,6,0.12) !important; }
.cell-flagged::after {
    content: ""; position: absolute; top: 4px; right: 4px;
    width: 6px; height: 6px; border-radius: 50%; background: var(--amber);
}
.cell-flagged { cursor: pointer; }

/* === Variance popover === */
.variance-popover {
    position: fixed;
    width: 360px;
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    box-shadow: 0 8px 24px rgba(0,0,0,0.12);
    z-index: 9000;
    padding: 14px 16px 12px;
}
.variance-popover[hidden] { display: none; }
.vr-pop-header {
    display: flex; justify-content: space-between; align-items: flex-start;
    margin-bottom: 8px;
}
.vr-pop-eyebrow { font-size: 10px; text-transform: uppercase; letter-spacing: 0.5px; color: var(--text-secondary); font-weight: 600; }
.vr-pop-target { font-size: 13px; font-weight: 600; color: var(--text); }
.vr-pop-stats > div {
    display: flex; justify-content: space-between;
    padding: 6px 0; border-bottom: 1px solid var(--border-light);
    font-size: 12px;
}
.vr-pop-stats > div:last-child { border: none; }
.vr-delta-up { color: var(--amber); font-weight: 600; }
.vr-delta-down { color: var(--green); font-weight: 600; }
.vr-pop-actions {
    display: flex; justify-content: space-between; align-items: center;
    margin-top: 10px;
}
```

- [ ] **Step 3: Smoke-test renders**

Reload. Console:
```javascript
const p = document.getElementById('variancePopover');
p.style.left = '200px'; p.style.top = '200px';
p.removeAttribute('hidden');
```
Expected: popover with eyebrow, target line, empty stats, two buttons. Hide: `p.setAttribute('hidden','');`

- [ ] **Step 4: Commit**

```bash
git add index.html css/styles.css
git commit -m "feat(variance): add variance flag CSS and popover HTML"
```

---

### Task 2.3: Apply variance flags in P&L and Cash Flow rendering

**Files:**
- Modify: `js/ui.js`

- [ ] **Step 1: In P&L renderer, compute and apply flag class**

Find `renderProfitLossSpreadsheet` in `js/ui.js`. Locate the loop that builds value cells per (category, month). For each row, you have access to the array of monthly values. For each cell at index `i`:

```javascript
const threshold = parseInt(DB.getMetaValue('variance_threshold_pct')) || 20;
const variance = Utils.computeVariance(monthlyValues, i, threshold);
// Check dismissal
const dismissedRaw = DB.getMetaValue('dismissed_variances') || '[]';
let dismissed = [];
try { dismissed = JSON.parse(dismissedRaw); } catch(e) {}
const targetKey = `pnl:${categoryId}:${month}`;
const isDismissed = dismissed.some(d => d.target_key === targetKey && d.value === monthlyValues[i]);
const cellClass = (variance.flagged && !isDismissed) ? 'cell-flagged' : '';
const dataAttrs = (variance.flagged && !isDismissed)
    ? `data-vr-tab="pnl" data-vr-cat="${categoryId}" data-vr-month="${month}" data-vr-value="${monthlyValues[i]}"`
    : '';

html += `<td class="${cellClass}" ${dataAttrs} title="${variance.flagged ? formatVarianceTooltip(variance) : ''}">${formatted}</td>`;
```

Add a small helper near the bottom of the file:

```javascript
function formatVarianceTooltip(v) {
    if (!v.flagged) return '';
    const pieces = [];
    if (v.deltaPrior !== null && isFinite(v.deltaPrior)) {
        pieces.push(`${v.deltaPrior >= 0 ? '+' : ''}${(v.deltaPrior * 100).toFixed(0)}% vs prior`);
    }
    if (v.deltaAvg3 !== null && isFinite(v.deltaAvg3)) {
        pieces.push(`${v.deltaAvg3 >= 0 ? '+' : ''}${(v.deltaAvg3 * 100).toFixed(0)}% vs 3-mo avg`);
    }
    return pieces.join(' · ');
}
```

- [ ] **Step 2: Same in Cash Flow renderer**

Find `renderCashFlowSpreadsheet`. Apply identical logic with `tab='cashflow'` in the `targetKey` and `data-vr-tab` attribute.

- [ ] **Step 3: Hide flags in view-only**

Wrap the flag application:

```javascript
const showFlags = !App.isViewOnly;
const cellClass = (showFlags && variance.flagged && !isDismissed) ? 'cell-flagged' : '';
```

- [ ] **Step 4: Smoke-test**

Reload. Switch to P&L. Look at any row where a month value is >20% different from the prior month. The cell should have an amber tint + corner dot. Hover → tooltip shows "+X% vs prior · +Y% vs 3-mo avg." Switch to view-only → flags disappear.

- [ ] **Step 5: Commit**

```bash
git add js/ui.js
git commit -m "feat(variance): apply flag styling to P&L and Cash Flow cells above threshold"
```

---

### Task 2.4: Variance popover controller

**Files:**
- Modify: `js/app.js`

- [ ] **Step 1: Add controller**

Append to `App` in `js/app.js` (after the period-note controller block):

```javascript
    // ============== VARIANCE POPOVER ==============

    _varianceState: { tab: null, categoryId: null, month: null, value: 0 },

    openVariancePopover(tab, categoryId, month, value, anchorEl) {
        const pop = document.getElementById('variancePopover');
        const target = document.getElementById('variancePopTarget');
        const stats = document.getElementById('variancePopStats');

        const cat = DB.db.exec('SELECT name FROM categories WHERE id = ?', [categoryId]);
        const catName = cat.length && cat[0].values.length ? cat[0].values[0][0] : '(unknown)';
        target.textContent = `${catName} · ${Utils.formatMonthDisplay(month)}`;

        // Compute deltas — need to fetch the monthly series
        const series = this._getCategoryMonthlySeries(tab, categoryId, month);
        const idx = series.months.indexOf(month);
        const threshold = parseInt(DB.getMetaValue('variance_threshold_pct')) || 20;
        const v = Utils.computeVariance(series.values, idx, threshold);

        const fmt = (d) => {
            if (d === null) return '<span style="color:#9ca3af">no data</span>';
            if (!isFinite(d)) return '<span class="vr-delta-up">∞</span>';
            const cls = d >= 0 ? 'vr-delta-up' : 'vr-delta-down';
            return `<span class="${cls}">${d >= 0 ? '+' : ''}${(d * 100).toFixed(0)}%</span>`;
        };
        stats.innerHTML = `
            <div><span>This month</span><strong>${Utils.formatCurrency(value)}</strong></div>
            <div><span>vs prior month</span>${fmt(v.deltaPrior)}</div>
            <div><span>vs 3-mo avg</span>${fmt(v.deltaAvg3)}</div>
        `;

        this._varianceState = { tab, categoryId, month, value };

        // Position
        pop.removeAttribute('hidden');
        if (anchorEl) {
            const r = anchorEl.getBoundingClientRect();
            const popW = 360;
            let left = r.left + (r.width / 2) - (popW / 2);
            const top = r.bottom + 8;
            left = Math.max(8, Math.min(left, window.innerWidth - popW - 8));
            pop.style.left = left + 'px';
            pop.style.top = top + 'px';
        }
    },

    _getCategoryMonthlySeries(tab, categoryId, currentMonth) {
        // Use existing rendering data sources to get the monthly series
        // for this category. P&L and Cash Flow expose this through their
        // existing data layer functions.
        const months = (function() {
            const tl = DB.getTimeline();
            const list = [];
            let [y, m] = tl.start.split('-').map(Number);
            const [endY, endM] = tl.end.split('-').map(Number);
            while (y < endY || (y === endY && m <= endM)) {
                list.push(`${y}-${String(m).padStart(2,'0')}`);
                m++; if (m > 12) { m = 1; y++; }
            }
            return list;
        })();
        const values = months.map(m => {
            // Sum the category's amount for this month
            const r = DB.db.exec(
                "SELECT COALESCE(SUM(amount), 0) FROM transactions " +
                "WHERE category_id = ? AND " +
                (tab === 'pnl' ? "month_due = ?" : "month_paid = ?"),
                [categoryId, m]
            );
            return r.length ? Number(r[0].values[0][0]) : 0;
        });
        return { months, values };
    },

    closeVariancePopover() {
        document.getElementById('variancePopover').setAttribute('hidden', '');
        this._varianceState = { tab: null, categoryId: null, month: null, value: 0 };
    },

    dismissCurrentVariance() {
        const { tab, categoryId, month, value } = this._varianceState;
        if (!tab) return;
        const targetKey = `${tab}:${categoryId}:${month}`;
        const raw = DB.getMetaValue('dismissed_variances') || '[]';
        let arr = []; try { arr = JSON.parse(raw); } catch(e) {}
        // Remove any prior entry for this target_key, then add new
        arr = arr.filter(d => d.target_key !== targetKey);
        arr.push({ target_key: targetKey, value: value });
        DB.setMetaValue('dismissed_variances', JSON.stringify(arr));
        this.closeVariancePopover();
        // Re-render the current tab so the flag disappears
        if (this.currentTab === 'pnl') this.refreshPnL();
        else if (this.currentTab === 'cashflow') this.refreshCashFlow();
    },

    explainCurrentVariance() {
        const { tab, month } = this._varianceState;
        this.closeVariancePopover();
        // Hand off to period note popover for the (tab, month)
        this.openPeriodNote(tab, month, null);
        // Pre-fill placeholder hint if textarea is empty
        const ta = document.getElementById('periodNotePopTextarea');
        if (!ta.value) {
            ta.placeholder = 'Why was this variance unusual?';
        }
    },
```

- [ ] **Step 2: Wire event listeners**

In `setupEventListeners()`:

```javascript
        document.getElementById('variancePopClose').addEventListener('click', () => this.closeVariancePopover());
        document.getElementById('variancePopDismiss').addEventListener('click', () => this.dismissCurrentVariance());
        document.getElementById('variancePopExplain').addEventListener('click', () => this.explainCurrentVariance());

        // Click-on-flagged-cell delegated handler
        document.addEventListener('click', (e) => {
            const cell = e.target.closest('.cell-flagged');
            if (!cell) return;
            const tab = cell.getAttribute('data-vr-tab');
            const cat = cell.getAttribute('data-vr-cat');
            const month = cell.getAttribute('data-vr-month');
            const value = parseFloat(cell.getAttribute('data-vr-value'));
            if (tab && cat && month) {
                this.openVariancePopover(tab, parseInt(cat), month, value, cell);
            }
        });
```

- [ ] **Step 3: Smoke-test full flow**

Reload. Click any flagged cell → variance popover opens with stats. Click "Dismiss flag" → flag disappears. Click another flagged cell → click "Explain this →" → variance popover closes, period note popover opens for that (tab, month).

- [ ] **Step 4: Commit**

```bash
git add js/app.js
git commit -m "feat(variance): add popover controller with dismiss and explain handoff"
```

---

### Phase 2 Complete — Manual Acceptance Test

- [ ] **Verify Phase 2:**
1. Set threshold to 20% in settings
2. Open P&L — at least one cell should be flagged amber
3. Hover flagged cell → tooltip shows %s
4. Click flagged cell → popover opens with stats
5. Click "Dismiss flag" → flag gone
6. Edit the underlying value (via journal) → flag returns (re-fired)
7. Click flagged cell → "Explain this →" → period note popover opens for same (tab, month)

---

## Phase 3: Insights Hub Tab

### Task 3.1: Add Insights tab to main tab strip

**Files:**
- Modify: `index.html`
- Modify: `js/app.js` (default tab order)

- [ ] **Step 1: Add tab button + content shell**

In `index.html` main tab strip (around line 54-126), add:

```html
<button type="button" class="main-tab-btn" data-tab="insights">📝 Insights</button>
```

And in tab content section, add:

```html
<div id="insightsTab" class="tab-content" hidden>
    <div class="insights-layout">
        <aside class="insights-sidebar" id="insightsSidebar"></aside>
        <main class="insights-content">
            <input id="insightsSearch" class="insight-search" placeholder="🔍 Search notes...">
            <div id="insightsFeed"></div>
        </main>
    </div>
</div>
```

- [ ] **Step 2: Update default tab order**

Find `setTabOrder` defaults in `js/database.js` or `js/app.js`. Add `'insights'` to the default ordered list, positioned after `balancesheet`.

- [ ] **Step 3: CSS**

Append to `css/styles.css` Notes section:

```css
/* === Insights Hub === */
.insights-layout { display: grid; grid-template-columns: 220px 1fr; min-height: 500px; }
.insights-sidebar {
    background: var(--surface-alt);
    border-right: 1px solid var(--border);
    padding: 16px;
}
.insights-content { padding: 16px; }
.insight-search {
    width: 100%; padding: 8px 12px; border: 1px solid var(--border);
    border-radius: var(--radius-sm); font-size: 13px; margin-bottom: 16px;
    box-sizing: border-box;
}
.filter-group { margin-bottom: 16px; }
.filter-label {
    font-size: 10px; text-transform: uppercase; letter-spacing: 0.6px;
    color: var(--text-secondary); margin-bottom: 6px; font-weight: 600;
}
.filter-pill {
    display: block; padding: 6px 10px; margin-bottom: 4px; font-size: 12px;
    color: var(--text); border-radius: 6px; cursor: pointer; user-select: none;
}
.filter-pill.active { background: rgba(59,130,246,0.12); color: var(--primary); font-weight: 600; }
.filter-pill:hover { background: var(--surface); }
.filter-count { float: right; color: var(--text-muted); font-size: 11px; }
.timeline-month {
    font-size: 11px; text-transform: uppercase; letter-spacing: 0.6px;
    color: var(--text-secondary); margin: 16px 0 8px; font-weight: 600;
}
.insight-card {
    background: var(--surface); border: 1px solid var(--border);
    border-radius: var(--radius-sm); padding: 12px 14px;
    margin-bottom: 8px; cursor: pointer; transition: 150ms ease;
}
.insight-card:hover { border-color: var(--primary); box-shadow: 0 2px 8px rgba(59,130,246,0.08); }
.insight-card-header {
    display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px;
}
.insight-tag {
    font-size: 10px; text-transform: uppercase; letter-spacing: 0.5px;
    padding: 2px 8px; border-radius: 4px; font-weight: 600;
}
.tag-period { background: var(--amber-bg); color: #92400e; }
.tag-anomaly { background: var(--blue-bg); color: #1e40af; }
.tag-chart { background: var(--green-bg); color: #065f46; }
.tag-txn { background: var(--surface-alt); color: var(--text-secondary); }
.insight-source { font-size: 11px; color: var(--text-muted); }
.insight-body { font-size: 13px; color: var(--text); line-height: 1.5; }
.insights-empty { text-align: center; padding: 48px 16px; color: var(--text-muted); }
```

- [ ] **Step 4: Smoke-test**

Reload. Click "📝 Insights" tab → shows empty layout with sidebar (empty) and search box.

- [ ] **Step 5: Commit**

```bash
git add index.html css/styles.css js/database.js js/app.js
git commit -m "feat(insights): scaffold Insights Hub tab and layout"
```

---

### Task 3.2: `getInsightsFeed` UNION query

**Files:**
- Modify: `js/database.js`

- [ ] **Step 1: Implement query**

Add to `DB`:

```javascript
    /**
     * Unified feed across notes + transactions for the Insights Hub.
     * @param {Object} filters { types?: Array<string>, sourceTabs?: Array<string>, timeRange?: 'all'|'quarter'|'year', search?: string }
     * @returns {Array<Object>} feed items
     */
    getInsightsFeed(filters = {}) {
        const sql = `
            SELECT
                'period' AS source_type,
                id AS source_id,
                target_key,
                body,
                substr(target_key, instr(target_key, ':') + 1) AS source_month,
                target_key AS source_label_raw,
                updated_at,
                pinned_to_share
            FROM notes WHERE target_type = 'period'
            UNION ALL
            SELECT
                'anomaly', id, target_key, body,
                substr(target_key, instr(target_key, ':') + 1),
                target_key, updated_at, pinned_to_share
            FROM notes WHERE target_type = 'anomaly'
            UNION ALL
            SELECT
                'chart', id, target_key, body,
                substr(target_key, instr(target_key, ':') + 1),
                target_key, updated_at, pinned_to_share
            FROM notes WHERE target_type = 'chart'
            UNION ALL
            SELECT
                'transaction',
                id,
                'txn:' || id,
                notes,
                substr(entry_date, 1, 7),
                'Journal · #' || id || ' · ' || entry_date,
                COALESCE(updated_at, created_at),
                0
            FROM transactions
            WHERE notes IS NOT NULL AND TRIM(notes) != ''
        `;
        const result = this.db.exec(sql);
        if (!result.length) return [];
        let items = result[0].values.map(r => ({
            source_type: r[0],
            source_id: r[1],
            target_key: r[2],
            body: r[3] || '',
            source_month: r[4] || '',
            source_label: this._buildSourceLabel(r[0], r[5]),
            updated_at: r[6],
            pinned_to_share: r[7] || 0
        }));

        // Apply filters
        if (filters.types && filters.types.length) {
            items = items.filter(i => filters.types.includes(i.source_type));
        }
        if (filters.sourceTabs && filters.sourceTabs.length) {
            items = items.filter(i => {
                if (i.source_type === 'period') {
                    const tab = i.target_key.split(':')[0];
                    return filters.sourceTabs.includes(tab);
                }
                if (i.source_type === 'transaction') return filters.sourceTabs.includes('journal');
                if (i.source_type === 'anomaly') return filters.sourceTabs.includes('dashboard');
                if (i.source_type === 'chart') return filters.sourceTabs.includes('dashboard');
                return false;
            });
        }
        if (filters.timeRange && filters.timeRange !== 'all') {
            const now = new Date();
            const cutoff = new Date(now);
            if (filters.timeRange === 'quarter') cutoff.setMonth(cutoff.getMonth() - 3);
            else if (filters.timeRange === 'year') cutoff.setFullYear(cutoff.getFullYear() - 1);
            const cutoffMonth = `${cutoff.getFullYear()}-${String(cutoff.getMonth()+1).padStart(2,'0')}`;
            items = items.filter(i => i.source_month >= cutoffMonth);
        }
        if (filters.search) {
            const q = filters.search.toLowerCase();
            items = items.filter(i => i.body.toLowerCase().includes(q));
        }

        // Sort: by source_month DESC, then updated_at DESC
        items.sort((a, b) => {
            if (a.source_month !== b.source_month) return b.source_month.localeCompare(a.source_month);
            return (b.updated_at || '').localeCompare(a.updated_at || '');
        });
        return items;
    },

    _buildSourceLabel(sourceType, raw) {
        if (sourceType === 'period') {
            const [tab, month] = raw.split(':');
            const tabLabel = { pnl: 'P&L', cashflow: 'Cash Flow', balancesheet: 'Balance Sheet' }[tab] || tab;
            return `${tabLabel} · ${Utils.formatMonthDisplay(month)}`;
        }
        if (sourceType === 'anomaly') {
            const [, month] = raw.split(':');
            return `Dashboard · ${Utils.formatMonthDisplay(month)}`;
        }
        if (sourceType === 'chart') {
            const [chartId, month] = raw.split(':');
            const chartName = { dashChartCashFlow: 'Cash Flow Trend', dashChartPnL: 'P&L Trend', dashChartBS: 'Balance Sheet Trend' }[chartId] || chartId;
            return `${chartName} · ${Utils.formatMonthDisplay(month)}`;
        }
        return raw; // transaction labels are pre-built
    },
```

- [ ] **Step 2: Smoke-test in console**

```javascript
DB.setNote('period', 'pnl:2026-03', 'Test period note');
DB.setNote('chart', 'dashChartPnL:2026-04', 'Test chart pin');
DB.getInsightsFeed({});
// Expected: array with 2 items, sorted newest-first
DB.getInsightsFeed({ search: 'period' });
// Expected: 1 item
DB.deleteNote('period','pnl:2026-03');
DB.deleteNote('chart','dashChartPnL:2026-04');
```

- [ ] **Step 3: Commit**

```bash
git add js/database.js
git commit -m "feat(insights): add getInsightsFeed UNION query with filters"
```

---

### Task 3.3: Insights tab renderer + filter UI

**Files:**
- Modify: `js/ui.js`
- Modify: `js/app.js`

- [ ] **Step 1: Renderer**

Add to `UI` in `js/ui.js`:

```javascript
    renderInsightsTab(filters, feed) {
        // Sidebar — counts come from full feed (no filter) for the type pills
        const fullFeed = DB.getInsightsFeed({});
        const counts = {
            all: fullFeed.length,
            period: fullFeed.filter(i => i.source_type === 'period').length,
            anomaly: fullFeed.filter(i => i.source_type === 'anomaly').length,
            chart: fullFeed.filter(i => i.source_type === 'chart').length,
            transaction: fullFeed.filter(i => i.source_type === 'transaction').length,
        };
        const tabCounts = {
            cashflow: fullFeed.filter(i => i.source_type === 'period' && i.target_key.startsWith('cashflow:')).length,
            pnl: fullFeed.filter(i => i.source_type === 'period' && i.target_key.startsWith('pnl:')).length,
            balancesheet: fullFeed.filter(i => i.source_type === 'period' && i.target_key.startsWith('balancesheet:')).length,
            dashboard: fullFeed.filter(i => i.source_type === 'anomaly' || i.source_type === 'chart').length,
            journal: fullFeed.filter(i => i.source_type === 'transaction').length,
        };
        const isType = (t) => filters.types && filters.types.includes(t) ? 'active' : '';
        const isTypeAll = (!filters.types || !filters.types.length) ? 'active' : '';
        const isTab = (t) => filters.sourceTabs && filters.sourceTabs.includes(t) ? 'active' : '';
        const isTime = (t) => (filters.timeRange || 'all') === t ? 'active' : '';

        const sidebar = `
            <div class="filter-group">
                <div class="filter-label">Type</div>
                <div class="filter-pill ${isTypeAll}" data-filter-type="all">All <span class="filter-count">${counts.all}</span></div>
                <div class="filter-pill ${isType('period')}" data-filter-type="period">📝 Period notes <span class="filter-count">${counts.period}</span></div>
                <div class="filter-pill ${isType('anomaly')}" data-filter-type="anomaly">🎯 Anomalies <span class="filter-count">${counts.anomaly}</span></div>
                <div class="filter-pill ${isType('chart')}" data-filter-type="chart">📊 Chart pins <span class="filter-count">${counts.chart}</span></div>
                <div class="filter-pill ${isType('transaction')}" data-filter-type="transaction">💬 Transaction <span class="filter-count">${counts.transaction}</span></div>
            </div>
            <div class="filter-group">
                <div class="filter-label">Source tab</div>
                <div class="filter-pill ${isTab('cashflow')}" data-filter-tab="cashflow">Cash Flow <span class="filter-count">${tabCounts.cashflow}</span></div>
                <div class="filter-pill ${isTab('pnl')}" data-filter-tab="pnl">P&L <span class="filter-count">${tabCounts.pnl}</span></div>
                <div class="filter-pill ${isTab('balancesheet')}" data-filter-tab="balancesheet">Balance Sheet <span class="filter-count">${tabCounts.balancesheet}</span></div>
                <div class="filter-pill ${isTab('dashboard')}" data-filter-tab="dashboard">Dashboard <span class="filter-count">${tabCounts.dashboard}</span></div>
                <div class="filter-pill ${isTab('journal')}" data-filter-tab="journal">Journal <span class="filter-count">${tabCounts.journal}</span></div>
            </div>
            <div class="filter-group">
                <div class="filter-label">Time range</div>
                <div class="filter-pill ${isTime('all')}" data-filter-time="all">All time</div>
                <div class="filter-pill ${isTime('quarter')}" data-filter-time="quarter">This quarter</div>
                <div class="filter-pill ${isTime('year')}" data-filter-time="year">This year</div>
            </div>
        `;
        document.getElementById('insightsSidebar').innerHTML = sidebar;

        // Feed grouped by month
        const feedEl = document.getElementById('insightsFeed');
        if (!feed.length) {
            feedEl.innerHTML = `<div class="insights-empty">No notes match these filters yet.</div>`;
            return;
        }
        const groups = {};
        feed.forEach(item => {
            (groups[item.source_month] = groups[item.source_month] || []).push(item);
        });
        const sortedMonths = Object.keys(groups).sort((a,b) => b.localeCompare(a));
        feedEl.innerHTML = sortedMonths.map(m => {
            const cards = groups[m].map(i => this.renderInsightCard(i)).join('');
            return `<div class="timeline-month">${Utils.formatMonthDisplay(m)}</div>${cards}`;
        }).join('');
    },

    renderInsightCard(item) {
        const tagMap = {
            period: { class: 'tag-period', label: '📝 Period Note' },
            anomaly: { class: 'tag-anomaly', label: '🎯 Anomaly' },
            chart: { class: 'tag-chart', label: '📊 Chart Pin' },
            transaction: { class: 'tag-txn', label: '💬 Transaction' }
        };
        const t = tagMap[item.source_type];
        const bodyHtml = Utils.renderMarkdownSafe(item.body);
        return `<div class="insight-card" data-source-type="${item.source_type}" data-target-key="${Utils.escapeHtml(item.target_key)}">
            <div class="insight-card-header">
                <span class="insight-tag ${t.class}">${t.label}</span>
                <span class="insight-source">${Utils.escapeHtml(item.source_label)}</span>
            </div>
            <div class="insight-body">${bodyHtml}</div>
        </div>`;
    },
```

- [ ] **Step 2: Controller**

Add to `App` in `js/app.js`:

```javascript
    // ============== INSIGHTS HUB ==============

    _insightsState: { types: [], sourceTabs: [], timeRange: 'all', search: '' },

    refreshInsights() {
        const f = this._insightsState;
        const feed = DB.getInsightsFeed(f);
        UI.renderInsightsTab(f, feed);
    },

    _toggleInsightFilter(kind, value) {
        if (kind === 'type') {
            if (value === 'all') { this._insightsState.types = []; }
            else {
                const arr = this._insightsState.types;
                const i = arr.indexOf(value);
                if (i >= 0) arr.splice(i, 1); else arr.push(value);
            }
        } else if (kind === 'tab') {
            const arr = this._insightsState.sourceTabs;
            const i = arr.indexOf(value);
            if (i >= 0) arr.splice(i, 1); else arr.push(value);
        } else if (kind === 'time') {
            this._insightsState.timeRange = value;
        }
        this.refreshInsights();
    },
```

In `setupEventListeners()`:

```javascript
        // Insights filters (delegated)
        document.addEventListener('click', (e) => {
            const pill = e.target.closest('.filter-pill');
            if (!pill || !pill.closest('#insightsSidebar')) return;
            if (pill.hasAttribute('data-filter-type'))
                this._toggleInsightFilter('type', pill.getAttribute('data-filter-type'));
            else if (pill.hasAttribute('data-filter-tab'))
                this._toggleInsightFilter('tab', pill.getAttribute('data-filter-tab'));
            else if (pill.hasAttribute('data-filter-time'))
                this._toggleInsightFilter('time', pill.getAttribute('data-filter-time'));
        });
        // Search
        let searchTimer = null;
        document.getElementById('insightsSearch').addEventListener('input', (e) => {
            clearTimeout(searchTimer);
            searchTimer = setTimeout(() => {
                this._insightsState.search = e.target.value;
                this.refreshInsights();
            }, 200);
        });
```

In the tab-switch logic, add a case:

```javascript
        else if (tab === 'insights') this.refreshInsights();
```

- [ ] **Step 3: Smoke-test**

Reload. Add a few notes (period, chart, txn). Switch to Insights tab. Filters render with counts. Click a type pill → narrows feed. Type in search box → debounced filter applies.

- [ ] **Step 4: Commit**

```bash
git add js/ui.js js/app.js
git commit -m "feat(insights): render Insights Hub feed with filters and search"
```

---

### Task 3.4: Cross-tab navigation from insight cards

**Files:**
- Modify: `js/app.js`

- [ ] **Step 1: Click-on-card handler**

In `setupEventListeners()`:

```javascript
        document.addEventListener('click', (e) => {
            const card = e.target.closest('.insight-card');
            if (!card) return;
            const sourceType = card.getAttribute('data-source-type');
            const targetKey = card.getAttribute('data-target-key');
            this._navigateToInsightSource(sourceType, targetKey);
        });
```

- [ ] **Step 2: Navigation router**

Add to `App`:

```javascript
    _navigateToInsightSource(sourceType, targetKey) {
        if (sourceType === 'period') {
            const [tab, month] = targetKey.split(':');
            this.switchTab(tab);
            // Open popover for the matching column header (best-effort)
            setTimeout(() => {
                const el = document.querySelector(`.pn-icon-wrap[data-pn-tab="${tab}"][data-pn-month="${month}"]`);
                if (el) {
                    el.scrollIntoView({ behavior: 'smooth', block: 'center' });
                    this.openPeriodNote(tab, month, el);
                }
            }, 100);
        } else if (sourceType === 'anomaly') {
            this.switchTab('dashboard');
            // Phase 4 implements expand+flash
        } else if (sourceType === 'chart') {
            this.switchTab('dashboard');
            // Phase 5 implements scroll+flash
        } else if (sourceType === 'transaction') {
            const id = targetKey.replace('txn:', '');
            this.switchTab('journal');
            setTimeout(() => {
                const row = document.querySelector(`tr[data-transaction-id="${id}"]`);
                if (row) {
                    row.scrollIntoView({ behavior: 'smooth', block: 'center' });
                    row.classList.add('flash-highlight');
                    setTimeout(() => row.classList.remove('flash-highlight'), 1500);
                }
            }, 100);
        }
    },
```

- [ ] **Step 3: Add flash highlight CSS**

Append to `css/styles.css`:

```css
@keyframes flash-highlight-anim {
    0%, 100% { background: transparent; }
    20%, 80% { background: rgba(59,130,246,0.15); }
}
.flash-highlight { animation: flash-highlight-anim 1.5s ease; }
```

- [ ] **Step 4: Smoke-test**

Add a P&L period note. Go to Insights tab. Click the card → tab switches to P&L, scrolls to that month, period popover opens. Add a transaction note. Click that card → switches to Journal, scrolls to row, row flashes blue.

- [ ] **Step 5: Commit**

```bash
git add js/app.js css/styles.css
git commit -m "feat(insights): cross-tab nav from insight cards to source location"
```

---

### Phase 3 Complete — Manual Acceptance Test

- [ ] **Verify Phase 3:**
1. Add notes across all 4 sources (period × 3 tabs, transaction note, fake chart pin via DB call, fake anomaly note via DB call)
2. Open Insights tab → all show, grouped by month, newest first
3. Type in search → narrows
4. Click type pill → narrows
5. Click time pill → narrows
6. Click any card → navigates to source, opens correct popover/scrolls correctly

---

## Phase 4: Anomaly Dashboard Widget

### Task 4.1: Anomaly rule functions + tests

**Files:**
- Modify: `js/utils.js`
- Modify: `tests/notes-analysis.test.html`

- [ ] **Step 1: Add tests for `runway_drop` (representative — others follow same shape)**

Append to `tests/notes-analysis.test.html`:

```javascript
const s3 = suite('Anomaly rules');
test(s3, 'runway_drop fires when runway shortens by >10 days', () => {
  // Mock context object with the inputs the rule reads
  const ctx = { snapshotMonth: '2026-04', priorRunwayDays: 142, currentRunwayDays: 124 };
  const a = Utils.anomalyRules.runwayDrop(ctx);
  assert(a, 'should fire');
  assertEq(a.signature, 'runway_drop');
  assertEq(a.severity, 'critical');
  assert(a.title.includes('18 days'), `title: ${a.title}`);
});
test(s3, 'runway_drop returns null when shortening <10 days', () => {
  const ctx = { snapshotMonth: '2026-04', priorRunwayDays: 142, currentRunwayDays: 138 };
  const a = Utils.anomalyRules.runwayDrop(ctx);
  assertEq(a, null);
});
test(s3, 'category_spike fires when current > 3x 6-mo avg', () => {
  const ctx = {
    snapshotMonth: '2026-04',
    categoryMonthlyTotals: {
      42: { name: 'Office Supplies', months: { '2025-10': 400, '2025-11': 380, '2025-12': 410, '2026-01': 420, '2026-02': 410, '2026-03': 460, '2026-04': 1840 } }
    }
  };
  const list = Utils.anomalyRules.categorySpike(ctx);
  assert(list.length > 0, 'should fire');
  assert(list[0].title.includes('Office Supplies'), 'mentions category');
  assertEq(list[0].severity, 'notable');
});
```

- [ ] **Step 2: Run failing tests (expect 3 failures)**

- [ ] **Step 3: Implement rule functions**

Append to `Utils` in `js/utils.js`:

```javascript
    /**
     * Anomaly rule functions. Each takes a context object and returns
     * either a single anomaly, an array of anomalies, or null.
     * Anomaly shape: { signature, severity, title, description, sourceTab, targetMonth, targetId, explanationStarter }
     */
    anomalyRules: {
        runwayDrop(ctx) {
            const drop = (ctx.priorRunwayDays || 0) - (ctx.currentRunwayDays || 0);
            if (drop <= 10) return null;
            return {
                signature: 'runway_drop',
                severity: 'critical',
                title: `Cash runway shortened by ${Math.round(drop)} days`,
                description: `From ${Math.round(ctx.priorRunwayDays)} days last month to ${Math.round(ctx.currentRunwayDays)} days now.`,
                sourceTab: 'cashflow',
                targetMonth: ctx.snapshotMonth,
                targetId: null,
                explanationStarter: 'Why did runway shorten this month?'
            };
        },
        categorySpike(ctx) {
            const out = [];
            for (const [catId, info] of Object.entries(ctx.categoryMonthlyTotals || {})) {
                const months = info.months;
                const cur = months[ctx.snapshotMonth];
                if (!cur) continue;
                // 6-month avg ending the month before snapshot
                const [y, m] = ctx.snapshotMonth.split('-').map(Number);
                const priorKeys = [];
                let yy = y, mm = m;
                for (let i = 0; i < 6; i++) {
                    mm--; if (mm < 1) { mm = 12; yy--; }
                    priorKeys.push(`${yy}-${String(mm).padStart(2,'0')}`);
                }
                const priors = priorKeys.map(k => months[k] || 0);
                const avg = priors.reduce((a,b)=>a+b,0) / 6;
                if (avg > 0 && cur >= 3 * avg) {
                    out.push({
                        signature: 'category_spike',
                        severity: 'notable',
                        title: `${info.name} spend ${(cur/avg).toFixed(1)}× the 6-month avg`,
                        description: `$${cur.toFixed(0)} this month vs $${avg.toFixed(0)} avg.`,
                        sourceTab: 'pnl',
                        targetMonth: ctx.snapshotMonth,
                        targetId: parseInt(catId),
                        explanationStarter: `Why was ${info.name} spending so high?`
                    });
                }
            }
            return out;
        },
        arAgingTrend(ctx) {
            const series = ctx.arDaysSeries || [];
            if (series.length < 3) return null;
            const last3 = series.slice(-3);
            if (last3[0] < last3[1] && last3[1] < last3[2]) {
                return {
                    signature: 'ar_aging_trend',
                    severity: 'critical',
                    title: 'AR Days Outstanding climbing 3 months in a row',
                    description: `${last3[0]} → ${last3[1]} → ${last3[2]} days.`,
                    sourceTab: 'dashboard',
                    targetMonth: ctx.snapshotMonth,
                    targetId: null,
                    explanationStarter: 'Why is AR aging worsening?'
                };
            }
            return null;
        },
        profitabilityFlip(ctx) {
            const cur = ctx.netIncomeCurrent;
            const prior = ctx.netIncomePrior;
            if (cur == null || prior == null) return null;
            if (prior <= 0 && cur > 0) {
                return {
                    signature: 'profitability_flip',
                    severity: 'positive',
                    title: 'First profitable month in recent history',
                    description: `Net income +$${cur.toFixed(0)}.`,
                    sourceTab: 'pnl',
                    targetMonth: ctx.snapshotMonth,
                    targetId: null,
                    explanationStarter: 'What drove the swing to profitability?'
                };
            }
            if (prior > 0 && cur <= 0) {
                return {
                    signature: 'profitability_flip',
                    severity: 'critical',
                    title: 'Month flipped to a loss',
                    description: `Net income $${cur.toFixed(0)}, down from $${prior.toFixed(0)}.`,
                    sourceTab: 'pnl',
                    targetMonth: ctx.snapshotMonth,
                    targetId: null,
                    explanationStarter: 'Why did this month swing negative?'
                };
            }
            return null;
        },
        dscrBreach(ctx) {
            if (ctx.dscrCurrent == null || ctx.dscrPrior == null) return null;
            if (ctx.dscrCurrent < 1.0 && ctx.dscrPrior >= 1.0) {
                return {
                    signature: 'dscr_breach',
                    severity: 'critical',
                    title: `DSCR dropped below 1.0`,
                    description: `Now ${ctx.dscrCurrent.toFixed(2)} (was ${ctx.dscrPrior.toFixed(2)}). Covenant risk.`,
                    sourceTab: 'dashboard',
                    targetMonth: ctx.snapshotMonth,
                    targetId: null,
                    explanationStarter: 'How will you address the DSCR breach?'
                };
            }
            return null;
        },
        recurringDisappeared(ctx) {
            const out = [];
            for (const [catId, info] of Object.entries(ctx.categoryMonthlyTotals || {})) {
                const months = info.months;
                const cur = months[ctx.snapshotMonth] || 0;
                if (cur !== 0) continue;
                // Check 6 prior consecutive months
                const [y, m] = ctx.snapshotMonth.split('-').map(Number);
                let yy = y, mm = m, allActive = true;
                for (let i = 0; i < 6; i++) {
                    mm--; if (mm < 1) { mm = 12; yy--; }
                    const k = `${yy}-${String(mm).padStart(2,'0')}`;
                    if (!months[k] || months[k] === 0) { allActive = false; break; }
                }
                if (allActive) {
                    out.push({
                        signature: 'recurring_disappeared',
                        severity: 'notable',
                        title: `${info.name} (recurring) has no entries this month`,
                        description: 'Was active 6 months in a row, missing this month.',
                        sourceTab: 'cashflow',
                        targetMonth: ctx.snapshotMonth,
                        targetId: parseInt(catId),
                        explanationStarter: `Did ${info.name} actually stop, or is it just not entered yet?`
                    });
                }
            }
            return out;
        }
    },

    /**
     * Run all anomaly rules and return a flat array, filtered against dismissed list.
     */
    computeAnomalies(ctx, dismissed) {
        const all = [];
        for (const [name, rule] of Object.entries(this.anomalyRules)) {
            const r = rule(ctx);
            if (!r) continue;
            if (Array.isArray(r)) all.push(...r);
            else all.push(r);
        }
        return all.filter(a => !dismissed.some(d =>
            d.signature === a.signature &&
            d.month === a.targetMonth &&
            (d.target_id == null || d.target_id === a.targetId)
        ));
    },
```

- [ ] **Step 4: Re-run tests**

Reload `tests/notes-analysis.test.html`. Expected: all tests pass.

- [ ] **Step 5: Commit**

```bash
git add js/utils.js tests/notes-analysis.test.html
git commit -m "feat(anomaly): add 6 anomaly rule functions with tests"
```

---

### Task 4.2: Anomaly widget UI on dashboard

**Files:**
- Modify: `index.html` (Dashboard section)
- Modify: `css/styles.css`
- Modify: `js/app.js`
- Modify: `js/ui.js`

- [ ] **Step 1: HTML shell**

In `index.html` Dashboard tab, add at the top of the dashboard content area:

```html
<div id="anomalyWidget" class="anomaly-widget" hidden>
    <div class="anomaly-widget-header" id="anomalyWidgetHeader">
        <div class="aw-title">
            <span>⚡ Worth Looking At — <span id="anomalyWidgetMonth">—</span></span>
            <span class="aw-count" id="anomalyWidgetCount">0</span>
        </div>
        <button type="button" class="aw-collapse-btn" id="anomalyWidgetCollapse">▾</button>
    </div>
    <div class="anomaly-widget-body" id="anomalyWidgetBody"></div>
</div>
```

- [ ] **Step 2: CSS**

Append:

```css
/* === Anomaly Widget === */
.anomaly-widget {
    background: linear-gradient(180deg, var(--amber-bg) 0%, var(--surface) 50%);
    border: 1px solid var(--amber);
    border-radius: var(--radius);
    margin-bottom: 16px;
    overflow: hidden;
}
[data-theme="dark"] .anomaly-widget {
    background: linear-gradient(180deg, rgba(217,119,6,0.12) 0%, var(--surface) 50%);
}
.anomaly-widget-header {
    display: flex; justify-content: space-between; align-items: center;
    padding: 12px 16px; cursor: pointer; user-select: none;
}
.aw-title { display: flex; align-items: center; gap: 12px; font-weight: 600; }
.aw-count {
    background: var(--amber); color: white; font-size: 11px;
    padding: 2px 8px; border-radius: 10px; font-weight: 600;
}
.aw-collapse-btn {
    background: transparent; border: none; cursor: pointer;
    color: var(--text-secondary); font-size: 14px;
    transition: transform 200ms;
}
.anomaly-widget.collapsed .aw-collapse-btn { transform: rotate(-90deg); }
.anomaly-widget.collapsed .anomaly-widget-body { display: none; }
.anomaly-widget-body { padding: 0 12px 12px; }
.anomaly-item {
    display: flex; gap: 12px; padding: 12px;
    border-radius: var(--radius-sm); margin-bottom: 8px;
    border: 1px solid var(--amber); background: var(--amber-bg);
}
[data-theme="dark"] .anomaly-item { background: rgba(217,119,6,0.12); }
.anomaly-item.critical { border-color: var(--red); background: var(--red-bg); }
[data-theme="dark"] .anomaly-item.critical { background: rgba(220,38,38,0.12); }
.anomaly-item.positive { border-color: var(--primary); background: var(--blue-bg); }
[data-theme="dark"] .anomaly-item.positive { background: rgba(59,130,246,0.12); }
.anomaly-icon-wrap {
    width: 32px; height: 32px; border-radius: var(--radius-sm);
    background: var(--surface);
    display: flex; align-items: center; justify-content: center;
    flex-shrink: 0; font-size: 16px;
}
.anomaly-content { flex: 1; }
.anomaly-title { font-size: 13px; font-weight: 600; margin: 0 0 2px; }
.anomaly-desc { font-size: 12px; color: var(--text-secondary); margin: 0 0 8px; }
.anomaly-actions { display: flex; gap: 8px; flex-wrap: wrap; }
.anomaly-btn {
    padding: 4px 10px; border-radius: 6px; font-size: 11px; font-weight: 600;
    border: 1px solid var(--border); background: var(--surface);
    cursor: pointer; color: var(--text-secondary);
}
.anomaly-btn:hover { color: var(--primary); border-color: var(--primary); }
.anomaly-btn.primary { background: var(--primary); color: white; border-color: var(--primary); }
.aw-show-more {
    text-align: center; padding: 8px; cursor: pointer;
    color: var(--primary); font-size: 12px; font-weight: 600;
}
```

- [ ] **Step 3: Renderer**

Add to `UI`:

```javascript
    renderAnomalyCard(a) {
        const iconMap = { critical: '⚠️', notable: '📊', positive: '✨' };
        const icon = iconMap[a.severity] || '⚡';
        return `<div class="anomaly-item ${a.severity}" data-anomaly-sig="${a.signature}" data-anomaly-month="${a.targetMonth}" data-anomaly-target="${a.targetId || ''}">
            <div class="anomaly-icon-wrap">${icon}</div>
            <div class="anomaly-content">
                <p class="anomaly-title">${Utils.escapeHtml(a.title)}</p>
                <p class="anomaly-desc">${Utils.escapeHtml(a.description)}</p>
                <div class="anomaly-actions">
                    <button class="anomaly-btn primary" data-anom-action="capture">Capture explanation</button>
                    <button class="anomaly-btn" data-anom-action="view">View ${a.sourceTab} →</button>
                    <button class="anomaly-btn" data-anom-action="dismiss">Dismiss</button>
                </div>
            </div>
        </div>`;
    },
```

- [ ] **Step 4: Controller**

Add to `App`:

```javascript
    // ============== ANOMALY WIDGET ==============

    _anomalyState: { items: [], expanded: false, snapshotMonth: null },

    _renderAnomalyWidget(snapshotMonth) {
        const ctx = this._buildAnomalyContext(snapshotMonth);
        const dismissedRaw = DB.getMetaValue('dismissed_anomalies') || '[]';
        let dismissed = []; try { dismissed = JSON.parse(dismissedRaw); } catch(e) {}
        const all = Utils.computeAnomalies(ctx, dismissed);
        // Sort by severity (critical > notable > positive)
        const order = { critical: 0, notable: 1, positive: 2 };
        all.sort((a, b) => order[a.severity] - order[b.severity]);

        const widget = document.getElementById('anomalyWidget');
        const body = document.getElementById('anomalyWidgetBody');
        const monthEl = document.getElementById('anomalyWidgetMonth');
        const countEl = document.getElementById('anomalyWidgetCount');

        if (!all.length) {
            widget.setAttribute('hidden', '');
            return;
        }
        if (this.isViewOnly) {
            // Hidden in view-only
            widget.setAttribute('hidden', '');
            return;
        }
        widget.removeAttribute('hidden');
        monthEl.textContent = Utils.formatMonthDisplay(snapshotMonth);
        countEl.textContent = all.length;
        this._anomalyState = { items: all, snapshotMonth, expanded: this._anomalyState.expanded };

        // Default collapsed
        const visible = this._anomalyState.expanded ? all : all.slice(0, 3);
        body.innerHTML = visible.map(a => UI.renderAnomalyCard(a)).join('') +
            (all.length > 3 && !this._anomalyState.expanded
                ? `<div class="aw-show-more" id="anomalyShowMore">Show ${all.length - 3} more →</div>`
                : '');
        widget.classList.toggle('collapsed', false); // ensure expanded shell
    },

    _buildAnomalyContext(snapshotMonth) {
        // Build the context object with the inputs all rules need.
        // Pull from existing DB methods used by KPIs.
        const [y, m] = snapshotMonth.split('-').map(Number);
        const prevMonth = (m === 1) ? `${y-1}-12` : `${y}-${String(m-1).padStart(2,'0')}`;

        // Runway
        let priorRunway = null, currentRunway = null;
        try {
            const kpi = this._computeKpiData(snapshotMonth);
            const kpiPrior = this._computeKpiData(prevMonth);
            currentRunway = kpi && kpi.runwayDays;
            priorRunway = kpiPrior && kpiPrior.runwayDays;
        } catch(e) { console.warn('runway calc failed', e); }

        // Net income
        let netIncomeCurrent = null, netIncomePrior = null;
        try {
            netIncomeCurrent = DB.getNetIncomeForMonth ? DB.getNetIncomeForMonth(snapshotMonth) : null;
            netIncomePrior = DB.getNetIncomeForMonth ? DB.getNetIncomeForMonth(prevMonth) : null;
        } catch(e) {}

        // DSCR
        let dscrCurrent = null, dscrPrior = null;
        try {
            const k1 = this._computeKpiData(snapshotMonth);
            const k0 = this._computeKpiData(prevMonth);
            dscrCurrent = k1 && k1.dscr;
            dscrPrior = k0 && k0.dscr;
        } catch(e) {}

        // AR aging trend (last 3 months including current)
        const arSeries = [];
        try {
            for (let i = 2; i >= 0; i--) {
                let yy = y, mm = m - i;
                while (mm < 1) { mm += 12; yy--; }
                const month = `${yy}-${String(mm).padStart(2,'0')}`;
                const k = this._computeKpiData(month);
                arSeries.push(k && k.arDays != null ? k.arDays : null);
            }
        } catch(e) {}

        // Category monthly totals (for category_spike, recurring_disappeared)
        const categoryMonthlyTotals = {};
        try {
            const cats = DB.db.exec("SELECT id, name FROM categories");
            if (cats.length) {
                cats[0].values.forEach(([id, name]) => {
                    const r = DB.db.exec(
                        "SELECT month_paid, SUM(amount) FROM transactions " +
                        "WHERE category_id = ? AND transaction_type = 'expense' " +
                        "AND month_paid IS NOT NULL GROUP BY month_paid",
                        [id]
                    );
                    if (r.length) {
                        const months = {};
                        r[0].values.forEach(([mo, sum]) => { months[mo] = Number(sum); });
                        categoryMonthlyTotals[id] = { name, months };
                    }
                });
            }
        } catch(e) { console.warn('category aggregation failed', e); }

        return {
            snapshotMonth,
            priorRunwayDays: priorRunway,
            currentRunwayDays: currentRunway,
            netIncomeCurrent, netIncomePrior,
            dscrCurrent, dscrPrior,
            arDaysSeries: arSeries.filter(v => v != null),
            categoryMonthlyTotals
        };
    },

    _toggleAnomalyWidgetExpand() {
        this._anomalyState.expanded = !this._anomalyState.expanded;
        this._renderAnomalyWidget(this._anomalyState.snapshotMonth);
    },

    _toggleAnomalyWidgetCollapse() {
        const w = document.getElementById('anomalyWidget');
        w.classList.toggle('collapsed');
    },

    _captureAnomalyExplanation(sig, month, targetId) {
        const a = this._anomalyState.items.find(x => x.signature === sig && x.targetMonth === month && (x.targetId == (targetId || null) || (!x.targetId && !targetId)));
        if (!a) return;
        const tab = a.sourceTab === 'pnl' || a.sourceTab === 'cashflow' || a.sourceTab === 'balancesheet' ? a.sourceTab : 'pnl';
        this.openPeriodNote(tab, month, null);
        const ta = document.getElementById('periodNotePopTextarea');
        if (!ta.value) ta.placeholder = a.explanationStarter || 'Why did this happen?';
    },

    _viewAnomalySource(sig, month, targetId) {
        const a = this._anomalyState.items.find(x => x.signature === sig && x.targetMonth === month);
        if (a && a.sourceTab) this.switchTab(a.sourceTab);
    },

    _dismissAnomaly(sig, month, targetId) {
        const raw = DB.getMetaValue('dismissed_anomalies') || '[]';
        let arr = []; try { arr = JSON.parse(raw); } catch(e) {}
        arr.push({ signature: sig, month, target_id: targetId ? parseInt(targetId) : null });
        DB.setMetaValue('dismissed_anomalies', JSON.stringify(arr));
        this._renderAnomalyWidget(this._anomalyState.snapshotMonth);
    },
```

- [ ] **Step 5: Wire to dashboard render and event listeners**

In `_renderDashboardSections` in `js/app.js` (around line 2323), add early call:

```javascript
        this._renderAnomalyWidget(snapshotMonth);
```

In `setupEventListeners()`:

```javascript
        document.getElementById('anomalyWidgetHeader').addEventListener('click', (e) => {
            if (e.target.closest('.aw-collapse-btn') || e.target === e.currentTarget) {
                this._toggleAnomalyWidgetCollapse();
            }
        });
        document.addEventListener('click', (e) => {
            const showMore = e.target.closest('#anomalyShowMore');
            if (showMore) { this._toggleAnomalyWidgetExpand(); return; }
            const item = e.target.closest('.anomaly-item');
            if (!item) return;
            const action = e.target.closest('[data-anom-action]')?.getAttribute('data-anom-action');
            if (!action) return;
            const sig = item.getAttribute('data-anomaly-sig');
            const month = item.getAttribute('data-anomaly-month');
            const targetId = item.getAttribute('data-anomaly-target') || null;
            if (action === 'capture') this._captureAnomalyExplanation(sig, month, targetId);
            else if (action === 'view') this._viewAnomalySource(sig, month, targetId);
            else if (action === 'dismiss') this._dismissAnomaly(sig, month, targetId);
        });
```

- [ ] **Step 6: Smoke-test**

Reload. Open Dashboard. If your data has any qualifying anomalies, the widget appears. Click "Capture explanation" → period note popover opens. Click Dismiss → anomaly disappears. Refresh — dismissed anomaly stays hidden for that month.

If no anomalies fire naturally, set test data: edit a few transactions to force a category spike or runway drop.

- [ ] **Step 7: Commit**

```bash
git add index.html css/styles.css js/app.js js/ui.js
git commit -m "feat(anomaly): add Worth Looking At dashboard widget"
```

---

### Phase 4 Complete — Manual Acceptance Test

- [ ] **Verify Phase 4:**
1. Dashboard shows widget when anomalies fire
2. Top 3 visible by default; "Show N more" expander works
3. Each card has 3 buttons (capture / view / dismiss)
4. Dismiss → card vanishes, persists across page reload
5. Capture → period note popover opens with starter placeholder
6. View → switches tab
7. Widget hidden in view-only mode

---

## Phase 5: Chart Annotations

### Task 5.1: Annotation mode toggle on dashboard charts

**Files:**
- Modify: `js/ui.js` (or wherever dashboard charts render their containers)
- Modify: `js/app.js`
- Modify: `css/styles.css`

- [ ] **Step 1: Inject "Add annotation" button into chart card headers**

Find dashboard chart sections in `js/app.js` `_renderDashboardSections` (around line 2323). For each of the three line-chart canvases (`dashChartCashFlow`, `dashChartPnL`, `dashChartBS`), the surrounding HTML wrapper contains a section title. Modify the wrapper HTML to include a button:

```html
<div class="dash-section-title-row">
    <h3 class="dash-section-title">Cash Flow Trend</h3>
    <button type="button" class="chart-annotate-btn" data-chart-id="dashChartCashFlow">📌 Add annotation</button>
</div>
```

(Apply the same pattern to PnL and BS chart sections, with their respective chart IDs.)

- [ ] **Step 2: CSS**

Append:

```css
.dash-section-title-row {
    display: flex; justify-content: space-between; align-items: center;
    margin-bottom: 8px;
}
.chart-annotate-btn {
    padding: 4px 10px; border-radius: 6px; font-size: 11px; font-weight: 600;
    border: 1px solid var(--border); background: var(--surface);
    cursor: pointer; color: var(--text-secondary);
}
.chart-annotate-btn.active { background: var(--primary); color: white; border-color: var(--primary); }
.chart-annotation-overlay {
    position: absolute; pointer-events: none;
    z-index: 5;
}
.chart-annotation-pin {
    position: absolute; width: 12px; height: 12px; border-radius: 50%;
    background: var(--amber); border: 2px solid var(--surface);
    box-shadow: 0 0 0 2px var(--amber);
    pointer-events: auto; cursor: pointer;
}
.chart-annotation-callout {
    position: absolute; background: var(--surface); border: 1px solid var(--border);
    border-radius: var(--radius-sm); padding: 8px 12px;
    font-size: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.08);
    max-width: 200px; line-height: 1.4;
    pointer-events: auto; cursor: move; user-select: none;
}
.chart-annotation-callout strong {
    display: block; color: var(--text); margin-bottom: 2px; font-size: 11px;
}
.chart-annotation-line {
    position: absolute; width: 1px; pointer-events: none;
    border-left: 1px dashed var(--amber); opacity: 0.6;
}
```

- [ ] **Step 3: Annotation mode controller**

Add to `App`:

```javascript
    // ============== CHART ANNOTATIONS ==============

    _chartAnnotationState: { activeChartId: null },

    _toggleChartAnnotateMode(chartId) {
        const btn = document.querySelector(`.chart-annotate-btn[data-chart-id="${chartId}"]`);
        if (this._chartAnnotationState.activeChartId === chartId) {
            this._chartAnnotationState.activeChartId = null;
            btn.classList.remove('active');
            btn.textContent = '📌 Add annotation';
        } else {
            // Deactivate any other
            document.querySelectorAll('.chart-annotate-btn.active').forEach(b => {
                b.classList.remove('active');
                b.textContent = '📌 Add annotation';
            });
            this._chartAnnotationState.activeChartId = chartId;
            btn.classList.add('active');
            btn.textContent = '✕ Cancel';
        }
    },
```

In `setupEventListeners()`:

```javascript
        document.addEventListener('click', (e) => {
            const btn = e.target.closest('.chart-annotate-btn');
            if (btn) this._toggleChartAnnotateMode(btn.getAttribute('data-chart-id'));
        });
```

- [ ] **Step 4: Smoke-test**

Reload. Dashboard shows three "📌 Add annotation" buttons. Click one → it activates (blue, "✕ Cancel"). Click again or another → toggles correctly.

- [ ] **Step 5: Commit**

```bash
git add js/app.js js/ui.js css/styles.css
git commit -m "feat(chart-annot): add annotation mode toggle to dashboard charts"
```

---

### Task 5.2: Click-on-point creates annotation

**Files:**
- Modify: `js/app.js` (chart click handlers — Chart.js exposes `onClick` in options)

- [ ] **Step 1: Hook Chart.js onClick**

In `js/app.js` find where `dashChartCashFlow`, `dashChartPnL`, `dashChartBS` are constructed (search for `new Chart(` near `_renderDashCashFlowChart`). Each Chart.js instance config has an `options` object. Add to each:

```javascript
options: {
    // ... existing options ...
    onClick: (evt, elements, chart) => {
        if (App._chartAnnotationState.activeChartId !== chart.canvas.id) return;
        if (!elements.length) return;
        const el = elements[0];
        const xKey = chart.data.labels[el.index]; // assume label is YYYY-MM
        App._createChartAnnotation(chart.canvas.id, xKey, el);
    }
}
```

(If the existing chart options already define `onClick`, merge by calling the existing handler first, then dispatch annotation logic only when the mode is active.)

- [ ] **Step 2: Annotation creation**

Add to `App`:

```javascript
    _createChartAnnotation(chartId, xKey, chartEl) {
        const targetKey = `${chartId}:${xKey}`;
        const existing = DB.getNote('chart', targetKey);
        const body = prompt('Annotation text:', existing ? existing.body : '');
        if (body === null) return; // cancelled
        if (!body.trim()) {
            DB.deleteNote('chart', targetKey);
        } else {
            // Default position offset
            const meta = JSON.stringify({ dx: 30, dy: -30 });
            DB.setNote('chart', targetKey, body, meta);
        }
        // Exit mode
        this._toggleChartAnnotateMode(chartId);
        this._renderChartAnnotations(chartId);
    },
```

- [ ] **Step 3: Smoke-test**

Reload. Activate annotation mode on Cash Flow chart. Click any data point → prompt asks for text. Type "Test annotation" → confirms. (Rendering happens in next task; verify the row exists:)

```javascript
DB.getNote('chart', 'dashChartCashFlow:2026-03')
// Expected: object with body 'Test annotation'
```

Cleanup: `DB.deleteNote('chart','dashChartCashFlow:2026-03');`

- [ ] **Step 4: Commit**

```bash
git add js/app.js
git commit -m "feat(chart-annot): create annotations on chart point click"
```

---

### Task 5.3: Render annotations as overlays + drag positioning

**Files:**
- Modify: `js/app.js`
- Modify: `js/ui.js`

- [ ] **Step 1: Annotation overlay renderer**

Add to `App`:

```javascript
    _renderChartAnnotations(chartId) {
        // Remove existing overlay for this chart
        const oldOverlay = document.getElementById('chartAnnotOverlay_' + chartId);
        if (oldOverlay) oldOverlay.remove();

        const canvas = document.getElementById(chartId);
        if (!canvas) return;
        const chart = this._dashCharts ? this._dashCharts[chartId] : null;
        if (!chart) return;

        // Fetch all annotations for this chart
        const result = DB.db.exec(
            "SELECT target_key, body, metadata FROM notes WHERE target_type='chart' AND target_key LIKE ?",
            [chartId + ':%']
        );
        if (!result.length || !result[0].values.length) return;

        const overlay = document.createElement('div');
        overlay.className = 'chart-annotation-overlay';
        overlay.id = 'chartAnnotOverlay_' + chartId;
        const parent = canvas.parentElement;
        parent.style.position = 'relative';
        parent.appendChild(overlay);

        // Match overlay size to canvas
        const rect = canvas.getBoundingClientRect();
        const parentRect = parent.getBoundingClientRect();
        overlay.style.left = (canvas.offsetLeft) + 'px';
        overlay.style.top = (canvas.offsetTop) + 'px';
        overlay.style.width = canvas.offsetWidth + 'px';
        overlay.style.height = canvas.offsetHeight + 'px';

        // For each annotation, find pixel coords of the data point
        const meta = chart.scales || {};
        const xScale = meta.x;
        const yScale = meta.y;
        if (!xScale || !yScale) return;

        result[0].values.forEach(([targetKey, body, metaStr]) => {
            const xKey = targetKey.substring(chartId.length + 1);
            const idx = chart.data.labels.indexOf(xKey);
            if (idx < 0) return; // orphan — not rendered
            const dataset = chart.data.datasets[0];
            const value = dataset.data[idx];
            const px = xScale.getPixelForValue(idx);
            const py = yScale.getPixelForValue(value);

            let dx = 30, dy = -30;
            if (metaStr) {
                try { const o = JSON.parse(metaStr); dx = o.dx || 30; dy = o.dy || -30; } catch(e) {}
            }

            // Pin
            const pin = document.createElement('div');
            pin.className = 'chart-annotation-pin';
            pin.style.left = (px - 6) + 'px';
            pin.style.top = (py - 6) + 'px';
            pin.title = body;
            overlay.appendChild(pin);

            // Dashed line to x-axis
            const line = document.createElement('div');
            line.className = 'chart-annotation-line';
            line.style.left = px + 'px';
            line.style.top = py + 'px';
            line.style.height = (yScale.getPixelForValue(0) - py) + 'px';
            overlay.appendChild(line);

            // Callout (hidden in view-only — show on hover only)
            const callout = document.createElement('div');
            callout.className = 'chart-annotation-callout';
            const monthLabel = Utils.formatMonthDisplay(xKey);
            callout.innerHTML = `<strong>📍 ${monthLabel}</strong>${Utils.escapeHtml(body)}`;
            callout.style.left = (px + dx) + 'px';
            callout.style.top = (py + dy) + 'px';
            callout.dataset.targetKey = targetKey;
            if (this.isViewOnly) {
                callout.style.display = 'none';
                pin.addEventListener('mouseenter', () => { callout.style.display = 'block'; });
                pin.addEventListener('mouseleave', () => { callout.style.display = 'none'; });
            }
            overlay.appendChild(callout);

            // Drag handler for callout
            this._wireCalloutDrag(callout, targetKey, px, py);
        });
    },

    _wireCalloutDrag(callout, targetKey, basePx, basePy) {
        let dragStart = null;
        callout.addEventListener('mousedown', (e) => {
            const rect = callout.getBoundingClientRect();
            dragStart = { x: e.clientX, y: e.clientY, left: rect.left, top: rect.top };
            e.preventDefault();
        });
        document.addEventListener('mousemove', (e) => {
            if (!dragStart) return;
            const dx = e.clientX - dragStart.x;
            const dy = e.clientY - dragStart.y;
            callout.style.left = (parseFloat(callout.style.left) + dx) + 'px';
            callout.style.top = (parseFloat(callout.style.top) + dy) + 'px';
            dragStart.x = e.clientX; dragStart.y = e.clientY;
        });
        document.addEventListener('mouseup', () => {
            if (!dragStart) return;
            // Persist new position relative to base point
            const newDx = parseFloat(callout.style.left) - basePx;
            const newDy = parseFloat(callout.style.top) - basePy;
            const existing = DB.getNote('chart', targetKey);
            if (existing) {
                DB.setNote('chart', targetKey, existing.body, JSON.stringify({ dx: newDx, dy: newDy }));
            }
            dragStart = null;
        });
    },
```

- [ ] **Step 2: Trigger render after each chart draws**

In `_renderDashCashFlowChart`, `_renderDashPnLChart`, `_renderDashBSChart` — after the `new Chart(...)` call, append:

```javascript
        // Render annotations on top
        setTimeout(() => this._renderChartAnnotations('dashChartCashFlow'), 50);
```

(Replace chart id per renderer.)

- [ ] **Step 3: Smoke-test**

Reload Dashboard. Activate annotation mode on Cash Flow chart. Click a data point → prompt → type text. Annotation appears with pin, dashed line, callout. Drag the callout → repositions. Refresh — position persists. In view-only mode, the callout hides; hover the pin → shows.

- [ ] **Step 4: Commit**

```bash
git add js/app.js js/ui.js
git commit -m "feat(chart-annot): render and drag-position annotations on dashboard charts"
```

---

### Phase 5 Complete — Manual Acceptance Test

- [ ] **Verify Phase 5:**
1. Three "📌 Add annotation" buttons on dashboard line charts
2. Activating mode + clicking a point → prompt for text → annotation renders
3. Pin, dashed line, callout all visible
4. Drag callout → moves; reload page → position persists
5. View-only mode → callout hidden, pin visible, hover pin shows callout
6. Annotation appears in Insights Hub feed as a "📊 Chart Pin" card
7. Click insight card → switches to Dashboard

---

## Final acceptance — full layer

- [ ] **End-to-end demo flow:**
1. **Period Note:** add a P&L note for current month
2. **Variance:** find a flagged cell on P&L (or force one), click it, click "Explain this →" → period note popover opens for that (tab, month) with placeholder hint
3. **Anomaly:** dashboard widget shows ≥1 anomaly; click "Capture explanation" → period note popover opens
4. **Chart annotation:** activate mode on Cash Flow chart, click a data point, type explanation
5. **Insights Hub:** shows all the above as cards, filterable by type/tab/time, searchable
6. **View-only:** generate a share, open it → footnotes appear under spreadsheets, chart annotations hover-only, anomaly widget hidden, Insights tab hidden
7. **Sync:** push to a sync group; pull on another browser → all notes round-trip

---

## Self-review: spec coverage

| Spec section | Covered in tasks |
|---|---|
| `notes` table schema | 1.1 |
| Markdown parser with safety | 1.2 |
| Notes CRUD | 1.3 |
| Period note icon (hover-only, per-tab scope) | 1.4, 1.7, 1.8 |
| Anchored popover with drag/resize | 1.4, 1.5, 1.6 |
| Pin to share + cap warning | 1.5, 1.10 |
| View-only footnotes | 1.9 |
| Settings UI for thresholds and cap | 1.11 |
| Variance compute + tests | 2.1 |
| Variance flag styling + popover | 2.2, 2.3 |
| Variance dismissal (re-fires on value change) | 2.4 |
| "Explain this →" handoff to Period Notes | 2.4 |
| Anomaly rule functions (6 rules) + tests | 4.1 |
| Anomaly widget UI, severity tiers, top-3 + expander | 4.2 |
| Anomaly capture/view/dismiss + per-month dismissal scope | 4.2 |
| Insights tab in main strip | 3.1 |
| `getInsightsFeed` UNION query | 3.2 |
| Insights filters/search/timeline-grouped feed | 3.3 |
| Cross-tab navigation from cards | 3.4 |
| Chart annotation mode + click-on-point creation | 5.1, 5.2 |
| Chart annotation render + drag positioning | 5.3 |
| Chart annotation view-only hover | 5.3 |
| `notes.metadata` column for chart positions | 1.1 (added at table creation; spec said Phase 5 but combining is safe and simpler) |
| View-only Insights Hub hidden via hardcoded check | 4.2 implements widget hide; Insights tab hide handled by `applyViewOnlyRestrictions` — confirm in 1.11 / acceptance test |

**Note on metadata column:** Spec said the `metadata` column lands in Phase 5. The plan adds it in Phase 1 schema for simplicity (idempotent migration; no other type uses it until Phase 5). This is a small simplification that avoids a Phase-5 schema migration step.

**Note on view-only Insights Hub hiding:** The acceptance tests verify the tab is hidden when `App.isViewOnly === true`. To make this enforced, add to `applyViewOnlyRestrictions` in `js/app.js`:

```javascript
        // Hide Insights tab in view-only (hardcoded, not via hidden_tabs)
        const insightsTabBtn = document.querySelector('.main-tab-btn[data-tab="insights"]');
        if (insightsTabBtn) insightsTabBtn.style.display = 'none';
```

Add this as part of Task 3.1 ("Insights tab scaffold") so it ships when the tab does.
