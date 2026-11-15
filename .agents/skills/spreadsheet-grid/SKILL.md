---
name: spreadsheet-grid
description: "Build an inline-editable spreadsheet/data grid UI component in vanilla HTML/CSS/JS with sticky headers, frozen columns, cell editing, drag-drop rows, sections, subtotals, and projected/actual views. Use when asked to: create a spreadsheet, build a data grid, add an editable table, make a spreadsheet view, create inline editing table, build an Excel-like grid, add sticky headers table, create frozen column table, build a financial spreadsheet, add drag-drop table rows, create sortable editable table, build cash flow/budget/P&L table, add a grid component, create a data entry grid, editable grid, data table with editing, spreadsheet component, frozen columns, financial spreadsheet, accounting spreadsheet, budget table, P&L table."
metadata:
  author: user
  version: "1.0.0"
  argument-hint: <config-description>
---

# SpreadsheetGrid Skill

You are generating a **self-contained SpreadsheetGrid class** in vanilla HTML/CSS/JS. This component creates inline-editable spreadsheet tables with sticky headers, frozen columns, drag-drop row reordering, section grouping, subtotals, and projected/actual view toggling. It must work in ANY project with zero dependencies.

---

## Step 1: Gather Configuration (Ask One at a Time)

Before writing any code, ask the user these questions **one at a time**, waiting for each answer:

1. **Column definitions** — Ask: "What columns do you need? For each column, tell me: display name, data key, type (text / number / currency / date / percentage), whether it's editable, whether it should be sticky (frozen), and optional width. Example: `Name | name | text | editable | sticky | 200px`"

2. **Data source** — Ask: "How will data be provided? Options: (a) a JS array of objects I'll wire up, (b) a callback function you'll implement, or (c) sample/mock data for prototyping."

3. **Section grouping** — Ask: "Should rows be grouped into sections? If yes, what field determines the section (e.g., a `category` field with values like 'Revenue', 'Expenses')? Or is it a flat list?"

4. **Subtotals and totals** — Ask: "Do you need subtotal rows per section and/or a grand total row? Which numeric columns should be summed? Do you need any computed rows (e.g., 'Gross Profit = Revenue - COGS')?"

5. **Row operations** — Ask: "Do you need: drag-drop reorder within sections? Add-row button? Delete-row button? Any constraints?"

6. **Projected vs actual toggle** — Ask: "Do you need a toggle between 'Projected' and 'Actual' views where each row can have both projected and actual values?"

7. **Container element** — Ask: "What is the ID of the HTML element this grid should render into? (e.g., `spreadsheet-container`)"

Once all answers are gathered, generate exactly two files.

---

## Step 2: Files to Generate

### File 1: `spreadsheet-grid.css`

All styles for the grid. Uses CSS custom properties with fallbacks so it can inherit from any host theme.

### File 2: `spreadsheet-grid.js`

The `SpreadsheetGrid` class. Fully self-contained — no imports, no dependencies. Attaches to a container by ID.

---

## Step 3: CSS Architecture

### 3.1 CSS Variables with Fallbacks

The grid must define its own CSS variables that fall back to common host-app variable names. This ensures theme compatibility without requiring the host to do anything special.

```css
.sg-grid {
  /* Surface & background */
  --sg-surface: var(--surface, #ffffff);
  --sg-surface-alt: var(--surface-alt, #f9fafb);
  --sg-surface-hover: var(--surface-hover, #f3f4f6);

  /* Borders */
  --sg-border: var(--border-light, #e5e7eb);
  --sg-border-strong: var(--border, #d1d5db);

  /* Text */
  --sg-text: var(--text, #1f2937);
  --sg-text-muted: var(--text-muted, #6b7280);
  --sg-text-heading: var(--text-heading, #111827);

  /* Primary color (needs RGB for alpha usage) */
  --sg-primary-rgb: var(--color-primary-rgb, 59, 130, 246);

  /* Semantic colors */
  --sg-positive: var(--green, #059669);
  --sg-negative: var(--red, #dc2626);

  /* Typography */
  --sg-font: var(--font, system-ui, -apple-system, sans-serif);
  --sg-mono: var(--mono, 'SF Mono', 'Cascadia Code', 'Consolas', monospace);

  /* Sizing */
  --sg-header-height: 40px;
  --sg-row-height: 36px;
  --sg-cell-padding: 8px 12px;

  /* Shadows */
  --sg-shadow-sticky: 2px 0 4px rgba(0, 0, 0, 0.06);
  --sg-shadow-header: 0 2px 4px rgba(0, 0, 0, 0.06);
}
```

### 3.2 Z-Index Layering Scheme

This is critical for sticky positioning to work correctly. There are three layers:

```css
/* Layer 1: Sticky first column (frozen column) */
.sg-grid td.sg-sticky,
.sg-grid td.sg-frozen {
  position: sticky;
  left: 0;
  z-index: 1;
  background: var(--sg-surface);
  box-shadow: var(--sg-shadow-sticky);
}

/* Layer 2: Sticky header row */
.sg-grid thead th {
  position: sticky;
  top: 0;
  z-index: 2;
  background: var(--sg-surface-alt);
  box-shadow: var(--sg-shadow-header);
}

/* Layer 3: Header + frozen column intersection (top-left corner) */
.sg-grid thead th.sg-sticky,
.sg-grid thead th.sg-frozen {
  position: sticky;
  left: 0;
  top: 0;
  z-index: 3;
  background: var(--sg-surface-alt);
}
```

If there are **multiple sticky columns**, each gets an incrementing `left` offset:

```css
/* Example: two frozen columns, first is 200px wide */
.sg-grid .sg-frozen-0 {
  left: 0;
}
.sg-grid .sg-frozen-1 {
  left: 200px; /* must match width of column 0 */
}
```

Generate these dynamically in JS based on the column config and inject via a `<style>` tag.

### 3.3 Core Table Styles

```css
.sg-grid {
  font-family: var(--sg-font);
  font-size: 14px;
  color: var(--sg-text);
  overflow-x: auto;
  overflow-y: auto;
  -webkit-overflow-scrolling: touch;
  border: 1px solid var(--sg-border);
  border-radius: 8px;
}

.sg-grid table {
  width: 100%;
  border-collapse: collapse;
  table-layout: fixed;
}

.sg-grid th,
.sg-grid td {
  padding: var(--sg-cell-padding);
  border-bottom: 1px solid var(--sg-border);
  text-align: left;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  height: var(--sg-row-height);
  box-sizing: border-box;
}

.sg-grid th {
  font-weight: 600;
  font-size: 12px;
  text-transform: uppercase;
  letter-spacing: 0.05em;
  color: var(--sg-text-muted);
  height: var(--sg-header-height);
  user-select: none;
}
```

### 3.4 Numeric and Currency Formatting Styles

```css
.sg-grid .sg-number,
.sg-grid .sg-currency,
.sg-grid .sg-percentage {
  text-align: right;
  font-family: var(--sg-mono);
  font-variant-numeric: tabular-nums;
  font-size: 13px;
}

.sg-grid .sg-negative {
  color: var(--sg-negative);
}

.sg-grid .sg-positive-val {
  color: var(--sg-positive);
}
```

### 3.5 Section Headers, Subtotals, Totals

```css
/* Section header row — spans all columns */
.sg-grid .sg-section-header td {
  background: var(--sg-surface-alt);
  font-weight: 700;
  font-size: 13px;
  color: var(--sg-text-heading);
  letter-spacing: 0.02em;
  border-bottom: 2px solid var(--sg-border-strong);
  padding: 10px 12px;
}

/* Subtotal row at end of section */
.sg-grid .sg-subtotal td {
  background: var(--sg-surface-alt);
  font-weight: 600;
  border-top: 1px solid var(--sg-border-strong);
  border-bottom: 2px solid var(--sg-border-strong);
  font-style: italic;
}

/* Grand total row at very bottom */
.sg-grid .sg-total td {
  background: rgba(var(--sg-primary-rgb), 0.06);
  font-weight: 700;
  border-top: 2px solid rgba(var(--sg-primary-rgb), 0.3);
  border-bottom: 2px solid rgba(var(--sg-primary-rgb), 0.3);
  font-size: 14px;
}
```

### 3.6 Editable Cell Styles

```css
.sg-grid .sg-editable {
  cursor: text;
  position: relative;
}

.sg-grid .sg-editable:hover {
  background: rgba(var(--sg-primary-rgb), 0.04);
  outline: 1px solid rgba(var(--sg-primary-rgb), 0.2);
  outline-offset: -1px;
}

.sg-grid .sg-cell-input {
  position: absolute;
  inset: 0;
  width: 100%;
  height: 100%;
  border: 2px solid rgba(var(--sg-primary-rgb), 0.5);
  border-radius: 0;
  background: var(--sg-surface);
  font-family: var(--sg-mono);
  font-size: 13px;
  padding: var(--sg-cell-padding);
  box-sizing: border-box;
  outline: none;
  text-align: inherit;
  color: var(--sg-text);
}
```

### 3.7 Drag-Drop Styles

```css
.sg-grid tr.sg-dragging {
  opacity: 0.4;
}

.sg-grid tr.sg-drag-over {
  border-top: 2px solid rgba(var(--sg-primary-rgb), 0.7);
}

.sg-grid tr[draggable="true"] {
  cursor: grab;
}

.sg-grid tr[draggable="true"]:active {
  cursor: grabbing;
}

.sg-grid .sg-drag-handle {
  cursor: grab;
  color: var(--sg-text-muted);
  user-select: none;
  font-size: 16px;
  line-height: 1;
}
```

### 3.8 Projected vs Actual Styles

```css
.sg-grid .sg-projected {
  font-style: italic;
  background: rgba(var(--sg-primary-rgb), 0.04);
}

.sg-grid .sg-view-toggle {
  display: inline-flex;
  gap: 0;
  border: 1px solid var(--sg-border-strong);
  border-radius: 6px;
  overflow: hidden;
  margin-bottom: 8px;
}

.sg-grid .sg-view-toggle button {
  padding: 6px 16px;
  font-size: 13px;
  border: none;
  background: var(--sg-surface);
  color: var(--sg-text-muted);
  cursor: pointer;
  transition: background 0.15s, color 0.15s;
}

.sg-grid .sg-view-toggle button.sg-active {
  background: rgba(var(--sg-primary-rgb), 0.1);
  color: rgb(var(--sg-primary-rgb));
  font-weight: 600;
}
```

### 3.9 Empty State

```css
.sg-grid .sg-empty-state {
  text-align: center;
  padding: 40px 20px;
  color: var(--sg-text-muted);
  font-size: 14px;
  font-style: italic;
}
```

### 3.10 Responsive Breakpoints

```css
@media (max-width: 768px) {
  .sg-grid {
    font-size: 13px;
  }
  .sg-grid th {
    font-size: 11px;
  }
  .sg-grid .sg-number,
  .sg-grid .sg-currency,
  .sg-grid .sg-percentage {
    font-size: 12px;
  }
}

@media (max-width: 480px) {
  .sg-grid {
    font-size: 12px;
    --sg-cell-padding: 6px 8px;
    --sg-row-height: 32px;
    --sg-header-height: 36px;
  }
}
```

### 3.11 Read-Only Mode

```css
.sg-grid.sg-readonly .sg-editable {
  cursor: default;
  pointer-events: none;
}

.sg-grid.sg-readonly .sg-editable:hover {
  background: transparent;
  outline: none;
}

.sg-grid.sg-readonly .sg-drag-handle {
  display: none;
}
```

### 3.12 Row Action Buttons (Add/Delete)

```css
.sg-grid .sg-row-actions {
  display: flex;
  gap: 4px;
  align-items: center;
}

.sg-grid .sg-row-btn {
  width: 24px;
  height: 24px;
  border: none;
  border-radius: 4px;
  background: transparent;
  cursor: pointer;
  color: var(--sg-text-muted);
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 14px;
  transition: background 0.15s, color 0.15s;
}

.sg-grid .sg-row-btn:hover {
  background: var(--sg-surface-hover);
  color: var(--sg-text);
}

.sg-grid .sg-row-btn-delete:hover {
  background: rgba(220, 38, 38, 0.1);
  color: var(--sg-negative);
}

.sg-grid .sg-add-row-btn {
  display: block;
  width: 100%;
  padding: 8px;
  background: transparent;
  border: 1px dashed var(--sg-border-strong);
  border-radius: 6px;
  color: var(--sg-text-muted);
  cursor: pointer;
  font-size: 13px;
  margin-top: 4px;
  transition: background 0.15s, color 0.15s, border-color 0.15s;
}

.sg-grid .sg-add-row-btn:hover {
  background: rgba(var(--sg-primary-rgb), 0.04);
  color: rgb(var(--sg-primary-rgb));
  border-color: rgba(var(--sg-primary-rgb), 0.3);
}
```

---

## Step 4: JavaScript Architecture

### 4.1 Class Skeleton

```js
class SpreadsheetGrid {
  /**
   * @param {string} containerId — ID of the DOM element to render into
   * @param {object} config — full configuration object
   * @param {Array} config.columns — column definitions
   * @param {Array} config.sections — data grouped by section
   * @param {object} [config.totals] — total row configuration
   * @param {Array} [config.computedRows] — computed row definitions
   * @param {boolean} [config.dragDrop] — enable drag-drop reordering
   * @param {boolean} [config.addRow] — show add-row button per section
   * @param {boolean} [config.deleteRow] — show delete button per row
   * @param {boolean} [config.projectedActual] — enable projected/actual toggle
   * @param {boolean} [config.readOnly] — disable editing and drag-drop
   * @param {string} [config.viewMode] — 'actual' or 'projected' (default: 'actual')
   */
  constructor(containerId, config) {
    this.containerId = containerId;
    this.container = document.getElementById(containerId);
    if (!this.container) {
      throw new Error(`SpreadsheetGrid: container "#${containerId}" not found.`);
    }
    this.config = Object.assign({
      columns: [],
      sections: [],
      totals: null,
      computedRows: [],
      dragDrop: false,
      addRow: false,
      deleteRow: false,
      projectedActual: false,
      readOnly: false,
      viewMode: 'actual',
    }, config);

    this._draggedRow = null;
    this._setupDone = false;
  }

  render() {
    const html = this._buildHTML();
    this.container.innerHTML = html;
    this._injectDynamicStyles();
    this._setupEventDelegation();
    if (this.config.readOnly) {
      this.container.querySelector('.sg-grid')?.classList.add('sg-readonly');
    }
  }

  setData(newSections) {
    this.config.sections = newSections;
    this.render();
  }

  setViewMode(mode) {
    this.config.viewMode = mode;
    this.render();
  }

  destroy() {
    this.container.innerHTML = '';
    this._setupDone = false;
    delete this.container.dataset.sgSetup;
  }

  // --- Private methods below ---
}
```

### 4.2 Column Configuration Shape

Each column in `config.columns` should be an object:

```js
{
  name: 'Account',        // display header
  key: 'account',         // data key in row objects
  type: 'text',           // 'text' | 'number' | 'currency' | 'date' | 'percentage'
  editable: false,        // can users edit this cell?
  sticky: true,           // freeze this column on horizontal scroll?
  width: '200px',         // explicit width (or null for auto)
  align: 'left',          // 'left' | 'right' | 'center' (auto-inferred from type if omitted)
}
```

### 4.3 Section Data Shape

```js
config.sections = [
  {
    id: 'revenue',
    label: 'Revenue',
    rows: [
      { id: 'row1', account: 'Product Sales', jan: 50000, feb: 52000, projected: { jan: 48000, feb: 51000 } },
      { id: 'row2', account: 'Service Revenue', jan: 15000, feb: 16000, projected: { jan: 14000, feb: 15500 } },
    ],
    subtotal: true,  // auto-calculate subtotals for numeric columns
  },
  {
    id: 'expenses',
    label: 'Expenses',
    rows: [ /* ... */ ],
    subtotal: true,
  }
];
```

### 4.4 HTML Generation via String Concatenation

Build the entire HTML as a string. Never use `document.createElement` for the initial render — string concatenation is faster for bulk generation.

```js
_buildHTML() {
  const cols = this.config.columns;
  const sections = this.config.sections;
  const showActions = this.config.deleteRow || this.config.dragDrop;

  // Empty state check
  const hasData = sections.some(s => s.rows && s.rows.length > 0);
  if (!hasData && !this.config.addRow) {
    return '<div class="sg-grid"><p class="sg-empty-state">No data available.</p></div>';
  }

  let html = '<div class="sg-grid">';

  // View mode toggle
  if (this.config.projectedActual) {
    html += this._buildViewToggle();
  }

  html += '<table>';

  // Colgroup for explicit widths
  html += '<colgroup>';
  if (this.config.dragDrop) html += '<col style="width:36px">';
  cols.forEach(col => {
    html += `<col${col.width ? ` style="width:${col.width}"` : ''}>`;
  });
  if (showActions) html += '<col style="width:60px">';
  html += '</colgroup>';

  // Thead
  html += '<thead><tr>';
  if (this.config.dragDrop) html += '<th class="sg-sticky sg-frozen-0"></th>';
  cols.forEach((col, i) => {
    const stickyClass = col.sticky ? ` sg-sticky sg-frozen-${this.config.dragDrop ? i + 1 : i}` : '';
    const alignClass = this._alignClass(col);
    html += `<th class="${alignClass}${stickyClass}">${this._escapeHtml(col.name)}</th>`;
  });
  if (showActions) html += '<th></th>';
  html += '</tr></thead>';

  // Tbody — iterate sections
  html += '<tbody>';
  sections.forEach(section => {
    html += this._buildSectionHTML(section);
  });

  // Computed rows (e.g., Gross Profit = Revenue - Expenses)
  this.config.computedRows.forEach(computed => {
    html += this._buildComputedRow(computed);
  });

  // Grand total row
  if (this.config.totals) {
    html += this._buildTotalRow();
  }

  html += '</tbody></table>';

  html += '</div>';
  return html;
}
```

### 4.5 Building a Section

```js
_buildSectionHTML(section) {
  const cols = this.config.columns;
  const totalCols = cols.length + (this.config.dragDrop ? 1 : 0) + ((this.config.deleteRow || this.config.dragDrop) ? 1 : 0);
  let html = '';

  // Section header
  html += `<tr class="sg-section-header" data-section="${section.id}">`;
  html += `<td colspan="${totalCols}">${this._escapeHtml(section.label)}</td>`;
  html += '</tr>';

  // Data rows
  if (section.rows) {
    section.rows.forEach((row, rowIndex) => {
      html += this._buildDataRow(row, section.id, rowIndex);
    });
  }

  // Add-row button
  if (this.config.addRow) {
    html += `<tr class="sg-add-row-tr" data-section="${section.id}">`;
    html += `<td colspan="${totalCols}">`;
    html += `<button class="sg-add-row-btn" data-section="${section.id}">+ Add Row</button>`;
    html += '</td></tr>';
  }

  // Subtotal row
  if (section.subtotal && section.rows && section.rows.length > 0) {
    html += this._buildSubtotalRow(section);
  }

  return html;
}
```

### 4.6 Building a Data Row

```js
_buildDataRow(row, sectionId, rowIndex) {
  const cols = this.config.columns;
  const mode = this.config.viewMode;
  const isProjected = mode === 'projected';
  const projectedClass = isProjected ? ' sg-projected' : '';

  let html = `<tr class="sg-data-row${projectedClass}" data-row-id="${row.id}" data-section="${sectionId}" data-row-index="${rowIndex}"${this.config.dragDrop ? ' draggable="true"' : ''}>`;

  // Drag handle column
  if (this.config.dragDrop) {
    html += `<td class="sg-sticky sg-frozen-0"><span class="sg-drag-handle">&#x2630;</span></td>`;
  }

  // Data cells
  cols.forEach((col, colIndex) => {
    let value;
    if (isProjected && row.projected && row.projected[col.key] !== undefined) {
      value = row.projected[col.key];
    } else {
      value = row[col.key];
    }

    const stickyClass = col.sticky ? ` sg-sticky sg-frozen-${this.config.dragDrop ? colIndex + 1 : colIndex}` : '';
    const typeClass = ` sg-${col.type || 'text'}`;
    const editableClass = col.editable && !this.config.readOnly ? ' sg-editable' : '';
    const valueClass = this._valueClass(col, value);

    const formatted = this._formatValue(value, col.type);

    html += `<td class="${typeClass}${stickyClass}${editableClass}${valueClass}" data-col-key="${col.key}" data-raw-value="${value ?? ''}">${formatted}</td>`;
  });

  // Action buttons
  if (this.config.deleteRow || this.config.dragDrop) {
    html += '<td class="sg-row-actions">';
    if (this.config.deleteRow) {
      html += `<button class="sg-row-btn sg-row-btn-delete" data-action="delete" data-row-id="${row.id}" title="Delete row">&#x2715;</button>`;
    }
    html += '</td>';
  }

  html += '</tr>';
  return html;
}
```

### 4.7 Subtotal and Total Row Generation

```js
_buildSubtotalRow(section) {
  const cols = this.config.columns;
  const mode = this.config.viewMode;
  const totalCols = (this.config.dragDrop ? 1 : 0);
  let html = '<tr class="sg-subtotal">';

  if (this.config.dragDrop) html += '<td></td>';

  cols.forEach((col, i) => {
    if (i === 0) {
      const stickyClass = col.sticky ? ` sg-sticky sg-frozen-${this.config.dragDrop ? 1 : 0}` : '';
      html += `<td class="${stickyClass}">Subtotal</td>`;
    } else if (['number', 'currency', 'percentage'].includes(col.type)) {
      const sum = this._sumColumn(section.rows, col.key, mode);
      const formatted = this._formatValue(sum, col.type);
      const valueClass = this._valueClass(col, sum);
      html += `<td class="sg-${col.type}${valueClass}">${formatted}</td>`;
    } else {
      html += '<td></td>';
    }
  });

  if (this.config.deleteRow || this.config.dragDrop) html += '<td></td>';
  html += '</tr>';
  return html;
}

_buildTotalRow() {
  const cols = this.config.columns;
  const mode = this.config.viewMode;
  let html = '<tr class="sg-total">';

  if (this.config.dragDrop) html += '<td></td>';

  cols.forEach((col, i) => {
    if (i === 0) {
      const stickyClass = col.sticky ? ` sg-sticky sg-frozen-${this.config.dragDrop ? 1 : 0}` : '';
      html += `<td class="${stickyClass}"><strong>${this.config.totals.label || 'Total'}</strong></td>`;
    } else if (['number', 'currency', 'percentage'].includes(col.type)) {
      const allRows = this.config.sections.flatMap(s => s.rows || []);
      const sum = this._sumColumn(allRows, col.key, mode);
      const formatted = this._formatValue(sum, col.type);
      const valueClass = this._valueClass(col, sum);
      html += `<td class="sg-${col.type}${valueClass}"><strong>${formatted}</strong></td>`;
    } else {
      html += '<td></td>';
    }
  });

  if (this.config.deleteRow || this.config.dragDrop) html += '<td></td>';
  html += '</tr>';
  return html;
}

_buildComputedRow(computed) {
  // computed = { label: 'Gross Profit', formula: (sections) => value, type: 'currency', class: 'sg-subtotal' }
  const cols = this.config.columns;
  let html = `<tr class="${computed.class || 'sg-subtotal'}">`;

  if (this.config.dragDrop) html += '<td></td>';

  cols.forEach((col, i) => {
    if (i === 0) {
      const stickyClass = col.sticky ? ` sg-sticky sg-frozen-${this.config.dragDrop ? 1 : 0}` : '';
      html += `<td class="${stickyClass}"><strong>${this._escapeHtml(computed.label)}</strong></td>`;
    } else if (['number', 'currency', 'percentage'].includes(col.type) && typeof computed.formula === 'function') {
      const value = computed.formula(this.config.sections, col.key, this.config.viewMode);
      const formatted = this._formatValue(value, col.type);
      const valueClass = this._valueClass(col, value);
      html += `<td class="sg-${col.type}${valueClass}"><strong>${formatted}</strong></td>`;
    } else {
      html += '<td></td>';
    }
  });

  if (this.config.deleteRow || this.config.dragDrop) html += '<td></td>';
  html += '</tr>';
  return html;
}
```

### 4.8 View Mode Toggle HTML

```js
_buildViewToggle() {
  const mode = this.config.viewMode;
  return `
    <div class="sg-view-toggle">
      <button class="${mode === 'actual' ? 'sg-active' : ''}" data-view-mode="actual">Actual</button>
      <button class="${mode === 'projected' ? 'sg-active' : ''}" data-view-mode="projected">Projected</button>
    </div>`;
}
```

### 4.9 Formatting Utilities

```js
_formatValue(value, type) {
  if (value === null || value === undefined || value === '') return '';

  switch (type) {
    case 'currency': {
      const num = typeof value === 'number' ? value : parseFloat(value);
      if (isNaN(num)) return this._escapeHtml(String(value));
      const isNeg = num < 0;
      const abs = Math.abs(num);
      const formatted = '$' + abs.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      return isNeg ? `(${formatted})` : formatted;
    }
    case 'number': {
      const num = typeof value === 'number' ? value : parseFloat(value);
      if (isNaN(num)) return this._escapeHtml(String(value));
      return num.toLocaleString('en-US', { maximumFractionDigits: 2 });
    }
    case 'percentage': {
      const num = typeof value === 'number' ? value : parseFloat(value);
      if (isNaN(num)) return this._escapeHtml(String(value));
      return num.toFixed(1) + '%';
    }
    case 'date':
      return this._escapeHtml(String(value));
    default:
      return this._escapeHtml(String(value));
  }
}

_valueClass(col, value) {
  if (!['number', 'currency', 'percentage'].includes(col.type)) return '';
  const num = typeof value === 'number' ? value : parseFloat(value);
  if (isNaN(num)) return '';
  if (num < 0) return ' sg-negative';
  return '';
}

_alignClass(col) {
  if (col.align) return `sg-align-${col.align}`;
  if (['number', 'currency', 'percentage'].includes(col.type)) return 'sg-align-right';
  return 'sg-align-left';
}

_escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

_sumColumn(rows, key, mode) {
  return rows.reduce((sum, row) => {
    let val;
    if (mode === 'projected' && row.projected && row.projected[key] !== undefined) {
      val = row.projected[key];
    } else {
      val = row[key];
    }
    const num = typeof val === 'number' ? val : parseFloat(val);
    return sum + (isNaN(num) ? 0 : num);
  }, 0);
}
```

### 4.10 Dynamic Styles Injection for Sticky Columns

```js
_injectDynamicStyles() {
  // Remove previously injected style tag
  const existingStyle = document.getElementById(`sg-dynamic-${this.containerId}`);
  if (existingStyle) existingStyle.remove();

  const stickyColumns = this.config.columns
    .map((col, i) => ({ col, index: this.config.dragDrop ? i + 1 : i }))
    .filter(({ col }) => col.sticky);

  if (stickyColumns.length === 0 && !this.config.dragDrop) return;

  let css = '';
  let leftOffset = 0;

  // Drag handle column is always first sticky if drag-drop is on
  if (this.config.dragDrop) {
    css += `#${this.containerId} .sg-frozen-0 { left: 0; }\n`;
    leftOffset = 36; // drag handle width
  }

  stickyColumns.forEach(({ col, index }) => {
    css += `#${this.containerId} .sg-frozen-${index} { left: ${leftOffset}px; }\n`;
    leftOffset += parseInt(col.width || '150', 10);
  });

  const style = document.createElement('style');
  style.id = `sg-dynamic-${this.containerId}`;
  style.textContent = css;
  document.head.appendChild(style);
}
```

### 4.11 Event Delegation Setup (Cell Editing)

This is the most critical pattern. Use a **one-time setup guard** so re-renders do not stack event listeners.

```js
_setupEventDelegation() {
  // ONE-TIME SETUP GUARD: prevent stacking listeners on re-render
  if (this.container.dataset.sgSetup === '1') return;
  this.container.dataset.sgSetup = '1';

  const grid = this;

  // --- Cell Editing ---
  this.container.addEventListener('click', function(e) {
    const cell = e.target.closest('.sg-editable');
    if (!cell) return;

    // Guard: already editing this cell
    if (cell.querySelector('.sg-cell-input')) return;

    const colKey = cell.dataset.colKey;
    const col = grid.config.columns.find(c => c.key === colKey);
    if (!col) return;

    const row = cell.closest('tr');
    const rowId = row ? row.dataset.rowId : null;

    // Get raw value and strip currency/percentage formatting
    const rawValue = cell.dataset.rawValue || cell.textContent.replace(/[^0-9.\-]/g, '');
    const originalText = cell.textContent;

    // Determine input type
    const inputType = ['number', 'currency', 'percentage'].includes(col.type) ? 'number' : 'text';

    // Create input element
    const input = document.createElement('input');
    input.type = inputType;
    input.className = 'sg-cell-input';
    input.value = rawValue;
    if (inputType === 'number') {
      input.step = 'any';
    }

    // Clear cell text and append input
    cell.textContent = '';
    cell.appendChild(input);
    input.focus();
    input.select();

    function save() {
      const newValue = inputType === 'number' ? parseFloat(input.value) : input.value;

      // Find the data row and update it
      let changed = false;
      for (const section of grid.config.sections) {
        const dataRow = section.rows.find(r => r.id === rowId);
        if (dataRow) {
          const mode = grid.config.viewMode;
          if (mode === 'projected' && dataRow.projected) {
            if (dataRow.projected[colKey] !== newValue) {
              dataRow.projected[colKey] = newValue;
              changed = true;
            }
          } else {
            if (dataRow[colKey] !== newValue) {
              dataRow[colKey] = newValue;
              changed = true;
            }
          }
          break;
        }
      }

      if (changed) {
        // Dispatch custom event
        grid.container.dispatchEvent(new CustomEvent('sg:celledit', {
          bubbles: true,
          detail: { rowId, colKey, value: newValue, viewMode: grid.config.viewMode }
        }));
      }

      // Re-render (this safely replaces everything)
      grid.render();
    }

    function cancel() {
      // Restore original text without saving
      cell.textContent = originalText;
    }

    input.addEventListener('blur', save);

    input.addEventListener('keydown', function(ke) {
      if (ke.key === 'Enter') {
        ke.preventDefault();
        input.blur(); // triggers save via blur handler
      } else if (ke.key === 'Escape') {
        ke.preventDefault();
        input.removeEventListener('blur', save); // prevent save
        cancel();
      }
    });
  });

  // --- View Mode Toggle ---
  this.container.addEventListener('click', function(e) {
    const toggleBtn = e.target.closest('[data-view-mode]');
    if (!toggleBtn) return;
    const mode = toggleBtn.dataset.viewMode;
    if (mode !== grid.config.viewMode) {
      grid.setViewMode(mode);
    }
  });

  // --- Add Row ---
  this.container.addEventListener('click', function(e) {
    const addBtn = e.target.closest('.sg-add-row-btn');
    if (!addBtn) return;
    const sectionId = addBtn.dataset.section;
    const section = grid.config.sections.find(s => s.id === sectionId);
    if (!section) return;

    // Create blank row with unique id
    const newRow = { id: 'row_' + Date.now() };
    grid.config.columns.forEach(col => {
      newRow[col.key] = col.type === 'text' ? '' : 0;
    });
    if (grid.config.projectedActual) {
      newRow.projected = {};
      grid.config.columns.forEach(col => {
        if (['number', 'currency', 'percentage'].includes(col.type)) {
          newRow.projected[col.key] = 0;
        }
      });
    }

    section.rows = section.rows || [];
    section.rows.push(newRow);

    grid.container.dispatchEvent(new CustomEvent('sg:rowadd', {
      bubbles: true,
      detail: { sectionId, row: newRow }
    }));

    grid.render();
  });

  // --- Delete Row ---
  this.container.addEventListener('click', function(e) {
    const delBtn = e.target.closest('.sg-row-btn-delete');
    if (!delBtn) return;
    const rowId = delBtn.dataset.rowId;

    for (const section of grid.config.sections) {
      const idx = (section.rows || []).findIndex(r => r.id === rowId);
      if (idx !== -1) {
        const removed = section.rows.splice(idx, 1)[0];
        grid.container.dispatchEvent(new CustomEvent('sg:rowdelete', {
          bubbles: true,
          detail: { sectionId: section.id, row: removed }
        }));
        break;
      }
    }

    grid.render();
  });

  // --- Drag-Drop Row Reordering ---
  if (this.config.dragDrop) {
    this.container.addEventListener('dragstart', function(e) {
      const row = e.target.closest('tr.sg-data-row');
      if (!row) return;
      grid._draggedRow = row;
      row.classList.add('sg-dragging');
      e.dataTransfer.effectAllowed = 'move';
      e.dataTransfer.setData('text/plain', row.dataset.rowId);
    });

    this.container.addEventListener('dragover', function(e) {
      e.preventDefault();
      const targetRow = e.target.closest('tr.sg-data-row');
      if (!targetRow || !grid._draggedRow) return;

      // Section constraint: only reorder within same section
      if (targetRow.dataset.section !== grid._draggedRow.dataset.section) return;

      e.dataTransfer.dropEffect = 'move';

      // Remove drag-over from all rows, then add to target
      grid.container.querySelectorAll('.sg-drag-over').forEach(el => el.classList.remove('sg-drag-over'));
      targetRow.classList.add('sg-drag-over');
    });

    this.container.addEventListener('drop', function(e) {
      e.preventDefault();
      const targetRow = e.target.closest('tr.sg-data-row');
      if (!targetRow || !grid._draggedRow) return;
      if (targetRow.dataset.section !== grid._draggedRow.dataset.section) return;
      if (targetRow === grid._draggedRow) return;

      const sectionId = grid._draggedRow.dataset.section;
      const draggedId = grid._draggedRow.dataset.rowId;
      const targetId = targetRow.dataset.rowId;

      // Update the data model
      const section = grid.config.sections.find(s => s.id === sectionId);
      if (section && section.rows) {
        const fromIdx = section.rows.findIndex(r => r.id === draggedId);
        const toIdx = section.rows.findIndex(r => r.id === targetId);
        if (fromIdx !== -1 && toIdx !== -1) {
          const [moved] = section.rows.splice(fromIdx, 1);
          section.rows.splice(toIdx, 0, moved);

          grid.container.dispatchEvent(new CustomEvent('sg:rowreorder', {
            bubbles: true,
            detail: { sectionId, rowId: draggedId, fromIndex: fromIdx, toIndex: toIdx }
          }));
        }
      }

      // Also move in DOM for immediate visual feedback
      targetRow.parentNode.insertBefore(grid._draggedRow, targetRow);

      // Re-render to sync fully
      grid.render();
    });

    this.container.addEventListener('dragend', function(e) {
      // Cleanup all drag classes
      grid.container.querySelectorAll('.sg-dragging').forEach(el => el.classList.remove('sg-dragging'));
      grid.container.querySelectorAll('.sg-drag-over').forEach(el => el.classList.remove('sg-drag-over'));
      grid._draggedRow = null;
    });
  }
}
```

---

## Step 5: Integration Example

After generating the two files, provide a usage example:

```html
<link rel="stylesheet" href="spreadsheet-grid.css">
<script src="spreadsheet-grid.js"></script>

<div id="my-grid"></div>

<script>
  const grid = new SpreadsheetGrid('my-grid', {
    columns: [
      { name: 'Account', key: 'account', type: 'text', sticky: true, width: '200px' },
      { name: 'Jan', key: 'jan', type: 'currency', editable: true, width: '120px' },
      { name: 'Feb', key: 'feb', type: 'currency', editable: true, width: '120px' },
      { name: 'Mar', key: 'mar', type: 'currency', editable: true, width: '120px' },
    ],
    sections: [
      {
        id: 'revenue',
        label: 'Revenue',
        subtotal: true,
        rows: [
          { id: 'r1', account: 'Product Sales', jan: 50000, feb: 52000, mar: 54000 },
          { id: 'r2', account: 'Services', jan: 15000, feb: 16000, mar: 17000 },
        ]
      },
      {
        id: 'expenses',
        label: 'Expenses',
        subtotal: true,
        rows: [
          { id: 'e1', account: 'Salaries', jan: -30000, feb: -30000, mar: -31000 },
          { id: 'e2', account: 'Rent', jan: -5000, feb: -5000, mar: -5000 },
        ]
      }
    ],
    computedRows: [
      {
        label: 'Net Income',
        class: 'sg-total',
        formula: (sections, colKey, viewMode) => {
          return sections.reduce((sum, section) => {
            return sum + section.rows.reduce((s, row) => {
              const val = viewMode === 'projected' && row.projected ? (row.projected[colKey] || 0) : (row[colKey] || 0);
              return s + (typeof val === 'number' ? val : 0);
            }, 0);
          }, 0);
        }
      }
    ],
    dragDrop: true,
    addRow: true,
    deleteRow: true,
    projectedActual: false,
  });

  grid.render();

  // Listen for changes
  document.getElementById('my-grid').addEventListener('sg:celledit', (e) => {
    console.log('Cell edited:', e.detail);
  });
</script>
```

---

## Step 6: Checklist Before Delivering

Before presenting the code to the user, verify:

- [ ] CSS file has all variable fallbacks defined in `.sg-grid`
- [ ] Z-index layers: 1 (sticky col), 2 (sticky header), 3 (intersection)
- [ ] Event delegation uses `container.dataset.sgSetup = '1'` one-time guard
- [ ] Cell editing strips formatting with `.replace(/[^0-9.\-]/g, '')`
- [ ] Cell input: blur saves, Enter triggers blur, Escape cancels (removes blur listener first)
- [ ] Drag-drop enforces same-section constraint via `dataset.section`
- [ ] All HTML built via string concatenation, assigned with single `innerHTML`
- [ ] `font-variant-numeric: tabular-nums` on numeric cells
- [ ] Negative values get `.sg-negative` class (red color)
- [ ] Currency formatted as `$X,XXX.XX` with parentheses for negatives
- [ ] Empty state shows when no data and no add-row
- [ ] `overflow-x: auto` on container, `-webkit-overflow-scrolling: touch`
- [ ] Read-only mode adds `.sg-readonly` and disables pointer-events on editables
- [ ] Responsive styles at 768px and 480px breakpoints
- [ ] Custom events dispatched: `sg:celledit`, `sg:rowreorder`, `sg:rowadd`, `sg:rowdelete`
- [ ] `destroy()` cleans up innerHTML and resets setup guard
