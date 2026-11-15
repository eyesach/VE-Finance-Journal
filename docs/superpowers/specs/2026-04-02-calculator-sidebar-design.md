# Calculator Sidebar — Design Spec

**Date:** 2026-04-02
**Status:** Draft

## Problem

When creating financial reports, the user needs to quickly find totals for certain months combined or across time periods. Currently there's no way to select multiple spreadsheet cells and see a sum, build custom expressions (add some values, subtract others, divide by a number), or do quick math on the data without exporting to a spreadsheet tool.

## Solution

A toggleable right-side calculator sidebar that lets users click cells across any tab to build a math expression with a live-updating result. Combines the quick-sum convenience of Google Sheets' status bar with the power of an editable formula bar.

## Design Decisions

| Decision | Choice | Alternatives Considered |
|----------|--------|------------------------|
| Activation | Dedicated "Calc" toggle button in header | Keyboard shortcut only, always-on |
| Placement | Right sidebar (~280px) | Bottom bar, floating panel |
| Expression style | Text-based formula bar (contenteditable) | Token/chip-based drag editor |
| Operator precedence | Standard math (×÷ before +−) | Left-to-right (basic calculator) |
| Cell interaction | Click to select (edit mode disabled) | Ctrl+Click additive (edit mode preserved) |

## Architecture

### 1. Calculator Mode Toggle

- A "Calc" button in the app header bar (near Save/Load/Export buttons)
- Clicking toggles calculator mode on/off
- When on: right sidebar slides in, main content area shrinks, all spreadsheet cells become selectable
- When off: sidebar slides out, content returns to full width, normal cell editing resumes
- Button gets an `active` CSS class when mode is on

### 2. Sidebar Layout (top to bottom)

1. **Header** — "Calculator" title + close (✕) button
2. **Result display** — large formatted currency number, live-updating
3. **Formula bar** — `contenteditable` div showing the expression
   - Cell-sourced values are colored green, hoverable to show source (category · month)
   - Manual numbers are plain/white
   - User can freely type: numbers, `+`, `−`, `×`/`*`, `÷`/`/`, `(`, `)`, spaces
   - Result updates live as expression changes
   - Shows "Error" for malformed expressions (unmatched parens, divide by zero)
4. **Cell reference list** — read-only list of all cell references in the formula
   - Each row: value + source label (e.g., "$5,200 — Revenue · Jan 2025")
   - Clicking a reference scrolls the formula bar to that span and applies a brief highlight flash (1s background animation)
5. **Manual input** — number input field + four operator buttons (+, −, ×, ÷) to insert manual values
6. **Footer** — Clear, Copy Result (formatted currency), Copy Formula (raw text)

### 3. Cell Selection by Tab

Clicking a numeric cell in calculator mode inserts its value into the formula bar at the cursor position.

| Tab | Cell selector | Data attributes | Label format |
|-----|--------------|-----------------|--------------|
| P&L | `.pnl-calc-cell` (new class, added to ALL month cells) | `data-cat-id`, `data-month` | "Revenue · Jan 2025" |
| Cash Flow | `.cf-calc-cell` (new class, added to ALL month cells — not just future) | `data-cat-id`, `data-month` | "Rent · Mar 2025" |
| Balance Sheet | `.bs-value` (new class) | `data-bs-key` (new attr) | "Cash", "Total AR" |
| Journal | `.txn-amount` (new class) | `data-txn-id` (new attr) | "Invoice #1042" |
| Budget | `.budget-amount` (new class) | `data-budget-id` (new attr) | "Marketing · Jan" |
| Loans | `.loan-amount` (new class) | `data-loan-id`, `data-month` | "Loan A · Mar" |
| Fixed Assets | `.asset-amount` (new class) | `data-asset-id` | "Equipment" |
| Projected Sales | `.ps-amount` (new class) | `data-ps-month` | "Projected · Feb" |

### 4. Formula Bar Behavior

**Inserting from cells:**
- Click a cell → value inserted at cursor position in formula bar
- If formula is empty, value becomes the entire expression
- Cell values rendered as colored spans with `data-*` attributes for source tracking
- Hover on a cell value in the formula shows a tooltip with category + month

**Editing:**
- Formula bar is a `contenteditable` div for inline color-coding
- Typing operators (`+`, `-`, `*`, `/`), numbers, and parentheses works naturally
- Cell-reference spans are **atomic** (`contenteditable="false"` on the span) — Backspace/Delete adjacent to one removes the entire span, cursor cannot enter it
- Paste events are intercepted to insert **plain text only** (strip HTML)
- Result recalculates on every `input` event

**Expression evaluation:**
- Custom recursive-descent parser handling: `+`, `−`, `×`/`*`, `÷`/`/`, parentheses
- `eval()` and `Function()` must NOT be used — custom parser only
- Standard operator precedence (multiply/divide before add/subtract)
- Returns formatted currency result or "Error" string
- Handles edge cases: empty expression (shows $0.00), divide by zero (Error), unmatched parens (Error)

### 5. Visual Feedback on Spreadsheets

When calculator mode is active:
- All numeric cells get a subtle hover effect (light dashed border) indicating selectability
- Cells currently referenced in the formula get a **blue highlight outline** (2px solid)
- Cursor changes to `crosshair` over selectable cells
- Existing cell edit behavior (click → inline input) is **disabled**
- On tabs without editable cells (Balance Sheet), the highlight still works on value cells

### 6. Keyboard & Interaction

- No modifier keys needed for basic selection (just click)
- `Escape` closes calculator mode
- Formula bar supports standard text editing (arrow keys, Home/End, Shift+select, Ctrl+A)
- Keyboard shortcuts (Ctrl+Z undo etc.) are NOT intercepted while formula bar is focused

### 7. State

- Calculator state (formula, selections) lives in memory only — clears when sidebar closes
- No persistence to database or IndexedDB
- Switching tabs preserves the current formula (can select cells across tabs)
- Cell values are **snapshots** captured at selection time — if underlying data changes (e.g., via quick-entry modal), the formula retains the original value
- Modals and quick-entry forms remain accessible in calc mode (only direct cell editing is disabled)

## Files to Modify

| File | Changes |
|------|---------|
| `index.html` | Add Calc button to header, add calculator sidebar HTML structure |
| `js/app.js` | Calculator mode toggle, cell click handler delegation, formula state management, expression evaluator |
| `js/ui.js` | Add `data-*` attributes and new CSS classes to numeric cells across all tab renderers |
| `css/styles.css` | Sidebar styles, calc-mode cell highlights, formula bar styling, responsive adjustments |

## Responsive Considerations

- On screens ≤1024px: if left sidebar is visible, it auto-collapses when calculator sidebar opens (to preserve spreadsheet width)
- On screens ≤768px: sidebar becomes a full-width bottom sheet (slides up from bottom) instead of a side panel
- On screens ≤480px: formula bar font size reduces, cell reference list collapses to count-only

## Out of Scope

- Saving/loading expressions
- Named cell references (like spreadsheet A1 notation)
- Functions (SUM, AVG, etc.) — the manual approach covers these use cases
- Cross-company calculations
