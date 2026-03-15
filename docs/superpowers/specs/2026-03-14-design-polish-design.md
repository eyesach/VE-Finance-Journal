# Design Polish & Visual Refinement Spec

**Date:** 2026-03-14
**Status:** Approved
**Scope:** Visual polish, micro-interactions, and element-level design refinements
**Approach:** Single CSS pass (Approach A) — one cohesive commit

---

## Overview

Elevate the Accounting Journal Calculator's visual quality from "functional" to "premium" through targeted design token updates, component refinements, and micro-detail additions. The direction is **Refined Professional + Modern tones** — think Stripe/Linear confidence with Vercel-style card depth.

Minor spacing nudges (1-2px from 4px grid snapping) are acceptable. No structural layout changes. No new features. Pure visual polish.

---

## Design Direction

- **Typography:** Keep DM Sans + DM Mono. Refine usage: tighter type scale, better weight contrast between hierarchy levels, tabular numbers for financial data.
- **Surfaces:** Defined Elevation — cards/panels get meaningful shadow depth (`0 2px 4px rgba(0,0,0,0.04), 0 8px 24px rgba(0,0,0,0.06)`) with near-invisible borders.
- **Buttons:** Soft Depth — primary buttons get a colored shadow glow, secondary buttons get subtle lift, everything feels tactile.
- **Tables:** Active Row Highlight — clean base with blue left-border accent + light blue background on hover/selected rows.

---

## Section 1: Design Token Updates

### Shadows
```css
--shadow-sm: 0 1px 3px rgba(0,0,0,0.04), 0 4px 12px rgba(0,0,0,0.03);
--shadow-md: 0 2px 4px rgba(0,0,0,0.04), 0 8px 24px rgba(0,0,0,0.06);
--shadow-primary: 0 1px 2px rgba(59,130,246,0.3), 0 2px 8px rgba(59,130,246,0.15);  /* new */
--shadow-focus: 0 0 0 2px rgba(59,130,246,0.3);  /* new */
```

### Dark mode shadow equivalents (in `[data-theme="dark"]`)
These **replace** the existing dark mode `--shadow-sm` and `--shadow-md` values.
```css
--shadow-sm: 0 1px 3px rgba(0,0,0,0.3), 0 4px 12px rgba(0,0,0,0.2);
--shadow-md: 0 2px 4px rgba(0,0,0,0.3), 0 8px 24px rgba(0,0,0,0.35);
--shadow-primary: 0 1px 2px rgba(96,165,250,0.3), 0 2px 8px rgba(96,165,250,0.15);
--shadow-focus: 0 0 0 2px rgba(96,165,250,0.3);
```

### Border Radius
- Cards/panels/summary cards: `--radius: 12px` (up from 8px)
- Buttons/inputs: `--radius-sm: 8px` (up from 5px)
- Badges/pills: keep at 4px (or 12px for pill variant)

### Transitions
Add to `:root` (no dark mode override needed — timing is theme-independent):
```css
--transition-fast: 150ms ease;  /* new — hover states */
--transition: 200ms ease;       /* new — layout shifts, modals */
```
**Usage pattern:** Always specify the property name: `transition: background var(--transition-fast), color var(--transition-fast);` — never use as the entire `transition` value to avoid unintended `all` transitions.

### Spacing (4px grid snap)
All spacing tokens snapped to 4px multiples:
- `--spacing-container: 20px` (was 18px)
- `--spacing-section-gap: 12px` (already on grid)
- `--spacing-card-padding: 16px 16px` (was 14px 16px — +2px vertical, acceptable nudge)
- `--spacing-tab-padding-v: 8px` (was 6px)
- `--spacing-tab-padding-h: 8px` (already on grid)
- `--spacing-form-padding-v: 8px` (was 7px)
- `--spacing-form-padding-h: 12px` (was 10px)
- `--spacing-modal-padding: 24px 24px` (was 22px 24px — +2px vertical, acceptable nudge)
- `--spacing-table-cell-v: 8px` (already on grid)
- `--spacing-table-cell-h: 16px` (was 14px)
- `--spacing-table-th-v: 8px` (already on grid)
- `--spacing-table-th-h: 16px` (was 14px)
- `--spacing-table-compact-v: 8px` (was 6px)
- `--spacing-table-compact-h: 12px` (already on grid)
- `--spacing-table-dense-v: 4px` (was 5px)
- `--spacing-table-dense-h: 12px` (was 10px)
- `--spacing-detail-padding: 20px` (was 18px)

---

## Section 2: Typography Refinement

### Type Scale (unchanged fonts, refined usage)
```css
--font-xs: 10.5px;    /* unchanged — section labels */
--font-sm: 12px;       /* unchanged — field labels, table body */
--font-base: 13px;     /* unchanged — body text */
--font-md: 14px;       /* unchanged */
--font-lg: 17px;       /* unchanged */
--font-xl: 22px;       /* unchanged — summary values */
```

### Label Hierarchy System
- **Section headers** (e.g., sidebar nav groups, summary card labels): `font-size: 10.5px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; color: var(--text-muted);`
- **Field labels** (e.g., form labels): `font-size: 12px; font-weight: 500; color: var(--text-secondary);`
- **Helper text** (e.g., descriptions, hints): `font-size: 11px; font-weight: 400; color: var(--text-muted);`

### Financial Number Treatment
```css
.summary-value,
.journal-table td[data-col="amount"],
[class*="amount"],
[class*="balance"],
[class*="total"] {
    font-variant-numeric: tabular-nums;
    letter-spacing: -0.3px;
}
```

Large summary values (--font-xl) get `font-weight: 700; letter-spacing: -0.5px;` for tighter, bolder rendering.

---

## Section 3: Component Refinements

### Summary Cards
- Border: `1px solid rgba(0,0,0,0.06)` (light mode), `1px solid rgba(255,255,255,0.06)` (dark)
- Shadow: `var(--shadow-md)`
- Radius: `var(--radius)` (12px from token)
- Label: uppercase 10.5px/600wt system

### Buttons
- **Primary:** `background: var(--primary); box-shadow: var(--shadow-primary); font-weight: 600; border-radius: 8px;`
- **Primary hover:** Darken background + intensify shadow
- **Secondary:** `background: var(--surface); border: 1px solid rgba(0,0,0,0.08); box-shadow: var(--shadow-sm); font-weight: 500;`
- **Ghost/text buttons:** No shadow, just color change on hover

### Inputs & Selects
- Border: `1px solid rgba(0,0,0,0.08)`
- Shadow: `0 1px 2px rgba(0,0,0,0.04)`
- Radius: `8px`
- Focus: `box-shadow: var(--shadow-focus); border-color: var(--primary);`

### Table Rows
- Base: subtle `1px solid var(--border-light)` dividers between rows
- Hover state: `background: var(--surface-hover); border-left: 3px solid var(--primary);`
- Selected state: same as hover but persists
- Transition: `background 150ms ease, border-left 150ms ease`
- Amount cells: mono font, tabular-nums, right-aligned

---

## Section 4: Micro-Details

### 1. Smooth Transitions
Normalize all `transition` properties:
- Hover states (color, background, shadow, opacity): `var(--transition-fast)` (150ms)
- Layout shifts (sidebar, modals, tab panels): `var(--transition)` (200ms)
- Remove inconsistent timing values (0.1s, 0.12s, 0.15s, 0.3s) throughout
- **Exception:** Sidebar collapse/expand animation keeps its own `0.25s ease` timing (coordinates with `.app-main` margin-left transition) — do NOT normalize this to `var(--transition)`

### 2. Focus Ring System
```css
*:focus-visible {
    outline: none;
    box-shadow: var(--shadow-focus);
}
```
Buttons that already have shadow on normal state: combine shadows on focus.

### 3. Tabular Number Rendering
Apply `font-variant-numeric: tabular-nums` to all monetary value displays. See Section 2 for selectors.

### 4. Semantic Color Consistency
Audit all views and enforce:
- Green (`var(--green)` / `var(--green-bg)`): income, paid, receivable, positive values
- Red (`var(--red)` / `var(--red-bg)`): expense, overdue, payable, negative values
- Amber (`var(--amber)` / `var(--amber-bg)`): pending, warning, attention needed

### 5. 4px Spacing Grid
See Section 1 spacing token updates. Additionally, audit inline styles and component-specific overrides for off-grid values.

### 6. Refined Label Hierarchy
See Section 2 label system. Apply consistently to:
- Sidebar nav group labels (`.sidebar-nav-group`): update from 9.5px/1.1px tracking to 10.5px/0.5px tracking to match the system
- Summary card labels (`.summary-label`)
- Form field labels
- Table column headers
- Modal section headers

### 7. Polished Empty States
For each major tab (Journal, Budget, Cash Flow, P&L, Balance Sheet, etc.), when no data exists:

**Important:** The existing codebase already has `<p class="empty-state">` elements in each tab (simple centered text). Replace each existing `<p class="empty-state">` with the richer block:

```html
<div class="empty-state">
    <svg class="empty-state-icon"><!-- contextual icon --></svg>
    <h3 class="empty-state-title">No journal entries yet</h3>
    <p class="empty-state-desc">Add your first entry to start tracking.</p>
    <button class="btn btn-primary empty-state-cta">+ New Entry</button>
</div>
```

The `.empty-state` CSS class is redefined from a simple `<p>` style to a flex-column centered container. All existing instances in `index.html` must be updated from `<p>` to `<div>` with the new child structure.

Styling: centered, 48px icon in muted color, 16px title, 13px description, primary CTA button. Vertical spacing on 4px grid.

### 8. Custom Scrollbars
Apply thin scrollbars globally, but **preserve the sidebar's existing white-tinted scrollbar** (it needs light colors against the dark sidebar background):

```css
/* Global scrollbar — light mode (for main content, modals, tables) */
* {
    scrollbar-width: thin;
    scrollbar-color: rgba(0,0,0,0.15) transparent;
}
*::-webkit-scrollbar { width: 6px; height: 6px; }
*::-webkit-scrollbar-track { background: transparent; }
*::-webkit-scrollbar-thumb { background: rgba(0,0,0,0.15); border-radius: 4px; }
*::-webkit-scrollbar-thumb:hover { background: rgba(0,0,0,0.25); }
```

**Sidebar override** (keep existing white-tinted style — higher specificity):
```css
.main-tabs {
    scrollbar-color: rgba(255,255,255,0.1) transparent;
}
.main-tabs::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.1); }
```

Dark mode global: swap to `rgba(255,255,255,0.1)` / `rgba(255,255,255,0.2)`.

### 9. Skeleton Loading States
CSS-only skeleton with shimmer animation:
```css
.skeleton {
    background: linear-gradient(90deg, var(--border-light) 25%, var(--surface) 50%, var(--border-light) 75%);
    background-size: 200% 100%;
    animation: skeleton-shimmer 1.5s infinite;
    border-radius: var(--radius-sm);
}
@keyframes skeleton-shimmer {
    0% { background-position: 200% 0; }
    100% { background-position: -200% 0; }
}
```
Apply to tab content areas during tab switches. Skeleton shapes match content layout (summary cards = 3 rectangles, table = rows of bars).

### 10. Toast Notification Polish
Enhance the existing toast/notification system (including the existing `undo-toast` at bottom-center). Do not replace the undo toast — polish its styling to match the new design language:

- **New toasts** (success/error/warning/info): Slide in from top-right with `transform: translateX(120%)` → `translateX(0)`
- **Existing undo toast**: Keep bottom-center position, update shadow to `var(--shadow-md)` and border-radius to `var(--radius)`
- Shadow: `var(--shadow-md)` for all toast types
- Structure: icon (success/error/warning/info) + message text
- Auto-dismiss progress bar: thin colored line at bottom that shrinks over duration
- Exit animation: slide out + fade
- Stack multiple toasts with 8px gap

---

## Files Modified

- `css/styles.css` — all token updates, component styles, new utility classes, scrollbar/focus/skeleton/toast/empty-state styles
- `index.html` — empty state markup for each tab, skeleton placeholder containers
- `js/ui.js` — skeleton show/hide on tab switch, toast animation logic, empty state visibility toggle

---

## What This Does NOT Change

- No layout changes (sidebar width, grid structure)
- No new features or functionality
- No JavaScript business logic
- No data model changes
- No theme system changes (themes still override tokens as before)
