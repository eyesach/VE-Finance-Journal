---
name: product-tour
description: "Generate an interactive tutorial/walkthrough system with spotlight highlighting, smart tooltip positioning, keyboard navigation, and multiple tour support. Zero dependencies. Use when asked to: add a tour, create a walkthrough, build onboarding, add tutorial, create guided tour, build product tour, add feature tour, create step-by-step guide, build interactive tutorial, add spotlight walkthrough, create tooltip tour, build onboarding flow, add user onboarding, create help walkthrough, build intro tour, add app tour, create first-time user experience, add coaching marks."
metadata:
  author: user
  version: "1.0.0"
  argument-hint: <tour-config>
---

# Product Tour / Interactive Walkthrough Generator

Generate a zero-dependency interactive tutorial/walkthrough system with spotlight highlighting, smart tooltip positioning, keyboard navigation, progress tracking, and multiple tour support. Works in ANY web application.

---

## STEP 1: Gather Requirements (Ask One at a Time)

Before generating any code, ask the user these questions **one at a time**, waiting for each answer:

### Question 1 — Tour Definitions
Ask: "What tours do you want? For each tour, I need:
- **ID** (e.g., `main-tour`, `settings-walkthrough`)
- **Label** (human-readable, shown in picker)
- **Description** (one sentence, shown in picker)
- **Steps** — for each step: target CSS selector, title, body text, and preferred tooltip position (`top`, `bottom`, `left`, `right`)"

If the user is unsure, offer to inspect their HTML/app to suggest logical tour steps.

### Question 2 — Styling Integration
Ask: "Does your app define CSS custom properties I should inherit? For example `--surface`, `--border`, `--text`, `--primary`. Also:
- Do you have a dark mode? What selector activates it (e.g., `[data-theme='dark']`, `.dark`)?
- What z-index base should I use? (default: 5000)"

### Question 3 — Trigger Behavior
Ask: "How should the tour start?
- Auto-trigger on first visit? (uses localStorage to remember)
- A help/tour button? (give me its CSS selector, or I can create one)
- A localStorage key name for tracking? (default: `product-tour-seen`)"

### Question 4 — Tab/Route Switching
Ask: "Does your app use tabs, pages, or routes? If a tour step targets an element on a different tab/page, I need a JavaScript callback function that switches to that tab. For example: `switchTab('settings')`. What function should I call, and what argument format does it expect?"

If the app has no tabs/routes, skip this.

### Question 5 — Scroll Container
Ask: "Is the main scrollable area the window/document, or a specific element? If it's a specific element, give me its CSS selector (e.g., `.main-content`, `#app-body`)."

---

## STEP 2: Generate Files

Generate these files in the user's project:

### File 1: `js/product-tour.js` (~400 lines)
### File 2: `css/product-tour.css` (~250 lines)
### File 3: `js/tour-config.js` (optional — separate config)

Also add `<link>` and `<script>` tags to the user's HTML entry point.

---

## STEP 3: Full Implementation Reference

Below is the COMPLETE specification for every component. Follow this EXACTLY.

---

### 3.1 — CSS: product-tour.css

#### 3.1.1 — Custom Properties for Theming

```css
:root {
  --pt-surface: var(--surface, #ffffff);
  --pt-surface-alt: var(--surface-alt, #f9fafb);
  --pt-border: var(--border, #e5e7eb);
  --pt-text: var(--text, #1f2937);
  --pt-text-secondary: var(--text-secondary, #6b7280);
  --pt-primary: var(--primary, #3b82f6);
  --pt-primary-hover: var(--primary-hover, #2563eb);
  --pt-primary-text: var(--primary-text, #ffffff);
  --pt-overlay-opacity: 0.55;
  --pt-radius: 8px;
  --pt-z-base: 5000;
  --pt-font: inherit;
  --pt-tooltip-max-width: 360px;
  --pt-tooltip-min-width: 260px;
  --pt-transition-speed: 0.3s;
  --pt-arrow-size: 12px;
  --pt-spotlight-padding: 6px;
  --pt-spotlight-radius: 6px;
}
```

These cascade from the host app. If the host defines `--surface`, the tour inherits it automatically. Every `var()` has a hardcoded fallback so the tour works standalone.

#### 3.1.2 — Overlay (Click Blocker)

```css
#product-tour-overlay {
  display: none;
  position: fixed;
  top: 0;
  left: 0;
  width: 100%;
  height: 100%;
  z-index: var(--pt-z-base);
  /* No background — the spotlight box-shadow creates the dimming */
}

#product-tour-overlay.active {
  display: block;
}
```

The overlay is transparent. Its only job is to block pointer events on everything below the tour layer. The actual dimming is done by the spotlight's box-shadow (see 3.1.3).

#### 3.1.3 — Spotlight (THE Core Visual Trick)

```css
#product-tour-spotlight {
  display: none;
  position: fixed;
  box-shadow: 0 0 0 9999px rgba(0, 0, 0, var(--pt-overlay-opacity));
  border-radius: var(--pt-spotlight-radius);
  pointer-events: none;
  z-index: calc(var(--pt-z-base) + 1);
  transition: top var(--pt-transition-speed) ease,
              left var(--pt-transition-speed) ease,
              width var(--pt-transition-speed) ease,
              height var(--pt-transition-speed) ease,
              border-radius var(--pt-transition-speed) ease;
}

#product-tour-spotlight.active {
  display: block;
}
```

**How this works:** The `box-shadow` with a `9999px` spread radius creates a solid dark overlay that covers the ENTIRE viewport — except for the element itself, which becomes the "cutout" or "window" in the overlay. The element's own `border-radius` is copied to `--pt-spotlight-radius` so the cutout matches the target's shape. This is far simpler and more performant than SVG clip-path or canvas-based approaches.

#### 3.1.4 — Target Highlight

```css
.product-tour-target-highlight {
  position: relative !important;
  z-index: calc(var(--pt-z-base) + 1) !important;
  pointer-events: auto !important;
}
```

This class is added to the current step's target element to lift it above the overlay so the user can interact with it if needed. The `!important` declarations override any existing positioning.

#### 3.1.5 — Tooltip

```css
#product-tour-tooltip {
  display: none;
  position: fixed;
  z-index: calc(var(--pt-z-base) + 2);
  background: var(--pt-surface);
  border: 1px solid var(--pt-border);
  border-radius: var(--pt-radius);
  box-shadow: 0 10px 25px rgba(0, 0, 0, 0.15), 0 4px 10px rgba(0, 0, 0, 0.08);
  padding: 0;
  max-width: var(--pt-tooltip-max-width);
  min-width: var(--pt-tooltip-min-width);
  font-family: var(--pt-font);
  color: var(--pt-text);
  --arrow-offset: 50%;
}

#product-tour-tooltip.active {
  display: block;
}

.product-tour-tooltip-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 14px 16px 0 16px;
}

.product-tour-tooltip-title {
  font-size: 15px;
  font-weight: 600;
  color: var(--pt-text);
  margin: 0;
  line-height: 1.3;
}

.product-tour-tooltip-close {
  background: none;
  border: none;
  cursor: pointer;
  color: var(--pt-text-secondary);
  font-size: 18px;
  line-height: 1;
  padding: 2px 4px;
  border-radius: 4px;
  transition: background 0.15s, color 0.15s;
  flex-shrink: 0;
  margin-left: 8px;
}

.product-tour-tooltip-close:hover {
  background: var(--pt-surface-alt);
  color: var(--pt-text);
}

.product-tour-tooltip-body {
  padding: 10px 16px 14px 16px;
  font-size: 13.5px;
  line-height: 1.55;
  color: var(--pt-text-secondary);
}

.product-tour-tooltip-footer {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 0 16px 14px 16px;
  gap: 8px;
}

.product-tour-progress {
  font-size: 12px;
  color: var(--pt-text-secondary);
  white-space: nowrap;
}

.product-tour-progress-bar {
  display: flex;
  gap: 4px;
  align-items: center;
}

.product-tour-progress-dot {
  width: 6px;
  height: 6px;
  border-radius: 50%;
  background: var(--pt-border);
  transition: background 0.2s;
}

.product-tour-progress-dot.completed {
  background: var(--pt-primary);
}

.product-tour-progress-dot.current {
  background: var(--pt-primary);
  width: 8px;
  height: 8px;
}

.product-tour-nav {
  display: flex;
  gap: 6px;
  margin-left: auto;
}

.product-tour-btn {
  padding: 6px 14px;
  border-radius: 6px;
  font-size: 13px;
  font-weight: 500;
  cursor: pointer;
  border: 1px solid var(--pt-border);
  background: var(--pt-surface);
  color: var(--pt-text);
  transition: background 0.15s, border-color 0.15s;
  white-space: nowrap;
}

.product-tour-btn:hover {
  background: var(--pt-surface-alt);
}

.product-tour-btn-primary {
  background: var(--pt-primary);
  color: var(--pt-primary-text);
  border-color: var(--pt-primary);
}

.product-tour-btn-primary:hover {
  background: var(--pt-primary-hover);
  border-color: var(--pt-primary-hover);
}
```

#### 3.1.6 — Arrow (Rotated Square Technique)

The arrow is a 12x12 square rotated 45 degrees with selective border hiding to create a triangle illusion. The `--arrow-offset` CSS variable is computed by JavaScript so the arrow always points at the target center, even when the tooltip is clamped to the viewport edge.

```css
#product-tour-tooltip .product-tour-arrow {
  position: absolute;
  width: var(--pt-arrow-size);
  height: var(--pt-arrow-size);
  background: var(--pt-surface);
  transform: rotate(45deg);
  border: 1px solid var(--pt-border);
}

/* Arrow on TOP of tooltip (tooltip is BELOW target) */
#product-tour-tooltip[data-position="bottom"] .product-tour-arrow {
  top: calc(var(--pt-arrow-size) / -2);
  left: var(--arrow-offset);
  transform: translateX(-50%) rotate(45deg);
  border-bottom: none;
  border-right: none;
}

/* Arrow on BOTTOM of tooltip (tooltip is ABOVE target) */
#product-tour-tooltip[data-position="top"] .product-tour-arrow {
  bottom: calc(var(--pt-arrow-size) / -2);
  left: var(--arrow-offset);
  transform: translateX(-50%) rotate(45deg);
  border-top: none;
  border-left: none;
}

/* Arrow on LEFT of tooltip (tooltip is RIGHT of target) */
#product-tour-tooltip[data-position="right"] .product-tour-arrow {
  left: calc(var(--pt-arrow-size) / -2);
  top: var(--arrow-offset);
  transform: translateY(-50%) rotate(45deg);
  border-top: none;
  border-right: none;
}

/* Arrow on RIGHT of tooltip (tooltip is LEFT of target) */
#product-tour-tooltip[data-position="left"] .product-tour-arrow {
  right: calc(var(--pt-arrow-size) / -2);
  top: var(--arrow-offset);
  transform: translateY(-50%) rotate(45deg);
  border-bottom: none;
  border-left: none;
}
```

#### 3.1.7 — Tour Picker Popover

```css
#product-tour-picker-backdrop {
  display: none;
  position: fixed;
  top: 0;
  left: 0;
  width: 100%;
  height: 100%;
  background: rgba(0, 0, 0, 0.3);
  z-index: calc(var(--pt-z-base) + 10);
}

#product-tour-picker-backdrop.active {
  display: block;
}

#product-tour-picker {
  position: fixed;
  top: 50%;
  left: 50%;
  transform: translate(-50%, -50%);
  background: var(--pt-surface);
  border: 1px solid var(--pt-border);
  border-radius: calc(var(--pt-radius) + 4px);
  box-shadow: 0 20px 40px rgba(0, 0, 0, 0.18);
  padding: 24px;
  min-width: 320px;
  max-width: 440px;
  z-index: calc(var(--pt-z-base) + 11);
  font-family: var(--pt-font);
}

.product-tour-picker-title {
  font-size: 17px;
  font-weight: 600;
  color: var(--pt-text);
  margin: 0 0 6px 0;
}

.product-tour-picker-subtitle {
  font-size: 13px;
  color: var(--pt-text-secondary);
  margin: 0 0 18px 0;
}

.product-tour-picker-list {
  list-style: none;
  padding: 0;
  margin: 0;
  display: flex;
  flex-direction: column;
  gap: 8px;
}

.product-tour-picker-item {
  padding: 12px 14px;
  border: 1px solid var(--pt-border);
  border-radius: var(--pt-radius);
  cursor: pointer;
  transition: background 0.15s, border-color 0.15s;
}

.product-tour-picker-item:hover {
  background: var(--pt-surface-alt);
  border-color: var(--pt-primary);
}

.product-tour-picker-item-label {
  font-size: 14px;
  font-weight: 600;
  color: var(--pt-text);
  margin: 0 0 2px 0;
}

.product-tour-picker-item-desc {
  font-size: 12.5px;
  color: var(--pt-text-secondary);
  margin: 0;
}

.product-tour-picker-close {
  position: absolute;
  top: 16px;
  right: 16px;
  background: none;
  border: none;
  font-size: 20px;
  cursor: pointer;
  color: var(--pt-text-secondary);
  padding: 2px 6px;
  border-radius: 4px;
  line-height: 1;
}

.product-tour-picker-close:hover {
  background: var(--pt-surface-alt);
  color: var(--pt-text);
}
```

#### 3.1.8 — Welcome Dialog (First Visit)

```css
.product-tour-welcome {
  text-align: center;
  max-width: 380px;
}

.product-tour-welcome h2 {
  font-size: 20px;
  font-weight: 700;
  margin: 0 0 8px 0;
  color: var(--pt-text);
}

.product-tour-welcome p {
  font-size: 14px;
  color: var(--pt-text-secondary);
  margin: 0 0 20px 0;
  line-height: 1.5;
}

.product-tour-welcome-actions {
  display: flex;
  gap: 10px;
  justify-content: center;
}
```

#### 3.1.9 — Animations

```css
@keyframes product-tour-fade-in {
  from { opacity: 0; transform: translate(-50%, -50%) scale(0.96); }
  to   { opacity: 1; transform: translate(-50%, -50%) scale(1); }
}

#product-tour-picker {
  animation: product-tour-fade-in 0.2s ease;
}

#product-tour-tooltip.active {
  animation: product-tour-tooltip-in 0.2s ease;
}

@keyframes product-tour-tooltip-in {
  from { opacity: 0; }
  to   { opacity: 1; }
}
```

---

### 3.2 — JavaScript: product-tour.js

#### 3.2.1 — Z-Index Layering Model

| Layer | z-index | Element ID / Class | Purpose |
|-------|---------|-------------------|---------|
| Overlay | base + 0 (5000) | `#product-tour-overlay` | Transparent click blocker over entire page |
| Spotlight | base + 1 (5001) | `#product-tour-spotlight` | Box-shadow cutout creating the dim effect |
| Target | base + 1 (5001) | `.product-tour-target-highlight` | Lifts target above overlay so it's visible/clickable |
| Tooltip | base + 2 (5002) | `#product-tour-tooltip` | The info panel with title, body, nav |
| Picker Backdrop | base + 10 (5010) | `#product-tour-picker-backdrop` | Dim behind the picker |
| Picker | base + 11 (5011) | `#product-tour-picker` | Tour selector popover |

#### 3.2.2 — Complete ProductTour Object

Generate the `ProductTour` object as an IIFE-based singleton. Here is the FULL structure:

```js
const ProductTour = (function () {
  'use strict';

  // --- Private State ---
  let config = {
    tours: {},
    switchTab: null,
    scrollContainer: null,
    welcomeKey: 'product-tour-seen',
    triggerSelector: null,
    zIndexBase: 5000,
    spotlightPadding: 6,
    viewportMargin: 8,
    arrowSize: 12,
    debounceMs: 100,
    settleMs: 80
  };

  let state = {
    active: false,
    currentTourId: null,
    currentStepIndex: 0,
    currentTour: null,
    previousTarget: null,
    resizeHandler: null
  };

  // --- DOM Element References (created once) ---
  let els = {};

  // ===========================================
  // DOM CREATION
  // ===========================================

  function createDOM() {
    if (document.getElementById('product-tour-overlay')) return;

    const overlay = document.createElement('div');
    overlay.id = 'product-tour-overlay';
    overlay.addEventListener('click', function (e) {
      e.stopPropagation();
    });

    const spotlight = document.createElement('div');
    spotlight.id = 'product-tour-spotlight';

    const tooltip = document.createElement('div');
    tooltip.id = 'product-tour-tooltip';
    tooltip.innerHTML = `
      <div class="product-tour-arrow"></div>
      <div class="product-tour-tooltip-header">
        <h4 class="product-tour-tooltip-title"></h4>
        <button class="product-tour-tooltip-close" aria-label="End tour">&times;</button>
      </div>
      <div class="product-tour-tooltip-body"></div>
      <div class="product-tour-tooltip-footer">
        <div class="product-tour-progress">
          <div class="product-tour-progress-bar"></div>
        </div>
        <div class="product-tour-nav">
          <button class="product-tour-btn product-tour-btn-prev">Back</button>
          <button class="product-tour-btn product-tour-btn-primary product-tour-btn-next">Next</button>
        </div>
      </div>
    `;

    const pickerBackdrop = document.createElement('div');
    pickerBackdrop.id = 'product-tour-picker-backdrop';

    const picker = document.createElement('div');
    picker.id = 'product-tour-picker';

    document.body.appendChild(overlay);
    document.body.appendChild(spotlight);
    document.body.appendChild(tooltip);
    document.body.appendChild(pickerBackdrop);
    document.body.appendChild(picker);

    els = {
      overlay,
      spotlight,
      tooltip,
      picker,
      pickerBackdrop,
      title: tooltip.querySelector('.product-tour-tooltip-title'),
      body: tooltip.querySelector('.product-tour-tooltip-body'),
      close: tooltip.querySelector('.product-tour-tooltip-close'),
      prevBtn: tooltip.querySelector('.product-tour-btn-prev'),
      nextBtn: tooltip.querySelector('.product-tour-btn-next'),
      progress: tooltip.querySelector('.product-tour-progress-bar'),
      arrow: tooltip.querySelector('.product-tour-arrow')
    };

    // Wire up tooltip buttons
    els.close.addEventListener('click', end);
    els.prevBtn.addEventListener('click', prevStep);
    els.nextBtn.addEventListener('click', nextStep);
    pickerBackdrop.addEventListener('click', hidePicker);
  }

  // ===========================================
  // TOOLTIP POSITIONING ALGORITHM
  // ===========================================
  //
  // This is the CRITICAL positioning logic.
  //
  // Step 1: Measure available space in all 4 directions from the target rect.
  // Step 2: Build a priority list: [preferred, bottom, right, top, left] deduplicated.
  // Step 3: For each direction, check if tooltip fits (width x height + arrow + margins).
  // Step 4: Pick the first direction that fits. If none fit, use the one with the most space.
  // Step 5: Center the tooltip on the perpendicular axis relative to the target.
  // Step 6: Clamp the tooltip position to the viewport with an 8px margin.
  // Step 7: Compute --arrow-offset so the arrow points at the target's center,
  //         even when the tooltip has been clamped away from center.

  function positionTooltip(targetRect, preferred) {
    const tt = els.tooltip;
    const arrow = els.arrow;

    // Ensure tooltip is visible but off-screen for measurement
    tt.style.visibility = 'hidden';
    tt.classList.add('active');
    tt.style.left = '0px';
    tt.style.top = '0px';

    const ttRect = tt.getBoundingClientRect();
    const ttW = ttRect.width;
    const ttH = ttRect.height;
    const margin = config.viewportMargin;
    const arrowGap = config.arrowSize / 2 + 6; // half arrow + gap
    const vw = window.innerWidth;
    const vh = window.innerHeight;

    // Step 1: Available space in each direction
    const spaceTop = targetRect.top;
    const spaceBottom = vh - targetRect.bottom;
    const spaceLeft = targetRect.left;
    const spaceRight = vw - targetRect.right;

    // Step 2: Priority order (deduplicated)
    const order = [preferred, 'bottom', 'right', 'top', 'left'];
    const seen = new Set();
    const priorities = [];
    for (const dir of order) {
      if (dir && !seen.has(dir)) {
        seen.add(dir);
        priorities.push(dir);
      }
    }

    // Step 3: Check fit
    function fits(dir) {
      switch (dir) {
        case 'bottom': return spaceBottom >= ttH + arrowGap + margin;
        case 'top':    return spaceTop >= ttH + arrowGap + margin;
        case 'left':   return spaceLeft >= ttW + arrowGap + margin;
        case 'right':  return spaceRight >= ttW + arrowGap + margin;
      }
      return false;
    }

    let chosen = priorities.find(fits);

    // Fallback: direction with most space
    if (!chosen) {
      const spaces = { top: spaceTop, bottom: spaceBottom, left: spaceLeft, right: spaceRight };
      chosen = Object.keys(spaces).reduce((a, b) => spaces[a] > spaces[b] ? a : b);
    }

    // Step 4 & 5: Compute position
    let top, left;
    const targetCenterX = targetRect.left + targetRect.width / 2;
    const targetCenterY = targetRect.top + targetRect.height / 2;

    switch (chosen) {
      case 'bottom':
        top = targetRect.bottom + arrowGap;
        left = targetCenterX - ttW / 2;
        break;
      case 'top':
        top = targetRect.top - ttH - arrowGap;
        left = targetCenterX - ttW / 2;
        break;
      case 'right':
        top = targetCenterY - ttH / 2;
        left = targetRect.right + arrowGap;
        break;
      case 'left':
        top = targetCenterY - ttH / 2;
        left = targetRect.left - ttW - arrowGap;
        break;
    }

    // Step 6: Clamp to viewport
    const clampedLeft = Math.max(margin, Math.min(left, vw - ttW - margin));
    const clampedTop = Math.max(margin, Math.min(top, vh - ttH - margin));

    // Step 7: Compute arrow offset
    let arrowOffset;
    if (chosen === 'top' || chosen === 'bottom') {
      // Arrow is on horizontal axis — offset from left edge of tooltip
      const idealArrowX = targetCenterX - clampedLeft;
      arrowOffset = Math.max(16, Math.min(idealArrowX, ttW - 16));
      arrowOffset = arrowOffset + 'px';
    } else {
      // Arrow is on vertical axis — offset from top edge of tooltip
      const idealArrowY = targetCenterY - clampedTop;
      arrowOffset = Math.max(16, Math.min(idealArrowY, ttH - 16));
      arrowOffset = arrowOffset + 'px';
    }

    // Apply
    tt.setAttribute('data-position', chosen);
    tt.style.left = clampedLeft + 'px';
    tt.style.top = clampedTop + 'px';
    tt.style.setProperty('--arrow-offset', arrowOffset);
    tt.style.visibility = '';
  }

  // ===========================================
  // SPOTLIGHT POSITIONING
  // ===========================================

  function positionSpotlight(targetEl) {
    const rect = targetEl.getBoundingClientRect();
    const pad = config.spotlightPadding;
    const computed = window.getComputedStyle(targetEl);
    const borderRadius = computed.borderRadius || '0px';

    els.spotlight.style.top = (rect.top - pad) + 'px';
    els.spotlight.style.left = (rect.left - pad) + 'px';
    els.spotlight.style.width = (rect.width + pad * 2) + 'px';
    els.spotlight.style.height = (rect.height + pad * 2) + 'px';
    els.spotlight.style.borderRadius = borderRadius;
  }

  // ===========================================
  // STEP LIFECYCLE
  // ===========================================
  //
  // 1. If step has `tab` property -> call switchTab callback
  // 2. Run prepare() callback if defined on the step
  // 3. Wait settleMs (80ms) for DOM to settle (tab switch, animations)
  // 4. Find target element; if not found or hidden -> auto-advance
  // 5. Scroll target into view if needed
  // 6. Position spotlight around target
  // 7. Position tooltip relative to target
  // 8. Update progress bar and button labels
  //
  // On leave (before moving to next/prev step):
  // - Run cleanup() callback if defined on the current step
  // - Remove highlight class from previous target

  function showStep(index, repositionOnly) {
    const tour = state.currentTour;
    if (!tour || index < 0 || index >= tour.steps.length) {
      end();
      return;
    }

    const step = tour.steps[index];
    const prevStep = state.currentStepIndex !== index ? tour.steps[state.currentStepIndex] : null;

    // Cleanup previous step (skip on reposition)
    if (!repositionOnly) {
      if (state.previousTarget) {
        state.previousTarget.classList.remove('product-tour-target-highlight');
        state.previousTarget = null;
      }
      if (prevStep && prevStep.cleanup && state.currentStepIndex !== index) {
        try { prevStep.cleanup(); } catch (e) { console.warn('Tour cleanup error:', e); }
      }

      state.currentStepIndex = index;

      // Tab switching
      if (step.tab && config.switchTab) {
        try { config.switchTab(step.tab); } catch (e) { console.warn('Tour switchTab error:', e); }
      }

      // Prepare callback
      if (step.prepare) {
        try { step.prepare(); } catch (e) { console.warn('Tour prepare error:', e); }
      }
    }

    // Wait for DOM settle, then position
    setTimeout(function () {
      const target = document.querySelector(step.selector);

      // Target not found or hidden -> auto-advance
      if (!target || target.offsetParent === null) {
        console.warn('Tour: target not found for selector "' + step.selector + '", auto-advancing.');
        if (index < tour.steps.length - 1) {
          showStep(index + 1);
        } else {
          end();
        }
        return;
      }

      // Scroll into view
      const container = config.scrollContainer
        ? document.querySelector(config.scrollContainer)
        : null;
      target.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'nearest' });

      // Brief delay after scroll for layout
      setTimeout(function () {
        // Highlight target
        target.classList.add('product-tour-target-highlight');
        state.previousTarget = target;

        // Position spotlight
        positionSpotlight(target);
        els.spotlight.classList.add('active');

        // Position tooltip
        const targetRect = target.getBoundingClientRect();
        els.title.textContent = step.title || '';
        els.body.innerHTML = step.text || '';
        positionTooltip(targetRect, step.position || 'bottom');

        // Update progress dots
        updateProgress(index, tour.steps.length);

        // Update button labels
        els.prevBtn.style.display = index === 0 ? 'none' : '';
        els.nextBtn.textContent = index === tour.steps.length - 1 ? 'Finish' : 'Next';

        // Show tooltip
        els.tooltip.classList.add('active');
      }, repositionOnly ? 0 : 40);
    }, repositionOnly ? 0 : config.settleMs);
  }

  function updateProgress(current, total) {
    let html = '';
    for (let i = 0; i < total; i++) {
      let cls = 'product-tour-progress-dot';
      if (i < current) cls += ' completed';
      if (i === current) cls += ' current';
      html += '<span class="' + cls + '"></span>';
    }
    els.progress.innerHTML = html;
  }

  // ===========================================
  // NAVIGATION
  // ===========================================

  function nextStep() {
    if (!state.active) return;
    const tour = state.currentTour;
    if (state.currentStepIndex < tour.steps.length - 1) {
      showStep(state.currentStepIndex + 1);
    } else {
      end();
    }
  }

  function prevStep() {
    if (!state.active || state.currentStepIndex <= 0) return;
    showStep(state.currentStepIndex - 1);
  }

  // ===========================================
  // KEYBOARD HANDLER
  // ===========================================

  function onKeyDown(e) {
    if (!state.active) return;
    switch (e.key) {
      case 'ArrowRight':
      case 'Enter':
        e.preventDefault();
        nextStep();
        break;
      case 'ArrowLeft':
        e.preventDefault();
        prevStep();
        break;
      case 'Escape':
        e.preventDefault();
        end();
        break;
    }
  }

  // ===========================================
  // RESIZE / SCROLL REPOSITION
  // ===========================================

  function setupResizeHandler() {
    let timeout;
    state.resizeHandler = function () {
      clearTimeout(timeout);
      timeout = setTimeout(function () {
        if (state.active) {
          showStep(state.currentStepIndex, true);
        }
      }, config.debounceMs);
    };

    window.addEventListener('resize', state.resizeHandler);
    window.addEventListener('scroll', state.resizeHandler, true); // true = capture for nested scrolls
  }

  function teardownResizeHandler() {
    if (state.resizeHandler) {
      window.removeEventListener('resize', state.resizeHandler);
      window.removeEventListener('scroll', state.resizeHandler, true);
      state.resizeHandler = null;
    }
  }

  // ===========================================
  // TOUR PICKER
  // ===========================================

  function showPicker() {
    createDOM();

    const tourIds = Object.keys(config.tours);
    if (tourIds.length === 0) return;

    // If only one tour, start it directly
    if (tourIds.length === 1) {
      start(tourIds[0]);
      return;
    }

    let html = `
      <button class="product-tour-picker-close" aria-label="Close">&times;</button>
      <h3 class="product-tour-picker-title">Choose a Tour</h3>
      <p class="product-tour-picker-subtitle">Select a guided walkthrough to get started.</p>
      <ul class="product-tour-picker-list">
    `;

    for (const id of tourIds) {
      const tour = config.tours[id];
      html += `
        <li class="product-tour-picker-item" data-tour-id="${id}">
          <div class="product-tour-picker-item-label">${tour.label || id}</div>
          <div class="product-tour-picker-item-desc">${tour.description || ''}</div>
        </li>
      `;
    }

    html += '</ul>';
    els.picker.innerHTML = html;

    // Wire events
    els.picker.querySelector('.product-tour-picker-close').addEventListener('click', hidePicker);
    els.picker.querySelectorAll('.product-tour-picker-item').forEach(function (item) {
      item.addEventListener('click', function () {
        const tourId = this.getAttribute('data-tour-id');
        hidePicker();
        start(tourId);
      });
    });

    els.pickerBackdrop.classList.add('active');
    els.picker.style.display = 'block';
  }

  function hidePicker() {
    if (els.pickerBackdrop) els.pickerBackdrop.classList.remove('active');
    if (els.picker) els.picker.style.display = 'none';
  }

  // ===========================================
  // WELCOME DIALOG (First Visit)
  // ===========================================

  function showWelcome(defaultTourId) {
    createDOM();

    const html = `
      <div class="product-tour-welcome">
        <h2>Welcome!</h2>
        <p>It looks like this is your first time here. Would you like a quick guided tour?</p>
        <div class="product-tour-welcome-actions">
          <button class="product-tour-btn product-tour-welcome-skip">Skip</button>
          <button class="product-tour-btn product-tour-btn-primary product-tour-welcome-start">Start Tour</button>
        </div>
      </div>
    `;

    els.picker.innerHTML = html;

    els.picker.querySelector('.product-tour-welcome-skip').addEventListener('click', function () {
      localStorage.setItem(config.welcomeKey, 'true');
      hidePicker();
    });

    els.picker.querySelector('.product-tour-welcome-start').addEventListener('click', function () {
      localStorage.setItem(config.welcomeKey, 'true');
      hidePicker();
      if (defaultTourId) {
        start(defaultTourId);
      } else {
        showPicker();
      }
    });

    els.pickerBackdrop.classList.add('active');
    els.picker.style.display = 'block';
  }

  // ===========================================
  // PUBLIC API
  // ===========================================

  function init(options) {
    Object.assign(config, options);

    createDOM();

    // Apply z-index base as CSS variable
    document.documentElement.style.setProperty('--pt-z-base', config.zIndexBase);

    // Keyboard
    document.addEventListener('keydown', onKeyDown);

    // Trigger button
    if (config.triggerSelector) {
      const trigger = document.querySelector(config.triggerSelector);
      if (trigger) {
        trigger.addEventListener('click', function (e) {
          e.preventDefault();
          showPicker();
        });
      }
    }

    // First-visit welcome
    if (config.welcomeKey && !localStorage.getItem(config.welcomeKey)) {
      const tourIds = Object.keys(config.tours);
      const defaultTour = tourIds.length === 1 ? tourIds[0] : null;
      // Slight delay so the page finishes rendering
      setTimeout(function () {
        showWelcome(defaultTour);
      }, 500);
    }
  }

  function start(tourId) {
    // End any current tour first (handles concurrent start)
    if (state.active) {
      end(true); // silent = true, no cleanup animation
    }

    const tour = config.tours[tourId];
    if (!tour || !tour.steps || tour.steps.length === 0) {
      console.warn('ProductTour: tour "' + tourId + '" not found or has no steps.');
      return;
    }

    state.active = true;
    state.currentTourId = tourId;
    state.currentTour = tour;
    state.currentStepIndex = 0;

    els.overlay.classList.add('active');
    setupResizeHandler();
    showStep(0);
  }

  function end(silent) {
    if (!state.active && !silent) return;

    // Cleanup current step
    if (state.currentTour && state.currentTour.steps[state.currentStepIndex]) {
      const step = state.currentTour.steps[state.currentStepIndex];
      if (step.cleanup) {
        try { step.cleanup(); } catch (e) { console.warn('Tour cleanup error:', e); }
      }
    }

    // Remove target highlight
    if (state.previousTarget) {
      state.previousTarget.classList.remove('product-tour-target-highlight');
      state.previousTarget = null;
    }

    // Hide everything
    els.overlay.classList.remove('active');
    els.spotlight.classList.remove('active');
    els.tooltip.classList.remove('active');

    teardownResizeHandler();

    state.active = false;
    state.currentTourId = null;
    state.currentTour = null;
    state.currentStepIndex = 0;
  }

  return {
    init: init,
    start: start,
    end: end,
    showPicker: showPicker,
    hidePicker: hidePicker
  };

})();
```

#### 3.2.3 — Tour Config File (tour-config.js)

Generate a separate config file with the tour definitions collected from the user. Structure:

```js
// Tour definitions — edit steps here
document.addEventListener('DOMContentLoaded', function () {
  ProductTour.init({
    tours: {
      'main-tour': {
        label: 'Getting Started',
        description: 'Learn the basics of the application.',
        steps: [
          {
            selector: '#some-element',
            title: 'Step Title',
            text: 'Description of what this element does.',
            position: 'bottom'
            // Optional:
            // tab: 'tab-id',        // triggers config.switchTab('tab-id')
            // prepare: function() {}, // runs before showing this step
            // cleanup: function() {}  // runs when leaving this step
          }
          // ... more steps
        ]
      }
      // ... more tours
    },
    switchTab: null,            // function(tabId) { ... } — or null if no tabs
    scrollContainer: null,      // CSS selector like '.main-content' or null for window
    welcomeKey: 'product-tour-seen',
    triggerSelector: '#help-tour-btn', // button that opens the picker
    zIndexBase: 5000
  });
});
```

---

## STEP 4: Integration Checklist

After generating the files, instruct the user to add these tags to their HTML (or add them automatically if you have access):

```html
<!-- Product Tour CSS -->
<link rel="stylesheet" href="css/product-tour.css">

<!-- Product Tour JS (before closing </body>) -->
<script src="js/product-tour.js"></script>
<script src="js/tour-config.js"></script>
```

---

## STEP 5: Edge Cases to Handle

Implement ALL of the following in the generated code:

1. **Target not found or hidden** (`offsetParent === null`): Auto-advance to the next step. If it is the last step, end the tour. Log a warning to console.

2. **Target inside a scrollable container**: Call `scrollIntoView({ behavior: 'smooth', block: 'nearest' })` and wait 80ms for scroll to settle before positioning.

3. **Responsive / window resize**: Debounced handler (100ms) on both `resize` and `scroll` (capture phase) events that re-calls `showStep(currentIndex, repositionOnly=true)`. The `repositionOnly` flag skips prepare/cleanup callbacks and tab switching.

4. **Concurrent `start()` calls**: If a tour is already active, `end()` it silently first, then start the new one.

5. **Single tour shortcut**: If only one tour is registered, `showPicker()` starts it directly instead of showing the picker UI.

6. **Progress indicator**: Dot-based progress bar showing completed dots (filled), current dot (filled + larger), and future dots (empty outline).

7. **Keyboard navigation**: `ArrowRight` / `Enter` = next, `ArrowLeft` = previous, `Escape` = end tour. Only active when `state.active` is true.

8. **Arrow always points at target**: The `--arrow-offset` CSS variable is recomputed every time the tooltip is positioned, accounting for viewport clamping. Arrow is clamped to stay at least 16px from tooltip edges.

9. **Clean teardown**: `end()` removes all classes, hides all elements, removes event listeners, and runs the final step's cleanup callback.

10. **No z-index conflicts**: All z-index values are computed from a single configurable base, using CSS `calc()` for the CSS layers and JS arithmetic for runtime positioning.

---

## STEP 6: Testing Guidance

After generating the code, tell the user:

1. Open the app in a browser
2. If `welcomeKey` is set and localStorage is clean, a welcome dialog should appear
3. Click "Start Tour" — verify spotlight highlights the first target
4. Press ArrowRight to advance, ArrowLeft to go back, Escape to exit
5. Resize the window — tooltip and spotlight should reposition smoothly
6. Open DevTools and delete localStorage entry to re-test welcome dialog
7. If any step's target is missing, verify it auto-advances with a console warning
