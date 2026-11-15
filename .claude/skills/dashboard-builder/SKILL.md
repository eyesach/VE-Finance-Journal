---
name: dashboard-builder
description: "Generate a complete interactive dashboard from 3 key metrics using Chart.js. Includes summary cards with trends, time series, distribution/comparison charts, filters, data table, dark mode, and auto-refresh. Works for ANY domain. Use when asked to: build/create/make a dashboard, create analytics dashboard, build metrics/KPI dashboard, create data dashboard, build monitoring dashboard, add Chart.js charts, create data visualization, build interactive charts, add summary cards with charts, create statistics dashboard, build real-time dashboard, create sales/performance/tracking dashboard, visualize data, build chart-based UI, create reporting page, charts and graphs, summary cards with charts, Chart.js dashboard, monitoring dashboard, KPI dashboard, real-time dashboard, performance dashboard, tracking dashboard."
metadata:
  author: user
  version: "1.0.0"
  argument-hint: <metric1, metric2, metric3>
---

# Dashboard Builder Skill

## Core Concept

The user provides **3 key metrics** from ANY domain. This skill builds a complete, polished, interactive dashboard around those 3 metrics. The skill adapts chart types, axis formatting, trend indicators, color coding, and mock data generation based on the metric types and domain context.

**This is domain-agnostic.** E-commerce, weather, SaaS, IoT, healthcare, sports, manufacturing, logistics — the same architecture works for all of them.

---

## Step 1: Gather All Requirements Upfront

Ask ALL of the following questions in a SINGLE message. Dashboards need all info upfront — do NOT ask one question at a time.

Present this as a numbered list the user can quickly fill in:

```
To build your dashboard, I need a few details:

1. **3 Key Metrics** — For each, provide:
   - Name (e.g., "Revenue", "Temperature", "Active Users")
   - Unit type: number | currency | percentage | duration
   - Is "lower better"? (e.g., churn rate, error rate, response time) — yes/no

2. **Data Source** — Pick one:
   - **mock** — I'll generate realistic sample data (default)
   - **api** — Provide endpoint URL + response shape
   - **static** — Provide a JSON array

3. **Categories/Segments** — How should data be segmented?
   (e.g., "North, South, East, West" or "Web, Mobile, API" or "ICU, ER, Outpatient")
   Default: 4 generic segments

4. **Time Granularity** — hourly | daily | weekly | monthly
   Default: daily

5. **Threshold/Target Lines** — Optional lines on the time series chart
   (e.g., "Revenue target: $50,000", "Max error rate: 2%", "Min temperature: 32")
```

If the user provides only the 3 metrics and nothing else, use these defaults:
- Data source: mock
- Segments: 4 contextually named segments based on the domain
- Granularity: daily
- Thresholds: none

---

## Step 2: Generate Files

Create exactly these 4 files:

```
index.html          — Single page, CDN imports, full layout
css/dashboard.css   — CSS custom properties, light/dark, responsive, all components
js/dashboard.js     — Dashboard controller: init, charts, filters, table, export
js/data.js          — DataSource: mock generator OR API adapter OR static loader
```

---

## Step 3: File Contents — Full Specifications

### 3A. index.html

Use this exact structure. Replace `{{TITLE}}`, `{{METRIC_1}}`, `{{METRIC_2}}`, `{{METRIC_3}}`, and `{{SEGMENTS}}` with user-provided values.

```html
<!DOCTYPE html>
<html lang="en" data-theme="light">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{{TITLE}} Dashboard</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;700&display=swap" rel="stylesheet">
  <link rel="stylesheet" href="css/dashboard.css">
</head>
<body>

  <!-- HEADER BAR -->
  <header class="header">
    <h1 class="header__title">{{TITLE}} Dashboard</h1>
    <div class="header__controls">
      <span class="header__updated" id="lastUpdated">Updated just now</span>
      <select id="autoRefresh" class="header__refresh" title="Auto-refresh interval">
        <option value="0">Auto-refresh: Off</option>
        <option value="5">Every 5s</option>
        <option value="15">Every 15s</option>
        <option value="30">Every 30s</option>
        <option value="60">Every 60s</option>
      </select>
      <button id="themeToggle" class="header__theme-btn" title="Toggle dark mode" aria-label="Toggle dark mode">
        <span class="icon-sun">&#9728;</span>
        <span class="icon-moon">&#9790;</span>
      </button>
    </div>
  </header>

  <!-- FILTER CONTROLS -->
  <section class="filters">
    <div class="filters__dates">
      <label>From <input type="date" id="dateFrom" class="filters__input"></label>
      <label>To <input type="date" id="dateTo" class="filters__input"></label>
      <div class="filters__presets">
        <button class="preset-btn" data-preset="today">Today</button>
        <button class="preset-btn" data-preset="7d">7d</button>
        <button class="preset-btn active" data-preset="30d">30d</button>
        <button class="preset-btn" data-preset="90d">90d</button>
        <button class="preset-btn" data-preset="ytd">YTD</button>
        <button class="preset-btn" data-preset="all">All</button>
      </div>
    </div>
    <div class="filters__segment">
      <label>Segment
        <select id="segmentFilter" class="filters__select">
          <option value="all">All Segments</option>
          <!-- Options populated by JS from segments array -->
        </select>
      </label>
    </div>
    <div class="filters__granularity">
      <span class="filters__label">Granularity</span>
      <label class="radio-label"><input type="radio" name="granularity" value="daily" checked> Daily</label>
      <label class="radio-label"><input type="radio" name="granularity" value="weekly"> Weekly</label>
      <label class="radio-label"><input type="radio" name="granularity" value="monthly"> Monthly</label>
    </div>
  </section>

  <!-- SUMMARY CARDS -->
  <section class="cards-row" id="summaryCards">
    <div class="card" data-metric="0">
      <span class="card__label">{{METRIC_1}}</span>
      <span class="card__value" id="cardValue0">—</span>
      <span class="card__trend" id="cardTrend0"></span>
      <canvas class="card__sparkline" id="spark0" width="120" height="40"></canvas>
    </div>
    <div class="card" data-metric="1">
      <span class="card__label">{{METRIC_2}}</span>
      <span class="card__value" id="cardValue1">—</span>
      <span class="card__trend" id="cardTrend1"></span>
      <canvas class="card__sparkline" id="spark1" width="120" height="40"></canvas>
    </div>
    <div class="card" data-metric="2">
      <span class="card__label">{{METRIC_3}}</span>
      <span class="card__value" id="cardValue2">—</span>
      <span class="card__trend" id="cardTrend2"></span>
      <canvas class="card__sparkline" id="spark2" width="120" height="40"></canvas>
    </div>
  </section>

  <!-- TIME SERIES CHART -->
  <section class="chart-full">
    <div class="chart-container">
      <h2 class="chart-title">Trends Over Time</h2>
      <canvas id="timeSeriesChart"></canvas>
    </div>
  </section>

  <!-- DISTRIBUTION + COMPARISON ROW -->
  <section class="charts-row" id="chartsRow">
    <div class="chart-container">
      <h2 class="chart-title">{{METRIC_1}} by Segment</h2>
      <canvas id="distributionChart"></canvas>
    </div>
    <div class="chart-container">
      <h2 class="chart-title">Metrics Comparison</h2>
      <canvas id="comparisonChart"></canvas>
    </div>
  </section>

  <!-- DATA TABLE -->
  <section class="table-section">
    <div class="table-header">
      <h2 class="chart-title">Detailed Data</h2>
      <button id="csvExport" class="btn btn--export">&#11123; Export CSV</button>
    </div>
    <div class="table-wrap">
      <table class="data-table" id="dataTable">
        <thead>
          <tr>
            <th data-sort="date" class="sortable">Date &#8597;</th>
            <th data-sort="segment" class="sortable">Segment &#8597;</th>
            <th data-sort="m0" class="sortable">{{METRIC_1}} &#8597;</th>
            <th data-sort="m1" class="sortable">{{METRIC_2}} &#8597;</th>
            <th data-sort="m2" class="sortable">{{METRIC_3}} &#8597;</th>
          </tr>
        </thead>
        <tbody id="tableBody"></tbody>
      </table>
    </div>
    <div class="table-pagination" id="pagination"></div>
  </section>

  <!-- EMPTY STATE (hidden by default) -->
  <section class="empty-state" id="emptyState" style="display:none;">
    <div class="empty-state__box">
      <p>No data available for the selected filters.</p>
      <button class="btn" onclick="Dashboard.loadSampleData()">Load Sample Data</button>
    </div>
  </section>

  <!-- CDN SCRIPTS -->
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-annotation@3.0.1/dist/chartjs-plugin-annotation.min.js"></script>

  <!-- APP SCRIPTS -->
  <script src="js/data.js"></script>
  <script src="js/dashboard.js"></script>
</body>
</html>
```

**Important HTML notes:**
- The `data-theme` attribute on `<html>` drives light/dark mode via CSS.
- All chart canvases must have unique IDs matching what `dashboard.js` references.
- Segment `<option>` elements are populated dynamically from the segments array in `js/data.js`.
- If there is only ONE segment, remove the distribution chart `<div>` entirely and give the comparison chart container the class `chart-full` instead of being inside `charts-row`.

---

### 3B. css/dashboard.css

```css
/* ============================================================
   CSS CUSTOM PROPERTIES (DESIGN TOKENS)
   ============================================================ */
:root {
  /* Colors — Light */
  --bg: #f0f2f5;
  --surface: #ffffff;
  --surface-hover: #f8f9fa;
  --text: #1a1a2e;
  --text-muted: #6c757d;
  --border: #e0e0e0;
  --primary: #4361ee;
  --primary-light: rgba(67, 97, 238, 0.1);
  --success: #2ec4b6;
  --danger: #e63946;
  --warning: #f4a261;
  --chart-1: #4361ee;
  --chart-2: #2ec4b6;
  --chart-3: #f4a261;
  --chart-4: #e63946;
  --chart-5: #7209b7;
  --chart-6: #3a86a7;
  --radius: 12px;
  --shadow: 0 1px 3px rgba(0,0,0,0.08), 0 1px 2px rgba(0,0,0,0.06);
  --shadow-lg: 0 4px 12px rgba(0,0,0,0.1);
  --font: 'DM Sans', system-ui, -apple-system, sans-serif;
  --transition: 0.2s ease;
}

[data-theme="dark"] {
  --bg: #0f0f1a;
  --surface: #1a1a2e;
  --surface-hover: #22223a;
  --text: #e8e8ed;
  --text-muted: #9a9ab0;
  --border: #2a2a40;
  --primary: #6381ff;
  --primary-light: rgba(99, 129, 255, 0.15);
  --shadow: 0 1px 3px rgba(0,0,0,0.3), 0 1px 2px rgba(0,0,0,0.2);
  --shadow-lg: 0 4px 12px rgba(0,0,0,0.4);
}

/* ============================================================
   RESET & BASE
   ============================================================ */
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

body {
  font-family: var(--font);
  background: var(--bg);
  color: var(--text);
  line-height: 1.5;
  min-height: 100vh;
  padding: 0 24px 48px;
  transition: background var(--transition), color var(--transition);
}

/* ============================================================
   HEADER
   ============================================================ */
.header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 20px 0;
  border-bottom: 1px solid var(--border);
  margin-bottom: 24px;
  flex-wrap: wrap;
  gap: 12px;
}

.header__title {
  font-size: 1.5rem;
  font-weight: 700;
  color: var(--text);
}

.header__controls {
  display: flex;
  align-items: center;
  gap: 12px;
}

.header__updated {
  font-size: 0.8rem;
  color: var(--text-muted);
}

.header__refresh {
  font-family: var(--font);
  font-size: 0.8rem;
  padding: 6px 10px;
  border-radius: 6px;
  border: 1px solid var(--border);
  background: var(--surface);
  color: var(--text);
  cursor: pointer;
}

.header__theme-btn {
  width: 36px;
  height: 36px;
  border-radius: 50%;
  border: 1px solid var(--border);
  background: var(--surface);
  cursor: pointer;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 1.1rem;
  transition: background var(--transition);
}
.header__theme-btn:hover { background: var(--surface-hover); }
[data-theme="light"] .icon-moon { display: none; }
[data-theme="dark"] .icon-sun { display: none; }

/* ============================================================
   FILTERS
   ============================================================ */
.filters {
  display: flex;
  align-items: flex-end;
  gap: 24px;
  padding: 16px 20px;
  background: var(--surface);
  border-radius: var(--radius);
  box-shadow: var(--shadow);
  margin-bottom: 24px;
  flex-wrap: wrap;
}

.filters__dates { display: flex; align-items: flex-end; gap: 10px; flex-wrap: wrap; }

.filters__input, .filters__select {
  font-family: var(--font);
  font-size: 0.85rem;
  padding: 7px 10px;
  border: 1px solid var(--border);
  border-radius: 6px;
  background: var(--bg);
  color: var(--text);
}

.filters__presets { display: flex; gap: 4px; }

.preset-btn {
  font-family: var(--font);
  font-size: 0.75rem;
  padding: 5px 10px;
  border: 1px solid var(--border);
  border-radius: 6px;
  background: var(--bg);
  color: var(--text-muted);
  cursor: pointer;
  transition: all var(--transition);
}
.preset-btn:hover { border-color: var(--primary); color: var(--primary); }
.preset-btn.active {
  background: var(--primary);
  color: white;
  border-color: var(--primary);
}

.filters__label { font-size: 0.8rem; color: var(--text-muted); margin-right: 8px; }

.filters__granularity { display: flex; align-items: center; gap: 6px; }

.radio-label {
  font-size: 0.82rem;
  cursor: pointer;
  display: flex;
  align-items: center;
  gap: 3px;
}

label { font-size: 0.82rem; color: var(--text-muted); display: flex; flex-direction: column; gap: 4px; }

/* ============================================================
   SUMMARY CARDS
   ============================================================ */
.cards-row {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 20px;
  margin-bottom: 24px;
}

.card {
  background: var(--surface);
  border-radius: var(--radius);
  box-shadow: var(--shadow);
  padding: 20px;
  display: flex;
  flex-direction: column;
  gap: 6px;
  transition: box-shadow var(--transition), transform var(--transition);
  position: relative;
  overflow: hidden;
}
.card:hover { box-shadow: var(--shadow-lg); transform: translateY(-2px); }

.card__label {
  font-size: 0.8rem;
  font-weight: 500;
  color: var(--text-muted);
  text-transform: uppercase;
  letter-spacing: 0.5px;
}

.card__value {
  font-size: 1.8rem;
  font-weight: 700;
  color: var(--text);
  line-height: 1.2;
}

.card__trend {
  font-size: 0.82rem;
  font-weight: 500;
  display: flex;
  align-items: center;
  gap: 4px;
}
.card__trend.up-good { color: var(--success); }
.card__trend.up-bad { color: var(--danger); }
.card__trend.down-good { color: var(--success); }
.card__trend.down-bad { color: var(--danger); }
.card__trend.neutral { color: var(--text-muted); }

.card__sparkline {
  align-self: flex-end;
  margin-top: 4px;
}

/* ============================================================
   CHART CONTAINERS
   ============================================================ */
.chart-full { margin-bottom: 24px; }

.charts-row {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 20px;
  margin-bottom: 24px;
}

.chart-container {
  background: var(--surface);
  border-radius: var(--radius);
  box-shadow: var(--shadow);
  padding: 20px;
}

.chart-title {
  font-size: 0.95rem;
  font-weight: 600;
  margin-bottom: 16px;
  color: var(--text);
}

.chart-full .chart-container { width: 100%; }

/* ============================================================
   DATA TABLE
   ============================================================ */
.table-section {
  background: var(--surface);
  border-radius: var(--radius);
  box-shadow: var(--shadow);
  padding: 20px;
  margin-bottom: 24px;
}

.table-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 16px;
}

.table-wrap { overflow-x: auto; }

.data-table {
  width: 100%;
  border-collapse: collapse;
  font-size: 0.85rem;
}

.data-table th {
  text-align: left;
  padding: 10px 14px;
  border-bottom: 2px solid var(--border);
  color: var(--text-muted);
  font-weight: 600;
  white-space: nowrap;
  user-select: none;
}

.data-table th.sortable { cursor: pointer; }
.data-table th.sortable:hover { color: var(--primary); }

.data-table td {
  padding: 9px 14px;
  border-bottom: 1px solid var(--border);
}

.data-table tbody tr:nth-child(even) { background: var(--surface-hover); }
.data-table tbody tr:hover { background: var(--primary-light); }

.table-pagination {
  display: flex;
  justify-content: center;
  gap: 6px;
  margin-top: 14px;
}

.table-pagination button {
  font-family: var(--font);
  font-size: 0.8rem;
  padding: 5px 11px;
  border: 1px solid var(--border);
  border-radius: 6px;
  background: var(--bg);
  color: var(--text);
  cursor: pointer;
}
.table-pagination button.active {
  background: var(--primary);
  color: white;
  border-color: var(--primary);
}

/* ============================================================
   BUTTONS
   ============================================================ */
.btn {
  font-family: var(--font);
  font-size: 0.82rem;
  font-weight: 500;
  padding: 8px 16px;
  border: 1px solid var(--border);
  border-radius: 8px;
  background: var(--surface);
  color: var(--text);
  cursor: pointer;
  transition: all var(--transition);
}
.btn:hover { border-color: var(--primary); color: var(--primary); }

.btn--export { display: flex; align-items: center; gap: 6px; }

/* ============================================================
   EMPTY STATE
   ============================================================ */
.empty-state {
  display: flex;
  justify-content: center;
  padding: 60px 20px;
}

.empty-state__box {
  text-align: center;
  padding: 40px;
  border: 2px dashed var(--border);
  border-radius: var(--radius);
  color: var(--text-muted);
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 16px;
}

/* ============================================================
   RESPONSIVE
   ============================================================ */
@media (max-width: 768px) {
  body { padding: 0 12px 32px; }
  .cards-row { grid-template-columns: 1fr; }
  .charts-row { grid-template-columns: 1fr; }
  .header { flex-direction: column; align-items: flex-start; }
  .filters { flex-direction: column; align-items: stretch; }
  .card__value { font-size: 1.4rem; }
}

@media (max-width: 480px) {
  .filters__presets { flex-wrap: wrap; }
  .header__controls { flex-wrap: wrap; }
}
```

**CSS implementation notes:**
- Every component uses `var(--*)` tokens so dark mode works automatically.
- Cards use `translateY(-2px)` on hover for a subtle lift effect.
- The trend classes (`up-good`, `up-bad`, `down-good`, `down-bad`) handle inverted metrics — apply the correct class in JS based on the metric's `invertedTrend` flag.
- Chart.js canvases are responsive by default; the container controls their size.

---

### 3C. js/data.js

This file defines the `DataSource` object. Adapt the implementation based on the user's chosen data source (mock, api, or static).

```js
/**
 * DataSource — Generates, fetches, or loads dashboard data.
 *
 * Data shape (array of objects):
 *   { date: 'YYYY-MM-DD', segment: 'SegmentName', m0: number, m1: number, m2: number }
 *
 * METRICS CONFIG (filled in per user input):
 */
const METRICS = [
  { key: 'm0', name: '{{METRIC_1_NAME}}', unit: '{{METRIC_1_UNIT}}', inverted: {{METRIC_1_INVERTED}} },
  { key: 'm1', name: '{{METRIC_2_NAME}}', unit: '{{METRIC_2_UNIT}}', inverted: {{METRIC_2_INVERTED}} },
  { key: 'm2', name: '{{METRIC_3_NAME}}', unit: '{{METRIC_3_UNIT}}', inverted: {{METRIC_3_INVERTED}} },
];

const SEGMENTS = [{{SEGMENTS_ARRAY}}];
// e.g., ['North', 'South', 'East', 'West']

const THRESHOLDS = [{{THRESHOLDS_ARRAY}}];
// e.g., [{ metric: 'm0', value: 50000, label: 'Revenue Target', color: '#e63946' }]

const CHART_COLORS = ['#4361ee', '#2ec4b6', '#f4a261', '#e63946', '#7209b7', '#3a86a7'];

/* ============================================================
   FORMAT ADAPTERS
   ============================================================ */
const FORMAT = {
  currency:   v => '$' + v.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 0 }),
  percentage: v => v.toFixed(1) + '%',
  number:     v => v.toLocaleString(undefined, { maximumFractionDigits: 0 }),
  duration:   v => v >= 3600 ? (v / 3600).toFixed(1) + 'h' : v >= 60 ? (v / 60).toFixed(1) + 'm' : v.toFixed(0) + 's',
};

function formatMetric(index, value) {
  return FORMAT[METRICS[index].unit](value);
}

/* ============================================================
   MOCK DATA GENERATOR
   ============================================================ */
const DataSource = {
  _raw: [],

  /**
   * Generate mock data for the given date range.
   * Uses realistic distributions per metric unit type:
   *   - currency:   random walk with seasonal variation (base 10000-100000)
   *   - percentage:  bounded 0-100, mean-reverting (base 5-80 depending on context)
   *   - number:      poisson-like with weekday/weekend patterns
   *   - duration:    log-normal distribution (seconds)
   */
  generateMock(startDate, endDate) {
    const data = [];
    const start = new Date(startDate);
    const end = new Date(endDate);
    const days = Math.round((end - start) / 86400000);

    // Base values per metric — adapt these to the domain
    const bases = METRICS.map(m => {
      switch (m.unit) {
        case 'currency':   return 10000 + Math.random() * 90000;
        case 'percentage': return 10 + Math.random() * 40;
        case 'number':     return 100 + Math.random() * 900;
        case 'duration':   return 60 + Math.random() * 300;
        default:           return 100;
      }
    });

    // Walk state per segment per metric
    const state = {};
    SEGMENTS.forEach(seg => {
      state[seg] = bases.map(b => b * (0.8 + Math.random() * 0.4));
    });

    for (let d = 0; d <= days; d++) {
      const date = new Date(start);
      date.setDate(date.getDate() + d);
      const dateStr = date.toISOString().slice(0, 10);
      const dayOfWeek = date.getDay(); // 0=Sun, 6=Sat
      const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
      // Seasonal factor: sine wave over the year
      const dayOfYear = Math.floor((date - new Date(date.getFullYear(), 0, 0)) / 86400000);
      const seasonal = 1 + 0.15 * Math.sin((dayOfYear / 365) * 2 * Math.PI);

      SEGMENTS.forEach(seg => {
        const row = { date: dateStr, segment: seg };

        METRICS.forEach((m, i) => {
          let prev = state[seg][i];
          let val;

          switch (m.unit) {
            case 'currency': {
              const drift = (Math.random() - 0.48) * prev * 0.06;
              const weekendFactor = isWeekend ? 0.7 : 1.0;
              val = Math.max(0, prev + drift) * seasonal * weekendFactor;
              break;
            }
            case 'percentage': {
              const noise = (Math.random() - 0.5) * 4;
              const meanRevert = (bases[i] - prev) * 0.05;
              val = Math.max(0, Math.min(100, prev + noise + meanRevert));
              break;
            }
            case 'number': {
              const lambda = prev * (isWeekend ? 0.6 : 1.0) * seasonal;
              // Simplified Poisson approximation
              val = Math.max(0, Math.round(lambda + (Math.random() - 0.5) * Math.sqrt(lambda) * 3));
              break;
            }
            case 'duration': {
              // Log-normal shift
              const logVal = Math.log(Math.max(1, prev));
              const noise = (Math.random() - 0.5) * 0.3;
              val = Math.max(1, Math.exp(logVal + noise));
              break;
            }
            default:
              val = prev + (Math.random() - 0.5) * 10;
          }

          state[seg][i] = val;
          row[m.key] = Math.round(val * 100) / 100;
        });

        data.push(row);
      });
    }

    this._raw = data;
    return data;
  },

  /**
   * For API data source — replace this method body:
   *
   * async fetchFromAPI() {
   *   const resp = await fetch('{{API_URL}}');
   *   const json = await resp.json();
   *   // Map json to standard shape: { date, segment, m0, m1, m2 }
   *   this._raw = json.map(item => ({
   *     date: item.{{DATE_FIELD}},
   *     segment: item.{{SEGMENT_FIELD}},
   *     m0: item.{{METRIC_1_FIELD}},
   *     m1: item.{{METRIC_2_FIELD}},
   *     m2: item.{{METRIC_3_FIELD}},
   *   }));
   *   return this._raw;
   * }
   */

  /**
   * For static JSON data source — replace generateMock call with:
   *
   * loadStatic() {
   *   this._raw = {{STATIC_JSON_DATA}};
   *   return this._raw;
   * }
   */

  /** Filter data by date range and segment */
  filter({ from, to, segment }) {
    return this._raw.filter(row => {
      if (from && row.date < from) return false;
      if (to && row.date > to) return false;
      if (segment && segment !== 'all' && row.segment !== segment) return false;
      return true;
    });
  },

  /** Aggregate data by time granularity */
  aggregate(data, granularity) {
    const buckets = {};

    data.forEach(row => {
      let key;
      const d = new Date(row.date);
      switch (granularity) {
        case 'weekly': {
          const weekStart = new Date(d);
          weekStart.setDate(d.getDate() - d.getDay());
          key = weekStart.toISOString().slice(0, 10);
          break;
        }
        case 'monthly':
          key = row.date.slice(0, 7);
          break;
        default: // daily
          key = row.date;
      }

      if (!buckets[key]) buckets[key] = { date: key, count: 0, m0: 0, m1: 0, m2: 0 };
      buckets[key].m0 += row.m0;
      buckets[key].m1 += row.m1;
      buckets[key].m2 += row.m2;
      buckets[key].count++;
    });

    // For percentages and durations, use average not sum
    return Object.values(buckets).map(b => {
      const result = { date: b.date };
      METRICS.forEach((m, i) => {
        if (m.unit === 'percentage' || m.unit === 'duration') {
          result[m.key] = b[m.key] / b.count;
        } else {
          result[m.key] = b[m.key];
        }
      });
      return result;
    }).sort((a, b) => a.date.localeCompare(b.date));
  },

  /** Get per-segment totals/averages for distribution and comparison charts */
  bySegment(data) {
    const segs = {};
    data.forEach(row => {
      if (!segs[row.segment]) segs[row.segment] = { segment: row.segment, count: 0, m0: 0, m1: 0, m2: 0 };
      segs[row.segment].m0 += row.m0;
      segs[row.segment].m1 += row.m1;
      segs[row.segment].m2 += row.m2;
      segs[row.segment].count++;
    });

    return Object.values(segs).map(s => {
      const result = { segment: s.segment };
      METRICS.forEach((m, i) => {
        result[m.key] = (m.unit === 'percentage' || m.unit === 'duration')
          ? s[m.key] / s.count
          : s[m.key];
      });
      return result;
    });
  }
};
```

**Data implementation notes:**
- The `_raw` array is the single source of truth. All charts read from filtered/aggregated views of it.
- `generateMock` produces 90 days of daily data by default (from 90 days ago to today). Adjust the range based on the chosen granularity.
- The mock generator creates **realistic patterns**: weekday/weekend variation, seasonal sine curves, mean-reverting percentages, and log-normal durations.
- When the user picks "api" data source, replace `generateMock` with `fetchFromAPI` and make the init call async.
- When the user picks "static", replace `generateMock` with `loadStatic` and embed their JSON.

---

### 3D. js/dashboard.js

```js
/**
 * Dashboard — Main controller.
 * Initializes data, binds events, renders all components.
 */
const Dashboard = {
  _charts: {},        // { timeSeries, distribution, comparison }
  _sparkCharts: [],   // [Chart, Chart, Chart]
  _sortCol: 'date',
  _sortAsc: false,
  _page: 1,
  _pageSize: 50,
  _refreshInterval: null,
  _filtered: [],

  /* ----------------------------------------------------------
     INIT
     ---------------------------------------------------------- */
  init() {
    // Generate or load data
    const end = new Date();
    const start = new Date();
    start.setDate(start.getDate() - 90);
    DataSource.generateMock(start.toISOString().slice(0, 10), end.toISOString().slice(0, 10));

    // Set default date inputs
    document.getElementById('dateFrom').value = start.toISOString().slice(0, 10);
    document.getElementById('dateTo').value = end.toISOString().slice(0, 10);

    // Populate segment dropdown
    const segSelect = document.getElementById('segmentFilter');
    SEGMENTS.forEach(s => {
      const opt = document.createElement('option');
      opt.value = s;
      opt.textContent = s;
      segSelect.appendChild(opt);
    });

    // Bind events
    this.bindEvents();

    // Initial render
    this.applyFilters();
  },

  /* ----------------------------------------------------------
     EVENT BINDING
     ---------------------------------------------------------- */
  bindEvents() {
    // Theme toggle
    document.getElementById('themeToggle').addEventListener('click', () => {
      const html = document.documentElement;
      html.dataset.theme = html.dataset.theme === 'dark' ? 'light' : 'dark';
      // Re-render charts for theme-aware grid colors
      this.applyFilters();
    });

    // Auto-refresh
    document.getElementById('autoRefresh').addEventListener('change', e => {
      this.setAutoRefresh(parseInt(e.target.value, 10));
    });

    // Filters
    document.getElementById('dateFrom').addEventListener('change', () => this.applyFilters());
    document.getElementById('dateTo').addEventListener('change', () => this.applyFilters());
    document.getElementById('segmentFilter').addEventListener('change', () => this.applyFilters());

    document.querySelectorAll('input[name="granularity"]').forEach(r => {
      r.addEventListener('change', () => this.applyFilters());
    });

    // Date presets
    document.querySelectorAll('.preset-btn').forEach(btn => {
      btn.addEventListener('click', e => {
        document.querySelectorAll('.preset-btn').forEach(b => b.classList.remove('active'));
        e.target.classList.add('active');
        const { from, to } = this.getPresetRange(e.target.dataset.preset);
        document.getElementById('dateFrom').value = from;
        document.getElementById('dateTo').value = to;
        this.applyFilters();
      });
    });

    // Table sort
    document.querySelectorAll('.data-table th.sortable').forEach(th => {
      th.addEventListener('click', () => {
        const col = th.dataset.sort;
        if (this._sortCol === col) this._sortAsc = !this._sortAsc;
        else { this._sortCol = col; this._sortAsc = true; }
        this.renderTable(this._filtered);
      });
    });

    // CSV export
    document.getElementById('csvExport').addEventListener('click', () => this.exportCSV());
  },

  /* ----------------------------------------------------------
     DATE PRESETS
     ---------------------------------------------------------- */
  getPresetRange(preset) {
    const today = new Date();
    const fmt = d => d.toISOString().slice(0, 10);
    let from;

    switch (preset) {
      case 'today':
        return { from: fmt(today), to: fmt(today) };
      case '7d':
        from = new Date(today); from.setDate(from.getDate() - 7);
        return { from: fmt(from), to: fmt(today) };
      case '30d':
        from = new Date(today); from.setDate(from.getDate() - 30);
        return { from: fmt(from), to: fmt(today) };
      case '90d':
        from = new Date(today); from.setDate(from.getDate() - 90);
        return { from: fmt(from), to: fmt(today) };
      case 'ytd':
        return { from: today.getFullYear() + '-01-01', to: fmt(today) };
      case 'all':
        return { from: '', to: '' };
      default:
        return { from: '', to: '' };
    }
  },

  /* ----------------------------------------------------------
     APPLY FILTERS (master render pipeline)
     ---------------------------------------------------------- */
  applyFilters() {
    const from = document.getElementById('dateFrom').value;
    const to = document.getElementById('dateTo').value;
    const segment = document.getElementById('segmentFilter').value;
    const granularity = document.querySelector('input[name="granularity"]:checked').value;

    const filtered = DataSource.filter({ from, to, segment });

    if (filtered.length === 0) {
      document.getElementById('emptyState').style.display = '';
      document.getElementById('summaryCards').style.display = 'none';
      document.querySelector('.chart-full').style.display = 'none';
      document.getElementById('chartsRow').style.display = 'none';
      document.querySelector('.table-section').style.display = 'none';
      return;
    }

    document.getElementById('emptyState').style.display = 'none';
    document.getElementById('summaryCards').style.display = '';
    document.querySelector('.chart-full').style.display = '';
    document.getElementById('chartsRow').style.display = '';
    document.querySelector('.table-section').style.display = '';

    this._filtered = filtered;

    const timeSeries = DataSource.aggregate(filtered, granularity);
    const bySegment = DataSource.bySegment(filtered);

    this.renderCards(timeSeries);
    this.renderTimeSeries(timeSeries);
    this.renderDistribution(bySegment);
    this.renderComparison(bySegment);
    this.renderTable(filtered);
    this.updateTimestamp();
  },

  /* ----------------------------------------------------------
     SUMMARY CARDS
     ---------------------------------------------------------- */
  renderCards(timeSeries) {
    METRICS.forEach((m, i) => {
      const values = timeSeries.map(r => r[m.key]);
      const current = values[values.length - 1] || 0;
      const previous = values[values.length - 2] || current;
      const change = previous !== 0 ? ((current - previous) / Math.abs(previous)) * 100 : 0;
      const isUp = change > 0;
      const isNeutral = Math.abs(change) < 0.5;

      // Value
      document.getElementById('cardValue' + i).textContent = formatMetric(i, current);

      // Trend with inverted logic
      const trendEl = document.getElementById('cardTrend' + i);
      if (isNeutral) {
        trendEl.textContent = '— 0%';
        trendEl.className = 'card__trend neutral';
      } else {
        const arrow = isUp ? '\u2191' : '\u2193';
        trendEl.textContent = arrow + ' ' + Math.abs(change).toFixed(1) + '%';
        if (m.inverted) {
          // Lower is better: up = bad, down = good
          trendEl.className = 'card__trend ' + (isUp ? 'up-bad' : 'down-good');
        } else {
          // Higher is better: up = good, down = bad
          trendEl.className = 'card__trend ' + (isUp ? 'up-good' : 'down-bad');
        }
      }

      // Sparkline
      this.renderSparkline(i, values);
    });
  },

  renderSparkline(index, values) {
    const canvas = document.getElementById('spark' + index);
    const ctx = canvas.getContext('2d');
    // Destroy previous
    if (this._sparkCharts[index]) { this._sparkCharts[index].destroy(); }

    this._sparkCharts[index] = new Chart(ctx, {
      type: 'line',
      data: {
        labels: values.map((_, i) => i),
        datasets: [{
          data: values,
          borderColor: CHART_COLORS[index],
          borderWidth: 1.5,
          pointRadius: 0,
          tension: 0.4,
          fill: false,
        }]
      },
      options: {
        responsive: false,
        plugins: { legend: { display: false }, tooltip: { enabled: false } },
        scales: { x: { display: false }, y: { display: false } },
        animation: false,
      }
    });
  },

  /* ----------------------------------------------------------
     TIME SERIES CHART
     ---------------------------------------------------------- */
  renderTimeSeries(timeSeries) {
    const ctx = document.getElementById('timeSeriesChart');

    // Destroy before create
    if (this._charts.timeSeries) { this._charts.timeSeries.destroy(); this._charts.timeSeries = null; }

    // Determine if we need dual Y-axis
    const unitCounts = {};
    METRICS.forEach(m => { unitCounts[m.unit] = (unitCounts[m.unit] || 0) + 1; });
    const unitTypes = Object.keys(unitCounts);
    const needsDualAxis = unitTypes.length > 1;

    let majorityUnit = METRICS[0].unit;
    let minorityUnit = null;
    if (needsDualAxis) {
      // Majority = most frequent unit type
      majorityUnit = unitTypes.reduce((a, b) => unitCounts[a] >= unitCounts[b] ? a : b);
      minorityUnit = unitTypes.find(u => u !== majorityUnit);
    }

    const isDark = document.documentElement.dataset.theme === 'dark';
    const gridColor = isDark ? 'rgba(255,255,255,0.08)' : 'rgba(0,0,0,0.06)';
    const textColor = isDark ? '#9a9ab0' : '#6c757d';

    const datasets = METRICS.map((m, i) => ({
      label: m.name,
      data: timeSeries.map(r => r[m.key]),
      borderColor: CHART_COLORS[i],
      backgroundColor: CHART_COLORS[i] + '20',
      borderWidth: 2,
      pointRadius: timeSeries.length > 60 ? 0 : 3,
      pointHoverRadius: 5,
      tension: 0.3,
      fill: false,
      spanGaps: true,
      yAxisID: (needsDualAxis && m.unit === minorityUnit) ? 'y1' : 'y',
    }));

    const scales = {
      x: {
        grid: { color: gridColor },
        ticks: { color: textColor, maxTicksLimit: 12 },
      },
      y: {
        position: 'left',
        grid: { color: gridColor },
        ticks: {
          color: textColor,
          callback: v => FORMAT[majorityUnit](v),
        },
      },
    };

    if (needsDualAxis) {
      scales.y1 = {
        position: 'right',
        grid: { drawOnChartArea: false },
        ticks: {
          color: textColor,
          callback: v => FORMAT[minorityUnit](v),
        },
      };
    }

    // Threshold annotations
    const annotations = THRESHOLDS.map(t => ({
      type: 'line',
      yMin: t.value,
      yMax: t.value,
      borderColor: t.color || '#e63946',
      borderWidth: 1.5,
      borderDash: [6, 4],
      label: {
        display: true,
        content: t.label,
        position: 'start',
        font: { size: 11 },
      },
      yScaleID: (needsDualAxis && METRICS.find(m => m.key === t.metric)?.unit === minorityUnit) ? 'y1' : 'y',
    }));

    this._charts.timeSeries = new Chart(ctx, {
      type: 'line',
      data: {
        labels: timeSeries.map(r => r.date),
        datasets,
      },
      options: {
        responsive: true,
        maintainAspectRatio: true,
        interaction: { mode: 'index', intersect: false },
        plugins: {
          legend: { labels: { color: textColor, usePointStyle: true, padding: 16 } },
          tooltip: {
            callbacks: {
              label: ctx => {
                const m = METRICS[ctx.datasetIndex];
                return m.name + ': ' + FORMAT[m.unit](ctx.parsed.y);
              }
            }
          },
          annotation: { annotations },
        },
        scales,
      }
    });
  },

  /* ----------------------------------------------------------
     DISTRIBUTION CHART (Doughnut)
     ---------------------------------------------------------- */
  renderDistribution(bySegment) {
    const ctx = document.getElementById('distributionChart');
    if (!ctx) return; // Skipped when single segment

    if (this._charts.distribution) { this._charts.distribution.destroy(); this._charts.distribution = null; }

    // Primary metric broken down by segment — top 6 + Other
    const sorted = [...bySegment].sort((a, b) => b.m0 - a.m0);
    let labels, values;
    if (sorted.length > 6) {
      const top6 = sorted.slice(0, 6);
      const otherSum = sorted.slice(6).reduce((sum, s) => sum + s.m0, 0);
      labels = [...top6.map(s => s.segment.length > 28 ? s.segment.slice(0, 28) + '...' : s.segment), 'Other'];
      values = [...top6.map(s => s.m0), otherSum];
    } else {
      labels = sorted.map(s => s.segment.length > 28 ? s.segment.slice(0, 28) + '...' : s.segment);
      values = sorted.map(s => s.m0);
    }

    const isDark = document.documentElement.dataset.theme === 'dark';
    const textColor = isDark ? '#9a9ab0' : '#6c757d';

    this._charts.distribution = new Chart(ctx, {
      type: 'doughnut',
      data: {
        labels,
        datasets: [{
          data: values,
          backgroundColor: CHART_COLORS.slice(0, labels.length).concat(['#999']),
          borderWidth: 2,
          borderColor: isDark ? '#1a1a2e' : '#ffffff',
        }]
      },
      options: {
        responsive: true,
        plugins: {
          legend: { position: 'bottom', labels: { color: textColor, padding: 12, usePointStyle: true } },
          tooltip: {
            callbacks: {
              label: ctx => {
                const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
                const pct = ((ctx.parsed / total) * 100).toFixed(1);
                return ctx.label + ': ' + formatMetric(0, ctx.parsed) + ' (' + pct + '%)';
              }
            }
          },
        },
      },
      plugins: [{
        // Custom plugin: draw percentage inside slices
        id: 'doughnutPercentLabels',
        afterDraw(chart) {
          const { ctx: c, data } = chart;
          const total = data.datasets[0].data.reduce((a, b) => a + b, 0);
          if (total === 0) return;

          chart.getDatasetMeta(0).data.forEach((arc, i) => {
            const pct = ((data.datasets[0].data[i] / total) * 100).toFixed(0);
            if (pct < 5) return; // Skip tiny slices

            const { x, y } = arc.tooltipPosition();
            c.save();
            c.font = 'bold 11px DM Sans, sans-serif';
            c.fillStyle = '#fff';
            c.textAlign = 'center';
            c.textBaseline = 'middle';
            c.fillText(pct + '%', x, y);
            c.restore();
          });
        }
      }]
    });
  },

  /* ----------------------------------------------------------
     COMPARISON CHART (Grouped Bar)
     ---------------------------------------------------------- */
  renderComparison(bySegment) {
    const ctx = document.getElementById('comparisonChart');
    if (this._charts.comparison) { this._charts.comparison.destroy(); this._charts.comparison = null; }

    const labels = bySegment.map(s => s.segment.length > 28 ? s.segment.slice(0, 28) + '...' : s.segment);
    const isDark = document.documentElement.dataset.theme === 'dark';
    const gridColor = isDark ? 'rgba(255,255,255,0.08)' : 'rgba(0,0,0,0.06)';
    const textColor = isDark ? '#9a9ab0' : '#6c757d';

    // Check if normalization is needed (ratio between max values > 100x)
    const maxVals = METRICS.map((m, i) => Math.max(...bySegment.map(s => s[m.key])));
    const ratio = Math.max(...maxVals) / Math.max(1, Math.min(...maxVals));
    const normalize = ratio > 100;

    const datasets = METRICS.map((m, i) => {
      let data = bySegment.map(s => s[m.key]);
      let label = m.name;
      if (normalize) {
        const maxVal = Math.max(...data, 1);
        data = data.map(v => (v / maxVal) * 100);
        label += ' (normalized %)';
      }
      return {
        label,
        data,
        backgroundColor: CHART_COLORS[i] + 'cc',
        borderColor: CHART_COLORS[i],
        borderWidth: 1,
        borderRadius: 4,
      };
    });

    this._charts.comparison = new Chart(ctx, {
      type: 'bar',
      data: { labels, datasets },
      options: {
        responsive: true,
        plugins: {
          legend: { labels: { color: textColor, usePointStyle: true, padding: 12 } },
          tooltip: {
            callbacks: {
              label: ctx => {
                if (normalize) return ctx.dataset.label + ': ' + ctx.parsed.y.toFixed(1) + '%';
                const m = METRICS[ctx.datasetIndex];
                return m.name + ': ' + FORMAT[m.unit](ctx.parsed.y);
              }
            }
          },
        },
        scales: {
          x: { grid: { display: false }, ticks: { color: textColor } },
          y: {
            grid: { color: gridColor },
            ticks: {
              color: textColor,
              callback: normalize ? (v => v + '%') : undefined,
            },
          },
        },
      }
    });
  },

  /* ----------------------------------------------------------
     DATA TABLE
     ---------------------------------------------------------- */
  renderTable(data) {
    const sorted = [...data].sort((a, b) => {
      let va = a[this._sortCol] ?? a.date;
      let vb = b[this._sortCol] ?? b.date;
      if (typeof va === 'string') return this._sortAsc ? va.localeCompare(vb) : vb.localeCompare(va);
      return this._sortAsc ? va - vb : vb - va;
    });

    const totalPages = Math.ceil(sorted.length / this._pageSize);
    if (this._page > totalPages) this._page = totalPages || 1;
    const start = (this._page - 1) * this._pageSize;
    const pageData = sorted.slice(start, start + this._pageSize);

    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = pageData.map(row =>
      '<tr>' +
        '<td>' + row.date + '</td>' +
        '<td>' + row.segment + '</td>' +
        '<td>' + formatMetric(0, row.m0) + '</td>' +
        '<td>' + formatMetric(1, row.m1) + '</td>' +
        '<td>' + formatMetric(2, row.m2) + '</td>' +
      '</tr>'
    ).join('');

    // Pagination
    const pagEl = document.getElementById('pagination');
    if (totalPages <= 1) { pagEl.innerHTML = ''; return; }

    let pagHTML = '';
    // Show max 7 page buttons
    const maxButtons = 7;
    let startPage = Math.max(1, this._page - 3);
    let endPage = Math.min(totalPages, startPage + maxButtons - 1);
    if (endPage - startPage < maxButtons - 1) startPage = Math.max(1, endPage - maxButtons + 1);

    if (this._page > 1) pagHTML += '<button onclick="Dashboard.goToPage(' + (this._page - 1) + ')">&#8592;</button>';
    for (let p = startPage; p <= endPage; p++) {
      pagHTML += '<button class="' + (p === this._page ? 'active' : '') + '" onclick="Dashboard.goToPage(' + p + ')">' + p + '</button>';
    }
    if (this._page < totalPages) pagHTML += '<button onclick="Dashboard.goToPage(' + (this._page + 1) + ')">&#8594;</button>';
    pagEl.innerHTML = pagHTML;
  },

  goToPage(p) {
    this._page = p;
    this.renderTable(this._filtered);
  },

  /* ----------------------------------------------------------
     CSV EXPORT
     ---------------------------------------------------------- */
  exportCSV() {
    const headers = ['Date', 'Segment', ...METRICS.map(m => m.name)];
    const rows = this._filtered.map(row =>
      [row.date, row.segment, ...METRICS.map((m, i) => row[m.key])].join(',')
    );
    const csv = [headers.join(','), ...rows].join('\n');
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'dashboard-export.csv';
    a.click();
    URL.revokeObjectURL(url);
  },

  /* ----------------------------------------------------------
     AUTO-REFRESH
     ---------------------------------------------------------- */
  setAutoRefresh(seconds) {
    if (this._refreshInterval) clearInterval(this._refreshInterval);
    this._refreshInterval = null;
    if (seconds > 0) {
      this._refreshInterval = setInterval(() => this.refresh(), seconds * 1000);
    }
  },

  refresh() {
    // Re-generate mock data (simulates new data arriving)
    const end = new Date();
    const start = new Date();
    start.setDate(start.getDate() - 90);
    DataSource.generateMock(start.toISOString().slice(0, 10), end.toISOString().slice(0, 10));
    this.applyFilters();
  },

  updateTimestamp() {
    document.getElementById('lastUpdated').textContent = 'Updated ' + new Date().toLocaleTimeString();
  },

  /* ----------------------------------------------------------
     LOAD SAMPLE DATA (for empty state)
     ---------------------------------------------------------- */
  loadSampleData() {
    const end = new Date();
    const start = new Date();
    start.setDate(start.getDate() - 30);
    DataSource.generateMock(start.toISOString().slice(0, 10), end.toISOString().slice(0, 10));
    document.getElementById('dateFrom').value = start.toISOString().slice(0, 10);
    document.getElementById('dateTo').value = end.toISOString().slice(0, 10);
    this.applyFilters();
  },
};

/* ----------------------------------------------------------
   BOOT
   ---------------------------------------------------------- */
document.addEventListener('DOMContentLoaded', () => Dashboard.init());
```

**JS implementation notes:**
- The **destroy-before-create** pattern is critical. Every `render*` method calls `.destroy()` on the previous Chart instance before creating a new one. Without this, Chart.js leaks canvases.
- The **dual Y-axis logic** counts how many metrics share each unit type. The majority unit type goes on the left Y-axis, the minority on the right with `grid: { drawOnChartArea: false }` to avoid overlapping gridlines.
- **Inverted trend logic**: In `renderCards`, when `m.inverted` is true, an upward change gets `up-bad` class (red), and a downward change gets `down-good` class (green). This is the opposite of the default behavior.
- **Normalization in comparison chart**: When the largest metric's max value is 100x+ the smallest, all values are normalized to percentages of their respective maxima. This prevents tiny bars next to huge bars.
- `spanGaps: true` on all line datasets handles missing data points.

---

## Edge Cases — Handle All of These

### Missing data
Set `spanGaps: true` on every line dataset so gaps in dates don't break the line.

### No data for current filters
Show the `#emptyState` section (dashed border box with "No data available" message and a "Load Sample Data" button). Hide all other sections.

### Single segment
If `SEGMENTS.length === 1`:
- Remove the distribution chart `<div>` from HTML entirely.
- Give the comparison chart the full width (`chart-full` class instead of `charts-row`).
- Hide the segment filter dropdown.

### Very long labels
Truncate any label longer than 28 characters: `label.length > 28 ? label.slice(0, 28) + '...' : label`

### All zero values
If every data point is zero for all metrics, show the charts but overlay a centered "No activity in this period" banner using a CSS pseudo-element or injected div.

### Large datasets (>50 rows in table)
Paginate with 50 rows per page. Show page buttons with a max of 7 visible at a time, with prev/next arrows.

### Theme switch mid-session
When the user toggles dark mode, call `applyFilters()` to re-render all charts with updated grid colors and text colors pulled from the current theme.

---

## Domain Adaptation Examples

These show how the SAME skill produces different dashboards:

### Example 1: E-Commerce Dashboard
**Input**: `Orders (number), Revenue (currency), Return Rate (percentage, inverted)`
- Segments: "Web, Mobile, In-Store, Marketplace"
- Thresholds: "Revenue target: $50,000"
- Mock data: Orders show weekday spikes, Revenue follows random walk, Return Rate mean-reverts around 8%
- Dual Y-axis: Left = number/currency, Right = percentage
- Trend cards: Return Rate down arrow is GREEN (inverted)

### Example 2: Weather Station Dashboard
**Input**: `Temperature (number), Humidity (percentage), Wind Speed (number)`
- Segments: "Station A, Station B, Station C"
- Thresholds: "Freeze warning: 32", "High wind: 25"
- Mock data: Temperature has seasonal sine, Humidity bounded 30-95%, Wind is log-normal
- Single Y-axis: All numbers/percentages on similar scales
- Trend cards: All default (higher = neutral, no inversion)

### Example 3: SaaS Metrics Dashboard
**Input**: `Active Users (number), MRR (currency), Churn Rate (percentage, inverted)`
- Segments: "Enterprise, Pro, Starter, Free"
- Thresholds: "MRR goal: $100,000", "Churn max: 5%"
- Mock data: Users grow with slight randomness, MRR random walk upward, Churn mean-reverts around 4%
- Dual Y-axis: Left = number/currency, Right = percentage
- Trend cards: Churn down arrow is GREEN (inverted)

---

## Checklist Before Delivering

Before presenting the final code, verify:

- [ ] All 4 files are generated (index.html, css/dashboard.css, js/dashboard.js, js/data.js)
- [ ] All `{{PLACEHOLDER}}` values are replaced with actual user values
- [ ] METRICS array matches user's 3 metrics exactly (name, unit, inverted flag)
- [ ] SEGMENTS array matches user's categories
- [ ] THRESHOLDS array includes user's target lines (or is empty array)
- [ ] Dark mode toggle works (CSS variables + chart re-render)
- [ ] Inverted metrics show correct trend colors (up=bad, down=good)
- [ ] Dual Y-axis logic is correct for the given unit types
- [ ] Date presets calculate correct ranges
- [ ] Table pagination works for datasets >50 rows
- [ ] CSV export includes all filtered data
- [ ] Empty state shows when no data matches filters
- [ ] Sparklines render in summary cards
- [ ] Charts use destroy-before-create pattern
- [ ] Responsive layout works (3-col to 1-col on mobile)
- [ ] All CDN scripts load (Chart.js 4.4.0 + annotation plugin 3.0.1)
- [ ] Google Fonts DM Sans loads
- [ ] No console errors on first load
