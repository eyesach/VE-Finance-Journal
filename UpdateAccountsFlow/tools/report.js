import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const TMP_DIR = path.join(__dirname, '..', '.tmp');

/**
 * Generate an HTML report from filtered sales data.
 *
 * @param {Sale[]} sales - Filtered sales array
 * @param {object} options - { dateLabel, sourceLabel, showMonthly, itemsCache }
 * @returns {string} path to generated report.html
 */
export function generateReport(sales, { dateLabel = '', sourceLabel = '', showMonthly = false, itemsCache = new Map() } = {}) {
  const totals = calculateTotals(sales);
  const onlineTotals = calculateTotals(sales.filter(s => s.source === 'online'));
  const tradeshowTotals = calculateTotals(sales.filter(s => s.source === 'tradeshow'));
  const products = groupByProduct(sales, itemsCache);
  const monthlyData = showMonthly ? groupByMonth(sales) : null;
  const hasItemData = products.some(p => p.description !== 'Unscraped Orders');

  const html = buildHTML({
    dateLabel,
    sourceLabel,
    totals,
    onlineTotals,
    tradeshowTotals,
    products,
    monthlyData,
    totalTransactions: sales.length,
    hasItemData,
  });

  if (!fs.existsSync(TMP_DIR)) {
    fs.mkdirSync(TMP_DIR, { recursive: true });
  }
  const outPath = path.join(TMP_DIR, 'report.html');
  fs.writeFileSync(outPath, html, 'utf-8');
  console.log(`  Report saved to ${outPath}`);
  return outPath;
}

function calculateTotals(sales) {
  let subtotal = 0, tax = 0, total = 0, shipping = 0, discount = 0;
  for (const s of sales) {
    subtotal += s.subtotal;
    tax += s.tax;
    total += s.total;
    shipping += s.shipping;
    discount += s.discount;
  }
  return { subtotal, tax, total, shipping, discount, count: sales.length };
}

/**
 * Group by product using scraped line items when available,
 * falling back to order-level data for unscraped transactions.
 */
function groupByProduct(sales, itemsCache) {
  const map = new Map();
  let unscrapedSubtotal = 0;
  let unscrapedTax = 0;
  let unscrapedCount = 0;

  for (const s of sales) {
    const items = itemsCache.get(s.transactionNo);

    if (items && items.length > 0) {
      // Use scraped line items
      for (const item of items) {
        const key = (item.name || 'Unknown Item').toLowerCase().trim();
        if (!map.has(key)) {
          map.set(key, {
            description: item.name || 'Unknown Item',
            productNumber: item.productNumber || null,
            count: 0,
            totalQuantity: 0,
            totalRevenue: 0,
            totalTax: 0,
            prices: [],
          });
        }
        const entry = map.get(key);
        entry.count++;
        entry.totalQuantity += item.quantity || 1;
        entry.totalRevenue += item.amount || 0;
        // Approximate tax per item (proportional to item amount / order subtotal)
        if (s.subtotal > 0) {
          entry.totalTax += s.tax * (item.amount / s.subtotal);
        }
        entry.prices.push(item.price || 0);
        if (item.productNumber && !entry.productNumber) {
          entry.productNumber = item.productNumber;
        }
      }
    } else {
      // Unscraped — accumulate into summary
      unscrapedSubtotal += s.subtotal;
      unscrapedTax += s.tax;
      unscrapedCount++;
    }
  }

  // Convert to array
  const products = [...map.values()].map(p => ({
    ...p,
    unitPrice: p.prices.length > 0 ? p.prices[0] : 0,
    priceConsistent: p.prices.every(pr => Math.abs(pr - p.prices[0]) < 0.01),
  }));
  products.sort((a, b) => b.totalQuantity - a.totalQuantity);

  // Add unscraped summary row if needed
  if (unscrapedCount > 0) {
    products.push({
      description: 'Unscraped Orders',
      productNumber: null,
      count: unscrapedCount,
      totalQuantity: unscrapedCount,
      totalRevenue: unscrapedSubtotal,
      totalTax: unscrapedTax,
      unitPrice: 0,
      priceConsistent: false,
      prices: [],
    });
  }

  return products;
}

function groupByMonth(sales) {
  const map = new Map();
  for (const s of sales) {
    const key = `${s.date.getFullYear()}-${String(s.date.getMonth() + 1).padStart(2, '0')}`;
    if (!map.has(key)) {
      map.set(key, { month: key, subtotal: 0, tax: 0, total: 0, count: 0 });
    }
    const entry = map.get(key);
    entry.subtotal += s.subtotal;
    entry.tax += s.tax;
    entry.total += s.total;
    entry.count++;
  }
  return [...map.values()].sort((a, b) => a.month.localeCompare(b.month));
}

function fmt(n) {
  return '$' + n.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function buildHTML({ dateLabel, sourceLabel, totals, onlineTotals, tradeshowTotals, products, monthlyData, totalTransactions, hasItemData }) {
  const showBothSources = onlineTotals.count > 0 && tradeshowTotals.count > 0;

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>VE Sales Report</title>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    background: #f8f9fa; color: #212529; padding: 24px;
  }
  .container { max-width: 960px; margin: 0 auto; }
  h1 { font-size: 1.6rem; margin-bottom: 4px; color: #4a90a4; }
  .date-label { color: #6c757d; margin-bottom: 24px; font-size: 0.95rem; }
  .cards { display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px; margin-bottom: 32px; }
  .card {
    background: #fff; border-radius: 8px; padding: 20px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.1); text-align: center;
  }
  .card .label { font-size: 0.8rem; text-transform: uppercase; color: #6c757d; letter-spacing: 0.5px; margin-bottom: 8px; }
  .card .value { font-size: 1.5rem; font-weight: 700; }
  .card .value.pretax { color: #4a90a4; }
  .card .value.tax { color: #dc3545; }
  .card .value.total { color: #28a745; }
  .card .sub { font-size: 0.75rem; color: #6c757d; margin-top: 4px; }
  .section { margin-bottom: 32px; }
  .section h2 { font-size: 1.1rem; margin-bottom: 12px; color: #333; border-bottom: 2px solid #4a90a4; padding-bottom: 6px; }
  table { width: 100%; border-collapse: collapse; background: #fff; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
  th { background: #4a90a4; color: #fff; padding: 10px 14px; text-align: left; font-size: 0.8rem; text-transform: uppercase; letter-spacing: 0.5px; }
  td { padding: 10px 14px; border-bottom: 1px solid #eef0f2; font-size: 0.9rem; }
  tr:last-child td { border-bottom: none; }
  tr:hover td { background: #f0f4f8; }
  .num { text-align: right; font-variant-numeric: tabular-nums; }
  .source-breakdown { display: grid; grid-template-columns: ${showBothSources ? '1fr 1fr' : '1fr'}; gap: 16px; margin-bottom: 32px; }
  .source-card {
    background: #fff; border-radius: 8px; padding: 16px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.1);
  }
  .source-card h3 { font-size: 0.9rem; color: #4a90a4; margin-bottom: 10px; }
  .source-card .stat { display: flex; justify-content: space-between; padding: 4px 0; font-size: 0.85rem; }
  .source-card .stat .label { color: #6c757d; }
  .total-row td { font-weight: 700; background: #eef0f2 !important; }
  .muted { color: #6c757d; font-style: italic; }
  .product-num { font-size: 0.75rem; color: #6c757d; }
</style>
</head>
<body>
<div class="container">
  <h1>VE Sales Report${sourceLabel ? ` — ${escapeHtml(sourceLabel)}` : ''}</h1>
  <div class="date-label">${dateLabel} &mdash; ${totalTransactions} transaction${totalTransactions !== 1 ? 's' : ''}</div>

  <div class="cards">
    <div class="card">
      <div class="label">Pretax Total</div>
      <div class="value pretax">${fmt(totals.subtotal)}</div>
    </div>
    <div class="card">
      <div class="label">Sales Tax</div>
      <div class="value tax">${fmt(totals.tax)}</div>
    </div>
    <div class="card">
      <div class="label">Post-Tax Total</div>
      <div class="value total">${fmt(totals.total)}</div>
      ${totals.shipping > 0 ? `<div class="sub">Shipping: ${fmt(totals.shipping)}</div>` : ''}
      ${totals.discount > 0 ? `<div class="sub">Discounts: -${fmt(totals.discount)}</div>` : ''}
    </div>
  </div>

  <div class="source-breakdown">
${onlineTotals.count > 0 ? `    <div class="source-card">
      <h3>Online Sales (Store Manager)</h3>
      <div class="stat"><span class="label">Transactions</span><span>${onlineTotals.count}</span></div>
      <div class="stat"><span class="label">Pretax</span><span>${fmt(onlineTotals.subtotal)}</span></div>
      <div class="stat"><span class="label">Tax</span><span>${fmt(onlineTotals.tax)}</span></div>
      <div class="stat"><span class="label">Total</span><span>${fmt(onlineTotals.total)}</span></div>
    </div>` : ''}
${tradeshowTotals.count > 0 ? `    <div class="source-card">
      <h3>Trade Show Sales (POS)</h3>
      <div class="stat"><span class="label">Transactions</span><span>${tradeshowTotals.count}</span></div>
      <div class="stat"><span class="label">Pretax</span><span>${fmt(tradeshowTotals.subtotal)}</span></div>
      <div class="stat"><span class="label">Tax</span><span>${fmt(tradeshowTotals.tax)}</span></div>
      <div class="stat"><span class="label">Total</span><span>${fmt(tradeshowTotals.total)}</span></div>
    </div>` : ''}
  </div>

${monthlyData ? `
  <div class="section">
    <h2>Monthly Breakdown</h2>
    <table>
      <thead><tr><th>Month</th><th class="num">Transactions</th><th class="num">Pretax</th><th class="num">Tax</th><th class="num">Total</th></tr></thead>
      <tbody>
${monthlyData.map(m => `        <tr><td>${m.month}</td><td class="num">${m.count}</td><td class="num">${fmt(m.subtotal)}</td><td class="num">${fmt(m.tax)}</td><td class="num">${fmt(m.total)}</td></tr>`).join('\n')}
        <tr class="total-row"><td>Total</td><td class="num">${totals.count}</td><td class="num">${fmt(totals.subtotal)}</td><td class="num">${fmt(totals.tax)}</td><td class="num">${fmt(totals.total)}</td></tr>
      </tbody>
    </table>
  </div>
` : ''}

  <div class="section">
    <h2>Product Breakdown${!hasItemData ? ' <span class="muted">(no product data scraped)</span>' : ''}</h2>
    <table>
      <thead><tr><th>Product</th><th class="num">Qty Sold</th><th class="num">Unit Price</th><th class="num">Total Revenue</th><th class="num">Tax (est.)</th></tr></thead>
      <tbody>
${products.map(p => {
  const desc = p.description === 'Unscraped Orders'
    ? `<span class="muted">${escapeHtml(p.description)} (${p.count} orders)</span>`
    : escapeHtml(p.description) + (p.productNumber ? `<br><span class="product-num">${escapeHtml(p.productNumber)}</span>` : '');
  const qty = p.description === 'Unscraped Orders' ? p.count : p.totalQuantity;
  const price = p.description === 'Unscraped Orders' ? '' : (p.priceConsistent ? fmt(p.unitPrice) : 'Varies');
  return `        <tr><td>${desc}</td><td class="num">${qty}</td><td class="num">${price}</td><td class="num">${fmt(p.totalRevenue)}</td><td class="num">${fmt(p.totalTax)}</td></tr>`;
}).join('\n')}
        <tr class="total-row"><td>Total</td><td class="num">${products.reduce((s, p) => s + (p.description === 'Unscraped Orders' ? p.count : p.totalQuantity), 0)}</td><td class="num"></td><td class="num">${fmt(totals.subtotal)}</td><td class="num">${fmt(totals.tax)}</td></tr>
      </tbody>
    </table>
  </div>

</div>
</body>
</html>`;
}

function escapeHtml(str) {
  return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
