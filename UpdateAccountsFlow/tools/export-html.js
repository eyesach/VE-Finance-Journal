/**
 * Generate a standalone interactive HTML report from sales data.
 * The HTML file includes date filters, source tabs, and sortable tables — no external dependencies.
 *
 * @param {Sale[]} sales - Sales array (date is a Date object)
 * @param {Map} itemsCache - Map of transactionNo → line items
 * @returns {string} Complete HTML string
 */
export function generateHtmlReport(sales, itemsCache) {
  // Serialize data for embedding in the HTML
  const round2 = (v) => Math.round(v * 100) / 100;
  const salesData = sales.map(s => ({
    transactionNo: s.transactionNo,
    date: s.date.toISOString().split('T')[0],
    billingName: s.billingName || '',
    source: s.source,
    subtotal: round2(s.subtotal),
    tax: round2(s.tax),
    shipping: round2(s.shipping),
    discount: round2(s.discount),
    total: round2(s.total),
  }));

  const itemsData = {};
  for (const s of sales) {
    const items = itemsCache.get(s.transactionNo);
    if (items && Array.isArray(items) && items.length > 0) {
      itemsData[s.transactionNo] = items.map(i => ({
        name: i.name || 'Unknown',
        productNumber: i.productNumber || null,
        price: round2(i.price || 0),
        quantity: i.quantity || 1,
        amount: round2(i.amount || 0),
        inferred: !!i.inferred,
      }));
    }
  }

  const dateRange = sales.length > 0 ? getDateRange(sales) : { from: '', to: '' };

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>VE Sales Report</title>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: #f8f9fa; color: #212529; padding: 24px; }
  .container { max-width: 1100px; margin: 0 auto; }
  h1 { font-size: 1.6rem; margin-bottom: 4px; color: #4a90a4; }
  .subtitle { color: #6c757d; margin-bottom: 20px; font-size: 0.9rem; }

  .controls { display: flex; flex-wrap: wrap; gap: 12px; align-items: flex-end; margin-bottom: 20px; padding: 16px; background: #fff; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
  .control-group { display: flex; flex-direction: column; gap: 4px; }
  .control-group label { font-size: 0.75rem; text-transform: uppercase; color: #6c757d; letter-spacing: 0.5px; font-weight: 600; }
  .controls select, .controls input[type="date"] { padding: 6px 10px; border: 1px solid #dee2e6; border-radius: 4px; font-size: 0.85rem; background: #fff; }
  .controls select:focus, .controls input:focus { outline: none; border-color: #4a90a4; }
  .spacer { flex: 1; }
  .reset-btn { padding: 6px 14px; background: #6c757d; color: #fff; border: none; border-radius: 4px; font-size: 0.8rem; cursor: pointer; font-weight: 600; }
  .reset-btn:hover { background: #565e64; }

  .cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 14px; margin-bottom: 20px; }
  .card { background: #fff; border-radius: 8px; padding: 18px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); text-align: center; }
  .card .label { font-size: 0.75rem; text-transform: uppercase; color: #6c757d; letter-spacing: 0.5px; margin-bottom: 6px; }
  .card .value { font-size: 1.4rem; font-weight: 700; }
  .card .value.pretax { color: #4a90a4; }
  .card .value.tax { color: #dc3545; }
  .card .value.total { color: #28a745; }
  .card .sub { font-size: 0.7rem; color: #6c757d; margin-top: 4px; }
  .card.hidden { display: none; }

  .source-breakdown { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin-bottom: 20px; }
  .source-card { background: #fff; border-radius: 8px; padding: 14px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
  .source-card h3 { font-size: 0.85rem; color: #4a90a4; margin-bottom: 8px; }
  .source-card .stat { display: flex; justify-content: space-between; padding: 3px 0; font-size: 0.82rem; }
  .source-card .stat .label { color: #6c757d; }

  .tabs { display: flex; gap: 0; margin-bottom: 0; border-bottom: 2px solid #4a90a4; }
  .tab-btn { padding: 8px 20px; background: #e8f0f4; border: none; cursor: pointer; font-size: 0.85rem; font-weight: 600; color: #4a90a4; border-radius: 6px 6px 0 0; }
  .tab-btn.active { background: #4a90a4; color: #fff; }
  .tab-content { display: none; }
  .tab-content.active { display: block; }

  .section { margin-bottom: 24px; }
  .section h2 { font-size: 1rem; margin-bottom: 10px; color: #333; border-bottom: 2px solid #4a90a4; padding-bottom: 5px; display: flex; justify-content: space-between; align-items: center; }
  .section h2 .count { font-size: 0.8rem; color: #6c757d; font-weight: 400; }
  table { width: 100%; border-collapse: collapse; background: #fff; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
  th { background: #4a90a4; color: #fff; padding: 9px 12px; text-align: left; font-size: 0.75rem; text-transform: uppercase; letter-spacing: 0.5px; cursor: pointer; user-select: none; white-space: nowrap; }
  th:hover { background: #3d7a8e; }
  th .arrow { margin-left: 4px; font-size: 0.65rem; }
  td { padding: 8px 12px; border-bottom: 1px solid #eef0f2; font-size: 0.85rem; }
  tr:last-child td { border-bottom: none; }
  tr:hover td { background: #f0f4f8; }
  .num { text-align: right; font-variant-numeric: tabular-nums; }
  .total-row td { font-weight: 700; background: #eef0f2 !important; }
  .muted { color: #6c757d; font-style: italic; }
  .product-num { font-size: 0.7rem; color: #6c757d; }
  .source-badge { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 0.7rem; font-weight: 600; }
  .source-badge.online { background: #d4edda; color: #155724; }
  .source-badge.tradeshow { background: #cce5ff; color: #004085; }

  .source-section { margin-bottom: 32px; }
  .source-section h3 { font-size: 1.05rem; color: #333; margin-bottom: 12px; padding: 8px 12px; border-radius: 6px; }
  .source-section h3.online-header { background: #d4edda; color: #155724; }
  .source-section h3.tradeshow-header { background: #cce5ff; color: #004085; }

  .generated { text-align: center; color: #adb5bd; font-size: 0.75rem; margin-top: 32px; padding-top: 16px; border-top: 1px solid #eef0f2; }

  @media print { body { padding: 12px; } .controls { display: none; } }
  @media (max-width: 768px) { .cards { grid-template-columns: 1fr 1fr; } .source-breakdown { grid-template-columns: 1fr; } .controls { flex-direction: column; } body { padding: 12px; } }
</style>
</head>
<body>
<div class="container">
  <h1>VE Sales Report</h1>
  <div class="subtitle" id="subtitle"></div>

  <div class="controls">
    <div class="control-group">
      <label>Source</label>
      <select id="filterSource">
        <option value="both">Both</option>
        <option value="online">Online</option>
        <option value="tradeshow">Trade Show</option>
      </select>
    </div>
    <div class="control-group">
      <label>From</label>
      <input type="date" id="filterFrom" value="${dateRange.from}">
    </div>
    <div class="control-group">
      <label>To</label>
      <input type="date" id="filterTo" value="${dateRange.to}">
    </div>
    <div class="control-group">
      <label>&nbsp;</label>
      <select id="filterPreset">
        <option value="">Preset...</option>
        <option value="all">All time</option>
        <option value="thisMonth">This month</option>
        <option value="lastMonth">Last month</option>
        <option value="thisQuarter">This quarter</option>
      </select>
    </div>
    <div class="spacer"></div>
    <div class="control-group">
      <label>&nbsp;</label>
      <button class="reset-btn" onclick="resetFilters()">Reset</button>
    </div>
  </div>

  <div class="cards" id="cards"></div>
  <div class="source-breakdown" id="sourceBreakdown"></div>

  <div class="tabs" id="viewTabs">
    <button class="tab-btn active" data-tab="bySource">By Source</button>
    <button class="tab-btn" data-tab="products">All Products</button>
    <button class="tab-btn" data-tab="transactions">All Transactions</button>
  </div>

  <div id="tab-bySource" class="tab-content active"></div>
  <div id="tab-products" class="tab-content"></div>
  <div id="tab-transactions" class="tab-content"></div>

  <div class="generated">Generated ${new Date().toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' })} — VE Sales Dashboard</div>
</div>

<script>
const ALL_SALES = ${JSON.stringify(salesData)};
const ITEMS_CACHE = ${JSON.stringify(itemsData)};

let filtered = [];
let productSortCol = 'qty', productSortDir = 'desc';
let txSortCol = 'date', txSortDir = 'desc';

function fmt(n) { return '$' + n.toFixed(2).replace(/\\B(?=(\\d{3})+(?!\\d))/g, ','); }
function fmtDate(iso) { const d = new Date(iso + 'T12:00:00'); return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }); }
function esc(s) { const d = document.createElement('div'); d.textContent = s; return d.innerHTML; }

function applyFilters() {
  const source = document.getElementById('filterSource').value;
  const from = document.getElementById('filterFrom').value;
  const to = document.getElementById('filterTo').value;
  filtered = ALL_SALES.filter(s => {
    if (source !== 'both' && s.source !== source) return false;
    if (from && s.date < from) return false;
    if (to && s.date > to) return false;
    return true;
  });
  renderAll();
}

function resetFilters() {
  document.getElementById('filterSource').value = 'both';
  document.getElementById('filterFrom').value = '';
  document.getElementById('filterTo').value = '';
  document.getElementById('filterPreset').value = '';
  applyFilters();
}

function applyPreset() {
  const preset = document.getElementById('filterPreset').value;
  const fromEl = document.getElementById('filterFrom');
  const toEl = document.getElementById('filterTo');
  const now = new Date();
  switch (preset) {
    case 'all': fromEl.value = ''; toEl.value = ''; break;
    case 'thisMonth': {
      const y = now.getFullYear(), m = String(now.getMonth()+1).padStart(2,'0');
      fromEl.value = y+'-'+m+'-01';
      toEl.value = y+'-'+m+'-'+String(new Date(y,now.getMonth()+1,0).getDate()).padStart(2,'0');
      break;
    }
    case 'lastMonth': {
      const d = new Date(now.getFullYear(), now.getMonth()-1, 1);
      const y = d.getFullYear(), m = String(d.getMonth()+1).padStart(2,'0');
      fromEl.value = y+'-'+m+'-01';
      toEl.value = y+'-'+m+'-'+String(new Date(y,d.getMonth()+1,0).getDate()).padStart(2,'0');
      break;
    }
    case 'thisQuarter': {
      const qm = Math.floor(now.getMonth()/3)*3;
      const y = now.getFullYear();
      fromEl.value = y+'-'+String(qm+1).padStart(2,'0')+'-01';
      const lm = qm+2;
      toEl.value = y+'-'+String(lm+1).padStart(2,'0')+'-'+String(new Date(y,lm+1,0).getDate()).padStart(2,'0');
      break;
    }
  }
  applyFilters();
}

function calcTotals(arr) {
  const r2 = v => Math.round(v * 100) / 100;
  let subtotal=0, tax=0, total=0, shipping=0, discount=0;
  for (const s of arr) { subtotal+=s.subtotal; tax+=s.tax; total+=s.total; shipping+=s.shipping; discount+=s.discount; }
  return { subtotal: r2(subtotal), tax: r2(tax), total: r2(total), shipping: r2(shipping), discount: r2(discount), count: arr.length };
}

function buildProducts(salesArr) {
  const map = new Map();
  let unscrapedSub=0, unscrapedTax=0, unscrapedCount=0;
  for (const s of salesArr) {
    const items = ITEMS_CACHE[s.transactionNo];
    if (items && items.length > 0) {
      for (const item of items) {
        const price = item.price||0;
        const key = (item.name||'Unknown').toLowerCase().trim()+'|'+price.toFixed(2);
        if (!map.has(key)) map.set(key, { name: item.name||'Unknown', productNumber: item.productNumber, price, qty:0, revenue:0, tax:0 });
        const e = map.get(key);
        e.qty += item.quantity||1;
        e.revenue += item.amount||0;
        if (s.subtotal>0) e.tax += Math.round(s.tax * ((item.amount||0)/s.subtotal) * 100) / 100;
        if (item.productNumber && !e.productNumber) e.productNumber = item.productNumber;
      }
    } else {
      unscrapedSub += s.subtotal; unscrapedTax += s.tax; unscrapedCount++;
    }
  }
  const products = [...map.values()];
  products.sort((a,b) => b.qty - a.qty);
  if (unscrapedCount > 0) products.push({ name: 'Unscraped Orders ('+unscrapedCount+')', productNumber: null, price: 0, qty: unscrapedCount, revenue: unscrapedSub, tax: unscrapedTax, _unsorted: true });
  return products;
}

function sortProducts(products, col, dir) {
  const d = dir === 'asc' ? 1 : -1;
  return products.filter(p=>!p._unsorted).sort((a,b) => {
    switch(col) {
      case 'name': return d * a.name.localeCompare(b.name);
      case 'qty': return d * (a.qty - b.qty);
      case 'price': return d * (a.price - b.price);
      case 'revenue': return d * (a.revenue - b.revenue);
      case 'tax': return d * (a.tax - b.tax);
      default: return 0;
    }
  }).concat(products.filter(p=>p._unsorted));
}

function renderAll() {
  const t = calcTotals(filtered);
  document.getElementById('subtitle').textContent = filtered.length + ' transaction' + (filtered.length!==1?'s':'') + ' shown';

  // Cards
  let cardsHtml = '<div class="card"><div class="label">Pretax Total</div><div class="value pretax">'+fmt(t.subtotal)+'</div><div class="sub">'+t.count+' transactions</div></div>';
  cardsHtml += '<div class="card"><div class="label">Sales Tax</div><div class="value tax">'+fmt(t.tax)+'</div></div>';
  cardsHtml += '<div class="card"><div class="label">Post-Tax Total</div><div class="value total">'+fmt(t.total)+'</div></div>';
  if (t.discount > 0) cardsHtml += '<div class="card"><div class="label">Discounts</div><div class="value" style="color:#e67e22">-'+fmt(t.discount)+'</div></div>';
  if (t.shipping > 0) cardsHtml += '<div class="card"><div class="label">Shipping</div><div class="value" style="color:#6f42c1">'+fmt(t.shipping)+'</div></div>';
  document.getElementById('cards').innerHTML = cardsHtml;

  // Source breakdown
  const online = filtered.filter(s => s.source === 'online');
  const tradeshow = filtered.filter(s => s.source === 'tradeshow');
  let sbHtml = '';
  if (online.length > 0) { const ot = calcTotals(online); sbHtml += '<div class="source-card"><h3>Online Sales (Store Manager)</h3><div class="stat"><span class="label">Transactions</span><span>'+ot.count+'</span></div><div class="stat"><span class="label">Pretax</span><span>'+fmt(ot.subtotal)+'</span></div><div class="stat"><span class="label">Tax</span><span>'+fmt(ot.tax)+'</span></div><div class="stat"><span class="label">Total</span><span>'+fmt(ot.total)+'</span></div></div>'; }
  if (tradeshow.length > 0) { const tt = calcTotals(tradeshow); sbHtml += '<div class="source-card"><h3>Trade Show Sales (POS)</h3><div class="stat"><span class="label">Transactions</span><span>'+tt.count+'</span></div><div class="stat"><span class="label">Pretax</span><span>'+fmt(tt.subtotal)+'</span></div><div class="stat"><span class="label">Tax</span><span>'+fmt(tt.tax)+'</span></div><div class="stat"><span class="label">Total</span><span>'+fmt(tt.total)+'</span></div></div>'; }
  document.getElementById('sourceBreakdown').innerHTML = sbHtml;

  renderBySource(online, tradeshow);
  renderProductsTab();
  renderTransactionsTab();
}

function productTableHtml(products, idPrefix) {
  let totalQty=0, totalRev=0, totalTax=0;
  let rows = '';
  for (const p of products) {
    const pnum = p.productNumber ? '<br><span class="product-num">'+esc(p.productNumber)+'</span>' : '';
    rows += '<tr><td>'+esc(p.name)+pnum+'</td><td class="num">'+p.qty+'</td><td class="num">'+(p._unsorted?'':fmt(p.price))+'</td><td class="num">'+fmt(p.revenue)+'</td><td class="num">'+fmt(p.tax)+'</td></tr>';
    totalQty += p.qty; totalRev += p.revenue; totalTax += p.tax;
  }
  rows += '<tr class="total-row"><td>Total</td><td class="num">'+totalQty+'</td><td class="num"></td><td class="num">'+fmt(totalRev)+'</td><td class="num">'+fmt(totalTax)+'</td></tr>';
  return '<table><thead><tr><th>Product</th><th class="num">Qty Sold</th><th class="num">Unit Price</th><th class="num">Total Revenue</th><th class="num">Tax (est.)</th></tr></thead><tbody>'+rows+'</tbody></table>';
}

function txTableHtml(salesArr, showSource) {
  const sorted = [...salesArr].sort((a,b) => {
    const d = txSortDir === 'asc' ? 1 : -1;
    switch(txSortCol) {
      case 'date': return d * a.date.localeCompare(b.date);
      case 'total': return d * (a.total - b.total);
      case 'subtotal': return d * (a.subtotal - b.subtotal);
      case 'tax': return d * (a.tax - b.tax);
      case 'txno': return d * a.transactionNo.localeCompare(b.transactionNo);
      default: return 0;
    }
  });
  let rows = '';
  for (const s of sorted) {
    const srcBadge = showSource ? '<td><span class="source-badge '+s.source+'">'+(s.source==='online'?'Online':'Trade Show')+'</span></td>' : '';
    rows += '<tr><td>'+esc(s.transactionNo)+'</td><td>'+fmtDate(s.date)+'</td>'+srcBadge+'<td>'+esc(s.billingName)+'</td><td class="num">'+fmt(s.subtotal)+'</td><td class="num">'+fmt(s.tax)+'</td><td class="num">'+fmt(s.total)+'</td></tr>';
  }
  const srcHeader = showSource ? '<th>Source</th>' : '';
  return '<table><thead><tr><th>Tx No</th><th>Date</th>'+srcHeader+'<th>Customer</th><th class="num">Subtotal</th><th class="num">Tax</th><th class="num">Total</th></tr></thead><tbody>'+rows+'</tbody></table>';
}

function renderBySource(online, tradeshow) {
  let html = '<div style="padding-top:16px;">';

  // Product summaries first: Online then Trade Show
  if (online.length > 0) {
    const products = buildProducts(online);
    const ot = calcTotals(online);
    html += '<div class="source-section"><h3 class="online-header">Online (Store Manager) — '+ot.count+' transactions, '+fmt(ot.total)+' total</h3>';
    html += '<div class="section"><h2>Products <span class="count">('+products.length+')</span></h2>' + productTableHtml(products, 'online') + '</div>';
    html += '</div>';
  }
  if (tradeshow.length > 0) {
    const products = buildProducts(tradeshow);
    const tt = calcTotals(tradeshow);
    html += '<div class="source-section"><h3 class="tradeshow-header">Trade Show (POS) — '+tt.count+' transactions, '+fmt(tt.total)+' total</h3>';
    html += '<div class="section"><h2>Products <span class="count">('+products.length+')</span></h2>' + productTableHtml(products, 'tradeshow') + '</div>';
    html += '</div>';
  }

  // Transaction lists after: Online then Trade Show
  if (online.length > 0) {
    html += '<div class="source-section"><h3 class="online-header">Online Transactions</h3>';
    html += '<div class="section"><h2>Transactions <span class="count">('+online.length+')</span></h2>' + txTableHtml(online, false) + '</div>';
    html += '</div>';
  }
  if (tradeshow.length > 0) {
    html += '<div class="source-section"><h3 class="tradeshow-header">Trade Show Transactions</h3>';
    html += '<div class="section"><h2>Transactions <span class="count">('+tradeshow.length+')</span></h2>' + txTableHtml(tradeshow, false) + '</div>';
    html += '</div>';
  }

  if (filtered.length === 0) html += '<p style="text-align:center;color:#6c757d;padding:40px;">No transactions match filters</p>';
  html += '</div>';
  document.getElementById('tab-bySource').innerHTML = html;
}

function renderProductsTab() {
  const products = sortProducts(buildProducts(filtered), productSortCol, productSortDir);
  let html = '<div style="padding-top:16px;"><div class="section"><h2>All Products <span class="count">('+products.length+')</span></h2>';
  html += productTableHtml(products, 'all');
  html += '</div></div>';
  document.getElementById('tab-products').innerHTML = html;
}

function renderTransactionsTab() {
  let html = '<div style="padding-top:16px;"><div class="section"><h2>All Transactions <span class="count">('+filtered.length+')</span></h2>';
  html += txTableHtml(filtered, true);
  html += '</div></div>';
  document.getElementById('tab-transactions').innerHTML = html;
}

// Tab switching
document.getElementById('viewTabs').addEventListener('click', e => {
  const btn = e.target.closest('.tab-btn');
  if (!btn) return;
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById('tab-' + btn.dataset.tab).classList.add('active');
});

// Sortable table headers (delegation)
document.addEventListener('click', e => {
  const th = e.target.closest('th');
  if (!th) return;
  const text = th.textContent.trim().replace(/[\\u25B2\\u25BC]/g, '').trim().toLowerCase();
  const colMap = { 'product': 'name', 'qty sold': 'qty', 'unit price': 'price', 'total revenue': 'revenue', 'tax (est.)': 'tax',
    'tx no': 'txno', 'date': 'date', 'customer': null, 'source': null, 'subtotal': 'subtotal', 'total': 'total' };
  const col = colMap[text];
  if (!col) return;

  // Determine if product or transaction sort
  const table = th.closest('table');
  const isProduct = th.textContent.includes('Product') || th.textContent.includes('Qty') || th.textContent.includes('Unit Price') || th.textContent.includes('Revenue');
  const isTx = th.textContent.includes('Tx') || th.textContent.includes('Date') || th.textContent.includes('Subtotal');

  if (['name','qty','price','revenue'].includes(col) || (col === 'tax' && isProduct)) {
    if (productSortCol === col) productSortDir = productSortDir === 'desc' ? 'asc' : 'desc';
    else { productSortCol = col; productSortDir = col === 'name' ? 'asc' : 'desc'; }
  } else {
    if (txSortCol === col) txSortDir = txSortDir === 'desc' ? 'asc' : 'desc';
    else { txSortCol = col; txSortDir = col === 'txno' ? 'asc' : 'desc'; }
  }
  renderAll();
});

// Filter event listeners
document.getElementById('filterSource').addEventListener('change', applyFilters);
document.getElementById('filterFrom').addEventListener('change', applyFilters);
document.getElementById('filterTo').addEventListener('change', applyFilters);
document.getElementById('filterPreset').addEventListener('change', applyPreset);

// Init
applyFilters();
</script>
</body>
</html>`;
}

function getDateRange(sales) {
  const dates = sales.map(s => s.date.toISOString().split('T')[0]).sort();
  return { from: dates[0], to: dates[dates.length - 1] };
}
