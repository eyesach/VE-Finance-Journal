import ExcelJS from 'exceljs';

/**
 * Generate an Excel workbook from sales data, grouped by month.
 *
 * @param {Sale[]} sales - Filtered sales array (date is a Date object)
 * @param {Map} itemsCache - Map of transactionNo → line items
 * @returns {Promise<Buffer>} Excel file buffer
 */
export async function generateExcelReport(sales, itemsCache) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'VE Sales Dashboard';
  wb.created = new Date();

  const currencyFmt = '$#,##0.00';
  const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4A90A4' } };
  const headerFont = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
  const boldFont = { bold: true };

  // Group sales by month
  const monthMap = new Map();
  for (const s of sales) {
    const key = `${s.date.getFullYear()}-${String(s.date.getMonth() + 1).padStart(2, '0')}`;
    if (!monthMap.has(key)) monthMap.set(key, []);
    monthMap.get(key).push(s);
  }
  const months = [...monthMap.keys()].sort();

  // --- Summary Sheet ---
  const summary = wb.addWorksheet('Summary', { tabColor: { argb: 'FF4A90A4' } });

  // Title
  summary.mergeCells('A1:F1');
  const titleCell = summary.getCell('A1');
  titleCell.value = 'VE Sales Report';
  titleCell.font = { bold: true, size: 16, color: { argb: 'FF4A90A4' } };

  // Date range
  summary.mergeCells('A2:F2');
  const dateRange = sales.length > 0
    ? formatDateRange(sales)
    : 'No data';
  summary.getCell('A2').value = dateRange;
  summary.getCell('A2').font = { size: 11, color: { argb: 'FF6C757D' } };

  // Overall totals
  const totals = calculateTotals(sales);
  const onlineTotals = calculateTotals(sales.filter(s => s.source === 'online'));
  const tradeshowTotals = calculateTotals(sales.filter(s => s.source === 'tradeshow'));

  let row = 4;
  summary.getCell(`A${row}`).value = 'Overall Summary';
  summary.getCell(`A${row}`).font = { bold: true, size: 13 };
  row++;

  const summaryData = [
    ['Transactions', totals.count],
    ['Pretax Total', totals.subtotal],
    ['Sales Tax', totals.tax],
    ['Post-Tax Total', totals.total],
  ];
  if (totals.discount > 0) summaryData.push(['Discounts', totals.discount]);
  if (totals.shipping > 0) summaryData.push(['Shipping', totals.shipping]);

  for (const [label, value] of summaryData) {
    summary.getCell(`A${row}`).value = label;
    summary.getCell(`A${row}`).font = boldFont;
    const valCell = summary.getCell(`B${row}`);
    valCell.value = value;
    if (typeof value === 'number' && label !== 'Transactions') {
      valCell.numFmt = currencyFmt;
    }
    row++;
  }

  // Source breakdown
  row += 1;
  if (onlineTotals.count > 0 || tradeshowTotals.count > 0) {
    summary.getCell(`A${row}`).value = 'Source Breakdown';
    summary.getCell(`A${row}`).font = { bold: true, size: 13 };
    row++;

    const sourceHeaders = ['Source', 'Transactions', 'Pretax', 'Tax', 'Total'];
    addHeaderRow(summary, row, sourceHeaders, headerFill, headerFont);
    row++;

    if (onlineTotals.count > 0) {
      addDataRow(summary, row, ['Online (Store Manager)', onlineTotals.count, onlineTotals.subtotal, onlineTotals.tax, onlineTotals.total], [null, null, currencyFmt, currencyFmt, currencyFmt]);
      row++;
    }
    if (tradeshowTotals.count > 0) {
      addDataRow(summary, row, ['Trade Show (POS)', tradeshowTotals.count, tradeshowTotals.subtotal, tradeshowTotals.tax, tradeshowTotals.total], [null, null, currencyFmt, currencyFmt, currencyFmt]);
      row++;
    }
    addDataRow(summary, row, ['Total', totals.count, totals.subtotal, totals.tax, totals.total], [null, null, currencyFmt, currencyFmt, currencyFmt]);
    summary.getRow(row).font = boldFont;
    row++;
  }

  // Monthly breakdown table
  if (months.length > 1) {
    row += 1;
    summary.getCell(`A${row}`).value = 'Monthly Breakdown';
    summary.getCell(`A${row}`).font = { bold: true, size: 13 };
    row++;

    const monthHeaders = ['Month', 'Source', 'Transactions', 'Pretax', 'Tax', 'Total'];
    addHeaderRow(summary, row, monthHeaders, headerFill, headerFont);
    row++;

    for (const m of months) {
      const mSales = monthMap.get(m);
      const mOnline = mSales.filter(s => s.source === 'online');
      const mTradeshow = mSales.filter(s => s.source === 'tradeshow');
      const mt = calculateTotals(mSales);

      if (mOnline.length > 0) {
        const ot = calculateTotals(mOnline);
        addDataRow(summary, row, [formatMonthLabel(m), 'Online', ot.count, ot.subtotal, ot.tax, ot.total], [null, null, null, currencyFmt, currencyFmt, currencyFmt]);
        row++;
      }
      if (mTradeshow.length > 0) {
        const tt = calculateTotals(mTradeshow);
        addDataRow(summary, row, [mOnline.length > 0 ? '' : formatMonthLabel(m), 'Trade Show', tt.count, tt.subtotal, tt.tax, tt.total], [null, null, null, currencyFmt, currencyFmt, currencyFmt]);
        row++;
      }
      // Month total row
      addDataRow(summary, row, [mOnline.length > 0 || mTradeshow.length > 0 ? '' : formatMonthLabel(m), 'Total', mt.count, mt.subtotal, mt.tax, mt.total], [null, null, null, currencyFmt, currencyFmt, currencyFmt]);
      summary.getRow(row).font = boldFont;
      row++;
    }
    // Grand total
    addDataRow(summary, row, ['Grand Total', '', totals.count, totals.subtotal, totals.tax, totals.total], [null, null, null, currencyFmt, currencyFmt, currencyFmt]);
    summary.getRow(row).font = boldFont;
  }

  // Overall product breakdown on summary sheet, split by source
  row += 2;
  summary.getCell(`A${row}`).value = 'Product Breakdown';
  summary.getCell(`A${row}`).font = { bold: true, size: 13 };
  row++;

  const prodSources = [];
  const allOnline = sales.filter(s => s.source === 'online');
  const allTradeshow = sales.filter(s => s.source === 'tradeshow');
  if (allOnline.length > 0) prodSources.push({ label: 'Online (Store Manager)', sales: allOnline });
  if (allTradeshow.length > 0) prodSources.push({ label: 'Trade Show (POS)', sales: allTradeshow });

  for (const src of prodSources) {
    summary.getCell(`A${row}`).value = src.label;
    summary.getCell(`A${row}`).font = { bold: true, size: 11 };
    row++;

    const prodHeaders = ['Product', 'Product #', 'Qty Sold', 'Unit Price', 'Total Revenue', 'Tax (est.)'];
    addHeaderRow(summary, row, prodHeaders, headerFill, headerFont);
    row++;

    const products = groupByProduct(src.sales, itemsCache);
    let totalQty = 0, totalRev = 0, totalTax = 0;
    for (const p of products) {
      const qty = p.description === 'Unscraped Orders' ? p.count : p.totalQuantity;
      const price = p.description === 'Unscraped Orders' ? null : (p.priceConsistent ? p.unitPrice : null);
      addDataRow(summary, row, [p.description, p.productNumber || '', qty, price, p.totalRevenue, p.totalTax], [null, null, null, currencyFmt, currencyFmt, currencyFmt]);
      totalQty += qty; totalRev += p.totalRevenue; totalTax += p.totalTax;
      row++;
    }
    addDataRow(summary, row, ['Total', '', totalQty, null, totalRev, totalTax], [null, null, null, null, currencyFmt, currencyFmt]);
    summary.getRow(row).font = boldFont;
    row += 2;
  }

  autoWidth(summary);

  // --- Per-Month Sheets ---
  const onlineHeaderFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF28A745' } };
  const tradeshowHeaderFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0D6EFD' } };
  const sectionFont = { bold: true, size: 12 };

  for (const m of months) {
    const mSales = monthMap.get(m);
    const sheetName = formatMonthLabel(m);
    const ws = wb.addWorksheet(sheetName);

    const mt = calculateTotals(mSales);
    const onlineSales = mSales.filter(s => s.source === 'online');
    const tradeshowSales = mSales.filter(s => s.source === 'tradeshow');
    let r = 1;

    // Month title + summary
    ws.mergeCells(`A${r}:G${r}`);
    ws.getCell(`A${r}`).value = sheetName;
    ws.getCell(`A${r}`).font = { bold: true, size: 14, color: { argb: 'FF4A90A4' } };
    r++;

    ws.getCell(`A${r}`).value = `${mt.count} transactions — Pretax: ${fmtCurrency(mt.subtotal)}, Tax: ${fmtCurrency(mt.tax)}, Total: ${fmtCurrency(mt.total)}`;
    ws.getCell(`A${r}`).font = { color: { argb: 'FF6C757D' } };
    r += 2;

    // Render each source that has data
    const sources = [];
    if (onlineSales.length > 0) sources.push({ label: 'Online (Store Manager)', sales: onlineSales, fill: onlineHeaderFill });
    if (tradeshowSales.length > 0) sources.push({ label: 'Trade Show (POS)', sales: tradeshowSales, fill: tradeshowHeaderFill });

    for (const src of sources) {
      const st = calculateTotals(src.sales);

      // Source header
      ws.getCell(`A${r}`).value = src.label;
      ws.getCell(`A${r}`).font = { bold: true, size: 13, color: { argb: 'FF333333' } };
      r++;
      ws.getCell(`A${r}`).value = `${st.count} transactions — Pretax: ${fmtCurrency(st.subtotal)}, Tax: ${fmtCurrency(st.tax)}, Total: ${fmtCurrency(st.total)}`;
      ws.getCell(`A${r}`).font = { size: 10, color: { argb: 'FF6C757D' } };
      r += 2;

      // Product breakdown for this source
      ws.getCell(`A${r}`).value = 'Products';
      ws.getCell(`A${r}`).font = sectionFont;
      r++;

      const prodHeaders = ['Product', 'Product #', 'Qty Sold', 'Unit Price', 'Total Revenue', 'Tax (est.)'];
      addHeaderRow(ws, r, prodHeaders, src.fill, headerFont);
      r++;

      const products = groupByProduct(src.sales, itemsCache);
      let totalQty = 0, totalRev = 0, totalTax = 0;

      for (const p of products) {
        const qty = p.description === 'Unscraped Orders' ? p.count : p.totalQuantity;
        const price = p.description === 'Unscraped Orders' ? null : (p.priceConsistent ? p.unitPrice : null);
        addDataRow(ws, r, [
          p.description,
          p.productNumber || '',
          qty,
          price,
          p.totalRevenue,
          p.totalTax,
        ], [null, null, null, currencyFmt, currencyFmt, currencyFmt]);
        if (p.description === 'Unscraped Orders') {
          ws.getCell(`A${r}`).font = { italic: true, color: { argb: 'FF6C757D' } };
        }
        totalQty += qty;
        totalRev += p.totalRevenue;
        totalTax += p.totalTax;
        r++;
      }

      addDataRow(ws, r, ['Total', '', totalQty, null, totalRev, totalTax], [null, null, null, null, currencyFmt, currencyFmt]);
      ws.getRow(r).font = boldFont;
      r += 2;

      // Transactions for this source
      ws.getCell(`A${r}`).value = 'Transactions';
      ws.getCell(`A${r}`).font = sectionFont;
      r++;

      const txHeaders = ['Tx No', 'Date', 'Customer', 'Subtotal', 'Tax', 'Total'];
      addHeaderRow(ws, r, txHeaders, src.fill, headerFont);
      r++;

      const sorted = [...src.sales].sort((a, b) => a.date - b.date);
      for (const s of sorted) {
        addDataRow(ws, r, [
          s.transactionNo,
          s.date,
          s.billingName || '',
          s.subtotal,
          s.tax,
          s.total,
        ], [null, 'mm/dd/yyyy', null, currencyFmt, currencyFmt, currencyFmt]);
        r++;
      }

      r += 2; // Gap before next source
    }

    autoWidth(ws);
  }

  return await wb.xlsx.writeBuffer();
}

// --- Helpers ---

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

function groupByProduct(sales, itemsCache) {
  const map = new Map();
  let unscrapedSubtotal = 0, unscrapedTax = 0, unscrapedCount = 0;

  for (const s of sales) {
    const items = itemsCache.get(s.transactionNo);
    if (items && items.length > 0 && !items.noItems) {
      for (const item of items) {
        const key = (item.name || 'Unknown Item').toLowerCase().trim();
        if (!map.has(key)) {
          map.set(key, {
            description: item.name || 'Unknown Item',
            productNumber: item.productNumber || null,
            count: 0, totalQuantity: 0, totalRevenue: 0, totalTax: 0, prices: [],
          });
        }
        const entry = map.get(key);
        entry.count++;
        entry.totalQuantity += item.quantity || 1;
        entry.totalRevenue += item.amount || 0;
        if (s.subtotal > 0) {
          entry.totalTax += s.tax * (item.amount / s.subtotal);
        }
        entry.prices.push(item.price || 0);
        if (item.productNumber && !entry.productNumber) entry.productNumber = item.productNumber;
      }
    } else if (!items || !items.noItems) {
      unscrapedSubtotal += s.subtotal;
      unscrapedTax += s.tax;
      unscrapedCount++;
    }
  }

  const products = [...map.values()].map(p => ({
    ...p,
    unitPrice: p.prices.length > 0 ? p.prices[0] : 0,
    priceConsistent: p.prices.every(pr => Math.abs(pr - p.prices[0]) < 0.01),
  }));
  products.sort((a, b) => b.totalQuantity - a.totalQuantity);

  if (unscrapedCount > 0) {
    products.push({
      description: 'Unscraped Orders',
      productNumber: null,
      count: unscrapedCount, totalQuantity: unscrapedCount,
      totalRevenue: unscrapedSubtotal, totalTax: unscrapedTax,
      unitPrice: 0, priceConsistent: false, prices: [],
    });
  }

  return products;
}

function fmtCurrency(n) {
  return '$' + n.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function formatMonthLabel(yyyyMm) {
  const [y, m] = yyyyMm.split('-');
  const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return `${monthNames[parseInt(m, 10) - 1]} ${y}`;
}

function formatDateRange(sales) {
  const dates = sales.map(s => s.date).sort((a, b) => a - b);
  const first = dates[0];
  const last = dates[dates.length - 1];
  const fmt = d => d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
  return `${fmt(first)} – ${fmt(last)} (${sales.length} transactions)`;
}

function addHeaderRow(ws, rowNum, values, fill, font) {
  for (let i = 0; i < values.length; i++) {
    const cell = ws.getCell(rowNum, i + 1);
    cell.value = values[i];
    cell.fill = fill;
    cell.font = font;
  }
}

function addDataRow(ws, rowNum, values, formats) {
  for (let i = 0; i < values.length; i++) {
    if (values[i] === null || values[i] === undefined) continue;
    const cell = ws.getCell(rowNum, i + 1);
    cell.value = values[i];
    if (formats && formats[i]) cell.numFmt = formats[i];
  }
}

function autoWidth(ws) {
  ws.columns.forEach(col => {
    let maxLen = 10;
    col.eachCell({ includeEmpty: false }, cell => {
      const len = cell.value ? String(cell.value).length : 0;
      if (len > maxLen) maxLen = len;
    });
    col.width = Math.min(maxLen + 3, 40);
  });
}
