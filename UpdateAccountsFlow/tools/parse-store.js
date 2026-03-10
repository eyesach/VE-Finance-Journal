import ExcelJS from 'exceljs';

/**
 * Parse Store Manager (online sales) Excel file.
 *
 * Known columns:
 * Transaction no | Date | Sales representative | Billing name | Billing company |
 * Billing address | Billing city | Billing state/province | Billing zip/postcode |
 * Billing country | Shipping name | Shipping company | ... | Subtotal | Tax |
 * Shipping | Discount | Promotion code | Total
 *
 * Returns Sale[] normalized objects.
 */
export async function parseStoreExcel(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const sheet = workbook.worksheets[0];
  if (!sheet) throw new Error('No worksheet found in Store Manager Excel');

  // Read header row to find column indices
  const headerRow = sheet.getRow(1);
  const headers = {};
  headerRow.eachCell((cell, colNumber) => {
    const val = String(cell.value || '').trim().toLowerCase();
    headers[val] = colNumber;
  });

  // Map known column names (case-insensitive)
  const cols = resolveColumns(headers);

  const sales = [];
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return; // skip header

    const transactionNo = getCellValue(row, cols.transactionNo);
    const dateRaw = getCellValue(row, cols.date);
    const billingName = getCellValue(row, cols.billingName);
    const subtotal = parseNumber(getCellValue(row, cols.subtotal));
    const tax = parseNumber(getCellValue(row, cols.tax));
    const shipping = parseNumber(getCellValue(row, cols.shipping));
    const discount = parseNumber(getCellValue(row, cols.discount));
    const total = parseNumber(getCellValue(row, cols.total));

    // Skip empty rows
    if (!transactionNo && !dateRaw && !total) return;

    const date = parseDate(dateRaw);
    if (!date) {
      console.warn(`  Warning: Skipping row ${rowNumber} — invalid date: ${dateRaw}`);
      return;
    }

    // Try to find a product/item description column
    const description = getCellValue(row, cols.description) || getCellValue(row, cols.billingName) || '';

    sales.push({
      transactionNo: String(transactionNo || ''),
      date,
      billingName: String(billingName || ''),
      description: String(description).trim(),
      subtotal: subtotal || 0,
      tax: tax || 0,
      shipping: shipping || 0,
      discount: Math.abs(discount || 0),
      total: total || 0,
      source: 'online',
    });
  });

  console.log(`  Parsed ${sales.length} online sales from Store Manager.`);

  // Parse line items from "Sales transaction items" sheet
  const itemsSheet = workbook.worksheets.find(s =>
    s.name.toLowerCase().includes('item')
  );
  const lineItems = new Map(); // transactionNo → LineItem[]
  if (itemsSheet) {
    const itemHeaders = {};
    itemsSheet.getRow(1).eachCell((cell, col) => {
      itemHeaders[String(cell.value || '').trim().toLowerCase()] = col;
    });

    const txCol = findCol(itemHeaders, ['transaction no', 'transaction_no', 'transactionno']);
    const nameCol = findCol(itemHeaders, ['item name', 'product name', 'name', 'item', 'product']);
    const numCol = findCol(itemHeaders, ['item number', 'product number', 'item no', 'product no']);
    const priceCol = findCol(itemHeaders, ['price', 'unit price']);
    const qtyCol = findCol(itemHeaders, ['quantity', 'qty']);
    const taxableCol = findCol(itemHeaders, ['taxable']);
    const amountCol = findCol(itemHeaders, ['amount', 'total', 'line total']);

    if (txCol && nameCol) {
      itemsSheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        const txNo = String(getCellValue(row, txCol) || '').trim();
        if (!txNo) return;

        const item = {
          name: String(getCellValue(row, nameCol) || 'Unknown').trim(),
          productNumber: numCol ? String(getCellValue(row, numCol) || '').trim() || null : null,
          price: parseNumber(getCellValue(row, priceCol)),
          quantity: parseInt(getCellValue(row, qtyCol) || '1', 10) || 1,
          taxable: taxableCol ? String(getCellValue(row, taxableCol) || '').toLowerCase() === 'yes' : false,
          amount: parseNumber(getCellValue(row, amountCol)),
        };

        if (!lineItems.has(txNo)) lineItems.set(txNo, []);
        lineItems.get(txNo).push(item);
      });
      console.log(`  Parsed ${lineItems.size} transactions with line items from Excel.`);
    }
  }

  return { sales, lineItems };
}

function resolveColumns(headers) {
  return {
    transactionNo: findCol(headers, ['transaction no', 'transaction_no', 'transactionno', 'trans no', 'order no', 'order number']),
    date: findCol(headers, ['date', 'order date', 'transaction date']),
    billingName: findCol(headers, ['billing name', 'billing_name', 'customer', 'name']),
    description: findCol(headers, ['description', 'item', 'product', 'item description', 'product name']),
    subtotal: findCol(headers, ['subtotal', 'sub total', 'sub-total', 'pretax']),
    tax: findCol(headers, ['tax', 'sales tax', 'tax amount']),
    shipping: findCol(headers, ['shipping', 'shipping cost', 'freight']),
    discount: findCol(headers, ['discount', 'discount amount']),
    total: findCol(headers, ['total', 'grand total', 'order total', 'amount']),
  };
}

function findCol(headers, candidates) {
  for (const name of candidates) {
    if (headers[name] !== undefined) return headers[name];
  }
  return null;
}

function getCellValue(row, colIndex) {
  if (!colIndex) return null;
  const cell = row.getCell(colIndex);
  if (cell.value === null || cell.value === undefined) return null;
  // Handle ExcelJS rich text, formulas, etc.
  if (typeof cell.value === 'object' && cell.value.result !== undefined) return cell.value.result;
  if (typeof cell.value === 'object' && cell.value.richText) {
    return cell.value.richText.map(r => r.text).join('');
  }
  return cell.value;
}

function parseNumber(val) {
  if (val === null || val === undefined) return 0;
  if (typeof val === 'number') return val;
  const cleaned = String(val).replace(/[$,\s]/g, '');
  const num = parseFloat(cleaned);
  return isNaN(num) ? 0 : num;
}

function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date) {
    return isNaN(val.getTime()) ? null : val;
  }
  const str = String(val).trim();
  // Try common formats: "03/06/2026", "03/06/2026 5:54 PM CT", "2026-03-06"
  const d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}
