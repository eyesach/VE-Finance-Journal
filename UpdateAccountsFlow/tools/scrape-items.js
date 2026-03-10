import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const CACHE_FILE = path.join(__dirname, '..', '.tmp', 'scraped-items.json');
const CONCURRENCY = 10; // number of parallel tabs

/** Check if a transaction has real cached data (non-empty items or marked as no-items) */
function isCached(cache, txNo) {
  const items = cache.get(txNo);
  if (!items) return false;
  if (Array.isArray(items) && items.length > 0) return true;
  // { noItems: true } means we scraped it but found nothing — don't re-scrape
  if (items.noItems) return true;
  return false;
}

/**
 * Load cached scraped items from disk.
 * Returns Map<transactionNo, LineItem[]>
 */
export function loadCache() {
  if (!fs.existsSync(CACHE_FILE)) return new Map();
  try {
    const data = JSON.parse(fs.readFileSync(CACHE_FILE, 'utf-8'));
    return new Map(Object.entries(data));
  } catch {
    return new Map();
  }
}

export function saveCache(cache) {
  const obj = Object.fromEntries(cache);
  const dir = path.dirname(CACHE_FILE);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(CACHE_FILE, JSON.stringify(obj, null, 2), 'utf-8');
}

/**
 * Build a price→product lookup from the existing cache.
 * Only includes prices that consistently map to a single product.
 */
function buildPriceLookup(cache) {
  const lookup = new Map(); // price (string, 2 decimals) → { name, productNumber, ambiguous }

  for (const items of cache.values()) {
    if (!Array.isArray(items)) continue;
    for (const item of items) {
      if (!item.price || item.price <= 0) continue;
      const key = item.price.toFixed(2);
      if (lookup.has(key)) {
        const existing = lookup.get(key);
        if (existing.name.toLowerCase() !== (item.name || '').toLowerCase()) {
          existing.ambiguous = true;
        }
        existing.count++;
      } else {
        lookup.set(key, {
          name: item.name || 'Unknown',
          productNumber: item.productNumber || null,
          count: 1,
          ambiguous: false,
        });
      }
    }
  }

  // Remove ambiguous entries
  for (const [key, val] of lookup) {
    if (val.ambiguous) lookup.delete(key);
  }

  return lookup;
}

/**
 * Smart scrape: match single-item orders by price, return unmatched for full scraping.
 *
 * @param {Sale[]} sales - All sales to process
 * @param {Map} cache - Existing scraped items cache
 * @param {function} onProgress - Optional callback(msg) for progress updates
 * @returns {{ matched: Sale[], unmatched: Sale[] }}
 */
export function smartMatch(sales, cache, onProgress) {
  const priceLookup = buildPriceLookup(cache);
  const log = onProgress || (() => {});

  log(`Price lookup built: ${priceLookup.size} unique price→product mappings`);

  const unscraped = sales.filter(s => !isCached(cache, s.transactionNo));
  log(`${unscraped.length} unscraped transactions to process`);

  if (priceLookup.size === 0) {
    log('No cached products to match against — all need scraping');
    return { matched: [], unmatched: unscraped };
  }

  const matched = [];
  const unmatched = [];

  for (const sale of unscraped) {
    const priceKey = sale.subtotal.toFixed(2);
    const product = priceLookup.get(priceKey);

    if (product) {
      cache.set(sale.transactionNo, [{
        name: product.name,
        productNumber: product.productNumber,
        price: sale.subtotal,
        quantity: 1,
        taxable: sale.tax > 0,
        amount: sale.subtotal,
        inferred: true,
      }]);
      matched.push(sale);
    } else {
      unmatched.push(sale);
    }
  }

  log(`Smart match: ${matched.length} matched by price, ${unmatched.length} need scraping`);
  saveCache(cache);

  return { matched, unmatched };
}

/**
 * Scrape line items from transaction detail pages using parallel tabs.
 *
 * @param {Page} page - Playwright page
 * @param {Sale[]} sales - Sales to scrape
 * @param {object} options - { sample, scrapeAll, useCache, onProgress }
 * @returns {Map} cache
 */
export async function scrapeTransactionItems(page, sales, options = {}) {
  const { sample, scrapeAll = false, useCache = false, onProgress } = options;
  const cache = loadCache();
  const log = onProgress || console.log.bind(console);

  if (useCache && cache.size > 0) {
    log(`Using cached data (${cache.size} transactions).`);
    return cache;
  }

  let toScrape = sales.filter(s => !isCached(cache, s.transactionNo));

  if (!scrapeAll && sample) {
    toScrape = toScrape.slice(0, sample);
  } else if (!scrapeAll && !sample) {
    log('No scraping flags set. Use --sample N or --scrape-all to scrape product details.');
    return cache;
  }

  if (toScrape.length === 0) {
    log('All transactions already cached.');
    return cache;
  }

  log(`Scraping ${toScrape.length} transactions (${CONCURRENCY} parallel tabs)...`);

  const onlineSales = toScrape.filter(s => s.source === 'online');
  const tradeshowSales = toScrape.filter(s => s.source === 'tradeshow');

  if (onlineSales.length > 0) {
    await scrapeSourceWithInference(page, onlineSales, 'Store Manager', cache, log, tradeshowSales);
  }
  const remainingTradeshow = tradeshowSales.filter(s => !isCached(cache, s.transactionNo));
  if (remainingTradeshow.length > 0) {
    await scrapeSourceWithInference(page, remainingTradeshow, 'Trade Show POS', cache, log, []);
  }

  saveCache(cache);
  log(`Scraping complete. ${cache.size} transactions cached.`);
  return cache;
}

/**
 * Scrape specific transactions by their transaction numbers.
 * Used by the server's smart scrape flow for unmatched transactions only.
 *
 * After scraping each source, re-runs price matching on remaining unscraped
 * transactions using newly discovered products to minimize total scraping.
 */
export async function scrapeSpecificTransactions(page, sales, cache, onProgress) {
  const log = onProgress || console.log.bind(console);

  if (sales.length === 0) {
    log('No transactions to scrape.');
    return cache;
  }

  let remaining = [...sales];
  log(`Scraping ${remaining.length} transactions (${CONCURRENCY} parallel tabs)...`);

  const onlineSales = remaining.filter(s => s.source === 'online');
  const tradeshowSales = remaining.filter(s => s.source === 'tradeshow');

  if (onlineSales.length > 0) {
    await scrapeSourceWithInference(page, onlineSales, 'Store Manager', cache, log, tradeshowSales);
  }

  // Re-filter tradeshow sales — some may have been inferred after online scraping
  const remainingTradeshow = tradeshowSales.filter(s => !isCached(cache, s.transactionNo));
  if (remainingTradeshow.length > 0) {
    if (remainingTradeshow.length < tradeshowSales.length) {
      log(`Skipping ${tradeshowSales.length - remainingTradeshow.length} tradeshow transactions (inferred from online scraping)`);
    }
    await scrapeSourceWithInference(page, remainingTradeshow, 'Trade Show POS', cache, log, []);
  }

  saveCache(cache);
  const totalInferred = [...cache.values()].filter(items => Array.isArray(items) && items.some(i => i.inferred)).length;
  log(`Scraping complete. ${cache.size} transactions cached (${totalInferred} inferred).`);
  return cache;
}

/**
 * After each batch of scraping, check if newly discovered products can be
 * used to infer remaining unscraped single-item transactions by price.
 */
function inferFromNewProducts(cache, remainingTxNos, allSalesMap, log) {
  const priceLookup = buildPriceLookup(cache);
  let inferredCount = 0;

  for (const txNo of [...remainingTxNos]) {
    if (isCached(cache, txNo)) {
      remainingTxNos.delete(txNo);
      continue;
    }
    const sale = allSalesMap.get(txNo);
    if (!sale) continue;

    const priceKey = sale.subtotal.toFixed(2);
    const product = priceLookup.get(priceKey);
    if (product) {
      cache.set(txNo, [{
        name: product.name,
        productNumber: product.productNumber,
        price: sale.subtotal,
        quantity: 1,
        taxable: sale.tax > 0,
        amount: sale.subtotal,
        inferred: true,
      }]);
      remainingTxNos.delete(txNo);
      inferredCount++;
    }
  }

  return inferredCount;
}

/**
 * Scrape transactions for a specific source using parallel tabs.
 * After each batch, infers remaining unscraped transactions using new price data.
 * @param {Sale[]} otherSourceSales - Sales from the other source to also try inferring
 */
async function scrapeSourceWithInference(page, sales, sourceName, cache, log, otherSourceSales = []) {
  log(`Navigating to ${sourceName}...`);
  const context = page.context();

  // Navigate to Marketplace Tools via VE Hub
  const currentUrl = page.url();
  if (!currentUrl.includes('portal.veinternational.org/portal')) {
    if (currentUrl.includes('hub.veinternational.org')) {
      await page.locator('text=Marketplace Tools').first().click();
      await page.waitForLoadState('networkidle');
      await page.waitForTimeout(1000);
    } else {
      await page.goto('https://hub.veinternational.org/', { waitUntil: 'networkidle', timeout: 30000 });
      await page.locator('text=Marketplace Tools').first().click();
      await page.waitForLoadState('networkidle');
      await page.waitForTimeout(1000);
    }
  }

  // Click into the right tool
  if (sourceName === 'Trade Show POS') {
    await page.locator('text=Trade Show POS').first().click();
    await page.waitForLoadState('networkidle');
    await page.locator('text=Sales').first().click();
    await page.waitForLoadState('networkidle');
  } else {
    await page.locator('text=Store Manager').first().click();
    await page.waitForLoadState('networkidle');
  }

  // Set date range to 1 year back / 1 year forward to ensure all transactions are visible
  await setDateRange(page, log);

  // Build a set of transaction numbers we need
  const needed = new Set(sales.map(s => s.transactionNo));

  // Collect transaction links across all pages (pagination)
  log(`Collecting transaction links for ${sourceName}...`);
  const txLinks = new Map(); // transactionNo -> href

  let pageNum = 1;
  while (true) {
    // Scan current page for matching transaction links
    const links = page.locator('a');
    const linkCount = await links.count();

    for (let i = 0; i < linkCount; i++) {
      const linkEl = links.nth(i);
      const text = (await linkEl.textContent().catch(() => '')) || '';
      const txNo = text.trim();
      if (needed.has(txNo) && !txLinks.has(txNo)) {
        const href = await linkEl.getAttribute('href').catch(() => null);
        if (href) {
          const fullUrl = href.startsWith('http') ? href : `https://portal.veinternational.org${href}`;
          txLinks.set(txNo, fullUrl);
        }
      }
    }

    // Check if we found all needed links
    if (txLinks.size >= needed.size) break;

    // Try to go to next page
    const nextBtn = page.locator('a:has-text("Next"), a:has-text("next"), a:has-text("»"), button:has-text("Next")').first();
    const hasNext = await nextBtn.isVisible().catch(() => false);

    if (!hasNext) break;

    try {
      await nextBtn.click();
      await page.waitForLoadState('networkidle');
      await page.waitForTimeout(500);
      pageNum++;
      log(`  Page ${pageNum}... (${txLinks.size}/${needed.size} links found)`);
    } catch {
      break; // pagination click failed, stop
    }
  }

  log(`Found ${txLinks.size}/${sales.length} transaction links across ${pageNum} page(s).`);

  if (txLinks.size === 0) return;

  // Build a map of all sales (this source + other source) for inference lookups
  const allSalesMap = new Map();
  for (const s of sales) allSalesMap.set(s.transactionNo, s);
  for (const s of otherSourceSales) allSalesMap.set(s.transactionNo, s);

  // Track which transactions still need scraping (can shrink via inference)
  const remainingTxNos = new Set(sales.filter(s => !isCached(cache, s.transactionNo)).map(s => s.transactionNo));
  // Also track other-source transactions for cross-source inference
  const otherRemainingTxNos = new Set(otherSourceSales.filter(s => !isCached(cache, s.transactionNo)).map(s => s.transactionNo));

  // Filter entries to only those still needed
  let entries = [...txLinks.entries()].filter(([txNo]) => remainingTxNos.has(txNo));
  let completed = 0;
  const originalTotal = entries.length;

  for (let batchStart = 0; batchStart < entries.length; batchStart += CONCURRENCY) {
    const batch = entries.slice(batchStart, batchStart + CONCURRENCY)
      .filter(([txNo]) => !isCached(cache, txNo)); // skip if already inferred

    if (batch.length === 0) continue;

    const tasks = batch.map(async ([txNo, url]) => {
      const tab = await context.newPage();
      try {
        await tab.goto(url, { waitUntil: 'domcontentloaded', timeout: 15000 });
        const items = await extractLineItems(tab);
        cache.set(txNo, items.length > 0 ? items : { noItems: true });
        remainingTxNos.delete(txNo);
      } catch (err) {
        // Silently skip failures
      } finally {
        await tab.close();
      }
      completed++;
      log(`[${completed}/${originalTotal}] Scraping ${sourceName}...`);
    });

    await Promise.all(tasks);
    saveCache(cache);

    // After each batch, try to infer remaining transactions using newly discovered prices
    const inferred = inferFromNewProducts(cache, remainingTxNos, allSalesMap, log);
    const otherInferred = inferFromNewProducts(cache, otherRemainingTxNos, allSalesMap, log);
    if (inferred > 0 || otherInferred > 0) {
      log(`  → Inferred ${inferred} ${sourceName} + ${otherInferred} other-source transactions from new prices`);
      saveCache(cache);
      // Re-filter remaining entries to skip newly inferred ones
      entries = entries.filter(([txNo]) => !isCached(cache, txNo));
      // Adjust batchStart since we modified entries array
      batchStart = -CONCURRENCY; // restart from beginning of filtered list
      completed = originalTotal - entries.length;
    }
  }
}

/**
 * Set the date range on the current sales list page (1 year back / 1 year forward).
 * Uses JS injection to avoid datepicker popup issues.
 */
async function setDateRange(page, log) {
  const now = new Date();
  const yearAgo = new Date(now.getFullYear() - 1, now.getMonth(), now.getDate());
  const yearAhead = new Date(now.getFullYear() + 1, now.getMonth(), now.getDate());
  const fmt = (d) => `${String(d.getMonth() + 1).padStart(2, '0')}/${String(d.getDate()).padStart(2, '0')}/${d.getFullYear()}`;
  const fromDate = fmt(yearAgo);
  const toDate = fmt(yearAhead);

  const count = await page.locator('input[type="text"]').count();
  if (count < 2) return;

  log(`Setting date range: ${fromDate} to ${toDate}...`);

  await page.evaluate(({ from, to }) => {
    const inputs = document.querySelectorAll('input[type="text"]');
    if (inputs[0]) inputs[0].value = from;
    if (inputs[1]) inputs[1].value = to;
  }, { from: fromDate, to: toDate });
  await page.waitForTimeout(300);

  const submitted = await page.evaluate(() => {
    const btn = document.getElementById('id_searchbtn')
      || document.querySelector('button[type="submit"]')
      || document.querySelector('input[value="Search"]');
    if (btn) { btn.click(); return true; }
    const form = document.querySelector('form');
    if (form) { form.submit(); return true; }
    return false;
  });

  if (submitted) {
    await page.waitForLoadState('networkidle');
    await page.waitForTimeout(1500);
    log('Date range set and search complete.');
  }
}

/**
 * Extract line items from a transaction detail page.
 */
async function extractLineItems(page) {
  const items = [];

  const tables = page.locator('table');
  const tableCount = await tables.count();

  for (let t = 0; t < tableCount; t++) {
    const table = tables.nth(t);
    const headerText = await table.locator('th').first().textContent().catch(() => '');

    if (headerText && headerText.trim().toLowerCase() === 'item') {
      const rows = table.locator('tbody tr, tr').filter({ hasNot: page.locator('th') });
      const rowCount = await rows.count();

      let currentItem = null;

      for (let r = 0; r < rowCount; r++) {
        const row = rows.nth(r);
        const cells = row.locator('td');
        const cellCount = await cells.count();

        if (cellCount < 2) continue;

        const firstCell = (await cells.nth(0).textContent() || '').trim();

        // Product/Item number sub-row
        const numMatch = firstCell.match(/^(?:product|item)\s*number:\s*(.+)/i);
        if (numMatch && currentItem) {
          currentItem.productNumber = numMatch[1].trim();
          continue;
        }

        // Skip subtotal/tax/total rows
        if (cellCount >= 2) {
          const lastCellLabel = cellCount >= 4
            ? (await cells.nth(cellCount - 2).textContent() || '').trim().toLowerCase()
            : '';
          if (['subtotal', 'total'].includes(lastCellLabel) || lastCellLabel.includes('tax')) {
            continue;
          }
        }

        // Regular item row — handle both 5-col (Item,Price,Qty,Taxable,Amount) and 4-col (Item,Price,Qty,Amount)
        if (firstCell && cellCount >= 4) {
          let price, quantity, taxable, amount;
          if (cellCount >= 5) {
            price = parseFloat((await cells.nth(1).textContent() || '0').replace(/[$,]/g, '')) || 0;
            quantity = parseInt((await cells.nth(2).textContent() || '1'), 10) || 1;
            taxable = (await cells.nth(3).textContent() || '').trim().toLowerCase() === 'yes';
            amount = parseFloat((await cells.nth(4).textContent() || '0').replace(/[$,]/g, '')) || 0;
          } else {
            // 4-column layout: Item, Price, Quantity, Amount (no Taxable)
            price = parseFloat((await cells.nth(1).textContent() || '0').replace(/[$,]/g, '')) || 0;
            quantity = parseInt((await cells.nth(2).textContent() || '1'), 10) || 1;
            taxable = false;
            amount = parseFloat((await cells.nth(3).textContent() || '0').replace(/[$,]/g, '')) || 0;
          }

          let name = firstCell;
          let productNumber = null;
          const prodMatch = name.match(/(?:product|item)\s*number:\s*(\S+)/i);
          if (prodMatch) {
            productNumber = prodMatch[1];
            name = name.replace(/\s*(?:product|item)\s*number:\s*\S+/i, '').trim();
          }

          currentItem = { name, price, quantity, taxable, amount, productNumber };
          items.push(currentItem);
        }
      }

      break;
    }
  }

  return items;
}
