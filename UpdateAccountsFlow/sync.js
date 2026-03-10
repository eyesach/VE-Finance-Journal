import yargs from 'yargs';
import { hideBin } from 'yargs/helpers';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';
import { launchBrowser, authenticate, navigateToMarketplace } from './tools/auth.js';
import { downloadExcelFiles } from './tools/download.js';
import { parseStoreExcel } from './tools/parse-store.js';
import { parseTradeshowExcel } from './tools/parse-tradeshow.js';
import { scrapeTransactionItems, loadCache, saveCache } from './tools/scrape-items.js';
import { generateReport } from './tools/report.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const TMP_DIR = path.join(__dirname, '.tmp');

const argv = yargs(hideBin(process.argv))
  .option('from', {
    type: 'string',
    describe: 'Start date (YYYY-MM-DD)',
  })
  .option('to', {
    type: 'string',
    describe: 'End date (YYYY-MM-DD)',
  })
  .option('month', {
    type: 'string',
    describe: 'Filter to a specific month (YYYY-MM)',
  })
  .option('all', {
    type: 'boolean',
    describe: 'Include all transactions (no date filter)',
    default: false,
  })
  .option('monthly', {
    type: 'boolean',
    describe: 'Show per-month breakdown in report',
    default: false,
  })
  .option('source', {
    type: 'string',
    describe: 'Filter by source: online, tradeshow, or both',
    default: 'both',
    choices: ['online', 'tradeshow', 'both'],
  })
  .option('sample', {
    type: 'number',
    describe: 'Scrape only first N transaction detail pages',
  })
  .option('scrape-all', {
    type: 'boolean',
    describe: 'Scrape all transaction detail pages for product info',
    default: false,
  })
  .option('use-cache', {
    type: 'boolean',
    describe: 'Use previously scraped product data from cache',
    default: false,
  })
  .option('login', {
    type: 'boolean',
    describe: 'Force manual login (opens browser for 2FA)',
    default: false,
  })
  .option('headed', {
    type: 'boolean',
    describe: 'Run browser in headed mode (visible)',
    default: false,
  })
  .option('skip-download', {
    type: 'boolean',
    describe: 'Skip download, reuse cached Excel files in .tmp/',
    default: false,
  })
  .help()
  .argv;

async function main() {
  console.log('\n=== VE Sales Report ===\n');

  const needsBrowser = !argv.skipDownload || argv.sample || argv.scrapeAll;
  const needsScraping = argv.sample || argv.scrapeAll;
  const sourceLabel = argv.source === 'both' ? 'Online + Trade Show' :
    argv.source === 'online' ? 'Online Sales' : 'Trade Show Sales';
  const totalSteps = needsScraping ? 5 : 4;
  let step = 1;

  // --- Download Phase ---
  let storeFile = findFile(TMP_DIR, 'vei_checkout');
  let tradeshowFile = findFile(TMP_DIR, 'tspos');
  let browserContext = null;
  let browserPage = null;

  if (!argv.skipDownload) {
    console.log(`[${step}/${totalSteps}] Downloading Excel files...`);
    const { context, page } = await launchBrowser({
      headed: argv.headed,
      login: argv.login,
    });
    browserContext = context;
    browserPage = await authenticate({ page, context, login: argv.login });
    await navigateToMarketplace(browserPage);
    const files = await downloadExcelFiles(browserPage);
    storeFile = files.storeFile;
    tradeshowFile = files.tradeshowFile;

    // Keep browser open if we need to scrape
    if (!needsScraping) {
      await context.close();
      browserContext = null;
      browserPage = null;
    }
    console.log('');
  } else {
    console.log(`[${step}/${totalSteps}] Skipping download — using cached files.`);
    if (argv.source !== 'tradeshow' && !fs.existsSync(storeFile)) {
      console.error(`  Error: ${storeFile} not found. Run without --skip-download first.`);
      process.exit(1);
    }
    if (argv.source !== 'online' && !fs.existsSync(tradeshowFile)) {
      console.error(`  Error: ${tradeshowFile} not found. Run without --skip-download first.`);
      process.exit(1);
    }
    console.log('');
  }
  step++;

  // --- Parse Phase ---
  console.log(`[${step}/${totalSteps}] Parsing Excel files...`);
  let allSales = [];
  const excelLineItems = new Map();
  if (argv.source !== 'tradeshow') {
    const { sales, lineItems } = await parseStoreExcel(storeFile);
    allSales.push(...sales);
    for (const [txNo, items] of lineItems) excelLineItems.set(txNo, items);
  }
  if (argv.source !== 'online') {
    const { sales, lineItems } = await parseTradeshowExcel(tradeshowFile);
    allSales.push(...sales);
    for (const [txNo, items] of lineItems) excelLineItems.set(txNo, items);
  }
  console.log(`  Total: ${allSales.length} transactions (${sourceLabel}).\n`);
  step++;

  // --- Populate cache from Excel line items ---
  let itemsCache = loadCache();
  if (excelLineItems.size > 0) {
    for (const [txNo, items] of excelLineItems) {
      itemsCache.set(txNo, items);
    }
    saveCache(itemsCache);
    console.log(`  Loaded ${excelLineItems.size} transactions with product data from Excel.\n`);
  }

  // --- Scrape Phase (optional, for any remaining unmatched) ---
  if (needsScraping) {
    console.log(`[${step}/${totalSteps}] Scraping transaction details for product info...`);

    // Launch browser if not already open
    if (!browserPage) {
      const { context, page } = await launchBrowser({
        headed: argv.headed,
        login: argv.login,
      });
      browserContext = context;
      browserPage = await authenticate({ page, context, login: argv.login });
    }

    itemsCache = await scrapeTransactionItems(browserPage, allSales, {
      sample: argv.sample,
      scrapeAll: argv.scrapeAll,
      useCache: argv.useCache,
    });

    await browserContext.close();
    browserContext = null;
    browserPage = null;
    console.log('');
    step++;
  } else if (argv.useCache) {
    console.log(`[${step}/${totalSteps}] Loading cached product data...`);
    itemsCache = loadCache();
    console.log(`  ${itemsCache.size} transactions in cache.\n`);
    step++;
  }

  // Close browser if still open
  if (browserContext) {
    await browserContext.close();
  }

  // --- Filter Phase ---
  console.log(`[${step}/${totalSteps}] Filtering by date...`);
  const { filtered, dateLabel } = filterByDate(allSales);
  console.log(`  ${filtered.length} transactions in range: ${dateLabel}\n`);
  step++;

  if (filtered.length === 0) {
    console.log('  No transactions found for this date range. Exiting.');
    process.exit(0);
  }

  // --- Report Phase ---
  console.log(`[${step}/${totalSteps}] Generating report...`);
  const reportPath = generateReport(filtered, {
    dateLabel,
    sourceLabel,
    showMonthly: argv.monthly || !!argv.month,
    itemsCache,
  });

  // Auto-open in browser
  try {
    const open = (await import('open')).default;
    await open(reportPath);
    console.log('  Report opened in browser.\n');
  } catch {
    console.log(`  Open the report manually: ${reportPath}\n`);
  }

  console.log('=== Done ===\n');
}

function filterByDate(sales) {
  let fromDate = null;
  let toDate = null;
  let dateLabel = '';

  if (argv.all) {
    dateLabel = 'All time';
    return { filtered: sales, dateLabel };
  }

  if (argv.month) {
    const [year, month] = argv.month.split('-').map(Number);
    fromDate = new Date(year, month - 1, 1);
    toDate = new Date(year, month, 0, 23, 59, 59, 999);
    const monthName = fromDate.toLocaleString('en-US', { month: 'long', year: 'numeric' });
    dateLabel = monthName;
  } else if (argv.from || argv.to) {
    if (argv.from) {
      fromDate = new Date(argv.from + 'T00:00:00');
    }
    if (argv.to) {
      toDate = new Date(argv.to + 'T23:59:59.999');
    }
    const fromStr = fromDate ? formatDate(fromDate) : 'beginning';
    const toStr = toDate ? formatDate(toDate) : 'present';
    dateLabel = `${fromStr} to ${toStr}`;
  } else {
    const now = new Date();
    fromDate = new Date(now.getFullYear(), now.getMonth(), 1);
    toDate = new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59, 999);
    const monthName = fromDate.toLocaleString('en-US', { month: 'long', year: 'numeric' });
    dateLabel = monthName;
  }

  const filtered = sales.filter(s => {
    if (fromDate && s.date < fromDate) return false;
    if (toDate && s.date > toDate) return false;
    return true;
  });

  return { filtered, dateLabel };
}

function findFile(dir, prefix) {
  if (!fs.existsSync(dir)) return path.join(dir, prefix + '.xlsx');
  const files = fs.readdirSync(dir).filter(f => f.startsWith(prefix) && f.endsWith('.xlsx'));
  if (files.length === 0) return path.join(dir, prefix + '.xlsx');
  files.sort((a, b) => fs.statSync(path.join(dir, b)).mtimeMs - fs.statSync(path.join(dir, a)).mtimeMs);
  return path.join(dir, files[0]);
}

function formatDate(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

main().catch(err => {
  console.error('\nError:', err.message);
  process.exit(1);
});
