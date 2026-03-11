import express from 'express';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';
import yargs from 'yargs';
import { hideBin } from 'yargs/helpers';
import { parseStoreExcel } from './tools/parse-store.js';
import { parseTradeshowExcel } from './tools/parse-tradeshow.js';
import { loadCache, saveCache, smartMatch, scrapeSpecificTransactions } from './tools/scrape-items.js';
import { launchBrowser, authenticate, navigateToMarketplace } from './tools/auth.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const TMP_DIR = path.join(__dirname, '.tmp');
const PORT = 3000;

const argv = yargs(hideBin(process.argv))
  .option('skip-download', {
    type: 'boolean',
    describe: 'Skip download, reuse cached Excel files in .tmp/',
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
  .option('port', {
    type: 'number',
    describe: 'Port for the web server',
    default: PORT,
  })
  .help()
  .argv;

// Global state
let allSales = [];
let itemsCache = new Map();
let scrapeStatus = { state: 'idle', message: '' };
let sseClients = [];

function findFile(dir, prefix) {
  if (!fs.existsSync(dir)) return path.join(dir, prefix + '.xlsx');
  const files = fs.readdirSync(dir).filter(f => f.startsWith(prefix) && f.endsWith('.xlsx'));
  if (files.length === 0) return path.join(dir, prefix + '.xlsx');
  files.sort((a, b) => fs.statSync(path.join(dir, b)).mtimeMs - fs.statSync(path.join(dir, a)).mtimeMs);
  return path.join(dir, files[0]);
}

async function parseExcelFiles() {
  const storeFile = findFile(TMP_DIR, 'vei_checkout');
  const tradeshowFile = findFile(TMP_DIR, 'tspos');

  console.log('Parsing Excel files...');
  allSales = [];
  const excelLineItems = new Map();

  if (fs.existsSync(storeFile)) {
    const { sales, lineItems } = await parseStoreExcel(storeFile);
    allSales.push(...sales);
    for (const [txNo, items] of lineItems) excelLineItems.set(txNo, items);
  }
  if (fs.existsSync(tradeshowFile)) {
    const { sales, lineItems } = await parseTradeshowExcel(tradeshowFile);
    allSales.push(...sales);
    for (const [txNo, items] of lineItems) excelLineItems.set(txNo, items);
  }

  console.log(`Total: ${allSales.length} transactions loaded.`);

  // Load existing scrape cache, then overlay Excel line items as the primary source
  itemsCache = loadCache();
  let excelOverrides = 0;
  for (const [txNo, items] of excelLineItems) {
    // Excel data is authoritative — overwrite inferred/scraped data
    if (!itemsCache.has(txNo) || !Array.isArray(itemsCache.get(txNo))) {
      excelOverrides++;
    }
    itemsCache.set(txNo, items);
  }
  saveCache(itemsCache);
  console.log(`Cache: ${itemsCache.size} transactions with product data (${excelLineItems.size} from Excel).`);
}

async function downloadFreshExcel({ fromDate, toDate } = {}) {
  const { downloadExcelFiles } = await import('./tools/download.js');
  const { context, page } = await launchBrowser({
    headed: argv.headed || true, // always headed for download (may need 2FA)
    login: argv.login,
  });
  const authPage = await authenticate({ page, context, login: argv.login });
  await navigateToMarketplace(authPage);
  await downloadExcelFiles(authPage, { fromDate, toDate });
  await context.close();
}

async function loadData() {
  console.log('\n=== VE Sales Dashboard ===\n');

  if (!argv.skipDownload) {
    console.log('Downloading Excel files...');
    await downloadFreshExcel();
    console.log('');
  }

  await parseExcelFiles();
  console.log('');
}

function sendSSE(event, data) {
  const msg = `event: ${event}\ndata: ${JSON.stringify(data)}\n\n`;
  for (const res of sseClients) {
    res.write(msg);
  }
}

function serializeSales(sales) {
  return sales.map(s => ({
    ...s,
    date: s.date.toISOString(),
  }));
}

function serializeCache(cache) {
  return Object.fromEntries(cache);
}

// Express app
const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// API: Get all sales data
app.get('/api/sales', (req, res) => {
  res.json(serializeSales(allSales));
});

// API: Get scraped items cache
app.get('/api/cache', (req, res) => {
  res.json(serializeCache(itemsCache));
});

// API: Product groups config
const GROUPS_FILE = path.join(TMP_DIR, 'product-groups.json');

function loadGroups() {
  if (!fs.existsSync(GROUPS_FILE)) return [];
  try { return JSON.parse(fs.readFileSync(GROUPS_FILE, 'utf-8')); } catch { return []; }
}

function saveGroups(groups) {
  if (!fs.existsSync(TMP_DIR)) fs.mkdirSync(TMP_DIR, { recursive: true });
  fs.writeFileSync(GROUPS_FILE, JSON.stringify(groups, null, 2));
}

app.get('/api/product-groups', (req, res) => {
  res.json(loadGroups());
});

app.put('/api/product-groups', (req, res) => {
  const groups = req.body;
  if (!Array.isArray(groups)) return res.status(400).json({ error: 'Expected array' });
  saveGroups(groups);
  res.json({ success: true });
});

// API: Export all data as JSON for import into main Accounting Journal app
app.get('/api/export', (req, res) => {
  try {
    const exportData = {
      exportVersion: 1,
      exportDate: new Date().toISOString(),
      companyName: 'VE International',
      sales: serializeSales(allSales),
      lineItems: serializeCache(itemsCache),
    };
    const exportFile = path.join(TMP_DIR, 've-sales-export.json');
    if (!fs.existsSync(TMP_DIR)) fs.mkdirSync(TMP_DIR, { recursive: true });
    fs.writeFileSync(exportFile, JSON.stringify(exportData, null, 2));
    res.setHeader('Content-Disposition', 'attachment; filename="ve-sales-export.json"');
    res.setHeader('Content-Type', 'application/json');
    res.json(exportData);
  } catch (err) {
    console.error('Export error:', err);
    res.status(500).json({ error: err.message });
  }
});

// API: Refresh Excel files (download fresh + re-parse)
let refreshStatus = { state: 'idle' };
app.get('/api/refresh', async (req, res) => {
  if (refreshStatus.state === 'running') {
    return res.status(409).json({ error: 'Refresh already in progress' });
  }
  refreshStatus = { state: 'running' };
  const { fromDate, toDate } = req.query;
  console.log(`Refreshing Excel files...${fromDate || toDate ? ` (${fromDate || 'start'} to ${toDate || 'now'})` : ''}`);
  try {
    await downloadFreshExcel({ fromDate, toDate });
    await parseExcelFiles();
    refreshStatus = { state: 'idle' };
    console.log('Refresh complete.\n');
    res.json({ success: true, totalSales: allSales.length });
  } catch (err) {
    refreshStatus = { state: 'idle' };
    console.error('Refresh error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// API: Get scrape status
app.get('/api/scrape/status', (req, res) => {
  res.json(scrapeStatus);
});

// API: Smart scrape via SSE
app.get('/api/scrape/start', async (req, res) => {
  // Set up SSE
  res.writeHead(200, {
    'Content-Type': 'text/event-stream',
    'Cache-Control': 'no-cache',
    'Connection': 'keep-alive',
  });
  sseClients.push(res);

  req.on('close', () => {
    sseClients = sseClients.filter(c => c !== res);
  });

  if (scrapeStatus.state === 'running') {
    res.write(`event: error\ndata: ${JSON.stringify({ message: 'Scrape already in progress' })}\n\n`);
    return;
  }

  scrapeStatus = { state: 'running', message: 'Starting smart scrape...' };

  // Send keepalive pings every 15s to prevent browser/proxy timeout
  const keepalive = setInterval(() => {
    res.write(': keepalive\n\n');
  }, 15000);

  const send = (event, data) => {
    res.write(`event: ${event}\ndata: ${JSON.stringify(data)}\n\n`);
  };

  const log = (msg) => {
    send('progress', { message: msg });
    console.log(`  [scrape] ${msg}`);
  };

  try {
    // Step 1: Smart match by price
    log('Step 1: Matching single-item orders by price...');
    const { matched, unmatched } = smartMatch(allSales, itemsCache, log);
    send('match-complete', { matched: matched.length, unmatched: unmatched.length });

    if (unmatched.length === 0) {
      log('All transactions matched! No scraping needed.');
      scrapeStatus = { state: 'idle', message: 'Complete — all matched by price' };
      send('complete', { totalCached: itemsCache.size });
      return;
    }

    // Step 2: Scrape remaining unmatched transactions
    log(`Step 2: Scraping ${unmatched.length} unmatched transactions...`);
    send('scrape-start', { count: unmatched.length });

    const { context, page } = await launchBrowser({
      headed: true, // always headed for scraping (needs 2FA login)
      login: argv.login,
    });

    let authPage;
    try {
      authPage = await authenticate({ page, context, login: argv.login });
      await navigateToMarketplace(authPage);
    } catch (err) {
      log(`Auth error: ${err.message}. Try restarting with --login --headed`);
      await context.close();
      scrapeStatus = { state: 'idle', message: 'Auth failed' };
      send('error', { message: 'Authentication failed. Restart with --login --headed' });
      return;
    }

    await scrapeSpecificTransactions(authPage, unmatched, itemsCache, log);
    await context.close();

    scrapeStatus = { state: 'idle', message: `Complete — ${itemsCache.size} transactions cached` };
    send('complete', { totalCached: itemsCache.size });
    log('Smart scrape complete!');
  } catch (err) {
    console.error('Scrape error:', err);
    scrapeStatus = { state: 'idle', message: `Error: ${err.message}` };
    send('error', { message: err.message });
  } finally {
    clearInterval(keepalive);
  }
});

// Start
async function main() {
  await loadData();

  const port = argv.port || PORT;
  app.listen(port, () => {
    console.log(`Dashboard running at http://localhost:${port}\n`);
  });

  // Auto-open in browser
  try {
    const open = (await import('open')).default;
    await open(`http://localhost:${port}`);
  } catch {
    // ignore
  }
}

main().catch(err => {
  console.error('Error:', err.message);
  process.exit(1);
});
