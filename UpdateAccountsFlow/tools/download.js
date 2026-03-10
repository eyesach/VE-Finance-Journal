import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const TMP_DIR = path.join(__dirname, '..', '.tmp');

const MARKETPLACE_URL = 'https://portal.veinternational.org/portal';

/**
 * Set the date range filters on the current page and click Search.
 *
 * The VE site has two plain text inputs labeled "From (mm/dd/yyyy)" and
 * "until (mm/dd/yyyy)" followed by a "Search" button.
 * Dates must be in MM/DD/YYYY format.
 */
async function setDateRange(page, fromDate, toDate) {
  // Default to 1 year behind and 1 year ahead from today
  if (!fromDate && !toDate) {
    const now = new Date();
    const yearAgo = new Date(now.getFullYear() - 1, now.getMonth(), now.getDate());
    const yearAhead = new Date(now.getFullYear() + 1, now.getMonth(), now.getDate());
    fromDate = fmtDate(yearAgo);
    toDate = fmtDate(yearAhead);
  }

  console.log(`  Setting date range: ${fromDate} to ${toDate}...`);

  // The VE site uses plain <input type="text"> fields — grab all text inputs on the page
  // and use the first two (From and Until)
  const allInputs = page.locator('input[type="text"]');
  const count = await allInputs.count();

  if (count < 2) {
    console.log(`  Warning: Found ${count} text input(s), expected at least 2. Skipping date filter.`);
    return;
  }

  // Set date values directly via JS to avoid triggering datepicker popups
  await page.evaluate(({ from, to }) => {
    const inputs = document.querySelectorAll('input[type="text"]');
    if (inputs[0]) inputs[0].value = from;
    if (inputs[1]) inputs[1].value = to;
  }, { from: fromDate, to: toDate });
  await page.waitForTimeout(300);

  // Submit the search form directly via JS (click or form submit)
  const submitted = await page.evaluate(() => {
    // Try clicking the search button
    const btn = document.getElementById('id_searchbtn')
      || document.querySelector('button[type="submit"]')
      || document.querySelector('input[value="Search"]');
    if (btn) { btn.click(); return true; }
    // Try submitting the form
    const form = document.querySelector('form');
    if (form) { form.submit(); return true; }
    return false;
  });

  if (submitted) {
    await page.waitForLoadState('networkidle');
    await page.waitForTimeout(1500);
    console.log('  Date range set and search complete.');
  } else {
    console.log('  Warning: Could not find Search button or form. Proceeding without date filter.');
  }
}

function fmtDate(d) {
  return `${String(d.getMonth() + 1).padStart(2, '0')}/${String(d.getDate()).padStart(2, '0')}/${d.getFullYear()}`;
}

/**
 * Navigate back to Marketplace Tools safely (avoids page.goBack() frame detach errors).
 */
async function navigateToMarketplace(page) {
  try {
    await page.goto(MARKETPLACE_URL, { waitUntil: 'networkidle', timeout: 30000 });
  } catch {
    await page.goto('https://hub.veinternational.org/', { waitUntil: 'networkidle', timeout: 30000 });
    await page.locator('text=Marketplace Tools').first().click();
    await page.waitForLoadState('networkidle');
    await page.waitForTimeout(1000);
  }
}

/**
 * Click the Excel download link and wait for the download event.
 * Tries multiple possible link texts.
 */
async function clickDownloadLink(page, timeout = 30000) {
  const candidates = [
    'Download sales transaction list (Excel)',
    'Download sales transaction list',
    'Export',
    'Excel',
    'Download',
  ];

  for (const text of candidates) {
    const link = page.locator(`a:has-text("${text}"), button:has-text("${text}")`).first();
    const visible = await link.isVisible().catch(() => false);
    if (visible) {
      const [download] = await Promise.all([
        page.waitForEvent('download', { timeout }),
        link.click(),
      ]);
      return download;
    }
  }

  throw new Error('Could not find Excel download link on page');
}

/**
 * Download Store Manager (online sales) Excel from Marketplace Tools.
 */
async function downloadStoreManager(page, { fromDate, toDate } = {}) {
  console.log('  Opening Store Manager...');
  await page.locator('text=Store Manager').first().click();
  await page.waitForLoadState('networkidle');

  await setDateRange(page, fromDate, toDate);

  const download = await clickDownloadLink(page);
  const filePath = path.join(TMP_DIR, 'vei_checkout.xlsx');
  await download.saveAs(filePath);
  console.log(`  Store Manager Excel saved to ${filePath}`);

  await navigateToMarketplace(page);
  return filePath;
}

/**
 * Download Trade Show POS Excel from Marketplace Tools.
 */
async function downloadTradeShowPOS(page, { fromDate, toDate } = {}) {
  console.log('  Opening Trade Show POS...');
  await page.locator('text=Trade Show POS').first().click();
  await page.waitForLoadState('networkidle');

  console.log('  Clicking Sales tab...');
  await page.locator('text=Sales').first().click();
  await page.waitForLoadState('networkidle');

  await setDateRange(page, fromDate, toDate);

  const download = await clickDownloadLink(page);
  const filePath = path.join(TMP_DIR, 'tspos.xlsx');
  await download.saveAs(filePath);
  console.log(`  Trade Show POS Excel saved to ${filePath}`);

  await navigateToMarketplace(page);
  return filePath;
}

/**
 * Download both Excel files from Marketplace Tools page.
 * @param {Page} page - Playwright page on Marketplace Tools
 * @param {object} options - { fromDate, toDate } in 'MM/DD/YYYY' format
 * @returns {{ storeFile, tradeshowFile }} paths
 */
export async function downloadExcelFiles(page, { fromDate, toDate } = {}) {
  if (!fs.existsSync(TMP_DIR)) {
    fs.mkdirSync(TMP_DIR, { recursive: true });
  }

  const storeFile = await downloadStoreManager(page, { fromDate, toDate });
  const tradeshowFile = await downloadTradeShowPOS(page, { fromDate, toDate });

  return { storeFile, tradeshowFile };
}
