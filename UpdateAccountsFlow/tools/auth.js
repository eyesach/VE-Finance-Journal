import { chromium } from 'playwright';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const STATE_DIR = path.join(__dirname, '..', 'playwright-state');

const CLEVER_URL = 'https://clever.com/oauth/authorize?channel=clever&client_id=4c63c1cf623dce82caac&confirmed=true&redirect_uri=https%3A%2F%2Fclever.com%2Fin%2Fauth_callback';

/**
 * Launch a persistent browser context.
 * On first run or --login, opens headed so user can complete Google SSO + 2FA.
 * On subsequent runs, reuses saved session cookies (headless).
 */
export async function launchBrowser({ headed = false, login = false }) {
  const headless = !headed && !login;

  const context = await chromium.launchPersistentContext(STATE_DIR, {
    headless,
    viewport: { width: 1280, height: 800 },
    args: ['--disable-blink-features=AutomationControlled'],
  });

  const page = context.pages()[0] || await context.newPage();
  return { context, page };
}

/**
 * Ensure we're authenticated. After login, navigate to VE Hub.
 * Returns a page on the VE Hub dashboard (hub.veinternational.org).
 */
export async function authenticate({ page, context, login = false }) {
  // Navigate to Clever
  await page.goto(CLEVER_URL, { waitUntil: 'domcontentloaded', timeout: 30000 });

  // Wait for navigation to settle — could land on Clever dashboard, VE Hub, or login page
  await page.waitForTimeout(3000);

  const currentUrl = page.url();

  // Already on VE Hub (session still valid from a previous run)
  if (currentUrl.includes('hub.veinternational.org')) {
    if (!login) {
      console.log('  Session valid — already on VE Hub.');
      return page;
    }
  }

  // Already on Clever dashboard (authenticated but not yet on VE Hub)
  if (currentUrl.includes('clever.com') && !currentUrl.includes('oauth/authorize')) {
    const isLoggedIn = await checkLoggedIn(page);
    if (isLoggedIn && !login) {
      console.log('  Session valid — on Clever dashboard. Navigating to VE Hub...');
      const hubPage = await navigateToVEHub(page);
      return hubPage;
    }
  }

  // Need manual login
  console.log('\n  Browser opened — please log in through Google SSO + 2FA.');
  console.log('  Waiting for you to reach the Clever dashboard or VE Hub...\n');

  // Wait for user to complete login (up to 5 minutes)
  // After SSO, they may land on Clever dashboard, VE Hub, or identity callback
  try {
    await page.waitForURL(
      url => {
        const u = typeof url === 'string' ? url : url.toString();
        return u.includes('clever.com/in/') ||
               u.includes('hub.veinternational.org') ||
               u.includes('identity.veinternational.org') ||
               u.includes('portal.veinternational.org');
      },
      { timeout: 300000 }
    );
    await page.waitForLoadState('networkidle');
    await page.waitForTimeout(3000);

    // Follow any remaining redirects (identity callback → hub)
    const postLoginUrl = page.url();
    if (postLoginUrl.includes('identity.veinternational.org')) {
      await page.waitForURL(url => {
        const u = typeof url === 'string' ? url : url.toString();
        return u.includes('hub.veinternational.org') || u.includes('clever.com/in/');
      }, { timeout: 30000 });
      await page.waitForLoadState('networkidle');
      await page.waitForTimeout(2000);
    }
  } catch (e) {
    // Log current URL for debugging
    console.log(`  Current URL: ${page.url()}`);
    throw new Error('Login timed out after 5 minutes. Please try again.');
  }

  console.log('  Login successful!');

  // If we're on Clever, navigate to VE Hub
  if (page.url().includes('clever.com')) {
    const hubPage = await navigateToVEHub(page);
    return hubPage;
  }

  return page;
}

async function checkLoggedIn(page) {
  try {
    await page.waitForSelector('text=Favorite resources', { timeout: 5000 });
    return true;
  } catch {
    return false;
  }
}

/**
 * From the Clever dashboard, click VE Hub to get to hub.veinternational.org.
 * VE Hub may open in a new tab or redirect in the same tab.
 */
async function navigateToVEHub(page) {
  console.log('  Navigating to VE Hub...');
  const context = page.context();

  // Listen for new pages (popup/new tab)
  const pagePromise = context.waitForEvent('page', { timeout: 15000 }).catch(() => null);

  const veHubLink = page.locator('text=VE Hub').first();
  await veHubLink.click();

  // Check if a new tab opened
  const newPage = await pagePromise;
  if (newPage) {
    await newPage.waitForLoadState('networkidle');
    console.log('  VE Hub loaded (new tab).');
    return newPage;
  }

  // Otherwise it redirected in the same tab
  await page.waitForURL('**/hub.veinternational.org/**', { timeout: 30000 });
  await page.waitForLoadState('networkidle');
  console.log('  VE Hub loaded.');
  return page;
}

/**
 * From VE Hub, navigate to Marketplace Tools.
 * Returns a page on the Marketplace Tools page (portal.veinternational.org).
 */
export async function navigateToMarketplace(page) {
  // If already on Marketplace Tools, skip
  if (page.url().includes('portal.veinternational.org/portal')) {
    console.log('  Already on Marketplace Tools.');
    return page;
  }

  // If on VE Hub, click Marketplace Tools
  console.log('  Navigating to Marketplace Tools...');
  const marketplaceLink = page.locator('text=Marketplace Tools').first();
  await marketplaceLink.click();
  await page.waitForLoadState('networkidle');
  await page.waitForTimeout(2000);
  console.log('  Marketplace Tools loaded.');

  return page;
}
