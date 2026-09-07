#!/usr/bin/env node
// Captures real SharePoint theme values (colors, fonts, spacing/radii, button style) from a
// live, authenticated page and writes them to scripts/sharepoint-theme.json.
//
// buildPageHtml() in sharepoint-publish.mjs reads that file (via sharepoint-theme.mjs) and uses
// the captured values instead of hand-picked defaults, so when SharePoint's theme/branding
// changes, re-running this script is enough to keep the static site visually in sync -- no
// manual CSS editing required.
//
// Usage:
//   node scripts/capture-sharepoint-theme.mjs [pageUrl]
//
// The first run opens a real Chromium window and pauses for you to sign in interactively
// (Microsoft login, MFA, etc.). Once signed in, press Enter in the terminal to continue; your
// session is saved to .cache/sharepoint-auth-state.json and reused (silently, headless) on
// subsequent runs until it expires, at which point the script will prompt you to sign in again.
import { mkdir, stat, writeFile } from 'node:fs/promises';
import path from 'node:path';
import readline from 'node:readline/promises';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const AUTH_STATE_PATH = path.join(__dirname, '..', '.cache', 'sharepoint-auth-state.json');
const THEME_OUTPUT_PATH = path.join(__dirname, 'sharepoint-theme.json');

const DEFAULT_TENANT_HOST = process.env.SP_TENANT_HOST || 'uebt.sharepoint.com';
const DEFAULT_SITE_PATH = process.env.SP_SITE_PATH || '/sites/GroveGuidance';
const DEFAULT_PAGE_URL = `https://${DEFAULT_TENANT_HOST}${DEFAULT_SITE_PATH}/SitePages/Home.aspx`;

async function fileExists(filePath) {
  try {
    await stat(filePath);
    return true;
  } catch {
    return false;
  }
}

async function main() {
  const pageUrl = process.argv[2] || DEFAULT_PAGE_URL;
  const hasStoredState = await fileExists(AUTH_STATE_PATH);

  let browser;
  let context;

  if (hasStoredState) {
    console.log(`Reusing saved sign-in from ${AUTH_STATE_PATH} (headless)...`);
    browser = await chromium.launch({ headless: true });
    context = await browser.newContext({ storageState: AUTH_STATE_PATH });
  } else {
    console.log('No saved SharePoint sign-in found. Opening a visible browser window...');
    browser = await chromium.launch({ headless: false });
    context = await browser.newContext();
  }

  const page = await context.newPage();
  await page.goto(pageUrl, { waitUntil: 'domcontentloaded', timeout: 60000 });

  // If we're bounced to a Microsoft login page, either the saved session expired or this is
  // the first run. Wait here for the user to sign in interactively.
  const isLoginPage = () => /login\.microsoftonline\.com|\/_forms\/default\.aspx/i.test(page.url());
  if (isLoginPage()) {
    if (hasStoredState) {
      console.log('Saved sign-in appears to have expired. Re-opening a visible browser window to sign in again...');
      await browser.close();
      browser = await chromium.launch({ headless: false });
      context = await browser.newContext();
      const retryPage = await context.newPage();
      await retryPage.goto(pageUrl, { waitUntil: 'domcontentloaded', timeout: 60000 });
      await waitForManualSignIn(retryPage, pageUrl);
      await context.storageState({ path: AUTH_STATE_PATH });
      await captureFromPage(retryPage);
      await browser.close();
      return;
    }
    await waitForManualSignIn(page, pageUrl);
    await mkdir(path.dirname(AUTH_STATE_PATH), { recursive: true });
    await context.storageState({ path: AUTH_STATE_PATH });
    console.log(`Saved sign-in to ${AUTH_STATE_PATH} for future headless runs.`);
  }

  await captureFromPage(page);
  await browser.close();
}

async function waitForManualSignIn(page, pageUrl) {
  console.log('Please sign in to SharePoint in the opened browser window (including any MFA prompt).');
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  await rl.question('Press Enter here once you have signed in and can see the SharePoint page... ');
  rl.close();
  // Give the SPA a moment to finish rendering the real page content after redirect.
  await page.goto(pageUrl, { waitUntil: 'networkidle', timeout: 60000 }).catch(() => {});
  await page.waitForTimeout(2000);
}

async function captureFromPage(page) {
  console.log('Waiting for SharePoint page content to render...');
  await page.waitForSelector('[data-automation-id="pageContentContainer"], .SPPageChrome, main', { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(1500);

  const captured = await page.evaluate(() => {
    function computed(selector, prop) {
      const el = document.querySelector(selector);
      if (!el) return null;
      const value = getComputedStyle(el)[prop];
      return value || null;
    }

    function firstNonNull(...values) {
      return values.find((v) => v !== null && v !== undefined && v !== '') ?? null;
    }

    // SharePoint's Fluent UI theme is exposed as CSS custom properties on the root/body in
    // modern communication-site themes (e.g. --themePrimary, --neutralLight). Read those
    // directly when present -- they are the most reliable source of the real brand colors.
    const rootStyle = getComputedStyle(document.documentElement);
    const bodyStyle = getComputedStyle(document.body);
    const cssVar = (name) => {
      const value = rootStyle.getPropertyValue(name) || bodyStyle.getPropertyValue(name);
      return value ? value.trim() : null;
    };

    // Hero/card image tiles: SharePoint renders these as buttons/anchors with a background
    // image and an overlaid caption. Try a few known automation-id / class selectors.
    const cardSelector = '[data-automation-id="heroCardContainer"], .heroCard, [class*="heroCard"]';
    const overlaySelector = '[class*="textOverlay"], [class*="titleOverlay"], [class*="descriptionOverlay"]';
    const buttonSelector = '.ms-Button--primary, [data-automation-id="callToActionButton"], button[class*="primaryButton"], a[class*="primaryButton"]';
    const headerSelector = '[data-automation-id="pageHeader"], .SPPageChrome-header, header';

    const cardEl = document.querySelector(cardSelector);
    const cardStyle = cardEl ? getComputedStyle(cardEl) : null;

    return {
      themePrimary: firstNonNull(cssVar('--themePrimary'), computed('a', 'color')),
      themeSecondary: firstNonNull(cssVar('--themeSecondary'), computed(buttonSelector, 'backgroundColor')),
      themeDark: firstNonNull(cssVar('--themeDark'), cardStyle ? cardStyle.backgroundColor : null),
      neutralPrimary: firstNonNull(cssVar('--neutralPrimary'), computed('body', 'color')),
      neutralSecondary: cssVar('--neutralSecondary'),
      neutralLight: firstNonNull(cssVar('--neutralLight'), computed('body', 'backgroundColor')),
      headerBackground: computed(headerSelector, 'backgroundColor'),
      headerBorder: computed(headerSelector, 'borderBottomColor'),
      fontFamily: computed('body', 'fontFamily'),
      bodyTextColor: computed('body', 'color'),
      cardCornerRadius: cardStyle ? cardStyle.borderRadius : null,
      buttonBackground: computed(buttonSelector, 'backgroundColor'),
      buttonTextColor: computed(buttonSelector, 'color'),
      buttonCornerRadius: computed(buttonSelector, 'borderRadius'),
    };
  });

  const result = {
    capturedAt: new Date().toISOString(),
    capturedFrom: page.url(),
    ...captured,
  };

  // Drop null values so buildPageHtml()'s defaults fill in anything we couldn't find
  // (SharePoint's DOM/class names can change; we prefer a safe fallback over a broken value).
  for (const [key, value] of Object.entries(result)) {
    if (value === null || value === undefined) delete result[key];
  }

  await writeFile(THEME_OUTPUT_PATH, `${JSON.stringify(result, null, 2)}\n`, 'utf8');
  console.log(`Captured theme written to ${THEME_OUTPUT_PATH}:`);
  console.log(JSON.stringify(result, null, 2));
}

main().catch((error) => {
  console.error('Theme capture failed:', error);
  process.exitCode = 1;
});
