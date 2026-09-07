// Theme tokens used by buildPageHtml() to style the generated static pages.
//
// These values are intentionally NOT hardcoded design decisions — they are meant to be
// captured directly from the live, authenticated SharePoint site via
// scripts/capture-sharepoint-theme.mjs, which writes scripts/sharepoint-theme.json.
// If that file exists, its values override the fallback defaults below, so re-running the
// capture script after a SharePoint theme/branding change is enough to keep the static site
// visually in sync — no manual CSS editing required.
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
export const THEME_FILE_PATH = path.join(__dirname, 'sharepoint-theme.json');

// Fallback values, used until a real capture has been run (or if a specific token wasn't
// found during capture). These approximate the current Grove Guidance branding by eye.
export const DEFAULT_THEME = {
  capturedAt: null,
  capturedFrom: null,
  fontFamily: 'Segoe UI, Arial, sans-serif',
  bodyTextColor: '#1b1b1b',
  themePrimary: '#0f4f8c',
  themeSecondary: '#0f4f5c',
  themeDark: '#333333',
  neutralPrimary: '#1b1b1b',
  neutralSecondary: '#444444',
  neutralLight: '#f3f2f1',
  headerBackground: '#f4f6f8',
  headerBorder: '#d5d9de',
  cardCornerRadius: '0px',
  cardOverlayGradient: 'linear-gradient(to top, rgba(0,0,0,.55), rgba(0,0,0,0) 70%)',
  buttonBackground: '#0f4f5c',
  buttonTextColor: '#ffffff',
  buttonCornerRadius: '2px',
};

let cachedTheme = null;

/**
 * Loads theme tokens, merging any captured overrides on top of the defaults.
 * Safe to call even if sharepoint-theme.json has never been generated.
 */
export async function loadTheme() {
  if (cachedTheme) return cachedTheme;

  let overrides = {};
  try {
    const raw = await readFile(THEME_FILE_PATH, 'utf8');
    overrides = JSON.parse(raw);
  } catch (error) {
    if (error?.code !== 'ENOENT') {
      console.warn(`Warning: failed to read ${THEME_FILE_PATH}: ${error.message}. Using default theme.`);
    }
  }

  cachedTheme = { ...DEFAULT_THEME, ...overrides };
  return cachedTheme;
}

/** Test-only helper to force loadTheme() to re-read the file on the next call. */
export function resetThemeCache() {
  cachedTheme = null;
}
