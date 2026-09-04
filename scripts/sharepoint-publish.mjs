#!/usr/bin/env node
import { createHash } from 'node:crypto';
import { mkdir, readFile, rm, stat, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

const DEFAULT_SP_TENANT_HOST = 'uebt.sharepoint.com';
const DEFAULT_SP_SITE_PATH = '/sites/GroveGuidance';
const MAX_RETRIES = 3;
const RETRY_DELAY_MS = 2000;

function normalizeSharePointSitePath(sitePathValue) {
  let sitePath = (sitePathValue || '').trim();
  if (!sitePath.startsWith('/')) {
    sitePath = `/${sitePath}`;
  }
  sitePath = sitePath.replace(/\/+$/, '') || '/';
  sitePath = sitePath.split(/\/(?:_api|_layouts)\b/i)[0] || '/';

  const segments = sitePath.split('/').filter(Boolean);
  while (segments.length > 0) {
    const lastSegment = segments[segments.length - 1];
    if (/\.aspx$/i.test(lastSegment) || /^forms$/i.test(lastSegment)) {
      segments.pop();
      continue;
    }
    if (/^sitepages$/i.test(lastSegment) || /^pages$/i.test(lastSegment)) {
      segments.pop();
    }
    break;
  }

  return segments.length > 0 ? `/${segments.join('/')}` : '/';
}

export function resolveSharePointLocation(tenantHostValue, sitePathValue) {
  let tenantHost = (tenantHostValue || '').trim();
  let derivedSitePath = '';

  if (/^https?:\/\//i.test(tenantHost)) {
    try {
      const parsed = new URL(tenantHost);
      tenantHost = parsed.host;
      derivedSitePath = parsed.pathname;
    } catch {
      tenantHost = '';
    }
  } else {
    const slashIndex = tenantHost.indexOf('/');
    if (slashIndex >= 0) {
      derivedSitePath = tenantHost.slice(slashIndex);
      tenantHost = tenantHost.slice(0, slashIndex);
    }
  }

  tenantHost = tenantHost.replace(/\/+$/, '').toLowerCase() || DEFAULT_SP_TENANT_HOST;

  let sitePath = (sitePathValue || '').trim();
  if (/^https?:\/\//i.test(sitePath)) {
    try {
      const parsed = new URL(sitePath);
      if (parsed.host) {
        tenantHost = parsed.host.toLowerCase();
      }
      sitePath = parsed.pathname;
    } catch {
      sitePath = '';
    }
  }

  sitePath = normalizeSharePointSitePath(sitePath || derivedSitePath || DEFAULT_SP_SITE_PATH);

  return { tenantHost, sitePath };
}

const resolvedSharePointLocation = resolveSharePointLocation(process.env.SP_TENANT_HOST, process.env.SP_SITE_PATH);

const config = {
  tenantId: process.env.SP_TENANT_ID,
  clientId: process.env.SP_CLIENT_ID,
  clientSecret: process.env.SP_CLIENT_SECRET,
  tenantHost: resolvedSharePointLocation.tenantHost,
  sitePath: resolvedSharePointLocation.sitePath,
  publicBaseUrl: (process.env.PUBLIC_BASE_URL || 'https://groveguidance.uebt.org').replace(/\/$/, ''),
  outputDir: path.resolve(process.cwd(), process.env.OUTPUT_DIR || 'dist'),
  stateFile: path.resolve(process.cwd(), process.env.STATE_FILE || '.cache/sharepoint-sync-state.json'),
  pilotPageLimit: process.env.PILOT_PAGE_LIMIT ? Number(process.env.PILOT_PAGE_LIMIT) : undefined,
};

function assertSyncEnvironment() {
  const requiredEnv = ['SP_TENANT_ID', 'SP_CLIENT_ID', 'SP_CLIENT_SECRET'];
  const missing = requiredEnv.filter((key) => !process.env[key]);
  if (missing.length > 0) {
    throw new Error(`Missing required environment variables: ${missing.join(', ')}`);
  }
}

const NAV_QUERIES = [
  `${config.sitePath}/_api/web/Navigation/QuickLaunch?$select=Title,Url`,
  `${config.sitePath}/_api/web/Navigation/TopNavigationBar?$select=Title,Url`,
];

async function exists(filePath) {
  try {
    await stat(filePath);
    return true;
  } catch {
    return false;
  }
}

async function ensureDir(dirPath) {
  await mkdir(dirPath, { recursive: true });
}

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function htmlEscape(value = '') {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function slugify(value = '') {
  return value
    .trim()
    .toLowerCase()
    .replace(/\.aspx$/i, '')
    .replace(/[^a-z0-9\-/]+/g, '-')
    .replace(/-+/g, '-')
    .replace(/^[-/]+|[-/]+$/g, '');
}

function routeFromPage(page) {
  const fileRef = page.FileRef || '';
  const markerIndex = fileRef.toLowerCase().indexOf('/sitepages/');
  const relativeRef = markerIndex >= 0 ? fileRef.slice(markerIndex + '/sitepages/'.length) : page.FileLeafRef || '';
  const clean = slugify(relativeRef);
  if (!clean || clean === 'home') {
    return '/';
  }
  return `/${clean}/`;
}

function outputFileFromRoute(route) {
  if (route === '/') {
    return path.join(config.outputDir, 'index.html');
  }
  const relative = route.replace(/^\//, '').replace(/\/$/, '');
  return path.join(config.outputDir, relative, 'index.html');
}

function canonicalFromRoute(route) {
  if (route === '/') {
    return `${config.publicBaseUrl}/`;
  }
  return `${config.publicBaseUrl}${route}`;
}

function extractBannerUrl(page) {
  const raw = page.BannerImageUrl;
  if (!raw) return undefined;
  if (typeof raw === 'string') {
    try {
      const parsed = JSON.parse(raw);
      return parsed.serverRelativeUrl || parsed.Url || parsed.url || raw;
    } catch {
      return raw;
    }
  }
  if (typeof raw === 'object') {
    return raw.serverRelativeUrl || raw.Url || raw.url;
  }
  return undefined;
}

async function getAccessToken() {
  const tokenUrl = `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type: 'client_credentials',
    client_id: config.clientId,
    client_secret: config.clientSecret,
    scope: `https://${config.tenantHost}/.default`,
  });
  const response = await fetch(tokenUrl, {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body,
  });
  if (!response.ok) {
    const errorBody = await response.text();
    console.error(`Token endpoint error (${response.status} ${response.statusText}):`, errorBody);
    throw new Error(`Could not obtain access token (${response.status} ${response.statusText})`);
  }
  const json = await response.json();
  return json.access_token;
}

async function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function sharePointRequest(token, relativeUrl, responseType = 'json', attempt = 1) {
  const url = relativeUrl.startsWith('http') ? relativeUrl : `https://${config.tenantHost}${relativeUrl}`;
  
  try {
    const response = await fetch(url, {
      headers: {
        authorization: 'Bearer ' + token,
        accept: 'application/json;odata=nometadata',
      },
    });
    
    if (!response.ok) {
      const errorBody = await response.text();
      console.error(`SharePoint request failed (${response.status} ${response.statusText}) for ${url}`);
      console.error(`Response body: ${errorBody.substring(0, 500)}`);
      
      // Retry on transient errors (500, 502, 503, 504)
      if ([500, 502, 503, 504].includes(response.status) && attempt < MAX_RETRIES) {
        console.log(`Retrying SharePoint request (attempt ${attempt + 1}/${MAX_RETRIES}) after ${RETRY_DELAY_MS}ms...`);
        await sleep(RETRY_DELAY_MS);
        return sharePointRequest(token, relativeUrl, responseType, attempt + 1);
      }
      
      throw new Error(`SharePoint request failed (${response.status} ${response.statusText}) for ${url}`);
    }
    
    if (responseType === 'buffer') {
      return Buffer.from(await response.arrayBuffer());
    }
    return response.json();
  } catch (error) {
    // Retry on network errors
    if (attempt < MAX_RETRIES && error.message && !error.message.includes('401')) {
      console.log(`Network error, retrying (attempt ${attempt + 1}/${MAX_RETRIES}):`, error.message);
      await sleep(RETRY_DELAY_MS);
      return sharePointRequest(token, relativeUrl, responseType, attempt + 1);
    }
    throw error;
  }
}

async function sharePointList(token, relativeUrl) {
  const values = [];
  let next = relativeUrl;
  while (next) {
    const json = await sharePointRequest(token, next);
    const batch = json.value || [];
    values.push(...batch);
    next = json['@odata.nextLink'] || json['odata.nextLink'] || undefined;
  }
  return values;
}

async function getPublishedPages(token) {
  const query = `${config.sitePath}/_api/web/lists/getbytitle('Site Pages')/items?` +
    '$select=Id,Title,FileLeafRef,FileRef,Modified,Created,FirstPublishedDate,Description,CanvasContent1,BannerImageUrl,Author/Title,Editor/Title,OData__ModerationStatus,PromotedState&' +
    '$expand=Author,Editor&$orderby=Modified desc&$top=5000';
  const rows = await sharePointList(token, query);
  const filtered = rows.filter((row) => {
    const moderation = row.OData__ModerationStatus ?? row._ModerationStatus;
    const promoted = row.PromotedState ?? 0;
    return moderation === 0 && promoted !== 2;
  });
  if (config.pilotPageLimit && Number.isFinite(config.pilotPageLimit)) {
    return filtered.slice(0, Math.max(1, config.pilotPageLimit));
  }
  return filtered;
}

async function getNavigation(token) {
  const links = [];
  for (const navQuery of NAV_QUERIES) {
    try {
      const entries = await sharePointList(token, navQuery);
      links.push(...entries);
    } catch {
      // Navigation extraction is best-effort.
    }
  }
  const unique = new Map();
  for (const item of links) {
    if (!item.Title || !item.Url) continue;
    const key = `${item.Title}|${item.Url}`;
    if (!unique.has(key)) {
      unique.set(key, { title: item.Title, url: item.Url });
    }
  }
  return [...unique.values()];
}

function normalizeToSharePointUrl(raw, pageFileRef) {
  if (!raw) return undefined;
  const value = raw.trim();
  if (!value || value.startsWith('#') || value.startsWith('mailto:') || value.startsWith('tel:') || value.startsWith('javascript:')) {
    return undefined;
  }
  if (value.startsWith('//')) {
    return normalizeToSharePointUrl(`https:${value}`, pageFileRef);
  }

  try {
    if (value.startsWith('http://') || value.startsWith('https://')) {
      const parsed = new URL(value);
      if (parsed.host.toLowerCase() !== config.tenantHost.toLowerCase()) {
        return undefined;
      }
      return `${parsed.pathname}${parsed.search}`;
    }
  } catch {
    return undefined;
  }

  if (value.startsWith('/')) {
    return value;
  }

  const pageDir = pageFileRef?.includes('/') ? pageFileRef.slice(0, pageFileRef.lastIndexOf('/')) : config.sitePath;
  const resolved = new URL(value, `https://${config.tenantHost}${pageDir}/`);
  if (resolved.host.toLowerCase() !== config.tenantHost.toLowerCase()) {
    return undefined;
  }
  return `${resolved.pathname}${resolved.search}`;
}

function collectAssetCandidates(page) {
  const content = [page.CanvasContent1 || '', page.Description || '', extractBannerUrl(page) || ''].join('\n');
  const candidates = new Set();
  const attrRegex = /(?:src|href|poster|data-src)=(["'])(.*?)\1/gi;
  const cssRegex = /url\((['"]?)(.*?)\1\)/gi;

  let match;
  while ((match = attrRegex.exec(content)) !== null) {
    candidates.add(match[2]);
  }
  while ((match = cssRegex.exec(content)) !== null) {
    candidates.add(match[2]);
  }
  return [...candidates]
    .map((candidate) => normalizeToSharePointUrl(candidate, page.FileRef || ''))
    .filter(Boolean);
}

function localAssetPathFromServerUrl(serverUrl) {
  const parsed = new URL(`https://${config.tenantHost}${serverUrl.startsWith('/') ? serverUrl : `/${serverUrl}`}`);
  const cleanPath = parsed.pathname.replace(/^\/+/, '');
  const queryHash = parsed.search ? `-${createHash('sha1').update(parsed.search).digest('hex').slice(0, 8)}` : '';
  const ext = path.extname(cleanPath);
  const base = ext ? cleanPath.slice(0, -ext.length) : cleanPath;
  return `/assets/${base}${queryHash}${ext}`;
}

function rewriteUrls(rawContent, replacements, pageRouteMap) {
  let content = rawContent || '';

  for (const [source, destination] of replacements.entries()) {
    const escaped = escapeRegExp(source);
    content = content.replace(new RegExp(escaped, 'g'), destination);
    if (source.startsWith('/')) {
      const absolute = `https://${config.tenantHost}${source}`;
      content = content.replace(new RegExp(escapeRegExp(absolute), 'g'), destination);
    }
  }

  content = content.replace(/href=(["'])(.*?)\1/gi, (_, quote, href) => {
    const normalized = normalizeToSharePointUrl(href, '');
    if (!normalized) return `href=${quote}${href}${quote}`;
    if (!/\.aspx($|\?)/i.test(normalized)) {
      return `href=${quote}${href}${quote}`;
    }
    const withoutQuery = normalized.split('?')[0].toLowerCase();
    const route = pageRouteMap.get(withoutQuery);
    if (!route) {
      return `href=${quote}${href}${quote}`;
    }
    return `href=${quote}${route}${quote}`;
  });

  return content;
}

function buildPageHtml({ title, description, content, canonicalUrl, navLinks }) {
  const navHtml = navLinks.length
    ? `<nav><ul>${navLinks
      .map((link) => `<li><a href="${htmlEscape(link.route)}">${htmlEscape(link.title)}</a></li>`)
      .join('')}</ul></nav>`
    : '';

  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>${htmlEscape(title || 'Grove Guidance')}</title>
  <meta name="description" content="${htmlEscape(description || '')}" />
  <link rel="canonical" href="${htmlEscape(canonicalUrl)}" />
  <style>
    body{font-family:Segoe UI,Arial,sans-serif;line-height:1.6;margin:0;color:#1b1b1b;background:#fff}
    header{padding:1rem 1.25rem;background:#f4f6f8;border-bottom:1px solid #d5d9de}
    nav ul{display:flex;flex-wrap:wrap;gap:.75rem;list-style:none;padding:0;margin:.75rem 0 0}
    nav a{text-decoration:none;color:#0f4f8c}
    main{max-width:1024px;margin:0 auto;padding:1.5rem}
    .unsupported{border-left:4px solid #ffb020;background:#fff3e0;padding:.75rem 1rem;margin:1rem 0}
  </style>
</head>
<body>
  <header>
    <h1>${htmlEscape(title || 'Grove Guidance')}</h1>
    ${navHtml}
  </header>
  <main>
    ${content}
  </main>
</body>
</html>`;
}

function mapNavigationLinks(navItems, pageRouteMap) {
  const links = [];
  for (const item of navItems) {
    const normalized = normalizeToSharePointUrl(item.url, '');
    if (!normalized) continue;
    const clean = normalized.split('?')[0].toLowerCase();
    const route = pageRouteMap.get(clean);
    if (route) {
      links.push({ title: item.title, route });
    }
  }
  return links;
}

function buildSitemap(routes) {
  const urls = routes
    .sort((a, b) => a.localeCompare(b))
    .map((route) => `<url><loc>${canonicalFromRoute(route)}</loc></url>`)
    .join('');
  return `<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">${urls}</urlset>`;
}

function buildRobots() {
  return `User-agent: *\nAllow: /\n\nSitemap: ${config.publicBaseUrl}/sitemap.xml\n`;
}

function buildStaticWebAppConfig() {
  return {
    globalHeaders: {
      'x-content-type-options': 'nosniff',
      'x-frame-options': 'SAMEORIGIN',
      'referrer-policy': 'strict-origin-when-cross-origin',
      'content-security-policy': "default-src 'self' https: data: blob:; img-src 'self' https: data: blob:; media-src 'self' https: data: blob:; style-src 'self' 'unsafe-inline' https:; script-src 'self' https:; object-src 'none'; base-uri 'self'; form-action 'self'; frame-ancestors 'none';",
    },
    routes: [
      {
        route: '/sitepages/*',
        statusCode: 301,
        redirect: '/',
      },
    ],
    navigationFallback: {
      rewrite: '/index.html',
      exclude: ['/assets/*', '/sitemap.xml', '/robots.txt'],
    },
    responseOverrides: {
      '404': {
        rewrite: '/404.html',
      },
    },
  };
}

async function loadState() {
  if (!(await exists(config.stateFile))) {
    return { pages: {}, assets: {}, lastSuccessfulSync: null };
  }
  const raw = await readFile(config.stateFile, 'utf8');
  return JSON.parse(raw);
}

async function saveState(state) {
  await ensureDir(path.dirname(config.stateFile));
  await writeFile(config.stateFile, `${JSON.stringify(state, null, 2)}\n`, 'utf8');
}

async function sync() {
  assertSyncEnvironment();
  await ensureDir(config.outputDir);
  const state = await loadState();
  const token = await getAccessToken();
  const [pages, navItems] = await Promise.all([getPublishedPages(token), getNavigation(token)]);

  const pageRouteMap = new Map();
  const pagesById = new Map();
  for (const page of pages) {
    const route = routeFromPage(page);
    const key = (page.FileRef || '').toLowerCase();
    pageRouteMap.set(key, route);
    pagesById.set(String(page.Id), page);
  }

  const seenRoutes = new Map();
  for (const page of pages) {
    const route = routeFromPage(page);
    if (seenRoutes.has(route)) {
      throw new Error(`Route collision detected: ${route} for pages ${seenRoutes.get(route)} and ${page.FileLeafRef}`);
    }
    seenRoutes.set(route, page.FileLeafRef);
  }

  const navLinks = mapNavigationLinks(navItems, pageRouteMap);

  const currentPageState = {};
  const currentAssetState = { ...state.assets };
  const publishedPageIds = new Set(pages.map((page) => String(page.Id)));

  for (const [pageId, previous] of Object.entries(state.pages || {})) {
    if (!publishedPageIds.has(pageId) && previous.route) {
      const previousOutput = outputFileFromRoute(previous.route);
      await rm(path.dirname(previousOutput), { recursive: true, force: true });
    }
  }

  const replacements = new Map();
  for (const page of pages) {
    for (const assetUrl of collectAssetCandidates(page)) {
      if (replacements.has(assetUrl)) continue;
      const localAsset = localAssetPathFromServerUrl(assetUrl);
      replacements.set(assetUrl, localAsset);

      const outputPath = path.join(config.outputDir, localAsset.replace(/^\//, ''));
      const previousAsset = currentAssetState[assetUrl];
      if (previousAsset && previousAsset.path === localAsset && await exists(outputPath)) {
        continue;
      }
      const content = await sharePointRequest(token, assetUrl, 'buffer');
      await ensureDir(path.dirname(outputPath));
      await writeFile(outputPath, content);
      currentAssetState[assetUrl] = {
        path: localAsset,
        fetchedAt: new Date().toISOString(),
      };
    }
  }

  for (const page of pages) {
    const pageId = String(page.Id);
    const route = routeFromPage(page);
    const outputFile = outputFileFromRoute(route);
    const fingerprint = createHash('sha1')
      .update(JSON.stringify({
        modified: page.Modified,
        title: page.Title,
        fileRef: page.FileRef,
        content: page.CanvasContent1,
        description: page.Description,
      }))
      .digest('hex');

    const previous = state.pages?.[pageId];
    const shouldRegenerate = !previous || previous.fingerprint !== fingerprint || !(await exists(outputFile));

    if (shouldRegenerate) {
      const rewrittenContent = rewriteUrls(page.CanvasContent1 || '', replacements, pageRouteMap) ||
        '<div class="unsupported">This page could not be fully rendered from SharePoint content.</div>';
      const html = buildPageHtml({
        title: page.Title || page.FileLeafRef,
        description: page.Description || '',
        content: rewrittenContent,
        canonicalUrl: canonicalFromRoute(route),
        navLinks,
      });
      await ensureDir(path.dirname(outputFile));
      await writeFile(outputFile, html, 'utf8');
    }

    currentPageState[pageId] = {
      route,
      fileRef: page.FileRef,
      modified: page.Modified,
      title: page.Title,
      fingerprint,
    };
  }

  const routes = [...Object.values(currentPageState).map((page) => page.route)];
  await writeFile(path.join(config.outputDir, 'sitemap.xml'), buildSitemap(routes), 'utf8');
  await writeFile(path.join(config.outputDir, 'robots.txt'), buildRobots(), 'utf8');
  await writeFile(path.join(config.outputDir, 'staticwebapp.config.json'), `${JSON.stringify(buildStaticWebAppConfig(), null, 2)}\n`, 'utf8');
  await writeFile(path.join(config.outputDir, '404.html'), buildPageHtml({
    title: 'Page Not Found',
    description: 'The requested page does not exist.',
    content: '<p>The requested page does not exist.</p><p><a href="/">Return to home</a></p>',
    canonicalUrl: `${config.publicBaseUrl}/404`,
    navLinks,
  }), 'utf8');

  await saveState({
    pages: currentPageState,
    assets: currentAssetState,
    lastSuccessfulSync: new Date().toISOString(),
  });

  await writeFile(path.join(config.outputDir, '.publish-manifest.json'), `${JSON.stringify({
    generatedAt: new Date().toISOString(),
    pageCount: pages.length,
    assetCount: Object.keys(currentAssetState).length,
    pilotPageLimit: config.pilotPageLimit || null,
  }, null, 2)}\n`, 'utf8');

  console.log(`Synced ${pages.length} published pages and ${Object.keys(currentAssetState).length} assets.`);
}

async function validate() {
  const errors = [];

  if (!(await exists(config.outputDir))) {
    throw new Error(`Output directory not found: ${config.outputDir}`);
  }

  const state = await loadState();
  const routes = new Set();
  for (const page of Object.values(state.pages || {})) {
    if (routes.has(page.route)) {
      errors.push(`Duplicate route in state: ${page.route}`);
    }
    routes.add(page.route);
    const outputFile = outputFileFromRoute(page.route);
    if (!(await exists(outputFile))) {
      errors.push(`Missing page output: ${outputFile}`);
      continue;
    }
    const html = await readFile(outputFile, 'utf8');
    if (!html.includes('<html')) {
      errors.push(`Invalid HTML output in ${outputFile}`);
    }

    const assetRefs = [...html.matchAll(/(?:src|href|poster)=(["'])(\/assets\/[^"']+)\1/gi)].map((match) => match[2]);
    for (const assetRef of assetRefs) {
      const absoluteAsset = path.join(config.outputDir, assetRef.replace(/^\//, ''));
      if (!(await exists(absoluteAsset))) {
        errors.push(`Missing asset referenced by ${outputFile}: ${assetRef}`);
      }
    }

    const internalLinks = [...html.matchAll(/href=(["'])(\/[a-z0-9\-_/]*?)\1/gi)].map((match) => match[2]);
    for (const link of internalLinks) {
      if (link.startsWith('/assets/') || link === '/') continue;
      const linkedPath = path.join(config.outputDir, link.replace(/^\//, ''), 'index.html');
      if (!(await exists(linkedPath))) {
        errors.push(`Broken internal link in ${outputFile}: ${link}`);
      }
    }
  }

  const requiredFiles = ['index.html', 'sitemap.xml', 'robots.txt', 'staticwebapp.config.json'];
  for (const file of requiredFiles) {
    const filePath = path.join(config.outputDir, file);
    if (!(await exists(filePath))) {
      errors.push(`Required output file missing: ${filePath}`);
    }
  }

  if (errors.length > 0) {
    throw new Error(`Validation failed:\n- ${errors.join('\n- ')}`);
  }

  console.log('Validation passed.');
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  const command = process.argv[2] || 'sync';
  if (command === 'sync') {
    await sync();
  } else if (command === 'validate') {
    await validate();
  } else {
    throw new Error(`Unknown command: ${command}`);
  }
}
