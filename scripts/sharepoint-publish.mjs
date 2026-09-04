#!/usr/bin/env node
import { createHash } from 'node:crypto';
import { mkdir, readFile, rm, stat, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

const DEFAULT_SP_TENANT_HOST = 'uebt.sharepoint.com';
const DEFAULT_SP_SITE_PATH = '/sites/GroveGuidance';
const MAX_RETRIES = 3;
const RETRY_DELAY_MS = 2000;
const SHAREPOINT_LIBRARY_SEGMENTS = new Set([
  'documents',
  'shared documents',
  'form templates',
  'siteassets',
  'site assets',
  'stylelibrary',
  'style library',
]);

function decodePathSegment(segment) {
  try {
    return decodeURIComponent(segment);
  } catch {
    return segment;
  }
}

function normalizePathCandidate(pathValue) {
  let value = (pathValue || '').trim();
  if (!value) {
    return '';
  }
  value = value.split('#', 1)[0] || '';
  try {
    value = decodeURIComponent(value);
  } catch {
    // Keep the raw candidate when decoding fails.
  }
  value = value.replace(/\\/g, '/');

  const siteMatch = value.match(/\/(?:sites|teams)\/[^/?#]+(?:\/[^?#]*)?/i);
  if (siteMatch) {
    value = siteMatch[0];
  }

  if (!value.startsWith('/')) {
    value = `/${value}`;
  }

  return value;
}

function normalizeSharePointSitePath(sitePathValue) {
  const rawSitePath = (sitePathValue || '').trim();
  const [rawPath = '', rawQuery = ''] = rawSitePath.split('?', 2);
  let sitePath = normalizePathCandidate(rawPath);

  if (!/\/(?:sites|teams)\//i.test(sitePath) && rawQuery) {
    const params = new URLSearchParams(rawQuery.split('#', 1)[0] || '');
    for (const key of ['id', 'RootFolder', 'rootfolder']) {
      const candidate = params.get(key);
      if (!candidate) continue;
      const normalizedCandidate = normalizePathCandidate(candidate);
      if (/\/(?:sites|teams)\//i.test(normalizedCandidate)) {
        sitePath = normalizedCandidate;
        break;
      }
    }
  }

  sitePath = sitePath.split(/[?#]/, 1)[0] || '';
  if (!sitePath.startsWith('/')) {
    sitePath = `/${sitePath}`;
  }
  sitePath = sitePath.replace(/\/+$/, '') || '/';
  sitePath = sitePath.split(/\/(?:_api|_layouts)\b/i)[0] || '/';

  const segments = sitePath.split('/').filter(Boolean);
  const libraryIndex = segments.findIndex((segment, index) => (
    index >= 2 && SHAREPOINT_LIBRARY_SEGMENTS.has(decodePathSegment(segment).toLowerCase())
  ));
  if (libraryIndex >= 0) {
    segments.splice(libraryIndex);
  }

  while (segments.length > 0) {
    const lastSegment = decodePathSegment(segments[segments.length - 1]);
    if (/\.[a-z0-9]+$/i.test(lastSegment) || /^forms$/i.test(lastSegment)) {
      segments.pop();
      continue;
    }
    if (/^sitepages$/i.test(lastSegment) || /^pages$/i.test(lastSegment)) {
      segments.pop();
      continue;
    }
    if (segments.length > 2 && SHAREPOINT_LIBRARY_SEGMENTS.has(lastSegment.toLowerCase())) {
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

export function normalizeGraphDriveKey(value) {
  return decodePathSegment(value || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '');
}

const SITE_PAGES_KEYS = new Set(['sitepages', 'pages', 'sitepagelibrary']);

function lastUrlPathSegment(webUrl) {
  if (!webUrl) return '';
  try {
    const pathname = new URL(webUrl).pathname;
    const segments = pathname.split('/').filter(Boolean);
    return segments.at(-1) || '';
  } catch {
    return '';
  }
}

export function findGraphSitePagesList(lists) {
  return (lists || []).find((list) => {
    const template = normalizeGraphDriveKey(list?.list?.template);
    if (SITE_PAGES_KEYS.has(template)) {
      return true;
    }

    const candidateKeys = [
      list?.displayName,
      list?.name,
      lastUrlPathSegment(list?.webUrl),
    ].map(normalizeGraphDriveKey);

    return candidateKeys.some((key) => SITE_PAGES_KEYS.has(key));
  });
}

export function resolveGraphSitePagesListId(lists, drives) {
  const sitePagesList = findGraphSitePagesList(lists);
  if (sitePagesList?.id) {
    return sitePagesList.id;
  }

  const sitePagesDrive = (drives || []).find((drive) => {
    const candidateKeys = [
      drive?.name,
      lastUrlPathSegment(drive?.webUrl),
    ].map(normalizeGraphDriveKey);

    return candidateKeys.some((key) => SITE_PAGES_KEYS.has(key));
  });

  return sitePagesDrive?.sharepointIds?.listId;
}

export function getGraphSiteListsRelativeUrl(siteId) {
  return `/sites/${siteId}/lists?$select=id,displayName,name,webUrl,list`;
}

export function splitGraphAssetServerRelativePath(serverUrl, sitePathValue = config.sitePath) {
  const cleanPath = (serverUrl || '').split(/[?#]/, 1)[0] || '';
  const normalizedSitePath = (sitePathValue || '').replace(/\/+$/, '');
  const prefix = normalizedSitePath && normalizedSitePath !== '/' ? `${normalizedSitePath}/` : '/';
  if (!cleanPath.toLowerCase().startsWith(prefix.toLowerCase())) {
    return undefined;
  }

  const segments = cleanPath.slice(prefix.length).split('/').filter(Boolean).map(decodePathSegment);
  if (segments.length < 2) {
    return undefined;
  }

  const [librarySegment, ...itemSegments] = segments;
  return {
    driveLookupKey: normalizeGraphDriveKey(librarySegment),
    itemPath: itemSegments.join('/'),
  };
}

export function normalizeGraphPageItem(item) {
  const fields = item?.fields || {};
  let fileRef = fields.FileRef;
  if (!fileRef && item?.webUrl) {
    try {
      fileRef = new URL(item.webUrl).pathname;
    } catch {
      fileRef = item.webUrl;
    }
  }

  return {
    Id: fields.Id ?? fields.ID ?? item?.id,
    Title: fields.Title ?? item?.name,
    FileLeafRef: fields.FileLeafRef ?? fields.LinkFilename ?? item?.name,
    FileRef: fileRef,
    Modified: fields.Modified ?? item?.lastModifiedDateTime,
    Created: fields.Created ?? item?.createdDateTime,
    FirstPublishedDate: fields.FirstPublishedDate,
    Description: fields.Description,
    CanvasContent1: fields.CanvasContent1,
    BannerImageUrl: fields.BannerImageUrl,
    OData__ModerationStatus: fields.OData__ModerationStatus ?? fields._ModerationStatus,
    PromotedState: fields.PromotedState,
  };
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

async function getAccessToken(scope) {
  const tokenUrl = `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type: 'client_credentials',
    client_id: config.clientId,
    client_secret: config.clientSecret,
    scope,
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

async function getGraphAccessToken() {
  return getAccessToken('https://graph.microsoft.com/.default');
}

async function getSharePointAccessToken() {
  return getAccessToken(`https://${config.tenantHost}/.default`);
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
      
      const error = new Error(`SharePoint request failed (${response.status} ${response.statusText}) for ${url}`);
      error.status = response.status;
      error.url = url;
      error.responseBody = errorBody;
      throw error;
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

async function graphRequest(token, relativeUrl, responseType = 'json', attempt = 1) {
  const url = relativeUrl.startsWith('http') ? relativeUrl : `https://graph.microsoft.com/v1.0${relativeUrl}`;

  try {
    const response = await fetch(url, {
      headers: {
        authorization: 'Bearer ' + token,
        accept: 'application/json',
      },
    });

    if (!response.ok) {
      const errorBody = await response.text();
      console.error(`Graph request failed (${response.status} ${response.statusText}) for ${url}`);
      console.error(`Response body: ${errorBody.substring(0, 500)}`);

      if ([500, 502, 503, 504].includes(response.status) && attempt < MAX_RETRIES) {
        console.log(`Retrying Graph request (attempt ${attempt + 1}/${MAX_RETRIES}) after ${RETRY_DELAY_MS}ms...`);
        await sleep(RETRY_DELAY_MS);
        return graphRequest(token, relativeUrl, responseType, attempt + 1);
      }

      const error = new Error(`Graph request failed (${response.status} ${response.statusText}) for ${url}`);
      error.status = response.status;
      error.url = url;
      error.responseBody = errorBody;
      throw error;
    }

    if (responseType === 'buffer') {
      return Buffer.from(await response.arrayBuffer());
    }
    return response.json();
  } catch (error) {
    if (attempt < MAX_RETRIES && error.message && !error.message.includes('401')) {
      console.log(`Network error, retrying (attempt ${attempt + 1}/${MAX_RETRIES}):`, error.message);
      await sleep(RETRY_DELAY_MS);
      return graphRequest(token, relativeUrl, responseType, attempt + 1);
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

async function graphList(token, relativeUrl) {
  const values = [];
  let next = relativeUrl;
  while (next) {
    const json = await graphRequest(token, next);
    const batch = json.value || [];
    values.push(...batch);
    next = json['@odata.nextLink'] || undefined;
  }
  return values;
}

let graphSiteContextPromise;

function firstDrivePathSegment(webUrl) {
  if (!webUrl) return undefined;
  try {
    const pathname = new URL(webUrl).pathname;
    const cleanSitePath = config.sitePath.replace(/\/+$/, '');
    const relative = pathname.toLowerCase().startsWith(`${cleanSitePath.toLowerCase()}/`)
      ? pathname.slice(cleanSitePath.length)
      : pathname;
    return relative.split('/').filter(Boolean)[0];
  } catch {
    return undefined;
  }
}

async function getGraphSiteContext(token) {
  if (!graphSiteContextPromise) {
    graphSiteContextPromise = (async () => {
      const site = await graphRequest(token, `/sites/${config.tenantHost}:${config.sitePath}?$select=id`);
      const [drives, lists] = await Promise.all([
        graphList(token, `/sites/${site.id}/drives?$select=id,name,webUrl,sharepointIds`),
        graphList(token, getGraphSiteListsRelativeUrl(site.id)),
      ]);

      const sitePagesListId = resolveGraphSitePagesListId(lists, drives);
      if (!sitePagesListId) {
        throw new Error(`Could not locate the SharePoint "Site Pages" list for ${config.sitePath}`);
      }

      const drivesByKey = new Map();
      for (const drive of drives) {
        for (const candidate of [drive.name, firstDrivePathSegment(drive.webUrl)]) {
          const key = normalizeGraphDriveKey(candidate);
          if (key && !drivesByKey.has(key)) {
            drivesByKey.set(key, drive);
          }
        }
      }

      return {
        siteId: site.id,
        sitePagesListId,
        drivesByKey,
      };
    })();
  }

  return graphSiteContextPromise;
}

async function getPublishedPages(graphToken, getOptionalSharePointToken = async () => undefined) {
  try {
    const { siteId, sitePagesListId } = await getGraphSiteContext(graphToken);
    const rows = await graphList(graphToken, `/sites/${siteId}/lists/${sitePagesListId}/items?$expand=fields`);
    const filtered = rows
      .map(normalizeGraphPageItem)
      .filter((row) => row.FileLeafRef && /\.aspx$/i.test(row.FileLeafRef))
      .filter((row) => {
        const moderation = Number(row.OData__ModerationStatus ?? row._ModerationStatus ?? 0);
        const promoted = Number(row.PromotedState ?? 0);
        return moderation === 0 && promoted !== 2;
      })
      .sort((left, right) => new Date(right.Modified || 0) - new Date(left.Modified || 0));
    if (config.pilotPageLimit && Number.isFinite(config.pilotPageLimit)) {
      return filtered.slice(0, Math.max(1, config.pilotPageLimit));
    }
    return filtered;
  } catch (error) {
    if (!isGraphAuthFailure(error)) {
      throw error;
    }
    console.warn(`Graph page sync unavailable (${error.message}); falling back to SharePoint REST.`);
  }

  const sharePointToken = await getOptionalSharePointToken();
  if (!sharePointToken) {
    throw new Error('Graph page sync failed and no SharePoint token was available for fallback.');
  }
  try {
    return await getPublishedPagesFromSharePoint(sharePointToken);
  } catch (error) {
    if (!isSharePointUnsupportedAppOnlyToken(error)) {
      throw error;
    }
    throw new Error(
      'Graph page sync was unauthorized and SharePoint REST fallback rejected the app-only token. '
      + 'In Entra, grant this app Microsoft Graph application access to the target site '
      + '(Sites.Read.All, or Sites.Selected plus a site-level grant) and admin-consent it. '
      + 'SharePoint REST app-only with a client secret is not supported by this site.',
    );
  }
}

async function getNavigation(getOptionalSharePointToken = async () => undefined) {
  const token = await getOptionalSharePointToken();
  if (!token) {
    return [];
  }
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

async function getAssetContent(graphToken, getOptionalSharePointToken, serverUrl) {
  let graphContext;
  try {
    graphContext = await getGraphSiteContext(graphToken);
  } catch (error) {
    if (!isGraphAuthFailure(error)) {
      throw error;
    }
  }
  const assetPath = splitGraphAssetServerRelativePath(serverUrl);
  if (graphContext && assetPath) {
    const drive = graphContext.drivesByKey.get(assetPath.driveLookupKey);
    if (drive) {
      const encodedItemPath = assetPath.itemPath.split('/').map((segment) => encodeURIComponent(segment)).join('/');
      try {
        return await graphRequest(
          graphToken,
          `/sites/${graphContext.siteId}/drives/${drive.id}/root:/${encodedItemPath}:/content`,
          'buffer',
        );
      } catch (error) {
        if (!isGraphAuthOrTransientFailure(error)) {
          throw error;
        }
      }
    }
  }

  const sharePointToken = await getOptionalSharePointToken();
  if (sharePointToken) {
    return sharePointRequest(sharePointToken, serverUrl, 'buffer');
  }

  throw new Error(`Could not resolve SharePoint asset for download: ${serverUrl}`);
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

function isGraphAuthFailure(error) {
  const message = String(error?.message || '');
  return /Graph request failed \((401|403)\b/i.test(message);
}

function isSharePointUnsupportedAppOnlyToken(error) {
  const message = String(error?.message || '');
  const responseBody = String(error?.responseBody || '');
  return (
    /SharePoint request failed \(401\b/i.test(message)
    && /Unsupported app only token\./i.test(responseBody)
  );
}

function isGraphAuthOrTransientFailure(error) {
  const message = String(error?.message || '');
  return isGraphAuthFailure(error) || /Graph request failed \((500|502|503|504)\b/i.test(message);
}

async function getPublishedPagesFromSharePoint(token) {
  const rows = await sharePointList(
    token,
    `${config.sitePath}/_api/web/lists/getByTitle('Site Pages')/items?$select=Id,Title,FileLeafRef,FileRef,Modified,Created,FirstPublishedDate,Description,CanvasContent1,BannerImageUrl,OData__ModerationStatus,PromotedState`,
  );
  const filtered = rows
    .filter((row) => row.FileLeafRef && /\.aspx$/i.test(row.FileLeafRef))
    .filter((row) => {
      const moderation = Number(row.OData__ModerationStatus ?? row._ModerationStatus ?? 0);
      const promoted = Number(row.PromotedState ?? 0);
      return moderation === 0 && promoted !== 2;
    })
    .sort((left, right) => new Date(right.Modified || 0) - new Date(left.Modified || 0));
  if (config.pilotPageLimit && Number.isFinite(config.pilotPageLimit)) {
    return filtered.slice(0, Math.max(1, config.pilotPageLimit));
  }
  return filtered;
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
  const graphToken = await getGraphAccessToken();
  let sharePointTokenPromise;
  const getOptionalSharePointToken = async () => {
    if (!sharePointTokenPromise) {
      sharePointTokenPromise = getSharePointAccessToken().catch(() => undefined);
    }
    return sharePointTokenPromise;
  };
  const [pages, navItems] = await Promise.all([
    getPublishedPages(graphToken, getOptionalSharePointToken),
    getNavigation(getOptionalSharePointToken),
  ]);

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
      const content = await getAssetContent(graphToken, getOptionalSharePointToken, assetUrl);
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
