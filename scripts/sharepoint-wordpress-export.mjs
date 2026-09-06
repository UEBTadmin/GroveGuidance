#!/usr/bin/env node
// One-time export of published GroveGuidance SharePoint pages into a WordPress WXR
// (WordPress eXtended RSS) file, plus a downloaded copy of every referenced media
// asset. This is NOT part of the ongoing sharepoint-publish.mjs sync pipeline - it's a
// standalone migration tool intended to be run once, with its output imported into a
// WordPress site via Tools -> Import -> WordPress (or `wp import`).
//
// Usage:
//   node ./scripts/sharepoint-wordpress-export.mjs
//
// Required environment variables (same as sharepoint-publish.mjs):
//   SP_TENANT_ID, SP_CLIENT_ID, SP_CLIENT_SECRET
// Optional:
//   SP_TENANT_HOST, SP_SITE_PATH, PUBLIC_BASE_URL, WP_EXPORT_DIR (default: wp-export)
//
// Output:
//   <WP_EXPORT_DIR>/wordpress-export.xml  - WXR file to import into WordPress
//   <WP_EXPORT_DIR>/media/...             - downloaded media assets, named to match
//                                            the <wp:attachment_url> placeholders in the XML
//   <WP_EXPORT_DIR>/navigation.json       - a simple nav menu (label + slug) derived from
//                                            the published pages, to recreate manually under
//                                            WordPress Appearance -> Menus (WXR does not
//                                            reliably carry SharePoint's QuickLaunch structure)

import { mkdir, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

import {
  collectAssetCandidates,
  getAssetContent,
  getGraphAccessToken,
  getPublishedPagesViaGraphPagesApi,
  navLinksFromPages,
  routeFromPage,
} from './sharepoint-publish.mjs';

const outputDir = path.resolve(process.cwd(), process.env.WP_EXPORT_DIR || 'wp-export');
const mediaDir = path.join(outputDir, 'media');
const publicBaseUrl = (process.env.PUBLIC_BASE_URL || 'https://groveguidance.uebt.org').replace(/\/$/, '');
const siteTitle = process.env.WP_SITE_TITLE || 'Grove Guidance';

function xmlEscape(value = '') {
  return String(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;');
}

function cdata(value = '') {
  return `<![CDATA[${String(value).replaceAll(']]>', ']]]]><![CDATA[>')}]]>`;
}

export function slugFromRoute(route) {
  if (route === '/') return 'home';
  return route.replace(/^\//, '').replace(/\/$/, '');
}

/** Deterministic, filesystem-safe local filename for a downloaded SharePoint asset URL. */
export function mediaFileNameForAsset(assetUrl) {
  const cleanPath = assetUrl.split('?')[0];
  const base = path.posix.basename(cleanPath) || 'asset';
  return base.replace(/[^a-zA-Z0-9._-]/g, '-');
}

async function downloadMedia(pages, graphToken) {
  const assetMap = new Map(); // sharePointUrl -> { fileName, localPath, mediaUrl }
  await mkdir(mediaDir, { recursive: true });

  for (const page of pages) {
    for (const assetUrl of collectAssetCandidates(page)) {
      if (assetMap.has(assetUrl)) continue;
      const fileName = mediaFileNameForAsset(assetUrl);
      let uniqueFileName = fileName;
      let counter = 1;
      const taken = new Set([...assetMap.values()].map((entry) => entry.fileName));
      while (taken.has(uniqueFileName)) {
        const ext = path.extname(fileName);
        const base = ext ? fileName.slice(0, -ext.length) : fileName;
        uniqueFileName = `${base}-${counter}${ext}`;
        counter += 1;
      }

      try {
        const content = await getAssetContent(graphToken, async () => undefined, assetUrl);
        await writeFile(path.join(mediaDir, uniqueFileName), content);
        assetMap.set(assetUrl, {
          fileName: uniqueFileName,
          mediaUrl: `${publicBaseUrl}/wp-content/uploads/grove-guidance-import/${uniqueFileName}`,
        });
        console.log(`Downloaded media: ${assetUrl} -> media/${uniqueFileName}`);
      } catch (error) {
        console.warn(`Skipping media asset (download failed): ${assetUrl} (${error.message})`);
      }
    }
  }

  return assetMap;
}

export function rewriteContentForWordPress(content, assetMap) {
  let rewritten = content || '';
  for (const [sharePointUrl, { mediaUrl }] of assetMap.entries()) {
    rewritten = rewritten.split(sharePointUrl).join(mediaUrl);
  }
  return rewritten;
}

function buildAttachmentItem(assetUrl, entry, pubDate) {
  return `    <item>
      <title>${xmlEscape(entry.fileName)}</title>
      <link>${xmlEscape(entry.mediaUrl)}</link>
      <pubDate>${xmlEscape(pubDate)}</pubDate>
      <dc:creator>${cdata('admin')}</dc:creator>
      <guid isPermaLink="false">${xmlEscape(entry.mediaUrl)}</guid>
      <description></description>
      <content:encoded>${cdata('')}</content:encoded>
      <excerpt:encoded>${cdata('')}</excerpt:encoded>
      <wp:post_id>0</wp:post_id>
      <wp:post_date>${xmlEscape(pubDate)}</wp:post_date>
      <wp:post_date_gmt>${xmlEscape(pubDate)}</wp:post_date_gmt>
      <wp:comment_status>closed</wp:comment_status>
      <wp:ping_status>closed</wp:ping_status>
      <wp:post_name>${xmlEscape(entry.fileName)}</wp:post_name>
      <wp:status>inherit</wp:status>
      <wp:post_parent>0</wp:post_parent>
      <wp:menu_order>0</wp:menu_order>
      <wp:post_type>attachment</wp:post_type>
      <wp:post_password></wp:post_password>
      <wp:is_sticky>0</wp:is_sticky>
      <wp:attachment_url>${xmlEscape(entry.mediaUrl)}</wp:attachment_url>
    </item>`;
}

function buildPageItem(page, assetMap, index) {
  const route = routeFromPage(page);
  const slug = slugFromRoute(route);
  const title = page.Title || page.FileLeafRef || slug;
  const content = rewriteContentForWordPress(page.CanvasContent1, assetMap);
  const pubDate = page.FirstPublishedDate || page.Created || new Date().toISOString();
  const modDate = page.Modified || pubDate;

  return `    <item>
      <title>${xmlEscape(title)}</title>
      <link>${xmlEscape(`${publicBaseUrl}${route}`)}</link>
      <pubDate>${xmlEscape(pubDate)}</pubDate>
      <dc:creator>${cdata('admin')}</dc:creator>
      <guid isPermaLink="false">${xmlEscape(`${publicBaseUrl}${route}?post_type=page#sharepoint-id-${page.Id}`)}</guid>
      <description></description>
      <content:encoded>${cdata(content)}</content:encoded>
      <excerpt:encoded>${cdata(page.Description || '')}</excerpt:encoded>
      <wp:post_id>${1000 + index}</wp:post_id>
      <wp:post_date>${xmlEscape(pubDate)}</wp:post_date>
      <wp:post_date_gmt>${xmlEscape(pubDate)}</wp:post_date_gmt>
      <wp:post_modified>${xmlEscape(modDate)}</wp:post_modified>
      <wp:post_modified_gmt>${xmlEscape(modDate)}</wp:post_modified_gmt>
      <wp:comment_status>closed</wp:comment_status>
      <wp:ping_status>closed</wp:ping_status>
      <wp:post_name>${xmlEscape(slug)}</wp:post_name>
      <wp:status>publish</wp:status>
      <wp:post_parent>0</wp:post_parent>
      <wp:menu_order>0</wp:menu_order>
      <wp:post_type>page</wp:post_type>
      <wp:post_password></wp:post_password>
      <wp:is_sticky>0</wp:is_sticky>
      <wp:postmeta>
        <wp:meta_key>_sharepoint_original_file_ref</wp:meta_key>
        <wp:meta_value>${cdata(page.FileRef || '')}</wp:meta_value>
      </wp:postmeta>
    </item>`;
}

export function buildWxr(pages, assetMap) {
  const now = new Date().toUTCString();
  const attachmentItems = [...assetMap.entries()]
    .map(([assetUrl, entry]) => buildAttachmentItem(assetUrl, entry, now))
    .join('\n');
  const pageItems = pages.map((page, index) => buildPageItem(page, assetMap, index)).join('\n');

  return `<?xml version="1.0" encoding="UTF-8"?>
<rss version="2.0"
  xmlns:excerpt="http://wordpress.org/export/1.2/excerpt/"
  xmlns:content="http://purl.org/rss/1.0/modules/content/"
  xmlns:wfw="http://wellformedweb.org/CommentAPI/"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:wp="http://wordpress.org/export/1.2/"
>
<channel>
  <title>${xmlEscape(siteTitle)}</title>
  <link>${xmlEscape(publicBaseUrl)}</link>
  <description>Exported from SharePoint site GroveGuidance</description>
  <pubDate>${now}</pubDate>
  <language>en-US</language>
  <wp:wxr_version>1.2</wp:wxr_version>
  <wp:base_site_url>${xmlEscape(publicBaseUrl)}</wp:base_site_url>
  <wp:base_blog_url>${xmlEscape(publicBaseUrl)}</wp:base_blog_url>
  <wp:author>
    <wp:author_id>1</wp:author_id>
    <wp:author_login>${cdata('admin')}</wp:author_login>
    <wp:author_email>${cdata('admin@example.com')}</wp:author_email>
    <wp:author_display_name>${cdata('admin')}</wp:author_display_name>
  </wp:author>
${attachmentItems}
${pageItems}
</channel>
</rss>
`;
}

async function main() {
  const requiredEnv = ['SP_TENANT_ID', 'SP_CLIENT_ID', 'SP_CLIENT_SECRET'];
  const missing = requiredEnv.filter((key) => !process.env[key]);
  if (missing.length > 0) {
    throw new Error(`Missing required environment variables: ${missing.join(', ')}`);
  }

  await mkdir(outputDir, { recursive: true });

  const graphToken = await getGraphAccessToken();
  const pages = await getPublishedPagesViaGraphPagesApi(graphToken);
  console.log(`Found ${pages.length} published page(s) to export.`);

  const assetMap = await downloadMedia(pages, graphToken);
  console.log(`Downloaded ${assetMap.size} media asset(s).`);

  const wxr = buildWxr(pages, assetMap);
  await writeFile(path.join(outputDir, 'wordpress-export.xml'), wxr, 'utf8');
  console.log(`Wrote ${path.join(outputDir, 'wordpress-export.xml')}`);

  const navLinks = navLinksFromPages(pages);
  await writeFile(
    path.join(outputDir, 'navigation.json'),
    `${JSON.stringify(navLinks, null, 2)}\n`,
    'utf8',
  );
  console.log(`Wrote ${path.join(outputDir, 'navigation.json')} (${navLinks.length} nav link(s) - recreate this menu manually under WordPress Appearance > Menus).`);

  console.log('\nNext steps:');
  console.log('1. In WordPress, go to Tools > Import > WordPress (install the importer plugin if prompted).');
  console.log(`2. Upload ${path.join(outputDir, 'wordpress-export.xml')} and import all pages, assigning them to a user.`);
  console.log(`3. Upload the files in ${mediaDir} to your Media Library (or place them at wp-content/uploads/grove-guidance-import/ to match the URLs already rewritten into the imported content).`);
  console.log(`4. Recreate the site navigation menu from ${path.join(outputDir, 'navigation.json')} under Appearance > Menus.`);
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  await main();
}
