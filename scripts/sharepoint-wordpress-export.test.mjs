import test from 'node:test';
import assert from 'node:assert/strict';

import {
  buildWxr,
  mediaFileNameForAsset,
  rewriteContentForWordPress,
  slugFromRoute,
} from './sharepoint-wordpress-export.mjs';

test('slugFromRoute maps the home route to "home" and strips slashes elsewhere', () => {
  assert.equal(slugFromRoute('/'), 'home');
  assert.equal(slugFromRoute('/first-login/'), 'first-login');
});

test('mediaFileNameForAsset produces a filesystem-safe filename from a SharePoint asset URL', () => {
  assert.equal(
    mediaFileNameForAsset('/sites/GroveGuidance/SiteAssets/SitePages/Home/1341352387Picture-7.jpg'),
    '1341352387Picture-7.jpg',
  );
  assert.equal(
    mediaFileNameForAsset('/sites/GroveGuidance/SiteAssets/banner.png?width=200'),
    'banner.png',
  );
});

test('rewriteContentForWordPress replaces SharePoint asset URLs with their WordPress media URLs', () => {
  const assetMap = new Map([
    ['/sites/GroveGuidance/SiteAssets/banner.png', { fileName: 'banner.png', mediaUrl: 'https://example.com/wp-content/uploads/grove-guidance-import/banner.png' }],
  ]);
  const rewritten = rewriteContentForWordPress(
    '<img src="/sites/GroveGuidance/SiteAssets/banner.png" />',
    assetMap,
  );
  assert.equal(rewritten, '<img src="https://example.com/wp-content/uploads/grove-guidance-import/banner.png" />');
});

test('buildWxr produces a WXR document with page and attachment items', () => {
  const pages = [
    {
      Id: '1',
      Title: 'Home',
      FileRef: '/sites/GroveGuidance/SitePages/Home.aspx',
      FileLeafRef: 'Home.aspx',
      CanvasContent1: '<p>Welcome</p>',
      Description: 'Welcome page',
      Created: '2026-01-01T00:00:00Z',
      Modified: '2026-01-02T00:00:00Z',
      FirstPublishedDate: '2026-01-01T00:00:00Z',
    },
  ];
  const assetMap = new Map([
    ['/sites/GroveGuidance/SiteAssets/banner.png', { fileName: 'banner.png', mediaUrl: 'https://example.com/wp-content/uploads/grove-guidance-import/banner.png' }],
  ]);

  const wxr = buildWxr(pages, assetMap);
  assert.match(wxr, /<wp:wxr_version>1\.2<\/wp:wxr_version>/);
  assert.match(wxr, /<title>Home<\/title>/);
  assert.match(wxr, /<wp:post_type>page<\/wp:post_type>/);
  assert.match(wxr, /<wp:post_type>attachment<\/wp:post_type>/);
  assert.match(wxr, /<wp:post_name>home<\/wp:post_name>/);
  assert.match(wxr, /Welcome/);
});
