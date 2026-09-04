import test from 'node:test';
import assert from 'node:assert/strict';

import { resolveSharePointLocation } from './sharepoint-publish.mjs';

test('resolveSharePointLocation trims Site Pages URLs to the site root', () => {
  assert.deepEqual(
    resolveSharePointLocation(
      'uebt.sharepoint.com/sites/GroveGuidance/SitePages/Home.aspx?viewid=123',
      '',
    ),
    {
      tenantHost: 'uebt.sharepoint.com',
      sitePath: '/sites/GroveGuidance',
    },
  );
});

test('resolveSharePointLocation trims document library URLs to the site root', () => {
  assert.deepEqual(
    resolveSharePointLocation(
      'uebt.sharepoint.com',
      '/sites/GroveGuidance/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FGroveGuidance%2FShared%20Documents',
    ),
    {
      tenantHost: 'uebt.sharepoint.com',
      sitePath: '/sites/GroveGuidance',
    },
  );
});

test('resolveSharePointLocation preserves subweb paths while trimming file URLs', () => {
  assert.deepEqual(
    resolveSharePointLocation(
      'https://uebt.sharepoint.com/sites/GroveGuidance/Subweb/SiteAssets/logo.png',
      '',
    ),
    {
      tenantHost: 'uebt.sharepoint.com',
      sitePath: '/sites/GroveGuidance/Subweb',
    },
  );
});

test('resolveSharePointLocation trims SharePoint sharing-link style paths to site root', () => {
  assert.deepEqual(
    resolveSharePointLocation(
      'uebt.sharepoint.com',
      '/:w:/r/sites/GroveGuidance/SitePages/Home.aspx?d=w123',
    ),
    {
      tenantHost: 'uebt.sharepoint.com',
      sitePath: '/sites/GroveGuidance',
    },
  );
});

test('resolveSharePointLocation derives the site path from id query parameter when needed', () => {
  assert.deepEqual(
    resolveSharePointLocation(
      'uebt.sharepoint.com',
      '/_layouts/15/Doc.aspx?sourcedoc=%7B123%7D&id=%2Fsites%2FGroveGuidance%2FShared%20Documents%2FGeneral',
    ),
    {
      tenantHost: 'uebt.sharepoint.com',
      sitePath: '/sites/GroveGuidance',
    },
  );
});

test('resolveSharePointLocation keeps subwebs from sharing-link style paths', () => {
  assert.deepEqual(
    resolveSharePointLocation(
      'uebt.sharepoint.com',
      '/:w:/r/sites/GroveGuidance/Subweb/SitePages/Guide.aspx',
    ),
    {
      tenantHost: 'uebt.sharepoint.com',
      sitePath: '/sites/GroveGuidance/Subweb',
    },
  );
});
