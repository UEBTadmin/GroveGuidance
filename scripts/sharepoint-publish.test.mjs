import test from 'node:test';
import assert from 'node:assert/strict';

import {
  normalizeGraphPageItem,
  resolveSharePointLocation,
  splitGraphAssetServerRelativePath,
} from './sharepoint-publish.mjs';

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

test('splitGraphAssetServerRelativePath extracts the library key and file path from encoded URLs', () => {
  assert.deepEqual(
    splitGraphAssetServerRelativePath(
      '/sites/GroveGuidance/Shared%20Documents/Policies/Guide.pdf?download=1',
      '/sites/GroveGuidance',
    ),
    {
      driveLookupKey: 'shareddocuments',
      itemPath: 'Policies/Guide.pdf',
    },
  );
});

test('normalizeGraphPageItem maps Graph list-item fields into the publisher page shape', () => {
  assert.deepEqual(
    normalizeGraphPageItem({
      id: '42',
      name: 'Home.aspx',
      webUrl: 'https://uebt.sharepoint.com/sites/GroveGuidance/SitePages/Home.aspx',
      lastModifiedDateTime: '2026-09-04T10:00:00Z',
      createdDateTime: '2026-09-04T09:00:00Z',
      fields: {
        Title: 'Home',
        FileLeafRef: 'Home.aspx',
        Description: 'Welcome',
        CanvasContent1: '<div>Welcome</div>',
        BannerImageUrl: '/sites/GroveGuidance/SiteAssets/banner.png',
        _ModerationStatus: '0',
        PromotedState: '1',
      },
    }),
    {
      Id: '42',
      Title: 'Home',
      FileLeafRef: 'Home.aspx',
      FileRef: '/sites/GroveGuidance/SitePages/Home.aspx',
      Modified: '2026-09-04T10:00:00Z',
      Created: '2026-09-04T09:00:00Z',
      FirstPublishedDate: undefined,
      Description: 'Welcome',
      CanvasContent1: '<div>Welcome</div>',
      BannerImageUrl: '/sites/GroveGuidance/SiteAssets/banner.png',
      OData__ModerationStatus: '0',
      PromotedState: '1',
    },
  );
});
