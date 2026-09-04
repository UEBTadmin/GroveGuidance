import test from 'node:test';
import assert from 'node:assert/strict';

import {
  findGraphSitePagesList,
  getGraphSiteListsRelativeUrl,
  normalizeGraphPageItem,
  resolveGraphSitePagesListId,
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

test('findGraphSitePagesList matches the Site Pages library by internal name', () => {
  assert.deepEqual(
    findGraphSitePagesList([
      { id: 'documents', displayName: 'Documents', name: 'Documents' },
      { id: 'pages', displayName: 'Seiten', name: 'SitePages' },
    ]),
    { id: 'pages', displayName: 'Seiten', name: 'SitePages' },
  );
});

test('findGraphSitePagesList matches the Site Pages library by webUrl path', () => {
  assert.deepEqual(
    findGraphSitePagesList([
      { id: 'documents', displayName: 'Documents', name: 'Documents' },
      {
        id: 'pages',
        displayName: 'Pages',
        name: 'pages',
        webUrl: 'https://uebt.sharepoint.com/sites/GroveGuidance/SitePages',
      },
    ]),
    {
      id: 'pages',
      displayName: 'Pages',
      name: 'pages',
      webUrl: 'https://uebt.sharepoint.com/sites/GroveGuidance/SitePages',
    },
  );
});

test('findGraphSitePagesList matches a Pages library webUrl path', () => {
  assert.deepEqual(
    findGraphSitePagesList([
      { id: 'documents', displayName: 'Documents', name: 'Documents' },
      {
        id: 'pages',
        displayName: 'Site content',
        name: 'ContentPages',
        webUrl: 'https://uebt.sharepoint.com/sites/GroveGuidance/Pages',
      },
    ]),
    {
      id: 'pages',
      displayName: 'Site content',
      name: 'ContentPages',
      webUrl: 'https://uebt.sharepoint.com/sites/GroveGuidance/Pages',
    },
  );
});

test('findGraphSitePagesList matches by Site Pages template when names are localized', () => {
  assert.deepEqual(
    findGraphSitePagesList([
      { id: 'documents', displayName: 'Documents', name: 'Documents' },
      {
        id: 'localized-pages',
        displayName: 'Siteinhalt',
        name: 'Dokumentbibliothek',
        webUrl: 'https://uebt.sharepoint.com/sites/GroveGuidance/Seitensammlung',
        list: { template: 'SitePageLibrary' },
      },
    ]),
    {
      id: 'localized-pages',
      displayName: 'Siteinhalt',
      name: 'Dokumentbibliothek',
      webUrl: 'https://uebt.sharepoint.com/sites/GroveGuidance/Seitensammlung',
      list: { template: 'SitePageLibrary' },
    },
  );
});

test('resolveGraphSitePagesListId falls back to Site Pages drive metadata when list discovery misses it', () => {
  assert.equal(
    resolveGraphSitePagesListId(
      [{ id: 'documents', displayName: 'Documents', name: 'Documents' }],
      [{
        id: 'pages-drive',
        name: 'Pages',
        webUrl: 'https://uebt.sharepoint.com/sites/GroveGuidance/SitePages',
        sharepointIds: { listId: 'site-pages-list-id' },
      }],
    ),
    'site-pages-list-id',
  );
});

test('resolveGraphSitePagesListId accepts Pages drive metadata as fallback', () => {
  assert.equal(
    resolveGraphSitePagesListId(
      [{ id: 'documents', displayName: 'Documents', name: 'Documents' }],
      [{
        id: 'pages-drive',
        name: 'Pages',
        webUrl: 'https://uebt.sharepoint.com/sites/GroveGuidance/Pages',
        sharepointIds: { listId: 'pages-list-id' },
      }],
    ),
    'pages-list-id',
  );
});

test('getGraphSiteListsRelativeUrl avoids unsupported Graph list expansion', () => {
  assert.equal(
    getGraphSiteListsRelativeUrl('site-id'),
    '/sites/site-id/lists?$select=id,displayName,name,webUrl,list',
  );
});
