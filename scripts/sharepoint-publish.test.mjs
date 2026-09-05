import test from 'node:test';
import assert from 'node:assert/strict';

import {
  collectAssetCandidates,
  findGraphSitePagesList,
  getGraphSiteListsRelativeUrl,
  isGraphSitePageItem,
  isPublishedGraphSitePage,
  normalizeGraphPageItem,
  normalizeGraphSitePage,
  renderCanvasLayoutHtml,
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

test('findGraphSitePagesList matches the Site Pages library by default view webUrl path', () => {
  assert.deepEqual(
    findGraphSitePagesList([
      { id: 'documents', displayName: 'Documents', name: 'Documents' },
      {
        id: 'pages',
        displayName: 'Websiteinhalt',
        name: 'Bibliothek',
        webUrl: 'https://uebt.sharepoint.com/sites/GroveGuidance/SitePages/Forms/AllItems.aspx',
      },
    ]),
    {
      id: 'pages',
      displayName: 'Websiteinhalt',
      name: 'Bibliothek',
      webUrl: 'https://uebt.sharepoint.com/sites/GroveGuidance/SitePages/Forms/AllItems.aspx',
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

test('findGraphSitePagesList matches by WebPageLibrary template when names are localized', () => {
  assert.deepEqual(
    findGraphSitePagesList([
      { id: 'documents', displayName: 'Documents', name: 'Documents' },
      {
        id: 'localized-web-pages',
        displayName: 'Websiteinhalte',
        name: 'Bibliothek',
        webUrl: 'https://uebt.sharepoint.com/sites/GroveGuidance/Seiten',
        list: { template: 'WebPageLibrary' },
      },
    ]),
    {
      id: 'localized-web-pages',
      displayName: 'Websiteinhalte',
      name: 'Bibliothek',
      webUrl: 'https://uebt.sharepoint.com/sites/GroveGuidance/Seiten',
      list: { template: 'WebPageLibrary' },
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

test('resolveGraphSitePagesListId accepts drive metadata when the webUrl is a library view page', () => {
  assert.equal(
    resolveGraphSitePagesListId(
      [{ id: 'documents', displayName: 'Documents', name: 'Documents' }],
      [{
        id: 'pages-drive',
        name: 'Websiteinhalt',
        webUrl: 'https://uebt.sharepoint.com/sites/GroveGuidance/SitePages/Forms/ByAuthor.aspx',
        sharepointIds: { listId: 'site-pages-list-id' },
      }],
    ),
    'site-pages-list-id',
  );
});

test('getGraphSiteListsRelativeUrl includes hidden lists without unsupported expansion', () => {
  assert.equal(
    getGraphSiteListsRelativeUrl('site-id'),
    '/sites/site-id/lists?includeHiddenLists=true&$select=id,displayName,name,webUrl,list',
  );
});

test('isGraphSitePageItem detects Site Page fields on Graph list items', () => {
  assert.equal(
    isGraphSitePageItem({
      name: 'Home.aspx',
      webUrl: 'https://uebt.sharepoint.com/sites/GroveGuidance/Seitensammlung/Home.aspx',
      fields: {
        FileLeafRef: 'Home.aspx',
        FileRef: '/sites/GroveGuidance/Seitensammlung/Home.aspx',
        CanvasContent1: '<div>Welcome</div>',
      },
    }),
    true,
  );
});

test('isGraphSitePageItem ignores non-page documents', () => {
  assert.equal(
    isGraphSitePageItem({
      name: 'Guide.pdf',
      webUrl: 'https://uebt.sharepoint.com/sites/GroveGuidance/Shared%20Documents/Guide.pdf',
      fields: {
        FileLeafRef: 'Guide.pdf',
        FileRef: '/sites/GroveGuidance/Shared%20Documents/Guide.pdf',
        CanvasContent1: '<div>Not a page</div>',
      },
    }),
    false,
  );
});

test('isPublishedGraphSitePage only accepts published pages', () => {
  assert.equal(isPublishedGraphSitePage({ publishingState: { level: 'published' } }), true);
  assert.equal(isPublishedGraphSitePage({ publishingState: { level: 'draft' } }), false);
  assert.equal(isPublishedGraphSitePage({}), false);
});

test('renderCanvasLayoutHtml renders textWebPart innerHtml', () => {
  const html = renderCanvasLayoutHtml({
    horizontalSections: [
      {
        columns: [
          {
            webparts: [
              { '@odata.type': '#microsoft.graph.textWebPart', innerHtml: '<h2>About</h2><p>Details</p>' },
            ],
          },
        ],
      },
    ],
  });
  assert.match(html, /<h2>About<\/h2><p>Details<\/p>/);
});

test('renderCanvasLayoutHtml renders standardWebPart title and quick-link items', () => {
  const html = renderCanvasLayoutHtml({
    horizontalSections: [
      {
        columns: [
          {
            webparts: [
              {
                '@odata.type': '#microsoft.graph.standardWebPart',
                data: {
                  properties: {
                    titleHTML: '<h2>Popular Guidance</h2>',
                    items: [{ title: 'First Login' }],
                    serverProcessedContent: {
                      searchablePlainTexts: [{ key: 'items[0].title', value: 'First Login' }],
                      links: [{ key: 'items[0].sourceItem.url', value: '/sites/GroveGuidance/SitePages/FirstLogin.aspx' }],
                    },
                  },
                },
              },
            ],
          },
        ],
      },
    ],
  });
  assert.match(html, /<h2>Popular Guidance<\/h2>/);
  assert.match(html, /href="\/sites\/GroveGuidance\/SitePages\/FirstLogin\.aspx"/);
  assert.match(html, />First Login<\/a>/);
});

test('renderCanvasLayoutHtml renders verticalSection webparts', () => {
  const html = renderCanvasLayoutHtml({
    horizontalSections: [],
    verticalSection: {
      webparts: [
        { '@odata.type': '#microsoft.graph.textWebPart', innerHtml: '<p>Sidebar</p>' },
      ],
    },
  });
  assert.match(html, /<p>Sidebar<\/p>/);
});

test('renderCanvasLayoutHtml returns empty string for missing canvasLayout', () => {
  assert.equal(renderCanvasLayoutHtml(undefined), '');
});

test('normalizeGraphSitePage maps a Graph site page into the publisher page shape', () => {
  const page = normalizeGraphSitePage({
    id: 'e40f8f2c-ae50-4282-8883-690e425229d9',
    name: 'Home.aspx',
    title: 'Home',
    webUrl: 'https://uebt.sharepoint.com/sites/GroveGuidance/SitePages/Home.aspx',
    description: 'Welcome',
    thumbnailWebUrl: 'https://uebt.sharepoint.com/thumb.jpg',
    lastModifiedDateTime: '2026-09-04T10:00:00Z',
    createdDateTime: '2026-09-04T09:00:00Z',
    publishingState: { level: 'published' },
    canvasLayout: {
      horizontalSections: [
        {
          columns: [
            { webparts: [{ '@odata.type': '#microsoft.graph.textWebPart', innerHtml: '<p>Welcome</p>' }] },
          ],
        },
      ],
    },
  });

  assert.equal(page.Id, 'e40f8f2c-ae50-4282-8883-690e425229d9');
  assert.equal(page.Title, 'Home');
  assert.equal(page.FileLeafRef, 'Home.aspx');
  assert.equal(page.FileRef, '/sites/GroveGuidance/SitePages/Home.aspx');
  assert.equal(page.Description, 'Welcome');
  assert.equal(page.BannerImageUrl, 'https://uebt.sharepoint.com/thumb.jpg');
  assert.equal(page.PublishingLevel, 'published');
  assert.match(page.CanvasContent1, /<p>Welcome<\/p>/);
});

test('collectAssetCandidates excludes internal .aspx page links from asset downloads', () => {
  const candidates = collectAssetCandidates({
    FileRef: '/sites/GroveGuidance/SitePages/Home.aspx',
    CanvasContent1: '<ul>'
      + '<li><a href="/sites/GroveGuidance/SitePages/FirstLogin.aspx">First Login</a></li>'
      + '<li><a href="/sites/GroveGuidance/SitePages/FirstLogin.aspx?query=1">First Login query</a></li>'
      + '<img src="/sites/GroveGuidance/SiteAssets/banner.png" />'
      + '</ul>',
  });
  assert.deepEqual(candidates, ['/sites/GroveGuidance/SiteAssets/banner.png']);
});
