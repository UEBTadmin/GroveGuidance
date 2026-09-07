# GroveGuidance

SharePoint-to-static cross-publishing pipeline for Grove Guidance.

This repository now runs an automated pipeline that pulls published pages from the SharePoint Communications Site and deploys them to Azure Static Web Apps.

## What the pipeline does

1. Extracts published pages from `Site Pages` in SharePoint.
2. Transforms page content into static HTML routes.
3. Mirrors referenced SharePoint-hosted assets (images/files/media/css/js) into `/assets`.
4. Rewrites internal links to the public static routes.
5. Produces static output in `/dist`:
   - page routes (`/index.html`, `/some-page/index.html`, ...)
   - `sitemap.xml`
   - `robots.txt`
   - `staticwebapp.config.json`
   - publish manifest (`.publish-manifest.json`)
6. Validates route uniqueness, HTML presence, internal links, and local asset references before deployment.
7. Deploys to Azure Static Web Apps.

## Matching SharePoint's visual theme

The static output does not have access to SharePoint's own stylesheet/JS (it's proprietary, auth-gated, and only rendered client-side by SharePoint's own SPA at runtime). Instead, `buildPageHtml()` in `scripts/sharepoint-publish.mjs` uses a small set of theme tokens (colors, fonts, corner radii, button/overlay styles) loaded from `scripts/sharepoint-theme.json` via `scripts/sharepoint-theme.mjs`, falling back to reasonable defaults if that file is absent.

To (re)capture the real values from a live, authenticated SharePoint page:

```
npm run theme:capture
```

- First run: opens a visible Chromium window and pauses for you to sign in interactively (including MFA). Your session is then saved to `.cache/sharepoint-auth-state.json` (gitignored) and reused headlessly on future runs until it expires.
- Writes the captured values to `scripts/sharepoint-theme.json`, which is committed to the repo so CI/CD always builds with the latest captured theme.
- Optionally pass a specific page URL to capture from, e.g. `node scripts/capture-sharepoint-theme.mjs "https://uebt.sharepoint.com/sites/GroveGuidance/SitePages/Home.aspx"`.

Whenever SharePoint's site theme/branding changes, re-run `npm run theme:capture`, review the diff in `scripts/sharepoint-theme.json`, and commit it — no manual CSS editing is required for the static site to follow along.

## Required GitHub secrets

Add these repository secrets:

- `SP_TENANT_ID`: Azure AD tenant ID.
- `SP_CLIENT_ID`: Azure AD app (client) ID.
- `SP_CLIENT_SECRET`: Azure AD app client secret.
- `SP_TENANT_HOST`: SharePoint host (for example `uebt.sharepoint.com`).
- `SP_SITE_PATH`: Site path (for example `/sites/GroveGuidance`).
- `AZURE_STATIC_WEB_APPS_API_TOKEN_VICTORIOUS_STONE_0501FF610`: Azure Static Web Apps deployment token.

`SP_TENANT_HOST` and `SP_SITE_PATH` can also be provided as full SharePoint URLs (including `SitePages`, library `Forms/AllItems.aspx`, and `/:w:/r/...` sharing-link formats); the publisher normalizes these values automatically to the SharePoint host and site root/subweb path.

## Azure AD app permissions

The publisher requests both:

- `https://graph.microsoft.com/.default` (Graph token)
- `https://{tenant-host}/.default` (SharePoint token)

Preferred mode uses Microsoft Graph application permissions that can read SharePoint site pages and files (for example `Sites.Read.All`, or `Sites.Selected` with a site-level grant). Grant admin consent after adding permissions.

If Graph site/list access is unauthorized, the publisher attempts SharePoint REST fallback for page and asset synchronization. Some SharePoint sites reject client-secret app-only REST tokens with `Unsupported app only token.`; in that case update Entra Graph application permissions/site grants (or switch to certificate-based SharePoint app-only auth if you require REST fallback).

## Workflow

Workflow file:

`/home/runner/work/GroveGuidance/GroveGuidance/.github/workflows/azure-static-web-apps-victorious-stone-0501ff610.yml`

Triggers:

- Push to `main`
- Scheduled every 6 hours
- Manual run (`workflow_dispatch`)

Manual pilot runs can pass `pilot_page_limit` to process only a subset of pages.

## Local run

```bash
npm run publish:sync
npm run publish:validate
```

Environment variables required locally match the secrets listed above.

## One-time WordPress migration export

`scripts/sharepoint-wordpress-export.mjs` is a standalone tool (not part of the ongoing sync
pipeline above) for migrating the published pages off SharePoint entirely, e.g. to a
self-hosted WordPress site. It reuses the same Graph Pages API access as the sync pipeline to
produce a WordPress WXR import file plus downloaded media assets:

```bash
npm run wordpress:export
```

Required environment variables: `SP_TENANT_ID`, `SP_CLIENT_ID`, `SP_CLIENT_SECRET` (same as
the sync pipeline). Optional: `SP_TENANT_HOST`, `SP_SITE_PATH`, `PUBLIC_BASE_URL`,
`WP_SITE_TITLE`, `WP_EXPORT_DIR` (default `wp-export`), `WP_MEDIA_SUBDIR` (subfolder under
`wp-content/uploads/` that the media will be served from; default: none, i.e. the uploads root).

Output (under `WP_EXPORT_DIR`):
- `wordpress-export.xml` - WXR file; import via WordPress **Tools > Import > WordPress**.
- `media/` - every referenced image/asset, downloaded and renamed to match the
  `<wp:attachment_url>` placeholders baked into the XML. Upload these so they are served from
  `wp-content/uploads/` (or the `WP_MEDIA_SUBDIR` subfolder, if set) to match the URLs already
  rewritten into the imported page content. The WordPress importer fetches each attachment URL
  over HTTP at import time, so the files must be in place **before** running the import.
- `navigation.json` - a simple label/slug list derived from the published pages, since WXR
  does not reliably carry SharePoint's QuickLaunch navigation structure. Recreate this menu
  manually under WordPress **Appearance > Menus**.

This is intended to be run once as a migration step, not on a schedule - after import,
WordPress (not SharePoint) becomes the source of truth for page content.
