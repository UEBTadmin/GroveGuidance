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

Preferred mode uses Microsoft Graph application permissions that can read SharePoint site pages and files (for example `Sites.Read.All`). If Graph site/list access is unauthorized, the publisher automatically falls back to SharePoint REST for page and asset synchronization so deployments can still proceed when SharePoint app-only permissions are available.

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
