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

`SP_TENANT_HOST` and `SP_SITE_PATH` can also be provided as full SharePoint URLs; the publisher normalizes these values automatically.

## Azure AD app permissions

The Azure AD app used by `SP_CLIENT_ID` must have SharePoint application permissions allowing read access to site pages and files (via `https://{tenant}.sharepoint.com/.default` token scope).

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
