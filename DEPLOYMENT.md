# Vercel deployment

This project now has a small Vercel web app around the original name card extractor.

## What users do

1. Open the Vercel URL.
2. Click **Connect Drive** and approve Google Drive read-only access.
3. Paste a Google Drive folder link.
4. Click **Sync Folder**.
5. Download the HubSpot-ready Excel export.

## What gets cached

Scans are stored in Supabase. The cache key is based on:

- Google Drive file ID
- `md5Checksum` when Drive provides it
- modified time and file size as fallback signals

That means the app skips cards already scanned, but a changed image can be scanned again and updates that Drive file's existing database row. Folder name alone is not used as the cache key because folder names can be renamed or duplicated.

## Required services

### Supabase

Yes, Supabase should be ready before deploying to Vercel. Vercel serverless functions do not have a durable local database, so a local JSON file will not work after deployment.

Run `supabase_schema.sql` in the Supabase SQL editor, then add these Vercel environment variables:

- `SUPABASE_URL`
- `SUPABASE_SERVICE_ROLE_KEY`

Keep `SUPABASE_SERVICE_ROLE_KEY` server-side only. Do not expose it in frontend JavaScript.

### Google Drive

You do not need a Google Drive plugin. You need a Google Cloud OAuth client and the Google Drive API enabled.

In Google Cloud:

1. Create or choose a Google Cloud project.
2. Enable **Google Drive API**.
3. Configure the OAuth consent screen.
4. Create an OAuth Client ID for a web application.
5. Add the redirect URI:
   - `https://YOUR-VERCEL-DOMAIN.vercel.app/api/oauth_callback`

Add these Vercel environment variables:

- `GOOGLE_CLIENT_ID`
- `GOOGLE_CLIENT_SECRET`
- `GOOGLE_REDIRECT_URI`

Set `GOOGLE_REDIRECT_URI` to the exact redirect URI above.

### OpenAI

Add:

- `OPENAI_API_KEY`
- `OPENAI_MODEL` optional, defaults to `gpt-4o`

### App secret

Add:

- `APP_SECRET`

Use a long random value. It signs the Google OAuth state parameter.

## Vercel notes

The sync API processes cards inside a serverless request. This is fine for small folders, but large folders can hit Vercel function time limits. If you expect many cards per folder, the next upgrade should be a queue/background worker so each image is processed separately.
