# Daxus LATAM to Notion Agent

This agent signs in to Daxus LATAM, extracts available course data with Playwright, and syncs it into a Notion database.

## GitHub Actions schedule

The workflow runs every Monday at 14:00 UTC and skips execution unless that Monday is the last Monday of the month. In America/Bogota, that is 9:00 AM.

It can also be run manually from the GitHub Actions tab with `workflow_dispatch`.

## Required GitHub secrets

Configure these repository secrets before enabling the workflow:

- `DAXUS_EMAIL`
- `DAXUS_PASSWORD`
- `NOTION_API_KEY`
- `NOTION_DATABASE_ID`

The local `.env` file is intentionally ignored and must not be committed.

## Local checks

```bash
npm ci
npm run check
```
