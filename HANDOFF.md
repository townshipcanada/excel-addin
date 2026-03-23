# Handoff Notes

## Current State

- Project is stable on `main` branch, clean working tree
- CI workflow exists at `.github/workflows/ci.yml` (validate manifest + build)
- No test suite configured yet (`npm test` is a no-op)
- Build: `npm run build` → `dist/`

## Known Gaps

- No automated tests
- CI only validates manifest existence, doesn't run `npm run validate`
- No staging/production deployment pipeline
- No `.env` or environment variable management — API key is stored client-side via Office settings API

## Recent Work

- Renamed namespace from `TOWNSHIP` to `TOWNSHIP_CANADA` (commit c570f82)
- Added CI workflow and LICENSE (commit eae7dc7)
- Migrated from Webpack to Vite (commit 6085118)
- Added trial API key auth flow (commit dc78dce)
