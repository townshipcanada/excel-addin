# Township Canada Excel Add-In

Microsoft Excel Add-In that converts Canadian legal land descriptions (township/range/section/meridian) to GPS coordinates.

## Tech Stack

- **Runtime**: Office.js (Excel JavaScript API)
- **Build**: Vite 6 with `vite-plugin-static-copy`
- **Language**: Vanilla JavaScript (no TypeScript, no framework)
- **Manifest**: Office XML manifest (`manifest.xml`)

## Project Structure

- `src/taskpane/` — Main UI panel (taskpane.js, taskpane.css)
- `src/functions/` — Excel custom functions (functions.js, functions.json)
- `src/commands/` — Ribbon commands (commands.js)
- `src/shared/config.js` — Shared configuration (API keys, endpoints)
- `src/shared/sampleData.js` — 100 hardcoded DLS entries for offline demo
- `taskpane.html`, `functions.html`, `commands.html` — Entry points
- `manifest.xml` — Office Add-In manifest

## Key Commands

- `npm run dev` — Start dev server (port 3000, HTTPS)
- `npm run build` — Production build to `dist/`
- `npm run validate` — Validate manifest.xml
- `npm run sideload` — Install certs and sideload into Excel

## Architecture Notes

- API key auth: users provide a trial or paid API key via Settings tab
- Offline demo: 100 sample DLS locations work without an API key (sampleData.js checked before API)
- Namespace: `TOWNSHIP_CANADA` (custom functions prefix)
- The add-in calls an external API to resolve land descriptions to coordinates
