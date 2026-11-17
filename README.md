# Township Canada Excel Add-In

Microsoft Excel Add-In for converting Canadian legal land descriptions (DLS, NTS, Geographic Townships) to GPS coordinates.

## Features

- **Custom Functions**: Use `=TOWNSHIP_CANADA.CONVERT()`, `=TOWNSHIP_CANADA.LAT()`, `=TOWNSHIP_CANADA.LNG()`, `=TOWNSHIP_CANADA.PROVINCE()` in any cell
- **Ribbon Button**: One-click batch conversion of selected cells
- **Task Pane Sidebar**: Batch convert by selection or column, manage API key settings
- **Try Offline**: 100 sample DLS locations work instantly without an API key
- **API Keys**: Get full access by choosing a plan on the [API plans page](https://townshipcanada.com/pricing?ref=excel-addin#api-plans), then connect the key in the task pane

## Custom Functions

| Function | Returns | Example |
| --- | --- | --- |
| `=TOWNSHIP_CANADA.CONVERT("NW-25-24-1-W5")` | `"52.123456, -114.654321"` | GPS coordinates as text |
| `=TOWNSHIP_CANADA.LAT("NW-25-24-1-W5")` | `52.123456` | Latitude only |
| `=TOWNSHIP_CANADA.LNG("NW-25-24-1-W5")` | `-114.654321` | Longitude only |
| `=TOWNSHIP_CANADA.PROVINCE("NW-25-24-1-W5")` | `"Alberta"` | Province name |

Supported formats: DLS (AB, SK, MB), NTS (BC), Geographic Townships (ON), River Lots, UWI, and FPS Grid.

## Development

```bash
npm install        # Install dependencies
npm run dev        # Start dev server (HTTPS for Office.js)
npm run build      # Build for production
npm run validate   # Validate manifest
npm run sideload   # Sideload into Excel for testing
```

## Architecture

```
src/
├── shared/config.js          # API URLs, storage helpers, API client
├── functions/
│   ├── functions.js           # Custom function implementations
│   └── functions.json         # Custom function metadata (Office.js registration)
├── taskpane/
│   ├── taskpane.html          # Task pane UI (batch convert + settings)
│   ├── taskpane.js            # Task pane logic
│   └── taskpane.css           # Styles
└── commands/commands.js       # Ribbon button handlers
```

## API Endpoints

Authentication via `X-API-Key` header (trial or paid key).

**Trial keys** (`tc_trial_...`) use `https://townshipcanada.com/api/integrations/trial`:

| Method | Endpoint | Purpose |
| --- | --- | --- |
| `GET` | `/search/legal-location?location={query}` | Single conversion |
| `POST` | `/batch/legal-location` | Batch conversion (up to 100 items per request) |
| `GET` | `/usage` | Usage quota check |

**Paid keys** (`tc_...`) use `https://developer.townshipcanada.com` with the same endpoint paths.

## Hosting

The add-in is served from `https://townshipcanada.com/integrations/excel/`, the web app's `public/integrations/excel/` folder. Every URL in `manifest.xml` points there, and `npm test` checks that each one exists in `dist/`.

To publish a change:

1. `npm run publish:web` builds and copies `dist/` into `../web/public/integrations/excel/` (set `WEB_DIR` if the web checkout is elsewhere).
2. Commit that folder in the web repo and merge it to `main`; Vercel deploys it.
3. Excel picks up the new files on the next load. A manifest change also needs the updated `manifest.xml` resubmitted or redistributed.

`npm run build` makes two passes: the pages (task pane and ribbon commands, with relative asset paths) and a standalone `functions.js` (one classic script, no imports) for the manifest's custom functions script URL. Custom functions run in the shared runtime, so the task pane page also registers them.

## AppSource Submission

1. Build: `npm run build`
2. Validate: `npm run validate`
3. Upload `dist/manifest.xml` and hosted assets to AppSource Partner Center
