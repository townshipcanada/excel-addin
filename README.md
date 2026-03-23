# Township Canada Excel Add-In

Microsoft Excel Add-In for converting Canadian legal land descriptions (DLS, NTS, Geographic Townships) to GPS coordinates.

## Features

- **Custom Functions**: Use `=TOWNSHIP_CANADA.CONVERT()`, `=TOWNSHIP_CANADA.LAT()`, `=TOWNSHIP_CANADA.LNG()`, `=TOWNSHIP_CANADA.PROVINCE()` in any cell
- **Ribbon Button**: One-click batch conversion of selected cells
- **Task Pane Sidebar**: Batch convert by selection or column, manage API key settings
- **Trial Keys**: Free 7-day trial (100 calls) at [townshipcanada.com/api/try](https://townshipcanada.com/api/try?ref=excel), unlimited with paid API key

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
| `POST` | `/batch/legal-location` | Batch conversion (up to 200 items) |
| `GET` | `/usage` | Usage quota check |

**Paid keys** (`tc_...`) use `https://developer.townshipcanada.com` with the same endpoint paths.

## AppSource Submission

1. Build: `npm run build`
2. Validate: `npm run validate`
3. Upload `dist/manifest.xml` and hosted assets to AppSource Partner Center
