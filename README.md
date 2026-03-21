# Township Canada Excel Add-In

Microsoft Excel Add-In for converting Canadian legal land descriptions (DLS, NTS, Geographic Townships) to GPS coordinates.

## Features

- **Custom Functions**: Use `=TOWNSHIP.CONVERT()`, `=TOWNSHIP.LAT()`, `=TOWNSHIP.LNG()`, `=TOWNSHIP.PROVINCE()` in any cell
- **Ribbon Button**: One-click batch conversion of selected cells
- **Task Pane Sidebar**: Batch convert by selection or column, manage API key settings
- **Trial Keys**: Free 7-day trial (100 calls) at [townshipcanada.com/api/try](https://townshipcanada.com/api/try?ref=excel), unlimited with paid API key

## Custom Functions

| Function                              | Example             | Returns                    |
| ------------------------------------- | ------------------- | -------------------------- |
| `=TOWNSHIP.CONVERT("NW-25-24-1-W5")`  | DLS quarter section | `"52.123456, -114.654321"` |
| `=TOWNSHIP.LAT("NW-25-24-1-W5")`      | Latitude only       | `52.123456`                |
| `=TOWNSHIP.LNG("NW-25-24-1-W5")`      | Longitude only      | `-114.654321`              |
| `=TOWNSHIP.PROVINCE("NW-25-24-1-W5")` | Province name       | `"Alberta"`                |

Supports DLS (AB, SK, MB), NTS (BC), Geographic Townships (ON), River Lots, UWI, and FPS Grid formats.

## Development

```bash
# Install dependencies
npm install

# Start dev server (with HTTPS for Office.js)
npm run dev

# Build for production
npm run build

# Validate manifest
npm run validate

# Sideload into Excel for testing
npm run sideload
```

## Architecture

```
src/
├── shared/
│   └── config.js            # Shared API URL, storage helpers, API client
├── functions/
│   ├── functions.js         # Custom function implementations
│   ├── functions.json       # Custom function metadata (Office.js registration)
│   └── functions.html       # Functions runtime page
├── taskpane/
│   ├── taskpane.html        # Task pane UI (batch convert + settings)
│   ├── taskpane.js          # Task pane logic
│   └── taskpane.css         # Styles (Township brand)
└── commands/
    ├── commands.html         # Commands runtime page
    └── commands.js           # Ribbon button handlers
```

## API Backend

Uses the same Township Canada integration API as the Google Sheets add-on:

- `POST /api/integrations/trial/convert` — Single conversion
- `POST /api/integrations/trial/convert-batch` — Batch conversion (up to 200 items)
- `GET /api/integrations/trial/usage` — Usage quota check

Authentication via `X-API-Key` header (trial or paid key).

## AppSource Submission

1. Build: `npm run build`
2. Validate: `npm run validate`
3. Upload `dist/manifest.xml` and hosted assets to AppSource Partner Center
4. Review process takes 4-6 weeks
