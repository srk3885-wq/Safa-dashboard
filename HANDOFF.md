# SAFA / SLS Incoming Shipments Dashboard Handoff

## Project Summary

This project is a lightweight web dashboard for SLS Distribution incoming shipments and inventory visibility. It is designed to replace the screenshot-style shipment summary with a cleaner, more professional dashboard that one inventory owner can update by uploading the raw shipment file or entering shipments manually.

The app currently runs as a small Node.js server with a static frontend. It stores shipment rows in a local JSON file and can be deployed online as long as the host provides persistent disk storage.

## Current Status

Completed:

- Minimal professional incoming-shipment dashboard.
- KPI cards for total inbound cases, next movement, air freight, and active transit.
- Product breakdown cards with flavor-level case bars.
- Shipment detail table.
- XLSX and CSV upload.
- Manual shipment entry form.
- Replace or merge import modes.
- Basic search, status, and freight filters.
- Shared backend save endpoint.
- Optional edit key through `ADMIN_PIN`.
- Clean dependency audit after switching from `xlsx` to `read-excel-file`.

Not yet done:

- Public deployment to a live domain.
- Real user accounts or role-based permissions.
- Database-backed production storage.
- Multi-user edit conflict handling beyond simple atomic JSON saves.

## File Map

- `index.html`: Main page structure.
- `styles.css`: Dashboard visual design and responsive layout.
- `app.js`: Frontend logic, parsing, filtering, metrics, manual entry, and rendering.
- `server.js`: Static file server and `/api/shipments` backend.
- `data/shipments.json`: Current stored shipment data.
- `package.json`: Node scripts and dependency list.
- `README.md`: Short run/deploy instructions.
- `.gitignore`: Local files to exclude from version control.

## How To Run Locally

Use Node.js 18 or newer.

```bash
npm install
npm start
```

Open:

```text
http://127.0.0.1:4173
```

The server uses port `4173` by default. To use another port:

```bash
PORT=5000 npm start
```

## Data Input Format

The importer was built around the raw file template:

```text
/Users/ai-home/Downloads/Shipment Status Template.xlsx
```

Expected columns:

- `Shipment Type`
- `Pick Up Date`
- `Brand`
- `Product Type`
- `Flavors`
- `Box Count`
- `Cargo Status`

The importer is forgiving about common header variations such as `Product`, `Cases`, `Status`, `Cartons`, and `Notes`.

The `Flavors` field supports line-by-line counts:

```text
White Gummy (new UI) 70ctns
Pineapple MTN Dew (new UI) 40ctns
Polar Ice (new UI) 40ctns
```

Those lines are split and used in the product card bars. If `Box Count` is missing, the dashboard totals the flavor lines.

## Data Model

Each shipment row is normalized to:

```json
{
  "id": "string",
  "shipmentType": "SEA or AIR",
  "pickupDate": "2026-06-01",
  "brand": "RAZ",
  "productType": "VUE POD",
  "flavors": "White Gummy 70ctns",
  "boxCount": 240,
  "cargoStatus": "VESSEL WILL DEPART ON 6/10 OR 6/11",
  "source": "Shipment Status Template.xlsx",
  "updatedAt": "2026-06-11T00:00:00.000Z"
}
```

Stored data lives in:

```text
data/shipments.json
```

## Key Logic

Import parsing:

- XLSX files are parsed in the browser with `read-excel-file`.
- CSV files are parsed by the frontend CSV parser in `app.js`.
- Only the first worksheet is imported.

Merge behavior:

- `Replace` overwrites all rows.
- `Merge` deduplicates rows using a fingerprint of shipment type, date, brand, product, flavors, box count, and status.

Status inference:

- The dashboard derives phase labels like `Scheduled`, `In Transit`, `Port Arrival`, `Trucking`, `Received`, and `Attention` from the shipment type and cargo status text.
- The relevant logic is in `inferPhase()` and `inferMode()` in `app.js`.

Date inference:

- The dashboard uses pickup date and status dates like `6/10` or `6/11`.
- If a year is missing, it uses the pickup-date year or current year.

## Backend API

Health:

```text
GET /api/health
```

Read shipments:

```text
GET /api/shipments
```

Save shipments:

```text
POST /api/shipments
Content-Type: application/json
```

Body:

```json
{
  "mode": "replace",
  "rows": []
}
```

Clear shipments:

```text
DELETE /api/shipments
```

If `ADMIN_PIN` is set, write requests must include:

```text
x-admin-pin: your-pin
```

## Online Deployment Notes

Recommended deployment options:

- Render
- Railway
- Fly.io
- VPS with Node.js and persistent disk

Required production settings:

- `NODE_ENV=production`
- `ADMIN_PIN=<private edit key>`
- Persistent disk mounted for `data/shipments.json`, or set `DATA_DIR` to the persistent storage path.

Example:

```bash
NODE_ENV=production ADMIN_PIN=change-this DATA_DIR=/data npm start
```

Important: do not deploy this to a serverless platform without persistent storage unless the save layer is changed to a database. Otherwise uploads may disappear between deployments or restarts.

## Security Notes

Current protection is intentionally simple:

- Anyone with the link can view the dashboard.
- Only users with `ADMIN_PIN` can upload, manually add, replace, merge, or clear data.

Before wider company use, consider:

- Real login/authentication.
- HTTPS-only deployment.
- Strong edit key stored as a host environment variable.
- Database storage with backups.
- Upload size limits and better file validation for production.

## Verification Already Performed

Performed during build:

- `node --check server.js`
- `node --check app.js`
- `npm audit --json`
- Parsed the provided XLSX template with `read-excel-file`.
- Verified the dashboard loads in the browser.
- Verified the Excel reader is ready in the page.
- Tested manual entry save and dashboard recalculation.
- Restored seed data after testing.
- Checked desktop and mobile layouts.

## Known Limitations

- Supports `.xlsx` and `.csv`; legacy `.xls` is not supported.
- Imports only the first worksheet.
- Data is stored in a JSON file, not a database.
- Simultaneous edits from multiple users are not deeply conflict-managed.
- Phase and freight status are inferred from text keywords, so unusual cargo-status wording may need new rules in `app.js`.

## Recommended Next Steps

1. Confirm the real ongoing raw data file from the inventory team.
2. Test 5-10 real shipment files and add any missing header aliases.
3. Pick the hosting target.
4. Deploy with persistent storage and `ADMIN_PIN`.
5. Add a simple backup/export routine for `data/shipments.json`.
6. If this becomes daily-critical, move storage from JSON to Postgres or SQLite.

## Maintenance Pointers

When raw file columns change:

- Update `mapHeader()` in `app.js`.

When status wording changes:

- Update `inferPhase()` and `inferMode()` in `app.js`.

When moving to a database:

- Replace `readData()` and `saveData()` in `server.js`.
- Keep the frontend data shape the same to avoid UI rewrites.
