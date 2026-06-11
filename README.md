# SAFA Dashboard

Incoming shipments + warehouse inventory dashboard. Drag-and-drop Excel upload, manual entries, low-stock alerts.

## Project layout

| File | What it is |
|---|---|
| `public/index.html` | Page structure (header, tabs, modals) |
| `public/styles.css` | All styling — colors, cards, tables |
| `public/app.js` | All logic — Excel parsing, views, manual entries |
| `server.js` | Tiny web server (Railway runs this) |

## Run locally

```
npm install
npm start          # opens on http://localhost:3000
```

## Deploy to Railway (first time)

1. Push this folder to a GitHub repo (github.com → New repository → upload these files).
2. Go to railway.app → New Project → **Deploy from GitHub repo** → pick the repo.
3. Railway auto-detects Node, runs `npm start`, and gives you a public URL
   (Settings → Networking → **Generate Domain**).

## Make changes later

Edit the file, commit/push to GitHub — Railway redeploys automatically in ~1 min.

Common edits:
- **Company name / title**: search "SAFA" in `public/index.html`
- **Colors**: `:root` variables at top of `public/styles.css`
- **Low-stock default**: `lowStockThreshold:50` in `public/app.js`
- **Column name aliases** (if the Excel headers change): `HEADER_ALIASES` / `INV_ALIASES` in `public/app.js`
- **Status keywords** (how cargo text maps to sea/air/port/truck): `deriveStatus()` in `public/app.js`

## Notes

- Data is stored in each user's browser (localStorage). The uploaded Excel is the source of truth — re-upload to refresh. Use Backup/Restore buttons to move data between computers.
