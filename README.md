# SLS Distribution Incoming Shipments

A minimal web dashboard for incoming goods, inventory, in-transit shipments, and manual shipment entries.

## Run locally

```bash
npm install
npm start
```

Open `http://localhost:4173`.

## Raw data format

The uploader supports `.xlsx` and `.csv` files with columns like:

- `Shipment Type`
- `Pick Up Date`
- `Brand`
- `Product Type`
- `Flavors`
- `Box Count`
- `Cargo Status`

The `Flavors` column can include multiple lines such as `White Gummy 70ctns`; the dashboard will split and total those lines automatically.

## Put it online

Install dependencies, then deploy it as a small Node app with:

```bash
npm install
npm start
```

For a public team link, set an `ADMIN_PIN` environment variable on the host. Viewers can see the dashboard, and the inventory person enters the same key before uploading or editing data.

Recommended hosts: Render, Railway, Fly.io, or any VPS with persistent disk enabled. If you deploy to a serverless platform without persistent storage, connect the save endpoint to a database first.
