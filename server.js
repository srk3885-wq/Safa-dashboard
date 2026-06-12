const http = require("node:http");
const fs = require("node:fs/promises");
const path = require("node:path");
const crypto = require("node:crypto");

const ROOT = __dirname;
const PORT = Number(process.env.PORT || 4173);
const HOST = process.env.HOST || (process.env.NODE_ENV === "production" ? "0.0.0.0" : "127.0.0.1");
const ADMIN_PIN = process.env.ADMIN_PIN || "";
const DASHBOARD_PASSWORD = process.env.DASHBOARD_PASSWORD || "";
const AUTH_SECRET =
  process.env.SESSION_SECRET ||
  crypto.createHash("sha256").update(`safa::${DASHBOARD_PASSWORD}`).digest("hex");
const AUTH_COOKIE = "safa_auth";
const AUTH_MAX_AGE_DAYS = 30;
const DATA_DIR = process.env.DATA_DIR
  ? path.resolve(process.env.DATA_DIR)
  : path.join(ROOT, "data");
const DATA_FILE = path.join(DATA_DIR, "shipments.json");
const MAX_BODY_BYTES = 5 * 1024 * 1024;

const staticFiles = new Map([
  ["/", { file: "index.html", type: "text/html; charset=utf-8" }],
  ["/index.html", { file: "index.html", type: "text/html; charset=utf-8" }],
  ["/styles.css", { file: "styles.css", type: "text/css; charset=utf-8" }],
  ["/app.js", { file: "app.js", type: "text/javascript; charset=utf-8" }],
  [
    "/vendor/read-excel-file.min.js",
    {
      file: path.join(ROOT, "node_modules/read-excel-file/bundle/read-excel-file.min.js"),
      type: "text/javascript; charset=utf-8"
    }
  ]
]);

function sendJson(res, status, payload) {
  res.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store"
  });
  res.end(JSON.stringify(payload));
}

function sendText(res, status, message) {
  res.writeHead(status, { "Content-Type": "text/plain; charset=utf-8" });
  res.end(message);
}

async function ensureDataFile() {
  await fs.mkdir(DATA_DIR, { recursive: true });
  try {
    await fs.access(DATA_FILE);
  } catch {
    await saveData({ rows: [], lastUpdated: null });
  }
}

async function readData() {
  await ensureDataFile();
  try {
    const text = await fs.readFile(DATA_FILE, "utf8");
    const parsed = JSON.parse(text);
    return {
      rows: Array.isArray(parsed.rows) ? parsed.rows.map(sanitizeRow).filter(Boolean) : [],
      lastUpdated: parsed.lastUpdated || null
    };
  } catch {
    return { rows: [], lastUpdated: null };
  }
}

async function saveData(data) {
  await fs.mkdir(DATA_DIR, { recursive: true });
  const payload = JSON.stringify(data, null, 2);
  const tmp = `${DATA_FILE}.tmp`;
  await fs.writeFile(tmp, payload);
  await fs.rename(tmp, DATA_FILE);
}

function parseBody(req) {
  return new Promise((resolve, reject) => {
    let size = 0;
    const chunks = [];

    req.on("data", (chunk) => {
      size += chunk.length;
      if (size > MAX_BODY_BYTES) {
        reject(new Error("Payload is too large."));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });

    req.on("end", () => {
      try {
        const text = Buffer.concat(chunks).toString("utf8");
        resolve(text ? JSON.parse(text) : {});
      } catch {
        reject(new Error("Invalid JSON."));
      }
    });

    req.on("error", reject);
  });
}

function canWrite(req) {
  if (!ADMIN_PIN) return true;
  return req.headers["x-admin-pin"] === ADMIN_PIN;
}

/* ---------- view password (login page) ---------- */
function expectedAuthToken() {
  return crypto.createHmac("sha256", AUTH_SECRET).update(`ok::${DASHBOARD_PASSWORD}`).digest("hex");
}

function getCookie(req, name) {
  const header = req.headers.cookie || "";
  const match = header.match(new RegExp(`(?:^|;\\s*)${name}=([^;]*)`));
  return match ? decodeURIComponent(match[1]) : null;
}

function isAuthed(req) {
  if (!DASHBOARD_PASSWORD) return true; // no password configured -> open
  const token = getCookie(req, AUTH_COOKIE) || "";
  const expected = expectedAuthToken();
  if (token.length !== expected.length) return false;
  return crypto.timingSafeEqual(Buffer.from(token), Buffer.from(expected));
}

function readRawBody(req) {
  return new Promise((resolve, reject) => {
    let size = 0;
    const chunks = [];
    req.on("data", (chunk) => {
      size += chunk.length;
      if (size > 64 * 1024) {
        reject(new Error("Payload is too large."));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });
    req.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    req.on("error", reject);
  });
}

function loginPage(error) {
  return `<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>SAFA — Sign in</title><style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;background:#f7f7f8;color:#0e0e10;display:flex;align-items:center;justify-content:center;min-height:100vh;-webkit-font-smoothing:antialiased}
.card{background:#fff;border:1px solid #e6e6ea;border-radius:16px;padding:40px 36px;width:94vw;max-width:380px;box-shadow:0 1px 2px rgba(14,14,16,.04),0 8px 32px rgba(14,14,16,.07)}
h1{font-size:20px;font-weight:800;letter-spacing:-.02em}
h1 span{color:#8a8a93;font-weight:600}
p{color:#8a8a93;font-size:13px;margin:6px 0 24px}
label{display:block;font-size:11px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:#8a8a93;margin-bottom:6px}
input{width:100%;border:1px solid #e6e6ea;border-radius:10px;padding:12px 14px;font-size:16px;outline:none;font-family:inherit}
input:focus{border-color:#0e0e10}
button{width:100%;margin-top:14px;background:#0e0e10;color:#fff;border:none;border-radius:10px;padding:12px;font-size:14px;font-weight:700;cursor:pointer;font-family:inherit}
button:hover{background:#26262b}
.err{background:#fee4e2;color:#b42318;border-radius:8px;padding:9px 12px;font-size:12.5px;font-weight:600;margin-bottom:14px}
</style></head><body>
<form class="card" method="POST" action="/login">
  <h1>SAFA <span>/ Incoming Shipments</span></h1>
  <p>Enter the team password to continue.</p>
  ${error ? '<div class="err">Wrong password — try again.</div>' : ""}
  <label for="pw">Password</label>
  <input id="pw" name="password" type="password" autofocus autocomplete="current-password">
  <button type="submit">Sign in</button>
</form></body></html>`;
}

function sendLoginPage(res, status, error) {
  res.writeHead(status, { "Content-Type": "text/html; charset=utf-8", "Cache-Control": "no-store" });
  res.end(loginPage(error));
}

async function handleAuthRoutes(req, res, pathname) {
  if (pathname === "/login" && req.method === "GET") {
    if (isAuthed(req)) {
      res.writeHead(302, { Location: "/" });
      res.end();
    } else {
      sendLoginPage(res, 200, false);
    }
    return true;
  }

  if (pathname === "/login" && req.method === "POST") {
    const body = await readRawBody(req);
    const given = String(new URLSearchParams(body).get("password") || "");
    const a = crypto.createHash("sha256").update(given).digest();
    const b = crypto.createHash("sha256").update(DASHBOARD_PASSWORD).digest();
    if (DASHBOARD_PASSWORD && crypto.timingSafeEqual(a, b)) {
      const secure =
        req.headers["x-forwarded-proto"] === "https" ? "; Secure" : "";
      res.writeHead(302, {
        Location: "/",
        "Set-Cookie": `${AUTH_COOKIE}=${expectedAuthToken()}; Path=/; HttpOnly; SameSite=Lax; Max-Age=${AUTH_MAX_AGE_DAYS * 86400}${secure}`
      });
      res.end();
    } else {
      sendLoginPage(res, 401, true);
    }
    return true;
  }

  if (pathname === "/logout") {
    res.writeHead(302, {
      Location: "/login",
      "Set-Cookie": `${AUTH_COOKIE}=; Path=/; HttpOnly; Max-Age=0`
    });
    res.end();
    return true;
  }

  return false;
}

function asText(value) {
  return String(value ?? "").trim();
}

function asNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const match = asText(value).replace(/,/g, "").match(/-?\d+(\.\d+)?/);
  return match ? Number(match[0]) : 0;
}

function sanitizeRow(row) {
  if (!row || typeof row !== "object") return null;

  const clean = {
    id: asText(row.id) || crypto.randomUUID(),
    shipmentType: asText(row.shipmentType),
    pickupDate: asText(row.pickupDate),
    brand: asText(row.brand),
    productType: asText(row.productType),
    flavors: asText(row.flavors),
    boxCount: asNumber(row.boxCount),
    cargoStatus: asText(row.cargoStatus),
    source: asText(row.source) || "Dashboard",
    updatedAt: asText(row.updatedAt) || new Date().toISOString()
  };

  if (!clean.brand && !clean.productType && !clean.cargoStatus && !clean.boxCount) {
    return null;
  }

  return clean;
}

function fingerprint(row) {
  return [
    row.shipmentType,
    row.pickupDate,
    row.brand,
    row.productType,
    row.flavors,
    row.boxCount,
    row.cargoStatus
  ]
    .map((part) => asText(part).toLowerCase())
    .join("|");
}

function mergeRows(existingRows, incomingRows) {
  const merged = new Map();

  for (const row of existingRows) {
    merged.set(fingerprint(row), row);
  }

  for (const row of incomingRows) {
    const key = fingerprint(row);
    const existing = merged.get(key);
    merged.set(key, {
      ...row,
      id: existing?.id || row.id,
      updatedAt: new Date().toISOString()
    });
  }

  return Array.from(merged.values());
}

async function serveStatic(req, res, pathname) {
  const staticFile = staticFiles.get(pathname);
  if (!staticFile) {
    sendText(res, 404, "Not found");
    return;
  }

  try {
    const filePath = path.isAbsolute(staticFile.file)
      ? staticFile.file
      : path.join(ROOT, staticFile.file);
    const body = await fs.readFile(filePath);
    res.writeHead(200, {
      "Content-Type": staticFile.type,
      "Cache-Control": "no-cache"
    });
    res.end(body);
  } catch {
    sendText(res, 500, "File could not be loaded.");
  }
}

async function handleApi(req, res, pathname) {
  if (pathname === "/api/health") {
    sendJson(res, 200, { ok: true, shared: true, protected: Boolean(ADMIN_PIN) });
    return;
  }

  if (pathname === "/api/shipments" && req.method === "GET") {
    const data = await readData();
    sendJson(res, 200, { ...data, protected: Boolean(ADMIN_PIN) });
    return;
  }

  if (pathname === "/api/shipments" && req.method === "POST") {
    if (!canWrite(req)) {
      sendJson(res, 401, { error: "Edit key is required." });
      return;
    }

    const body = await parseBody(req);
    const incomingRows = Array.isArray(body.rows)
      ? body.rows.map(sanitizeRow).filter(Boolean)
      : [];
    const mode = body.mode === "merge" ? "merge" : "replace";
    const current = await readData();
    const rows = mode === "merge" ? mergeRows(current.rows, incomingRows) : incomingRows;
    const next = { rows, lastUpdated: new Date().toISOString() };

    await saveData(next);
    sendJson(res, 200, next);
    return;
  }

  if (pathname === "/api/shipments" && req.method === "DELETE") {
    if (!canWrite(req)) {
      sendJson(res, 401, { error: "Edit key is required." });
      return;
    }

    const next = { rows: [], lastUpdated: new Date().toISOString() };
    await saveData(next);
    sendJson(res, 200, next);
    return;
  }

  sendJson(res, 404, { error: "Unknown API route." });
}

const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url, `http://${req.headers.host || "localhost"}`);
    const pathname = url.pathname;

    if (await handleAuthRoutes(req, res, pathname)) {
      return;
    }

    if (!isAuthed(req)) {
      if (pathname.startsWith("/api/")) {
        sendJson(res, 401, { error: "Sign in required." });
      } else {
        res.writeHead(302, { Location: "/login" });
        res.end();
      }
      return;
    }

    if (pathname.startsWith("/api/")) {
      await handleApi(req, res, pathname);
      return;
    }

    await serveStatic(req, res, pathname);
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Server error." });
  }
});

server.listen(PORT, HOST, () => {
  const displayHost = HOST === "0.0.0.0" ? "localhost" : HOST;
  console.log(`SAFA inventory dashboard running at http://${displayHost}:${PORT}`);
});
