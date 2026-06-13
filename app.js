const STORE_KEY = "safa-inventory-rows-v1";
const PIN_KEY = "safa-inventory-edit-key-v1";
const INV_STORE_KEY = "safa-warehouse-items-v1";
const THRESHOLD_KEY = "safa-low-threshold-v1";
const ARCHIVE_AFTER_DAYS = 14;

const SAMPLE_ROWS = [
  {
    id: "template-raz-vue-pod-2026-06-01",
    shipmentType: "SEA or AIR",
    pickupDate: "2026-06-01",
    brand: "RAZ",
    productType: "VUE POD",
    flavors:
      "White Gummy (new UI) 70ctns\nPineapple MTN Dew (new UI) 40ctns\nPolar Ice (new UI) 40ctns\nTriple Berry Lime (new UI) 80ctns\nWhite Gummy (new UI) 10ctns",
    boxCount: 240,
    cargoStatus: "VESSEL WILL DEPART ON 6/10 OR 6/11",
    source: "Shipment Status Template.xlsx"
  }
];

const state = {
  rows: [],
  inventory: [],
  snapshots: [],
  showArchive: false,
  threshold: Number(localStorage.getItem("safa-low-threshold-v1")) || 50,
  serverOnline: false,
  importMode: "replace",
  filters: {
    query: "",
    phase: "all",
    mode: "all"
  }
};

const els = {};

document.addEventListener("DOMContentLoaded", init);

async function init() {
  cacheElements();
  markExcelReaderState();
  bindEvents();

  els.adminPin.value = localStorage.getItem(PIN_KEY) || "";
  await loadInitialData();
  render();
}

function markExcelReaderState() {
  document.documentElement.dataset.excelReader = window.readXlsxFile ? "ready" : "missing";
}

function cacheElements() {
  Object.assign(els, {
    metrics: document.querySelector("#metrics"),
    productGrid: document.querySelector("#productGrid"),
    shipmentRows: document.querySelector("#shipmentRows"),
    rowCount: document.querySelector("#rowCount"),
    asOfDate: document.querySelector("#asOfDate"),
    connectionBadge: document.querySelector("#connectionBadge"),
    saveState: document.querySelector("#saveState"),
    fileInput: document.querySelector("#fileInput"),
    fileDrop: document.querySelector("#fileDrop"),
    adminPin: document.querySelector("#adminPin"),
    boardReport: document.querySelector("#boardReport"),
    exportJson: document.querySelector("#exportJson"),
    clearData: document.querySelector("#clearData"),
    manualForm: document.querySelector("#manualForm"),
    searchInput: document.querySelector("#searchInput"),
    phaseFilter: document.querySelector("#phaseFilter"),
    modeFilter: document.querySelector("#modeFilter"),
    exportCsv: document.querySelector("#exportCsv"),
    archiveToggle: document.querySelector("#archiveToggle"),
    calendarGrid: document.querySelector("#calendarGrid"),
    calRange: document.querySelector("#calRange"),
    invFileInput: document.querySelector("#invFileInput"),
    invFileDrop: document.querySelector("#invFileDrop"),
    invThreshold: document.querySelector("#invThreshold"),
    invForm: document.querySelector("#invForm"),
    invRows: document.querySelector("#invRows"),
    invCount: document.querySelector("#invCount"),
    importDialog: document.querySelector("#importDialog"),
    importSummary: document.querySelector("#importSummary"),
    importMerge: document.querySelector("#importMerge"),
    importReplace: document.querySelector("#importReplace"),
    importCancel: document.querySelector("#importCancel")
  });
}

function bindEvents() {
  els.fileInput.addEventListener("change", () => {
    const file = els.fileInput.files?.[0];
    if (file) handleFile(file);
    els.fileInput.value = "";
  });

  els.fileDrop.addEventListener("dragover", (event) => {
    event.preventDefault();
    els.fileDrop.classList.add("is-dragging");
  });

  els.fileDrop.addEventListener("dragleave", () => {
    els.fileDrop.classList.remove("is-dragging");
  });

  els.fileDrop.addEventListener("drop", (event) => {
    event.preventDefault();
    els.fileDrop.classList.remove("is-dragging");
    const file = event.dataTransfer.files?.[0];
    if (file) handleFile(file);
  });

  document.querySelectorAll("[data-import-mode]").forEach((button) => {
    button.addEventListener("click", () => {
      state.importMode = button.dataset.importMode;
      document.querySelectorAll("[data-import-mode]").forEach((item) => {
        item.classList.toggle("active", item === button);
      });
    });
  });

  els.adminPin.addEventListener("input", () => {
    localStorage.setItem(PIN_KEY, els.adminPin.value);
  });

  els.boardReport.addEventListener("click", openBoardReport);
  els.exportCsv.addEventListener("click", exportCsv);
  els.exportJson.addEventListener("click", exportJson);

  els.archiveToggle.addEventListener("click", () => {
    state.showArchive = !state.showArchive;
    render();
  });

  els.invFileInput.addEventListener("change", () => {
    const file = els.invFileInput.files?.[0];
    if (file) handleInventoryFile(file);
    els.invFileInput.value = "";
  });
  els.invFileDrop.addEventListener("dragover", (event) => {
    event.preventDefault();
    els.invFileDrop.classList.add("is-dragging");
  });
  els.invFileDrop.addEventListener("dragleave", () => els.invFileDrop.classList.remove("is-dragging"));
  els.invFileDrop.addEventListener("drop", (event) => {
    event.preventDefault();
    els.invFileDrop.classList.remove("is-dragging");
    const file = event.dataTransfer.files?.[0];
    if (file) handleInventoryFile(file);
  });

  els.invThreshold.value = state.threshold;
  els.invThreshold.addEventListener("change", () => {
    state.threshold = Math.max(0, Number(els.invThreshold.value) || 0);
    localStorage.setItem(THRESHOLD_KEY, String(state.threshold));
    render();
  });

  els.invForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const data = new FormData(els.invForm);
    const item = {
      brand: cleanText(data.get("brand")),
      product: cleanText(data.get("product")),
      flavor: cleanText(data.get("flavor")),
      stock: Math.max(0, Number(data.get("stock")) || 0)
    };
    try {
      await applyIncomingInventory([item], "merge");
      els.invForm.reset();
      setStatus("Inventory item added");
    } catch (error) {
      setStatus(error.message || "Inventory save failed");
    }
  });

  els.invRows.addEventListener("change", async (event) => {
    const input = event.target.closest("input[data-stock-id]");
    if (!input) return;
    const item = state.inventory.find((entry) => entry.id === input.dataset.stockId);
    if (!item) return;
    item.stock = Math.max(0, Number(input.value) || 0);
    try {
      await applyIncomingInventory(state.inventory, "replace");
      setStatus("Stock updated");
    } catch (error) {
      setStatus(error.message || "Stock update failed");
    }
  });

  els.invRows.addEventListener("click", async (event) => {
    const button = event.target.closest("button[data-del-id]");
    if (!button) return;
    const item = state.inventory.find((entry) => entry.id === button.dataset.delId);
    if (!item || !window.confirm(`Delete ${item.brand} ${item.product} | ${item.flavor}?`)) return;
    try {
      await applyIncomingInventory(state.inventory.filter((entry) => entry.id !== item.id), "replace");
      setStatus("Inventory item removed");
    } catch (error) {
      setStatus(error.message || "Delete failed");
    }
  });
  els.clearData.addEventListener("click", clearData);
  els.manualForm.addEventListener("submit", handleManualEntry);

  els.searchInput.addEventListener("input", () => {
    state.filters.query = els.searchInput.value.trim().toLowerCase();
    render();
  });

  els.phaseFilter.addEventListener("change", () => {
    state.filters.phase = els.phaseFilter.value;
    render();
  });

  els.modeFilter.addEventListener("change", () => {
    state.filters.mode = els.modeFilter.value;
    render();
  });
}

async function loadInitialData() {
  setStatus("Loading");

  try {
    const response = await fetch("/api/shipments", { cache: "no-store" });
    if (!response.ok) throw new Error("Shared data unavailable.");
    const data = await response.json();
    state.serverOnline = true;
    state.rows = normalizeRows(data.rows?.length ? data.rows : SAMPLE_ROWS);
    try {
      const invResponse = await fetch("/api/inventory", { cache: "no-store" });
      const invData = invResponse.ok ? await invResponse.json() : { items: [] };
      state.inventory = Array.isArray(invData.items) ? invData.items : [];
    } catch {
      state.inventory = [];
    }
    try {
      const snapResponse = await fetch("/api/snapshots", { cache: "no-store" });
      const snapData = snapResponse.ok ? await snapResponse.json() : { snapshots: [] };
      state.snapshots = Array.isArray(snapData.snapshots) ? snapData.snapshots : [];
    } catch {
      state.snapshots = [];
    }
    setStatus("Shared data");
  } catch {
    state.serverOnline = false;
    const saved = readLocalRows();
    state.rows = normalizeRows(saved.length ? saved : SAMPLE_ROWS);
    try {
      state.inventory = JSON.parse(localStorage.getItem(INV_STORE_KEY) || "[]");
    } catch {
      state.inventory = [];
    }
    setStatus("Local data");
  }
}

function readLocalRows() {
  try {
    const rows = JSON.parse(localStorage.getItem(STORE_KEY) || "[]");
    return Array.isArray(rows) ? rows : [];
  } catch {
    return [];
  }
}

function persistLocal() {
  localStorage.setItem(STORE_KEY, JSON.stringify(state.rows));
}


function askImportMode(count, label, preferred = "replace") {
  const dialog = els.importDialog;
  if (!dialog || typeof dialog.showModal !== "function") {
    const replace = window.confirm(`${count} ${label} found.\n\nOK = REPLACE all existing data\nCancel = ADD to existing data`);
    return Promise.resolve(replace ? "replace" : "merge");
  }
  els.importSummary.textContent = `${formatNumber(count)} ${label} found in the file, and there is already data loaded. What should happen to the existing data?`;
  els.importMerge.classList.toggle("suggested", preferred === "merge");
  els.importReplace.classList.toggle("suggested", preferred === "replace");
  return new Promise((resolve) => {
    const finish = (value) => {
      cleanup();
      if (dialog.open) dialog.close();
      resolve(value);
    };
    const onMerge = () => finish("merge");
    const onReplace = () => finish("replace");
    const onCancel = () => finish(null);
    function cleanup() {
      els.importMerge.removeEventListener("click", onMerge);
      els.importReplace.removeEventListener("click", onReplace);
      els.importCancel.removeEventListener("click", onCancel);
      dialog.removeEventListener("cancel", onCancel);
    }
    els.importMerge.addEventListener("click", onMerge);
    els.importReplace.addEventListener("click", onReplace);
    els.importCancel.addEventListener("click", onCancel);
    dialog.addEventListener("cancel", onCancel);
    dialog.showModal();
  });
}

async function handleFile(file) {
  try {
    setStatus("Reading file");
    const rows = await parseFile(file);
    if (!rows.length) throw new Error("No usable shipment rows were found.");

    let mode = "replace";
    if (state.rows.length) {
      mode = await askImportMode(rows.length, "shipment rows", state.importMode);
      if (!mode) {
        setStatus("Import cancelled — nothing changed");
        return;
      }
    }
    await applyIncomingRows(rows, mode);
    setStatus(
      mode === "merge"
        ? `${rows.length} row${rows.length === 1 ? "" : "s"} added to existing data`
        : `${rows.length} row${rows.length === 1 ? "" : "s"} imported`
    );
  } catch (error) {
    setStatus(error.message || "Import failed");
  }
}

async function applyIncomingRows(rows, mode) {
  const incomingRows = normalizeRows(rows);

  if (state.serverOnline) {
    const response = await fetch("/api/shipments", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-admin-pin": els.adminPin.value
      },
      body: JSON.stringify({ mode, rows: incomingRows })
    });

    const data = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(data.error || "Could not save shared data.");

    state.rows = normalizeRows(data.rows || []);
  } else {
    state.rows = mode === "merge" ? mergeRows(state.rows, incomingRows) : incomingRows;
    persistLocal();
  }

  render();
}

async function handleManualEntry(event) {
  event.preventDefault();

  const data = new FormData(els.manualForm);
  const row = normalizeRow({
    shipmentType: data.get("shipmentType"),
    pickupDate: data.get("pickupDate"),
    brand: data.get("brand"),
    productType: data.get("productType"),
    flavors: data.get("flavors"),
    boxCount: data.get("boxCount"),
    cargoStatus: data.get("cargoStatus"),
    source: "Manual"
  });

  try {
    await applyIncomingRows([row], "merge");
    els.manualForm.reset();
    setStatus("Manual entry saved");
  } catch (error) {
    setStatus(error.message || "Manual entry failed");
  }
}

async function clearData() {
  if (!window.confirm("Clear all dashboard data?")) return;

  try {
    if (state.serverOnline) {
      const response = await fetch("/api/shipments", {
        method: "DELETE",
        headers: { "x-admin-pin": els.adminPin.value }
      });
      const data = await response.json().catch(() => ({}));
      if (!response.ok) throw new Error(data.error || "Could not clear shared data.");
    } else {
      localStorage.removeItem(STORE_KEY);
    }

    state.rows = [];
    render();
    setStatus("Cleared");
  } catch (error) {
    setStatus(error.message || "Clear failed");
  }
}



/* ================= inventory ================= */
async function handleInventoryFile(file) {
  try {
    setStatus("Reading stock file");
    const items = await parseInventoryFile(file);
    if (!items.length) throw new Error("No usable stock rows were found.");

    let mode = "replace";
    if (state.inventory.length) {
      mode = await askImportMode(items.length, "stock SKUs", "replace");
      if (!mode) {
        setStatus("Import cancelled — nothing changed");
        return;
      }
    }
    await applyIncomingInventory(items, mode);
    setStatus(
      mode === "merge"
        ? `${items.length} SKU${items.length === 1 ? "" : "s"} added/updated in stock`
        : `${items.length} SKU${items.length === 1 ? "" : "s"} imported`
    );
  } catch (error) {
    setStatus(error.message || "Stock import failed");
  }
}

async function parseInventoryFile(file) {
  const name = (file.name || "").toLowerCase();
  let table;
  if (name.endsWith(".csv") || file.type === "text/csv") {
    const text = await file.text();
    table = text.split(/\r?\n/).map((line) => splitCsvLine(line));
  } else {
    if (!window.readXlsxFile) throw new Error("Excel reader is not available.");
    table = await window.readXlsxFile(file);
  }
  return inventoryTableToItems(table);
}

function splitCsvLine(line) {
  const out = [];
  let cur = "";
  let quoted = false;
  for (let i = 0; i < line.length; i += 1) {
    const ch = line[i];
    if (ch === '"') {
      if (quoted && line[i + 1] === '"') {
        cur += '"';
        i += 1;
      } else {
        quoted = !quoted;
      }
    } else if (ch === "," && !quoted) {
      out.push(cur);
      cur = "";
    } else {
      cur += ch;
    }
  }
  out.push(cur);
  return out;
}

function inventoryTableToItems(table) {
  if (!Array.isArray(table) || !table.length) return [];
  let headerIndex = -1;
  let nameCol = -1;
  let stockCol = -1;
  for (let i = 0; i < Math.min(table.length, 10); i += 1) {
    const cells = (table[i] || []).map((cell) => cleanText(cell).toLowerCase().replace(/[^a-z]/g, ""));
    const nIdx = cells.findIndex((cell) => /^(productname|product|itemname|item|sku|description|name)/.test(cell) && cell);
    const sIdx = cells.findIndex((cell) => /^(currentstock|stock|onhand|qtyonhand|inventory|available|balance|quantity|qty|units)/.test(cell) && cell);
    if (nIdx >= 0 && sIdx >= 0 && nIdx !== sIdx) {
      headerIndex = i;
      nameCol = nIdx;
      stockCol = sIdx;
      break;
    }
  }
  if (headerIndex < 0) return [];

  const items = [];
  for (let i = headerIndex + 1; i < table.length; i += 1) {
    const row = table[i] || [];
    const rawName = cleanText(row[nameCol]);
    if (!rawName || /^(grand\s*)?total/i.test(rawName)) continue;
    const stock = Math.max(0, parseNumber(row[stockCol]));
    items.push({ ...parseStockProductName(rawName), stock });
  }
  return items;
}

function parseStockProductName(text) {
  let brand = "—";
  let product = "—";
  let flavor = cleanText(text);
  const pipeIndex = flavor.indexOf("|");
  if (pipeIndex >= 0) {
    const left = flavor.slice(0, pipeIndex).trim();
    flavor = flavor.slice(pipeIndex + 1).trim() || "—";
    const dashIndex = left.indexOf(" - ");
    if (dashIndex >= 0) {
      brand = left.slice(0, dashIndex).trim() || "—";
      product = left.slice(dashIndex + 3).trim() || "—";
    } else {
      product = left || "—";
    }
  } else {
    const dashIndex = flavor.indexOf(" - ");
    if (dashIndex >= 0) {
      brand = flavor.slice(0, dashIndex).trim() || "—";
      flavor = flavor.slice(dashIndex + 3).trim() || "—";
    }
  }
  return { brand, product, flavor };
}

async function applyIncomingInventory(items, mode) {
  const payloadItems = items.map((item) => ({
    id: item.id,
    brand: cleanText(item.brand) || "—",
    product: cleanText(item.product) || "—",
    flavor: cleanText(item.flavor) || "—",
    stock: Math.max(0, Number(item.stock) || 0)
  }));

  if (state.serverOnline) {
    const response = await fetch("/api/inventory", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-admin-pin": els.adminPin.value
      },
      body: JSON.stringify({ mode, items: payloadItems })
    });
    const data = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(data.error || "Could not save inventory.");
    state.inventory = Array.isArray(data.items) ? data.items : [];
  } else {
    const withIds = payloadItems.map((item) => ({ ...item, id: item.id || crypto.randomUUID() }));
    if (mode === "merge") {
      const map = new Map(state.inventory.map((item) => [invKey(item), item]));
      withIds.forEach((item) => {
        const existing = map.get(invKey(item));
        map.set(invKey(item), { ...item, id: existing?.id || item.id });
      });
      state.inventory = Array.from(map.values());
    } else {
      state.inventory = withIds;
    }
    localStorage.setItem(INV_STORE_KEY, JSON.stringify(state.inventory));
  }
  render();
}

function invKey(item) {
  return [item.brand, item.product, item.flavor].map((part) => cleanText(part).toLowerCase()).join("|");
}

function normKey(value) {
  return cleanText(value).toUpperCase().replace(/[^A-Z0-9]/g, "");
}

function inboundByBrandFlavor(rows) {
  const map = new Map();
  rows
    .filter((row) => row.phase !== "Received")
    .forEach((row) => {
      row.lines.forEach((line) => {
        const key = `${normKey(row.brand)}|${normKey(line.name)}`;
        map.set(key, (map.get(key) || 0) + (line.cases || 0));
      });
    });
  return map;
}

function stockFlag(item) {
  if (!item.stock) return { label: "OUT", cls: "flag-out" };
  if (item.stock <= state.threshold) return { label: "LOW", cls: "flag-low" };
  return { label: "OK", cls: "flag-ok" };
}

function renderInventory(enrichedRows) {
  if (!els.invRows) return;
  const items = state.inventory.slice().sort((a, b) => (a.stock || 0) - (b.stock || 0));
  const inbound = inboundByBrandFlavor(enrichedRows);
  const totalUnits = items.reduce((sum, item) => sum + (item.stock || 0), 0);
  const low = items.filter((item) => item.stock > 0 && item.stock <= state.threshold).length;
  const out = items.filter((item) => !item.stock).length;

  els.invCount.textContent = items.length
    ? `${formatNumber(totalUnits)} units · ${items.length} SKUs · ${low} low · ${out} out`
    : "No stock loaded";

  if (!items.length) {
    els.invRows.innerHTML = `<tr><td colspan="7" class="empty-cell">Upload the stock file (Product Name / Current Stock) or add items manually.</td></tr>`;
    return;
  }

  els.invRows.innerHTML = items
    .map((item) => {
      const flag = stockFlag(item);
      const inboundCases = inbound.get(`${normKey(item.brand)}|${normKey(item.flavor)}`) || 0;
      return `<tr>
        <td><strong>${escapeHtml(item.brand)}</strong></td>
        <td>${escapeHtml(item.product)}</td>
        <td>${escapeHtml(item.flavor)}</td>
        <td class="num"><input class="stock-input" type="number" min="0" value="${item.stock || 0}" data-stock-id="${escapeHtml(item.id)}"></td>
        <td class="num muted-cell" title="Inbound cartons matched by brand + flavor">${inboundCases ? `+${formatNumber(inboundCases)}` : "—"}</td>
        <td><span class="flag ${flag.cls}">${flag.label}</span></td>
        <td class="num"><button class="row-del" type="button" data-del-id="${escapeHtml(item.id)}" title="Delete">✕</button></td>
      </tr>`;
    })
    .join("");
}

/* ================= calendar ================= */
function renderCalendar(rows) {
  if (!els.calendarGrid) return;
  const start = startOfDay(new Date());
  start.setDate(start.getDate() - ((start.getDay() + 6) % 7)); // back to Monday
  const days = 28;
  const events = new Map();
  rows
    .filter((row) => row.relevantDate && row.phase !== "Received")
    .forEach((row) => {
      if (!events.has(row.relevantDate)) events.set(row.relevantDate, []);
      events.get(row.relevantDate).push(row);
    });

  const end = new Date(start);
  end.setDate(end.getDate() + days - 1);
  els.calRange.textContent = `${shortDate(toIsoDate(start))} – ${shortDate(toIsoDate(end))}`;

  const todayIso = toIsoDate(new Date());
  const weekdays = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
    .map((day) => `<div class="cal-head">${day}</div>`)
    .join("");

  const cells = [];
  for (let i = 0; i < days; i += 1) {
    const date = new Date(start);
    date.setDate(date.getDate() + i);
    const iso = toIsoDate(date);
    const dayEvents = events.get(iso) || [];
    const cases = sumCases(dayEvents);
    const isToday = iso === todayIso;
    const isPast = iso < todayIso;
    const monthLabel = date.getDate() === 1 || i === 0 ? `<span class="cal-month">${date.toLocaleDateString(undefined, { month: "short" })}</span>` : "";
    const pop = dayEvents.length
      ? `<div class="cal-pop"><div class="cal-pop-date">${date.toLocaleDateString(undefined, { weekday: "long", month: "long", day: "numeric" })} · ${formatNumber(cases)} cases</div>${dayEvents
          .map(
            (row) => `<div class="cal-pop-row"><strong>${escapeHtml(row.brand)} ${escapeHtml(row.productType)}</strong><span>${formatNumber(row.caseCount)} cases · ${escapeHtml(row.phase)} · ${escapeHtml(row.mode)}</span>${row.cargoStatus ? `<em>${escapeHtml(row.cargoStatus)}</em>` : ""}</div>`
          )
          .join("")}</div>`
      : "";
    cells.push(
      `<div class="cal-day${isToday ? " today" : ""}${isPast ? " past" : ""}${dayEvents.length ? " has-events" : ""}">
        <span class="cal-num">${monthLabel}${date.getDate()}</span>
        ${dayEvents.length ? `<span class="cal-badge">${formatNumber(cases)}</span><span class="cal-ships">${dayEvents.length} shipment${dayEvents.length === 1 ? "" : "s"}</span>` : ""}
        ${pop}
      </div>`
    );
  }

  els.calendarGrid.innerHTML = weekdays + cells.join("");
}

/* ================= archive ================= */
function isArchived(row) {
  if (row.phase !== "Received") return false;
  const reference = row.relevantDate || (row.updatedAt ? row.updatedAt.slice(0, 10) : null);
  if (!reference) return false;
  const age = (startOfDay(new Date()) - new Date(`${reference}T00:00:00`)) / 86400000;
  return age > ARCHIVE_AFTER_DAYS;
}

/* ================= csv export ================= */
function csvEscape(value) {
  const text = String(value ?? "");
  return /[",\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

function exportCsv() {
  const rows = state.rows.map(enrichRow);
  const header = ["Shipment Type", "Pick Up Date", "Brand", "Product Type", "Flavors", "Box Count", "Cargo Status", "Phase", "Freight", "Source"];
  const lines = [header.join(",")];
  rows.forEach((row) => {
    lines.push(
      [row.shipmentType, row.pickupDate, row.brand, row.productType, row.flavors, row.caseCount, row.cargoStatus, row.phase, row.mode, row.source]
        .map(csvEscape)
        .join(",")
    );
  });
  if (state.inventory.length) {
    lines.push("");
    lines.push(["Inventory — Brand", "Product (SPU)", "Flavor (SKU)", "Current Stock", "Alert"].join(","));
    state.inventory
      .slice()
      .sort((a, b) => (a.stock || 0) - (b.stock || 0))
      .forEach((item) => {
        lines.push([item.brand, item.product, item.flavor, item.stock || 0, stockFlag(item).label].map(csvEscape).join(","));
      });
  }
  const blob = new Blob([`\ufeff${lines.join("\n")}`], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `shipment-inventory-${toIsoDate(new Date())}.csv`;
  link.click();
  URL.revokeObjectURL(url);
}

function openBoardReport() {
  const rows = state.rows.map(enrichRow);
  if (!rows.length) {
    window.alert("No shipment data to report yet.");
    return;
  }

  const total = sumCases(rows);
  const phases = new Map();
  const modes = new Map();
  const brands = new Map();
  rows.forEach((row) => {
    phases.set(row.phase, (phases.get(row.phase) || 0) + row.caseCount);
    modes.set(row.mode, (modes.get(row.mode) || 0) + row.caseCount);
    brands.set(row.brand || "—", (brands.get(row.brand || "—") || 0) + row.caseCount);
  });
  const groups = groupProducts(rows);
  const attention = rows.filter((row) => row.phase === "Attention");
  const inTransit = phases.get("In Transit") || 0;
  const received = phases.get("Received") || 0;

  const barRows = (map) => {
    const entries = Array.from(map.entries()).sort((a, b) => b[1] - a[1]);
    const max = Math.max(...entries.map((e) => e[1]), 1);
    const sum = entries.reduce((acc, e) => acc + e[1], 0) || 1;
    return entries
      .map(
        ([name, value]) => `<div class="brow"><div class="btop"><span>${escapeHtml(name)}</span><b>${formatNumber(value)} <i>${Math.round((value / sum) * 100)}%</i></b></div><div class="bar"><span style="width:${Math.max((value / max) * 100, 2)}%"></span></div></div>`
      )
      .join("");
  };

  const productCards = groups
    .map((group) => {
      const top = group.lines.slice(0, 6);
      const max = Math.max(...top.map((line) => line.cases), 1);
      const flavorRows = top
        .map(
          (line) => `<div class="frow"><div class="btop"><span>${escapeHtml(line.name)}</span><b>${formatNumber(line.cases)}</b></div><div class="bar slim"><span style="width:${Math.max((line.cases / max) * 100, 2)}%"></span></div></div>`
        )
        .join("");
      const more = group.lines.length > 6 ? `<div class="more">+ ${group.lines.length - 6} more flavors</div>` : "";
      return `<div class="pcard"><div class="pbrand">${escapeHtml(group.brand)} · ${escapeHtml(group.phase)}</div><div class="pname">${escapeHtml(group.product)}</div>${flavorRows}${more}<div class="ptotal"><b>${formatNumber(group.total)}</b> cases · ${group.shipments} shipment${group.shipments === 1 ? "" : "s"}</div></div>`;
    })
    .join("");

  const tableRows = rows
    .slice()
    .sort((a, b) => (b.relevantDate || "").localeCompare(a.relevantDate || ""))
    .map(
      (row) => `<tr><td>${escapeHtml(row.mode)}</td><td>${escapeHtml(shortDate(row.pickupDate) || "—")}</td><td><b>${escapeHtml(row.brand)}</b></td><td>${escapeHtml(row.productType)}</td><td class="num">${formatNumber(row.caseCount)}</td><td>${escapeHtml(row.phase)}</td><td class="muted">${escapeHtml(row.cargoStatus || "—")}</td></tr>`
    )
    .join("");

  const attentionBlock = attention.length
    ? `<div class="section attention"><h3>Needs attention (${attention.length})</h3>${attention
        .map(
          (row) => `<div class="att-row"><b>${escapeHtml(row.brand)} ${escapeHtml(row.productType)}</b> — ${formatNumber(row.caseCount)} cases · ${escapeHtml(row.cargoStatus || "no status")}</div>`
        )
        .join("")}</div>`
    : "";

  const weekAgo = (() => {
    const target = new Date();
    target.setDate(target.getDate() - 6);
    const targetIso = toIsoDate(target);
    const candidates = (state.snapshots || []).filter((snap) => snap.date <= targetIso);
    return candidates.length ? candidates[candidates.length - 1] : null;
  })();
  const fmtDelta = (now, then) => {
    const diff = now - then;
    if (!diff) return "unchanged";
    return `${diff > 0 ? "+" : "−"}${formatNumber(Math.abs(diff))}`;
  };
  const stockTotal = state.inventory.reduce((sum, item) => sum + (item.stock || 0), 0);
  const trendBlock = weekAgo
    ? `<div class="section"><h3>Since ${escapeHtml(shortDate(weekAgo.date))}</h3><div class="trend"><span><b>${fmtDelta(total, weekAgo.totalCases)}</b> cases inbound</span><span><b>${fmtDelta(rows.length, weekAgo.shipmentCount)}</b> shipments</span>${state.inventory.length ? `<span><b>${fmtDelta(stockTotal, weekAgo.totalStock || 0)}</b> units on hand</span>` : ""}</div></div>`
    : "";
  const lowItems = state.inventory.filter((item) => item.stock > 0 && item.stock <= state.threshold).sort((a, b) => a.stock - b.stock);
  const outItems = state.inventory.filter((item) => !item.stock);
  const inboundMap = inboundByBrandFlavor(rows);
  const invList = (items) => items
    .slice(0, 12)
    .map((item) => {
      const inboundCases = inboundMap.get(`${normKey(item.brand)}|${normKey(item.flavor)}`) || 0;
      return `<div class="att-row"><b>${escapeHtml(item.brand)} ${escapeHtml(item.flavor)}</b> — ${formatNumber(item.stock || 0)} in stock${inboundCases ? ` · ${formatNumber(inboundCases)} inbound` : " · nothing inbound"}</div>`;
    })
    .join("") || '<div class="att-row">None</div>';
  const inventoryBlock = state.inventory.length
    ? `<div class="cols"><div class="section"><h3>Warehouse — ${formatNumber(stockTotal)} units · ${state.inventory.length} SKUs</h3><div class="trend"><span><b>${lowItems.length}</b> low stock</span><span><b>${outItems.length}</b> out of stock</span></div></div><div class="section"><h3>Low stock (top)</h3>${invList(lowItems)}</div><div class="section attention"><h3>Out of stock</h3>${invList(outItems)}</div></div>`
    : "";
  const generated = new Date();
  const dateLong = generated.toLocaleDateString(undefined, { weekday: "long", year: "numeric", month: "long", day: "numeric" });

  const html = `<!DOCTYPE html>
<html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Board Overview — ${dateLong}</title><style>
*{margin:0;padding:0;box-sizing:border-box}
@page{margin:13mm}
body{font-family:'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;background:#fff;color:#0e0e10;-webkit-font-smoothing:antialiased;padding:clamp(20px,4vw,56px)}
.sheet{max-width:1000px;margin:0 auto}
.eyebrow{font-size:11px;font-weight:800;letter-spacing:.12em;text-transform:uppercase;color:#8a8a93}
h1{font-size:clamp(26px,4vw,40px);font-weight:850;letter-spacing:-.02em;margin:4px 0 2px}
.subtitle{color:#8a8a93;font-size:14px;margin-bottom:28px}
.printbtn{position:fixed;top:18px;right:18px;background:#0e0e10;color:#fff;border:none;border-radius:10px;padding:11px 18px;font-size:13px;font-weight:700;cursor:pointer;font-family:inherit}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin-bottom:28px}
.kpi{border:1px solid #e6e6ea;border-radius:14px;padding:16px 18px}
.kpi.dark{background:#0e0e10;border-color:#0e0e10;color:#fff}
.kpi .lbl{font-size:10px;font-weight:800;letter-spacing:.09em;text-transform:uppercase;color:#8a8a93}
.kpi.dark .lbl{color:#9c9ca6}
.kpi .val{font-size:26px;font-weight:850;letter-spacing:-.02em;margin-top:6px;font-variant-numeric:tabular-nums}
.cols{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:14px;margin-bottom:28px}
.section{border:1px solid #e6e6ea;border-radius:14px;padding:20px;break-inside:avoid;margin-bottom:14px}
.section h3{font-size:11px;font-weight:800;letter-spacing:.09em;text-transform:uppercase;color:#8a8a93;margin-bottom:14px}
.brow,.frow{margin-bottom:11px}
.btop{display:flex;justify-content:space-between;gap:10px;font-size:13px;margin-bottom:5px}
.btop b{font-variant-numeric:tabular-nums;font-weight:700}
.btop i{font-style:normal;color:#8a8a93;font-size:11px;font-weight:600}
.bar{height:7px;background:#f0f0f3;border-radius:99px;overflow:hidden}
.bar.slim{height:5px}
.bar span{display:block;height:100%;background:#0e0e10;border-radius:99px}
.pgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:14px;margin-bottom:28px}
.pcard{border:1px solid #e6e6ea;border-radius:14px;padding:18px 20px;break-inside:avoid}
.pbrand{font-size:10px;font-weight:800;letter-spacing:.1em;text-transform:uppercase;color:#8a8a93}
.pname{font-size:17px;font-weight:820;margin:2px 0 14px}
.ptotal{border-top:1px solid #f0f0f3;margin-top:12px;padding-top:11px;font-size:13px;color:#8a8a93}
.ptotal b{color:#0e0e10;font-size:15px}
.more{font-size:11.5px;color:#8a8a93}
table{width:100%;border-collapse:collapse;font-size:12.5px}
th{font-size:10px;font-weight:800;letter-spacing:.08em;text-transform:uppercase;color:#8a8a93;text-align:left;padding:9px 10px;border-bottom:1px solid #e6e6ea}
td{padding:9px 10px;border-bottom:1px solid #f0f0f3;vertical-align:top}
td.num{text-align:right;font-weight:700;font-variant-numeric:tabular-nums}
td.muted{color:#8a8a93;max-width:280px}
.attention{border-color:#f4c7c3;background:#fffafa}
.att-row{font-size:13px;padding:7px 0;border-bottom:1px solid #f7e3e1}
.att-row:last-child{border-bottom:none}
.trend{display:flex;gap:22px;flex-wrap:wrap;font-size:13px;color:#8a8a93}
.trend b{color:#0e0e10;font-size:16px;font-variant-numeric:tabular-nums}
.foot{margin-top:26px;font-size:11.5px;color:#8a8a93;display:flex;justify-content:space-between;gap:12px;flex-wrap:wrap}
@media print{.printbtn{display:none}body{padding:0}}
</style></head><body>
<button class="printbtn" onclick="window.print()">Print / Save as PDF</button>
<div class="sheet">
  <div class="eyebrow">Shipment / Inventory</div>
  <h1>Board Overview</h1>
  <div class="subtitle">${dateLong} · ${rows.length} active shipment${rows.length === 1 ? "" : "s"} · live data snapshot</div>
  <div class="kpis">
    <div class="kpi dark"><div class="lbl">Total cases inbound</div><div class="val">${formatNumber(total)}</div></div>
    <div class="kpi"><div class="lbl">In transit</div><div class="val">${formatNumber(inTransit)}</div></div>
    <div class="kpi"><div class="lbl">Received</div><div class="val">${formatNumber(received)}</div></div>
    <div class="kpi"><div class="lbl">Brands</div><div class="val">${brands.size}</div></div>
    <div class="kpi"><div class="lbl">Needs attention</div><div class="val">${attention.length}</div></div>
  </div>
  ${trendBlock}
  ${attentionBlock}
  <div class="cols">
    <div class="section"><h3>Cases by status</h3>${barRows(phases)}</div>
    <div class="section"><h3>Cases by brand</h3>${barRows(brands)}</div>
    <div class="section"><h3>Cases by freight</h3>${barRows(modes)}</div>
  </div>
  <div class="eyebrow" style="margin-bottom:12px">Product breakdown — brand / SPU / flavor</div>
  <div class="pgrid">${productCards}</div>
  ${inventoryBlock}
  <div class="section"><h3>All shipments</h3><table><thead><tr><th>Mode</th><th>Pick up</th><th>Brand</th><th>Product</th><th style="text-align:right">Cases</th><th>Status</th><th>Cargo status</th></tr></thead><tbody>${tableRows}</tbody></table></div>
  <div class="foot"><span>Generated ${generated.toLocaleString()}</span><span>Internal — prepared for board review</span></div>
</div>
</body></html>`;

  const win = window.open("", "_blank");
  if (!win) {
    window.alert("Please allow pop-ups for this site to open the report.");
    return;
  }
  win.document.write(html);
  win.document.close();
}

function exportJson() {
  const blob = new Blob(
    [
      JSON.stringify(
        {
          exportedAt: new Date().toISOString(),
          rows: state.rows
        },
        null,
        2
      )
    ],
    { type: "application/json" }
  );
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `safa-shipments-${toIsoDate(new Date())}.json`;
  link.click();
  URL.revokeObjectURL(url);
}

async function parseFile(file) {
  const name = file.name.toLowerCase();
  if (name.endsWith(".csv")) {
    const text = await file.text();
    return csvToObjects(text).map((row, index) =>
      normalizeRow(row, { source: file.name, rowNumber: index + 2 })
    );
  }

  if (!window.readXlsxFile) {
    throw new Error("Excel reader is still loading. Try again in a moment.");
  }

  const tableRows = await window.readXlsxFile(file);
  const rawRows = tableRowsToObjects(tableRows);

  return rawRows.map((row, index) =>
    normalizeRow(row, { source: file.name, rowNumber: index + 2 })
  );
}

function tableRowsToObjects(tableRows) {
  const rows = (tableRows || []).filter((row) =>
    row.some((cell) => cleanText(cell))
  );
  const headers = rows.shift() || [];

  return rows.map((row) => {
    const object = {};
    headers.forEach((header, index) => {
      object[cleanText(header) || `Column ${index + 1}`] = row[index] ?? "";
    });
    return object;
  });
}

function csvToObjects(text) {
  const rows = [];
  let current = [];
  let field = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    const next = text[i + 1];

    if (char === '"' && inQuotes && next === '"') {
      field += '"';
      i += 1;
      continue;
    }

    if (char === '"') {
      inQuotes = !inQuotes;
      continue;
    }

    if (char === "," && !inQuotes) {
      current.push(field);
      field = "";
      continue;
    }

    if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") i += 1;
      current.push(field);
      rows.push(current);
      current = [];
      field = "";
      continue;
    }

    field += char;
  }

  current.push(field);
  rows.push(current);

  const cleanRows = rows.filter((row) => row.some((cell) => String(cell).trim()));
  const headers = cleanRows.shift() || [];
  return cleanRows.map((row) => {
    const object = {};
    headers.forEach((header, index) => {
      object[header] = row[index] ?? "";
    });
    return object;
  });
}

function normalizeRows(rows) {
  return rows.map((row) => normalizeRow(row)).filter(Boolean);
}

function normalizeRow(input, meta = {}) {
  const mapped = {};

  Object.entries(input || {}).forEach(([key, value]) => {
    const field = mapHeader(key);
    if (field) mapped[field] = value;
  });

  const flavors = cleanText(mapped.flavors ?? input.flavors);
  const flavorTotal = parseFlavorLines(flavors).reduce((sum, line) => sum + line.cases, 0);
  const boxCount = parseNumber(mapped.boxCount ?? input.boxCount) || flavorTotal;
  const cargoStatus = cleanText(mapped.cargoStatus ?? input.cargoStatus);
  const shipmentType = cleanText(mapped.shipmentType ?? input.shipmentType);

  const row = {
    id: cleanText(input.id) || buildId(),
    shipmentType,
    pickupDate: normalizeDate(mapped.pickupDate ?? input.pickupDate),
    brand: cleanText(mapped.brand ?? input.brand).toUpperCase(),
    productType: cleanText(mapped.productType ?? input.productType).toUpperCase(),
    flavors,
    boxCount,
    cargoStatus,
    source: cleanText(meta.source ?? input.source) || "Upload",
    updatedAt: cleanText(input.updatedAt) || new Date().toISOString()
  };

  if (!row.brand && !row.productType && !row.cargoStatus && !row.boxCount) {
    return null;
  }

  return row;
}

function mapHeader(header) {
  const key = cleanText(header).toLowerCase().replace(/[^a-z0-9]/g, "");
  const map = {
    shipmenttype: "shipmentType",
    shiptype: "shipmentType",
    freighttype: "shipmentType",
    type: "shipmentType",
    pickupdate: "pickupDate",
    pickup: "pickupDate",
    date: "pickupDate",
    etd: "pickupDate",
    eta: "pickupDate",
    brand: "brand",
    producttype: "productType",
    product: "productType",
    item: "productType",
    sku: "productType",
    flavors: "flavors",
    flavor: "flavors",
    flavours: "flavors",
    flavour: "flavors",
    details: "flavors",
    boxcount: "boxCount",
    cases: "boxCount",
    casecount: "boxCount",
    cartons: "boxCount",
    ctns: "boxCount",
    cargostatus: "cargoStatus",
    status: "cargoStatus",
    notes: "cargoStatus",
    remarks: "cargoStatus"
  };

  return map[key] || null;
}

function parseFlavorLines(text) {
  const lines = cleanText(text)
    .split(/\n|;/)
    .map((line) => line.trim())
    .filter(Boolean);

  return lines.map((line) => {
    const match = line
      .replace(/\s+/g, " ")
      .match(/^(.*?)(?:\s+|-)?(\d+(?:,\d{3})*(?:\.\d+)?)\s*(ctns?|cartons?|cases?|boxes?)?\s*$/i);
    const cases = match ? parseNumber(match[2]) : 0;
    const name = match ? cleanText(match[1]) : line;
    return { name: name || line, cases };
  });
}

function enrichRow(row) {
  const lines = parseFlavorLines(row.flavors);
  const lineTotal = lines.reduce((sum, line) => sum + line.cases, 0);
  const caseCount = Number(row.boxCount) || lineTotal || 0;
  const mode = inferMode(row);
  const phase = inferPhase(row);
  const relevantDate = getRelevantDate(row);

  return {
    ...row,
    caseCount,
    lines: lines.length ? lines : [{ name: row.cargoStatus || row.shipmentType || "Shipment", cases: caseCount }],
    mode,
    phase,
    relevantDate
  };
}

function inferMode(row) {
  const shipmentType = cleanText(row.shipmentType).toLowerCase();
  const status = cleanText(row.cargoStatus).toLowerCase();
  const text = `${shipmentType} ${status}`;

  if (/sea\s*or\s*air|air\s*or\s*sea/.test(shipmentType)) {
    if (/vessel|sea|ocean|port|sail/.test(status)) return "Sea";
    if (/air|flight|flying/.test(status)) return "Air";
    return "Mixed";
  }

  if (/air|flight|flying/.test(text)) return "Air";
  if (/truck|dray|clip/.test(text)) return "Truck";
  if (/sea|vessel|ocean|port|sail/.test(text)) return "Sea";
  return "Mixed";
}

function inferPhase(row) {
  const text = `${row.shipmentType} ${row.cargoStatus}`.toLowerCase();

  if (/hold|delay|customs|problem|issue/.test(text)) return "Attention";
  if (/received|delivered|warehouse|arrived warehouse|stocked/.test(text)) return "Received";
  if (/truck|in truck|dray/.test(text)) return "Trucking";
  if (/port arrival|arriving|arrived at port|at port/.test(text)) return "Port Arrival";
  if (/departed|in transit|on water|sailing|sailed|flying|flight/.test(text)) return "In Transit";
  if (/will depart|scheduled|booking|booked|pending/.test(text)) return "Scheduled";
  if (/vessel|depart/.test(text)) return "In Transit";

  return "Scheduled";
}

function getRelevantDate(row) {
  const yearHint = row.pickupDate ? Number(row.pickupDate.slice(0, 4)) : new Date().getFullYear();
  const dates = [
    ...extractStatusDates(row.cargoStatus, yearHint),
    row.pickupDate
  ].filter(Boolean);

  if (!dates.length) return "";

  const today = startOfDay(new Date()).getTime();
  const sorted = dates
    .map((date) => new Date(`${date}T00:00:00`))
    .filter((date) => !Number.isNaN(date.getTime()))
    .sort((a, b) => a - b);
  const future = sorted.find((date) => date.getTime() >= today);
  return toIsoDate(future || sorted[sorted.length - 1]);
}

function extractStatusDates(text, fallbackYear) {
  const dates = [];
  const regex = /(\d{1,2})[\/.-](\d{1,2})(?:[\/.-](\d{2,4}))?/g;
  let match;

  while ((match = regex.exec(text || ""))) {
    const month = Number(match[1]);
    const day = Number(match[2]);
    const rawYear = match[3] ? Number(match[3]) : fallbackYear;
    const year = rawYear < 100 ? 2000 + rawYear : rawYear;
    const date = new Date(year, month - 1, day);
    if (!Number.isNaN(date.getTime())) dates.push(toIsoDate(date));
  }

  return dates;
}

function render() {
  const enrichedRows = state.rows.map(enrichRow);
  const archivedRows = enrichedRows.filter(isArchived);
  const activeRows = enrichedRows.filter((row) => !isArchived(row));
  const visibleRows = state.showArchive ? enrichedRows : activeRows;
  const filteredRows = filterRows(visibleRows);

  if (els.archiveToggle) {
    els.archiveToggle.textContent = state.showArchive
      ? `Hide archived (${archivedRows.length})`
      : `Show archived (${archivedRows.length})`;
    els.archiveToggle.classList.toggle("active", state.showArchive);
  }
  renderCalendar(activeRows);
  renderInventory(activeRows);

  els.asOfDate.textContent = formatDate(new Date());
  els.connectionBadge.textContent = state.serverOnline ? "Shared" : "Local";
  els.rowCount.textContent = `${filteredRows.length} shipment${filteredRows.length === 1 ? "" : "s"}`;

  renderMetrics(enrichedRows);
  renderProducts(filteredRows);
  renderTable(filteredRows);
}

function filterRows(rows) {
  return rows.filter((row) => {
    const searchText = [
      row.brand,
      row.productType,
      row.flavors,
      row.cargoStatus,
      row.shipmentType,
      row.source
    ]
      .join(" ")
      .toLowerCase();

    const queryMatch = !state.filters.query || searchText.includes(state.filters.query);
    const phaseMatch = state.filters.phase === "all" || row.phase === state.filters.phase;
    const modeMatch = state.filters.mode === "all" || row.mode === state.filters.mode;
    return queryMatch && phaseMatch && modeMatch;
  });
}

function renderMetrics(rows) {
  const activeRows = rows.filter((row) => row.phase !== "Received");
  const total = sumCases(activeRows);
  const air = sumCases(activeRows.filter((row) => row.mode === "Air"));
  const moving = sumCases(
    activeRows.filter((row) => ["In Transit", "Trucking", "Port Arrival"].includes(row.phase))
  );
  const nextGroup = getNextGroup(activeRows);

  const metrics = [
    {
      label: "Total Cases Inbound",
      value: formatNumber(total),
      sub: "Across active shipments",
      dark: true
    },
    {
      label: nextGroup ? `Next Move ${shortDate(nextGroup.date)}` : "Next Move",
      value: formatNumber(nextGroup?.cases || 0),
      sub: nextGroup ? nextGroup.label : "No dated shipment",
      dark: false
    },
    {
      label: "Air Freight",
      value: formatNumber(air),
      sub: `${rows.filter((row) => row.mode === "Air").length} air shipment${rows.filter((row) => row.mode === "Air").length === 1 ? "" : "s"}`,
      dark: false
    },
    {
      label: "In Transit Now",
      value: formatNumber(moving),
      sub: "Moving, at port, or on truck",
      dark: false
    }
  ];

  els.metrics.innerHTML = metrics
    .map(
      (metric) => `
        <article class="metric-card${metric.dark ? " dark" : ""}">
          <span class="label">${escapeHtml(metric.label)}</span>
          <span class="value">${escapeHtml(metric.value)}</span>
          <span class="sub">${escapeHtml(metric.sub)}</span>
        </article>
      `
    )
    .join("");
}

function renderProducts(rows) {
  const groups = groupProducts(rows);

  if (!groups.length) {
    els.productGrid.innerHTML = `<div class="empty">No shipments match the current view.</div>`;
    return;
  }

  els.productGrid.innerHTML = groups
    .map((group) => {
      const maxLine = Math.max(...group.lines.map((line) => line.cases), 1);
      const visibleLines = group.lines.slice(0, 5);
      const phaseClass = phaseClassName(group.phase);
      const rowsHtml = visibleLines
        .map((line) => {
          const width = Math.max(4, Math.round((line.cases / maxLine) * 100));
          return `
            <div class="breakdown-row">
              <div class="breakdown-label">
                <span>${escapeHtml(line.name)}</span>
                <span>${formatNumber(line.cases)}</span>
              </div>
              <div class="bar-track"><div class="bar-fill" style="width:${width}%"></div></div>
            </div>
          `;
        })
        .join("");

      return `
        <article class="product-card">
          <div class="product-brand">${escapeHtml(group.brand || "Unassigned")}</div>
          <div class="product-title">${escapeHtml(group.product || "Product")}</div>
          <div class="product-meta">
            <span>${group.shipments} shipment${group.shipments === 1 ? "" : "s"}</span>
            <span class="phase ${phaseClass}">${escapeHtml(group.phase)}</span>
          </div>
          <div class="breakdown-list">${rowsHtml}</div>
          <div class="product-total">${formatNumber(group.total)} cases <span>total ${escapeHtml(group.product || "product")}</span></div>
        </article>
      `;
    })
    .join("");
}

function renderTable(rows) {
  const sorted = [...rows].sort((a, b) => {
    const aTime = a.relevantDate ? new Date(a.relevantDate).getTime() : 0;
    const bTime = b.relevantDate ? new Date(b.relevantDate).getTime() : 0;
    return bTime - aTime || b.caseCount - a.caseCount;
  });

  if (!sorted.length) {
    els.shipmentRows.innerHTML = `
      <tr>
        <td colspan="6" class="muted">No shipment rows to show.</td>
      </tr>
    `;
    return;
  }

  els.shipmentRows.innerHTML = sorted
    .map(
      (row) => `
        <tr>
          <td>
            <strong>${escapeHtml(row.brand || "Unassigned")}</strong>
            <div class="muted">${escapeHtml(row.productType || "Product")}</div>
          </td>
          <td>${formatNumber(row.caseCount)}</td>
          <td>${escapeHtml(row.mode)}</td>
          <td>
            <span class="phase ${phaseClassName(row.phase)}">${escapeHtml(row.phase)}</span>
            <div class="muted">${escapeHtml(row.cargoStatus || row.shipmentType || "")}</div>
          </td>
          <td>${row.relevantDate ? escapeHtml(formatDate(new Date(`${row.relevantDate}T00:00:00`))) : "--"}</td>
          <td>${escapeHtml(row.source || "Dashboard")}</td>
        </tr>
      `
    )
    .join("");
}

function groupProducts(rows) {
  const groups = new Map();

  rows.forEach((row) => {
    const key = `${row.brand}|${row.productType}`;
    if (!groups.has(key)) {
      groups.set(key, {
        brand: row.brand,
        product: row.productType,
        total: 0,
        shipments: 0,
        phases: new Map(),
        lines: new Map()
      });
    }

    const group = groups.get(key);
    group.total += row.caseCount;
    group.shipments += 1;
    group.phases.set(row.phase, (group.phases.get(row.phase) || 0) + row.caseCount);

    const lines = row.lines.length ? row.lines : [{ name: row.cargoStatus || "Shipment", cases: row.caseCount }];
    lines.forEach((line) => {
      const name = line.name || row.cargoStatus || "Shipment";
      const cases = line.cases || (lines.length === 1 ? row.caseCount : 0);
      group.lines.set(name, (group.lines.get(name) || 0) + cases);
    });
  });

  return Array.from(groups.values())
    .map((group) => ({
      ...group,
      phase: topPhase(group.phases),
      lines: Array.from(group.lines, ([name, cases]) => ({ name, cases })).sort((a, b) => b.cases - a.cases)
    }))
    .sort((a, b) => b.total - a.total);
}

function getNextGroup(rows) {
  const dated = rows
    .filter((row) => row.relevantDate)
    .sort((a, b) => new Date(a.relevantDate) - new Date(b.relevantDate));

  if (!dated.length) return null;

  const date = dated[0].relevantDate;
  const sameDate = dated.filter((row) => row.relevantDate === date);
  return {
    date,
    cases: sumCases(sameDate),
    label: sameDate[0].phase
  };
}

function topPhase(phases) {
  const ordered = Array.from(phases.entries()).sort((a, b) => b[1] - a[1]);
  return ordered[0]?.[0] || "Scheduled";
}

function sumCases(rows) {
  return rows.reduce((sum, row) => sum + (Number(row.caseCount) || 0), 0);
}

function mergeRows(existingRows, incomingRows) {
  const map = new Map();

  existingRows.forEach((row) => {
    map.set(fingerprint(row), row);
  });

  incomingRows.forEach((row) => {
    const key = fingerprint(row);
    const existing = map.get(key);
    map.set(key, { ...row, id: existing?.id || row.id });
  });

  return Array.from(map.values());
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
    .map((part) => cleanText(part).toLowerCase())
    .join("|");
}

function phaseClassName(phase) {
  if (phase === "In Transit" || phase === "Trucking") return "transit";
  if (phase === "Port Arrival") return "arrival";
  if (phase === "Attention") return "attention";
  if (phase === "Scheduled") return "scheduled";
  return "";
}

function normalizeDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return toIsoDate(value);
  }

  const text = cleanText(value);
  if (!text) return "";

  const numericDate = text.match(/^(\d{1,2})[\/.-](\d{1,2})(?:[\/.-](\d{2,4}))?$/);
  if (numericDate) {
    const month = Number(numericDate[1]);
    const day = Number(numericDate[2]);
    const rawYear = numericDate[3] ? Number(numericDate[3]) : new Date().getFullYear();
    const year = rawYear < 100 ? 2000 + rawYear : rawYear;
    return toIsoDate(new Date(year, month - 1, day));
  }

  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? "" : toIsoDate(parsed);
}

function startOfDay(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function toIsoDate(date) {
  const normalized = startOfDay(date);
  const year = normalized.getFullYear();
  const month = String(normalized.getMonth() + 1).padStart(2, "0");
  const day = String(normalized.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function formatDate(date) {
  return new Intl.DateTimeFormat("en-US", {
    month: "short",
    day: "numeric",
    year: "numeric"
  }).format(date);
}

function shortDate(isoDate) {
  if (!isoDate) return "";
  const date = new Date(`${isoDate}T00:00:00`);
  return new Intl.DateTimeFormat("en-US", {
    month: "short",
    day: "numeric"
  })
    .format(date)
    .toUpperCase();
}

function formatNumber(value) {
  return new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 }).format(Number(value) || 0);
}

function parseNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const match = cleanText(value).replace(/,/g, "").match(/-?\d+(\.\d+)?/);
  return match ? Number(match[0]) : 0;
}

function cleanText(value) {
  return String(value ?? "").replace(/\r/g, "").trim();
}

function escapeHtml(value) {
  return cleanText(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function buildId() {
  if (window.crypto?.randomUUID) return window.crypto.randomUUID();
  return `row-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function setStatus(message) {
  els.saveState.textContent = message;
}
