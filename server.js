// RaceLine — Order & Stock Desk
// A small self-contained server: run it on one office PC, everyone else
// connects to it over the WiFi network from their phone or computer.

const express = require("express");
const fs = require("fs");
const path = require("path");
const os = require("os");
const crypto = require("crypto");
const multer = require("multer");
const XLSX = require("xlsx");

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 5 * 1024 * 1024 } });

// Reads the first sheet of an uploaded .xlsx/.xls file into plain row objects.
function parseExcelRows(buffer) {
  const workbook = XLSX.read(buffer, { type: "buffer" });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(sheet, { defval: "" });
}

// Looks up a value in a spreadsheet row by trying several possible column
// header spellings, since real-world exports never agree on header names.
function pick(row, ...aliases) {
  const keys = Object.keys(row);
  for (const alias of aliases) {
    const found = keys.find((k) => k.trim().toLowerCase() === alias.trim().toLowerCase());
    if (found !== undefined && row[found] !== "") return row[found];
  }
  return "";
}

const PORT = process.env.PORT || 3000;
const DATA_FILE = path.join(__dirname, "data", "data.json");
const SEED_FILE = path.join(__dirname, "data", "seed.json");

// ---------- storage ----------
// Plain JSON file on disk. A tiny in-process write queue keeps concurrent
// requests from corrupting the file (no native/database dependency needed).

if (!fs.existsSync(DATA_FILE)) {
  fs.copyFileSync(SEED_FILE, DATA_FILE);
}

function readData() {
  const data = JSON.parse(fs.readFileSync(DATA_FILE, "utf8"));
  // Backfill for data files saved before "other supplier" purchases existed.
  if (!Array.isArray(data.otherRequests)) data.otherRequests = [];
  if (!Array.isArray(data.otherStockIns)) data.otherStockIns = [];
  if (!Array.isArray(data.otherWriteOffs)) data.otherWriteOffs = [];
  return data;
}

let writeQueue = Promise.resolve();
function withData(mutator) {
  writeQueue = writeQueue.then(async () => {
    const data = readData();
    const result = mutator(data);
    fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2));
    return result;
  });
  return writeQueue;
}

function newId(prefix) {
  const rand = crypto.randomUUID ? crypto.randomUUID() : Math.random().toString(36).slice(2);
  return `${prefix}-${rand}`;
}

function nowIso() {
  return new Date().toISOString();
}

// ---------- app ----------
const app = express();
app.use(express.json({ limit: "2mb" }));
app.use(express.static(path.join(__dirname, "public")));

// Whole shared state in one call — dataset is small, polling this is cheap.
app.get("/api/state", (req, res) => {
  res.json(readData());
});

app.get("/api/export", (req, res) => {
  res.setHeader("Content-Disposition", 'attachment; filename="raceline-backup.json"');
  res.json(readData());
});

// Restores the whole dataset from a previously-exported backup file — used
// to carry real data across a redeploy on a host without a persistent disk
// (a fresh deploy resets data/data.json to the seed file).
app.post("/api/restore", async (req, res) => {
  const incoming = req.body;
  if (!incoming || typeof incoming !== "object" || !Array.isArray(incoming.team) || !Array.isArray(incoming.products)) {
    return res.status(400).json({ error: "That doesn't look like a RaceLine backup file." });
  }
  await withData((data) => {
    for (const key of Object.keys(data)) delete data[key];
    Object.assign(data, incoming);
    if (!Array.isArray(data.otherRequests)) data.otherRequests = [];
    if (!Array.isArray(data.otherStockIns)) data.otherStockIns = [];
    if (!Array.isArray(data.otherWriteOffs)) data.otherWriteOffs = [];
  });
  res.json({ ok: true });
});

// ---- team ----
app.post("/api/team", async (req, res) => {
  const { name, role } = req.body || {};
  if (!name || !name.trim()) return res.status(400).json({ error: "Name is required." });
  if (!["owner", "order-taker", "biller", "dispatch"].includes(role)) {
    return res.status(400).json({ error: "Invalid role." });
  }
  const member = { id: newId("team"), name: name.trim(), role };
  await withData((data) => data.team.push(member));
  res.json(member);
});

app.put("/api/team/:id", async (req, res) => {
  const { name, role } = req.body || {};
  const result = await withData((data) => {
    const m = data.team.find((t) => t.id === req.params.id);
    if (!m) return null;
    if (name && name.trim()) m.name = name.trim();
    if (role) m.role = role;
    return m;
  });
  if (!result) return res.status(404).json({ error: "Team member not found." });
  res.json(result);
});

app.delete("/api/team/:id", async (req, res) => {
  await withData((data) => {
    data.team = data.team.filter((t) => t.id !== req.params.id);
  });
  res.json({ ok: true });
});

// Orders still waiting on a part (taken = promised, billed = confirmed but
// not yet shipped), oldest first — first come, first served.
function pendingDemand(data, partNumber) {
  return data.orders
    .filter((o) => (o.status === "taken" || o.status === "billed") && o.items.some((it) => it.partNumber === partNumber))
    .map((o) => {
      const item = o.items.find((it) => it.partNumber === partNumber);
      return { orderId: o.id, customerName: o.customerName, qty: item.qty, status: o.status, takenAt: o.takenAt };
    })
    .sort((a, b) => new Date(a.takenAt) - new Date(b.takenAt));
}

// ---- products ----
function normalizeProduct(p) {
  const partNumber = String(p.partNumber || p.id || "").trim();
  return {
    id: partNumber,
    partNumber,
    description: String(p.description || "").trim(),
    category: String(p.category || "").trim(),
    unit: String(p.unit || "pcs").trim() || "pcs",
    stock: Number(p.stock) || 0,
    reorderLevel: Number(p.reorderLevel) || 0,
    location: String(p.location || "").trim(),
    sample: false,
  };
}

app.post("/api/products", async (req, res) => {
  const body = req.body;
  const list = Array.isArray(body) ? body : [body];
  const errors = [];
  const saved = await withData((data) => {
    const out = [];
    for (const raw of list) {
      const p = normalizeProduct(raw);
      if (!p.partNumber) {
        errors.push(`Skipped a row with no part number.`);
        continue;
      }
      const existing = data.products.find((x) => x.id === p.id);
      if (existing) {
        Object.assign(existing, p, { sample: existing.sample && p.stock === existing.stock });
        out.push(existing);
      } else {
        data.products.push(p);
        out.push(p);
      }
    }
    return out;
  });
  res.json({ saved, errors });
});

app.put("/api/products/:id", async (req, res) => {
  const result = await withData((data) => {
    const p = data.products.find((x) => x.id === req.params.id);
    if (!p) return null;
    const { description, category, unit, stock, reorderLevel, location } = req.body || {};
    if (description !== undefined) p.description = String(description).trim();
    if (category !== undefined) p.category = String(category).trim();
    if (unit !== undefined) p.unit = String(unit).trim();
    if (stock !== undefined) p.stock = Number(stock) || 0;
    if (reorderLevel !== undefined) p.reorderLevel = Number(reorderLevel) || 0;
    if (location !== undefined) p.location = String(location).trim();
    p.sample = false;
    return p;
  });
  if (!result) return res.status(404).json({ error: "Item not found." });
  res.json(result);
});

app.delete("/api/products/:id", async (req, res) => {
  await withData((data) => {
    data.products = data.products.filter((p) => p.id !== req.params.id);
  });
  res.json({ ok: true });
});

// Bulk-import products from an uploaded Excel file (.xlsx/.xls). Tries a
// handful of common header spellings so real-world exports don't need
// reformatting first.
app.post("/api/products/import-excel", upload.single("file"), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: "No file uploaded." });
  let rows;
  try {
    rows = parseExcelRows(req.file.buffer);
  } catch (e) {
    return res.status(400).json({ error: "Could not read that file. Make sure it's a .xlsx or .xls file." });
  }
  const errors = [];
  const saved = await withData((data) => {
    const out = [];
    for (const row of rows) {
      const partNumber = String(pick(row, "Part Number", "Part No", "PartNo", "SKU", "Item Code", "Item")).trim();
      if (!partNumber) continue;
      const p = normalizeProduct({
        partNumber,
        description: pick(row, "Description", "Item Name", "Name"),
        category: pick(row, "Category", "Group"),
        unit: pick(row, "Unit", "UOM"),
        stock: pick(row, "Stock", "Qty", "Quantity", "Closing Stock", "Closing Balance"),
        reorderLevel: pick(row, "Reorder Level", "Reorder", "Min Stock"),
        location: pick(row, "Location", "Godown", "Rack"),
      });
      const existing = data.products.find((x) => x.id === p.id);
      if (existing) {
        Object.assign(existing, p, { sample: false });
        out.push(existing);
      } else {
        data.products.push(p);
        out.push(p);
      }
    }
    return out;
  });
  if (saved.length === 0) errors.push("No rows with a recognizable part number were found in that file.");
  res.json({ saved, errors, rowsRead: rows.length });
});

app.post("/api/products/clear-samples", async (req, res) => {
  const removed = await withData((data) => {
    const before = data.products.length;
    data.products = data.products.filter((p) => !p.sample);
    return before - data.products.length;
  });
  res.json({ removed });
});

// ---- customers ----
function normalizeCustomer(c) {
  return {
    id: c.id && !c.id.startsWith("new") ? c.id : newId("cust"),
    name: String(c.name || "").trim(),
    contact: String(c.contact || "").trim(),
    area: String(c.area || "").trim(),
    creditPeriodDays: c.creditPeriodDays !== undefined ? Number(c.creditPeriodDays) || 0 : 0,
    outstandingAmount: c.outstandingAmount !== undefined ? Number(c.outstandingAmount) || 0 : 0,
    outstandingAsOf: c.outstandingAmount !== undefined ? nowIso() : null,
    sample: false,
  };
}

app.post("/api/customers", async (req, res) => {
  const body = req.body;
  const list = Array.isArray(body) ? body : [body];
  const saved = await withData((data) => {
    const out = [];
    for (const raw of list) {
      const c = normalizeCustomer(raw);
      if (!c.name) continue;
      data.customers.push(c);
      out.push(c);
    }
    return out;
  });
  res.json({ saved });
});

app.put("/api/customers/:id", async (req, res) => {
  const result = await withData((data) => {
    const c = data.customers.find((x) => x.id === req.params.id);
    if (!c) return null;
    const { name, contact, area, creditPeriodDays, outstandingAmount } = req.body || {};
    if (name !== undefined) c.name = String(name).trim();
    if (contact !== undefined) c.contact = String(contact).trim();
    if (area !== undefined) c.area = String(area).trim();
    if (creditPeriodDays !== undefined) c.creditPeriodDays = Number(creditPeriodDays) || 0;
    if (outstandingAmount !== undefined) {
      c.outstandingAmount = Number(outstandingAmount) || 0;
      c.outstandingAsOf = nowIso();
    }
    c.sample = false;
    return c;
  });
  if (!result) return res.status(404).json({ error: "Customer not found." });
  res.json(result);
});

app.post("/api/customers/clear-samples", async (req, res) => {
  const removed = await withData((data) => {
    const before = data.customers.length;
    data.customers = data.customers.filter((c) => !c.sample);
    return before - data.customers.length;
  });
  res.json({ removed });
});

// Bulk-import customers from an uploaded Excel file.
app.post("/api/customers/import-excel", upload.single("file"), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: "No file uploaded." });
  let rows;
  try {
    rows = parseExcelRows(req.file.buffer);
  } catch (e) {
    return res.status(400).json({ error: "Could not read that file. Make sure it's a .xlsx or .xls file." });
  }
  const saved = await withData((data) => {
    const out = [];
    for (const row of rows) {
      const name = String(pick(row, "Name", "Customer Name", "Party", "Party Name")).trim();
      if (!name) continue;
      const c = normalizeCustomer({
        name,
        contact: pick(row, "Contact", "Phone", "Mobile"),
        area: pick(row, "Area", "City", "Location"),
        creditPeriodDays: pick(row, "Credit Period", "Credit Days"),
        outstandingAmount: pick(row, "Outstanding", "Outstanding Amount", "Closing Balance", "Balance"),
      });
      data.customers.push(c);
      out.push(c);
    }
    return out;
  });
  res.json({ saved, rowsRead: rows.length });
});

// ---- orders ----
app.post("/api/orders", async (req, res) => {
  const { customerId, newCustomer, items, takenBy, notes } = req.body || {};
  if (!Array.isArray(items) || items.length === 0) {
    return res.status(400).json({ error: "Add at least one item." });
  }
  if (!takenBy) return res.status(400).json({ error: "Missing who is taking the order." });

  const result = await withData((data) => {
    let customer = null;
    if (customerId === "__new__" && newCustomer && newCustomer.name) {
      customer = normalizeCustomer(newCustomer);
      data.customers.push(customer);
    } else {
      customer = data.customers.find((c) => c.id === customerId);
    }
    if (!customer) return { error: "Select or add a customer." };

    const warnings = [];
    const orderItems = [];
    for (const it of items) {
      const product = data.products.find((p) => p.id === it.partNumber);
      const qty = Number(it.qty) || 0;
      if (!product || qty <= 0) continue;
      if (qty > product.stock) {
        warnings.push(`${product.partNumber}: only ${product.stock} ${product.unit} on record, order has ${qty}.`);
      }
      orderItems.push({ partNumber: product.partNumber, description: product.description, qty, unit: product.unit });
    }
    if (orderItems.length === 0) return { error: "No valid items on this order." };

    const order = {
      id: newId("ord"),
      customerId: customer.id,
      customerName: customer.name,
      items: orderItems,
      status: "taken",
      notes: (notes || "").trim(),
      takenBy,
      takenAt: nowIso(),
      billedBy: null,
      billedAt: null,
      tallyInvoiceNo: null,
      dispatchedBy: null,
      dispatchedAt: null,
      cancelledBy: null,
      cancelledAt: null,
    };
    data.orders.unshift(order);
    return { order, warnings };
  });

  if (result.error) return res.status(400).json({ error: result.error });
  res.json(result);
});

app.put("/api/orders/:id/bill", async (req, res) => {
  const { billedBy, tallyInvoiceNo } = req.body || {};
  const result = await withData((data) => {
    const o = data.orders.find((x) => x.id === req.params.id);
    if (!o) return { error: "Order not found." };
    if (o.status !== "taken") return { error: `Order is already ${o.status}.` };
    o.status = "billed";
    o.billedBy = billedBy || "Unknown";
    o.billedAt = nowIso();
    o.tallyInvoiceNo = (tallyInvoiceNo || "").trim() || null;
    return { order: o };
  });
  if (result.error) return res.status(400).json({ error: result.error });
  res.json(result.order);
});

app.put("/api/orders/:id/dispatch", async (req, res) => {
  const { dispatchedBy } = req.body || {};
  const result = await withData((data) => {
    const o = data.orders.find((x) => x.id === req.params.id);
    if (!o) return { error: "Order not found." };
    if (o.status !== "billed") return { error: `Order must be billed before dispatch (currently ${o.status}).` };
    for (const item of o.items) {
      const product = data.products.find((p) => p.id === item.partNumber);
      if (product) product.stock = Math.max(0, product.stock - item.qty);
    }
    o.status = "dispatched";
    o.dispatchedBy = dispatchedBy || "Unknown";
    o.dispatchedAt = nowIso();
    return { order: o };
  });
  if (result.error) return res.status(400).json({ error: result.error });
  res.json(result.order);
});

app.put("/api/orders/:id/cancel", async (req, res) => {
  const { cancelledBy } = req.body || {};
  const result = await withData((data) => {
    const o = data.orders.find((x) => x.id === req.params.id);
    if (!o) return { error: "Order not found." };
    if (o.status === "dispatched") return { error: "Already dispatched orders can't be cancelled here." };
    if (o.status === "cancelled") return { error: "Already cancelled." };
    o.status = "cancelled";
    o.cancelledBy = cancelledBy || "Unknown";
    o.cancelledAt = nowIso();
    return { order: o };
  });
  if (result.error) return res.status(400).json({ error: result.error });
  res.json(result.order);
});

// ---- SKF purchases: requests sent, stock received, pending balance ----
// No PO numbers exist on either side (orders go out as an emailed Excel
// sheet, SKF's bill references nothing back), so matching is done by
// running per-part-number ledger instead of matching documents:
//   pending = everything ever requested − everything ever received or written off
function normalizeLineItems(data, items) {
  const out = [];
  for (const it of items || []) {
    const product = data.products.find((p) => p.id === it.partNumber);
    const qty = Number(it.qty) || 0;
    if (!product || qty <= 0) continue;
    out.push({ partNumber: product.partNumber, description: product.description, unit: product.unit, qty });
  }
  return out;
}

app.post("/api/skf-requests", async (req, res) => {
  const { items, loggedBy, note } = req.body || {};
  const result = await withData((data) => {
    const lineItems = normalizeLineItems(data, items);
    if (lineItems.length === 0) return { error: "Add at least one part number." };
    const request = { id: newId("skfreq"), date: nowIso(), loggedBy: loggedBy || "Unknown", note: (note || "").trim(), items: lineItems };
    data.skfRequests.push(request);
    return { request };
  });
  if (result.error) return res.status(400).json({ error: result.error });
  res.json(result.request);
});

app.post("/api/stock-in", async (req, res) => {
  const { items, receivedBy, note } = req.body || {};
  const result = await withData((data) => {
    const lineItems = normalizeLineItems(data, items);
    if (lineItems.length === 0) return { error: "Add at least one part number." };
    const allocations = [];
    for (const item of lineItems) {
      const product = data.products.find((p) => p.id === item.partNumber);
      product.stock += item.qty;
      product.sample = false;
      const demand = pendingDemand(data, item.partNumber);
      let remaining = item.qty;
      const covered = [];
      for (const d of demand) {
        if (remaining <= 0) break;
        const take = Math.min(remaining, d.qty);
        remaining -= take;
        covered.push({ ...d, covered: take });
      }
      allocations.push({ partNumber: item.partNumber, description: item.description, qtyReceived: item.qty, covered, leftover: remaining });
    }
    const stockIn = { id: newId("stockin"), date: nowIso(), receivedBy: receivedBy || "Unknown", note: (note || "").trim(), items: lineItems };
    data.stockIns.push(stockIn);
    return { stockIn, allocations };
  });
  if (result.error) return res.status(400).json({ error: result.error });
  res.json(result);
});

app.post("/api/skf-writeoffs", async (req, res) => {
  const { partNumber, qty, note, writtenOffBy } = req.body || {};
  const result = await withData((data) => {
    const product = data.products.find((p) => p.id === partNumber);
    if (!product) return { error: "Unknown part number." };
    const q = Number(qty) || 0;
    if (q <= 0) return { error: "Enter a quantity to write off." };
    const writeOff = { id: newId("wo"), date: nowIso(), partNumber, description: product.description, qty: q, note: (note || "").trim(), writtenOffBy: writtenOffBy || "Unknown" };
    data.skfWriteOffs.push(writeOff);
    return { writeOff };
  });
  if (result.error) return res.status(400).json({ error: result.error });
  res.json(result.writeOff);
});

// ---- Other-supplier purchases ----
// Same "no PO numbers, running ledger" pattern as SKF above, but for any
// other party (they sell the same catalog part numbers as SKF). Kept as
// entirely separate collections/endpoints from the SKF ones so the
// existing SKF flow can't be affected by this.
app.post("/api/other-requests", async (req, res) => {
  const { party, items, loggedBy, note } = req.body || {};
  const p = String(party || "").trim();
  const result = await withData((data) => {
    if (!p) return { error: "Enter a supplier/party name." };
    const lineItems = normalizeLineItems(data, items);
    if (lineItems.length === 0) return { error: "Add at least one part number." };
    const request = { id: newId("othreq"), party: p, date: nowIso(), loggedBy: loggedBy || "Unknown", note: (note || "").trim(), items: lineItems };
    data.otherRequests.push(request);
    return { request };
  });
  if (result.error) return res.status(400).json({ error: result.error });
  res.json(result.request);
});

app.post("/api/other-stock-in", async (req, res) => {
  const { party, items, receivedBy, note } = req.body || {};
  const p = String(party || "").trim();
  const result = await withData((data) => {
    if (!p) return { error: "Enter a supplier/party name." };
    const lineItems = normalizeLineItems(data, items);
    if (lineItems.length === 0) return { error: "Add at least one part number." };
    const allocations = [];
    for (const item of lineItems) {
      const product = data.products.find((x) => x.id === item.partNumber);
      product.stock += item.qty;
      product.sample = false;
      const demand = pendingDemand(data, item.partNumber);
      let remaining = item.qty;
      const covered = [];
      for (const d of demand) {
        if (remaining <= 0) break;
        const take = Math.min(remaining, d.qty);
        remaining -= take;
        covered.push({ ...d, covered: take });
      }
      allocations.push({ partNumber: item.partNumber, description: item.description, qtyReceived: item.qty, covered, leftover: remaining });
    }
    const stockIn = { id: newId("othstockin"), party: p, date: nowIso(), receivedBy: receivedBy || "Unknown", note: (note || "").trim(), items: lineItems };
    data.otherStockIns.push(stockIn);
    return { stockIn, allocations };
  });
  if (result.error) return res.status(400).json({ error: result.error });
  res.json(result);
});

app.post("/api/other-writeoffs", async (req, res) => {
  const { party, partNumber, qty, note, writtenOffBy } = req.body || {};
  const p = String(party || "").trim();
  const result = await withData((data) => {
    if (!p) return { error: "Enter a supplier/party name." };
    const product = data.products.find((x) => x.id === partNumber);
    if (!product) return { error: "Unknown part number." };
    const q = Number(qty) || 0;
    if (q <= 0) return { error: "Enter a quantity to write off." };
    const writeOff = { id: newId("othwo"), party: p, date: nowIso(), partNumber, description: product.description, qty: q, note: (note || "").trim(), writtenOffBy: writtenOffBy || "Unknown" };
    data.otherWriteOffs.push(writeOff);
    return { writeOff };
  });
  if (result.error) return res.status(400).json({ error: result.error });
  res.json(result.writeOff);
});

// Bulk-import an "other supplier" request straight from an emailed/WhatsApp
// Excel sheet — one file becomes one logged request against that party.
app.post("/api/other-requests/import-excel", upload.single("file"), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: "No file uploaded." });
  const party = String((req.body && req.body.party) || "").trim();
  if (!party) return res.status(400).json({ error: "Enter a supplier/party name." });
  let rows;
  try {
    rows = parseExcelRows(req.file.buffer);
  } catch (e) {
    return res.status(400).json({ error: "Could not read that file. Make sure it's a .xlsx or .xls file." });
  }
  const result = await withData((data) => {
    const items = [];
    for (const row of rows) {
      const partNumber = String(pick(row, "Part Number", "Part No", "PartNo", "SKU", "Item Code", "Item")).trim();
      const qty = Number(pick(row, "Qty", "Quantity", "Order Qty")) || 0;
      if (!partNumber || qty <= 0) continue;
      items.push({ partNumber, qty });
    }
    const lineItems = normalizeLineItems(data, items);
    if (lineItems.length === 0) return { error: "No valid part numbers with quantities found in that file." };
    const request = { id: newId("othreq"), party, date: nowIso(), loggedBy: (req.body && req.body.loggedBy) || "Unknown", note: "Imported from Excel", items: lineItems };
    data.otherRequests.push(request);
    return { request, rowsRead: rows.length };
  });
  if (result.error) return res.status(400).json({ error: result.error });
  res.json(result);
});

app.listen(PORT, "0.0.0.0", () => {
  const nets = os.networkInterfaces();
  const addresses = [];
  for (const name of Object.keys(nets)) {
    for (const net of nets[name]) {
      if (net.family === "IPv4" && !net.internal) addresses.push(net.address);
    }
  }
  console.log("");
  console.log("  RaceLine is running.");
  console.log("");
  console.log(`  On this PC:      http://localhost:${PORT}`);
  for (const addr of addresses) {
    console.log(`  On the network:  http://${addr}:${PORT}`);
  }
  console.log("");
  console.log("  Open the network address on any phone or PC on the same WiFi.");
  console.log("  Leave this window open — closing it stops RaceLine.");
  console.log("");
});
