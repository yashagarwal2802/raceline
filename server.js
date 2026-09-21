// RaceLine — Order & Stock Desk
// A small self-contained server: run it on one office PC, everyone else
// connects to it over the WiFi network from their phone or computer.

const express = require("express");
const fs = require("fs");
const path = require("path");
const os = require("os");
const crypto = require("crypto");
const net = require("net");
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

// One-time auth secret bootstrap. Generated once and saved into the data
// file so login tokens stay valid across server restarts/redeploys — this
// runs a single time at startup, never inside readData() (which runs on
// every request), otherwise every token would be invalidated constantly.
(function ensureAuthSecret() {
  const raw = JSON.parse(fs.readFileSync(DATA_FILE, "utf8"));
  if (!raw.authSecret) {
    raw.authSecret = crypto.randomBytes(32).toString("hex");
    fs.writeFileSync(DATA_FILE, JSON.stringify(raw, null, 2));
  }
})();

function readData() {
  const data = JSON.parse(fs.readFileSync(DATA_FILE, "utf8"));
  // Backfill for data files saved before "other supplier" purchases existed.
  if (!Array.isArray(data.otherRequests)) data.otherRequests = [];
  if (!Array.isArray(data.otherStockIns)) data.otherStockIns = [];
  if (!Array.isArray(data.otherWriteOffs)) data.otherWriteOffs = [];
  // Backfill per-person login fields for team members saved before PIN
  // login existed.
  for (const m of data.team || []) {
    if (m.pinHash === undefined) m.pinHash = null;
    if (m.pinSalt === undefined) m.pinSalt = null;
    if (m.deviceId === undefined) m.deviceId = null;
    if (m.pendingDeviceId === undefined) m.pendingDeviceId = null;
    if (m.pendingDeviceRequestedAt === undefined) m.pendingDeviceRequestedAt = null;
    if (typeof m.sessionVersion !== "number") m.sessionVersion = 1;
  }
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

// ---------- login / PIN auth ----------
// Every team member has a 6-digit PIN and is locked to one device (browser)
// until the Owner approves a change. Sessions are stateless HMAC-signed
// tokens that expire after 24 hours; bumping a person's sessionVersion
// invalidates every token already issued to them (used for PIN resets,
// device-change approval and instant "revoke access").

function hashPin(pin, salt) {
  const s = salt || crypto.randomBytes(16).toString("hex");
  const hash = crypto.scryptSync(String(pin), s, 64).toString("hex");
  return { hash, salt: s };
}

function pinMatches(pin, member) {
  if (!member.pinHash || !member.pinSalt) return false;
  const { hash } = hashPin(pin, member.pinSalt);
  const a = Buffer.from(hash, "hex");
  const b = Buffer.from(member.pinHash, "hex");
  return a.length === b.length && crypto.timingSafeEqual(a, b);
}

function signToken(memberId, issuedAt, sessionVersion, secret) {
  const payload = `${memberId}.${issuedAt}.${sessionVersion}`;
  const sig = crypto.createHmac("sha256", secret).update(payload).digest("hex");
  return `${payload}.${sig}`;
}

function verifyToken(token, data) {
  if (!token || typeof token !== "string") return null;
  const parts = token.split(".");
  if (parts.length !== 4) return null;
  const [memberId, issuedAtStr, sessionVersionStr, sig] = parts;
  const payload = `${memberId}.${issuedAtStr}.${sessionVersionStr}`;
  const expectedSig = crypto.createHmac("sha256", data.authSecret).update(payload).digest("hex");
  const sigBuf = Buffer.from(sig, "hex");
  const expBuf = Buffer.from(expectedSig, "hex");
  if (sigBuf.length !== expBuf.length || !crypto.timingSafeEqual(sigBuf, expBuf)) return null;
  const issuedAt = Number(issuedAtStr);
  if (!issuedAt || Date.now() - issuedAt > 24 * 60 * 60 * 1000) return null; // expired after 24h
  const member = data.team.find((m) => m.id === memberId);
  if (!member) return null;
  if (member.sessionVersion !== Number(sessionVersionStr)) return null; // reset/revoked since
  return member;
}

// What the picker/team screens are allowed to see about each person — never
// the PIN hash/salt, device id, or the server's auth secret.
function publicTeamMember(m) {
  return {
    id: m.id,
    name: m.name,
    role: m.role,
    hasPin: !!m.pinHash,
    pendingDevice: !!m.pendingDeviceId,
    pendingDeviceRequestedAt: m.pendingDeviceRequestedAt || null,
  };
}

function requireOwner(req, res, next) {
  if (!req.member || req.member.role !== "owner") return res.status(403).json({ error: "Owner only." });
  next();
}

// Temporary diagnostic: checks whether THIS SERVER (i.e. Render, not your
// browser/laptop) can open a raw TCP connection to the Tally cloud server's
// ODBC port. Visit this URL directly in a browser to test it — it needs no
// login since it exposes nothing but a yes/no reachability result. Defaults
// to the Tally-on-cloud address/port already on file; pass ?host=...&port=...
// to test a different one.
app.get("/api/tally-test", (req, res) => {
  const host = String(req.query.host || "v60020.22164.tallyprimecloud.in");
  const port = Number(req.query.port) || 9537;
  const start = Date.now();
  const socket = new net.Socket();
  let settled = false;
  const finish = (result) => {
    if (settled) return;
    settled = true;
    socket.destroy();
    res.json({ host, port, ms: Date.now() - start, ...result });
  };
  socket.setTimeout(8000);
  socket.once("connect", () => finish({ ok: true, message: "Connected! Render can reach this host and port." }));
  socket.once("timeout", () => finish({ ok: false, message: "Timed out — nothing answered. Usually means it's still blocked/firewalled from this server, or the address/port is wrong." }));
  socket.once("error", (err) => finish({ ok: false, message: `Connection error: ${err.message}` }));
  socket.connect(port, host);
});

// Shown on the picker screen before anyone is logged in — no sensitive data.
app.get("/api/team-roster", (req, res) => {
  const data = readData();
  res.json({ team: data.team.map(publicTeamMember) });
});

// First-time PIN setup for a team member who doesn't have one yet (existing
// team members created before PIN login existed, or someone newly added by
// the Owner). Also binds this device as their one allowed device. Refuses
// once a PIN already exists — from then on /api/login is the only way in,
// and only the Owner can reset a PIN via /api/team/:id/pin.
app.post("/api/team/:id/set-own-pin", async (req, res) => {
  const { pin, deviceId } = req.body || {};
  if (!/^\d{6}$/.test(String(pin || ""))) return res.status(400).json({ error: "PIN must be exactly 6 digits." });
  if (!deviceId) return res.status(400).json({ error: "Missing device id." });
  const result = await withData((data) => {
    const m = data.team.find((t) => t.id === req.params.id);
    if (!m) return { error: "Team member not found." };
    if (m.pinHash) return { error: "A PIN is already set for this person. Ask the Owner to reset it if needed." };
    const { hash, salt } = hashPin(pin);
    m.pinHash = hash;
    m.pinSalt = salt;
    m.deviceId = deviceId;
    m.sessionVersion = (m.sessionVersion || 1) + 1;
    const token = signToken(m.id, Date.now(), m.sessionVersion, data.authSecret);
    return { token, member: publicTeamMember(m) };
  });
  if (result.error) return res.status(400).json({ error: result.error });
  res.json(result);
});

app.post("/api/login", async (req, res) => {
  const { memberId, pin, deviceId } = req.body || {};
  if (!deviceId) return res.status(400).json({ error: "Missing device id." });
  const result = await withData((data) => {
    const m = data.team.find((t) => t.id === memberId);
    if (!m) return { error: "not-found" };
    if (!m.pinHash) return { error: "no-pin" };
    if (!pinMatches(pin, m)) return { error: "Incorrect PIN." };
    if (!m.deviceId) {
      m.deviceId = deviceId; // first login after a device change/reset — bind it
    } else if (m.deviceId !== deviceId) {
      m.pendingDeviceId = deviceId;
      m.pendingDeviceRequestedAt = nowIso();
      return { error: "device-pending" };
    }
    const token = signToken(m.id, Date.now(), m.sessionVersion, data.authSecret);
    return { token, member: publicTeamMember(m) };
  });
  if (result.error === "not-found") return res.status(404).json({ error: "Team member not found." });
  if (result.error === "no-pin") return res.status(400).json({ error: "no-pin" });
  if (result.error === "device-pending") {
    return res.status(403).json({ error: "device-pending", message: "This is a new device for this account. Ask the Owner to approve it from the Team tab." });
  }
  if (result.error) return res.status(400).json({ error: result.error });
  res.json(result);
});

// Everything below this point requires a valid, unexpired, non-revoked
// token — attached by the client as the X-Raceline-Token header.
app.use("/api", (req, res, next) => {
  const data = readData();
  const token = req.header("X-Raceline-Token");
  const member = verifyToken(token, data);
  if (!member) return res.status(401).json({ error: "Please log in again." });
  req.member = member;
  next();
});

// Whole shared state in one call — dataset is small, polling this is cheap.
app.get("/api/state", (req, res) => {
  const data = readData();
  res.json({ ...data, team: data.team.map(publicTeamMember) });
});

app.get("/api/export", requireOwner, (req, res) => {
  res.setHeader("Content-Disposition", 'attachment; filename="raceline-backup.json"');
  res.json(readData());
});

// Restores the whole dataset from a previously-exported backup file — used
// to carry real data across a redeploy on a host without a persistent disk
// (a fresh deploy resets data/data.json to the seed file).
app.post("/api/restore", requireOwner, async (req, res) => {
  const incoming = req.body;
  if (!incoming || typeof incoming !== "object" || !Array.isArray(incoming.team) || !Array.isArray(incoming.products)) {
    return res.status(400).json({ error: "That doesn't look like a RaceLine backup file." });
  }
  await withData((data) => {
    const keepAuthSecret = data.authSecret; // in case this is an older backup taken before PIN login existed
    for (const key of Object.keys(data)) delete data[key];
    Object.assign(data, incoming);
    if (!Array.isArray(data.otherRequests)) data.otherRequests = [];
    if (!Array.isArray(data.otherStockIns)) data.otherStockIns = [];
    if (!Array.isArray(data.otherWriteOffs)) data.otherWriteOffs = [];
    if (!data.authSecret) data.authSecret = keepAuthSecret || crypto.randomBytes(32).toString("hex");
    // An older backup's team members won't have PIN/device fields yet — that's
    // fine, readData() backfills them on every read; everyone just sets a
    // fresh PIN the next time they log in.
  });
  res.json({ ok: true });
});

// ---- team ----
app.post("/api/team", requireOwner, async (req, res) => {
  const { name, role } = req.body || {};
  if (!name || !name.trim()) return res.status(400).json({ error: "Name is required." });
  if (!["owner", "order-taker", "biller", "dispatch"].includes(role)) {
    return res.status(400).json({ error: "Invalid role." });
  }
  const member = {
    id: newId("team"), name: name.trim(), role,
    pinHash: null, pinSalt: null, deviceId: null,
    pendingDeviceId: null, pendingDeviceRequestedAt: null, sessionVersion: 1,
  };
  await withData((data) => data.team.push(member));
  res.json(publicTeamMember(member));
});

app.put("/api/team/:id", requireOwner, async (req, res) => {
  const { name, role } = req.body || {};
  const result = await withData((data) => {
    const m = data.team.find((t) => t.id === req.params.id);
    if (!m) return null;
    if (name && name.trim()) m.name = name.trim();
    if (role) m.role = role;
    return m;
  });
  if (!result) return res.status(404).json({ error: "Team member not found." });
  res.json(publicTeamMember(result));
});

app.delete("/api/team/:id", requireOwner, async (req, res) => {
  await withData((data) => {
    data.team = data.team.filter((t) => t.id !== req.params.id);
  });
  res.json({ ok: true });
});

// Owner resets someone's PIN — also clears their device lock (so the first
// login after a reset binds whatever device they log in from next) and
// bumps sessionVersion so any token they already have stops working.
app.put("/api/team/:id/pin", requireOwner, async (req, res) => {
  const { pin } = req.body || {};
  if (!/^\d{6}$/.test(String(pin || ""))) return res.status(400).json({ error: "PIN must be exactly 6 digits." });
  const result = await withData((data) => {
    const m = data.team.find((t) => t.id === req.params.id);
    if (!m) return { error: "Team member not found." };
    const { hash, salt } = hashPin(pin);
    m.pinHash = hash;
    m.pinSalt = salt;
    m.deviceId = null;
    m.pendingDeviceId = null;
    m.pendingDeviceRequestedAt = null;
    m.sessionVersion = (m.sessionVersion || 1) + 1;
    return { member: m };
  });
  if (result.error) return res.status(400).json({ error: result.error });
  res.json(publicTeamMember(result.member));
});

// Owner approves a pending device change — the new device becomes the one
// allowed device, and any session on the old device is invalidated.
app.post("/api/team/:id/approve-device", requireOwner, async (req, res) => {
  const result = await withData((data) => {
    const m = data.team.find((t) => t.id === req.params.id);
    if (!m) return { error: "Team member not found." };
    if (!m.pendingDeviceId) return { error: "No pending device request for this person." };
    m.deviceId = m.pendingDeviceId;
    m.pendingDeviceId = null;
    m.pendingDeviceRequestedAt = null;
    m.sessionVersion = (m.sessionVersion || 1) + 1;
    return { member: m };
  });
  if (result.error) return res.status(400).json({ error: result.error });
  res.json(publicTeamMember(result.member));
});

app.post("/api/team/:id/deny-device", requireOwner, async (req, res) => {
  const result = await withData((data) => {
    const m = data.team.find((t) => t.id === req.params.id);
    if (!m) return { error: "Team member not found." };
    m.pendingDeviceId = null;
    m.pendingDeviceRequestedAt = null;
    return { member: m };
  });
  if (result.error) return res.status(400).json({ error: result.error });
  res.json(publicTeamMember(result.member));
});

// Instantly logs this person out everywhere (bumps sessionVersion so every
// token already issued to them stops working) and clears their device lock,
// so they'll need a fresh PIN entry — and Owner approval again if that's a
// new device — to get back in.
app.post("/api/team/:id/revoke", requireOwner, async (req, res) => {
  const result = await withData((data) => {
    const m = data.team.find((t) => t.id === req.params.id);
    if (!m) return { error: "Team member not found." };
    m.sessionVersion = (m.sessionVersion || 1) + 1;
    m.deviceId = null;
    m.pendingDeviceId = null;
    m.pendingDeviceRequestedAt = null;
    return { member: m };
  });
  if (result.error) return res.status(400).json({ error: result.error });
  res.json(publicTeamMember(result.member));
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

// Bulk-import a request to SKF straight from the Excel sheet you're about
// to email them — same idea as the other-supplier Excel import, one file
// becomes one logged request.
app.post("/api/skf-requests/import-excel", upload.single("file"), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: "No file uploaded." });
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
    const request = { id: newId("skfreq"), date: nowIso(), loggedBy: (req.body && req.body.loggedBy) || "Unknown", note: "Imported from Excel", items: lineItems };
    data.skfRequests.push(request);
    return { request, rowsRead: rows.length };
  });
  if (result.error) return res.status(400).json({ error: result.error });
  res.json(result);
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
