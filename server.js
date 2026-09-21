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
  // Backfill for the Tally sync — which vouchers have already been applied
  // to stock (by GUID, so re-running a sync never double-counts), plus a
  // running log of what the sync has done for the Dashboard to show.
  if (!Array.isArray(data.tallySyncedVoucherGuids)) data.tallySyncedVoucherGuids = [];
  if (!Array.isArray(data.tallySyncLog)) data.tallySyncLog = [];
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

// ---------- Tally sync ----------
// Confirmed working setup, from testing: v60020.22164.tallyprimecloud.in:9537,
// company "BMA - (from 1-Apr-26)". Talks to Tally's XML/HTTP gateway directly
// (no ODBC driver needed) — a plain HTTP POST of a small TDL request.
const TALLY_HOST = "v60020.22164.tallyprimecloud.in";
const TALLY_PORT = 9537;
const TALLY_COMPANY = "BMA - (from 1-Apr-26)";

// Sends a Tally XML request and resolves with the raw response body.
function queryTally(xmlRequest, { host = TALLY_HOST, port = TALLY_PORT, timeout = 20000 } = {}) {
  const http = require("http");
  return new Promise((resolve, reject) => {
    const request = http.request(
      { host, port, path: "/", method: "POST", headers: { "Content-Type": "text/xml", "Content-Length": Buffer.byteLength(xmlRequest) }, timeout },
      (resp) => {
        let body = "";
        resp.on("data", (chunk) => { body += chunk; });
        resp.on("end", () => resolve(body));
      }
    );
    request.on("timeout", () => { request.destroy(); reject(new Error("Timed out waiting for Tally to respond.")); });
    request.on("error", (err) => reject(err));
    request.write(xmlRequest);
    request.end();
  });
}

// Builds the request for every Sales/Purchase voucher (with line items) in
// a date range, for a given company.
function buildVoucherQueryXml({ company = TALLY_COMPANY, from, to }) {
  return [
    "<ENVELOPE>",
    "<HEADER>",
    "<VERSION>1</VERSION>",
    "<TALLYREQUEST>Export</TALLYREQUEST>",
    "<TYPE>Collection</TYPE>",
    "<ID>SyncVouchers</ID>",
    "</HEADER>",
    "<BODY>",
    "<DESC>",
    "<STATICVARIABLES>",
    "<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>",
    `<SVCURRENTCOMPANY>${company.replace(/&/g, "&amp;")}</SVCURRENTCOMPANY>`,
    `<SVFROMDATE>${from}</SVFROMDATE>`,
    `<SVTODATE>${to}</SVTODATE>`,
    "</STATICVARIABLES>",
    "<TDL>",
    "<TDLMESSAGE>",
    '<COLLECTION NAME="SyncVouchers" ISINITIALIZE="Yes">',
    "<TYPE>Voucher</TYPE>",
    "<FILTER>OnlySalesOrPurchase</FILTER>",
    "<FILTER>DateInRange</FILTER>",
    "<FETCH>DATE</FETCH>",
    "<FETCH>VOUCHERTYPENAME</FETCH>",
    "<FETCH>VOUCHERNUMBER</FETCH>",
    "<FETCH>PARTYLEDGERNAME</FETCH>",
    "<FETCH>GUID</FETCH>",
    "<FETCH>ALLINVENTORYENTRIES.LIST</FETCH>",
    "</COLLECTION>",
    '<SYSTEM TYPE="Formulae" NAME="OnlySalesOrPurchase">$VoucherTypeName = "Sales" OR $VoucherTypeName = "Purchase"</SYSTEM>',
    // A plain Voucher collection ignores SVFROMDATE/SVTODATE on its own — it
    // needs an explicit date filter referencing those same variables.
    '<SYSTEM TYPE="Formulae" NAME="DateInRange">$Date &gt;= ##SVFROMDATE AND $Date &lt;= ##SVTODATE</SYSTEM>',
    "</TDLMESSAGE>",
    "</TDL>",
    "</DESC>",
    "</BODY>",
    "</ENVELOPE>",
  ].join("");
}

// Pulls the leading number out of Tally's compound quantity strings, e.g.
// " 6 NOS = 0 Box" -> 6, "-3 NOS" -> 3 (direction comes from voucher type,
// not the sign here, since Tally's +/- convention varies by voucher/config).
function parseTallyQty(qtyStr) {
  const m = String(qtyStr || "").match(/(-?[\d.]+)/);
  return m ? Math.abs(parseFloat(m[1])) : 0;
}

// Turns Tally's raw XML into a plain-JS list of vouchers with their line
// items. Deliberately simple regex parsing rather than a full XML parser —
// Tally's export is flat enough that this is reliable and needs no new
// dependency; the one wrinkle (BATCHALLOCATIONS.LIST repeating the same
// qty/rate fields inside each inventory entry) is handled by only reading
// up to the first BATCHALLOCATIONS.LIST marker, since the entry's own
// top-level fields always appear before it.
function parseTallyVouchers(xmlBody) {
  const voucherBlocks = xmlBody.match(/<VOUCHER[^>]*>[\s\S]*?<\/VOUCHER>/g) || [];
  const vouchers = [];
  for (const block of voucherBlocks) {
    const guid = (block.match(/<GUID>([\s\S]*?)<\/GUID>/) || [])[1];
    const voucherType = (block.match(/<VOUCHERTYPENAME>([\s\S]*?)<\/VOUCHERTYPENAME>/) || [])[1];
    const voucherNumber = (block.match(/<VOUCHERNUMBER>([\s\S]*?)<\/VOUCHERNUMBER>/) || [])[1];
    const party = (block.match(/<PARTYLEDGERNAME[^>]*>([\s\S]*?)<\/PARTYLEDGERNAME>/) || [])[1];
    const dateRaw = (block.match(/<DATE[^>]*>([\s\S]*?)<\/DATE>/) || [])[1]; // YYYYMMDD
    if (!guid || !voucherType) continue;
    const entryBlocks = block.match(/<ALLINVENTORYENTRIES\.LIST>[\s\S]*?<\/ALLINVENTORYENTRIES\.LIST>/g) || [];
    const items = [];
    for (const entry of entryBlocks) {
      const beforeBatch = entry.split("<BATCHALLOCATIONS.LIST>")[0];
      const stockItemName = (beforeBatch.match(/<STOCKITEMNAME[^>]*>([\s\S]*?)<\/STOCKITEMNAME>/) || [])[1];
      const qtyStr = (beforeBatch.match(/<ACTUALQTY[^>]*>([\s\S]*?)<\/ACTUALQTY>/) || [])[1];
      if (!stockItemName) continue;
      const qty = parseTallyQty(qtyStr);
      if (qty > 0) items.push({ stockItemName: stockItemName.trim(), qty });
    }
    vouchers.push({
      guid, voucherType, voucherNumber: voucherNumber || "", party: (party || "").trim(),
      date: dateRaw ? `${dateRaw.slice(0, 4)}-${dateRaw.slice(4, 6)}-${dateRaw.slice(6, 8)}` : "",
      items,
    });
  }
  return vouchers;
}

// Matches a Tally stock item name to a RaceLine product by exact part
// number (case/whitespace insensitive) — RaceLine's catalog was originally
// imported from Tally, so names should line up directly.
function findProductForStockItem(data, stockItemName) {
  const norm = (s) => String(s || "").trim().toUpperCase();
  return data.products.find((p) => norm(p.partNumber) === norm(stockItemName)) || null;
}

// The core of the sync: for a set of parsed vouchers not yet applied,
// works out (but does not apply, unless apply=true) what would change.
function computeSyncPlan(data, vouchers, { apply } = {}) {
  const already = new Set(data.tallySyncedVoucherGuids);
  const newVouchers = vouchers.filter((v) => !already.has(v.guid));
  const purchases = [];
  const sales = [];
  const unmatched = [];
  for (const v of newVouchers) {
    const lineResults = [];
    for (const item of v.items) {
      const product = findProductForStockItem(data, item.stockItemName);
      if (!product) {
        unmatched.push({ voucherType: v.voucherType, voucherNumber: v.voucherNumber, stockItemName: item.stockItemName, qty: item.qty });
        continue;
      }
      const direction = v.voucherType === "Purchase" ? 1 : -1;
      const before = product.stock;
      const after = v.voucherType === "Purchase" ? before + item.qty : Math.max(0, before - item.qty);
      lineResults.push({ partNumber: product.partNumber, description: product.description, qty: item.qty, stockBefore: before, stockAfter: after });
      if (apply) {
        product.stock = after;
        product.sample = false;
      }
    }
    const entry = { guid: v.guid, voucherType: v.voucherType, voucherNumber: v.voucherNumber, party: v.party, date: v.date, items: lineResults };
    if (v.voucherType === "Purchase") purchases.push(entry);
    else sales.push(entry);
    if (apply) data.tallySyncedVoucherGuids.push(v.guid);
  }
  return { totalVouchersChecked: vouchers.length, newVouchersFound: newVouchers.length, purchases, sales, unmatched };
}

// The date-range filter turned out not to work reliably against Tally's
// plain voucher collection (it silently returned everything regardless of
// the range asked for) — and separately, asking for item-level detail on
// ALL vouchers in one big request came back mostly empty even though the
// data is there. Splitting into two small, reliable steps fixes both:
// 1) fetch every voucher's header info (cheap, proven reliable), 2) fetch
// item-level detail in small batches (~20 at a time, by exact ID) only for
// vouchers we haven't already synced.

function buildVoucherHeadersXml({ company = TALLY_COMPANY } = {}) {
  return [
    "<ENVELOPE>",
    "<HEADER>",
    "<VERSION>1</VERSION>",
    "<TALLYREQUEST>Export</TALLYREQUEST>",
    "<TYPE>Collection</TYPE>",
    "<ID>SyncVoucherHeaders</ID>",
    "</HEADER>",
    "<BODY>",
    "<DESC>",
    "<STATICVARIABLES>",
    "<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>",
    `<SVCURRENTCOMPANY>${company.replace(/&/g, "&amp;")}</SVCURRENTCOMPANY>`,
    "</STATICVARIABLES>",
    "<TDL>",
    "<TDLMESSAGE>",
    '<COLLECTION NAME="SyncVoucherHeaders" ISINITIALIZE="Yes">',
    "<TYPE>Voucher</TYPE>",
    "<FILTER>OnlySalesOrPurchase</FILTER>",
    "<FETCH>DATE</FETCH>",
    "<FETCH>VOUCHERTYPENAME</FETCH>",
    "<FETCH>VOUCHERNUMBER</FETCH>",
    "<FETCH>PARTYLEDGERNAME</FETCH>",
    "<FETCH>GUID</FETCH>",
    "</COLLECTION>",
    '<SYSTEM TYPE="Formulae" NAME="OnlySalesOrPurchase">$VoucherTypeName = "Sales" OR $VoucherTypeName = "Purchase"</SYSTEM>',
    "</TDLMESSAGE>",
    "</TDL>",
    "</DESC>",
    "</BODY>",
    "</ENVELOPE>",
  ].join("");
}

function buildVoucherItemsBatchXml({ company = TALLY_COMPANY, guids }) {
  const condition = guids.map((g) => `$Guid = "${g}"`).join(" OR ");
  return [
    "<ENVELOPE>",
    "<HEADER>",
    "<VERSION>1</VERSION>",
    "<TALLYREQUEST>Export</TALLYREQUEST>",
    "<TYPE>Collection</TYPE>",
    "<ID>SyncVoucherItemsBatch</ID>",
    "</HEADER>",
    "<BODY>",
    "<DESC>",
    "<STATICVARIABLES>",
    "<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>",
    `<SVCURRENTCOMPANY>${company.replace(/&/g, "&amp;")}</SVCURRENTCOMPANY>`,
    "</STATICVARIABLES>",
    "<TDL>",
    "<TDLMESSAGE>",
    '<COLLECTION NAME="SyncVoucherItemsBatch" ISINITIALIZE="Yes">',
    "<TYPE>Voucher</TYPE>",
    "<FILTER>GuidInBatch</FILTER>",
    "<FETCH>GUID</FETCH>",
    "<FETCH>ALLINVENTORYENTRIES.LIST</FETCH>",
    "</COLLECTION>",
    `<SYSTEM TYPE="Formulae" NAME="GuidInBatch">${condition}</SYSTEM>`,
    "</TDLMESSAGE>",
    "</TDL>",
    "</DESC>",
    "</BODY>",
    "</ENVELOPE>",
  ].join("");
}

function parseVoucherHeaders(xmlBody) {
  const voucherBlocks = xmlBody.match(/<VOUCHER[^>]*>[\s\S]*?<\/VOUCHER>/g) || [];
  const out = [];
  for (const block of voucherBlocks) {
    const guid = (block.match(/<GUID>([\s\S]*?)<\/GUID>/) || [])[1];
    const voucherType = (block.match(/<VOUCHERTYPENAME>([\s\S]*?)<\/VOUCHERTYPENAME>/) || [])[1];
    const voucherNumber = (block.match(/<VOUCHERNUMBER>([\s\S]*?)<\/VOUCHERNUMBER>/) || [])[1];
    const party = (block.match(/<PARTYLEDGERNAME[^>]*>([\s\S]*?)<\/PARTYLEDGERNAME>/) || [])[1];
    const dateRaw = (block.match(/<DATE[^>]*>([\s\S]*?)<\/DATE>/) || [])[1];
    if (!guid || !voucherType) continue;
    out.push({
      guid, voucherType, voucherNumber: voucherNumber || "", party: (party || "").trim(),
      date: dateRaw ? `${dateRaw.slice(0, 4)}-${dateRaw.slice(4, 6)}-${dateRaw.slice(6, 8)}` : "",
    });
  }
  return out;
}

function parseVoucherItemsOnly(xmlBody) {
  const voucherBlocks = xmlBody.match(/<VOUCHER[^>]*>[\s\S]*?<\/VOUCHER>/g) || [];
  const map = {};
  for (const block of voucherBlocks) {
    const guid = (block.match(/<GUID>([\s\S]*?)<\/GUID>/) || [])[1];
    if (!guid) continue;
    const entryBlocks = block.match(/<ALLINVENTORYENTRIES\.LIST>[\s\S]*?<\/ALLINVENTORYENTRIES\.LIST>/g) || [];
    const items = [];
    for (const entry of entryBlocks) {
      const beforeBatch = entry.split("<BATCHALLOCATIONS.LIST>")[0];
      const stockItemName = (beforeBatch.match(/<STOCKITEMNAME[^>]*>([\s\S]*?)<\/STOCKITEMNAME>/) || [])[1];
      const qtyStr = (beforeBatch.match(/<ACTUALQTY[^>]*>([\s\S]*?)<\/ACTUALQTY>/) || [])[1];
      if (!stockItemName) continue;
      const qty = parseTallyQty(qtyStr);
      if (qty > 0) items.push({ stockItemName: stockItemName.trim(), qty });
    }
    map[guid] = items;
  }
  return map;
}

function chunkArray(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}

// Runs an async worker over a list with at most `limit` requests to Tally
// in flight at once — used for the individual-voucher retry pass below so
// it doesn't take forever if many vouchers need it, without hammering the
// Tally server with 100+ simultaneous connections either.
async function runWithConcurrency(items, limit, worker) {
  let idx = 0;
  async function next() {
    while (idx < items.length) {
      const i = idx++;
      await worker(items[i], i);
    }
  }
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, next));
}

// The main entry point the sync endpoints use: every voucher not already
// synced, each with its item lines filled in via small batched requests.
// Tally's batched item export has occasionally (not always — seen once out
// of several tests) come back with items missing for vouchers that do have
// them, for reasons that were never fully pinned down. As a safety net,
// anything that comes back empty from the batch pass gets double-checked
// with its own individual request before being accepted as genuinely
// item-less, so an intermittent Tally hiccup can't silently under-count.
async function fetchNewVouchersWithItems(alreadySyncedGuids) {
  const already = new Set(alreadySyncedGuids);
  const headersBody = await queryTally(buildVoucherHeadersXml({}));
  const headers = parseVoucherHeaders(headersBody);
  const newHeaders = headers.filter((h) => !already.has(h.guid));
  const batches = chunkArray(newHeaders, 20);
  const itemsMap = {};
  for (const batch of batches) {
    const body = await queryTally(buildVoucherItemsBatchXml({ guids: batch.map((h) => h.guid) }));
    Object.assign(itemsMap, parseVoucherItemsOnly(body));
  }
  const missing = newHeaders.filter((h) => !(itemsMap[h.guid] && itemsMap[h.guid].length > 0));
  if (missing.length > 0) {
    await runWithConcurrency(missing, 5, async (h) => {
      try {
        const body = await queryTally(buildVoucherItemsBatchXml({ guids: [h.guid] }));
        const single = parseVoucherItemsOnly(body);
        if (single[h.guid] && single[h.guid].length > 0) itemsMap[h.guid] = single[h.guid];
      } catch (e) {
        // Leave it as empty — the sync will show it as an item-less
        // voucher rather than silently drop it or crash the whole sync.
      }
    });
  }
  return newHeaders.map((h) => ({ ...h, items: itemsMap[h.guid] || [], itemsNeededRetry: missing.some((m) => m.guid === h.guid) }));
}

// Diagnostic: shows, in one short summary, which Tally stock item names DO
// and DON'T match a RaceLine part number, without dumping every voucher's
// full detail. Fetches the same not-yet-synced vouchers as the real sync
// preview, then groups their line items by stock item name so repeats only
// count once — much shorter to read than scrolling through 182 vouchers.
app.get("/api/tally-match-test", async (req, res) => {
  try {
    const data = readData();
    const vouchers = await fetchNewVouchersWithItems(data.tallySyncedVoucherGuids);
    const seen = new Map();
    for (const v of vouchers) {
      for (const item of v.items) {
        const key = item.stockItemName;
        if (!seen.has(key)) seen.set(key, { name: key, count: 0, matched: !!findProductForStockItem(data, key) });
        seen.get(key).count += 1;
      }
    }
    const all = Array.from(seen.values());
    const unmatched = all.filter((x) => !x.matched).sort((a, b) => b.count - a.count);
    const matched = all.filter((x) => x.matched);
    res.json({
      ok: true,
      totalVouchersChecked: vouchers.length,
      distinctStockItemNames: all.length,
      matchedDistinctCount: matched.length,
      unmatchedDistinctCount: unmatched.length,
      unmatchedSample: unmatched.slice(0, 25),
      matchedSample: matched.slice(0, 8),
    });
  } catch (e) {
    res.status(500).json({ ok: false, message: e.message });
  }
});

// ---------- Tally catalog baseline sync ----------
// Yash confirmed that Tally's more specific item names (clearance-class
// codes like /C3, pack-size notes like "(PACK 56PC)") are genuinely
// different products, not the same part written differently — so instead
// of loosening the match, RaceLine's product list needs an entry for each
// exact Tally stock item name. This seeds/updates that list directly from
// Tally's own stock item master (name + current closing balance — the
// actual real-time stock number Tally has on file right now), so every
// item lines up exactly for future voucher syncing.
function buildStockItemsXml({ company = TALLY_COMPANY } = {}) {
  return [
    "<ENVELOPE>",
    "<HEADER>",
    "<VERSION>1</VERSION>",
    "<TALLYREQUEST>Export</TALLYREQUEST>",
    "<TYPE>Collection</TYPE>",
    "<ID>FullStockItems</ID>",
    "</HEADER>",
    "<BODY>",
    "<DESC>",
    "<STATICVARIABLES>",
    "<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>",
    `<SVCURRENTCOMPANY>${company.replace(/&/g, "&amp;")}</SVCURRENTCOMPANY>`,
    "</STATICVARIABLES>",
    "<TDL>",
    "<TDLMESSAGE>",
    '<COLLECTION NAME="FullStockItems" ISINITIALIZE="Yes">',
    "<TYPE>StockItem</TYPE>",
    "<FETCH>NAME</FETCH>",
    "<FETCH>CLOSINGBALANCE</FETCH>",
    "</COLLECTION>",
    "</TDLMESSAGE>",
    "</TDL>",
    "</DESC>",
    "</BODY>",
    "</ENVELOPE>",
  ].join("");
}

function parseStockItems(xmlBody) {
  const blocks = xmlBody.match(/<STOCKITEM[^>]*>[\s\S]*?<\/STOCKITEM>/g) || [];
  const out = [];
  for (const block of blocks) {
    const name = (block.match(/<NAME>([\s\S]*?)<\/NAME>/) || [])[1];
    const balRaw = (block.match(/<CLOSINGBALANCE>([\s\S]*?)<\/CLOSINGBALANCE>/) || [])[1];
    if (!name) continue;
    out.push({ name: name.trim(), closingBalance: parseTallyQty(balRaw) });
  }
  return out;
}

function computeCatalogPlan(data, stockItems) {
  const norm = (s) => String(s || "").trim().toUpperCase();
  const newProducts = [];
  const stockChanges = [];
  for (const item of stockItems) {
    const existing = data.products.find((p) => norm(p.partNumber) === norm(item.name));
    if (!existing) newProducts.push({ partNumber: item.name, stock: item.closingBalance });
    else if (Number(existing.stock) !== item.closingBalance) {
      stockChanges.push({ partNumber: existing.partNumber, stockBefore: existing.stock, stockAfter: item.closingBalance });
    }
  }
  return { totalTallyItems: stockItems.length, newProducts, stockChanges };
}

// Dry run — shows exactly what the catalog baseline sync would do (how many
// brand-new products it would create, how many existing stock numbers it
// would correct) without changing anything. Safe to run anytime.
app.get("/api/tally-catalog-preview", async (req, res) => {
  try {
    const body = await queryTally(buildStockItemsXml({}));
    const stockItems = parseStockItems(body);
    const data = readData();
    const plan = computeCatalogPlan(data, stockItems);
    res.json({
      ok: true,
      totalTallyItems: plan.totalTallyItems,
      newProductsCount: plan.newProducts.length,
      stockChangesCount: plan.stockChanges.length,
      newProductsSample: plan.newProducts.slice(0, 20),
      stockChangesSample: plan.stockChanges.slice(0, 20),
    });
  } catch (e) {
    res.status(500).json({ ok: false, message: e.message });
  }
});

// Actually applies the catalog baseline: creates a RaceLine product for
// every Tally stock item that doesn't already have one (using Tally's exact
// name as the part number), and corrects the stock number on every existing
// product to match Tally's current closing balance. Also marks every
// Purchase/Sales voucher that exists in Tally right now as already synced —
// their effect is already baked into today's closing balance, so the
// regular voucher sync should only apply vouchers created after this point,
// never these. Owner only, since it changes real stock for many items.
app.post("/api/tally-catalog-sync", async (req, res) => {
  const token = req.header("X-Raceline-Token");
  const data0 = readData();
  const member = verifyToken(token, data0);
  if (!member || member.role !== "owner") return res.status(403).json({ error: "Owner only — log in and try again from the Dashboard." });
  try {
    const stockBody = await queryTally(buildStockItemsXml({}));
    const stockItems = parseStockItems(stockBody);
    const headersBody = await queryTally(buildVoucherHeadersXml({}));
    const allHeaderGuids = parseVoucherHeaders(headersBody).map((h) => h.guid);
    const result = await withData((data) => {
      const norm = (s) => String(s || "").trim().toUpperCase();
      let created = 0;
      let updated = 0;
      for (const item of stockItems) {
        const existing = data.products.find((p) => norm(p.partNumber) === norm(item.name));
        if (!existing) {
          data.products.push(normalizeProduct({ partNumber: item.name, description: item.name, unit: "pcs", stock: item.closingBalance }));
          created += 1;
        } else if (Number(existing.stock) !== item.closingBalance) {
          existing.stock = item.closingBalance;
          existing.sample = false;
          updated += 1;
        }
      }
      const already = new Set(data.tallySyncedVoucherGuids);
      let markedSynced = 0;
      for (const guid of allHeaderGuids) {
        if (!already.has(guid)) {
          data.tallySyncedVoucherGuids.push(guid);
          markedSynced += 1;
        }
      }
      const logEntry = {
        id: newId("tsync"), at: nowIso(), by: member.name,
        newVouchersFound: 0, purchasesApplied: 0, salesApplied: 0, unmatchedCount: 0,
        note: `Catalog baseline sync: ${created} new products created, ${updated} stock numbers corrected, ${markedSynced} existing vouchers marked as already reflected in today's balances.`,
      };
      data.tallySyncLog.unshift(logEntry);
      if (data.tallySyncLog.length > 200) data.tallySyncLog.length = 200;
      return { created, updated, markedSynced, totalTallyItems: stockItems.length };
    });
    res.json({ ok: true, ...result });
  } catch (e) {
    res.status(500).json({ ok: false, message: e.message });
  }
});

// Diagnostic: tries fetching item-level detail for the same first group of
// not-yet-synced vouchers using several different batch sizes (2, 5, 10, 20),
// and reports how many of each batch actually came back with items filled
// in. This tells us the largest batch size Tally can reliably hand back
// item detail for, so the real sync can use that number instead of guessing.
app.get("/api/tally-batch-size-test", async (req, res) => {
  try {
    const data = readData();
    const already = new Set(data.tallySyncedVoucherGuids);
    const headersBody = await queryTally(buildVoucherHeadersXml({}));
    const headers = parseVoucherHeaders(headersBody);
    const newHeaders = headers.filter((h) => !already.has(h.guid));
    const testSizes = [1, 2, 5, 10, 20];
    const results = {};
    for (const size of testSizes) {
      const subset = newHeaders.slice(0, size);
      if (subset.length === 0) continue;
      const body = await queryTally(buildVoucherItemsBatchXml({ guids: subset.map((h) => h.guid) }));
      const itemsMap = parseVoucherItemsOnly(body);
      const withItems = subset.filter((h) => (itemsMap[h.guid] || []).length > 0).length;
      results[size] = { requested: subset.length, gotItemsFor: withItems };
    }
    res.json({ ok: true, totalNewVouchers: newHeaders.length, results });
  } catch (e) {
    res.status(500).json({ ok: false, message: e.message });
  }
});

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

// A step further than /api/tally-test: actually sends an HTTP request (Tally's
// XML/ODBC gateway speaks HTTP under the hood) and shows what comes back, so
// we can tell whether this is genuinely talking to Tally versus something
// else answering on that port. Visit directly in a browser to run it.
app.get("/api/tally-http-test", (req, res) => {
  const host = String(req.query.host || "v60020.22164.tallyprimecloud.in");
  const port = Number(req.query.port) || 9537;
  const http = require("http");
  const start = Date.now();
  const request = http.get({ host, port, path: "/", timeout: 8000 }, (resp) => {
    let body = "";
    resp.on("data", (chunk) => { if (body.length < 2000) body += chunk; });
    resp.on("end", () => {
      res.json({ host, port, ms: Date.now() - start, ok: true, statusCode: resp.statusCode, headers: resp.headers, bodyPreview: body.slice(0, 2000) });
    });
  });
  request.on("timeout", () => {
    request.destroy();
    res.json({ host, port, ms: Date.now() - start, ok: false, message: "Connected at the TCP level but no HTTP response came back in time." });
  });
  request.on("error", (err) => {
    res.json({ host, port, ms: Date.now() - start, ok: false, message: `Request error: ${err.message}` });
  });
});

// Asks Tally for the actual list of companies it has loaded, so we can
// confirm "BMA" is spelled exactly the way Tally has it before the real
// sync is built around that name. Visit directly in a browser to run it.
app.get("/api/tally-companies-test", (req, res) => {
  const host = String(req.query.host || "v60020.22164.tallyprimecloud.in");
  const port = Number(req.query.port) || 9537;
  const http = require("http");
  // "List of Companies" isn't a built-in report name in this Tally version —
  // this instead defines a small inline TDL collection asking Tally to list
  // every object of type "Company" it has open, which is the reliable way
  // to do this across Tally versions.
  const xmlRequest = [
    "<ENVELOPE>",
    "<HEADER>",
    "<VERSION>1</VERSION>",
    "<TALLYREQUEST>Export</TALLYREQUEST>",
    "<TYPE>Collection</TYPE>",
    "<ID>ListOfCompanies</ID>",
    "</HEADER>",
    "<BODY>",
    "<DESC>",
    "<STATICVARIABLES>",
    "<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>",
    "</STATICVARIABLES>",
    "<TDL>",
    "<TDLMESSAGE>",
    '<COLLECTION NAME="ListOfCompanies" ISINITIALIZE="Yes">',
    "<TYPE>Company</TYPE>",
    "<FETCH>NAME</FETCH>",
    "</COLLECTION>",
    "</TDLMESSAGE>",
    "</TDL>",
    "</DESC>",
    "</BODY>",
    "</ENVELOPE>",
  ].join("");
  const start = Date.now();
  const request = http.request(
    { host, port, path: "/", method: "POST", headers: { "Content-Type": "text/xml", "Content-Length": Buffer.byteLength(xmlRequest) }, timeout: 10000 },
    (resp) => {
      let body = "";
      resp.on("data", (chunk) => { if (body.length < 8000) body += chunk; });
      resp.on("end", () => {
        res.json({ host, port, ms: Date.now() - start, ok: true, statusCode: resp.statusCode, bodyPreview: body.slice(0, 8000) });
      });
    }
  );
  request.on("timeout", () => {
    request.destroy();
    res.json({ host, port, ms: Date.now() - start, ok: false, message: "Timed out waiting for a response." });
  });
  request.on("error", (err) => {
    res.json({ host, port, ms: Date.now() - start, ok: false, message: `Request error: ${err.message}` });
  });
  request.write(xmlRequest);
  request.end();
});

// Pulls a small real slice of data (stock item names) from a specific
// company, to confirm we can query actual business data — not just company
// names — before the real sync gets built. Visit directly in a browser;
// pass ?company=... to test a different one, ?limit=... to see more rows.
app.get("/api/tally-stock-test", (req, res) => {
  const host = String(req.query.host || "v60020.22164.tallyprimecloud.in");
  const port = Number(req.query.port) || 9537;
  const company = String(req.query.company || "BMA - (from 1-Apr-26)");
  const limit = Number(req.query.limit) || 25;
  const http = require("http");
  const xmlRequest = [
    "<ENVELOPE>",
    "<HEADER>",
    "<VERSION>1</VERSION>",
    "<TALLYREQUEST>Export</TALLYREQUEST>",
    "<TYPE>Collection</TYPE>",
    "<ID>StockItemsTest</ID>",
    "</HEADER>",
    "<BODY>",
    "<DESC>",
    "<STATICVARIABLES>",
    "<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>",
    `<SVCURRENTCOMPANY>${company.replace(/&/g, "&amp;")}</SVCURRENTCOMPANY>`,
    "</STATICVARIABLES>",
    "<TDL>",
    "<TDLMESSAGE>",
    '<COLLECTION NAME="StockItemsTest" ISINITIALIZE="Yes">',
    "<TYPE>StockItem</TYPE>",
    "<FETCH>NAME</FETCH>",
    "<FETCH>CLOSINGBALANCE</FETCH>",
    "</COLLECTION>",
    "</TDLMESSAGE>",
    "</TDL>",
    "</DESC>",
    "</BODY>",
    "</ENVELOPE>",
  ].join("");
  const start = Date.now();
  const request = http.request(
    { host, port, path: "/", method: "POST", headers: { "Content-Type": "text/xml", "Content-Length": Buffer.byteLength(xmlRequest) }, timeout: 15000 },
    (resp) => {
      let body = "";
      resp.on("data", (chunk) => { body += chunk; });
      resp.on("end", () => {
        const matches = body.match(/<STOCKITEM[^>]*>[\s\S]*?<\/STOCKITEM>/g) || [];
        res.json({
          host, port, company, ms: Date.now() - start, ok: true, statusCode: resp.statusCode,
          totalItemsFound: matches.length,
          sample: matches.slice(0, limit).map((m) => m.replace(/\s+/g, " ").trim()),
        });
      });
    }
  );
  request.on("timeout", () => {
    request.destroy();
    res.json({ host, port, company, ms: Date.now() - start, ok: false, message: "Timed out waiting for a response." });
  });
  request.on("error", (err) => {
    res.json({ host, port, company, ms: Date.now() - start, ok: false, message: `Request error: ${err.message}` });
  });
  request.write(xmlRequest);
  request.end();
});

// Pulls basic Sales + Purchase voucher info (date, type, number, party —
// not line items yet) from a company/date range, to confirm we can read
// actual transactions before building line-item detail into the real sync.
// Visit directly in a browser; ?company=, ?from=YYYYMMDD, ?to=YYYYMMDD.
app.get("/api/tally-vouchers-test", (req, res) => {
  const host = String(req.query.host || "v60020.22164.tallyprimecloud.in");
  const port = Number(req.query.port) || 9537;
  const company = String(req.query.company || "BMA - (from 1-Apr-26)");
  const from = String(req.query.from || "20260401");
  const to = String(req.query.to || new Date().toISOString().slice(0, 10).replace(/-/g, ""));
  const http = require("http");
  const xmlRequest = [
    "<ENVELOPE>",
    "<HEADER>",
    "<VERSION>1</VERSION>",
    "<TALLYREQUEST>Export</TALLYREQUEST>",
    "<TYPE>Collection</TYPE>",
    "<ID>VouchersTest</ID>",
    "</HEADER>",
    "<BODY>",
    "<DESC>",
    "<STATICVARIABLES>",
    "<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>",
    `<SVCURRENTCOMPANY>${company.replace(/&/g, "&amp;")}</SVCURRENTCOMPANY>`,
    `<SVFROMDATE>${from}</SVFROMDATE>`,
    `<SVTODATE>${to}</SVTODATE>`,
    "</STATICVARIABLES>",
    "<TDL>",
    "<TDLMESSAGE>",
    '<COLLECTION NAME="VouchersTest" ISINITIALIZE="Yes">',
    "<TYPE>Voucher</TYPE>",
    "<FILTER>OnlySalesOrPurchase</FILTER>",
    "<FETCH>DATE</FETCH>",
    "<FETCH>VOUCHERTYPENAME</FETCH>",
    "<FETCH>VOUCHERNUMBER</FETCH>",
    "<FETCH>PARTYLEDGERNAME</FETCH>",
    "<FETCH>GUID</FETCH>",
    "</COLLECTION>",
    '<SYSTEM TYPE="Formulae" NAME="OnlySalesOrPurchase">$VoucherTypeName = "Sales" OR $VoucherTypeName = "Purchase"</SYSTEM>',
    "</TDLMESSAGE>",
    "</TDL>",
    "</DESC>",
    "</BODY>",
    "</ENVELOPE>",
  ].join("");
  const start = Date.now();
  const request = http.request(
    { host, port, path: "/", method: "POST", headers: { "Content-Type": "text/xml", "Content-Length": Buffer.byteLength(xmlRequest) }, timeout: 20000 },
    (resp) => {
      let body = "";
      resp.on("data", (chunk) => { body += chunk; });
      resp.on("end", () => {
        const matches = body.match(/<VOUCHER[^>]*>[\s\S]*?<\/VOUCHER>/g) || [];
        const lineError = body.match(/<LINEERROR>([\s\S]*?)<\/LINEERROR>/);
        res.json({
          host, port, company, from, to, ms: Date.now() - start, ok: true, statusCode: resp.statusCode,
          lineError: lineError ? lineError[1] : null,
          totalVouchersFound: matches.length,
          sample: matches.slice(0, 15).map((m) => m.replace(/\s+/g, " ").trim()),
          rawPreviewIfNoVouchers: matches.length === 0 ? body.slice(0, 3000) : undefined,
        });
      });
    }
  );
  request.on("timeout", () => {
    request.destroy();
    res.json({ host, port, company, ms: Date.now() - start, ok: false, message: "Timed out waiting for a response." });
  });
  request.on("error", (err) => {
    res.json({ host, port, company, ms: Date.now() - start, ok: false, message: `Request error: ${err.message}` });
  });
  request.write(xmlRequest);
  request.end();
});

// Pulls a couple of full vouchers WITH their line items (part number, qty)
// as raw, unstripped XML — so we can see exactly how Tally structures the
// item list before writing the real parsing logic. ?voucherNumber= narrows
// to one specific voucher if needed; otherwise shows the first few found.
app.get("/api/tally-voucher-items-test", (req, res) => {
  const host = String(req.query.host || "v60020.22164.tallyprimecloud.in");
  const port = Number(req.query.port) || 9537;
  const company = String(req.query.company || "BMA - (from 1-Apr-26)");
  const from = String(req.query.from || "20260401");
  const to = String(req.query.to || new Date().toISOString().slice(0, 10).replace(/-/g, ""));
  const limit = Number(req.query.limit) || 2;
  const http = require("http");
  const xmlRequest = [
    "<ENVELOPE>",
    "<HEADER>",
    "<VERSION>1</VERSION>",
    "<TALLYREQUEST>Export</TALLYREQUEST>",
    "<TYPE>Collection</TYPE>",
    "<ID>VoucherItemsTest</ID>",
    "</HEADER>",
    "<BODY>",
    "<DESC>",
    "<STATICVARIABLES>",
    "<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>",
    `<SVCURRENTCOMPANY>${company.replace(/&/g, "&amp;")}</SVCURRENTCOMPANY>`,
    `<SVFROMDATE>${from}</SVFROMDATE>`,
    `<SVTODATE>${to}</SVTODATE>`,
    "</STATICVARIABLES>",
    "<TDL>",
    "<TDLMESSAGE>",
    '<COLLECTION NAME="VoucherItemsTest" ISINITIALIZE="Yes">',
    "<TYPE>Voucher</TYPE>",
    "<FILTER>OnlySalesOrPurchase</FILTER>",
    "<FETCH>DATE</FETCH>",
    "<FETCH>VOUCHERTYPENAME</FETCH>",
    "<FETCH>VOUCHERNUMBER</FETCH>",
    "<FETCH>PARTYLEDGERNAME</FETCH>",
    "<FETCH>GUID</FETCH>",
    "<FETCH>ALLINVENTORYENTRIES.LIST</FETCH>",
    "</COLLECTION>",
    '<SYSTEM TYPE="Formulae" NAME="OnlySalesOrPurchase">$VoucherTypeName = "Sales" OR $VoucherTypeName = "Purchase"</SYSTEM>',
    "</TDLMESSAGE>",
    "</TDL>",
    "</DESC>",
    "</BODY>",
    "</ENVELOPE>",
  ].join("");
  const start = Date.now();
  const request = http.request(
    { host, port, path: "/", method: "POST", headers: { "Content-Type": "text/xml", "Content-Length": Buffer.byteLength(xmlRequest) }, timeout: 20000 },
    (resp) => {
      let body = "";
      resp.on("data", (chunk) => { body += chunk; });
      resp.on("end", () => {
        const matches = body.match(/<VOUCHER[^>]*>[\s\S]*?<\/VOUCHER>/g) || [];
        res.json({
          host, port, company, ms: Date.now() - start, ok: true, statusCode: resp.statusCode,
          totalVouchersFound: matches.length,
          fullVouchersRaw: matches.slice(0, limit),
        });
      });
    }
  );
  request.on("timeout", () => {
    request.destroy();
    res.json({ host, port, company, ms: Date.now() - start, ok: false, message: "Timed out waiting for a response." });
  });
  request.on("error", (err) => {
    res.json({ host, port, company, ms: Date.now() - start, ok: false, message: `Request error: ${err.message}` });
  });
  request.write(xmlRequest);
  request.end();
});

// Shows exactly what the real sync WOULD do, without changing any RaceLine
// data — every Purchase/Sales voucher not yet synced, matched against the
// product catalog, with the stock change each line would cause and a list
// of any stock item names that don't match a RaceLine part number. Safe to
// run as often as you like; visit directly in a browser. ?from=YYYYMMDD to
// override the default (company start, 1-Apr-2026).
app.get("/api/tally-sync-preview", async (req, res) => {
  try {
    const data = readData();
    const vouchers = await fetchNewVouchersWithItems(data.tallySyncedVoucherGuids);
    const neededRetryCount = vouchers.filter((v) => v.itemsNeededRetry).length;
    const plan = computeSyncPlan(data, vouchers, { apply: false });
    res.json({ ok: true, totalHeadersFetched: vouchers.length, neededRetryCount, ...plan });
  } catch (e) {
    res.status(500).json({ ok: false, message: e.message });
  }
});

// Actually applies the sync: increases stock for new Purchase vouchers,
// decreases it for new Sales vouchers, and remembers which voucher GUIDs
// have been applied so re-running this never double-counts. Owner only,
// since it changes real stock. Run /api/tally-sync-preview first to check
// the plan looks right.
app.post("/api/tally-sync-run", async (req, res) => {
  const token = req.header("X-Raceline-Token");
  const data0 = readData();
  const member = verifyToken(token, data0);
  if (!member || member.role !== "owner") return res.status(403).json({ error: "Owner only — log in and try again from the Dashboard." });
  try {
    const data0b = readData();
    const vouchers = await fetchNewVouchersWithItems(data0b.tallySyncedVoucherGuids);
    const result = await withData((data) => {
      const plan = computeSyncPlan(data, vouchers, { apply: true });
      const logEntry = {
        id: newId("tsync"), at: nowIso(), by: member.name,
        newVouchersFound: plan.newVouchersFound, purchasesApplied: plan.purchases.length,
        salesApplied: plan.sales.length, unmatchedCount: plan.unmatched.length,
      };
      data.tallySyncLog.unshift(logEntry);
      if (data.tallySyncLog.length > 200) data.tallySyncLog.length = 200;
      return plan;
    });
    res.json({ ok: true, ...result });
  } catch (e) {
    res.status(500).json({ ok: false, message: e.message });
  }
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
