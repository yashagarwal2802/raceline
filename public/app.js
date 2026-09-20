// RaceLine — Order & Stock Desk. Vanilla JS, no build step, no framework.
(() => {
  "use strict";

  // ---------- tiny helpers ----------
  const $ = (sel, root = document) => root.querySelector(sel);
  const el = (html) => {
    const t = document.createElement("template");
    t.innerHTML = html.trim();
    return t.content.firstElementChild;
  };
  const esc = (s) =>
    String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

  function relTime(iso) {
    if (!iso) return "";
    const d = new Date(iso);
    const diffMs = Date.now() - d.getTime();
    const min = Math.floor(diffMs / 60000);
    if (min < 1) return "just now";
    if (min < 60) return `${min}m ago`;
    const hr = Math.floor(min / 60);
    if (hr < 24) return `${hr}h ago`;
    const days = Math.floor(hr / 24);
    if (days < 7) return `${days}d ago`;
    return d.toLocaleDateString("en-IN", { day: "numeric", month: "short" });
  }
  function localDateStr(iso) {
    return new Date(iso).toLocaleDateString("en-CA");
  }
  const inrFormatter = new Intl.NumberFormat("en-IN", { style: "currency", currency: "INR", maximumFractionDigits: 0 });
  function inr(n) {
    return inrFormatter.format(Number(n) || 0);
  }

  function toast(msg, isError) {
    const t = el(`<div class="toast${isError ? " error" : ""}">${esc(msg)}</div>`);
    document.body.appendChild(t);
    setTimeout(() => t.remove(), 3400);
  }

  async function api(path, opts) {
    const res = await fetch(path, {
      method: opts?.method || "GET",
      headers: opts?.body ? { "Content-Type": "application/json" } : undefined,
      body: opts?.body ? JSON.stringify(opts.body) : undefined,
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || "Something went wrong.");
    return data;
  }

  // For uploading a real Excel file — no Content-Type header, the browser
  // sets the multipart boundary itself.
  async function apiUpload(path, formData) {
    const res = await fetch(path, { method: "POST", body: formData });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || "Something went wrong.");
    return data;
  }

  // ---------- state ----------
  let state = { team: [], products: [], customers: [], orders: [], skfRequests: [], stockIns: [], skfWriteOffs: [], otherRequests: [], otherStockIns: [], otherWriteOffs: [] };
  let identityId = localStorage.getItem("raceline_identity") || null;
  let activeTab = null;
  let cart = [];
  let selectedProduct = null;
  let ordersFilter = null; // set on first entry to Orders, defaults per role
  let ordersSearch = "";
  let stockSearch = "";
  let editingProductId = null;
  let editingCustomerId = null;
  let showAddProduct = false;
  let showBulkImportProducts = false;
  let showAddCustomer = false;
  let showBulkImportCustomers = false;
  let showAddTeam = false;
  let showSkfRequestForm = false;
  let showStockInForm = false;
  let lastStockInResult = null;
  let showOtherRequestForm = false;
  let showOtherStockInForm = false;
  let lastOtherStockInResult = null;

  const ROLE_LABEL = { owner: "Owner", "order-taker": "Order taker", biller: "Biller", dispatch: "Dispatch" };
  const TABS = {
    owner: [
      ["dashboard", "Dashboard"],
      ["new-order", "New order"],
      ["orders", "Orders"],
      ["stock", "Stock"],
      ["purchases", "Purchases"],
      ["other-purchases", "Other purchases"],
      ["customers", "Customers"],
      ["team", "Team"],
    ],
    "order-taker": [
      ["new-order", "New order"],
      ["orders", "Orders"],
      ["stock", "Stock"],
    ],
    biller: [
      ["orders", "Orders"],
      ["stock", "Stock"],
    ],
    dispatch: [
      ["orders", "Orders"],
      ["stock", "Stock"],
      ["purchases", "Purchases"],
      ["other-purchases", "Other purchases"],
    ],
  };
  const DEFAULT_TAB = { owner: "dashboard", "order-taker": "new-order", biller: "orders", dispatch: "orders" };
  const DEFAULT_ORDERS_FILTER = { owner: "all", "order-taker": "all", biller: "taken", dispatch: "billed" };

  function me() {
    return state.team.find((t) => t.id === identityId) || null;
  }

  // ---------- data loading ----------
  async function loadState({ silent } = {}) {
    try {
      const fresh = await api("/api/state");
      state = fresh;
      if (!silent) render();
    } catch (e) {
      if (!silent) toast("Could not reach the server. Is RaceLine still running?", true);
    }
  }

  function isEditingSomething() {
    const active = document.activeElement;
    const main = $("main") || $("#app");
    return active && ["INPUT", "TEXTAREA", "SELECT"].includes(active.tagName) && main && main.contains(active);
  }

  // ---------- render root ----------
  function render() {
    const app = $("#app");
    const person = me();
    if (!person) {
      app.innerHTML = "";
      app.appendChild(renderPicker());
      return;
    }
    if (!activeTab || !TABS[person.role].some(([id]) => id === activeTab)) {
      activeTab = DEFAULT_TAB[person.role];
    }
    if (ordersFilter === null) ordersFilter = DEFAULT_ORDERS_FILTER[person.role];

    app.innerHTML = "";
    app.appendChild(renderShell(person));
  }

  // ---------- identity picker ----------
  function renderPicker() {
    const wrap = el(`<div class="picker-wrap"></div>`);
    const card = el(`
      <div class="picker-card">
        <div class="brand"><span class="mark">RaceLine</span></div>
        <div class="picker-sub">Order &amp; Stock Desk — who's this?</div>
        <div class="member-grid"></div>
      </div>
    `);
    const grid = $(".member-grid", card);
    if (state.team.length === 0) {
      grid.appendChild(el(`<div class="empty-note">No team members set up yet.</div>`));
    }
    for (const t of state.team) {
      const btn = el(`
        <button class="member-btn" data-id="${t.id}">
          <span class="name">${esc(t.name)}</span>
          <span class="role-chip">${ROLE_LABEL[t.role] || t.role}</span>
        </button>
      `);
      btn.addEventListener("click", () => {
        identityId = t.id;
        localStorage.setItem("raceline_identity", identityId);
        activeTab = null;
        ordersFilter = null;
        render();
      });
      grid.appendChild(btn);
    }
    card.appendChild(el(`<div class="empty-note" style="padding-top:16px;">Don't see your name? Ask the Owner to add you under the Team tab.</div>`));
    wrap.appendChild(card);
    return wrap;
  }

  // ---------- app shell ----------
  function renderShell(person) {
    const root = el(`<div></div>`);
    const topbar = el(`
      <div class="topbar">
        <div class="brand"><span class="mark">RaceLine</span><span class="tag">Order &amp; Stock Desk</span></div>
        <div class="topbar-spacer"></div>
        <div class="who">
          <span class="name">${esc(person.name)}</span>
          <span class="role-chip">${ROLE_LABEL[person.role]}</span>
          <button class="link-btn" id="switch-btn">Switch</button>
        </div>
      </div>
    `);
    $("#switch-btn", topbar).addEventListener("click", () => {
      identityId = null;
      localStorage.removeItem("raceline_identity");
      render();
    });

    const tabs = el(`<div class="tabs"></div>`);
    for (const [id, label] of TABS[person.role]) {
      const b = el(`<button class="tab${id === activeTab ? " active" : ""}">${label}</button>`);
      b.addEventListener("click", () => {
        activeTab = id;
        render();
      });
      tabs.appendChild(b);
    }

    const main = el(`<main></main>`);
    const view = {
      dashboard: renderDashboard,
      "new-order": renderNewOrder,
      orders: renderOrdersView,
      stock: renderStockView,
      purchases: renderPurchasesView,
      "other-purchases": renderOtherPurchasesView,
      customers: renderCustomersView,
      team: renderTeamView,
    }[activeTab];
    main.appendChild(view(person));

    root.appendChild(topbar);
    root.appendChild(tabs);
    root.appendChild(main);
    return root;
  }

  // ---------- Dashboard ----------
  function renderDashboard(person) {
    const wrap = el(`<div></div>`);
    wrap.appendChild(el(`<div class="view-header"><h1>Dashboard</h1><div class="sub">${new Date().toLocaleDateString("en-IN", { weekday: "long", day: "numeric", month: "long" })}</div></div>`));

    const today = new Date().toLocaleDateString("en-CA");
    const takenToday = state.orders.filter((o) => localDateStr(o.takenAt) === today).length;
    const pendingBill = state.orders.filter((o) => o.status === "taken").length;
    const pendingDispatch = state.orders.filter((o) => o.status === "billed").length;
    const lowStock = state.products.filter((p) => p.stock <= p.reorderLevel);

    const stats = el(`<div class="grid-stats"></div>`);
    stats.appendChild(el(`<div class="stat-tile"><div class="num">${takenToday}</div><div class="label">Orders today</div></div>`));
    stats.appendChild(el(`<div class="stat-tile${pendingBill ? " warn" : ""}"><div class="num">${pendingBill}</div><div class="label">Pending billing</div></div>`));
    stats.appendChild(el(`<div class="stat-tile${pendingDispatch ? " warn" : ""}"><div class="num">${pendingDispatch}</div><div class="label">Pending dispatch</div></div>`));
    stats.appendChild(el(`<div class="stat-tile${lowStock.length ? " crit" : ""}"><div class="num">${lowStock.length}</div><div class="label">Low stock items</div></div>`));
    wrap.appendChild(stats);

    // Low stock panel
    const lsPanel = el(`<div class="panel"><h2>Reorder soon</h2></div>`);
    if (lowStock.length === 0) {
      lsPanel.appendChild(el(`<div class="empty">Nothing at or below reorder level.</div>`));
    } else {
      const tw = el(`<div class="table-wrap"><table class="data"><thead><tr><th>Part number</th><th>Description</th><th>Stock</th><th>Reorder at</th></tr></thead><tbody></tbody></table></div>`);
      const tbody = $("tbody", tw);
      lowStock
        .sort((a, b) => a.stock - a.reorderLevel - (b.stock - b.reorderLevel))
        .forEach((p) => {
          tbody.appendChild(
            el(`<tr>
              <td class="mono">${esc(p.partNumber)}</td>
              <td>${esc(p.description)}</td>
              <td class="mono">${stockDot(p)}${p.stock} ${esc(p.unit)}</td>
              <td class="mono">${p.reorderLevel}</td>
            </tr>`)
          );
        });
      lsPanel.appendChild(tw);
    }
    wrap.appendChild(lsPanel);

    // Slow movers: dispatched in last 90 days
    const cutoff = Date.now() - 90 * 24 * 3600 * 1000;
    const movedRecently = new Set();
    state.orders
      .filter((o) => o.status === "dispatched" && new Date(o.dispatchedAt).getTime() >= cutoff)
      .forEach((o) => o.items.forEach((it) => movedRecently.add(it.partNumber)));
    const slow = state.products.filter((p) => !movedRecently.has(p.partNumber) && !p.sample).sort((a, b) => b.stock - a.stock).slice(0, 10);

    const smPanel = el(`<div class="panel"><h2>Slow movers <span style="font-weight:500;color:var(--ink-muted);font-size:12.5px;">— no dispatch in 90 days</span></h2></div>`);
    const hasHistory = state.orders.some((o) => o.status === "dispatched");
    if (!hasHistory) {
      smPanel.appendChild(el(`<div class="empty">No dispatch history yet — this fills in as you use RaceLine.</div>`));
    } else if (slow.length === 0) {
      smPanel.appendChild(el(`<div class="empty">Everything has moved in the last 90 days.</div>`));
    } else {
      const tw = el(`<div class="table-wrap"><table class="data"><thead><tr><th>Part number</th><th>Description</th><th>Stock on hand</th></tr></thead><tbody></tbody></table></div>`);
      const tbody = $("tbody", tw);
      slow.forEach((p) => {
        tbody.appendChild(el(`<tr><td class="mono">${esc(p.partNumber)}</td><td>${esc(p.description)}</td><td class="mono">${p.stock} ${esc(p.unit)}</td></tr>`));
      });
      smPanel.appendChild(tw);
    }
    wrap.appendChild(smPanel);

    const backupRow = el(`<div style="margin-top:14px;display:flex;gap:10px;align-items:center;flex-wrap:wrap;"></div>`);
    backupRow.appendChild(el(`<a class="btn small" href="/api/export" download="raceline-backup.json">Export backup (JSON)</a>`));
    if (person && person.role === "owner") {
      const restoreInput = el(`<input type="file" accept=".json" style="display:none;" />`);
      const restoreBtn = el(`<button class="btn small ghost">Restore from backup…</button>`);
      restoreBtn.addEventListener("click", () => restoreInput.click());
      restoreInput.addEventListener("change", async () => {
        const file = restoreInput.files[0];
        if (!file) return;
        if (!confirm("This replaces ALL current data (stock, customers, orders, everything) with this backup file's contents. This can't be undone. Continue?")) {
          restoreInput.value = "";
          return;
        }
        try {
          const text = await file.text();
          const parsed = JSON.parse(text);
          await api("/api/restore", { method: "POST", body: parsed });
          toast("Restored from backup.");
          await loadState();
        } catch (e) {
          toast(e.message || "Could not restore that file.", true);
        } finally {
          restoreInput.value = "";
        }
      });
      backupRow.appendChild(restoreBtn);
      backupRow.appendChild(restoreInput);
    }
    wrap.appendChild(backupRow);
    return wrap;
  }

  function stockDot(p) {
    const cls = p.stock <= 0 ? "out" : p.stock <= p.reorderLevel ? "low" : "ok";
    return `<span class="stock-dot ${cls}"></span>`;
  }

  // ---------- New order ----------
  function renderNewOrder(person) {
    const wrap = el(`<div></div>`);
    wrap.appendChild(el(`<div class="view-header"><h1>New order</h1><div class="sub">Taken by ${esc(person.name)}</div></div>`));

    // customer
    const custPanel = el(`<div class="panel"></div>`);
    custPanel.appendChild(el(`<div class="field"><label>Customer</label></div>`));
    const sel = el(`<select id="no-customer"><option value="">Select customer…</option></select>`);
    state.customers
      .slice()
      .sort((a, b) => a.name.localeCompare(b.name))
      .forEach((c) => sel.appendChild(el(`<option value="${c.id}">${esc(c.name)}${c.area ? " — " + esc(c.area) : ""}</option>`)));
    sel.appendChild(el(`<option value="__new__">+ New customer…</option>`));
    custPanel.appendChild(sel);
    const newCustFields = el(`
      <div class="form-grid" style="margin-top:10px;" hidden>
        <div class="field"><label>Name</label><input type="text" id="no-new-name" /></div>
        <div class="field"><label>Contact</label><input type="tel" id="no-new-contact" /></div>
        <div class="field"><label>Area</label><input type="text" id="no-new-area" /></div>
      </div>
    `);
    sel.addEventListener("change", () => {
      newCustFields.hidden = sel.value !== "__new__";
    });
    custPanel.appendChild(newCustFields);
    wrap.appendChild(custPanel);

    // item search — a live dropdown under the input, filtered as you type,
    // each row showing current stock so the order-taker sees availability
    // without leaving the field.
    const itemPanel = el(`<div class="panel"></div>`);
    itemPanel.appendChild(el(`<div class="field"><label>Add item</label></div>`));
    const comboWrap = el(`<div class="combo"></div>`);
    const search = el(`<input type="search" id="no-item-search" placeholder="Type a part number or description…" autocomplete="off" />`);
    const results = el(`<div class="combo-results" hidden></div>`);
    comboWrap.appendChild(search);
    comboWrap.appendChild(results);
    itemPanel.appendChild(comboWrap);
    const addRow = el(`
      <div class="item-picker-row" style="margin-top:10px;" hidden>
        <div><strong class="mono" id="no-selected-pn"></strong><div id="no-selected-desc" style="font-size:12.5px;color:var(--ink-muted)"></div></div>
        <div class="qty field"><label>Qty</label><input type="number" id="no-qty" min="1" value="1" /></div>
        <button class="btn primary small" id="no-add-btn">Add</button>
      </div>
    `);
    itemPanel.appendChild(addRow);

    function pickProduct(p) {
      selectedProduct = p;
      $("#no-selected-pn", addRow).textContent = p.partNumber;
      $("#no-selected-desc", addRow).textContent = `${p.description} — ${p.stock} ${p.unit} in stock`;
      addRow.hidden = false;
      results.hidden = true;
      search.value = p.partNumber;
      $("#no-qty", addRow).value = 1;
      $("#no-qty", addRow).focus();
    }
    function renderResults() {
      const q = search.value.trim().toLowerCase();
      if (!q) {
        results.hidden = true;
        results.innerHTML = "";
        return;
      }
      const matches = state.products
        .filter((p) => p.partNumber.toLowerCase().includes(q) || p.description.toLowerCase().includes(q))
        .slice(0, 8);
      results.innerHTML = "";
      results.hidden = false;
      if (matches.length === 0) {
        results.appendChild(el(`<div class="combo-empty">No matching parts.</div>`));
        return;
      }
      matches.forEach((p) => {
        const row = el(`
          <button type="button" class="combo-row">
            <span><span class="mono">${esc(p.partNumber)}</span> — ${esc(p.description)}</span>
            <span class="mono" style="color:var(--ink-muted);white-space:nowrap;">${stockDot(p)}${p.stock} ${esc(p.unit)}</span>
          </button>
        `);
        row.addEventListener("mousedown", (e) => e.preventDefault()); // keep focus so blur doesn't beat the click
        row.addEventListener("click", () => pickProduct(p));
        results.appendChild(row);
      });
    }
    search.addEventListener("input", renderResults);
    search.addEventListener("focus", renderResults);
    search.addEventListener("blur", () => setTimeout(() => (results.hidden = true), 150));

    $("#no-add-btn", addRow).addEventListener("click", () => {
      if (!selectedProduct) return;
      const qty = Math.max(1, Number($("#no-qty", addRow).value) || 1);
      const existing = cart.find((c) => c.partNumber === selectedProduct.partNumber);
      if (existing) existing.qty += qty;
      else cart.push({ partNumber: selectedProduct.partNumber, description: selectedProduct.description, unit: selectedProduct.unit, stock: selectedProduct.stock, qty });
      selectedProduct = null;
      search.value = "";
      results.innerHTML = "";
      addRow.hidden = true;
      renderCart();
    });
    wrap.appendChild(itemPanel);

    // cart
    const cartPanel = el(`<div class="panel"><h2>Order items</h2><div id="no-cart"></div></div>`);
    wrap.appendChild(cartPanel);
    function renderCart() {
      const holder = $("#no-cart", cartPanel);
      if (cart.length === 0) {
        holder.innerHTML = `<div class="empty">No items added yet.</div>`;
        return;
      }
      const table = el(`<table class="cart-table"><thead><tr><th>Part number</th><th>Description</th><th>Qty</th><th></th></tr></thead><tbody></tbody></table>`);
      const tbody = $("tbody", table);
      cart.forEach((c, i) => {
        const tr = el(`<tr>
          <td class="mono">${esc(c.partNumber)}</td>
          <td>${esc(c.description)}${c.qty > c.stock ? `<div style="color:var(--warning);font-size:12px;">Only ${c.stock} ${esc(c.unit)} on record</div>` : ""}</td>
          <td><input type="number" min="1" value="${c.qty}" style="width:70px;" data-idx="${i}" class="cart-qty" /></td>
          <td><button class="btn small ghost" data-idx="${i}" class="cart-remove">Remove</button></td>
        </tr>`);
        tbody.appendChild(tr);
      });
      holder.innerHTML = "";
      holder.appendChild(table);
      table.querySelectorAll(".cart-qty").forEach((inp) =>
        inp.addEventListener("change", (e) => {
          const idx = Number(e.target.dataset.idx);
          cart[idx].qty = Math.max(1, Number(e.target.value) || 1);
          renderCart();
        })
      );
      table.querySelectorAll("button[data-idx]").forEach((btn) =>
        btn.addEventListener("click", (e) => {
          const idx = Number(e.target.dataset.idx);
          cart.splice(idx, 1);
          renderCart();
        })
      );
    }
    renderCart();

    // notes + submit
    const submitPanel = el(`<div class="panel"></div>`);
    submitPanel.appendChild(el(`<div class="field"><label>Notes (optional)</label><textarea id="no-notes" placeholder="Delivery instructions, special notes…"></textarea></div>`));
    const submitBtn = el(`<button class="btn primary" style="margin-top:12px;">Place order</button>`);
    submitPanel.appendChild(submitBtn);
    wrap.appendChild(submitPanel);

    submitBtn.addEventListener("click", async () => {
      if (cart.length === 0) return toast("Add at least one item first.", true);
      const customerId = sel.value;
      if (!customerId) return toast("Select a customer.", true);
      let newCustomer = null;
      if (customerId === "__new__") {
        const name = $("#no-new-name", newCustFields).value.trim();
        if (!name) return toast("Enter the new customer's name.", true);
        newCustomer = { name, contact: $("#no-new-contact", newCustFields).value.trim(), area: $("#no-new-area", newCustFields).value.trim() };
      }
      submitBtn.disabled = true;
      try {
        const res = await api("/api/orders", {
          method: "POST",
          body: {
            customerId,
            newCustomer,
            items: cart.map((c) => ({ partNumber: c.partNumber, qty: c.qty })),
            takenBy: person.name,
            notes: $("#no-notes", submitPanel).value,
          },
        });
        cart = [];
        toast(res.warnings && res.warnings.length ? `Order placed — ${res.warnings.join(" ")}` : "Order placed.");
        activeTab = "orders";
        ordersFilter = "taken";
        await loadState();
      } catch (e) {
        toast(e.message, true);
      } finally {
        submitBtn.disabled = false;
      }
    });

    return wrap;
  }

  // ---------- Orders ----------
  const FILTER_LABEL = { all: "All", taken: "Taken", billed: "Billed", dispatched: "Dispatched", cancelled: "Cancelled" };

  function renderOrdersView(person) {
    const wrap = el(`<div></div>`);
    wrap.appendChild(el(`<div class="view-header"><h1>Orders</h1></div>`));

    const filterRow = el(`<div class="filter-row"></div>`);
    Object.keys(FILTER_LABEL).forEach((f) => {
      const chip = el(`<button class="chip${ordersFilter === f ? " active" : ""}">${FILTER_LABEL[f]}</button>`);
      chip.addEventListener("click", () => {
        ordersFilter = f;
        renderList();
        updateChips();
      });
      filterRow.appendChild(chip);
    });
    function updateChips() {
      [...filterRow.children].forEach((c, i) => c.classList.toggle("active", Object.keys(FILTER_LABEL)[i] === ordersFilter));
    }
    wrap.appendChild(filterRow);

    const search = el(`<input type="search" placeholder="Search customer, order id, or part number…" value="${esc(ordersSearch)}" style="margin-bottom:14px;" />`);
    search.addEventListener("input", () => {
      ordersSearch = search.value;
      renderList();
    });
    wrap.appendChild(search);

    const list = el(`<div class="stack"></div>`);
    wrap.appendChild(list);

    function renderList() {
      const q = ordersSearch.trim().toLowerCase();
      let orders = state.orders.filter((o) => ordersFilter === "all" || o.status === ordersFilter);
      if (q) {
        orders = orders.filter(
          (o) =>
            o.customerName.toLowerCase().includes(q) ||
            o.id.toLowerCase().includes(q) ||
            o.items.some((it) => it.partNumber.toLowerCase().includes(q))
        );
      }
      list.innerHTML = "";
      if (orders.length === 0) {
        list.appendChild(el(`<div class="empty">No orders here.</div>`));
        return;
      }
      orders.forEach((o) => list.appendChild(renderOrderCard(o, person)));
    }
    renderList();

    return wrap;
  }

  function renderOrderCard(o, person) {
    const card = el(`
      <div class="order-card">
        <div class="order-top">
          <div>
            <div class="order-cust">${esc(o.customerName)}</div>
            <div class="order-meta">#${o.id.slice(-6)} · taken by ${esc(o.takenBy)} · ${relTime(o.takenAt)}</div>
          </div>
          <span class="pill ${o.status}">${o.status}</span>
        </div>
        <div class="order-items"></div>
      </div>
    `);
    const itemsHolder = $(".order-items", card);
    o.items.forEach((it) => {
      itemsHolder.appendChild(el(`<div class="row"><span class="pn">${esc(it.partNumber)} <span style="color:var(--ink-faint)">${esc(it.description)}</span></span><span>${it.qty} ${esc(it.unit)}</span></div>`));
    });
    if (o.notes) card.appendChild(el(`<div class="order-notes">${esc(o.notes)}</div>`));
    if (o.tallyInvoiceNo) card.appendChild(el(`<div class="order-meta">Tally invoice: <span class="mono">${esc(o.tallyInvoiceNo)}</span></div>`));
    if (o.status === "billed" || o.status === "dispatched") card.appendChild(el(`<div class="order-meta">Billed by ${esc(o.billedBy || "")} · ${relTime(o.billedAt)}</div>`));
    if (o.status === "dispatched") card.appendChild(el(`<div class="order-meta">Dispatched by ${esc(o.dispatchedBy || "")} · ${relTime(o.dispatchedAt)}</div>`));
    if (o.status === "cancelled") card.appendChild(el(`<div class="order-meta">Cancelled by ${esc(o.cancelledBy || "")}</div>`));

    const actions = el(`<div class="order-actions"></div>`);
    const canBill = (person.role === "owner" || person.role === "biller") && o.status === "taken";
    const canDispatch = (person.role === "owner" || person.role === "dispatch") && o.status === "billed";
    const canCancel = person.role === "owner" && (o.status === "taken" || o.status === "billed");

    if (canBill) {
      const b = el(`<button class="btn primary small">Mark billed</button>`);
      b.addEventListener("click", async () => {
        const inv = prompt("Tally invoice number (optional):", "");
        if (inv === null) return;
        try {
          await api(`/api/orders/${o.id}/bill`, { method: "PUT", body: { billedBy: person.name, tallyInvoiceNo: inv } });
          toast("Marked as billed.");
          await loadState();
        } catch (e) {
          toast(e.message, true);
        }
      });
      actions.appendChild(b);
    }
    if (canDispatch) {
      const b = el(`<button class="btn primary small">Mark dispatched</button>`);
      b.addEventListener("click", async () => {
        if (!confirm("Dispatch this order? Stock on hand will be reduced.")) return;
        try {
          await api(`/api/orders/${o.id}/dispatch`, { method: "PUT", body: { dispatchedBy: person.name } });
          toast("Marked as dispatched.");
          await loadState();
        } catch (e) {
          toast(e.message, true);
        }
      });
      actions.appendChild(b);
    }
    if (canCancel) {
      const b = el(`<button class="btn small danger">Cancel</button>`);
      b.addEventListener("click", async () => {
        if (!confirm("Cancel this order?")) return;
        try {
          await api(`/api/orders/${o.id}/cancel`, { method: "PUT", body: { cancelledBy: person.name } });
          toast("Order cancelled.");
          await loadState();
        } catch (e) {
          toast(e.message, true);
        }
      });
      actions.appendChild(b);
    }
    if (actions.children.length) card.appendChild(actions);
    return card;
  }

  // ---------- Stock ----------
  function renderStockView(person) {
    const wrap = el(`<div></div>`);
    const isOwner = person.role === "owner";
    const canEditRow = person.role === "owner" || person.role === "dispatch";
    wrap.appendChild(el(`<div class="view-header"><h1>Stock</h1><div class="sub">${state.products.length} SKF part numbers</div></div>`));

    const sampleCount = state.products.filter((p) => p.sample).length;
    if (isOwner && sampleCount > 0) {
      const banner = el(`<div class="banner"><span>${sampleCount} sample part${sampleCount === 1 ? "" : "s"} loaded so the app isn't empty. Replace with your real catalog, then clear the samples.</span></div>`);
      const clearBtn = el(`<button class="btn small">Clear samples</button>`);
      clearBtn.addEventListener("click", async () => {
        const r = await api("/api/products/clear-samples", { method: "POST" });
        toast(`Removed ${r.removed} sample item(s).`);
        await loadState();
      });
      banner.appendChild(clearBtn);
      wrap.appendChild(banner);
    }

    if (isOwner) {
      const actionsRow = el(`<div class="section-actions"></div>`);
      const addBtn = el(`<button class="btn">+ Add item</button>`);
      const importBtn = el(`<button class="btn">Bulk import</button>`);
      addBtn.addEventListener("click", () => {
        showAddProduct = !showAddProduct;
        showBulkImportProducts = false;
        rerenderStockBody();
      });
      importBtn.addEventListener("click", () => {
        showBulkImportProducts = !showBulkImportProducts;
        showAddProduct = false;
        rerenderStockBody();
      });
      actionsRow.appendChild(addBtn);
      actionsRow.appendChild(importBtn);
      wrap.appendChild(actionsRow);
    }

    const bodyHolder = el(`<div></div>`);
    wrap.appendChild(bodyHolder);

    function rerenderStockBody() {
      bodyHolder.innerHTML = "";
      if (isOwner && showAddProduct) bodyHolder.appendChild(renderAddProductForm());
      if (isOwner && showBulkImportProducts) bodyHolder.appendChild(renderBulkImportProducts());

      const search = el(`<input type="search" placeholder="Search part number or description…" value="${esc(stockSearch)}" style="margin:12px 0;" />`);
      search.addEventListener("input", () => {
        stockSearch = search.value;
        renderTable();
      });
      bodyHolder.appendChild(search);

      const tableHolder = el(`<div></div>`);
      bodyHolder.appendChild(tableHolder);

      function renderTable() {
        const q = stockSearch.trim().toLowerCase();
        const items = state.products
          .filter((p) => !q || p.partNumber.toLowerCase().includes(q) || p.description.toLowerCase().includes(q))
          .sort((a, b) => a.partNumber.localeCompare(b.partNumber));
        tableHolder.innerHTML = "";
        if (items.length === 0) {
          tableHolder.appendChild(el(`<div class="empty">No parts match.</div>`));
          return;
        }
        const tw = el(`<div class="table-wrap"><table class="data"><thead><tr>
          <th>Part number</th><th>Description</th><th>Category</th><th>Location</th><th>Stock</th><th>Reorder at</th>${canEditRow ? "<th></th>" : ""}
        </tr></thead><tbody></tbody></table></div>`);
        const tbody = $("tbody", tw);
        items.forEach((p) => tbody.appendChild(renderProductRow(p, canEditRow, renderTable)));
        tableHolder.appendChild(tw);
      }
      renderTable();
    }
    rerenderStockBody();

    return wrap;
  }

  function renderProductRow(p, canEdit, onSaved) {
    if (editingProductId === p.id) {
      const tr = el(`<tr>
        <td class="mono">${esc(p.partNumber)}</td>
        <td><input type="text" value="${esc(p.description)}" id="ep-desc" /></td>
        <td><input type="text" value="${esc(p.category)}" id="ep-cat" style="width:120px;" /></td>
        <td><input type="text" value="${esc(p.location)}" id="ep-loc" style="width:100px;" placeholder="e.g. Rack A1" /></td>
        <td><input type="number" value="${p.stock}" id="ep-stock" style="width:80px;" /></td>
        <td><input type="number" value="${p.reorderLevel}" id="ep-reorder" style="width:80px;" /></td>
        <td style="display:flex;gap:6px;"><button class="btn small primary">Save</button><button class="btn small ghost">Cancel</button></td>
      </tr>`);
      $("button.primary", tr).addEventListener("click", async () => {
        try {
          await api(`/api/products/${encodeURIComponent(p.id)}`, {
            method: "PUT",
            body: {
              description: $("#ep-desc", tr).value,
              category: $("#ep-cat", tr).value,
              location: $("#ep-loc", tr).value,
              stock: $("#ep-stock", tr).value,
              reorderLevel: $("#ep-reorder", tr).value,
            },
          });
          editingProductId = null;
          await loadState({ silent: true });
          toast("Stock updated.");
          onSaved();
        } catch (e) {
          toast(e.message, true);
        }
      });
      $("button.ghost", tr).addEventListener("click", () => {
        editingProductId = null;
        onSaved();
      });
      return tr;
    }
    const tr = el(`<tr>
      <td class="mono">${esc(p.partNumber)}${p.sample ? '<span class="sample-tag">Sample</span>' : ""}</td>
      <td>${esc(p.description)}</td>
      <td>${esc(p.category)}</td>
      <td class="mono">${p.location ? esc(p.location) : '<span style="color:var(--ink-faint);">—</span>'}</td>
      <td class="mono">${stockDot(p)}${p.stock} ${esc(p.unit)}</td>
      <td class="mono">${p.reorderLevel}</td>
      ${canEdit ? `<td><button class="btn small">Edit</button></td>` : ""}
    </tr>`);
    if (canEdit) {
      $("button", tr).addEventListener("click", () => {
        editingProductId = p.id;
        onSaved();
      });
    }
    return tr;
  }

  function renderAddProductForm() {
    const panel = el(`
      <div class="panel">
        <h2>Add a part</h2>
        <div class="form-grid">
          <div class="field"><label>Part number</label><input type="text" id="ap-pn" placeholder="e.g. 6205-2RS1" /></div>
          <div class="field"><label>Description</label><input type="text" id="ap-desc" /></div>
          <div class="field"><label>Category</label><input type="text" id="ap-cat" placeholder="Deep groove ball…" /></div>
          <div class="field"><label>Unit</label><input type="text" id="ap-unit" value="pcs" /></div>
          <div class="field"><label>Stock on hand</label><input type="number" id="ap-stock" value="0" /></div>
          <div class="field"><label>Reorder level</label><input type="number" id="ap-reorder" value="0" /></div>
          <div class="field"><label>Location</label><input type="text" id="ap-loc" placeholder="e.g. Rack A1" /></div>
        </div>
        <button class="btn primary" style="margin-top:12px;">Save part</button>
      </div>
    `);
    $("button", panel).addEventListener("click", async () => {
      const partNumber = $("#ap-pn", panel).value.trim();
      if (!partNumber) return toast("Enter a part number.", true);
      try {
        await api("/api/products", {
          method: "POST",
          body: {
            partNumber,
            description: $("#ap-desc", panel).value,
            category: $("#ap-cat", panel).value,
            unit: $("#ap-unit", panel).value,
            stock: $("#ap-stock", panel).value,
            reorderLevel: $("#ap-reorder", panel).value,
            location: $("#ap-loc", panel).value,
          },
        });
        showAddProduct = false;
        toast("Part added.");
        await loadState();
      } catch (e) {
        toast(e.message, true);
      }
    });
    return panel;
  }

  function renderBulkImportProducts() {
    const panel = el(`
      <div class="panel">
        <h2>Bulk import parts</h2>
        <div class="sub" style="margin-bottom:8px;">One part per line: <code class="mono">part number, description, category, unit, stock, reorder level, location</code></div>
        <textarea id="bi-text" rows="6" placeholder="6205-2RS1, Deep groove ball bearing, Deep groove ball, pcs, 84, 20, Rack A1"></textarea>
        <button class="btn primary" id="bi-paste-btn" style="margin-top:10px;">Import</button>
        <div class="sub" style="margin:16px 0 10px;border-top:1px solid var(--border);padding-top:12px;">— or upload an Excel file (.xlsx/.xls), e.g. exported from Tally —</div>
        <input type="file" id="bi-file" accept=".xlsx,.xls" />
        <button class="btn" id="bi-file-btn" style="margin-top:10px;display:block;">Upload file</button>
      </div>
    `);
    $("#bi-paste-btn", panel).addEventListener("click", async () => {
      const lines = $("#bi-text", panel).value.split("\n").map((l) => l.trim()).filter(Boolean);
      const rows = lines.map((l) => {
        const [partNumber, description, category, unit, stock, reorderLevel, location] = l.split(",").map((s) => (s || "").trim());
        return { partNumber, description, category, unit, stock, reorderLevel, location };
      });
      if (rows.length === 0) return toast("Paste at least one row.", true);
      try {
        const res = await api("/api/products", { method: "POST", body: rows });
        showBulkImportProducts = false;
        toast(`Imported ${res.saved.length} part(s).${res.errors.length ? " " + res.errors.length + " skipped." : ""}`);
        await loadState();
      } catch (e) {
        toast(e.message, true);
      }
    });
    $("#bi-file-btn", panel).addEventListener("click", async () => {
      const file = $("#bi-file", panel).files[0];
      if (!file) return toast("Choose a file first.", true);
      const fd = new FormData();
      fd.append("file", file);
      const btn = $("#bi-file-btn", panel);
      btn.disabled = true;
      try {
        const res = await apiUpload("/api/products/import-excel", fd);
        showBulkImportProducts = false;
        toast(`Imported ${res.saved.length} part(s) from ${res.rowsRead} row(s).${res.errors.length ? " " + res.errors.length + " skipped." : ""}`);
        await loadState();
      } catch (e) {
        toast(e.message, true);
      } finally {
        btn.disabled = false;
      }
    });
    return panel;
  }

  // ---------- Customers ----------
  function renderCustomersView() {
    const wrap = el(`<div></div>`);
    wrap.appendChild(el(`<div class="view-header"><h1>Customers</h1><div class="sub">${state.customers.length} on your line</div></div>`));

    const sampleCount = state.customers.filter((c) => c.sample).length;
    if (sampleCount > 0) {
      const banner = el(`<div class="banner"><span>${sampleCount} sample customer${sampleCount === 1 ? "" : "s"} loaded as a placeholder. Add your real dealer list, then clear the samples.</span></div>`);
      const clearBtn = el(`<button class="btn small">Clear samples</button>`);
      clearBtn.addEventListener("click", async () => {
        const r = await api("/api/customers/clear-samples", { method: "POST" });
        toast(`Removed ${r.removed} sample customer(s).`);
        await loadState();
      });
      banner.appendChild(clearBtn);
      wrap.appendChild(banner);
    }

    const actionsRow = el(`<div class="section-actions"></div>`);
    const addBtn = el(`<button class="btn">+ Add customer</button>`);
    const importBtn = el(`<button class="btn">Bulk import</button>`);
    addBtn.addEventListener("click", () => {
      showAddCustomer = !showAddCustomer;
      showBulkImportCustomers = false;
      rerender();
    });
    importBtn.addEventListener("click", () => {
      showBulkImportCustomers = !showBulkImportCustomers;
      showAddCustomer = false;
      rerender();
    });
    actionsRow.appendChild(addBtn);
    actionsRow.appendChild(importBtn);
    wrap.appendChild(actionsRow);

    const bodyHolder = el(`<div></div>`);
    wrap.appendChild(bodyHolder);

    function rerender() {
      bodyHolder.innerHTML = "";
      if (showAddCustomer) bodyHolder.appendChild(renderAddCustomerForm(rerender));
      if (showBulkImportCustomers) bodyHolder.appendChild(renderBulkImportCustomers());

      const totalOutstanding = state.customers.reduce((sum, c) => sum + (c.outstandingAmount || 0), 0);
      if (totalOutstanding > 0) {
        const stats = el(`<div class="grid-stats" style="margin-bottom:14px;"></div>`);
        stats.appendChild(el(`<div class="stat-tile warn"><div class="num" style="font-size:20px;">${inr(totalOutstanding)}</div><div class="label">Total outstanding</div></div>`));
        bodyHolder.appendChild(stats);
      }

      const tw = el(`<div class="table-wrap" style="margin-top:12px;"><table class="data"><thead><tr><th>Name</th><th>Contact</th><th>Area</th><th>Credit period</th><th>Outstanding</th><th></th></tr></thead><tbody></tbody></table></div>`);
      const tbody = $("tbody", tw);
      if (state.customers.length === 0) {
        bodyHolder.appendChild(el(`<div class="empty">No customers yet.</div>`));
      } else {
        state.customers
          .slice()
          .sort((a, b) => (b.outstandingAmount || 0) - (a.outstandingAmount || 0) || a.name.localeCompare(b.name))
          .forEach((c) => tbody.appendChild(renderCustomerRow(c, rerender)));
        bodyHolder.appendChild(tw);
      }
    }
    rerender();

    return wrap;
  }

  function renderCustomerRow(c, onSaved) {
    if (editingCustomerId === c.id) {
      const tr = el(`<tr>
        <td><input type="text" value="${esc(c.name)}" id="ec-name" /></td>
        <td><input type="tel" value="${esc(c.contact)}" id="ec-contact" /></td>
        <td><input type="text" value="${esc(c.area)}" id="ec-area" /></td>
        <td><input type="number" value="${c.creditPeriodDays || 0}" id="ec-credit" style="width:70px;" /> days</td>
        <td><input type="number" value="${c.outstandingAmount || 0}" id="ec-outstanding" style="width:100px;" /></td>
        <td style="display:flex;gap:6px;"><button class="btn small primary">Save</button><button class="btn small ghost">Cancel</button></td>
      </tr>`);
      $("button.primary", tr).addEventListener("click", async () => {
        await api(`/api/customers/${c.id}`, {
          method: "PUT",
          body: {
            name: $("#ec-name", tr).value,
            contact: $("#ec-contact", tr).value,
            area: $("#ec-area", tr).value,
            creditPeriodDays: $("#ec-credit", tr).value,
            outstandingAmount: $("#ec-outstanding", tr).value,
          },
        });
        editingCustomerId = null;
        await loadState({ silent: true });
        onSaved();
      });
      $("button.ghost", tr).addEventListener("click", () => {
        editingCustomerId = null;
        onSaved();
      });
      return tr;
    }
    const tr = el(`<tr>
      <td>${esc(c.name)}${c.sample ? '<span class="sample-tag">Sample</span>' : ""}</td>
      <td>${esc(c.contact)}</td>
      <td>${esc(c.area)}</td>
      <td class="mono">${c.creditPeriodDays ? c.creditPeriodDays + "d" : "—"}</td>
      <td class="mono">${c.outstandingAmount ? inr(c.outstandingAmount) : "—"}${c.outstandingAsOf ? `<div style="font-size:11px;color:var(--ink-faint);">as of ${relTime(c.outstandingAsOf)}</div>` : ""}</td>
      <td><button class="btn small">Edit</button></td>
    </tr>`);
    $("button", tr).addEventListener("click", () => {
      editingCustomerId = c.id;
      onSaved();
    });
    return tr;
  }

  function renderAddCustomerForm(onDone) {
    const panel = el(`
      <div class="panel">
        <h2>Add a customer</h2>
        <div class="form-grid">
          <div class="field"><label>Name</label><input type="text" id="ac-name" /></div>
          <div class="field"><label>Contact</label><input type="tel" id="ac-contact" /></div>
          <div class="field"><label>Area</label><input type="text" id="ac-area" /></div>
          <div class="field"><label>Credit period (days)</label><input type="number" id="ac-credit" value="0" /></div>
          <div class="field"><label>Outstanding amount (₹)</label><input type="number" id="ac-outstanding" value="0" /></div>
        </div>
        <button class="btn primary" style="margin-top:12px;">Save customer</button>
      </div>
    `);
    $("button", panel).addEventListener("click", async () => {
      const name = $("#ac-name", panel).value.trim();
      if (!name) return toast("Enter a name.", true);
      await api("/api/customers", {
        method: "POST",
        body: {
          name,
          contact: $("#ac-contact", panel).value,
          area: $("#ac-area", panel).value,
          creditPeriodDays: $("#ac-credit", panel).value,
          outstandingAmount: $("#ac-outstanding", panel).value,
        },
      });
      showAddCustomer = false;
      toast("Customer added.");
      await loadState();
    });
    return panel;
  }

  function renderBulkImportCustomers() {
    const panel = el(`
      <div class="panel">
        <h2>Bulk import customers</h2>
        <div class="sub" style="margin-bottom:8px;">One customer per line: <code class="mono">name, contact, area</code></div>
        <textarea id="bic-text" rows="6" placeholder="Shree Balaji Bearing House, 98200 00001, Local"></textarea>
        <button class="btn primary" id="bic-paste-btn" style="margin-top:10px;">Import</button>
        <div class="sub" style="margin:16px 0 10px;border-top:1px solid var(--border);padding-top:12px;">— or upload an Excel file (.xlsx/.xls), e.g. exported from Tally —</div>
        <input type="file" id="bic-file" accept=".xlsx,.xls" />
        <button class="btn" id="bic-file-btn" style="margin-top:10px;display:block;">Upload file</button>
      </div>
    `);
    $("#bic-paste-btn", panel).addEventListener("click", async () => {
      const lines = $("#bic-text", panel).value.split("\n").map((l) => l.trim()).filter(Boolean);
      const rows = lines.map((l) => {
        const [name, contact, area] = l.split(",").map((s) => (s || "").trim());
        return { name, contact, area };
      });
      if (rows.length === 0) return toast("Paste at least one row.", true);
      const res = await api("/api/customers", { method: "POST", body: rows });
      showBulkImportCustomers = false;
      toast(`Imported ${res.saved.length} customer(s).`);
      await loadState();
    });
    $("#bic-file-btn", panel).addEventListener("click", async () => {
      const file = $("#bic-file", panel).files[0];
      if (!file) return toast("Choose a file first.", true);
      const fd = new FormData();
      fd.append("file", file);
      const btn = $("#bic-file-btn", panel);
      btn.disabled = true;
      try {
        const res = await apiUpload("/api/customers/import-excel", fd);
        showBulkImportCustomers = false;
        toast(`Imported ${res.saved.length} customer(s) from ${res.rowsRead} row(s).`);
        await loadState();
      } catch (e) {
        toast(e.message, true);
      } finally {
        btn.disabled = false;
      }
    });
    return panel;
  }

  // ---------- reusable: pick part numbers + quantities ----------
  function createItemBuilder() {
    let items = [];
    let selected = null;
    const wrap = el(`<div></div>`);
    const comboWrap = el(`<div class="combo"></div>`);
    const search = el(`<input type="search" placeholder="Type a part number or description…" autocomplete="off" />`);
    const results = el(`<div class="combo-results" hidden></div>`);
    comboWrap.appendChild(search);
    comboWrap.appendChild(results);
    const addRow = el(`
      <div class="item-picker-row" style="margin-top:10px;" hidden>
        <div><strong class="mono" id="ib-pn"></strong><div id="ib-desc" style="font-size:12.5px;color:var(--ink-muted)"></div></div>
        <div class="qty field"><label>Qty</label><input type="number" min="1" value="1" /></div>
        <button class="btn primary small">Add</button>
      </div>
    `);
    const listHolder = el(`<div style="margin-top:10px;"></div>`);
    wrap.appendChild(comboWrap);
    wrap.appendChild(addRow);
    wrap.appendChild(listHolder);

    function renderResults() {
      const q = search.value.trim().toLowerCase();
      if (!q) {
        results.hidden = true;
        results.innerHTML = "";
        return;
      }
      const matches = state.products.filter((p) => p.partNumber.toLowerCase().includes(q) || p.description.toLowerCase().includes(q)).slice(0, 8);
      results.innerHTML = "";
      results.hidden = false;
      if (matches.length === 0) {
        results.appendChild(el(`<div class="combo-empty">No matching parts.</div>`));
        return;
      }
      matches.forEach((p) => {
        const row = el(`
          <button type="button" class="combo-row">
            <span><span class="mono">${esc(p.partNumber)}</span> — ${esc(p.description)}</span>
            <span class="mono" style="color:var(--ink-muted);white-space:nowrap;">${stockDot(p)}${p.stock} ${esc(p.unit)}</span>
          </button>
        `);
        row.addEventListener("mousedown", (e) => e.preventDefault());
        row.addEventListener("click", () => {
          selected = p;
          $("#ib-pn", addRow).textContent = p.partNumber;
          $("#ib-desc", addRow).textContent = p.description;
          addRow.hidden = false;
          results.hidden = true;
          search.value = p.partNumber;
          $("input[type=number]", addRow).value = 1;
          $("input[type=number]", addRow).focus();
        });
        results.appendChild(row);
      });
    }
    search.addEventListener("input", renderResults);
    search.addEventListener("focus", renderResults);
    search.addEventListener("blur", () => setTimeout(() => (results.hidden = true), 150));

    $("button", addRow).addEventListener("click", () => {
      if (!selected) return;
      const qty = Math.max(1, Number($("input[type=number]", addRow).value) || 1);
      const existing = items.find((i) => i.partNumber === selected.partNumber);
      if (existing) existing.qty += qty;
      else items.push({ partNumber: selected.partNumber, description: selected.description, unit: selected.unit, qty });
      selected = null;
      search.value = "";
      results.innerHTML = "";
      addRow.hidden = true;
      renderList();
    });

    function renderList() {
      listHolder.innerHTML = "";
      if (items.length === 0) {
        listHolder.appendChild(el(`<div class="empty">No parts added yet.</div>`));
        return;
      }
      const table = el(`<table class="cart-table"><thead><tr><th>Part number</th><th>Description</th><th>Qty</th><th></th></tr></thead><tbody></tbody></table>`);
      const tbody = $("tbody", table);
      items.forEach((c, i) => {
        const tr = el(`<tr>
          <td class="mono">${esc(c.partNumber)}</td>
          <td>${esc(c.description)}</td>
          <td><input type="number" min="1" value="${c.qty}" style="width:70px;" /></td>
          <td><button class="btn small ghost">Remove</button></td>
        </tr>`);
        $("input", tr).addEventListener("change", (e) => {
          items[i].qty = Math.max(1, Number(e.target.value) || 1);
        });
        $("button", tr).addEventListener("click", () => {
          items.splice(i, 1);
          renderList();
        });
        tbody.appendChild(tr);
      });
      listHolder.appendChild(table);
    }
    renderList();

    return { el: wrap, getItems: () => items, clear: () => { items = []; renderList(); } };
  }

  // ---------- SKF pending ledger (mirrors the server's FIFO logic) ----------
  function skfPendingSummary() {
    const byPart = {};
    function ensure(pn, desc) {
      if (!byPart[pn]) byPart[pn] = { partNumber: pn, description: desc || "", requests: [], totalReceived: 0, totalWrittenOff: 0 };
      return byPart[pn];
    }
    state.skfRequests.forEach((r) => r.items.forEach((it) => ensure(it.partNumber, it.description).requests.push({ date: r.date, qty: it.qty })));
    state.stockIns.forEach((s) => s.items.forEach((it) => (ensure(it.partNumber, it.description).totalReceived += it.qty)));
    state.skfWriteOffs.forEach((w) => (ensure(w.partNumber, w.description).totalWrittenOff += w.qty));

    const result = [];
    Object.values(byPart).forEach((p) => {
      const requests = p.requests.slice().sort((a, b) => new Date(a.date) - new Date(b.date));
      let toConsume = p.totalReceived + p.totalWrittenOff;
      let oldestSince = null;
      let totalRequested = 0;
      requests.forEach((r) => {
        totalRequested += r.qty;
        let remaining = r.qty;
        if (toConsume > 0) {
          const consumed = Math.min(toConsume, remaining);
          toConsume -= consumed;
          remaining -= consumed;
        }
        if (remaining > 0 && oldestSince === null) oldestSince = r.date;
      });
      const pending = Math.max(0, totalRequested - p.totalReceived - p.totalWrittenOff);
      result.push({ partNumber: p.partNumber, description: p.description, requested: totalRequested, received: p.totalReceived, writtenOff: p.totalWrittenOff, pending, oldestSince: pending > 0 ? oldestSince : null });
    });
    return result.sort((a, b) => b.pending - a.pending);
  }

  // ---------- Purchases (SKF) ----------
  function renderPurchasesView(person) {
    const wrap = el(`<div></div>`);
    wrap.appendChild(el(`<div class="view-header"><h1>Purchases</h1><div class="sub">Tracking orders placed with SKF</div></div>`));

    const canRequest = person.role === "owner";
    const canReceive = person.role === "owner" || person.role === "dispatch";

    const actionsRow = el(`<div class="section-actions"></div>`);
    const reqBtn = el(`<button class="btn">+ Log request to SKF</button>`);
    const recvBtn = el(`<button class="btn primary">Log stock received</button>`);
    if (canRequest) actionsRow.appendChild(reqBtn);
    if (canReceive) actionsRow.appendChild(recvBtn);
    wrap.appendChild(actionsRow);

    const bodyHolder = el(`<div></div>`);
    wrap.appendChild(bodyHolder);

    function rerenderBody() {
      bodyHolder.innerHTML = "";
      if (canRequest && showSkfRequestForm) bodyHolder.appendChild(renderSkfRequestForm(person, () => { showSkfRequestForm = false; rerenderBody(); }));
      if (canReceive && showStockInForm) bodyHolder.appendChild(renderStockInForm(person, () => { showStockInForm = false; rerenderBody(); }));
      if (lastStockInResult) bodyHolder.appendChild(renderLastStockInPanel());
      bodyHolder.appendChild(renderPendingPanel(person));
      bodyHolder.appendChild(renderSkfActivityPanel());
    }
    reqBtn.addEventListener("click", () => {
      showSkfRequestForm = !showSkfRequestForm;
      showStockInForm = false;
      rerenderBody();
    });
    recvBtn.addEventListener("click", () => {
      showStockInForm = !showStockInForm;
      showSkfRequestForm = false;
      rerenderBody();
    });
    rerenderBody();

    return wrap;
  }

  function renderSkfRequestForm(person, onDone) {
    const panel = el(`<div class="panel"><h2>Log a request to SKF</h2><div class="sub" style="margin-bottom:8px;">What you're emailing them, so RaceLine can track it.</div></div>`);

    const tabsRow = el(`<div class="filter-row"></div>`);
    const manualChip = el(`<button class="chip active">Pick parts</button>`);
    const excelChip = el(`<button class="chip">Upload Excel</button>`);
    tabsRow.appendChild(manualChip);
    tabsRow.appendChild(excelChip);
    panel.appendChild(tabsRow);

    const manualSection = el(`<div style="margin-top:10px;"></div>`);
    const builder = createItemBuilder();
    manualSection.appendChild(builder.el);
    manualSection.appendChild(el(`<div class="field" style="margin-top:10px;"><label>Note (optional)</label><textarea id="skf-req-note" placeholder="e.g. sent via email to SKF Pune depot"></textarea></div>`));
    const manualBtn = el(`<button class="btn primary" style="margin-top:12px;">Log request</button>`);
    manualSection.appendChild(manualBtn);

    const excelSection = el(`<div style="margin-top:10px;" hidden></div>`);
    excelSection.appendChild(el(`<div class="sub" style="margin-bottom:8px;">Upload the Excel sheet you're sending/sent SKF, listing part numbers and quantities.</div>`));
    const fileInput = el(`<input type="file" id="skf-req-file" accept=".xlsx,.xls" />`);
    excelSection.appendChild(fileInput);
    const excelBtn = el(`<button class="btn primary" style="margin-top:12px;display:block;">Upload &amp; log request</button>`);
    excelSection.appendChild(excelBtn);

    panel.appendChild(manualSection);
    panel.appendChild(excelSection);

    manualChip.addEventListener("click", () => {
      manualChip.classList.add("active");
      excelChip.classList.remove("active");
      manualSection.hidden = false;
      excelSection.hidden = true;
    });
    excelChip.addEventListener("click", () => {
      excelChip.classList.add("active");
      manualChip.classList.remove("active");
      excelSection.hidden = false;
      manualSection.hidden = true;
    });

    manualBtn.addEventListener("click", async () => {
      const items = builder.getItems();
      if (items.length === 0) return toast("Add at least one part.", true);
      try {
        await api("/api/skf-requests", { method: "POST", body: { items, loggedBy: person.name, note: $("#skf-req-note", panel).value } });
        toast("Request logged.");
        onDone();
        await loadState();
      } catch (e) {
        toast(e.message, true);
      }
    });

    excelBtn.addEventListener("click", async () => {
      const file = fileInput.files[0];
      if (!file) return toast("Choose an Excel file first.", true);
      const fd = new FormData();
      fd.append("file", file);
      fd.append("loggedBy", person.name);
      excelBtn.disabled = true;
      try {
        const res = await apiUpload("/api/skf-requests/import-excel", fd);
        toast(`Request logged from ${res.rowsRead} row(s).`);
        onDone();
        await loadState();
      } catch (e) {
        toast(e.message, true);
      } finally {
        excelBtn.disabled = false;
      }
    });

    return panel;
  }

  function renderStockInForm(person, onDone) {
    const panel = el(`<div class="panel"><h2>Log stock received</h2><div class="sub" style="margin-bottom:8px;">What actually arrived from SKF today.</div></div>`);
    const builder = createItemBuilder();
    panel.appendChild(builder.el);
    panel.appendChild(el(`<div class="field" style="margin-top:10px;"><label>Note (optional)</label><textarea id="stockin-note" placeholder="e.g. partial delivery, balance pending"></textarea></div>`));
    const btn = el(`<button class="btn primary" style="margin-top:12px;">Log receipt</button>`);
    panel.appendChild(btn);
    btn.addEventListener("click", async () => {
      const items = builder.getItems();
      if (items.length === 0) return toast("Add at least one part.", true);
      try {
        const res = await api("/api/stock-in", { method: "POST", body: { items, receivedBy: person.name, note: $("#stockin-note", panel).value } });
        lastStockInResult = res;
        toast("Stock received logged.");
        onDone();
        await loadState();
      } catch (e) {
        toast(e.message, true);
      }
    });
    return panel;
  }

  function renderLastStockInPanel() {
    const panel = el(`<div class="panel"></div>`);
    const header = el(`<div style="display:flex;justify-content:space-between;align-items:baseline;"><h2>Just received — who to fulfill first</h2></div>`);
    const closeBtn = el(`<button class="link-btn">Dismiss</button>`);
    closeBtn.addEventListener("click", () => {
      lastStockInResult = null;
      render();
    });
    header.appendChild(closeBtn);
    panel.appendChild(header);
    lastStockInResult.allocations.forEach((a) => {
      const block = el(`<div style="margin-top:10px;padding-top:10px;border-top:1px dashed var(--border);"></div>`);
      block.appendChild(el(`<div><span class="mono">${esc(a.partNumber)}</span> — received ${a.qtyReceived}, ${esc(a.description)}</div>`));
      if (a.covered.length === 0) {
        block.appendChild(el(`<div class="empty" style="padding:6px 0;">No pending customer orders for this part.</div>`));
      } else {
        a.covered.forEach((c) => {
          block.appendChild(el(`<div class="order-meta">→ ${esc(c.customerName)} (${c.status}, ordered ${relTime(c.takenAt)}) — covers ${c.covered} of ${c.qty}</div>`));
        });
      }
      if (a.leftover > 0) block.appendChild(el(`<div class="order-meta" style="color:var(--success);">${a.leftover} left over after covering pending orders.</div>`));
      panel.appendChild(block);
    });
    return panel;
  }

  function renderPendingPanel(person) {
    const panel = el(`<div class="panel"><h2>Pending from SKF</h2></div>`);
    const rows = skfPendingSummary().filter((r) => r.pending > 0 || r.requested > 0);
    if (rows.length === 0) {
      panel.appendChild(el(`<div class="empty">No requests logged yet.</div>`));
      return panel;
    }
    const tw = el(`<div class="table-wrap"><table class="data"><thead><tr>
      <th>Part number</th><th>Requested</th><th>Received</th><th>Written off</th><th>Pending</th><th>Outstanding since</th>${person.role === "owner" ? "<th></th>" : ""}
    </tr></thead><tbody></tbody></table></div>`);
    const tbody = $("tbody", tw);
    rows.forEach((r) => {
      const tr = el(`<tr>
        <td class="mono">${esc(r.partNumber)}${r.pending > 0 ? "" : ""}</td>
        <td class="mono">${r.requested}</td>
        <td class="mono">${r.received}</td>
        <td class="mono">${r.writtenOff}</td>
        <td class="mono">${r.pending > 0 ? `<span class="stock-dot low"></span>` : ""}${r.pending}</td>
        <td>${r.oldestSince ? relTime(r.oldestSince) : "—"}</td>
      </tr>`);
      if (person.role === "owner") {
        const td = el(`<td></td>`);
        if (r.pending > 0) {
          const btn = el(`<button class="btn small danger">Write off</button>`);
          btn.addEventListener("click", async () => {
            const qtyStr = prompt(`Write off how many units of ${r.partNumber}? (pending: ${r.pending})`, String(r.pending));
            if (qtyStr === null) return;
            const qty = Number(qtyStr);
            if (!qty || qty <= 0) return toast("Enter a valid quantity.", true);
            const note = prompt("Reason (optional):", "") || "";
            try {
              await api("/api/skf-writeoffs", { method: "POST", body: { partNumber: r.partNumber, qty, note, writtenOffBy: person.name } });
              toast("Written off.");
              await loadState();
            } catch (e) {
              toast(e.message, true);
            }
          });
          td.appendChild(btn);
        }
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    });
    panel.appendChild(tw);
    return panel;
  }

  function renderSkfActivityPanel() {
    const panel = el(`<div class="panel"><h2>Recent activity</h2></div>`);
    const entries = [
      ...state.skfRequests.map((r) => ({ type: "Requested", date: r.date, by: r.loggedBy, items: r.items })),
      ...state.stockIns.map((s) => ({ type: "Received", date: s.date, by: s.receivedBy, items: s.items })),
    ]
      .sort((a, b) => new Date(b.date) - new Date(a.date))
      .slice(0, 8);
    if (entries.length === 0) {
      panel.appendChild(el(`<div class="empty">Nothing logged yet.</div>`));
      return panel;
    }
    entries.forEach((e) => {
      const summary = e.items.map((it) => `${it.partNumber} ×${it.qty}`).join(", ");
      panel.appendChild(el(`<div class="order-meta" style="padding:5px 0;"><strong>${e.type}</strong> · ${esc(e.by)} · ${relTime(e.date)} — ${esc(summary)}</div>`));
    });
    return panel;
  }

  // ---------- Other-supplier pending ledger (same FIFO logic, keyed by party+part) ----------
  function otherPendingSummary() {
    const byKey = {};
    function ensure(party, pn, desc) {
      const key = party + "||" + pn;
      if (!byKey[key]) byKey[key] = { party, partNumber: pn, description: desc || "", requests: [], totalReceived: 0, totalWrittenOff: 0 };
      return byKey[key];
    }
    state.otherRequests.forEach((r) => r.items.forEach((it) => ensure(r.party, it.partNumber, it.description).requests.push({ date: r.date, qty: it.qty })));
    state.otherStockIns.forEach((s) => s.items.forEach((it) => (ensure(s.party, it.partNumber, it.description).totalReceived += it.qty)));
    state.otherWriteOffs.forEach((w) => (ensure(w.party, w.partNumber, w.description).totalWrittenOff += w.qty));

    const result = [];
    Object.values(byKey).forEach((p) => {
      const requests = p.requests.slice().sort((a, b) => new Date(a.date) - new Date(b.date));
      let toConsume = p.totalReceived + p.totalWrittenOff;
      let oldestSince = null;
      let totalRequested = 0;
      requests.forEach((r) => {
        totalRequested += r.qty;
        let remaining = r.qty;
        if (toConsume > 0) {
          const consumed = Math.min(toConsume, remaining);
          toConsume -= consumed;
          remaining -= consumed;
        }
        if (remaining > 0 && oldestSince === null) oldestSince = r.date;
      });
      const pending = Math.max(0, totalRequested - p.totalReceived - p.totalWrittenOff);
      result.push({ party: p.party, partNumber: p.partNumber, description: p.description, requested: totalRequested, received: p.totalReceived, writtenOff: p.totalWrittenOff, pending, oldestSince: pending > 0 ? oldestSince : null });
    });
    return result.sort((a, b) => b.pending - a.pending || a.party.localeCompare(b.party));
  }

  function getKnownParties() {
    const set = new Set();
    state.otherRequests.forEach((r) => set.add(r.party));
    state.otherStockIns.forEach((s) => set.add(s.party));
    return [...set].sort();
  }

  // ---------- Purchases (other suppliers) ----------
  function renderOtherPurchasesView(person) {
    const wrap = el(`<div></div>`);
    wrap.appendChild(el(`<div class="view-header"><h1>Other purchases</h1><div class="sub">Tracking orders placed with suppliers other than SKF — same catalog part numbers</div></div>`));

    const canRequest = person.role === "owner";
    const canReceive = person.role === "owner" || person.role === "dispatch";

    const actionsRow = el(`<div class="section-actions"></div>`);
    const reqBtn = el(`<button class="btn">+ Log request to a party</button>`);
    const recvBtn = el(`<button class="btn primary">Log stock received</button>`);
    if (canRequest) actionsRow.appendChild(reqBtn);
    if (canReceive) actionsRow.appendChild(recvBtn);
    wrap.appendChild(actionsRow);

    const bodyHolder = el(`<div></div>`);
    wrap.appendChild(bodyHolder);

    function rerenderBody() {
      bodyHolder.innerHTML = "";
      if (canRequest && showOtherRequestForm) bodyHolder.appendChild(renderOtherRequestForm(person, () => { showOtherRequestForm = false; rerenderBody(); }));
      if (canReceive && showOtherStockInForm) bodyHolder.appendChild(renderOtherStockInForm(person, () => { showOtherStockInForm = false; rerenderBody(); }));
      if (lastOtherStockInResult) bodyHolder.appendChild(renderLastOtherStockInPanel());
      bodyHolder.appendChild(renderOtherPendingPanel(person));
      bodyHolder.appendChild(renderOtherActivityPanel());
    }
    reqBtn.addEventListener("click", () => {
      showOtherRequestForm = !showOtherRequestForm;
      showOtherStockInForm = false;
      rerenderBody();
    });
    recvBtn.addEventListener("click", () => {
      showOtherStockInForm = !showOtherStockInForm;
      showOtherRequestForm = false;
      rerenderBody();
    });
    rerenderBody();

    return wrap;
  }

  function renderOtherRequestForm(person, onDone) {
    const panel = el(`<div class="panel"><h2>Log a request to a party</h2><div class="sub" style="margin-bottom:8px;">What you're ordering by WhatsApp/phone, so RaceLine can track it.</div></div>`);
    const knownParties = getKnownParties();
    panel.appendChild(el(`
      <div class="field">
        <label>Party / supplier name</label>
        <input type="text" id="oth-req-party" list="oth-party-list" placeholder="e.g. Bansal Bearings" />
        <datalist id="oth-party-list">${knownParties.map((p) => `<option value="${esc(p)}"></option>`).join("")}</datalist>
      </div>
    `));

    const tabsRow = el(`<div class="filter-row" style="margin-top:10px;"></div>`);
    const manualChip = el(`<button class="chip active">Pick parts</button>`);
    const excelChip = el(`<button class="chip">Upload Excel</button>`);
    tabsRow.appendChild(manualChip);
    tabsRow.appendChild(excelChip);
    panel.appendChild(tabsRow);

    const manualSection = el(`<div style="margin-top:10px;"></div>`);
    const builder = createItemBuilder();
    manualSection.appendChild(builder.el);
    manualSection.appendChild(el(`<div class="field" style="margin-top:10px;"><label>Note (optional)</label><textarea id="oth-req-note" placeholder="e.g. sent via WhatsApp"></textarea></div>`));
    const manualBtn = el(`<button class="btn primary" style="margin-top:12px;">Log request</button>`);
    manualSection.appendChild(manualBtn);

    const excelSection = el(`<div style="margin-top:10px;" hidden></div>`);
    excelSection.appendChild(el(`<div class="sub" style="margin-bottom:8px;">Upload the Excel sheet you're sending/sent this party, listing part numbers and quantities.</div>`));
    const fileInput = el(`<input type="file" id="oth-req-file" accept=".xlsx,.xls" />`);
    excelSection.appendChild(fileInput);
    const excelBtn = el(`<button class="btn primary" style="margin-top:12px;display:block;">Upload &amp; log request</button>`);
    excelSection.appendChild(excelBtn);

    panel.appendChild(manualSection);
    panel.appendChild(excelSection);

    manualChip.addEventListener("click", () => {
      manualChip.classList.add("active");
      excelChip.classList.remove("active");
      manualSection.hidden = false;
      excelSection.hidden = true;
    });
    excelChip.addEventListener("click", () => {
      excelChip.classList.add("active");
      manualChip.classList.remove("active");
      excelSection.hidden = false;
      manualSection.hidden = true;
    });

    manualBtn.addEventListener("click", async () => {
      const party = $("#oth-req-party", panel).value.trim();
      if (!party) return toast("Enter a party/supplier name.", true);
      const items = builder.getItems();
      if (items.length === 0) return toast("Add at least one part.", true);
      try {
        await api("/api/other-requests", { method: "POST", body: { party, items, loggedBy: person.name, note: $("#oth-req-note", panel).value } });
        toast("Request logged.");
        onDone();
        await loadState();
      } catch (e) {
        toast(e.message, true);
      }
    });

    excelBtn.addEventListener("click", async () => {
      const party = $("#oth-req-party", panel).value.trim();
      if (!party) return toast("Enter a party/supplier name.", true);
      const file = fileInput.files[0];
      if (!file) return toast("Choose an Excel file first.", true);
      const fd = new FormData();
      fd.append("file", file);
      fd.append("party", party);
      fd.append("loggedBy", person.name);
      excelBtn.disabled = true;
      try {
        const res = await apiUpload("/api/other-requests/import-excel", fd);
        toast(`Request logged from ${res.rowsRead} row(s).`);
        onDone();
        await loadState();
      } catch (e) {
        toast(e.message, true);
      } finally {
        excelBtn.disabled = false;
      }
    });

    return panel;
  }

  function renderOtherStockInForm(person, onDone) {
    const panel = el(`<div class="panel"><h2>Log stock received</h2><div class="sub" style="margin-bottom:8px;">What actually arrived from another party today.</div></div>`);
    const knownParties = getKnownParties();
    panel.appendChild(el(`
      <div class="field">
        <label>Party / supplier name</label>
        <input type="text" id="oth-in-party" list="oth-party-list-2" placeholder="e.g. Bansal Bearings" />
        <datalist id="oth-party-list-2">${knownParties.map((p) => `<option value="${esc(p)}"></option>`).join("")}</datalist>
      </div>
    `));
    const builder = createItemBuilder();
    panel.appendChild(builder.el);
    panel.appendChild(el(`<div class="field" style="margin-top:10px;"><label>Note (optional)</label><textarea id="oth-in-note" placeholder="e.g. partial delivery, balance pending"></textarea></div>`));
    const btn = el(`<button class="btn primary" style="margin-top:12px;">Log receipt</button>`);
    panel.appendChild(btn);
    btn.addEventListener("click", async () => {
      const party = $("#oth-in-party", panel).value.trim();
      if (!party) return toast("Enter a party/supplier name.", true);
      const items = builder.getItems();
      if (items.length === 0) return toast("Add at least one part.", true);
      try {
        const res = await api("/api/other-stock-in", { method: "POST", body: { party, items, receivedBy: person.name, note: $("#oth-in-note", panel).value } });
        lastOtherStockInResult = res;
        toast("Stock received logged.");
        onDone();
        await loadState();
      } catch (e) {
        toast(e.message, true);
      }
    });
    return panel;
  }

  function renderLastOtherStockInPanel() {
    const panel = el(`<div class="panel"></div>`);
    const header = el(`<div style="display:flex;justify-content:space-between;align-items:baseline;"><h2>Just received — who to fulfill first</h2></div>`);
    const closeBtn = el(`<button class="link-btn">Dismiss</button>`);
    closeBtn.addEventListener("click", () => {
      lastOtherStockInResult = null;
      render();
    });
    header.appendChild(closeBtn);
    panel.appendChild(header);
    lastOtherStockInResult.allocations.forEach((a) => {
      const block = el(`<div style="margin-top:10px;padding-top:10px;border-top:1px dashed var(--border);"></div>`);
      block.appendChild(el(`<div><span class="mono">${esc(a.partNumber)}</span> — received ${a.qtyReceived}, ${esc(a.description)}</div>`));
      if (a.covered.length === 0) {
        block.appendChild(el(`<div class="empty" style="padding:6px 0;">No pending customer orders for this part.</div>`));
      } else {
        a.covered.forEach((c) => {
          block.appendChild(el(`<div class="order-meta">→ ${esc(c.customerName)} (${c.status}, ordered ${relTime(c.takenAt)}) — covers ${c.covered} of ${c.qty}</div>`));
        });
      }
      if (a.leftover > 0) block.appendChild(el(`<div class="order-meta" style="color:var(--success);">${a.leftover} left over after covering pending orders.</div>`));
      panel.appendChild(block);
    });
    return panel;
  }

  function renderOtherPendingPanel(person) {
    const panel = el(`<div class="panel"><h2>Pending from other parties</h2></div>`);
    const rows = otherPendingSummary().filter((r) => r.pending > 0 || r.requested > 0);
    if (rows.length === 0) {
      panel.appendChild(el(`<div class="empty">No requests logged yet.</div>`));
      return panel;
    }
    const tw = el(`<div class="table-wrap"><table class="data"><thead><tr>
      <th>Party</th><th>Part number</th><th>Requested</th><th>Received</th><th>Written off</th><th>Pending</th><th>Outstanding since</th>${person.role === "owner" ? "<th></th>" : ""}
    </tr></thead><tbody></tbody></table></div>`);
    const tbody = $("tbody", tw);
    rows.forEach((r) => {
      const tr = el(`<tr>
        <td>${esc(r.party)}</td>
        <td class="mono">${esc(r.partNumber)}</td>
        <td class="mono">${r.requested}</td>
        <td class="mono">${r.received}</td>
        <td class="mono">${r.writtenOff}</td>
        <td class="mono">${r.pending > 0 ? `<span class="stock-dot low"></span>` : ""}${r.pending}</td>
        <td>${r.oldestSince ? relTime(r.oldestSince) : "—"}</td>
      </tr>`);
      if (person.role === "owner") {
        const td = el(`<td></td>`);
        if (r.pending > 0) {
          const btn = el(`<button class="btn small danger">Write off</button>`);
          btn.addEventListener("click", async () => {
            const qtyStr = prompt(`Write off how many units of ${r.partNumber} from ${r.party}? (pending: ${r.pending})`, String(r.pending));
            if (qtyStr === null) return;
            const qty = Number(qtyStr);
            if (!qty || qty <= 0) return toast("Enter a valid quantity.", true);
            const note = prompt("Reason (optional):", "") || "";
            try {
              await api("/api/other-writeoffs", { method: "POST", body: { party: r.party, partNumber: r.partNumber, qty, note, writtenOffBy: person.name } });
              toast("Written off.");
              await loadState();
            } catch (e) {
              toast(e.message, true);
            }
          });
          td.appendChild(btn);
        }
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    });
    panel.appendChild(tw);
    return panel;
  }

  function renderOtherActivityPanel() {
    const panel = el(`<div class="panel"><h2>Recent activity</h2></div>`);
    const entries = [
      ...state.otherRequests.map((r) => ({ type: "Requested", party: r.party, date: r.date, by: r.loggedBy, items: r.items })),
      ...state.otherStockIns.map((s) => ({ type: "Received", party: s.party, date: s.date, by: s.receivedBy, items: s.items })),
    ]
      .sort((a, b) => new Date(b.date) - new Date(a.date))
      .slice(0, 8);
    if (entries.length === 0) {
      panel.appendChild(el(`<div class="empty">Nothing logged yet.</div>`));
      return panel;
    }
    entries.forEach((e) => {
      const summary = e.items.map((it) => `${it.partNumber} ×${it.qty}`).join(", ");
      panel.appendChild(el(`<div class="order-meta" style="padding:5px 0;"><strong>${e.type}</strong> · ${esc(e.party)} · ${esc(e.by)} · ${relTime(e.date)} — ${esc(summary)}</div>`));
    });
    return panel;
  }

  // ---------- Team ----------
  function renderTeamView() {
    const wrap = el(`<div></div>`);
    wrap.appendChild(el(`<div class="view-header"><h1>Team</h1><div class="sub">${state.team.length} people</div></div>`));

    const addBtn = el(`<button class="btn" style="margin-bottom:14px;">+ Add team member</button>`);
    const bodyHolder = el(`<div></div>`);
    addBtn.addEventListener("click", () => {
      showAddTeam = !showAddTeam;
      rerender();
    });
    wrap.appendChild(addBtn);
    wrap.appendChild(bodyHolder);

    function rerender() {
      bodyHolder.innerHTML = "";
      if (showAddTeam) bodyHolder.appendChild(renderAddTeamForm());
      const list = el(`<div class="stack"></div>`);
      state.team.forEach((t) => {
        const row = el(`
          <div class="panel" style="display:flex;align-items:center;justify-content:space-between;gap:10px;">
            <div><strong>${esc(t.name)}</strong> <span class="role-chip">${ROLE_LABEL[t.role]}</span></div>
            <div style="display:flex;gap:8px;"></div>
          </div>
        `);
        const actions = $("div > div", row) || row.lastElementChild;
        const roleSel = el(`<select style="width:auto;">${Object.keys(ROLE_LABEL).map((r) => `<option value="${r}"${r === t.role ? " selected" : ""}>${ROLE_LABEL[r]}</option>`).join("")}</select>`);
        roleSel.addEventListener("change", async () => {
          await api(`/api/team/${t.id}`, { method: "PUT", body: { role: roleSel.value } });
          toast("Role updated.");
          await loadState();
        });
        const removeBtn = el(`<button class="btn small danger">Remove</button>`);
        removeBtn.addEventListener("click", async () => {
          if (!confirm(`Remove ${t.name} from the team?`)) return;
          await api(`/api/team/${t.id}`, { method: "DELETE" });
          toast("Removed.");
          await loadState();
        });
        actions.appendChild(roleSel);
        actions.appendChild(removeBtn);
        list.appendChild(row);
      });
      bodyHolder.appendChild(list);
    }
    rerender();

    return wrap;
  }

  function renderAddTeamForm() {
    const panel = el(`
      <div class="panel">
        <h2>Add team member</h2>
        <div class="form-grid">
          <div class="field"><label>Name</label><input type="text" id="at-name" /></div>
          <div class="field"><label>Role</label>
            <select id="at-role">
              <option value="order-taker">Order taker</option>
              <option value="biller">Biller</option>
              <option value="dispatch">Dispatch</option>
              <option value="owner">Owner</option>
            </select>
          </div>
        </div>
        <button class="btn primary" style="margin-top:12px;">Save</button>
      </div>
    `);
    $("button", panel).addEventListener("click", async () => {
      const name = $("#at-name", panel).value.trim();
      if (!name) return toast("Enter a name.", true);
      await api("/api/team", { method: "POST", body: { name, role: $("#at-role", panel).value } });
      showAddTeam = false;
      toast("Team member added.");
      await loadState();
    });
    return panel;
  }

  // ---------- boot ----------
  async function boot() {
    if ("serviceWorker" in navigator) {
      navigator.serviceWorker.register("sw.js").catch(() => {}); // enables "Add to Home Screen"
    }
    await loadState({ silent: true });
    render();
    setInterval(async () => {
      if (isEditingSomething()) return;
      await loadState({ silent: true });
      render();
    }, 4000);
  }
  boot();
})();
