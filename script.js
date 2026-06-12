/**
 * CineStock — Restaurant Inventory Management System
 * script.js
 *
 * Architecture:
 *   STATE  →  localStorage persistence
 *   PARSE  →  SheetJS Excel parser (Recipes_Combined.xlsx structure)
 *   UI     →  Tab navigation, modals, tables, cards
 *   SELL   →  Deducts recipe ingredients from inventory
 */

'use strict';

/* ═══════════════════════════════════════════════
   1. STATE MANAGEMENT (localStorage)
═══════════════════════════════════════════════ */

const KEYS = {
  inventory : 'cinestock_inventory',   // { [name]: { unit, stock, threshold } }
  menuItems : 'cinestock_menu',        // [ { name, price, category, recipe:[{ingredient,unit,qty}] } ]
  sales     : 'cinestock_sales',       // [ { ts, item, qty, price, ingredients } ]
  dailyUsage: 'cinestock_daily',       // { date: { [ingredient]: totalUsed } }
};

/** Load JSON from localStorage, return defaultVal on miss/parse error */
function load(key, defaultVal = null) {
  try {
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : defaultVal;
  } catch { return defaultVal; }
}

/** Save value to localStorage as JSON */
function save(key, val) {
  localStorage.setItem(key, JSON.stringify(val));
}

/* ═══════════════════════════════════════════════
   2. EXCEL PARSER  (SheetJS)
   Handles the specific layout of Recipes_Combined.xlsx:
   - Sheet "Recipe Items"   → menu items with ingredient recipes
   - Sheet "Non Recipe Items" → items sold as-is (tracked as inventory units)
═══════════════════════════════════════════════ */

/**
 * Parse "Recipe Items" sheet.
 * Layout pattern per item:
 *   Row A: col[1] = Item Name (non-NaN)
 *   Row B: col[2]='MENU', col[3]='UNIT', col[4]='QTY', col[5]=Price
 *   Rows C+: col[2]=ingredient, col[3]=unit, col[4]=qty  (until next item name)
 */
function parseRecipeSheet(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
  const menuItems = [];
  let current = null;

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const col1 = r[1] ? String(r[1]).trim() : '';
    const col2 = r[2] ? String(r[2]).trim() : '';
    const col3 = r[3] ? String(r[3]).trim() : '';
    const col4 = r[4];
    const col5 = r[5];

    // ── New menu item detected: col1 has text, col2 is empty or not 'MENU'
    if (col1 && col1 !== 'NaN' && col2.toUpperCase() !== 'MENU' && !isIngredientRow(col2)) {
      if (current) menuItems.push(current);

      // Derive category from name keywords
      const cat = detectCategory(col1);
      current = { name: col1, category: cat, price: 0, recipe: [] };
    }

    // ── Price/header row: col2 === 'MENU'
    if (col2.toUpperCase() === 'MENU' && current) {
      const price = parseFloat(col5);
      if (!isNaN(price)) current.price = price;
    }

    // ── Ingredient row: col2 is ingredient name, col3 is unit, col4 is qty
    if (col2 && col2.toUpperCase() !== 'MENU' && col3 && col4 !== null && current) {
      const qty = parseFloat(col4);
      if (!isNaN(qty) && qty > 0 && !col2.toLowerCase().includes('wastage') && !col2.toLowerCase().includes('wadtage')) {
        current.recipe.push({
          ingredient: col2.trim(),
          unit: col3.trim(),
          qty: qty,
        });
      }
    }
  }
  if (current) menuItems.push(current);
  return menuItems.filter(m => m.recipe.length > 0);
}

/** Helper: does this string look like an ingredient name row (not a header/item)? */
function isIngredientRow(s) {
  if (!s) return false;
  const skip = ['menu', 'unit', 'qty', 'price', 'isb', 'lhr', 'khi', 'arena', 'rcgp', 'vendor'];
  return !skip.includes(s.toLowerCase());
}

/** Derive category from item name */
function detectCategory(name) {
  const n = name.toLowerCase();
  if (n.includes('pop-corn') || n.includes('popcorn')) return 'Popcorn';
  if (n.includes('nachos'))   return 'Snacks';
  if (n.includes('fries') || n.includes('curly') || n.includes('garlic') || n.includes('french') || n.includes('cheesy')) return 'Fries';
  if (n.includes('coffee') || n.includes('latte') || n.includes('cappuccino')) return 'Coffee';
  if (n.includes('tea'))      return 'Tea';
  if (n.includes('chocolate')) return 'Hot Drinks';
  if (n.includes('juice') || n.includes('slush')) return 'Cold Drinks';
  if (n.includes('coke') || n.includes('cola') || n.includes('sprite')) return 'Soda';
  if (n.includes('hot dog'))  return 'Hot Food';
  if (n.includes('nugget') || n.includes('wings')) return 'Chicken';
  return 'Other';
}

/**
 * Parse "Non Recipe Items" sheet.
 * These items are tracked as single-unit inventory (no sub-ingredients).
 * Layout: row[1]=serial, row[2]=item name, various price columns.
 */
function parseNonRecipeSheet(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
  const items = [];

  for (let i = 5; i < rows.length; i++) {  // data starts ~row 5
    const r = rows[i];
    const name = r[2] ? String(r[2]).trim() : '';
    if (!name || name === 'NaN') continue;

    // Try to get selling price (SP column, typically col index 9 or 10)
    let price = 0;
    for (const idx of [10, 9, 7, 6]) {
      const v = parseFloat(r[idx]);
      if (!isNaN(v) && v > 50) { price = v; break; }
    }

    items.push({
      name,
      category: detectCategory(name),
      price,
      recipe: [{ ingredient: name, unit: 'pcs', qty: 1 }],  // 1 unit = 1 sold
    });
  }
  return items;
}

/**
 * Extract all unique ingredients from parsed menu items
 * and seed inventory with defaults if not already present.
 */
function seedInventoryFromMenu(menuItems) {
  const inv = load(KEYS.inventory, {});
  let added = 0;

  menuItems.forEach(item => {
    item.recipe.forEach(r => {
      if (!inv[r.ingredient]) {
        // Default starting stock: 5 kg / 5000 g / 100 pieces based on unit
        const defaultStock = getDefaultStock(r.unit, r.qty);
        inv[r.ingredient] = {
          unit      : r.unit,
          stock     : defaultStock,
          threshold : Math.max(defaultStock * 0.15, r.qty * 5), // 15% of default or 5× recipe qty
        };
        added++;
      }
    });
  });

  save(KEYS.inventory, inv);
  return added;
}

/** Sensible default stock for a newly imported ingredient */
function getDefaultStock(unit, recipeQty) {
  const u = (unit || '').toLowerCase();
  if (['grm', 'g', 'gm', 'gram'].includes(u)) return 5000;
  if (['kg', 'kgs', 'kilogram'].includes(u))   return 5;
  if (['ml', 'mls'].includes(u))               return 3000;
  if (['ltr', 'l', 'litre', 'liter'].includes(u)) return 5;
  if (['no', 'pcs', 'piece', 'pieces'].includes(u)) return 200;
  return Math.ceil(recipeQty * 50);  // generic: 50× recipe qty
}

/* ═══════════════════════════════════════════════
   3. SELL LOGIC
═══════════════════════════════════════════════ */

/**
 * Check if a menu item can be sold (all ingredients in stock)
 * Returns { ok: bool, missing: [ { ingredient, needed, have } ] }
 */
function checkStock(menuItem, qty = 1) {
  const inv = load(KEYS.inventory, {});
  const missing = [];

  menuItem.recipe.forEach(r => {
    const needed = r.qty * qty;
    const ingData = inv[r.ingredient];
    const have = ingData ? ingData.stock : 0;
    if (have < needed) {
      missing.push({ ingredient: r.ingredient, needed, have, unit: r.unit });
    }
  });

  return { ok: missing.length === 0, missing };
}

/**
 * Execute a sale: deduct ingredients, log sale, update daily usage.
 * Returns updated inventory.
 */
function executeSale(menuItem, qty = 1) {
  const inv = load(KEYS.inventory, {});
  const today = todayStr();
  const daily = load(KEYS.dailyUsage, {});
  if (!daily[today]) daily[today] = {};

  const usedIngredients = [];

  menuItem.recipe.forEach(r => {
    const deduct = r.qty * qty;
    if (inv[r.ingredient]) {
      inv[r.ingredient].stock = Math.max(0, inv[r.ingredient].stock - deduct);
    } else {
      // Create entry if missing (shouldn't happen, but safe)
      inv[r.ingredient] = { unit: r.unit, stock: 0, threshold: r.qty * 5 };
    }

    // Track daily usage
    daily[today][r.ingredient] = (daily[today][r.ingredient] || 0) + deduct;

    usedIngredients.push({ ingredient: r.ingredient, unit: r.unit, qty: deduct });
  });

  save(KEYS.inventory, inv);
  save(KEYS.dailyUsage, daily);

  // Log sale
  const sales = load(KEYS.sales, []);
  sales.unshift({
    ts         : new Date().toISOString(),
    item       : menuItem.name,
    qty        : qty,
    price      : menuItem.price,
    revenue    : menuItem.price * qty,
    ingredients: usedIngredients,
  });
  save(KEYS.sales, sales);

  return inv;
}

/* ═══════════════════════════════════════════════
   4. UI HELPERS
═══════════════════════════════════════════════ */

/** Format number with up to 3 decimal places, stripping trailing zeros */
function fmt(n) {
  const f = parseFloat(n);
  if (isNaN(f)) return '0';
  return parseFloat(f.toFixed(3)).toString();
}

/** Today as YYYY-MM-DD */
function todayStr() {
  return new Date().toISOString().slice(0, 10);
}

/** Friendly datetime */
function fmtDateTime(iso) {
  const d = new Date(iso);
  return d.toLocaleDateString('en-PK', { day: '2-digit', month: 'short', year: 'numeric' })
    + ' ' + d.toLocaleTimeString('en-PK', { hour: '2-digit', minute: '2-digit' });
}

/** Show a toast notification */
function toast(msg, type = 'info') {
  const icons = { success: '✅', error: '❌', warning: '⚠️', info: 'ℹ️' };
  const el = document.createElement('div');
  el.className = `toast toast-${type}`;
  el.innerHTML = `<span class="toast-icon">${icons[type] || 'ℹ️'}</span>
                  <span class="toast-msg">${msg}</span>`;
  document.getElementById('toast-container').appendChild(el);
  setTimeout(() => el.remove(), 4000);
}

/** Get stock status */
function stockStatus(stock, threshold) {
  if (stock <= 0)         return 'out';
  if (stock <= threshold) return 'low';
  return 'ok';
}

/** Status label HTML */
function statusHTML(status) {
  const map = { ok: ['✅ In Stock', 'status-ok'], low: ['⚠️ Low Stock', 'status-low'], out: ['❌ Out of Stock', 'status-out'] };
  const [label, cls] = map[status] || map.ok;
  return `<span class="${cls}">${label}</span>`;
}

/* ═══════════════════════════════════════════════
   5. TAB NAVIGATION
═══════════════════════════════════════════════ */

function initTabs() {
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const tabId = btn.dataset.tab;

      // Update active button
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');

      // Show active panel
      document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
      document.getElementById(`tab-${tabId}`).classList.add('active');

      // Refresh relevant panels
      if (tabId === 'dashboard') renderDashboard();
      if (tabId === 'menu')      renderMenu();
      if (tabId === 'inventory') renderInventory();
      if (tabId === 'sales')     renderSales();
    });
  });
}

/* ═══════════════════════════════════════════════
   6. DASHBOARD RENDER
═══════════════════════════════════════════════ */

function renderDashboard() {
  const inv     = load(KEYS.inventory, {});
  const menu    = load(KEYS.menuItems, []);
  const sales   = load(KEYS.sales, []);
  const daily   = load(KEYS.dailyUsage, {});
  const today   = todayStr();
  const todayUsage = daily[today] || {};

  // KPIs
  const entries   = Object.entries(inv);
  const totalIng  = entries.length;
  const inStock   = entries.filter(([,v]) => stockStatus(v.stock, v.threshold) === 'ok').length;
  const lowStock  = entries.filter(([,v]) => stockStatus(v.stock, v.threshold) === 'low').length;
  const outStock  = entries.filter(([,v]) => stockStatus(v.stock, v.threshold) === 'out').length;
  const todaySales= sales.filter(s => s.ts.startsWith(today)).reduce((a, s) => a + s.qty, 0);

  document.getElementById('kpi-total-ingredients').textContent = totalIng;
  document.getElementById('kpi-in-stock').textContent          = inStock;
  document.getElementById('kpi-low-stock').textContent         = lowStock;
  document.getElementById('kpi-out-stock').textContent         = outStock;
  document.getElementById('kpi-menu-items').textContent        = menu.length;
  document.getElementById('kpi-today-sales').textContent       = todaySales;
  document.getElementById('today-date').textContent            = new Date().toLocaleDateString('en-PK', { weekday:'long', day:'numeric', month:'long', year:'numeric' });
  document.getElementById('low-stock-count').textContent       = `${lowStock + outStock} items`;

  // Low Stock Table
  const lowTbody = document.getElementById('low-stock-tbody');
  const alerts   = entries.filter(([,v]) => stockStatus(v.stock, v.threshold) !== 'ok');
  if (alerts.length === 0) {
    lowTbody.innerHTML = `<tr><td colspan="6" class="empty-msg">No low stock items 🎉 Everything is well stocked!</td></tr>`;
  } else {
    lowTbody.innerHTML = alerts.map(([name, v]) => {
      const status = stockStatus(v.stock, v.threshold);
      return `<tr>
        <td><strong>${name}</strong></td>
        <td>${v.unit}</td>
        <td>${fmt(v.stock)} ${v.unit}</td>
        <td>${fmt(v.threshold)} ${v.unit}</td>
        <td>${statusHTML(status)}</td>
        <td><button class="btn btn-primary btn-sm" onclick="openStockModal('${encodeURIComponent(name)}')">📦 Add Stock</button></td>
      </tr>`;
    }).join('');
  }

  // Daily Usage Table
  const usageTbody = document.getElementById('daily-usage-tbody');
  const usageEntries = Object.entries(todayUsage);
  if (usageEntries.length === 0) {
    usageTbody.innerHTML = `<tr><td colspan="4" class="empty-msg">No usage recorded today. Start selling!</td></tr>`;
  } else {
    usageTbody.innerHTML = usageEntries.map(([name, used]) => {
      const v = inv[name] || { unit: '', stock: 0 };
      return `<tr>
        <td><strong>${name}</strong></td>
        <td>${fmt(used)} ${v.unit}</td>
        <td>${v.unit}</td>
        <td>${fmt(v.stock)} ${v.unit}</td>
      </tr>`;
    }).join('');
  }
}

/* ═══════════════════════════════════════════════
   7. MENU RENDER
═══════════════════════════════════════════════ */

function renderMenu(searchVal = '', filterCat = 'all') {
  const menu = load(KEYS.menuItems, []);
  const inv  = load(KEYS.inventory, {});
  const grid = document.getElementById('menu-grid');

  // Populate category filter
  const filterSel = document.getElementById('menu-filter');
  const cats = ['all', ...new Set(menu.map(m => m.category))];
  filterSel.innerHTML = cats.map(c => `<option value="${c}">${c === 'all' ? 'All Categories' : c}</option>`).join('');
  filterSel.value = filterCat;

  if (menu.length === 0) {
    grid.innerHTML = `<div class="empty-msg" style="grid-column:1/-1;padding:40px;text-align:center;">
      No menu items loaded.<br><br>
      <button class="btn btn-primary" onclick="switchTab('import')">📥 Import Excel File</button>
    </div>`;
    return;
  }

  // Filter
  let items = menu;
  if (searchVal) {
    const q = searchVal.toLowerCase();
    items = items.filter(m => m.name.toLowerCase().includes(q) || m.category.toLowerCase().includes(q));
  }
  if (filterCat !== 'all') {
    items = items.filter(m => m.category === filterCat);
  }

  if (items.length === 0) {
    grid.innerHTML = `<div class="empty-msg" style="grid-column:1/-1;padding:40px;text-align:center;">No items match your search.</div>`;
    return;
  }

  grid.innerHTML = items.map(item => {
    const { ok } = checkStock(item, 1);
    const recipeLines = item.recipe.slice(0, 4).map(r =>
      `<strong>${r.ingredient}</strong>: ${fmt(r.qty)} ${r.unit}`
    ).join('<br>');
    const more = item.recipe.length > 4 ? `<br>+${item.recipe.length - 4} more...` : '';

    return `<div class="menu-card">
      <div>
        <div class="menu-card-cat">${item.category}</div>
        <div class="menu-card-name">${item.name}</div>
      </div>
      <div class="menu-card-price">PKR ${item.price.toLocaleString()} <span>/ serving</span></div>
      <div class="menu-card-recipe">${recipeLines}${more}</div>
      <div class="menu-card-footer">
        <button class="sell-btn" ${!ok ? 'disabled title="Insufficient stock"' : ''}
          onclick="openSellModal(${JSON.stringify(JSON.stringify(item)).slice(1,-1)})">
          ${ok ? '🛒 Sell' : '❌ Out of Stock'}
        </button>
      </div>
    </div>`;
  }).join('');
}

/* ═══════════════════════════════════════════════
   8. INVENTORY RENDER
═══════════════════════════════════════════════ */

function renderInventory(searchVal = '', filterStatus = 'all') {
  const inv = load(KEYS.inventory, {});
  const tbody = document.getElementById('inv-tbody');
  let entries = Object.entries(inv);

  // Filter
  if (searchVal) {
    const q = searchVal.toLowerCase();
    entries = entries.filter(([name]) => name.toLowerCase().includes(q));
  }
  if (filterStatus !== 'all') {
    entries = entries.filter(([, v]) => stockStatus(v.stock, v.threshold) === filterStatus);
  }

  // Sort: out → low → ok
  const order = { out: 0, low: 1, ok: 2 };
  entries.sort((a, b) => order[stockStatus(a[1].stock, a[1].threshold)] - order[stockStatus(b[1].stock, b[1].threshold)]);

  if (entries.length === 0) {
    tbody.innerHTML = `<tr><td colspan="6" class="empty-msg">No ingredients found.</td></tr>`;
    return;
  }

  tbody.innerHTML = entries.map(([name, v]) => {
    const status = stockStatus(v.stock, v.threshold);
    const pct    = v.threshold > 0 ? Math.min(100, (v.stock / (v.threshold * 5)) * 100) : 100;
    const fillCls = status === 'ok' ? 'fill-ok' : status === 'low' ? 'fill-low' : 'fill-out';
    const encName = encodeURIComponent(name);

    return `<tr>
      <td><strong>${name}</strong></td>
      <td>${v.unit}</td>
      <td>
        <div class="stock-bar-wrap">
          <div class="stock-bar"><div class="stock-bar-fill ${fillCls}" style="width:${pct}%"></div></div>
          <span class="stock-num">${fmt(v.stock)}</span>
        </div>
      </td>
      <td>${fmt(v.threshold)} ${v.unit}</td>
      <td>${statusHTML(status)}</td>
      <td style="display:flex;gap:6px;flex-wrap:wrap;">
        <button class="btn btn-primary btn-sm" onclick="openStockModal('${encName}')">📦 Add Stock</button>
        <button class="btn btn-danger btn-sm" onclick="deleteIngredient('${encName}')">🗑️</button>
      </td>
    </tr>`;
  }).join('');
}

/* ═══════════════════════════════════════════════
   9. SALES LOG RENDER
═══════════════════════════════════════════════ */

function renderSales(searchVal = '') {
  let sales = load(KEYS.sales, []);
  const tbody = document.getElementById('sales-tbody');

  if (searchVal) {
    const q = searchVal.toLowerCase();
    sales = sales.filter(s => s.item.toLowerCase().includes(q));
  }

  if (sales.length === 0) {
    tbody.innerHTML = `<tr><td colspan="6" class="empty-msg">No sales recorded yet. Start selling!</td></tr>`;
    return;
  }

  tbody.innerHTML = sales.map(s => {
    const ingList = s.ingredients.map(i => `${i.ingredient}: ${fmt(i.qty)} ${i.unit}`).join(', ');
    return `<tr>
      <td style="white-space:nowrap;">${fmtDateTime(s.ts)}</td>
      <td><strong>${s.item}</strong></td>
      <td>${s.qty}</td>
      <td>PKR ${s.price.toLocaleString()}</td>
      <td><strong style="color:var(--teal)">PKR ${s.revenue.toLocaleString()}</strong></td>
      <td style="font-size:0.75rem;color:var(--text-muted)">${ingList}</td>
    </tr>`;
  }).join('');
}

/* ═══════════════════════════════════════════════
   10. SELL MODAL
═══════════════════════════════════════════════ */

let _sellItem = null;  // currently selected menu item for sale

/** Open sell confirmation modal */
function openSellModal(itemJSON) {
  const item = JSON.parse(itemJSON);
  _sellItem = item;

  document.getElementById('sell-modal-title').textContent = `Sell: ${item.name}`;
  document.getElementById('sell-qty').value = 1;
  document.getElementById('sell-stock-warning').
