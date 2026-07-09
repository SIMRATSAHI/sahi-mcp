require('dotenv').config();
const express = require('express');
const cors = require('cors');
const path = require('path');
const crypto = require('crypto');
const { Pool } = require('pg');
const multer = require('multer');
const xlsx = require('xlsx');

const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: { rejectUnauthorized: false }
});

// --- Multi-site inventory config ---
// org_id 1 = HQ / Shanghai warehouse (source of all stock). Buyers below are
// treated as their own stock-holding sites once stock is transferred to them.
const HQ_ORG_ID = 1;
const BUYER_ORG_MAP = { CANADA: 2, UK: 3 };

// --- Additive, idempotent schema migration (safe to run every boot) ---
async function ensureSchema() {
  try {
    await pool.query(`ALTER TABLE po_line_items ADD COLUMN IF NOT EXISTS received_qty NUMERIC NOT NULL DEFAULT 0`);
  } catch (err) {
    console.error('Schema migration warning:', err.message);
  }
  try {
    await pool.query(`
      CREATE TABLE IF NOT EXISTS users (
        id SERIAL PRIMARY KEY,
        email TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        password_salt TEXT NOT NULL,
        role TEXT NOT NULL CHECK (role IN ('ADMIN', 'BUYER', 'ACCOUNTS', 'LOGISTICS')),
        name TEXT,
        is_active BOOLEAN NOT NULL DEFAULT true,
        created_at TIMESTAMP NOT NULL DEFAULT NOW()
      )
    `);
    // Widen the role CHECK constraint if the table already existed from an earlier
    // version (before the LOGISTICS role was added).
    await pool.query(`ALTER TABLE users DROP CONSTRAINT IF EXISTS users_role_check`);
    await pool.query(`ALTER TABLE users ADD CONSTRAINT users_role_check CHECK (role IN ('ADMIN', 'BUYER', 'ACCOUNTS', 'LOGISTICS'))`);
  } catch (err) {
    console.error('Users table migration warning:', err.message);
  }
  try {
    // Who created each document - shown directly on the document/list itself.
    await pool.query(`ALTER TABLE purchase_orders ADD COLUMN IF NOT EXISTS created_by TEXT`);
    await pool.query(`ALTER TABLE customer_orders ADD COLUMN IF NOT EXISTS created_by TEXT`);
    await pool.query(`ALTER TABLE item_master ADD COLUMN IF NOT EXISTS created_by TEXT`);
  } catch (err) {
    console.error('created_by migration warning:', err.message);
  }
  try {
    // Full audit trail of every meaningful action, independent of the documents
    // themselves (covers deletes, receipts, password resets, logins, etc).
    await pool.query(`
      CREATE TABLE IF NOT EXISTS activity_log (
        id SERIAL PRIMARY KEY,
        user_email TEXT,
        user_role TEXT,
        action TEXT NOT NULL,
        reference_id TEXT,
        details JSONB,
        created_at TIMESTAMP NOT NULL DEFAULT NOW()
      )
    `);
  } catch (err) {
    console.error('activity_log migration warning:', err.message);
  }
}
ensureSchema();

// Record an action in the audit trail. Never lets a logging failure break the
// actual request - logging is best-effort.
async function logActivity(user, action, referenceId, details) {
  try {
    await pool.query(
      `INSERT INTO activity_log (user_email, user_role, action, reference_id, details) VALUES ($1, $2, $3, $4, $5)`,
      [user ? user.email : null, user ? user.role : null, action, referenceId || null, details ? JSON.stringify(details) : null]
    );
  } catch (err) {
    console.error('Failed to write activity log:', err.message);
  }
}

// --- Auth: lightweight, dependency-free (no bcrypt/express-session needed) ---
// Passwords hashed with Node's built-in scrypt; sessions are random tokens kept
// in an in-memory Map and handed to the browser as an HttpOnly cookie.
const sessions = new Map(); // sessionId -> { userId, email, role, name }

function hashPassword(password, existingSalt) {
  const salt = existingSalt || crypto.randomBytes(16).toString('hex');
  const hash = crypto.scryptSync(password, salt, 64).toString('hex');
  return { salt, hash };
}

function verifyPassword(password, salt, hash) {
  const check = crypto.scryptSync(password, salt, 64).toString('hex');
  const a = Buffer.from(check, 'hex');
  const b = Buffer.from(hash, 'hex');
  return a.length === b.length && crypto.timingSafeEqual(a, b);
}

function parseCookies(req) {
  const header = req.headers.cookie;
  const cookies = {};
  if (!header) return cookies;
  header.split(';').forEach(pair => {
    const idx = pair.indexOf('=');
    if (idx === -1) return;
    cookies[pair.slice(0, idx).trim()] = decodeURIComponent(pair.slice(idx + 1).trim());
  });
  return cookies;
}

function getSessionUser(req) {
  const sid = parseCookies(req)['sahi_session'];
  if (!sid) return null;
  return sessions.get(sid) || null;
}

function defaultLandingFor(role) {
  if (role === 'ADMIN') return '/index.html';
  if (role === 'BUYER') return '/create-po.html';
  if (role === 'ACCOUNTS') return '/accounts.html';
  if (role === 'LOGISTICS') return '/receive-po.html';
  return '/login.html';
}

// Gate a static HTML page: redirects (browser-friendly) instead of JSON errors.
function requireAuthPage(rolesAllowed) {
  return (req, res, next) => {
    const user = getSessionUser(req);
    if (!user) return res.redirect('/login.html');
    if (rolesAllowed && !rolesAllowed.includes(user.role)) {
      return res.redirect(defaultLandingFor(user.role));
    }
    req.user = user;
    next();
  };
}

// Gate an API endpoint: returns JSON errors instead of redirecting.
function requireAuthApi(rolesAllowed) {
  return (req, res, next) => {
    const user = getSessionUser(req);
    if (!user) return res.status(401).json({ error: 'Not authenticated' });
    if (rolesAllowed && !rolesAllowed.includes(user.role)) {
      return res.status(403).json({ error: 'Not authorized for this action' });
    }
    req.user = user;
    next();
  };
}

// Increase (or decrease, with a negative delta) on-hand quantity for a sku at a site.
async function adjustInventory(client, sku, org_id, delta) {
  const existing = await client.query("SELECT quantity_on_hand FROM inventory WHERE sku = $1 AND org_id = $2 FOR UPDATE", [sku, org_id]);
  if (existing.rows.length > 0) {
    await client.query("UPDATE inventory SET quantity_on_hand = quantity_on_hand + $1 WHERE sku = $2 AND org_id = $3", [delta, sku, org_id]);
  } else {
    await client.query("INSERT INTO inventory (sku, org_id, quantity_on_hand) VALUES ($1, $2, $3)", [sku, org_id, delta]);
  }
}

async function getOnHand(client, sku, org_id) {
  const r = await client.query("SELECT quantity_on_hand FROM inventory WHERE sku = $1 AND org_id = $2", [sku, org_id]);
  return r.rows.length > 0 ? parseFloat(r.rows[0].quantity_on_hand) : 0;
}

const app = express();
app.use(cors());
app.use(express.json());

// --- Role-gated pages (must be registered before express.static, so static
// never serves these files directly without the auth/role check running first) ---
// Admin: everything. Buyer: procurement (Create PO / Item Master / Receive Stock).
// Accounts: read-only PO & financials view.
app.get(['/', '/index.html'], requireAuthPage(['ADMIN']), (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));
app.get(['/create-po.html', '/create-po'], requireAuthPage(['ADMIN', 'BUYER']), (req, res) => res.sendFile(path.join(__dirname, 'public', 'create-po.html')));
app.get(['/items.html'], requireAuthPage(['ADMIN', 'BUYER']), (req, res) => res.sendFile(path.join(__dirname, 'public', 'items.html')));
app.get(['/receive-po.html', '/receive-po'], requireAuthPage(['ADMIN', 'BUYER', 'LOGISTICS']), (req, res) => res.sendFile(path.join(__dirname, 'public', 'receive-po.html')));
app.get(['/transfer.html'], requireAuthPage(['ADMIN']), (req, res) => res.sendFile(path.join(__dirname, 'public', 'transfer.html')));
app.get(['/accounts.html'], requireAuthPage(['ADMIN', 'ACCOUNTS']), (req, res) => res.sendFile(path.join(__dirname, 'public', 'accounts.html')));
app.get(['/admin-users.html'], requireAuthPage(['ADMIN']), (req, res) => res.sendFile(path.join(__dirname, 'public', 'admin-users.html')));

app.use(express.static('public'));

// --- Auth API ---
app.get('/api/setup-status', async (req, res) => {
  try {
    const r = await pool.query('SELECT COUNT(*)::int AS count FROM users');
    res.json({ needsSetup: r.rows[0].count === 0 });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to check setup status' });
  }
});

app.post('/api/setup', async (req, res) => {
  try {
    const countRes = await pool.query('SELECT COUNT(*)::int AS count FROM users');
    if (countRes.rows[0].count > 0) return res.status(403).json({ error: 'Setup already completed - use the login page.' });
    const { email, password, name } = req.body;
    if (!email || !password) return res.status(400).json({ error: 'email and password are required' });
    if (password.length < 8) return res.status(400).json({ error: 'Password must be at least 8 characters.' });
    const { salt, hash } = hashPassword(password);
    await pool.query(
      'INSERT INTO users (email, password_hash, password_salt, role, name) VALUES ($1, $2, $3, $4, $5)',
      [email, hash, salt, 'ADMIN', name || 'Admin']
    );
    res.json({ success: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to complete setup' });
  }
});

app.post('/api/login', async (req, res) => {
  try {
    const { email, password } = req.body;
    if (!email || !password) return res.status(400).json({ error: 'email and password are required' });
    const r = await pool.query('SELECT * FROM users WHERE email = $1 AND is_active = true', [email]);
    if (r.rows.length === 0) return res.status(401).json({ error: 'Invalid email or password' });
    const u = r.rows[0];
    if (!verifyPassword(password, u.password_salt, u.password_hash)) {
      return res.status(401).json({ error: 'Invalid email or password' });
    }
    const sid = crypto.randomBytes(32).toString('hex');
    sessions.set(sid, { userId: u.id, email: u.email, role: u.role, name: u.name });
    res.setHeader('Set-Cookie', `sahi_session=${sid}; HttpOnly; Path=/; SameSite=Lax; Max-Age=${60 * 60 * 24 * 7}`);
    logActivity({ email: u.email, role: u.role }, 'LOGIN', null, null);
    res.json({ success: true, role: u.role, name: u.name, landing: defaultLandingFor(u.role) });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Login failed' });
  }
});

app.post('/api/logout', (req, res) => {
  const sid = parseCookies(req)['sahi_session'];
  if (sid) sessions.delete(sid);
  res.setHeader('Set-Cookie', `sahi_session=; HttpOnly; Path=/; Max-Age=0`);
  res.json({ success: true });
});

app.get('/api/me', (req, res) => {
  const user = getSessionUser(req);
  if (!user) return res.status(401).json({ error: 'Not authenticated' });
  res.json({ email: user.email, role: user.role, name: user.name });
});

app.get('/api/admin/users', requireAuthApi(['ADMIN']), async (req, res) => {
  try {
    const r = await pool.query('SELECT id, email, role, name, is_active, created_at FROM users ORDER BY created_at');
    res.json(r.rows);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to fetch users' });
  }
});

// Full audit trail - who did what, when. Admin only.
app.get('/api/activity-log', requireAuthApi(['ADMIN']), async (req, res) => {
  try {
    const r = await pool.query('SELECT * FROM activity_log ORDER BY created_at DESC LIMIT 200');
    res.json(r.rows);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to fetch activity log' });
  }
});

app.post('/api/admin/users', requireAuthApi(['ADMIN']), async (req, res) => {
  try {
    const { email, password, role, name } = req.body;
    if (!email || !password || !role) return res.status(400).json({ error: 'email, password, and role are required' });
    if (!['ADMIN', 'BUYER', 'ACCOUNTS', 'LOGISTICS'].includes(role)) return res.status(400).json({ error: 'Invalid role' });
    if (password.length < 8) return res.status(400).json({ error: 'Password must be at least 8 characters.' });
    const { salt, hash } = hashPassword(password);
    await pool.query(
      'INSERT INTO users (email, password_hash, password_salt, role, name) VALUES ($1, $2, $3, $4, $5)',
      [email, hash, salt, role, name || null]
    );
    logActivity(req.user, 'USER_CREATED', email, { role, name: name || null });
    res.json({ success: true });
  } catch (err) {
    if (err.code === '23505') return res.status(409).json({ error: 'A user with that email already exists' });
    console.error(err);
    res.status(500).json({ error: 'Failed to create user' });
  }
});

app.patch('/api/admin/users/:id/active', requireAuthApi(['ADMIN']), async (req, res) => {
  try {
    const { is_active } = req.body;
    const r = await pool.query('UPDATE users SET is_active = $1 WHERE id = $2 RETURNING id, email, is_active', [!!is_active, req.params.id]);
    if (r.rows.length === 0) return res.status(404).json({ error: 'User not found' });
    logActivity(req.user, is_active ? 'USER_ENABLED' : 'USER_DISABLED', r.rows[0].email, null);
    res.json({ success: true, user: r.rows[0] });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to update user' });
  }
});

// Admin-only password reset for any user. Regular users have no self-service reset -
// this is intentional: only Admin can set a new password for a teammate's account.
app.patch('/api/admin/users/:id/password', requireAuthApi(['ADMIN']), async (req, res) => {
  try {
    const { password } = req.body;
    if (!password || password.length < 8) return res.status(400).json({ error: 'Password must be at least 8 characters.' });
    const { salt, hash } = hashPassword(password);
    const r = await pool.query(
      'UPDATE users SET password_hash = $1, password_salt = $2 WHERE id = $3 RETURNING id, email',
      [hash, salt, req.params.id]
    );
    if (r.rows.length === 0) return res.status(404).json({ error: 'User not found' });
    // Any existing sessions for this user become stale credentials-wise but stay logged
    // in until they log out - acceptable for an internal tool; not invalidating sessions here.
    logActivity(req.user, 'PASSWORD_RESET', r.rows[0].email, null);
    res.json({ success: true, user: r.rows[0] });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to reset password' });
  }
});

app.get('/api/buyers', requireAuthApi(['ADMIN', 'BUYER', 'ACCOUNTS', 'LOGISTICS']), async (req, res) => {
  const result = await pool.query("SELECT code, name, currency, exchange_rate_to_usd FROM buyers");
  res.json(result.rows);
});

app.get('/api/vendors', requireAuthApi(['ADMIN', 'BUYER', 'ACCOUNTS', 'LOGISTICS']), async (req, res) => {
  try {
    const result = await pool.query("SELECT * FROM vendors ORDER BY vendor_code");
    res.json(result.rows);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Failed to fetch vendors" });
  }
});

app.post('/api/vendors', requireAuthApi(['ADMIN', 'BUYER']), async (req, res) => {
  try {
    const { vendor_code, vendor_name, category, city, contact, email } = req.body;
    if (!vendor_code || !vendor_name || !category) {
      return res.status(400).json({ error: "vendor_code, vendor_name, and category are required" });
    }
    const result = await pool.query(
      `INSERT INTO vendors (vendor_code, vendor_name, category, city, contact, email, is_active)
       VALUES ($1, $2, $3, $4, $5, $6, true)
       ON CONFLICT (vendor_code) DO NOTHING
       RETURNING *`,
      [vendor_code, vendor_name, category, city || null, contact || null, email || null]
    );
    if (result.rows.length === 0) return res.status(409).json({ error: "Vendor code already exists" });
    logActivity(req.user, 'VENDOR_CREATED', vendor_code, { vendor_name });
    res.json({ success: true, vendor: result.rows[0] });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Failed to create vendor" });
  }
});

app.get('/api/items', requireAuthApi(['ADMIN', 'BUYER', 'ACCOUNTS', 'LOGISTICS']), async (req, res) => {
  const { category } = req.query;
  let query = "SELECT sku, friendly_name, category_code, year_code, collection_code, department_code, department_name, color_code, material, std_cost_rmb, std_cost_rmb/7.0 as std_cost_usd, status, barcode, vendor_item_number FROM item_master WHERE status='ACTIVE'";
  let params = [];

  if (category) {
    query += " AND category_code = $1";
    params.push(category);
  }

  try {
    const result = await pool.query(query, params);
    res.json(result.rows);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Failed to fetch items" });
  }
});

app.post('/api/create-invoice', requireAuthApi(['ADMIN']), async (req, res) => {
  const { buyer_code, po_reference, sku, item_name, qty, markup_percent, std_cost_usd } = req.body;
  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    const buyerRes = await client.query("SELECT name, currency, exchange_rate_to_usd FROM buyers WHERE code = $1", [buyer_code]);
    if (buyerRes.rows.length === 0) { await client.query('ROLLBACK'); return res.status(400).json({ error: "Buyer not found" }); }
    const buyer = buyerRes.rows[0];

    const onHand = await getOnHand(client, sku, HQ_ORG_ID);
    if (onHand < qty) {
      await client.query('ROLLBACK');
      return res.status(400).json({ error: `Insufficient stock at HQ: only ${onHand} units of ${sku} available.` });
    }

    const markupMultiplier = 1 + (markup_percent / 100);
    const unit_price_usd = std_cost_usd * markupMultiplier;
    const unit_price_local = unit_price_usd * buyer.exchange_rate_to_usd;
    const total_local = unit_price_local * qty;
    const order_id = `${buyer_code.substring(0,2)}-${Date.now().toString().slice(-6)}`;

    await adjustInventory(client, sku, HQ_ORG_ID, -qty);
    await client.query(`
      INSERT INTO customer_orders
      (order_id, po_reference, buyer_code, sku, item_name, qty, markup_percent, unit_cost_usd, unit_price_usd, unit_price_local, total_local, order_currency, created_by)
      VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13)
    `, [order_id, po_reference, buyer_code, sku, item_name, qty, markup_percent, std_cost_usd, unit_price_usd, unit_price_local, total_local, buyer.currency, req.user.email]);
    await client.query('COMMIT');
    logActivity(req.user, 'INVOICE_CREATED', order_id, { buyer_code, sku, qty, total_local: total_local.toFixed(2), currency: buyer.currency });

    res.json({
      success: true,
      order_id,
      po_reference,
      buyer_code,
      buyer_name: buyer.name,
      sku,
      item_name,
      qty,
      markup_percent,
      unit_price_local: unit_price_local.toFixed(2),
      total_local: total_local.toFixed(2),
      currency: buyer.currency
    });
  } catch (err) {
    await client.query('ROLLBACK');
    console.error(err);
    res.status(500).json({ error: "Failed to create invoice" });
  } finally {
    client.release();
  }
});

app.post('/api/create-po', requireAuthApi(['ADMIN', 'BUYER']), async (req, res) => {
  const { vendor_code, po_date, po_currency, exchange_rate, items } = req.body;
  const po_id = `PO-${Date.now().toString().slice(-6)}`;
  
  try {
    const vendorRes = await pool.query("SELECT vendor_code, category FROM vendors WHERE vendor_code = $1", [vendor_code]);
    if (vendorRes.rows.length === 0) return res.status(400).json({ error: "Vendor not found" });
    const vendorCategory = vendorRes.rows[0].category;

    let yearCode = 'A', collectionCode = 'PS', departmentCode = '1', colorCode = 'SLV', material = 'Glass';
    if (vendorCategory === 'JW') { material = 'Sterling Silver'; colorCode = 'SLV'; }
    if (vendorCategory === 'AP') { material = 'Cotton'; colorCode = 'BLK'; }
    if (vendorCategory === 'HB') { material = 'Leather'; colorCode = 'BLK'; }

    for (const item of items) {
      const exists = await pool.query("SELECT sku FROM item_master WHERE sku = $1", [item.sku]);
      if (exists.rows.length === 0) {
        const newSku = `${vendorCategory}${yearCode}${collectionCode}-${departmentCode}${Date.now().toString().slice(-3)}-${colorCode}`;
        await pool.query(
          `INSERT INTO item_master (sku, friendly_name, category_code, year_code, collection_code, department_code, color_code, hs_code, std_cost_rmb, description, material, created_by)
           VALUES ($1, $2, $3, $4, $5, $6, $7, '9505100090', $8, $9, $10, $11)`,
          [newSku, `Auto-created ${newSku}`, vendorCategory, yearCode, collectionCode, departmentCode, colorCode, item.unit_price * 7.0, 'New Item', material, req.user.email]
        );
        await pool.query("INSERT INTO inventory (sku, org_id, quantity_on_hand) VALUES ($1, 1, 0)", [newSku]);
        item.sku = newSku;
      }
    }

    let total_usd = 0, total_rmb = 0;
    const exchange_rate_to_rmb = parseFloat(exchange_rate) || 7.0;

    for (const item of items) {
      if (po_currency === 'RMB') {
        const line_total_rmb = item.qty * item.unit_price;
        const line_total_usd = line_total_rmb / exchange_rate_to_rmb;
        total_rmb += line_total_rmb;
        total_usd += line_total_usd;
      } else {
        const line_total_usd = item.qty * item.unit_price;
        const line_total_rmb = line_total_usd * exchange_rate_to_rmb;
        total_usd += line_total_usd;
        total_rmb += line_total_rmb;
      }
    }

    const deposit_usd = total_usd * 0.30;
    const balance_usd = total_usd * 0.70;

    await pool.query(
      `INSERT INTO purchase_orders (po_id, vendor_code, po_date, invoice_currency, exchange_rate_to_rmb, status, total_rmb, total_usd, deposit_usd, balance_usd, created_by)
       VALUES ($1, $2, $3, $4, $5, 'ISSUED', $6, $7, $8, $9, $10)`,
      [po_id, vendor_code, po_date, po_currency, exchange_rate_to_rmb, total_rmb, total_usd, deposit_usd, balance_usd, req.user.email]
    );

    for (const item of items) {
      const costRmb = po_currency === 'RMB' ? item.unit_price : item.unit_price * exchange_rate_to_rmb;
      await pool.query(
        `INSERT INTO po_line_items (po_id, sku, quantity, unit_price_foreign, unit_cost_rmb, hs_code)
         VALUES ($1, $2, $3, $4, $5, $6)`,
        [po_id, item.sku, item.qty, item.unit_price, costRmb, '9505100090']
      );
    }

    const lineItemsRes = await pool.query(`
      SELECT li.sku, li.quantity, li.unit_price_foreign, li.unit_cost_rmb, 
             m.friendly_name as item_name, m.color_code as color, m.description, m.material, m.barcode, 
             m.vendor_item_number, m.hs_code
      FROM po_line_items li
      JOIN item_master m ON li.sku = m.sku
      WHERE li.po_id = $1
    `, [po_id]);

    const deposit_rmb = deposit_usd * exchange_rate_to_rmb;
    const balance_rmb = balance_usd * exchange_rate_to_rmb;

    logActivity(req.user, 'PO_CREATED', po_id, { vendor_code, total_rmb: total_rmb.toFixed(2), total_usd: total_usd.toFixed(2), line_count: items.length });

    res.json({
      success: true,
      po_id,
      vendor_code, 
      po_date, 
      po_currency, 
      exchange_rate_to_rmb,
      total_usd: total_usd.toFixed(2), 
      total_rmb: total_rmb.toFixed(2),
      deposit_usd: deposit_usd.toFixed(2),
      balance_usd: balance_usd.toFixed(2),
      deposit_rmb: deposit_rmb.toFixed(2),
      balance_rmb: balance_rmb.toFixed(2),
      items: lineItemsRes.rows
    });

  } catch (err) {
    console.error('PO Creation Error:', err);
    res.status(500).json({ error: "Failed to create PO" });
  }
});

app.post('/api/transfer-stock', requireAuthApi(['ADMIN']), async (req, res) => {
  const { buyer_code, sku, qty, markup_percent, cost_usd } = req.body;
  const client = await pool.connect();
  try {
    await client.query('BEGIN');

    const buyerRes = await client.query("SELECT name, currency, exchange_rate_to_usd FROM buyers WHERE code = $1", [buyer_code]);
    if (buyerRes.rows.length === 0) { await client.query('ROLLBACK'); return res.status(400).json({ error: "Buyer not found" }); }
    const buyer = buyerRes.rows[0];

    const itemRes = await client.query("SELECT friendly_name FROM item_master WHERE sku = $1", [sku]);
    const item_name = itemRes.rows.length > 0 ? itemRes.rows[0].friendly_name : sku;

    const onHand = await getOnHand(client, sku, HQ_ORG_ID);
    if (onHand < qty) {
      await client.query('ROLLBACK');
      return res.status(400).json({ error: `Insufficient stock at HQ: only ${onHand} units of ${sku} available.` });
    }

    const destOrgId = BUYER_ORG_MAP[buyer_code];

    const markupMultiplier = 1 + (markup_percent / 100);
    const unit_price_usd = cost_usd * markupMultiplier;
    const unit_price_local = unit_price_usd * buyer.exchange_rate_to_usd;
    const total_local = unit_price_local * qty;
    const order_id = `TR-${Date.now().toString().slice(-6)}`;

    // Move stock: out of HQ, into the destination site's own inventory (true multi-site ledger)
    await adjustInventory(client, sku, HQ_ORG_ID, -qty);
    if (destOrgId) {
      await adjustInventory(client, sku, destOrgId, qty);
    }

    await client.query(`
      INSERT INTO customer_orders
      (order_id, po_reference, buyer_code, sku, item_name, qty, markup_percent, unit_cost_usd, unit_price_usd, unit_price_local, total_local, order_currency, created_by)
      VALUES ($1, 'TRANSFER', $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12)
    `, [order_id, buyer_code, sku, item_name, qty, markup_percent, cost_usd, unit_price_usd, unit_price_local, total_local, buyer.currency, req.user.email]);

    await client.query('COMMIT');
    logActivity(req.user, 'STOCK_TRANSFERRED', order_id, { buyer_code, sku, qty });
    res.json({
      success: true,
      order_id,
      message: `Transferred ${qty} units.`,
      buyer_code,
      buyer_name: buyer.name,
      sku,
      item_name,
      qty,
      markup_percent,
      unit_price_local: unit_price_local.toFixed(2),
      total_local: total_local.toFixed(2),
      currency: buyer.currency
    });
  } catch (err) {
    await client.query('ROLLBACK');
    console.error(err);
    res.status(500).json({ error: "Transfer failed" });
  } finally {
    client.release();
  }
});

app.get('/api/purchase-orders', requireAuthApi(['ADMIN', 'BUYER', 'ACCOUNTS', 'LOGISTICS']), async (req, res) => {
  try {
    const result = await pool.query(`
      SELECT po_id, vendor_code, po_date, status, total_usd, total_rmb
      FROM purchase_orders ORDER BY po_date DESC, po_id DESC
    `);
    res.json(result.rows);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Failed to fetch purchase orders" });
  }
});

// Full PO detail for reprinting a PDF later (create-po.html only shows the PDF once, at
// creation time - this lets any page pull the same data back out to regenerate it).
app.get('/api/purchase-orders/:po_id/full', requireAuthApi(['ADMIN', 'BUYER', 'ACCOUNTS', 'LOGISTICS']), async (req, res) => {
  try {
    const poRes = await pool.query("SELECT * FROM purchase_orders WHERE po_id = $1", [req.params.po_id]);
    if (poRes.rows.length === 0) return res.status(404).json({ error: "PO not found" });
    const po = poRes.rows[0];

    const lineItemsRes = await pool.query(`
      SELECT li.sku, li.quantity, li.unit_price_foreign, li.unit_cost_rmb, li.received_qty,
             m.friendly_name as item_name, m.color_code as color, m.description, m.material, m.barcode,
             m.vendor_item_number, m.hs_code
      FROM po_line_items li
      JOIN item_master m ON li.sku = m.sku
      WHERE li.po_id = $1
      ORDER BY li.sku
    `, [req.params.po_id]);

    res.json({
      success: true,
      po_id: po.po_id,
      vendor_code: po.vendor_code,
      po_date: po.po_date,
      po_currency: po.invoice_currency,
      exchange_rate_to_rmb: po.exchange_rate_to_rmb,
      status: po.status,
      total_usd: parseFloat(po.total_usd).toFixed(2),
      total_rmb: parseFloat(po.total_rmb).toFixed(2),
      deposit_usd: parseFloat(po.deposit_usd).toFixed(2),
      balance_usd: parseFloat(po.balance_usd).toFixed(2),
      deposit_rmb: (parseFloat(po.deposit_usd) * parseFloat(po.exchange_rate_to_rmb)).toFixed(2),
      balance_rmb: (parseFloat(po.balance_usd) * parseFloat(po.exchange_rate_to_rmb)).toFixed(2),
      items: lineItemsRes.rows
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Failed to fetch PO detail" });
  }
});

app.get('/api/purchase-orders/:po_id/lines', requireAuthApi(['ADMIN', 'BUYER', 'ACCOUNTS', 'LOGISTICS']), async (req, res) => {
  try {
    const result = await pool.query(`
      SELECT li.sku, li.quantity, li.unit_cost_rmb, li.received_qty, m.friendly_name, m.barcode
      FROM po_line_items li
      JOIN item_master m ON li.sku = m.sku
      WHERE li.po_id = $1
      ORDER BY li.sku
    `, [req.params.po_id]);
    res.json(result.rows);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Failed to fetch PO line items" });
  }
});

app.delete('/api/purchase-orders/:po_id', requireAuthApi(['ADMIN']), async (req, res) => {
  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    const poId = req.params.po_id;
    const check = await client.query("SELECT po_id, status FROM purchase_orders WHERE po_id = $1", [poId]);
    if (check.rows.length === 0) { await client.query('ROLLBACK'); return res.status(404).json({ error: "PO not found" }); }
    const receivedCheck = await client.query("SELECT COALESCE(SUM(received_qty), 0) AS total_received FROM po_line_items WHERE po_id = $1", [poId]);
    if (parseFloat(receivedCheck.rows[0].total_received) > 0) {
      await client.query('ROLLBACK');
      return res.status(400).json({ error: "Cannot delete a PO that already has received stock against it." });
    }
    await client.query("DELETE FROM po_line_items WHERE po_id = $1", [poId]);
    await client.query("DELETE FROM purchase_orders WHERE po_id = $1", [poId]);
    await client.query('COMMIT');
    logActivity(req.user, 'PO_DELETED', poId, null);
    res.json({ success: true, po_id: poId });
  } catch (err) {
    await client.query('ROLLBACK');
    console.error(err);
    res.status(500).json({ error: "Failed to delete PO" });
  } finally {
    client.release();
  }
});

// Goods receipt: moves ordered stock from "on order" into real, sellable HQ inventory.
// Without this step, inventory.quantity_on_hand never increases and every sale/transfer
// would eventually go negative.
app.post('/api/receive-po', requireAuthApi(['ADMIN', 'BUYER', 'LOGISTICS']), async (req, res) => {
  const { po_id, lines } = req.body; // lines: [{ sku, qty_received }]
  if (!po_id || !Array.isArray(lines) || lines.length === 0) {
    return res.status(400).json({ error: "po_id and at least one line are required" });
  }
  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    const receivedLines = [];

    for (const line of lines) {
      const qty = parseFloat(line.qty_received);
      if (!qty || qty <= 0) continue;

      const lineRes = await client.query(
        "SELECT quantity, received_qty FROM po_line_items WHERE po_id = $1 AND sku = $2 FOR UPDATE",
        [po_id, line.sku]
      );
      if (lineRes.rows.length === 0) continue;

      const ordered = parseFloat(lineRes.rows[0].quantity);
      const alreadyReceived = parseFloat(lineRes.rows[0].received_qty) || 0;
      const remaining = ordered - alreadyReceived;
      const receiveQty = Math.min(qty, remaining);
      if (receiveQty <= 0) continue;

      await client.query(
        "UPDATE po_line_items SET received_qty = received_qty + $1 WHERE po_id = $2 AND sku = $3",
        [receiveQty, po_id, line.sku]
      );
      await adjustInventory(client, line.sku, HQ_ORG_ID, receiveQty);
      receivedLines.push({ sku: line.sku, received: receiveQty });
    }

    const allLines = await client.query(
      "SELECT quantity, received_qty FROM po_line_items WHERE po_id = $1",
      [po_id]
    );
    const fullyReceived = allLines.rows.length > 0 && allLines.rows.every(r => parseFloat(r.received_qty) >= parseFloat(r.quantity));
    const anyReceived = allLines.rows.some(r => parseFloat(r.received_qty) > 0);
    const newStatus = fullyReceived ? 'RECEIVED' : (anyReceived ? 'PARTIAL' : 'ISSUED');
    await client.query("UPDATE purchase_orders SET status = $1 WHERE po_id = $2", [newStatus, po_id]);

    await client.query('COMMIT');
    logActivity(req.user, 'STOCK_RECEIVED', po_id, { status: newStatus, lines: receivedLines });
    res.json({ success: true, po_id, status: newStatus, received: receivedLines });
  } catch (err) {
    await client.query('ROLLBACK');
    console.error(err);
    res.status(500).json({ error: "Failed to receive PO" });
  } finally {
    client.release();
  }
});

const upload = multer({ storage: multer.memoryStorage() });
app.post('/api/upload-invoice', requireAuthApi(['ADMIN', 'BUYER']), upload.single('invoiceFile'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "No file uploaded" });
    
    const { currency } = req.body;
    const workbook = xlsx.read(req.file.buffer);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rawRows = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: "" });

    let headerRowIdx = -1, colItemNo = -1, colFriendly = -1, colPrice = -1;
    let colDeptCode = -1, colDeptName = -1, colColor = -1, colMaterial = -1;

    for (let i = 0; i < rawRows.length; i++) {
      const row = rawRows[i];
      const itemNoIndex = row.findIndex(cell => String(cell).trim().includes('Item No.') || String(cell).trim().includes('货号') || String(cell).trim().includes('ITEM NO'));
      if (itemNoIndex !== -1) {
        headerRowIdx = i;
        colItemNo = itemNoIndex;
        colFriendly = row.findIndex(cell => String(cell).trim().includes('Friendly Name') || String(cell).trim().includes('Item Name') || String(cell).trim().includes('Name'));
        colPrice = row.findIndex(cell => String(cell).trim().includes('UNTI PRICE') || String(cell).trim().includes('UNIT PRICE') || String(cell).trim().includes('单价') || String(cell).trim().includes('Price'));
        colDeptCode = row.findIndex(cell => String(cell).trim().includes('Department Code'));
        colDeptName = row.findIndex(cell => String(cell).trim().includes('Department Name'));
        colColor = row.findIndex(cell => String(cell).trim().includes('Color') && !String(cell).trim().includes('Department'));
        colMaterial = row.findIndex(cell => String(cell).trim().includes('Material'));

        if (colFriendly === -1 || colDeptCode === -1 || colColor === -1) {
          for (let j = i+1; j < Math.min(i+4, rawRows.length); j++) {
            const nextRow = rawRows[j];
            if (colFriendly === -1) colFriendly = nextRow.findIndex(cell => String(cell).trim().includes('Friendly Name') || String(cell).trim().includes('Item Name') || String(cell).trim().includes('Name'));
            if (colDeptCode === -1) colDeptCode = nextRow.findIndex(cell => String(cell).trim().includes('Department Code'));
            if (colDeptName === -1) colDeptName = nextRow.findIndex(cell => String(cell).trim().includes('Department Name'));
            if (colColor === -1) colColor = nextRow.findIndex(cell => String(cell).trim().includes('Color'));
            if (colMaterial === -1) colMaterial = nextRow.findIndex(cell => String(cell).trim().includes('Material'));
          }
        }
        break;
      }
    }

    if (headerRowIdx === -1 || colItemNo === -1 || colPrice === -1) {
      return res.status(400).json({ error: "Could not find required columns (Item No., Unit Price)." });
    }

    const items = [];
    for (let i = headerRowIdx + 1; i < rawRows.length; i++) {
      const row = rawRows[i];
      const itemNo = row[colItemNo] ? String(row[colItemNo]).trim() : '';
      const friendlyName = (colFriendly !== -1 && row[colFriendly]) ? String(row[colFriendly]).trim() : itemNo;
      const deptCode = (colDeptCode !== -1 && row[colDeptCode]) ? String(row[colDeptCode]).trim() : '01';
      const deptName = (colDeptName !== -1 && row[colDeptName]) ? String(row[colDeptName]).trim() : 'Default';
      const colorCode = (colColor !== -1 && row[colColor]) ? String(row[colColor]).trim() : 'SLV';
      const material = (colMaterial !== -1 && row[colMaterial]) ? String(row[colMaterial]).trim() : 'GL';
      
      let priceStr = row[colPrice] ? String(row[colPrice]) : '';
      priceStr = priceStr.replace(/[^0-9.]/g, '');
      const price = parseFloat(priceStr) || 0;

      if (!itemNo || price === 0) continue;

      items.push({
        vendorItem: itemNo,
        sku: `NEW-${itemNo}`,
        friendlyName: friendlyName,
        deptCode: deptCode,
        deptName: deptName,
        colorCode: colorCode,
        material: material,
        unit_price: price,
        currency: currency,
        qty: 96
      });
    }
    res.json({ success: true, items });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Failed to parse Excel file" });
  }
});

// --- FIXED SAVE MAPPINGS (STOPS MULTIPLYING PRICE BY 7.0) ---
app.post('/api/save-mappings', requireAuthApi(['ADMIN', 'BUYER']), async (req, res) => {
  const { mappings } = req.body;
  if (!mappings || !Array.isArray(mappings)) return res.status(400).json({ error: "Invalid mappings data" });

  const results = [];
  const errors = [];
  const client = await pool.connect();

  try {
    await client.query('BEGIN');

    for (let mapIdx = 0; mapIdx < mappings.length; mapIdx++) {
      const item = mappings[mapIdx];
      const { vendorItem, selectedSku, price, qty, catCode, yearCode, colCode, colorCode, friendlyName, material } = item;
      // Defensive fallback - a missing department code must never end up as the literal
      // string "undefined" inside a SKU (this happened when the frontend didn't pass it through).
      const deptCode = (item.deptCode || '01').toString().trim() || '01';
      const deptName = item.deptName || 'Default';

      if (selectedSku.startsWith('NEW-')) {
        // Offset by loop index so multiple items saved in the same request (same millisecond)
        // never collide on the same 3-digit sequence.
        const sequence = ((Date.now() + mapIdx) % 1000).toString().padStart(3, '0');
        const newSku = `${catCode}${yearCode}${colCode}-${deptCode}${sequence}-${colorCode}`;
        
        try {
          const barcodeRes = await client.query("SELECT next_barcode FROM barcode_sequence WHERE id = 1 FOR UPDATE");
          let currentBarcode = barcodeRes.rows[0]?.next_barcode;
          if (!currentBarcode) {
            currentBarcode = '6941181218000';
            await client.query("INSERT INTO barcode_sequence (id, next_barcode) VALUES (1, $1) ON CONFLICT (id) DO NOTHING", [String(Number(currentBarcode) + 1)]);
          }

          // CORRECTED: Saving 'price' as is (RMB). Removed 'price * 7.0'.
          await client.query(
            `INSERT INTO item_master (sku, friendly_name, category_code, year_code, collection_code, department_code, department_name, color_code, hs_code, std_cost_rmb, description, material, barcode, vendor_item_number, created_by)
             VALUES ($1, $2, $3, $4, $5, $6, $7, $8, '9505100090', $9, $10, $11, $12, $13, $14)
             ON CONFLICT (sku) DO NOTHING`,
            [newSku, friendlyName || `Imported ${vendorItem}`, catCode, yearCode, colCode, deptCode, deptName, colorCode, price, `Imported from vendor ${vendorItem}`, material, currentBarcode, vendorItem, req.user.email]
          );

          await client.query("UPDATE barcode_sequence SET next_barcode = LPAD((CAST(next_barcode AS BIGINT) + 1)::TEXT, 13, '0') WHERE id = 1");
          await client.query("INSERT INTO inventory (sku, org_id, quantity_on_hand) VALUES ($1, 1, 0) ON CONFLICT DO NOTHING", [newSku]);
          results.push({ vendorItem, sku: newSku, barcode: currentBarcode, status: 'created' });
        } catch (err) { errors.push({ vendorItem, error: err.message }); await client.query('ROLLBACK'); throw err; }
      } else {
        results.push({ vendorItem, sku: selectedSku, status: 'mapped' });
      }
    }
    await client.query('COMMIT');
  } catch (err) { return res.status(500).json({ error: "Failed to save mappings", details: err.message }); }
  finally { client.release(); }
  const createdCount = results.filter(r => r.status === 'created').length;
  if (createdCount > 0) {
    logActivity(req.user, 'ITEMS_CREATED', null, { count: createdCount, skus: results.filter(r => r.status === 'created').map(r => r.sku) });
  }
  res.json({ success: true, results, errors });
});

// Utility: patch vendor_item_number on an existing item_master row (used to backfill
// rows created before the save-mappings vendor_item_number bug was fixed).
app.patch('/api/items/:sku/vendor-item-number', requireAuthApi(['ADMIN', 'BUYER']), async (req, res) => {
  try {
    const { vendor_item_number } = req.body;
    if (!vendor_item_number) return res.status(400).json({ error: "vendor_item_number is required" });
    const result = await pool.query(
      "UPDATE item_master SET vendor_item_number = $1 WHERE sku = $2 RETURNING sku, vendor_item_number",
      [vendor_item_number, req.params.sku]
    );
    if (result.rows.length === 0) return res.status(404).json({ error: "SKU not found" });
    res.json({ success: true, ...result.rows[0] });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Failed to patch vendor_item_number" });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`🚀 Portal running on port ${PORT}`));