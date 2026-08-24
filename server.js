require('dotenv').config();
const express = require('express');
const cors = require('cors');
const path = require('path');
const crypto = require('crypto');
const { Pool } = require('pg');
const multer = require('multer');
const xlsx = require('xlsx');
const fs = require('fs');
const nodemailer = require('nodemailer');
const PDFDocument = require('pdfkit');

// ── Email config (for B2B order notifications) ──
const EMAIL_HOST = process.env.EMAIL_HOST || 'smtp.gmail.com';
const EMAIL_PORT = parseInt(process.env.EMAIL_PORT || '587');
const EMAIL_USER = process.env.EMAIL_USER || 'B2B-order@sahilondon.com';
const EMAIL_PASS = process.env.EMAIL_PASS || '';
const ORDER_EMAIL = process.env.ORDER_EMAIL || 'B2B-order@sahilondon.com';

console.log(`[EMAIL CONFIG] user=${EMAIL_USER}, host=${EMAIL_HOST}:${EMAIL_PORT}, pass=${EMAIL_PASS ? 'SET(' + EMAIL_PASS.length + 'chars)' : 'MISSING'}, order_to=${ORDER_EMAIL}`);

const mailer = nodemailer.createTransport({
  host: EMAIL_HOST,
  port: EMAIL_PORT,
  secure: EMAIL_PORT === 465,
  auth: { user: EMAIL_USER, pass: EMAIL_PASS }
});

// ── PDF Generator for B2B Order Confirmations ──
function generateOrderPDF(order) {
  return new Promise((resolve, reject) => {
    try {
      const doc = new PDFDocument({ margin: 40, size: 'A4' });
      const chunks = [];
      doc.on('data', c => chunks.push(c));
      doc.on('end', () => resolve(Buffer.concat(chunks)));
      doc.on('error', reject);

      // Colors — SAHI palette
      const SKY_BLUE = '#87CEEB';
      const PASTEL_PINK = '#FFB7C5';
      const LIGHT_GREY = '#F2F2F2';
      const WHITE = '#ffffff';
      const BLACK = '#111111';
      const GREY = '#666666';

      const imgDir = path.join(__dirname, 'public', 'images', 'b2b');
      const publicImgDir = path.join(__dirname, 'public', 'images');

      // Extract base SKU (strip variant suffix like -OWT-1 or -BLU)
      function baseSku(sku) {
        if (!sku) return '';
        const m = sku.match(/^([A-Z]+[0-9]+)/);
        return m ? m[1] : sku;
      }

      // Find local image file for a SKU
      function findImage(sku) {
        const b = baseSku(sku);
        const candidates = [
          path.join(imgDir, `${b}.jpg`),
          path.join(imgDir, `${b}.png`),
          path.join(publicImgDir, `${b}_model.jpg`),
        ];
        for (const c of candidates) {
          if (fs.existsSync(c)) return c;
        }
        return null;
      }

      // ── HEADER BAR ──
      doc.rect(0, 0, 595, 100).fill(SKY_BLUE);
      doc.fillColor(BLACK).fontSize(28).font('Helvetica-Bold')
        .text('SAHI LONDON', 40, 18);
      doc.fontSize(8.5).font('Helvetica')
        .text('1231, Niagara On The Lake', 40, 48)
        .text('Ontario L0S 1J0, Canada', 40, 61)
        .text('HST #732146907', 40, 74);
      doc.fontSize(9).font('Helvetica-Bold').fillColor(BLACK)
        .text('ORDER', 350, 22, { width: 205, align: 'right' });
      doc.fontSize(20).font('Helvetica-Bold').fillColor(BLACK)
        .text(`#${order.id}`, 350, 35, { width: 205, align: 'right' });
      const orderDate = order.created_at
        ? new Date(order.created_at).toLocaleDateString('en-CA', { year: 'numeric', month: 'long', day: 'numeric' })
        : new Date().toLocaleDateString('en-CA', { year: 'numeric', month: 'long', day: 'numeric' });
      doc.fontSize(9).font('Helvetica-Bold').fillColor(BLACK)
        .text('DATE', 350, 64, { width: 205, align: 'right' });
      doc.fontSize(10).font('Helvetica').fillColor(BLACK)
        .text(orderDate, 350, 78, { width: 205, align: 'right' });

      // ── ADDRESS SECTION ──
      const addrY = 112;
      const addrW = 255;
      const addrH = 86;

      // BILL TO — Customer billing details
      doc.roundedRect(40, addrY, addrW, addrH, 5).fill(LIGHT_GREY).stroke('#cccccc');
      const buyerName = order.company_name || 'LARKIN GREENHOUSES INC';
      const buyerContact = order.contact_name || 'Manon st Pierre';
      doc.fillColor(BLACK).fontSize(10).font('Helvetica-Bold')
        .text('BILL TO', 52, addrY + 8);
      doc.fontSize(9).font('Helvetica')
        .text(buyerName, 52, addrY + 24)
        .text(`Contact: ${buyerContact}`, 52, addrY + 38)
        .text(order.email || '', 52, addrY + 52)
        .text(order.phone || '', 52, addrY + 66);

      // SHIP TO — Delivery address
      doc.roundedRect(315, addrY, addrW, addrH, 5).fill(LIGHT_GREY).stroke('#cccccc');
      const shipTo = order.shipping_address || '';
      doc.fillColor(BLACK).fontSize(10).font('Helvetica-Bold')
        .text('SHIP TO', 327, addrY + 8);
      doc.fontSize(9).font('Helvetica')
        .text(shipTo || 'Shipping address not provided', 327, addrY + 24, { width: addrW - 24, lineGap: 2 });

      // ── ITEMS TABLE ──
      const tableY = addrY + 116;
      const cols = [40, 80, 155, 360, 415, 485];   // x positions
      //                                Img   SKU   Product         Qty  UnitPr  Total
      const colHeaders = ['', 'SKU', 'PRODUCT', 'QTY', 'UNIT PRICE', 'TOTAL'];
      const colAligns  = ['center', 'left', 'left', 'center', 'right', 'right'];

      // Table header
      doc.roundedRect(40, tableY, 515, 20, 4).fill(SKY_BLUE);
      doc.fillColor(BLACK).fontSize(8).font('Helvetica-Bold');
      colHeaders.forEach((h, i) => {
        const x = cols[i];
        const w = i < 4 ? (cols[i + 1] - cols[i]) : (cols[i + 1] ? cols[i + 1] - cols[i] : 70);
        doc.text(h, x, tableY + 5, { width: w, align: colAligns[i] });
      });

      // Table rows — 10 items per page, header repeats on each page
      let rowY = tableY + 22;
      const items = order.items || [];
      const rowH = 28;
      const ITEMS_PER_PAGE = 10;
      doc.font('Helvetica').fontSize(7.5);

      // Helper: draw header bar + table header on a page
      function drawTableHeader(pageTopY) {
        // Header bar
        doc.rect(0, 0, 595, 100).fill(SKY_BLUE);
        doc.fillColor(BLACK).fontSize(28).font('Helvetica-Bold')
          .text('SAHI LONDON', 40, 18);
        doc.fontSize(8.5).font('Helvetica')
          .text('1231, Niagara On The Lake', 40, 48)
          .text('Ontario L0S 1J0, Canada', 40, 61)
          .text('HST #732146907', 40, 74);
        doc.fontSize(9).font('Helvetica-Bold').fillColor(BLACK)
          .text('ORDER', 350, 22, { width: 205, align: 'right' });
        doc.fontSize(20).font('Helvetica-Bold').fillColor(BLACK)
          .text(`#${order.id}`, 350, 35, { width: 205, align: 'right' });
        doc.fontSize(9).font('Helvetica-Bold').fillColor(BLACK)
          .text('DATE', 350, 64, { width: 205, align: 'right' });
        doc.fontSize(10).font('Helvetica').fillColor(BLACK)
          .text(orderDate, 350, 78, { width: 205, align: 'right' });

        // Table header
        doc.roundedRect(40, pageTopY, 515, 20, 4).fill(SKY_BLUE);
        doc.fillColor(BLACK).fontSize(8).font('Helvetica-Bold');
        colHeaders.forEach((h, i) => {
          const x = cols[i];
          const w = i < 4 ? (cols[i + 1] - cols[i]) : (cols[i + 1] ? cols[i + 1] - cols[i] : 70);
          doc.text(h, x, pageTopY + 5, { width: w, align: colAligns[i] });
        });
      }

      for (let i = 0; i < items.length; i++) {
        // Page break after every ITEMS_PER_PAGE rows
        if (i > 0 && i % ITEMS_PER_PAGE === 0) {
          doc.addPage();
          // Redraw header on new page
          const newTableY = 112;
          drawTableHeader(newTableY);
          rowY = newTableY + 22;
          doc.font('Helvetica').fontSize(7.5);
        }

        const it = items[i];
        const bg = i % 2 === 0 ? WHITE : '#F7F9FB';
        doc.roundedRect(40, rowY, 515, rowH, 0).fill(bg);

        // Product image thumbnail (leftmost column)
        const imgPath = findImage(it.sku);
        if (imgPath) {
          try {
            doc.image(imgPath, cols[0] + 4, rowY + 3, { width: 22, height: 22 });
          } catch (_) { /* ignore broken images */ }
        }

        // SKU, Product, Qty, Unit Price, Total
        doc.fillColor(BLACK);
        const texts = [
          { v: it.sku || '', x: cols[1], w: 75, a: 'left' },
          { v: it.product_name || '', x: cols[2], w: 200, a: 'left' },
          { v: String(it.quantity || 0), x: cols[3], w: 45, a: 'center' },
          { v: `$${Number(it.unit_price || 0).toFixed(2)}`, x: cols[4], w: 65, a: 'right' },
          { v: `$${Number(it.total || 0).toFixed(2)}`, x: cols[5], w: 70, a: 'right' },
        ];
        texts.forEach(t => {
          doc.text(t.v, t.x, rowY + 6, { width: t.w, align: t.a });
        });

        rowY += rowH;
      }

      // ── TOTALS ──
      // Keep totals on the last item page so Order #11 stays 2 pages, not 3.
      const totalBarY = rowY + 10;
      const hstAmt = order.hst_amount ? Number(order.hst_amount) : 0;
      const grandTotal = order.grand_total ? Number(order.grand_total) : Number(order.total_amount || 0);

      // Subtotal
      doc.fillColor(BLACK).fontSize(10).font('Helvetica')
        .text(`Subtotal:`, 300, totalBarY, { width: 120, align: 'right' });
      doc.fillColor(BLACK).fontSize(10).font('Helvetica')
        .text(`CAD $${Number(order.total_amount || 0).toFixed(2)}`, 425, totalBarY, { width: 130, align: 'right' });

      // HST
      if (hstAmt > 0) {
        doc.fillColor(BLACK).fontSize(10).font('Helvetica')
          .text(`HST (13%, Ontario):`, 300, totalBarY + 16, { width: 120, align: 'right' });
        doc.fillColor(BLACK).fontSize(10).font('Helvetica')
          .text(`CAD $${hstAmt.toFixed(2)}`, 425, totalBarY + 16, { width: 130, align: 'right' });
      }

      // Grand Total
      doc.roundedRect(300, totalBarY + (hstAmt > 0 ? 36 : 20), 255, 24, 6).fill(LIGHT_GREY).stroke('#cccccc');
      doc.fillColor(BLACK).fontSize(15).font('Helvetica-Bold')
        .text(`GRAND TOTAL: CAD $${grandTotal.toFixed(2)}`, 310, totalBarY + (hstAmt > 0 ? 40 : 24), { width: 235, align: 'right' });

      // ── NOTES ──
      if (order.notes) {
        const notesY = totalBarY + (hstAmt > 0 ? 76 : 54);
        doc.fillColor(BLACK).fontSize(8).font('Helvetica')
          .text('NOTES:', 40, notesY);
        doc.fillColor(BLACK)
          .text(order.notes, 40, notesY + 12, { width: 515 });
      }

      // ── FOOTER ──
      doc.fillColor(BLACK).fontSize(7.5).font('Helvetica')
        .text(`SAHI London  •  B2B Order #${order.id}  •  ${new Date().toISOString().split('T')[0]}`, 40, 800, { width: 515, align: 'center' });

      doc.end();
    } catch (e) {
      reject(e);
    }
  });
}

// Ensure public/images directory exists (needed for image uploads on fresh deployments)
fs.mkdirSync(path.join(__dirname, 'public', 'images'), { recursive: true });

// Optional: Tesseract.js for OCR of image-based invoices. If not installed,
// image upload returns an error asking the user to upload Excel/CSV instead.
let Tesseract = null;
try { Tesseract = require('tesseract.js'); } catch (e) { console.log('Tesseract.js not installed — image OCR unavailable, Excel/CSV only'); }

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

  // CBM (cubic meters) per unit — needed on PO PDF for shipping/logistics.
  try {
    await pool.query(`ALTER TABLE item_master ADD COLUMN IF NOT EXISTS cbm NUMERIC DEFAULT 0`);
  } catch (err) {
    console.error('cbm column migration warning:', err.message);
  }

  // Invoice reference on PO — every PO is created against a vendor invoice.
  try {
    await pool.query(`ALTER TABLE purchase_orders ADD COLUMN IF NOT EXISTS invoice_reference TEXT`);
    await pool.query(`ALTER TABLE purchase_orders ADD COLUMN IF NOT EXISTS invoice_id INTEGER REFERENCES vendor_invoices(id) ON DELETE SET NULL`);
  } catch (err) {
    console.error('invoice reference columns migration warning:', err.message);
  }

  // --- Phase 1: Collaboration workflow schema additions ---

  // Optimistic locking: version column on editable tables.
  // Every update must WHERE-clause on the current version and increment it.
  // If the row was changed by someone else, the UPDATE affects 0 rows and the
  // server returns a 409 Conflict.
  try {
    await pool.query(`ALTER TABLE item_master ADD COLUMN IF NOT EXISTS version INTEGER NOT NULL DEFAULT 0`);
    await pool.query(`ALTER TABLE purchase_orders ADD COLUMN IF NOT EXISTS version INTEGER NOT NULL DEFAULT 0`);
  } catch (err) {
    console.error('version column migration warning:', err.message);
  }

  // PO approval workflow: new statuses beyond ISSUED/PARTIAL/RECEIVED.
  // DRAFT = creator is still editing (only creator can see/edit)
  // SUBMITTED = sent for approval (locked, approver acts next)
  // APPROVED = manager approved, ready to dispatch to supplier
  // REJECTED = sent back to creator for changes
  // ISSUED = dispatched to supplier (existing behaviour, now requires APPROVED first)
  try {
    await pool.query(`ALTER TABLE purchase_orders ADD COLUMN IF NOT EXISTS approved_by TEXT`);
    await pool.query(`ALTER TABLE purchase_orders ADD COLUMN IF NOT EXISTS approved_at TIMESTAMP`);
    await pool.query(`ALTER TABLE purchase_orders ADD COLUMN IF NOT EXISTS submitted_at TIMESTAMP`);
    await pool.query(`ALTER TABLE purchase_orders ADD COLUMN IF NOT EXISTS rejection_reason TEXT`);
  } catch (err) {
    console.error('PO approval columns migration warning:', err.message);
  }

  // Pessimistic record locking: when a user opens an item for editing, a row is
  // inserted here. Other users see "locked by X" and cannot edit until the lock
  // is released or auto-expires (15-minute TTL).
  try {
    await pool.query(`
      CREATE TABLE IF NOT EXISTS record_locks (
        id SERIAL PRIMARY KEY,
        table_name TEXT NOT NULL,
        record_id TEXT NOT NULL,
        locked_by_email TEXT NOT NULL,
        locked_by_name TEXT,
        locked_at TIMESTAMP NOT NULL DEFAULT NOW(),
        expires_at TIMESTAMP NOT NULL,
        UNIQUE (table_name, record_id)
      )
    `);
  } catch (err) {
    console.error('record_locks migration warning:', err.message);
  }

  // In-app notifications: created when a PO changes hands between roles.
  try {
    await pool.query(`
      CREATE TABLE IF NOT EXISTS notifications (
        id SERIAL PRIMARY KEY,
        target_email TEXT NOT NULL,
        target_role TEXT,
        type TEXT NOT NULL,
        title TEXT NOT NULL,
        body TEXT,
        reference_id TEXT,
        is_read BOOLEAN NOT NULL DEFAULT false,
        created_at TIMESTAMP NOT NULL DEFAULT NOW()
      )
    `);
    await pool.query(`CREATE INDEX IF NOT EXISTS idx_notifications_unread ON notifications (target_email, is_read) WHERE is_read = false`);
  } catch (err) {
    console.error('notifications migration warning:', err.message);
  }

  // --- Phase 2: Auto-PO agent schema ---

  // Vendor invoices: stores uploaded invoice metadata. Each invoice goes through
  // a lifecycle: UPLOADED → PROCESSING → MATCHED → REVIEWED → PO_CREATED.
  try {
    await pool.query(`
      CREATE TABLE IF NOT EXISTS vendor_invoices (
        id SERIAL PRIMARY KEY,
        invoice_number TEXT,
        vendor_code TEXT,
        invoice_date DATE,
        file_name TEXT NOT NULL,
        file_type TEXT NOT NULL,
        status TEXT NOT NULL DEFAULT 'UPLOADED' CHECK (status IN ('UPLOADED', 'PROCESSING', 'MATCHED', 'REVIEWED', 'PO_CREATED', 'ERROR')),
        total_amount NUMERIC,
        currency TEXT DEFAULT 'RMB',
        po_id TEXT,
        error_message TEXT,
        created_by TEXT,
        created_at TIMESTAMP NOT NULL DEFAULT NOW()
      )
    `);
  } catch (err) {
    console.error('vendor_invoices migration warning:', err.message);
  }

  // Invoice line items: each row is one line from the parsed invoice, with its
  // match result against the Item Master.
  try {
    await pool.query(`
      CREATE TABLE IF NOT EXISTS invoice_line_items (
        id SERIAL PRIMARY KEY,
        invoice_id INTEGER NOT NULL REFERENCES vendor_invoices(id) ON DELETE CASCADE,
        seq INTEGER NOT NULL DEFAULT 0,
        vendor_item_number TEXT,
        description TEXT,
        quantity NUMERIC NOT NULL DEFAULT 1,
        unit_price NUMERIC NOT NULL DEFAULT 0,
        line_total NUMERIC NOT NULL DEFAULT 0,
        matched_sku TEXT,
        match_confidence INTEGER NOT NULL DEFAULT 0,
        match_status TEXT NOT NULL DEFAULT 'REVIEW_NEEDED' CHECK (match_status IN ('AUTO_MATCHED', 'REVIEW_NEEDED', 'CONFIRMED', 'REJECTED', 'NEW_ITEM')),
        match_method TEXT,
        created_at TIMESTAMP NOT NULL DEFAULT NOW()
      )
    `);
    await pool.query(`CREATE INDEX IF NOT EXISTS idx_invoice_items_invoice ON invoice_line_items (invoice_id)`);
  } catch (err) {
    console.error('invoice_line_items migration warning:', err.message);
  }

  // Vendors table — used by PO creation, invoice upload, and item master.
  // Must exist before any vendor-related endpoint is called.
  try {
    await pool.query(`
      CREATE TABLE IF NOT EXISTS vendors (
        vendor_code TEXT PRIMARY KEY,
        vendor_name TEXT NOT NULL,
        category TEXT,
        city TEXT,
        contact TEXT,
        email TEXT,
        is_active BOOLEAN NOT NULL DEFAULT true,
        created_at TIMESTAMP NOT NULL DEFAULT NOW()
      )
    `);
    // Seed a few known vendors if the table is empty.
    const vCount = await pool.query('SELECT COUNT(*) as cnt FROM vendors');
    if (parseInt(vCount.rows[0].cnt) === 0) {
      const seedVendors = [
        ['VS001', 'Vendor Sample 1', 'HD', 'Shanghai', '', ''],
        ['VS002', 'Vendor Sample 2', 'JW', 'Yiwu', '', ''],
        ['VS003', 'Vendor Sample 3', 'AP', 'Guangzhou', '', '']
      ];
      for (const [code, name, cat, city, contact, email] of seedVendors) {
        await pool.query(
          `INSERT INTO vendors (vendor_code, vendor_name, category, city, contact, email, is_active)
           VALUES ($1, $2, $3, $4, $5, $6, true) ON CONFLICT (vendor_code) DO NOTHING`,
          [code, name, cat, city, contact, email]
        );
      }
      console.log('Seeded vendors table with sample data');
    }
  } catch (err) {
    console.error('vendors table migration warning:', err.message);
  }

  // Catalogue visitor tracking — WhatsApp + Company for B2B wholesale catalogue access
  try {
    await pool.query(`
      CREATE TABLE IF NOT EXISTS catalogue_visitors (
        id SERIAL PRIMARY KEY,
        whatsapp TEXT NOT NULL,
        company TEXT NOT NULL,
        created_at TIMESTAMP NOT NULL DEFAULT NOW(),
        visit_count INTEGER NOT NULL DEFAULT 1
      )
    `);
  } catch (err) {
    console.error('catalogue_visitors migration warning:', err.message);
  }

  // B2B customer accounts — self-registration with company details (HST/GST etc.)
  try {
    await pool.query(`
      CREATE TABLE IF NOT EXISTS b2b_customers (
        id SERIAL PRIMARY KEY,
        user_id INTEGER REFERENCES users(id) ON DELETE CASCADE UNIQUE,
        company_name TEXT NOT NULL,
        contact_name TEXT NOT NULL,
        phone TEXT NOT NULL,
        address TEXT,
        city TEXT,
        province TEXT,
        postal_code TEXT,
        hst_gst_number TEXT,
        created_at TIMESTAMP NOT NULL DEFAULT NOW()
      )
    `);
  } catch (err) {
    console.error('b2b_customers migration warning:', err.message);
  }

  // B2B orders
  try {
    await pool.query(`
      CREATE TABLE IF NOT EXISTS b2b_orders (
        id SERIAL PRIMARY KEY,
        user_id INTEGER REFERENCES users(id) ON DELETE SET NULL,
        company_name TEXT NOT NULL,
        contact_name TEXT NOT NULL,
        email TEXT NOT NULL,
        phone TEXT,
        hst_gst_number TEXT,
        shipping_address TEXT,
        notes TEXT,
        total_amount NUMERIC NOT NULL DEFAULT 0,
        item_count INTEGER NOT NULL DEFAULT 0,
        status TEXT NOT NULL DEFAULT 'NEW' CHECK (status IN ('DRAFT', 'NEW', 'PROCESSING', 'SHIPPED', 'COMPLETED', 'CANCELLED', 'DELETED')),
        created_at TIMESTAMP NOT NULL DEFAULT NOW()
      )
    `);
  } catch (err) {
    console.error('b2b_orders migration warning:', err.message);
  }

  // HST and grand total columns for B2B orders (added after initial table creation)
  try {
    await pool.query(`ALTER TABLE b2b_orders ADD COLUMN IF NOT EXISTS hst_amount NUMERIC NOT NULL DEFAULT 0`);
    await pool.query(`ALTER TABLE b2b_orders ADD COLUMN IF NOT EXISTS grand_total NUMERIC NOT NULL DEFAULT 0`);
  } catch (err) {
    console.error('b2b_orders hst/grand_total migration warning:', err.message);
  }

  // Soft-delete column for B2B orders
  try {
    await pool.query(`ALTER TABLE b2b_orders ADD COLUMN IF NOT EXISTS deleted BOOLEAN NOT NULL DEFAULT FALSE`);
  } catch (err) {
    console.error('b2b_orders deleted column migration warning:', err.message);
  }

  // Update CHECK constraint to include DRAFT and DELETED statuses (for existing tables)
  try {
    await pool.query(`
      DO $$
      BEGIN
        IF EXISTS (SELECT 1 FROM pg_constraint WHERE conname = 'b2b_orders_status_check') THEN
          ALTER TABLE b2b_orders DROP CONSTRAINT b2b_orders_status_check;
        END IF;
        ALTER TABLE b2b_orders ADD CONSTRAINT b2b_orders_status_check
          CHECK (status IN ('DRAFT', 'NEW', 'PROCESSING', 'SHIPPED', 'COMPLETED', 'CANCELLED', 'DELETED'));
      END $$;
    `);
  } catch (err) {
    console.error('b2b_orders status constraint update warning:', err.message);
  }

  // B2B order line items
  try {
    await pool.query(`
      CREATE TABLE IF NOT EXISTS b2b_order_items (
        id SERIAL PRIMARY KEY,
        order_id INTEGER REFERENCES b2b_orders(id) ON DELETE CASCADE,
        sku TEXT NOT NULL,
        product_name TEXT NOT NULL,
        quantity INTEGER NOT NULL,
        unit_price NUMERIC NOT NULL DEFAULT 0,
        total NUMERIC NOT NULL DEFAULT 0
      )
    `);
  } catch (err) {
    console.error('b2b_order_items migration warning:', err.message);
  }

  // B2B customers table
  try {
    await pool.query(`
      CREATE TABLE IF NOT EXISTS b2b_customers (
        id SERIAL PRIMARY KEY,
        company_name TEXT NOT NULL,
        contact_name TEXT NOT NULL,
        email TEXT DEFAULT '',
        phone TEXT DEFAULT '',
        hst_gst_number TEXT DEFAULT '',
        shipping_address TEXT DEFAULT '',
        notes TEXT,
        card_image TEXT DEFAULT '',
        created_at TIMESTAMPTZ DEFAULT NOW()
      )
    `);
  } catch (err) {
    console.error('b2b_customers migration warning:', err.message);
  }
  // Add missing columns if table already existed with older schema
  const customerColumns = [
    'company_name', 'contact_name', 'email', 'phone',
    'hst_gst_number', 'shipping_address', 'notes', 'card_image'
  ];
  for (const col of customerColumns) {
    try {
      await pool.query(`ALTER TABLE b2b_customers ADD COLUMN IF NOT EXISTS ${col} TEXT ${col === 'company_name' || col === 'contact_name' ? "NOT NULL DEFAULT ''" : "DEFAULT ''"}`);
    } catch (err) {
      // Column may already exist with NOT NULL, ignore
    }
  }

  // Widen role constraint to include B2B_CUSTOMER
  try {
    await pool.query(`ALTER TABLE users DROP CONSTRAINT IF EXISTS users_role_check`);
    await pool.query(`ALTER TABLE users ADD CONSTRAINT users_role_check CHECK (role IN ('ADMIN', 'BUYER', 'ACCOUNTS', 'LOGISTICS', 'B2B_CUSTOMER'))`);
  } catch (err) {
    console.error('B2B role migration warning:', err.message);
  }

  // Opening inventory table — records migrated stock counts
  try {
    await pool.query(`
      CREATE TABLE IF NOT EXISTS opening_inventory (
        id SERIAL PRIMARY KEY,
        sku TEXT NOT NULL,
        barcode TEXT,
        friendly_name TEXT,
        color TEXT,
        qty INTEGER NOT NULL,
        org_id INTEGER NOT NULL DEFAULT 1,
        created_by TEXT,
        created_at TIMESTAMP NOT NULL DEFAULT NOW()
      )
    `);
    await pool.query(`CREATE INDEX IF NOT EXISTS idx_oi_sku ON opening_inventory(sku)`);
    await pool.query(`CREATE INDEX IF NOT EXISTS idx_oi_barcode ON opening_inventory(barcode)`);
    await pool.query(`CREATE INDEX IF NOT EXISTS idx_oi_org ON opening_inventory(org_id)`);
  } catch (err) {
    console.error('opening_inventory migration warning:', err.message);
  }

  // Add columns to item_master for bulk import support
  try {
    await pool.query(`ALTER TABLE item_master ADD COLUMN IF NOT EXISTS collection TEXT`);
    await pool.query(`ALTER TABLE item_master ADD COLUMN IF NOT EXISTS category TEXT`);
    await pool.query(`ALTER TABLE item_master ADD COLUMN IF NOT EXISTS original_qty INTEGER DEFAULT 0`);
    await pool.query(`ALTER TABLE item_master ADD COLUMN IF NOT EXISTS balance_qty INTEGER DEFAULT 0`);
  } catch (err) {
    console.error('item_master bulk columns migration warning:', err.message);
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
    if (!user) {
      const originalUrl = req.originalUrl || req.url;
      return res.redirect('/login.html?redirect=' + encodeURIComponent(originalUrl));
    }
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

// --- Record locking helpers (pessimistic lock for concurrent edit protection) ---
const LOCK_TTL_MINUTES = 15;

// Try to acquire a lock on a record. Returns true on success, or the current
// holder's info if the lock is held by someone else (and not expired).
async function acquireLock(tableName, recordId, user) {
  // First, purge expired locks for this record.
  await pool.query(`DELETE FROM record_locks WHERE table_name = $1 AND record_id = $2 AND expires_at < NOW()`, [tableName, recordId]);

  try {
    await pool.query(
      `INSERT INTO record_locks (table_name, record_id, locked_by_email, locked_by_name, expires_at)
       VALUES ($1, $2, $3, $4, NOW() + INTERVAL '${LOCK_TTL_MINUTES} minutes')`,
      [tableName, recordId, user.email, user.name || user.email]
    );
    return { success: true };
  } catch (err) {
    if (err.code === '23505') {
      // Unique constraint violation - lock already held.
      const existing = await pool.query(
        `SELECT locked_by_email, locked_by_name, expires_at FROM record_locks WHERE table_name = $1 AND record_id = $2`,
        [tableName, recordId]
      );
      return { success: false, lockedBy: existing.rows[0] };
    }
    throw err;
  }
}

// Release a lock - only the holder (or an admin) can release.
async function releaseLock(tableName, recordId, user) {
  if (user.role === 'ADMIN') {
    await pool.query(`DELETE FROM record_locks WHERE table_name = $1 AND record_id = $2`, [tableName, recordId]);
  } else {
    await pool.query(`DELETE FROM record_locks WHERE table_name = $1 AND record_id = $2 AND locked_by_email = $3`, [tableName, recordId, user.email]);
  }
}

// Check who holds the lock (if anyone). Returns null if unlocked.
async function checkLock(tableName, recordId) {
  // Purge expired locks first.
  await pool.query(`DELETE FROM record_locks WHERE table_name = $1 AND record_id = $2 AND expires_at < NOW()`, [tableName, recordId]);
  const r = await pool.query(
    `SELECT locked_by_email, locked_by_name, expires_at FROM record_locks WHERE table_name = $1 AND record_id = $2`,
    [tableName, recordId]
  );
  return r.rows.length > 0 ? r.rows[0] : null;
}

// Renew (refresh) an existing lock so it doesn't expire while the user is still editing.
async function renewLock(tableName, recordId, user) {
  const r = await pool.query(
    `UPDATE record_locks SET expires_at = NOW() + INTERVAL '${LOCK_TTL_MINUTES} minutes'
     WHERE table_name = $1 AND record_id = $2 AND locked_by_email = $3 RETURNING expires_at`,
    [tableName, recordId, user.email]
  );
  return r.rows.length > 0;
}

// --- Notification helpers ---
async function createNotification(targetEmail, targetRole, type, title, body, referenceId) {
  try {
    await pool.query(
      `INSERT INTO notifications (target_email, target_role, type, title, body, reference_id)
       VALUES ($1, $2, $3, $4, $5, $6)`,
      [targetEmail, targetRole || null, type, title, body || null, referenceId || null]
    );
  } catch (err) {
    console.error('Failed to create notification:', err.message);
  }
}

// Notify all users with a given role (e.g., all ADMINs when a PO is submitted).
async function notifyRole(pool, role, type, title, body, referenceId) {
  try {
    const users = await pool.query('SELECT email FROM users WHERE role = $1 AND is_active = true', [role]);
    for (const u of users.rows) {
      await createNotification(u.email, role, type, title, body, referenceId);
    }
  } catch (err) {
    console.error('Failed to notify role:', err.message);
  }
}

// --- Phase 2: Item matching engine ---
// Matches a vendor invoice line item to an existing Item Master row.
// Strategy (in priority order):
//   1. Exact vendor_item_number match → 100% confidence
//   2. Exact friendly_name match → 95% confidence
//   3. Fuzzy friendly_name match (Levenshtein) → 60-94% confidence
//   4. No match → 0%, needs human review

function levenshteinDistance(a, b) {
  const m = a.length, n = b.length;
  const d = Array.from({ length: m + 1 }, () => new Array(n + 1).fill(0));
  for (let i = 0; i <= m; i++) d[i][0] = i;
  for (let j = 0; j <= n; j++) d[0][j] = j;
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + cost);
    }
  }
  return d[m][n];
}

function calculateSimilarity(a, b) {
  if (!a || !b) return 0;
  if (a === b) return 100;
  const maxLen = Math.max(a.length, b.length);
  if (maxLen === 0) return 100;
  const dist = levenshteinDistance(a.toLowerCase().trim(), b.toLowerCase().trim());
  return Math.round((1 - dist / maxLen) * 100);
}

async function matchItemToMaster(vendorItemNumber, description, pool) {
  // 1. Exact match by vendor_item_number
  if (vendorItemNumber) {
    const r = await pool.query(
      'SELECT sku, friendly_name, vendor_item_number, std_cost_rmb FROM item_master WHERE vendor_item_number = $1 AND status = $2',
      [vendorItemNumber, 'ACTIVE']
    );
    if (r.rows.length > 0) {
      return { sku: r.rows[0].sku, confidence: 100, method: 'vendor_item_number_exact', item: r.rows[0] };
    }
  }

  // 2. Exact match by friendly_name
  if (description) {
    const r = await pool.query(
      'SELECT sku, friendly_name, vendor_item_number, std_cost_rmb FROM item_master WHERE TRIM(LOWER(friendly_name)) = TRIM(LOWER($1)) AND status = $2',
      [description, 'ACTIVE']
    );
    if (r.rows.length > 0) {
      return { sku: r.rows[0].sku, confidence: 95, method: 'name_exact', item: r.rows[0] };
    }
  }

  // 3. Fuzzy match by friendly_name (token-aware: check if description is a substring or vice versa)
  if (description && description.length >= 3) {
    const r = await pool.query(
      'SELECT sku, friendly_name, vendor_item_number, std_cost_rmb FROM item_master WHERE status = $1',
      ['ACTIVE']
    );
    let bestMatch = null;
    let bestScore = 0;
    const descLower = description.toLowerCase().trim();
    for (const item of r.rows) {
      const nameLower = (item.friendly_name || '').toLowerCase().trim();
      // Substring match bonus
      if (nameLower && (nameLower.includes(descLower) || descLower.includes(nameLower))) {
        const score = Math.max(85, calculateSimilarity(description, item.friendly_name));
        if (score > bestScore) { bestScore = score; bestMatch = item; }
        continue;
      }
      const score = calculateSimilarity(description, item.friendly_name);
      if (score > bestScore) { bestScore = score; bestMatch = item; }
    }
    if (bestMatch && bestScore >= 60) {
      return { sku: bestMatch.sku, confidence: bestScore, method: 'name_fuzzy', item: bestMatch };
    }
  }

  return { sku: null, confidence: 0, method: 'none', item: null };
}

// Parse OCR text into line items. Uses heuristics: looks for lines that contain
// a price (number with decimals) and optionally a quantity and item number.
function parseOcrLineItems(text) {
  const lines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);
  const items = [];

  for (const line of lines) {
    // Skip very short lines or lines that look like headers/totals
    if (line.length < 5) continue;
    if (/^(total|subtotal|tax|vat|amount|balance|deposit|grand total)/i.test(line)) continue;

    // Extract price: number with 2 decimal places (e.g., 12.50, 1,234.00)
    const priceMatch = line.match(/(\d{1,3}(?:[,]\d{3})*\.\d{2}|\d+\.\d{2})/);
    if (!priceMatch) continue;

    let priceStr = priceMatch[1].replace(/,/g, '');
    const unitPrice = parseFloat(priceStr);
    if (!unitPrice || unitPrice <= 0) continue;

    // Try to extract quantity (integer before the price)
    const qtyMatch = line.match(/(\d+)\s*(?:pcs|pieces|units|qty|quantity|x)?[\s]*$/i);
    const quantity = qtyMatch ? parseInt(qtyMatch[1]) : 1;

    // Remove numbers from the line to get the description
    let description = line
      .replace(/\d{1,3}(?:[,]\d{3})*\.\d{2}/g, '')
      .replace(/\d+/g, '')
      .replace(/\s+/g, ' ')
      .trim();

    // Try to extract an item number (alphanumeric code with hyphens)
    const itemNoMatch = line.match(/([A-Z]{1,4}[-]?\d{2,6}[-]?[A-Z0-9]{0,4})/i);
    const vendorItemNumber = itemNoMatch ? itemNoMatch[1] : null;

    if (description.length < 2) description = vendorItemNumber || `Line ${items.length + 1}`;

    items.push({
      vendor_item_number: vendorItemNumber,
      description,
      quantity,
      unit_price: unitPrice,
      line_total: unitPrice * quantity
    });
  }

  return items;
}

// --- PO status state machine ---
// Defines which transitions are legal and who can trigger them.
const PO_TRANSITIONS = {
  DRAFT:     { SUBMITTED: ['ADMIN', 'BUYER'] },
  SUBMITTED: { APPROVED: ['ADMIN'], REJECTED: ['ADMIN'] },
  APPROVED:  { ISSUED: ['ADMIN', 'BUYER'] },
  REJECTED:  { SUBMITTED: ['ADMIN', 'BUYER'], DRAFT: ['ADMIN', 'BUYER'] },
  ISSUED:    { PARTIAL: ['ADMIN', 'BUYER', 'LOGISTICS'], RECEIVED: ['ADMIN', 'BUYER', 'LOGISTICS'] },
  PARTIAL:   { RECEIVED: ['ADMIN', 'BUYER', 'LOGISTICS'] },
  RECEIVED:  { CLOSED: ['ADMIN'] },
  CLOSED:    {}
};

function canTransitionPO(currentStatus, newStatus, userRole) {
  const allowed = PO_TRANSITIONS[currentStatus];
  if (!allowed || !allowed[newStatus]) return false;
  return allowed[newStatus].includes(userRole);
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
app.get(['/invoices.html'], requireAuthPage(['ADMIN', 'BUYER']), (req, res) => res.sendFile(path.join(__dirname, 'public', 'invoices.html')));
app.get(['/admin-users.html'], requireAuthPage(['ADMIN']), (req, res) => res.sendFile(path.join(__dirname, 'public', 'admin-users.html')));
app.get(['/b2b-orders.html'], requireAuthPage(['ADMIN']), (req, res) => res.sendFile(path.join(__dirname, 'public', 'b2b-orders.html')));
app.get(['/b2b-customers.html'], requireAuthPage(['ADMIN']), (req, res) => res.sendFile(path.join(__dirname, 'public', 'b2b-customers.html')));
app.get(['/create-b2b-order.html'], requireAuthPage(['ADMIN']), (req, res) => res.sendFile(path.join(__dirname, 'public', 'create-b2b-order.html')));
app.get('/lookbook.html', (req, res) => res.sendFile(path.join(__dirname, 'public', 'lookbook.html')));

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

app.patch('/api/admin/users/:id/role', requireAuthApi(['ADMIN']), async (req, res) => {
  try {
    const { role } = req.body;
    if (!role || !['ADMIN', 'BUYER', 'ACCOUNTS', 'LOGISTICS', 'B2B_CUSTOMER'].includes(role)) {
      return res.status(400).json({ error: 'Invalid role.' });
    }
    const r = await pool.query(
      'UPDATE users SET role = $1 WHERE id = $2 RETURNING id, email, role',
      [role, req.params.id]
    );
    if (r.rows.length === 0) return res.status(404).json({ error: 'User not found' });
    logActivity(req.user, 'ROLE_CHANGED', r.rows[0].email, { new_role: role });
    res.json({ success: true, user: r.rows[0] });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to update role' });
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
  let query = "SELECT sku, friendly_name, category_code, year_code, collection_code, department_code, department_name, color_code, material, std_cost_rmb, std_cost_rmb/7.0 as std_cost_usd, status, barcode, vendor_item_number, hs_code, cbm FROM item_master WHERE status='ACTIVE'";
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
  const { vendor_code, po_date, po_currency, exchange_rate, items, invoice_reference, invoice_id } = req.body;
  const po_id = `PO-${Date.now().toString().slice(-6)}`;
  
  try {
    const vendorRes = await pool.query("SELECT vendor_code, category FROM vendors WHERE vendor_code = $1", [vendor_code]);
    if (vendorRes.rows.length === 0) return res.status(400).json({ error: "Vendor not found" });
    const vendorCategory = vendorRes.rows[0].category;

    // Validate linked invoice if provided
    let resolvedInvoiceId = invoice_id || null;
    let resolvedInvoiceRef = invoice_reference || null;
    if (resolvedInvoiceId) {
      const invRes = await pool.query('SELECT id, invoice_number, vendor_code, status FROM vendor_invoices WHERE id = $1', [resolvedInvoiceId]);
      if (invRes.rows.length === 0) return res.status(400).json({ error: "Invoice not found" });
      const inv = invRes.rows[0];
      if (inv.vendor_code && inv.vendor_code !== vendor_code) {
        return res.status(400).json({ error: `Invoice vendor ${inv.vendor_code} does not match selected vendor ${vendor_code}` });
      }
      if (!resolvedInvoiceRef) resolvedInvoiceRef = inv.invoice_number;
      if (['PO_CREATED', 'ERROR'].includes(inv.status)) {
        return res.status(400).json({ error: `Invoice status is ${inv.status}. Cannot link to PO.` });
      }
    }

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
      `INSERT INTO purchase_orders (po_id, vendor_code, po_date, invoice_currency, exchange_rate_to_rmb, status, total_rmb, total_usd, deposit_usd, balance_usd, created_by, invoice_reference, invoice_id)
       VALUES ($1, $2, $3, $4, $5, 'DRAFT', $6, $7, $8, $9, $10, $11, $12)`,
      [po_id, vendor_code, po_date, po_currency, exchange_rate_to_rmb, total_rmb, total_usd, deposit_usd, balance_usd, req.user.email, resolvedInvoiceRef, resolvedInvoiceId]
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
             m.vendor_item_number, m.hs_code, m.cbm
      FROM po_line_items li
      JOIN item_master m ON li.sku = m.sku
      WHERE li.po_id = $1
    `, [po_id]);

    const deposit_rmb = deposit_usd * exchange_rate_to_rmb;
    const balance_rmb = balance_usd * exchange_rate_to_rmb;

    logActivity(req.user, 'PO_CREATED', po_id, { vendor_code, total_rmb: total_rmb.toFixed(2), total_usd: total_usd.toFixed(2), line_count: items.length, status: 'DRAFT', invoice_reference: resolvedInvoiceRef, invoice_id: resolvedInvoiceId });

    res.json({
      success: true,
      po_id,
      vendor_code,
      po_date,
      po_currency,
      exchange_rate_to_rmb,
      status: 'DRAFT',
      total_usd: total_usd.toFixed(2),
      total_rmb: total_rmb.toFixed(2),
      deposit_usd: deposit_usd.toFixed(2),
      balance_usd: balance_usd.toFixed(2),
      deposit_rmb: deposit_rmb.toFixed(2),
      balance_rmb: balance_rmb.toFixed(2),
      invoice_reference: resolvedInvoiceRef,
      invoice_id: resolvedInvoiceId,
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
    const { status } = req.query;
    // Try with Phase 1 columns first; fall back to base columns if migration hasn't run
    const fullSelect = 'po_id, vendor_code, po_date, status, total_usd, total_rmb, created_by, approved_by, submitted_at, approved_at, rejection_reason, version, invoice_reference, invoice_id';
    const baseSelect = 'po_id, vendor_code, po_date, status, total_usd, total_rmb, created_by, invoice_reference, invoice_id';
    let query, params, useFallback = false;
    try {
      query = `SELECT ${fullSelect} FROM purchase_orders ${status ? 'WHERE status = $1' : ''} ORDER BY po_date DESC, po_id DESC`;
      params = status ? [status] : [];
      const result = await pool.query(query, params);
      return res.json(result.rows);
    } catch (colErr) {
      // Phase 1 columns don't exist yet — fall back to base query
      console.error('PO query with Phase 1 columns failed, falling back:', colErr.message);
      useFallback = true;
    }
    if (useFallback) {
      query = `SELECT ${baseSelect} FROM purchase_orders ${status ? 'WHERE status = $1' : ''} ORDER BY po_date DESC, po_id DESC`;
      params = status ? [status] : [];
      const result = await pool.query(query, params);
      return res.json(result.rows);
    }
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
             m.vendor_item_number, m.hs_code, m.cbm
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
      invoice_reference: po.invoice_reference,
      invoice_id: po.invoice_id,
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
    // Only DRAFT POs can be deleted. Once submitted/approved/issued, they must be voided via an amendment.
    if (!['DRAFT', 'REJECTED'].includes(check.rows[0].status)) {
      await client.query('ROLLBACK');
      return res.status(400).json({ error: `Cannot delete a PO with status ${check.rows[0].status}. Only DRAFT or REJECTED POs can be deleted.` });
    }
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

    // State machine guard: can only receive against ISSUED or PARTIAL POs.
    const poStatusRes = await client.query('SELECT status FROM purchase_orders WHERE po_id = $1', [po_id]);
    if (poStatusRes.rows.length === 0) { await client.query('ROLLBACK'); return res.status(404).json({ error: "PO not found" }); }
    if (!['ISSUED', 'PARTIAL'].includes(poStatusRes.rows[0].status)) {
      await client.query('ROLLBACK');
      return res.status(409).json({ error: `Cannot receive against a PO with status ${poStatusRes.rows[0].status}. PO must be ISSUED first.` });
    }

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

    // Notify the PO creator that stock has been received.
    const creatorRes = await pool.query('SELECT created_by FROM purchase_orders WHERE po_id = $1', [po_id]);
    if (creatorRes.rows[0]?.created_by && creatorRes.rows[0].created_by !== req.user.email) {
      await createNotification(creatorRes.rows[0].created_by, null, 'STOCK_RECEIVED',
        `Stock received for PO ${po_id}`, `${newStatus === 'RECEIVED' ? 'Fully received' : 'Partially received'} by ${req.user.name || req.user.email}`, po_id);
    }

    res.json({ success: true, po_id, status: newStatus, received: receivedLines });
  } catch (err) {
    await client.query('ROLLBACK');
    console.error(err);
    res.status(500).json({ error: "Failed to receive PO" });
  } finally {
    client.release();
  }
});

// --- PO State Machine: submit, approve, reject, issue ---

// Submit a DRAFT PO for approval. Notifies all ADMINs.
app.post('/api/purchase-orders/:po_id/submit', requireAuthApi(['ADMIN', 'BUYER']), async (req, res) => {
  try {
    const r = await pool.query('SELECT status, created_by, version FROM purchase_orders WHERE po_id = $1', [req.params.po_id]);
    if (r.rows.length === 0) return res.status(404).json({ error: 'PO not found' });
    const po = r.rows[0];
    if (!canTransitionPO(po.status, 'SUBMITTED', req.user.role)) {
      return res.status(409).json({ error: `Cannot submit a PO with status ${po.status}` });
    }
    // Only the creator or an admin can submit.
    if (po.created_by !== req.user.email && req.user.role !== 'ADMIN') {
      return res.status(403).json({ error: 'Only the PO creator or an admin can submit it' });
    }
    await pool.query(
      `UPDATE purchase_orders SET status = 'SUBMITTED', submitted_at = NOW(), version = version + 1 WHERE po_id = $1`,
      [req.params.po_id]
    );
    logActivity(req.user, 'PO_SUBMITTED', req.params.po_id, { previous_status: po.status });
    await notifyRole(pool, 'ADMIN', 'PO_SUBMITTED', `PO ${req.params.po_id} awaiting approval`, `Submitted by ${req.user.name || req.user.email}`, req.params.po_id);
    res.json({ success: true, po_id: req.params.po_id, status: 'SUBMITTED' });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to submit PO' });
  }
});

// Approve a SUBMITTED PO. Only ADMIN can approve.
app.post('/api/purchase-orders/:po_id/approve', requireAuthApi(['ADMIN']), async (req, res) => {
  try {
    const r = await pool.query('SELECT status, version FROM purchase_orders WHERE po_id = $1', [req.params.po_id]);
    if (r.rows.length === 0) return res.status(404).json({ error: 'PO not found' });
    const po = r.rows[0];
    if (!canTransitionPO(po.status, 'APPROVED', req.user.role)) {
      return res.status(409).json({ error: `Cannot approve a PO with status ${po.status}` });
    }
    await pool.query(
      `UPDATE purchase_orders SET status = 'APPROVED', approved_by = $1, approved_at = NOW(), version = version + 1 WHERE po_id = $2`,
      [req.user.email, req.params.po_id]
    );
    logActivity(req.user, 'PO_APPROVED', req.params.po_id, null);

    // Notify the creator that their PO was approved.
    const creatorRes = await pool.query('SELECT created_by FROM purchase_orders WHERE po_id = $1', [req.params.po_id]);
    if (creatorRes.rows[0]?.created_by) {
      await createNotification(creatorRes.rows[0].created_by, null, 'PO_APPROVED', `PO ${req.params.po_id} approved`, `Approved by ${req.user.name || req.user.email}. Ready to issue.`, req.params.po_id);
    }
    res.json({ success: true, po_id: req.params.po_id, status: 'APPROVED', approved_by: req.user.email });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to approve PO' });
  }
});

// Reject a SUBMITTED PO. Sends it back to DRAFT with a reason.
app.post('/api/purchase-orders/:po_id/reject', requireAuthApi(['ADMIN']), async (req, res) => {
  const { reason } = req.body;
  try {
    const r = await pool.query('SELECT status, version FROM purchase_orders WHERE po_id = $1', [req.params.po_id]);
    if (r.rows.length === 0) return res.status(404).json({ error: 'PO not found' });
    const po = r.rows[0];
    if (!canTransitionPO(po.status, 'REJECTED', req.user.role)) {
      return res.status(409).json({ error: `Cannot reject a PO with status ${po.status}` });
    }
    await pool.query(
      `UPDATE purchase_orders SET status = 'REJECTED', rejection_reason = $1, version = version + 1 WHERE po_id = $2`,
      [reason || null, req.params.po_id]
    );
    logActivity(req.user, 'PO_REJECTED', req.params.po_id, { reason });

    // Notify the creator that their PO was rejected.
    const creatorRes = await pool.query('SELECT created_by FROM purchase_orders WHERE po_id = $1', [req.params.po_id]);
    if (creatorRes.rows[0]?.created_by) {
      await createNotification(creatorRes.rows[0].created_by, null, 'PO_REJECTED', `PO ${req.params.po_id} rejected`, reason ? `Reason: ${reason}` : `Rejected by ${req.user.name || req.user.email}`, req.params.po_id);
    }
    res.json({ success: true, po_id: req.params.po_id, status: 'REJECTED', reason });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to reject PO' });
  }
});

// Issue an APPROVED PO to the supplier (dispatch). Transitions to ISSUED.
app.post('/api/purchase-orders/:po_id/issue', requireAuthApi(['ADMIN', 'BUYER']), async (req, res) => {
  try {
    const r = await pool.query('SELECT status, version FROM purchase_orders WHERE po_id = $1', [req.params.po_id]);
    if (r.rows.length === 0) return res.status(404).json({ error: 'PO not found' });
    const po = r.rows[0];
    if (!canTransitionPO(po.status, 'ISSUED', req.user.role)) {
      return res.status(409).json({ error: `Cannot issue a PO with status ${po.status}. Must be APPROVED first.` });
    }
    await pool.query(
      `UPDATE purchase_orders SET status = 'ISSUED', version = version + 1 WHERE po_id = $1`,
      [req.params.po_id]
    );
    logActivity(req.user, 'PO_ISSUED', req.params.po_id, null);
    res.json({ success: true, po_id: req.params.po_id, status: 'ISSUED' });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to issue PO' });
  }
});

// --- Item Master: single-item fetch, edit with optimistic locking, record locking ---

// Get a single item for editing (includes version for optimistic locking).
app.get('/api/items/:sku', requireAuthApi(['ADMIN', 'BUYER']), async (req, res) => {
  try {
    const r = await pool.query(
      `SELECT sku, friendly_name, category_code, year_code, collection_code, department_code, department_name,
              color_code, material, std_cost_rmb, status, barcode, vendor_item_number, hs_code, cbm, description, version
       FROM item_master WHERE sku = $1`,
      [req.params.sku]
    );
    if (r.rows.length === 0) return res.status(404).json({ error: 'Item not found' });

    // Attach lock info so the frontend knows if someone else is editing.
    const lock = await checkLock('item_master', req.params.sku);
    res.json({ ...r.rows[0], locked_by: lock });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to fetch item' });
  }
});

// Update an item with optimistic locking. The client must send the version it
// last read; if the server's version is higher, return 409 Conflict.
app.put('/api/items/:sku', requireAuthApi(['ADMIN', 'BUYER']), async (req, res) => {
  const { friendly_name, category_code, year_code, collection_code, department_code, department_name,
          color_code, material, std_cost_rmb, status, hs_code, description, cbm, version } = req.body;

  if (version === undefined || version === null) {
    return res.status(400).json({ error: 'version is required for optimistic locking' });
  }

  try {
    const r = await pool.query(
      `UPDATE item_master SET
         friendly_name = COALESCE($1, friendly_name),
         category_code = COALESCE($2, category_code),
         year_code = COALESCE($3, year_code),
         collection_code = COALESCE($4, collection_code),
         department_code = COALESCE($5, department_code),
         department_name = COALESCE($6, department_name),
         color_code = COALESCE($7, color_code),
         material = COALESCE($8, material),
         std_cost_rmb = COALESCE($9, std_cost_rmb),
         status = COALESCE($10, status),
         hs_code = COALESCE($11, hs_code),
         description = COALESCE($12, description),
         cbm = COALESCE($13, cbm),
         version = version + 1
       WHERE sku = $14 AND version = $15
       RETURNING sku, friendly_name, version`,
      [friendly_name, category_code, year_code, collection_code, department_code, department_name,
       color_code, material, std_cost_rmb, status, hs_code, description, cbm, req.params.sku, version]
    );

    if (r.rows.length === 0) {
      // Either the row doesn't exist, or the version doesn't match (someone else edited it).
      const exists = await pool.query('SELECT sku FROM item_master WHERE sku = $1', [req.params.sku]);
      if (exists.rows.length === 0) return res.status(404).json({ error: 'Item not found' });
      return res.status(409).json({
        error: 'This item was modified by another user. Please refresh and try again.',
        conflict: true
      });
    }

    logActivity(req.user, 'ITEM_UPDATED', req.params.sku, { updated_fields: Object.keys(req.body).filter(k => k !== 'version') });

    // Release the edit lock on save.
    await releaseLock('item_master', req.params.sku, req.user);

    res.json({ success: true, item: r.rows[0] });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to update item' });
  }
});

// Upload a product image for an Item Master row.  The file is saved to
// public/images/<vendor_item_number>.jpg so the PO PDF generator can
// embed it.  Falls back to <sku>.jpg if vendor_item_number is blank.
const imageUpload = multer({
  storage: multer.diskStorage({
    destination: path.join(__dirname, 'public', 'images'),
    filename: (req, file, cb) => {
      cb(null, `tmp_${Date.now()}_${file.originalname.replace(/[^a-zA-Z0-9._-]/g, '_')}`);
    }
  }),
  limits: { fileSize: 5 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const ok = /jpeg|jpg|png|webp/.test(file.mimetype);
    cb(ok ? null : new Error('Only JPEG, PNG, or WebP images allowed'), ok);
  }
});

app.post('/api/items/:sku/upload-image', requireAuthApi(['ADMIN', 'BUYER']), imageUpload.single('image'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No image file uploaded' });

    const itemRes = await pool.query('SELECT sku, vendor_item_number FROM item_master WHERE sku = $1', [req.params.sku]);
    if (itemRes.rows.length === 0) return res.status(404).json({ error: 'Item not found' });

    const item = itemRes.rows[0];
    const baseName = (item.vendor_item_number || item.sku).replace(/[^a-zA-Z0-9._-]/g, '_');
    const newPath = path.join(__dirname, 'public', 'images', `${baseName}.jpg`);

    try { fs.unlinkSync(newPath); } catch (e) {}
    fs.renameSync(req.file.path, newPath);

    logActivity(req.user, 'ITEM_IMAGE_UPLOADED', req.params.sku, { filename: `${baseName}.jpg` });

    res.json({ success: true, image_url: `/images/${baseName}.jpg` });
  } catch (err) {
    console.error('Image upload error:', err);
    res.status(500).json({ error: 'Failed to upload image: ' + err.message });
  }
});

// Acquire an edit lock on an item (pessimistic locking).
app.post('/api/items/:sku/lock', requireAuthApi(['ADMIN', 'BUYER']), async (req, res) => {
  try {
    const result = await acquireLock('item_master', req.params.sku, req.user);
    if (result.success) {
      logActivity(req.user, 'ITEM_LOCKED', req.params.sku, null);
      res.json({ success: true, locked_by: { email: req.user.email, name: req.user.name || req.user.email } });
    } else {
      res.status(409).json({
        error: `This item is being edited by ${result.lockedBy.locked_by_name || result.lockedBy.locked_by_email}`,
        locked_by: result.lockedBy
      });
    }
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to acquire lock' });
  }
});

// Release an edit lock.
app.delete('/api/items/:sku/lock', requireAuthApi(['ADMIN', 'BUYER']), async (req, res) => {
  try {
    await releaseLock('item_master', req.params.sku, req.user);
    res.json({ success: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to release lock' });
  }
});

// Renew (heartbeat) an edit lock so it doesn't expire while the user is still editing.
app.post('/api/items/:sku/lock/renew', requireAuthApi(['ADMIN', 'BUYER']), async (req, res) => {
  try {
    const renewed = await renewLock('item_master', req.params.sku, req.user);
    if (renewed) {
      res.json({ success: true });
    } else {
      res.status(404).json({ error: 'No active lock found for this user on this item' });
    }
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to renew lock' });
  }
});

// --- Notifications ---

// Get the current user's notifications (most recent 50).
app.get('/api/notifications', requireAuthApi(['ADMIN', 'BUYER', 'ACCOUNTS', 'LOGISTICS']), async (req, res) => {
  try {
    const r = await pool.query(
      `SELECT id, type, title, body, reference_id, is_read, created_at
       FROM notifications WHERE target_email = $1
       ORDER BY created_at DESC LIMIT 50`,
      [req.user.email]
    );
    res.json(r.rows);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to fetch notifications' });
  }
});

// Get unread notification count (for the nav badge).
app.get('/api/notifications/unread-count', requireAuthApi(['ADMIN', 'BUYER', 'ACCOUNTS', 'LOGISTICS']), async (req, res) => {
  try {
    const r = await pool.query(
      'SELECT COUNT(*)::int AS count FROM notifications WHERE target_email = $1 AND is_read = false',
      [req.user.email]
    );
    res.json({ count: r.rows[0].count });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to fetch unread count' });
  }
});

// Mark a single notification as read.
app.post('/api/notifications/:id/read', requireAuthApi(['ADMIN', 'BUYER', 'ACCOUNTS', 'LOGISTICS']), async (req, res) => {
  try {
    await pool.query('UPDATE notifications SET is_read = true WHERE id = $1 AND target_email = $2', [req.params.id, req.user.email]);
    res.json({ success: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to mark notification as read' });
  }
});

// Mark all notifications as read.
app.post('/api/notifications/read-all', requireAuthApi(['ADMIN', 'BUYER', 'ACCOUNTS', 'LOGISTICS']), async (req, res) => {
  try {
    await pool.query('UPDATE notifications SET is_read = true WHERE target_email = $1 AND is_read = false', [req.user.email]);
    res.json({ success: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to mark all as read' });
  }
});

const upload = multer({ storage: multer.memoryStorage() });
app.post('/api/upload-invoice', requireAuthApi(['ADMIN', 'BUYER']), upload.single('invoiceFile'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "No file uploaded" });

    const { currency } = req.body;

    // Strip BOM from CSV files — xlsx.read doesn't always handle UTF-8 BOM correctly,
    // causing the first column header to be "\ufeffItem No." which won't match.
    let fileBuffer = req.file.buffer;
    if (fileBuffer.length >= 3 && fileBuffer[0] === 0xEF && fileBuffer[1] === 0xBB && fileBuffer[2] === 0xBF) {
      fileBuffer = fileBuffer.subarray(3); // Remove UTF-8 BOM
    }

    const workbook = xlsx.read(fileBuffer, { type: 'buffer' });
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
      // Build a helpful error message showing what columns were found
      const foundHeaders = rawRows.length > 0 ? (rawRows[0] || []).map(c => String(c).trim()).filter(Boolean).join(', ') : '(empty file)';
      return res.status(400).json({
        error: `Could not find required columns. Your file must have "Item No." and "UNIT PRICE" columns. Found headers: ${foundHeaders}`
      });
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

// ============================================================
// Phase 2: Auto-PO Agent — Invoice upload, matching, PO generation
// ============================================================

// Upload a vendor invoice (Excel/CSV or image). The file is parsed, line items
// are extracted, and each is matched against the Item Master. Results are stored
// in vendor_invoices + invoice_line_items for the review screen.
app.post('/api/invoices/upload', requireAuthApi(['ADMIN', 'BUYER']), upload.single('invoiceFile'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

    const { vendor_code, currency, invoice_date } = req.body;
    if (!vendor_code) return res.status(400).json({ error: 'vendor_code is required' });

    // Verify vendor exists
    const vendorRes = await pool.query('SELECT vendor_code, vendor_name FROM vendors WHERE vendor_code = $1', [vendor_code]);
    if (vendorRes.rows.length === 0) return res.status(400).json({ error: 'Vendor not found' });

    const fileName = req.file.originalname;
    const fileExt = fileName.split('.').pop().toLowerCase();
    const isImage = ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp'].includes(fileExt);
    const isExcel = ['xlsx', 'xls', 'csv'].includes(fileExt);

    if (!isImage && !isExcel) {
      return res.status(400).json({ error: 'Unsupported file type. Please upload Excel (.xlsx/.xls/.csv) or image (.png/.jpg/.jpeg).' });
    }

    // Create the invoice record
    const invoiceRes = await pool.query(
      `INSERT INTO vendor_invoices (invoice_number, vendor_code, invoice_date, file_name, file_type, status, currency, created_by)
       VALUES ($1, $2, $3, $4, $5, 'PROCESSING', $6, $7) RETURNING id`,
      [
        req.body.invoice_number || `INV-${Date.now().toString().slice(-6)}`,
        vendor_code,
        invoice_date || new Date().toISOString().split('T')[0],
        fileName,
        isImage ? 'IMAGE' : 'EXCEL',
        currency || 'RMB',
        req.user.email
      ]
    );
    const invoiceId = invoiceRes.rows[0].id;

    let lineItems = [];

    if (isExcel) {
      // Parse Excel/CSV using the same proven logic from /api/upload-invoice
      const workbook = xlsx.read(req.file.buffer);
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rawRows = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: '' });

      let headerRowIdx = -1, colItemNo = -1, colFriendly = -1, colPrice = -1, colQty = -1;

      for (let i = 0; i < rawRows.length; i++) {
        const row = rawRows[i];
        const itemNoIndex = row.findIndex(cell =>
          String(cell).trim().includes('Item No.') || String(cell).trim().includes('货号') || String(cell).trim().includes('ITEM NO')
        );
        if (itemNoIndex !== -1) {
          headerRowIdx = i;
          colItemNo = itemNoIndex;
          colFriendly = row.findIndex(cell =>
            String(cell).trim().includes('Friendly Name') || String(cell).trim().includes('Item Name') || String(cell).trim().includes('Name')
          );
          colPrice = row.findIndex(cell =>
            String(cell).trim().includes('UNIT PRICE') || String(cell).trim().includes('UNTI PRICE') || String(cell).trim().includes('单价') || String(cell).trim().includes('Price')
          );
          colQty = row.findIndex(cell =>
            String(cell).trim().includes('QTY') || String(cell).trim().includes('Quantity') || String(cell).trim().includes('数量')
          );
          // Check next 3 rows for header continuation
          if (colFriendly === -1 || colPrice === -1) {
            for (let j = i + 1; j < Math.min(i + 4, rawRows.length); j++) {
              const nextRow = rawRows[j];
              if (colFriendly === -1) colFriendly = nextRow.findIndex(cell => String(cell).trim().includes('Friendly Name') || String(cell).trim().includes('Item Name') || String(cell).trim().includes('Name'));
              if (colPrice === -1) colPrice = nextRow.findIndex(cell => String(cell).trim().includes('UNIT PRICE') || String(cell).trim().includes('Price'));
              if (colQty === -1) colQty = nextRow.findIndex(cell => String(cell).trim().includes('QTY') || String(cell).trim().includes('Quantity'));
            }
          }
          break;
        }
      }

      if (headerRowIdx === -1 || colItemNo === -1 || colPrice === -1) {
        await pool.query(`UPDATE vendor_invoices SET status = 'ERROR', error_message = 'Could not find required columns (Item No., Unit Price)' WHERE id = $1`, [invoiceId]);
        return res.status(400).json({ error: 'Could not find required columns (Item No., Unit Price) in the Excel file.' });
      }

      for (let i = headerRowIdx + 1; i < rawRows.length; i++) {
        const row = rawRows[i];
        const itemNo = row[colItemNo] ? String(row[colItemNo]).trim() : '';
        const friendlyName = (colFriendly !== -1 && row[colFriendly]) ? String(row[colFriendly]).trim() : itemNo;
        let priceStr = row[colPrice] ? String(row[colPrice]) : '';
        priceStr = priceStr.replace(/[^0-9.]/g, '');
        const price = parseFloat(priceStr) || 0;
        const qty = (colQty !== -1 && row[colQty]) ? parseInt(String(row[colQty]).replace(/[^0-9]/g, '')) || 1 : 1;

        if (!itemNo || price === 0) continue;

        lineItems.push({
          vendor_item_number: itemNo,
          description: friendlyName || itemNo,
          quantity: qty,
          unit_price: price,
          line_total: price * qty
        });
      }
    } else if (isImage) {
      // OCR with Tesseract.js
      if (!Tesseract) {
        await pool.query(`UPDATE vendor_invoices SET status = 'ERROR', error_message = 'Image OCR not available. Please upload Excel/CSV instead.' WHERE id = $1`, [invoiceId]);
        return res.status(400).json({ error: 'Image OCR is not available on this server. Please upload an Excel or CSV file instead.' });
      }

      try {
        const { data: { text } } = await Tesseract.recognize(req.file.buffer, 'eng');
        lineItems = parseOcrLineItems(text);

        if (lineItems.length === 0) {
          await pool.query(`UPDATE vendor_invoices SET status = 'ERROR', error_message = 'OCR completed but no line items could be extracted from the image.' WHERE id = $1`, [invoiceId]);
          return res.status(400).json({ error: 'OCR completed but no line items could be extracted. Try uploading an Excel/CSV version of the invoice for better results.' });
        }
      } catch (ocrErr) {
        console.error('OCR error:', ocrErr.message);
        await pool.query(`UPDATE vendor_invoices SET status = 'ERROR', error_message = $1 WHERE id = $2`, [ocrErr.message, invoiceId]);
        return res.status(500).json({ error: 'OCR processing failed: ' + ocrErr.message });
      }
    }

    if (lineItems.length === 0) {
      await pool.query(`UPDATE vendor_invoices SET status = 'ERROR', error_message = 'No line items found in the file.' WHERE id = $1`, [invoiceId]);
      return res.status(400).json({ error: 'No line items found in the uploaded file.' });
    }

    // Match each line item to the Item Master
    let totalAmount = 0;
    let autoMatchedCount = 0;
    let reviewNeededCount = 0;

    for (let i = 0; i < lineItems.length; i++) {
      const li = lineItems[i];
      totalAmount += li.line_total;

      const match = await matchItemToMaster(li.vendor_item_number, li.description, pool);

      const matchStatus = match.confidence >= 80 ? 'AUTO_MATCHED' : 'REVIEW_NEEDED';
      if (matchStatus === 'AUTO_MATCHED') autoMatchedCount++; else reviewNeededCount++;

      await pool.query(
        `INSERT INTO invoice_line_items (invoice_id, seq, vendor_item_number, description, quantity, unit_price, line_total, matched_sku, match_confidence, match_status, match_method)
         VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11)`,
        [invoiceId, i, li.vendor_item_number, li.description, li.quantity, li.unit_price, li.line_total, match.sku, match.confidence, matchStatus, match.method]
      );
    }

    // Update invoice status
    await pool.query(
      `UPDATE vendor_invoices SET status = 'MATCHED', total_amount = $1 WHERE id = $2`,
      [totalAmount, invoiceId]
    );

    logActivity(req.user, 'INVOICE_UPLOADED', String(invoiceId), {
      vendor_code, file_name: fileName, line_count: lineItems.length,
      auto_matched: autoMatchedCount, review_needed: reviewNeededCount
    });

    // Notify admins that a new invoice is ready for review
    await notifyRole(pool, 'ADMIN', 'INVOICE_UPLOADED',
      `New invoice from ${vendorRes.rows[0].vendor_name}`,
      `${autoMatchedCount} auto-matched, ${reviewNeededCount} need review. Uploaded by ${req.user.name || req.user.email}`,
      String(invoiceId)
    );

    res.json({
      success: true,
      invoice_id: invoiceId,
      status: 'MATCHED',
      total_items: lineItems.length,
      auto_matched: autoMatchedCount,
      review_needed: reviewNeededCount,
      total_amount: totalAmount.toFixed(2)
    });
  } catch (err) {
    console.error('Invoice upload error:', err);
    res.status(500).json({ error: 'Failed to process invoice: ' + err.message });
  }
});

// List all vendor invoices
app.get('/api/invoices', requireAuthApi(['ADMIN', 'BUYER']), async (req, res) => {
  try {
    const r = await pool.query(`
      SELECT vi.id, vi.invoice_number, vi.vendor_code, v.vendor_name, vi.invoice_date,
             vi.file_name, vi.file_type, vi.status, vi.total_amount, vi.currency,
             vi.po_id, vi.created_by, vi.created_at, vi.error_message,
             (SELECT COUNT(*) FROM invoice_line_items WHERE invoice_id = vi.id) AS total_items,
             (SELECT COUNT(*) FROM invoice_line_items WHERE invoice_id = vi.id AND match_status = 'AUTO_MATCHED') AS auto_matched,
             (SELECT COUNT(*) FROM invoice_line_items WHERE invoice_id = vi.id AND match_status = 'REVIEW_NEEDED') AS review_needed
      FROM vendor_invoices vi
      LEFT JOIN vendors v ON vi.vendor_code = v.vendor_code
      ORDER BY vi.created_at DESC
    `);
    res.json(r.rows);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to fetch invoices' });
  }
});

// Get invoice details with all line items and their match info
app.get('/api/invoices/:id', requireAuthApi(['ADMIN', 'BUYER']), async (req, res) => {
  try {
    const invRes = await pool.query(`
      SELECT vi.*, v.vendor_name
      FROM vendor_invoices vi
      LEFT JOIN vendors v ON vi.vendor_code = v.vendor_code
      WHERE vi.id = $1
    `, [req.params.id]);

    if (invRes.rows.length === 0) return res.status(404).json({ error: 'Invoice not found' });

    const itemsRes = await pool.query(`
      SELECT ili.*, im.friendly_name AS matched_name, im.barcode AS matched_barcode,
             im.std_cost_rmb AS matched_cost, im.material AS matched_material
      FROM invoice_line_items ili
      LEFT JOIN item_master im ON ili.matched_sku = im.sku
      WHERE ili.invoice_id = $1
      ORDER BY ili.seq
    `, [req.params.id]);

    res.json({ ...invRes.rows[0], line_items: itemsRes.rows });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to fetch invoice details' });
  }
});

// Update a line item's match (confirm, change SKU, mark as new, or reject)
app.put('/api/invoice-items/:id', requireAuthApi(['ADMIN', 'BUYER']), async (req, res) => {
  const { matched_sku, match_status } = req.body;
  if (!match_status || !['AUTO_MATCHED', 'REVIEW_NEEDED', 'CONFIRMED', 'REJECTED', 'NEW_ITEM'].includes(match_status)) {
    return res.status(400).json({ error: 'Invalid match_status' });
  }

  try {
    // If changing to a specific SKU, verify it exists
    let confidence = null;
    let method = null;
    if (matched_sku && match_status !== 'NEW_ITEM' && match_status !== 'REJECTED') {
      const itemRes = await pool.query('SELECT sku FROM item_master WHERE sku = $1', [matched_sku]);
      if (itemRes.rows.length === 0) return res.status(404).json({ error: 'SKU not found in Item Master' });
      confidence = 100;
      method = 'manual';
    }

    const r = await pool.query(
      `UPDATE invoice_line_items
       SET matched_sku = $1, match_status = $2, match_confidence = COALESCE($3, match_confidence), match_method = COALESCE($4, match_method)
       WHERE id = $5 RETURNING *`,
      [matched_sku || null, match_status, confidence, method, req.params.id]
    );

    if (r.rows.length === 0) return res.status(404).json({ error: 'Line item not found' });
    res.json({ success: true, item: r.rows[0] });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to update line item' });
  }
});

// Search Item Master for manual matching (autocomplete)
app.get('/api/invoices/search-items', requireAuthApi(['ADMIN', 'BUYER']), async (req, res) => {
  const { q } = req.query;
  if (!q || q.length < 1) return res.json([]);

  try {
    const r = await pool.query(
      `SELECT sku, friendly_name, category_code, color_code, material, std_cost_rmb, vendor_item_number, barcode
       FROM item_master
       WHERE status = 'ACTIVE' AND (
         sku ILIKE $1 OR friendly_name ILIKE $1 OR vendor_item_number ILIKE $1 OR barcode ILIKE $1
       )
       ORDER BY friendly_name LIMIT 20`,
      [`%${q}%`]
    );
    res.json(r.rows);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Search failed' });
  }
});

// Create a DRAFT PO from an invoice's reviewed line items.
// Only CONFIRMED and AUTO_MATCHED items are included. NEW_ITEM items are created
// in the Item Master first (same as the existing save-mappings flow).
app.post('/api/invoices/:id/create-po', requireAuthApi(['ADMIN', 'BUYER']), async (req, res) => {
  const { po_date, po_currency, exchange_rate } = req.body;
  const invoiceId = req.params.id;
  const client = await pool.connect();

  try {
    await client.query('BEGIN');

    // Get invoice
    const invRes = await client.query('SELECT * FROM vendor_invoices WHERE id = $1', [invoiceId]);
    if (invRes.rows.length === 0) {
      await client.query('ROLLBACK');
      return res.status(404).json({ error: 'Invoice not found' });
    }
    const invoice = invRes.rows[0];
    if (!['MATCHED', 'REVIEWED'].includes(invoice.status)) {
      await client.query('ROLLBACK');
      return res.status(400).json({ error: `Cannot create PO from invoice with status ${invoice.status}. Review the invoice first.` });
    }

    // Get all line items that are ready (CONFIRMED or AUTO_MATCHED)
    const itemsRes = await client.query(
      `SELECT * FROM invoice_line_items WHERE invoice_id = $1 AND match_status IN ('CONFIRMED', 'AUTO_MATCHED') AND matched_sku IS NOT NULL ORDER BY seq`,
      [invoiceId]
    );

    if (itemsRes.rows.length === 0) {
      await client.query('ROLLBACK');
      return res.status(400).json({ error: 'No confirmed line items to create a PO. Please review and confirm at least one item.' });
    }

    // Get vendor category for SKU generation
    const vendorRes = await client.query('SELECT vendor_code, category FROM vendors WHERE vendor_code = $1', [invoice.vendor_code]);
    const vendorCategory = vendorRes.rows[0]?.category || 'GN';

    let yearCode = 'A', collectionCode = 'PS', departmentCode = '1', colorCode = 'SLV', material = 'Glass';
    if (vendorCategory === 'JW') { material = 'Sterling Silver'; colorCode = 'SLV'; }
    if (vendorCategory === 'AP') { material = 'Cotton'; colorCode = 'BLK'; }
    if (vendorCategory === 'HB') { material = 'Leather'; colorCode = 'BLK'; }

    // Process NEW_ITEM items: create them in Item Master
    const newItemRes = await client.query(
      `SELECT * FROM invoice_line_items WHERE invoice_id = $1 AND match_status = 'NEW_ITEM' ORDER BY seq`,
      [invoiceId]
    );

    for (let i = 0; i < newItemRes.rows.length; i++) {
      const li = newItemRes.rows[i];
      const newSku = `${vendorCategory}${yearCode}${collectionCode}-${departmentCode}${((Date.now() + i) % 1000).toString().padStart(3, '0')}-${colorCode}`;

      const barcodeRes = await client.query('SELECT next_barcode FROM barcode_sequence WHERE id = 1 FOR UPDATE');
      let currentBarcode = barcodeRes.rows[0]?.next_barcode || '6941181218000';

      await client.query(
        `INSERT INTO item_master (sku, friendly_name, category_code, year_code, collection_code, department_code, color_code, hs_code, std_cost_rmb, description, material, barcode, vendor_item_number, created_by)
         VALUES ($1, $2, $3, $4, $5, $6, $7, '9505100090', $8, $9, $10, $11, $12, $13)
         ON CONFLICT (sku) DO NOTHING`,
        [newSku, li.description, vendorCategory, yearCode, collectionCode, departmentCode, colorCode, li.unit_price, `Auto-created from invoice ${invoice.invoice_number}`, material, currentBarcode, li.vendor_item_number, req.user.email]
      );

      await client.query('UPDATE barcode_sequence SET next_barcode = LPAD((CAST(next_barcode AS BIGINT) + 1)::TEXT, 13, \'0\') WHERE id = 1');
      await client.query('INSERT INTO inventory (sku, org_id, quantity_on_hand) VALUES ($1, 1, 0) ON CONFLICT DO NOTHING', [newSku]);

      // Update the line item with the new SKU
      await client.query('UPDATE invoice_line_items SET matched_sku = $1, match_status = $2 WHERE id = $3', [newSku, 'CONFIRMED', li.id]);
    }

    // Re-fetch all confirmed items (now includes newly created ones)
    const allItemsRes = await client.query(
      `SELECT * FROM invoice_line_items WHERE invoice_id = $1 AND match_status IN ('CONFIRMED', 'AUTO_MATCHED') AND matched_sku IS NOT NULL ORDER BY seq`,
      [invoiceId]
    );

    // Calculate totals
    const exchange_rate_to_rmb = parseFloat(exchange_rate) || 7.0;
    let total_usd = 0, total_rmb = 0;

    for (const item of allItemsRes.rows) {
      if (po_currency === 'RMB') {
        total_rmb += parseFloat(item.quantity) * parseFloat(item.unit_price);
        total_usd += (parseFloat(item.quantity) * parseFloat(item.unit_price)) / exchange_rate_to_rmb;
      } else {
        total_usd += parseFloat(item.quantity) * parseFloat(item.unit_price);
        total_rmb += parseFloat(item.quantity) * parseFloat(item.unit_price) * exchange_rate_to_rmb;
      }
    }

    const deposit_usd = total_usd * 0.30;
    const balance_usd = total_usd * 0.70;
    const po_id = `PO-${Date.now().toString().slice(-6)}`;

    // Create the PO
    await client.query(
      `INSERT INTO purchase_orders (po_id, vendor_code, po_date, invoice_currency, exchange_rate_to_rmb, status, total_rmb, total_usd, deposit_usd, balance_usd, created_by, invoice_reference, invoice_id)
       VALUES ($1, $2, $3, $4, $5, 'DRAFT', $6, $7, $8, $9, $10, $11, $12)`,
      [po_id, invoice.vendor_code, po_date || new Date().toISOString().split('T')[0], po_currency || invoice.currency, exchange_rate_to_rmb, total_rmb, total_usd, deposit_usd, balance_usd, req.user.email, invoice.invoice_number, invoiceId]
    );

    // Create PO line items
    for (const item of allItemsRes.rows) {
      const costRmb = (po_currency || invoice.currency) === 'RMB' ? parseFloat(item.unit_price) : parseFloat(item.unit_price) * exchange_rate_to_rmb;
      await client.query(
        `INSERT INTO po_line_items (po_id, sku, quantity, unit_price_foreign, unit_cost_rmb, hs_code)
         VALUES ($1, $2, $3, $4, $5, '9505100090')`,
        [po_id, item.matched_sku, parseFloat(item.quantity), parseFloat(item.unit_price), costRmb]
      );
    }

    // Mark invoice as PO_CREATED
    await client.query('UPDATE vendor_invoices SET status = $1, po_id = $2 WHERE id = $3', ['PO_CREATED', po_id, invoiceId]);

    // Mark all line items as CONFIRMED
    await client.query('UPDATE invoice_line_items SET match_status = $1 WHERE invoice_id = $2 AND match_status IN (\'CONFIRMED\', \'AUTO_MATCHED\')', ['CONFIRMED', invoiceId]);

    await client.query('COMMIT');

    logActivity(req.user, 'PO_CREATED_FROM_INVOICE', po_id, { invoice_id: invoiceId, vendor_code: invoice.vendor_code, total_usd: total_usd.toFixed(2), line_count: allItemsRes.rows.length });

    // Notify admins
    await notifyRole(pool, 'ADMIN', 'PO_FROM_INVOICE',
      `PO ${po_id} created from invoice`,
      `Auto-generated from invoice ${invoice.invoice_number} by ${req.user.name || req.user.email}. ${allItemsRes.rows.length} items, total $${total_usd.toFixed(2)}`,
      po_id
    );

    res.json({
      success: true,
      po_id,
      status: 'DRAFT',
      total_items: allItemsRes.rows.length,
      total_usd: total_usd.toFixed(2),
      total_rmb: total_rmb.toFixed(2),
      invoice_reference: invoice.invoice_number,
      invoice_id: invoiceId,
      message: `PO ${po_id} created with ${allItemsRes.rows.length} items. Submit it for approval when ready.`
    });
  } catch (err) {
    await client.query('ROLLBACK');
    console.error('Create PO from invoice error:', err);
    res.status(500).json({ error: 'Failed to create PO: ' + err.message });
  } finally {
    client.release();
  }
});

// Health check endpoint — useful for debugging Render deployment issues
app.get('/api/health', async (req, res) => {
  const checks = { server: 'ok', database: 'unknown', tables: {} };
  try {
    const result = await pool.query('SELECT 1 as ok');
    checks.database = result.rows.length > 0 ? 'ok' : 'error';
    // Check if Phase 1 columns exist
    try {
      await pool.query('SELECT version FROM purchase_orders LIMIT 1');
      checks.tables.po_version_column = 'exists';
    } catch (e) {
      checks.tables.po_version_column = 'MISSING';
    }
    try {
      await pool.query('SELECT 1 FROM notifications LIMIT 1');
      checks.tables.notifications = 'exists';
    } catch (e) {
      checks.tables.notifications = 'MISSING';
    }
    try {
      await pool.query('SELECT 1 FROM record_locks LIMIT 1');
      checks.tables.record_locks = 'exists';
    } catch (e) {
      checks.tables.record_locks = 'MISSING';
    }
  } catch (e) {
    checks.database = 'error: ' + e.message;
  }
  res.json(checks);
});

// ── Catalogue visitor tracking (WhatsApp-gated B2B catalogue access) ──
app.post('/api/register', async (req, res) => {
  try {
    const { whatsapp, company } = req.body;
    if (!whatsapp || whatsapp.trim().length < 5) {
      return res.status(400).json({ error: 'Please enter a valid WhatsApp number.' });
    }
    if (!company || company.trim().length < 1) {
      return res.status(400).json({ error: 'Please enter your company name.' });
    }
    const wa = whatsapp.trim();
    const co = company.trim();
    const existing = await pool.query('SELECT id, visit_count FROM catalogue_visitors WHERE whatsapp = $1', [wa]);
    if (existing.rows.length > 0) {
      await pool.query(
        'UPDATE catalogue_visitors SET company = $1, visit_count = visit_count + 1, created_at = NOW() WHERE id = $2',
        [co, existing.rows[0].id]
      );
    } else {
      await pool.query('INSERT INTO catalogue_visitors (whatsapp, company) VALUES ($1, $2)', [wa, co]);
    }
    console.log(`[CATALOGUE VISIT] ${co} — WhatsApp: ${wa}`);
    res.json({ success: true });
  } catch (err) {
    console.error('Register error:', err.message);
    res.status(500).json({ error: 'Server error' });
  }
});

app.get('/api/customers', async (req, res) => {
  try {
    const result = await pool.query(
      'SELECT id, whatsapp, company, created_at, visit_count FROM catalogue_visitors ORDER BY created_at DESC'
    );
    res.json({ count: result.rows.length, customers: result.rows });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/stats', async (req, res) => {
  try {
    const total = await pool.query('SELECT COUNT(*) as count FROM catalogue_visitors');
    const unique = await pool.query('SELECT COUNT(DISTINCT whatsapp) as count FROM catalogue_visitors');
    const totalVisits = await pool.query('SELECT SUM(visit_count) as count FROM catalogue_visitors');
    res.json({
      total: parseInt(total.rows[0].count),
      unique: parseInt(unique.rows[0].count),
      totalVisits: parseInt(totalVisits.rows[0].count) || 0
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── B2B Customer Auth ──
app.post('/api/b2b/register', async (req, res) => {
  try {
    const { email, password, company_name, contact_name, phone, address, city, province, postal_code, hst_gst_number } = req.body;
    const missing = [];
    if (!email) missing.push('email');
    if (!password || password.length < 6) missing.push('password (min 6 chars)');
    if (!company_name) missing.push('company_name');
    if (!contact_name) missing.push('contact_name');
    if (!phone) missing.push('phone');
    if (missing.length) return res.status(400).json({ error: `Missing: ${missing.join(', ')}` });

    // Check duplicate email
    const dup = await pool.query('SELECT id FROM users WHERE email = $1', [email]);
    if (dup.rows.length > 0) return res.status(409).json({ error: 'An account with this email already exists.' });

    const { salt, hash } = hashPassword(password);
    const client = await pool.connect();
    try {
      await client.query('BEGIN');
      const u = await client.query(
        'INSERT INTO users (email, password_hash, password_salt, role, name, is_active) VALUES ($1, $2, $3, $4, $5, true) RETURNING id',
        [email, hash, salt, 'B2B_CUSTOMER', contact_name]
      );
      await client.query(
        `INSERT INTO b2b_customers (user_id, company_name, contact_name, phone, address, city, province, postal_code, hst_gst_number)
         VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9)`,
        [u.rows[0].id, company_name, contact_name, phone, address || null, city || null, province || null, postal_code || null, hst_gst_number || null]
      );
      await client.query('COMMIT');
      logActivity({ email, role: 'B2B_CUSTOMER' }, 'B2B_REGISTER', null, { company_name, contact_name });
      console.log(`[B2B REGISTER] ${company_name} — ${email} (${contact_name})`);
      res.json({ success: true, message: 'Account created. You can now log in.' });
    } catch (e) {
      await client.query('ROLLBACK');
      throw e;
    } finally {
      client.release();
    }
  } catch (err) {
    console.error('B2B register error:', err.message);
    res.status(500).json({ error: 'Registration failed' });
  }
});

app.post('/api/b2b/login', async (req, res) => {
  try {
    const { email, password } = req.body;
    if (!email || !password) return res.status(400).json({ error: 'Email and password are required.' });
    const r = await pool.query("SELECT * FROM users WHERE email = $1 AND role = 'B2B_CUSTOMER' AND is_active = true", [email]);
    if (r.rows.length === 0) return res.status(401).json({ error: 'Invalid email or password.' });
    const u = r.rows[0];
    if (!verifyPassword(password, u.password_salt, u.password_hash)) {
      return res.status(401).json({ error: 'Invalid email or password.' });
    }
    // Load company info
    const co = await pool.query('SELECT * FROM b2b_customers WHERE user_id = $1', [u.id]);
    const sid = crypto.randomBytes(32).toString('hex');
    sessions.set(sid, {
      userId: u.id,
      email: u.email,
      role: u.role,
      name: u.name,
      company: co.rows.length ? co.rows[0].company_name : '',
      hst_gst: co.rows.length ? co.rows[0].hst_gst_number || '' : ''
    });
    res.setHeader('Set-Cookie', `sahi_session=${sid}; HttpOnly; Path=/; SameSite=Lax; Max-Age=${60 * 60 * 24 * 7}`);
    logActivity({ email: u.email, role: u.role }, 'B2B_LOGIN', null, null);
    res.json({ success: true, name: u.name, company: co.rows[0]?.company_name || '' });
  } catch (err) {
    console.error('B2B login error:', err.message);
    res.status(500).json({ error: 'Login failed' });
  }
});

app.get('/api/b2b/me', (req, res) => {
  const user = getSessionUser(req);
  if (!user || user.role !== 'B2B_CUSTOMER') return res.status(401).json({ error: 'Not authenticated' });
  res.json({ email: user.email, name: user.name, company: user.company || '', hst_gst: user.hst_gst || '' });
});

// ── B2B Ordering ──
app.post('/api/b2b/order', async (req, res) => {
  const user = getSessionUser(req);
  if (!user) return res.status(401).json({ error: 'Please log in to place an order.' });

  try {
    const { items, notes } = req.body;
    if (!items || !Array.isArray(items) || items.length === 0) {
      return res.status(400).json({ error: 'Cart is empty.' });
    }

    // Load full company details
    const co = await pool.query('SELECT * FROM b2b_customers WHERE user_id = $1', [user.userId]);
    const company = co.rows.length ? co.rows[0] : null;
    const companyName = company ? company.company_name : (user.company || user.name);
    const contactName = company ? company.contact_name : user.name;
    const phone = company ? company.phone : '';
    const hstGst = company ? company.hst_gst_number || '' : (user.hst_gst || '');
    const shipping = company
      ? [company.address, company.city, company.province, company.postal_code].filter(Boolean).join(', ')
      : '';

    let totalAmount = 0;
    let itemCount = 0;
    for (const item of items) {
      totalAmount += (item.unit_price || 0) * (item.quantity || 0);
      itemCount += item.quantity || 0;
    }

    // Calculate HST (13% for Ontario-based customers) and grand total
    const hstAmount = parseFloat((totalAmount * 0.13).toFixed(2));
    const grandTotal = parseFloat((totalAmount + hstAmount).toFixed(2));

    const client = await pool.connect();
    let orderId;
    try {
      await client.query('BEGIN');
      const ord = await client.query(
        `INSERT INTO b2b_orders (user_id, company_name, contact_name, email, phone, hst_gst_number, shipping_address, notes, total_amount, item_count, hst_amount, grand_total)
         VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12) RETURNING id`,
        [user.userId, companyName, contactName, user.email, phone, hstGst, shipping, notes || null, totalAmount, itemCount, hstAmount, grandTotal]
      );
      orderId = ord.rows[0].id;
      for (const item of items) {
        const lineTotal = (item.unit_price || 0) * (item.quantity || 0);
        await client.query(
          'INSERT INTO b2b_order_items (order_id, sku, product_name, quantity, unit_price, total) VALUES ($1, $2, $3, $4, $5, $6)',
          [orderId, item.sku, item.product_name, item.quantity, item.unit_price || 0, lineTotal]
        );
      }
      await client.query('COMMIT');
    } catch (e) {
      await client.query('ROLLBACK');
      throw e;
    } finally {
      client.release();
    }

    // Build order summary for email
    let itemsTable = items.map(i =>
      `${i.sku}\t${i.product_name}\t${i.quantity}\t$${Number(i.unit_price || 0).toFixed(2)}\t$${(Number(i.unit_price || 0) * i.quantity).toFixed(2)}`
    ).join('\n');

    const emailBody = `NEW B2B ORDER #${orderId}

Company: ${companyName}
Contact: ${contactName}
Email: ${user.email}
Phone: ${phone}
HST/GST: ${hstGst || 'N/A'}
Shipping: ${shipping || 'N/A'}
Notes: ${notes || 'None'}
Items: ${itemCount} | Subtotal: CAD $${totalAmount.toFixed(2)} | HST (13%): CAD $${hstAmount.toFixed(2)} | Grand Total: CAD $${grandTotal.toFixed(2)}

--- ORDER LINES ---
SKU\tProduct\tQty\tUnit Price\tLine Total
${itemsTable}

View in admin: https://sahi-mcp.onrender.com/index.html
`;

    // Generate PDF
    const orderForPdf = {
      id: orderId,
      company_name: companyName,
      contact_name: contactName,
      email: user.email,
      phone: phone,
      hst_gst_number: hstGst || '',
      shipping_address: shipping || '',
      notes: notes || '',
      total_amount: totalAmount,
      hst_amount: hstAmount,
      grand_total: grandTotal,
      item_count: itemCount,
      status: 'NEW',
      created_at: new Date().toISOString(),
      items: items.map(i => ({
        sku: i.sku,
        product_name: i.product_name,
        quantity: i.quantity,
        unit_price: Number(i.unit_price || 0),
        total: (Number(i.unit_price || 0) * i.quantity)
      }))
    };
    const pdfBuffer = await generateOrderPDF(orderForPdf);

    // Build CSV content
    const csvHeader = 'SKU,Product Name,Quantity,Unit Price (CAD),Line Total (CAD)';
    const csvRows = items.map(i => {
      const name = `"${(i.product_name || '').replace(/"/g, '""')}"`;
      return `${i.sku},${name},${i.quantity},${Number(i.unit_price||0).toFixed(2)},${(Number(i.unit_price||0)*i.quantity).toFixed(2)}`;
    });
    const csvContent = '\uFEFF' + [csvHeader, ...csvRows].join('\n'); // BOM for Excel

    // Send email to B2B-order@sahilondon.com
    try {
      if (EMAIL_PASS) {
        await mailer.sendMail({
          from: EMAIL_USER,
          to: ORDER_EMAIL,
          subject: `NEW B2B Order #${orderId} — ${companyName} (CAD $${totalAmount.toFixed(2)})`,
          text: emailBody,
          attachments: [
            {
              filename: `SAHI_B2B_Order_${orderId}.pdf`,
              content: pdfBuffer,
              contentType: 'application/pdf'
            },
            {
              filename: `SAHI_B2B_Order_${orderId}.csv`,
              content: csvContent,
              contentType: 'text/csv; charset=utf-8'
            }
          ]
        });
        console.log(`[B2B ORDER] #${orderId} emailed to ${ORDER_EMAIL} (with PDF)`);
      } else {
        console.log(`[B2B ORDER] #${orderId} — EMAIL SKIPPED (no EMAIL_PASS set)`);
        console.log(emailBody);
      }
    } catch (mailErr) {
      console.error(`[B2B ORDER] #${orderId} EMAIL FAILED — user: ${EMAIL_USER}, host: ${EMAIL_HOST}:${EMAIL_PORT}`);
      console.error('Error details:', mailErr.message, mailErr.code, mailErr.command || '');
      if (mailErr.response) console.error('Gmail response:', mailErr.response);
      console.error('Full error:', JSON.stringify(mailErr, Object.getOwnPropertyNames(mailErr)));
      console.log(emailBody);
    }

    logActivity({ email: user.email, role: 'B2B_CUSTOMER' }, 'B2B_ORDER', String(orderId), { total: totalAmount, items: itemCount });
    res.json({ success: true, orderId, total: totalAmount, message: `Order #${orderId} placed successfully!` });
  } catch (err) {
    console.error('B2B order error:', err.message);
    res.status(500).json({ error: 'Failed to place order.' });
  }
});

// View own orders
app.get('/api/b2b/orders', async (req, res) => {
  const user = getSessionUser(req);
  if (!user) return res.status(401).json({ error: 'Not authenticated' });

  try {
    const ord = await pool.query(
      'SELECT id, company_name, total_amount, item_count, status, notes, created_at FROM b2b_orders WHERE user_id = $1 ORDER BY created_at DESC LIMIT 50',
      [user.userId]
    );
    res.json(ord.rows);
  } catch (err) {
    res.status(500).json({ error: 'Failed to fetch orders' });
  }
});

// Admin: create a B2B order on behalf of a customer
app.post('/api/b2b/admin/create-order', requireAuthApi(['ADMIN']), async (req, res) => {
  try {
    const { company_name, contact_name, email, phone, hst_gst_number, shipping_address, notes, items } = req.body;
    if (!items || !Array.isArray(items) || items.length === 0) {
      return res.status(400).json({ error: 'Cart is empty.' });
    }
    if (!company_name || !contact_name) {
      return res.status(400).json({ error: 'Company name and contact name are required.' });
    }

    let totalAmount = 0;
    let itemCount = 0;
    for (const item of items) {
      totalAmount += (item.unit_price || 0) * (item.quantity || 0);
      itemCount += item.quantity || 0;
    }

    const hstAmount = parseFloat((totalAmount * 0.13).toFixed(2));
    const grandTotal = parseFloat((totalAmount + hstAmount).toFixed(2));

    const client = await pool.connect();
    let orderId;
    try {
      await client.query('BEGIN');
      const ord = await client.query(
        `INSERT INTO b2b_orders (user_id, company_name, contact_name, email, phone, hst_gst_number, shipping_address, notes, total_amount, item_count, hst_amount, grand_total, status)
         VALUES (NULL, $1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, 'DRAFT') RETURNING id`,
        [company_name, contact_name, email || '', phone || '', hst_gst_number || '', shipping_address || '', notes || null, totalAmount, itemCount, hstAmount, grandTotal]
      );
      orderId = ord.rows[0].id;
      for (const item of items) {
        const lineTotal = (item.unit_price || 0) * (item.quantity || 0);
        await client.query(
          'INSERT INTO b2b_order_items (order_id, sku, product_name, quantity, unit_price, total) VALUES ($1, $2, $3, $4, $5, $6)',
          [orderId, item.sku, item.product_name, item.quantity, item.unit_price || 0, lineTotal]
        );
      }
      await client.query('COMMIT');
    } catch (e) {
      await client.query('ROLLBACK');
      throw e;
    } finally {
      client.release();
    }

    res.json({ success: true, orderId, total: totalAmount, hst: hstAmount, grandTotal, itemCount });
  } catch (err) {
    console.error('Admin create order error:', err.message);
    res.status(500).json({ error: 'Failed to create order: ' + err.message });
  }
});

// Admin: view all B2B orders
app.get('/api/b2b/admin/orders', requireAuthApi(['ADMIN']), async (req, res) => {
  try {
    const ord = await pool.query(
      'SELECT o.*, array_agg(json_build_object(\'sku\', i.sku, \'product_name\', i.product_name, \'quantity\', i.quantity, \'unit_price\', i.unit_price, \'total\', i.total)) as items FROM b2b_orders o LEFT JOIN b2b_order_items i ON i.order_id = o.id GROUP BY o.id ORDER BY o.created_at DESC LIMIT 100'
    );
    res.json(ord.rows);
  } catch (err) {
    res.status(500).json({ error: 'Failed to fetch orders' });
  }
});

// Admin: soft-delete a B2B order
app.delete('/api/b2b/admin/orders/:id', requireAuthApi(['ADMIN']), async (req, res) => {
  try {
    const { id } = req.params;
    const result = await pool.query(
      'UPDATE b2b_orders SET deleted = TRUE WHERE id = $1 AND deleted = FALSE RETURNING *',
      [id]
    );
    if (result.rows.length === 0) return res.status(404).json({ error: 'Order not found or already deleted' });
    res.json({ success: true, message: 'Order marked as deleted' });
  } catch (err) {
    console.error('B2B order delete error:', err.message);
    res.status(500).json({ error: 'Failed to delete order' });
  }
});

// Admin: restore a soft-deleted B2B order
app.post('/api/b2b/admin/orders/:id/restore', requireAuthApi(['ADMIN']), async (req, res) => {
  try {
    const { id } = req.params;
    const result = await pool.query(
      'UPDATE b2b_orders SET deleted = FALSE WHERE id = $1 AND deleted = TRUE RETURNING *',
      [id]
    );
    if (result.rows.length === 0) return res.status(404).json({ error: 'Order not found or not deleted' });
    res.json({ success: true, message: 'Order restored' });
  } catch (err) {
    console.error('B2B order restore error:', err.message);
    res.status(500).json({ error: 'Failed to restore order' });
  }
});

// ===== B2B Customer Management =====

// List all customers
app.get('/api/b2b/admin/customers', requireAuthApi(['ADMIN']), async (req, res) => {
  try {
    const result = await pool.query('SELECT * FROM b2b_customers ORDER BY company_name');
    res.json(result.rows);
  } catch (err) {
    console.error('B2B customers list error:', err.message);
    res.status(500).json({ error: 'Failed to fetch customers' });
  }
});

// Get single customer
app.get('/api/b2b/admin/customers/:id', requireAuthApi(['ADMIN']), async (req, res) => {
  try {
    const result = await pool.query('SELECT * FROM b2b_customers WHERE id = $1', [req.params.id]);
    if (result.rows.length === 0) return res.status(404).json({ error: 'Customer not found' });
    res.json(result.rows[0]);
  } catch (err) {
    console.error('B2B customer get error:', err.message);
    res.status(500).json({ error: 'Failed to fetch customer' });
  }
});

// Upload business card photo
const cardUpload = multer({
  storage: multer.diskStorage({
    destination: path.join(__dirname, 'public', 'images', 'cards'),
    filename: (req, file, cb) => {
      cb(null, `card_${Date.now()}_${file.originalname.replace(/[^a-zA-Z0-9._-]/g, '_')}`);
    }
  }),
  limits: { fileSize: 10 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const ok = /jpeg|jpg|png|webp/.test(file.mimetype);
    cb(ok ? null : new Error('Only JPEG, PNG, or WebP images allowed'), ok);
  }
});

// Ensure cards directory exists
const cardsDir = path.join(__dirname, 'public', 'images', 'cards');
try { fs.mkdirSync(cardsDir, { recursive: true }); } catch (e) {}

app.post('/api/b2b/admin/customers/upload-card', requireAuthApi(['ADMIN']), cardUpload.single('card'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No image uploaded' });
    const imageUrl = `/images/cards/${req.file.filename}`;
    res.json({ success: true, card_image: imageUrl });
  } catch (err) {
    console.error('Card upload error:', err.message);
    res.status(500).json({ error: 'Failed to upload card: ' + err.message });
  }
});

// Create customer
app.post('/api/b2b/admin/customers', requireAuthApi(['ADMIN']), async (req, res) => {
  try {
    const { company_name, contact_name, email, phone, hst_gst_number, shipping_address, notes, card_image } = req.body;
    if (!contact_name) {
      return res.status(400).json({ error: 'Contact name is required.' });
    }
    if (!email && !phone) {
      return res.status(400).json({ error: 'Please provide at least an email or phone number.' });
    }
    const result = await pool.query(
      `INSERT INTO b2b_customers (company_name, contact_name, email, phone, hst_gst_number, shipping_address, notes, card_image)
       VALUES ($1, $2, $3, $4, $5, $6, $7, $8) RETURNING *`,
      [company_name || '', contact_name, email || '', phone || '', hst_gst_number || '', shipping_address || '', notes || null, card_image || '']
    );
    res.json(result.rows[0]);
  } catch (err) {
    console.error('B2B customer create error:', err.message);
    res.status(500).json({ error: 'Failed to create customer: ' + err.message });
  }
});

// Update customer
app.put('/api/b2b/admin/customers/:id', requireAuthApi(['ADMIN']), async (req, res) => {
  try {
    const { company_name, contact_name, email, phone, hst_gst_number, shipping_address, notes, card_image } = req.body;
    if (!contact_name) {
      return res.status(400).json({ error: 'Contact name is required.' });
    }
    if (!email && !phone) {
      return res.status(400).json({ error: 'Please provide at least an email or phone number.' });
    }
    const result = await pool.query(
      `UPDATE b2b_customers SET company_name=$1, contact_name=$2, email=$3, phone=$4, hst_gst_number=$5, shipping_address=$6, notes=$7, card_image=$8
       WHERE id=$9 RETURNING *`,
      [company_name || '', contact_name, email || '', phone || '', hst_gst_number || '', shipping_address || '', notes || null, card_image !== undefined ? card_image : '', req.params.id]
    );
    if (result.rows.length === 0) return res.status(404).json({ error: 'Customer not found' });
    res.json(result.rows[0]);
  } catch (err) {
    console.error('B2B customer update error:', err.message);
    res.status(500).json({ error: 'Failed to update customer' });
  }
});

// Delete customer
app.delete('/api/b2b/admin/customers/:id', requireAuthApi(['ADMIN']), async (req, res) => {
  try {
    const result = await pool.query('DELETE FROM b2b_customers WHERE id = $1 RETURNING id', [req.params.id]);
    if (result.rows.length === 0) return res.status(404).json({ error: 'Customer not found' });
    res.json({ success: true, message: 'Customer deleted' });
  } catch (err) {
    console.error('B2B customer delete error:', err.message);
    res.status(500).json({ error: 'Failed to delete customer' });
  }
});

// Admin: export B2B order as Shopify-compatible CSV
app.get('/api/b2b/admin/orders/:id/csv', requireAuthApi(['ADMIN']), async (req, res) => {
  try {
    const { id } = req.params;
    const ord = await pool.query(
      'SELECT * FROM b2b_orders WHERE id = $1',
      [id]
    );
    if (ord.rows.length === 0) return res.status(404).json({ error: 'Order not found' });
    const order = ord.rows[0];
    if (order.deleted) return res.status(403).json({ error: 'Cannot export CSV for a deleted order' });

    const items = await pool.query(
      'SELECT * FROM b2b_order_items WHERE order_id = $1 ORDER BY id',
      [id]
    );

    // Shopify-compatible CSV columns
    const header = [
      'SKU', 'Product Name', 'Quantity', 'Unit Price (CAD)', 'Line Total (CAD)',
      'Order ID', 'Company', 'Contact', 'Email', 'Phone',
      'HST/GST', 'Shipping Address', 'Notes', 'Order Date', 'Status'
    ];

    const rows = items.rows.map(it => [
      it.sku,
      `"${(it.product_name || '').replace(/"/g, '""')}"`,
      it.quantity,
      Number(it.unit_price).toFixed(2),
      Number(it.total).toFixed(2),
      order.id,
      `"${(order.company_name || '').replace(/"/g, '""')}"`,
      `"${(order.contact_name || '').replace(/"/g, '""')}"`,
      order.email,
      order.phone || '',
      order.hst_gst_number || '',
      `"${(order.shipping_address || '').replace(/"/g, '""')}"`,
      `"${(order.notes || '').replace(/"/g, '""')}"`,
      order.created_at ? new Date(order.created_at).toISOString().split('T')[0] : '',
      order.status
    ]);

    const csv = [header.join(','), ...rows.map(r => r.join(','))].join('\n');

    res.setHeader('Content-Type', 'text/csv; charset=utf-8');
    res.setHeader('Content-Disposition', `attachment; filename="B2B_Order_${id}_Shopify.csv"`);
    res.send('\uFEFF' + csv); // BOM for Excel UTF-8 compatibility
  } catch (err) {
    console.error('CSV export error:', err.message);
    res.status(500).json({ error: 'Failed to export CSV' });
  }
});

// Admin: download B2B order as PDF
app.get('/api/b2b/admin/orders/:id/pdf', requireAuthApi(['ADMIN']), async (req, res) => {
  try {
    const { id } = req.params;
    const ord = await pool.query('SELECT * FROM b2b_orders WHERE id = $1', [id]);
    if (ord.rows.length === 0) return res.status(404).json({ error: 'Order not found' });
    const order = ord.rows[0];
    if (order.deleted) return res.status(403).json({ error: 'Cannot download PDF for a deleted order' });

    const items = await pool.query('SELECT * FROM b2b_order_items WHERE order_id = $1 ORDER BY id', [id]);

    const pdfBuffer = await generateOrderPDF({
      ...order,
      items: items.rows
    });

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="SAHI_B2B_Order_${id}.pdf"`);
    res.send(pdfBuffer);
  } catch (err) {
    console.error('PDF export error:', err.message);
    res.status(500).json({ error: 'Failed to export PDF' });
  }
});

// Admin: publish a DRAFT B2B order (change status from DRAFT to NEW)
app.patch('/api/b2b/admin/orders/:id/publish', requireAuthApi(['ADMIN']), async (req, res) => {
  try {
    const { id } = req.params;

    const check = await pool.query('SELECT id, status, company_name, contact_name, email, phone, total_amount, item_count FROM b2b_orders WHERE id = $1', [id]);
    if (check.rows.length === 0) {
      return res.status(404).json({ error: 'Order not found' });
    }

    const order = check.rows[0];

    if (order.status !== 'DRAFT') {
      return res.status(400).json({ error: `Order #${id} is already published (status: ${order.status}). Only DRAFT orders can be published.` });
    }

    await pool.query("UPDATE b2b_orders SET status = 'NEW' WHERE id = $1", [id]);

    logActivity(req.user, 'B2B_ORDER_PUBLISHED', String(id), { company: order.company_name, total: order.total_amount });

    console.log(`[B2B ORDER] #${id} published by ${req.user.email || 'admin'} — status changed DRAFT → NEW`);

    res.json({
      success: true,
      orderId: id,
      status: 'NEW',
      message: `Order #${id} has been published successfully. It is now active.`
    });
  } catch (err) {
    console.error('Publish B2B order error:', err.message);
    res.status(500).json({ error: 'Failed to publish order' });
  }
});

// Admin: update B2B order status (NEW → PROCESSING → SHIPPED → COMPLETED)
app.patch('/api/b2b/admin/orders/:id/status', requireAuthApi(['ADMIN']), async (req, res) => {
  try {
    const { id } = req.params;
    const { status } = req.body;

    const validStatuses = ['NEW', 'PROCESSING', 'SHIPPED', 'COMPLETED', 'CANCELLED'];
    if (!validStatuses.includes(status)) {
      return res.status(400).json({ error: `Invalid status. Must be one of: ${validStatuses.join(', ')}` });
    }

    const check = await pool.query('SELECT id, status FROM b2b_orders WHERE id = $1', [id]);
    if (check.rows.length === 0) {
      return res.status(404).json({ error: 'Order not found' });
    }

    await pool.query('UPDATE b2b_orders SET status = $1 WHERE id = $2', [status, id]);
    logActivity(req.user, 'B2B_STATUS_UPDATE', String(id), { from: check.rows[0].status, to: status });

    res.json({ success: true, orderId: id, status, message: `Order #${id} status updated to ${status}` });
  } catch (err) {
    console.error('Status update error:', err.message);
    res.status(500).json({ error: 'Failed to update order status' });
  }
});

// ── Send Thank You Email to customers ──
app.post('/api/b2b/admin/send-thank-you', requireAuthApi(['ADMIN']), async (req, res) => {
  const { emails, customer_name } = req.body;
  
  if (!emails || !Array.isArray(emails) || emails.length === 0) {
    return res.status(400).json({ error: 'At least one email address is required' });
  }

  if (!EMAIL_PASS) {
    return res.status(500).json({ error: 'Email not configured (EMAIL_PASS missing)' });
  }

  const htmlContent = `<!DOCTYPE html>
<html lang="en">
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"></head>
<body style="margin:0;padding:0;background-color:#f4f4f4;font-family:Arial,Helvetica,sans-serif;">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:#f4f4f4;padding:20px 0;">
    <tr><td align="center">
      <table role="presentation" width="600" cellpadding="0" cellspacing="0" style="max-width:600px;background-color:#ffffff;">
        <tr><td style="background:#87CEEB;padding:35px 30px 30px;text-align:center;">
          <div style="font-size:32px;font-weight:bold;color:#000;letter-spacing:6px;margin-bottom:2px;">SAHI</div>
          <div style="font-size:11px;color:#000;letter-spacing:8px;text-transform:uppercase;margin-bottom:12px;">LONDON</div>
          <div style="font-size:12px;color:#000;font-style:italic;letter-spacing:1px;">Made by hand. Worn with intent.</div>
        </td></tr>
        <tr><td style="padding:35px 30px 20px;">
          <p style="font-size:16px;color:#000;margin:0 0 20px;">Dear ${customer_name || 'Valued Partner'},</p>
          <p style="font-size:14px;color:#000;line-height:1.7;margin:0 0 16px;">Thank you so much for visiting the SAHI stand at Toronto Market Week. It was a pleasure meeting you and sharing our latest collections with you.</p>
          <p style="font-size:14px;color:#000;line-height:1.7;margin:0 0 16px;">We truly appreciate the time you took to explore our pieces. Each SAHI creation is crafted with care, blending contemporary design with artisanal quality, and we would be delighted to welcome you into the SAHI family.</p>
        </td></tr>
        <tr><td style="padding:0 30px 20px;">
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#FFB7C5;border-radius:8px;">
            <tr><td style="padding:20px 25px;text-align:center;">
              <p style="font-size:15px;color:#000;margin:0;font-weight:bold;">Visit Our Online Catalogue</p>
              <p style="font-size:13px;color:#000;margin:6px 0 16px;">Explore our full range of apparel and accessories</p>
              <a href="https://sahi-mcp.onrender.com/catalogue.html?sales" target="_blank" rel="noopener" style="display:inline-block;background:#87CEEB;color:#000;font-size:14px;font-weight:bold;text-decoration:none;padding:12px 30px;border-radius:6px;">View Catalogue</a>
            </td></tr>
          </table>
        </td></tr>
        <tr><td style="padding:10px 30px 25px;">
          <p style="font-size:14px;color:#000;line-height:1.7;margin:0 0 12px;">If you have any questions about our products, pricing, or would like to place an order, please do not hesitate to reach out. We are here to help.</p>
          <p style="font-size:14px;color:#000;line-height:1.7;margin:0 0 8px;">Warm regards,</p>
          <p style="font-size:14px;color:#000;margin:0;font-weight:bold;">The SAHI London Team</p>
        </td></tr>
        <tr><td style="background:#f4f4f4;padding:25px 30px;">
          <p style="font-size:13px;color:#000;margin:0 0 15px;font-weight:bold;text-align:center;">Connect With Us</p>
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0">
            <tr>
              <td width="50%" style="padding:0 8px;">
                <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:8px;">
                  <tr><td style="padding:15px;text-align:center;">
                    <div style="font-size:28px;margin-bottom:5px;">&#128247;</div>
                    <div style="font-size:11px;color:#666;margin-bottom:3px;">Instagram</div>
                    <a href="https://www.instagram.com/sahi_london/" target="_blank" rel="noopener" style="font-size:14px;color:#000;font-weight:bold;text-decoration:none;">@sahi_london</a>
                  </td></tr>
                </table>
              </td>
              <td width="50%" style="padding:0 8px;">
                <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:8px;">
                  <tr><td style="padding:15px;text-align:center;">
                    <div style="font-size:28px;margin-bottom:5px;">&#128172;</div>
                    <div style="font-size:11px;color:#666;margin-bottom:3px;">WhatsApp</div>
                    <a href="https://wa.me/447518130383" target="_blank" rel="noopener" style="font-size:14px;color:#000;font-weight:bold;text-decoration:none;">+44 7518 130383</a>
                  </td></tr>
                </table>
              </td>
            </tr>
          </table>
        </td></tr>
        <tr><td style="background:#87CEEB;padding:20px 30px;text-align:center;">
          <p style="font-size:14px;color:#000;margin:0 0 4px;font-weight:bold;letter-spacing:2px;">SAHI LONDON</p>
          <p style="font-size:11px;color:#000;margin:0;line-height:1.6;">1231 Niagara On The Lake, Ontario L0S 1J0, Canada<br>Instagram: <a href="https://www.instagram.com/sahi_london/" target="_blank" rel="noopener" style="color:#000;font-weight:bold;text-decoration:none;">@sahi_london</a> &middot; WhatsApp: <a href="https://wa.me/447518130383" target="_blank" rel="noopener" style="color:#000;font-weight:bold;text-decoration:none;">+44 7518 130383</a></p>
        </td></tr>
      </table>
    </td></tr>
  </table>
</body>
</html>`;

  const results = { sent: [], failed: [] };

  for (const email of emails) {
    try {
      await mailer.sendMail({
        from: EMAIL_USER,
        to: email,
        cc: 'info@sahilondon.com',
        subject: 'Thank You from SAHI London — Toronto Market Week',
        html: htmlContent
      });
      results.sent.push(email);
      console.log(`[THANK YOU EMAIL] Sent to ${email} (CC: info@sahilondon.com)`);
    } catch (err) {
      results.failed.push({ email, error: err.message });
      console.error(`[THANK YOU EMAIL] Failed for ${email}:`, err.message);
    }
  }

  logActivity(req.user, 'THANK_YOU_EMAIL_SENT', null, { sent: results.sent.length, failed: results.failed.length });

  res.json({
    success: true,
    sent: results.sent.length,
    failed: results.failed.length,
    details: results
  });
});

// ════════════════════════════════════════════════════════════════
// BULK ITEM IMPORT + OPENING INVENTORY
// ════════════════════════════════════════════════════════════════

// POST /api/items/bulk — bulk create items from Excel
app.post('/api/items/bulk', requireAuthApi(['ADMIN', 'BUYER']), async (req, res) => {
  const { items } = req.body;
  if (!items || !Array.isArray(items) || items.length === 0) {
    return res.status(400).json({ error: 'No items provided' });
  }

  let imported = 0;
  let skipped = 0;
  let errors = 0;
  const details = [];

  for (const item of items) {
    try {
      const existing = await pool.query('SELECT sku FROM item_master WHERE sku = $1', [item.sku]);
      if (existing.rows.length > 0) {
        skipped++;
        details.push({ sku: item.sku, status: 'skipped', reason: 'SKU already exists' });
        continue;
      }

      await pool.query(
        `INSERT INTO item_master (sku, barcode, friendly_name, collection, category, original_qty, balance_qty, status, created_by)
         VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9)
         ON CONFLICT (sku) DO NOTHING`,
        [
          item.sku,
          item.barcode || null,
          item.friendly_name || item.sku,
          item.collection || null,
          item.category || null,
          item.original_qty || item.qty || 0,
          item.balance || item.qty || 0,
          item.status || 'PENDING_IMAGE',
          req.user ? req.user.email : 'system'
        ]
      );

      // Also create an inventory row so the item shows up in stock views
      await pool.query(
        `INSERT INTO inventory (sku, org_id, quantity_on_hand) VALUES ($1, 1, $2) ON CONFLICT DO NOTHING`,
        [item.sku, item.balance || item.qty || 0]
      );

      imported++;
      details.push({ sku: item.sku, status: 'created' });
    } catch (err) {
      errors++;
      details.push({ sku: item.sku, status: 'error', error: err.message });
    }
  }

  logActivity(req.user, 'BULK_ITEM_IMPORT', null, { imported, skipped, errors, total: items.length });

  res.json({ success: true, imported, skipped, errors, details });
});

// GET /api/opening-inventory — list all saved opening inventory entries
app.get('/api/opening-inventory', requireAuthApi(['ADMIN', 'BUYER']), async (req, res) => {
  try {
    const result = await pool.query(`
      SELECT oi.*, im.friendly_name, im.color
      FROM opening_inventory oi
      LEFT JOIN item_master im ON oi.sku = im.sku
      ORDER BY oi.created_at DESC
    `);
    res.json(result.rows);
  } catch (err) {
    console.error('Error fetching opening inventory:', err);
    res.status(500).json({ error: err.message });
  }
});

// POST /api/opening-inventory — save a batch of scanned items as opening stock
app.post('/api/opening-inventory', requireAuthApi(['ADMIN', 'BUYER']), async (req, res) => {
  const { items, org_id } = req.body;
  if (!items || !Array.isArray(items) || items.length === 0) {
    return res.status(400).json({ error: 'No items provided' });
  }
  const targetOrg = org_id || 1;
  const createdBy = req.user ? req.user.email : 'system';

  try {
    let saved = 0;
    let totalQty = 0;

    for (const item of items) {
      const existing = await pool.query(
        'SELECT id, qty FROM opening_inventory WHERE sku = $1 AND org_id = $2',
        [item.sku, targetOrg]
      );

      if (existing.rows.length > 0) {
        await pool.query(
          'UPDATE opening_inventory SET qty = qty + $1, created_by = $2, created_at = NOW() WHERE id = $3',
          [item.qty, createdBy, existing.rows[0].id]
        );
      } else {
        await pool.query(
          `INSERT INTO opening_inventory (sku, barcode, friendly_name, color, qty, org_id, created_by)
           VALUES ($1, $2, $3, $4, $5, $6, $7)`,
          [item.sku, item.barcode || null, item.friendly_name || null, item.color || null, item.qty, targetOrg, createdBy]
        );
      }
      saved++;
      totalQty += item.qty;
    }

    logActivity(req.user, 'OPENING_INVENTORY_SAVE', null, { saved, totalQty, org_id: targetOrg });

    res.json({ success: true, saved, total_qty: totalQty });
  } catch (err) {
    console.error('Error saving opening inventory:', err);
    res.status(500).json({ error: err.message });
  }
});

// GET /api/opening-inventory/summary — summary by org
app.get('/api/opening-inventory/summary', requireAuthApi(['ADMIN', 'BUYER']), async (req, res) => {
  try {
    const result = await pool.query(`
      SELECT org_id, COUNT(*) as item_count, SUM(qty) as total_qty
      FROM opening_inventory
      GROUP BY org_id
      ORDER BY org_id
    `);
    res.json(result.rows);
  } catch (err) {
    console.error('Error fetching opening inventory summary:', err);
    res.status(500).json({ error: err.message });
  }
});

// PUT /api/items/:sku/barcode — update barcode for existing item
app.put('/api/items/:sku/barcode', requireAuthApi(['ADMIN', 'BUYER']), async (req, res) => {
  const { sku } = req.params;
  const { barcode } = req.body;
  if (!barcode) return res.status(400).json({ error: 'Barcode required' });

  try {
    const result = await pool.query(
      'UPDATE item_master SET barcode = $1 WHERE sku = $2 RETURNING *',
      [barcode, sku]
    );
    if (result.rows.length === 0) {
      return res.status(404).json({ error: 'Item not found' });
    }
    res.json({ success: true, item: result.rows[0] });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`🚀 Portal running on port ${PORT}`));
