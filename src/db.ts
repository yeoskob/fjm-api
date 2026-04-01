import Database from 'better-sqlite3';
import { generateId } from './utils/id';


export const db = new Database('fjm.db');

db.pragma('foreign_keys = ON');
db.pragma('journal_mode = WAL');
db.pragma('synchronous = NORMAL');
db.pragma('temp_store = MEMORY');

db.exec(`
  CREATE TABLE IF NOT EXISTS users (
    id TEXT PRIMARY KEY,
    name TEXT NOT NULL,
    username TEXT UNIQUE NOT NULL,
    password TEXT NOT NULL,
    role TEXT NOT NULL
  );

  CREATE TABLE IF NOT EXISTS inquiries (
    id TEXT PRIMARY KEY,
    rfq_no TEXT,
    tanggal TEXT NOT NULL,
    customer TEXT NOT NULL,
    sales_pic TEXT NOT NULL,
    nama_barang TEXT NOT NULL,
    spesifikasi TEXT,
    qty REAL,
    target_price REAL,
    deadline_quotation TEXT,
    lampiran TEXT,
    supplier TEXT,
    harga_beli REAL,
    lead_time TEXT,
    moq REAL,
    stock_availability TEXT,
    term_pembayaran TEXT,
    harga_jual REAL,
    margin REAL,
    lead_time_customer TEXT,
    validitas_quotation TEXT,
    catatan_quotation TEXT,
    status TEXT NOT NULL DEFAULT 'new_inquiry',
    created_at TEXT NOT NULL,
    created_by TEXT NOT NULL,
    updated_at TEXT,
    updated_by TEXT
  );

  CREATE TABLE IF NOT EXISTS inquiry_items (
    id TEXT PRIMARY KEY,
    inquiry_id TEXT NOT NULL,
    coupa_row_index INTEGER,
    lot_id TEXT,
    lot_name TEXT,
    lot_expected_quantity REAL,
    lot_quantity_note TEXT,
    coupa_item_id TEXT,
    item_name TEXT,
    item_quantity REAL,
    item_uom TEXT,
    item_need_by_date TEXT,
    item_manufacturer_name TEXT,
    item_manufacturer_part_number TEXT,
    item_classification_of_goods TEXT,
    item_extended_description TEXT,
    item_fiscal_code TEXT,
    coupa_bid_id TEXT,
    bid_capacity REAL,
    bid_price_amount REAL,
    bid_price_currency TEXT,
    bid_lead_time TEXT,
    bid_supplier_item_name TEXT,
    bid_item_part_number TEXT,
    bid_item_description TEXT,
    bid_shipping_term TEXT,
    target_price REAL,
    supplier TEXT,
    harga_beli REAL,
    lead_time TEXT,
    moq REAL,
    stock_availability TEXT,
    term_pembayaran TEXT,
    harga_jual REAL,
    margin REAL,
    lead_time_customer TEXT,
    validitas_quotation TEXT,
    catatan_quotation TEXT,
    FOREIGN KEY(inquiry_id) REFERENCES inquiries(id) ON DELETE CASCADE
  );

  CREATE TABLE IF NOT EXISTS coupa_files (
    inquiry_id TEXT PRIMARY KEY,
    file_name TEXT NOT NULL,
    file_data BLOB NOT NULL,
    created_at TEXT NOT NULL,
    FOREIGN KEY(inquiry_id) REFERENCES inquiries(id) ON DELETE CASCADE
  );

  CREATE TABLE IF NOT EXISTS activity_log (
    id TEXT PRIMARY KEY,
    inquiry_id TEXT NOT NULL,
    action TEXT NOT NULL,
    old_status TEXT,
    new_status TEXT,
    note TEXT,
    done_by TEXT NOT NULL,
    done_by_name TEXT NOT NULL,
    created_at TEXT NOT NULL,
    FOREIGN KEY(inquiry_id) REFERENCES inquiries(id) ON DELETE CASCADE
  );

  CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL
  );

  CREATE TABLE IF NOT EXISTS inquiry_notes (
    id TEXT PRIMARY KEY,
    inquiry_id TEXT NOT NULL,
    note TEXT NOT NULL,
    created_by TEXT NOT NULL,
    created_by_name TEXT NOT NULL,
    created_at TEXT NOT NULL,
    FOREIGN KEY(inquiry_id) REFERENCES inquiries(id) ON DELETE CASCADE
  );

  CREATE TABLE IF NOT EXISTS products (
    id TEXT PRIMARY KEY,
    name TEXT NOT NULL,
    image_data_url TEXT,
    proposed_price REAL NOT NULL,
    approved_price REAL,
    approved_source_id TEXT,
    lead_time_minutes INTEGER NOT NULL,
    status TEXT NOT NULL,
    created_at TEXT NOT NULL,
    created_by TEXT NOT NULL,
    approved_at TEXT
  );

  CREATE TABLE IF NOT EXISTS product_sources (
    id TEXT PRIMARY KEY,
    product_id TEXT NOT NULL,
    label TEXT,
    url TEXT,
    price REAL,
    FOREIGN KEY(product_id) REFERENCES products(id) ON DELETE CASCADE
  );
`);

// Indexes to speed up common lookups (notes/comments, items, activity, dashboards)
db.exec(`
  CREATE INDEX IF NOT EXISTS idx_inquiry_items_inquiry_id ON inquiry_items (inquiry_id);
  CREATE INDEX IF NOT EXISTS idx_activity_log_inquiry_id ON activity_log (inquiry_id);
  CREATE INDEX IF NOT EXISTS idx_inquiry_notes_inquiry_id ON inquiry_notes (inquiry_id);
  CREATE INDEX IF NOT EXISTS idx_inquiry_notes_inquiry_item_id ON inquiry_notes (inquiry_id, item_id);
  CREATE INDEX IF NOT EXISTS idx_inquiries_created_at ON inquiries (created_at);
  CREATE INDEX IF NOT EXISTS idx_inquiries_status ON inquiries (status);
  CREATE INDEX IF NOT EXISTS idx_inquiries_sales_pic ON inquiries (sales_pic);
`);

const ensureProductsSchema = () => {
  const columns = db.prepare("PRAGMA table_info('products')").all() as Array<{ name: string }>;
  const hasApprovedSource = columns.some((c) => c.name === 'approved_source_id');
  if (!hasApprovedSource) {
    db.exec('ALTER TABLE products ADD COLUMN approved_source_id TEXT');
  }
};

const ensureInquiriesCoupaColumns = () => {
  const columns = db.prepare("PRAGMA table_info('inquiries')").all() as Array<{ name: string }>;
  const names = columns.map((c) => c.name);
  if (!names.includes('coupa_source')) {
    db.exec('ALTER TABLE inquiries ADD COLUMN coupa_source INTEGER NOT NULL DEFAULT 0');
  }
  if (!names.includes('coupa_file_name')) {
    db.exec('ALTER TABLE inquiries ADD COLUMN coupa_file_name TEXT');
  }
};

const seedUsers = () => {
  const count = db.prepare('SELECT COUNT(*) as count FROM users').get() as { count: number };
  if (count.count > 0) {
    return;
  }

  const insert = db.prepare(
    'INSERT INTO users (id, name, username, password, role) VALUES (@id, @name, @username, @password, @role)'
  );

  const seed = [
    { id: generateId(), name: 'Administrator', username: 'admin', password: 'admin', role: 'admin' },
    { id: generateId(), name: 'Sourcing User', username: 'sourcing', password: 'sourcing', role: 'sourcing' },
    { id: generateId(), name: 'Marketing User', username: 'marketing', password: 'marketing', role: 'marketing' },
    { id: generateId(), name: 'Sales User', username: 'sales', password: 'sales', role: 'marketing' },
  ];

  const tx = db.transaction((rows) => {
    for (const row of rows) {
      insert.run(row);
    }
  });

  tx(seed);
};

const normalizeSalesRole = () => {
  db.prepare("UPDATE users SET role = 'marketing' WHERE role = 'sales'").run();
};

const ensureInquiryItemsFromExisting = () => {
  const itemCount = db.prepare('SELECT COUNT(*) as count FROM inquiry_items').get() as { count: number };
  if (itemCount.count > 0) {
    return;
  }

  const inquiries = db.prepare('SELECT * FROM inquiries').all() as Array<Record<string, unknown>>;
  if (inquiries.length === 0) {
    return;
  }

  const insert = db.prepare(
    `INSERT INTO inquiry_items (
      id, inquiry_id, item_name, item_quantity, item_extended_description, target_price, item_need_by_date,
      supplier, harga_beli, lead_time, moq, stock_availability, term_pembayaran, harga_jual, margin,
      lead_time_customer, validitas_quotation, catatan_quotation
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
  );

  const tx = db.transaction((rows: Array<Record<string, unknown>>) => {
    for (const row of rows) {
      insert.run(
        generateId(),
        row['id'],
        row['nama_barang'] ?? null,
        row['qty'] ?? null,
        row['spesifikasi'] ?? null,
        row['target_price'] ?? null,
        row['deadline_quotation'] ?? null,
        row['supplier'] ?? null,
        row['harga_beli'] ?? null,
        row['lead_time'] ?? null,
        row['moq'] ?? null,
        row['stock_availability'] ?? null,
        row['term_pembayaran'] ?? null,
        row['harga_jual'] ?? null,
        row['margin'] ?? null,
        row['lead_time_customer'] ?? null,
        row['validitas_quotation'] ?? null,
        row['catatan_quotation'] ?? null
      );
    }
  });

  tx(inquiries);
};

const ensureItemImageColumn = () => {
  const columns = db.prepare("PRAGMA table_info('inquiry_items')").all() as Array<{ name: string }>;
  if (!columns.some((c) => c.name === 'item_image')) {
    db.exec('ALTER TABLE inquiry_items ADD COLUMN item_image TEXT');
  }
};

const ensureSessionsTable = () => {
  db.exec(`
    CREATE TABLE IF NOT EXISTS sessions (
      token TEXT PRIMARY KEY,
      user_id TEXT NOT NULL,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `);
};

const ensureSourcingPicColumn = () => {
  const columns = db.prepare("PRAGMA table_info('inquiries')").all() as Array<{ name: string }>;
  if (!columns.some((c) => c.name === 'sourcing_pic')) {
    db.exec('ALTER TABLE inquiries ADD COLUMN sourcing_pic TEXT');
  }
};

const ensureNotesItemIdColumn = () => {
  const columns = db.prepare("PRAGMA table_info('inquiry_notes')").all() as Array<{ name: string }>;
  if (!columns.some((c) => c.name === 'item_id')) {
    db.exec('ALTER TABLE inquiry_notes ADD COLUMN item_id TEXT');
  }
};

const ensurePriceApprovedColumn = () => {
  const columns = db.prepare("PRAGMA table_info('inquiry_items')").all() as Array<{ name: string }>;
  if (!columns.some((c) => c.name === 'price_approved')) {
    db.exec('ALTER TABLE inquiry_items ADD COLUMN price_approved INTEGER NOT NULL DEFAULT 0');
  }
  if (!columns.some((c) => c.name === 'approved_price')) {
    db.exec('ALTER TABLE inquiry_items ADD COLUMN approved_price REAL');
  }
};

ensureProductsSchema();
ensureInquiriesCoupaColumns();
ensureInquiryItemsFromExisting();
ensureItemImageColumn();
ensureSessionsTable();
ensureSourcingPicColumn();
ensureNotesItemIdColumn();
ensurePriceApprovedColumn();
normalizeSalesRole();
seedUsers();

// Seed default settings if not present
db.prepare(`INSERT OR IGNORE INTO settings (key, value) VALUES ('default_margin_pct', '20')`).run();

