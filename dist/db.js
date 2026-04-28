"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.db = void 0;
const better_sqlite3_1 = __importDefault(require("better-sqlite3"));
const path_1 = __importDefault(require("path"));
const id_1 = require("./utils/id");
const DB_PATH = process.env['DB_PATH'] ?? path_1.default.resolve(__dirname, '../fjm.db');
exports.db = new better_sqlite3_1.default(DB_PATH);
exports.db.pragma('foreign_keys = ON');
exports.db.pragma('journal_mode = WAL');
exports.db.pragma('synchronous = NORMAL');
exports.db.pragma('temp_store = MEMORY');
exports.db.pragma('wal_autocheckpoint = 1000'); // checkpoint every 1000 pages (~4 MB)
// Checkpoint any leftover WAL from previous run into the main DB on startup
exports.db.pragma('wal_checkpoint(TRUNCATE)');
// Detect fresh install BEFORE running any CREATE TABLE, so we can skip
// already-applied migrations on existing databases.
const isFreshInstall = (() => {
    try {
        exports.db.prepare('SELECT 1 FROM inquiries LIMIT 1').get();
        return false;
    }
    catch {
        return true;
    }
})();
// ─── Base schema (fresh installs only — existing DBs already have these) ────
exports.db.exec(`
  CREATE TABLE IF NOT EXISTS schema_version (
    version INTEGER NOT NULL DEFAULT 0
  );

  CREATE TABLE IF NOT EXISTS users (
    id TEXT PRIMARY KEY,
    name TEXT NOT NULL,
    username TEXT UNIQUE NOT NULL,
    password TEXT NOT NULL,
    role TEXT NOT NULL
  );

  CREATE TABLE IF NOT EXISTS roles (
    name TEXT PRIMARY KEY,
    menus TEXT NOT NULL,
    tabs TEXT NOT NULL DEFAULT '{}',
    created_at TEXT NOT NULL
  );

  CREATE TABLE IF NOT EXISTS inquiries (
    id TEXT PRIMARY KEY,
    rfq_no TEXT,
    tanggal TEXT NOT NULL,
    customer TEXT NOT NULL,
    sales_pic TEXT NOT NULL,
    sourcing_pic TEXT,
    coupa_source INTEGER NOT NULL DEFAULT 0,
    organization TEXT NOT NULL DEFAULT 'FJM',
    coupa_file_name TEXT,
    nama_barang TEXT,
    spesifikasi TEXT,
    qty REAL,
    target_price REAL,
    deadline_quotation TEXT,
    lampiran TEXT,
    status TEXT NOT NULL DEFAULT 'new_inquiry',
    sent_incomplete INTEGER NOT NULL DEFAULT 0,
    sent_incomplete_reason TEXT,
    created_at TEXT NOT NULL,
    created_by TEXT NOT NULL,
    updated_at TEXT,
    updated_by TEXT,
    price_approval_started_at TEXT
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
    item_image TEXT,
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
    price_approved INTEGER NOT NULL DEFAULT 0,
    needs_price_review INTEGER NOT NULL DEFAULT 0,
    review_status TEXT NOT NULL DEFAULT 'pending',
    review_round INTEGER NOT NULL DEFAULT 0,
    approved_price REAL,
    alternate_name TEXT,
    ppn_type TEXT,
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

  CREATE TABLE IF NOT EXISTS organizations (
    id TEXT PRIMARY KEY,
    code TEXT UNIQUE NOT NULL,
    created_at TEXT NOT NULL,
    created_by TEXT
  );

  CREATE TABLE IF NOT EXISTS notifications (
    id TEXT PRIMARY KEY,
    type TEXT NOT NULL,
    inquiry_id TEXT NOT NULL,
    rfq_no TEXT,
    message TEXT NOT NULL,
    triggered_by TEXT NOT NULL,
    triggered_by_name TEXT NOT NULL,
    created_at TEXT NOT NULL,
    read_at TEXT
  );

  CREATE TABLE IF NOT EXISTS inquiry_notes (
    id TEXT PRIMARY KEY,
    inquiry_id TEXT NOT NULL,
    item_id TEXT,
    note TEXT NOT NULL,
    created_by TEXT NOT NULL,
    created_by_name TEXT NOT NULL,
    created_at TEXT NOT NULL,
    FOREIGN KEY(inquiry_id) REFERENCES inquiries(id) ON DELETE CASCADE
  );

  CREATE TABLE IF NOT EXISTS sessions (
    token TEXT PRIMARY KEY,
    user_id TEXT NOT NULL,
    created_at TEXT NOT NULL DEFAULT (datetime('now'))
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
// ─── Indexes ─────────────────────────────────────────────────────────────────
exports.db.exec(`
  CREATE INDEX IF NOT EXISTS idx_inquiry_items_inquiry_id    ON inquiry_items  (inquiry_id);
  CREATE INDEX IF NOT EXISTS idx_activity_log_inquiry_id     ON activity_log   (inquiry_id);
  CREATE INDEX IF NOT EXISTS idx_inquiry_notes_inquiry_id    ON inquiry_notes  (inquiry_id);
  CREATE INDEX IF NOT EXISTS idx_inquiry_notes_item_id       ON inquiry_notes  (inquiry_id, item_id);
  CREATE INDEX IF NOT EXISTS idx_inquiries_created_at        ON inquiries      (created_at);
  CREATE INDEX IF NOT EXISTS idx_inquiries_status            ON inquiries      (status);
  CREATE INDEX IF NOT EXISTS idx_inquiries_sales_pic         ON inquiries      (sales_pic);
`);
// ─── Migrations ───────────────────────────────────────────────────────────────
// Each migration runs exactly once, identified by its version number.
// Fresh installs jump straight to LATEST_VERSION (all columns already in CREATE TABLE above).
// Existing DBs run only the migrations they haven't seen yet.
const LATEST_VERSION = 20;
const cols = (table) => exports.db.prepare(`PRAGMA table_info('${table}')`).all().map((c) => c.name);
const migrations = [
    {
        // Add Coupa columns to inquiries
        version: 1,
        run: () => {
            if (!cols('inquiries').includes('coupa_source'))
                exports.db.exec('ALTER TABLE inquiries ADD COLUMN coupa_source INTEGER NOT NULL DEFAULT 0');
            if (!cols('inquiries').includes('coupa_file_name'))
                exports.db.exec('ALTER TABLE inquiries ADD COLUMN coupa_file_name TEXT');
        },
    },
    {
        // Migrate single-item inquiries table into inquiry_items
        version: 2,
        run: () => {
            const itemCount = exports.db.prepare('SELECT COUNT(*) as count FROM inquiry_items').get().count;
            if (itemCount > 0)
                return;
            const inquiries = exports.db.prepare('SELECT * FROM inquiries').all();
            if (inquiries.length === 0)
                return;
            const insert = exports.db.prepare(`INSERT INTO inquiry_items (
          id, inquiry_id, item_name, item_quantity, item_extended_description,
          target_price, item_need_by_date, supplier, harga_beli, lead_time, moq,
          stock_availability, term_pembayaran, harga_jual, margin,
          lead_time_customer, validitas_quotation, catatan_quotation
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`);
            const tx = exports.db.transaction((rows) => {
                for (const row of rows) {
                    insert.run((0, id_1.generateId)(), row['id'], row['nama_barang'] ?? null, row['qty'] ?? null, row['spesifikasi'] ?? null, row['target_price'] ?? null, row['deadline_quotation'] ?? null, row['supplier'] ?? null, row['harga_beli'] ?? null, row['lead_time'] ?? null, row['moq'] ?? null, row['stock_availability'] ?? null, row['term_pembayaran'] ?? null, row['harga_jual'] ?? null, row['margin'] ?? null, row['lead_time_customer'] ?? null, row['validitas_quotation'] ?? null, row['catatan_quotation'] ?? null);
                }
            });
            tx(inquiries);
        },
    },
    {
        // Add item_image to inquiry_items
        version: 3,
        run: () => {
            if (!cols('inquiry_items').includes('item_image'))
                exports.db.exec('ALTER TABLE inquiry_items ADD COLUMN item_image TEXT');
        },
    },
    {
        // Create sessions table
        version: 4,
        run: () => {
            exports.db.exec(`
        CREATE TABLE IF NOT EXISTS sessions (
          token TEXT PRIMARY KEY,
          user_id TEXT NOT NULL,
          created_at TEXT NOT NULL DEFAULT (datetime('now'))
        )
      `);
        },
    },
    {
        // Add sourcing_pic to inquiries
        version: 5,
        run: () => {
            if (!cols('inquiries').includes('sourcing_pic'))
                exports.db.exec('ALTER TABLE inquiries ADD COLUMN sourcing_pic TEXT');
        },
    },
    {
        // Add item_id to inquiry_notes
        version: 6,
        run: () => {
            if (!cols('inquiry_notes').includes('item_id'))
                exports.db.exec('ALTER TABLE inquiry_notes ADD COLUMN item_id TEXT');
        },
    },
    {
        // Add price approval columns to inquiry_items
        version: 7,
        run: () => {
            const c = cols('inquiry_items');
            if (!c.includes('price_approved'))
                exports.db.exec('ALTER TABLE inquiry_items ADD COLUMN price_approved INTEGER NOT NULL DEFAULT 0');
            if (!c.includes('approved_price'))
                exports.db.exec('ALTER TABLE inquiry_items ADD COLUMN approved_price REAL');
            if (!c.includes('alternate_name'))
                exports.db.exec('ALTER TABLE inquiry_items ADD COLUMN alternate_name TEXT');
        },
    },
    {
        // Add approved_source_id to products
        version: 8,
        run: () => {
            if (!cols('products').includes('approved_source_id'))
                exports.db.exec('ALTER TABLE products ADD COLUMN approved_source_id TEXT');
        },
    },
    {
        // Normalize legacy 'sales' role → 'marketing' (one-time)
        version: 9,
        run: () => {
            exports.db.prepare("UPDATE users SET role = 'marketing' WHERE role = 'sales'").run();
        },
    },
    {
        // Add tabs column to roles
        version: 10,
        run: () => {
            const cols = exports.db.prepare("PRAGMA table_info(roles)").all().map((r) => r.name);
            if (!cols.includes('tabs'))
                exports.db.exec("ALTER TABLE roles ADD COLUMN tabs TEXT NOT NULL DEFAULT '{}'");
        },
    },
    {
        // Add needs_price_review column to inquiry_items
        version: 12,
        run: () => {
            if (!cols('inquiry_items').includes('needs_price_review'))
                exports.db.exec('ALTER TABLE inquiry_items ADD COLUMN needs_price_review INTEGER NOT NULL DEFAULT 0');
        },
    },
    {
        // Add sourcing_missed column to inquiry_items
        version: 13,
        run: () => {
            if (!cols('inquiry_items').includes('sourcing_missed'))
                exports.db.exec('ALTER TABLE inquiry_items ADD COLUMN sourcing_missed INTEGER NOT NULL DEFAULT 0');
        },
    },
    {
        // Add item review workflow columns to inquiry_items
        version: 14,
        run: () => {
            const itemCols = cols('inquiry_items');
            if (!itemCols.includes('review_status')) {
                exports.db.exec("ALTER TABLE inquiry_items ADD COLUMN review_status TEXT NOT NULL DEFAULT 'pending'");
            }
            if (!itemCols.includes('review_round')) {
                exports.db.exec('ALTER TABLE inquiry_items ADD COLUMN review_round INTEGER NOT NULL DEFAULT 0');
            }
            exports.db.exec(`UPDATE inquiry_items
         SET review_status = CASE
           WHEN price_approved = 1 THEN 'approved'
           WHEN needs_price_review = 1 THEN 'review'
           ELSE 'pending'
         END
         WHERE review_status IS NULL OR review_status = '' OR review_status = 'pending'`);
            const inquiryCols = cols('inquiries');
            if (!inquiryCols.includes('sent_incomplete')) {
                exports.db.exec('ALTER TABLE inquiries ADD COLUMN sent_incomplete INTEGER NOT NULL DEFAULT 0');
            }
            if (!inquiryCols.includes('sent_incomplete_reason')) {
                exports.db.exec('ALTER TABLE inquiries ADD COLUMN sent_incomplete_reason TEXT');
            }
        },
    },
    {
        // Add organization to inquiries
        version: 15,
        run: () => {
            if (!cols('inquiries').includes('organization')) {
                exports.db.exec("ALTER TABLE inquiries ADD COLUMN organization TEXT NOT NULL DEFAULT 'FJM'");
            }
            exports.db.exec("UPDATE inquiries SET organization = 'FJM' WHERE organization IS NULL OR organization = ''");
        },
    },
    {
        // Add sourcing_missed to inquiries (RFQ-level missed tracking)
        version: 17,
        run: () => {
            if (!cols('inquiries').includes('sourcing_missed'))
                exports.db.exec('ALTER TABLE inquiries ADD COLUMN sourcing_missed INTEGER NOT NULL DEFAULT 0');
        },
    },
    {
        // Create organizations master table and seed defaults
        version: 16,
        run: () => {
            exports.db.exec(`
        CREATE TABLE IF NOT EXISTS organizations (
          id TEXT PRIMARY KEY,
          code TEXT UNIQUE NOT NULL,
          created_at TEXT NOT NULL,
          created_by TEXT
        )
      `);
            const now = new Date().toISOString();
            const insertOrg = exports.db.prepare('INSERT OR IGNORE INTO organizations (id, code, created_at, created_by) VALUES (?, ?, ?, ?)');
            insertOrg.run((0, id_1.generateId)(), 'FJM', now, 'system');
            insertOrg.run((0, id_1.generateId)(), 'FMI', now, 'system');
            insertOrg.run((0, id_1.generateId)(), 'FSA', now, 'system');
        },
    },
    {
        // Create persistent notification queue for push events
        version: 17,
        run: () => {
            exports.db.exec(`
        CREATE TABLE IF NOT EXISTS notifications (
          id TEXT PRIMARY KEY,
          type TEXT NOT NULL,
          inquiry_id TEXT NOT NULL,
          rfq_no TEXT,
          message TEXT NOT NULL,
          triggered_by TEXT NOT NULL,
          triggered_by_name TEXT NOT NULL,
          created_at TEXT NOT NULL,
          read_at TEXT
        )
      `);
            exports.db.exec(`CREATE INDEX IF NOT EXISTS idx_notifications_read_at ON notifications (read_at)`);
        },
    },
    {
        // Add price_approval_started_at to track when an inquiry entered price_approval
        version: 18,
        run: () => {
            if (!cols('inquiries').includes('price_approval_started_at'))
                exports.db.exec('ALTER TABLE inquiries ADD COLUMN price_approval_started_at TEXT');
        },
    },
    {
        // Add ppn_type to inquiry_items
        version: 19,
        run: () => {
            if (!cols('inquiry_items').includes('ppn_type'))
                exports.db.exec('ALTER TABLE inquiry_items ADD COLUMN ppn_type TEXT');
        },
    },
    {
        // Migrate previously-flagged missed RFQs to status = 'missed'
        version: 20,
        run: () => {
            exports.db.exec(`UPDATE inquiries SET status = 'missed' WHERE status = 'rfq' AND sourcing_missed = 1`);
        },
    },
    {
        // Add recipient_username to notifications for per-user targeting
        version: 21,
        run: () => {
            if (!cols('notifications').includes('recipient_username')) {
                exports.db.exec('ALTER TABLE notifications ADD COLUMN recipient_username TEXT');
            }
            exports.db.exec('CREATE INDEX IF NOT EXISTS idx_notifications_recipient ON notifications (recipient_username)');
        },
    },
];
const runMigrations = () => {
    const versionRow = exports.db.prepare('SELECT version FROM schema_version').get();
    if (!versionRow) {
        // schema_version table is empty — decide whether fresh install or legacy DB
        if (isFreshInstall) {
            exports.db.prepare('INSERT INTO schema_version (version) VALUES (?)').run(LATEST_VERSION);
            return;
        }
        else {
            exports.db.prepare('INSERT INTO schema_version (version) VALUES (0)').run();
        }
    }
    let version = exports.db.prepare('SELECT version FROM schema_version').get().version;
    if (version >= LATEST_VERSION)
        return;
    for (const m of migrations) {
        if (m.version > version) {
            m.run();
            exports.db.prepare('UPDATE schema_version SET version = ?').run(m.version);
            version = m.version;
        }
    }
};
runMigrations();
// Safety net for environments where schema_version is ahead but a column is missing.
const ensureInquiryItemsColumns = () => {
    const c = cols('inquiry_items');
    if (!c.includes('needs_price_review')) {
        exports.db.exec('ALTER TABLE inquiry_items ADD COLUMN needs_price_review INTEGER NOT NULL DEFAULT 0');
    }
    if (!c.includes('review_status')) {
        exports.db.exec("ALTER TABLE inquiry_items ADD COLUMN review_status TEXT NOT NULL DEFAULT 'pending'");
    }
    if (!c.includes('review_round')) {
        exports.db.exec('ALTER TABLE inquiry_items ADD COLUMN review_round INTEGER NOT NULL DEFAULT 0');
    }
};
ensureInquiryItemsColumns();
const ensureInquiriesColumns = () => {
    const c = cols('inquiries');
    if (!c.includes('sent_incomplete')) {
        exports.db.exec('ALTER TABLE inquiries ADD COLUMN sent_incomplete INTEGER NOT NULL DEFAULT 0');
    }
    if (!c.includes('sent_incomplete_reason')) {
        exports.db.exec('ALTER TABLE inquiries ADD COLUMN sent_incomplete_reason TEXT');
    }
    if (!c.includes('organization')) {
        exports.db.exec("ALTER TABLE inquiries ADD COLUMN organization TEXT NOT NULL DEFAULT 'FJM'");
    }
    exports.db.exec("UPDATE inquiries SET organization = 'FJM' WHERE organization IS NULL OR organization = ''");
    if (!c.includes('sourcing_missed'))
        exports.db.exec('ALTER TABLE inquiries ADD COLUMN sourcing_missed INTEGER NOT NULL DEFAULT 0');
};
ensureInquiriesColumns();
const ensureOrganizationsTable = () => {
    exports.db.exec(`
    CREATE TABLE IF NOT EXISTS organizations (
      id TEXT PRIMARY KEY,
      code TEXT UNIQUE NOT NULL,
      created_at TEXT NOT NULL,
      created_by TEXT
    )
  `);
};
ensureOrganizationsTable();
const ensureNotificationsColumns = () => {
    const c = cols('notifications');
    if (!c.includes('recipient_username')) {
        exports.db.exec('ALTER TABLE notifications ADD COLUMN recipient_username TEXT');
    }
    exports.db.exec('CREATE INDEX IF NOT EXISTS idx_notifications_recipient ON notifications (recipient_username)');
};
ensureNotificationsColumns();
// ─── Seed data ────────────────────────────────────────────────────────────────
const ensureDefaultRoles = () => {
    const defaults = [
        { name: 'admin', menus: ['marketing', 'sourcing', 'dashboard', 'pricelist', 'admin', 'purchasing'] },
        { name: 'manager', menus: ['pricelist'] },
        { name: 'sourcing', menus: ['sourcing'] },
        { name: 'marketing', menus: ['marketing'] },
    ];
    const insert = exports.db.prepare('INSERT OR IGNORE INTO roles (name, menus, created_at) VALUES (?, ?, ?)');
    const now = new Date().toISOString();
    const tx = exports.db.transaction((rows) => {
        for (const row of rows)
            insert.run(row.name, JSON.stringify(row.menus), now);
    });
    tx(defaults);
};
const seedUsers = () => {
    const count = exports.db.prepare('SELECT COUNT(*) as count FROM users').get();
    if (count.count > 0)
        return;
    const insert = exports.db.prepare('INSERT INTO users (id, name, username, password, role) VALUES (@id, @name, @username, @password, @role)');
    const seed = [
        { id: (0, id_1.generateId)(), name: 'Administrator', username: 'admin', password: 'admin', role: 'admin' },
        { id: (0, id_1.generateId)(), name: 'Sourcing User', username: 'sourcing', password: 'sourcing', role: 'sourcing' },
        { id: (0, id_1.generateId)(), name: 'Marketing User', username: 'marketing', password: 'marketing', role: 'marketing' },
        { id: (0, id_1.generateId)(), name: 'Sales User', username: 'sales', password: 'sales', role: 'marketing' },
    ];
    const tx = exports.db.transaction((rows) => { for (const row of rows)
        insert.run(row); });
    tx(seed);
};
const seedOrganizations = () => {
    const now = new Date().toISOString();
    const insert = exports.db.prepare('INSERT OR IGNORE INTO organizations (id, code, created_at, created_by) VALUES (?, ?, ?, ?)');
    const tx = exports.db.transaction(() => {
        insert.run((0, id_1.generateId)(), 'FJM', now, 'system');
        insert.run((0, id_1.generateId)(), 'FMI', now, 'system');
        insert.run((0, id_1.generateId)(), 'FSA', now, 'system');
    });
    tx();
};
ensureDefaultRoles();
seedUsers();
seedOrganizations();
exports.db.prepare(`INSERT OR IGNORE INTO settings (key, value) VALUES ('default_margin_pct', '20')`).run();
