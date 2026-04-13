"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.inquiriesRouter = void 0;
const express_1 = require("express");
const db_1 = require("../db");
const id_1 = require("../utils/id");
const XLSX = __importStar(require("xlsx"));
exports.inquiriesRouter = (0, express_1.Router)();
function generateRfqNo() {
    const d = new Date();
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    const rand = Math.floor(Math.random() * 9000) + 1000;
    return `RFQ-${y}${m}${day}-${rand}`;
}
function logActivity(inquiryId, action, oldStatus, newStatus, note, doneBy, doneByName) {
    db_1.db.prepare(`INSERT INTO activity_log (id, inquiry_id, action, old_status, new_status, note, done_by, done_by_name, created_at)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`).run((0, id_1.generateId)(), inquiryId, action, oldStatus, newStatus, note, doneBy, doneByName, new Date().toISOString());
}
function mapItem(row) {
    return {
        id: row['id'],
        inquiryId: row['inquiry_id'],
        coupaRowIndex: row['coupa_row_index'],
        lotId: row['lot_id'],
        lotName: row['lot_name'],
        lotExpectedQuantity: row['lot_expected_quantity'],
        lotQuantityNote: row['lot_quantity_note'],
        coupaItemId: row['coupa_item_id'],
        itemName: row['item_name'],
        itemQuantity: row['item_quantity'],
        itemUom: row['item_uom'],
        itemNeedByDate: parseExcelDate(row['item_need_by_date']),
        itemManufacturerName: row['item_manufacturer_name'],
        itemManufacturerPartNumber: row['item_manufacturer_part_number'],
        itemClassificationOfGoods: row['item_classification_of_goods'],
        itemExtendedDescription: row['item_extended_description'],
        itemFiscalCode: row['item_fiscal_code'],
        itemImage: row['item_image'],
        coupaBidId: row['coupa_bid_id'],
        bidCapacity: row['bid_capacity'],
        bidPriceAmount: row['bid_price_amount'],
        bidPriceCurrency: row['bid_price_currency'],
        bidLeadTime: row['bid_lead_time'],
        bidSupplierItemName: row['bid_supplier_item_name'],
        alternateName: row['alternate_name'],
        bidItemPartNumber: row['bid_item_part_number'],
        bidItemDescription: row['bid_item_description'],
        bidShippingTerm: row['bid_shipping_term'],
        supplier: row['supplier'],
        hargaBeli: row['harga_beli'],
        leadTime: row['lead_time'],
        moq: row['moq'],
        stockAvailability: row['stock_availability'],
        termPembayaran: row['term_pembayaran'],
        hargaJual: row['harga_jual'],
        approvedPrice: row['approved_price'],
        margin: row['margin'],
        leadTimeCustomer: row['lead_time_customer'],
        validitasQuotation: row['validitas_quotation'],
        catatanQuotation: row['catatan_quotation'],
        priceApproved: row['price_approved'] === 1,
        needsPriceReview: row['needs_price_review'] === 1,
        reviewStatus: row['review_status'] ?? 'pending',
        reviewRound: Number(row['review_round'] ?? 0),
        sourcingMissed: row['sourcing_missed'] === 1,
    };
}
function mapInquiry(row, items) {
    return {
        id: row['id'],
        rfqNo: row['rfq_no'],
        tanggal: row['tanggal'],
        customer: row['customer'],
        salesPic: row['sales_pic'],
        sourcingPic: row['sourcing_pic'] ?? null,
        status: row['status'],
        coupaSource: row['coupa_source'] === 1,
        coupaFileName: row['coupa_file_name'] ?? null,
        createdAt: row['created_at'],
        createdBy: row['created_by'],
        updatedAt: row['updated_at'],
        updatedBy: row['updated_by'],
        sentIncomplete: row['sent_incomplete'] === 1,
        sentIncompleteReason: row['sent_incomplete_reason'] ?? null,
        needByDate: row['deadline_quotation'] ?? null,
        items: items.map(mapItem),
    };
}
function deriveCustomerFromFilename(fileName) {
    const clean = fileName.replace(/\.[^.]+$/, '');
    const parts = clean.split('#').map((p) => p.trim()).filter(Boolean);
    if (parts.length >= 2) {
        const name = parts[1].replace(/[-_]+/g, ' ').trim();
        const eventId = parts[2] ? parts[2].trim() : null;
        return eventId ? `${name} #${eventId}` : name;
    }
    return clean.replace(/[-_]+/g, ' ').trim();
}
function parseCoupaFieldMap(sheet) {
    const map = {};
    const range = sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : null;
    if (!range) {
        return map;
    }
    // Primary: read JSON field_name metadata from row 0
    for (let c = range.s.c; c <= range.e.c; c += 1) {
        const cellAddress = XLSX.utils.encode_cell({ r: 0, c });
        const cell = sheet[cellAddress];
        if (!cell || cell.v == null)
            continue;
        const raw = String(cell.v);
        if (!raw.includes('field_name'))
            continue;
        try {
            const parsed = JSON.parse(raw);
            if (parsed?.field_name) {
                map[parsed.field_name] = c;
            }
        }
        catch { /* ignore */ }
    }
    // Fallback: scan row 4 (human-readable column labels) for known fields not found via JSON
    const labelFallbacks = [
        ['item.need_by_date', 'need by date'],
    ];
    for (const [fieldName, labelSubstr] of labelFallbacks) {
        if (map[fieldName] != null)
            continue; // already mapped
        for (let c = range.s.c; c <= range.e.c; c += 1) {
            const cell = sheet[XLSX.utils.encode_cell({ r: 4, c })];
            if (!cell || cell.v == null)
                continue;
            if (String(cell.v).toLowerCase().includes(labelSubstr)) {
                map[fieldName] = c;
                break;
            }
        }
    }
    return map;
}
function readSheetCell(sheet, rowIndex, colIndex) {
    if (colIndex == null)
        return null;
    const cell = sheet[XLSX.utils.encode_cell({ r: rowIndex, c: colIndex })];
    if (!cell)
        return null;
    return cell.v ?? null;
}
function setSheetCell(sheet, rowIndex, colIndex, value) {
    if (colIndex == null || value == null || value === '')
        return;
    const addr = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
    const cell = sheet[addr] ?? {};
    cell.v = value;
    cell.t = typeof value === 'number' ? 'n' : 's';
    sheet[addr] = cell;
}
function parseExcelDate(value) {
    if (value == null || value === '')
        return null;
    // Excel date serial number
    const num = typeof value === 'number' ? value : Number(value);
    if (!isNaN(num) && num > 40000) {
        const date = new Date((num - 25569) * 86400 * 1000);
        return isNaN(date.getTime()) ? null : date.toISOString().slice(0, 10);
    }
    if (typeof value === 'string') {
        // Already an ISO date string stored in DB — return date part only
        if (/^\d{4}-\d{2}-\d{2}/.test(value))
            return value.slice(0, 10);
        // Try general date parse (handles "30-Mar-2026", "2026/03/30", etc.)
        const d = new Date(value);
        return isNaN(d.getTime()) ? null : d.toISOString().slice(0, 10);
    }
    return null;
}
function normalizeLeadTime(value) {
    if (value == null)
        return null;
    if (typeof value === 'number')
        return Number.isFinite(value) ? value : null;
    const trimmed = String(value).trim();
    if (!trimmed)
        return null;
    const num = Number.parseFloat(trimmed.replace(/[^0-9.]/g, ''));
    return Number.isFinite(num) ? num : null;
}
function allItemsSourced(items) {
    if (items.length === 0)
        return false;
    return items.every((item) => item['supplier'] && item['harga_beli'] != null && item['lead_time']);
}
function allItemsApproved(items) {
    if (items.length === 0)
        return false;
    return items.every((item) => item['price_approved'] === 1);
}
function recalcInquiryStatus(inquiryId, doneBy, doneByName) {
    // Manual transition is now required for price_approval -> price_approved.
    // Keep this function as a no-op because older call sites still invoke it.
    void inquiryId;
    void doneBy;
    void doneByName;
}
// GET /inquiries
exports.inquiriesRouter.get('/', (_req, res) => {
    const rows = db_1.db.prepare('SELECT * FROM inquiries ORDER BY created_at DESC').all();
    const items = db_1.db.prepare('SELECT * FROM inquiry_items ORDER BY coupa_row_index ASC, id ASC').all();
    const logs = db_1.db.prepare('SELECT * FROM activity_log ORDER BY created_at ASC').all();
    const itemsByInquiry = new Map();
    for (const item of items) {
        const id = String(item['inquiry_id']);
        const list = itemsByInquiry.get(id) ?? [];
        list.push(item);
        itemsByInquiry.set(id, list);
    }
    const logsByInquiry = new Map();
    for (const log of logs) {
        const id = String(log['inquiry_id']);
        const list = logsByInquiry.get(id) ?? [];
        list.push({
            id: log['id'],
            inquiryId: log['inquiry_id'],
            action: log['action'],
            oldStatus: log['old_status'],
            newStatus: log['new_status'],
            note: log['note'],
            doneBy: log['done_by'],
            doneByName: log['done_by_name'],
            createdAt: log['created_at'],
        });
        logsByInquiry.set(id, list);
    }
    const result = rows.map((row) => ({
        ...mapInquiry(row, itemsByInquiry.get(String(row['id'])) ?? []),
        activityLog: logsByInquiry.get(String(row['id'])) ?? [],
    }));
    res.json(result);
});
// GET /inquiries/dashboard/user
exports.inquiriesRouter.get('/dashboard/user', (req, res) => {
    const name = String(req.query['name'] ?? '').trim();
    if (!name) {
        res.status(400).json({ error: 'name query param is required.' });
        return;
    }
    const startOfMonth = new Date();
    startOfMonth.setDate(1);
    startOfMonth.setHours(0, 0, 0, 0);
    const startOfMonthIso = startOfMonth.toISOString();
    // Sales stats
    const total = db_1.db.prepare('SELECT COUNT(*) as c FROM inquiries WHERE sales_pic = ?').get(name).c;
    const thisMonthSales = db_1.db.prepare('SELECT COUNT(*) as c FROM inquiries WHERE sales_pic = ? AND created_at >= ?').get(name, startOfMonthIso).c;
    const quotationSent = db_1.db.prepare(`SELECT COUNT(*) as c FROM inquiries WHERE sales_pic = ? AND status IN ('quotation_sent','deal')`).get(name).c;
    const sentIncomplete = db_1.db.prepare(`SELECT COUNT(*) as c FROM inquiries
     WHERE sales_pic = ? AND status IN ('quotation_sent','deal') AND sent_incomplete = 1`).get(name).c;
    const sentIncompleteRate = quotationSent > 0 ? +((sentIncomplete / quotationSent) * 100).toFixed(1) : 0;
    const deals = db_1.db.prepare(`SELECT COUNT(*) as c FROM inquiries WHERE sales_pic = ? AND status = 'deal'`).get(name).c;
    const lost = db_1.db.prepare(`SELECT COUNT(*) as c FROM inquiries WHERE sales_pic = ? AND status = 'lost'`).get(name).c;
    const active = db_1.db.prepare(`SELECT COUNT(*) as c FROM inquiries WHERE sales_pic = ? AND status NOT IN ('deal','lost')`).get(name).c;
    const conversionRate = total > 0 ? +((deals / total) * 100).toFixed(1) : 0;
    const statusBreakdown = db_1.db.prepare('SELECT status, COUNT(*) as count FROM inquiries WHERE sales_pic = ? GROUP BY status ORDER BY count DESC').all(name);
    // Sourcing stats
    const itemsSourced = db_1.db.prepare(`SELECT COUNT(*) as c FROM activity_log WHERE done_by_name = ? AND action = 'Sourcing info submitted'`).get(name).c;
    const inquiriesContributed = db_1.db.prepare(`SELECT COUNT(DISTINCT inquiry_id) as c FROM activity_log WHERE done_by_name = ? AND action = 'Sourcing info submitted'`).get(name).c;
    const thisMonthSourcing = db_1.db.prepare(`SELECT COUNT(*) as c FROM activity_log WHERE done_by_name = ? AND action = 'Sourcing info submitted' AND created_at >= ?`).get(name, startOfMonthIso).c;
    // Per-user item state breakdown — scoped to RFQs assigned to this sourcing user
    const userItemsTerisi = db_1.db.prepare(`SELECT COUNT(*) as c FROM inquiry_items ii
     JOIN inquiries i ON i.id = ii.inquiry_id
     WHERE i.sourcing_pic = ?
       AND ii.supplier IS NOT NULL AND ii.harga_beli IS NOT NULL AND ii.lead_time IS NOT NULL
       AND ii.sourcing_missed = 0`).get(name).c;
    const userItemsMissed = db_1.db.prepare(`SELECT COUNT(*) as c FROM inquiry_items ii
     JOIN inquiries i ON i.id = ii.inquiry_id
     WHERE i.sourcing_pic = ? AND ii.sourcing_missed = 1`).get(name).c;
    const userItemsTidakTerisi = db_1.db.prepare(`SELECT COUNT(*) as c FROM inquiry_items ii
     JOIN inquiries i ON i.id = ii.inquiry_id
     WHERE i.sourcing_pic = ?
       AND (ii.supplier IS NULL OR ii.harga_beli IS NULL OR ii.lead_time IS NULL)`).get(name).c;
    // Manager stats
    const approvalsTotal = db_1.db.prepare(`SELECT COUNT(*) as c FROM activity_log WHERE done_by_name = ? AND action = 'Price approved'`).get(name).c;
    const approvalsThisMonth = db_1.db.prepare(`SELECT COUNT(*) as c FROM activity_log WHERE done_by_name = ? AND action = 'Price approved' AND created_at >= ?`).get(name, startOfMonthIso).c;
    const inquiriesApproved = db_1.db.prepare(`SELECT COUNT(DISTINCT inquiry_id) as c FROM activity_log WHERE done_by_name = ? AND action = 'Price approved'`).get(name).c;
    res.json({
        salesStats: { total, thisMonth: thisMonthSales, quotationSent, sentIncomplete, sentIncompleteRate, deals, lost, active, conversionRate, statusBreakdown },
        sourcingStats: { itemsSourced, inquiriesContributed, thisMonth: thisMonthSourcing, itemsTerisi: userItemsTerisi, itemsMissed: userItemsMissed, itemsTidakTerisi: userItemsTidakTerisi },
        managerStats: { approvalsTotal, approvalsThisMonth, inquiriesApproved },
    });
});
// GET /inquiries/dashboard
exports.inquiriesRouter.get('/dashboard', (_req, res) => {
    const startOfMonth = new Date();
    startOfMonth.setDate(1);
    startOfMonth.setHours(0, 0, 0, 0);
    const startOfMonthIso = startOfMonth.toISOString();
    // Sales stats
    const total = db_1.db.prepare('SELECT COUNT(*) as c FROM inquiries').get().c;
    const thisMonth = db_1.db.prepare('SELECT COUNT(*) as c FROM inquiries WHERE created_at >= ?').get(startOfMonthIso).c;
    const quotationSent = db_1.db.prepare(`SELECT COUNT(*) as c FROM inquiries WHERE status IN ('quotation_sent','deal')`).get().c;
    const sentIncomplete = db_1.db.prepare(`SELECT COUNT(*) as c FROM inquiries WHERE status IN ('quotation_sent','deal') AND sent_incomplete = 1`).get().c;
    const sentIncompleteRate = quotationSent > 0 ? +((sentIncomplete / quotationSent) * 100).toFixed(1) : 0;
    const deals = db_1.db.prepare(`SELECT COUNT(*) as c FROM inquiries WHERE status = 'deal'`).get().c;
    const lost = db_1.db.prepare(`SELECT COUNT(*) as c FROM inquiries WHERE status = 'lost'`).get().c;
    const conversionRate = total > 0 ? +((deals / total) * 100).toFixed(1) : 0;
    const topSales = db_1.db.prepare(`SELECT sales_pic, COUNT(*) as deal_count FROM inquiries WHERE status = 'deal' GROUP BY sales_pic ORDER BY deal_count DESC LIMIT 5`).all();
    const statusBreakdown = db_1.db.prepare(`SELECT status, COUNT(*) as count FROM inquiries GROUP BY status ORDER BY count DESC`).all();
    // Item state breakdown — single source of truth from inquiry_items current state
    const itemsTerisi = db_1.db.prepare(`SELECT COUNT(*) as c FROM inquiry_items WHERE supplier IS NOT NULL AND harga_beli IS NOT NULL AND lead_time IS NOT NULL AND sourcing_missed = 0`).get().c;
    const itemsMissed = db_1.db.prepare(`SELECT COUNT(*) as c FROM inquiry_items WHERE sourcing_missed = 1`).get().c;
    const itemsTidakTerisi = db_1.db.prepare(`SELECT COUNT(*) as c FROM inquiry_items WHERE (supplier IS NULL OR harga_beli IS NULL OR lead_time IS NULL)`).get().c;
    // Sourcing stats — use inquiry_items for totals so stat cards match the pie chart
    const sourcingPending = db_1.db.prepare(`SELECT COUNT(*) as c FROM inquiries WHERE status = 'rfq'`).get().c;
    // Total sourced = items with supplier data filled (terisi + missed)
    const sourcingItemsTotal = itemsTerisi + itemsMissed;
    // "This month" approximated from activity_log (deduplication not possible without item timestamps)
    const sourcingItemsThisMonth = db_1.db.prepare(`SELECT COUNT(DISTINCT inquiry_id) as c FROM activity_log WHERE action = 'Sourcing info submitted' AND created_at >= ?`).get(startOfMonthIso).c;
    const topSourcers = db_1.db.prepare(`SELECT done_by_name as sourcing_pic, COUNT(*) as items_count
     FROM activity_log WHERE action = 'Sourcing info submitted'
     GROUP BY done_by_name ORDER BY items_count DESC LIMIT 5`).all();
    res.json({
        total, thisMonth, quotationSent, sentIncomplete, sentIncompleteRate, deals, lost, conversionRate, topSales, statusBreakdown,
        sourcingPending, sourcingItemsThisMonth, sourcingItemsTotal, topSourcers,
        itemsTerisi, itemsTidakTerisi, itemsMissed,
    });
});
// POST /inquiries/import-coupa
exports.inquiriesRouter.post('/import-coupa', (req, res) => {
    const { fileBase64, fileName, createdBy, createdByName } = req.body;
    if (!fileBase64 || !fileName || !createdBy) {
        res.status(400).json({ error: 'fileBase64, fileName, createdBy are required.' });
        return;
    }
    let workbook;
    try {
        const buffer = Buffer.from(String(fileBase64), 'base64');
        workbook = XLSX.read(buffer, { type: 'buffer', cellStyles: true });
    }
    catch {
        res.status(400).json({ error: 'Invalid Excel file.' });
        return;
    }
    const sheet = workbook.Sheets['Items and Services'];
    if (!sheet) {
        res.status(400).json({ error: 'Items and Services sheet not found.' });
        return;
    }
    const fieldMap = parseCoupaFieldMap(sheet);
    const itemIdCol = fieldMap['item.id'];
    const bidIdCol = fieldMap['bid.id'];
    if (itemIdCol == null || bidIdCol == null) {
        res.status(400).json({ error: 'Coupa field mapping missing item.id or bid.id.' });
        return;
    }
    const range = sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : null;
    if (!range) {
        res.status(400).json({ error: 'Sheet is empty.' });
        return;
    }
    const items = [];
    for (let r = 5; r <= range.e.r; r += 1) {
        const itemId = readSheetCell(sheet, r, itemIdCol);
        const bidId = readSheetCell(sheet, r, bidIdCol);
        if (itemId == null && bidId == null) {
            continue;
        }
        const toNumber = (value) => {
            if (value == null || value === '')
                return null;
            const num = Number(value);
            return Number.isFinite(num) ? num : null;
        };
        items.push({
            coupa_row_index: r + 1,
            lot_id: readSheetCell(sheet, r, fieldMap['lot.id']),
            lot_name: readSheetCell(sheet, r, fieldMap['lot.name']),
            lot_expected_quantity: toNumber(readSheetCell(sheet, r, fieldMap['lot.expected_quantity'])),
            lot_quantity_note: readSheetCell(sheet, r, fieldMap['lot.quantity_note']),
            coupa_item_id: itemId,
            item_name: readSheetCell(sheet, r, fieldMap['item.name']),
            item_quantity: toNumber(readSheetCell(sheet, r, fieldMap['item.quantity'])),
            item_uom: readSheetCell(sheet, r, fieldMap['item.uom']),
            item_need_by_date: parseExcelDate(readSheetCell(sheet, r, fieldMap['item.need_by_date'])),
            item_manufacturer_name: readSheetCell(sheet, r, fieldMap['item.manufacturer_name']),
            item_manufacturer_part_number: readSheetCell(sheet, r, fieldMap['item.manufacturer_part_number']),
            item_classification_of_goods: readSheetCell(sheet, r, fieldMap['item.classification_of_goods']),
            item_extended_description: readSheetCell(sheet, r, fieldMap['item.extended_description']),
            item_fiscal_code: readSheetCell(sheet, r, fieldMap['item.fiscal_code']),
            coupa_bid_id: bidId,
            bid_capacity: toNumber(readSheetCell(sheet, r, fieldMap['bid.capacity'])),
            bid_price_amount: toNumber(readSheetCell(sheet, r, fieldMap['bid.price_amount'])),
            bid_price_currency: readSheetCell(sheet, r, fieldMap['bid.price_currency']),
            bid_lead_time: readSheetCell(sheet, r, fieldMap['bid.lead_time']),
            bid_supplier_item_name: readSheetCell(sheet, r, fieldMap['bid.supplier_item_name']),
            bid_item_part_number: readSheetCell(sheet, r, fieldMap['bid.item_part_number']),
            bid_item_description: readSheetCell(sheet, r, fieldMap['bid.item_description']),
            bid_shipping_term: readSheetCell(sheet, r, fieldMap['bid.shipping_term']),
        });
    }
    if (items.length === 0) {
        res.status(400).json({ error: 'No data rows found.' });
        return;
    }
    const id = (0, id_1.generateId)();
    const rfqNo = generateRfqNo();
    const tanggal = new Date().toISOString().split('T')[0];
    const createdAt = new Date().toISOString();
    const customer = deriveCustomerFromFilename(String(fileName));
    const salesPic = String(createdByName ?? createdBy);
    const firstItemName = String(items[0]?.item_name ?? 'Multiple items');
    const namaBarang = items.length > 1 ? `${firstItemName} +${items.length - 1} items` : firstItemName;
    const tx = db_1.db.transaction(() => {
        db_1.db.prepare(`INSERT INTO inquiries (id, rfq_no, tanggal, customer, sales_pic, nama_barang, status, coupa_source, coupa_file_name, created_at, created_by)
       VALUES (?, ?, ?, ?, ?, ?, 'new_inquiry', 1, ?, ?, ?)`).run(id, rfqNo, tanggal, customer, salesPic, namaBarang, String(fileName), createdAt, createdBy);
        const insertItem = db_1.db.prepare(`INSERT INTO inquiry_items (
        id, inquiry_id, coupa_row_index, lot_id, lot_name, lot_expected_quantity, lot_quantity_note,
        coupa_item_id, item_name, item_quantity, item_uom, item_need_by_date, item_manufacturer_name,
        item_manufacturer_part_number, item_classification_of_goods, item_extended_description, item_fiscal_code,
        coupa_bid_id, bid_capacity, bid_price_amount, bid_price_currency, bid_lead_time,
        bid_supplier_item_name, bid_item_part_number, bid_item_description, bid_shipping_term,
        harga_jual, alternate_name
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`);
        for (const item of items) {
            insertItem.run((0, id_1.generateId)(), id, item.coupa_row_index ?? null, item.lot_id ?? null, item.lot_name ?? null, item.lot_expected_quantity ?? null, item.lot_quantity_note ?? null, item.coupa_item_id ?? null, item.item_name ?? null, item.item_quantity ?? null, item.item_uom ?? null, item.item_need_by_date ?? null, item.item_manufacturer_name ?? null, item.item_manufacturer_part_number ?? null, item.item_classification_of_goods ?? null, item.item_extended_description ?? null, item.item_fiscal_code ?? null, item.coupa_bid_id ?? null, item.bid_capacity ?? null, item.bid_price_amount ?? null, item.bid_price_currency ?? null, item.bid_lead_time ?? null, item.bid_supplier_item_name ?? null, item.bid_item_part_number ?? null, item.bid_item_description ?? null, item.bid_shipping_term ?? null, item.bid_price_amount ?? null, item.bid_supplier_item_name ?? null);
        }
        db_1.db.prepare(`INSERT INTO coupa_files (inquiry_id, file_name, file_data, created_at) VALUES (?, ?, ?, ?)`).run(id, String(fileName), Buffer.from(String(fileBase64), 'base64'), createdAt);
    });
    tx();
    logActivity(id, 'Coupa file imported', null, 'new_inquiry', String(fileName), String(createdBy), String(createdByName ?? createdBy));
    res.status(201).json({ id, rfqNo, itemCount: items.length });
});
function generateExcelFilename(customer) {
    // customer may be "Pt Merdeka Copper Gold Event #58847" → "en#Pt Merdeka Copper Gold Event#58847"
    const match = customer.match(/\s*#(\d+)\s*$/);
    const eventNo = match ? match[1] : '';
    const customerName = match ? customer.replace(/\s*#\d+\s*$/, '').trim() : customer;
    return `en#${customerName}${eventNo ? '#' + eventNo : ''}`;
}
function buildCoupaFormatExcel(items) {
    // Exact Coupa column layout (verified from real Coupa file):
    // Col 0:  (special start entry, no field_name)
    // Col 1:  lot.id              Col 2:  lot.name
    // Col 3:  lot.expected_qty    Col 4:  lot.quantity_note
    // Col 5:  item.id             Col 6:  item.name
    // Col 7:  item.quantity       Col 8:  item.uom
    // Col 9:  item.need_by_date   Col 10: item.manufacturer_name
    // Col 11: item.manufacturer_part_number
    // Col 12: item.classification_of_goods
    // Col 13: item.extended_description
    // Col 14: item.fiscal_code    Col 15: bid.id
    // Col 16: bid.capacity        Col 17: bid.price_amount
    // Col 18: bid.price_currency  Col 19: bid.lead_time
    // Col 20: bid.supplier_item_name
    // Col 21: bid.item_part_number
    // Col 22: bid.item_description
    // Col 23: bid.shipping_term
    const NUM_COLS = 24;
    const ws = {};
    const setCell = (r, c, v) => {
        if (v == null)
            return;
        ws[XLSX.utils.encode_cell({ r, c })] = { v, t: typeof v === 'number' ? 'n' : 's' };
    };
    // Row 0: JSON field_name metadata — matches real Coupa format exactly (with offset:5)
    setCell(0, 0, JSON.stringify({ start: true, layout: 'table', name: 'supplier/response_lines', locale: 'en' }));
    const fieldCols = [
        [1, 'lot.id'], [2, 'lot.name'],
        [3, 'lot.expected_quantity'], [4, 'lot.quantity_note'],
        [5, 'item.id'], [6, 'item.name'],
        [7, 'item.quantity'], [8, 'item.uom'],
        [9, 'item.need_by_date'], [10, 'item.manufacturer_name'],
        [11, 'item.manufacturer_part_number'],
        [12, 'item.classification_of_goods'],
        [13, 'item.extended_description'],
        [14, 'item.fiscal_code'], [15, 'bid.id'],
        [16, 'bid.capacity'], [17, 'bid.price_amount'],
        [18, 'bid.price_currency'], [19, 'bid.lead_time'],
        [20, 'bid.supplier_item_name'],
        [21, 'bid.item_part_number'],
        [22, 'bid.item_description'],
        [23, 'bid.shipping_term'],
    ];
    fieldCols.forEach(([c, field]) => setCell(0, c, JSON.stringify({ field_name: field, offset: 5 })));
    // Row 1: Instruction text (matches real Coupa)
    setCell(1, 1, 'The yellow cells below are your input fields. You can upload this file to save the information you entered in the editable cells to your response.');
    // Row 2: Note (matches real Coupa)
    setCell(2, 1, 'NOTE: This Excel file is locked to ensure it uploads correctly, and you must still click "submit" after uploading to submit!');
    // Row 3: Section headers — exact positions from real Coupa file
    setCell(3, 1, 'Lot');
    setCell(3, 3, 'Lot Fields');
    setCell(3, 5, 'Item / Service');
    setCell(3, 7, 'Item / Service Fields');
    setCell(3, 15, 'Supplier Response Fields');
    // Row 4: Column labels — exact labels from real Coupa file
    const colLabels = [
        [1, 'Lot ID (Text)'],
        [2, 'Lot Name (Text)'],
        [3, 'Expected Quantity (Integer)'],
        [4, 'Quantity Note (Text)'],
        [5, 'Item ID (Text)'],
        [6, 'Item Description (Text)'],
        [7, 'Expected Quantity (Number)'],
        [8, 'Unit of Measurement (Text)'],
        [9, 'Need by Date (Date)'],
        [10, 'Manufacturer Name (Text)'],
        [11, 'Manufacturer Part Number (Text)'],
        [12, 'Classification Of Goods (Text)'],
        [13, 'Description (Text)'],
        [14, 'Fiscal Code (Text)'],
        [15, 'Bid ID (Text)'],
        [16, 'Capacity (Number)'],
        [17, 'Unit Bid Price (Number)'],
        [18, 'Bid Price Currency (Text)'],
        [19, 'Lead Time (Integer)'],
        [20, 'Supplier Item Name (Text)'],
        [21, 'Item Part Number (Text)'],
        [22, 'Item Description (Text)'],
        [23, 'Shipping Terms (Text)'],
    ];
    colLabels.forEach(([c, label]) => setCell(4, c, label));
    // Row 5+: item data
    items.forEach((item, i) => {
        const r = 5 + i;
        const itemId = (item['coupa_item_id'] ?? item['id']);
        const bidId = item['coupa_bid_id'];
        const itemName = item['item_name'];
        const itemQty = item['item_quantity'];
        const approvedPrice = (item['approved_price'] ?? item['harga_jual'] ?? item['bid_price_amount']);
        const bidCapacity = (item['bid_capacity'] ?? itemQty);
        const bidCurrency = (item['bid_price_currency'] ?? 'IDR');
        const leadTime = normalizeLeadTime((item['lead_time_customer'] ?? item['lead_time'] ?? item['bid_lead_time']));
        const catatan = item['catatan_quotation'];
        const description = catatan || item['item_extended_description'] || null;
        const shipping = item['term_pembayaran'] || item['bid_shipping_term'];
        const supplierItemName = (item['bid_supplier_item_name'] ?? item['alternate_name'] ?? null);
        const bidItemPartNumber = item['bid_item_part_number'];
        const bidItemDescription = (item['bid_item_description'] ?? description);
        // Lot columns
        setCell(r, 3, itemQty); // lot.expected_quantity
        // Item columns
        setCell(r, 5, itemId != null ? String(itemId) : null);
        setCell(r, 6, itemName);
        setCell(r, 7, itemQty);
        setCell(r, 8, item['item_uom']);
        setCell(r, 9, parseExcelDate(item['item_need_by_date']));
        setCell(r, 10, item['item_manufacturer_name']);
        setCell(r, 11, item['item_manufacturer_part_number']);
        setCell(r, 12, item['item_classification_of_goods']);
        setCell(r, 13, item['item_extended_description']);
        setCell(r, 14, item['item_fiscal_code']);
        // Bid / supplier response columns
        setCell(r, 15, bidId != null ? String(bidId) : null);
        setCell(r, 16, bidCapacity); // bid.capacity
        setCell(r, 17, approvedPrice); // bid.price_amount
        setCell(r, 18, bidCurrency); // bid.price_currency
        setCell(r, 19, leadTime); // bid.lead_time
        setCell(r, 20, supplierItemName); // bid.supplier_item_name
        setCell(r, 21, bidItemPartNumber); // bid.item_part_number
        setCell(r, 22, bidItemDescription); // bid.item_description
        setCell(r, 23, shipping); // bid.shipping_term
    });
    const lastRow = 5 + Math.max(items.length - 1, 0);
    ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: lastRow, c: NUM_COLS - 1 } });
    // Hide row 0 (JSON metadata) — invisible to users, preserved for re-import
    ws['!rows'] = [{ hidden: true }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Items and Services');
    return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}
// GET /inquiries/:id/export-coupa
exports.inquiriesRouter.get('/:id/export-coupa', (req, res) => {
    const { id } = req.params;
    const inquiryRow = db_1.db.prepare('SELECT id, status, customer FROM inquiries WHERE id = ?').get(id);
    if (!inquiryRow) {
        res.status(404).json({ error: 'Not found.' });
        return;
    }
    const items = db_1.db.prepare(`SELECT * FROM inquiry_items WHERE inquiry_id = ? ORDER BY coupa_row_index ASC, id ASC`).all(id);
    const fileRow = db_1.db.prepare('SELECT file_name, file_data FROM coupa_files WHERE inquiry_id = ?').get(id);
    let output;
    let safeName;
    if (fileRow) {
        // Fill original Coupa Excel with sourcing/approval data
        const workbook = XLSX.read(fileRow.file_data, { type: 'buffer', cellStyles: true });
        const sheet = workbook.Sheets['Items and Services'];
        if (!sheet) {
            res.status(400).json({ error: 'Items and Services sheet not found.' });
            return;
        }
        const fieldMap = parseCoupaFieldMap(sheet);
        for (const item of items) {
            const rowIndex = item['coupa_row_index'];
            if (rowIndex == null)
                continue;
            const row = Number(rowIndex) - 1;
            const approvedPrice = (item['approved_price'] ?? item['harga_jual'] ?? item['bid_price_amount']);
            const leadTimeCustomer = item['lead_time_customer'];
            const leadTimeFallback = item['lead_time'];
            const leadTime = normalizeLeadTime(leadTimeCustomer ?? leadTimeFallback ?? item['bid_lead_time']);
            const catatan = item['catatan_quotation'] ?? null;
            const description = catatan || item['item_extended_description'] || null;
            const itemName = item['item_name'];
            const itemQty = item['item_quantity'];
            const bidCapacity = (item['bid_capacity'] ?? itemQty);
            const bidCurrency = (item['bid_price_currency'] ?? null);
            const supplierItemName = (item['bid_supplier_item_name'] ?? item['alternate_name'] ?? null);
            const bidItemPartNumber = item['bid_item_part_number'];
            const bidItemDescription = (item['bid_item_description'] ?? description);
            setSheetCell(sheet, row, fieldMap['bid.capacity'], bidCapacity ?? null);
            setSheetCell(sheet, row, fieldMap['bid.price_amount'], approvedPrice ?? null);
            setSheetCell(sheet, row, fieldMap['bid.price_currency'], bidCurrency ?? null);
            setSheetCell(sheet, row, fieldMap['bid.lead_time'], leadTime ?? null);
            setSheetCell(sheet, row, fieldMap['bid.supplier_item_name'], supplierItemName ?? null);
            setSheetCell(sheet, row, fieldMap['bid.item_part_number'], bidItemPartNumber ?? null);
            setSheetCell(sheet, row, fieldMap['bid.item_description'], bidItemDescription ?? null);
            const shipping = item['term_pembayaran'] || item['bid_shipping_term'];
            if (shipping)
                setSheetCell(sheet, row, fieldMap['bid.shipping_term'], shipping);
        }
        delete sheet['!autofilter'];
        output = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx', cellStyles: true });
        safeName = fileRow.file_name.endsWith('.xlsx') ? fileRow.file_name : fileRow.file_name + '.xlsx';
    }
    else {
        // No original file — generate a simple Excel with a derived filename
        output = buildCoupaFormatExcel(items);
        const baseName = generateExcelFilename(inquiryRow.customer);
        safeName = `${baseName}.xlsx`;
    }
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${safeName}"`);
    res.send(output);
});
// POST /inquiries
exports.inquiriesRouter.post('/', (req, res) => {
    const { customer, salesPic, namaBarang, spesifikasi, qty, itemUom, itemNeedByDate, itemManufacturerName, itemManufacturerPartNumber, itemClassificationOfGoods, itemImage, deadlineQuotation, lampiran, createdBy, createdByName } = req.body;
    if (!customer || !salesPic || !namaBarang || !createdBy) {
        res.status(400).json({ error: 'customer, salesPic, namaBarang, createdBy are required.' });
        return;
    }
    const id = (0, id_1.generateId)();
    const rfqNo = generateRfqNo();
    const tanggal = new Date().toISOString().split('T')[0];
    const createdAt = new Date().toISOString();
    const needByDate = itemNeedByDate ?? deadlineQuotation ?? null;
    db_1.db.prepare(`INSERT INTO inquiries (id, rfq_no, tanggal, customer, sales_pic, nama_barang, spesifikasi, qty, deadline_quotation, lampiran, status, created_at, created_by)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'new_inquiry', ?, ?)`).run(id, rfqNo, tanggal, customer, salesPic, namaBarang, spesifikasi ?? null, qty ?? null, needByDate, lampiran ?? null, createdAt, createdBy);
    db_1.db.prepare(`INSERT INTO inquiry_items (
      id, inquiry_id, item_name, item_quantity, item_uom, item_need_by_date,
      item_manufacturer_name, item_manufacturer_part_number, item_classification_of_goods,
      item_extended_description, item_image
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`).run((0, id_1.generateId)(), id, namaBarang, qty ?? null, itemUom ?? null, needByDate, itemManufacturerName ?? null, itemManufacturerPartNumber ?? null, itemClassificationOfGoods ?? null, spesifikasi ?? null, itemImage ?? null);
    logActivity(id, 'Inquiry created', null, 'new_inquiry', null, String(createdBy), String(createdByName ?? createdBy));
    res.status(201).json({ id, rfqNo, tanggal, status: 'new_inquiry', createdAt });
});
// PUT /inquiries/:id
exports.inquiriesRouter.put('/:id', (req, res) => {
    const { id } = req.params;
    const inquiry = db_1.db.prepare('SELECT id, status FROM inquiries WHERE id = ?').get(id);
    if (!inquiry) {
        res.status(404).json({ error: 'Not found.' });
        return;
    }
    if (!['new_inquiry', 'rfq'].includes(inquiry.status)) {
        res.status(400).json({ error: 'Cannot edit inquiry at this stage.' });
        return;
    }
    const { customer, salesPic, namaBarang, spesifikasi, qty, itemUom, itemNeedByDate, itemManufacturerName, itemManufacturerPartNumber, itemClassificationOfGoods, deadlineQuotation, lampiran, updatedBy, updatedByName } = req.body;
    const needByDate = itemNeedByDate ?? deadlineQuotation ?? null;
    db_1.db.prepare(`UPDATE inquiries SET
       customer = COALESCE(?, customer), sales_pic = COALESCE(?, sales_pic),
       nama_barang = COALESCE(?, nama_barang), spesifikasi = ?, qty = ?,
       deadline_quotation = ?, lampiran = ?,
       updated_at = ?, updated_by = ?
     WHERE id = ?`).run(customer ?? null, salesPic ?? null, namaBarang ?? null, spesifikasi ?? null, qty ?? null, needByDate, lampiran ?? null, new Date().toISOString(), updatedBy ?? null, id);
    const itemCount = db_1.db.prepare('SELECT COUNT(*) as c FROM inquiry_items WHERE inquiry_id = ?').get(id).c;
    if (itemCount === 1) {
        db_1.db.prepare(`UPDATE inquiry_items SET
         item_name = COALESCE(?, item_name),
         item_extended_description = ?,
         item_quantity = ?,
         item_uom = ?,
         item_need_by_date = ?,
         item_manufacturer_name = ?,
         item_manufacturer_part_number = ?,
         item_classification_of_goods = ?
       WHERE inquiry_id = ?`).run(namaBarang ?? null, spesifikasi ?? null, qty ?? null, itemUom ?? null, needByDate, itemManufacturerName ?? null, itemManufacturerPartNumber ?? null, itemClassificationOfGoods ?? null, id);
    }
    logActivity(id, 'Inquiry updated', inquiry.status, inquiry.status, null, String(updatedBy ?? ''), String(updatedByName ?? updatedBy ?? ''));
    res.json({ ok: true });
});
// POST /inquiries/:id/send-rfq — Sales qualifies and sends to Sourcing
exports.inquiriesRouter.post('/:id/send-rfq', (req, res) => {
    const { id } = req.params;
    const inquiry = db_1.db.prepare('SELECT id, status FROM inquiries WHERE id = ?').get(id);
    if (!inquiry) {
        res.status(404).json({ error: 'Not found.' });
        return;
    }
    if (inquiry.status !== 'new_inquiry') {
        res.status(400).json({ error: 'Only new inquiries can be sent to sourcing.' });
        return;
    }
    const { doneBy, doneByName, note } = req.body;
    db_1.db.prepare('UPDATE inquiries SET status = ?, updated_at = ?, updated_by = ? WHERE id = ?')
        .run('rfq', new Date().toISOString(), doneBy, id);
    logActivity(id, 'RFQ sent to Sourcing', 'new_inquiry', 'rfq', String(note ?? ''), String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''));
    res.json({ ok: true });
});
// POST /inquiries/:id/items — Marketing adds a new item
exports.inquiriesRouter.post('/:id/items', (req, res) => {
    const { id } = req.params;
    const inquiry = db_1.db.prepare('SELECT id, status FROM inquiries WHERE id = ?').get(id);
    if (!inquiry) {
        res.status(404).json({ error: 'Not found.' });
        return;
    }
    if (inquiry.status !== 'new_inquiry') {
        res.status(400).json({ error: 'Inquiry must be in new_inquiry status.' });
        return;
    }
    const { itemName, itemQuantity, itemUom, itemNeedByDate, itemManufacturerName, itemManufacturerPartNumber, itemClassificationOfGoods, itemExtendedDescription, itemImage, doneBy, doneByName } = req.body;
    if (!itemName) {
        res.status(400).json({ error: 'itemName is required.' });
        return;
    }
    const newId = (0, id_1.generateId)();
    db_1.db.prepare(`INSERT INTO inquiry_items (id, inquiry_id, item_name, item_quantity, item_uom, item_need_by_date,
      item_manufacturer_name, item_manufacturer_part_number, item_classification_of_goods,
      item_extended_description, item_image)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`).run(newId, id, itemName, itemQuantity ?? null, itemUom ?? null, itemNeedByDate ?? null, itemManufacturerName ?? null, itemManufacturerPartNumber ?? null, itemClassificationOfGoods ?? null, itemExtendedDescription ?? null, itemImage ?? null);
    logActivity(id, 'Item added by marketing', null, null, null, String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''));
    res.json({ ok: true, id: newId });
});
// PATCH /inquiries/:id/items/:itemId — Marketing updates target price / image / need-by-date (new_inquiry only)
exports.inquiriesRouter.patch('/:id/items/:itemId', (req, res) => {
    const { id, itemId } = req.params;
    const inquiry = db_1.db.prepare('SELECT id, status FROM inquiries WHERE id = ?').get(id);
    if (!inquiry) {
        res.status(404).json({ error: 'Not found.' });
        return;
    }
    if (inquiry.status !== 'new_inquiry') {
        res.status(400).json({ error: 'Inquiry must be in new_inquiry status.' });
        return;
    }
    const item = db_1.db.prepare('SELECT id FROM inquiry_items WHERE id = ? AND inquiry_id = ?').get(itemId, id);
    if (!item) {
        res.status(404).json({ error: 'Item not found.' });
        return;
    }
    const { itemImage, doneBy, doneByName } = req.body;
    db_1.db.prepare('UPDATE inquiry_items SET item_image = ? WHERE id = ?')
        .run(itemImage ?? null, itemId);
    logActivity(id, 'Item reviewed by marketing', null, null, null, String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''));
    res.json({ ok: true });
});
// PATCH /inquiries/:id/need-by-date — Sales updates RFQ-level need-by date from any status
exports.inquiriesRouter.patch('/:id/need-by-date', (req, res) => {
    const { id } = req.params;
    const inquiry = db_1.db.prepare('SELECT id FROM inquiries WHERE id = ?').get(id);
    if (!inquiry) {
        res.status(404).json({ error: 'Not found.' });
        return;
    }
    const { needByDate, doneBy, doneByName } = req.body;
    db_1.db.prepare('UPDATE inquiries SET deadline_quotation = ? WHERE id = ?')
        .run(needByDate ?? null, id);
    logActivity(id, 'Need-by date updated', null, null, null, String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''));
    res.json({ ok: true });
});
// PATCH /inquiries/:id/items/:itemId/harga-jual — Sales/Marketing updates selling price
exports.inquiriesRouter.patch('/:id/items/:itemId/harga-jual', (req, res) => {
    const { id, itemId } = req.params;
    const inquiry = db_1.db.prepare('SELECT id, status FROM inquiries WHERE id = ?').get(id);
    if (!inquiry) {
        res.status(404).json({ error: 'Not found.' });
        return;
    }
    if (!['price_approved', 'quotation_sent', 'deal'].includes(inquiry.status)) {
        res.status(400).json({ error: 'Inquiry must be in price_approved, quotation_sent, or deal status.' });
        return;
    }
    const item = db_1.db.prepare('SELECT id, harga_beli, approved_price, review_status FROM inquiry_items WHERE id = ? AND inquiry_id = ?').get(itemId, id);
    if (!item) {
        res.status(404).json({ error: 'Item not found.' });
        return;
    }
    const { hargaJual, doneBy, doneByName } = req.body;
    if (!hargaJual) {
        res.status(400).json({ error: 'hargaJual is required.' });
        return;
    }
    const nextPrice = Number(hargaJual);
    const margin = item.harga_beli != null ? nextPrice - item.harga_beli : null;
    const needsReview = item.approved_price != null ? (nextPrice < item.approved_price ? 1 : 0) : 0;
    const nextReviewStatus = item.review_status === 'rejected'
        ? 'rejected'
        : (needsReview ? 'review' : 'approved');
    db_1.db.prepare('UPDATE inquiry_items SET harga_jual = ?, margin = ?, needs_price_review = ?, review_status = ? WHERE id = ?')
        .run(nextPrice, margin, needsReview, nextReviewStatus, itemId);
    logActivity(id, 'Harga jual updated by marketing', null, null, null, String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''));
    res.json({ ok: true });
});
// POST /inquiries/:id/sourcing-info — Sourcing fills supplier data
exports.inquiriesRouter.post('/:id/sourcing-info', (req, res) => {
    const { id } = req.params;
    const inquiry = db_1.db.prepare('SELECT id, status FROM inquiries WHERE id = ?').get(id);
    if (!inquiry) {
        res.status(404).json({ error: 'Not found.' });
        return;
    }
    if (!['rfq', 'price_approval'].includes(inquiry.status)) {
        res.status(400).json({ error: 'Inquiry must be in RFQ status.' });
        return;
    }
    const { supplier, hargaBeli, leadTime, moq, stockAvailability, termPembayaran, doneBy, doneByName } = req.body;
    if (!supplier || hargaBeli === undefined || !leadTime) {
        res.status(400).json({ error: 'supplier, hargaBeli, leadTime are required.' });
        return;
    }
    const item = db_1.db.prepare('SELECT id, price_approved FROM inquiry_items WHERE inquiry_id = ? ORDER BY id LIMIT 1').get(id);
    if (!item) {
        res.status(400).json({ error: 'No items found.' });
        return;
    }
    if (item.price_approved) {
        res.status(400).json({ error: 'Item already approved, cannot edit.' });
        return;
    }
    db_1.db.prepare(`UPDATE inquiry_items SET supplier = ?, harga_beli = ?, lead_time = ?, moq = ?,
       stock_availability = ?, term_pembayaran = ? WHERE id = ?`).run(supplier, hargaBeli, leadTime, moq ?? null, stockAvailability ?? null, termPembayaran ?? null, item.id);
    logActivity(id, 'Sourcing info submitted', 'rfq', 'rfq', null, String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''));
    recalcInquiryStatus(id, String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''));
    res.json({ ok: true });
});
exports.inquiriesRouter.post('/:id/items/:itemId/sourcing-info', (req, res) => {
    const { id, itemId } = req.params;
    const inquiry = db_1.db.prepare('SELECT id, status FROM inquiries WHERE id = ?').get(id);
    if (!inquiry) {
        res.status(404).json({ error: 'Not found.' });
        return;
    }
    if (!['rfq', 'price_approval'].includes(inquiry.status)) {
        res.status(400).json({ error: 'Inquiry must be in RFQ status.' });
        return;
    }
    const item = db_1.db.prepare('SELECT id, price_approved, item_need_by_date FROM inquiry_items WHERE id = ? AND inquiry_id = ?').get(itemId, id);
    if (!item) {
        res.status(404).json({ error: 'Item not found.' });
        return;
    }
    if (item.price_approved) {
        res.status(400).json({ error: 'Item already approved, cannot edit.' });
        return;
    }
    const { supplier, hargaBeli, leadTime, moq, stockAvailability, termPembayaran, alternateName, doneBy, doneByName } = req.body;
    if (!supplier || hargaBeli === undefined || !leadTime) {
        res.status(400).json({ error: 'supplier, hargaBeli, leadTime are required.' });
        return;
    }
    // Mark as missed if submitted after the item's need-by date
    let sourcingMissed = 0;
    if (item.item_need_by_date) {
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        const needBy = new Date(item.item_need_by_date);
        needBy.setHours(0, 0, 0, 0);
        if (today > needBy)
            sourcingMissed = 1;
    }
    db_1.db.prepare(`UPDATE inquiry_items SET supplier = ?, harga_beli = ?, lead_time = ?, moq = ?,
       stock_availability = ?, term_pembayaran = ?, alternate_name = ?, sourcing_missed = ? WHERE id = ?`).run(supplier, hargaBeli, leadTime, moq ?? null, stockAvailability ?? null, termPembayaran ?? null, alternateName ?? null, sourcingMissed, itemId);
    logActivity(id, 'Sourcing info submitted', 'rfq', 'rfq', null, String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''));
    recalcInquiryStatus(id, String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''));
    res.json({ ok: true });
});
// POST /inquiries/:id/send-to-price-approval — Sourcing manually submits for price approval
exports.inquiriesRouter.post('/:id/send-to-price-approval', (req, res) => {
    const { id } = req.params;
    const inquiry = db_1.db.prepare('SELECT id, status FROM inquiries WHERE id = ?').get(id);
    if (!inquiry) {
        res.status(404).json({ error: 'Not found.' });
        return;
    }
    if (inquiry.status !== 'rfq') {
        res.status(400).json({ error: 'Inquiry must be in rfq status.' });
        return;
    }
    const { doneBy, doneByName, note } = req.body;
    if (!doneBy) {
        res.status(400).json({ error: 'doneBy is required.' });
        return;
    }
    db_1.db.prepare('UPDATE inquiries SET status = ?, updated_at = ?, updated_by = ? WHERE id = ?')
        .run('price_approval', new Date().toISOString(), String(doneBy), id);
    logActivity(id, 'Sent to Price Approval', 'rfq', 'price_approval', note ? String(note) : null, String(doneBy), String(doneByName ?? doneBy));
    res.json({ ok: true });
});
// POST /inquiries/:id/return-to-sourcing — Pricelist sends unfilled items back to sourcing
exports.inquiriesRouter.post('/:id/return-to-sourcing', (req, res) => {
    const { id } = req.params;
    const inquiry = db_1.db.prepare('SELECT id, status FROM inquiries WHERE id = ?').get(id);
    if (!inquiry) {
        res.status(404).json({ error: 'Not found.' });
        return;
    }
    if (inquiry.status !== 'price_approval') {
        res.status(400).json({ error: 'Inquiry must be in price_approval status.' });
        return;
    }
    const items = db_1.db.prepare('SELECT price_approved, harga_beli FROM inquiry_items WHERE inquiry_id = ?').all(id);
    const hasUnsourced = items.some((i) => !i.price_approved && !i.harga_beli);
    if (!hasUnsourced) {
        res.status(400).json({ error: 'No unfilled items to return to sourcing.' });
        return;
    }
    const { doneBy, doneByName } = req.body;
    if (!doneBy) {
        res.status(400).json({ error: 'doneBy is required.' });
        return;
    }
    db_1.db.prepare('UPDATE inquiries SET status = ?, updated_at = ?, updated_by = ? WHERE id = ?')
        .run('rfq', new Date().toISOString(), String(doneBy), id);
    logActivity(id, 'Returned to Sourcing', 'price_approval', 'rfq', null, String(doneBy), String(doneByName ?? doneBy));
    res.json({ ok: true });
});
// POST /inquiries/:id/send-to-sent — Marketing sends quotation to customer (price_approved → quotation_sent)
exports.inquiriesRouter.post('/:id/send-to-sent', (req, res) => {
    const { id } = req.params;
    const inquiry = db_1.db.prepare('SELECT id, status FROM inquiries WHERE id = ?').get(id);
    if (!inquiry) {
        res.status(404).json({ error: 'Not found.' });
        return;
    }
    if (inquiry.status !== 'price_approved') {
        res.status(400).json({ error: 'Inquiry must be in price_approved status.' });
        return;
    }
    const { doneBy, doneByName, incompleteReason } = req.body;
    if (!doneBy) {
        res.status(400).json({ error: 'doneBy is required.' });
        return;
    }
    const items = db_1.db.prepare('SELECT item_name, price_approved, review_status, needs_price_review, harga_jual, approved_price FROM inquiry_items WHERE inquiry_id = ?')
        .all(id);
    const unresolved = items.filter((i) => i.price_approved !== 1 ||
        i.review_status === 'rejected' ||
        i.needs_price_review === 1 ||
        (i.harga_jual != null && i.approved_price != null && i.harga_jual < i.approved_price));
    const isIncomplete = unresolved.length > 0;
    db_1.db.prepare('UPDATE inquiries SET status = ?, sent_incomplete = ?, sent_incomplete_reason = ?, updated_at = ?, updated_by = ? WHERE id = ?')
        .run('quotation_sent', isIncomplete ? 1 : 0, (isIncomplete && String(incompleteReason ?? '').trim()) ? String(incompleteReason).trim() : null, new Date().toISOString(), String(doneBy), id);
    const action = isIncomplete ? 'Quotation sent to customer (incomplete)' : 'Quotation sent to customer';
    const note = isIncomplete
        ? (String(incompleteReason ?? '').trim()
            ? `Sent with ${unresolved.length} unresolved item(s). ${String(incompleteReason).trim()}`
            : `Sent with ${unresolved.length} unresolved item(s).`)
        : null;
    logActivity(id, action, 'price_approved', 'quotation_sent', note, String(doneBy), String(doneByName ?? doneBy));
    res.json({ ok: true });
});
// POST /inquiries/:id/return-to-price-approval — Marketing sends back for price review (price_approved → price_approval)
exports.inquiriesRouter.post('/:id/return-to-price-approval', (req, res) => {
    const { id } = req.params;
    const inquiry = db_1.db.prepare('SELECT id, status FROM inquiries WHERE id = ?').get(id);
    if (!inquiry) {
        res.status(404).json({ error: 'Not found.' });
        return;
    }
    if (inquiry.status !== 'price_approved') {
        res.status(400).json({ error: 'Inquiry must be in price_approved status.' });
        return;
    }
    const { doneBy, doneByName, negotiationReason, reviewReason } = req.body;
    if (!doneBy) {
        res.status(400).json({ error: 'doneBy is required.' });
        return;
    }
    const reason = String(reviewReason ?? negotiationReason ?? '').trim();
    if (!reason) {
        res.status(400).json({ error: 'reviewReason is required.' });
        return;
    }
    const items = db_1.db.prepare(`SELECT id, price_approved, needs_price_review, review_status, review_round, harga_jual, approved_price
     FROM inquiry_items
     WHERE inquiry_id = ?`).all(id);
    const itemsNeedingReview = items.filter((item) => item.price_approved !== 1 ||
        item.harga_jual == null ||
        item.review_status === 'rejected' ||
        item.needs_price_review === 1 ||
        (item.harga_jual != null && item.approved_price != null && item.harga_jual < item.approved_price));
    if (!itemsNeedingReview.length) {
        res.status(400).json({ error: 'No item needs price review.' });
        return;
    }
    const reopeningRejectedCount = itemsNeedingReview.filter((item) => item.review_status === 'rejected').length;
    const reviewIds = new Set(itemsNeedingReview.map((item) => item.id));
    const setNeedsReview = db_1.db.prepare('UPDATE inquiry_items SET price_approved = 0, needs_price_review = 1, review_status = ?, review_round = ? WHERE id = ?');
    const keepApproved = db_1.db.prepare("UPDATE inquiry_items SET price_approved = 1, needs_price_review = 0, review_status = 'approved' WHERE id = ?");
    const insertItemNote = db_1.db.prepare('INSERT INTO inquiry_notes (id, inquiry_id, item_id, note, created_by, created_by_name, created_at) VALUES (?, ?, ?, ?, ?, ?, ?)');
    const applyReviewRouting = db_1.db.transaction(() => {
        const now = new Date().toISOString();
        for (const item of items) {
            if (reviewIds.has(item.id)) {
                const nextRound = Number(item.review_round ?? 0) + 1;
                setNeedsReview.run('review', nextRound, item.id);
                insertItemNote.run((0, id_1.generateId)(), id, item.id, `Price Review requested. Reason: ${reason}`, String(doneBy), String(doneByName ?? doneBy), now);
            }
            else {
                keepApproved.run(item.id);
            }
        }
    });
    applyReviewRouting();
    db_1.db.prepare('UPDATE inquiries SET status = ?, sourcing_pic = ?, updated_at = ?, updated_by = ? WHERE id = ?')
        .run('price_approval', String(doneByName ?? doneBy), new Date().toISOString(), String(doneBy), id);
    logActivity(id, `Returned to Price Approval for review (${itemsNeedingReview.length} item${itemsNeedingReview.length > 1 ? 's' : ''})`, 'price_approved', 'price_approval', `Items: ${itemsNeedingReview.length}. Reason: ${reason}${reopeningRejectedCount > 0 ? ' (includes rejected item negotiation reopen)' : ''}`, String(doneBy), String(doneByName ?? doneBy));
    res.json({ ok: true });
});
// POST /inquiries/:id/send-to-price-approved — Manager manually submits to Price Approved
exports.inquiriesRouter.post('/:id/send-to-price-approved', (req, res) => {
    const { id } = req.params;
    const inquiry = db_1.db.prepare('SELECT id, status FROM inquiries WHERE id = ?').get(id);
    if (!inquiry) {
        res.status(404).json({ error: 'Not found.' });
        return;
    }
    if (inquiry.status !== 'price_approval') {
        res.status(400).json({ error: 'Inquiry must be in price_approval status.' });
        return;
    }
    const { doneBy, doneByName } = req.body;
    if (!doneBy) {
        res.status(400).json({ error: 'doneBy is required.' });
        return;
    }
    db_1.db.prepare('UPDATE inquiries SET status = ?, updated_at = ?, updated_by = ? WHERE id = ?')
        .run('price_approved', new Date().toISOString(), String(doneBy), id);
    logActivity(id, 'Sent to Price Approved', 'price_approval', 'price_approved', null, String(doneBy), String(doneByName ?? doneBy));
    res.json({ ok: true });
});
// POST /inquiries/:id/approve — Manager approves price
exports.inquiriesRouter.post('/:id/approve', (req, res) => {
    const { id } = req.params;
    const inquiry = db_1.db.prepare('SELECT id, status FROM inquiries WHERE id = ?').get(id);
    if (!inquiry) {
        res.status(404).json({ error: 'Not found.' });
        return;
    }
    if (inquiry.status !== 'price_approval') {
        res.status(400).json({ error: 'Inquiry must be in price_approval status.' });
        return;
    }
    const { hargaJual, leadTimeCustomer, validitasQuotation, catatanQuotation, doneBy, doneByName } = req.body;
    if (!hargaJual) {
        res.status(400).json({ error: 'hargaJual is required.' });
        return;
    }
    const item = db_1.db.prepare('SELECT id, harga_beli FROM inquiry_items WHERE inquiry_id = ? ORDER BY id LIMIT 1').get(id);
    if (!item) {
        res.status(400).json({ error: 'No items found.' });
        return;
    }
    const margin = item.harga_beli != null ? Number(hargaJual) - item.harga_beli : null;
    db_1.db.prepare(`UPDATE inquiry_items SET harga_jual = ?, approved_price = ?, margin = ?, lead_time_customer = ?,
       validitas_quotation = ?, catatan_quotation = ?, price_approved = 1, needs_price_review = 0, review_status = 'approved' WHERE id = ?`).run(hargaJual, hargaJual, margin, leadTimeCustomer ?? null, validitasQuotation ?? null, catatanQuotation ?? null, item.id);
    logActivity(id, 'Price approved', 'price_approval', 'price_approval', String(catatanQuotation ?? ''), String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''));
    recalcInquiryStatus(id, String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''));
    res.json({ ok: true });
});
exports.inquiriesRouter.post('/:id/items/:itemId/approve', (req, res) => {
    const { id, itemId } = req.params;
    const inquiry = db_1.db.prepare('SELECT id, status FROM inquiries WHERE id = ?').get(id);
    if (!inquiry) {
        res.status(404).json({ error: 'Not found.' });
        return;
    }
    if (inquiry.status !== 'price_approval') {
        res.status(400).json({ error: 'Inquiry must be in price_approval status.' });
        return;
    }
    const item = db_1.db.prepare('SELECT id, harga_beli FROM inquiry_items WHERE id = ? AND inquiry_id = ?').get(itemId, id);
    if (!item) {
        res.status(404).json({ error: 'Item not found.' });
        return;
    }
    const { hargaJual, leadTimeCustomer, validitasQuotation, catatanQuotation, doneBy, doneByName } = req.body;
    if (!hargaJual) {
        res.status(400).json({ error: 'hargaJual is required.' });
        return;
    }
    const margin = item.harga_beli != null ? Number(hargaJual) - item.harga_beli : null;
    db_1.db.prepare(`UPDATE inquiry_items SET harga_jual = ?, approved_price = ?, margin = ?, lead_time_customer = ?,
       validitas_quotation = ?, catatan_quotation = ?, price_approved = 1, needs_price_review = 0, review_status = 'approved' WHERE id = ?`).run(hargaJual, hargaJual, margin, leadTimeCustomer ?? null, validitasQuotation ?? null, catatanQuotation ?? null, itemId);
    logActivity(id, 'Price approved', 'price_approval', 'price_approval', String(catatanQuotation ?? ''), String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''));
    recalcInquiryStatus(id, String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''));
    res.json({ ok: true });
});
// POST /inquiries/:id/items/:itemId/reject — Manager rejects a price proposal
exports.inquiriesRouter.post('/:id/items/:itemId/reject', (req, res) => {
    const { id, itemId } = req.params;
    const inquiry = db_1.db.prepare('SELECT id, status FROM inquiries WHERE id = ?').get(id);
    if (!inquiry) {
        res.status(404).json({ error: 'Not found.' });
        return;
    }
    if (inquiry.status !== 'price_approval') {
        res.status(400).json({ error: 'Inquiry must be in price_approval status.' });
        return;
    }
    const item = db_1.db.prepare('SELECT id, review_round, item_name, review_status, needs_price_review FROM inquiry_items WHERE id = ? AND inquiry_id = ?').get(itemId, id);
    if (!item) {
        res.status(404).json({ error: 'Item not found.' });
        return;
    }
    if (item.review_status !== 'review' && item.needs_price_review !== 1) {
        res.status(400).json({ error: 'Only items in review status can be rejected.' });
        return;
    }
    const { doneBy, doneByName, reason } = req.body;
    const rejectReason = String(reason ?? '').trim();
    if (!doneBy) {
        res.status(400).json({ error: 'doneBy is required.' });
        return;
    }
    if (!rejectReason) {
        res.status(400).json({ error: 'reason is required.' });
        return;
    }
    const nextRound = Number(item.review_round ?? 0) + 1;
    db_1.db.prepare(`UPDATE inquiry_items
     SET price_approved = 0,
         needs_price_review = 1,
         review_status = 'rejected',
         review_round = ?
     WHERE id = ?`).run(nextRound, itemId);
    logActivity(id, `Price rejected (${item.item_name ?? 'Item'})`, 'price_approval', 'price_approval', rejectReason, String(doneBy), String(doneByName ?? doneBy));
    db_1.db.prepare(`INSERT INTO inquiry_notes (id, inquiry_id, item_id, note, created_by, created_by_name, created_at)
     VALUES (?, ?, ?, ?, ?, ?, ?)`).run((0, id_1.generateId)(), id, itemId, `Reject reason: ${rejectReason}`, String(doneBy), String(doneByName ?? doneBy), new Date().toISOString());
    res.json({ ok: true });
});
// POST /inquiries/:id/close — Sales or Manager closes as deal/lost
exports.inquiriesRouter.post('/:id/close', (req, res) => {
    const { id } = req.params;
    const inquiry = db_1.db.prepare('SELECT id, status FROM inquiries WHERE id = ?').get(id);
    if (!inquiry) {
        res.status(404).json({ error: 'Not found.' });
        return;
    }
    const closeableStatuses = ['new_inquiry', 'rfq', 'quotation_sent'];
    if (!closeableStatuses.includes(inquiry.status)) {
        res.status(400).json({ error: 'Cannot close inquiry at this stage.' });
        return;
    }
    const { outcome, doneBy, doneByName, note } = req.body;
    if (!['deal', 'lost'].includes(String(outcome))) {
        res.status(400).json({ error: 'outcome must be deal or lost.' });
        return;
    }
    const oldStatus = inquiry.status;
    db_1.db.prepare('UPDATE inquiries SET status = ?, updated_at = ?, updated_by = ? WHERE id = ?')
        .run(outcome, new Date().toISOString(), doneBy, id);
    logActivity(id, `Closed as ${String(outcome)}`, oldStatus, String(outcome), String(note ?? ''), String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''));
    res.json({ ok: true });
});
// POST /inquiries/:id/ready-to-purchase — move a deal to ready_to_purchase
exports.inquiriesRouter.post('/:id/ready-to-purchase', (req, res) => {
    const { id } = req.params;
    const inquiry = db_1.db.prepare('SELECT id, status FROM inquiries WHERE id = ?').get(id);
    if (!inquiry) {
        res.status(404).json({ error: 'Not found.' });
        return;
    }
    if (inquiry.status !== 'deal') {
        res.status(400).json({ error: 'Only deal inquiries can be moved to Ready to Purchase.' });
        return;
    }
    const { doneBy, doneByName } = req.body;
    db_1.db.prepare('UPDATE inquiries SET status = ?, updated_at = ?, updated_by = ? WHERE id = ?')
        .run('ready_to_purchase', new Date().toISOString(), doneBy, id);
    logActivity(id, 'Moved to Ready to Purchase', 'deal', 'ready_to_purchase', null, String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''));
    res.json({ ok: true });
});
// GET /inquiries/:id/notes
exports.inquiriesRouter.get('/:id/notes', (req, res) => {
    const { id } = req.params;
    const notes = db_1.db.prepare('SELECT id, inquiry_id, item_id, note, created_by, created_by_name, created_at FROM inquiry_notes WHERE inquiry_id = ? ORDER BY created_at ASC').all(id);
    res.json(notes.map((n) => ({
        id: n['id'],
        inquiryId: n['inquiry_id'],
        itemId: n['item_id'] ?? null,
        note: n['note'],
        createdBy: n['created_by'],
        createdByName: n['created_by_name'],
        createdAt: n['created_at'],
    })));
});
// POST /inquiries/:id/notes
exports.inquiriesRouter.post('/:id/notes', (req, res) => {
    const { id } = req.params;
    const { note, doneBy, doneByName, role } = req.body;
    if (!note || !String(note).trim()) {
        res.status(400).json({ error: 'Note cannot be empty.' });
        return;
    }
    const inquiry = db_1.db.prepare('SELECT id, sales_pic, sourcing_pic FROM inquiries WHERE id = ?')
        .get(id);
    if (!inquiry) {
        res.status(404).json({ error: 'Not found.' });
        return;
    }
    const isAdminOrManager = role === 'admin' || role === 'manager';
    const isAssigned = doneByName === inquiry.sales_pic || doneByName === inquiry.sourcing_pic;
    if (!isAdminOrManager && !isAssigned) {
        res.status(403).json({ error: 'Only assigned users can add comments.' });
        return;
    }
    const noteId = (0, id_1.generateId)();
    db_1.db.prepare('INSERT INTO inquiry_notes (id, inquiry_id, note, created_by, created_by_name, created_at) VALUES (?, ?, ?, ?, ?, ?)').run(noteId, id, String(note).trim(), doneBy, doneByName, new Date().toISOString());
    res.json({ id: noteId });
});
// GET /inquiries/:id/items/:itemId/notes
exports.inquiriesRouter.get('/:id/items/:itemId/notes', (req, res) => {
    const { id, itemId } = req.params;
    const notes = db_1.db.prepare('SELECT id, inquiry_id, item_id, note, created_by, created_by_name, created_at FROM inquiry_notes WHERE inquiry_id = ? AND item_id = ? ORDER BY created_at ASC').all(id, itemId);
    res.json(notes.map((n) => ({
        id: n['id'],
        inquiryId: n['inquiry_id'],
        itemId: n['item_id'],
        note: n['note'],
        createdBy: n['created_by'],
        createdByName: n['created_by_name'],
        createdAt: n['created_at'],
    })));
});
// POST /inquiries/:id/items/:itemId/notes
exports.inquiriesRouter.post('/:id/items/:itemId/notes', (req, res) => {
    const { id, itemId } = req.params;
    const { note, doneBy, doneByName, role } = req.body;
    if (!note || !String(note).trim()) {
        res.status(400).json({ error: 'Note cannot be empty.' });
        return;
    }
    const inquiry = db_1.db.prepare('SELECT id, sales_pic, sourcing_pic FROM inquiries WHERE id = ?')
        .get(id);
    if (!inquiry) {
        res.status(404).json({ error: 'Not found.' });
        return;
    }
    const isAdminOrManager = role === 'admin' || role === 'manager';
    const isAssigned = doneByName === inquiry.sales_pic || doneByName === inquiry.sourcing_pic;
    if (!isAdminOrManager && !isAssigned) {
        res.status(403).json({ error: 'Only assigned users can add comments.' });
        return;
    }
    const noteId = (0, id_1.generateId)();
    db_1.db.prepare('INSERT INTO inquiry_notes (id, inquiry_id, item_id, note, created_by, created_by_name, created_at) VALUES (?, ?, ?, ?, ?, ?, ?)').run(noteId, id, itemId, String(note).trim(), doneBy, doneByName, new Date().toISOString());
    res.json({ id: noteId });
});
// PATCH /inquiries/:id/assign-sales — admin/manager only
exports.inquiriesRouter.patch('/:id/assign-sales', (req, res) => {
    const { id } = req.params;
    const { salesPic, doneBy, doneByName, role } = req.body;
    if (role !== 'admin' && role !== 'manager') {
        res.status(403).json({ error: 'Only admin or manager can reassign Sales PIC.' });
        return;
    }
    if (!salesPic || !String(salesPic).trim()) {
        res.status(400).json({ error: 'salesPic is required.' });
        return;
    }
    const inquiry = db_1.db.prepare('SELECT id FROM inquiries WHERE id = ?').get(id);
    if (!inquiry) {
        res.status(404).json({ error: 'Not found.' });
        return;
    }
    db_1.db.prepare('UPDATE inquiries SET sales_pic = ?, updated_at = ?, updated_by = ? WHERE id = ?')
        .run(String(salesPic).trim(), new Date().toISOString(), doneBy, id);
    logActivity(id, `Sales PIC reassigned to: ${String(salesPic)}`, null, null, '', String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''));
    res.json({ ok: true });
});
// PATCH /inquiries/:id/assign-sourcing
// Admin/manager can assign any sourcing user; sourcing can self-assign if unassigned
exports.inquiriesRouter.patch('/:id/assign-sourcing', (req, res) => {
    const { id } = req.params;
    const { sourcingPic, doneBy, doneByName, role } = req.body;
    const inquiry = db_1.db.prepare('SELECT id, sourcing_pic FROM inquiries WHERE id = ?').get(id);
    if (!inquiry) {
        res.status(404).json({ error: 'Not found.' });
        return;
    }
    const isSourcing = role === 'sourcing';
    const isAdminOrManager = role === 'admin' || role === 'manager';
    if (!isAdminOrManager && !isSourcing) {
        res.status(403).json({ error: 'Not authorized.' });
        return;
    }
    // Sourcing can only self-assign and only if not already assigned
    if (isSourcing && !isAdminOrManager) {
        if (inquiry.sourcing_pic) {
            res.status(400).json({ error: 'Already assigned to another sourcing user.' });
            return;
        }
        if (String(sourcingPic) !== String(doneByName)) {
            res.status(403).json({ error: 'Sourcing can only assign themselves.' });
            return;
        }
    }
    db_1.db.prepare('UPDATE inquiries SET sourcing_pic = ?, updated_at = ?, updated_by = ? WHERE id = ?')
        .run(sourcingPic ?? null, new Date().toISOString(), doneBy, id);
    logActivity(id, `Sourcing assigned: ${String(sourcingPic ?? 'unassigned')}`, null, null, '', String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''));
    res.json({ ok: true });
});
