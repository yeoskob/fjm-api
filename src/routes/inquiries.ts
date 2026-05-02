import { Router, Request, Response } from 'express';
import { db } from '../db';
import { generateId } from '../utils/id';
import * as XLSX from 'xlsx';
import { insertAndBroadcast, usernameForPic } from './notifications';

export const inquiriesRouter = Router();

function generateRfqNo(): string {
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  const rand = Math.floor(Math.random() * 9000) + 1000;
  return `RFQ-${y}${m}${day}-${rand}`;
}

function logActivity(
  inquiryId: string,
  action: string,
  oldStatus: string | null,
  newStatus: string | null,
  note: string | null,
  doneBy: string,
  doneByName: string
): void {
  db.prepare(
    `INSERT INTO activity_log (id, inquiry_id, action, old_status, new_status, note, done_by, done_by_name, created_at)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
  ).run(generateId(), inquiryId, action, oldStatus, newStatus, note, doneBy, doneByName, new Date().toISOString());
}

function mapItem(row: Record<string, unknown>) {
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
    supplierUrl: row['supplier_url'],
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
    reviewStatus: row['review_status'] === 'rejected' ? 'review' : (row['review_status'] ?? 'pending'),
    reviewRound: Number(row['review_round'] ?? 0),
    sourcingMissed: row['sourcing_missed'] === 1,
    ppnType: row['ppn_type'] ?? null,
  };
}

function mapInquiry(row: Record<string, unknown>, items: Array<Record<string, unknown>>) {
  const itemNeedByDate = items
    .map((item) => item['item_need_by_date'])
    .filter((date): date is string => typeof date === 'string' && date !== '')
    .sort()[0] ?? null;
  const needByDate = row['deadline_quotation'] || itemNeedByDate;

  return {
    id: row['id'],
    rfqNo: row['rfq_no'],
    tanggal: row['tanggal'],
    customer: row['customer'],
    salesPic: row['sales_pic'],
    sourcingPic: row['sourcing_pic'] ?? null,
    status: row['status'],
    coupaSource: row['coupa_source'] === 1,
    organization: row['organization'] ?? 'FJM',
    coupaFileName: row['coupa_file_name'] ?? null,
    createdAt: row['created_at'],
    createdBy: row['created_by'],
    updatedAt: row['updated_at'],
    updatedBy: row['updated_by'],
    sentIncomplete: row['sent_incomplete'] === 1,
    sentIncompleteReason: row['sent_incomplete_reason'] ?? null,
    needByDate,
    sourcingMissed: row['sourcing_missed'] === 1,
    priceApprovalStartedAt: row['price_approval_started_at'] ?? null,
    items: items.map(mapItem),
  };
}

function normalizeOrganization(value: unknown): string | null {
  const v = String(value ?? '').trim().toUpperCase();
  if (!v) return null;
  const exists = db.prepare('SELECT id FROM organizations WHERE code = ?').get(v) as { id: string } | undefined;
  return exists ? v : null;
}

function deriveCustomerFromFilename(fileName: string): string {
  const clean = fileName.replace(/\.[^.]+$/, '');
  const parts = clean.split('#').map((p) => p.trim()).filter(Boolean);
  if (parts.length >= 2) {
    const name = parts[1].replace(/[-_]+/g, ' ').trim();
    const eventId = parts[2] ? parts[2].trim() : null;
    return eventId ? `${name} #${eventId}` : name;
  }
  return clean.replace(/[-_]+/g, ' ').trim();
}

function parseCoupaFieldMap(sheet: XLSX.WorkSheet): Record<string, number> {
  const map: Record<string, number> = {};
  const range = sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : null;
  if (!range) {
    return map;
  }

  // Primary: read JSON field_name metadata from row 0
  for (let c = range.s.c; c <= range.e.c; c += 1) {
    const cellAddress = XLSX.utils.encode_cell({ r: 0, c });
    const cell = sheet[cellAddress];
    if (!cell || cell.v == null) continue;
    const raw = String(cell.v);
    if (!raw.includes('field_name')) continue;
    try {
      const parsed = JSON.parse(raw);
      if (parsed?.field_name) {
        map[parsed.field_name] = c;
      }
    } catch { /* ignore */ }
  }

  // Fallback: scan row 4 (human-readable column labels) for known fields not found via JSON
  const labelFallbacks: Array<[string, string]> = [
    ['item.need_by_date', 'need by date'],
  ];
  for (const [fieldName, labelSubstr] of labelFallbacks) {
    if (map[fieldName] != null) continue; // already mapped
    for (let c = range.s.c; c <= range.e.c; c += 1) {
      const cell = sheet[XLSX.utils.encode_cell({ r: 4, c })];
      if (!cell || cell.v == null) continue;
      if (String(cell.v).toLowerCase().includes(labelSubstr)) {
        map[fieldName] = c;
        break;
      }
    }
  }

  return map;
}

function readSheetCell(sheet: XLSX.WorkSheet, rowIndex: number, colIndex?: number) {
  if (colIndex == null) return null;
  const cell = sheet[XLSX.utils.encode_cell({ r: rowIndex, c: colIndex })];
  if (!cell) return null;
  return cell.v ?? null;
}

function setSheetCell(sheet: XLSX.WorkSheet, rowIndex: number, colIndex?: number, value?: string | number | null) {
  if (colIndex == null || value == null || value === '') return;
  const addr = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
  const cell = sheet[addr] ?? {};
  cell.v = value;
  cell.t = typeof value === 'number' ? 'n' : 's';
  sheet[addr] = cell as XLSX.CellObject;
}

function parseExcelDate(value: unknown): string | null {
  if (value == null || value === '') return null;
  // Excel date serial number
  const num = typeof value === 'number' ? value : Number(value);
  if (!isNaN(num) && num > 40000) {
    const date = new Date((num - 25569) * 86400 * 1000);
    return isNaN(date.getTime()) ? null : date.toISOString().slice(0, 10);
  }
  if (typeof value === 'string') {
    // Already an ISO date string stored in DB — return date part only
    if (/^\d{4}-\d{2}-\d{2}/.test(value)) return value.slice(0, 10);
    // Try general date parse (handles "30-Mar-2026", "2026/03/30", etc.)
    const d = new Date(value);
    return isNaN(d.getTime()) ? null : d.toISOString().slice(0, 10);
  }
  return null;
}

function normalizeLeadTime(value?: string | number | null): number | null {
  if (value == null) return null;
  if (typeof value === 'number') return Number.isFinite(value) ? value : null;
  const trimmed = String(value).trim();
  if (!trimmed) return null;
  const num = Number.parseFloat(trimmed.replace(/[^0-9.]/g, ''));
  return Number.isFinite(num) ? num : null;
}



function recalcInquiryStatus(inquiryId: string, doneBy: string, doneByName: string) {
  // Manual transition is now required for price_approval -> price_approved.
  // Keep this function as a no-op because older call sites still invoke it.
  void inquiryId;
  void doneBy;
  void doneByName;
}

function autoMarkMissedRfqs(): void {
  db.prepare(`
    UPDATE inquiries
    SET sourcing_missed = 1, status = 'missed', updated_at = datetime('now')
    WHERE sourcing_missed = 0
      AND status = 'rfq'
      AND EXISTS (
        SELECT 1 FROM inquiry_items ii
        WHERE ii.inquiry_id = inquiries.id AND ii.item_need_by_date IS NOT NULL
          AND date(ii.item_need_by_date) < date('now')
      )
      AND (
        (sourcing_pic IS NULL OR sourcing_pic = '')
        OR NOT EXISTS (
          SELECT 1 FROM inquiry_items ii
          WHERE ii.inquiry_id = inquiries.id
            AND COALESCE(ii.supplier,'') NOT IN ('', '-')
            AND ii.harga_beli IS NOT NULL AND ii.harga_beli > 0
            AND COALESCE(ii.lead_time,'') != ''
            AND ii.ppn_type IS NOT NULL
        )
      )
  `).run();
}

function autoMarkUnsentRfqs(): void {
  db.prepare(`
    UPDATE inquiries
    SET status = 'unsent', updated_at = datetime('now')
    WHERE status = 'price_approved'
      AND EXISTS (
        SELECT 1 FROM inquiry_items
        WHERE inquiry_id = inquiries.id
          AND item_need_by_date IS NOT NULL
          AND date(item_need_by_date) < date('now')
      )
  `).run();
}

// GET /inquiries/report
inquiriesRouter.get('/report', (req: Request, res: Response) => {
  autoMarkMissedRfqs();
  autoMarkUnsentRfqs();

  const { month, salesPic } = req.query as { month?: string; salesPic?: string };

  const params: unknown[] = [];
  let where = `WHERE 1=1`;
  if (salesPic) { where += ` AND i.sales_pic = ?`; params.push(salesPic); }
  if (month)    { where += ` AND strftime('%Y-%m', i.tanggal) = ?`; params.push(month); }

  const rows = db.prepare(`
    SELECT
      i.id, i.rfq_no, i.customer, i.sales_pic, i.sourcing_pic, i.tanggal, i.status,
      COALESCE(NULLIF(i.deadline_quotation, ''), MIN(ii.item_need_by_date)) AS need_by_date,
      CAST(julianday(COALESCE(NULLIF(i.deadline_quotation, ''), MIN(ii.item_need_by_date))) - julianday(date(i.tanggal)) AS INTEGER) AS timeline_days,
      CAST(julianday(date(al_sent.sent_at, '+7 hours')) - julianday(date(i.tanggal)) AS INTEGER) AS days_taken
    FROM inquiries i
    LEFT JOIN inquiry_items ii
      ON ii.inquiry_id = i.id AND ii.item_need_by_date IS NOT NULL
    LEFT JOIN (
      SELECT inquiry_id, MIN(created_at) AS sent_at FROM activity_log
      WHERE action LIKE 'Quotation sent%' GROUP BY inquiry_id
    ) al_sent ON al_sent.inquiry_id = i.id
    ${where}
    GROUP BY i.id
    ORDER BY i.tanggal DESC
  `).all(...params) as Array<Record<string, unknown>>;

  res.json({ rows });
});

// GET /inquiries/report/sourcing
inquiriesRouter.get('/report/sourcing', (req: Request, res: Response) => {
  autoMarkMissedRfqs();
  autoMarkUnsentRfqs();

  const { month, sourcingPic } = req.query as { month?: string; sourcingPic?: string };

  const params: unknown[] = [];
  let where = `WHERE 1=1`;
  if (sourcingPic) { where += ` AND i.sourcing_pic = ?`; params.push(sourcingPic); }
  if (month)       { where += ` AND strftime('%Y-%m', i.tanggal) = ?`; params.push(month); }

  const rows = db.prepare(`
    SELECT
      i.id, i.rfq_no, i.customer, i.sales_pic, i.sourcing_pic, i.tanggal, i.status,
      COALESCE(NULLIF(i.deadline_quotation, ''), MIN(ii.item_need_by_date)) AS need_by_date,
      COUNT(ii.id) AS total_items,
      SUM(CASE WHEN COALESCE(ii.supplier,'') != '' AND ii.harga_beli IS NOT NULL AND COALESCE(ii.lead_time,'') != '' THEN 1 ELSE 0 END) AS sourced_items
    FROM inquiries i
    LEFT JOIN inquiry_items ii ON ii.inquiry_id = i.id
    ${where}
    GROUP BY i.id
    ORDER BY i.tanggal DESC
  `).all(...params) as Array<Record<string, unknown>>;

  res.json({ rows });
});

// GET /inquiries/report/export
inquiriesRouter.get('/report/export', (req: Request, res: Response) => {
  autoMarkMissedRfqs();
  autoMarkUnsentRfqs();

  const { month, salesPic, status, search, audience, dateField, dateFrom, dateTo } = req.query as {
    month?: string; salesPic?: string; status?: string; search?: string; audience?: string;
    dateField?: string; dateFrom?: string; dateTo?: string;
  };
  const isPurchasingExport = audience === 'purchasing';

  const params: unknown[] = [];
  let where = `WHERE 1=1`;
  if (salesPic) { where += ` AND i.sales_pic = ?`; params.push(salesPic); }
  if (month)    { where += ` AND strftime('%Y-%m', i.tanggal) = ?`; params.push(month); }
  if (status)   { where += ` AND i.status = ?`; params.push(status); }
  if (search && String(search).trim()) {
    const like = `%${String(search).trim()}%`;
    where += ` AND (i.rfq_no LIKE ? OR i.customer LIKE ?)`;
    params.push(like, like);
  }
  if (dateFrom || dateTo) {
    const dateExpr = dateField === 'need_by_date'
      ? `COALESCE(NULLIF(i.deadline_quotation, ''), (SELECT MIN(item_need_by_date) FROM inquiry_items WHERE inquiry_id = i.id))`
      : `date(i.tanggal)`;
    if (dateFrom) { where += ` AND ${dateExpr} >= ?`; params.push(String(dateFrom)); }
    if (dateTo)   { where += ` AND ${dateExpr} <= ?`; params.push(String(dateTo)); }
  }

  const summaryRows = db.prepare(`
    SELECT
      i.id, i.rfq_no, i.customer, i.sales_pic, i.sourcing_pic, i.tanggal, i.status,
      COALESCE(NULLIF(i.deadline_quotation, ''), MIN(ii.item_need_by_date)) AS need_by_date,
      CAST(julianday(COALESCE(NULLIF(i.deadline_quotation, ''), MIN(ii.item_need_by_date))) - julianday(date(i.tanggal)) AS INTEGER) AS timeline_days,
      CAST(julianday(date(al_sent.sent_at, '+7 hours')) - julianday(date(i.tanggal)) AS INTEGER) AS days_taken
    FROM inquiries i
    LEFT JOIN inquiry_items ii ON ii.inquiry_id = i.id AND ii.item_need_by_date IS NOT NULL
    LEFT JOIN (
      SELECT inquiry_id, MIN(created_at) AS sent_at FROM activity_log
      WHERE action LIKE 'Quotation sent%' GROUP BY inquiry_id
    ) al_sent ON al_sent.inquiry_id = i.id
    ${where}
    GROUP BY i.id ORDER BY i.tanggal DESC
  `).all(...params) as Array<Record<string, unknown>>;

  const ids = summaryRows.map((r) => r['id'] as string);
  const itemRows = ids.length > 0
    ? db.prepare(`
        SELECT i.rfq_no, i.customer, i.sales_pic, i.sourcing_pic, i.tanggal,
          ii.item_name, ii.item_quantity, ii.item_uom, ii.item_need_by_date,
          ii.supplier, ii.supplier_url, ii.harga_beli, ii.harga_jual, ii.margin,
          ii.lead_time, ii.moq, ii.stock_availability, ii.ppn_type
        FROM inquiry_items ii
        JOIN inquiries i ON i.id = ii.inquiry_id
        WHERE ii.inquiry_id IN (${ids.map(() => '?').join(',')})
        ORDER BY i.tanggal DESC, ii.coupa_row_index ASC, ii.id ASC
      `).all(...ids) as Array<Record<string, unknown>>
    : [];

  const statusLabels: Record<string, string> = {
    new_inquiry: 'New Inquiry', rfq: 'RFQ to Sourcing',
    price_approval: 'Price Approval', price_approved: 'Price Approved',
    quotation_sent: 'Quotation Sent', follow_up: 'Negotiation',
    ready_to_purchase: 'Ready to Purchase', missed: 'Missed', unsent: 'Unsent',
  };

  const dateFieldLabel = dateField === 'need_by_date' ? 'Need By Date' : 'Inquiry Date';
  const dateRangeStr = dateFrom || dateTo
    ? `${dateFieldLabel}: ${dateFrom || '…'} to ${dateTo || '…'}`
    : '';
  const metadata: unknown[][] = [
    [`Exported: ${new Date().toISOString().slice(0, 19).replace('T', ' ')}`],
  ];
  if (month)     metadata.push([`Month: ${month}`]);
  if (salesPic)  metadata.push([`Sales PIC: ${salesPic}`]);
  if (status)    metadata.push([`Status: ${statusLabels[String(status)] ?? status}`]);
  if (dateRangeStr) metadata.push([dateRangeStr]);
  metadata.push([]);

  // Sheet 1: Summary
  const summaryData = [
    ...metadata,
    ['RFQ No', 'Customer', 'Sales PIC', 'Sourcing PIC', 'Inquiry Date', 'Need By Date', 'Timeline (days)', 'Days Taken', 'Status'],
    ...summaryRows.map((r) => [
      r['rfq_no'], r['customer'], r['sales_pic'], r['sourcing_pic'] ?? '',
      r['tanggal'] ? String(r['tanggal']).slice(0, 10) : '',
      r['need_by_date'] ? String(r['need_by_date']).slice(0, 10) : '',
      r['timeline_days'] ?? '',
      r['days_taken'] ?? '',
      statusLabels[String(r['status'] ?? '')] ?? r['status'],
    ]),
  ];

  // Sheet 2: Items
  const itemHeaders = isPurchasingExport
    ? ['RFQ No', 'Customer', 'Sales PIC', 'Sourcing PIC', 'Inquiry Date', 'Item Name', 'Qty', 'UOM', 'Need By Date', 'Supplier', 'Supplier URL', 'Harga Beli', 'Lead Time', 'MOQ', 'Stock', 'PPN Type']
    : ['RFQ No', 'Customer', 'Sales PIC', 'Sourcing PIC', 'Inquiry Date', 'Item Name', 'Qty', 'UOM', 'Need By Date', 'Supplier', 'Harga Beli', 'Harga Jual', 'Margin (%)', 'Lead Time', 'MOQ', 'Stock', 'PPN Type'];

  const itemsData = [
    ...metadata,
    itemHeaders,
    ...itemRows.map((r) => (
      isPurchasingExport
        ? [
            r['rfq_no'], r['customer'], r['sales_pic'], r['sourcing_pic'] ?? '',
            r['tanggal'] ? String(r['tanggal']).slice(0, 10) : '',
            r['item_name'] ?? '', r['item_quantity'] ?? '', r['item_uom'] ?? '',
            r['item_need_by_date'] ? String(r['item_need_by_date']).slice(0, 10) : '',
            r['supplier'] ?? '', r['supplier_url'] ?? '', r['harga_beli'] ?? '',
            r['lead_time'] ?? '', r['moq'] ?? '',
            r['stock_availability'] ?? '', r['ppn_type'] ?? '',
          ]
        : [
            r['rfq_no'], r['customer'], r['sales_pic'], r['sourcing_pic'] ?? '',
            r['tanggal'] ? String(r['tanggal']).slice(0, 10) : '',
            r['item_name'] ?? '', r['item_quantity'] ?? '', r['item_uom'] ?? '',
            r['item_need_by_date'] ? String(r['item_need_by_date']).slice(0, 10) : '',
            r['supplier'] ?? '', r['harga_beli'] ?? '', r['harga_jual'] ?? '',
            r['margin'] ?? '', r['lead_time'] ?? '', r['moq'] ?? '',
            r['stock_availability'] ?? '', r['ppn_type'] ?? '',
          ]
    )),
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summaryData), 'Summary');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(itemsData), 'Items');

  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }) as Buffer;
  const filename = `report${month ? '-' + month : ''}.xlsx`;
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  res.send(buf);
});

inquiriesRouter.get('/', (_req: Request, res: Response) => {
  autoMarkMissedRfqs();
  autoMarkUnsentRfqs();
  const rows = db.prepare('SELECT * FROM inquiries ORDER BY created_at DESC').all() as Array<Record<string, unknown>>;
  const items = db.prepare('SELECT * FROM inquiry_items ORDER BY coupa_row_index ASC, id ASC').all() as Array<Record<string, unknown>>;
  const logs = db.prepare('SELECT * FROM activity_log ORDER BY created_at ASC').all() as Array<Record<string, unknown>>;

  const itemsByInquiry = new Map<string, Array<Record<string, unknown>>>();
  for (const item of items) {
    const id = String(item['inquiry_id']);
    const list = itemsByInquiry.get(id) ?? [];
    list.push(item);
    itemsByInquiry.set(id, list);
  }

  const logsByInquiry = new Map<string, unknown[]>();
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
inquiriesRouter.get('/dashboard/user', (req: Request, res: Response) => {
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
  const total = (db.prepare(`SELECT COUNT(*) as c FROM inquiries WHERE sales_pic = ? AND status != 'missed'`).get(name) as { c: number }).c;
  const thisMonthSales = (db.prepare(`SELECT COUNT(*) as c FROM inquiries WHERE sales_pic = ? AND created_at >= ? AND status != 'missed'`).get(name, startOfMonthIso) as { c: number }).c;
  const quotationSent = (db.prepare(`SELECT COUNT(*) as c FROM inquiries WHERE sales_pic = ? AND status IN ('quotation_sent','ready_to_purchase')`).get(name) as { c: number }).c;
  const unsent = (db.prepare(
    `SELECT COUNT(*) as c FROM inquiries WHERE sales_pic = ? AND status = 'unsent'`
  ).get(name) as { c: number }).c;
  const active = (db.prepare(`SELECT COUNT(*) as c FROM inquiries WHERE sales_pic = ? AND status NOT IN ('quotation_sent','ready_to_purchase')`).get(name) as { c: number }).c;
  const conversionRate = total > 0 ? +((quotationSent / total) * 100).toFixed(1) : 0;
  const statusBreakdown = db.prepare(
    'SELECT status, COUNT(*) as count FROM inquiries WHERE sales_pic = ? GROUP BY status ORDER BY count DESC'
  ).all(name) as Array<{ status: string; count: number }>;

  // Sourcing stats
  const itemsSourced = (db.prepare(
    `SELECT COUNT(*) as c FROM activity_log WHERE done_by_name = ? AND action = 'Sourcing info submitted'`
  ).get(name) as { c: number }).c;
  const inquiriesContributed = (db.prepare(
    `SELECT COUNT(DISTINCT inquiry_id) as c FROM activity_log WHERE done_by_name = ? AND action = 'Sourcing info submitted'`
  ).get(name) as { c: number }).c;
  const thisMonthSourcing = (db.prepare(
    `SELECT COUNT(*) as c FROM activity_log WHERE done_by_name = ? AND action = 'Sourcing info submitted' AND created_at >= ?`
  ).get(name, startOfMonthIso) as { c: number }).c;

  // Per-user item state breakdown — scoped to RFQs this sourcing user has contributed to
  // (matches inquiriesContributed, which is also based on activity_log)
  const contributedInquiriesClause = `ii.inquiry_id IN (
    SELECT DISTINCT inquiry_id FROM activity_log
    WHERE done_by_name = ? AND action = 'Sourcing info submitted'
  )`;
  const userItemsTerisi = (db.prepare(
    `SELECT COUNT(*) as c FROM inquiry_items ii
     JOIN inquiries i ON i.id = ii.inquiry_id
     WHERE ${contributedInquiriesClause} AND i.sourcing_missed = 0
       AND COALESCE(ii.supplier,'') != '' AND ii.harga_beli IS NOT NULL AND COALESCE(ii.lead_time,'') != ''`
  ).get(name) as { c: number }).c;
  const userItemsMissed = (db.prepare(
    `SELECT COUNT(*) as c FROM inquiry_items ii
     JOIN inquiries i ON i.id = ii.inquiry_id
     WHERE ${contributedInquiriesClause} AND i.sourcing_missed = 1`
  ).get(name) as { c: number }).c;
  const userItemsTidakTerisi = (db.prepare(
    `SELECT COUNT(*) as c FROM inquiry_items ii
     JOIN inquiries i ON i.id = ii.inquiry_id
     WHERE ${contributedInquiriesClause} AND i.sourcing_missed = 0
       AND (COALESCE(ii.supplier,'') = '' OR ii.harga_beli IS NULL OR COALESCE(ii.lead_time,'') = '')`
  ).get(name) as { c: number }).c;

  // Manager stats
  const approvalsTotal = (db.prepare(
    `SELECT COUNT(DISTINCT inquiry_id) as c FROM activity_log WHERE done_by_name = ? AND action = 'Price approved'`
  ).get(name) as { c: number }).c;
  const approvalsThisMonth = (db.prepare(
    `SELECT COUNT(DISTINCT inquiry_id) as c FROM activity_log WHERE done_by_name = ? AND action = 'Price approved' AND created_at >= ?`
  ).get(name, startOfMonthIso) as { c: number }).c;
  const inquiriesApproved = (db.prepare(
    `SELECT COUNT(DISTINCT inquiry_id) as c FROM activity_log WHERE done_by_name = ? AND action = 'Price approved'`
  ).get(name) as { c: number }).c;

  res.json({
    salesStats: { total, thisMonth: thisMonthSales, quotationSent, unsent, active, conversionRate, statusBreakdown },
    sourcingStats: { itemsSourced, inquiriesContributed, thisMonth: thisMonthSourcing, itemsTerisi: userItemsTerisi, itemsMissed: userItemsMissed, itemsTidakTerisi: userItemsTidakTerisi },
    managerStats: { approvalsTotal, approvalsThisMonth, inquiriesApproved },
  });
});

// GET /inquiries/dashboard
inquiriesRouter.get('/dashboard', (_req: Request, res: Response) => {
  autoMarkMissedRfqs();
  autoMarkUnsentRfqs();
  const startOfMonth = new Date();
  startOfMonth.setDate(1);
  startOfMonth.setHours(0, 0, 0, 0);
  const startOfMonthIso = startOfMonth.toISOString();

  // Sales stats
  const total = (db.prepare(`SELECT COUNT(*) as c FROM inquiries WHERE status != 'missed'`).get() as { c: number }).c;
  const thisMonth = (db.prepare(`SELECT COUNT(*) as c FROM inquiries WHERE created_at >= ? AND status != 'missed'`).get(startOfMonthIso) as { c: number }).c;
  const quotationSent = (db.prepare(`SELECT COUNT(*) as c FROM inquiries WHERE status IN ('quotation_sent','ready_to_purchase')`).get() as { c: number }).c;
  const unsent = (db.prepare(
    `SELECT COUNT(*) as c FROM inquiries WHERE status = 'unsent'`
  ).get() as { c: number }).c;
  const conversionRate = total > 0 ? +((quotationSent / total) * 100).toFixed(1) : 0;

  const topSales = db.prepare(
    `SELECT sales_pic, COUNT(*) as sent_count FROM inquiries WHERE status IN ('quotation_sent','ready_to_purchase') GROUP BY sales_pic ORDER BY sent_count DESC LIMIT 5`
  ).all() as Array<{ sales_pic: string; sent_count: number }>;

  const topMarketing = db.prepare(
    `SELECT sales_pic, COUNT(*) as sent_count FROM inquiries
     WHERE status IN ('quotation_sent', 'ready_to_purchase')
     GROUP BY sales_pic ORDER BY sent_count DESC LIMIT 5`
  ).all() as Array<{ sales_pic: string; sent_count: number }>;

  const statusBreakdown = db.prepare(
    `SELECT status, COUNT(*) as count FROM inquiries GROUP BY status ORDER BY count DESC`
  ).all() as Array<{ status: string; count: number }>;

  // RFQ-level sourcing breakdown
  const rfqsMissed = (db.prepare(`SELECT COUNT(*) as c FROM inquiries WHERE status = 'missed'`).get() as { c: number }).c;
  const rfqsMissedUnassigned = (db.prepare(`SELECT COUNT(*) as c FROM inquiries WHERE status = 'missed' AND (sourcing_pic IS NULL OR sourcing_pic = '')`).get() as { c: number }).c;
  const itemsMissedUnassigned = (db.prepare(`SELECT COUNT(*) as c FROM inquiry_items WHERE inquiry_id IN (SELECT id FROM inquiries WHERE status = 'missed' AND (sourcing_pic IS NULL OR sourcing_pic = ''))`).get() as { c: number }).c;

  // Item state breakdown — scoped to non-missed RFQs only
  const itemsTerisi = (db.prepare(
    `SELECT COUNT(*) as c FROM inquiry_items ii
     JOIN inquiries i ON i.id = ii.inquiry_id
     WHERE i.sourcing_missed = 0
       AND COALESCE(ii.supplier,'') != '' AND ii.harga_beli IS NOT NULL AND COALESCE(ii.lead_time,'') != ''`
  ).get() as { c: number }).c;
  const itemsMissed = (db.prepare(
    `SELECT COUNT(*) as c FROM inquiry_items WHERE inquiry_id IN (SELECT id FROM inquiries WHERE sourcing_missed = 1)`
  ).get() as { c: number }).c;
  const itemsTidakTerisi = (db.prepare(
    `SELECT COUNT(*) as c FROM inquiry_items ii
     JOIN inquiries i ON i.id = ii.inquiry_id
     WHERE i.sourcing_missed = 0
       AND (COALESCE(ii.supplier,'') = '' OR ii.harga_beli IS NULL OR COALESCE(ii.lead_time,'') = '')`
  ).get() as { c: number }).c;

  // Sourcing stats
  const sourcingPending = (db.prepare(`SELECT COUNT(*) as c FROM inquiries WHERE status = 'rfq' AND sourcing_missed = 0`).get() as { c: number }).c;
  const sourcingItemsTotal = itemsTerisi + itemsMissed;
  // "This month" approximated from activity_log (deduplication not possible without item timestamps)
  const sourcingItemsThisMonth = (db.prepare(
    `SELECT COUNT(DISTINCT inquiry_id) as c FROM activity_log WHERE action = 'Sourcing info submitted' AND created_at >= ?`
  ).get(startOfMonthIso) as { c: number }).c;
  const topSourcers = db.prepare(
    `SELECT done_by_name as sourcing_pic, COUNT(*) as items_count
     FROM activity_log WHERE action = 'Sourcing info submitted'
     GROUP BY done_by_name ORDER BY items_count DESC LIMIT 5`
  ).all() as Array<{ sourcing_pic: string; items_count: number }>;

  const urgentRfqs = db.prepare(
    `SELECT i.id, i.rfq_no, i.customer, i.sourcing_pic,
       COALESCE(NULLIF(i.deadline_quotation, ''), MIN(ii.item_need_by_date)) AS need_by_date,
       CAST(julianday(COALESCE(NULLIF(i.deadline_quotation, ''), MIN(ii.item_need_by_date))) - julianday(date('now', 'localtime')) AS INTEGER) AS days_left
     FROM inquiries i
     LEFT JOIN inquiry_items ii ON ii.inquiry_id = i.id
     WHERE i.status = 'rfq'
       AND i.sourcing_missed = 0
     GROUP BY i.id
     HAVING need_by_date IS NOT NULL AND need_by_date != ''
       AND days_left >= 0 AND days_left <= 1
     ORDER BY need_by_date ASC LIMIT 8`
  ).all() as Array<{ id: string; rfq_no: string; customer: string; sourcing_pic: string | null; need_by_date: string; days_left: number }>;

  res.json({
    total, thisMonth, quotationSent, unsent, conversionRate, topSales, topMarketing, statusBreakdown,
    sourcingPending, sourcingItemsThisMonth, sourcingItemsTotal, topSourcers, urgentRfqs, rfqsMissed, rfqsMissedUnassigned,
    itemsTerisi, itemsTidakTerisi, itemsMissed, itemsMissedUnassigned,
  });
});

// GET /inquiries/:id
inquiriesRouter.get('/:id([^/]{1,})', (req: Request, res: Response) => {
  const { id } = req.params;
  const row = db.prepare('SELECT * FROM inquiries WHERE id = ?').get(id) as Record<string, unknown> | undefined;
  if (!row) { res.status(404).json({ error: 'Not found.' }); return; }

  const items = db.prepare('SELECT * FROM inquiry_items WHERE inquiry_id = ? ORDER BY coupa_row_index ASC, id ASC').all(id) as Array<Record<string, unknown>>;
  const logs = (db.prepare('SELECT * FROM activity_log WHERE inquiry_id = ? ORDER BY created_at ASC').all(id) as Array<Record<string, unknown>>)
    .map((l) => ({
      id: l['id'], inquiryId: l['inquiry_id'], action: l['action'],
      oldStatus: l['old_status'], newStatus: l['new_status'], note: l['note'],
      doneBy: l['done_by'], doneByName: l['done_by_name'], createdAt: l['created_at'],
    }));

  res.json({ ...mapInquiry(row, items), activityLog: logs });
});

// POST /inquiries/import-coupa
inquiriesRouter.post('/import-coupa', (req: Request, res: Response) => {
  const { fileBase64, fileName, createdBy, createdByName, organization, needByDate } = req.body as Record<string, unknown>;
  const org = normalizeOrganization(organization);
  if (!fileBase64 || !fileName || !createdBy || !org) {
    res.status(400).json({ error: 'fileBase64, fileName, createdBy, organization are required. Organization must exist in Settings.' });
    return;
  }
  const overrideNeedByDate = typeof needByDate === 'string' && needByDate.trim() ? needByDate.trim() : null;
  if (!overrideNeedByDate) {
    res.status(400).json({ error: 'needByDate is required.' });
    return;
  }

  let workbook: XLSX.WorkBook;
  try {
    const buffer = Buffer.from(String(fileBase64), 'base64');
    workbook = XLSX.read(buffer, { type: 'buffer', cellStyles: true });
  } catch {
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

  const items: Array<Record<string, unknown>> = [];
  for (let r = 5; r <= range.e.r; r += 1) {
    const itemId = readSheetCell(sheet, r, itemIdCol);
    const bidId = readSheetCell(sheet, r, bidIdCol);
    if (itemId == null && bidId == null) {
      continue;
    }

    const toNumber = (value: unknown) => {
      if (value == null || value === '') return null;
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
      item_need_by_date: overrideNeedByDate,
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

  const id = generateId();
  const rfqNo = generateRfqNo();
  const tanggal = new Date().toISOString().split('T')[0];
  const createdAt = new Date().toISOString();
  const customer = deriveCustomerFromFilename(String(fileName));
  const salesPic = String(createdByName ?? createdBy);
  const firstItemName = String(items[0]?.item_name ?? 'Multiple items');
  const namaBarang = items.length > 1 ? `${firstItemName} +${items.length - 1} items` : firstItemName;

  const tx = db.transaction(() => {
    db.prepare(
      `INSERT INTO inquiries (id, rfq_no, tanggal, customer, sales_pic, nama_barang, deadline_quotation, status, coupa_source, organization, coupa_file_name, created_at, created_by)
       VALUES (?, ?, ?, ?, ?, ?, ?, 'new_inquiry', 1, ?, ?, ?, ?)`
    ).run(id, rfqNo, tanggal, customer, salesPic, namaBarang, overrideNeedByDate, org, String(fileName), createdAt, createdBy);

    const insertItem = db.prepare(
      `INSERT INTO inquiry_items (
        id, inquiry_id, coupa_row_index, lot_id, lot_name, lot_expected_quantity, lot_quantity_note,
        coupa_item_id, item_name, item_quantity, item_uom, item_need_by_date, item_manufacturer_name,
        item_manufacturer_part_number, item_classification_of_goods, item_extended_description, item_fiscal_code,
        coupa_bid_id, bid_capacity, bid_price_amount, bid_price_currency, bid_lead_time,
        bid_supplier_item_name, bid_item_part_number, bid_item_description, bid_shipping_term,
        harga_jual, alternate_name
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
    );

    for (const item of items) {
      insertItem.run(
        generateId(),
        id,
        item.coupa_row_index ?? null,
        item.lot_id ?? null,
        item.lot_name ?? null,
        item.lot_expected_quantity ?? null,
        item.lot_quantity_note ?? null,
        item.coupa_item_id ?? null,
        item.item_name ?? null,
        item.item_quantity ?? null,
        item.item_uom ?? null,
        item.item_need_by_date ?? null,
        item.item_manufacturer_name ?? null,
        item.item_manufacturer_part_number ?? null,
        item.item_classification_of_goods ?? null,
        item.item_extended_description ?? null,
        item.item_fiscal_code ?? null,
        item.coupa_bid_id ?? null,
        item.bid_capacity ?? null,
        item.bid_price_amount ?? null,
        item.bid_price_currency ?? null,
        item.bid_lead_time ?? null,
        item.bid_supplier_item_name ?? null,
        item.bid_item_part_number ?? null,
        item.bid_item_description ?? null,
        item.bid_shipping_term ?? null,
        item.bid_price_amount ?? null,
        item.bid_supplier_item_name ?? null
      );
    }

    db.prepare(
      `INSERT INTO coupa_files (inquiry_id, file_name, file_data, created_at) VALUES (?, ?, ?, ?)`
    ).run(id, String(fileName), Buffer.from(String(fileBase64), 'base64'), createdAt);
  });

  tx();

  logActivity(id, 'Coupa file imported', null, 'new_inquiry', String(fileName), String(createdBy), String(createdByName ?? createdBy));
  res.status(201).json({ id, rfqNo, itemCount: items.length });
});

function generateExcelFilename(customer: string): string {
  // customer may be "Pt Merdeka Copper Gold Event #58847" → "en#Pt Merdeka Copper Gold Event#58847"
  const match = customer.match(/\s*#(\d+)\s*$/);
  const eventNo = match ? match[1] : '';
  const customerName = match ? customer.replace(/\s*#\d+\s*$/, '').trim() : customer;
  return `en#${customerName}${eventNo ? '#' + eventNo : ''}`;
}

function buildCoupaFormatExcel(items: Array<Record<string, unknown>>): Buffer {
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

  const ws: XLSX.WorkSheet = {};

  const setCell = (r: number, c: number, v: string | number | null) => {
    if (v == null) return;
    ws[XLSX.utils.encode_cell({ r, c })] = { v, t: typeof v === 'number' ? 'n' : 's' } as XLSX.CellObject;
  };

  // Row 0: JSON field_name metadata — matches real Coupa format exactly (with offset:5)
  setCell(0, 0,  JSON.stringify({ start: true, layout: 'table', name: 'supplier/response_lines', locale: 'en' }));
  const fieldCols: Array<[number, string]> = [
    [1,  'lot.id'],             [2,  'lot.name'],
    [3,  'lot.expected_quantity'], [4, 'lot.quantity_note'],
    [5,  'item.id'],            [6,  'item.name'],
    [7,  'item.quantity'],      [8,  'item.uom'],
    [9,  'item.need_by_date'],  [10, 'item.manufacturer_name'],
    [11, 'item.manufacturer_part_number'],
    [12, 'item.classification_of_goods'],
    [13, 'item.extended_description'],
    [14, 'item.fiscal_code'],   [15, 'bid.id'],
    [16, 'bid.capacity'],       [17, 'bid.price_amount'],
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
  setCell(3, 1,  'Lot');
  setCell(3, 3,  'Lot Fields');
  setCell(3, 5,  'Item / Service');
  setCell(3, 7,  'Item / Service Fields');
  setCell(3, 15, 'Supplier Response Fields');

  // Row 4: Column labels — exact labels from real Coupa file
  const colLabels: Array<[number, string]> = [
    [1,  'Lot ID (Text)'],
    [2,  'Lot Name (Text)'],
    [3,  'Expected Quantity (Integer)'],
    [4,  'Quantity Note (Text)'],
    [5,  'Item ID (Text)'],
    [6,  'Item Description (Text)'],
    [7,  'Expected Quantity (Number)'],
    [8,  'Unit of Measurement (Text)'],
    [9,  'Need by Date (Date)'],
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
    const itemId   = (item['coupa_item_id'] ?? item['id']) as string | number | null;
    const bidId    = item['coupa_bid_id'] as string | number | null;
    const itemName = item['item_name'] as string | null;
    const itemQty  = item['item_quantity'] as number | null;
    const approvedPrice = (item['approved_price'] ?? item['harga_jual'] ?? item['bid_price_amount']) as number | null;
    const bidCapacity = (item['bid_capacity'] ?? itemQty) as number | null;
    const bidCurrency = (item['bid_price_currency'] ?? 'IDR') as string | null;
    const leadTime  = normalizeLeadTime((item['lead_time_customer'] ?? item['lead_time'] ?? item['bid_lead_time']) as string | number | null);
    const catatan   = item['catatan_quotation'] as string | null;
    const description = catatan || (item['item_extended_description'] as string | null) || null;
    const shipping    = (item['term_pembayaran'] as string | null) || (item['bid_shipping_term'] as string | null);
    const supplierItemName = (item['bid_supplier_item_name'] ?? item['alternate_name'] ?? null) as string | null;
    const bidItemPartNumber = item['bid_item_part_number'] as string | null;
    const bidItemDescription = (item['bid_item_description'] ?? description) as string | null;

    // Lot columns
    setCell(r, 3, itemQty);  // lot.expected_quantity
    // Item columns
    setCell(r, 5,  itemId != null ? String(itemId) : null);
    setCell(r, 6,  itemName);
    setCell(r, 7,  itemQty);
    setCell(r, 8,  item['item_uom'] as string | null);
    setCell(r, 9,  parseExcelDate(item['item_need_by_date']));
    setCell(r, 10, item['item_manufacturer_name'] as string | null);
    setCell(r, 11, item['item_manufacturer_part_number'] as string | null);
    setCell(r, 12, item['item_classification_of_goods'] as string | null);
    setCell(r, 13, item['item_extended_description'] as string | null);
    setCell(r, 14, item['item_fiscal_code'] as string | null);
    // Bid / supplier response columns
    setCell(r, 15, bidId != null ? String(bidId) : null);
    setCell(r, 16, bidCapacity);      // bid.capacity
    setCell(r, 17, approvedPrice);    // bid.price_amount
    setCell(r, 18, bidCurrency);      // bid.price_currency
    setCell(r, 19, leadTime);         // bid.lead_time
    setCell(r, 20, supplierItemName); // bid.supplier_item_name
    setCell(r, 21, bidItemPartNumber); // bid.item_part_number
    setCell(r, 22, bidItemDescription); // bid.item_description
    setCell(r, 23, shipping);    // bid.shipping_term
  });

  const lastRow = 5 + Math.max(items.length - 1, 0);
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: lastRow, c: NUM_COLS - 1 } });

  // Hide row 0 (JSON metadata) — invisible to users, preserved for re-import
  ws['!rows'] = [{ hidden: true }];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Items and Services');
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }) as Buffer;
}

// GET /inquiries/:id/export-coupa
inquiriesRouter.get('/:id/export-coupa', (req: Request, res: Response) => {
  const { id } = req.params;
  const inquiryRow = db.prepare('SELECT id, status, customer FROM inquiries WHERE id = ?').get(id) as { id: string; status: string; customer: string } | undefined;
  if (!inquiryRow) { res.status(404).json({ error: 'Not found.' }); return; }

  const items = db.prepare(
    `SELECT * FROM inquiry_items WHERE inquiry_id = ? ORDER BY coupa_row_index ASC, id ASC`
  ).all(id) as Array<Record<string, unknown>>;

  const fileRow = db.prepare('SELECT file_name, file_data FROM coupa_files WHERE inquiry_id = ?').get(id) as { file_name: string; file_data: Buffer } | undefined;

  let output: Buffer;
  let safeName: string;

  if (fileRow) {
    // Fill original Coupa Excel with sourcing/approval data
    const workbook = XLSX.read(fileRow.file_data, { type: 'buffer', cellStyles: true });
    const sheet = workbook.Sheets['Items and Services'];
    if (!sheet) { res.status(400).json({ error: 'Items and Services sheet not found.' }); return; }

    const fieldMap = parseCoupaFieldMap(sheet);

    for (const item of items) {
      const rowIndex = item['coupa_row_index'];
      if (rowIndex == null) continue;
      const row = Number(rowIndex) - 1;

      const approvedPrice = ((item['approved_price'] ?? item['harga_jual'] ?? item['bid_price_amount']) as number | null);
      const leadTimeCustomer = item['lead_time_customer'] as string | number | null;
      const leadTimeFallback = item['lead_time'] as string | number | null;
      const leadTime = normalizeLeadTime(leadTimeCustomer ?? leadTimeFallback ?? (item['bid_lead_time'] as string | number | null));
      const catatan = (item['catatan_quotation'] as string | null) ?? null;
      const description = catatan || (item['item_extended_description'] as string | null) || null;
      const itemQty = item['item_quantity'] as number | null;
      const bidCapacity = (item['bid_capacity'] ?? itemQty) as number | null;
      const bidCurrency = (item['bid_price_currency'] ?? null) as string | null;
      const supplierItemName = (item['bid_supplier_item_name'] ?? item['alternate_name'] ?? null) as string | null;
      const bidItemPartNumber = item['bid_item_part_number'] as string | null;
      const bidItemDescription = (item['bid_item_description'] ?? description) as string | null;

      setSheetCell(sheet, row, fieldMap['bid.capacity'], bidCapacity ?? null);
      setSheetCell(sheet, row, fieldMap['bid.price_amount'], approvedPrice ?? null);
      setSheetCell(sheet, row, fieldMap['bid.price_currency'], bidCurrency ?? null);
      setSheetCell(sheet, row, fieldMap['bid.lead_time'], leadTime ?? null);
      setSheetCell(sheet, row, fieldMap['bid.supplier_item_name'], supplierItemName ?? null);
      setSheetCell(sheet, row, fieldMap['bid.item_part_number'], bidItemPartNumber ?? null);
      setSheetCell(sheet, row, fieldMap['bid.item_description'], bidItemDescription ?? null);

      const shipping = (item['term_pembayaran'] as string | null) || (item['bid_shipping_term'] as string | null);
      if (shipping) setSheetCell(sheet, row, fieldMap['bid.shipping_term'], shipping);
    }

    delete sheet['!autofilter'];
    output = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx', cellStyles: true }) as Buffer;
    safeName = fileRow.file_name.endsWith('.xlsx') ? fileRow.file_name : fileRow.file_name + '.xlsx';
  } else {
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
inquiriesRouter.post('/', (req: Request, res: Response) => {
  const { customer, salesPic, namaBarang, spesifikasi, qty, itemUom, itemNeedByDate, itemManufacturerName, itemManufacturerPartNumber, itemClassificationOfGoods, itemImage, deadlineQuotation, lampiran, createdBy, createdByName, organization } =
    req.body as Record<string, unknown>;
  const org = normalizeOrganization(organization);

  if (!customer || !salesPic || !namaBarang || !createdBy || !org) {
    res.status(400).json({ error: 'customer, salesPic, namaBarang, createdBy, organization are required. Organization must exist in Settings.' });
    return;
  }

  const id = generateId();
  const rfqNo = generateRfqNo();
  const tanggal = new Date().toISOString().split('T')[0];
  const createdAt = new Date().toISOString();
  const needByDate = itemNeedByDate ?? deadlineQuotation ?? null;

  db.prepare(
    `INSERT INTO inquiries (id, rfq_no, tanggal, customer, sales_pic, nama_barang, spesifikasi, qty, deadline_quotation, lampiran, organization, status, created_at, created_by)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'new_inquiry', ?, ?)`
  ).run(id, rfqNo, tanggal, customer, salesPic, namaBarang, spesifikasi ?? null, qty ?? null, needByDate, lampiran ?? null, org, createdAt, createdBy);

  db.prepare(
    `INSERT INTO inquiry_items (
      id, inquiry_id, item_name, item_quantity, item_uom, item_need_by_date,
      item_manufacturer_name, item_manufacturer_part_number, item_classification_of_goods,
      item_extended_description, item_image
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
  ).run(generateId(), id, namaBarang, qty ?? null, itemUom ?? null, needByDate,
    itemManufacturerName ?? null, itemManufacturerPartNumber ?? null, itemClassificationOfGoods ?? null,
    spesifikasi ?? null, itemImage ?? null);

  logActivity(id, 'Inquiry created', null, 'new_inquiry', null, String(createdBy), String(createdByName ?? createdBy));

  res.status(201).json({ id, rfqNo, tanggal, status: 'new_inquiry', createdAt });
});

// PUT /inquiries/:id
inquiriesRouter.put('/:id', (req: Request, res: Response) => {
  const { id } = req.params;
  const inquiry = db.prepare('SELECT id, status FROM inquiries WHERE id = ?').get(id) as { id: string; status: string } | undefined;

  if (!inquiry) { res.status(404).json({ error: 'Not found.' }); return; }
  if (!['new_inquiry', 'rfq'].includes(inquiry.status)) {
    res.status(400).json({ error: 'Cannot edit inquiry at this stage.' }); return;
  }

  const body = req.body as Record<string, unknown>;
  const { customer, salesPic, namaBarang, spesifikasi, qty, itemUom, itemNeedByDate, itemManufacturerName, itemManufacturerPartNumber, itemClassificationOfGoods, itemImage, deadlineQuotation, lampiran, updatedBy, updatedByName } =
    body;
  const org = req.body ? normalizeOrganization(body['organization']) : null;
  if (body['organization'] != null && !org) {
    res.status(400).json({ error: 'organization must exist in Settings.' }); return;
  }

  const needByDate = itemNeedByDate ?? deadlineQuotation ?? null;
  const rawItems = Array.isArray(body['items']) ? body['items'] as Array<Record<string, unknown>> : null;
  if (rawItems && rawItems.length === 0) {
    res.status(400).json({ error: 'At least one item is required.' }); return;
  }

  const updateInquiry = db.prepare(
    `UPDATE inquiries SET
       customer = COALESCE(?, customer), sales_pic = COALESCE(?, sales_pic),
       organization = COALESCE(?, organization),
       nama_barang = COALESCE(?, nama_barang), spesifikasi = ?, qty = ?,
       deadline_quotation = ?, lampiran = ?,
       updated_at = ?, updated_by = ?
     WHERE id = ?`
  );
  const updateSingleItem = db.prepare(
    `UPDATE inquiry_items SET
       item_name = COALESCE(?, item_name),
       item_extended_description = ?,
       item_quantity = ?,
       item_uom = ?,
       item_need_by_date = ?,
       item_manufacturer_name = ?,
       item_manufacturer_part_number = ?,
       item_classification_of_goods = ?,
       item_image = ?
     WHERE inquiry_id = ?`
  );
  const updateItem = db.prepare(
    `UPDATE inquiry_items SET
       item_name = ?,
       item_extended_description = ?,
       item_quantity = ?,
       item_uom = ?,
       item_need_by_date = ?,
       item_image = ?
     WHERE id = ? AND inquiry_id = ?`
  );
  const insertItem = db.prepare(
    `INSERT INTO inquiry_items (id, inquiry_id, item_name, item_quantity, item_uom, item_need_by_date,
      item_extended_description, item_image)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
  );
  const deleteMissingItems = db.prepare(
    `DELETE FROM inquiry_items
     WHERE inquiry_id = ? AND id NOT IN (${rawItems?.map(() => '?').join(',') || "''"})`
  );
  const runUpdate = db.transaction(() => {
    updateInquiry.run(customer ?? null, salesPic ?? null, org ?? null, namaBarang ?? null, spesifikasi ?? null, qty ?? null, needByDate, lampiran ?? null, new Date().toISOString(), updatedBy ?? null, id);

    if (rawItems) {
      const keptIds: string[] = [];
      for (const item of rawItems) {
        const itemId = typeof item['id'] === 'string' && item['id'] ? String(item['id']) : generateId();
        const existing = db.prepare('SELECT id FROM inquiry_items WHERE id = ? AND inquiry_id = ?').get(itemId, id) as { id: string } | undefined;
        const itemName = String(item['itemName'] ?? '').trim();
        const itemQuantity = item['itemQuantity'] ?? null;
        const itemUomValue = String(item['itemUom'] ?? '').trim();
        const itemDescription = item['itemExtendedDescription'] == null ? null : String(item['itemExtendedDescription']).trim();
        const itemImageValue = item['itemImage'] ?? null;
        keptIds.push(itemId);
        if (existing) {
          updateItem.run(itemName, itemDescription, itemQuantity, itemUomValue, needByDate, itemImageValue, itemId, id);
        } else {
          insertItem.run(itemId, id, itemName, itemQuantity, itemUomValue, needByDate, itemDescription, itemImageValue);
        }
      }
      deleteMissingItems.run(id, ...keptIds);
    } else {
      const itemCount = (db.prepare('SELECT COUNT(*) as c FROM inquiry_items WHERE inquiry_id = ?').get(id) as { c: number }).c;
      if (itemCount === 1) {
        updateSingleItem.run(namaBarang ?? null, spesifikasi ?? null, qty ?? null, itemUom ?? null, needByDate,
          itemManufacturerName ?? null, itemManufacturerPartNumber ?? null, itemClassificationOfGoods ?? null, itemImage ?? null, id);
      }
    }
  });
  runUpdate();

  logActivity(id, 'Inquiry updated', inquiry.status, inquiry.status, null, String(updatedBy ?? ''), String(updatedByName ?? updatedBy ?? ''));
  res.json({ ok: true });
});

// POST /inquiries/:id/send-rfq — Sales qualifies and sends to Sourcing
inquiriesRouter.post('/:id/send-rfq', (req: Request, res: Response) => {
  const { id } = req.params;
  const inquiry = db.prepare('SELECT id, status FROM inquiries WHERE id = ?').get(id) as { id: string; status: string } | undefined;

  if (!inquiry) { res.status(404).json({ error: 'Not found.' }); return; }
  if (inquiry.status !== 'new_inquiry') {
    res.status(400).json({ error: 'Only new inquiries can be sent to sourcing.' }); return;
  }

  const { doneBy, doneByName, note } = req.body as Record<string, unknown>;
  db.prepare('UPDATE inquiries SET status = ?, updated_at = ?, updated_by = ? WHERE id = ?')
    .run('rfq', new Date().toISOString(), doneBy, id);

  logActivity(id, 'RFQ sent to Sourcing', 'new_inquiry', 'rfq', String(note ?? ''), String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''));
  res.json({ ok: true });
});

// POST /inquiries/:id/items — Marketing adds a new item
inquiriesRouter.post('/:id/items', (req: Request, res: Response) => {
  const { id } = req.params;
  const inquiry = db.prepare('SELECT id, status FROM inquiries WHERE id = ?').get(id) as { id: string; status: string } | undefined;
  if (!inquiry) { res.status(404).json({ error: 'Not found.' }); return; }
  if (inquiry.status !== 'new_inquiry') { res.status(400).json({ error: 'Inquiry must be in new_inquiry status.' }); return; }

  const { itemName, itemQuantity, itemUom, itemNeedByDate, itemManufacturerName, itemManufacturerPartNumber, itemClassificationOfGoods, itemExtendedDescription, itemImage, doneBy, doneByName } =
    req.body as Record<string, unknown>;
  if (!itemName) { res.status(400).json({ error: 'itemName is required.' }); return; }

  const newId = generateId();
  db.prepare(
    `INSERT INTO inquiry_items (id, inquiry_id, item_name, item_quantity, item_uom, item_need_by_date,
      item_manufacturer_name, item_manufacturer_part_number, item_classification_of_goods,
      item_extended_description, item_image)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
  ).run(newId, id, itemName, itemQuantity ?? null, itemUom ?? null, itemNeedByDate ?? null,
    itemManufacturerName ?? null, itemManufacturerPartNumber ?? null, itemClassificationOfGoods ?? null,
    itemExtendedDescription ?? null, itemImage ?? null);

  logActivity(id, 'Item added by marketing', null, null, null, String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''));
  res.json({ ok: true, id: newId });
});

// PATCH /inquiries/:id/items/:itemId — Marketing updates target price / image / need-by-date (new_inquiry only)
inquiriesRouter.patch('/:id/items/:itemId', (req: Request, res: Response) => {
  const { id, itemId } = req.params;
  const inquiry = db.prepare('SELECT id, status FROM inquiries WHERE id = ?').get(id) as { id: string; status: string } | undefined;
  if (!inquiry) { res.status(404).json({ error: 'Not found.' }); return; }
  if (inquiry.status !== 'new_inquiry') { res.status(400).json({ error: 'Inquiry must be in new_inquiry status.' }); return; }

  const item = db.prepare('SELECT id FROM inquiry_items WHERE id = ? AND inquiry_id = ?').get(itemId, id) as { id: string } | undefined;
  if (!item) { res.status(404).json({ error: 'Item not found.' }); return; }

  const { itemImage, doneBy, doneByName } = req.body as Record<string, unknown>;
  db.prepare('UPDATE inquiry_items SET item_image = ? WHERE id = ?')
    .run(itemImage ?? null, itemId);

  logActivity(id, 'Item reviewed by marketing', null, null, null, String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''));
  res.json({ ok: true });
});

// PATCH /inquiries/:id/need-by-date — Sales updates RFQ-level need-by date from any status
inquiriesRouter.patch('/:id/need-by-date', (req: Request, res: Response) => {
  const { id } = req.params;
  const inquiry = db.prepare('SELECT id FROM inquiries WHERE id = ?').get(id) as { id: string } | undefined;
  if (!inquiry) { res.status(404).json({ error: 'Not found.' }); return; }

  const { needByDate, doneBy, doneByName } = req.body as Record<string, unknown>;
  db.prepare('UPDATE inquiries SET deadline_quotation = ? WHERE id = ?')
    .run(needByDate ?? null, id);

  logActivity(id, 'Need-by date updated', null, null, null, String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''));
  res.json({ ok: true });
});

// PATCH /inquiries/:id/items/:itemId/harga-jual — Sales/Marketing updates selling price
inquiriesRouter.patch('/:id/items/:itemId/harga-jual', (req: Request, res: Response) => {
  const { id, itemId } = req.params;
  const inquiry = db.prepare('SELECT id, status FROM inquiries WHERE id = ?').get(id) as { id: string; status: string } | undefined;
  if (!inquiry) { res.status(404).json({ error: 'Not found.' }); return; }
  if (!['price_approved', 'quotation_sent', 'deal'].includes(inquiry.status)) {
    res.status(400).json({ error: 'Inquiry must be in price_approved, quotation_sent, or deal status.' }); return;
  }
  const item = db.prepare(
    'SELECT id, harga_beli, approved_price, review_status FROM inquiry_items WHERE id = ? AND inquiry_id = ?'
  ).get(itemId, id) as { id: string; harga_beli: number | null; approved_price: number | null; review_status: string | null } | undefined;
  if (!item) { res.status(404).json({ error: 'Item not found.' }); return; }

  const { hargaJual, doneBy, doneByName } = req.body as Record<string, unknown>;
  if (!hargaJual) { res.status(400).json({ error: 'hargaJual is required.' }); return; }

  const nextPrice = Number(hargaJual);
  const margin = item.harga_beli != null ? nextPrice - item.harga_beli : null;
  const needsReview = item.approved_price != null ? (nextPrice < item.approved_price ? 1 : 0) : 0;
  const nextReviewStatus =
    item.review_status === 'rejected'
      ? 'rejected'
      : (needsReview ? 'review' : 'approved');
  db.prepare('UPDATE inquiry_items SET harga_jual = ?, margin = ?, needs_price_review = ?, review_status = ? WHERE id = ?')
    .run(nextPrice, margin, needsReview, nextReviewStatus, itemId);

  logActivity(id, 'Harga jual updated by marketing', null, null, null, String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''));
  res.json({ ok: true });
});

// POST /inquiries/:id/sourcing-info — Sourcing fills supplier data
inquiriesRouter.post('/:id/sourcing-info', (req: Request, res: Response) => {
  const { id } = req.params;
  const inquiry = db.prepare('SELECT id, status FROM inquiries WHERE id = ?').get(id) as { id: string; status: string } | undefined;

  if (!inquiry) { res.status(404).json({ error: 'Not found.' }); return; }
  if (!['rfq', 'price_approval'].includes(inquiry.status)) {
    res.status(400).json({ error: 'Inquiry must be in RFQ status.' }); return;
  }

  const { supplier, supplierUrl, hargaBeli, leadTime, moq, stockAvailability, termPembayaran, ppnType, doneBy, doneByName } =
    req.body as Record<string, unknown>;

  const supplierStr = String(supplier ?? '').trim();
  if (!supplierStr || hargaBeli === undefined || !leadTime) {
    res.status(400).json({ error: 'supplier, hargaBeli, leadTime are required.' }); return;
  }
  if (!/[A-Za-z0-9]/.test(supplierStr)) {
    res.status(400).json({ error: 'Supplier name is invalid.' }); return;
  }
  if (!ppnType) {
    res.status(400).json({ error: 'ppnType is required.' }); return;
  }

  const item = db.prepare('SELECT id, price_approved FROM inquiry_items WHERE inquiry_id = ? ORDER BY id LIMIT 1').get(id) as { id: string; price_approved: number } | undefined;
  if (!item) { res.status(400).json({ error: 'No items found.' }); return; }
  if (item.price_approved) { res.status(400).json({ error: 'Item already approved, cannot edit.' }); return; }

  db.prepare(
    `UPDATE inquiry_items SET supplier = ?, supplier_url = ?, harga_beli = ?, lead_time = ?, moq = ?,
       stock_availability = ?, term_pembayaran = ?, ppn_type = ? WHERE id = ?`
  ).run(supplierStr, supplierUrl ?? null, hargaBeli, leadTime, moq ?? null, stockAvailability ?? null, termPembayaran ?? null, ppnType ?? null, item.id);

  logActivity(id, 'Sourcing info submitted', 'rfq', 'rfq', null, String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''));
  recalcInquiryStatus(id, String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''));
  res.json({ ok: true });
});

inquiriesRouter.post('/:id/items/:itemId/sourcing-info', (req: Request, res: Response) => {
  const { id, itemId } = req.params;
  const inquiry = db.prepare('SELECT id, status, sourcing_missed FROM inquiries WHERE id = ?').get(id) as { id: string; status: string; sourcing_missed: number } | undefined;
  if (!inquiry) { res.status(404).json({ error: 'Not found.' }); return; }
  if (!['rfq', 'price_approval'].includes(inquiry.status)) {
    res.status(400).json({ error: 'Inquiry must be in RFQ status.' }); return;
  }
  if (inquiry.sourcing_missed) { res.status(400).json({ error: 'RFQ is marked as missed and can no longer be filled.' }); return; }

  const item = db.prepare('SELECT id, price_approved FROM inquiry_items WHERE id = ? AND inquiry_id = ?').get(itemId, id) as { id: string; price_approved: number } | undefined;
  if (!item) { res.status(404).json({ error: 'Item not found.' }); return; }
  if (item.price_approved) { res.status(400).json({ error: 'Item already approved, cannot edit.' }); return; }

  const { supplier, supplierUrl, hargaBeli, leadTime, moq, stockAvailability, termPembayaran, alternateName, ppnType, doneBy, doneByName } =
    req.body as Record<string, unknown>;

  const supplierStr = String(supplier ?? '').trim();
  if (!supplierStr || hargaBeli === undefined || !leadTime) {
    res.status(400).json({ error: 'supplier, hargaBeli, leadTime are required.' }); return;
  }
  if (!/[A-Za-z0-9]/.test(supplierStr)) {
    res.status(400).json({ error: 'Supplier name is invalid.' }); return;
  }
  if (!ppnType) {
    res.status(400).json({ error: 'ppnType is required.' }); return;
  }

  db.prepare(
    `UPDATE inquiry_items SET supplier = ?, supplier_url = ?, harga_beli = ?, lead_time = ?, moq = ?,
       stock_availability = ?, term_pembayaran = ?, alternate_name = ?, ppn_type = ? WHERE id = ?`
  ).run(supplierStr, supplierUrl ?? null, hargaBeli, leadTime, moq ?? null, stockAvailability ?? null, termPembayaran ?? null, alternateName ?? null, ppnType ?? null, itemId);

  logActivity(id, 'Sourcing info submitted', 'rfq', 'rfq', null, String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''));
  recalcInquiryStatus(id, String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''));
  res.json({ ok: true });
});

// POST /inquiries/:id/send-to-price-approval — Sourcing manually submits for price approval
inquiriesRouter.post('/:id/send-to-price-approval', (req: Request, res: Response) => {
  const { id } = req.params;
  const inquiry = db.prepare('SELECT id, status FROM inquiries WHERE id = ?').get(id) as { id: string; status: string } | undefined;
  if (!inquiry) { res.status(404).json({ error: 'Not found.' }); return; }
  if (inquiry.status !== 'rfq') {
    res.status(400).json({ error: 'Inquiry must be in rfq status.' }); return;
  }

  const submittedCount = (db.prepare(`
    SELECT COUNT(*) as c
    FROM inquiry_items
    WHERE inquiry_id = ?
      AND COALESCE(supplier, '') != ''
      AND harga_beli IS NOT NULL
      AND COALESCE(lead_time, '') != ''
      AND COALESCE(sourcing_missed, 0) = 0
  `).get(id) as { c: number }).c;
  if (submittedCount === 0) {
    res.status(400).json({ error: 'At least one item must be sourced before sending to Price Approval.' }); return;
  }

  const { doneBy, doneByName, note } = req.body as Record<string, unknown>;
  if (!doneBy) { res.status(400).json({ error: 'doneBy is required.' }); return; }

  const rfqRow = db.prepare('SELECT rfq_no FROM inquiries WHERE id = ?').get(id) as { rfq_no: string | null } | undefined;
  const now = new Date().toISOString();
  db.prepare('UPDATE inquiries SET status = ?, updated_at = ?, updated_by = ?, price_approval_started_at = ? WHERE id = ?')
    .run('price_approval', now, String(doneBy), now, id);
  logActivity(id, 'Sent to Price Approval', 'rfq', 'price_approval', note ? String(note) : null, String(doneBy), String(doneByName ?? doneBy));
  insertAndBroadcast(
    'price_approval', id, rfqRow?.rfq_no ?? null,
    `${rfqRow?.rfq_no ?? 'RFQ'} needs price approval — submitted by ${String(doneByName ?? doneBy)}`,
    String(doneBy), String(doneByName ?? doneBy),
  );
  res.json({ ok: true });
});

// POST /inquiries/:id/return-to-sourcing — Pricelist sends unfilled items back to sourcing
inquiriesRouter.post('/:id/return-to-sourcing', (req: Request, res: Response) => {
  const { id } = req.params;
  const inquiry = db.prepare('SELECT id, status, rfq_no, sourcing_pic FROM inquiries WHERE id = ?').get(id) as { id: string; status: string; rfq_no: string | null; sourcing_pic: string | null } | undefined;
  if (!inquiry) { res.status(404).json({ error: 'Not found.' }); return; }
  if (inquiry.status !== 'price_approval' && inquiry.status !== 'follow_up') {
    res.status(400).json({ error: 'Inquiry must be in price_approval or follow_up status.' }); return;
  }

  const { doneBy, doneByName } = req.body as Record<string, unknown>;
  if (!doneBy) { res.status(400).json({ error: 'doneBy is required.' }); return; }

  db.prepare('UPDATE inquiries SET status = ?, updated_at = ?, updated_by = ?, price_approval_started_at = NULL WHERE id = ?')
    .run('rfq', new Date().toISOString(), String(doneBy), id);
  logActivity(id, 'Returned to Sourcing', inquiry.status, 'rfq', null, String(doneBy), String(doneByName ?? doneBy));

  const sourcingRecipient = usernameForPic(inquiry.sourcing_pic);
  if (sourcingRecipient) {
    insertAndBroadcast(
      'return_to_sourcing', id, inquiry.rfq_no ?? null,
      `${inquiry.rfq_no ?? 'RFQ'} returned to Sourcing by ${String(doneByName ?? doneBy)}`,
      String(doneBy), String(doneByName ?? doneBy),
      sourcingRecipient,
    );
  }
  res.json({ ok: true });
});

// POST /inquiries/:id/send-to-sent — Marketing sends quotation to customer (price_approved → quotation_sent)
inquiriesRouter.post('/:id/send-to-sent', (req: Request, res: Response) => {
  const { id } = req.params;
  const inquiry = db.prepare('SELECT id, status FROM inquiries WHERE id = ?').get(id) as { id: string; status: string } | undefined;
  if (!inquiry) { res.status(404).json({ error: 'Not found.' }); return; }
  if (inquiry.status !== 'price_approved') {
    res.status(400).json({ error: 'Inquiry must be in price_approved status.' }); return;
  }

  const { doneBy, doneByName, incompleteReason } = req.body as Record<string, unknown>;
  if (!doneBy) { res.status(400).json({ error: 'doneBy is required.' }); return; }

  const items = db.prepare('SELECT item_name, price_approved, review_status, needs_price_review, harga_jual, approved_price FROM inquiry_items WHERE inquiry_id = ?')
    .all(id) as Array<{
      item_name: string | null;
      price_approved: number;
      review_status: string | null;
      needs_price_review: number;
      harga_jual: number | null;
      approved_price: number | null;
    }>;
  const unresolved = items.filter((i) =>
    i.price_approved !== 1 ||
    i.review_status === 'rejected' ||
    i.needs_price_review === 1 ||
    (i.harga_jual != null && i.approved_price != null && i.harga_jual < i.approved_price)
  );
  const isIncomplete = unresolved.length > 0;
  db.prepare('UPDATE inquiries SET status = ?, sent_incomplete = ?, sent_incomplete_reason = ?, updated_at = ?, updated_by = ? WHERE id = ?')
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

// POST /inquiries/:id/return-to-price-approval — Marketing sends back for price review (price_approved or quotation_sent → follow_up)
inquiriesRouter.post('/:id/return-to-price-approval', (req: Request, res: Response) => {
  const { id } = req.params;
  const inquiry = db.prepare('SELECT id, status FROM inquiries WHERE id = ?').get(id) as { id: string; status: string } | undefined;
  if (!inquiry) { res.status(404).json({ error: 'Not found.' }); return; }

  const { doneBy, doneByName, negotiationReason, reviewReason } = req.body as Record<string, unknown>;
  if (!doneBy) { res.status(400).json({ error: 'doneBy is required.' }); return; }
  const reason = String(reviewReason ?? negotiationReason ?? '').trim();
  if (!reason) { res.status(400).json({ error: 'reviewReason is required.' }); return; }

  const items = db.prepare(
    `SELECT id, price_approved, needs_price_review, review_status, review_round, harga_jual, approved_price
     FROM inquiry_items
     WHERE inquiry_id = ?`
  ).all(id) as Array<{
    id: string;
    price_approved: number;
    needs_price_review: number;
    review_status: string | null;
    review_round: number | null;
    harga_jual: number | null;
    approved_price: number | null;
  }>;

  const filteredItems = items.filter((item) =>
    item.price_approved !== 1 ||
    item.harga_jual == null ||
    item.review_status === 'review' ||
    item.needs_price_review === 1 ||
    (item.harga_jual != null && item.approved_price != null && item.harga_jual < item.approved_price)
  );

  // If nothing matches the automatic filter, send every item back for review.
  const itemsNeedingReview = filteredItems.length ? filteredItems : items;

  if (!itemsNeedingReview.length) {
    res.status(400).json({ error: 'Inquiry has no items.' }); return;
  }

  const reopeningReviewCount = itemsNeedingReview.filter((item) => item.review_status === 'review').length;

  const reviewIds = new Set(itemsNeedingReview.map((item) => item.id));
  const setNeedsReview = db.prepare('UPDATE inquiry_items SET price_approved = 0, needs_price_review = 1, review_status = ?, review_round = ? WHERE id = ?');
  const keepApproved = db.prepare("UPDATE inquiry_items SET price_approved = 1, needs_price_review = 0, review_status = 'approved' WHERE id = ?");
  const insertItemNote = db.prepare(
    'INSERT INTO inquiry_notes (id, inquiry_id, item_id, note, created_by, created_by_name, created_at) VALUES (?, ?, ?, ?, ?, ?, ?)'
  );

  const applyReviewRouting = db.transaction(() => {
    const now = new Date().toISOString();
    for (const item of items) {
      if (reviewIds.has(item.id)) {
        const nextRound = Number(item.review_round ?? 0) + 1;
        setNeedsReview.run('review', nextRound, item.id);
        insertItemNote.run(
          generateId(),
          id,
          item.id,
          `Price Review requested. Reason: ${reason}`,
          String(doneBy),
          String(doneByName ?? doneBy),
          now
        );
      } else {
        keepApproved.run(item.id);
      }
    }
  });
  applyReviewRouting();

  const rfqRowReview = db.prepare('SELECT rfq_no FROM inquiries WHERE id = ?').get(id) as { rfq_no: string | null } | undefined;
  const nowReview = new Date().toISOString();
  const oldStatus = inquiry.status;
  db.prepare('UPDATE inquiries SET status = ?, sourcing_pic = ?, updated_at = ?, updated_by = ?, price_approval_started_at = ? WHERE id = ?')
    .run('follow_up', String(doneByName ?? doneBy), nowReview, String(doneBy), nowReview, id);
  logActivity(
    id,
    `Sent to Price Review (${itemsNeedingReview.length} item${itemsNeedingReview.length > 1 ? 's' : ''})`,
    oldStatus,
    'follow_up',
    `Items: ${itemsNeedingReview.length}. Reason: ${reason}${reopeningReviewCount > 0 ? ' (includes negotiation review items)' : ''}`,
    String(doneBy),
    String(doneByName ?? doneBy)
  );
  insertAndBroadcast(
    'price_review', id, rfqRowReview?.rfq_no ?? null,
    `${rfqRowReview?.rfq_no ?? 'RFQ'} sent for price review — ${String(doneByName ?? doneBy)} (${itemsNeedingReview.length} item${itemsNeedingReview.length > 1 ? 's' : ''})`,
    String(doneBy), String(doneByName ?? doneBy),
  );
  res.json({ ok: true });
});

// POST /inquiries/:id/send-to-price-approved — Manager manually submits to Price Approved
inquiriesRouter.post('/:id/send-to-price-approved', (req: Request, res: Response) => {
  const { id } = req.params;
  const inquiry = db.prepare('SELECT id, status, rfq_no, sales_pic FROM inquiries WHERE id = ?').get(id) as { id: string; status: string; rfq_no: string | null; sales_pic: string | null } | undefined;
  if (!inquiry) { res.status(404).json({ error: 'Not found.' }); return; }
  if (inquiry.status !== 'price_approval' && inquiry.status !== 'follow_up') {
    res.status(400).json({ error: 'Inquiry must be in price_approval or follow_up status.' }); return;
  }

  const { doneBy, doneByName } = req.body as Record<string, unknown>;
  if (!doneBy) { res.status(400).json({ error: 'doneBy is required.' }); return; }

  db.prepare('UPDATE inquiries SET status = ?, updated_at = ?, updated_by = ?, price_approval_started_at = NULL WHERE id = ?')
    .run('price_approved', new Date().toISOString(), String(doneBy), id);
  logActivity(id, 'Sent to Price Approved', inquiry.status, 'price_approved', null, String(doneBy), String(doneByName ?? doneBy));

  const salesRecipient = usernameForPic(inquiry.sales_pic);
  if (salesRecipient) {
    insertAndBroadcast(
      'price_approved', id, inquiry.rfq_no ?? null,
      `${inquiry.rfq_no ?? 'RFQ'} is Price Approved — ready to send quotation`,
      String(doneBy), String(doneByName ?? doneBy),
      salesRecipient,
    );
  }
  res.json({ ok: true });
});

// POST /inquiries/:id/approve — Manager approves price
inquiriesRouter.post('/:id/approve', (req: Request, res: Response) => {
  const { id } = req.params;
  const inquiry = db.prepare('SELECT id, status FROM inquiries WHERE id = ?').get(id) as { id: string; status: string } | undefined;

  if (!inquiry) { res.status(404).json({ error: 'Not found.' }); return; }
  if (inquiry.status !== 'price_approval' && inquiry.status !== 'follow_up') {
    res.status(400).json({ error: 'Inquiry must be in price_approval or follow_up status.' }); return;
  }

  const { hargaJual, leadTimeCustomer, validitasQuotation, catatanQuotation, doneBy, doneByName } =
    req.body as Record<string, unknown>;

  if (!hargaJual) { res.status(400).json({ error: 'hargaJual is required.' }); return; }

  const item = db.prepare('SELECT id, harga_beli FROM inquiry_items WHERE inquiry_id = ? ORDER BY id LIMIT 1').get(id) as { id: string; harga_beli: number | null } | undefined;
  if (!item) { res.status(400).json({ error: 'No items found.' }); return; }

  const margin = item.harga_beli != null ? Number(hargaJual) - item.harga_beli : null;

  db.prepare(
    `UPDATE inquiry_items SET harga_jual = ?, approved_price = ?, margin = ?, lead_time_customer = ?,
       validitas_quotation = ?, catatan_quotation = ?, price_approved = 1, needs_price_review = 0, review_status = 'approved' WHERE id = ?`
  ).run(hargaJual, hargaJual, margin, leadTimeCustomer ?? null, validitasQuotation ?? null, catatanQuotation ?? null, item.id);

  logActivity(id, 'Price approved', inquiry.status, inquiry.status, String(catatanQuotation ?? ''), String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''));
  recalcInquiryStatus(id, String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''));
  res.json({ ok: true });
});

inquiriesRouter.post('/:id/items/:itemId/approve', (req: Request, res: Response) => {
  const { id, itemId } = req.params;
  const inquiry = db.prepare('SELECT id, status FROM inquiries WHERE id = ?').get(id) as { id: string; status: string } | undefined;
  if (!inquiry) { res.status(404).json({ error: 'Not found.' }); return; }
  if (inquiry.status !== 'price_approval' && inquiry.status !== 'follow_up') {
    res.status(400).json({ error: 'Inquiry must be in price_approval or follow_up status.' }); return;
  }

  const item = db.prepare('SELECT id, harga_beli FROM inquiry_items WHERE id = ? AND inquiry_id = ?').get(itemId, id) as { id: string; harga_beli: number | null } | undefined;
  if (!item) { res.status(404).json({ error: 'Item not found.' }); return; }

  const { hargaJual, leadTimeCustomer, validitasQuotation, catatanQuotation, doneBy, doneByName } =
    req.body as Record<string, unknown>;

  if (!hargaJual) { res.status(400).json({ error: 'hargaJual is required.' }); return; }

  const margin = item.harga_beli != null ? Number(hargaJual) - item.harga_beli : null;

  db.prepare(
    `UPDATE inquiry_items SET harga_jual = ?, approved_price = ?, margin = ?, lead_time_customer = ?,
       validitas_quotation = ?, catatan_quotation = ?, price_approved = 1, needs_price_review = 0, review_status = 'approved' WHERE id = ?`
  ).run(hargaJual, hargaJual, margin, leadTimeCustomer ?? null, validitasQuotation ?? null, catatanQuotation ?? null, itemId);

  logActivity(id, 'Price approved', inquiry.status, inquiry.status, String(catatanQuotation ?? ''), String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''));
  recalcInquiryStatus(id, String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''));
  res.json({ ok: true });
});

// POST /inquiries/:id/items/:itemId/reject — Manager sets counter price for negotiation
inquiriesRouter.post('/:id/items/:itemId/reject', (req: Request, res: Response) => {
  const { id, itemId } = req.params;
  const inquiry = db.prepare('SELECT id, status FROM inquiries WHERE id = ?').get(id) as { id: string; status: string } | undefined;
  if (!inquiry) { res.status(404).json({ error: 'Not found.' }); return; }
  if (inquiry.status !== 'price_approval' && inquiry.status !== 'follow_up') {
    res.status(400).json({ error: 'Inquiry must be in price_approval or follow_up status.' }); return;
  }

  const item = db.prepare(
    'SELECT id, review_round, item_name, review_status, needs_price_review FROM inquiry_items WHERE id = ? AND inquiry_id = ?'
  ).get(itemId, id) as {
    id: string;
    review_round: number | null;
    item_name: string | null;
    review_status: string | null;
    needs_price_review: number;
  } | undefined;
  if (!item) { res.status(404).json({ error: 'Item not found.' }); return; }
  if (item.review_status !== 'review' && item.needs_price_review !== 1) {
    res.status(400).json({ error: 'Only items in review status can be updated.' }); return;
  }

  const { doneBy, doneByName, reason, counterPrice } = req.body as Record<string, unknown>;
  const negotiationReason = String(reason ?? '').trim();
  const nextCounterPrice = Number(counterPrice);
  if (!doneBy) { res.status(400).json({ error: 'doneBy is required.' }); return; }
  if (!Number.isFinite(nextCounterPrice) || nextCounterPrice <= 0) {
    res.status(400).json({ error: 'counterPrice is required.' }); return;
  }

  const nextRound = Number(item.review_round ?? 0) + 1;
  db.prepare(
    `UPDATE inquiry_items
     SET price_approved = 0,
         needs_price_review = 1,
         review_status = 'review',
         approved_price = ?,
         review_round = ?
     WHERE id = ?`
  ).run(nextCounterPrice, nextRound, itemId);

  logActivity(
    id,
    `Counter price updated (${item.item_name ?? 'Item'})`,
    inquiry.status,
    inquiry.status,
    `Counter price: Rp ${nextCounterPrice.toLocaleString('id-ID')}${negotiationReason ? `. ${negotiationReason}` : ''}`,
    String(doneBy),
    String(doneByName ?? doneBy)
  );

  db.prepare(
    `INSERT INTO inquiry_notes (id, inquiry_id, item_id, note, created_by, created_by_name, created_at)
     VALUES (?, ?, ?, ?, ?, ?, ?)`
  ).run(
    generateId(),
    id,
    itemId,
    `Counter price set to Rp ${nextCounterPrice.toLocaleString('id-ID')}${negotiationReason ? `. Note: ${negotiationReason}` : ''}`,
    String(doneBy),
    String(doneByName ?? doneBy),
    new Date().toISOString()
  );

  res.json({ ok: true });
});

// POST /inquiries/:id/ready-to-purchase — move a quotation_sent inquiry to ready_to_purchase
inquiriesRouter.post('/:id/ready-to-purchase', (req: Request, res: Response) => {
  const { id } = req.params;
  const inquiry = db.prepare('SELECT id, status FROM inquiries WHERE id = ?').get(id) as { id: string; status: string } | undefined;

  if (!inquiry) { res.status(404).json({ error: 'Not found.' }); return; }
  if (!['quotation_sent', 'follow_up'].includes(inquiry.status)) {
    res.status(400).json({ error: 'Only quotation_sent or follow_up inquiries can be moved to Ready to Purchase.' }); return;
  }

  const { doneBy, doneByName } = req.body as Record<string, unknown>;
  const oldStatus = inquiry.status;

  db.prepare('UPDATE inquiries SET status = ?, updated_at = ?, updated_by = ? WHERE id = ?')
    .run('ready_to_purchase', new Date().toISOString(), doneBy, id);

  logActivity(id, 'Moved to Ready to Purchase', oldStatus, 'ready_to_purchase', null, String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''));
  res.json({ ok: true });
});

// GET /inquiries/:id/notes
inquiriesRouter.get('/:id/notes', (req: Request, res: Response) => {
  const { id } = req.params;
  const notes = db.prepare(
    'SELECT id, inquiry_id, item_id, note, created_by, created_by_name, created_at FROM inquiry_notes WHERE inquiry_id = ? ORDER BY created_at ASC'
  ).all(id) as Array<Record<string, unknown>>;
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
inquiriesRouter.post('/:id/notes', (req: Request, res: Response) => {
  const { id } = req.params;
  const { note, doneBy, doneByName } = req.body as Record<string, unknown>;
  const authUser = (req as any).user as { role: string };

  if (!note || !String(note).trim()) {
    res.status(400).json({ error: 'Note cannot be empty.' }); return;
  }

  const inquiry = db.prepare('SELECT id, sales_pic, sourcing_pic FROM inquiries WHERE id = ?')
    .get(id) as { id: string; sales_pic: string; sourcing_pic: string | null } | undefined;
  if (!inquiry) { res.status(404).json({ error: 'Not found.' }); return; }

  const isAdminOrManager = authUser.role === 'admin' || authUser.role === 'manager';
  const isAssigned = doneByName === inquiry.sales_pic || doneByName === inquiry.sourcing_pic;

  if (!isAdminOrManager && !isAssigned) {
    res.status(403).json({ error: 'Only assigned users can add comments.' }); return;
  }

  const noteId = generateId();
  db.prepare(
    'INSERT INTO inquiry_notes (id, inquiry_id, note, created_by, created_by_name, created_at) VALUES (?, ?, ?, ?, ?, ?)'
  ).run(noteId, id, String(note).trim(), doneBy, doneByName, new Date().toISOString());

  res.json({ id: noteId });
});

// GET /inquiries/:id/items/:itemId/notes
inquiriesRouter.get('/:id/items/:itemId/notes', (req: Request, res: Response) => {
  const { id, itemId } = req.params;
  const notes = db.prepare(
    'SELECT id, inquiry_id, item_id, note, created_by, created_by_name, created_at FROM inquiry_notes WHERE inquiry_id = ? AND item_id = ? ORDER BY created_at ASC'
  ).all(id, itemId) as Array<Record<string, unknown>>;
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
inquiriesRouter.post('/:id/items/:itemId/notes', (req: Request, res: Response) => {
  const { id, itemId } = req.params;
  const { note, doneBy, doneByName } = req.body as Record<string, unknown>;
  const authUser = (req as any).user as { role: string };

  if (!note || !String(note).trim()) {
    res.status(400).json({ error: 'Note cannot be empty.' }); return;
  }

  const inquiry = db.prepare('SELECT id, sales_pic, sourcing_pic FROM inquiries WHERE id = ?')
    .get(id) as { id: string; sales_pic: string; sourcing_pic: string | null } | undefined;
  if (!inquiry) { res.status(404).json({ error: 'Not found.' }); return; }

  const isAdminOrManager = authUser.role === 'admin' || authUser.role === 'manager';
  const isAssigned = doneByName === inquiry.sales_pic || doneByName === inquiry.sourcing_pic;

  if (!isAdminOrManager && !isAssigned) {
    res.status(403).json({ error: 'Only assigned users can add comments.' }); return;
  }

  const noteId = generateId();
  db.prepare(
    'INSERT INTO inquiry_notes (id, inquiry_id, item_id, note, created_by, created_by_name, created_at) VALUES (?, ?, ?, ?, ?, ?, ?)'
  ).run(noteId, id, itemId, String(note).trim(), doneBy, doneByName, new Date().toISOString());

  res.json({ id: noteId });
});

// PATCH /inquiries/:id/assign-sales — admin/manager only
inquiriesRouter.patch('/:id/assign-sales', (req: Request, res: Response) => {
  const { id } = req.params;
  const { salesPic, doneBy, doneByName } = req.body as Record<string, unknown>;
  const authUser = (req as any).user as { role: string };

  if (authUser.role !== 'admin' && authUser.role !== 'manager') {
    res.status(403).json({ error: 'Only admin or manager can reassign Sales PIC.' }); return;
  }
  if (!salesPic || !String(salesPic).trim()) {
    res.status(400).json({ error: 'salesPic is required.' }); return;
  }

  const inquiry = db.prepare('SELECT id, rfq_no, sales_pic FROM inquiries WHERE id = ?')
    .get(id) as { id: string; rfq_no: string | null; sales_pic: string | null } | undefined;
  if (!inquiry) { res.status(404).json({ error: 'Not found.' }); return; }

  const newPic = String(salesPic).trim();
  db.prepare('UPDATE inquiries SET sales_pic = ?, updated_at = ?, updated_by = ? WHERE id = ?')
    .run(newPic, new Date().toISOString(), doneBy, id);

  logActivity(id, `Sales PIC reassigned to: ${newPic}`, null, null, '', String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''));

  if (newPic !== (inquiry.sales_pic ?? '')) {
    const recipient = usernameForPic(newPic);
    if (recipient && recipient !== String(doneBy ?? '')) {
      insertAndBroadcast(
        'assigned_sales', id, inquiry.rfq_no ?? null,
        `${inquiry.rfq_no ?? 'RFQ'} assigned to you by ${String(doneByName ?? doneBy)}`,
        String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''),
        recipient,
      );
    }
  }

  res.json({ ok: true });
});

// PATCH /inquiries/:id/assign-sourcing
// Admin/manager can assign any sourcing user; sourcing can self-assign if unassigned
inquiriesRouter.patch('/:id/assign-sourcing', (req: Request, res: Response) => {
  const { id } = req.params;
  const { sourcingPic, doneBy, doneByName } = req.body as Record<string, unknown>;
  const authUser = (req as any).user as { role: string; menus: string[] };

  const inquiry = db.prepare('SELECT id, rfq_no, sourcing_pic FROM inquiries WHERE id = ?')
    .get(id) as { id: string; rfq_no: string | null; sourcing_pic: string | null } | undefined;
  if (!inquiry) { res.status(404).json({ error: 'Not found.' }); return; }

  const isAdminOrManager = authUser.role === 'admin' || authUser.role === 'manager';
  const hasSourcingMenu = authUser.menus?.includes('sourcing');

  if (!isAdminOrManager && !hasSourcingMenu) {
    res.status(403).json({ error: 'Not authorized.' }); return;
  }

  // Non-admin/manager sourcing users can only self-assign if unassigned
  if (hasSourcingMenu && !isAdminOrManager) {
    if (inquiry.sourcing_pic) {
      res.status(400).json({ error: 'Already assigned to another sourcing user.' }); return;
    }
    if (String(sourcingPic) !== String(doneByName)) {
      res.status(403).json({ error: 'Sourcing can only assign themselves.' }); return;
    }
  }

  const newPic = sourcingPic ? String(sourcingPic).trim() : null;
  db.prepare('UPDATE inquiries SET sourcing_pic = ?, updated_at = ?, updated_by = ? WHERE id = ?')
    .run(newPic, new Date().toISOString(), doneBy, id);

  logActivity(id, `Sourcing assigned: ${String(sourcingPic ?? 'unassigned')}`, null, null, '', String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''));

  // Notify the newly-assigned sourcing user (skip self-assignment and unassignment)
  if (newPic && newPic !== (inquiry.sourcing_pic ?? '') && newPic !== String(doneByName ?? '')) {
    const recipient = usernameForPic(newPic);
    if (recipient) {
      insertAndBroadcast(
        'assigned_sourcing', id, inquiry.rfq_no ?? null,
        `${inquiry.rfq_no ?? 'RFQ'} assigned to you by ${String(doneByName ?? doneBy)}`,
        String(doneBy ?? ''), String(doneByName ?? doneBy ?? ''),
        recipient,
      );
    }
  }

  res.json({ ok: true });
});




