import { Router, Request, Response } from 'express';
import { db } from '../db';
import { generateId } from '../utils/id';

export const notificationsRouter = Router();

// ── SSE client registry ──────────────────────────────────────────────────────

interface SseClient {
  res: Response;
  username: string;
  role: string;
}
const sseClients = new Set<SseClient>();

// ── Types ────────────────────────────────────────────────────────────────────

export type NotificationType =
  | 'price_approval'
  | 'price_review'
  | 'price_approved'
  | 'return_to_sourcing'
  | 'assigned_sales'
  | 'assigned_sourcing';

export interface NotificationRecord {
  id: string;
  type: string;
  inquiry_id: string;
  rfq_no: string | null;
  message: string;
  triggered_by: string;
  triggered_by_name: string;
  created_at: string;
  read_at: string | null;
  recipient_username: string | null;
}

function formatNotificationTitle(rfqNo: string | null | undefined, customer: string | null | undefined): string | null {
  const rfq = String(rfqNo ?? '').trim();
  if (!rfq) return null;

  const prefixedRfq = rfq.startsWith('#') ? rfq : `#${rfq}`;
  const customerName = String(customer ?? '').trim();
  return customerName ? `${prefixedRfq}-${customerName}` : prefixedRfq;
}

function notificationTitleForInquiry(inquiryId: string, rfqNo: string | null | undefined): string | null {
  const row = db.prepare('SELECT customer FROM inquiries WHERE id = ?').get(inquiryId) as { customer: string | null } | undefined;
  return formatNotificationTitle(rfqNo, row?.customer ?? null);
}

// ── Shared helper called from other routers ──────────────────────────────────

export function insertAndBroadcast(
  type: NotificationType,
  inquiryId: string,
  rfqNo: string | null,
  message: string,
  triggeredBy: string,
  triggeredByName: string,
  recipientUsername: string | null = null,
): void {
  const id = generateId();
  const createdAt = new Date().toISOString();

  db.prepare(
    `INSERT INTO notifications (id, type, inquiry_id, rfq_no, message, triggered_by, triggered_by_name, created_at, recipient_username)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
  ).run(id, type, inquiryId, rfqNo ?? null, message, triggeredBy, triggeredByName, createdAt, recipientUsername ?? null);

  const notif: NotificationRecord = {
    id, type, inquiry_id: inquiryId, rfq_no: notificationTitleForInquiry(inquiryId, rfqNo), message,
    triggered_by: triggeredBy, triggered_by_name: triggeredByName,
    created_at: createdAt, read_at: null, recipient_username: recipientUsername ?? null,
  };

  const payload = `data: ${JSON.stringify(notif)}\n\n`;
  for (const client of sseClients) {
    if (recipientUsername && client.username !== recipientUsername) continue;
    client.res.write(payload);
  }
}

/** Resolve a PIC field (which historically stores display name) to a username. */
export function usernameForPic(pic: string | null | undefined): string | null {
  if (!pic) return null;
  const row = db.prepare('SELECT username FROM users WHERE name = ? OR username = ?').get(pic, pic) as { username: string } | undefined;
  return row?.username ?? null;
}

// ── REST endpoints ───────────────────────────────────────────────────────────

// GET /notifications — unread notifications visible to the current user.
notificationsRouter.get('/', (req: Request, res: Response) => {
  const user = (req as any).user as { username: string; role: string };
  const rows = db.prepare(
    `SELECT n.*, i.customer AS customer
     FROM notifications n
     LEFT JOIN inquiries i ON i.id = n.inquiry_id
     WHERE n.read_at IS NULL
       AND (n.recipient_username IS NULL OR n.recipient_username = ?)
     ORDER BY n.created_at DESC`
  ).all(user.username) as Array<NotificationRecord & { customer: string | null }>;
  res.json(rows.map(({ customer, ...row }) => ({
    ...row,
    rfq_no: formatNotificationTitle(row.rfq_no, customer),
  })));
});

// POST /notifications/read-all — mark every unread notification visible to
// the current user as read (broadcasts stay for others, targeted-to-me get cleared)
notificationsRouter.post('/read-all', (req: Request, res: Response) => {
  const user = (req as any).user as { username: string; role: string };
  const now = new Date().toISOString();
  db.prepare(
    `UPDATE notifications SET read_at = ?
     WHERE read_at IS NULL AND recipient_username = ?`
  ).run(now, user.username);
  // Broadcast rows (recipient NULL) we mark read for everyone — that matches
  // the prior behaviour where "mark all read" cleared the queue globally.
  db.prepare(
    `UPDATE notifications SET read_at = ?
     WHERE read_at IS NULL AND recipient_username IS NULL`
  ).run(now);
  res.json({ ok: true });
});

// POST /notifications/:id/read — mark a single notification as read
notificationsRouter.post('/:id/read', (req: Request, res: Response) => {
  const { id } = req.params;
  db.prepare(`UPDATE notifications SET read_at = ? WHERE id = ?`)
    .run(new Date().toISOString(), id);
  res.json({ ok: true });
});

// GET /notifications/stream — SSE endpoint
notificationsRouter.get('/stream', (req: Request, res: Response) => {
  const user = (req as any).user as { username: string; role: string };

  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  res.flushHeaders();
  res.write(':connected\n\n');

  const client: SseClient = { res, username: user.username, role: user.role };
  sseClients.add(client);

  const heartbeat = setInterval(() => res.write(':heartbeat\n\n'), 25_000);

  req.on('close', () => {
    clearInterval(heartbeat);
    sseClients.delete(client);
  });
});
