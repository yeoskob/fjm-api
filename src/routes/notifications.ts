import { Router, Request, Response } from 'express';
import { db } from '../db';
import { generateId } from '../utils/id';

export const notificationsRouter = Router();

// ── In-memory SSE client registry ─────────────────────────────────────────────

const sseClients = new Set<Response>();

// ── Shared helper called from other routers ────────────────────────────────────

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
}

export function insertAndBroadcast(
  type: 'price_approval' | 'price_review',
  inquiryId: string,
  rfqNo: string | null,
  message: string,
  triggeredBy: string,
  triggeredByName: string,
): void {
  const id = generateId();
  const createdAt = new Date().toISOString();

  db.prepare(
    `INSERT INTO notifications (id, type, inquiry_id, rfq_no, message, triggered_by, triggered_by_name, created_at)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
  ).run(id, type, inquiryId, rfqNo ?? null, message, triggeredBy, triggeredByName, createdAt);

  const notif: NotificationRecord = {
    id, type, inquiry_id: inquiryId, rfq_no: rfqNo ?? null, message,
    triggered_by: triggeredBy, triggered_by_name: triggeredByName,
    created_at: createdAt, read_at: null,
  };

  const payload = `data: ${JSON.stringify(notif)}\n\n`;
  for (const client of sseClients) {
    client.write(payload);
  }
}

// ── REST endpoints ─────────────────────────────────────────────────────────────

// GET /notifications — all unread, newest first
notificationsRouter.get('/', (_req: Request, res: Response) => {
  const rows = db.prepare(
    `SELECT * FROM notifications WHERE read_at IS NULL ORDER BY created_at DESC`
  ).all() as NotificationRecord[];
  res.json(rows);
});

// POST /notifications/read-all — mark every unread notification as read
notificationsRouter.post('/read-all', (_req: Request, res: Response) => {
  db.prepare(`UPDATE notifications SET read_at = ? WHERE read_at IS NULL`)
    .run(new Date().toISOString());
  res.json({ ok: true });
});

// POST /notifications/:id/read — mark a single notification as read
notificationsRouter.post('/:id/read', (req: Request, res: Response) => {
  const { id } = req.params;
  db.prepare(`UPDATE notifications SET read_at = ? WHERE id = ?`)
    .run(new Date().toISOString(), id);
  res.json({ ok: true });
});

// GET /notifications/stream — SSE endpoint (token passed as ?token= because
// the browser EventSource API cannot set custom headers; requireAuth accepts it)
notificationsRouter.get('/stream', (req: Request, res: Response) => {
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  res.flushHeaders();

  // Immediate acknowledgement so the client knows the connection is live
  res.write(':connected\n\n');

  sseClients.add(res);

  // Keep-alive ping every 25 s (proxies typically close idle connections at 30 s)
  const heartbeat = setInterval(() => res.write(':heartbeat\n\n'), 25_000);

  req.on('close', () => {
    clearInterval(heartbeat);
    sseClients.delete(res);
  });
});
