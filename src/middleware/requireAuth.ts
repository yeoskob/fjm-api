import { Request, Response, NextFunction } from 'express';
import { db } from '../db';

export function requireAuth(req: Request, res: Response, next: NextFunction): void {
  const auth = req.headers['authorization'];
  // Also accept token as query param for EventSource (SSE) connections which
  // cannot set custom headers in the browser.
  const token = auth?.startsWith('Bearer ') ? auth.slice(7)
    : (req.query['token'] as string | undefined) ?? null;

  if (!token) {
    res.status(401).json({ error: 'Unauthorized.' });
    return;
  }

  const user = db
    .prepare(`SELECT u.id, u.name, u.username, u.role, r.menus
              FROM users u
              JOIN sessions s ON u.id = s.user_id
              LEFT JOIN roles r ON r.name = u.role
              WHERE s.token = ?`)
    .get(token) as { id: string; name: string; username: string; role: string; menus: string | null } | undefined;

  if (!user) {
    res.status(401).json({ error: 'Session expired. Please log in again.' });
    return;
  }

  let menus: string[] = [];
  try { menus = user.menus ? JSON.parse(user.menus) : []; } catch { menus = []; }
  (req as any).user = { ...user, menus };
  next();
}
