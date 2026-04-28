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
    .prepare('SELECT users.id, users.name, users.username, users.role FROM users JOIN sessions ON users.id = sessions.user_id WHERE sessions.token = ?')
    .get(token) as { id: string; name: string; username: string; role: string } | undefined;

  if (!user) {
    res.status(401).json({ error: 'Session expired. Please log in again.' });
    return;
  }

  (req as any).user = user;
  next();
}
