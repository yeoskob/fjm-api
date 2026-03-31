import { Router, Request, Response } from 'express';
import { randomUUID } from 'crypto';
import { db } from '../db';

export const authRouter = Router();

authRouter.post('/login', (req: Request, res: Response) => {
  const { username, password } = req.body as { username?: string; password?: string };
  if (!username || !password) {
    res.status(400).json({ error: 'Username and password are required.' });
    return;
  }

  const user = db
    .prepare('SELECT id, name, username, role, password FROM users WHERE LOWER(username) = LOWER(?)')
    .get(username) as { id: string; name: string; username: string; role: string; password: string } | undefined;

  if (!user || user.password !== password) {
    res.status(401).json({ error: 'Invalid username or password.' });
    return;
  }

  const token = db.transaction((userId: string) => {
    // Enforce single active session per user by clearing old sessions.
    db.prepare('DELETE FROM sessions WHERE user_id = ?').run(userId);
    const nextToken = randomUUID();
    db.prepare('INSERT INTO sessions (token, user_id) VALUES (?, ?)').run(nextToken, userId);
    return nextToken;
  })(user.id);

  res.json({ id: user.id, name: user.name, username: user.username, role: user.role, token });
});

authRouter.post('/logout', (req: Request, res: Response) => {
  const auth = req.headers['authorization'];
  const headerToken = auth?.startsWith('Bearer ') ? auth.slice(7) : null;
  const bodyToken = (req.body as { token?: string })?.token ?? null;
  const token = headerToken ?? bodyToken;
  if (token) {
    db.prepare('DELETE FROM sessions WHERE token = ?').run(token);
  }
  res.json({ ok: true });
});
