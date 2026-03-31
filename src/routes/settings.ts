import { Router, Request, Response } from 'express';
import { db } from '../db';

export const settingsRouter = Router();

// GET /settings
settingsRouter.get('/', (_req: Request, res: Response) => {
  const rows = db.prepare('SELECT key, value FROM settings').all() as Array<{ key: string; value: string }>;
  const result: Record<string, string> = {};
  for (const row of rows) result[row.key] = row.value;
  res.json(result);
});

// PUT /settings/:key
settingsRouter.put('/:key', (req: Request, res: Response) => {
  const { key } = req.params;
  const { value } = req.body as { value: string };
  if (value == null) { res.status(400).json({ error: 'value is required.' }); return; }
  db.prepare('INSERT INTO settings (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value')
    .run(key, String(value));
  res.json({ key, value: String(value) });
});
