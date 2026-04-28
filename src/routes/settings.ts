import { Router, Request, Response } from 'express';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { db } from '../db';
import { generateId } from '../utils/id';

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

// GET /settings/backup (admin only)
settingsRouter.get('/backup', async (req: Request, res: Response) => {
  const authUser = (req as any).user as { id: string; username: string; role: string } | undefined;
  if (!authUser || authUser.role !== 'admin') {
    res.status(403).json({ error: 'Only admin can backup the database.' });
    return;
  }

  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  const filename = `fjm-db-backup-${timestamp}.db`;
  const backupPath = path.join(os.tmpdir(), filename);

  try {
    await db.backup(backupPath);
    res.download(backupPath, filename, (err) => {
      fs.promises.rm(backupPath, { force: true }).catch(() => undefined);
      if (err && !res.headersSent) {
        res.status(500).json({ error: 'Failed to download database backup.' });
      }
    });
  } catch (err) {
    fs.promises.rm(backupPath, { force: true }).catch(() => undefined);
    res.status(500).json({ error: 'Failed to create database backup.' });
  }
});

// GET /settings/organizations
settingsRouter.get('/organizations', (_req: Request, res: Response) => {
  const rows = db.prepare(
    'SELECT id, code, created_at, created_by FROM organizations ORDER BY code ASC'
  ).all() as Array<{ id: string; code: string; created_at: string; created_by: string | null }>;
  res.json(rows.map((row) => ({
    id: row.id,
    code: row.code,
    createdAt: row.created_at,
    createdBy: row.created_by,
  })));
});

// POST /settings/organizations (admin only)
settingsRouter.post('/organizations', (req: Request, res: Response) => {
  const authUser = (req as any).user as { id: string; username: string; role: string } | undefined;
  if (!authUser || authUser.role !== 'admin') {
    res.status(403).json({ error: 'Only admin can manage organizations.' });
    return;
  }

  const { code } = req.body as { code?: string };
  const normalizedCode = String(code ?? '').trim().toUpperCase();
  if (!normalizedCode) {
    res.status(400).json({ error: 'code is required.' });
    return;
  }
  if (!/^[A-Z0-9_-]{2,20}$/.test(normalizedCode)) {
    res.status(400).json({ error: 'code must be 2-20 chars (A-Z, 0-9, _, -).' });
    return;
  }

  const existing = db.prepare('SELECT id FROM organizations WHERE code = ?').get(normalizedCode) as { id: string } | undefined;
  if (existing) {
    res.status(409).json({ error: 'Organization already exists.' });
    return;
  }

  const id = generateId();
  const createdAt = new Date().toISOString();
  db.prepare(
    'INSERT INTO organizations (id, code, created_at, created_by) VALUES (?, ?, ?, ?)'
  ).run(id, normalizedCode, createdAt, authUser.username);

  res.status(201).json({
    id,
    code: normalizedCode,
    createdAt,
    createdBy: authUser.username,
  });
});
