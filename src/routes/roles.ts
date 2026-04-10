import { Router, Request, Response } from 'express';
import { db } from '../db';

export const rolesRouter = Router();


const getUserRole = (req: Request): string | undefined => {
  return (req as { user?: { role?: string } }).user?.role;
};

const requireAdmin = (req: Request, res: Response): boolean => {
  if (getUserRole(req) !== 'admin') {
    res.status(403).json({ error: 'Admin access required.' });
    return false;
  }
  return true;
};

const normalizeMenus = (menus: unknown): string[] => {
  if (!Array.isArray(menus)) return [];
  const cleaned = menus
    .map((menu) => String(menu).trim())
    .filter((menu) => menu.length > 0);
  return Array.from(new Set(cleaned));
};

function parseTabs(raw: string | null): Record<string, string[]> {
  if (!raw) return {};
  try {
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === 'object' && !Array.isArray(parsed) ? parsed : {};
  } catch { return {}; }
}

rolesRouter.get('/', (_req: Request, res: Response) => {
  const rows = db.prepare('SELECT name, menus, tabs FROM roles ORDER BY name').all() as Array<{ name: string; menus: string; tabs: string }>;
  const roles = rows.map((row) => {
    let menus: string[] = [];
    try {
      const parsed = JSON.parse(row.menus);
      menus = Array.isArray(parsed) ? parsed : [];
    } catch { menus = []; }
    return { name: row.name, menus, tabs: parseTabs(row.tabs) };
  });
  res.json(roles);
});

rolesRouter.post('/', (req: Request, res: Response) => {
  if (!requireAdmin(req, res)) return;

  const { name, menus, tabs } = req.body as { name?: string; menus?: unknown; tabs?: unknown };
  const roleName = String(name ?? '').trim();
  if (!roleName) {
    res.status(400).json({ error: 'Role name is required.' });
    return;
  }

  const normalizedMenus = normalizeMenus(menus);
  if (normalizedMenus.length === 0) {
    res.status(400).json({ error: 'Select at least one menu.' });
    return;
  }

  const existing = db.prepare('SELECT name FROM roles WHERE name = ?').get(roleName) as { name: string } | undefined;
  if (existing) {
    res.status(409).json({ error: 'Role already exists.' });
    return;
  }

  const tabsJson = JSON.stringify(tabs && typeof tabs === 'object' && !Array.isArray(tabs) ? tabs : {});
  db.prepare('INSERT INTO roles (name, menus, tabs, created_at) VALUES (?, ?, ?, ?)')
    .run(roleName, JSON.stringify(normalizedMenus), tabsJson, new Date().toISOString());

  res.status(201).json({ name: roleName, menus: normalizedMenus, tabs: parseTabs(tabsJson) });
});

rolesRouter.put('/:name', (req: Request, res: Response) => {
  if (!requireAdmin(req, res)) return;

  const { name } = req.params;
  const { menus, tabs } = req.body as { menus?: unknown; tabs?: unknown };
  const normalizedMenus = normalizeMenus(menus);
  if (normalizedMenus.length === 0) {
    res.status(400).json({ error: 'Select at least one menu.' });
    return;
  }

  const existing = db.prepare('SELECT name FROM roles WHERE name = ?').get(name) as { name: string } | undefined;
  if (!existing) {
    res.status(404).json({ error: 'Role not found.' });
    return;
  }

  const tabsJson = JSON.stringify(tabs && typeof tabs === 'object' && !Array.isArray(tabs) ? tabs : {});
  db.prepare('UPDATE roles SET menus = ?, tabs = ? WHERE name = ?').run(JSON.stringify(normalizedMenus), tabsJson, name);
  res.json({ name, menus: normalizedMenus, tabs: parseTabs(tabsJson) });
});

rolesRouter.delete('/:name', (req: Request, res: Response) => {
  if (!requireAdmin(req, res)) return;

  const { name } = req.params;
  const inUse = db.prepare('SELECT COUNT(*) as c FROM users WHERE role = ?').get(name) as { c: number };
  if (inUse.c > 0) {
    res.status(400).json({ error: 'Role is assigned to users.' });
    return;
  }

  db.prepare('DELETE FROM roles WHERE name = ?').run(name);
  res.status(204).send();
});
