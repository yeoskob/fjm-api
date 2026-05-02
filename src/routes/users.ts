import { Router, Request, Response } from 'express';
import { db } from '../db';
import { generateId } from '../utils/id';

export const usersRouter = Router();

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

const roleExists = (role: string): boolean => {
  const existing = db.prepare('SELECT name FROM roles WHERE name = ?').get(role) as { name: string } | undefined;
  return Boolean(existing);
};

usersRouter.get('/', (_req: Request, res: Response) => {
  const rows = db.prepare(
    `SELECT u.id, u.name, u.username, u.role, r.menus
     FROM users u LEFT JOIN roles r ON r.name = u.role ORDER BY u.name`
  ).all() as Array<{ id: string; name: string; username: string; role: string; menus: string | null }>;
  const users = rows.map((u) => {
    let menus: string[] = [];
    try { menus = u.menus ? JSON.parse(u.menus) : []; } catch { menus = []; }
    return { id: u.id, name: u.name, username: u.username, role: u.role, menus };
  });
  res.json(users);
});

usersRouter.post('/', (req: Request, res: Response) => {
  if (!requireAdmin(req, res)) {
    return;
  }

  const { name, username, password, role } = req.body as {
    name?: string;
    username?: string;
    password?: string;
    role?: string;
  };

  if (!name || !username || !password || !role) {
    res.status(400).json({ error: 'Missing required fields.' });
    return;
  }

  if (!roleExists(role)) {
    res.status(400).json({ error: 'Role not found.' });
    return;
  }

  const existing = db
    .prepare('SELECT id FROM users WHERE LOWER(username) = LOWER(?)')
    .get(username) as { id: string } | undefined;

  if (existing) {
    res.status(409).json({ error: 'Username already exists.' });
    return;
  }

  const id = generateId();
  db.prepare('INSERT INTO users (id, name, username, password, role) VALUES (?, ?, ?, ?, ?)').run(
    id,
    name,
    username,
    password,
    role
  );

  res.status(201).json({ id, name, username, role });
});

usersRouter.put('/:id', (req: Request, res: Response) => {
  if (!requireAdmin(req, res)) {
    return;
  }

  const { id } = req.params;
  const { name, username, password, role } = req.body as {
    name?: string;
    username?: string;
    password?: string;
    role?: string;
  };

  if (!name || !username || !role) {
    res.status(400).json({ error: 'Missing required fields.' });
    return;
  }

  if (!roleExists(role)) {
    res.status(400).json({ error: 'Role not found.' });
    return;
  }

  const current = db.prepare('SELECT id FROM users WHERE id = ?').get(id) as { id: string } | undefined;
  if (!current) {
    res.status(404).json({ error: 'User not found.' });
    return;
  }

  const duplicate = db
    .prepare('SELECT id FROM users WHERE LOWER(username) = LOWER(?) AND id != ?')
    .get(username, id) as { id: string } | undefined;

  if (duplicate) {
    res.status(409).json({ error: 'Username already exists.' });
    return;
  }

  if (password) {
    db.prepare('UPDATE users SET name = ?, username = ?, password = ?, role = ? WHERE id = ?').run(
      name,
      username,
      password,
      role,
      id
    );
  } else {
    db.prepare('UPDATE users SET name = ?, username = ?, role = ? WHERE id = ?').run(name, username, role, id);
  }

  res.json({ id, name, username, role });
});

usersRouter.delete('/:id', (req: Request, res: Response) => {
  if (!requireAdmin(req, res)) {
    return;
  }

  const { id } = req.params;
  db.prepare('DELETE FROM users WHERE id = ?').run(id);
  res.status(204).send();
});
