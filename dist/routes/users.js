"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.usersRouter = void 0;
const express_1 = require("express");
const db_1 = require("../db");
const id_1 = require("../utils/id");
exports.usersRouter = (0, express_1.Router)();
exports.usersRouter.get('/', (_req, res) => {
    const users = db_1.db.prepare('SELECT id, name, username, role FROM users ORDER BY name').all();
    res.json(users);
});
exports.usersRouter.post('/', (req, res) => {
    const { name, username, password, role } = req.body;
    if (!name || !username || !password || !role) {
        res.status(400).json({ error: 'Missing required fields.' });
        return;
    }
    const existing = db_1.db
        .prepare('SELECT id FROM users WHERE LOWER(username) = LOWER(?)')
        .get(username);
    if (existing) {
        res.status(409).json({ error: 'Username already exists.' });
        return;
    }
    const id = (0, id_1.generateId)();
    db_1.db.prepare('INSERT INTO users (id, name, username, password, role) VALUES (?, ?, ?, ?, ?)').run(id, name, username, password, role);
    res.status(201).json({ id, name, username, role });
});
exports.usersRouter.put('/:id', (req, res) => {
    const { id } = req.params;
    const { name, username, password, role } = req.body;
    if (!name || !username || !role) {
        res.status(400).json({ error: 'Missing required fields.' });
        return;
    }
    const current = db_1.db.prepare('SELECT id FROM users WHERE id = ?').get(id);
    if (!current) {
        res.status(404).json({ error: 'User not found.' });
        return;
    }
    const duplicate = db_1.db
        .prepare('SELECT id FROM users WHERE LOWER(username) = LOWER(?) AND id != ?')
        .get(username, id);
    if (duplicate) {
        res.status(409).json({ error: 'Username already exists.' });
        return;
    }
    if (password) {
        db_1.db.prepare('UPDATE users SET name = ?, username = ?, password = ?, role = ? WHERE id = ?').run(name, username, password, role, id);
    }
    else {
        db_1.db.prepare('UPDATE users SET name = ?, username = ?, role = ? WHERE id = ?').run(name, username, role, id);
    }
    res.json({ id, name, username, role });
});
exports.usersRouter.delete('/:id', (req, res) => {
    const { id } = req.params;
    db_1.db.prepare('DELETE FROM users WHERE id = ?').run(id);
    res.status(204).send();
});
