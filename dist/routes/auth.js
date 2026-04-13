"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.authRouter = void 0;
const express_1 = require("express");
const crypto_1 = require("crypto");
const db_1 = require("../db");
exports.authRouter = (0, express_1.Router)();
exports.authRouter.post('/login', (req, res) => {
    const { username, password } = req.body;
    if (!username || !password) {
        res.status(400).json({ error: 'Username and password are required.' });
        return;
    }
    const user = db_1.db
        .prepare('SELECT id, name, username, role, password FROM users WHERE LOWER(username) = LOWER(?)')
        .get(username);
    if (!user || user.password !== password) {
        res.status(401).json({ error: 'Invalid username or password.' });
        return;
    }
    const roleRow = db_1.db.prepare('SELECT menus, tabs FROM roles WHERE name = ?').get(user.role);
    let menus = [];
    let tabs = {};
    if (roleRow?.menus) {
        try {
            const parsed = JSON.parse(roleRow.menus);
            if (Array.isArray(parsed))
                menus = parsed.map((m) => String(m));
        }
        catch {
            menus = [];
        }
    }
    if (roleRow?.tabs) {
        try {
            const parsed = JSON.parse(roleRow.tabs);
            if (parsed && typeof parsed === 'object' && !Array.isArray(parsed))
                tabs = parsed;
        }
        catch {
            tabs = {};
        }
    }
    const token = db_1.db.transaction((userId) => {
        // Enforce single active session per user by clearing old sessions.
        db_1.db.prepare('DELETE FROM sessions WHERE user_id = ?').run(userId);
        const nextToken = (0, crypto_1.randomUUID)();
        db_1.db.prepare('INSERT INTO sessions (token, user_id) VALUES (?, ?)').run(nextToken, userId);
        return nextToken;
    })(user.id);
    res.json({ id: user.id, name: user.name, username: user.username, role: user.role, menus, tabs, token });
});
exports.authRouter.post('/logout', (req, res) => {
    const auth = req.headers['authorization'];
    const headerToken = auth?.startsWith('Bearer ') ? auth.slice(7) : null;
    const bodyToken = req.body?.token ?? null;
    const token = headerToken ?? bodyToken;
    if (token) {
        db_1.db.prepare('DELETE FROM sessions WHERE token = ?').run(token);
    }
    res.json({ ok: true });
});
