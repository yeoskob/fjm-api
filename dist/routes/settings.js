"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.settingsRouter = void 0;
const express_1 = require("express");
const fs_1 = __importDefault(require("fs"));
const os_1 = __importDefault(require("os"));
const path_1 = __importDefault(require("path"));
const db_1 = require("../db");
const id_1 = require("../utils/id");
exports.settingsRouter = (0, express_1.Router)();
// GET /settings
exports.settingsRouter.get('/', (_req, res) => {
    const rows = db_1.db.prepare('SELECT key, value FROM settings').all();
    const result = {};
    for (const row of rows)
        result[row.key] = row.value;
    res.json(result);
});
// PUT /settings/:key
exports.settingsRouter.put('/:key', (req, res) => {
    const { key } = req.params;
    const { value } = req.body;
    if (value == null) {
        res.status(400).json({ error: 'value is required.' });
        return;
    }
    db_1.db.prepare('INSERT INTO settings (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value')
        .run(key, String(value));
    res.json({ key, value: String(value) });
});
// GET /settings/backup (admin only)
exports.settingsRouter.get('/backup', async (req, res) => {
    const authUser = req.user;
    if (!authUser || authUser.role !== 'admin') {
        res.status(403).json({ error: 'Only admin can backup the database.' });
        return;
    }
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const filename = `fjm-db-backup-${timestamp}.db`;
    const backupPath = path_1.default.join(os_1.default.tmpdir(), filename);
    try {
        await db_1.db.backup(backupPath);
        res.download(backupPath, filename, (err) => {
            fs_1.default.promises.rm(backupPath, { force: true }).catch(() => undefined);
            if (err && !res.headersSent) {
                res.status(500).json({ error: 'Failed to download database backup.' });
            }
        });
    }
    catch (err) {
        fs_1.default.promises.rm(backupPath, { force: true }).catch(() => undefined);
        res.status(500).json({ error: 'Failed to create database backup.' });
    }
});
// GET /settings/organizations
exports.settingsRouter.get('/organizations', (_req, res) => {
    const rows = db_1.db.prepare('SELECT id, code, created_at, created_by FROM organizations ORDER BY code ASC').all();
    res.json(rows.map((row) => ({
        id: row.id,
        code: row.code,
        createdAt: row.created_at,
        createdBy: row.created_by,
    })));
});
// POST /settings/organizations (admin only)
exports.settingsRouter.post('/organizations', (req, res) => {
    const authUser = req.user;
    if (!authUser || authUser.role !== 'admin') {
        res.status(403).json({ error: 'Only admin can manage organizations.' });
        return;
    }
    const { code } = req.body;
    const normalizedCode = String(code ?? '').trim().toUpperCase();
    if (!normalizedCode) {
        res.status(400).json({ error: 'code is required.' });
        return;
    }
    if (!/^[A-Z0-9_-]{2,20}$/.test(normalizedCode)) {
        res.status(400).json({ error: 'code must be 2-20 chars (A-Z, 0-9, _, -).' });
        return;
    }
    const existing = db_1.db.prepare('SELECT id FROM organizations WHERE code = ?').get(normalizedCode);
    if (existing) {
        res.status(409).json({ error: 'Organization already exists.' });
        return;
    }
    const id = (0, id_1.generateId)();
    const createdAt = new Date().toISOString();
    db_1.db.prepare('INSERT INTO organizations (id, code, created_at, created_by) VALUES (?, ?, ?, ?)').run(id, normalizedCode, createdAt, authUser.username);
    res.status(201).json({
        id,
        code: normalizedCode,
        createdAt,
        createdBy: authUser.username,
    });
});
