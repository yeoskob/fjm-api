"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.settingsRouter = void 0;
const express_1 = require("express");
const db_1 = require("../db");
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
