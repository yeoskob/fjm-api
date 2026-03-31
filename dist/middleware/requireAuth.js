"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.requireAuth = void 0;
const db_1 = require("../db");
function requireAuth(req, res, next) {
    const auth = req.headers['authorization'];
    const token = auth?.startsWith('Bearer ') ? auth.slice(7) : null;
    if (!token) {
        res.status(401).json({ error: 'Unauthorized.' });
        return;
    }
    const user = db_1.db
        .prepare('SELECT users.id, users.name, users.username, users.role FROM users JOIN sessions ON users.id = sessions.user_id WHERE sessions.token = ?')
        .get(token);
    if (!user) {
        res.status(401).json({ error: 'Session expired. Please log in again.' });
        return;
    }
    req.user = user;
    next();
}
exports.requireAuth = requireAuth;
