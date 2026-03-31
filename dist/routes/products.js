"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.productsRouter = void 0;
const express_1 = require("express");
const db_1 = require("../db");
const id_1 = require("../utils/id");
exports.productsRouter = (0, express_1.Router)();
exports.productsRouter.get('/', (_req, res) => {
    const products = db_1.db
        .prepare(`SELECT id, name, image_data_url as imageDataUrl, proposed_price as proposedPrice,
              approved_price as approvedPrice, approved_source_id as approvedSourceId,
              lead_time_minutes as leadTimeMinutes, status,
              created_at as createdAt, created_by as createdBy, approved_at as approvedAt
       FROM products
       ORDER BY created_at DESC`)
        .all();
    const sourceRows = db_1.db
        .prepare('SELECT id, product_id as productId, label, url, price FROM product_sources')
        .all();
    const sourcesByProduct = new Map();
    for (const source of sourceRows) {
        const list = sourcesByProduct.get(source.productId) ?? [];
        list.push({ id: source.id, label: source.label, url: source.url, price: source.price ?? undefined });
        sourcesByProduct.set(source.productId, list);
    }
    const payload = products.map((product) => ({
        ...product,
        sources: sourcesByProduct.get(String(product.id)) ?? [],
    }));
    res.json(payload);
});
exports.productsRouter.post('/', (req, res) => {
    const { name, imageDataUrl, proposedPrice, leadTimeMinutes, sources, createdBy } = req.body;
    if (!name || proposedPrice === undefined || leadTimeMinutes === undefined || !createdBy) {
        res.status(400).json({ error: 'Missing required fields.' });
        return;
    }
    const id = (0, id_1.generateId)();
    const createdAt = new Date().toISOString();
    const insertProduct = db_1.db.prepare(`INSERT INTO products (id, name, image_data_url, proposed_price, lead_time_minutes, status, created_at, created_by)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?)`);
    const insertSource = db_1.db.prepare('INSERT INTO product_sources (id, product_id, label, url, price) VALUES (?, ?, ?, ?, ?)');
    const tx = db_1.db.transaction(() => {
        insertProduct.run(id, name, imageDataUrl ?? null, proposedPrice, leadTimeMinutes, 'pending', createdAt, createdBy);
        (sources ?? []).slice(0, 3).forEach((source) => {
            insertSource.run((0, id_1.generateId)(), id, source.label ?? null, source.url ?? null, source.price ?? null);
        });
    });
    tx();
    res.status(201).json({
        id,
        name,
        imageDataUrl,
        proposedPrice,
        leadTimeMinutes,
        status: 'pending',
        createdAt,
        createdBy,
        sources: sources ?? [],
    });
});
exports.productsRouter.put('/:id', (req, res) => {
    const { id } = req.params;
    const { name, imageDataUrl, proposedPrice, leadTimeMinutes, sources } = req.body;
    if (proposedPrice === undefined || proposedPrice <= 0) {
        res.status(400).json({ error: 'Proposed price must be greater than 0.' });
        return;
    }
    if (leadTimeMinutes === undefined || leadTimeMinutes <= 0) {
        res.status(400).json({ error: 'Lead time must be greater than 0.' });
        return;
    }
    const product = db_1.db
        .prepare('SELECT id, status FROM products WHERE id = ?')
        .get(id);
    if (!product) {
        res.status(404).json({ error: 'Product not found.' });
        return;
    }
    if (product.status !== 'pending') {
        res.status(400).json({ error: 'Only pending products can be edited.' });
        return;
    }
    const updateName = name?.trim();
    if (updateName !== undefined && updateName.length === 0) {
        res.status(400).json({ error: 'Name cannot be empty.' });
        return;
    }
    const updateProduct = db_1.db.prepare('UPDATE products SET name = COALESCE(?, name), image_data_url = ?, proposed_price = ?, lead_time_minutes = ?, approved_source_id = NULL WHERE id = ?');
    const insertSource = db_1.db.prepare('INSERT INTO product_sources (id, product_id, label, url, price) VALUES (?, ?, ?, ?, ?)');
    const tx = db_1.db.transaction(() => {
        updateProduct.run(updateName ?? null, imageDataUrl ?? null, proposedPrice, leadTimeMinutes, id);
        db_1.db.prepare('DELETE FROM product_sources WHERE product_id = ?').run(id);
        (sources ?? []).slice(0, 3).forEach((source) => {
            insertSource.run((0, id_1.generateId)(), id, source.label ?? null, source.url ?? null, source.price ?? null);
        });
    });
    tx();
    res.json({ ok: true });
});
exports.productsRouter.post('/:id/approve', (req, res) => {
    const { id } = req.params;
    const { approvedPrice, approvedSourceId } = req.body;
    const product = db_1.db.prepare('SELECT id FROM products WHERE id = ?').get(id);
    if (!product) {
        res.status(404).json({ error: 'Product not found.' });
        return;
    }
    const approvedAt = new Date().toISOString();
    db_1.db.prepare('UPDATE products SET status = ?, approved_price = ?, approved_source_id = ?, approved_at = ? WHERE id = ?').run('approved', approvedPrice ?? null, approvedSourceId ?? null, approvedAt, id);
    res.json({ ok: true });
});
