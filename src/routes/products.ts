import { Router, Request, Response } from 'express';
import { db } from '../db';
import { generateId } from '../utils/id';

export const productsRouter = Router();

productsRouter.get('/', (_req: Request, res: Response) => {
  const products = db
    .prepare(
      `SELECT id, name, image_data_url as imageDataUrl, proposed_price as proposedPrice,
              approved_price as approvedPrice, approved_source_id as approvedSourceId,
              lead_time_minutes as leadTimeMinutes, status,
              created_at as createdAt, created_by as createdBy, approved_at as approvedAt
       FROM products
       ORDER BY created_at DESC`
    )
    .all() as Array<Record<string, unknown>>;

  const sourceRows = db
    .prepare('SELECT id, product_id as productId, label, url, price FROM product_sources')
    .all() as Array<{ id: string; productId: string; label: string; url: string; price: number | null }>;

  const sourcesByProduct = new Map<string, Array<Record<string, unknown>>>();
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

productsRouter.post('/', (req: Request, res: Response) => {
  const { name, imageDataUrl, proposedPrice, leadTimeMinutes, sources, createdBy } = req.body as {
    name?: string;
    imageDataUrl?: string;
    proposedPrice?: number;
    leadTimeMinutes?: number;
    sources?: Array<{ label?: string; url?: string; price?: number }>;
    createdBy?: string;
  };

  if (!name || proposedPrice === undefined || leadTimeMinutes === undefined || !createdBy) {
    res.status(400).json({ error: 'Missing required fields.' });
    return;
  }

  const id = generateId();
  const createdAt = new Date().toISOString();

  const insertProduct = db.prepare(
    `INSERT INTO products (id, name, image_data_url, proposed_price, lead_time_minutes, status, created_at, created_by)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
  );

  const insertSource = db.prepare(
    'INSERT INTO product_sources (id, product_id, label, url, price) VALUES (?, ?, ?, ?, ?)'
  );

  const tx = db.transaction(() => {
    insertProduct.run(
      id,
      name,
      imageDataUrl ?? null,
      proposedPrice,
      leadTimeMinutes,
      'pending',
      createdAt,
      createdBy
    );

    (sources ?? []).slice(0, 3).forEach((source) => {
      insertSource.run(generateId(), id, source.label ?? null, source.url ?? null, source.price ?? null);
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

productsRouter.put('/:id', (req: Request, res: Response) => {
  const { id } = req.params;
  const { name, imageDataUrl, proposedPrice, leadTimeMinutes, sources } = req.body as {
    name?: string;
    imageDataUrl?: string | null;
    proposedPrice?: number;
    leadTimeMinutes?: number;
    sources?: Array<{ label?: string; url?: string; price?: number }>;
  };

  if (proposedPrice === undefined || proposedPrice <= 0) {
    res.status(400).json({ error: 'Proposed price must be greater than 0.' });
    return;
  }

  if (leadTimeMinutes === undefined || leadTimeMinutes <= 0) {
    res.status(400).json({ error: 'Lead time must be greater than 0.' });
    return;
  }

  const product = db
    .prepare('SELECT id, status FROM products WHERE id = ?')
    .get(id) as { id: string; status: string } | undefined;

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

  const updateProduct = db.prepare(
    'UPDATE products SET name = COALESCE(?, name), image_data_url = ?, proposed_price = ?, lead_time_minutes = ?, approved_source_id = NULL WHERE id = ?'
  );

  const insertSource = db.prepare(
    'INSERT INTO product_sources (id, product_id, label, url, price) VALUES (?, ?, ?, ?, ?)'
  );

  const tx = db.transaction(() => {
    updateProduct.run(updateName ?? null, imageDataUrl ?? null, proposedPrice, leadTimeMinutes, id);
    db.prepare('DELETE FROM product_sources WHERE product_id = ?').run(id);
    (sources ?? []).slice(0, 3).forEach((source) => {
      insertSource.run(generateId(), id, source.label ?? null, source.url ?? null, source.price ?? null);
    });
  });

  tx();
  res.json({ ok: true });
});

productsRouter.post('/:id/approve', (req: Request, res: Response) => {
  const { id } = req.params;
  const { approvedPrice, approvedSourceId } = req.body as { approvedPrice?: number; approvedSourceId?: string };

  const product = db.prepare('SELECT id FROM products WHERE id = ?').get(id) as { id: string } | undefined;
  if (!product) {
    res.status(404).json({ error: 'Product not found.' });
    return;
  }

  const approvedAt = new Date().toISOString();

  db.prepare(
    'UPDATE products SET status = ?, approved_price = ?, approved_source_id = ?, approved_at = ? WHERE id = ?'
  ).run('approved', approvedPrice ?? null, approvedSourceId ?? null, approvedAt, id);

  res.json({ ok: true });
});
