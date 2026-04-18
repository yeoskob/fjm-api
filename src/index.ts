import express from 'express';
import cors from 'cors';
import { authRouter } from './routes/auth';
import { usersRouter } from './routes/users';
import { productsRouter } from './routes/products';
import { inquiriesRouter } from './routes/inquiries';
import { rolesRouter } from './routes/roles';
import { settingsRouter } from './routes/settings';
import { notificationsRouter } from './routes/notifications';
import { requireAuth } from './middleware/requireAuth';
import './db';

const app = express();
app.use(cors());
app.use(express.json({ limit: '5mb' }));

app.get('/health', (_req, res) => {
  res.json({ ok: true });
});

app.use('/auth', authRouter);
app.use('/users', requireAuth, usersRouter);
app.use('/products', requireAuth, productsRouter);
app.use('/inquiries', requireAuth, inquiriesRouter);
app.use('/roles', requireAuth, rolesRouter);
app.use('/settings', requireAuth, settingsRouter);
app.use('/notifications', requireAuth, notificationsRouter);

const port = Number(process.env.PORT) || 4000;
app.listen(port, "0.0.0.0", () => {
  console.log(`API listening on ${port}`);
});
