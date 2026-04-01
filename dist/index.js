"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const express_1 = __importDefault(require("express"));
const cors_1 = __importDefault(require("cors"));
const auth_1 = require("./routes/auth");
const users_1 = require("./routes/users");
const products_1 = require("./routes/products");
const inquiries_1 = require("./routes/inquiries");
const settings_1 = require("./routes/settings");
const requireAuth_1 = require("./middleware/requireAuth");
require("./db");
const app = (0, express_1.default)();
app.use((0, cors_1.default)());
app.use(express_1.default.json({ limit: '5mb' }));
app.get('/health', (_req, res) => {
    res.json({ ok: true });
});
app.use('/auth', auth_1.authRouter);
app.use('/users', requireAuth_1.requireAuth, users_1.usersRouter);
app.use('/products', requireAuth_1.requireAuth, products_1.productsRouter);
app.use('/inquiries', requireAuth_1.requireAuth, inquiries_1.inquiriesRouter);
app.use('/settings', requireAuth_1.requireAuth, settings_1.settingsRouter);
const port = Number(process.env.PORT) || 4000;
app.listen(port, "0.0.0.0", () => {
    console.log(`API listening on ${port}`);
});
