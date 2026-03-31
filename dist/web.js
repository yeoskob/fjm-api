"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const express_1 = __importDefault(require("express"));
const path_1 = __importDefault(require("path"));
const http_proxy_middleware_1 = require("http-proxy-middleware");
const app = (0, express_1.default)();
const distPath = path_1.default.resolve(__dirname, '../../fjm-app/dist/fjm-app');
app.use(express_1.default.static(distPath));
app.use('/api', (0, http_proxy_middleware_1.createProxyMiddleware)({
    target: 'http://localhost:4000',
    changeOrigin: true,
    pathRewrite: { '^/api': '' },
}));
app.get('*', (_req, res) => {
    res.sendFile(path_1.default.join(distPath, 'index.html'));
});
app.listen(80, () => {
    console.log('Web server running on http://<YOUR_PUBLIC_IP>');
});
