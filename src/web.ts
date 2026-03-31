import express from 'express';
import path from 'path';
import { createProxyMiddleware } from 'http-proxy-middleware';

const app = express();

const distPath = path.resolve(__dirname, '../../fjm-app/dist/fjm-app');
app.use(express.static(distPath));

app.use(
  '/api',
  createProxyMiddleware({
    target: 'http://localhost:4000',
    changeOrigin: true,
    pathRewrite: { '^/api': '' },
  })
);

app.get('*', (_req, res) => {
  res.sendFile(path.join(distPath, 'index.html'));
});

app.listen(80, () => {
  console.log('Web server running on http://<YOUR_PUBLIC_IP>');
});
