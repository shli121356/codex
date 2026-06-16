import react from '@vitejs/plugin-react';
import fs from 'node:fs';
import path from 'node:path';
import { defineConfig } from 'vite';

const WATCH_DIR = path.resolve('C:/Users/Administrator/Desktop/表格测试');

function isWorkbook(fileName) {
  return /\.(xlsx|xls)$/i.test(fileName) && !fileName.startsWith('~$');
}

function listWorkbooks() {
  if (!fs.existsSync(WATCH_DIR)) return [];
  return fs.readdirSync(WATCH_DIR)
    .filter(isWorkbook)
    .map((name) => {
      const fullPath = path.join(WATCH_DIR, name);
      const stat = fs.statSync(fullPath);
      return {
        name,
        size: stat.size,
        mtimeMs: stat.mtimeMs,
        url: `/api/watched-workbooks/file/${encodeURIComponent(name)}`
      };
    })
    .sort((a, b) => a.name.localeCompare(b.name, 'zh-CN'));
}

function sendJson(res, payload, statusCode = 200) {
  res.statusCode = statusCode;
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.end(JSON.stringify(payload));
}

function configureWorkbookApi(server) {
  const clients = new Set();
  let watcher;
  let notifyTimer;

  function notifyClients() {
    clearTimeout(notifyTimer);
    notifyTimer = setTimeout(() => {
      const payload = JSON.stringify({ files: listWorkbooks(), folder: WATCH_DIR });
      for (const res of clients) {
        res.write(`event: workbooks\n`);
        res.write(`data: ${payload}\n\n`);
      }
    }, 240);
  }

  if (fs.existsSync(WATCH_DIR)) {
    watcher = fs.watch(WATCH_DIR, notifyClients);
  }

  server.middlewares.use((req, res, next) => {
    const reqUrl = new URL(req.url || '/', 'http://localhost');

    if (reqUrl.pathname === '/api/watched-workbooks') {
      sendJson(res, { files: listWorkbooks(), folder: WATCH_DIR });
      return;
    }

    if (reqUrl.pathname === '/api/watched-workbooks/events') {
      res.writeHead(200, {
        'Content-Type': 'text/event-stream; charset=utf-8',
        'Cache-Control': 'no-cache, no-transform',
        Connection: 'keep-alive'
      });
      res.write(`event: workbooks\n`);
      res.write(`data: ${JSON.stringify({ files: listWorkbooks(), folder: WATCH_DIR })}\n\n`);
      clients.add(res);
      req.on('close', () => clients.delete(res));
      return;
    }

    if (reqUrl.pathname.startsWith('/api/watched-workbooks/file/')) {
      const fileName = path.basename(decodeURIComponent(reqUrl.pathname.replace('/api/watched-workbooks/file/', '')));
      const filePath = path.resolve(WATCH_DIR, fileName);
      if (!filePath.startsWith(WATCH_DIR) || !isWorkbook(fileName) || !fs.existsSync(filePath)) {
        sendJson(res, { error: 'Workbook not found' }, 404);
        return;
      }
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      fs.createReadStream(filePath).pipe(res);
      return;
    }

    next();
  });

  server.httpServer?.once('close', () => {
    clearTimeout(notifyTimer);
    watcher?.close();
    for (const res of clients) res.end();
    clients.clear();
  });
}

export default defineConfig({
  plugins: [
    react(),
    {
      name: 'local-workbook-watch-api',
      configureServer: configureWorkbookApi
    }
  ]
});
