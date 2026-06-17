import react from '@vitejs/plugin-react';
import { execFile } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { promisify } from 'node:util';
import { defineConfig } from 'vite';

const DEFAULT_WATCH_DIR = path.resolve('C:/Users/Administrator/Desktop/表格测试');
const execFileAsync = promisify(execFile);
let watchDir = DEFAULT_WATCH_DIR;

function isWorkbook(fileName) {
  return /\.(xlsx|xls)$/i.test(fileName) && !fileName.startsWith('~$');
}

function listWorkbooks() {
  try {
    if (!fs.existsSync(watchDir)) return [];
    return fs.readdirSync(watchDir)
      .filter(isWorkbook)
      .map((name) => {
        try {
          const fullPath = path.join(watchDir, name);
          const stat = fs.statSync(fullPath);
          if (!stat.isFile()) return null;
          return {
            name,
            size: stat.size,
            mtimeMs: stat.mtimeMs,
            url: `/api/watched-workbooks/file/${encodeURIComponent(name)}`
          };
        } catch {
          return null;
        }
      })
      .filter(Boolean)
      .sort((a, b) => a.name.localeCompare(b.name, 'zh-CN'));
  } catch {
    return [];
  }
}

function readJsonBody(req) {
  return new Promise((resolve, reject) => {
    let body = '';
    req.on('data', chunk => {
      body += chunk;
      if (body.length > 100_000) reject(new Error('Request body too large'));
    });
    req.on('end', () => {
      try {
        resolve(body ? JSON.parse(body) : {});
      } catch (error) {
        reject(error);
      }
    });
    req.on('error', reject);
  });
}

function sendJson(res, payload, statusCode = 200) {
  res.statusCode = statusCode;
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.end(JSON.stringify(payload));
}

function isInsideWatchDir(filePath) {
  const relative = path.relative(watchDir, filePath);
  return relative && !relative.startsWith('..') && !path.isAbsolute(relative);
}

async function openFolderPicker() {
  const script = [
    'Add-Type -AssemblyName System.Windows.Forms',
    'Add-Type -AssemblyName System.Drawing',
    '$dialog = New-Object System.Windows.Forms.FolderBrowserDialog',
    "$dialog.Description = '选择需要监听的表格文件夹'",
    `$dialog.SelectedPath = '${watchDir.replace(/'/g, "''")}'`,
    '$dialog.ShowNewFolderButton = $true',
    "if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8; Write-Output $dialog.SelectedPath }"
  ].join('; ');
  const { stdout } = await execFileAsync('powershell.exe', ['-NoProfile', '-STA', '-Command', script], {
    encoding: 'utf8',
    windowsHide: false,
    timeout: 120000
  });
  return stdout.trim();
}

function configureWorkbookApi(server) {
  const clients = new Set();
  let watcher;
  let notifyTimer;

  function folderPayload() {
    return { files: listWorkbooks(), folder: watchDir, exists: fs.existsSync(watchDir) };
  }

  function notifyClients() {
    clearTimeout(notifyTimer);
    notifyTimer = setTimeout(() => {
      const payload = JSON.stringify(folderPayload());
      for (const res of clients) {
        res.write(`event: workbooks\n`);
        res.write(`data: ${payload}\n\n`);
      }
    }, 240);
  }

  function startWatcher() {
    watcher?.close();
    watcher = undefined;
    try {
      if (fs.existsSync(watchDir)) {
        watcher = fs.watch(watchDir, notifyClients);
        watcher.on('error', notifyClients);
      }
    } catch {
      watcher = undefined;
    }
    notifyClients();
  }

  function setWatchDir(nextDir) {
    const resolved = path.resolve(String(nextDir || '').trim());
    try {
      if (!resolved || !fs.existsSync(resolved) || !fs.statSync(resolved).isDirectory()) {
        return false;
      }
    } catch {
      return false;
    }
    watchDir = resolved;
    startWatcher();
    return true;
  }

  startWatcher();

  server.middlewares.use(async (req, res, next) => {
    const reqUrl = new URL(req.url || '/', 'http://localhost');

    if (reqUrl.pathname === '/api/watched-workbooks') {
      sendJson(res, folderPayload());
      return;
    }

    if (reqUrl.pathname === '/api/watched-workbooks/folder' && req.method === 'POST') {
      try {
        const body = await readJsonBody(req);
        if (!setWatchDir(body.path)) {
          sendJson(res, { error: 'Folder not found or not a directory', folder: watchDir }, 400);
          return;
        }
        sendJson(res, folderPayload());
      } catch (error) {
        sendJson(res, { error: 'Invalid folder request' }, 400);
      }
      return;
    }

    if (reqUrl.pathname === '/api/watched-workbooks/select-folder' && req.method === 'POST') {
      try {
        const selectedPath = await openFolderPicker();
        if (!selectedPath) {
          sendJson(res, { cancelled: true, ...folderPayload() });
          return;
        }
        if (!setWatchDir(selectedPath)) {
          sendJson(res, { error: 'Selected folder is unavailable', folder: watchDir }, 400);
          return;
        }
        sendJson(res, folderPayload());
      } catch (error) {
        sendJson(res, { error: 'Folder picker unavailable', detail: error.message }, 500);
      }
      return;
    }

    if (reqUrl.pathname === '/api/watched-workbooks/events') {
      res.writeHead(200, {
        'Content-Type': 'text/event-stream; charset=utf-8',
        'Cache-Control': 'no-cache, no-transform',
        Connection: 'keep-alive'
      });
      res.write(`event: workbooks\n`);
      res.write(`data: ${JSON.stringify(folderPayload())}\n\n`);
      clients.add(res);
      req.on('close', () => clients.delete(res));
      return;
    }

    if (reqUrl.pathname.startsWith('/api/watched-workbooks/file/')) {
      const fileName = path.basename(decodeURIComponent(reqUrl.pathname.replace('/api/watched-workbooks/file/', '')));
      const filePath = path.resolve(watchDir, fileName);
      if (!isInsideWatchDir(filePath) || !isWorkbook(fileName) || !fs.existsSync(filePath)) {
        sendJson(res, { error: 'Workbook not found' }, 404);
        return;
      }
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      const stream = fs.createReadStream(filePath);
      stream.on('error', () => {
        if (!res.headersSent) sendJson(res, { error: 'Workbook read failed' }, 500);
        else res.destroy();
      });
      stream.pipe(res);
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
