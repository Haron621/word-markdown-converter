import { createServer } from 'node:http';
import { readFile, stat } from 'node:fs/promises';
import { extname, join, normalize } from 'node:path';

const host = '127.0.0.1';
const port = 3000;
const root = join(import.meta.dirname, 'dist');

const types = {
  '.css': 'text/css; charset=utf-8',
  '.html': 'text/html; charset=utf-8',
  '.ico': 'image/x-icon',
  '.js': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.svg': 'image/svg+xml',
  '.webp': 'image/webp',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
};

createServer(async (request, response) => {
  try {
    const url = new URL(request.url ?? '/', `http://${host}:${port}`);
    const relative = decodeURIComponent(url.pathname).replace(/^\/+/, '');
    let filePath = normalize(join(root, relative || 'index.html'));
    if (!filePath.startsWith(root)) throw new Error('Invalid path');

    const fileStat = await stat(filePath).catch(() => null);
    if (fileStat?.isDirectory()) filePath = join(filePath, 'index.html');

    const body = await readFile(filePath);
    response.writeHead(200, {
      'Cache-Control': 'no-store',
      'Content-Type': types[extname(filePath).toLowerCase()] ?? 'application/octet-stream',
    });
    response.end(body);
  } catch {
    response.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
    response.end('Файл не найден');
  }
}).listen(port, host, () => {
  console.log(`Word to Markdown запущен: http://${host}:${port}/ru/`);
  console.log('Чтобы остановить сервис, закройте это окно или нажмите Ctrl+C.');
});
