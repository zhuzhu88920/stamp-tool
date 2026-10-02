'use strict';
/*
 * 文件盖章 · 本地版 —— 零依赖 Node 服务
 * - 托管前端（web/ 目录，SEA 模式读嵌入资源）
 * - /api/seals        列出印章目录下的 PNG（{ok,data:[{key,name}]}，与云端同构）
 * - /api/seal?key=    返回印章图片字节（对应云端 /api/r2-file）
 * - /api/seal-upload  原始字节 + X-Filename 头，写入印章目录（对应云端 /api/seal-upload）
 * - /api/seal-del     删除印章（对应云端 /api/r2-del）
 * - /api/stampdir     GET 读 / POST 设置印章目录；持久化到程序旁 StampTool.json，下次免设
 * 首次运行自动生成 StampTool.json 与默认 stamps/ 目录；双击启动后自动打开浏览器。
 */
const http = require('http');
const fs = require('fs');
const path = require('path');
const { exec } = require('child_process');

let sea = null;
let IS_SEA = false;
try {
  sea = require('node:sea');
  IS_SEA = typeof sea.isSea === 'function' ? sea.isSea() : false;
} catch (e) { IS_SEA = false; }

const PROJECT_ROOT = __dirname.replace(/[\\/]src$/, '');
const WEB_DIR = path.join(PROJECT_ROOT, 'web');
const CONFIG_DIR = IS_SEA ? path.dirname(process.execPath) : PROJECT_ROOT;
const CONFIG_FILE = path.join(CONFIG_DIR, 'StampTool.json');
const DEFAULT_STAMPS_DIR = path.join(CONFIG_DIR, 'stamps');

const SEAL_EXT = new Set(['.png', '.jpg', '.jpeg']);
const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.mjs': 'application/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
  '.bcmap': 'application/octet-stream',
  '.pfb': 'application/octet-stream',
  '.ttf': 'font/ttf',
};

// ---------- 配置（首次运行自动生成） ----------
function loadConfig() {
  try {
    if (fs.existsSync(CONFIG_FILE)) {
      const c = JSON.parse(fs.readFileSync(CONFIG_FILE, 'utf8'));
      if (c && typeof c === 'object') return c;
    }
  } catch (e) {}
  const fresh = { stampDir: '' };
  try { fs.writeFileSync(CONFIG_FILE, JSON.stringify(fresh, null, 2), 'utf8'); } catch (e) {}
  return fresh;
}

function saveConfig(cfg) {
  try { fs.writeFileSync(CONFIG_FILE, JSON.stringify(cfg, null, 2), 'utf8'); }
  catch (e) { console.error('保存 StampTool.json 失败:', e.message); }
}

function ensureDefaultDir() {
  try { if (!fs.existsSync(DEFAULT_STAMPS_DIR)) fs.mkdirSync(DEFAULT_STAMPS_DIR, { recursive: true }); }
  catch (e) {}
}

function resolveStampDir(cfg) {
  const raw = (cfg.stampDir || '').trim();
  if (raw) {
    // path.resolve 兼容两种存法：相对路径（相对 exe 目录）→ 绝对；绝对路径 → 原样。
    // 再 resolve 归一化消除正反斜杠差异（用户配置 "d:/seals" → "d:\seals"）
    try {
      const n = path.resolve(CONFIG_DIR, raw);
      if (fs.existsSync(n) && fs.statSync(n).isDirectory()) return n;
    } catch (e) {}
  }
  ensureDefaultDir();
  return DEFAULT_STAMPS_DIR;
}

// 绝对路径 → json 存储值：在 exe 目录内则存相对路径（正斜杠，绿色软件搬家不失效），
// 目录外退化为绝对路径
function toStorableDir(absDir) {
  if (!absDir) return '';
  const rel = path.relative(CONFIG_DIR, absDir);
  if (!rel || rel.startsWith('..') || path.isAbsolute(rel)) return absDir;
  return rel.replace(/\\/g, '/');
}

// 校验拼出的路径仍在印章目录内（须先 resolve 归一化，否则 Windows 下
// path.join 产生的反斜杠与配置中的正斜杠 startsWith 比较恒假）
function inDir(dir, name) {
  const fp = path.resolve(dir, name);
  return fp !== dir && fp.startsWith(dir + path.sep) ? fp : null;
}

function listStamps(dir) {
  try {
    return fs.readdirSync(dir)
      .filter((f) => SEAL_EXT.has(path.extname(f).toLowerCase()))
      .sort((a, b) => a.localeCompare(b, 'zh'))
      .map((f) => ({ key: f, name: f.replace(/\.(png|jpe?g)$/i, '') }));
  } catch (e) { return []; }
}

function safeName(name) {
  const base = path.basename(String(name || ''));
  if (!base || base === '.' || base === '..') return null;
  if (!SEAL_EXT.has(path.extname(base).toLowerCase())) return null;
  return base;
}

// ---------- 静态资源：dev 磁盘 / SEA 嵌入 ----------
function readAssetBytes(rel) {
  if (IS_SEA) {
    try {
      const a = sea.getAsset('web/' + rel);
      if (a == null) return null;
      return Buffer.isBuffer(a) ? a : Buffer.from(a);   // SEA 返回 ArrayBuffer，必须转 Buffer
    } catch (e) { return null; }
  }
  try {
    const p = path.join(WEB_DIR, rel);
    if (!p.startsWith(WEB_DIR)) return null;
    return fs.readFileSync(p);
  } catch (e) { return null; }
}

// ---------- 响应工具（防重复发送/客户端断开导致进程崩溃） ----------
function sendJson(res, code, obj) {
  if (res.writableEnded || res.headersSent) return;
  try {
    res.writeHead(code, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(Buffer.from(JSON.stringify(obj), 'utf8'));
  } catch (e) {}
}
function sendBytes(res, code, buf, mime) {
  if (res.writableEnded || res.headersSent) return;
  try {
    res.writeHead(code, { 'Content-Type': mime || 'application/octet-stream' });
    res.end(buf);
  } catch (e) {}
}

function readBody(req, capMB) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let size = 0;
    req.on('data', (c) => {
      size += c.length;
      if (size > (capMB || 20) * 1024 * 1024) { reject(new Error('文件过大')); req.destroy(); return; }
      chunks.push(c);
    });
    req.on('end', () => resolve(Buffer.concat(chunks)));
    req.on('error', reject);
  });
}

// ---------- 浏览器存活跟踪（SSE）：全部页面关闭后自动退出服务 ----------
// 页面开一条 EventSource 长连接；连接数归零且曾有过连接、持续 8s → process.exit。
// 用 SSE 而非定时心跳：后台标签页的 setInterval 会被浏览器节流到分钟级，会误判；
// 长连接不受节流影响，页面关闭时 TCP 立即断开。8s 宽限容忍刷新时的重连间隙。
const sseClients = new Set();
let hadClient = false;
let exitTimer = null;

function armExitCheck() {
  if (exitTimer) return;
  exitTimer = setTimeout(() => {
    exitTimer = null;
    if (hadClient && sseClients.size === 0) {
      console.log('浏览器已全部关闭，服务自动退出');
      process.exit(0);
    }
  }, 8000);
}

function disarmExitCheck() {
  if (exitTimer) { clearTimeout(exitTimer); exitTimer = null; }
}

function serveEvents(req, res) {
  if (res.writableEnded || res.headersSent) return;
  try {
    res.writeHead(200, {
      'Content-Type': 'text/event-stream; charset=utf-8',
      'Cache-Control': 'no-cache',
      'Connection': 'keep-alive',
    });
    res.write(':ok\n\n');
  } catch (e) { return; }
  sseClients.add(res);
  hadClient = true;
  disarmExitCheck();
  const keep = setInterval(() => { try { res.write(':ping\n\n'); } catch (e) {} }, 25000);
  res.on('close', () => {
    clearInterval(keep);
    sseClients.delete(res);
    if (sseClients.size === 0 && hadClient) armExitCheck();
  });
}

// ---------- API ----------
async function handleApi(req, res, url) {
  const p = url.pathname;

  if (p === '/api/events') return serveEvents(req, res);

  if (p === '/api/health') return sendJson(res, 200, { ok: true, data: { sea: IS_SEA } });

  if (p === '/api/seals' && req.method === 'GET') {
    const cfg = loadConfig();
    return sendJson(res, 200, { ok: true, data: listStamps(resolveStampDir(cfg)) });
  }

  if (p === '/api/seal' && req.method === 'GET') {
    const name = safeName(url.searchParams.get('key'));
    if (!name) return sendJson(res, 400, { ok: false, error: '文件名不合法' });
    const dir = resolveStampDir(loadConfig());
    const fp = inDir(dir, name);
    if (!fp || !fs.existsSync(fp)) return sendJson(res, 404, { ok: false, error: '找不到印章' });
    return sendBytes(res, 200, fs.readFileSync(fp), MIME[path.extname(name).toLowerCase()] || 'application/octet-stream');
  }

  if (p === '/api/seal-upload' && req.method === 'POST') {
    const name = safeName(decodeURIComponent(req.headers['x-filename'] || ''));
    if (!name) return sendJson(res, 400, { ok: false, error: '仅支持 PNG / JPG 文件' });
    const buf = await readBody(req, 20);
    if (!buf.length) return sendJson(res, 400, { ok: false, error: '空文件' });
    const dir = resolveStampDir(loadConfig());
    fs.writeFileSync(path.join(dir, name), buf);
    return sendJson(res, 200, { ok: true, data: { dir } });
  }

  if (p === '/api/seal-del' && req.method === 'POST') {
    const body = JSON.parse((await readBody(req, 1)).toString('utf8') || '{}');
    const name = safeName(body.key);
    if (!name) return sendJson(res, 400, { ok: false, error: '文件名不合法' });
    const dir = resolveStampDir(loadConfig());
    const fp = inDir(dir, name);
    if (!fp || !fs.existsSync(fp)) return sendJson(res, 404, { ok: false, error: '找不到印章' });
    fs.unlinkSync(fp);
    return sendJson(res, 200, { ok: true, data: {} });
  }

  if (p === '/api/stampdir' && req.method === 'GET') {
    const cfg = loadConfig();
    return sendJson(res, 200, { ok: true, data: { dir: resolveStampDir(cfg) } });
  }

  if (p === '/api/stampdir' && req.method === 'POST') {
    const body = JSON.parse((await readBody(req, 1)).toString('utf8') || '{}');
    const dir = (body.dir || '').trim();
    let abs = '';
    if (dir) {
      abs = path.resolve(dir);
      if (!fs.existsSync(abs) || !fs.statSync(abs).isDirectory()) {
        return sendJson(res, 400, { ok: false, error: '目录不存在或不是文件夹：' + dir });
      }
    }
    const cfg = loadConfig();
    cfg.stampDir = toStorableDir(abs);   // exe 目录内存相对路径，外部存绝对
    saveConfig(cfg);
    return sendJson(res, 200, { ok: true, data: { dir: resolveStampDir(cfg) } });
  }

  return sendJson(res, 404, { ok: false, error: '未找到接口' });
}

function serveStatic(req, res, url) {
  let rel = decodeURIComponent(url.pathname);
  if (rel === '/' || rel === '') rel = '/index.html';
  rel = rel.replace(/^\/+/, '');
  const data = readAssetBytes(rel);
  if (data == null) return sendBytes(res, 404, Buffer.from('Not Found'), 'text/plain; charset=utf-8');
  return sendBytes(res, 200, data, MIME[path.extname(rel).toLowerCase()] || 'application/octet-stream');
}

const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url, 'http://127.0.0.1');
    if (url.pathname.startsWith('/api/')) await handleApi(req, res, url);
    else serveStatic(req, res, url);
  } catch (e) {
    if (!res.headersSent && !res.writableEnded) {
      try { sendJson(res, 500, { ok: false, error: e.message }); } catch (_) {}
    }
  }
});

function findPort(base, cb) {
  let port = base;
  const tryOnce = () => {
    const s = http.createServer();
    s.once('error', () => { s.close(); port++; if (port - base < 50) tryOnce(); else cb(null); });
    s.once('listening', () => { s.close(); cb(port); });
    s.listen(port, '127.0.0.1');
  };
  tryOnce();
}

loadConfig();          // 首次运行自动生成 StampTool.json
ensureDefaultDir();    // 首次运行自动建默认印章目录

const NO_OPEN = process.argv.includes('--no-open');
findPort(8731, (port) => {
  if (!port) { console.error('无法找到可用端口'); process.exit(1); }
  server.listen(port, '127.0.0.1', () => {
    const url = 'http://127.0.0.1:' + port + '/';
    console.log('文件盖章 已启动: ' + url);
    if (!NO_OPEN) {
      const cmd = process.platform === 'win32'
        ? 'cmd /c start "" "' + url + '"'
        : (process.platform === 'darwin' ? 'open "' + url + '"' : 'xdg-open "' + url + '"');
      exec(cmd, () => {});
    }
  });
});
