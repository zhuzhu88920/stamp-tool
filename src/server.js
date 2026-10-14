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
const os = require('os');
const path = require('path');
const { exec, execFile, execFileSync, spawn } = require('child_process');

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

// ---------- Office 能力检测 + 常驻转换宿主（COM 实探权威 + 实例复用） ----------
// 设计目标：避免每次转换都冷启动 Excel/Word（含 Quit 约 4s 开销），对标线上常驻实例。
// 方案：启动时拉起一个常驻 PowerShell 进程，初始化并保活 Word/Excel COM 实例；
//       /api/convert 通过 stdin 向该进程派发任务、stdout 回收结果。
//       COM 单线程 → 任务串行；宿主崩溃自动重启一次，失败则回退到「每次新开进程」的旧方案。
//       卸载残留骗不过 COM 实探，能力判定权威。
const OFFICE_CAP = { word: false, excel: false, checked: false };

// 常驻宿主脚本（无反引号、无 ${}，可安全放进 JS 模板字符串）。
// 初始化 Word/Excel → 打印 READY<tab><word><tab><excel> → 循环读 stdin 任务行。
// 任务行（制表符分隔，路径均为 ASCII 临时名）：  <id><tab><src><tab><dst><tab><kind>
// 结果行：  <id><tab>OK   或   <id><tab>ERR<tab><msg>
const HOST_PS = `
$ErrorActionPreference='Stop'
$word=$null; $excel=$null
function GetWord { if(-not $word){ $word=New-Object -ComObject Word.Application; $word.Visible=$false; $word.DisplayAlerts=0 }; $word }
function GetExcel { if(-not $excel){ $excel=New-Object -ComObject Excel.Application; $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.AskToUpdateLinks=$false; $excel.AlertBeforeOverwriting=$false }; $excel }
$wok=$false; $eok=$false
try { $w=GetWord; if($w.Version){$wok=$true} } catch { $word=$null }
try { $e=GetExcel; if($e.Version){$eok=$true} } catch { $excel=$null }
[Console]::Out.Write('READY' + [char]9 + $wok.ToString().ToLower() + [char]9 + $eok.ToString().ToLower() + [char]10); [Console]::Out.Flush()
while($true){
  $line=[Console]::In.ReadLine()
  if($null -eq $line){ break }
  if($line -eq ''){ continue }
  $parts=$line -split [char]9
  $id=$parts[0]; $src=$parts[1]; $dst=$parts[2]; $kind=$parts[3]
  try {
    if($kind -eq 'word'){
      $app=GetWord
      $doc=$app.Documents.Open($src,$false,$true)
      # ExportAsFixedFormat(...,Item=0=wdExportDocumentContent) 仅导出正文，排除批注/修订等标记
      try { $doc.ExportAsFixedFormat($dst,17,$false,0,0,'','',0) } finally { $doc.Close($false) }
    } else {
      $app=GetExcel
      $wb=$app.Workbooks.Open($src,0)
      try {
        # 若任意表设了打印区域，仅导出这些表（隐藏其余表，Workbook.ExportAsFixedFormat 自动跳过隐藏表）；
        # 一张都没设才整本导出（各表导已用区域）。打印区域是 per-sheet 的，整本导出不会自动裁剪未设区域的表。
        $hasAnyPA=$false
        foreach($ws in $wb.Worksheets){ if($ws.PageSetup.PrintArea -and $ws.PageSetup.PrintArea -ne ''){ $hasAnyPA=$true; break } }
        if($hasAnyPA){
          foreach($ws in $wb.Worksheets){
            $paSet=($ws.PageSetup.PrintArea -and $ws.PageSetup.PrintArea -ne '')
            if(-not $paSet){ $ws.Visible=0 }   # xlSheetHidden
          }
        }
        # IgnorePrintAreas=0(False)：尊重各表已设打印区域
        $wb.ExportAsFixedFormat(0,$dst,0,0,0)
      } finally { $wb.Close($false) }
    }
    [Console]::Out.Write($id + [char]9 + 'OK' + [char]10); [Console]::Out.Flush()
  } catch {
    [Console]::Out.Write($id + [char]9 + 'ERR' + [char]9 + ((($_.Exception.Message -split [char]10)[0]).Replace([char]9,' ')) + [char]10); [Console]::Out.Flush()
  }
}
if($word){ try{ $word.Quit() }catch{} }
if($excel){ try{ $excel.Quit() }catch{} }
`;

// 旧方案兜底：每次新开 powershell 进程转换（宿主不可用时）。Word: SaveAs(17=wdFormatPDF)；
// Excel: ExportAsFixedFormat(0=xlTypePDF)。临时目录 guid 命名，路径经环境变量传入，规避中文/空格编码。
const CONVERT_PS = `
$ErrorActionPreference='Stop'
$src=$env:ST_SRC; $dst=$env:ST_DST; $kind=$env:ST_KIND
$app = if($kind -eq 'word'){ New-Object -ComObject Word.Application } else { New-Object -ComObject Excel.Application }
try {
  if($kind -eq 'word'){
    $app.Visible=$false
    $app.DisplayAlerts=0
    $doc=$app.Documents.Open($src,$false,$true)
    # ExportAsFixedFormat(...,Item=0) 仅导出正文，排除批注/修订等标记
    $doc.ExportAsFixedFormat($dst,17,$false,0,0,'','',0)
    $doc.Close($false)
  } else {
    $app.Visible=$false
    $app.DisplayAlerts=$false
    $app.AskToUpdateLinks=$false
    $app.AlertBeforeOverwriting=$false
    $wb=$app.Workbooks.Open($src,0)
    try {
      # 若任意表设了打印区域，仅导出这些表（隐藏其余表，导出时自动跳过）；全没设才整本导出
      $hasAnyPA=$false
      foreach($ws in $wb.Worksheets){ if($ws.PageSetup.PrintArea -and $ws.PageSetup.PrintArea -ne ''){ $hasAnyPA=$true; break } }
      if($hasAnyPA){
        foreach($ws in $wb.Worksheets){
          $paSet=($ws.PageSetup.PrintArea -and $ws.PageSetup.PrintArea -ne '')
          if(-not $paSet){ $ws.Visible=0 }   # xlSheetHidden
        }
      }
      # IgnorePrintAreas=0(False)：尊重各表已设打印区域
      $wb.ExportAsFixedFormat(0,$dst,0,0,0)
    } finally { if($wb){ try{ $wb.Close($false) }catch{} } }
  }
} finally { try { $app.Quit() } catch {} }
`;

let hostProc = null;
let hostReady = false;
const jobMap = new Map();
let jobSeq = 0;
let jobQueue = Promise.resolve();
let hostStdoutBuf = '';
let shuttingDown = false;

function hostEncodeCmd() {
  const b64 = Buffer.from(HOST_PS, 'utf16le').toString('base64');
  return ['-NoProfile', '-STA', '-EncodedCommand', b64];
}

function spawnHost() {
  return new Promise((resolve) => {
    let settled = false;
    const finish = (ok) => { if (!settled) { settled = true; resolve(ok); } };
    let proc;
    try {
      proc = spawn('powershell.exe', hostEncodeCmd(), { windowsHide: true, stdio: ['pipe', 'pipe', 'pipe'] });
    } catch (e) { return finish(false); }
    hostProc = proc;
    hostReady = false;
    hostStdoutBuf = '';

    const processLine = (line) => {
      if (line.startsWith('READY')) {
        const parts = line.split('\t');
        OFFICE_CAP.word = parts[1] === 'true';
        OFFICE_CAP.excel = parts[2] === 'true';
        OFFICE_CAP.checked = true;
        hostReady = true;
        if (OFFICE_CAP.word || OFFICE_CAP.excel) finish(true);
        else { try { proc.kill(); } catch (e) {} hostProc = null; finish(false); }
        return;
      }
      const parts = line.split('\t');
      if (parts.length < 2) return;
      const id = parseInt(parts[0], 10);
      const job = jobMap.get(id);
      if (!job) return;
      clearTimeout(job.timer);
      jobMap.delete(id);
      if (parts[1] === 'OK') job.resolve();
      else job.reject(new Error(parts.slice(2).join('\t') || '转换失败'));
    };

    proc.stdout.on('data', (d) => {
      hostStdoutBuf += d.toString('utf8');
      let idx;
      while ((idx = hostStdoutBuf.indexOf('\n')) >= 0) {
        const line = hostStdoutBuf.slice(0, idx).replace(/\r$/, '');
        hostStdoutBuf = hostStdoutBuf.slice(idx + 1);
        processLine(line);
      }
    });
    proc.stderr.on('data', () => {});
    proc.on('error', () => finish(false));
    proc.on('exit', () => {
      hostProc = null;
      hostReady = false;
      for (const [, j] of jobMap) { clearTimeout(j.timer); j.reject(new Error('转换宿主已退出')); }
      jobMap.clear();
      finish(false);
      if ((OFFICE_CAP.word || OFFICE_CAP.excel) && !shuttingDown) spawnHost();
    });
    setTimeout(() => finish(false), 15000);
  });
}

function fallbackConvert(kind, src, dst) {
  return new Promise((resolve, reject) => {
    execFile('powershell.exe', ['-NoProfile', '-STA', '-Command', CONVERT_PS], {
      env: Object.assign({}, process.env, { ST_SRC: src, ST_DST: dst, ST_KIND: kind }),
      timeout: 180000, windowsHide: true,
    }, (err) => {
      if (err) return reject(err);
      if (!fs.existsSync(dst)) return reject(new Error('转换未生成 PDF'));
      resolve();
    });
  });
}

// ---------- Office 文档转 PDF（常驻宿主优先，失败回退每次新开进程） ----------
function convertOffice(buf, kind, srcExt) {
  const task = () => new Promise((resolve, reject) => {
    let dir;
    let src, dst;   // 提到 executor 作用域，供异步回调 finish 访问（try 块内 const 外部不可见）
    const finish = (err) => {
      if (err) { try { fs.rmSync(dir, { recursive: true, force: true }); } catch (_) {} return reject(err); }
      try {
        if (!fs.existsSync(dst)) throw new Error('转换未生成 PDF');
        resolve(fs.readFileSync(dst));
      } catch (e) { try { fs.rmSync(dir, { recursive: true, force: true }); } catch (_) {} reject(e); }
      finally { try { fs.rmSync(dir, { recursive: true, force: true }); } catch (_) {} }
    };
    try {
      dir = fs.mkdtempSync(path.join(os.tmpdir(), 'stamptool-'));
      // 必须保留原始扩展名：老格式(.doc/.xls)内容存成 .docx/.xlsx 时，
      // Word/Excel 会判定「文件格式与文件扩展名不匹配」而拒绝打开（因其按扩展名选解析器）。
      src = path.join(dir, 'in' + (srcExt || (kind === 'word' ? '.docx' : '.xlsx')));
      dst = path.join(dir, 'out.pdf');
      fs.writeFileSync(src, buf);

      // 无常驻宿主 → 直接回退兜底方案
      if (!(hostProc && hostReady)) {
        return fallbackConvert(kind, src, dst).then(() => finish(null), finish);
      }
      // 派发任务给常驻宿主（stdin 写一行，stdout 回一行）
      const id = ++jobSeq;
      const timer = setTimeout(() => {
        jobMap.delete(id);
        finish(new Error('Office 转换超时（>170s）'));
      }, 170000);
      jobMap.set(id, {
        resolve: () => { clearTimeout(timer); finish(null); },
        reject: (e) => { clearTimeout(timer); finish(e); },
        timer,
      });
      try {
        hostProc.stdin.write(id + '\t' + src + '\t' + dst + '\t' + kind + '\n');
      } catch (e) {
        clearTimeout(timer); jobMap.delete(id);
        fallbackConvert(kind, src, dst).then(() => finish(null), finish);
      }
    } catch (e) {
      if (dir) { try { fs.rmSync(dir, { recursive: true, force: true }); } catch (_) {} }
      reject(e);
    }
  });
  // 串行锁：COM 不可并发，同一时刻只允许一个转换
  const run = jobQueue.then(task, task);
  jobQueue = run.then(() => {}, () => {});
  return run;
}

function startConverter() {
  return spawnHost().then((ok) => {
    if (!ok) { OFFICE_CAP.word = false; OFFICE_CAP.excel = false; OFFICE_CAP.checked = true; }
    return OFFICE_CAP;
  });
}

process.on('exit', () => {
  shuttingDown = true;
  try { if (hostProc) hostProc.kill(); } catch (e) {}
});

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

  if (p === '/api/health') return sendJson(res, 200, { ok: true, data: { sea: IS_SEA, office: OFFICE_CAP.word || OFFICE_CAP.excel, word: OFFICE_CAP.word, excel: OFFICE_CAP.excel } });

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

  // ---------- Office 文档转 PDF（本地直调 MS Office COM） ----------
  if (p === '/api/convert' && req.method === 'POST') {
    const rawName = decodeURIComponent(req.headers['x-filename'] || '');
    const ext = (rawName.split('.').pop() || '').toLowerCase();
    const isWord = ext === 'doc' || ext === 'docx';
    const isExcel = ext === 'xls' || ext === 'xlsx';
    if (!isWord && !isExcel) return sendJson(res, 400, { ok: false, error: '仅支持 Word / Excel 文档（.doc / .docx / .xls / .xlsx）' });
    if (isWord && !OFFICE_CAP.word) return sendJson(res, 400, { ok: false, error: '本地无 MS Word，不支持文档转 PDF 直出，仅支持 PDF 盖章' });
    if (isExcel && !OFFICE_CAP.excel) return sendJson(res, 400, { ok: false, error: '本地无 MS Excel，不支持文档转 PDF 直出，仅支持 PDF 盖章' });
    let buf;
    try { buf = await readBody(req, 50); }
    catch (e) { return sendJson(res, 413, { ok: false, error: '文件过大（>50MB）' }); }
    if (!buf.length) return sendJson(res, 400, { ok: false, error: '空文件' });
    try {
      const pdf = await convertOffice(buf, isWord ? 'word' : 'excel', '.' + ext);
      return sendBytes(res, 200, pdf, 'application/pdf');
    } catch (e) {
      return sendJson(res, 500, { ok: false, error: 'Office 转换失败：' + (e && e.message ? e.message : '未知错误') });
    }
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

function startServer() {
  const capTxt = (OFFICE_CAP.word ? 'Word' : '') + (OFFICE_CAP.excel ? (OFFICE_CAP.word ? '+Excel' : 'Excel') : (OFFICE_CAP.word || OFFICE_CAP.excel ? '' : '无'));
  findPort(8731, (port) => {
    if (!port) { console.error('无法找到可用端口'); process.exit(1); }
    server.listen(port, '127.0.0.1', () => {
      const url = 'http://127.0.0.1:' + port + '/';
      console.log('文件盖章 已启动: ' + url + '  (Office: ' + capTxt + ')');
      if (!NO_OPEN) {
        const cmd = process.platform === 'win32'
          ? 'cmd /c start "" "' + url + '"'
          : (process.platform === 'darwin' ? 'open "' + url + '"' : 'xdg-open "' + url + '"');
        exec(cmd, () => {});
      }
    });
  });
}

// 先拉起常驻转换宿主并完成 Office 能力检测（COM 实探权威），再起服务，保证页面首屏即拿到正确能力。
startConverter().then(startServer).catch(startServer);
