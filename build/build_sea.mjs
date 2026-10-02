'use strict';
/*
 * SEA 单文件打包：node src/server.js + web/ 全部静态资源 → build/StampTool.exe
 * 流程：
 *  1. 递归扫描 web/ → 生成 sea-config.json 的 assets 键值映射（键=运行时 sea.getAsset 名，值=磁盘路径）
 *  2. node --experimental-sea-config 生成 blob
 *  3. 复制 node.exe → build/StampTool.exe
 *  4. postject 注入 blob（sentinel fuse 打开 SEA 开关）
 * 用法：node build/build_sea.mjs
 */
import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const BUILD_DIR = path.join(ROOT, 'build');
const WEB_DIR = path.join(ROOT, 'web');
const NODE_EXE = process.execPath;
const SEA_CONFIG = path.join(BUILD_DIR, 'sea-config.json');
const SEA_BLOB = path.join(BUILD_DIR, 'sea-prep.blob');
const OUT_EXE = path.join(BUILD_DIR, 'StampTool.exe');
const SENTINEL_FUSE = 'NODE_SEA_FUSE_fce680ab2cc467b6e072b8b5df1996b2';

// 递归扫描 web/，返回 { "web/<相对路径(正斜杠)>": "<绝对路径>" }
function walkAssets(dir, base, out) {
  for (const f of fs.readdirSync(dir, { withFileTypes: true })) {
    const abs = path.join(dir, f.name);
    const rel = base ? base + '/' + f.name : f.name;   // 资源键一律用正斜杠，与 sea.getAsset('web/xxx') 对应
    if (f.isDirectory()) walkAssets(abs, rel, out);
    else out['web/' + rel] = abs;
  }
  return out;
}

console.log('[1/4] 生成 sea-config.json（assets 映射）…');
const assets = walkAssets(WEB_DIR, '', {});
fs.mkdirSync(BUILD_DIR, { recursive: true });
fs.writeFileSync(SEA_CONFIG, JSON.stringify({
  main: 'src/server.js',
  output: 'build/sea-prep.blob',
  disableExperimentalSEAWarning: true,
  useSnapshot: false,
  useCodeCache: false,
  assets,
}, null, 2), 'utf8');
console.log(`      共 ${Object.keys(assets).length} 个静态资源`);

console.log('[2/4] 生成 SEA blob …');
execFileSync(NODE_EXE, ['--experimental-sea-config', SEA_CONFIG], { cwd: ROOT, stdio: 'inherit' });

console.log('[3/4] 复制 node.exe → StampTool.exe …');
fs.rmSync(OUT_EXE, { force: true });
fs.copyFileSync(NODE_EXE, OUT_EXE);

console.log('[4/4] postject 注入 …');
// 不走 .cmd shim（execFileSync 无法直接 spawn .cmd），直接用 node 调 postject CLI
const postjectCli = path.join(ROOT, 'node_modules', 'postject', 'dist', 'cli.js');
execFileSync(NODE_EXE,
  [postjectCli, OUT_EXE, 'NODE_SEA_BLOB', SEA_BLOB, '--sentinel-fuse', SENTINEL_FUSE],
  { cwd: ROOT, stdio: 'inherit' });

// 可选 UPX 压缩：build/tools/ 下存在 upx.exe 时自动执行（86.9MB → ~22.4MB）。
// 已验证 UPX 压缩后的 SEA exe 功能完好；注意部分杀软可能对 UPX 壳误报，加白即可。
let upx = null;
for (const dir of fs.readdirSync(BUILD_DIR, { withFileTypes: true })) {
  if (dir.isDirectory() && dir.name.startsWith('upx-')) {
    const cand = path.join(BUILD_DIR, dir.name, 'upx.exe');
    if (fs.existsSync(cand)) { upx = cand; break; }
  }
}
if (upx) {
  console.log('[+UPX] 压缩 …');
  execFileSync(upx, ['--best', '--lzma', '-q', OUT_EXE], { cwd: ROOT, stdio: 'inherit' });
} else {
  console.log('[跳过 UPX] 未找到 build/tools/upx-*/upx.exe，输出未压缩 exe');
}

const mb = (fs.statSync(OUT_EXE).size / 1024 / 1024).toFixed(1);
console.log(`\n完成：build/StampTool.exe（${mb} MB）`);
console.log('双击运行 → 自动开浏览器 http://127.0.0.1:8731/；测试可加 --no-open');
