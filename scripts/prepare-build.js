/**
 * 打包前准备：把前端构建产物复制到项目根 dist/
 * （electron-builder 以项目根为基准打包 files，主进程按 __dirname/dist 加载页面）
 */
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const SRC = path.join(ROOT, 'web-ui', 'dist');
const DEST = path.join(ROOT, 'dist');
const BACKEND = path.join(ROOT, 'backend-server', 'server', 'server.exe');

function fail(msg) {
  console.error(`[prepare] ❌ ${msg}`);
  process.exit(1);
}

if (!fs.existsSync(path.join(SRC, 'index.html'))) {
  fail('前端未构建，请先执行 npm run build:frontend');
}
if (!fs.existsSync(BACKEND)) {
  fail('后端引擎未打包，请先执行 python build_backend.py');
}

fs.rmSync(DEST, { recursive: true, force: true });
fs.cpSync(SRC, DEST, { recursive: true });

const sizeMB = (dir) => {
  let total = 0;
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, entry.name);
    total += entry.isDirectory() ? sizeMB(p) : fs.statSync(p).size;
  }
  return total;
};

console.log(`[prepare] ✅ 前端产物已就位 dist/ (${(sizeMB(DEST) / 1024 / 1024).toFixed(1)} MB)`);
console.log(`[prepare] ✅ 后端引擎 backend-server/server/server.exe (${(fs.statSync(BACKEND).size / 1024 / 1024).toFixed(1)} MB)`);
