// 从微信读书页面 HTML 中求值 window.__NUXT__，输出完整 JSON。
// 用法: node nuxt_extract.js <html_file>
//
// 安全说明（重要）：
//   vm 不是安全沙箱，只能用来隔离「受控输入」，不能用来执行来路不明的代码。
//   这里的输入必须是调用方自己从 weread.qq.com 下载并落盘的 HTML，
//   且调用方（fetch_booklist.py）已校验只用固定域名抓取。
//   下方另加两道限制：拒收含进程/模块加载的赋值语句；执行限 8 秒；禁用动态代码生成。
'use strict';
const fs = require('fs');
const vm = require('vm');

const htmlPath = process.argv[2];
if (!htmlPath) {
  console.error('USAGE: node nuxt_extract.js <html_file>');
  process.exit(64);
}

const html = fs.readFileSync(htmlPath, 'utf8');
const idx = html.indexOf('window.__NUXT__=');
if (idx < 0) {
  console.error('NO_NUXT');
  process.exit(1);
}
let end = html.indexOf('</script>', idx);
if (end < 0) end = html.length;
let code = html.slice(idx, end).trim();
if (code.endsWith(';')) code = code.slice(0, -1);

// 正常产物只是普通对象字面量赋值。出现下面这些记号说明不是预期形态，直接拒收。
const FORBIDDEN = [/\brequire\s*\(/, /\bimport\s*\(/, /\bFunction\s*\(/, /\beval\s*\(/];
for (const rule of FORBIDDEN) {
  if (rule.test(code)) {
    console.error('REJECTED_SUSPICIOUS_ASSIGNMENT');
    process.exit(3);
  }
}

// 沙箱只给页面脚本可能碰到的最小对象；不含 require / process / fs / 网络。
const sandbox = Object.create(null);
sandbox.window = Object.create(null);
sandbox.console = { log() {}, warn() {}, error() {} };
sandbox.navigator = { userAgent: 'node' };
vm.createContext(sandbox, { codeGeneration: { strings: false, wasm: false } });

try {
  vm.runInContext(code, sandbox, { timeout: 8000, displayErrors: true });
} catch (e) {
  console.error('EVAL_ERR:' + e.message);
  process.exit(2);
}

const nuxt = sandbox.window.__NUXT__;
if (!nuxt || typeof nuxt !== 'object') {
  console.error('NO_NUXT_VALUE');
  process.exit(4);
}
process.stdout.write(JSON.stringify(nuxt));
