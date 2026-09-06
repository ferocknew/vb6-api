#!/usr/bin/env node
/**
 * vb6ide-api skill.js — VB6 IDE 插件 REST API 请求工具集
 * 用法：node skill.js <command> [args]
 * 固化请求规则：端口探测、UTF-8、JSON 信封解析、CRLF 规范化、超时与错误定性。
 * 所有写操作的代码体从文件或 stdin 读取（杜绝 shell 转义破坏 body）。
 */
'use strict';
const fs = require('fs');

const PORT_BASE = 8306;
const PORT_MAX_TRIES = 20;
const TIMEOUT_MS = 30000;
const TRUNCATE_DEFAULT = 60000; // 大响应默认截断字符数，--full 取全量

// ---------- 参数解析 ----------
// 通用规则：--key value 成对收（值并入 opts.key）；--cleanup/--full 两个开关型例外收 true
function parseArgs(argv) {
  const cmd = argv[0];
  const rest = argv.slice(1);
  const opts = { _: [] };
  for (let i = 0; i < rest.length; i++) {
    const a = rest[i];
    if (a === '--cleanup' || a === '--full') {
      opts[a.slice(2)] = true;
    } else if (a.startsWith('--')) {
      opts[a.slice(2)] = rest[++i];
    } else opts._.push(a);
  }
  return { cmd, opts };
}

// ---------- 端口探测 ----------
let cachedPort = null;
async function findPort() {
  if (cachedPort) return cachedPort;
  if (process.env.VB6IDE_PORT) {
    cachedPort = Number(process.env.VB6IDE_PORT);
    return cachedPort;
  }
  for (let i = 0; i < PORT_MAX_TRIES; i++) {
    const port = PORT_BASE + i;
    try {
      const r = await req(port, 'GET', '/api/status');
      if (r.envelope && r.envelope.ok) { cachedPort = port; return port; }
    } catch (_) { /* 端口无服务，继续 */ }
  }
  fail('服务不可达：8306 起 ' + PORT_MAX_TRIES + ' 个端口均无应答。请确认 VB6 IDE 已打开、插件已加载且服务已启动（工具栏按钮）。');
}

// ---------- 请求核心（固定格式） ----------
async function req(port, method, path, bodyObj, rawBody) {
  const ctl = new AbortController();
  const timer = setTimeout(() => ctl.abort(), TIMEOUT_MS);
  const headers = {};
  let body;
  if (rawBody !== undefined) {
    body = rawBody; // 已是字符串（代码体），node fetch 按 UTF-8 发送
    headers['Content-Type'] = 'application/json; charset=utf-8';
  } else if (bodyObj !== undefined) {
    // 服务端 BUG-A 绕行（2026-09-06 实测）：body 链路 UTF-8 解码失效（非 ASCII 变 ?），
    // 但 JsonUnescape 的 \uXXXX → ChrW$ 路径正确。故 stringify 后把非 ASCII 全部转义为 \uXXXX
    // （在 JSON 文本层面替换，单反斜杠才是合法 Unicode 转义）。服务端修复后此行可移除。
    body = JSON.stringify(bodyObj).replace(/[-￿]/g, c => '\\u' + c.charCodeAt(0).toString(16).padStart(4, '0'));
    headers['Content-Type'] = 'application/json; charset=utf-8';
  }
  let res;
  try {
    res = await fetch('http://127.0.0.1:' + port + path, { method, headers, body, signal: ctl.signal });
  } catch (e) {
    clearTimeout(timer);
    throw e;
  }
  clearTimeout(timer);
  const text = await res.text();
  let env = null;
  try { env = JSON.parse(text); } catch (_) { /* 非 JSON（如 /docs HTML），透传 */ }
  return { status: res.status, text, envelope: env };
}

async function call(method, path, bodyObj, rawBody) {
  const port = await findPort();
  const r = await req(port, method, path, bodyObj, rawBody);
  if (r.envelope) {
    if (!r.envelope.ok) {
      fail('HTTP ' + r.status + ' ok:false — ' + r.envelope.message);
    }
    return r.envelope;
  }
  if (r.status !== 200) fail('HTTP ' + r.status + ' 非 JSON 响应：' + r.text.slice(0, 200));
  // 200 但非合法 JSON = 服务端信封违规（2026-09-06 实测：bps 空账本返回 "data":"count":0 缺 {），
  // 必须显式报错而非静默透传畸形体
  fail('服务端 200 响应非合法 JSON（信封违规）：' + r.text.slice(0, 200));
}

// ---------- 代码体读取与规范化（规则 1/3 固化） ----------
function readCodeFrom(fileOrDash) {
  let text;
  if (fileOrDash === '-') {
    text = fs.readFileSync(0, 'utf8'); // stdin
  } else {
    if (!fs.existsSync(fileOrDash)) fail('文件不存在：' + fileOrDash);
    text = fs.readFileSync(fileOrDash, 'utf8');
  }
  // CRLF 规范化（LF → CRLF）+ 保证尾换行
  text = text.replace(/\r?\n/g, '\r\n');
  if (text.length && !text.endsWith('\r\n')) text += '\r\n';
  // 规则 1 硬校验：禁止声明头（提前拦截，不依赖服务端剥除）
  const firstLine = text.split('\r\n')[0] || '';
  if (/^\s*(Option\s+Explicit|Attribute\s+)/i.test(firstLine)) {
    fail('写入内容首行是 "' + firstLine.trim() + '"：禁止写入 Option Explicit / Attribute 声明头（规则 1），请从过程主体或正文开始。');
  }
  return text;
}

// ---------- 输出 ----------
function fail(msg) { console.error('ERROR: ' + msg); process.exit(1); }
function out(data, opts) {
  let s = typeof data === 'string' ? data : JSON.stringify(data, null, 2);
  const limit = opts && opts.max ? Number(opts.max) : TRUNCATE_DEFAULT;
  if (!(opts && opts.full) && s.length > limit) {
    s = s.slice(0, limit) + '\n...[截断，共 ' + s.length + ' 字符；--full 查看全量，或 --max N 调整]';
  }
  console.log(s);
}

// ---------- 读 JSON 文件（props/style 用，非代码体：不做 CRLF 归一与声明头拦截） ----------
function readJsonFrom(fileOrDash, what) {
  let text;
  if (fileOrDash === '-') text = fs.readFileSync(0, 'utf8');
  else if (/^\s*\{/.test(fileOrDash)) text = fileOrDash; // 内联 JSON 字面量（如 '{"left":100}'）直接用，不当文件路径
  else {
    if (!fs.existsSync(fileOrDash)) fail('文件不存在：' + fileOrDash);
    text = fs.readFileSync(fileOrDash, 'utf8');
  }
  try { return JSON.parse(text); } catch (e) { fail(what + '不是合法 JSON：' + e.message); }
}

// ---------- 命令实现 ----------
const commands = {
  async status() {
    cachedPort = null;
    const port = await findPort();
    const env = await call('GET', '/api/status');
    console.log('port: ' + port);
    out(env.data);
  },
  async project(o) { out((await call('GET', '/api/project')).data, o); },
  async components(o) { out((await call('GET', '/api/components')).data, o); },
  async snapshot(o) { out((await call('GET', '/api/project/snapshot')).data, o); },

  async 'code-get'(o) {
    const [name] = o._;
    if (!name) fail('用法：code-get <module> [--start N --end M]');
    let q = '';
    if (o.start) q += 'start=' + encodeURIComponent(o.start);
    if (o.end) q += (q ? '&' : '') + 'end=' + encodeURIComponent(o.end);
    out((await call('GET', '/api/modules/' + encodeURIComponent(name) + '/code' + (q ? '?' + q : ''))).data, o);
  },
  async 'module-create'(o) {
    const [name, type] = o._;
    if (!name) fail('用法：module-create <name> [type]  type: module|class|form|mdiform|usercontrol|proppage');
    out((await call('POST', '/api/modules', { name, type: type || 'module' })).data, o);
  },
  async 'code-put'(o) {
    const [name, file] = o._;
    if (!name || !file) fail('用法：code-put <module> <file|-> [--type module|class]   代码体从文件或 stdin(-) 读取');
    const code = readCodeFrom(file);
    const body = o.type ? { code, type: o.type } : { code };
    out((await call('PUT', '/api/modules/' + encodeURIComponent(name) + '/code', body)).data, o);
  },
  async 'module-delete'(o) {
    const [name] = o._;
    if (!name) fail('用法：module-delete <name>   不可逆，删前先 code-get 确认');
    out((await call('DELETE', '/api/modules/' + encodeURIComponent(name))).data, o);
  },
  async 'module-import'(o) {
    const [path] = o._;
    if (!path) fail('用法：module-import <绝对路径>   支持 .bas .cls .frm .ctl .pag .dob');
    out((await call('POST', '/api/modules/import', { path })).data, o);
  },

  async procs(o) {
    const [name] = o._;
    if (!name) fail('用法：procs <module>');
    out((await call('GET', '/api/modules/' + encodeURIComponent(name) + '/procedures')).data, o);
  },
  async 'proc-get'(o) {
    const [name, proc] = o._;
    if (!name || !proc) fail('用法：proc-get <module> <proc>   Property 加 .Get/.Let/.Set 后缀');
    out((await call('GET', '/api/modules/' + encodeURIComponent(name) + '/procedures/' + encodeURIComponent(proc))).data, o);
  },
  async 'proc-put'(o) {
    const [name, proc, file] = o._;
    if (!name || !proc || !file) fail('用法：proc-put <module> <proc> <file|->   注意：操作范围含过程上方紧邻注释');
    const code = readCodeFrom(file);
    out((await call('PUT', '/api/modules/' + encodeURIComponent(name) + '/procedures/' + encodeURIComponent(proc), { code })).data, o);
  },
  async 'proc-delete'(o) {
    const [name, proc] = o._;
    if (!name || !proc) fail('用法：proc-delete <module> <proc> [--cleanup]   不可逆');
    out((await call('DELETE', '/api/modules/' + encodeURIComponent(name) + '/procedures/' + encodeURIComponent(proc) + (o.cleanup ? '?cleanup=1' : ''))).data, o);
  },

  async ide(o) { out((await call('GET', '/api/ide')).data, o); },

  async 'lines-get'(o) {
    const [name] = o._;
    if (!name) fail('用法：lines-get <module> [--start N --end M]');
    let q = '';
    if (o.start) q += 'start=' + encodeURIComponent(o.start);
    if (o.end) q += (q ? '&' : '') + 'end=' + encodeURIComponent(o.end);
    out((await call('GET', '/api/modules/' + encodeURIComponent(name) + '/lines' + (q ? '?' + q : ''))).data, o);
  },
  async 'lines-insert'(o) {
    const [name, file] = o._;
    if (!name || !file || !o.start) fail('用法：lines-insert <module> --start N <file|->   在第 N 行之前插入');
    const code = readCodeFrom(file);
    out((await call('POST', '/api/modules/' + encodeURIComponent(name) + '/lines', { start: o.start, code })).data, o);
  },
  async 'lines-replace'(o) {
    const [name, file] = o._;
    if (!name || !file || !o.start || !o.end) fail('用法：lines-replace <module> --start N --end M <file|->');
    const code = readCodeFrom(file);
    out((await call('PATCH', '/api/modules/' + encodeURIComponent(name) + '/lines', { start: o.start, end: o.end, code })).data, o);
  },
  async 'lines-delete'(o) {
    const [name] = o._;
    if (!name || !o.start || !o.end) fail('用法：lines-delete <module> --start N --end M   不可逆，删前先读确认');
    out((await call('DELETE', '/api/modules/' + encodeURIComponent(name) + '/lines?start=' + encodeURIComponent(o.start) + '&end=' + encodeURIComponent(o.end))).data, o);
  },

  async refs(o) { out((await call('GET', '/api/references')).data, o); },
  async 'ref-add'(o) {
    if (o.guid) out((await call('POST', '/api/references', { guid: o.guid, major: o.major || '1', minor: o.minor || '0' })).data, o);
    else if (o.path) out((await call('POST', '/api/references', { path: o.path })).data, o);
    else if (o.name) out((await call('POST', '/api/references', { name: o.name })).data, o);
    else fail('用法：ref-add --guid {GUID} [--major N --minor N] | --path D:\\dll路径 | --name "库名"');
  },
  async 'ref-del'(o) {
    const [guid] = o._;
    if (!guid) fail('用法：ref-del <GUID>   如 {2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}');
    out((await call('DELETE', '/api/references/' + encodeURIComponent(guid))).data, o);
  },
  async typelibs(o) {
    const [kw] = o._;
    out((await call('GET', '/api/typelibs' + (kw ? '?keyword=' + encodeURIComponent(kw) : ''))).data, o);
  },

  async 'form-get'(o) {
    const [name] = o._;
    if (!name) fail('用法：form-get <form>');
    out((await call('GET', '/api/forms/' + encodeURIComponent(name))).data, o);
  },
  async 'ctrl-add'(o) {
    const [form, name, cls] = o._;
    if (!form || !name || !cls) fail('用法：ctrl-add <form> <name> <className> [--left N --top N --width N --height N --caption S]   如 cls=VB.CommandButton');
    const body = { name, className: cls };
    for (const k of ['left', 'top', 'width', 'height']) if (o[k]) body[k] = o[k];
    if (o.caption) body.caption = o.caption;
    out((await call('POST', '/api/forms/' + encodeURIComponent(form) + '/controls', body)).data, o);
  },
  async 'ctrl-set'(o) {
    const [form, ctrl, file] = o._;
    if (!form || !ctrl || !file) fail('用法：ctrl-set <form> <ctrl> <props.json|->   文件形如 {"props":{"Caption":"X","Left":100}}');
    const props = readJsonFrom(file, 'props 文件');
    out((await call('PATCH', '/api/forms/' + encodeURIComponent(form) + '/controls/' + encodeURIComponent(ctrl), { props })).data, o);
  },
  async 'ctrl-del'(o) {
    const [form, ctrl] = o._;
    if (!form || !ctrl) fail('用法：ctrl-del <form> <ctrl>   不可逆，删前先 form-get 确认');
    out((await call('DELETE', '/api/forms/' + encodeURIComponent(form) + '/controls/' + encodeURIComponent(ctrl))).data, o);
  },
  async 'form-align'(o) {
    const [form] = o._;
    if (!form || !o.mode || !o.controls) fail('用法：form-align <form> --mode M --controls a,b,c [--value N]   mode: left|right|top|bottom|center-h|center-v|width|height|space-h|space-v');
    const body = { mode: o.mode, controls: o.controls.split(',').map(s => s.trim()).filter(Boolean) };
    if (o.value) body.value = o.value;
    out((await call('POST', '/api/forms/' + encodeURIComponent(form) + '/align', body)).data, o);
  },
  async 'form-style'(o) {
    const [form, file] = o._;
    if (!form || !file) fail('用法：form-style <form> <style.json|->   形如 {"caption":"X","width":6000,"height":4000,"startUpPosition":1}');
    const style = readJsonFrom(file, 'style 文件');
    out((await call('PATCH', '/api/forms/' + encodeURIComponent(form) + '/style', style)).data, o);
  },

  async 'win-show'(o) {
    const [w] = o._;
    out((await call('POST', '/api/ui/immediate-window', { window: w || 'immediate' })).data, o);
  },
  async save(o) { out((await call('POST', '/api/ui/save-project', {})).data, o); },

  async 'debug-status'(o) { out((await call('GET', '/api/debug/status')).data, o); },
  async 'debug-run'(o) { out((await call('POST', '/api/debug/run', {})).data, o); },
  async 'debug-stop'(o) { out((await call('POST', '/api/debug/stop', {})).data, o); },
  async bps(o) { out((await call('GET', '/api/debug/breakpoints')).data, o); },
  async 'bp-set'(o) {
    const [mod, line] = o._;
    if (!mod || !line) fail('用法：bp-set <module> <line>   toggle 语义：该行已有断点则本次为取消');
    const n = Number(line);
    if (!Number.isInteger(n) || n < 1) fail('line 必须为正整数（1 基准）：' + line);
    out((await call('POST', '/api/debug/breakpoints', { module: mod, line: n })).data, o);
  },
  async 'bp-del'(o) {
    const [mod, line] = o._;
    if (mod && line) {
      const n = Number(line);
      if (!Number.isInteger(n) || n < 1) fail('line 必须为正整数（1 基准）：' + line);
      out((await call('DELETE', '/api/debug/breakpoints', { module: mod, line: n })).data, o);
    }
    else out((await call('DELETE', '/api/debug/breakpoints')).data, o); // 无参数 = 清除全部断点（0.1.9 起必须发空 body：{} 是非空 body 会被 400 拒）
  },
  async 'bp-clear'(o) { out((await call('DELETE', '/api/debug/breakpoints')).data, o); }, // 空 body 全清（0.1.9 契约：非空 body 解析不出 module 必 400）
  async 'debug-output'(o) { out((await call('GET', '/api/debug/output')).data, o); },
};

// ---------- 入口 ----------
(async () => {
  const argv = process.argv.slice(2);
  if (!argv.length || argv[0] === 'help' || argv[0] === '--help') {
    console.log('vb6ide-api 工具集（node skill.js <command>）\n');
    console.log('  status                                   探测端口与服务状态（其余命令的前置）');
    console.log('  project | components | snapshot | ide     工程信息（只读；snapshot 大输出默认截断）');
    console.log('  code-get <mod> [--start N --end M]       读模块代码（1 基准）');
    console.log('  module-create <name> [type]              新建模块');
    console.log('  code-put <mod> <file|-> [--type t]       整体写入模块（仅标准/类）');
    console.log('  module-delete <mod>                      删除模块（不可逆）');
    console.log('  module-import <path>                     导入模块文件（.bas/.cls）');
    console.log('  procs <mod>                              过程列表');
    console.log('  proc-get <mod> <proc>                    读单过程');
    console.log('  proc-put <mod> <proc> <file|->           整体替换过程');
    console.log('  proc-delete <mod> <proc> [--cleanup]     删除过程（不可逆）');
    console.log('  lines-get <mod> [--start --end]          读行范围（语义同 code-get）');
    console.log('  lines-insert <mod> --start N <file|->    第 N 行前插入');
    console.log('  lines-replace <mod> --start --end <f|->  替换行范围（空代码=删除）');
    console.log('  lines-delete <mod> --start --end         删除行范围（不可逆）');
    console.log('  refs | ref-add | ref-del | typelibs      引用库管理（ref-add 三模式 guid/path/name）');
    console.log('  form-get <form>                          窗体信息+控件清单（坐标 twips）');
    console.log('  ctrl-add <form> <name> <cls> [坐标]      添加控件');
    console.log('  ctrl-set <form> <ctrl> <props.json|->    设置控件属性');
    console.log('  ctrl-del <form> <ctrl>                   删除控件（不可逆）');
    console.log('  form-align <form> --mode M --controls .. 对齐/等距（10 种 mode）');
    console.log('  form-style <form> <style.json|->         窗体样式（caption/尺寸/启动位）');
    console.log('  win-show [immediate|locals|...]          打开 IDE 工具窗口');
    console.log('  save                                     保存当前工程');
    console.log('  debug-status | debug-run | debug-stop    调试三态（run=启动/继续）');
    console.log('  bps | bp-set | bp-del | bp-clear         断点（自记账；bp-set 是 toggle）');
    console.log('  debug-output                             读立即窗口输出');
    console.log('\n公共项：--full 全量输出 / --max N 截断字符数 / VB6IDE_PORT 指定端口');
    console.log('bp-del 无参数=清除全部断点（发空 body；服务端 0.1.9 起拒绝非空畸形 body 全清）；bp-set 同行重复调用会取消断点（toggle 语义）');
    return;
  }
  const { cmd, opts } = parseArgs(argv);
  const fn = commands[cmd];
  if (!fn) fail('未知命令：' + cmd + '（node skill.js help 查看清单）');
  await fn(opts);
})().catch(e => fail(e && e.message ? e.message : String(e)));
