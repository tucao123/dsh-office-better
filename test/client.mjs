/**
 * 客户端半边的无浏览器验证：用桩 window/__ModuleLoader__/React/slots 把 lib/client.js 跑起来，
 * 检查 2 个 Slot 注册、按钮渲染、面板 iframe 的 src 是否正确、文件列表能否解析。
 * 运行：node test/client.mjs
 */
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import vm from 'node:vm';

const here = dirname(fileURLToPath(import.meta.url));
const source = readFileSync(join(here, '..', 'lib', 'client.js'), 'utf8');

const problems = [];

/* ------------------------------------------------------------- 桩环境 */

const h = (type, props, ...children) => ({ type, props: props ?? {}, children: children.flat() });
const React = {
  createElement: h,
  Fragment: 'Fragment',
  useState: (initial) => [typeof initial === 'function' ? initial() : initial, () => {}],
  useEffect: () => {},
};

const styleNodes = [];
const documentStub = {
  head: { appendChild: (node) => styleNodes.push(node) },
  createElement: () => ({ setAttribute() {}, textContent: '' }),
};

const fetched = [];
const fetchStub = async (url) => {
  fetched.push(url);
  if (url.includes('/files')) {
    return {
      ok: true,
      status: 200,
      json: async () => ({
        files: [
          { name: '报表.pptx', type: 'pptx', size: 20480, mtime: Date.now() },
          { name: '说明.docx', type: 'docx', size: 5120, mtime: Date.now() },
        ],
      }),
    };
  }
  return { ok: false, status: 404, json: async () => ({}) };
};

const sandbox = {
  window: {
    __ModuleLoader__: { load: (definition) => { sandbox.__def = definition; } },
    setInterval: () => 0,
    clearInterval: () => {},
  },
  document: documentStub,
  fetch: fetchStub,
  console,
  setTimeout,
};
vm.createContext(sandbox);
vm.runInContext(source, sandbox);

const def = sandbox.__def;
if (!def) problems.push('__ModuleLoader__.load 未被调用');
if (def && def.id !== 'dsh-office-better') problems.push(`bundle id 错误：${def && def.id}`);

const mod = def.factory((name) => {
  if (name === 'react') return React;
  throw new Error(`未预期的 require("${name}")`);
});
if (typeof mod.apply !== 'function') problems.push('exports.apply 缺失');
if (!Array.isArray(mod.inject) || !mod.inject.includes('slots')) problems.push('exports.inject 未声明 slots');

/* ------------------------------------------------------------- 桩 ctx */

const registrations = [];
const slots = {
  inject(name, callback) { callback(); },
  register(options, component) {
    registrations.push({ options, component });
    return () => {};
  },
};
const effects = [];
const ctx = {
  get: (name) => (name === 'slots' ? slots : undefined),
  effect: (fn) => { effects.push(fn); return () => {}; },
};
mod.apply(ctx);

const bySlot = Object.fromEntries(registrations.map((r) => [r.options.name, r]));
const headerReg = bySlot['conversation.session.header.utilities'];
const overlayReg = bySlot['shell.overlay'];
if (!headerReg) problems.push('未注册会话头部按钮');
if (!overlayReg) problems.push('未注册 shell.overlay 面板');
if (headerReg && headerReg.options.id !== 'office-better-toggle') problems.push('按钮 id 不对');
if (overlayReg && overlayReg.options.id !== 'office-better-panel') problems.push('面板 id 不对');

/* ------------------------------------------------------------- 渲染检查 */

const flatText = (node) => {
  if (node == null) return '';
  if (typeof node === 'string' || typeof node === 'number') return String(node);
  if (Array.isArray(node)) return node.map(flatText).join(' ');
  return flatText(node.children);
};

const findAll = (node, predicate, acc = []) => {
  if (node == null || typeof node !== 'object') return acc;
  if (Array.isArray(node)) { node.forEach((n) => findAll(n, predicate, acc)); return acc; }
  if (predicate(node)) acc.push(node);
  findAll(node.children, predicate, acc);
  return acc;
};

// 注册时是 (props) => h(Comp, props) 的包装，这里展开一层再渲染
const render = (registered, props) => {
  const el = registered.component(props);
  if (el && typeof el.type === 'function') return el.type(el.props ?? {});
  return el;
};

const closed = render(overlayReg, {});
if (closed !== null) problems.push('面板未打开时不应渲染');

const button = render(headerReg, { sessionId: 'session-test' });
if (button.props.className !== 'dob-btn') problems.push('按钮样式类不对');
if (flatText(button).trim() !== 'Office') problems.push(`按钮文字异常：${flatText(button).trim()}`);
if (typeof button.props.onClick !== 'function') problems.push('按钮没有 onClick');

button.props.onClick();
await new Promise((resolve) => setTimeout(resolve, 40));
if (!fetched.some((u) => u.includes('/files?session=session-test'))) problems.push('点击后没有请求文件列表');

const opened = render(overlayReg, {});
const html = flatText(opened);
if (!html.includes('Office 预览')) problems.push('面板标题缺失');
if (!html.includes('报表.pptx')) problems.push('文件列表没有渲染文件名');

const iframe = findAll(opened, (n) => n.type === 'iframe')[0];
if (!iframe) problems.push('面板里没有 iframe');
else {
  const src = iframe.props.src;
  const want = '/api/office-better/preview?session=session-test&file=';
  if (!src.startsWith(want)) problems.push(`iframe src 前缀不对：${src}`);
  if (!src.includes(encodeURIComponent('报表.pptx'))) problems.push(`iframe src 未编码文件名：${src}`);
}

if (styleNodes.length === 0) problems.push('未注入样式表');
else if (!String(styleNodes[0].textContent).includes('.dob-panel')) problems.push('样式表内容异常');

/* ------------------------------------------------------------- 输出 */

console.log(`注册：${registrations.map((r) => `${r.options.name}#${r.options.id}`).join(', ')}`);
console.log(`请求：${fetched.map((u) => u.replace(/\?.*/, '?…')).join(', ')}`);
console.log(`iframe：${iframe ? iframe.props.src : '(无)'}`);
if (problems.length > 0) {
  console.error('\nFAIL:');
  for (const p of problems) console.error(' - ' + p);
  process.exit(1);
}
console.log('\nCLIENT OK');
