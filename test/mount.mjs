/**
 * 挂载冒烟测试：用桩 ctx 跑 apply()，验证
 *  1. 9 个工具全部注册
 *  2. 工具能真正产出 pptx/docx/xlsx
 *  3. 预览路由返回的 HTML 里确实有 SVG 幻灯片
 * 运行：node test/mount.mjs
 */
import { writeFileSync, mkdirSync, rmSync, existsSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { apply, Config } from '../lib/index.js';

const here = dirname(fileURLToPath(import.meta.url));
const wsRoot = join(here, '..', 'out', 'ws');
rmSync(wsRoot, { recursive: true, force: true });
mkdirSync(wsRoot, { recursive: true });

const tools = [];
const routes = [];
const logs = [];
const fakeServer = {
  register(route) {
    routes.push(route);
    return () => routes.splice(routes.indexOf(route), 1);
  },
};
const ctx = {
  tools: { register(def) { tools.push(def); } },
  inject(names, callback) {
    const webCtx = {
      get: (key) => (key === 'webServer' ? fakeServer : undefined),
      effect: (fn) => { const d = fn(); return d; },
      logger: { info: (m) => logs.push(m), warn: (m) => logs.push(m) },
    };
    callback(webCtx);
  },
  effect: (fn) => fn(),
  logger: { info: (m) => logs.push(m), warn: (m) => logs.push(m) },
};

const config = { workspaceDir: wsRoot, previewPrefix: '/api/office-better' };
apply(ctx, config);

const problems = [];
const names = tools.map((t) => t.name);
const expected = ['office_deck', 'office_deck_add', 'office_lint', 'office_preview', 'office_docx', 'office_xlsx', 'office_read', 'office_list', 'office_guide'];
for (const name of expected) if (!names.includes(name)) problems.push(`缺少工具 ${name}`);
if (tools.length !== expected.length) problems.push(`工具数 ${tools.length}，期望 ${expected.length}`);
if (routes.length !== 1 || routes[0].kind !== 'prefix') problems.push('预览路由未注册');

const byName = Object.fromEntries(tools.map((t) => [t.name, t]));
const exec = { agent: { session: { id: 'test', header: { cwd: wsRoot } } } };
const run = async (name, args) => byName[name].execute(args, exec);

// 1) pptx
const deck = {
  template: 'consulting',
  slides: [
    { layout: 'cover', title: 'Office Better', subtitle: '自研 OOXML 生成器冒烟测试', eyebrow: 'v0.1' },
    { layout: 'cover', title: 'x', hidden: true },
    { layout: 'kpi', title: '关键指标', metrics: [{ value: '0', label: '外部进程', note: '零 spawn' }, { value: '3', label: '文档格式' }, { value: '15', label: '幻灯片版式' }] },
    { layout: 'ending', title: '完成' },
  ],
};
const deckResult = await run('office_deck', { filename: 'mount.pptx', deck });
if (deckResult.slides !== 3) problems.push(`pptx 页数 ${deckResult.slides}（hidden 页应被跳过，期望 3）`);
if ((deckResult.errors ?? []).length > 0) problems.push(`pptx lint 有 ${deckResult.errors.length} 个错误`);
if (!existsSync(join(wsRoot, 'mount.pptx'))) problems.push('mount.pptx 未落盘');
if (deckResult.text.includes('预览')) { /* 有预览提示 */ } else problems.push('返回里没有预览地址');

// 2) deck_add
const addResult = await run('office_deck_add', {
  filename: 'mount.pptx',
  slides: [{ layout: 'quote', quote: '版面即数据，数据即版面。', author: 'dsh-office-better' }],
  at: 3,
});
if (addResult.slides !== 4) problems.push(`追加后页数 ${addResult.slides}，期望 4`);

// 3) docx
const docxResult = await run('office_docx', {
  filename: 'mount.docx',
  doc: {
    title: '测试文档',
    blocks: [
      { type: 'heading', level: 1, text: '一、背景' },
      { type: 'paragraph', text: '这是一段测试正文，用来验证 docx 生成。' },
      { type: 'bullets', items: ['要点一', '要点二'] },
      { type: 'numbered', items: ['第一步', '第二步'] },
      { type: 'table', headers: ['项目', '数值'], rows: [['零外部依赖', '✓'], ['版式', '15 种']] },
      { type: 'quote', text: '引用一句话。' },
      { type: 'code', text: 'const x = 1;' },
    ],
  },
});
if (docxResult.blocks !== 7) problems.push(`docx 块数 ${docxResult.blocks}，期望 7`);

// 4) xlsx
const xlsxResult = await run('office_xlsx', {
  filename: 'mount.xlsx',
  book: {
    sheets: [
      { name: '数据', rows: [['月份', '收入', '增速'], ['1月', 120, '=B2*0.1'], ['2月', 150, '=B3*0.1']], widths: [10, 12, 12] },
      { name: '说明', header: false, rows: [['这是第二个工作表']] },
    ],
  },
});
if (xlsxResult.sheets !== 2) problems.push(`xlsx 工作表 ${xlsxResult.sheets}，期望 2`);

// 5) read 回读
const readResult = await run('office_read', { filename: 'mount.pptx' });
if (!readResult.text.includes('Office Better')) problems.push('office_read 没读出封面标题');
const readXlsx = await run('office_read', { filename: 'mount.xlsx' });
if (!readXlsx.text.includes('收入')) problems.push('office_read 没读出 xlsx 单元格');

// 6) lint / list / guide
const lintResult = await run('office_lint', { filename: 'mount.pptx' });
if (!lintResult.text.includes('检查完成')) problems.push('office_lint 输出异常');
const listResult = await run('office_list', {});
if (listResult.files.length !== 3) problems.push(`office_list 列出 ${listResult.files.length} 个文件，期望 3`);
const guideResult = await run('office_guide', {});
if (!guideResult.text.includes('字数红线')) problems.push('office_guide 内容异常');

// 7) 预览路由
const handler = routes[0].handler;
const callRoute = async (url) => {
  const res = { status: 0, headers: {}, body: '', writeHead(code, headers) { this.status = code; this.headers = headers; }, end(chunk) { if (chunk) this.body += chunk.toString(); } };
  await handler({ url }, res);
  return res;
};
const previewRes = await callRoute('/api/office-better/preview?session=test&file=mount.pptx');
if (previewRes.status !== 200) problems.push(`预览状态 ${previewRes.status}`);
if (!previewRes.body.includes('<svg')) problems.push('预览里没有 SVG');
if (!previewRes.body.includes('Office Better')) problems.push('预览里没有封面文字');
const filesRes = await callRoute('/api/office-better/files?session=test');
if (filesRes.status !== 200 || !filesRes.body.includes('mount.pptx')) problems.push('files 路由异常');
const indexRes = await callRoute('/api/office-better/index?session=test');
if (!indexRes.body.includes('下载')) problems.push('index 路由异常');
const missingRes = await callRoute('/api/office-better/preview?session=test&file=nope.pptx');
if (missingRes.status !== 404) problems.push(`缺文件应 404，实际 ${missingRes.status}`);

console.log(`工具 ${tools.length} 个：${names.join(', ')}`);
console.log(`路由 ${routes.length} 条：${routes[0]?.path}`);
console.log(`预览 HTML ${(previewRes.body.length / 1024).toFixed(1)} KB，SVG ${(previewRes.body.match(/<svg/g) ?? []).length} 个`);
console.log(`日志：${logs.join(' | ')}`);
writeFileSync(join(here, '..', 'out', 'mount-report.txt'), `problems=${problems.length}\n${problems.join('\n')}`);
if (problems.length > 0) {
  console.error('\nFAIL:');
  for (const p of problems) console.error(' - ' + p);
  process.exit(1);
}
console.log('\nMOUNT OK');
