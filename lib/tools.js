/**
 * @module tools
 * 9 个工具：deck / deck_add / lint / preview / docx / xlsx / read / list / guide。
 * 所有工具都是纯进程内计算（写文件 + 生成 OOXML），不 spawn 任何外部进程。
 */
import { writeFileSync, readFileSync, existsSync } from 'node:fs';
import { defineTool } from '@deepseek-ai/dsh-tools';
import { compileDeck, summarize, catalog } from './pptx/deck.js';
import { LIMITS, MAX_SLIDES } from './pptx/spec.js';
import { LAYOUT_IDS, renderSlide, DARK_LAYOUTS } from './pptx/layouts.js';
import { THEMES, TEMPLATES, resolveStyle } from './pptx/theme.js';
import { previewHtml } from './pptx/preview.js';
import { lintPages } from './pptx/lint.js';
import { parseDocSpec, buildDocx } from './docx/docx.js';
import { parseSheetSpec, buildXlsx } from './xlsx/xlsx.js';
import { readOffice, outlineToText } from './read/outline.js';
import { previewHtmlForOutline } from './routes.js';

const textCard = (...lines) => [{ type: 'text', text: lines.filter(Boolean).join('\n') }];

function sessionOf(exec) {
  const id = exec?.agent?.session?.id;
  return typeof id === 'string' && id ? id : 'default';
}

function cwdOf(exec) {
  const cwd = exec?.agent?.session?.header?.cwd;
  return typeof cwd === 'string' && cwd ? cwd : undefined;
}

const out = (text, extra = {}) => ({
  schema: { type: 'object', additionalProperties: true },
  render: (_args, value) => textCard(value.text),
});

function guideText() {
  const themeList = Object.values(THEMES).map((t) => `  ${t.id.padEnd(18)} ${t.name}`).join('\n');
  const tplList = Object.values(TEMPLATES).map((t) => `  ${t.id.padEnd(18)} ${t.name}（默认主题 ${t.defaultTheme}）`).join('\n');
  const limits = Object.entries(LIMITS).map(([layout, spec]) => (
    `  ${layout.padEnd(9)} ${Object.entries(spec).map(([k, v]) => `${k}≤${v}`).join(' ')}`
  )).join('\n');
  return [
    '# dsh-office-better 设计约束',
    '',
    '## 流程',
    '1. office_deck 传 DeckSpec（只填内容，坐标/字号/配色由版式引擎算）',
    '2. 返回的 text 里已含 lint 结果；有 error 必须修，warning 建议修',
    '3. office_preview 拿到预览 URL（浏览器直接看，和 pptx 同源 IR）',
    '',
    `## 画布`, '  960x540pt（16:9），边距 72pt，正文区 y=124..466',
    '',
    '## 版式（15 种）', `  ${LAYOUT_IDS.join(', ')}`,
    '',
    '## 主题', themeList,
    '',
    '## 模板', tplList,
    '',
    '## 字数红线（超限会在校验期直接报错，不会画出溢出）', limits,
    '',
    `## 上限`, `  slides ≤ ${MAX_SLIDES}；深色整幅页（cover/section/ending）之间至少隔一页浅色内容页；不要连续两页同版式`,
    '',
    '## DeckSpec 示例',
    '{"template":"product-launch","theme":"tech-cyan","slides":[',
    '  {"layout":"cover","eyebrow":"2026 规划","title":"产品名\\nV2","subtitle":"一句话说明"},',
    '  {"layout":"bullets","title":"三个判断","columns":2,"items":[{"title":"要点","desc":"一句解释"}]},',
    '  {"layout":"ending","title":"谢谢"}]}',
  ].join('\n');
}

/**
 * 注册全部工具。
 * @param {import('@deepseek-ai/cordis').Context} ctx
 * @param {{workspace: import('./workspace.js').Workspace, config: object}} deps
 */
export function registerTools(ctx, deps) {
  const { workspace } = deps;
  const prefix = deps.config.previewPrefix ?? '/api/office-better';

  const previewUrl = (sessionId) => `${prefix}/index?session=${encodeURIComponent(sessionId)}`;

  // ---------------------------------------------------------------- deck
  ctx.tools.register(defineTool({
    name: 'office_deck',
    description: '用 DeckSpec 生成 pptx（16:9，15 种版式，8 套主题，6 套模板）。只填内容，排版由引擎计算；返回 lint 结果与预览地址。',
    parameters: {
      filename: { type: 'string', required: true, description: '文件名，必须以 .pptx 结尾' },
      deck: { type: 'object', additionalProperties: true, required: true, description: 'DeckSpec：{template?,theme?,slides:[{layout,...}]}，先跑 office_guide 看约束' },
      title: { type: 'string', description: '文档标题（元数据）' },
      author: { type: 'string', description: '作者（元数据）' },
    },
    output: out(),
    isConcurrencySafe: () => false,
    async execute(args, exec) {
      const sessionId = sessionOf(exec);
      const cwd = cwdOf(exec);
      const abs = workspace.resolve(sessionId, args.filename, cwd);
      const compiled = compileDeck(args.deck, { title: args.title, author: args.author });
      writeFileSync(abs, compiled.pptx);
      writeFileSync(workspace.sidecarPath(sessionId, args.filename, cwd), JSON.stringify({
        spec: args.deck,
        options: { title: args.title, author: args.author },
        savedAt: new Date().toISOString(),
      }, null, 2));
      const text = [
        summarize(compiled, args.filename),
        `路径：${abs}`,
        `预览：${previewUrl(sessionId)}`,
      ].join('\n');
      return {
        text,
        file: args.filename,
        path: abs,
        slides: compiled.pages.length,
        errors: compiled.lint.errors,
        warnings: compiled.lint.warnings,
        previewUrl: previewUrl(sessionId),
      };
    },
  }));

  // ---------------------------------------------------------------- deck_add
  ctx.tools.register(defineTool({
    name: 'office_deck_add',
    description: '给已由 office_deck 生成的 pptx 追加页面（沿用同一份 DeckSpec 与主题，不会换配色）。',
    parameters: {
      filename: { type: 'string', required: true, description: '已存在的 .pptx' },
      slides: { type: 'array', required: true, items: { type: 'object', additionalProperties: true }, description: '要追加的 slide 对象数组（与 office_deck 的 slides 同结构）' },
      at: { type: 'integer', description: '插入位置（1-based，缺省追加到末尾）' },
    },
    output: out(),
    isConcurrencySafe: () => false,
    async execute(args, exec) {
      const sessionId = sessionOf(exec);
      const cwd = cwdOf(exec);
      const abs = workspace.resolve(sessionId, args.filename, cwd);
      if (!existsSync(abs)) throw new Error(`文件不存在：${args.filename}。请先用 office_deck 生成。`);
      const sidecarPath = workspace.sidecarPath(sessionId, args.filename, cwd);
      if (!existsSync(sidecarPath)) {
        throw new Error(`${args.filename} 不是本插件生成的（没有 IR 旁挂文件），无法原地追加。请用 office_deck 重新生成。`);
      }
      const stored = JSON.parse(readFileSync(sidecarPath, 'utf8'));
      const spec = stored.spec;
      const slides = Array.isArray(args.slides) ? args.slides : [args.slides];
      const at = Number.isInteger(args.at) ? Math.max(0, Math.min(slides.length, args.at - 1)) : spec.slides.length;
      spec.slides.splice(at, 0, ...slides);
      const compiled = compileDeck(spec, stored.options ?? {});
      writeFileSync(abs, compiled.pptx);
      writeFileSync(sidecarPath, JSON.stringify({ ...stored, spec, savedAt: new Date().toISOString() }, null, 2));
      return {
        text: [summarize(compiled, args.filename), `已在第 ${at + 1} 页插入 ${slides.length} 页`, `预览：${previewUrl(sessionId)}`].join('\n'),
        file: args.filename,
        path: abs,
        slides: compiled.pages.length,
        errors: compiled.lint.errors,
        warnings: compiled.lint.warnings,
        previewUrl: previewUrl(sessionId),
      };
    },
  }));

  // ---------------------------------------------------------------- lint
  ctx.tools.register(defineTool({
    name: 'office_lint',
    description: '只做检查不写文件：对 DeckSpec 或已生成的文件做溢出/越界/对比度/字号/密度/层级检查。',
    parameters: {
      deck: { type: 'object', additionalProperties: true, description: '要检查的 DeckSpec（与 filename 二选一）' },
      filename: { type: 'string', description: '已生成的文件名（读 IR 旁挂文件）' },
    },
    output: out(),
    isConcurrencySafe: () => true,
    async execute(args, exec) {
      const sessionId = sessionOf(exec);
      const cwd = cwdOf(exec);
      let compiled;
      if (args.deck) {
        compiled = compileDeck(args.deck, {});
      } else if (args.filename) {
        const sidecarPath = workspace.sidecarPath(sessionId, args.filename, cwd);
        if (!existsSync(sidecarPath)) throw new Error(`${args.filename} 没有 IR 旁挂文件，无法 lint。`);
        const stored = JSON.parse(readFileSync(sidecarPath, 'utf8'));
        compiled = compileDeck(stored.spec, stored.options ?? {});
      } else {
        throw new Error('deck 和 filename 至少要给一个');
      }
      const { errors, warnings, stats } = compiled.lint;
      const lines = [
        `检查完成：${compiled.pages.length} 页，${errors.length} 个错误，${warnings.length} 个提醒`,
        `最小字号 ${stats.smallestFont}pt · 最低对比度 ${stats.minContrast.toFixed(2)}:1 · 形状 ${stats.shapes} 个`,
      ];
      for (const e of errors) lines.push(`✗ [${e.code}] 第 ${e.slide} 页 ${e.message}`);
      for (const w of warnings) lines.push(`! [${w.code}] 第 ${w.slide} 页 ${w.message}`);
      if (errors.length === 0 && warnings.length === 0) lines.push('通过：没有发现版式问题。');
      return { text: lines.join('\n'), errors, warnings, stats };
    },
  }));

  // ---------------------------------------------------------------- preview
  ctx.tools.register(defineTool({
    name: 'office_preview',
    description: '拿到会话文档的浏览器预览地址（预览由同一份版式 IR 渲染，和 pptx 一致，不启动外部进程）。',
    parameters: {
      filename: { type: 'string', description: '可选：指定文件则直接给该文件的预览地址' },
    },
    output: out(),
    isConcurrencySafe: () => true,
    async execute(args, exec) {
      const sessionId = sessionOf(exec);
      const url = args.filename
        ? `${prefix}/preview?session=${encodeURIComponent(sessionId)}&file=${encodeURIComponent(args.filename)}`
        : previewUrl(sessionId);
      return { text: `预览地址：${url}\n（在 DSH Web 里直接打开；页面自带文件列表与下载链接）`, previewUrl: url };
    },
  }));

  // ---------------------------------------------------------------- docx
  ctx.tools.register(defineTool({
    name: 'office_docx',
    description: '用块列表生成 docx（heading/paragraph/bullets/numbered/table/quote/code/divider/pagebreak）。',
    parameters: {
      filename: { type: 'string', required: true, description: '文件名，必须以 .docx 结尾' },
      doc: { type: 'object', additionalProperties: true, required: true, description: 'DocSpec：{title?,subtitle?,author?,accent?,blocks:[...]}' },
      author: { type: 'string', description: '作者（元数据）' },
    },
    output: out(),
    isConcurrencySafe: () => false,
    async execute(args, exec) {
      const sessionId = sessionOf(exec);
      const cwd = cwdOf(exec);
      const abs = workspace.resolve(sessionId, args.filename, cwd);
      const spec = parseDocSpec(args.doc);
      if (args.author) spec.author = args.author;
      const buffer = buildDocx(spec);
      writeFileSync(abs, buffer);
      return {
        text: [
          `文件：${args.filename}（${(buffer.length / 1024).toFixed(1)} KB）`,
          `块数：${spec.blocks.length} · A4 版心 1440 twips 边距`,
          `路径：${abs}`,
          `预览：${prefix}/preview?session=${encodeURIComponent(sessionId)}&file=${encodeURIComponent(args.filename)}`,
        ].join('\n'),
        file: args.filename,
        path: abs,
        blocks: spec.blocks.length,
      };
    },
  }));

  // ---------------------------------------------------------------- xlsx
  ctx.tools.register(defineTool({
    name: 'office_xlsx',
    description: '生成 xlsx（多工作表、表头样式、冻结首行、列宽、公式）。',
    parameters: {
      filename: { type: 'string', required: true, description: '文件名，必须以 .xlsx 结尾' },
      book: { type: 'object', additionalProperties: true, required: true, description: 'BookSpec：{title?,sheets:[{name,rows,header?,widths?}]}；字符串以 = 开头视为公式' },
      author: { type: 'string', description: '作者（元数据）' },
    },
    output: out(),
    isConcurrencySafe: () => false,
    async execute(args, exec) {
      const sessionId = sessionOf(exec);
      const cwd = cwdOf(exec);
      const abs = workspace.resolve(sessionId, args.filename, cwd);
      const spec = parseSheetSpec(args.book);
      if (args.author) spec.author = args.author;
      const buffer = buildXlsx(spec);
      writeFileSync(abs, buffer);
      return {
        text: [
          `文件：${args.filename}（${(buffer.length / 1024).toFixed(1)} KB）`,
          `工作表：${spec.sheets.map((s) => `${s.name}(${s.rows.length} 行)`).join(', ')}`,
          `路径：${abs}`,
        ].join('\n'),
        file: args.filename,
        path: abs,
        sheets: spec.sheets.length,
      };
    },
  }));

  // ---------------------------------------------------------------- read
  ctx.tools.register(defineTool({
    name: 'office_read',
    description: '读取 docx / xlsx / pptx 的文本大纲（段落样式、幻灯页文本、单元格值）。',
    parameters: {
      filename: { type: 'string', required: true, description: '会话内已存在的文件名' },
      maxLines: { type: 'integer', description: '最多返回多少行（默认 400）' },
    },
    output: out(),
    isConcurrencySafe: () => true,
    async execute(args, exec) {
      const sessionId = sessionOf(exec);
      const cwd = cwdOf(exec);
      const abs = workspace.resolve(sessionId, args.filename, cwd);
      if (!existsSync(abs)) throw new Error(`文件不存在：${args.filename}`);
      const kind = args.filename.slice(args.filename.lastIndexOf('.') + 1).toLowerCase();
      const outline = readOffice(readFileSync(abs), kind);
      const text = outlineToText(outline, { maxLines: args.maxLines ?? 400 });
      return {
        text: `${args.filename}：${outline.count} ${outline.unit}\n\n${text}`,
        kind: outline.kind,
        count: outline.count,
        outline,
      };
    },
  }));

  // ---------------------------------------------------------------- list
  ctx.tools.register(defineTool({
    name: 'office_list',
    description: '列出当前会话工作区里的 Office 文档。',
    parameters: {},
    output: out(),
    isConcurrencySafe: () => true,
    async execute(_args, exec) {
      const sessionId = sessionOf(exec);
      const files = workspace.listFiles(sessionId, cwdOf(exec));
      if (files.length === 0) return { text: '会话工作区里还没有 Office 文档。', files };
      const lines = files.map((f) => `- ${f.name}（${f.type}，${(f.size / 1024).toFixed(1)} KB）`);
      return { text: `${files.length} 个文件：\n${lines.join('\n')}`, files };
    },
  }));

  // ---------------------------------------------------------------- guide
  ctx.tools.register(defineTool({
    name: 'office_guide',
    description: '生成 PPT 前先看：版式清单、主题/模板、字数红线、DeckSpec 示例。',
    parameters: {},
    output: out(),
    isConcurrencySafe: () => true,
    async execute() {
      return { text: guideText(), catalog: catalog() };
    },
  }));

  return { previewUrl };
}

/** 供路由层复用：把 outline 渲染成 HTML 预览。 */
export { previewHtml, previewHtmlForOutline, lintPages, resolveStyle, renderSlide, DARK_LAYOUTS };
