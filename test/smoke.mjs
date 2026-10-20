/**
 * 冒烟测试：编译一份 8 页 deck，校验 lint、ZIP 结构与 XML 良构性，并落盘产物。
 * 运行：node test/smoke.mjs
 */
import { writeFileSync, mkdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { compileDeck } from '../lib/pptx/deck.js';
import { unzipSync } from '../lib/util/zip.js';

const here = dirname(fileURLToPath(import.meta.url));
const outDir = join(here, '..', 'out');
mkdirSync(outDir, { recursive: true });

const spec = {
  template: 'product-launch',
  theme: 'tech-cyan',
  footer: 'DeepSeek Flash V4.1 · 模型简介',
  slides: [
    { layout: 'cover', eyebrow: '模型速览', title: 'DeepSeek Flash\nV4.1', subtitle: '轻量、高速、面向高并发的通用大模型 —— 能力与落地速览' },
    { layout: 'section', number: '01', title: '认识 Flash V4.1', subtitle: '它是什么，以及它擅长什么' },
    {
      layout: 'bullets',
      title: '核心特点：快、省、稳',
      columns: 2,
      items: [
        { title: '响应更快', desc: '轻量架构优先压低单次延迟，适合高频交互' },
        { title: '成本更低', desc: '单位调用成本可控，长尾请求易于规模化' },
        { title: '吞吐稳定', desc: '面向高并发优化，短请求场景表现均衡' },
        { title: '能力通用', desc: '覆盖写作、问答、代码、摘要等日常任务' },
        { title: '长文友好', desc: '适合处理较长文档与多轮对话上下文' },
        { title: '接入简单', desc: '接口风格贴近主流方案，迁移成本低' },
      ],
    },
    {
      layout: 'cards',
      title: '四个典型落地场景',
      cards: [
        { title: '智能客服', tag: '高频', desc: '高频短请求的问答与工单分流，响应快、单次成本可控。' },
        { title: '内容生产', tag: '批量', desc: '批量生成文案、摘要、标题与改写，适合流水线式生产。' },
        { title: '代码辅助', tag: '开发', desc: '代码补全、注释生成与片段重写，服务日常开发流程。' },
        { title: '数据抽取', tag: '结构化', desc: '从半结构化文本中抽取字段，大批量处理的性价比之选。' },
      ],
    },
    {
      layout: 'compare',
      title: '怎么选：Flash 还是 Pro',
      left: { title: 'Flash · 轻量高速', points: ['优先响应速度与成本', '适合高频、短请求任务', '大规模接入性价比高', '复杂推理交给 Pro 兜底'] },
      right: { title: 'Pro · 旗舰深度', points: ['优先复杂推理与深度思考', '适合长文分析与难题求解', '单次调用成本更高', '关键任务的质量保障'] },
    },
    {
      layout: 'steps',
      title: '五步完成接入',
      steps: [
        { title: '明确场景', desc: '判断是否强推理；高频短请求优先 Flash' },
        { title: '接入接口', desc: '按官方文档配置密钥、模型名与超时' },
        { title: '固化提示词', desc: '约束输出结构，减少多轮反复' },
        { title: '灰度验证', desc: '小流量对比效果与耗时，再全量' },
        { title: '持续观察', desc: '监控失败率与延迟，按需切换档位' },
      ],
    },
    {
      layout: 'table',
      title: '场景速查表',
      headers: ['场景', '建议档位', '理由', '关注点'],
      rows: [
        ['客服问答', 'Flash', '高频短请求，成本敏感', '响应速度'],
        ['长文总结', 'Flash / Pro', '视文档长度与深度要求', '上下文长度'],
        ['复杂推理', 'Pro', '需要多步思考与自检', '准确率'],
        ['批量抽取', 'Flash', '单条成本决定总体预算', '单位成本'],
        ['代码评审', 'Pro', '需要全局理解整个仓库', '上下文窗口'],
      ],
    },
    { layout: 'ending', title: '谢谢', subtitle: '具体参数与能力以官方模型文档为准' },
  ],
};

const compiled = compileDeck(spec, { title: 'DeepSeek Flash V4.1 简介' });
const pptxPath = join(outDir, 'smoke.pptx');
writeFileSync(pptxPath, compiled.pptx);
writeFileSync(join(outDir, 'smoke.preview.html'), compiled.html);

// —— 断言 ——
const problems = [];
if (compiled.pages.length !== spec.slides.length) problems.push(`页数不符：${compiled.pages.length}`);
if (compiled.lint.errors.length > 0) problems.push(`lint 错误 ${compiled.lint.errors.length} 条`);
if (compiled.pptx.length < 8000) problems.push(`pptx 只有 ${compiled.pptx.length} 字节，疑似为空壳`);

const zip = unzipSync(compiled.pptx);
const names = zip.names();
for (const required of ['[Content_Types].xml', 'ppt/presentation.xml', 'ppt/theme/theme1.xml', 'ppt/slides/slide8.xml']) {
  if (!names.includes(required)) problems.push(`缺少 part: ${required}`);
}
// 每个 slide part 都应有对应的 rels
for (let i = 1; i <= compiled.pages.length; i += 1) {
  if (!names.includes(`ppt/slides/_rels/slide${i}.xml.rels`)) problems.push(`slide${i} 缺 rels`);
}
// 粗查 XML 闭合（真的良构性由 pwsh 的 [xml] 校验）
for (const name of names.filter((x) => x.endsWith('.xml'))) {
  const text = zip.text(name);
  if (!text.startsWith('<?xml')) problems.push(`${name} 缺 XML 声明`);
  const open = (text.match(/<[^/!?][^>]*>/g) ?? []).length;
  const close = (text.match(/<\/[^>]+>/g) ?? []).length;
  const selfClose = (text.match(/<[^>]*\/>/g) ?? []).length;
  if (open !== close + selfClose) problems.push(`${name} 标签不平衡：open=${open} close=${close} self=${selfClose}`);
}

const sizeKb = (compiled.pptx.length / 1024).toFixed(1);
console.log(`pptx: ${pptxPath} (${sizeKb} KB, ${compiled.pages.length} 页)`);
console.log(`preview: ${join(outDir, 'smoke.preview.html')}`);
console.log(`lint: ${compiled.lint.errors.length} errors / ${compiled.lint.warnings.length} warnings`);
for (const w of compiled.lint.warnings.slice(0, 5)) console.log(`  ! ${w.code} 第${w.slide}页 ${w.message}`);
for (const e of compiled.lint.errors.slice(0, 5)) console.log(`  ✗ ${e.code} 第${e.slide}页 ${e.message}`);
console.log(`最小字号 ${compiled.lint.stats.smallestFont}pt · 最低对比度 ${compiled.lint.stats.minContrast.toFixed(2)}:1 · 形状 ${compiled.lint.stats.shapes}`);

/* ---------------------------------------------------------------- 样式覆盖 */

const styled = compileDeck({
  template: 'consulting',
  theme: 'tech-cyan',
  style: {
    footer: '自定义页脚',
    pageNumber: false,
    background: 'FFF8F0',
    transition: 'fade',
    meta: { title: '样式覆盖测试', author: 'QA', keywords: 'k1,k2', description: '覆盖说明' },
  },
  slides: [
    { layout: 'cover', title: '样式封面', subtitle: '副标题' },
    {
      layout: 'bullets',
      title: '要点',
      items: [{ title: '一条', desc: '说明' }],
      background: 'F0F7FF',
      transition: 'push',
    },
    { layout: 'cards', title: '未覆盖的浅色页', cards: [{ title: '卡片', desc: '说明' }] },
    { layout: 'ending', title: '完' },
  ],
});
writeFileSync(join(outDir, 'styled.pptx'), styled.pptx);

const pageText = (page) => page.shapes
  .filter((s) => s.k === 'text' && s.paras)
  .flatMap((s) => s.paras.map((p) => p.runs.map((r) => r.t).join('')))
  .join(' | ');

// 封面是深色整幅页：全局 background 只作用于浅色内容页，深色页保留主题深底
if (styled.pages[0].pageBg !== '07161C') problems.push(`深色封面底色被误改：${styled.pages[0].pageBg}`);
if (styled.pages[2].pageBg !== 'FFF8F0') problems.push(`style.background 未作用到浅色页：${styled.pages[2].pageBg}`);
if (styled.pages[1].pageBg !== 'F0F7FF') problems.push(`单页 background 未覆盖：${styled.pages[1].pageBg}`);
if (!pageText(styled.pages[1]).includes('自定义页脚')) problems.push('style.footer 未生效');
if (pageText(styled.pages[1]).includes('02')) problems.push('pageNumber:false 仍然画了页码');
if (styled.pages[1].transition !== 'push') problems.push(`单页 transition 未生效：${styled.pages[1].transition}`);
if (styled.pages[0].transition !== 'fade') problems.push(`全局 transition 未生效：${styled.pages[0].transition}`);

const styledZip = unzipSync(styled.pptx);
const coreXml = styledZip.text('docProps/core.xml');
if (!coreXml.includes('样式覆盖测试')) problems.push('style.meta.title 未写入 core.xml');
if (!coreXml.includes(`<cp:keywords>k1,k2</cp:keywords>`)) problems.push('style.meta.keywords 未写入 core.xml');
if (!coreXml.includes('覆盖说明')) problems.push('style.meta.description 未写入 core.xml');
if (!styledZip.text('ppt/slides/slide2.xml').includes('<p:transition')) problems.push('第 2 页缺 transition 标记');

const plain = compileDeck({
  slides: [{ layout: 'bullets', title: '仅一页', items: [{ title: 'a' }] }],
  style: { pageNumber: true },
});
const plainLabel = plain.pages[0].shapes.find((s) => s.label === '页码');
const plainText = plainLabel ? plainLabel.paras[0].runs[0].t : '(无)';
if (plainText !== '1') problems.push(`pageNumber:true 应输出 "1"，实际 "${plainText}"`);

let rejectedStyleKey = false;
try {
  compileDeck({ style: { pageNumer: true }, slides: [{ layout: 'ending' }] });
} catch (err) {
  rejectedStyleKey = /未知字段/.test(err.message);
}
if (!rejectedStyleKey) problems.push('未知 style 字段未报错');

console.log(`样式覆盖：深色封面 ${styled.pages[0].pageBg} · 浅色页 ${styled.pages[2].pageBg} · 单页覆盖 ${styled.pages[1].pageBg} · 页码 ${styled.style.pageNumberFormat === null ? '关闭' : '开启'} · meta ${coreXml.includes('QA') ? 'ok' : 'missing'}`);

if (problems.length > 0) {
  console.error('\nFAIL:');
  for (const p of problems) console.error(' - ' + p);
  process.exit(1);
}
console.log('\nOK');
