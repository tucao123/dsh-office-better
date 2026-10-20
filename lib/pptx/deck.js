/**
 * @module pptx/deck
 * DeckSpec -> IR 页 -> (pptx | 预览 HTML | lint)。
 * 编排层：唯一负责把"内容"变成"产物"的地方。
 */
import { parseDeckSpec, DeckSpecError } from './spec.js';
import { resolveStyle, CANVAS, THEMES, TEMPLATES } from './theme.js';
import { renderSlide, DARK_LAYOUTS, LAYOUT_IDS } from './layouts.js';
import { buildPptx } from './writer.js';
import { previewHtml } from './preview.js';
import { lintPages } from './lint.js';

export { DeckSpecError, LAYOUT_IDS };

/**
 * 编译一份 deck。
 * @param {any} spec DeckSpec（对象或 JSON 字符串）
 * @param {{footer?:string, title?:string, author?:string, transition?:string}} options
 * @returns {{deck:object, style:object, pages:Array<object>, lint:object, pptx:Buffer, html:string}}
 */
export function compileDeck(spec, options = {}) {
  const deck = parseDeckSpec(spec);
  const styleCfg = deck.style ?? {};
  const styleOptions = {
    template: deck.template,
    theme: deck.theme,
    footer: deck.footer ?? styleCfg.footer ?? options.footer ?? defaultFooter(deck),
    background: styleCfg.background,
    pageNumber: styleCfg.pageNumber,
  };
  const style = resolveStyle(styleOptions);
  const visible = deck.slides.filter((slide) => slide.hidden !== true);
  const transition = styleCfg.transition ?? options.transition;

  const pages = visible.map((slide, index) => {
    const dark = DARK_LAYOUTS.has(slide.layout);
    // 单页 background 优先于全局/主题底色；深色整幅页默认仍用自己的深色底。
    const pageBg = slide.background
      ?? (dark ? (slide.layout === 'section' ? style.darkBg2 : style.darkBg) : style.pageBg);
    const shapes = renderSlide(slide, { style, pageNo: index + 1, total: visible.length });
    return {
      layout: slide.layout,
      pageBg,
      transition: slide.transition ?? transition,
      shapes,
      title: pageTitleOf(slide),
    };
  });

  const lint = lintPages(pages);
  const metaCfg = styleCfg.meta ?? {};
  const meta = {
    title: metaCfg.title ?? options.title ?? pages[0]?.title ?? '演示文稿',
    author: metaCfg.author ?? options.author ?? 'dsh-office-better',
    subject: metaCfg.subject ?? (deck.template ? `template:${deck.template}` : undefined),
    keywords: metaCfg.keywords ?? (deck.theme ? `theme:${deck.theme}` : undefined),
    description: metaCfg.description,
  };
  const pptx = buildPptx({ pages, style, meta });
  const html = previewHtml({
    title: meta.title,
    pages,
    meta: { template: deck.template, theme: deck.theme },
  });
  return { deck, style, pages, lint, pptx, html };
}

/** 页脚默认文案：优先用封面标题，避免每页重复长标题。 */
function defaultFooter(deck) {
  const cover = deck.slides.find((slide) => slide.layout === 'cover');
  if (cover?.title) return String(cover.title).replace(/\n/g, ' ');
  const first = deck.slides[0];
  return first?.title ? String(first.title) : null;
}

function pageTitleOf(slide) {
  if (slide.layout === 'quote') return slide.quote;
  return slide.title ?? slide.quote ?? slide.layout;
}

/** 供工具层使用的简短报告。 */
export function summarize(compiled, fileName) {
  const { lint, pages, style } = compiled;
  const lines = [
    `文件：${fileName}`,
    `页数：${pages.length} 页 · 版式：${pages.map((p) => p.layout).join(', ')}`,
    `主题：${style.theme.name}（${style.theme.id}）${style.template ? ` · 模板：${style.template.name}` : ''}`,
    `画布：${CANVAS.w}x${CANVAS.h}pt（16:9）`,
    `检查：${lint.errors.length} 个错误，${lint.warnings.length} 个提醒`,
  ];
  for (const issue of lint.errors.slice(0, 8)) lines.push(`  ✗ [${issue.code}] 第 ${issue.slide} 页 ${issue.message}`);
  for (const issue of lint.warnings.slice(0, 8)) lines.push(`  ! [${issue.code}] 第 ${issue.slide} 页 ${issue.message}`);
  if (lint.errors.length > 8 || lint.warnings.length > 8) lines.push('  …（其余省略，可用 office_lint 查看完整列表）');
  return lines.join('\n');
}

/** 主题/模板清单（给 guide 工具用）。 */
export function catalog() {
  return {
    themes: Object.values(THEMES).map(({ id, name }) => ({ id, name })),
    templates: Object.values(TEMPLATES).map(({ id, name }) => ({ id, name })),
    layouts: LAYOUT_IDS,
  };
}
