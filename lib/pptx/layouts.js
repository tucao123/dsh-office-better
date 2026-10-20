/**
 * @module pptx/layouts
 * 15 种版式的布局引擎：输入「只含内容」的 slide 对象，输出**形状 IR**。
 *
 * IR 是本插件的单一真源：PPTX 写入器、SVG 预览、lint 全部消费同一份 IR。
 * 因此预览不会和成品不一致，溢出也不是"看截图才发现"，而是布局期就能算出来。
 */
import { CANVAS } from './theme.js';
import { wrap, measure } from '../ir/text.js';

const W = CANVAS.w;
const H = CANVAS.h;
const M = CANVAS.margin;
const CW = W - 2 * M;
const BODY_TOP = CANVAS.bodyTop;
const BODY_BOTTOM = CANVAS.bodyBottom;
const BODY_H = BODY_BOTTOM - BODY_TOP;
const GAP = CANVAS.gap;

/* ------------------------------------------------------------------ 基础形状 */

function box({ x, y, w, h, fill = null, line = null, lineWidth = 0.75, adj = null, radius = 'rect', fillAlpha, lineAlpha, role = 'shape', label = '' }) {
  const k = radius === 'round' ? 'round' : radius === 'ellipse' ? 'ellipse' : 'rect';
  return { k, x, y, w, h, fill, line, lineWidth, adj, fillAlpha, lineAlpha, role, label };
}

function seg(x, y, w, h, fill) {
  return box({ x, y, w, h, fill, role: 'accent', label: '分隔线' });
}

function dot(x, y, size, fill) {
  return box({ x, y, w: size, h: size, fill, radius: 'ellipse', role: 'accent', label: '标记点' });
}

/** 折行后自动缩放：给定最大高度，挑能整段塞下的最大字号。 */
export function flow(text, { size, min = 8.5, maxW, maxH = Infinity, lh = 1.35, bold = false }) {
  let s = size;
  for (;;) {
    const lines = wrap(text, s, maxW, bold);
    const height = lines.length * s * lh;
    if (height <= maxH || s <= min) return { size: s, lines, height };
    s = Math.max(min, Math.round((s - 0.5) * 10) / 10);
  }
}

/** 单行自适应：缩字号，实在放不下才截断（截断会被 lint 记为 warn）。 */
export function fitLine(text, { size, min = 9, maxW, bold = false }) {
  const raw = String(text ?? '');
  let s = size;
  while (s > min && measure(raw, s, bold) > maxW) s = Math.max(min, Math.round((s - 0.5) * 10) / 10);
  if (measure(raw, s, bold) <= maxW) return { text: raw, size: s, truncated: false };
  let out = '';
  for (const ch of raw) {
    if (measure(`${out}${ch}…`, s, bold) > maxW) break;
    out += ch;
  }
  return { text: `${out}…`, size: s, truncated: true };
}

/**
 * 文本块形状。
 * @param {{x:number,y:number,w:number,lines:string[]|Array<{runs:object[],align?:string}>,
 *          size:number,color:string,bold?:boolean,align?:string,lh?:number,anchor?:string,
 *          label?:string,role?:string,h?:number,spaceBefore?:number}} p
 */
function textBlock(p) {
  const {
    x, y, w, lines, size, color, bold = false, align = 'l', lh = 1.35,
    anchor = 't', label = '', role = 'body', h, spaceBefore = 0,
  } = p;
  const paras = lines.map((line, i) => {
    if (typeof line === 'string') {
      return { runs: [{ t: line, size, bold, color }], align, lineHeight: lh, spaceBefore: i === 0 ? spaceBefore : 0 };
    }
    return {
      runs: line.runs,
      align: line.align ?? align,
      lineHeight: lh,
      spaceBefore: i === 0 ? spaceBefore : 0,
    };
  });
  const contentH = lines.reduce((sum, _l, i) => sum + size * lh + (i === 0 ? spaceBefore : 0), 0);
  // 文本形状没有填充/描边，高度只是容纳用的框：给 8% + 1.2pt 余量，
  // 用来吸收 PowerPoint 在行距之外追加的首行上沿/末行下沿（实测约 0.8pt/块）。
  const boxH = Math.max(h ?? 0, Math.ceil(contentH * 1.08 + 1.2));
  return {
    k: 'text', x, y, w, h: boxH, anchor, paras, label, role, size, color, contentH,
  };
}

/* ------------------------------------------------------------------ 页面外衣 */

/** 页眉：标题 + 强调短条 + 分隔线。 */
function header(title, style) {
  const shapes = [];
  if (style.topBar) shapes.push(box({ x: 0, y: 0, w: W, h: 6, fill: style.accent }));
  if (style.corner) shapes.push(box({ x: W - 120, y: 0, w: 120, h: 120, fill: style.accentSoft, radius: 'ellipse', fillAlpha: 0.55, role: 'deco' }));
  shapes.push(box({ x: M, y: 56, w: 5.5, h: 26, fill: style.accent, role: 'accent' }));
  const fit = fitLine(title, { size: 26, min: 17, maxW: CW - 28, bold: true });
  shapes.push(textBlock({
    x: M + 20, y: 44, w: CW - 20, h: 40, lines: [fit.text], size: fit.size,
    color: style.ink.t1, bold: true, anchor: 'ctr', lh: 1.2, label: '标题', role: 'title',
  }));
  shapes.push(seg(M, CANVAS.dividerY, CW, 1, style.border));
  return shapes;
}

/** 页脚 + 页码。 */
function chrome(style, pageNo, total, dark = false) {
  const shapes = [];
  const ink = dark ? style.darkInk : style.ink;
  if (style.railWidth > 0 && !dark) {
    shapes.push(box({ x: 0, y: 0, w: style.railWidth, h: H, fill: style.accent, role: 'deco' }));
  }
  if (style.footer) {
    const fit = fitLine(style.footer, { size: 9.5, min: 8, maxW: 520 });
    shapes.push(textBlock({
      x: M, y: CANVAS.footerY, w: 520, h: 14, lines: [fit.text], size: fit.size,
      color: ink.t2, label: '页脚', role: 'caption',
    }));
  }
  if (style.pageNumberFormat) {
    shapes.push(textBlock({
      x: W - M - 140, y: CANVAS.footerY, w: 140, h: 14,
      lines: [style.pageNumberFormat(pageNo, total)], size: 9.5, color: ink.t2,
      align: 'r', label: '页码', role: 'caption',
    }));
  }
  return shapes;
}

/* ------------------------------------------------------------------ 版式实现 */

/** 封面：深色整幅，标题块 + 副标题 + 装饰椭圆。 */
function cover(slide, ctx) {
  const s = ctx.style;
  const ink = s.darkInk;
  const shapes = [
    box({ x: 0, y: 0, w: W, h: H, fill: s.darkBg, role: 'bg' }),
    box({ x: 690, y: 80, w: 320, h: 320, line: s.accent2, lineWidth: 1, lineAlpha: 0.22, radius: 'ellipse', role: 'deco' }),
    box({ x: 570, y: 190, w: 440, h: 440, line: s.accent2, lineWidth: 0.75, lineAlpha: 0.12, radius: 'ellipse', role: 'deco' }),
  ];
  const titleRaw = String(slide.title ?? '').split('\n').map((t) => t.trim()).filter(Boolean);
  if (slide.eyebrow) {
    shapes.push(box({ x: M, y: 172, w: 54, h: 4, fill: s.accent2, role: 'accent' }));
    shapes.push(textBlock({
      x: M, y: 148, w: 420, h: 18, lines: [String(slide.eyebrow)], size: 11.5,
      color: s.accentDark, bold: true, label: '眉标', role: 'caption',
    }));
  }
  let y = 206;
  if (titleRaw.length > 0) {
    const isVersion = (t) => /^v?\d+(\.\d+)*$/i.test(t);
    const size = titleRaw.length >= 3 ? 44 : 52;
    const lines = titleRaw.map((t) => ({
      runs: [{ t, size, bold: true, color: isVersion(t) ? s.accentDark : ink.t1 }],
    }));
    const blockH = titleRaw.length * size * 1.12;
    shapes.push(textBlock({
      x: M, y, w: 640, h: blockH, lines, size, color: ink.t1, lh: 1.12,
      label: '主标题', role: 'title',
    }));
    y += blockH + 26;
  }
  if (slide.subtitle) {
    const fit = flow(slide.subtitle, { size: 16.5, min: 12, maxW: 600, maxH: 70, lh: 1.45 });
    shapes.push(textBlock({
      x: M, y, w: 600, h: fit.height, lines: fit.lines, size: fit.size,
      color: ink.t2, lh: 1.45, label: '副标题', role: 'body',
    }));
    y += fit.height + 22;
  }
  if (slide.meta) {
    shapes.push(textBlock({
      x: M, y, w: 600, h: 16, lines: [String(slide.meta)], size: 11, color: ink.t2, label: '署名', role: 'caption',
    }));
  }
  shapes.push(seg(M, 438, 280, 1, ink.t3));
  return shapes;
}

/** 章节页：编号 + 标题 + 副标题。 */
function section(slide, ctx) {
  const s = ctx.style;
  const ink = s.darkInk;
  const shapes = [box({ x: 0, y: 0, w: W, h: H, fill: s.darkBg2, role: 'bg' })];
  if (slide.number != null) {
    shapes.push(textBlock({
      x: M, y: 92, w: 320, h: 116, lines: [String(slide.number)], size: 88,
      color: s.accentDark, bold: true, lh: 1.1, label: '章节号', role: 'title',
    }));
  }
  const fit = fitLine(slide.title ?? '', { size: 36, min: 20, maxW: 640, bold: true });
  shapes.push(textBlock({
    x: M, y: 240, w: 660, h: 52, lines: [fit.text], size: fit.size, color: ink.t1,
    bold: true, lh: 1.2, label: '章节标题', role: 'title',
  }));
  shapes.push(seg(M, 308, 58, 3, s.accent2));
  if (slide.subtitle) {
    const sub = flow(slide.subtitle, { size: 15, min: 11, maxW: 620, maxH: 60, lh: 1.4 });
    shapes.push(textBlock({
      x: M, y: 328, w: 620, h: sub.height, lines: sub.lines, size: sub.size,
      color: ink.t2, lh: 1.4, label: '章节副标题', role: 'body',
    }));
  }
  return shapes;
}

/** 要点：1-2 栏，可 1-8 条。 */
function bullets(slide, ctx) {
  const s = ctx.style;
  const items = slide.items ?? [];
  const cols = slide.columns === 2 || items.length > 4 ? 2 : 1;
  const rows = Math.ceil(items.length / cols);
  const colW = (CW - (cols - 1) * 32) / cols;
  const rowH = (BODY_H - GAP * (rows - 1)) / rows;
  const shapes = header(slide.title ?? '要点', s);
  items.forEach((item, i) => {
    const cx = M + (i % cols) * (colW + 32);
    const cy = BODY_TOP + Math.floor(i / cols) * (rowH + GAP);
    shapes.push(dot(cx, cy + 8, 6, s.accent));
    const tf = fitLine(item.title ?? '', { size: 16.5, min: 11, maxW: colW - 20, bold: true });
    shapes.push(textBlock({
      x: cx + 18, y: cy, w: colW - 18, h: tf.size * 1.3, lines: [tf.text], size: tf.size,
      color: s.ink.t1, bold: true, lh: 1.3, label: `要点 ${i + 1} 标题`, role: 'heading',
    }));
    if (item.desc) {
      const df = flow(item.desc, {
        size: 12.5, min: 9, maxW: colW - 18, maxH: Math.max(20, rowH - tf.size * 1.3 - 8), lh: 1.4,
      });
      shapes.push(textBlock({
        x: cx + 18, y: cy + tf.size * 1.34, w: colW - 18, h: df.height, lines: df.lines,
        size: df.size, color: s.ink.t2, lh: 1.4, label: `要点 ${i + 1} 说明`, role: 'body',
      }));
    }
  });
  return shapes;
}

/** 卡片：1-6 张，列数自动（≤3 单行、4 张 2x2、5-6 张 3 列）。 */
function cards(slide, ctx) {
  const s = ctx.style;
  const list = slide.cards ?? [];
  const cols = slide.columns ?? (list.length <= 3 ? list.length : list.length === 4 ? 2 : 3);
  const rows = Math.ceil(list.length / cols);
  const cardW = (CW - GAP * (cols - 1)) / cols;
  const cardH = (BODY_H - GAP * (rows - 1)) / rows;
  const shapes = header(slide.title ?? '卡片', s);
  list.forEach((card, i) => {
    const cx = M + (i % cols) * (cardW + GAP);
    const cy = BODY_TOP + Math.floor(i / cols) * (cardH + GAP);
    shapes.push(box({
      x: cx, y: cy, w: cardW, h: cardH, fill: s.surface, line: s.border, adj: 0.05,
      radius: 'round', role: 'card', label: `卡片 ${i + 1}`,
    }));
    shapes.push(seg(cx + 20, cy + 20, 22, 3, s.accent));
    if (card.tag) {
      const tag = fitLine(card.tag, { size: 10, min: 8, maxW: cardW - 60 });
      shapes.push(textBlock({
        x: cx + 20, y: cy + 18, w: cardW - 40, h: 14, lines: [tag.text],
        size: tag.size, color: s.accent, align: 'r', label: `卡片 ${i + 1} 标签`, role: 'caption',
      }));
    }
    const tf = fitLine(card.title ?? '', { size: 16, min: 10.5, maxW: cardW - 40, bold: true });
    shapes.push(textBlock({
      x: cx + 20, y: cy + 32, w: cardW - 40, h: tf.size * 1.3, lines: [tf.text], size: tf.size,
      color: s.ink.t1, bold: true, lh: 1.3, label: `卡片 ${i + 1} 标题`, role: 'heading',
    }));
    if (card.desc) {
      const df = flow(card.desc, {
        size: 11.8, min: 8.5, maxW: cardW - 40, maxH: Math.max(18, cardH - 32 - tf.size * 1.34 - 14), lh: 1.42,
      });
      shapes.push(textBlock({
        x: cx + 20, y: cy + 36 + tf.size * 1.34, w: cardW - 40, h: df.height, lines: df.lines,
        size: df.size, color: s.ink.t2, lh: 1.42, label: `卡片 ${i + 1} 说明`, role: 'body',
      }));
    }
  });
  return shapes;
}

/** 数字锚点：1-4 个巨型数字。 */
function kpi(slide, ctx) {
  const s = ctx.style;
  const metrics = (slide.metrics ?? []).slice(0, 4);
  const n = Math.max(1, metrics.length);
  const colW = (CW - GAP * (n - 1)) / n;
  const shapes = header(slide.title ?? '关键指标', s);
  const cardH = Math.min(190, BODY_H - 10);
  metrics.forEach((metric, i) => {
    const cx = M + i * (colW + GAP);
    shapes.push(box({
      x: cx, y: BODY_TOP + 6, w: colW, h: cardH, fill: s.surface, line: s.border,
      adj: 0.05, radius: 'round', role: 'card', label: `指标 ${i + 1}`,
    }));
    shapes.push(seg(cx, BODY_TOP + 6, colW, 3, s.accent));
    const vf = fitLine(metric.value ?? '', { size: 46, min: 20, maxW: colW - 32, bold: true });
    shapes.push(textBlock({
      x: cx + 16, y: BODY_TOP + 34, w: colW - 32, h: vf.size * 1.2, lines: [vf.text],
      size: vf.size, color: s.accent, bold: true, lh: 1.2, label: `指标 ${i + 1} 数值`, role: 'title',
    }));
    const lf = flow(metric.label ?? '', { size: 12.5, min: 9, maxW: colW - 32, maxH: 40, lh: 1.35 });
    shapes.push(textBlock({
      x: cx + 16, y: BODY_TOP + 40 + vf.size * 1.24, w: colW - 32, h: lf.height,
      lines: lf.lines, size: lf.size, color: s.ink.t1, bold: true, lh: 1.35,
      label: `指标 ${i + 1} 名称`, role: 'heading',
    }));
    if (metric.note) {
      const nf = flow(metric.note, { size: 10.5, min: 8.5, maxW: colW - 32, maxH: 40, lh: 1.4 });
      shapes.push(textBlock({
        x: cx + 16, y: BODY_TOP + 44 + vf.size * 1.24 + lf.height, w: colW - 32, h: nf.height,
        lines: nf.lines, size: nf.size, color: s.ink.t2, lh: 1.4, label: `指标 ${i + 1} 注释`, role: 'caption',
      }));
    }
  });
  return shapes;
}

/** 步骤：1-5 步，横向卡片 + 连接符。 */
function steps(slide, ctx) {
  const s = ctx.style;
  const list = (slide.steps ?? []).slice(0, 5);
  const n = Math.max(1, list.length);
  const cardW = (CW - GAP * (n - 1)) / n;
  const shapes = header(slide.title ?? '实施路径', s);
  const cardH = Math.min(196, BODY_H - 12);
  list.forEach((step, i) => {
    const cx = M + i * (cardW + GAP);
    const cy = BODY_TOP + 8;
    shapes.push(box({
      x: cx, y: cy, w: cardW, h: cardH, fill: s.surface, line: s.border, adj: 0.06,
      radius: 'round', role: 'card', label: `步骤 ${i + 1}`,
    }));
    shapes.push(textBlock({
      x: cx + 16, y: cy + 16, w: cardW - 32, h: 26, lines: [String(i + 1).padStart(2, '0')],
      size: 19, color: s.accent, bold: true, lh: 1.2, label: `步骤 ${i + 1} 序号`, role: 'caption',
    }));
    const tf = fitLine(step.title ?? '', { size: 13.5, min: 10, maxW: cardW - 32, bold: true });
    shapes.push(textBlock({
      x: cx + 16, y: cy + 52, w: cardW - 32, h: tf.size * 1.3, lines: [tf.text], size: tf.size,
      color: s.ink.t1, bold: true, lh: 1.3, label: `步骤 ${i + 1} 标题`, role: 'heading',
    }));
    if (step.desc) {
      const df = flow(step.desc, {
        size: 10.8, min: 8.5, maxW: cardW - 32, maxH: Math.max(18, cardH - 56 - tf.size * 1.34 - 12), lh: 1.42,
      });
      shapes.push(textBlock({
        x: cx + 16, y: cy + 56 + tf.size * 1.34, w: cardW - 32, h: df.height, lines: df.lines,
        size: df.size, color: s.ink.t2, lh: 1.42, label: `步骤 ${i + 1} 说明`, role: 'body',
      }));
    }
    if (i < list.length - 1) {
      shapes.push(seg(cx + cardW + 3, cy + cardH / 2, GAP - 6, 1.5, s.ink.t3));
    }
  });
  return shapes;
}

/** 对比：左右两栏要点。 */
function compare(slide, ctx) {
  const s = ctx.style;
  const shapes = header(slide.title ?? '对比', s);
  const sides = [
    { data: slide.left ?? {}, tint: s.accent },
    { data: slide.right ?? {}, tint: s.ink.t2 },
  ];
  const panelW = (CW - GAP) / 2;
  const panelH = Math.min(BODY_H - 8, 316);
  sides.forEach(({ data, tint }, i) => {
    const px = M + i * (panelW + GAP);
    const py = BODY_TOP + 4;
    shapes.push(box({
      x: px, y: py, w: panelW, h: panelH, fill: s.surface, line: s.border, adj: 0.05,
      radius: 'round', role: 'card', label: `对比栏 ${i + 1}`,
    }));
    shapes.push(box({ x: px, y: py, w: panelW, h: 4, fill: tint, adj: 0, role: 'accent' }));
    const tf = fitLine(data.title ?? '', { size: 16, min: 11, maxW: panelW - 48, bold: true });
    shapes.push(textBlock({
      x: px + 24, y: py + 22, w: panelW - 48, h: tf.size * 1.3, lines: [tf.text], size: tf.size,
      color: tint, bold: true, lh: 1.3, label: `对比栏 ${i + 1} 标题`, role: 'heading',
    }));
    shapes.push(seg(px + 24, py + 60, panelW - 48, 1, s.border));
    const points = (data.points ?? []).slice(0, 6);
    const rowH = Math.max(28, (panelH - 78) / Math.max(1, points.length));
    points.forEach((point, j) => {
      const yy = py + 74 + j * rowH;
      shapes.push(dot(px + 24, yy + 6, 5, tint));
      const pf = flow(point, {
        size: 13, min: 9, maxW: panelW - 58, maxH: rowH - 4, lh: 1.35,
      });
      shapes.push(textBlock({
        x: px + 38, y: yy, w: panelW - 58, h: pf.height, lines: pf.lines, size: pf.size,
        color: s.ink.t1, lh: 1.35, label: `对比栏 ${i + 1} 第 ${j + 1} 条`, role: 'body',
      }));
    });
  });
  return shapes;
}

/** 时间线：1-5 个横向节点。 */
function timeline(slide, ctx) {
  const s = ctx.style;
  const events = (slide.events ?? []).slice(0, 5);
  const n = Math.max(1, events.length);
  const colW = CW / n;
  const axisY = BODY_TOP + 96;
  const shapes = header(slide.title ?? '时间线', s);
  shapes.push(seg(M + 10, axisY, CW - 20, 2, s.border));
  events.forEach((event, i) => {
    const cx = M + i * colW + colW / 2;
    shapes.push(dot(cx - 6, axisY - 5, 12, s.accent));
    const dfit = fitLine(event.date ?? '', { size: 11, min: 8.5, maxW: colW - 24, bold: true });
    shapes.push(textBlock({
      x: cx - colW / 2 + 8, y: axisY - 34, w: colW - 16, h: 16, lines: [dfit.text], size: dfit.size,
      color: s.accent, bold: true, align: 'ctr', label: `节点 ${i + 1} 时间`, role: 'caption',
    }));
    const tf = flow(event.title ?? '', { size: 13, min: 9.5, maxW: colW - 24, maxH: 34, lh: 1.3, bold: true });
    shapes.push(textBlock({
      x: cx - colW / 2 + 8, y: axisY + 18, w: colW - 16, h: tf.height, lines: tf.lines, size: tf.size,
      color: s.ink.t1, bold: true, align: 'ctr', lh: 1.3, label: `节点 ${i + 1} 标题`, role: 'heading',
    }));
    if (event.desc) {
      const ef = flow(event.desc, { size: 10.8, min: 8.5, maxW: colW - 24, maxH: 76, lh: 1.4 });
      shapes.push(textBlock({
        x: cx - colW / 2 + 8, y: axisY + 24 + tf.height, w: colW - 16, h: ef.height,
        lines: ef.lines, size: ef.size, color: s.ink.t2, align: 'ctr', lh: 1.4,
        label: `节点 ${i + 1} 说明`, role: 'body',
      }));
    }
  });
  return shapes;
}

/** 引文：一句原话 + 出处。 */
function quote(slide, ctx) {
  const s = ctx.style;
  const shapes = [];
  shapes.push(textBlock({
    x: M - 6, y: 118, w: 120, h: 90, lines: ['“'], size: 96, color: s.accent, bold: true,
    lh: 1.0, label: '引号装饰', role: 'deco',
  }));
  const qf = flow(slide.quote ?? '', { size: 24, min: 13, maxW: CW - 80, maxH: 200, lh: 1.45 });
  shapes.push(textBlock({
    x: M + 40, y: 196, w: CW - 80, h: qf.height, lines: qf.lines, size: qf.size,
    color: s.ink.t1, lh: 1.45, label: '引文', role: 'title',
  }));
  if (slide.author) {
    shapes.push(seg(M + 40, 196 + qf.height + 24, 44, 2, s.accent));
    const af = fitLine(`—— ${slide.author}${slide.role ? `，${slide.role}` : ''}`, {
      size: 12.5, min: 9, maxW: CW - 80,
    });
    shapes.push(textBlock({
      x: M + 40, y: 196 + qf.height + 36, w: CW - 80, h: 18, lines: [af.text], size: af.size,
      color: s.ink.t2, label: '引文出处', role: 'caption',
    }));
  }
  return shapes;
}

/** 表格：列宽按内容测量分配，行高自适应。 */
function table(slide, ctx) {
  const s = ctx.style;
  const headers = (slide.headers ?? []).slice(0, 5);
  const rows = (slide.rows ?? []).slice(0, 8);
  const shapes = header(slide.title ?? '明细', s);
  const cols = headers.length || Math.max(1, ...rows.map((r) => r.length));
  const pad = 16;
  // 按内容自然宽度分配列宽（自动排版里最容易崩、也最值得算清楚的一步）
  const natural = [];
  for (let c = 0; c < cols; c += 1) {
    let widest = measure(String(headers[c] ?? ''), 11.5, true);
    for (const row of rows) widest = Math.max(widest, measure(String(row[c] ?? ''), 12));
    natural.push(Math.max(64, Math.min(340, widest + pad * 2 + 4)));
  }
  const sum = natural.reduce((a, b) => a + b, 0);
  const scale = sum > CW ? CW / sum : 1;
  const widths = natural.map((v) => v * scale);
  const leftover = CW - widths.reduce((a, b) => a + b, 0);
  if (leftover > 0) widths[widths.length - 1] += leftover;

  const headH = 34;
  const rowH = Math.min(46, (BODY_H - headH) / Math.max(1, rows.length));
  const totalH = headH + rowH * rows.length;
  const ty = BODY_TOP + Math.max(0, Math.min(18, (BODY_H - totalH) / 2));

  shapes.push(box({ x: M, y: ty, w: CW, h: headH, fill: s.darkBg2, adj: 0, role: 'card', label: '表头' }));
  let x = M;
  headers.forEach((cell, c) => {
    const f = fitLine(cell, { size: 11.5, min: 8.5, maxW: widths[c] - pad * 2, bold: true });
    shapes.push(textBlock({
      x: x + pad, y: ty, w: widths[c] - pad * 2, h: headH, lines: [f.text], size: f.size,
      color: 'FFFFFF', bold: true, anchor: 'ctr', label: `表头 ${c + 1}`, role: 'heading',
    }));
    x += widths[c];
  });
  rows.forEach((row, r) => {
    const ry = ty + headH + r * rowH;
    shapes.push(box({
      x: M, y: ry, w: CW, h: rowH, fill: r % 2 === 0 ? s.surface : s.pageBg, adj: 0, role: 'card', label: `第 ${r + 1} 行`,
    }));
    shapes.push(seg(M, ry + rowH - 0.8, CW, 0.8, s.border));
    let cx = M;
    for (let c = 0; c < cols; c += 1) {
      const cell = String(row[c] ?? '');
      const bold = c === 0;
      const f = fitLine(cell, { size: 12, min: 8.5, maxW: widths[c] - pad * 2, bold });
      shapes.push(textBlock({
        x: cx + pad, y: ry, w: widths[c] - pad * 2, h: rowH, lines: [f.text], size: f.size,
        color: bold ? s.ink.t1 : s.ink.t2, bold, anchor: 'ctr',
        label: `第 ${r + 1} 行第 ${c + 1} 列`, role: bold ? 'heading' : 'body',
      }));
      cx += widths[c];
    }
  });
  return shapes;
}

/** 目录：1-6 个章节。 */
function agenda(slide, ctx) {
  const s = ctx.style;
  const items = (slide.items ?? []).slice(0, 6);
  const shapes = header(slide.title ?? '目录', s);
  const rowH = (BODY_H - GAP * (items.length - 1)) / Math.max(1, items.length);
  items.forEach((item, i) => {
    const y = BODY_TOP + i * (rowH + GAP);
    shapes.push(textBlock({
      x: M, y, w: 44, h: rowH, lines: [String(i + 1).padStart(2, '0')], size: 20,
      color: s.accent, bold: true, anchor: 'ctr', lh: 1.2, label: `目录 ${i + 1} 序号`, role: 'caption',
    }));
    const tf = fitLine(item.title ?? '', { size: 15, min: 10.5, maxW: CW - 80, bold: true });
    shapes.push(textBlock({
      x: M + 52, y: item.desc ? y + 4 : y, w: CW - 80, h: tf.size * 1.3, lines: [tf.text],
      size: tf.size, color: s.ink.t1, bold: true, lh: 1.3, label: `目录 ${i + 1} 标题`, role: 'heading',
    }));
    if (item.desc) {
      const df = fitLine(item.desc, { size: 11.5, min: 8.5, maxW: CW - 80 });
      shapes.push(textBlock({
        x: M + 52, y: y + 8 + tf.size * 1.34, w: CW - 80, h: 15, lines: [df.text], size: df.size,
        color: s.ink.t2, label: `目录 ${i + 1} 说明`, role: 'body',
      }));
    }
    if (i < items.length - 1) shapes.push(seg(M + 52, y + rowH + GAP / 2 - 0.5, CW - 52, 1, s.border));
  });
  return shapes;
}

/** SWOT：四象限。 */
function swot(slide, ctx) {
  const s = ctx.style;
  const quads = [
    { key: 's', label: '优势 S', tint: s.accent },
    { key: 'w', label: '劣势 W', tint: s.ink.t2 },
    { key: 'o', label: '机会 O', tint: s.accent2 },
    { key: 't', label: '威胁 T', tint: s.ink.t2 },
  ];
  const shapes = header(slide.title ?? 'SWOT', s);
  const qW = (CW - GAP) / 2;
  const qH = (BODY_H - GAP) / 2;
  quads.forEach((quad, i) => {
    const qx = M + (i % 2) * (qW + GAP);
    const qy = BODY_TOP + Math.floor(i / 2) * (qH + GAP);
    shapes.push(box({
      x: qx, y: qy, w: qW, h: qH, fill: s.surface, line: s.border, adj: 0.05,
      radius: 'round', role: 'card', label: quad.label,
    }));
    shapes.push(box({ x: qx, y: qy, w: 4, h: qH, fill: quad.tint, adj: 0, role: 'accent' }));
    shapes.push(textBlock({
      x: qx + 20, y: qy + 16, w: qW - 40, h: 18, lines: [quad.label], size: 12,
      color: quad.tint, bold: true, label: `${quad.label} 标签`, role: 'heading',
    }));
    const list = (slide[quad.key] ?? []).slice(0, 4);
    list.forEach((entry, j) => {
      const yy = qy + 42 + j * 22;
      shapes.push(dot(qx + 20, yy + 5, 4.5, quad.tint));
      const f = fitLine(entry, { size: 11.5, min: 8.5, maxW: qW - 56 });
      shapes.push(textBlock({
        x: qx + 32, y: yy, w: qW - 52, h: 16, lines: [f.text], size: f.size,
        color: s.ink.t2, label: `${quad.label} 第 ${j + 1} 条`, role: 'body',
      }));
    });
  });
  return shapes;
}

/** 报价/方案对比：1-4 个方案列。 */
function pricing(slide, ctx) {
  const s = ctx.style;
  const plans = (slide.plans ?? []).slice(0, 4);
  const n = Math.max(1, plans.length);
  const planW = (CW - GAP * (n - 1)) / n;
  const shapes = header(slide.title ?? '方案', s);
  const cardH = Math.min(BODY_H - 10, 300);
  plans.forEach((plan, i) => {
    const px = M + i * (planW + GAP);
    const py = BODY_TOP + 6;
    const hot = plan.highlight === true;
    shapes.push(box({
      x: px, y: py, w: planW, h: cardH, fill: hot ? s.accentSoft : s.surface,
      line: hot ? s.accent : s.border, lineWidth: hot ? 1.25 : 0.75, adj: 0.05,
      radius: 'round', role: 'card', label: `方案 ${i + 1}`,
    }));
    const nf = fitLine(plan.name ?? '', { size: 14, min: 10, maxW: planW - 40, bold: true });
    shapes.push(textBlock({
      x: px + 20, y: py + 20, w: planW - 40, h: 20, lines: [nf.text], size: nf.size,
      color: hot ? s.accent : s.ink.t1, bold: true, label: `方案 ${i + 1} 名称`, role: 'heading',
    }));
    if (plan.price) {
      const prf = fitLine(plan.price, { size: 28, min: 13, maxW: planW - 40, bold: true });
      shapes.push(textBlock({
        x: px + 20, y: py + 48, w: planW - 40, h: 38, lines: [prf.text], size: prf.size,
        color: hot ? s.accent : s.ink.t1, bold: true, lh: 1.2, label: `方案 ${i + 1} 价格`, role: 'title',
      }));
    }
    if (plan.tag) {
      const tag = fitLine(plan.tag, { size: 10, min: 8, maxW: planW - 40 });
      shapes.push(textBlock({
        x: px + 20, y: py + 90, w: planW - 40, h: 14, lines: [tag.text], size: tag.size,
        color: s.ink.t2, label: `方案 ${i + 1} 标注`, role: 'caption',
      }));
    }
    shapes.push(seg(px + 20, py + 112, planW - 40, 1, s.border));
    const features = (plan.features ?? []).slice(0, 5);
    features.forEach((feature, j) => {
      const yy = py + 126 + j * 26;
      shapes.push(dot(px + 20, yy + 5, 4, hot ? s.accent : s.ink.t2));
      const f = fitLine(feature, { size: 11.2, min: 8.5, maxW: planW - 56 });
      shapes.push(textBlock({
        x: px + 32, y: yy, w: planW - 52, h: 16, lines: [f.text], size: f.size,
        color: s.ink.t2, label: `方案 ${i + 1} 第 ${j + 1} 项`, role: 'body',
      }));
    });
  });
  return shapes;
}

/** 路线图：1-5 个阶段。 */
function roadmap(slide, ctx) {
  const s = ctx.style;
  const phases = (slide.phases ?? []).slice(0, 5);
  const n = Math.max(1, phases.length);
  const cardW = (CW - GAP * (n - 1)) / n;
  const shapes = header(slide.title ?? '路线图', s);
  const barY = BODY_TOP + 10;
  phases.forEach((_phase, i) => {
    const cx = M + i * (cardW + GAP);
    shapes.push(box({
      x: cx, y: barY, w: cardW, h: 5, fill: s.accent, adj: 0,
      fillAlpha: 0.35 + (i / Math.max(1, n - 1)) * 0.65, role: 'accent',
    }));
  });
  const cardH = Math.min(206, BODY_H - 40);
  phases.forEach((phase, i) => {
    const cx = M + i * (cardW + GAP);
    const cy = barY + 22;
    shapes.push(box({
      x: cx, y: cy, w: cardW, h: cardH, fill: s.surface, line: s.border, adj: 0.05,
      radius: 'round', role: 'card', label: `阶段 ${i + 1}`,
    }));
    const pf = fitLine(phase.phase ?? `阶段 ${i + 1}`, { size: 10.5, min: 8, maxW: cardW - 32, bold: true });
    shapes.push(textBlock({
      x: cx + 16, y: cy + 16, w: cardW - 32, h: 14, lines: [pf.text], size: pf.size,
      color: s.accent, bold: true, label: `阶段 ${i + 1} 标签`, role: 'caption',
    }));
    const tf = fitLine(phase.title ?? '', { size: 13, min: 9.5, maxW: cardW - 32, bold: true });
    shapes.push(textBlock({
      x: cx + 16, y: cy + 38, w: cardW - 32, h: tf.size * 1.3, lines: [tf.text], size: tf.size,
      color: s.ink.t1, bold: true, lh: 1.3, label: `阶段 ${i + 1} 标题`, role: 'heading',
    }));
    if (phase.desc) {
      const df = flow(phase.desc, {
        size: 10.8, min: 8.5, maxW: cardW - 32, maxH: Math.max(18, cardH - 42 - tf.size * 1.34 - 12), lh: 1.4,
      });
      shapes.push(textBlock({
        x: cx + 16, y: cy + 42 + tf.size * 1.34, w: cardW - 32, h: df.height, lines: df.lines,
        size: df.size, color: s.ink.t2, lh: 1.4, label: `阶段 ${i + 1} 说明`, role: 'body',
      }));
    }
  });
  return shapes;
}

/** 结尾：深色整幅致谢。 */
function ending(slide, ctx) {
  const s = ctx.style;
  const ink = s.darkInk;
  const shapes = [
    box({ x: 0, y: 0, w: W, h: H, fill: s.darkBg, role: 'bg' }),
    box({ x: 300, y: 150, w: 360, h: 250, line: s.accent2, lineWidth: 0.75, lineAlpha: 0.16, radius: 'ellipse', role: 'deco' }),
  ];
  const title = slide.title ?? '谢谢';
  shapes.push(textBlock({
    x: M, y: 208, w: CW, h: 56, lines: [title], size: 42, color: ink.t1, bold: true,
    align: 'ctr', lh: 1.2, label: '结束语', role: 'title',
  }));
  if (slide.subtitle) {
    const f = flow(slide.subtitle, { size: 13.5, min: 9.5, maxW: CW - 120, maxH: 60, lh: 1.4 });
    shapes.push(textBlock({
      x: M + 60, y: 282, w: CW - 120, h: f.height, lines: f.lines, size: f.size,
      color: ink.t2, align: 'ctr', lh: 1.4, label: '结束副标题', role: 'body',
    }));
  }
  return shapes;
}

export const LAYOUTS = {
  cover, section, bullets, cards, kpi, steps, compare, timeline,
  quote, table, agenda, swot, pricing, roadmap, ending,
};

export const LAYOUT_IDS = Object.keys(LAYOUTS);

/** 深色整幅版式：它们自己画背景，且用深色墨。 */
export const DARK_LAYOUTS = new Set(['cover', 'section', 'ending']);

/**
 * 渲染一页为 IR。
 * @param {{layout:string}} slide 已校验的 slide
 * @param {{style:object, pageNo:number, total:number}} ctx
 * @returns {Array<object>} shapes
 */
export function renderSlide(slide, ctx) {
  const fn = LAYOUTS[slide.layout];
  if (!fn) throw new Error(`未知版式 "${slide.layout}"。可选：${LAYOUT_IDS.join(', ')}`);
  const dark = DARK_LAYOUTS.has(slide.layout);
  const shapes = fn(slide, ctx);
  // 浅色内容页补底色；深色页底色由版式自己画
  if (!dark) shapes.unshift(box({ x: 0, y: 0, w: W, h: H, fill: ctx.style.pageBg, role: 'bg' }));
  for (const extra of chrome(ctx.style, ctx.pageNo, ctx.total, dark)) shapes.push(extra);
  return shapes;
}
