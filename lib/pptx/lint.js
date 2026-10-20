/**
 * @module pptx/lint
 * 数值化版式检查。
 *
 * 原实现的自检闭环是「渲染 PNG → 把图回传给模型 → 模型用眼睛找溢出」。
 * 那需要外部渲染器、需要多模态模型、而且是主观判断。这里改成对 IR 做静态分析：
 * 溢出、越界、对比度、字号、密度、层级 —— 全部可判定、可回归、可在 CI 里跑。
 */
import { CANVAS } from './theme.js';
import { measure } from '../ir/text.js';

const W = CANVAS.w;
const H = CANVAS.h;

/** sRGB 相对亮度。 */
function luminance(hex) {
  const num = parseInt(hex, 16);
  const channels = [(num >> 16) & 0xff, (num >> 8) & 0xff, num & 0xff].map((v) => {
    const c = v / 255;
    return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
  });
  return 0.2126 * channels[0] + 0.7152 * channels[1] + 0.0722 * channels[2];
}

/** WCAG 对比度（1:1 ~ 21:1）。 */
export function contrastRatio(hexA, hexB) {
  const a = luminance(String(hexA).replace('#', ''));
  const b = luminance(String(hexB).replace('#', ''));
  const [hi, lo] = a > b ? [a, b] : [b, a];
  return (hi + 0.05) / (lo + 0.05);
}

function contains(shape, px, py) {
  return px >= shape.x && px <= shape.x + shape.w && py >= shape.y && py <= shape.y + shape.h;
}

/** 找到文本下方最上层的实心填充（决定对比度比较基准）。 */
function backgroundUnder(shapes, index, px, py, fallback) {
  let bg = fallback;
  for (let i = 0; i < index; i += 1) {
    const shape = shapes[i];
    if (!shape.fill) continue;
    if (shape.fillAlpha != null && shape.fillAlpha < 0.98) continue;
    if (shape.k === 'text') continue;
    if (contains(shape, px, py)) bg = shape.fill;
  }
  return bg;
}

/**
 * 检查一组已渲染的页面 IR。
 * @param {Array<{shapes:object[], pageBg:string}>} pages
 * @returns {{errors:Array<object>, warnings:Array<object>, stats:object}}
 */
export function lintPages(pages) {
  const errors = [];
  const warnings = [];
  const stats = { shapes: 0, textBlocks: 0, smallestFont: Infinity, minContrast: Infinity };

  pages.forEach((page, index) => {
    const slideNo = index + 1;
    const shapes = page.shapes ?? [];
    const bgBase = page.pageBg ?? 'FFFFFF';
    stats.shapes += shapes.length;
    let maxTitle = 0;
    let maxBody = 0;
    let visualBlocks = 0;

    shapes.forEach((shape, shapeIndex) => {
      if (shape.role === 'bg' || shape.role === 'deco') return;
      if (shape.k !== 'text') visualBlocks += 1;

      // 1) 越界
      if (shape.x < -0.5 || shape.y < -0.5 || shape.x + shape.w > W + 0.5 || shape.y + shape.h > H + 0.5) {
        errors.push({
          code: 'OUT_OF_CANVAS',
          slide: slideNo,
          shape: shape.label,
          message: `「${shape.label}」超出画布：x=${shape.x.toFixed(1)} y=${shape.y.toFixed(1)} w=${shape.w.toFixed(1)} h=${shape.h.toFixed(1)}（画布 ${W}x${H}）`,
        });
      }
      if (shape.k !== 'text' || !shape.paras) return;

      stats.textBlocks += 1;
      const size = shape.size ?? 12;
      stats.smallestFont = Math.min(stats.smallestFont, size);
      if (shape.role === 'title') maxTitle = Math.max(maxTitle, size);
      if (shape.role === 'body') maxBody = Math.max(maxBody, size);

      // 2) 纵向溢出（布局期已折行，高度可直接比较）
      const contentH = shape.contentH ?? 0;
      if (contentH > shape.h + 0.8) {
        errors.push({
          code: 'TEXT_OVERFLOW',
          slide: slideNo,
          shape: shape.label,
          message: `「${shape.label}」文本高 ${contentH.toFixed(1)}pt 超出文本框 ${shape.h.toFixed(1)}pt`,
        });
      }

      // 3) 横向超宽 / 被截断
      shape.paras.forEach((para, paraIndex) => {
        const text = para.runs.map((r) => r.t).join('');
        const width = para.runs.reduce((sum, r) => sum + measure(r.t, r.size ?? size, r.bold), 0);
        if (width > shape.w + 1) {
          warnings.push({
            code: 'LINE_TOO_WIDE',
            slide: slideNo,
            shape: shape.label,
            message: `「${shape.label}」第 ${paraIndex + 1} 行宽 ${width.toFixed(1)}pt 超出文本框 ${shape.w.toFixed(1)}pt（PowerPoint 可能二次折行）`,
          });
        }
        if (text.trim().endsWith('…')) {
          warnings.push({
            code: 'TEXT_TRUNCATED',
            slide: slideNo,
            shape: shape.label,
            message: `「${shape.label}」内容被截断（"${text.slice(0, 24)}…"），请缩短或拆页`,
          });
        }
      });

      // 4) 字号过小
      if (size < 8.5) {
        warnings.push({
          code: 'FONT_TOO_SMALL',
          slide: slideNo,
          shape: shape.label,
          message: `「${shape.label}」字号 ${size}pt 小于 8.5pt，投影/打印可能看不清`,
        });
      }

      // 5) 对比度
      const firstRun = shape.paras[0]?.runs?.[0];
      const textColor = firstRun?.color ?? shape.color;
      if (textColor) {
        const bg = backgroundUnder(shapes, shapeIndex, shape.x + shape.w / 2, shape.y + shape.h / 2, bgBase);
        const ratio = contrastRatio(textColor, bg);
        stats.minContrast = Math.min(stats.minContrast, ratio);
        const large = (firstRun?.size ?? size) >= 18 || ((firstRun?.size ?? size) >= 14 && firstRun?.bold);
        const threshold = large ? 3 : 4.5;
        if (ratio < threshold) {
          warnings.push({
            code: 'LOW_CONTRAST',
            slide: slideNo,
            shape: shape.label,
            message: `「${shape.label}」对比度 ${ratio.toFixed(2)}:1 低于 ${threshold}:1（文字 #${textColor} / 背景 #${bg}）`,
          });
        }
      }
    });

    // 6) 页面密度
    if (visualBlocks > 16) {
      warnings.push({
        code: 'TOO_DENSE',
        slide: slideNo,
        shape: '页面',
        message: `第 ${slideNo} 页有 ${visualBlocks} 个视觉块，超过 16 个建议拆页`,
      });
    }
    // 7) 层级：标题必须明显大于正文
    if (maxTitle > 0 && maxBody > 0 && maxTitle < maxBody * 1.15) {
      warnings.push({
        code: 'WEAK_HIERARCHY',
        slide: slideNo,
        shape: '页面',
        message: `第 ${slideNo} 页标题 ${maxTitle}pt 与正文 ${maxBody}pt 差距不足 15%，层级不清晰`,
      });
    }
  });

  if (!Number.isFinite(stats.smallestFont)) stats.smallestFont = 0;
  if (!Number.isFinite(stats.minContrast)) stats.minContrast = 21;
  return { errors, warnings, stats };
}
