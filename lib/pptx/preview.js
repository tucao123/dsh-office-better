/**
 * @module pptx/preview
 * IR -> SVG/HTML 预览。
 *
 * 预览不启动任何子进程、不占端口：它就是同一份 IR 的另一种投影。
 * 所以预览和导出的 pptx 天然一致（原实现要靠 officecli watch 起一个 HTTP
 * 服务器再反代 iframe，那套 Host/Origin 改写和 5s 端口上报都是额外故障点）。
 */
import { CANVAS } from './theme.js';
import { measure } from '../ir/text.js';

const W = CANVAS.w;
const H = CANVAS.h;
const FONT = "'Segoe UI','Microsoft YaHei','PingFang SC','Hiragino Sans GB',sans-serif";

function escXml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function n(value) {
  return Math.round(Number(value) * 100) / 100;
}

/** 文本块 -> 若干 <text>。逐行逐 run 输出，位置与 pptx 的折行完全一致。 */
function textSvg(shape) {
  const lines = shape.paras ?? [];
  const size = shape.size ?? 12;
  const lh = lines[0]?.lineHeight ?? 1.35;
  const blockH = shape.contentH ?? lines.reduce((sum, p) => sum + size * (p.lineHeight ?? 1.35) + (p.spaceBefore ?? 0), 0);
  let top = shape.y;
  if (shape.anchor === 'ctr') top = shape.y + (shape.h - blockH) / 2;
  else if (shape.anchor === 'b') top = shape.y + shape.h - blockH;

  const out = [];
  let cursor = top;
  for (const para of lines) {
    const lineSize = para.runs?.[0]?.size ?? size;
    const lineHeightPx = lineSize * (para.lineHeight ?? lh);
    const baseline = cursor + (para.spaceBefore ?? 0) + lineHeightPx / 2 + lineSize * 0.33;
    const runs = para.runs ?? [];
    const totalW = runs.reduce((sum, r) => sum + measure(r.t, r.size ?? lineSize, r.bold), 0);
    let x = shape.x;
    if (para.align === 'ctr') x = shape.x + (shape.w - totalW) / 2;
    else if (para.align === 'r') x = shape.x + shape.w - totalW;
    for (const run of runs) {
      const runSize = run.size ?? lineSize;
      out.push(`<text x="${n(x)}" y="${n(baseline)}" font-size="${n(runSize)}" fill="#${run.color ?? shape.color ?? '111111'}"`
        + ` font-family="${FONT}"${run.bold ? ' font-weight="600"' : ''}`
        + `${run.spc ? ` letter-spacing="${n(run.spc)}"` : ''}>${escXml(run.t)}</text>`);
      x += measure(run.t, runSize, run.bold);
    }
    cursor += (para.spaceBefore ?? 0) + lineHeightPx;
  }
  return out.join('');
}

/** 单个形状 -> SVG 元素。 */
export function shapeToSvg(shape) {
  if (shape.k === 'text') return textSvg(shape);
  const fill = shape.fill ? `fill="#${shape.fill}"` : 'fill="none"';
  const fillOpacity = shape.fill && shape.fillAlpha != null ? ` fill-opacity="${n(shape.fillAlpha)}"` : '';
  const stroke = shape.line
    ? `stroke="#${shape.line}" stroke-width="${n(shape.lineWidth ?? 0.75)}"`
    : 'stroke="none"';
  const strokeOpacity = shape.line && shape.lineAlpha != null ? ` stroke-opacity="${n(shape.lineAlpha)}"` : '';
  if (shape.k === 'ellipse') {
    return `<ellipse cx="${n(shape.x + shape.w / 2)}" cy="${n(shape.y + shape.h / 2)}"`
      + ` rx="${n(shape.w / 2)}" ry="${n(shape.h / 2)}" ${fill}${fillOpacity} ${stroke}${strokeOpacity}/>`;
  }
  const r = shape.k === 'round' ? n(Math.min(shape.w, shape.h) * (shape.adj ?? 0.06)) : 0;
  return `<rect x="${n(shape.x)}" y="${n(shape.y)}" width="${n(shape.w)}" height="${n(shape.h)}"`
    + `${r ? ` rx="${r}" ry="${r}"` : ''} ${fill}${fillOpacity} ${stroke}${strokeOpacity}/>`;
}

/** 一页 -> 完整 SVG。 */
export function slideToSvg(page) {
  const bg = page.background ?? page.pageBg ?? 'FFFFFF';
  const body = page.shapes.map(shapeToSvg).join('');
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" role="img">`
    + `<rect x="0" y="0" width="${W}" height="${H}" fill="#${bg}"/>${body}</svg>`;
}

/**
 * 整份预览 HTML（自包含，可直接用浏览器打开）。
 * @param {{title?:string, pages:Array<object>, meta?:object}} input
 */
export function previewHtml(input) {
  const title = input.title ?? '预览';
  const slides = input.pages.map((page, i) => (
    `<figure class="slide" id="s${i + 1}"><figcaption>${i + 1} / ${input.pages.length}</figcaption>`
    + `${slideToSvg(page)}</figure>`
  )).join('');
  return `<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>${escXml(title)} · 预览</title>
<style>
  :root { color-scheme: dark; }
  * { box-sizing: border-box; }
  body { margin: 0; padding: 28px 20px 60px; background: #10131a; color: #cfd6e4;
    font: 13px/1.6 ${FONT}; }
  header { max-width: 1000px; margin: 0 auto 22px; display: flex; align-items: baseline; gap: 12px; flex-wrap: wrap; }
  header h1 { font-size: 16px; margin: 0; color: #e8edf6; font-weight: 600; }
  header .meta { color: #7c8798; font-size: 12px; }
  .slides { max-width: 1000px; margin: 0 auto; display: flex; flex-direction: column; gap: 22px; }
  figure.slide { margin: 0; background: #fff; border-radius: 10px; overflow: hidden;
    box-shadow: 0 10px 30px rgba(0,0,0,.4); position: relative; }
  figure.slide svg { display: block; width: 100%; height: auto; }
  figure.slide figcaption { position: absolute; right: 8px; bottom: 6px; font-size: 10px;
    color: rgba(120,130,150,.85); background: rgba(255,255,255,.75); padding: 1px 6px; border-radius: 4px; }
  footer { max-width: 1000px; margin: 26px auto 0; color: #6b7688; font-size: 11.5px; }
  @media print {
    body { background: #fff; padding: 0; }
    figure.slide { box-shadow: none; break-after: page; border-radius: 0; }
    header, footer, figcaption { display: none; }
  }
</style></head>
<body>
<header>
  <h1>${escXml(title)}</h1>
  <span class="meta">${input.pages.length} 页 · 由 dsh-office-better 从同一份 IR 渲染（与导出的 pptx 一致）</span>
  ${input.meta?.template ? `<span class="meta">模板 ${escXml(input.meta.template)}</span>` : ''}
  ${input.meta?.theme ? `<span class="meta">主题 ${escXml(input.meta.theme)}</span>` : ''}
</header>
<div class="slides">${slides}</div>
<footer>提示：Ctrl/Cmd + P 可打印为 PDF；预览直接来自版式 IR，不经过任何外部渲染器。</footer>
</body></html>`;
}

/** 文档列表页（浏览器里点开就能看）。 */
export function indexHtml(files, sessionLabel) {
  const rows = files.length === 0
    ? '<p class="empty">当前会话工作区里还没有 Office 文档。让 Agent 生成一个，或把已有文件放进来。</p>'
    : `<table><thead><tr><th>文件</th><th>类型</th><th>大小</th><th>修改时间</th><th></th></tr></thead><tbody>${
      files.map((f) => `<tr><td>${escXml(f.name)}</td><td>${escXml(f.type)}</td>`
        + `<td>${(f.size / 1024).toFixed(1)} KB</td><td>${new Date(f.mtime).toLocaleString('zh-CN')}</td>`
        + `<td><a href="preview?file=${encodeURIComponent(f.name)}">预览</a>`
        + ` · <a href="raw?file=${encodeURIComponent(f.name)}">下载</a></td></tr>`).join('')
    }</tbody></table>`;
  return `<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8"/><title>Office 文档 · ${escXml(sessionLabel)}</title>
<style>
  body { margin: 0; padding: 32px; background: #10131a; color: #cfd6e4;
    font: 13px/1.6 ${FONT}; }
  h1 { font-size: 17px; color: #e8edf6; margin: 0 0 4px; }
  .sub { color: #7c8798; margin-bottom: 22px; }
  table { border-collapse: collapse; width: 100%; max-width: 860px; }
  th, td { text-align: left; padding: 9px 12px; border-bottom: 1px solid #232a36; font-size: 12.5px; }
  th { color: #9aa5b8; font-weight: 500; }
  a { color: #6fd3f2; text-decoration: none; }
  a:hover { text-decoration: underline; }
  .empty { color: #7c8798; }
</style></head>
<body><h1>Office 文档</h1><div class="sub">${escXml(sessionLabel)}</div>${rows}</body></html>`;
}
