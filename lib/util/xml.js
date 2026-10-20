/**
 * @module util/xml
 * OOXML 文本与形状原语：转义、单位换算、填充/描边、段落与 run 拼装。
 *
 * 设计约束：所有几何量以 **pt** 为单位传给本模块，内部换算 EMU。
 * 1pt = 12700 EMU；画布 960x540pt = 12192000x6858000 EMU（16:9，13.333x7.5 英寸）。
 */

export const EMU_PER_PT = 12700;
export const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
export const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
export const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const NS_S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';

/** XML 头（standalone） —— 所有 part 一律用它开头。 */
export const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

/** XML/A 命名空间声明串，供 pptx 各 part 复用。 */
export const NS_DECL_PPT = `xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"`;

/** 转义 XML 文本与属性值。 */
export function esc(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/** pt -> EMU（整数）。 */
export function emu(pt) {
  return Math.round(Number(pt) * EMU_PER_PT);
}

/** 0..1 比例的 alpha -> OOXML 的 0..100000 整数。 */
export function alphaOf(value) {
  return Math.max(0, Math.min(100000, Math.round(Number(value) * 100000)));
}

/** 归一化颜色：接受 'RRGGBB' / '#RRGGBB' / 'F00'。 */
export function normHex(color) {
  if (color == null) return null;
  let c = String(color).trim().replace(/^#/, '');
  if (c.length === 3) c = c.split('').map((ch) => ch + ch).join('');
  if (!/^[0-9a-fA-F]{6}$/.test(c)) {
    const short = c.toUpperCase();
    const named = { WHITE: 'FFFFFF', BLACK: '000000' }[short];
    if (named) return named;
    throw new Error(`非法颜色值: ${JSON.stringify(color)}（需为 RRGGBB 或 #RRGGBB）`);
  }
  return c.toUpperCase();
}

/** <a:solidFill> 或 <a:noFill>。 */
export function fillXml(color, alpha) {
  if (color == null) return '<a:noFill/>';
  const hex = normHex(color);
  const a = alpha == null ? '' : `<a:alpha val="${alphaOf(alpha)}"/>`;
  return `<a:solidFill><a:srgbClr val="${hex}">${a}</a:srgbClr></a:solidFill>`;
}

/** <a:ln> 或 <a:noFill> 包围的描边。 */
export function lineXml(color, widthPt = 0.75, alpha) {
  if (color == null) return '<a:ln><a:noFill/></a:ln>';
  return `<a:ln w="${emu(widthPt)}" cap="flat" cmpd="sng" algn="ctr">`
    + fillXml(color, alpha)
    + '<a:prstDash val="solid"/></a:ln>';
}

/**
 * 单个 run 的 <a:rPr>。
 * @param {{size?:number,bold?:boolean,italic?:boolean,color?:string,spc?:number,fontLat?:string,fontEa?:string}} p
 */
export function runPropsXml(p = {}) {
  const attrs = [`lang="zh-CN"`, `altLang="en-US"`];
  if (p.size != null) attrs.push(`sz="${Math.round(Number(p.size) * 100)}"`);
  if (p.bold) attrs.push('b="1"');
  if (p.italic) attrs.push('i="1"');
  if (p.spc) attrs.push(`spc="${Math.round(Number(p.spc) * 100)}"`);
  if (p.underline) attrs.push('u="sng"');
  attrs.push('dirty="0"');
  return `<a:rPr ${attrs.join(' ')}>`
    + fillXml(p.color ?? '000000')
    + `<a:latin typeface="${esc(p.fontLat ?? 'Segoe UI')}"/>`
    + `<a:ea typeface="${esc(p.fontEa ?? 'Microsoft YaHei')}"/>`
    + '</a:rPr>';
}

/**
 * 一个物理行 <a:p>。IR 里的段落已经由布局引擎折行，因此这里一个 para = 一行。
 * @param {{runs:Array<{t:string,size?:number,bold?:boolean,color?:string,italic?:boolean,spc?:number}>,
 *          align?:'l'|'ctr'|'r'|'just', lineHeight?:number, spaceBefore?:number}} para
 */
export function paraXml(para) {
  const align = para.align ?? 'l';
  const maxRunSize = (para.runs ?? []).reduce((m, r) => Math.max(m, r.size ?? 0), 12);
  // 用**绝对行距**（spcPts）而不是百分比：
  // spcPct 是相对字体的自然行高（YaHei ≈ 1.32em），会让 PowerPoint 的换行高度
  // 与布局期计算不一致；spcPts 直接把行高钉死成布局引擎算出的 pt 值，所见即所得。
  const lineHeightPt = (para.lineHeight ?? 1.2) * (maxRunSize || 12);
  let ppr = `<a:pPr algn="${align}">`;
  ppr += `<a:lnSpc><a:spcPts val="${Math.round(lineHeightPt * 100)}"/></a:lnSpc>`;
  ppr += `<a:spcBef><a:spcPts val="${Math.round((para.spaceBefore ?? 0) * 100)}"/></a:spcBef>`;
  ppr += '<a:spcAft><a:spcPts val="0"/></a:spcAft>';
  ppr += '</a:pPr>';
  const runs = (para.runs ?? [])
    .filter((r) => r && r.t != null && String(r.t) !== '')
    .map((r) => `<a:r>${runPropsXml(r)}<a:t>${esc(r.t)}</a:t></a:r>`)
    .join('');
  return `<a:p>${ppr}${runs}</a:p>`;
}

/** 形状的 <p:spPr>（xfrm + prstGeom + fill + ln）。 */
export function shapePropsXml(shape) {
  const { x, y, w, h } = shape;
  const rot = shape.rot ? ` rot="${Math.round(shape.rot * 60000)}"` : '';
  const xfrm = `<a:xfrm${rot}><a:off x="${emu(x)}" y="${emu(y)}"/>`
    + `<a:ext cx="${emu(Math.max(0.01, w))}" cy="${emu(Math.max(0.01, h))}"/></a:xfrm>`;
  const prst = shape.k === 'ellipse' ? 'ellipse'
    : shape.k === 'round' ? 'roundRect'
      : shape.k === 'line' ? 'line' : 'rect';
  const adj = typeof shape.adj === 'number'
    ? `<a:avLst><a:gd name="adj" fmla="val ${Math.round(shape.adj * 100000)}"/></a:avLst>`
    : '<a:avLst/>';
  const geom = `<a:prstGeom prst="${prst}">${adj}</a:prstGeom>`;
  return `<p:spPr>${xfrm}${geom}${fillXml(shape.fill, shape.fillAlpha)}`
    + `${lineXml(shape.line, shape.lineWidth ?? 0.75, shape.lineAlpha)}</p:spPr>`;
}

/** 空 spTree（主控/版式用）。 */
export const EMPTY_SP_TREE = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
  + '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
  + '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
