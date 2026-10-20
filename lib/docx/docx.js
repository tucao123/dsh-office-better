/**
 * @module docx/docx
 * WordprocessingML 生成与校验（纯手写 OOXML）。
 *
 * 输入是一份块列表（heading / paragraph / bullets / numbered / table / quote /
 * code / divider / pagebreak），输出 .docx 字节流。
 * 与原实现的差别：那里 docx 走的是 officecli 的"路径寻址 + 属性读写"原语，
 * 依赖外部进程且需要模型自己拼路径；这里是一次性声明式生成 + 结构化校验。
 */
import { zipSync } from '../util/zip.js';
import { XML_HEAD, esc } from '../util/xml.js';

const TWIP = 20; // 1pt = 20 twips
const HALF = 2; // 字号半磅

export class DocxSpecError extends Error {
  constructor(message) {
    super(message);
    this.name = 'DocxSpecError';
  }
}

const fail = (where, message) => {
  throw new DocxSpecError(`${where}: ${message}`);
};

const BLOCK_TYPES = ['heading', 'paragraph', 'bullets', 'numbered', 'table', 'quote', 'code', 'divider', 'pagebreak'];

const text = (value, where, max = 400) => {
  if (typeof value !== 'string') fail(where, `必须是字符串（收到 ${typeof value}）`);
  if (value.length > max) fail(where, `超长 ${value.length} 字，上限 ${max} 字`);
  return value;
};

/**
 * 校验 DocSpec。
 * @param {any} input
 * @returns {{title?:string, subtitle?:string, author?:string, accent?:string, blocks:Array<object>}}
 */
export function parseDocSpec(input) {
  if (typeof input === 'string') {
    try {
      return parseDocSpec(JSON.parse(input));
    } catch (err) {
      throw new DocxSpecError(`doc 是字符串但 JSON 解析失败：${err.message}`);
    }
  }
  if (!input || typeof input !== 'object') {
    throw new DocxSpecError('doc 必须是对象，形如 {"title":"…","blocks":[{"type":"heading","text":"…"}]}');
  }
  const blocks = input.blocks;
  if (!Array.isArray(blocks) || blocks.length === 0) fail('blocks', '必须是非空数组');
  if (blocks.length > 200) fail('blocks', `最多 200 个块，收到 ${blocks.length} 个`);
  const parsed = blocks.map((block, i) => {
    const where = `blocks[${i}]`;
    if (!block || typeof block !== 'object') fail(where, '必须是对象');
    const type = String(block.type ?? '');
    if (!BLOCK_TYPES.includes(type)) fail(`${where}.type`, `未知类型 "${type}"，可选：${BLOCK_TYPES.join(', ')}`);
    switch (type) {
      case 'heading': {
        const level = Number(block.level ?? 1);
        if (![1, 2, 3].includes(level)) fail(`${where}.level`, '只能是 1、2、3');
        return { type, level, text: text(block.text, `${where}.text`, 120) };
      }
      case 'paragraph':
        return { type, text: text(block.text, `${where}.text`, 2000), bold: block.bold === true, align: block.align };
      case 'quote':
      case 'code':
        return { type, text: text(block.text, `${where}.text`, 3000) };
      case 'bullets':
      case 'numbered': {
        const items = block.items;
        if (!Array.isArray(items) || items.length === 0) fail(`${where}.items`, '必须是非空数组');
        if (items.length > 30) fail(`${where}.items`, `最多 30 项，收到 ${items.length} 项`);
        return { type, items: items.map((item, j) => text(item, `${where}.items[${j}]`, 400)) };
      }
      case 'table': {
        const headers = block.headers;
        const rows = block.rows;
        if (!Array.isArray(headers) || headers.length === 0) fail(`${where}.headers`, '必须是非空数组');
        if (headers.length > 8) fail(`${where}.headers`, `最多 8 列，收到 ${headers.length} 列`);
        if (!Array.isArray(rows) || rows.length === 0) fail(`${where}.rows`, '必须是非空数组');
        if (rows.length > 100) fail(`${where}.rows`, `最多 100 行，收到 ${rows.length} 行`);
        return {
          type,
          headers: headers.map((h, j) => text(h, `${where}.headers[${j}]`, 40)),
          rows: rows.map((row, r) => {
            if (!Array.isArray(row)) fail(`${where}.rows[${r}]`, '必须是数组');
            if (row.length > headers.length) fail(`${where}.rows[${r}]`, `有 ${row.length} 列，超过表头 ${headers.length} 列`);
            return row.map((cell, c) => text(String(cell ?? ''), `${where}.rows[${r}][${c}]`, 200));
          }),
        };
      }
      case 'divider':
      case 'pagebreak':
        return { type };
      default:
        fail(`${where}.type`, '未实现');
    }
  });
  return {
    title: input.title != null ? text(input.title, 'title', 80) : undefined,
    subtitle: input.subtitle != null ? text(input.subtitle, 'subtitle', 120) : undefined,
    author: input.author != null ? text(input.author, 'author', 40) : undefined,
    accent: input.accent,
    blocks: parsed,
  };
}

/* --------------------------------------------------------------- XML 辅助 */

const wRun = (content, { bold = false, italic = false, size, color, mono = false } = {}) => {
  const rpr = '<w:rPr>'
    + `<w:rFonts w:ascii="${mono ? 'Consolas' : 'Segoe UI'}" w:hAnsi="${mono ? 'Consolas' : 'Segoe UI'}" w:eastAsia="Microsoft YaHei"/>`
    + (bold ? '<w:b/>' : '')
    + (italic ? '<w:i/>' : '')
    + (color ? `<w:color w:val="${color}"/>` : '')
    + (size ? `<w:sz w:val="${Math.round(size * HALF)}"/><w:szCs w:val="${Math.round(size * HALF)}"/>` : '')
    + '</w:rPr>';
  return `<w:r>${rpr}<w:t xml:space="preserve">${esc(content)}</w:t></w:r>`;
};

const wPara = (runs, { style, spacing, indent, align, keepNext } = {}) => {
  const ppr = '<w:pPr>'
    + (style ? `<w:pStyle w:val="${style}"/>` : '')
    + (spacing ? `<w:spacing ${Object.entries(spacing).map(([k, v]) => `w:${k}="${v}"`).join(' ')}/>` : '')
    + (indent ? `<w:ind ${Object.entries(indent).map(([k, v]) => `w:${k}="${v}"`).join(' ')}/>` : '')
    + (align ? `<w:jc w:val="${align}"/>` : '')
    + (keepNext ? '<w:keepNext/>' : '')
    + '</w:pPr>';
  return `<w:p>${ppr}${runs}</w:p>`;
};

const wTable = (headers, rows, accent) => {
  const cols = headers.length;
  const gridW = Math.floor(9638 / cols);
  const grid = '<w:tblGrid>' + headers.map(() => `<w:gridCol w:w="${gridW}"/>`).join('') + '</w:tblGrid>';
  const border = (side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="D8DEE8"/>`;
  const tblPr = '<w:tblPr>'
    + '<w:tblW w:w="5000" w:type="pct"/>'
    + `<w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(border).join('')}</w:tblBorders>`
    + '<w:tblCellMar><w:top w:w="80" w:type="dxa"/><w:left w:w="120" w:type="dxa"/>'
    + '<w:bottom w:w="80" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tblCellMar>'
    + '</w:tblPr>';
  const cell = (value, { head = false } = {}) => {
    const shade = head ? `<w:shd w:val="clear" w:color="auto" w:fill="${accent}"/>` : '';
    const runs = wRun(String(value), { bold: head, size: 10.5, color: head ? 'FFFFFF' : '1F2937' });
    return `<w:tc><w:tcPr><w:tcW w:w="${gridW}" w:type="dxa"/>${shade}</w:tcPr>`
      + `${wPara(runs, { spacing: { before: '40', after: '40', line: '260', lineRule: 'atLeast' } })}</w:tc>`;
  };
  const headRow = `<w:tr><w:trPr><w:tblHeader/></w:trPr>${headers.map((h) => cell(h, { head: true })).join('')}</w:tr>`;
  const bodyRows = rows.map((row) => {
    const cells = [];
    for (let i = 0; i < cols; i += 1) cells.push(cell(row[i] ?? ''));
    return `<w:tr>${cells.join('')}</w:tr>`;
  }).join('');
  return `<w:tbl>${tblPr}${grid}${headRow}${bodyRows}</w:tbl>`
    + wPara('', { spacing: { before: '0', after: '120' } });
};

/* --------------------------------------------------------------- 打包 */

const stylesXml = (accent, ink) => `${XML_HEAD}<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">`
  + '<w:docDefaults><w:rPrDefault><w:rPr>'
  + '<w:rFonts w:ascii="Segoe UI" w:hAnsi="Segoe UI" w:eastAsia="Microsoft YaHei" w:cs="Segoe UI"/>'
  + '<w:sz w:val="21"/><w:szCs w:val="21"/><w:color w:val="1F2937"/>'
  + '</w:rPr></w:rPrDefault>'
  + '<w:pPrDefault><w:pPr><w:spacing w:after="140" w:line="300" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
  + '</w:docDefaults>'
  + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
  + ['1', '2', '3'].map((lv) => {
    const size = { 1: 21, 2: 16, 3: 13 }[lv];
    return `<w:style w:type="paragraph" w:styleId="Heading${lv}"><w:name w:val="heading ${lv}"/><w:basedOn w:val="Normal"/>`
      + '<w:next w:val="Normal"/><w:qFormat/>'
      + `<w:pPr><w:keepNext/><w:spacing w:before="${lv === '1' ? 300 : 220}" w:after="100"/><w:outlineLvl w:val="${Number(lv) - 1}"/></w:pPr>`
      + `<w:rPr><w:b/><w:color w:val="${ink}"/><w:sz w:val="${size * 2}"/><w:szCs w:val="${size * 2}"/></w:rPr></w:style>`;
  }).join('')
  + `<w:style w:type="paragraph" w:styleId="DocTitle"><w:name w:val="Doc Title"/><w:basedOn w:val="Normal"/><w:qFormat/>`
  + `<w:pPr><w:spacing w:before="0" w:after="160"/></w:pPr>`
  + `<w:rPr><w:b/><w:color w:val="${ink}"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr></w:style>`
  + `<w:style w:type="paragraph" w:styleId="Quote2"><w:name w:val="Quote 2"/><w:basedOn w:val="Normal"/>`
  + `<w:pPr><w:ind w:left="360"/><w:pBdr><w:left w:val="single" w:sz="18" w:space="8" w:color="${accent}"/></w:pBdr></w:pPr>`
  + `<w:rPr><w:i/><w:color w:val="4B5563"/></w:rPr></w:style>`
  + '<w:style w:type="paragraph" w:styleId="CodeBlock"><w:name w:val="Code Block"/><w:basedOn w:val="Normal"/>'
  + '<w:pPr><w:shd w:val="clear" w:color="auto" w:fill="F3F5F9"/><w:spacing w:before="60" w:after="60" w:line="260" w:lineRule="auto"/></w:pPr>'
  + '<w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="19"/></w:rPr></w:style>'
  + '</w:styles>';

const numberingXml = `${XML_HEAD}<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">`
  + '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
  + '<w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="425" w:hanging="285"/></w:pPr></w:lvl></w:abstractNum>'
  + '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
  + '<w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="425" w:hanging="285"/></w:pPr></w:lvl></w:abstractNum>'
  + '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
  + '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>'
  + '</w:numbering>';

/**
 * 生成 docx。
 * @param {{title?:string, subtitle?:string, author?:string, accent?:string, blocks:Array<object>}} doc 已校验
 * @returns {Buffer}
 */
export function buildDocx(doc) {
  const accent = (doc.accent ?? '1D4ED8').replace('#', '');
  const ink = '111827';
  const body = [];
  if (doc.subtitle) body.push(wPara(wRun(doc.subtitle, { size: 11, color: '6B7280' }), { spacing: { before: '0', after: '40' } }));
  if (doc.title) body.push(wPara(wRun(doc.title), { style: 'DocTitle' }));

  for (const block of doc.blocks) {
    switch (block.type) {
      case 'heading':
        body.push(wPara(wRun(block.text), { style: `Heading${block.level}`, keepNext: true }));
        break;
      case 'paragraph':
        body.push(wPara(wRun(block.text, { bold: block.bold }), {
          align: block.align === 'center' ? 'center' : block.align === 'right' ? 'right' : undefined,
        }));
        break;
      case 'quote':
        body.push(wPara(wRun(block.text), { style: 'Quote2' }));
        break;
      case 'code':
        body.push(wPara(wRun(block.text, { mono: true }), { style: 'CodeBlock' }));
        break;
      case 'bullets':
      case 'numbered':
        for (const item of block.items) {
          body.push(wPara(wRun(item), {
            spacing: { before: '0', after: '60' },
            indent: { left: '425', hanging: '285' },
            ...(block.type === 'numbered' ? {} : {}),
          }).replace('<w:pPr>', `<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${block.type === 'bullets' ? 1 : 2}"/></w:numPr>`));
        }
        break;
      case 'table':
        body.push(wTable(block.headers, block.rows, accent));
        break;
      case 'divider':
        body.push(wPara('', {
          spacing: { before: '120', after: '120' },
        }).replace('<w:pPr>', '<w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="D8DEE8"/></w:pBdr>'));
        break;
      case 'pagebreak':
        body.push(`<w:p><w:r><w:br w:type="page"/></w:r></w:p>`);
        break;
      default:
        break;
    }
  }

  const sectPr = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
    + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="851" w:footer="992" w:gutter="0"/>'
    + '<w:docGrid w:linePitch="312"/></w:sectPr>';

  const documentXml = `${XML_HEAD}<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`
    + `<w:body>${body.join('')}${sectPr}</w:body></w:document>`;

  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  const contentTypes = `${XML_HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    + '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    + '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
    + '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
    + '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
    + '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
    + '</Types>';

  const rootRels = `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
    + '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
    + '</Relationships>';

  const documentRels = `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
    + '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
    + '</Relationships>';

  const settingsXml = `${XML_HEAD}<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">`
    + '<w:zoom w:percent="100"/><w:defaultTabStop w:val="420"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
    + '</w:settings>';

  const core = `${XML_HEAD}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">`
    + `<dc:title>${esc(doc.title ?? '文档')}</dc:title>`
    + `<dc:creator>${esc(doc.author ?? 'dsh-office-better')}</dc:creator>`
    + `<cp:lastModifiedBy>${esc(doc.author ?? 'dsh-office-better')}</cp:lastModifiedBy>`
    + `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>`
    + `<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>`
    + '</cp:coreProperties>';

  const app = `${XML_HEAD}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">`
    + '<Application>dsh-office-better</Application></Properties>';

  return zipSync([
    { name: '[Content_Types].xml', data: contentTypes },
    { name: '_rels/.rels', data: rootRels },
    { name: 'docProps/core.xml', data: core },
    { name: 'docProps/app.xml', data: app },
    { name: 'word/document.xml', data: documentXml },
    { name: 'word/_rels/document.xml.rels', data: documentRels },
    { name: 'word/styles.xml', data: stylesXml(accent, ink) },
    { name: 'word/numbering.xml', data: numberingXml },
    { name: 'word/settings.xml', data: settingsXml },
  ]);
}
