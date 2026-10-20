/**
 * @module xlsx/xlsx
 * SpreadsheetML 生成与校验（纯手写 OOXML）。
 * 字符串用 inlineStr，省掉 sharedStrings 这一层；公式直接写 <f>。
 */
import { zipSync } from '../util/zip.js';
import { XML_HEAD, esc } from '../util/xml.js';

export class XlsxSpecError extends Error {
  constructor(message) {
    super(message);
    this.name = 'XlsxSpecError';
  }
}

const fail = (where, message) => {
  throw new XlsxSpecError(`${where}: ${message}`);
};

const colLetter = (index) => {
  let n = index + 1;
  let out = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
};

/**
 * 校验 XlsxSpec。
 * @param {any} input
 * @returns {{title?:string, author?:string, sheets:Array<object>}}
 */
export function parseSheetSpec(input) {
  if (typeof input === 'string') {
    try {
      return parseSheetSpec(JSON.parse(input));
    } catch (err) {
      throw new XlsxSpecError(`sheets 是字符串但 JSON 解析失败：${err.message}`);
    }
  }
  if (!input || typeof input !== 'object') {
    throw new XlsxSpecError('spec 必须是对象，形如 {"sheets":[{"name":"Sheet1","rows":[["A","B"],[1,2]]}]}');
  }
  const sheets = input.sheets;
  if (!Array.isArray(sheets) || sheets.length === 0) fail('sheets', '必须是非空数组');
  if (sheets.length > 20) fail('sheets', `最多 20 个工作表，收到 ${sheets.length} 个`);
  const seen = new Set();
  const parsed = sheets.map((sheet, i) => {
    const where = `sheets[${i}]`;
    if (!sheet || typeof sheet !== 'object') fail(where, '必须是对象');
    const rawName = String(sheet.name ?? `Sheet${i + 1}`);
    if (rawName.length > 31) fail(`${where}.name`, `工作表名最长 31 字符，收到 ${rawName.length}`);
    if (/[[\]:*?/\\]/.test(rawName)) fail(`${where}.name`, '工作表名不能包含 [ ] : * ? / \\');
    if (seen.has(rawName)) fail(`${where}.name`, `工作表名重复："${rawName}"`);
    seen.add(rawName);
    const rows = sheet.rows;
    if (!Array.isArray(rows) || rows.length === 0) fail(`${where}.rows`, '必须是非空数组');
    if (rows.length > 5000) fail(`${where}.rows`, `最多 5000 行，收到 ${rows.length} 行`);
    const widths = sheet.widths;
    if (widths != null) {
      if (!Array.isArray(widths) || widths.some((w) => typeof w !== 'number' || w <= 0)) {
        fail(`${where}.widths`, '必须是正数数组（列宽，单位近似字符数）');
      }
    }
    return {
      name: rawName,
      header: sheet.header !== false,
      widths,
      rows: rows.map((row, r) => {
        if (!Array.isArray(row)) fail(`${where}.rows[${r}]`, '必须是数组');
        if (row.length > 64) fail(`${where}.rows[${r}]`, `最多 64 列，收到 ${row.length} 列`);
        return row.map((cell) => cell ?? '');
      }),
    };
  });
  return {
    title: input.title != null ? String(input.title) : undefined,
    author: input.author != null ? String(input.author) : undefined,
    accent: input.accent,
    sheets: parsed,
  };
}

function cellXml(ref, value, styleId) {
  const s = styleId ? ` s="${styleId}"` : '';
  if (value === null || value === undefined || value === '') return `<c r="${ref}"${s}/>`;
  if (typeof value === 'number' && Number.isFinite(value)) return `<c r="${ref}"${s}><v>${value}</v></c>`;
  if (typeof value === 'boolean') return `<c r="${ref}"${s} t="b"><v>${value ? 1 : 0}</v></c>`;
  const text = String(value);
  if (text.startsWith('=') && text.length > 1) {
    return `<c r="${ref}"${s}><f>${esc(text.slice(1))}</f></c>`;
  }
  return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${esc(text)}</t></is></c>`;
}

function sheetXml(sheet) {
  const cols = Math.max(1, ...sheet.rows.map((row) => row.length));
  const colsXml = sheet.widths
    ? `<cols>${sheet.widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('')}</cols>`
    : '';
  const rowsXml = sheet.rows.map((row, r) => {
    const rowNumber = r + 1;
    const isHeader = sheet.header && r === 0;
    const cells = [];
    for (let c = 0; c < cols; c += 1) {
      const styleId = isHeader ? 1 : 2;
      cells.push(cellXml(`${colLetter(c)}${rowNumber}`, row[c], styleId));
    }
    const ht = isHeader ? ' ht="22" customHeight="1"' : '';
    return `<row r="${rowNumber}"${ht}>${cells.join('')}</row>`;
  }).join('');
  const freeze = sheet.header
    ? '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
    : '';
  const dimension = `A1:${colLetter(cols - 1)}${sheet.rows.length}`;
  return `${XML_HEAD}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`
    + `<dimension ref="${dimension}"/>${freeze}${colsXml}<sheetData>${rowsXml}</sheetData></worksheet>`;
}

function stylesXml(accent) {
  return `${XML_HEAD}<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">`
    + '<fonts count="3">'
    + '<font><sz val="11"/><color theme="1"/><name val="Segoe UI"/><charset val="134"/></font>'
    + '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Segoe UI"/><charset val="134"/></font>'
    + '<font><sz val="11"/><color rgb="FF1F2937"/><name val="Segoe UI"/><charset val="134"/></font>'
    + '</fonts>'
    + '<fills count="3">'
    + '<fill><patternFill patternType="none"/></fill>'
    + '<fill><patternFill patternType="gray125"/></fill>'
    + `<fill><patternFill patternType="solid"><fgColor rgb="FF${accent}"/><bgColor indexed="64"/></patternFill></fill>`
    + '</fills>'
    + '<borders count="2">'
    + '<border><left/><right/><top/><bottom/><diagonal/></border>'
    + '<border><left style="thin"><color rgb="FFD8DEE8"/></left><right style="thin"><color rgb="FFD8DEE8"/></right>'
    + '<top style="thin"><color rgb="FFD8DEE8"/></top><bottom style="thin"><color rgb="FFD8DEE8"/></bottom><diagonal/></border>'
    + '</borders>'
    + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    + '<cellXfs count="3">'
    + '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    + '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"><alignment vertical="center"/></xf>'
    + '<xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"><alignment vertical="center" wrapText="1"/></xf>'
    + '</cellXfs>'
    + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    + '</styleSheet>';
}

/**
 * 生成 xlsx。
 * @param {{title?:string, author?:string, accent?:string, sheets:Array<object>}} spec 已校验
 * @returns {Buffer}
 */
export function buildXlsx(spec) {
  const accent = (spec.accent ?? '1D4ED8').replace('#', '');
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  const sheetCount = spec.sheets.length;

  const contentTypes = `${XML_HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    + spec.sheets.map((_s, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')
    + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    + '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
    + '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
    + '</Types>';

  const rootRels = `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
    + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
    + '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
    + '</Relationships>';

  const workbook = `${XML_HEAD}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`
    + '<bookViews><workbookView activeTab="0"/></bookViews>'
    + '<sheets>'
    + spec.sheets.map((s, i) => `<sheet name="${esc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')
    + '</sheets></workbook>';

  const workbookRels = `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
    + spec.sheets.map((_s, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')
    + `<Relationship Id="rId${sheetCount + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`
    + '</Relationships>';

  const core = `${XML_HEAD}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">`
    + `<dc:title>${esc(spec.title ?? '工作簿')}</dc:title>`
    + `<dc:creator>${esc(spec.author ?? 'dsh-office-better')}</dc:creator>`
    + `<cp:lastModifiedBy>${esc(spec.author ?? 'dsh-office-better')}</cp:lastModifiedBy>`
    + `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>`
    + `<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>`
    + '</cp:coreProperties>';

  const app = `${XML_HEAD}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">`
    + '<Application>dsh-office-better</Application></Properties>';

  const entries = [
    { name: '[Content_Types].xml', data: contentTypes },
    { name: '_rels/.rels', data: rootRels },
    { name: 'docProps/core.xml', data: core },
    { name: 'docProps/app.xml', data: app },
    { name: 'xl/workbook.xml', data: workbook },
    { name: 'xl/_rels/workbook.xml.rels', data: workbookRels },
    { name: 'xl/styles.xml', data: stylesXml(accent) },
  ];
  spec.sheets.forEach((sheet, i) => {
    entries.push({ name: `xl/worksheets/sheet${i + 1}.xml`, data: sheetXml(sheet) });
  });
  return zipSync(entries);
}
