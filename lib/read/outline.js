/**
 * @module read/outline
 * 反向读取：把 docx / xlsx / pptx 抽成结构化大纲。
 * 只做"读得出来"的部分（文本、样式名、单元格），用于审阅和二次编辑的输入。
 */
import { unzipSync } from '../util/zip.js';

function decodeXml(value) {
  return String(value)
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

function allText(xml, tag) {
  const out = [];
  const re = new RegExp(`<${tag}[^>]*>([\\s\\S]*?)</${tag}>`, 'g');
  let m = re.exec(xml);
  while (m) {
    out.push(decodeXml(m[1].replace(/<[^>]+>/g, '')));
    m = re.exec(xml);
  }
  return out;
}

/** PPTX：逐页取文本。 */
function readPptx(zip) {
  const slideNames = zip.names()
    .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort((a, b) => Number(a.match(/(\d+)/)[1]) - Number(b.match(/(\d+)/)[1]));
  const slides = slideNames.map((name, i) => {
    const texts = allText(zip.text(name), 'a:t').map((t) => t.trim()).filter(Boolean);
    return { index: i + 1, lines: texts };
  });
  return { kind: 'pptx', count: slides.length, unit: '页', sections: slides };
}

/** DOCX：逐段取文本与样式。 */
function readDocx(zip) {
  const xml = zip.text('word/document.xml');
  const paras = [];
  const re = /<w:p\b([\s\S]*?)<\/w:p>/g;
  let m = re.exec(xml);
  const styleNames = new Map();
  if (zip.has('word/styles.xml')) {
    const styles = zip.text('word/styles.xml');
    const sre = /<w:style\b[^>]*w:styleId="([^"]+)"[^>]*>([\s\S]*?)<\/w:style>/g;
    let sm = sre.exec(styles);
    while (sm) {
      const styleId = sm[1];
      if (styleId.startsWith('Heading')) styleNames.set(styleId, `H${styleId.replace(/\D/g, '')}`);
      else if (styleId === 'DocTitle') styleNames.set(styleId, 'H0');
      sm = sre.exec(styles);
    }
  }
  while (m) {
    const inner = m[1];
    const styleMatch = inner.match(/<w:pStyle w:val="([^"]+)"\/>/);
    const texts = allText(inner.replace(/<w:pPr>[\s\S]*?<\/w:pPr>/g, ''), 'w:t');
    const text = texts.join('').trim();
    if (text) {
      paras.push({
        style: styleMatch ? (styleNames.get(styleMatch[1]) ?? styleMatch[1]) : 'P',
        text,
      });
    }
    m = re.exec(xml);
  }
  return { kind: 'docx', count: paras.length, unit: '段', sections: [{ index: 1, lines: paras.map((p) => `[${p.style}] ${p.text}`) }] };
}

/** XLSX：逐表逐行取单元格。 */
function readXlsx(zip) {
  const shared = [];
  if (zip.has('xl/sharedStrings.xml')) {
    const sst = zip.text('xl/sharedStrings.xml');
    const re = /<si>([\s\S]*?)<\/si>/g;
    let m = re.exec(sst);
    while (m) {
      shared.push(allText(m[1], 't').join(''));
      m = re.exec(sst);
    }
  }
  const names = [];
  if (zip.has('xl/workbook.xml')) {
    const wb = zip.text('xl/workbook.xml');
    const re = /<sheet\b[^>]*name="([^"]+)"/g;
    let m = re.exec(wb);
    while (m) {
      names.push(decodeXml(m[1]));
      m = re.exec(wb);
    }
  }
  const sheetParts = zip.names()
    .filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n))
    .sort((a, b) => Number(a.match(/(\d+)/)[1]) - Number(b.match(/(\d+)/)[1]));
  const sections = sheetParts.map((name, i) => {
    const xml = zip.text(name);
    const lines = [];
    const rowRe = /<row\b[^>]*>([\s\S]*?)<\/row>/g;
    let rm = rowRe.exec(xml);
    while (rm) {
      const cells = [];
      const cellRe = /<c\b([^>]*)>([\s\S]*?)<\/c>/g;
      let cm = cellRe.exec(rm[1]);
      while (cm) {
        const attrs = cm[1];
        const body = cm[2];
        const ref = (attrs.match(/r="([A-Z]+\d+)"/) ?? [])[1] ?? '';
        const type = (attrs.match(/t="([^"]+)"/) ?? [])[1] ?? 'n';
        let value = '';
        if (type === 'inlineStr') value = allText(body, 't').join('');
        else if (type === 's') {
          const idx = Number((body.match(/<v>([\s\S]*?)<\/v>/) ?? [])[1]);
          value = shared[idx] ?? '';
        } else {
          value = decodeXml((body.match(/<v>([\s\S]*?)<\/v>/) ?? [])[1] ?? '');
          const formula = (body.match(/<f>([\s\S]*?)<\/f>/) ?? [])[1];
          if (formula) value = `=${decodeXml(formula)}${value ? ` → ${value}` : ''}`;
        }
        if (value !== '') cells.push(`${ref}=${value}`);
        cm = cellRe.exec(rm[1]);
      }
      if (cells.length > 0) lines.push(cells.join(' | '));
      rm = rowRe.exec(xml);
    }
    return { index: i + 1, name: names[i] ?? `Sheet${i + 1}`, lines };
  });
  return { kind: 'xlsx', count: sections.length, unit: '个工作表', sections };
}

/**
 * 读任意 OOXML 文件。
 * @param {Buffer} buf
 * @param {'pptx'|'docx'|'xlsx'} kind
 */
export function readOffice(buf, kind) {
  const zip = unzipSync(buf);
  if (kind === 'pptx') return readPptx(zip);
  if (kind === 'docx') return readDocx(zip);
  if (kind === 'xlsx') return readXlsx(zip);
  throw new Error(`不支持的类型 ${kind}（只支持 pptx / docx / xlsx）`);
}

/** 大纲 -> 纯文本（给工具输出用）。 */
export function outlineToText(outline, { maxLines = 400 } = {}) {
  const parts = [];
  let budget = maxLines;
  for (const section of outline.sections) {
    const head = outline.kind === 'xlsx' ? `## ${section.name}` : `## ${outline.unit} ${section.index}`;
    parts.push(head);
    for (const line of section.lines) {
      if (budget <= 0) {
        parts.push('…（已截断）');
        return parts.join('\n');
      }
      parts.push(line);
      budget -= 1;
    }
  }
  return parts.join('\n');
}
