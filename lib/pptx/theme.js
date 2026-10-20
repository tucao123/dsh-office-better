/**
 * @module pptx/theme
 * 主题与模板。主题决定配色，模板决定"外衣"（页码格式、页脚、色轨、深色封面）。
 * 两者正交：换模板不动版式，换版式不动模板 —— 与原实现一致的设计取舍，
 * 区别在于这里的 token 是纯数据，不依赖任何外部渲染器。
 */

/** 画布与栅格（pt）。960x540 = 16:9。 */
export const CANVAS = Object.freeze({
  w: 960,
  h: 540,
  margin: 72,
  gap: 16,
  titleY: 46,
  dividerY: 96,
  bodyTop: 124,
  bodyBottom: 466,
  footerY: 500,
});

/**
 * 八个主题。每个主题给出两套墨色：浅色页（page*）与深色页（dark*）。
 * @type {Record<string, any>}
 */
export const THEMES = {
  'business-blue': {
    id: 'business-blue',
    name: '商务蓝',
    pageBg: 'F7F9FC', surface: 'FFFFFF', border: 'E2E8F0',
    pageT1: '0F172A', pageT2: '5B6B80', pageT3: '94A3B8',
    accent: '1D4ED8', accent2: '3B82F6', accentSoft: 'E8EFFC',
    darkBg: '0B1F3F', darkBg2: '112A52', darkT1: 'FFFFFF', darkT2: 'B9C8DC', darkT3: '6B7F9B',
  },
  'academic-crimson': {
    id: 'academic-crimson',
    name: '学术深红',
    pageBg: 'FBF9F7', surface: 'FFFFFF', border: 'E7DED6',
    pageT1: '1C1414', pageT2: '6B5B58', pageT3: 'A29490',
    accent: '8C1D2F', accent2: 'B5304A', accentSoft: 'F7E9EC',
    darkBg: '2A0E14', darkBg2: '3A1520', darkT1: 'FFFFFF', darkT2: 'D8C2C6', darkT3: '8E6A70',
  },
  'tech-cyan': {
    id: 'tech-cyan',
    name: '科技青',
    pageBg: 'F4FAFB', surface: 'FFFFFF', border: 'DBE9EC',
    pageT1: '0B2027', pageT2: '4A6670', pageT3: '8AA3AB',
    accent: '0E7490', accent2: '22D3EE', accentSoft: 'E0F7FC',
    darkBg: '07161C', darkBg2: '0C2229', darkT1: 'FFFFFF', darkT2: 'A9C2CB', darkT3: '5F7C86',
  },
  'warm-orange': {
    id: 'warm-orange',
    name: '暖橙',
    pageBg: 'FEFAF5', surface: 'FFFFFF', border: 'F3E4D3',
    pageT1: '2B1A0C', pageT2: '7A5B44', pageT3: 'B79B85',
    accent: 'C2410C', accent2: 'EA580C', accentSoft: 'FDEBDD',
    darkBg: '2A1608', darkBg2: '3B200C', darkT1: 'FFFFFF', darkT2: 'E5C9B4', darkT3: '9A7355',
  },
  'gov-red': {
    id: 'gov-red',
    name: '政务红',
    pageBg: 'FCFAF9', surface: 'FFFFFF', border: 'EFE3E0',
    pageT1: '1A1110', pageT2: '635451', pageT3: 'A2928F',
    accent: 'B91C1C', accent2: 'DC2626', accentSoft: 'FBE9E9',
    darkBg: '2B0E0E', darkBg2: '3D1414', darkT1: 'FFFFFF', darkT2: 'E0C4C4', darkT3: '96706F',
  },
  'minimal-gray': {
    id: 'minimal-gray',
    name: '极简灰',
    pageBg: 'FAFAFA', surface: 'FFFFFF', border: 'E5E5E5',
    pageT1: '111111', pageT2: '5C5C5C', pageT3: '9C9C9C',
    accent: '111111', accent2: '6B7280', accentSoft: 'F0F0F0',
    darkBg: '141414', darkBg2: '1F1F1F', darkT1: 'FFFFFF', darkT2: 'B8B8B8', darkT3: '7A7A7A',
  },
  'nature-green': {
    id: 'nature-green',
    name: '自然绿',
    pageBg: 'F7FBF7', surface: 'FFFFFF', border: 'DEEBDE',
    pageT1: '0F2014', pageT2: '4E6853', pageT3: '8CA893',
    accent: '15803D', accent2: '22A554', accentSoft: 'E4F5E9',
    darkBg: '0B1D12', darkBg2: '122A1B', darkT1: 'FFFFFF', darkT2: 'B4CDBB', darkT3: '6B8A74',
  },
  'luxury-black': {
    id: 'luxury-black',
    name: '奢华黑金',
    pageBg: 'FBF9F4', surface: 'FFFFFF', border: 'EDE6D6',
    pageT1: '14110B', pageT2: '6B614A', pageT3: 'A79C82',
    accent: '9A7B34', accent2: 'C8A24A', accentSoft: 'F6EFDD',
    darkBg: '0B0B0C', darkBg2: '171614', darkT1: 'F7F3EA', darkT2: 'C9BFA6', darkT3: '8A8069',
  },
};

export const DEFAULT_THEME = 'business-blue';

/** 取主题；未知 id 抛错并列出可选值。 */
export function getTheme(id) {
  const theme = THEMES[id ?? DEFAULT_THEME];
  if (!theme) {
    throw new Error(`未知主题 "${id}"。可选：${Object.keys(THEMES).join(', ')}`);
  }
  return theme;
}

/**
 * 六个模板。装饰性能力全部是"给版式函数用的数据"，不是硬编码。
 * rail：左侧色轨宽度；pageNumber：页码格式；footerTone：页脚墨色。
 */
export const TEMPLATES = {
  consulting: {
    id: 'consulting', name: '咨询简报', defaultTheme: 'business-blue',
    rail: { width: 10 }, pageNumber: '{n} / {total}', footer: true,
  },
  'product-launch': {
    id: 'product-launch', name: '产品发布', defaultTheme: 'tech-cyan',
    rail: null, pageNumber: '{n} / {total}', footer: true, corner: true, coverDark: true,
  },
  'academic-defense': {
    id: 'academic-defense', name: '学术答辩', defaultTheme: 'academic-crimson',
    rail: null, pageNumber: '{n} / {total}', footer: true, plain: true,
  },
  minimal: {
    id: 'minimal', name: '极简', defaultTheme: 'minimal-gray',
    rail: null, pageNumber: null, footer: false, plain: true,
  },
  'gov-report': {
    id: 'gov-report', name: '政务报告', defaultTheme: 'gov-red',
    rail: null, pageNumber: '{n} / {total}', footer: true, topBar: true,
  },
  'annual-report': {
    id: 'annual-report', name: '年度报告', defaultTheme: 'luxury-black',
    rail: { width: 12 }, pageNumber: '{n} / {total}', footer: true, coverDark: true,
  },
};

export function getTemplate(id) {
  if (id == null) return null;
  const tpl = TEMPLATES[id];
  if (!tpl) throw new Error(`未知模板 "${id}"。可选：${Object.keys(TEMPLATES).join(', ')}`);
  return tpl;
}

/**
 * 深色页上的强调色。浅色页的 accent2 放到深色底上往往不够对比
 * （实测 business-blue 的 3B82F6 在 0B1F3F 上只有 4.45:1），
 * 所以深色页的文字强调统一走这张表（每个都 ≥ 6:1）。
 */
const DARK_ACCENT = {
  'business-blue': '60A5FA',
  'academic-crimson': 'F08A9B',
  'tech-cyan': '22D3EE',
  'warm-orange': 'FB923C',
  'gov-red': 'F87171',
  'minimal-gray': '9CA3AF',
  'nature-green': '4ADE80',
  'luxury-black': 'C8A24A',
};

/**
 * 把主题 + 模板解析成版式函数用的样式上下文。
 * @returns {{
 *   theme:object, template:object|null, pageBg:string, darkBg:string, darkBg2:string,
 *   ink:object, darkInk:object, accent:string, accent2:string, accentDark:string, accentSoft:string,
 *   surface:string, border:string, footer:string|null, pageNumberFormat:((n:number,total:number)=>string)|null,
 *   railWidth:number, corner:boolean, topBar:boolean
 * }}
 */
export function resolveStyle(options = {}) {
  const template = getTemplate(options.template);
  const theme = getTheme(options.theme ?? template?.defaultTheme);
  const footer = options.footer !== undefined
    ? options.footer
    : (template?.footer === false ? null : null);
  const fmt = template?.pageNumber ?? '{n} / {total}';
  return {
    theme,
    template,
    pageBg: theme.pageBg,
    darkBg: theme.darkBg,
    darkBg2: theme.darkBg2,
    ink: { t1: theme.pageT1, t2: theme.pageT2, t3: theme.pageT3 },
    darkInk: { t1: theme.darkT1, t2: theme.darkT2, t3: theme.darkT3 },
    accent: theme.accent,
    accent2: theme.accent2,
    accentDark: theme.accentDark ?? DARK_ACCENT[theme.id] ?? theme.accent2,
    accentSoft: theme.accentSoft,
    surface: theme.surface,
    border: theme.border,
    footer,
    pageNumberFormat: fmt === null ? null
      : (n, total) => String(fmt).replace('{n}', String(n).padStart(2, '0')).replace('{total}', String(total).padStart(2, '0')),
    railWidth: template?.rail?.width ?? 0,
    corner: template?.corner === true,
    topBar: template?.topBar === true,
  };
}
