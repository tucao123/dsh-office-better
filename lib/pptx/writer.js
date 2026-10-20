/**
 * @module pptx/writer
 * IR -> PPTX 字节流。纯 node:zlib + 手写 OOXML，不依赖任何外部可执行文件。
 */
import { zipSync } from '../util/zip.js';
import {
  XML_HEAD, NS_DECL_PPT, esc, emu, fillXml, paraXml, shapePropsXml, EMPTY_SP_TREE,
} from '../util/xml.js';
import { CANVAS } from './theme.js';

/** 单个形状 -> <p:sp>。role==='bg' 由 slide 背景承担，跳过。 */
export function shapeXml(shape, id) {
  const name = esc(shape.label || shape.role || `shape${id}`);
  const head = `<p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>`;
  const spPr = shapePropsXml(shape);
  if (shape.k !== 'text' || !shape.paras) return `<p:sp>${head}${spPr}</p:sp>`;
  const anchor = shape.anchor === 'ctr' ? 'ctr' : shape.anchor === 'b' ? 'b' : 't';
  const body = `<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="${anchor}" anchorCtr="0"><a:noAutofit/></a:bodyPr>`
    + `<a:lstStyle/>${shape.paras.map(paraXml).join('')}</p:txBody>`;
  return `<p:sp>${head}${spPr}${body}</p:sp>`;
}

/** 一页 -> slide part XML。 */
export function slideXml(page, index) {
  const bg = page.background ? fillXml(page.background) : fillXml(page.pageBg ?? 'FFFFFF');
  const shapes = page.shapes
    .filter((shape) => shape.role !== 'bg')
    .map((shape, i) => shapeXml(shape, i + 2))
    .join('');
  const transition = page.transition === 'fade'
    ? '<p:transition spd="med"><p:fade/></p:transition>'
    : page.transition === 'push' ? '<p:transition spd="med"><p:push/></p:transition>' : '';
  const tree = `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>`
    + '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
    + '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
  return `${XML_HEAD}<p:sld ${NS_DECL_PPT}>`
    + `<p:cSld><p:bg><p:bgPr>${bg}<a:effectLst/></p:bgPr></p:bg><p:spTree>${tree}${shapes}</p:spTree></p:cSld>`
    + `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${transition}</p:sld>`;
}

const EMPTY_PLACEHOLDER_TREE = EMPTY_SP_TREE;

function themeXml(style) {
  const solid = (v) => `<a:solidFill><a:srgbClr val="${v}"/></a:solidFill>`;
  const ln = (w, v) => `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr">${solid(v)}<a:prstDash val="solid"/></a:ln>`;
  return `${XML_HEAD}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="OfficeBetter">`
    + '<a:themeElements><a:clrScheme name="OfficeBetter">'
    + `<a:dk1><a:srgbClr val="${style.ink.t1}"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>`
    + `<a:dk2><a:srgbClr val="${style.theme.darkBg2}"/></a:dk2><a:lt2><a:srgbClr val="${style.pageBg}"/></a:lt2>`
    + `<a:accent1><a:srgbClr val="${style.accent}"/></a:accent1><a:accent2><a:srgbClr val="${style.accent2}"/></a:accent2>`
    + `<a:accent3><a:srgbClr val="${style.accentSoft}"/></a:accent3><a:accent4><a:srgbClr val="${style.ink.t2}"/></a:accent4>`
    + `<a:accent5><a:srgbClr val="${style.border}"/></a:accent5><a:accent6><a:srgbClr val="${style.ink.t2}"/></a:accent6>`
    + `<a:hlink><a:srgbClr val="${style.accent}"/></a:hlink><a:folHlink><a:srgbClr val="${style.ink.t2}"/></a:folHlink>`
    + '</a:clrScheme>'
    + '<a:fontScheme name="OfficeBetter">'
    + '<a:majorFont><a:latin typeface="Segoe UI"/><a:ea typeface="Microsoft YaHei"/><a:cs typeface=""/></a:majorFont>'
    + '<a:minorFont><a:latin typeface="Segoe UI"/><a:ea typeface="Microsoft YaHei"/><a:cs typeface=""/></a:minorFont>'
    + '</a:fontScheme>'
    + '<a:fmtScheme name="OfficeBetter">'
    + `<a:fillStyleLst>${solid(style.surface)}${solid(style.pageBg)}${solid(style.border)}</a:fillStyleLst>`
    + `<a:lnStyleLst>${ln(9525, style.ink.t1)}${ln(19050, style.ink.t1)}${ln(28575, style.ink.t1)}</a:lnStyleLst>`
    + '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle>'
    + '<a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
    + `<a:bgFillStyleLst>${solid('FFFFFF')}${solid(style.pageBg)}${solid(style.border)}</a:bgFillStyleLst>`
    + '</a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>';
}

/**
 * 组装整份 pptx。
 * @param {{pages:Array<{shapes:object[], pageBg:string, background?:string, transition?:string}>,
 *          style:object, meta?:{title?:string, author?:string, subject?:string, keywords?:string}}} input
 * @returns {Buffer}
 */
export function buildPptx(input) {
  const { pages, style } = input;
  const meta = input.meta ?? {};
  const total = pages.length;
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');

  const overrides = [
    ['/ppt/presentation.xml', 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'],
    ['/ppt/slideMasters/slideMaster1.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'],
    ['/ppt/slideLayouts/slideLayout1.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'],
    ['/ppt/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml'],
    ['/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml'],
    ['/docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml'],
    ...pages.map((_p, i) => [`/ppt/slides/slide${i + 1}.xml`, 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml']),
  ];

  const contentTypes = `${XML_HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + overrides.map(([part, type]) => `<Override PartName="${part}" ContentType="${type}"/>`).join('')
    + '</Types>';

  const rootRels = `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>'
    + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
    + '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
    + '</Relationships>';

  const core = `${XML_HEAD}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">`
    + `<dc:title>${esc(meta.title ?? '演示文稿')}</dc:title>`
    + `<dc:creator>${esc(meta.author ?? 'dsh-office-better')}</dc:creator>`
    + `<dc:subject>${esc(meta.subject ?? '')}</dc:subject>`
    + `<dc:description>${esc(meta.description ?? '')}</dc:description>`
    + `<cp:keywords>${esc(meta.keywords ?? '')}</cp:keywords>`
    + `<cp:lastModifiedBy>${esc(meta.author ?? 'dsh-office-better')}</cp:lastModifiedBy>`
    + `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>`
    + `<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>`
    + '</cp:coreProperties>';

  const app = `${XML_HEAD}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">`
    + `<Application>dsh-office-better</Application><Slides>${total}</Slides>`
    + '<PresentationFormat>Widescreen</PresentationFormat></Properties>';

  const slideIds = pages.map((_p, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`).join('');
  const presentation = `${XML_HEAD}<p:presentation ${NS_DECL_PPT} saveSubsetFonts="1">`
    + '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
    + `<p:sldIdLst>${slideIds}</p:sldIdLst>`
    + `<p:sldSz cx="${emu(CANVAS.w)}" cy="${emu(CANVAS.h)}"/><p:notesSz cx="6858000" cy="9144000"/>`
    + '<p:defaultTextStyle><a:defPPr><a:defRPr lang="zh-CN"/></a:defPPr></p:defaultTextStyle>'
    + '</p:presentation>';

  const presentationRels = `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>'
    + pages.map((_p, i) => `<Relationship Id="rId${i + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`).join('')
    + `<Relationship Id="rId${total + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>`
    + '</Relationships>';

  const master = `${XML_HEAD}<p:sldMaster ${NS_DECL_PPT}>`
    + `<p:cSld><p:bg><p:bgPr>${fillXml(style.pageBg)}<a:effectLst/></p:bgPr></p:bg>`
    + `<p:spTree>${EMPTY_PLACEHOLDER_TREE}</p:spTree></p:cSld>`
    + '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
    + '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>'
    + '<p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr lang="zh-CN" sz="2800" b="1">'
    + `<a:solidFill><a:srgbClr val="${style.ink.t1}"/></a:solidFill>`
    + '<a:latin typeface="Segoe UI"/><a:ea typeface="Microsoft YaHei"/></a:defRPr></a:lvl1pPr></p:titleStyle>'
    + '<p:bodyStyle><a:lvl1pPr><a:defRPr lang="zh-CN" sz="1600">'
    + `<a:solidFill><a:srgbClr val="${style.ink.t2}"/></a:solidFill>`
    + '<a:latin typeface="Segoe UI"/><a:ea typeface="Microsoft YaHei"/></a:defRPr></a:lvl1pPr></p:bodyStyle>'
    + '<p:otherStyle><a:lvl1pPr><a:defRPr lang="zh-CN" sz="1600">'
    + '<a:solidFill><a:srgbClr val="333333"/></a:solidFill>'
    + '<a:latin typeface="Segoe UI"/><a:ea typeface="Microsoft YaHei"/></a:defRPr></a:lvl1pPr></p:otherStyle>'
    + '</p:txStyles></p:sldMaster>';

  const masterRels = `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
    + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>'
    + '</Relationships>';

  const layout = `${XML_HEAD}<p:sldLayout ${NS_DECL_PPT} type="blank" preserve="1">`
    + `<p:cSld name="Blank"><p:spTree>${EMPTY_PLACEHOLDER_TREE}</p:spTree></p:cSld>`
    + '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>';

  const layoutRels = `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>'
    + '</Relationships>';

  const slideRels = `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
    + '</Relationships>';

  const entries = [
    { name: '[Content_Types].xml', data: contentTypes },
    { name: '_rels/.rels', data: rootRels },
    { name: 'docProps/core.xml', data: core },
    { name: 'docProps/app.xml', data: app },
    { name: 'ppt/presentation.xml', data: presentation },
    { name: 'ppt/_rels/presentation.xml.rels', data: presentationRels },
    { name: 'ppt/theme/theme1.xml', data: themeXml(style) },
    { name: 'ppt/slideMasters/slideMaster1.xml', data: master },
    { name: 'ppt/slideMasters/_rels/slideMaster1.xml.rels', data: masterRels },
    { name: 'ppt/slideLayouts/slideLayout1.xml', data: layout },
    { name: 'ppt/slideLayouts/_rels/slideLayout1.xml.rels', data: layoutRels },
  ];
  pages.forEach((page, i) => {
    entries.push({ name: `ppt/slides/slide${i + 1}.xml`, data: slideXml(page, i) });
    entries.push({ name: `ppt/slides/_rels/slide${i + 1}.xml.rels`, data: slideRels });
  });
  return zipSync(entries);
}
