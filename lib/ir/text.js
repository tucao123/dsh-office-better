/**
 * @module ir/text
 * 字体度量与折行。
 *
 * 这是本插件相对原实现的核心改进之一：**折行在布局期就确定**。
 * 原实现把文本整段交给 officecli/PowerPoint 去自动换行，于是「会不会溢出」
 * 只有截图之后才知道，模型只能靠看图反复返工（它的 README 把"截图回传"当成
 * 卖点，本质是在给缺度量这件事打补丁）。这里有确定的宽度估算：
 *   - 布局期折行 -> IR 里存的是物理行
 *   - PPTX 用同样的行切开 -> 所见即所得
 *   - SVG 预览用同一份行 -> 预览和成品一致
 *   - lint 用同一份行 -> 溢出变成可计算的数字，而不是看图
 */

/** 需要按 1em 计算的字符区间（CJK、全角标点、谚文、假名等）。 */
const WIDE_RE = /[\u1100-\u115F\u2E80-\u303E\u3041-\u33FF\u3400-\u4DBF\u4E00-\u9FFF\uA000-\uA4CF\uAC00-\uD7A3\uF900-\uFAFF\uFE10-\uFE19\uFE30-\uFE6F\uFF00-\uFF60\uFFE0-\uFFE6\u3000-\u303F]/;

/** 拉丁字符的近似宽度（em）。基于 Segoe UI 平均字宽经验值。 */
const NARROW = /[iIlj.,;:'!|`ft()\[\]{}]/;
const DIGIT = /[0-9]/;
const UPPER = /[A-Z]/;
const SPACE = /[\s\u00A0]/;

/** 单个字符宽度（pt）。bold 略宽。 */
export function charWidth(ch, size, bold = false) {
  let ratio;
  if (WIDE_RE.test(ch)) ratio = 1.0;
  else if (SPACE.test(ch)) ratio = 0.3;
  else if (NARROW.test(ch)) ratio = 0.3;
  else if (DIGIT.test(ch)) ratio = 0.56;
  else if (UPPER.test(ch)) ratio = 0.66;
  else ratio = 0.53;
  return size * ratio * (bold ? 1.045 : 1);
}

/** 一段文本的自然宽度（pt）。 */
export function measure(text, size, bold = false) {
  let width = 0;
  for (const ch of String(text)) width += charWidth(ch, size, bold);
  return width;
}

/** 不允许出现在行首的标点（避头尾）。 */
const NO_LINE_START = '，。、；：？！）】》」』”’,.;:?!)]}%';

/**
 * 贪心折行：返回物理行数组。
 * 规则：
 *  - 超宽就断；断点优先落在空格/连字符/中文之间；
 *  - 避头尾：不让 NO_LINE_START 里的标点落到行首；
 *  - 单字符行也放不下时直接容忍（由 lint 报告 overflow），不静默丢字。
 * @param {string} text
 * @param {number} size 字号 pt
 * @param {number} maxWidth 可用宽度 pt
 * @param {boolean} bold
 * @returns {string[]}
 */
export function wrap(text, size, maxWidth, bold = false) {
  const source = String(text ?? '');
  if (source === '') return [''];
  const lines = [];
  let line = '';
  let width = 0;
  const flush = (forceBreakAt = -1) => {
    if (forceBreakAt >= 0) {
      lines.push(line.slice(0, forceBreakAt));
      line = line.slice(forceBreakAt);
      width = measure(line, size, bold);
    } else {
      lines.push(line);
      line = '';
      width = 0;
    }
  };
  for (const ch of source) {
    if (ch === '\n') { flush(); continue; }
    const cw = charWidth(ch, size, bold);
    if (width + cw <= maxWidth || line === '') {
      line += ch;
      width += cw;
      continue;
    }
    // 需要断行：若末字符是空格可直接断，否则回退到最近的软断点
    let cut = line.length;
    if (!SPACE.test(line[line.length - 1])) {
      const soft = Math.max(line.lastIndexOf(' '), line.lastIndexOf('-'), line.lastIndexOf('/'));
      if (soft > 0) cut = soft + 1;
    }
    // 避头尾：下一行首个字符是标点时，把上一行最后一个字符一起带下去
    let carry = '';
    if (NO_LINE_START.includes(ch) && cut > 1) {
      carry = line.slice(cut - 1);
      cut -= 1;
    }
    if (cut <= 0) cut = line.length;
    lines.push((line.slice(0, cut) + carry).replace(/\s+$/, ''));
    line = line.slice(cut);
    width = measure(line, size, bold);
    line += ch;
    width += cw;
  }
  if (line !== '') lines.push(line.replace(/\s+$/, ''));
  return lines.length > 0 ? lines : [''];
}

/** 行数 -> 文本块高度（pt）。 */
export function blockHeight(lineCount, size, lineHeight = 1.35, spaceBefore = 0) {
  return lineCount * size * lineHeight + spaceBefore;
}
