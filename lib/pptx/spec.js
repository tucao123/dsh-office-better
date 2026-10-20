/**
 * @module pptx/spec
 * DeckSpec 校验与字数红线。
 *
 * 与原实现的差别：那边把字数上限写在"设计指南"里让模型自己遵守（违反也要等
 * 排版跑出来才发现溢出）；这里在**输入阶段**就硬校验并抛出可操作的错误，
 * 每个字段的超限都会指名道姓地告诉你允许多少字。
 */

/** 每页上限。 */
export const MAX_SLIDES = 40;

/** 各版式的字数红线（超出即报错）。 */
export const LIMITS = {
  cover: { title: 24, subtitle: 44, eyebrow: 12, meta: 28 },
  section: { title: 18, subtitle: 40, number: 4 },
  bullets: { items: 8, itemTitle: 22, itemDesc: 46 },
  cards: { cards: 6, cardTitle: 14, cardDesc: 64 },
  kpi: { metrics: 4, value: 8, label: 12, note: 30 },
  steps: { steps: 5, stepTitle: 12, stepDesc: 40 },
  compare: { points: 6, point: 32, title: 16 },
  timeline: { events: 5, eventTitle: 14, eventDesc: 42, date: 12 },
  quote: { quote: 80, author: 24, role: 24 },
  table: { headers: 5, header: 10, rows: 8, cell: 40 },
  agenda: { items: 6, itemTitle: 16, itemDesc: 42 },
  swot: { perQuadrant: 4, entry: 20 },
  pricing: { plans: 4, name: 14, price: 10, features: 5, feature: 14 },
  roadmap: { phases: 5, phase: 8, phaseTitle: 14, phaseDesc: 40 },
  ending: { title: 14, subtitle: 44 },
};

export class DeckSpecError extends Error {
  constructor(message) {
    super(message);
    this.name = 'DeckSpecError';
  }
}

function fail(where, message) {
  throw new DeckSpecError(`${where}: ${message}`);
}

function str(value, where, max, { required = false } = {}) {
  if (value == null || value === '') {
    if (required) fail(where, '为必填字符串');
    return undefined;
  }
  if (typeof value !== 'string' && typeof value !== 'number') fail(where, `必须是字符串（收到 ${typeof value}）`);
  const text = String(value);
  if (max != null && text.length > max) {
    fail(where, `超长 ${text.length} 字，上限 ${max} 字。请精简："${text.slice(0, 20)}…"`);
  }
  return text;
}

function arr(value, where, max, { min = 1, required = true } = {}) {
  if (value == null) {
    if (required) fail(where, '为必填数组');
    return [];
  }
  if (!Array.isArray(value)) fail(where, `必须是数组（收到 ${typeof value}）`);
  if (value.length < min) fail(where, `至少要 ${min} 项，收到 ${value.length} 项`);
  if (value.length > max) fail(where, `最多 ${max} 项，收到 ${value.length} 项`);
  return value;
}

/** 校验并归一化单页。 */
export function parseSlide(raw, index) {
  const where = `slides[${index}]`;
  if (!raw || typeof raw !== 'object') fail(where, '必须是对象，形如 {"layout":"cover","title":"…"}');
  const layout = String(raw.layout ?? '');
  if (!LIMITS[layout]) {
    fail(`${where}.layout`, `未知版式 "${layout}"。可选：${Object.keys(LIMITS).join(', ')}`);
  }
  const L = LIMITS[layout];
  const out = { layout, transition: raw.transition, background: raw.background };
  if (raw.hidden != null) out.hidden = raw.hidden === true;

  switch (layout) {
    case 'cover':
      out.title = str(raw.title, `${where}.title`, L.title, { required: true });
      out.subtitle = str(raw.subtitle, `${where}.subtitle`, L.subtitle);
      out.eyebrow = str(raw.eyebrow, `${where}.eyebrow`, L.eyebrow);
      out.meta = str(raw.meta, `${where}.meta`, L.meta);
      break;
    case 'section':
      out.title = str(raw.title, `${where}.title`, L.title, { required: true });
      out.number = str(raw.number, `${where}.number`, L.number);
      out.subtitle = str(raw.subtitle, `${where}.subtitle`, L.subtitle);
      break;
    case 'bullets':
      out.title = str(raw.title, `${where}.title`, 30, { required: true });
      out.columns = raw.columns === 2 ? 2 : 1;
      out.items = arr(raw.items, `${where}.items`, L.items).map((item, i) => {
        if (!item || typeof item !== 'object') fail(`${where}.items[${i}]`, '必须是 {title, desc?}');
        return {
          title: str(item.title, `${where}.items[${i}].title`, L.itemTitle, { required: true }),
          desc: str(item.desc, `${where}.items[${i}].desc`, L.itemDesc),
        };
      });
      break;
    case 'cards':
      out.title = str(raw.title, `${where}.title`, 30, { required: true });
      if (raw.columns != null) out.columns = Number(raw.columns);
      out.cards = arr(raw.cards, `${where}.cards`, L.cards).map((card, i) => {
        if (!card || typeof card !== 'object') fail(`${where}.cards[${i}]`, '必须是 {title, desc?, tag?}');
        return {
          title: str(card.title, `${where}.cards[${i}].title`, L.cardTitle, { required: true }),
          desc: str(card.desc, `${where}.cards[${i}].desc`, L.cardDesc),
          tag: str(card.tag, `${where}.cards[${i}].tag`, 8),
        };
      });
      break;
    case 'kpi':
      out.title = str(raw.title, `${where}.title`, 30, { required: true });
      out.metrics = arr(raw.metrics, `${where}.metrics`, L.metrics).map((metric, i) => {
        if (!metric || typeof metric !== 'object') fail(`${where}.metrics[${i}]`, '必须是 {value, label, note?}');
        return {
          value: str(metric.value, `${where}.metrics[${i}].value`, L.value, { required: true }),
          label: str(metric.label, `${where}.metrics[${i}].label`, L.label, { required: true }),
          note: str(metric.note, `${where}.metrics[${i}].note`, L.note),
        };
      });
      break;
    case 'steps':
      out.title = str(raw.title, `${where}.title`, 30, { required: true });
      out.steps = arr(raw.steps, `${where}.steps`, L.steps).map((step, i) => {
        if (!step || typeof step !== 'object') fail(`${where}.steps[${i}]`, '必须是 {title, desc?}');
        return {
          title: str(step.title, `${where}.steps[${i}].title`, L.stepTitle, { required: true }),
          desc: str(step.desc, `${where}.steps[${i}].desc`, L.stepDesc),
        };
      });
      break;
    case 'compare': {
      out.title = str(raw.title, `${where}.title`, 30, { required: true });
      const sideOf = (side, name) => {
        if (!side || typeof side !== 'object') fail(`${where}.${name}`, '必须是 {title, points[]}');
        return {
          title: str(side.title, `${where}.${name}.title`, L.title, { required: true }),
          points: arr(side.points, `${where}.${name}.points`, L.points).map((point, i) => str(point, `${where}.${name}.points[${i}]`, L.point, { required: true })),
        };
      };
      out.left = sideOf(raw.left, 'left');
      out.right = sideOf(raw.right, 'right');
      break;
    }
    case 'timeline':
      out.title = str(raw.title, `${where}.title`, 30, { required: true });
      out.events = arr(raw.events, `${where}.events`, L.events).map((event, i) => {
        if (!event || typeof event !== 'object') fail(`${where}.events[${i}]`, '必须是 {date, title, desc?}');
        return {
          date: str(event.date, `${where}.events[${i}].date`, L.date, { required: true }),
          title: str(event.title, `${where}.events[${i}].title`, L.eventTitle, { required: true }),
          desc: str(event.desc, `${where}.events[${i}].desc`, L.eventDesc),
        };
      });
      break;
    case 'quote':
      out.quote = str(raw.quote, `${where}.quote`, L.quote, { required: true });
      out.author = str(raw.author, `${where}.author`, L.author);
      out.role = str(raw.role, `${where}.role`, L.role);
      break;
    case 'table':
      out.title = str(raw.title, `${where}.title`, 30, { required: true });
      out.headers = arr(raw.headers, `${where}.headers`, L.headers).map((h, i) => str(h, `${where}.headers[${i}]`, L.header, { required: true }));
      out.rows = arr(raw.rows, `${where}.rows`, L.rows).map((row, r) => {
        if (!Array.isArray(row)) fail(`${where}.rows[${r}]`, '必须是字符串数组');
        if (row.length > out.headers.length) fail(`${where}.rows[${r}]`, `有 ${row.length} 列，超过表头 ${out.headers.length} 列`);
        return row.map((cell, c) => str(cell, `${where}.rows[${r}][${c}]`, L.cell) ?? '');
      });
      break;
    case 'agenda':
      out.title = str(raw.title, `${where}.title`, 30, { required: true });
      out.items = arr(raw.items, `${where}.items`, L.items).map((item, i) => {
        if (!item || typeof item !== 'object') fail(`${where}.items[${i}]`, '必须是 {title, desc?}');
        return {
          title: str(item.title, `${where}.items[${i}].title`, L.itemTitle, { required: true }),
          desc: str(item.desc, `${where}.items[${i}].desc`, L.itemDesc),
        };
      });
      break;
    case 'swot': {
      out.title = str(raw.title, `${where}.title`, 30, { required: true });
      for (const key of ['s', 'w', 'o', 't']) {
        out[key] = arr(raw[key], `${where}.${key}`, L.perQuadrant, { required: false, min: 0 })
          .map((entry, i) => str(entry, `${where}.${key}[${i}]`, L.entry, { required: true }));
      }
      break;
    }
    case 'pricing':
      out.title = str(raw.title, `${where}.title`, 30, { required: true });
      out.plans = arr(raw.plans, `${where}.plans`, L.plans).map((plan, i) => {
        if (!plan || typeof plan !== 'object') fail(`${where}.plans[${i}]`, '必须是 {name, price, features[]}');
        return {
          name: str(plan.name, `${where}.plans[${i}].name`, L.name, { required: true }),
          price: str(plan.price, `${where}.plans[${i}].price`, L.price, { required: true }),
          tag: str(plan.tag, `${where}.plans[${i}].tag`, 12),
          highlight: plan.highlight === true,
          features: arr(plan.features, `${where}.plans[${i}].features`, L.features).map((f, j) => str(f, `${where}.plans[${i}].features[${j}]`, L.feature, { required: true })),
        };
      });
      break;
    case 'roadmap':
      out.title = str(raw.title, `${where}.title`, 30, { required: true });
      out.phases = arr(raw.phases, `${where}.phases`, L.phases).map((phase, i) => {
        if (!phase || typeof phase !== 'object') fail(`${where}.phases[${i}]`, '必须是 {phase, title, desc?}');
        return {
          phase: str(phase.phase, `${where}.phases[${i}].phase`, L.phase, { required: true }),
          title: str(phase.title, `${where}.phases[${i}].title`, L.phaseTitle, { required: true }),
          desc: str(phase.desc, `${where}.phases[${i}].desc`, L.phaseDesc),
        };
      });
      break;
    case 'ending':
      out.title = str(raw.title, `${where}.title`, L.title) ?? '谢谢';
      out.subtitle = str(raw.subtitle, `${where}.subtitle`, L.subtitle);
      break;
    default:
      fail(`${where}.layout`, `未实现的版式 "${layout}"`);
  }
  return out;
}

/**
 * 校验整份 DeckSpec。
 * @param {any} input
 * @returns {{template?:string, theme?:string, footer?:string, style?:object, slides:Array<object>}}
 */
export function parseDeckSpec(input) {
  if (typeof input === 'string') {
    try {
      return parseDeckSpec(JSON.parse(input));
    } catch (err) {
      if (err instanceof DeckSpecError) throw err;
      throw new DeckSpecError(`deck 是字符串但 JSON 解析失败：${err.message}`);
    }
  }
  if (!input || typeof input !== 'object') {
    throw new DeckSpecError('deck 必须是对象，形如 {"template":"consulting","theme":"business-blue","slides":[…]}');
  }
  const slides = arr(input.slides, 'slides', MAX_SLIDES, { min: 1 });
  if (input.template != null && typeof input.template !== 'string') fail('template', '必须是字符串');
  if (input.theme != null && typeof input.theme !== 'string') fail('theme', '必须是字符串');
  if (input.footer != null && typeof input.footer !== 'string') fail('footer', '必须是字符串');
  if (input.style != null && typeof input.style !== 'object') fail('style', '必须是对象');
  return {
    template: input.template,
    theme: input.theme,
    footer: input.footer,
    style: input.style,
    slides: slides.map((slide, i) => parseSlide(slide, i)),
  };
}
