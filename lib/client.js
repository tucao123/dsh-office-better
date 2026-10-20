/**
 * dsh-office-better 的 Client half：会话头部挂一个「Office」按钮，
 * 点开在右侧浮层里嵌 iframe 预览（同一个 /api/office-better 路由，零子进程）。
 *
 * 打包契约（与宿主 dsh-client-modules 一致）：
 *   window.__ModuleLoader__.load({ id, factory(require) })
 *   factory 返回 module.exports，其中 apply(ctx) 是 cordis 插件体，
 *   inject 是需要的 cordis 服务列表（slots / timer）。
 */
window.__ModuleLoader__.load({
  id: 'dsh-office-better',
  factory: (require) => {
    var module = { exports: {} };
    var exports = module.exports;

    const React = require('react');
    const h = React.createElement;

    const PREFIX = '/api/office-better';

    /* ---------------------------------------------------------------- 极简 store */

    const store = {
      open: false,
      sessionId: null,
      file: null,
      files: [],
      loading: false,
      error: null,
      updatedAt: 0,
    };
    const listeners = new Set();
    const emit = () => { for (const fn of listeners) fn(); };
    const patch = (next) => { Object.assign(store, next); emit(); };

    function useStore() {
      const [, force] = React.useState(0);
      React.useEffect(() => {
        const fn = () => force((n) => n + 1);
        listeners.add(fn);
        return () => { listeners.delete(fn); };
      }, []);
      return store;
    }

    async function loadFiles(sessionId) {
      if (!sessionId) return;
      patch({ loading: true, error: null });
      try {
        const res = await fetch(`${PREFIX}/files?session=${encodeURIComponent(sessionId)}&_=${Date.now()}`);
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const data = await res.json();
        const files = Array.isArray(data.files) ? data.files : [];
        const keep = store.file && files.some((f) => f.name === store.file) ? store.file : (files[0]?.name ?? null);
        patch({ files, file: keep, loading: false, updatedAt: Date.now() });
      } catch (err) {
        patch({ loading: false, error: err instanceof Error ? err.message : String(err) });
      }
    }

    const previewUrl = (sessionId, file) => `${PREFIX}/preview?session=${encodeURIComponent(sessionId)}&file=${encodeURIComponent(file)}`;

    /* ---------------------------------------------------------------- 样式 */

    let styleEl = null;
    function ensureStyle() {
      if (styleEl !== null || typeof document === 'undefined') return;
      const css = `
        .dob-panel{position:fixed;top:0;right:0;bottom:0;width:min(58vw,760px);min-width:420px;
          display:flex;flex-direction:column;pointer-events:auto;z-index:60;
          background:var(--dsw-alias-bg-layer-1,#141821);color:var(--dsw-alias-label-primary,#e8edf6);
          border-left:1px solid var(--dsw-alias-border-l1,#252b36);
          box-shadow:-18px 0 48px rgba(0,0,0,.45);font:13px/1.5 'Segoe UI','Microsoft YaHei',sans-serif}
        .dob-bar{display:flex;align-items:center;gap:8px;padding:8px 10px;
          border-bottom:1px solid var(--dsw-alias-border-l1,#252b36);flex:none}
        .dob-title{font-weight:600;font-size:12.5px;letter-spacing:.02em}
        .dob-dot{width:6px;height:6px;border-radius:50%;background:#22d3ee;flex:none}
        .dob-spacer{flex:1}
        .dob-btn{appearance:none;border:1px solid var(--dsw-alias-border-l1,#2a3140);
          background:var(--dsw-alias-bg-layer-2,#1b212c);color:inherit;border-radius:6px;
          padding:3px 8px;font-size:11.5px;cursor:pointer;line-height:1.6}
        .dob-btn:hover{background:var(--dsw-alias-bg-layer-3,#232a37)}
        .dob-btn[aria-pressed="true"]{border-color:#0e7490;color:#7ad9ee}
        .dob-list{display:flex;gap:6px;overflow-x:auto;padding:7px 10px;flex:none;
          border-bottom:1px solid var(--dsw-alias-border-l1,#252b36)}
        .dob-file{appearance:none;border:1px solid transparent;background:transparent;color:inherit;
          border-radius:6px;padding:3px 8px;font-size:11.5px;cursor:pointer;white-space:nowrap;
          max-width:220px;overflow:hidden;text-overflow:ellipsis}
        .dob-file:hover{background:var(--dsw-alias-bg-layer-2,#1b212c)}
        .dob-file[aria-current="true"]{background:#0e7490;color:#fff}
        .dob-empty{padding:18px 12px;color:var(--dsw-alias-label-secondary,#93a0b4);font-size:12px}
        .dob-frame{flex:1;border:0;background:#0b0f16;width:100%}
        .dob-err{color:#f87171;font-size:11.5px;padding:0 10px 6px}
      `;
      styleEl = document.createElement('style');
      styleEl.setAttribute('data-dsh-office-better', '');
      styleEl.textContent = css;
      document.head.appendChild(styleEl);
    }

    /* ---------------------------------------------------------------- 面板 */

    function Panel() {
      const s = useStore();
      ensureStyle();
      React.useEffect(() => {
        if (!s.open) return;
        if (s.sessionId) void loadFiles(s.sessionId);
        const t = window.setInterval(() => { if (store.open && store.sessionId) void loadFiles(store.sessionId); }, 3000);
        return () => window.clearInterval(t);
      }, [s.open, s.sessionId]);
      if (!s.open) return null;
      const file = s.file;
      return h('div', { className: 'dob-panel', role: 'region', 'aria-label': 'Office 文档预览' },
        h('div', { className: 'dob-bar' },
          h('span', { className: 'dob-dot' }),
          h('span', { className: 'dob-title' }, 'Office 预览'),
          h('span', { className: 'dob-spacer' }),
          h('button', {
            className: 'dob-btn', type: 'button', title: '立即刷新',
            onClick: () => { if (s.sessionId) void loadFiles(s.sessionId); },
          }, s.loading ? '刷新中…' : '刷新'),
          h('a', {
            className: 'dob-btn', href: `${PREFIX}/index?session=${encodeURIComponent(s.sessionId ?? '')}`,
            target: '_blank', rel: 'noreferrer', title: '在新标签页打开',
          }, '新标签'),
          h('button', {
            className: 'dob-btn', type: 'button', title: '关闭',
            onClick: () => patch({ open: false }),
          }, '关闭'),
        ),
        s.error ? h('div', { className: 'dob-err' }, `取文件列表失败：${s.error}`) : null,
        h('div', { className: 'dob-list' },
          s.files.length === 0
            ? h('span', { className: 'dob-empty' }, '还没有 Office 文档。让 Agent 生成一个（office_deck / office_docx / office_xlsx）。')
            : s.files.map((f) => h('button', {
              key: f.name, type: 'button', className: 'dob-file',
              'aria-current': f.name === file ? 'true' : 'false',
              title: `${f.name} · ${(f.size / 1024).toFixed(1)} KB`,
              onClick: () => patch({ file: f.name }),
            }, f.name)),
        ),
        file
          ? h('iframe', {
            className: 'dob-frame',
            key: `${s.sessionId}:${file}`,
            src: previewUrl(s.sessionId, file),
            title: `预览 ${file}`,
          })
          : h('div', { className: 'dob-empty' }, '选一个文件看预览。列表每 3 秒自动刷新。'),
      );
    }

    /* ---------------------------------------------------------------- 触发按钮 */

    function HeaderButton(props) {
      const s = useStore();
      // sessionId 变化时登记到 store（放在 effect 里，避免渲染期触发别的组件 setState）
      React.useEffect(() => {
        if (props.sessionId) patch({ sessionId: props.sessionId });
      }, [props.sessionId]);
      return h('button', {
        type: 'button',
        className: 'dob-btn',
        'aria-pressed': s.open ? 'true' : 'false',
        title: '打开 Office 文档实时预览',
        onClick: () => {
          const next = !store.open;
          patch({ open: next, sessionId: props.sessionId });
          if (next) void loadFiles(props.sessionId);
        },
      }, s.open ? 'Office ●' : 'Office');
    }

    /* ---------------------------------------------------------------- 插件体 */

    function apply(ctx) {
      const slots = ctx.get('slots');
      if (slots === undefined) return;
      slots.inject('conversation.session.header.utilities', () => slots.register(
        { name: 'conversation.session.header.utilities', id: 'office-better-toggle', order: 10 },
        (props) => h(HeaderButton, props),
      ));
      slots.inject('shell.overlay', () => slots.register(
        { name: 'shell.overlay', id: 'office-better-panel', order: 20 },
        () => h(Panel),
      ));
      ctx.effect(() => () => {
        if (styleEl !== null && styleEl.parentNode) styleEl.parentNode.removeChild(styleEl);
        styleEl = null;
      }, 'dsh-office-better(client): styles');
    }

    exports.apply = apply;
    exports.inject = ['slots'];
    return module.exports;
  },
});
