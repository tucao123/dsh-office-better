/**
 * @module routes
 * HTTP 预览路由。挂在 webServer 的 prefix 上，只有 session+file 两个参数。
 *
 * pptx 走「旁挂 DeckSpec -> 重编译 -> SVG」，docx/xlsx 走文本大纲渲染。
 * 不启动子进程、不反代、不需要 Host/Origin 改写 —— 原实现那一整套
 * （officecli watch + 端口上报 + SSE 双通道 + iframe 代理）在这里完全不需要。
 */
import { existsSync, readFileSync, statSync } from 'node:fs';
import { esc } from './util/xml.js';
import { previewHtml } from './pptx/preview.js';
import { compileDeck } from './pptx/deck.js';
import { readOffice } from './read/outline.js';

const HTML_HEAD = (title) => `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>${esc(title)} · 预览</title>
<style>
  body { margin:0; padding:26px 20px 60px; background:#10131a; color:#cfd6e4;
    font:13px/1.65 'Segoe UI','Microsoft YaHei',sans-serif; }
  .wrap { max-width:1000px; margin:0 auto; }
  h1 { font-size:16px; color:#e8edf6; margin:0 0 6px; font-weight:600; }
  .meta { color:#7c8798; font-size:12px; margin-bottom:20px; }
  pre { background:#171b24; border:1px solid #232a36; border-radius:8px; padding:16px;
    overflow:auto; font:12px/1.7 Consolas,monospace; color:#c8d2e0; white-space:pre-wrap; }
  .note { background:#1b2230; border-left:3px solid #6fd3f2; padding:10px 14px; border-radius:0 6px 6px 0;
    color:#a9b6c8; margin-bottom:16px; }
  a { color:#6fd3f2; }
</style></head><body><div class="wrap">`;

/** 文本大纲预览页（docx / xlsx / 非本插件生成的 pptx）。 */
export function previewHtmlForOutline(fileName, outline, note) {
  const body = outline.sections.map((section) => {
    const title = outline.kind === 'xlsx' ? section.name : `${outline.unit} ${section.index}`;
    return `<h2 style="font-size:13px;color:#9fb0c6;margin:18px 0 8px">${esc(title)}</h2>`
      + `<pre>${esc(section.lines.join('\n') || '（空）')}</pre>`;
  }).join('');
  return `${HTML_HEAD(fileName)}<h1>${esc(fileName)}</h1>`
    + `<div class="meta">${outline.kind} · ${outline.count} ${outline.unit}</div>`
    + (note ? `<div class="note">${esc(note)}</div>` : '')
    + body
    + '</div></body></html>';
}

function sendHtml(res, html, code = 200) {
  const data = Buffer.from(html, 'utf8');
  res.writeHead(code, { 'content-type': 'text/html; charset=utf-8', 'content-length': data.length });
  res.end(data);
}

function sendJson(res, code, body) {
  const data = Buffer.from(JSON.stringify(body), 'utf8');
  res.writeHead(code, { 'content-type': 'application/json; charset=utf-8', 'content-length': data.length });
  res.end(data);
}

function sendError(res, code, message) {
  sendHtml(res, `${HTML_HEAD('出错')}<h1>出错了（HTTP ${code}）</h1><div class="note">${esc(message)}</div></div></body></html>`, code);
}

/**
 * 注册预览路由。
 * @param {import('@deepseek-ai/cordis').Context} ctx
 * @param {{workspace: import('./workspace.js').Workspace, config: object}} deps
 * @returns {() => void} 注销函数
 */
export function registerRoutes(ctx, deps) {
  const webServer = ctx.get('webServer');
  if (webServer === undefined) {
    ctx.logger.warn('dsh-office-better: webServer 不可用，预览接口未注册（工具仍可用）');
    return () => {};
  }
  const prefix = (deps.config.previewPrefix ?? '/api/office-better').replace(/\/$/, '');
  const { workspace } = deps;

  return webServer.register({
    kind: 'prefix',
    path: prefix,
    async handler(req, res) {
      const url = new URL(req.url ?? '/', 'http://127.0.0.1');
      const sub = url.pathname.startsWith(prefix) ? url.pathname.slice(prefix.length) : url.pathname;
      const sessionId = url.searchParams.get('session') ?? 'default';
      const file = url.searchParams.get('file') ?? '';
      try {
        if (sub === '/files') {
          return sendJson(res, 200, { files: workspace.listFiles(sessionId) });
        }
        if (sub === '/index' || sub === '/' || sub === '') {
          const files = workspace.listFiles(sessionId);
          return sendHtml(res, indexHtmlFor(files, sessionId));
        }
        if (!file) return sendError(res, 400, '缺少 file 参数');
        const abs = workspace.resolve(sessionId, file);
        if (!existsSync(abs)) return sendError(res, 404, `文件不存在：${file}`);

        if (sub === '/raw') {
          const stat = statSync(abs);
          const buf = readFileSync(abs);
          res.writeHead(200, {
            'content-type': 'application/vnd.openxmlformats-officedocument',
            'content-length': buf.length,
            'content-disposition': `attachment; filename*=UTF-8''${encodeURIComponent(file)}`,
            'last-modified': stat.mtime.toUTCString(),
          });
          return res.end(buf);
        }

        if (sub === '/preview') {
          const kind = file.slice(file.lastIndexOf('.') + 1).toLowerCase();
          if (kind === 'pptx') {
            const sidecar = workspace.sidecarPath(sessionId, file);
            if (existsSync(sidecar)) {
              const stored = JSON.parse(readFileSync(sidecar, 'utf8'));
              const compiled = compileDeck(stored.spec, stored.options ?? {});
              return sendHtml(res, compiled.html);
            }
            const outline = readOffice(readFileSync(abs), kind);
            return sendHtml(res, previewHtmlForOutline(
              file,
              outline,
              '这个 pptx 不是本插件生成的（没有 IR 旁挂文件），只能展示文本大纲；渲染预览需要 office_deck 重新生成。',
            ));
          }
          const outline = readOffice(readFileSync(abs), kind);
          return sendHtml(res, previewHtmlForOutline(file, outline));
        }

        return sendError(res, 404, `未知路由 ${sub}`);
      } catch (err) {
        return sendError(res, 500, err instanceof Error ? err.message : String(err));
      }
    },
  });
}

function indexHtmlFor(files, sessionId) {
  const rows = files.length === 0
    ? '<div class="note">当前会话工作区里还没有 Office 文档。用 office_deck / office_docx / office_xlsx 生成一个。</div>'
    : '<pre>' + files.map((f) => `${f.name}\n  ${f.type} · ${(f.size / 1024).toFixed(1)} KB · ${new Date(f.mtime).toLocaleString('zh-CN')}`
      + `\n  预览 ${prefixLink('preview', f.name, sessionId)}   下载 ${prefixLink('raw', f.name, sessionId)}`).join('\n\n') + '</pre>';
  return `${HTML_HEAD('Office 文档')}<h1>Office 文档</h1>`
    + `<div class="meta">会话 ${esc(sessionId)} · ${files.length} 个文件</div>${rows}`
    + '<div class="meta" style="margin-top:22px">预览由版式 IR 直接渲染，与导出的 pptx 一致；不依赖任何外部渲染器。</div>'
    + '</div></body></html>';
}

function prefixLink(kind, file, sessionId) {
  return `<a href="${kind}?session=${encodeURIComponent(sessionId)}&file=${encodeURIComponent(file)}">${kind === 'preview' ? '打开' : '下载'}</a>`;
}
