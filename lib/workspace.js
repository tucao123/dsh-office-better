/**
 * @module workspace
 * 会话隔离的文件工作区。
 *
 * 目录优先级：会话 cwd（DSH 工作区）> 已登记 cwd > 配置根/<sessionId>。
 * 与原实现的关键差别：**HTTP 预览路由也能拿到 cwd**。原实现的路由只有 sessionId，
 * 只有在工具先跑过一次之后才知道真实目录，所以面板在生成成功之前永远是空列表；
 * 这里工具运行时会登记 cwd，且路由回退逻辑一致，列表/预览不再依赖执行顺序。
 */
import { existsSync, mkdirSync, readdirSync, statSync } from 'node:fs';
import { isAbsolute, join, resolve } from 'node:path';

const FILENAME_RE = /^[^\\/:*?"<>|\r\n]+\.(docx|xlsx|pptx)$/i;
const SESSION_RE = /^[a-zA-Z0-9_-]+$/;
const SIDECAR_DIR = '.office-better';

export class Workspace {
  constructor(options = {}) {
    this.root = resolve(options.workspaceDir || join(process.env.TEMP ?? '/tmp', 'dsh-office-better'));
    this.registry = new Map();
  }

  sanitizeSession(sessionId) {
    return typeof sessionId === 'string' && SESSION_RE.test(sessionId) ? sessionId : 'default';
  }

  /** 会话目录（惰性创建）。 */
  sessionDir(sessionId, cwd) {
    const sid = this.sanitizeSession(sessionId);
    let dir;
    if (typeof cwd === 'string' && isAbsolute(cwd)) {
      this.registry.set(sid, cwd);
      dir = cwd;
    } else if (this.registry.has(sid)) {
      dir = this.registry.get(sid);
    } else {
      dir = join(this.root, sid);
    }
    if (!existsSync(dir)) mkdirSync(dir, { recursive: true });
    return dir;
  }

  /** 文件名校验：只允许简单文件名 + 三个白名单扩展名。 */
  resolve(sessionId, filename, cwd) {
    if (typeof filename !== 'string' || !FILENAME_RE.test(filename)) {
      throw new Error(`非法文件名 ${JSON.stringify(filename)}：必须是以 .docx / .xlsx / .pptx 结尾的简单文件名（不含路径分隔符）。`);
    }
    if (filename.includes('..')) throw new Error(`非法文件名 ${filename}：不允许 ".."`);
    return join(this.sessionDir(sessionId, cwd), filename);
  }

  /** 列出会话内的 Office 文件。 */
  listFiles(sessionId, cwd) {
    const dir = this.sessionDir(sessionId, cwd);
    const files = [];
    for (const name of readdirSync(dir)) {
      if (!FILENAME_RE.test(name)) continue;
      const stat = statSync(join(dir, name));
      if (!stat.isFile()) continue;
      files.push({
        name,
        type: name.slice(name.lastIndexOf('.') + 1).toLowerCase(),
        size: stat.size,
        mtime: stat.mtimeMs,
      });
    }
    files.sort((a, b) => b.mtime - a.mtime);
    return files;
  }

  /** IR 旁挂文件路径（用于 deck 追加/精确预览）。 */
  sidecarPath(sessionId, filename, cwd) {
    const dir = join(this.sessionDir(sessionId, cwd), SIDECAR_DIR);
    if (!existsSync(dir)) mkdirSync(dir, { recursive: true });
    return join(dir, `${filename}.json`);
  }
}
