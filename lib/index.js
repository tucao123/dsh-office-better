/**
 * dsh-office-better —— DSH 的 Office 文档插件（PPT / DOCX / XLSX）。
 *
 * 设计要点（相对 dsh-officecli 的取舍，详见 ARCHITECTURE.md）：
 *  1. 零外部进程：OOXML 由 node:zlib + 手写 XML 直接产出，不再 spawn officecli
 *  2. IR 单一真源：pptx / 预览 / lint 消费同一份形状 IR，预览与产物必然一致
 *  3. 布局期折行 + 绝对行距（spcPts）：溢出在生成前就算出来，不靠截图回看
 *  4. 预览是纯 HTTP 路由：不启动 watch 子进程、不反代 iframe、不占额外端口
 */
import Schema from '@deepseek-ai/schemastery';
import { Workspace } from './workspace.js';
import { registerTools } from './tools.js';
import { registerRoutes } from './routes.js';

export const name = 'dsh-office-better';

/** 只硬依赖 tools；webServer 可选（headless profile 没有它时工具照常工作）。 */
export const inject = ['tools'];

export const Config = Schema.object({
  workspaceDir: Schema.string()
    .default('')
    .description('工作区根目录（空 = 系统临时目录下的 dsh-office-better；有会话 cwd 时优先用会话 cwd）'),
  previewPrefix: Schema.string()
    .default('/api/office-better')
    .description('预览路由前缀'),
});

export function apply(ctx, config) {
  const options = {
    workspaceDir: typeof config.workspaceDir === 'string' && config.workspaceDir.trim()
      ? config.workspaceDir.trim()
      : undefined,
    previewPrefix: config.previewPrefix || '/api/office-better',
  };
  const workspace = new Workspace({ workspaceDir: options.workspaceDir });
  const deps = { workspace, config: options };

  // 预览路由挂在 webServer 上；用 ctx.inject 而不是同步 ctx.get，
  // 避免"插件激活早于 webServer 提供"导致路由永远注册不上。
  ctx.inject(['webServer'], (webCtx) => {
    webCtx.effect(() => registerRoutes(webCtx, deps), 'dsh-office-better: preview routes');
    webCtx.logger.info(`dsh-office-better: 预览接口已挂载 ${options.previewPrefix}`);
  });

  registerTools(ctx, deps);
  ctx.logger.info('dsh-office-better: 9 个工具就绪（纯进程内 OOXML，无外部依赖）');

  ctx.effect(() => () => {
    ctx.logger.info('dsh-office-better: 已卸载');
  }, 'dsh-office-better: teardown');
}
