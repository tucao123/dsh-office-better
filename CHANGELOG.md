# 更新日志

## 0.3.0

- 支持从 npm 安装：`plugin_manager({ action: "install_bundle", target: "dsh-office-better@0.3.0" })`
- README 安装章节拆成「从 npm 装」与「本地检出装」两条路径
- `publishConfig.access = public`，补齐 `author`

## 0.2.4

- 实现 `deck.style` 覆盖（此前只有 `transition` 生效）：
  - `footer` / `pageNumber`（false 隐藏 / true 纯数字 / 模板串）/ `background` / `transition` / `meta`
  - 单页 `background`、`transition` 优先于全局；深色整幅页保留主题深底
  - 未知 style 字段直接报错，不再静默忽略
- 修正 `dc:description`（此前被错写成 keywords），新增 `cp:keywords`
- 补齐插件展示元数据：`locale/zh.json`、`locale/en.json`、`icon.svg`
- `peerDependencies` 放宽到 `^0.1.7-rc.1 || ^0.2.0-rc.1`，适配 dsh 0.2.1-rc.2 的版本门禁
- 测试：`test/smoke.mjs` 新增 10 条样式断言

## 0.2.2

- 修掉深色页文字对比度：新增 `DARK_ACCENT` 表（8 套主题各给一个 ≥6:1 的深底强调色），封面眉标/版本行与章节号改用 `accentDark`
  - 实测：business-blue 4.45:1 → 6.44:1、academic-crimson 2.98 → 7.52、gov-red 3.71 → 6.48、minimal-gray 3.81 → 7.26

## 0.2.1

- 客户端半边 `lib/client.js`：会话头部 `Office` 按钮 + `shell.overlay` 浮层预览（文件列表 3 秒轮询 + iframe 指向 `/preview`）
- `test/client.mjs`：桩 `__ModuleLoader__`/`React`/`slots`/`fetch` 的无浏览器验证

## 0.1.0

- 首个版本：`office_deck` / `deck_add` / `lint` / `preview` / `docx` / `xlsx` / `read` / `list` / `guide` 九个工具
- 15 种幻灯片版式、8 套主题、6 套模板
- 单一真源形状 IR：pptx 写入器、SVG 预览、lint 三路消费
- 布局期折行 + 绝对行距（`spcPts`），零外部进程、零运行时依赖
