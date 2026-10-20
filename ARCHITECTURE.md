# 架构与逆向笔记

这份文档记录两件事：**原插件 dsh-officecli v0.1.1 是怎么做的**，以及**本插件为什么这样重写**。

---

## 1. 原插件拆解

### 1.1 数据流

```
模型（只写内容）
   ↓  DeckSpec {template?, theme?, slides:[{layout, ...content}]}
插件设计层（算坐标/字号/配色/折行策略）
   ↓  OfficeCLI 命令（add shape --prop …，一次 batch 原子提交）
officecli.exe  ← 外部进程，Windows 上只有 .cmd/.ps1 shim
   ↓
pptx
```

预览是另一条链路：

```
office_* 工具执行 → 事件总线 → SSE → 侧边栏面板
面板点文件 → GET /api/officecli/watch?session=&file=
   → WatchManager spawn("officecli watch <file> --port 0")
   → 从 stdout 抓 "Watch: http://127.0.0.1:PORT"
   → 插件反代 /api/officecli/watch/<session>/* 到该端口，iframe 嵌入
```

### 1.2 它的模块职责（逆向所得）

| 模块 | 职责 |
|---|---|
| `lib/index.js` | cordis 插件入口：`inject:['tools']`，`ctx.inject(['webServer'])` 挂路由，挂 `tools/execute`、`tools/result` 广播运行态 |
| `lib/service.js` | `spawn(officecliPath, args, {shell:false})` + `--json` 信封解析，30s/60s 超时，ENOENT 翻译成"未安装"提示 |
| `lib/watch.js` | 每会话一个 watch 子进程，端口自报、`/api/switch` 切文件、空闲退出重启、`taskkill /T /F` 杀进程树 |
| `lib/routes.js` | `/api/officecli` 前缀路由：`/files`、`/events`(SSE)、`/watch`、`/watch/<sid>/*` 代理 |
| `lib/proxy.js` | 反代 watch 服务：改写 Host/Origin 绕开它的门禁，把内嵌 JS 里的根相对路径重写成代理前缀 |
| `lib/workspace.js` | 会话目录解析（会话 cwd > 已登记 cwd > `<root>/<sessionId>`）、文件名白名单防路径逃逸 |
| `lib/pptx/` | **设计层**：`grid`（960×540pt / 12 栏 / 8pt 基线）、`theme`（8 主题）、`templates`（6 套外衣）、`layouts`（15 版式）、`checklist`（字数红线与自检清单）、`shape`（IR → `--prop` 命令） |
| `lib/tools/` | 12 个底层原语工具（create/add/set/remove/move/get/query/view/dump/screenshot/list…）+ 3 个高层 PPT 工具 |

### 1.3 值得继承的设计

1. **内容与排版分离**：模型只填字段，坐标由引擎算。这是它对"AI 生成的 PPT 排版崩坏"的正确解法。
2. **字数红线前置**：把每个字段的上限随指南下发给模型（虽然只做到"告知"，没做到"强制"）。
3. **模板与版式正交**：换模板不动版式，换版式不动模板。
4. **会话隔离的工作区**：文件落进 DSH 会话 cwd，而不是某个全局目录。
5. **文字节流**：工具返回结构化摘要而不是整份文档。

### 1.4 它的四个结构性窟窿

| 窟窿 | 具体表现 | 根因 |
|---|---|---|
| 外部二进制单点 | Windows `spawn` ENOENT；npm 上同名包是另一个产品 | 把 OOXML 生成外包给了 CLI |
| 预览双进程 | watch 服务器 + 反代 + Host/Origin 改写 + 5s 端口上报 | 预览必须由外部渲染器提供 |
| 自检靠肉眼 | "把 PNG 回传给模型，模型看图找溢出" | 布局期没有文本度量，不知道会不会溢出 |
| 列表依赖执行顺序 | 面板 `/files` 只有 sessionId、无 cwd，工具没成功跑过就是空列表 | cwd 只存在于工具上下文，没在路由层回退 |

---

## 2. 重写决策

| 维度 | 原实现 | 本实现 | 理由 |
|---|---|---|---|
| OOXML 生成 | `officecli` 进程 + `--prop` 命令 | 进程内 `node:zlib` + 手写 XML | 去掉单点故障、超时、版本漂移 |
| 排版中间产物 | 直接生成 CLI 命令 | **形状 IR**（`{k,x,y,w,h,fill,line,paras,role}`） | 一份数据喂三个消费者 |
| 折行 | 交给 PowerPoint 自动换行 | **布局期折行**（`ir/text.js`，含避头尾） | 高度可计算 → 溢出可判定 |
| 行距 | `spcPct`（相对字体自然行高） | **`spcPts`（绝对 pt）** | 让 PowerPoint 的行盒 = 引擎模型 |
| 溢出检查 | 截图回传 + 多模态肉眼 | `pptx/lint.js` 静态分析 | 可回归、可 CI、不需要看图 |
| 预览 | watch 子进程 + 反代 iframe | 同一份 IR 渲染 SVG，纯 HTTP 路由 | 零子进程、零端口、必然一致 |
| 字数红线 | 写在指南里"请遵守" | 校验期**硬报错**并指名允许字数 | 违反在输入阶段就被挡下 |
| 文档类型 | docx/xlsx 靠 CLI 原语逐步编辑 | 声明式块列表 / 工作表一次性生成 | 结构化文档不需要布局数学，也不需要外部进程 |
| 读取 | `office_get`/`office_dump` 读 CLI 的路径树 | 解压 + XML 抽取文本大纲 | 自足 |

保留：模板/版式正交、15 种版式命名、8 套主题、字数红线、会话隔离、`ctx.inject(['webServer'])` 的可选依赖写法。

---

## 3. 形状 IR

所有版式函数的输出都是 `shape[]`：

```js
{
  k: 'rect' | 'round' | 'ellipse' | 'text',
  x, y, w, h,                  // pt，画布 960×540
  fill, fillAlpha,             // 'RRGGBB'
  line, lineWidth, lineAlpha,
  adj,                         // 圆角半径占短边比例
  role: 'bg'|'card'|'accent'|'title'|'heading'|'body'|'caption'|'deco',
  label: '卡片 2 说明',         // 人可读，lint 直接引用
  paras: [{ runs:[{t,size,bold,color,spc}], align, lineHeight, spaceBefore }],
  size, color, contentH, anchor
}
```

- `role` 同时驱动三件事：写入器跳过 `bg`、lint 的对比度基准与层级检查、预览的绘制顺序。
- `paras` 里每一段就是**一个物理行**（折行已在布局期完成），因此 pptx 与 SVG 的换行完全一致。
- `label` 让 lint 的报错是"「卡片 2 说明」文本高 30.7pt 超出文本框"，而不是"shape id=37 越界"。

---

## 4. 文本度量与折行

`lib/ir/text.js`：

- **宽度模型**：CJK/全角 = 1.0em，空格/窄字符 ≈ 0.3em，数字 0.56em，大写 0.66em，其余 0.53em；粗体 ×1.045。
- **折行**：贪心 + 软断点（空格/`-`/`/`）+ **避头尾**（`，。、；：？！）】」` 等不落行首）。
- **`flow()`**：给定 `maxW` / `maxH`，从目标字号往下逐 0.5pt 试算，取能整段塞下的最大字号 —— 这是"内容多就自动缩字号"而不是"溢出/截断"。
- **`fitLine()`**：单行场景缩字号，实在不行才截断并记 `warn TEXT_TRUNCATED`。

### 行距：为什么必须用 `spcPts`

第一版用 `spcPct`（比如 130%），结果 PowerPoint 的实测行盒比模型大得多：`spcPct` 是**相对字体自然行高**（YaHei ≈ 1.32em），130% → 实际 1.72em。于是预览和成品的行距不一致，也正是原实现不得不靠截图兜底的原因。

改成 `spcPts = lh × size`（绝对 pt）后，PowerPoint 的行盒与引擎模型一致。实测（PowerPoint COM 逐形状比对 `TextFrame2.TextRange.BoundHeight` 与形状高度）：

| 阶段 | 溢出形状数（8 页 deck） |
|---|---|
| `spcPct` + 紧贴内容的高度 | 48 |
| 改 `spcPts` 之后 | 5 |
| 再给文本框 8% + 1.2pt 余量（吸收首行上沿/末行下沿约 0.8pt/块） | **0** |

---

## 5. lint 规则

| code | 级别 | 判据 |
|---|---|---|
| `OUT_OF_CANVAS` | error | 形状超出 960×540（容差 0.5pt） |
| `TEXT_OVERFLOW` | error | `contentH > h + 0.8`（高度来自折行模型，不是估的） |
| `LINE_TOO_WIDE` | warn | 单行宽 > 盒宽 1pt（PowerPoint 会二次折行） |
| `TEXT_TRUNCATED` | warn | 行尾出现截断省略号 |
| `FONT_TOO_SMALL` | warn | 字号 < 8.5pt |
| `LOW_CONTRAST` | warn | 文字与其下方最上层实心填充的 WCAG 对比度 < 4.5:1（大字号 < 3:1） |
| `TOO_DENSE` | warn | 单页视觉块 > 16 |
| `WEAK_HIERARCHY` | warn | 最大标题字号 < 正文最大字号 × 1.15 |

对比度的"下方填充"是**按绘制顺序回溯**求出来的（跳过 `< 0.98` 的透明填充），所以卡片上的文字会和卡片底色比，而不是和页面底色比。

---

## 6. 预览

| | 原实现 | 本实现 |
|---|---|---|
| 渲染器 | `officecli watch` 子进程（浏览器渲染） | `pptx/preview.js` 直接吐 `<svg>` |
| 传输 | 反代 + Host/Origin 改写 + SSE 双通道 | 一个 prefix 路由，4 个端点 |
| 一致性来源 | 同一个文件被两个渲染器解释 | 同一份 IR 的两种投影 |
| 故障点 | 端口、超时、门禁、进程树 | 无（纯函数 + 文件读） |

路由（`/api/office-better`）：

- `GET /files?session=` → JSON 列表
- `GET /index?session=` → 文件列表 HTML（打开/下载）
- `GET /preview?session=&file=` → pptx 走旁挂 DeckSpec 重编译；docx/xlsx 走大纲
- `GET /raw?session=&file=` → 原始字节

旁挂文件：`<sessionDir>/.office-better/<file>.pptx.json`，存原始 DeckSpec + 选项。`office_deck_add` 靠它原地追加，预览靠它精确重放。

---

## 7. 校验与回归

`test/smoke.mjs`：编译 8 页 deck → 断言 lint 无 error → 断言 ZIP 必需 part 与每个 `slideN.xml.rels` 齐全 → 粗查每个 XML 的标签配平（严格良构性由外部工具复核）。

`test/mount.mjs`：桩 `ctx`（`tools.register` / `ctx.inject` / `webServer.register` / `logger`）跑 `apply()`，端到端验证 9 个工具与 4 个路由，包括 hidden 页跳过、`deck_add` 插入位置、docx 块数、xlsx 工作表、`office_read` 回读、404 行为。

外部复核：`[xml]` 解析全部 part（pptx 27 / docx 9 / xlsx 9 个 part 全部良构）；PowerPoint COM `Presentations.Open` 无修复提示，逐形状几何比对溢出为 0。

---

## 8. 客户端半边（侧边栏预览，v0.2）

静态客户端插件的打包契约由宿主 `@deepseek-ai/dsh-client-modules` 定义，逆向所得：

- 宿主在**启动时**扫描 Loader 条目里声明了 `dsh.client` 的包，解析 `exports["./client"]`，按内容 sha1 生成 rev，暴露为 `/plugins/<包名>/client.js?rev=<hash>`；
- bundle 必须是 `window.__ModuleLoader__.load({ id, factory(require) })` 形态，factory 返回的 `module.exports` 里要有 `apply(ctx)`，可选 `inject`（cordis 服务名数组）；
- `dsh.client.inject` 声明的是**模块依赖**（如 `@deepseek-ai/dsh-client-ui-commands`），不是服务；
- **插件集合变化只在重启后生效**——所以新增客户端 bundle 必须重启 DSH 进程（服务端的新工具/路由可以热加载）。

槽位选择（先用 `Slots.listSubTree` 查证的）：

| 槽位 | 形态 | 用途 |
|---|---|---|
| `conversation.session.header.utilities` | list，scope=session | 放触发按钮；标准 props 带 `sessionId`，正好拿到当前会话 |
| `shell.overlay` | list，scope=root，click-through | 放浮层面板；新 id 是**新增**而非替换，不遮挡宿主任何 UI |

没有选 `details`（右栏）：它是 single 槽，"占位即接管"，会连带失去宿主的工具详情面板。浮层方案对宿主零侵入。

面板本身：文件列表（3 秒轮询 `/files`）+ `iframe` 指向 `/preview`。因为预览 HTML 由宿主同源路由提供，`iframe` 不需要任何跨域/凭据处理，也不需要 SSE —— 原实现那套"双通道 SSE + Host/Origin 改写"彻底不需要。

无浏览器验证：`test/client.mjs` 用桩 `window.__ModuleLoader__` / `React` / `slots` / `fetch` 实跑 `lib/client.js`，断言两个 Slot 注册、按钮标签、点击后的 `/files` 请求、面板标题与文件列表、以及 iframe `src` 的 session/文件名编码。

---

## 9. 已知限制

1. **不含图片/SVG/图表**：`pptx` 只画形状与文本；`add shape` 类原语暂不暴露。需要图片时得自己往 zip 里塞 media part（规划中）。
2. **字宽度量是估算**：非 CJK、非 Segoe UI/YaHei 的字体可能有 ±5% 偏差；偏差会先被 lint 的 `LINE_TOO_WIDE` 抓到，而不是直接溢出。
3. **不修改第三方 Office 文件**：`office_read` 只读；原地编辑只支持本插件自己生成、且有旁挂 IR 的文件。
4. **侧边栏面板需要重启才出现**：客户端 bundle 由宿主启动时扫描（见上节）；面板打开后列表靠 3 秒轮询刷新，不是推送式。
5. **docx 无图片、脚注、页眉页脚、目录域**；xlsx 无图表、条件格式、数据校验。

---

## 10. DeckSpec 样式覆盖（0.2.4）

`deck.style` 之前只有 `transition` 被读取，其余字段（`footer` / `pageNumber` / `background` / `meta`）虽然写进了规范却没人消费——典型的"声明了但没实现"。0.2.4 补齐：

- `parseStyleBlock()` 在输入期校验：未知字段直接报错（防 `pageNumer` 这类拼写错误静默失效），`pageNumber` 只接受 `false | true | 模板串`。
- `resolveStyle({ pageNumber, background })` 统一产出 `pageNumberFormat` 与 `pageBg`，版式层不需要知道来源。
- 单页 `background` 优先于全局；**深色整幅页（cover/section/ending）默认仍用自己的深底**，只有显式给单页 `background` 才覆盖。
- `lintPages` 的对比度基准取自 `page.pageBg`，所以换底色后 lint 跟着换基准——测试里浅色页改成 `FFF8F0` 仍然 0 warning，就是这条链的证据。
- `style.meta` 落到 docProps：`dc:title` / `dc:creator` / `cp:keywords` / `dc:description`（此前 `dc:description` 被错写成 keywords）。

`test/smoke.mjs` 覆盖：全局底色、单页覆盖、深色页不受全局影响、页码关闭、`pageNumber: true` 输出纯数字、自定义页脚、单页与全局转场、meta 写进 `core.xml`、未知字段报错。

---

## 11. 路线图

- [ ] 图片与形状原语（`office_image`、自由形状）
- [x] 客户端侧边栏面板（Slot 注册 + iframe 指向 `/preview`）
- [x] DeckSpec 逐页样式覆盖（`deck.style` + 单页 `background`/`transition`）
- [ ] 面板改为事件推送（宿主侧 SSE `/events`），替掉 3 秒轮询
- [ ] docx 图片、页眉页脚、目录域
- [ ] xlsx 数字格式、条件格式
- [ ] `office_deck` 支持多主题混排
- [ ] lint 阈值可配置（每份 deck 自带 `style.lint`）
