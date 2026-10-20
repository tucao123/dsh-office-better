# dsh-office-better

DSH（DeepSeek Harness）的 Office 文档插件：**确定性**生成 pptx / docx / xlsx，附带自建预览与数值化版式检查。

- **零外部进程**：OOXML 由 `node:zlib` + 手写 XML 直接产出，不 spawn 任何 CLI，不需要 PATH 上有什么可执行文件。
- **单一真源（IR）**：幻灯片先编译成形状 IR（坐标/字号/折行全部算好），pptx 写入器、SVG 预览、lint 三者消费同一份 IR —— 预览与产物必然一致。
- **布局期折行 + 绝对行距**：溢出在生成之前就能算出来，不靠"渲染截图 → 用眼睛找"。
- **预览是纯 HTTP 路由**：不启动 watch 子进程、不反代 iframe、不占额外端口。

## 为什么重写（而不是修 dsh-officecli）

[violetdream/dsh-officecli](https://github.com/violetdream/dsh-officecli) 的架构是 **模型 → DeckSpec → 插件设计层 → `officecli` 可执行文件 → pptx**。设计层（坐标/字号/配色）是它的价值所在，但整条链路压在一个外部二进制上，于是：

1. Windows 上 npm 只生成 `officecli.cmd` / `.ps1`，而插件用 `spawn(..., { shell: false })`——`CreateProcess` 只解析真正的 PE 可执行文件，直接 `ENOENT`；
2. 即使 spawn 通了，npm 上的 `officecli` 是另一个产品（自然语言生成器），没有 `batch` / `watch` / `add --prop` 这些原语；
3. 预览要靠 `officecli watch` 起一个 HTTP 服务再反代 iframe，多一层进程、端口、Host/Origin 改写、5s 端口上报超时；
4. "会不会溢出"只有截图回传之后才能判断，需要多模态模型 + 反复返工。

本插件保留它正确的部分（内容与排版分离、字数红线、模板/主题正交），把**渲染和检查搬进进程内**。详见 [ARCHITECTURE.md](ARCHITECTURE.md)。

## 工具（9 个）

| 工具 | 作用 |
|---|---|
| `office_guide` | 先看约束：15 种版式、8 套主题、6 套模板、逐字段字数红线、DeckSpec 示例 |
| `office_deck` | DeckSpec → pptx（同时返回 lint 结果与预览地址） |
| `office_deck_add` | 给已有 deck 追加页面（沿用同一份 IR 与主题） |
| `office_lint` | 只检查不写文件：溢出 / 越界 / 对比度 / 字号 / 密度 / 层级 |
| `office_preview` | 拿到浏览器预览地址（预览由同一份 IR 渲染） |
| `office_docx` | 块列表 → docx（标题/正文/列表/表格/引用/代码/分页） |
| `office_xlsx` | 多工作表 → xlsx（表头样式、冻结首行、列宽、公式） |
| `office_read` | 读 docx / xlsx / pptx 的文本大纲 |
| `office_list` | 列出会话工作区里的 Office 文档 |

## 安装

```bash
# 1) 打包（在插件目录）
npm pack

# 2) 装进 profile（用 tarball，别用 link：link 会让宿主包解析失败）
cd "$DSH_HOME/profiles/web"        # Windows: C:\Users\<你>\.dsh\profiles\web
pnpm add <上一步生成的 dsh-office-better-0.1.0.tgz>

# 3) 把 bundle 写进 package.json 的 dsh.profile.bundles
#    （dsh plugin add 只安装，不改配置）
```

```json
{
  "dsh": {
    "profile": {
      "bundles": ["...", "dsh-office-better"]
    }
  }
}
```

改完 host 组合需要**重启 DSH 进程**才生效。

## 快速开始

```
office_deck({
  filename: "季度汇报.pptx",
  deck: {
    template: "consulting",
    theme: "business-blue",
    slides: [
      { layout: "cover", eyebrow: "2026 Q3", title: "季度经营回顾", subtitle: "增长、留存与下季度动作" },
      { layout: "kpi", title: "三个关键数", metrics: [
        { value: "18%", label: "营收同比" },
        { value: "92%", label: "续费率" },
        { value: "1.2w", label: "新增客户" }]},
      { layout: "bullets", title: "本季度三件事", columns: 2, items: [
        { title: "渠道结构优化", desc: "直营占比提升到 64%，获客成本下降 11%" }]},
      { layout: "table", title: "分区域表现",
        headers: ["区域", "营收", "同比"], rows: [["华东", "4,200 万", "+21%"]]},
      { layout: "ending", title: "谢谢" }
    ]
  }
})
```

返回值里带着 `errors` / `warnings` 和预览地址；有 error 就改 spec 重新生成，不用截图看。

## 15 种版式

`cover`、`section`、`bullets`、`cards`、`kpi`、`steps`、`compare`、`timeline`、`quote`、`table`、`agenda`、`swot`、`pricing`、`roadmap`、`ending`

## 8 套主题 / 6 套模板

主题：`business-blue`、`academic-crimson`、`tech-cyan`、`warm-orange`、`gov-red`、`minimal-gray`、`nature-green`、`luxury-black`

模板：`consulting`、`product-launch`、`academic-defense`、`minimal`、`gov-report`、`annual-report`

## 预览

两种入口，读同一份 `/api/office-better` 路由：

**① 侧边栏浮层（v0.2 起）**：会话头部右上角多一个 `Office` 按钮，点开右侧浮层——文件列表（每 3 秒自动刷新）+ 内嵌 iframe 预览，切换文件即时换，不占独立标签页。

**② 浏览器页面**：`office_preview` 给的地址（`/api/office-better/index?session=…`）直接打开：

- 列表页 → 每个文件的「打开 / 下载」
- pptx 预览：读旁挂的 DeckSpec → 重新编译 → 逐页 `<svg>`，与导出的 pptx 同源
- docx / xlsx 预览：文本大纲

> 客户端半边由宿主 `dsh-client-modules` 在**启动时**扫描 `dsh.client` 声明并提供
> `/plugins/<包名>/client.js`，所以新增/升级客户端 bundle 后需要重启 DSH 进程；
> 只改服务端（工具、路由）可以热加载。

## 测试

```bash
node test/smoke.mjs   # 编译 8 页 deck，校验 ZIP/XML 结构与 lint
node test/mount.mjs   # 桩 ctx 跑 apply()，验证 9 个工具 + 预览路由端到端
node test/client.mjs  # 桩 __ModuleLoader__/React/slots，验证 2 个 Slot 注册与 iframe src
```

产物同时在 PowerPoint 里验证过：`Presentations.Open` 无修复提示，逐形状比对
`TextFrame2.TextRange.BoundHeight/BoundWidth` 与形状尺寸，溢出数为 0。

## 目录

```
lib/
  index.js          cordis 插件入口（工具 + 可选 webServer 路由）
  tools.js          9 个工具定义
  routes.js         预览 HTTP 路由
  workspace.js      会话隔离工作区
  util/xml.js       OOXML 原语（转义/EMU/填充/描边/段落）
  util/zip.js       ZIP 读写（CRC32 + deflateRaw）
  ir/text.js        字宽度量与折行（避头尾）
  pptx/theme.js     8 主题 + 6 模板
  pptx/spec.js      DeckSpec 校验与字数红线
  pptx/layouts.js   15 种版式 → 形状 IR
  pptx/lint.js      溢出/越界/对比度/字号/密度/层级
  pptx/preview.js   IR → SVG/HTML
  pptx/writer.js    IR → pptx
  docx/docx.js      块列表 → docx
  xlsx/xlsx.js      工作表 → xlsx
  read/outline.js   docx/xlsx/pptx → 文本大纲
```

## 许可

MIT
