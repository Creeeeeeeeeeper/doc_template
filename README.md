# Doc Template

一个用 Tauri 2 + Vite 写的桌面应用，把"在 Word 里挖空 + 别人填表"这件事拆成三个模式：

1. **制作模板** 📝：导入任意 `.docx`，在原格式预览里点光标 / 选文字，插入 `{@field}`（文字）或 `{%field}`（图片）占位符，保存为模板。占位符在编辑区渲染为高亮块，支持双击编辑、右键删除、Backspace 整块删除。
2. **修改模板** 🔧：导入做好的模板，重命名 / 删除字段、改描述 —— 重命名会自动同步文档里所有占位符。双击任意占位符可修改名称和该位置的样式。
3. **填写模板** ✏️：导入模板，应用自动识别字段并生成表单。同一字段名出现在多处时，各处独立保存字体/字号/颜色，文本内容联动同步。图片字段直接上传 PNG/JPG，最后导出填好的 `.docx`。

整个流程都在本地完成，不需要服务端、不需要 Word/WPS 安装。打包后是个 ~10MB 的单文件桌面程序。

---

## 目录

- [功能特性](#功能特性)
- [快速开始](#快速开始)
- [使用说明](#使用说明)
- [模板语法](#模板语法)
- [项目结构](#项目结构)
- [架构与技术栈](#架构与技术栈)
- [DOCX 引擎工作原理](#docx-引擎工作原理)
- [测试与调试](#测试与调试)
- [已知陷阱与设计决策](#已知陷阱与设计决策)
- [后续可拓展](#后续可拓展)
- [已实现 / 完成](#已实现--完成)

---

## 功能特性

| 类别 | 能力 |
|---|---|
| 制作模板 | 在原格式预览里点光标 / 选文字插入占位符 |
|  | 占位符渲染为高亮块（蓝色=文字，粉色=图片），`contenteditable` 编辑 |
|  | Backspace / Delete 命中占位符时整块删除，不按字符删 |
|  | 右键占位符弹出"删除占位符"菜单 |
|  | 双击占位符弹出编辑框，可修改名称和该位置的字体 / 字号 / 颜色 |
|  | 鼠标悬停占位符 0.5s 显示详情 tooltip（名称、类型、当前样式） |
|  | Word 视图预览（按页渲染，遵循原文档页面尺寸与页边距） |
|  | 预览缩放：默认适应宽度、底部缩放开关、`Ctrl + 滚轮`缩放 |
|  | 段落级编辑（保留 pPr、第一个 rPr 作为默认样式） |
|  | 编辑区失焦自动取消预览高亮 |
|  | 字段统一管理：名称、描述 |
|  | 复杂段落（含图片、超链接、域）警告并保护 |
| 修改模板 | 字段重命名 —— 同步替换文档里所有 `{@old}`→`{@new}` |
|  | 字段删除 —— 一键抹掉所有占位符 + 元数据 |
|  | 字段汇总区显示出现次数（`name ×3`）和未使用元数据 |
|  | 双击占位符可修改名称（全局同步重命名）和该位置样式 |
|  | per-occurrence 样式：同一字段名在不同位置可有不同字体 / 字号 / 颜色 |
|  | 修改后默认覆盖原文件（不再追加 `-template` 后缀） |
| 填写模板 | 自动识别字段并生成表单 |
|  | 同一字段多处出现时，各处独立渲染样式控件，文本输入联动同步 |
|  | 每处独立选字体 / 字号 / 颜色（系统字体全量，中文优先排序） |
|  | 中文字号名（初号 ~ 八号）+ 数值磅值（5 ~ 72） |
|  | 图片字段支持 PNG / JPEG / GIF |
| 自动检测 | 插入字段时自动读取光标处的字体 / 字号 / 颜色当默认 |
| 字体枚举 | Rust 端解析 `name` 表，免权限手势、含中文家族名、覆盖用户字体目录 |
| 兼容性 | 输出符合 OOXML 规范，Word / WPS / LibreOffice 都能打开 |
| 体积 | 打包后约 10MB（Tauri，复用系统 WebView2） |

---

## 快速开始

### 环境要求

- **Node.js** 18+
- **Rust** 1.77+
- **Windows**：MSVC build tools + WebView2 Runtime（Win10/11 默认已装）
- **macOS**：Xcode Command Line Tools
- **Linux**：`webkit2gtk` 等 Tauri 依赖（参考 [Tauri 文档](https://tauri.app/start/prerequisites/)）

### 安装与开发

```bash
git clone <repo>
cd doc_template
npm install
npm run tauri:dev
```

首次启动需要编译 Rust 端，约 1.5 ~ 3 分钟。之后热重载只动前端，~1 秒。

### 打包发行版

```bash
npm run tauri:build
```

产物在 `src-tauri/target/release/bundle/`：
- Windows：`.msi` 安装包 + `.exe`
- macOS：`.dmg` + `.app`
- Linux：`.deb` / `.AppImage`

### 仅前端调试（不带 Tauri）

```bash
npm run dev
```

注意：`pickDocx` / `readBytesFromPath` / `saveBytesViaDialog` 走 Tauri 命令，浏览器里点不动。但 docx 引擎本身（`src/docx.js`）是平台无关的，可单独跑：

```bash
node scripts/smoke-test.js
```

---

## 使用说明

### 制作模板

1. 顶部点 `📝 制作模板`
2. `选择文档 (.docx)` 导入任意 Word 文档
3. 应用左半边渲染原格式预览，右半边把每段拆成可编辑的卡片
   - 预览按 Word 页面视图渲染：保留页面尺寸（A4/Letter 等）与页边距
   - 首次加载默认「适应宽度」，避免一进来就左右滚动
   - 底部有「预览缩放」控制（`-` / `适应宽度` / `+`），也支持 `Ctrl + 滚轮`
4. **插入字段**有两种方式：
   - **预览里**：点光标定位，或者选中要替换的一段文字，点工具栏 `+ 文字字段` / `+ 图片字段`，弹窗里填名称 / 描述 → 占位符插到那个位置（选中的文字会被覆盖）
   - **卡片里**：直接在编辑区改文本，手写 `{@name}` 也行，应用会自动登记字段元数据
5. **占位符操作**：
   - 占位符渲染为高亮块（蓝色=文字，粉色=图片），不可被拆散编辑
   - **Backspace / Delete**：光标紧贴占位符时按删除键，整块删除
   - **右键**：右键占位符弹出"删除占位符"菜单
   - **双击**：双击占位符弹出编辑框，可修改名称（全局同步）和该位置的字体 / 字号 / 颜色
   - **悬停 tooltip**：鼠标悬停 0.5s 显示占位符详情（名称、类型、当前样式）
6. 卡片头部的"字段汇总"列出所有字段。点字段名可以再编辑描述
7. 右侧段落卡片的编辑区会自动随内容增高，聚焦时左侧预览高亮对应段落，失焦自动取消
8. 改完点 `刷新预览` 看效果，最后 `保存为模板`

#### 自动读取原格式

打开"添加字段"弹窗时，应用会去看光标位置原来的 `<w:rPr>`，把字体 / 字号 / 颜色读出来当默认值。比如标题段是"方正小标宋简体 / 三号"，正文段是"仿宋 / 小四"，分别在两段插入字段时，弹窗里默认就是对应那一套。

源码：`src/docx.js` 的 `getRunStyleAt(paragraphXml, charOffset)` —— 累加 `<w:t>` 长度找到对应 run，再从该 run 的 `<w:rPr>` 抽 font / size / color。字体优先 `eastAsia`，回退 `ascii` / `hAnsi`；字号是半磅（OOXML 规范），除以 2 转 pt。

### 修改模板

「制作模板」是一次性的 —— 占位符插完就走。但模板上线后总会发现要改：字段名取得不准、描述要补、默认字号要从小四改成五号、有的字段根本用不到。这时切到 `🔧 修改模板`：

1. 顶部点 `🔧 修改模板`
2. `选择模板 (.docx)` 导入做好的模板（即带 `template/fields.json` 的那种）
3. UI 跟「制作模板」完全相同 —— 同样的双栏布局、同样能加新占位符
4. 字段汇总区每个 tag 是 **`name ×N` + `×` 删除按钮**：
   - **`×N`** 是这个字段在文档里出现的次数；如果是 `0` 就说明只剩元数据没在文档里用
   - **点 tag 名字** → 弹出编辑对话框（描述 / 默认格式 / 名字都能改）
   - **点 `×` 按钮** → 弹 `confirm`：删了占位符也删元数据，不可撤销
5. **重命名字段**：点 tag 名字打开对话框，把"名称"改成新名 → 确定。应用会一次性 regex 替换文档里所有 `{@old}` → `{@new}` 和 `{%old}` → `{%new}`，并把 `fieldMeta` 的键迁过去
   - 如果新名跟另一个**已有**字段撞了，对话框会拦下来提示，不会让两个字段合并
6. 改完点 `保存修改`，默认覆盖原文件（系统对话框还是会让你确认）

#### 修改 vs 制作

两个 tab 共用同一套 DOM 和 state —— 切来切去状态不丢。区别只在三处：

| | 制作模板 | 修改模板 |
|---|---|---|
| 加载按钮 | 选择文档 (.docx) | 选择模板 (.docx) |
| 保存按钮 | 保存为模板 | 保存修改 |
| 默认文件名 | `xxx-template.docx` | 同输入文件名（覆盖） |

字段重命名 / 删除两个功能在两个 tab 里都能用 —— 只是「修改模板」更明确地告诉用户："你正在改一个已有模板"。

### 填写模板

1. 顶部点 `✏️ 填写模板`
2. `选择模板 (.docx)` 导入做好的模板
3. 应用扫描 `{@…}` / `{%…}`，按字段顺序生成表单
4. **文字字段**：textarea + 字体下拉 + 字号下拉 + 颜色选择器
   - 字体下拉支持搜索（系统字体全量，中文优先排序，含用户字体目录）
   - 字号有"中文字号"和"磅值"两组，覆盖 Word 字号下拉的所有选项
   - 字段卡片顶部显示字段描述（如果模板里设过）
   - **同一字段多处出现**时，每处渲染为独立卡片（标注"样式 1/2"等），各自有独立的字体/字号/颜色控件，但文本输入框联动——修改一个，其他自动同步
5. **图片字段**：选 PNG/JPEG/GIF 图片
6. `生成并保存 DOCX` 选输出路径

字段卡片的初始字体 / 字号 / 颜色，来自模板里存的 per-occurrence 样式（`template/fields.json` 的 `occStyles`）。第一次填的时候不用每个字段重新选。

---

## 模板语法

| 占位符 | 类型 | 用法 | 备注 |
|---|---|---|---|
| `{@name}` | 文字 | `姓名：{@name}` | 行内可用，可跨多个 run |
| `{%avatar}` | 图片 | `{%avatar}` | 建议单独占一行；默认 240×240 像素 |

字段名必须是 `\w+`（字母 / 数字 / 下划线）。中文 / 空格 / 符号都不行。

**字段元数据**（描述）保存在 `.docx` 包内的 `template/fields.json`。**per-occurrence 样式**（每个占位符位置的字体/字号/颜色）也存在同一个 JSON 里：

```json
{
  "version": 3,
  "fields": [
    { "name": "name", "type": "text", "description": "姓名" },
    { "name": "avatar", "type": "image", "description": "证件照" }
  ],
  "occStyles": [
    { "pIdx": 0, "occ": 0, "name": "name", "sigil": "@", "font": "方正小标宋简体", "size": 16, "sizeLabel": "三号", "color": "#000000" },
    { "pIdx": 5, "occ": 0, "name": "name", "sigil": "@", "font": "仿宋", "size": 12, "sizeLabel": "小四", "color": "#000000" }
  ]
}
```

同一字段名（如 `name`）在不同位置可以有不同的字体/字号/颜色。填写时文本内容联动同步，但各位置保持各自样式。

兼容 v2：早期模板的 `defaultFont`/`defaultSize`/`defaultColor` 字段已移至 `occStyles`，旧模板加载时会回退到默认值。

---

## 项目结构

```
.
├── index.html                  # 应用入口（含字段元数据 modal）
├── src/
│   ├── main.js                 # UI 状态机、Tauri 命令调用、表单渲染
│   ├── docx.js                 # DOCX 引擎（解析、改写、字段元数据）
│   ├── styles.css
│   └── assets/
├── src-tauri/
│   ├── src/lib.rs              # read_file_bytes / save_bytes / list_fonts 命令
│   ├── tauri.conf.json
│   ├── Cargo.toml
│   ├── icons/
│   └── capabilities/           # Tauri 2 权限声明
├── scripts/
│   ├── make-sample-template.js # 生成 sample-template.docx
│   ├── smoke-test.js           # 端到端测试（35 个 check）
│   └── inspect.js              # 检查任意 docx 内部结构、跑 XML 校验
├── package.json
└── vite.config.js
```

---

## 架构与技术栈

```
┌──────────────────────────────────────────────────────────┐
│                  Tauri WebView (Chromium)                 │
│  ┌────────────────────────────────────────────────────┐   │
│  │  index.html / src/main.js (UI 层)                  │   │
│  │   ├ 三个 Tab: 制作模板 / 修改模板 / 填写模板        │   │
│  │   ├ docx-preview        ── 原格式预览              │   │
│  │   ├ FontPicker / Modal  ── 自定义控件              │   │
│  │   └ src/docx.js         ── DOCX 引擎               │   │
│  │       ├ PizZip          ── docx = zip + xml        │   │
│  │       ├ Docxtemplater   ── {field} 替换            │   │
│  │       └ Image Module    ── 图片注入                │   │
│  └────────────────────────────────────────────────────┘   │
│                          ↕ invoke                         │
│  ┌────────────────────────────────────────────────────┐   │
│  │  src-tauri/src/lib.rs (Rust 层)                    │   │
│  │  ├ read_file_bytes(path) -> Vec<u8>               │   │
│  │  ├ save_bytes(path, bytes) -> ()                  │   │
│  │  └ list_fonts() -> Vec<String>                    │   │
│  └────────────────────────────────────────────────────┘   │
└──────────────────────────────────────────────────────────┘
```

| 层 | 依赖 | 职责 |
|---|---|---|
| UI | docx-preview ^0.3.7 | 渲染原格式预览（分页、页面尺寸、页边距） |
| 引擎 | pizzip ^3.1.6 | docx 是 zip 包，PizZip 用来读 / 写 |
|  | docxtemplater ^3.50.0 | `{field}` 标准替换 |
|  | docxtemplater-image-module-free ^1.1.1 | 把字节流注入成 `<w:drawing>` |
| 桌面壳 | @tauri-apps/api ^2.1.1 | 与 Rust 后端通信 |
|  | @tauri-apps/plugin-dialog ^2.0.1 | 系统文件 / 保存对话框 |
| 后端 | tauri 2.x | 应用框架，文件 IO |
|  | ttf-parser ^0.21 | 解析系统字体的 `name` 表抽家族名 |
|  | walkdir ^2 | 递归扫字体目录 |

---

## DOCX 引擎工作原理

`.docx` 本质是个 zip 包，里面是若干 XML：

```
mydoc.docx (zip)
├── [Content_Types].xml         # 每个 part 的 MIME 类型声明
├── _rels/.rels                 # 包级关系
└── word/
    ├── document.xml            # 主体内容（段落 / 表格 / 图片占位）
    ├── _rels/document.xml.rels # document.xml 的关系（链接、图片）
    ├── styles.xml
    ├── settings.xml
    ├── theme/theme1.xml
    └── media/image1.png        # 内嵌图片二进制
```

下面是这个项目里需要 hack 的核心点。

### 段落解析

`parseParagraphs(zip)` 把 `word/document.xml` 里每个 `<w:p>` 抽成一个对象：

```js
{
  index, originalXml, originalText, currentText,
  dirty, pStart, pEnd, hasComplex, selfClosing
}
```

- `originalText` 是把段内所有 `<w:t>` 拼起来；这是给用户编辑的那一行文本
- `pStart` / `pEnd` 是段落 XML 在 `document.xml` 里的字节偏移，用于 in-place 替换
- 既匹配 `<w:p>...</w:p>`，又匹配自闭合的 `<w:p/>`（空段落），保证段数与 `docx-preview` 渲染出的 `<p>` 数量一致
- `hasComplex` 检测有 `<w:drawing>` / `<w:hyperlink>` / `<w:fldChar>` / `<w:instrText>` 的段落，UI 会标 `⚠ 复杂段落`，提示编辑后可能丢失非文本内容

### 段落写回（编辑模式）

`buildParagraphXml(paragraph)` 重写改过的段，**保策略是放弃 run 级别**：

1. 保留原段的 `<w:pPr>`（段落属性：对齐、缩进、行距）
2. 取段内**第一个** `<w:rPr>`（run 属性：字体、字号、颜色）当默认样式
3. 把整个段落塞进**单个** `<w:r>`，新文本里的 `\n` 用 `<w:br/>` 处理

代价：原段如果是"红色 + 蓝色"两段不同颜色的文字，编辑后会变成全部用第一段的颜色。

收益：实现简单、可靠，不会因为部分修改打乱 run 边界。

`applyParagraphEdits` 按 `pStart` 倒序替换，避免改前段位移影响后段偏移。

### 占位符样式注入（填写模式）

如果直接把 `{@name}` 给 docxtemplater，它只会做文本替换，新文本继承所在 run 的样式 —— 也就是模板里写 `{@name}` 那个位置原本的样式，用户在表单里选的字体 / 字号 / 颜色没法生效。

所以填写前先做一遍预处理（`preprocessTemplateForFill`）。该函数按出现顺序跟踪每个字段名的 occurrence index，从 `styleMap[name]` 数组中取对应的 per-occurrence 样式：

```
原: <w:r><w:rPr>原样式</w:rPr><w:t>姓名：{@name}，欢迎</w:t></w:r>

后: <w:r><w:rPr>原样式</w:rPr><w:t xml:space="preserve">姓名：</w:t></w:r>
    <w:r><w:rPr>第1处name的样式</w:rPr><w:t xml:space="preserve">{name}</w:t></w:r>
    <w:r><w:rPr>原样式</w:rPr><w:t xml:space="preserve">，欢迎</w:t></w:r>
```

`styleMap[name]` 可以是单个样式对象（向后兼容）或样式数组（per-occurrence）。切完 run、把 `{@name}` 替换成 `{name}`，再交给 docxtemplater 跑标准替换。

`buildRPr({font, size, color})` 拼出符合 OOXML 的 `<w:rPr>`：

```xml
<w:rPr>
  <w:rFonts w:ascii="..." w:hAnsi="..." w:cs="..." w:eastAsia="..."/>
  <w:sz w:val="halfPt"/>
  <w:szCs w:val="halfPt"/>
  <w:color w:val="HEX"/>
</w:rPr>
```

四个 `w:rFonts` 属性都设同一字体，否则中英混排时 Word 可能切回 `Calibri`。`w:sz` 是半磅整数（12pt → 24）。

### 字段元数据持久化

应用要在 `.docx` 里塞两类东西：

1. 字段描述（给操作员看的提示）
2. per-occurrence 样式（每个占位符位置的字体 / 字号 / 颜色）

存哪里？方案对比：

| 方案 | 优点 | 缺点 |
|---|---|---|
| `template/fields.json`（**当前**） | 简单、独立、易调试 | 需在 `[Content_Types].xml` 声明 `.json` 扩展名，否则 Word 拒打开 |
| `docProps/custom.xml` | OOXML 标准，Word 保留 | 格式有 schema 限制，写起来繁琐 |
| 编码到占位符里（如 `{@name,仿宋,小四,#000}`） | 不需额外文件 | 描述里有逗号会爆；docxtemplater 会把逗号当参数 |

选了 JSON 方案，关键是 `buildTemplate` 写入 fields.json 时同步加 content type 声明。per-occurrence 样式以 `occStyles` 数组存储，每条记录包含段落索引（`pIdx`）、出现序号（`occ`）、字段名（`name`）和样式属性。

```js
function ensureJsonContentType(zip) {
  const ct = zip.file("[Content_Types].xml").asText();
  if (/Extension="json"/i.test(ct)) return;
  zip.file(
    "[Content_Types].xml",
    ct.replace(
      /<\/Types>/,
      '<Default Extension="json" ContentType="application/json"/></Types>'
    )
  );
}
```

`renderFilled` 输出前会**删掉** `template/fields.json` —— 填好的文档不需要编辑器元数据，留着只是徒增 part。

### 自动读取光标处格式

插入字段弹窗的"默认格式"会从光标位置自动填好。`getRunStyleAt(paragraphXml, charOffset)`：

1. 用 regex 匹配段内所有 `<w:r>...</w:r>`
2. 累加每个 run 内 `<w:t>` 的总字符数
3. 找到包含 `charOffset` 的那个 run，解析 `<w:rPr>` 里的字体 / 字号 / 颜色

字体名优先级：`w:eastAsia` > `w:ascii` > `w:hAnsi`。中文文档里 `w:eastAsia` 才是中文字体名。

### 字段重命名 / 删除

修改模式的 rename 和 delete 都是在 **段落文本层** 做的（`state.paragraphs[].currentText`），不直接动 XML —— 等到「保存修改」时才走 `buildTemplate` 把改过的段重写回 `document.xml`。这样既复用了已有的段落写回管线，也避免了对 run 边界的精确 diff。

实现都在 `src/main.js`：

```js
// renameField('partyA', 'partyB')
//   {@partyA}  → {@partyB}
//   {%partyA}  → {%partyB}   ← image 占位符也跟着变
//   fieldMeta.delete('partyA') + fieldMeta.set('partyB', oldMeta)
const re = new RegExp(`\\{([@%])${escapeRegex(oldName)}\\}`, "g");
p.currentText = p.currentText.replace(re, `{$1${newName}}`);
```

```js
// deleteField('partyA')
//   {@partyA} | {%partyA} → ''   ← 占位符直接抹掉
//   fieldMeta.delete('partyA')
const re = new RegExp(`\\{[@%]${escapeRegex(name)}\\}`, "g");
p.currentText = p.currentText.replace(re, "");
```

碰撞校验放在弹窗的 `onSubmit` 里：如果用户把 `partyA` 改成 `partyB` 但 `partyB` 已存在为另一个独立字段，对话框拦下来不让走，避免两个字段被合并。

---

## 测试与调试

### Smoke test

35 个端到端 check，覆盖填写流程、编辑回写、namespace 完整性、ContentTypes 合规：

```bash
node scripts/smoke-test.js
```

运行后产物：
- `smoke-output.docx` —— 填写流程的输出
- `smoke-edited-template.docx` —— 编辑后保存的模板
- `smoke-edited-filled.docx` —— 用上面那个模板填出来的最终 docx

### Inspect 工具

要看任意 docx 包里有什么、ContentTypes 长啥样、每个 XML part 是否良构：

```bash
node scripts/inspect.js path/to/file.docx
```

输出：

```
========== file.docx ==========
Files:
  [Content_Types].xml
  word/document.xml
  ...

[Content_Types].xml:
  <?xml ...?>
  ...

word/_rels/document.xml.rels:
  ...

word/document.xml:
  ...

(no XML errors logged above => all parts well-formed)
```

如果某个 part 解析失败会列出错误。

### 生成测试模板

```bash
node scripts/make-sample-template.js   # 写出 sample-template.docx
```

里面带 `{@name}` `{@title}` `{@joinDate}` `{@bio}` 文字字段和 `{%avatar}` 图片字段。

---

## 已知陷阱与设计决策

### 1. `[Content_Types].xml` 必须声明每个 part 的扩展

OOXML 规范要求每个 part 都有 content type，要么走 `<Default Extension="...">`，要么走 `<Override PartName="...">`。

漏掉的下场：Word 提示"文件已损坏，是否修复？"，WPS 直接拒开。

我们写 `template/fields.json` 时同步加 `<Default Extension="json" ContentType="application/json"/>`。

### 2. 必须声明完整的 namespace

如果 docx 里有图片（`<w:drawing>` / `<wp:inline>` / `r:embed`）但 `<w:document>` 上没声明 `wp:` `r:` 前缀，Word 会报错。

`scripts/make-sample-template.js` 里完整列出了所有 Word 期望的 namespace：

```js
const NS = ' xmlns:wpc="..." xmlns:cx="..." xmlns:mc="..."'
         + ' xmlns:o="..." xmlns:r="..." xmlns:m="..."'
         + ' xmlns:v="..." xmlns:wp14="..." xmlns:wp="..."'
         + ' xmlns:w10="..." xmlns:w="..." xmlns:w14="..."'
         + ' xmlns:w15="..." xmlns:w16se="..." xmlns:wpg="..."'
         + ' xmlns:wpi="..." xmlns:wne="..." xmlns:wps="..."'
         + ' mc:Ignorable="w14 w15 w16se wp14"';
```

### 3. 必须有 `<w:sectPr>`、`word/styles.xml`、`word/settings.xml`

少任何一个，Word 都可能拒开或显示空白。`make-sample-template.js` 里都补齐了。

### 4. 字号是半磅，不是磅

`<w:sz w:val="24"/>` 表示 12pt，不是 24pt。`buildRPr` 里 `Math.round(size * 2)` 转换。

### 5. 中文字体名要写 `w:eastAsia`

`<w:rFonts w:ascii="宋体" .../>` 单独这样写，中英文混排时 Word 可能用 `Calibri` 显示中文。四个属性都设同一字体最稳：

```xml
<w:rFonts w:ascii="..." w:hAnsi="..." w:cs="..." w:eastAsia="..."/>
```

### 6. 字体枚举走 Rust 端

`window.queryLocalFonts()`（Chromium Local Font Access API）有两个问题：
- 要用户激活手势才能调，没点击就调会被拒
- 在 Tauri WebView 里偶尔只返回 PostScript / 英文名（"FangZheng XiaoBiaoSong"）—— Word 显示的是中文家族名（"方正小标宋简体"），用户在我们的列表里就找不到
- 用户安装到 `%LOCALAPPDATA%\Microsoft\Windows\Fonts` 的字体可能漏掉

所以字体枚举挪到了 Rust 端（`src-tauri/src/lib.rs::list_fonts`）：
- 用 `walkdir` 扫所有系统字体目录（Windows / macOS / Linux 都覆盖，含 per-user 目录）
- 用 `ttf-parser` 直接解析每个 `.ttf/.otf/.ttc/.otc` 的 `name` 表
- 抽 nameID=1（Family）和 nameID=16（Preferred Family），所有语言的版本都收 —— 这样 Word 看得到的中文名我们也看得到
- TTC 集合包按 face index 逐个解
- 返回去重排序的家族名列表，前端通过 `invoke('list_fonts')` 拿

前端仍保留 `queryLocalFonts` 和内置 50 个常用字体作为两级回退，以防 Rust 端调用失败。

### 7. 字体下拉框

Chromium 的 `<datalist>` 弹出框只显示约 4 项，长字体列表不能滚动。所以我们自己写了 `createFontPicker`：可搜索的 `<input>` + 自定义 `<ul>`，支持键盘上下选 + Enter 选中，打开时滚到当前值。

### 8. PizZip 重新生成时文件顺序会变

原 docx 里 `[Content_Types].xml` 通常是 zip 第一项，PizZip 重写后顺序可能变。**OOXML 不要求 `[Content_Types].xml` 必须第一项**，Word 也不挑顺序。

### 9. 编辑模式段落处理是 lossy 的

`buildParagraphXml` 把整段塞进单个 `<w:r>` + 第一个 `<w:rPr>`。原段如果有部分加粗 / 部分变色，编辑后会丢。复杂段落 UI 标了 ⚠ 警告，但不阻止用户操作。

更严谨的做法：保留原 run 边界、智能 diff 文本。但实现复杂度上升一个量级，目前这个简化版够用。

### 10. 预览缩放为什么用 CSS `zoom`

最初用 `transform: scale(...)` 时，视觉缩小了，但布局宽度不变，会留下“看不见但可横向滚动”的空间，页面也容易出现偏右。

现在改为对预览 wrapper 使用 CSS `zoom`：缩放会直接影响布局尺寸，适应宽度时不会残留横向滚动，页面居中更稳定。

---

## 后续可拓展

- **PDF 导出**：调用本机 LibreOffice `soffice --convert-to pdf`，或集成 `docx-to-pdf` 这类纯 Rust 库
- **Run 级编辑**：保留原段落里的 run 边界，做精确 diff，避免丢部分加粗 / 变色
- **表格内字段 UI**：目前已能识别和填表格里的 `{@field}`，但段落卡片把所有段都铺开，表格看不出来
- **字段值持久化**：把表单值存为 JSON，方便重复填同一模板（比如月度报告）
- **段落新增 / 删除**：现在只能改文本，加段要在 Word 里编辑后再导入
- **字段重排序**：当前字段顺序是按 `{@…}` 在文档里的出现顺序生成；想要自定义表单顺序需要拖拽 UI
- **批量填写**：上传 Excel / CSV，按行批量生成 docx
- **校验规则**：字段加 regex / 必填 / 最长字符等约束
- **撤销/重做**：修改模板时误删字段后没法 undo，得重新打开文件

## 已实现 / 完成

- ✅ 三个模式：制作 / 修改 / 填写
- ✅ 占位符高亮块渲染（contenteditable，蓝色=文字，粉色=图片）
- ✅ Backspace / Delete 整块删除占位符，右键菜单删除
- ✅ 双击占位符编辑（名称 + per-occurrence 样式）
- ✅ 鼠标悬停占位符 0.5s 显示详情 tooltip
- ✅ per-occurrence 样式：同一字段名在不同位置可有不同字体 / 字号 / 颜色
- ✅ 填写模式多处出现的字段：文本联动同步，样式独立
- ✅ 字段重命名（同步替换文档占位符）
- ✅ 字段删除（占位符 + 元数据一起抹掉）
- ✅ 字段汇总区显示出现次数
- ✅ Rust 端字体枚举（含中文家族名、用户字体目录）
- ✅ 自动读取光标处的字体 / 字号 / 颜色作为字段默认格式
- ✅ OOXML content type 合规（fields.json 不再让 Word 拒打开）
- ✅ 原格式预览支持分页和 Word 页面比例（页面尺寸 / 页边距）
- ✅ 预览缩放（适应宽度、底部缩放按钮、`Ctrl + 滚轮`）
- ✅ 编辑区聚焦高亮预览对应段落，失焦自动取消
- ✅ 字体选择弹窗不再自动展开（改为点击触发）

---

## License

MIT
