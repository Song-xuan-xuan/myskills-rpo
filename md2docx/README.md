> ⚠️ **精简版，测试中**。命令已合并为单入口（`md2docx.py <子命令>`），
> 自检 137 项通过，但尚未在真实报告上完整验收。遇到问题请先看 `docs/TESTING.md`。

# md2docx —— Markdown 写报告，一键出「排版可控 + 编号活」的 Word

```
report.md ──pandoc+模板──> report.docx ──docxfix──> 可交付.docx
```

三条设计原则：

1. **md 只负责内容**，不模拟 Word 格式；格式全部落在「档案（profile）」里
2. **编号用 Word 原生域**（SEQ / REF / TOC），不是写死的文字，改了会自动重排
3. **每一处格式都能改**，改档案文件即可，不用打开 Word 点样式

---

## 零、它在哪

```
<工作目录>/.agents/skills/md2docx/   ← skill 本体（pi 会自动加载）
（如果同时放了一份软链接，老路径也能用）
```

`SKILL.md` 是给 agent 看的操作说明，`README.md`（本文件）是给人看的完整手册，
`docs/TESTING.md` 是三步验收清单。

## 一、30 秒上手

产物**只有一个 docx**，不会在你目录里留侧车文件、缓存或临时目录：

```
报告.md  →  报告.docx        （就这些）
```

```bash
python3 md2docx.py make 报告.md                 # 转换 + 质检 + 格式摘要（推荐）
python3 md2docx.py make 报告.md --pdf           # 顺带导 PDF（需要本机 Word）
python3 md2docx.py make 报告.md --profiles technical,official   # 一份 md 出多版
```

只要转换、不要质检：

```bash
python3 md2docx.py 报告.md 交付版.docx --profile technical
```

### 先跑自检

```bash
python3 selftest.py
```

它会检查依赖、把每套 profile 的模板编译一遍、转四份样例，然后验证
分节数 / 页码格式 / 目录域 / SEQ 域 / REF 域 / 页脚页码域 / 斜体残留 /
图片是否超出版心 / 表格是否重复表头 / 标题颜色 / 编号分隔，
最后再验证「格式清单能生成」「`--set` 能改单个键」「`--reference` 能用你自己的模板」，
最后打印 PASS 清单（当前 57 项全绿）。

---

## 二、格式到底在哪配置？（重点）

**没有任何格式写死在代码里。** 全部在 `profiles/*.json` 里，分七块：

```
profile.json
├── page                 页面：纸张、四边距、页眉页脚距离、首页是否不同
├── doc_default          文档默认字体字号（兜底）
├── header / footer      页眉文字、页脚页码格式
├── cover                封面：留白、额外行、是否编页码
├── front_matter         目录/图目录/表目录：标题、页码格式（罗马/阿拉伯）
├── body_paging          正文页码格式与起始号
├── captions + refs      图/表/公式的编号外形 & 正文引用外形
├── toc                  目录层级、是否自动插入、是否另起一页
├── heading_numbering    标题自动编号（1 / 1.1 / 一、 / 第1章…）
├── table                表格线框、表头、单元格对齐
└── styles               每一个 Word 样式：字体/字号/缩进/行距/间距
```

### 2.0 三种改格式的方式

| 方式 | 怎么做 | 什么时候用 |
|---|---|---|
| **A. 改档案 JSON** | `configure.py --set 路径=值` 或直接编辑 JSON | 绝大多数情况 |
| **B. 用自己的 Word 模板** | 在 Word 里改样式另存为 docx，转换时 `--reference 你的模板.docx` | 已有单位模板，想用 Word 界面调 |
| **C. 问 agent** | 直接说「题注加粗」「正文改成仿宋三号」 | 不确定改哪个键时 |

改完随时看当前格式：

```bash
python3 configure.py --report <档案名或路径>
```

输出是一份中文清单（页面/正文/各级标题/题注/表格/代码块/页眉页脚/封面/目录），
**读的是实际编译出来的模板**，所以是真正生效的效果，不是 JSON 里写了什么就显示什么。

> ⚠️ 不要改 `build/*.docx`。模板每次转换都重新编译（否则不同项目同名 profile 会串格式），
> 手改会被冲掉。要改就用 A 或 B。

### 2.1 正文的「顶格 / 缩进 / 段前后间距」怎么设

```json
"styles": {
  "Normal": {
    "size_pt": "小四",              // 字号（可写中文号数）
    "align": "justify",             // left | center | right | justify
    "line_spacing": 1.5,            // 倍数行距；固定行距改用 line_spacing_pt
    "space_before_pt": 0,           // 段前（磅）
    "space_after_pt": 0             // 段后（磅）
  },
  "FirstParagraph": { "first_line_indent_chars": 2 },   // 标题后第一段
  "BodyText":       { "first_line_indent_chars": 2 }    // 其余正文段
}
```

- **缩进单位是「字符」**，不是厘米——`first_line_indent_chars: 2` 就是首行缩进 2 字符，换字号不会错位
- 顶格就是 `0`
- 悬挂缩进：`hanging_indent_chars`
- 左右缩进：`left_indent_chars` / `right_indent_chars`

> ⚠️ **最关键的一条**：首行缩进只加在 `BodyText` + `FirstParagraph`，
> **绝不能加在 `Normal`**。加在 Normal 会连标题、表格、图注一起缩进——
> 这正是大部分人从 md 转 Word 踩的坑。

### 2.2 每一块内容对应哪个 Word 样式

| md 里的东西 | Word 样式名 | 可调项 |
|---|---|---|
| `#` … `######` | `Heading1` … `Heading6` | 字体/字号/对齐/段前后 |
| 标题后第一段 | `FirstParagraph` | 缩进 |
| 其余正文段 | `BodyText` | 缩进/行距 |
| 表格单元格、紧凑列表 | `Compact` | 字号/对齐 |
| 图注 | `ImageCaption` | 字号/对齐/间距 |
| 表注 | `TableCaption` | 同上 |
| 图片所在段落 | `CaptionedFigure` | 居中、与下段同页 |
| `> 引用` | `BlockText` | 缩进/字体 |
| ```代码块``` | `SourceCode` | 字号/底色/边框 |
| 脚注 | `FootnoteText` | 字号 |
| 文档标题 | `Title` / `Subtitle` / `Author` / `Date` | 全部 |

### 2.3 标题编号：想用 `1.` / `1.1.` / `一、` 都行

```json
"heading_numbering": {
  "enabled": true,
  "levels": [
    { "format": "%1、",   "num_fmt": "chineseCounting", "indent_chars": 2 },  // 一、
    { "format": "（%2）", "num_fmt": "chineseCounting", "indent_chars": 2 },  // （一）
    { "format": "%3.",    "num_fmt": "decimal",         "indent_chars": 2 }   // 1.
  ]
}
```

`format` 里 `%1 %2 %3` 是各级序号，其余字符原样输出。常见写法：

| 想要的效果 | `format` |
|---|---|
| 1 / 1.1 / 1.1.1 | `"%1"` / `"%1.%2"` / `"%1.%2.%3"` |
| 1. / 1.1. / 1.1.1. | `"%1."` / `"%1.%2."` / `"%1.%2.%3."` |
| 一、/（一）/ 1. | `"%1、"` / `"（%2）"` / `"%3."` |
| 第1章 / 1.1 | `"第%1章"` / `"%1.%2"` |
| 不要编号 | `"enabled": false` |

`num_fmt` 支持：`decimal`、`chineseCounting`、`upperRoman`、`lowerRoman`、`upperLetter`、`lowerLetter`。

### 2.4 图 / 表 / 公式的编号与引用

```json
"captions": {
  "fig": "图 {n}",        // 题注长什么样
  "tbl": "表 {n}",
  "eq":  "({n})",         // 公式编号
  "chapter": false        // true = 按章编号（图 2-1）
},
"refs": {
  "fig": "{n}",           // 正文里 {{fig:x}} 渲染成什么
  "tbl": "{n}",
  "eq":  "({n})"
}
```

`{n}` 是自动编号占位符，前后可以随便加字：

| 想显示 | `captions.fig` / `refs.fig` |
|---|---|
| 图 1 | `"图 {n}"` |
| 图1 | `"图{n}"` |
| 图 2-1（按章） | `"图 {n}"` + `"chapter": true` |
| Figure 1 | `"Figure {n}"` |
| 只出数字 1 | `"{n}"` |

**正文写法**（关键约定）：

```markdown
![系统架构图](img/arch.png){#fig:arch}      <!-- 自动变成「图 1 系统架构图」 -->
见图 {{fig:arch}} 所示。                     <!-- refs.fig="{n}" 时 → 「见图 1 所示」 -->
```

- `refs` 用 `"{n}"` 时，正文要自己写「见图」二字（和 LaTeX `\ref` 一样）
- `refs` 改成 `"图 {n}"` 时，正文写「见 {{fig:arch}}」→「见 图 1」
- 题注前缀和正文引用**是两套独立配置**，可以各玩各的

### 2.4b 交叉引用：图 / 表 / 公式 / 章节

正文里用 `{{类型:锚点}}`（**双大括号**），四种类型：

| 写法 | 指向 | 默认渲染 |
|---|---|---|
| `{{fig:arch}}` | 图题注的编号 | `1` |
| `{{tbl:stat}}` | 表题注的编号 | `1` |
| `{{eq:cap}}` | 公式编号 | `(1)` |
| `{{sec:detail}}` | 章节（标题） | `1.1`（章号） |

> **为什么用双大括号而不是 `@fig:arch`**：pandoc 的 `--citeproc` 会把任何
> `@key` 当成文献引用吃掉，`@fig:arch` 会变成 `[fig:arch?]`。
> `{{fig:arch}}` 不会被碰，参考文献和交叉引用才能共存。
> （旧的 `@fig:arch` 写法仍兼容，但没有参考文献的文档里也能用。）

章节引用还可以用 `{t}` 取标题文字：

```json
"refs": { "sec": "见「{t}」（第 {n} 节）" }
```

| 占位 | 含义 |
|---|---|
| `{n}` | 编号（章节=章号，图表=自动编号） |
| `{t}` | 标题文字（只有章节有意义） |

锚点怎么写：

```markdown
## 关键细节 {#sec:detail}      ← 章节锚点（写在标题后面）
![说明](a.png){#fig:arch}      ← 图锚点
: 表说明 {#tbl:stat}            ← 表锚点
::: {#eq:cap}                  ← 公式锚点（必须用围栏）
```

### 2.4c 参考文献

md 头部写 `bibliography:`，其余全自动：

```markdown
---
title: 我的报告
bibliography: refs.bib        ← BibTeX 文件，和 md 同目录
---

正文写 @vaswani2017 就会变成 [1]。

# 参考文献

::: {#refs}
:::
```

内置三套国标 CSL（`csl/` 目录，已离线）：

| 配置 | 效果 |
|---|---|
| `chinese-gb7714-2005-numeric` | GB/T 7714-2005 数字式 `[1]`（默认） |
| `chinese-gb7714-2005-author-date` | GB/T 7714-2005 作者-年制 `(作者, 2020)` |
| `chinese-gb7714-1987-numeric` | GB/T 7714-1987 数字式 |

```json
"citations": { "csl": "chinese-gb7714-2005-numeric", "number_bibliography": true }
```

**编号是活的**：`[1] [2] …` 不是写死的文字，而是 `SEQ 文献` 域；正文角标 `[1]`
是指向书目的 `REF` 域。在 Word 里删掉第 7 条文献、`Ctrl+A → F9`，
后面的编号**和正文里所有角标**都会自动重排。

（关掉就设 `number_bibliography: false`，回到 citeproc 写死的静态编号。
作者-年制风格没有 `[n]`，本功能自动跳过。）

参考文献条目的样式是 Word 的 `Bibliography` 样式（悬挂缩进），可在 `styles` 里调。

### 2.4d 宽表横排

列特别多的表可以单独放到横向页：

```json
"landscape": ["tbl:wide"]
```

md 里给那张表一个锚点：

```markdown
| ...30 列... |

: 宽表说明 {#tbl:wide}
```

转换后这张表连同它的题注会单独占一页（横向），其余页面不受影响。

### 2.12 摘要与关键词

md 头部写摘要，转换后自动**单独占一页**放在封面之后、目录之前：

```markdown
---
title: 我的报告
abstract: |
  这里是摘要正文，可以写多段。
  缩进、行距跟正文样式走。
keywords: [数据平台, 调度, 监控]
---
```

```json
"abstract": {
  "enabled": true,
  "title": "摘　要",              // 摘要标题
  "title_style": "TOCHeading",
  "keywords_title": "关键词",
  "style": "Abstract"             // 摘要正文用哪个样式
}
```

关键词会渲染成一行「关键词：甲；乙；丙」，样式同上。

### 2.13 脚注 / 尾注

```json
"notes": {
  "footnote": { "pos": "pageBottom", "num_fmt": "decimal",
                "num_start": 1, "num_restart": "continuous" },
  "endnote":  { "pos": "docEnd", "num_fmt": "lowerRoman" }
}
```

| 字段 | 可选值 |
|---|---|
| `pos` | `pageBottom`（页底）/ `beneathText`（正文下方）；尾注 `docEnd` / `sectEnd` |
| `num_fmt` | `decimal` `lowerRoman` `upperRoman` `chineseCounting` `symbol` |
| `num_restart` | `continuous`（连续）/ `eachPage`（每页）/ `eachSect`（每节） |

md 里照常写 `脚注[^1]` + `[^1]: 内容`。

### 2.14 列表编号

```json
"lists": {
  "ordered": { "num_fmt": "decimal", "format": "%1.", "indent_chars": 2 },
  "bullet":  { "char": "•" }
}
```

`format` 里 `%1` 是层级序号（`%1.` → `1.`，`%1)` → `1)`，`（%1）` → `（1）`），
配合 `num_fmt: chineseCounting` 就是「一、」。

### 2.15 表格列宽 / 行跨页

```json
"table": {
  "cant_split": true,                        // 整行不允许被分页拆开
  "col_widths": { "tbl:metric": [30, 35, 35] },   // 按表锚点指定列宽比例
  "per_table": {
    "tbl:wide": { "preset": "full_grid", "col_widths": [15, 20, 20, 45] }
  }
}
```

锚点就是题注里的 `{#tbl:xxx}`。指定了列宽的表会自动切成固定布局。

### 2.16 文档级开关

```json
"document": {
  "track_changes": false,     // 打开后接收者的修改会被记成修订
  "protect": null,            // readOnly / comments / trackedChanges / forms
  "default_tab_cm": 1.27,
  "zoom": 100
}
```

### 2.17 首字下沉 / 文档网格

```json
"drop_cap": { "enabled": true, "lines": 3, "chars": 2 },
"page": { "doc_grid": { "type": "lines", "lines_per_page": 22 } }
```

公文常用「每页 22 行」，`lines_per_page` 会按版心高度自动算出行距。

### 2.18 交叉引用带页码

`refs` 模板里可以用 `{p}` 输出页码：

```json
"refs": { "sec": "第 {n} 节（第 {p} 页）" }
```

| 占位 | 含义 |
|---|---|
| `{n}` | 编号 |
| `{t}` | 标题文字（仅章节） |
| `{p}` | 所在页码 |

### 2.19 附录编号（附录 A / A.1）

```json
"appendix": {
  "enabled": true,
  "label": "附录",           // 一级编号前缀，渲染成「附录 A」
  "marker": "sec:app",       // 附录起点标题的锚点
  "end_marker": "sec:endappendix"   // 可选：从这里回到正常章节编号
}
```

md 里**不要手写「附录 A」**，只写标题并加锚点：

```markdown
# 数据说明 {#sec:app}

附录正文。

## 明细表 {#sec:app_data}

# 补充实验

第二个附录——一级标题自动变成「附录 B」。

# 结论 {#sec:endappendix}

回到正常编号（第 2 章）。
```

效果：

| 标题 | 显示 |
|---|---|
| `# 数据说明 {#sec:app}` | **附录 A** 数据说明 |
| `## 明细表 {#sec:app_data}` | **A.1** 明细表 |
| `# 补充实验` | **附录 B** 补充实验 |
| 正文 `见表 {{sec:app_data}}` | 见表 A.1 |

附录里的图、表如果用 `captions.chapter`，章号也会自动取 A / B。

### 2.20 封面 logo

```json
"cover": {
  "logo": "assets/logo.png",          // 标题上方
  "logo_width_cm": 4.0,
  "logo_align": "center",
  "logo_space_after_pt": 18,
  "logo_bottom": "assets/seal.png",   // 页面底部（可选）
  "logo_bottom_width_cm": 3.0
}
```

路径相对 md 所在目录。

### 2.21 一份 md 出多版

```bash
python3 md2docx.py batch 报告.md                                   # 三套内置档案全出
python3 md2docx.py batch 报告.md --profiles technical,academic      # 指定档案
python3 md2docx.py batch 报告.md --outdir dist --prefix 2026Q3      # 指定目录和前缀
python3 md2docx.py batch 报告.md --profiles a.json,b.json --sample   # 顺带出样张
```

输出：

```
dist/报告_technical.docx
dist/报告_official.docx
dist/报告_academic.docx
```

「给领导一版、给甲方一版、投稿一版」不用再跑三次。

### 2.5 表格线框

```json
"table": {
  "preset": "three_line",       // three_line | full_grid | outline | none
  "borders": {                  // 任意一边都能覆盖
    "top":    { "sz": 12 },     // sz 单位 1/8 磅：12 = 1.5 磅
    "bottom": { "sz": 12 },
    "header_bottom": { "sz": 6, "color": "000000" }
  },
  "header": { "bold": true, "repeat": true, "valign": "center", "shading": "F2F2F2" },
  "cell":   { "valign": "center" }
}
```

- `preset` 给现成的四套；`borders` 里写哪边就覆盖哪边
- 线的样式：`single` / `double` / `dotted` / `dashed` / `none`
- `repeat: true` 表格跨页自动重复表头

### 2.6 封面

```json
"cover": {
  "enabled": true,
  "top_space_cm": 6.5,        // 标题距页面顶端
  "above": ["内部资料", "编号：RPT-2026-017"],   // 标题上方若干行
  "below": ["数据智能研究中心"],                  // 日期下方若干行
  "bottom_space_cm": 2.0,
  "restyle": true,            // 把 Title/Subtitle/Author/Date 换成 CoverTitle 等样式
  "page_number": false,       // 封面不编页码
  "header": false             // 封面不要页眉
}
```

封面的标题/副标题/作者/日期来自 md 的 YAML 头：

```markdown
---
title: 面向多源异构数据的高效融合与推理方法研究报告
subtitle: 关键技术、实验验证与工程实践
author: 数据智能研究中心
date: 2026 年 10 月
---
```

它们会被自动搬进封面节，样式换成 `CoverTitle` / `CoverSubtitle` / `CoverAuthor` / `CoverDate`
（这四个样式同样在 `styles` 里可调，想加下划线、改字号都行）。

### 2.7 目录 / 图目录 / 表目录

```json
"front_matter": {
  "toc_title": "目　录",           // null 则不写标题
  "page_number": true,
  "page_num_fmt": "upperRoman",    // 目录页用 I II III
  "page_num_start": 1,
  "header": false,                 // 目录页不要页眉
  "figures": { "enabled": true, "title": "图目录", "label": "图" },
  "tables":  { "enabled": true, "title": "表目录", "label": "表" }
}
```

- 图目录/表目录是 Word 的 `TOC \c "图"` 域，跟着题注走，增删图表后 F9 自动更新
- `toc.levels` 控制主目录显示到几级标题（`"1-3"`）

### 2.8 分节页码（封面不编号 / 目录罗马 / 正文从 1 开始）

文档会被自动切成三节：

| 节 | 内容 | 页码 | 页眉 |
|---|---|---|---|
| 1 | 封面 | 无 | 无 |
| 2 | 目录 / 图目录 / 表目录 | 罗马数字，从 I 开始 | 默认无 |
| 3 | 正文 | 阿拉伯数字，从 1 开始 | 有 |

```json
"body_paging": { "page_num_fmt": "decimal", "page_num_start": 1 }
```

页脚里可用三个占位：`{PAGE}` 当前页、`{NUMPAGES}` 全文页数、`{SECTIONPAGES}` 本节页数。
想要「第 1 页 共 5 页（本节）」就把 `footer.format` 写成
`"第 {PAGE} 页 共 {SECTIONPAGES} 页"`。

实现方式：向文档插入两个分节符（`sectPr`），每节独立页脚部件 +
`w:pgNumType` 指定编号格式与起始号。页脚域**不带** `\* ARABIC`，否则会覆盖罗马格式。

### 2.9 页眉页脚

页眉文字、字号、位置都可配：

```json
"header": { "text": "XX 单位内部报告", "size_pt": "小五", "align": "center",
            "bottom_border": true, "use_chapter": true },
"footer": { "format": "第 {PAGE} 页 共 {NUMPAGES} 页", "align": "center" }
```

- `use_chapter: true` → 页眉显示**当前章节标题**（Word 的 STYLEREF 域，章节变了自动跟着变）
- `even_text: "..."` → 偶数页用另一段页眉

> ⚠️ **奇偶页页眉的副作用**：一开启 `even_text`，Word 会让每一节都从**奇数页**开始，
> 于是封面/目录/正文之间会多出空白页。这是双面印刷时「每章从右页开始」的惯例，
> 单面打印就不想要。不想留白页就别设 `even_text`。

### 2.10 目录条目样式

Word 更新目录后，条目用的是 `TOC1`/`TOC2`/`TOC3` 样式。pandoc 默认模板里没有这几个样式，
不定义的话目录会用 Word 的默认字体（Calibri），和正文对不上。现在自动生成并可配：

```json
"TOC1": { "size_pt": "小四", "bold": true, "space_before_pt": 6 },
"TOC2": { "size_pt": "小四" },
"TOC3": { "size_pt": "小四" }
```

### 2.11 中文排版细节

正式中文报告这几项默认都该开，现在统一在 `styles.Normal` 里：

```json
"Normal": {
  "widow_control": true,     // 孤行控制（段落不只剩一行在页尾/页首）
  "kinsoku": true,           // 按中文习惯控制首尾字符
  "overflow_punct": true,    // 允许标点溢出边界
  "auto_space_de": true,     // 中英文之间自动加间距
  "auto_space_dn": true      // 中文与数字之间自动加间距
}
```

另外 pandoc 默认模板里标题带主题蓝（`accent1`），正式报告要黑色——已自动改掉。


```json
"header": { "text": "XX 单位内部报告", "align": "center", "bottom_border": true },
"footer": { "page_number": true, "format": "第 {PAGE} 页 共 {NUMPAGES} 页" },
"page":   { "first_page_different": true }
```

`{PAGE}` / `{NUMPAGES}` 会变成 Word 域，不是死数字。

---

## 三、不同报告用不同格式：profile 档案

```bash
profiles/technical.json   技术报告（默认）
profiles/official.json    党政机关公文（仿宋三号、一、（一）、全框线、— 1 —）
profiles/academic.json    学位论文（宋体小四、图 2-1 按章编号、标题分页）
```

用法：

```bash
python3 md2docx.py report.md out.docx --profile official
```

要做新格式？**复制一份改**，或走问答生成（下一节）。

---

## 四、让 agent / 用户问答着把格式定下来

格式是视觉问题，一次问 68 项没人受得了。现在是**分轮 + 样张确认**的闭环：

### 4.1 从现成模板反推（有模板时最快）

```bash
python3 configure.py --from-docx 单位模板.docx --out 客户A.json
python3 configure.py --report 客户A.json
```

会从 docx 里读出：正文字体字号、各级标题、题注、缩进、行距、段前后距、
页边距、页眉文字、页脚页码格式，直接落成档案。

### 4.2 分轮提问（没有模板时）

```bash
python3 configure.py --ask --md 报告.md --round 1
```

输出的是 `ask_user_question` 能直接吃的 JSON：

```json
{ "round": 1, "title": "先定骨架…",
  "current": {"num_style": "1 / 1.1 / 1.1.1"},
  "questions": [
    {"id": "num_style", "header": "标题编号",
     "question": "标题怎么自动编号？",
     "options": [{"label": "1 / 1.1 / 1.1.1（当前）", "description": "当前值"}, …]}
  ] }
```

**会先扫一遍 md**，只问这份文档用得上的项：没有图就不问图注，没有
`bibliography:` 就不问参考文献风格，没有脚注就不问脚注位置。

五个轮次：

| 轮次 | 主题 |
|---|---|
| 1 | 从哪套起步、标题编号、封面 |
| 2 | 正文字体/字号/行距/缩进/对齐、各级标题字号 |
| 3 | 图、表、公式的编号与正文引用外形、表格线框 |
| 4 | 纸张边距、页眉页脚、目录、页码 |
| 5 | 摘要、脚注、列表、参考文献、修订模式 |

每轮答完就落档案：

```bash
echo '{"body_font": "仿宋", "body_size": "三号"}' > a1.json
python3 configure.py --base technical --answers a1.json --out 报告.profile.json
```

### 4.3 出样张让用户用眼睛确认

```bash
python3 configure.py --sample 报告.profile.json
# → build/样张_报告.profile.docx
```

样张一页半，包含：三级标题、两段正文、加粗/行内代码、列表、引用块、
图注、表注+三线表、带编号公式、代码块、脚注。**不包含封面和目录**，避免浪费时间。

### 4.4 看改了什么

```bash
python3 configure.py --diff profiles/technical.json 客户A.json
#   54 处差异（左=档案A，右=档案B）：
#     styles.BodyText.first_line_indent_chars  2 -> 0
#     footer.format  '第 {PAGE} 页' -> '— {PAGE} —'
```

### 4.5 一句话改一项 / 批量改

```bash
python3 configure.py --base x.json --out x.json --set styles.Caption.bold=true
python3 configure.py --base x.json --out x.json --answers a2.json
python3 configure.py --show x.json        # 打印当前格式清单
```

## 五、docxfix.py 干了什么

pandoc 只会写「图 1」这种死文字。docxfix 把它换成域：

| 处理 | 实现 | 效果 |
|---|---|---|
| 图/表题注 | `SEQ` 域 | 增删图表后 F9 自动重排 |
| 公式编号 | 居中制表位 + 右对齐 `SEQ` 域 | 公式居中、编号靠右 |
| 交叉引用 | `REF` 域 | 编号变了自动跟着变 |
| 目录 | `TOC` 域 | 可更新、带页码 |
| 标题编号 | Word 多级列表绑定 Heading1-4 | 改顺序自动重排 |
| 表格线框/表头 | 按 profile 直接写进 XML | 换预设即换样式 |
| 图片溢出 | 超过版心等比缩小 | 不撑破页面 |
| 斜体 | 全局强制关闭 | 满足「全文不要斜体」 |
| 自动刷新 | `updateFields` | 打开时自动更新域 |

---

## 六、md 写法速查

```markdown
---
title: 报告标题
subtitle: 副标题
author: 部门
date: 2026 年 10 月
---

{{TOC}}                                   <!-- 目录占位（profile 里 auto:true 可省） -->

# 一级标题                                 <!-- 自动编号 -->
## 二级标题

正文段落，首行自动缩进。

![架构图](img/a.png){#fig:arch width=13cm}
见图 {{fig:arch}}。

| 指标 | 目标 | 实测 |
|:---|---:|---:|
| 时延 | 200 ms | 185 ms |

: 关键指标 {#tbl:stat}                     <!-- 表注写在表格下面 -->

见表 {{tbl:stat}}。

::: {#eq:cap}                             <!-- 公式必须用围栏才能拿到 id -->
$$
Q = \frac{N \cdot B}{T_{p}} \times \eta
$$
:::

由式 {{eq:cap}} 可知。

> 引用块

```python
print("代码块")
```
```

规则：

- `{#fig:x}` `{#tbl:x}` `{#eq:x}` 是锚点，**必须唯一**
- 公式必须 `::: {#eq:x}` 围栏
- 图片用**相对路径**（`assets/a.png`），不要用剪贴板临时路径
- 别再手写「图 1」「1.1」这类编号

---

## 七、交付验收：md2docx.py verify

```bash
python3 md2docx.py verify 报告.docx          # 人话报告
python3 md2docx.py verify 报告.docx --json   # 机器可读
python3 md2docx.py verify dist/              # 整个目录
```

一次把「程序能判断的」全部判掉，只剩少数需要眼睛的：

| 类别 | 查什么 |
|---|---|
| **[错误]** | 引用找不到落点、题注缺 SEQ 域、跨部件引用悬空、元素顺序/重复、书签不成对、域不配对、页脚缺页码域、未替换的 `{{fig:x}}` |
| **[提醒]** | 图/表题注可能被分页拆开（缺 `keepNext`）、图片有效 dpi < 150、长表没设跨页重复表头、目录条目缺制表位、关键词带方括号、锚点重名、表格列数与网格不一致 |
| **[需肉眼]** | 通常 0~3 行：题注是否同页、编号是否连续（F9 之后再看） |

**为什么需要它**：结构对不对、编号会不会重排，本来都能程序判断。
以前只能导 PDF 逐页看，20 页要看 20 张图；现在一条命令就有结论。

## 七点五、只看一张图：联系表（**唯一允许的看图方式**）

> agent 注意：**禁止把 PDF 逐页转 PNG 再一张张 read**。
> 18 页报告那样看要烧几十万 token，而 `md2docx.py verify` 已经判完了能判的。
> 交付时只给 docx 路径；确实要看图，就一次出一张联系表。

```bash
python3 md2docx.py preview 报告.docx --sheet          # 自动挑关键页
python3 md2docx.py preview 报告.docx --pages 1,3,8    # 指定页
python3 md2docx.py preview 报告.docx --sheet --cols 3 # 一行三页
```

自动挑的是：封面、目录页、正文首页、最后一页。输出
`报告_联系表.png`，一张图看完，不用逐页截图。

**md2docx.py preview 默认不会启动 Word**（原因见下一节）。所以：

```bash
# 已有 PDF（你自己导出的）→ 直接拿来拼联系表，不碰 Word
python3 md2docx.py preview 报告.docx --sheet

# 需要脚本去调 Word 渲染时才加 --word
python3 md2docx.py preview 报告.docx --sheet --word

# 诊断 Word 状态 / 清理本工具的残留
python3 md2docx.py preview --doctor
python3 md2docx.py preview --cleanup
```

需要 Pillow（`pip install pillow`）；没有的话会退化成导出整份 PDF。

### 7.9 verify 的检查清单

```bash
python3 md2docx.py verify 报告.docx          # 单文件
python3 md2docx.py verify dist/              # 整个目录
```

分两层：

**结构合规**（Word 会不会说「文件已损坏」）

| 检查 | 说明 |
|---|---|
| 元素顺序 | `cantSplit` 在 `tblHeader` 前、`headerReference` 最前、`settings.xml` 顺序 |
| 重复元素 | 同一个父元素下不能有两个 `footnotePr` 之类 |
| 书签配对 | `bookmarkStart` / `bookmarkEnd` 必须成对 |
| 跨部件引用 | `w:endnote` 指向的部件必须存在、`numId` / `abstractNumId` 必须已定义、`Content_Types` 不能登记幽灵部件 |

**交付质量**（交出去能不能用）

| 检查 | 说明 |
|---|---|
| 目录制表位 | TOC1-9 必须有右对齐点线制表位，否则页码不右对齐、条目全顶格 |
| 交叉引用落点 | 每个 `REF xxx` 都要能找到对应书签，否则 F9 后变「错误！未找到引用源」 |
| 图片宽度 | 不能超版心 |
| 页脚页码域 | 页脚里要有 `PAGE` 域 |
| 题注序号域 | 题注必须是 `SEQ` 域，否则增删图表不重排 |
| 文献编号域 | 书目编号必须是 `SEQ 文献` 域，正文角标必须是 `REF` 域，否则删条目不重排 |
| 角标跳转 | 文献角标应能点开跳到条目 |
| 兼容性模式 | 缺 `compatibilityMode=15` 会显示「兼容性模式」 |
| 关键词格式 | 不该带方括号 |

模板类文件（没有域、段落很少）会自动跳过交付质检。

## 九、几个容易困惑的默认行为

| 现象 | 原因 | 怎么办 |
|---|---|---|
| Word 打开问「是否更新域」 | `settings.xml` 里 `updateFields=true` | 选「是」，编号和目录才会刷新 |
| 不更新域时目录只有一行占位文字 | 页码要排版后才知道，没法预填 | 同上；或手动 Ctrl+A → F9 |
| 标题栏显示「兼容性模式」 | 已补 `compatibilityMode=15`，旧文件才会显示 | 重新生成 |
| 一条命令跑完想再改格式 | 改档案（profile）不要改 docx | `configure.py --set 键=值` 后重跑 |

## 十、Word 报「文件可能已经损坏」怎么办

**先自查：**

```bash
python3 md2docx.py verify 报告.docx
```

它检查 Word 最在意的四类问题：

| 检查项 | 说明 |
|---|---|
| 重复子元素 | `sectPr` 里不能有两个 `footnotePr`、`trPr` 里不能有两个 `tblHeader` |
| 元素顺序 | `cantSplit` 必须排在 `tblHeader` 前；`headerReference` 必须最前；`settings.xml` 里 `trackChanges` 不能排在 `updateFields` 后面 |
| 书签成对 | 每个 `bookmarkStart` 必须有同 id 的 `bookmarkEnd` |
| 样式存在 | 段落引用的样式必须在 `styles.xml` 里存在 |

**为什么偏偏 Word 报错，LibreOffice / python-docx 都正常？**
Word 严格按 OOXML schema 校验，另外两个很宽容。
所以「能打开」不等于「合规」。

**本工具已做的防护：**
每次转换的最后一步是 `normalize_structure()`，把所有容器的子元素按 schema 排序，
再由 `validate()` 复查一遍，有问题会在终端打印出来。

**旧文件（本次修复之前生成的）就地修：**

```bash
cd <skill 目录>
python3 - <<'EOF' 旧文件.docx
import sys, importlib.util
SK = os.path.dirname(os.path.abspath("docxfix.py"))   # 或在 skill 目录里执行
sys.path.insert(0, SK)
spec = importlib.util.spec_from_file_location("dx", SK + "/docxfix.py")
dx = importlib.util.module_from_spec(spec); spec.loader.exec_module(dx)
from docx import Document
d = Document(sys.argv[1])
print("修正", dx.normalize_structure(d), "处顺序问题")
dx.validate(d); d.save(sys.argv[1])
EOF
```

（重复元素无法靠排序修复，那种只能重新生成。）

## 十一、Word 弹「此文件被锁定 / 文件已损坏」怎么办

这两种弹窗都是**脚本驱动 Word** 造成的，不是你机器坏了。

**根因**：

1. Word 在 macOS 上只有一个实例。两个会话（或两个 agent）同时让 Word 打开文件，
   后打开的那个就会看到「锁定者：另一用户」。
2. `md2docx.py preview` 会把待预览的 docx 复制到 Word 沙箱容器里再打开。
   如果上一次的副本没清掉、或那份文档还开在 Word 里，同名文件再次打开就报锁定。
3. 这版 Word 的 AppleScript **没有可用的 close 命令**，文档关不掉、越积越多。

**skill 现在的做法**（不需要你做任何事）：

| 措施 | 效果 |
|---|---|
| 默认**不启动 Word** | agent 跑流程时根本不会碰 Word |
| 文件名 = 时间 + 进程号 + 随机 | 两个会话永不撞名 |
| 开跑前清掉自己的临时文件和 `~$` 锁文件 | 不会读到别人的残留 |
| 转完用 Cmd+W 关掉刚打开的文档 | 不再累积 |
| 有比 docx 新的 PDF 就直接复用 | 迭代时不重复渲染 |
| **不会冷启动 Word** | 冷启动的 Word 会卡在引导页，脚本拿不到文档（`-2700`） |

**你需要做的**：

```bash
python3 md2docx.py preview --doctor        # 看 Word 进程 / 打开的文档 / 本工具残留
python3 md2docx.py preview --cleanup       # 清掉本工具留下的临时文件
```

如果 `--doctor` 显示「Word 有关不掉的文档」或渲染报 `-2700` / `missing value`：
在 Dock 上右键 Word → 退出（退不掉就用活动监视器强制退出），然后
**自己双击 docx 看效果**，或「文件 → 导出为 PDF」后不带 `--word` 再跑一次
（会自动复用那份 PDF 拼联系表）。

真正判断格式对不对，用 `md2docx.py verify`，它不需要任何渲染器。

### 11.5 高级开关（档案里没写，但代码支持）

拷到档案里就能用，不需要改代码：

| 键 | 作用 |
|---|---|
| `page.gutter_cm` | 装订线（左侧留白加宽） |
| `page.landscape` | 整篇横向（单表横排用 `landscape: ["tbl:x"]`） |
| `page.doc_grid.{type,line_pitch,char_space}` | 文档网格：固定行距/字距 |
| `styles.<X>.char_spacing_pt` | 字间距（磅），负值收紧 |
| `styles.<X>.char_scale` | 字宽百分比（如 90 = 压扁 10%） |
| `table.cell_margin_cm` | 单元格边距（上下左右） |
| `table.col_widths` / `table.col_widths_mode` | 列宽：比例数组 / `"auto"` 按内容分配 |
| `table.per_table` | 单张表单独覆盖线框/表头/列宽（按 `tbl:x` 锚点） |
| `header.use_chapter` | 页眉显示当前一级标题（STYLEREF 域） |
| `header.even_text` | 偶数页页眉；**设了会启用奇偶页不同** |
| `notes.footnote.separator_ids` | 脚注分隔线样式 id |
| `citations.link_citations` | 文献角标是否可点击跳转（默认 true） |
| `document.compatibility_mode` | 兼容模式版本号（默认 15） |

## 十二、已知限制

1. pandoc 会重写 `numbering.xml`，所以标题编号必须在 docxfix 阶段做，不能写进模板
2. 域不会自己更新，若 Word 未自动刷新：**Ctrl+A → F9**
3. 转 PDF 前必须先更新域
4. 公式 id 只能通过 `::: {#eq:x}` 围栏获得
5. `--profile` 生成的档案会丢掉 JSON 注释（手写档案保留注释）
6. 中文字体名要在目标机器上存在（`仿宋`/`方正小标宋简体` 等 Windows 字体 macOS 往往没有）
7. WPS 对域和多级列表的兼容性未验证，交付前建议用 WPS 开一遍
8. `--reference` 用的模板 docx 必须含 pandoc 那套样式（ImageCaption/TableCaption 等），
   否则题注功能失效。**用本工具生成的 docx 当模板就没问题**
9. 同一层级里混用中文序号（`一、`）和阿拉伯（`1.1`）时，
   Word 会把上级序号也按中文渲染（`三.1`）；docxfix 会自动加 `isLgl` 修正为 `3.1`，
   但如果你手写 `lvlText` 又不想要这个行为，得去掉 `isLgl`

---

## 十三、目录结构

```
md2docx/
├── SKILL.md               skill 入口（agent 读这个）
├── README.md              完整手册（人读这个）
├── requirements.txt       Python 依赖
├── .gitignore
│
├── md2docx.py             唯一入口：convert / make / batch / verify / preview / styles
├── configure.py           格式档案：问答 / 反推模板 / 样张 / 对比 / 格式清单
├── _common.py             共用小函数（路径、退出、档案解析）
│
├── build_template.py      引擎①：profile → reference.docx（样式/页面/页眉页脚）
├── docxfix.py             引擎②：静态文字 → Word 域（题注/引用/目录/分节/表格）
│
├── questions.json         问答清单（72 条，可自行增删）
├── profiles/              三套内置档案：technical / official / academic
├── csl/                   三套国标参考文献样式（离线）
│
├── selftest.py            自检整条管线（138 项）
├── docs/
│   ├── TESTING.md         人工验收脚本（给人看的）
│   └── REPO-HINT.md       放进多 skill 仓库时的组织与安装方式
└── examples/              样例：sample-report.md / complex-sample.md / sample.bib / *.png
```

一条命令记法：

```bash
python3 md2docx.py <md>          # 转换
python3 md2docx.py make <md>     # 转换 + 验收（推荐）
python3 md2docx.py verify x.docx # 验收
python3 md2docx.py preview x.docx --sheet   # 看效果（一张图）
python3 configure.py --list      # 定制格式
```
