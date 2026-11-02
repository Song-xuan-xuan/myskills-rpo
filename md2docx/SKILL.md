---
name: md2docx
description: 把 Markdown 报告转成排版规范、编号可自动重排的 Word(docx)。当用户要把 md 稿子做成正式 Word 报告、需要图/表/公式自动题注与交叉引用、需要封面和分节页码、要讨论或定制字体字号缩进页眉页脚格式、或用户要改已有报告的排版格式时使用。
---

# md2docx：Markdown → 规范 Word 报告

本目录自带完整工具链，**不要自己重新写**。

```bash
# 依赖（只需一次）
brew install pandoc            # 或 apt install pandoc
pip install -r requirements.txt

# 自检（换机器后先跑这个，138 项）
python3 selftest.py
```

**只有一个入口**（下面常用这几条就够）：

| 做什么 | 命令 |
|---|---|
| 转换 | `python3 md2docx.py 报告.md` |
| 转换 + 验收（推荐） | `python3 md2docx.py make 报告.md` |
| 一份 md 出多版 | `python3 md2docx.py make 报告.md --profiles technical,official` |
| 交付验收 | `python3 md2docx.py verify 报告.docx` |
| 看效果（只出一张图） | `python3 md2docx.py preview 报告.docx --sheet` |
| 定制格式 | `python3 configure.py --list` / `--ask` / `--from-docx` / `--report` |

目录里 `examples/` 是样例 md，`docs/TESTING.md` 是人工验收脚本（给人看的）。

## 产物约定（重要）

**转换只产出 docx 一个文件**，不要在工作目录里生成侧车 JSON、缓存目录、临时副本。
需要额外产物时只在用户明确要求时做：

| 产物 | 什么时候出现 |
|---|---|
| `报告.docx` | 每次转换，**唯一默认产物** |
| `报告.pdf` / `报告_联系表.png` | 仅 `md2docx.py make --pdf` / `md2docx.py preview --sheet` |
| `dist/*.docx` | 仅 `md2docx.py batch`（一份 md 出多版） |
| `样张_*.docx` | 仅 `configure.py --sample` |

## 什么时候用

- 用户要「把这份 md 转成 Word 报告 / 交稿版本」
- 用户抱怨 Word 里正文缩进把标题表格也缩进了、图表没有自动编号、交叉引用要手改
- 用户要「题注加粗」「正文改成仿宋三号」「标题编号用 1. 1.1.」这类格式调整
- 用户要给不同项目套不同格式（公文体、学位论文体、公司模板）

## 四阶段工作流（顺序不能颠倒，不要一步到位）

```
① 内容     大纲 → md        ← 只谈内容，不谈格式，不转换
② 定稿     用户确认 md      ← 硬门禁：没确认不许往下走
③ 格式     问格式 / 套模板   ← 内容定了才谈格式
④ 导出     docx + verify    ← 用户要图才渲染
```

### 阶段 ①：先把内容做对 —— 不碰格式

用户说「帮我写个报告」「调研一下 X 出一份报告」时，**默认只交付 md**：

1. 讨论大纲 → 用户认可
2. 写 / 改 md（内容、结构、图表的锚点 `{#fig:x}`、引用 `{{fig:x}}`）
3. **停下来，把 md 交给用户看**

这个阶段禁止做的事：

- ❌ 不要问字体字号页边距（内容还没定，问了也白问）
- ❌ 不要转换 docx
- ❌ 不要开 Word / 导 PDF / 截图

### 阶段 ②：等用户确认内容 —— 硬门禁

**必须看到用户明确说「内容可以了 / md 没问题 / 就这么定」才进入阶段 ③。**
用户没说，就停在阶段 ①，可以问一句：

> md 已经写好（路径），内容你看一遍？确认后我再套格式导 Word。

（用户一开始就说「我已经有 md，直接转 Word」→ 视为已确认，直接进阶段 ③。）

### 阶段 ③：这才开始谈格式

先判断用户是「只想出稿」还是「要按单位模板排」：

#### 情况 A：用户手里有现成的 Word 模板

直接用反推，比逐项问快得多：

```bash
python3 $SKILL/configure.py --from-docx 单位模板.docx --out 客户A.json
python3 $SKILL/configure.py --report 客户A.json     # 念给用户听，确认识别对不对
```

它会读出字体、字号、缩进、行距、页边距、页眉文字，直接落成档案。

#### 情况 B：没有模板 —— 分轮次问，别一口气问完

**先扫 md，只问这份文档用得上的项**（没有图就不问图注，没有参考文献就不问 CSL）：

```bash
python3 $SKILL/configure.py --ask --md 报告.md --round 1
```

输出的是 `ask_user_question` 直接能用的 JSON（含 `id`/`header`/`question`/`options`），
把它喂给 `ask_user_question` 工具即可。五个轮次：

| 轮次 | 问什么 |
|---|---|
| 1 | 从哪套起步、标题怎么编号、要不要封面 |
| 2 | 正文字体 / 字号 / 行距 / 首行缩进 / 对齐 |
| 3 | 图、表、公式的编号外形与正文引用外形 |
| 4 | 纸张边距、页眉页脚、目录、页码 |
| 5 | 摘要、脚注、列表、参考文献、修订模式 |

**每轮问完就落档案，不要攒到最后。**

```bash
echo '{"body_font": "仿宋", "body_size": "三号", "body_line": "固定28.8磅"}' > /tmp/a1.json
python3 $SKILL/configure.py --base technical --answers /tmp/a1.json --out 报告.profile.json
```

#### 关键一步：生成样张让用户看

**格式是视觉问题，光念文字用户没概念。** 每问完一轮，出一张样张：

```bash
python3 $SKILL/configure.py --sample 报告.profile.json
# → 当前目录/样张_报告.docx（一页半：各级标题、正文、图注、表注、公式、脚注、列表、代码块）
```

把**文件路径**给用户，让他自己打开看，问「这版可以吗？哪里还要改」。
**不要自己转 PDF 截图去看**——样张是给用户看的。

#### 情况 C：用户直接说清要求

「正文仿宋三号、标题编号用一、 不要封面」——不必再问，`--set` 批量设好，
出样张，然后念一遍 `configure.py --report` 让用户确认。

### 阶段 ④：转换 + 自动验收

```bash
# 本文件所在目录就是 skill 根目录，下面用 $SKILL 指代（别硬编码绝对路径）
SKILL="$(cd "$(dirname "${BASH_SOURCE[0]:-$0}")" && pwd)"   # 交互式里可直接写实际路径

# 一条命令：转换 + 质检 + 摘要（推荐）
python3 $SKILL/md2docx.py make 报告.md

# 要多个版本（技术版/公文版/论文版）
python3 $SKILL/md2docx.py make 报告.md --profiles technical,official,academic

# 只要一版、指定档案
python3 $SKILL/md2docx.py 报告.md 交付版.docx --profile 报告.profile.json
```

`--profile` 可以给档案名，也可以给路径（相对当前目录）。

**到此为止就交付，不要自己去看渲染效果**——`md2docx.py make` 已经跑了 `verify`。
只有用户说「给我看看长什么样」时才渲染，且**一次只出一张联系表**：

```bash
python3 $SKILL/md2docx.py preview 交付版.docx --sheet   # 需要现成 PDF 或显式 --word
```

### 阶段 ⑤：按反馈改格式（改档案，不改代码）

用户说「题注加粗」「正文改成仿宋」「标题编号用 1. 1.1.」时，**改档案，不改代码**：

```bash
# 改单个键
python3 $SKILL/configure.py --base <md名>.profile.json --out <md名>.profile.json \
        --set styles.Caption.bold=true

# 改一组（有现成翻译的项）
echo '{"num_style": "1. / 1.1. / 1.1.1."}' > /tmp/a.json
python3 $SKILL/configure.py --base <md名>.profile.json --out <md名>.profile.json \
        --answers /tmp/a.json
```

改完重跑第 2 步，然后念一遍新的清单。

## 格式从哪来：两条路

| 路线 | 怎么改 | 适合 |
|---|---|---|
| **A. profile JSON**（默认） | 改 `styles.*` / `captions` / `refs` / `table` / `page` 等键 | 绝大多数情况 |
| **B. 自己的 Word 模板** | 在 Word 里改样式另存为 docx，转换时 `--reference 你的模板.docx` | 用户已有单位模板，想用 Word 的样式界面改 |

路线 B 的前提：模板 docx 必须是**本工具链生成的**（或含 pandoc 那套样式名的），
否则 pandoc 的 ImageCaption/TableCaption 等样式会缺失，题注功能会失效。

> ⚠️ **不要改 `build/*.docx`**。模板每次转换都会重新编译（避免不同项目串格式），
> 手改会被静默冲掉。要改就用路线 A 或 B。

### 常用「人话 → 配置键」对照

| 用户说 | 改哪个键 |
|---|---|
| 题注加粗 | `styles.Caption.bold`（图注表注都跟着变） |
| 题注字体/字号 | `styles.Caption.east_asia_font` / `.size_pt` |
| 正文换字体 | `styles.Normal.east_asia_font` / `.ascii_font` |
| 正文字号 | `styles.Normal.size_pt`（中文名如「三号」或磅值如 16） |
| 正文行距 | `styles.Normal.line_spacing`（倍数）或 `.line_spacing_pt`（固定磅） |
| 正文首行缩进 | `styles.BodyText.first_line_indent_chars`（`FirstParagraph` 继承它） |
| 正文两端对齐 | `styles.Normal.align` = `justify` |
| 标题字号/字体 | `styles.Heading1..6.size_pt` / `.east_asia_font` |
| 标题加粗 | `styles.Heading1..6.bold` |
| 标题编号形式 | `heading_numbering.levels[].format` + `.num_fmt`（用 `--answers` 传 `num_style` 更省事） |
| 图号长成「图 1」还是「(1)」 | `captions.fig` / `.tbl` / `.eq`，`{n}` 是编号占位 |
| 正文引用显示「图 1」还是「1」 | `refs.fig` / `.tbl` / `.eq` |
| 图号按章（2-1） | `captions.chapter` = true |
| 表格线框 | `table.preset`（three_line / full_grid / outline / none）+ `table.borders` |
| 表头底色 | `table.header.shading`（色值如 F2F2F2） |
| 页眉文字 | `header.text` |
| 页码格式 | `footer.format`（`{PAGE}` `{NUMPAGES}` `{SECTIONPAGES}`） |
| 封面加行 | `cover.above` / `cover.below`（数组） |
| 要不要封面 | `cover.enabled` |
| 目录层级 | `toc.levels`（如 "1-3"） |
| 目录页码罗马数字 | `front_matter.page_num_fmt` = `upperRoman` |
| 图目录/表目录 | `front_matter.figures.enabled` / `.tables.enabled` |

## md 写法：三条硬规则

1. **不要手写编号**（「图 1」「1.1」「表 2」），一律用锚点 + 引用：
   ```markdown
   ![说明](assets/a.png){#fig:arch}      → 图注自动编号
   见图 @fig:arch。                        → REF 域，编号变了自己跟着变

   | A | B |
   |---|---|
   | 1 | 2 |

   : 表说明 {#tbl:x}                      → 表注写在表格下方

   ::: {#eq:cap}                          → 公式必须用围栏，否则拿不到 id
   $$
   Q = \frac{N B}{T}
   $$
   :::
   ```
2. **首行缩进只加在 `BodyText` / `FirstParagraph`**，绝不能加到 `Normal`——
   否则标题、表格、图注会一起缩进（这是最常见的用户痛点）。
3. **图片用相对路径**，不要用剪贴板临时路径（`/var/folders/...`），否则换机器就丢图。

## 摘要 / 脚注 / 列表 / 表格列宽 / 文档开关

| 用户说 | 改哪个键 |
|---|---|
| 要摘要页 | md 头部写 `abstract: \|` 和 `keywords: [甲, 乙]`；`abstract.enabled` 控制 |
| 摘要标题 | `abstract.title` |
| 脚注放页底 / 每页重排 | `notes.footnote.pos` / `.num_restart` |
| 脚注用一二三编号 | `notes.footnote.num_fmt` = `chineseCounting` |
| 有序列表用 1) 或一、 | `lists.ordered.format`（`%1` 是层级号）+ `.num_fmt` |
| 无序列表符号 | `lists.bullet.char` |
| 某张表列宽 | `table.col_widths`：`{"tbl:锚点": [30, 70]}` |
| 表格行不被拆页 | `table.cant_split` = true |
| 单张表换样式 | `table.per_table`：`{"tbl:x": {"preset": "full_grid", "col_widths": [...]}}` |
| 首字下沉 | `drop_cap.enabled` |
| 每页固定 22 行（公文） | `page.doc_grid.lines_per_page` |
| 开修订模式 | `document.track_changes` |
| 限制编辑 | `document.protect` = `readOnly` / `comments` |
| 引用带页码 | `refs.sec` = `"第 {n} 节（第 {p} 页）"` |
| 超链接颜色（去蓝） | `styles.Hyperlink.color` = `000000` |
| 三线表粗细 | `table.borders.top.sz` / `.bottom.sz`（单位 1/8 磅，12=1.5磅）、`table.borders.header_bottom.sz` |
| 长表要不要禁拆行 | `table.cant_split`：短表开、超过 20 行的长表别开（会挤走版面） |
| 列宽按内容自动分 | `table.col_widths` = `{"tbl:x": "auto"}` 或 `table.col_widths_mode` = `auto` |
| 列表跟正文一样大、表格小一号 | `styles.Compact`（列表）/ `styles.TableText`（表格）分开调 |
| 图目录/表目录默认要不要 | `front_matter.figures.enabled` / `.tables.enabled`（默认：技术报告关、学位论文开） |
| 附录自动编号 A / A.1 | `appendix.enabled`，起点标题加 `{#sec:app}` |
| 封面 logo | `cover.logo` + `cover.logo_width_cm` |
| 一份 md 出多版 | `md2docx.py make 报告.md --profiles a,b,c` |
| 参考文献删条目后编号自动重排 | 默认开；`citations.number_bibliography: false` 可关 |
| 列表字号 vs 表格字号 | `styles.Compact`（列表）/ `styles.TableText`（表格）分开调 |
| 列宽按内容自动分 | `table.col_widths` = `{"tbl:x": "auto"}` |

## 交付前检查：**先跑 verify，不要截图**

```bash
python3 $SKILL/md2docx.py make 报告.md          # 转换 + 自动 verify + 打印格式摘要
```

只要转换不要摘要时：

```bash
python3 $SKILL/md2docx.py verify out.docx   # 程序化验收（推荐）
python3 $SKILL/md2docx.py styles out.docx   # 逐段看用了哪个样式、缩进多少
python3 $SKILL/configure.py --report <profile>   # 当前格式清单（念给用户听）
```

`md2docx.py verify` 会把能判的全判掉，分三类输出：

- **[错误]**：必须修（引用没落点、题注缺 SEQ 域、跨部件引用悬空、书页脚缺页码…）
- **[提醒]**：可能不是你要的效果（图/表题注可能被拆页、图片 dpi 偏低、长表未设表头重复…）
- **[需肉眼]**：程序判不了的少数几项，通常 0~3 行

### 关于「看效果」的硬规矩（很费 token，务必遵守）

**禁止用 read 打开单页图片。** 一份 18 页的报告如果逐页转 PNG 再一张张 read，
会烧掉几十万 token，而 `md2docx.py verify` 早就能把能判的都判完了。

允许的行为只有两种：

| 场景 | 做法 | 开销 |
|---|---|---|
| 交付给用户看 | 只给 **docx 路径**，让用户自己打开 | 0 |
| 确实需要图（用户要求 / 定位某个视觉 bug） | `md2docx.py preview --sheet`，**一次 1 张** | 1 张图 |

```bash
python3 $SKILL/md2docx.py preview 报告.docx --sheet     # 封面+目录+正文首末页，拼成一张
python3 $SKILL/md2docx.py preview 报告.docx --pages 1,3 # 只要特定页（也拼成一张）
```

**不要**：
- ❌ `pdftoppm` 逐页导出后一张张 `read`
- ❌ 每改一版格式就重新渲染一次
- ❌ 自己开 Word / 导 PDF（Word 冷启动会卡死，还会弹「文件被锁定」）

判断格式对不对，先跑这两个（都不依赖渲染器）：

```bash
python3 $SKILL/md2docx.py verify 报告.docx
python3 $SKILL/configure.py --report 报告.profile.json
```

## 踩过的坑（别重犯）

- pandoc 会重写 `numbering.xml`，**标题编号必须在 docxfix 后处理阶段做**，写进 reference.docx 会被丢掉
- 页脚域不能带 `\* ARABIC`，否则会盖掉分节的罗马数字格式
- pandoc 的 `TOCHeading` 样式 `basedOn Heading1`，会把目录标题编号成「1」并把正文顶成「2」，必须改挂 Normal
- pandoc 默认模板里 Heading1-9 带**主题蓝色**，正式报告要黑色（build_template 已统一改成黑）
- 上级标题用中文序号（`一、`）时，下级写 `%1.%2` 会渲染成「三.1」，
  docxfix 会自动加 `isLgl` 让它变成「3.1」
- 中文序号后面不该跟空格（`一、 标题`），docxfix 会自动把 `suff` 设成 `nothing`
- pandoc 3.x 是 `--syntax-highlighting=none`，2.x 是 `--no-highlight`（脚本已自动判断）
- 域不会自动更新：交付前 **Ctrl+A → F9**，或走 `md2docx.py preview`（它内部会 F9）
- 预览要把文件复制到 **Word 自己的沙箱容器**
  （`~/Library/Containers/com.microsoft.Word/Data/Documents/`），
  否则 Word 会弹「授予访问权限」
- 这版 Word 的 AppleScript **不支持 close 关闭文档**（会报错），
  所以 preview 用「唯一文件名」避免重复预览时读到旧内容；
  预览后 Word 里会留一个临时窗口（同一份内容只留一个）
- 交叉引用用 `{{fig:x}}` 而不是 `@fig:x`：citeproc 会把 `@key` 当文献引用吃掉
- pandoc 把表注放在表**前面**，docxfix 只把「紧邻题注的那个书签」搬进题注；
  标题的书签必须留在原地，否则章节交叉引用会指到图上去
- 开启「奇偶页页眉」后，Word 会让每节从奇数页开始，节与节之间多出空白页
  （双面印刷惯例；单面打印就别设 `header.even_text`）
- `refs.*` 别同时带前缀又在正文写「见**表**」：`refs.tbl="表 {n}"` +
  正文「见表 {{tbl:x}}」会得到「见表 表 1」。默认 `refs.*` 只出数字（`{n}`），
  正文自己写「见表 {{tbl:x}}」；想省事就把 `refs.tbl` 设成 `"表 {n}"` 并只写「见 {{tbl:x}}」
- 锚点写错时不会静默变成「?」了：保留原文并在终端列出错误的锚点
- `settings.xml` 带 `updateFields=true`，Word 打开会问「是否更新域」，选「是」。
  不选则目录只显示占位文字——因为页码要排版后才知道，没法预填
- 附录**不要手写**「附录 A」：写了会变成「附录 A 附录 A 数据说明」。
  只写标题 + `{#sec:app}` 锚点，编号前缀由 `appendix.label` 提供
- 附录里的二级编号要用每一级自己的 `num_fmt`：一级 upperLetter、二级 decimal，
  否则会算成「1.1」而不是「A.1」（docxfix 已处理）
- `-f markdown-smart` 必须开，否则中文引号被转坏
- **Word 对 OOXML 元素顺序极严**：`trPr` 里 `cantSplit` 必须在 `tblHeader` 前，
  `sectPr` 里 `headerReference` 必须最前，`settings.xml` 里 `trackChanges` 不能排在
  `updateFields` 后。顺序错了 Word 直接弹「文件可能已经损坏」，而 LibreOffice /
  python-docx 照常打开。docxfix 每次输出前会 `normalize_structure()` 统一排序，
  交付前用 `md2docx.py verify` 复查
- `footnotePr` / `endnotePr` 在 `sectPr` 里只能有一个：先删旧的再插新的，
  否则重复元素同样触发「已损坏」
- Word 单实例：两个会话同时预览会互相踩（「文件被锁定，锁者：另一用户」）。
  `md2docx.py preview` 有文件锁排队，并会清 `~$` 锁文件、关模态弹窗

## 环境依赖

pandoc 2.x/3.x、python3 + `python-docx`、macOS 上 `md2docx.py preview` 用 Microsoft Word 导 PDF。

Word 的 AppleScript 偶尔会失灵（打开无反应、active document 是别人的文件）。
`md2docx.py preview` 已经做了：唯一文件名、按文件名确认、清锁文件、关弹窗、失败重启 Word 重试。
若仍失败，直接让对方打开 docx 看效果，不要卡在预览上。
