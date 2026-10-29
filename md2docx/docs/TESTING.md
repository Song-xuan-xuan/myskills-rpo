# 人工验收脚本

> **这份文档是给人看的**（安装后逐条走一遍，确认功能都在）。
> agent 不要执行这里的手工步骤，它只需要跑 `verify.py`。

## 前提

把本目录放到 pi 能扫描到的地方（如 `.agents/skills/md2docx/` 或
`~/.pi/agent/skills/md2docx/`），在该目录或其父目录里开新会话。
会话已开着的话，先说 `/reload` 让 skill 重新加载。

---

## 第 1 步：确认 skill 被发现了（10 秒）

新会话里直接说：

> 帮我看看有哪些可用技能

或者强制调用：

> /skill:md2docx --list-profiles

期望输出：

```
可用 profile：
  academic       学位论文
  official       公文
  technical      技术报告
```

看到这个就说明 skill 加载成功、工具链可用。

---

## 第 2 步：管线自检（1 分钟）

```bash
python3 .agents/skills/md2docx/selftest.py
```

期望：`通过 136 项，失败 0 项`

它会检查依赖、编译模板、转样例，再验分节/页码/域/斜体/图宽/表头/标题颜色/编号分隔，
以及格式清单、`--set`、`--reference`、目录条目样式、中文排版、奇偶页页眉、横排页、
参考文献、章节交叉引用。

---

## 第 3 步：看格式清单（30 秒）

```bash
python3 .agents/skills/md2docx/format_report.py technical
```

期望看到一份中文清单（页面/正文/标题1-6/图题注/表题注/公式/表格/代码块/页眉页脚/封面/目录），
比如：

```
正文　　宋体 · Times New Roman(西文) · 小四(12磅) · 两端对齐 · 首行缩进2字符 · 行距1.5倍
标题1　　黑体 · 三号(16磅) · 居中 · 段前18 段后12 · 编号「1」
图题注　宋体 · 五号(10.5磅) · 居中 · 顶格 · 编号「图 1」 · 正文引用「1」
交叉引用 eq→(1) · fig→1 · sec→1 · tbl→1
```

---

## 第 4 步：真实转换 + 肉眼验收（2 分钟）—— **给人看的，agent 请跳过**

```bash
mkdir -p /tmp/try && cd /tmp/try
cp .agents/skills/md2docx/examples/{sample-report.md,sample.bib,flow.png} .

python3 .agents/skills/md2docx/md2docx.py sample-report.md out.docx
python3 .agents/skills/md2docx/preview.py out.docx
```

这一步是**人工验收脚本**：用眼睛确认版式。agent 不要执行这一步
（它应该只跑 `verify.py`，需要图时用 `preview.py --sheet` 出一张）。

按顺序看 7 处：

| 页 | 看什么 | 期望 |
|---|---|---|
| 1 | 封面 | 无页码、无页眉；标题/副标题/作者/日期在 |
| 2 | 目录 | 从「1 项目概况」开始；页码 `I`；下方有「图目录」 |
| 3 | 正文首页 | 页脚「第 1 页」；页眉有字；首行缩进 2 字符 |
| 3 | 图 | 图注在图下，图与图注没被拆页 |
| 4 | 公式 | 居中，编号 (1)(2)(3) 靠右 |
| 5 | 长表 | 跨页处表头重复出现 |
| 全文 | 引用 | 「见表 1」「见图 2」数字对得上；文献是 [1] [2]，末尾有国标格式的参考文献 |

> 预览用的是本机 Microsoft Word。Word 里会留下一个临时文档窗口
> （Word 的脚本接口不支持自动关闭），自己关掉即可。

---

## 第 5 步：验证配置真的驱动格式（1 分钟）

```bash
cp .agents/skills/md2docx/profiles/technical.json /tmp/try/mine.json
# 编辑 /tmp/try/mine.json：
#   styles.BodyText.first_line_indent_chars: 2 → 0
#   styles.CoverTitle.size_pt: "小一" → "二号"

cd /tmp/try
python3 .agents/skills/md2docx/md2docx.py sample-report.md t.docx --profile mine.json
python3 .agents/skills/md2docx/preview.py t.docx
```

缩进消失 + 封面标题变小 = 配置链路通。

也可以用 `--set` 不改文件：

```bash
python3 .agents/skills/md2docx/configure.py \
    --base /tmp/try/mine.json --out /tmp/try/mine.json \
    --set styles.Caption.bold=true
```

改完会自动打印新的格式清单。

---

## 第 5.5 步：看样张（1 分钟）

```bash
python3 .agents/skills/md2docx/configure.py --sample technical
open build/样张_technical.docx
```

一页半，把标题层级、正文缩进、图注、表注、公式编号、脚注、列表、代码块
一次看完。**这是确认格式最快的方式**，比读配置清单直观。

换成自己的档案再出一次：

```bash
python3 .agents/skills/md2docx/configure.py --sample /tmp/try/mine.json
```

## 第 5.6 步：从现成 Word 模板反推（有模板时）

```bash
python3 .agents/skills/md2docx/configure.py --from-docx 单位模板.docx --out 单位.json
python3 .agents/skills/md2docx/format_report.py 单位.json
```

核对识别结果对不对，不对就用 `--set` 改。

## 第 6 步：走一遍「agent 先问格式」的流程（关键）

这是本 skill 的核心设计。在新会话里说：

> 我要写一份项目总结报告，转成 Word

**期望 agent 先问格式**（而不是直接生成），问完再转，转完把格式清单念给你听。

若 agent 直接就生成了，说明 SKILL.md 没生效：

- 会话是旧快照 → 说 `/reload`
- 工作目录不在 skill 所在目录的祖先链里 → skill 扫不到

---

## 第 6.5 步：用你自己的报告

三个前提：

1. md 开头有 YAML 头（`title` / `subtitle` / `author` / `date`），否则封面是空的
2. 图片用**相对路径**，且和 md 的相对位置正确
3. 编号不手写，交叉引用写 `{{fig:x}}` / `{{tbl:x}}` / `{{eq:x}}` / `{{sec:x}}`

```bash
cd 你的报告目录
python3 .agents/skills/md2docx/md2docx.py 报告.md 交付版.docx --profile technical
```

---


## 第 7 步：新能力速验（3 分钟）

### 一份 md 出多版

```bash
python3 .agents/skills/md2docx/batch.py sample-report.md --outdir /tmp/dist
ls /tmp/dist     # 三个版本：technical / official / academic
```

### 附录自动编号

在 md 末尾追加：

```markdown
# 数据说明 {#sec:app}

附录正文。

## 明细表

补充明细。
```

档案里打开：

```bash
python3 .agents/skills/md2docx/configure.py --base technical --out /tmp/app.json \
        --set appendix.enabled=true --set appendix.marker=sec:app
python3 .agents/skills/md2docx/md2docx.py sample-report.md /tmp/app.docx --profile /tmp/app.json
```

打开后应看到「附录 A 数据说明」「A.1 明细表」。

### 封面 logo

```bash
python3 .agents/skills/md2docx/configure.py --base technical --out /tmp/logo.json \
        --set cover.logo=flow.png --set cover.logo_width_cm=3.5
python3 .agents/skills/md2docx/md2docx.py sample-report.md /tmp/logo.docx --profile /tmp/logo.json
```

封面顶部应出现图片，标题在其下方。

### 从模板反推

```bash
python3 .agents/skills/md2docx/configure.py --from-docx sample-report.docx --out /tmp/rev.json
python3 .agents/skills/md2docx/format_report.py /tmp/rev.json
```

应识别出：宋体小四、黑体三号标题、首行缩进 2 字符、1.5 倍行距、页边距 3.17cm、页眉文字。

### 结构合规（Word 修复弹窗的根因检查）

```bash
python3 .agents/skills/md2docx/verify.py /tmp/app.docx /tmp/logo.docx
```


## 第 8 步：一键交付（推荐给用户）

```bash
python3 .agents/skills/md2docx/make.py sample-report.md
```

一条命令做完：转换 → 结构合规 → 交付质检 → 打印本次用了哪套格式。
退出码 0 才算可以直接交。

多版本：

```bash
python3 .agents/skills/md2docx/make.py sample-report.md --profiles technical,official,academic
```

## 第 9 步：验证「不再靠截图」

以前验收要导 PDF、逐页看。现在：

```bash
python3 .agents/skills/md2docx/verify.py 你的报告.docx
```

输出分三类：[错误] 必须修 / [提醒] 可能不是你要的效果 / [需肉眼] 通常 0~3 行。
绿色且没有[错误]就可以交。

确实需要看效果时，只导一张图：

```bash
# 默认不启动 Word：若同目录已有 PDF 就直接用它拼图
python3 .agents/skills/md2docx/preview.py 你的报告.docx --sheet

# 确实要脚本调 Word 渲染时才加 --word
python3 .agents/skills/md2docx/preview.py 你的报告.docx --sheet --word
# → 你的报告_联系表.png（封面 + 目录 + 正文首末页拼成一张）
```

Word 弹「文件被锁定」时先看状态：

```bash
python3 .agents/skills/md2docx/preview.py --doctor
python3 .agents/skills/md2docx/preview.py --cleanup
```

## 交付前

- WPS 打开验一遍（域与多级列表的兼容性）
- 转 PDF 走 `preview.py`（内部会 F9 更新域），不要直接「另存为 PDF」
- 想更换档案：`--profile official` / `--profile academic` / 自定义 JSON

---

## 出问题怎么定位

```bash
python3 .agents/skills/md2docx/tools/show_styles.py out.docx    # 每段的样式/字体/缩进
python3 .agents/skills/md2docx/md2docx.py --list-profiles
python3 .agents/skills/md2docx/selftest.py
python3 .agents/skills/md2docx/configure.py --list        # 55 条配置问题
python3 .agents/skills/md2docx/configure.py --list --core # 15 条核心问题
python3 .agents/skills/md2docx/format_report.py technical # 当前格式清单
```
