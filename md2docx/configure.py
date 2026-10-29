#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
configure.py —— 定格式 / 生成档案（profile）

【最常用】从现成 Word 模板反推
    python3 configure.py --from-docx 单位模板.docx --out 客户A.json

【agent 分轮提问】
    python3 configure.py --ask --md 报告.md --round 1     # 输出 ask_user_question 能用的 JSON
    python3 configure.py --ask --md 报告.md               # 五轮全部输出
    python3 configure.py --list [--core]                   # 纯文本问题清单

【出样张，让用户用眼睛确认】
    python3 configure.py --sample 客户A.json

【落档案 / 改档案】
    python3 configure.py --answers answers.json --out 客户A.json
    python3 configure.py --base 客户A.json --out 客户A.json --set styles.Caption.bold=true
    python3 configure.py --show 客户A.json                 # 打印格式清单
    python3 configure.py --diff technical.json 客户A.json   # 看改了什么

【人机对话（终端里一路回车）】
    python3 configure.py --out 客户A.json
"""
import copy, json, os, re, shutil, subprocess, sys, tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import asklib
QUESTIONS = os.path.join(HERE, "questions.json")
PROFILE_DIR = os.path.join(HERE, "profiles")


def strip_json_comments(text):
    out, i, n, in_str = [], 0, len(text), False
    while i < n:
        c = text[i]
        if in_str:
            out.append(c)
            if c == "\\" and i + 1 < n:
                out.append(text[i + 1]); i += 2; continue
            if c == '"':
                in_str = False
            i += 1; continue
        if c == '"':
            in_str = True; out.append(c); i += 1; continue
        if c == "/" and i + 1 < n and text[i + 1] == "/":
            while i < n and text[i] != "\n":
                i += 1
            continue
        if c == "/" and i + 1 < n and text[i + 1] == "*":
            i += 2
            while i + 1 < n and not (text[i] == "*" and text[i + 1] == "/"):
                i += 1
            i += 2; continue
        out.append(c); i += 1
    return "".join(out)


def load_json(path):
    return json.loads(strip_json_comments(open(path, encoding="utf-8").read()))


def get_path(d, path):
    cur = d
    for part in path.split("."):
        if not isinstance(cur, dict) or part not in cur:
            return None
        cur = cur[part]
    return cur


def set_path(d, path, value):
    parts = path.split(".")
    cur = d
    for p in parts[:-1]:
        cur = cur.setdefault(p, {})
    cur[parts[-1]] = value


def resolve_profile(name):
    if not name:
        return os.path.join(PROFILE_DIR, "technical.json")
    for cand in (name, os.path.join(PROFILE_DIR, name + ".json")):
        if os.path.exists(cand):
            return cand
    p = os.path.join(PROFILE_DIR, name + ".json")
    if os.path.exists(p):
        return p
    raise SystemExit("找不到 profile: %s" % name)


# ---------- 选项 -> 配置 的映射 ----------

NUM_STYLES = {
    "1 / 1.1 / 1.1.1": [
        {"format": "%1", "num_fmt": "decimal", "indent_chars": 0},
        {"format": "%1.%2", "num_fmt": "decimal", "indent_chars": 0},
        {"format": "%1.%2.%3", "num_fmt": "decimal", "indent_chars": 0},
    ],
    "1. / 1.1. / 1.1.1.": [
        {"format": "%1.", "num_fmt": "decimal", "indent_chars": 0},
        {"format": "%1.%2.", "num_fmt": "decimal", "indent_chars": 0},
        {"format": "%1.%2.%3.", "num_fmt": "decimal", "indent_chars": 0},
    ],
    "一、 / （一） / 1.": [
        {"format": "%1、", "num_fmt": "chineseCounting", "indent_chars": 2},
        {"format": "（%2）", "num_fmt": "chineseCounting", "indent_chars": 2},
        {"format": "%3.", "num_fmt": "decimal", "indent_chars": 2},
    ],
    "第1章 / 1.1 / 1.1.1": [
        {"format": "第%1章", "num_fmt": "chineseCounting", "indent_chars": 0},
        {"format": "%1.%2", "num_fmt": "decimal", "indent_chars": 0},
        {"format": "%1.%2.%3", "num_fmt": "decimal", "indent_chars": 0},
    ],
    "不自动编号": None,
}

LINE_SPACING = {
    "单倍": ("line_spacing", 1.0),
    "1.25倍": ("line_spacing", 1.25),
    "1.5倍": ("line_spacing", 1.5),
    "2倍": ("line_spacing", 2.0),
    "固定20磅": ("line_spacing_pt", 20),
    "固定28.8磅": ("line_spacing_pt", 28.8),
}

CAPTION_PRESETS = {
    "图 1": ("图 {n}", False), "图1": ("图{n}", False),
    "图 1-1（按章）": ("图 {n}", True),
    "Figure 1": ("Figure {n}", False), "(1)": ("({n})", False),
    "表 1": ("表 {n}", False), "表1": ("表{n}", False),
    "表 1-1（按章）": ("表 {n}", True), "Table 1": ("Table {n}", False),
    "(1)": ("({n})", False), "（1）": ("（{n}）", False),
    "式 1": ("式 {n}", False), "(1-1)": ("({n})", True),
}

PAPER = {"A4": (21.0, 29.7), "A5": (14.8, 21.0), "Letter": (21.59, 27.94)}


def apply_answers(profile, answers, questions):
    qmap = {q["id"]: q for q in questions}
    chapter_flag = False
    for qid, val in answers.items():
        q = qmap.get(qid)
        if q is None:
            continue
        path = q.get("path")
        if not path:
            continue

        # 脚注位置
        if qid == "footnote_pos":
            set_path(profile, "notes.footnote.pos",
                     "pageBottom" if val.startswith("页底") else "beneathText")
            continue
        # 脚注编号格式
        if qid == "footnote_num":
            set_path(profile, "notes.footnote.num_fmt",
                     {"1, 2, 3": "decimal", "i, ii, iii": "lowerRoman",
                      "I, II, III": "upperRoman", "一, 二, 三": "chineseCounting"
                      }.get(val, "decimal"))
            continue
        # 脚注重排
        if qid == "footnote_restart":
            set_path(profile, "notes.footnote.num_restart",
                     {"全文连续": "continuous", "每页重排": "eachPage",
                      "每章重排": "eachSect"}.get(val, "continuous"))
            continue
        # 有序列表编号
        if qid == "list_ordered":
            fmt, nf = {"1. 2. 3.": ("%1.", "decimal"),
                       "1) 2) 3)": ("%1)", "decimal"),
                       "(1) (2) (3)": ("(%1)", "decimal"),
                       "一、二、三、": ("%1、", "chineseCounting")
                       }.get(val, ("%1.", "decimal"))
            set_path(profile, "lists.ordered.format", fmt)
            set_path(profile, "lists.ordered.num_fmt", nf)
            continue
        # 文档保护
        if qid == "doc_protect":
            set_path(profile, "document.protect",
                     {"不限制": None, "只读": "readOnly", "仅允许批注": "comments",
                      "仅允许修订": "trackedChanges"}.get(val, None))
            continue
        # 每页行数
        if qid == "grid_lines":
            if val:
                set_path(profile, "page.doc_grid.type", "lines")
                set_path(profile, "page.doc_grid.lines_per_page", int(val))
            else:
                set_path(profile, "page.doc_grid", None)
            continue
        # 参考文献风格
        if qid == "csl_style":
            set_path(profile, "citations.csl",
                     {"国标2005 数字 [1]": "chinese-gb7714-2005-numeric",
                      "国标2005 作者-年 (作者, 2020)": "chinese-gb7714-2005-author-date",
                      "国标1987 数字 [1]": "chinese-gb7714-1987-numeric"
                      }.get(val, "chinese-gb7714-2005-numeric"))
            continue
        # 章节交叉引用的外形
        if qid == "sec_ref_style":
            set_path(profile, "refs.sec",
                     {"只出章号（1.1）": "{n}", "第 1.1 节": "第 {n} 节",
                      "「章节标题」": "「{t}」"}.get(val, "{n}"))
            continue
        # 页眉显示当前章节标题
        if qid == "header_use_chapter":
            set_path(profile, "header.use_chapter", bool(val))
            continue
        # 标题是否加粗：一次改到所有标题层级
        if qid == "h_bold":
            for lvl in range(1, 7):
                set_path(profile, "styles.Heading%d.bold" % lvl, bool(val))
            set_path(profile, "heading_numbering.bold", bool(val))
            continue

        # —— 需要翻译成配置的几项 ——
        if qid == "paper":
            w, h = PAPER.get(val, PAPER["A4"])
            set_path(profile, "page.width_cm", w)
            set_path(profile, "page.height_cm", h)
            continue
        if qid == "body_line":
            key, num = LINE_SPACING.get(val, ("line_spacing", 1.5))
            set_path(profile, "styles.Normal.line_spacing_pt", None)
            set_path(profile, "styles.Normal.line_spacing", None)
            set_path(profile, "styles.Normal." + key, num)
            continue
        if qid == "body_align":
            set_path(profile, "styles.Normal.align",
                     "justify" if val == "两端对齐" else "left")
            continue
        if qid in ("h1_align",):
            set_path(profile, path.replace(".align_choice", ".align"),
                     "center" if val == "居中" else "left")
            continue
        if qid == "num_style":
            if NUM_STYLES.get(val) is None:
                set_path(profile, "heading_numbering.enabled", False)
            else:
                set_path(profile, "heading_numbering.enabled", True)
                set_path(profile, "heading_numbering.levels", NUM_STYLES[val])
                set_path(profile, "heading_numbering.format", None)
            continue
        if qid in ("fig_caption", "tbl_caption", "eq_caption"):
            kind = {"fig_caption": "fig", "tbl_caption": "tbl", "eq_caption": "eq"}[qid]
            tpl, chap = CAPTION_PRESETS.get(val, ("{n}", False))
            set_path(profile, "captions." + kind, tpl)
            if chap:
                chapter_flag = True
            continue
        if qid == "ref_style":
            if val.startswith("只出数字"):
                set_path(profile, "refs", {"fig": "{n}", "tbl": "{n}", "eq": "({n})"})
            else:
                set_path(profile, "refs", {"fig": "图 {n}", "tbl": "表 {n}", "eq": "式 ({n})"})
            continue
        if qid == "table_style":
            set_path(profile, "table.preset",
                     {"三线表": "three_line", "全框线": "full_grid",
                      "只有外框": "outline", "无框线": "none"}.get(val, "three_line"))
            continue
        if qid == "table_header_shading":
            set_path(profile, "table.header.shading", val or None)
            continue
        if qid == "toc_levels":
            set_path(profile, "toc.levels", "1-3" if val == "不要目录" else val)
            set_path(profile, "toc.auto", val != "不要目录")
            continue
        if qid == "footer_format":
            if val == "不要页码":
                set_path(profile, "footer.page_number", False)
            else:
                set_path(profile, "footer.page_number", True)
                set_path(profile, "footer.format", val)
            continue
        # —— 封面 / 目录 / 页码 ——
        if qid in ("cover_above", "cover_below"):
            lines = [x.strip() for x in str(val).split("/") if x.strip()]
            set_path(profile, path, lines)
            continue
        if qid == "front_page_fmt":
            set_path(profile, "front_matter.page_num_fmt",
                     {"罗马数字 I II III": "upperRoman",
                      "小写罗马 i ii iii": "lowerRoman",
                      "阿拉伯数字 1 2 3": "decimal",
                      "不显示页码": None}.get(val, "upperRoman"))
            if val == "不显示页码":
                set_path(profile, "front_matter.page_number", False)
            else:
                set_path(profile, "front_matter.page_number", True)
            continue
        if qid == "body_page_fmt":
            set_path(profile, "body_paging.page_num_fmt",
                     {"阿拉伯数字 1 2 3": "decimal",
                      "罗马数字 I II III": "upperRoman"}.get(val, "decimal"))
            continue

        if qid == "body_indent":
            cmap = {"顶格（0 字符）": 0, "顶格": 0, "0": 0,
                    "1 字符": 1, "1": 1, "2 字符": 2, "2": 2}
            ch = float(cmap.get(str(val), val))
            set_path(profile, "styles.BodyText.first_line_indent_chars", ch)
            set_path(profile, "styles.FirstParagraph.first_line_indent_chars", ch)
            continue

        set_path(profile, path, val)

    if chapter_flag:
        set_path(profile, "captions.chapter", True)
    return profile


# ---------- 三种运行模式 ----------

def do_list(questions, base, core_only=False):
    """打印所有问题（agent 读这个去问用户）"""
    group = None
    for q in questions:
        if core_only and not q.get("core"):
            continue
        if q.get("group") != group:
            group = q.get("group")
            print("\n## %s" % group)
        cur = None if q.get("derived") else (
            get_path(base, q["path"]) if q.get("path") else None)
        default = q.get("default")
        if cur is not None and q.get("type") != "choice":
            default = cur
        choices = q.get("choices")
        line = "- [%s]%s %s" % (q["id"], " *" if q.get("core") else "", q["ask"])
        if choices:
            line += "  可选：%s" % " / ".join(choices)
        line += "  默认：%s" % (default if default is not None else "（空）")
        print(line)
    print("\n* 标记的是核心问题（agent 优先问这几个）")
    print("\n把答案写成 JSON：{\"body_font\": \"仿宋\", \"body_size\": \"三号\", ...}")
    print("然后：python3 configure.py --answers answers.json --out profiles/新档案.json")


def do_interactive(questions, base, out):
    answers = {}
    group = None
    print("=" * 60)
    print("交互配置：一路回车即用默认值")
    print("=" * 60)
    for q in questions:
        if q.get("group") != group:
            group = q.get("group")
            print("\n【%s】" % group)
        cur = None if q.get("derived") else (
            get_path(base, q["path"]) if q.get("path") else None)
        default = cur if (cur is not None and q.get("type") != "choice") else q.get("default")
        choices = q.get("choices")
        prompt = "  %s" % q["ask"]
        if choices:
            prompt += "\n    (%s)" % " / ".join(choices)
        prompt += "\n    [%s]: " % (default if default is not None else "")
        try:
            val = input(prompt).strip()
        except EOFError:
            print("\n（没有可读的输入，改用 --answers answers.json 或 --ask）",
                  file=sys.stderr)
            return answers
        if val == "":
            if default is None or default == "":
                continue
            val = default
        answers[q["id"]] = val
    return answers



def do_ask(questions, base, argv):
    """输出 ask_user_question 能直接用的 JSON（按轮次）"""
    md = None
    if "--md" in argv:
        md = argv[argv.index("--md") + 1]
    feats = asklib.scan_md(md)
    rounds = [int(argv[argv.index("--round") + 1])] if "--round" in argv else [1, 2, 3, 4, 5]
    limit = 4
    if "--limit" in argv:
        limit = int(argv[argv.index("--limit") + 1])
    payload = []
    for r in rounds:
        item = asklib.build_ask(questions, base, r, feats, limit)
        if item:
            payload.append(item)
    print(json.dumps({"md": md, "features": sorted(feats), "rounds": payload},
                     ensure_ascii=False, indent=2))
    return 0


def do_from_docx(base, argv, out):
    """从现成的 Word 模板反推档案"""
    src = None
    if "--from-docx" in argv:
        i = argv.index("--from-docx")
        src = argv[i + 1]
    if not src or not os.path.exists(src):
        raise SystemExit("找不到模板文件：%s" % src)
    frag = asklib.from_docx(src)

    profile = copy.deepcopy(base)
    # 只吸收档案里本来就有的样式键，避免把 47 个无关样式全塞进去
    keep_styles = set((base.get("styles") or {}).keys())
    for section, val in frag.items():
        if section == "styles":
            for sid, props in val.items():
                if sid in keep_styles:
                    profile.setdefault("styles", {}).setdefault(sid, {}).update(props)
            continue
        if isinstance(val, dict) and isinstance(profile.get(section), dict):
            profile[section].update(val)
        else:
            profile[section] = val

    if not out:
        out = os.path.join(PROFILE_DIR, os.path.splitext(os.path.basename(src))[0] + ".json")
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    with open(out, "w", encoding="utf-8") as fh:
        json.dump(profile, fh, ensure_ascii=False, indent=2)

    print("✓ 从模板反推出档案：%s" % out)
    if frag.get("page"):
        print("  页面：%s" % frag["page"])
    if frag.get("styles"):
        print("  识别到 %d 个样式的格式（已吸收 %d 个）" % (
            len(frag["styles"]), len(profile.get("styles") or {})))
        for sid in ("Normal", "Heading1", "Heading2", "BodyText",
                    "Caption", "ImageCaption", "TableCaption", "Compact"):
            if sid in frag["styles"]:
                print("    %-14s %s" % (sid, frag["styles"][sid]))
    if frag.get("header"):
        print("  页眉：%s" % frag["header"].get("text"))
    print("\n核对一遍再转：python3 format_report.py %s" % out)
    return 0


def do_diff(argv):
    args = [a for a in argv if not a.startswith("--")]
    if len(args) < 2:
        raise SystemExit("用法：--diff 档案A 档案B")
    a = load_json(resolve_profile(args[0]))
    b = load_json(resolve_profile(args[1]))
    lines = asklib.diff_profiles(a, b)
    if not lines:
        print("两个档案没有差异")
        return 0
    print("%d 处差异（左=档案A，右=档案B）：" % len(lines))
    for line in lines:
        print("  " + line)
    return 0


SAMPLE_MD = r"""---
title: 样式样张
subtitle: 用于确认各项格式
author: 格式校验
date: 2026 年 10 月
abstract: |
  这一页专门用来肉眼确认各部分的格式：标题层级、正文缩进与行距、
  图注、表注、公式编号、脚注、列表、代码块。
keywords: [样张, 排版]
---

# 一级标题

这是「标题后第一段」，应首行缩进两字符；行距、字号与下方正文一致。

这是第二段正文，用来确认段落之间的间距。**这里是加粗**，*这里是斜体*（应不生效），
`这里是行内代码`，另有上标脚注[^1]。

## 二级标题

### 三级标题

#### 四级标题

有序列表：

1. 第一项
2. 第二项

无序列表：

- 项目甲
- 项目乙

> 这是一段引用块，用来说明出处。

正文中的交叉引用：见图 1、表 1、式 (1)。

![示意图](flow.png){#fig:sample width=10cm}

| 环节 | 目标 | 实测 | 说明 |
|:---|---:|---:|:---|
| 数据采集 | 5 min | 4.2 min | 达标 |
| 清洗入库 | 8 min | 6.9 min | 达标 |

: 示例表格 {#tbl:sample}

吞吐上限由下式给出：

::: {#eq:sample}
$$
Q = \left( \sum_{i=1}^{n} \frac{B_i}{T_i} \right) \times \eta
$$
:::

```python
def hello(name):
    return f"代码块应使用等宽字体：{name}"
```

[^1]: 这是脚注内容，应出现在页底，编号格式由档案决定。
"""


def do_sample(base, argv, out_dir):
    """生成一页样式样张，用来肉眼确认格式"""
    import subprocess
    tmp = tempfile.mkdtemp(prefix="md2docx-sample.")
    md = os.path.join(tmp, "样张.md")
    with open(md, "w", encoding="utf-8") as fh:
        fh.write(SAMPLE_MD)
    # 样张里要有一张图（SAMPLE_MD 里引用 examples/flow.png，所以照搬目录结构）
    src_img = os.path.join(HERE, "examples", "flow.png")
    if os.path.exists(src_img):
        os.makedirs(os.path.join(tmp, "examples"), exist_ok=True)
        shutil.copy(src_img, os.path.join(tmp, "examples", "flow.png"))
    name = argv[argv.index("--sample") + 1] if "--sample" in argv and \
        argv.index("--sample") + 1 < len(argv) else None
    prof = resolve_profile(name) if name and not name.startswith("-") else \
        os.path.join(PROFILE_DIR, "technical.json")

    # 样张只关心「各部分长什么样」，把封面/目录/摘要关掉，压到一两页
    sample_cfg = copy.deepcopy(load_json(prof))
    sample_cfg.setdefault("cover", {})["enabled"] = False
    sample_cfg.setdefault("toc", {})["auto"] = False
    sample_cfg.setdefault("abstract", {})["enabled"] = False
    fm = sample_cfg.setdefault("front_matter", {})
    for k in ("figures", "tables"):
        fm.setdefault(k, {})["enabled"] = False
    prof_sample = os.path.join(tmp, "sample-profile.json")
    with open(prof_sample, "w", encoding="utf-8") as fh:
        json.dump(sample_cfg, fh, ensure_ascii=False, indent=2)

    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, "样张_%s.docx" % os.path.basename(prof).replace(".json", ""))
    r = subprocess.run([sys.executable, os.path.join(HERE, "md2docx.py"),
                        md, out, "--profile", prof_sample], capture_output=True, text=True)
    shutil.rmtree(tmp, ignore_errors=True)
    if r.returncode != 0:
        print(r.stderr or r.stdout, file=sys.stderr)
        raise SystemExit("样张生成失败")
    print("✓ 样张：%s" % out)
    print("  打开确认格式，或用 python3 preview.py \"%s\" 导 PDF 看" % out)
    return 0


def main():
    argv = sys.argv[1:]
    if not argv or "-h" in argv or "--help" in argv:
        print(__doc__.strip())
        return 0
    questions = load_json(QUESTIONS)

    base_name = None
    if "--base" in argv:
        base_name = argv[argv.index("--base") + 1]
    out = None
    if "--out" in argv:
        out = argv[argv.index("--out") + 1]

    base_path = resolve_profile(base_name)
    base = copy.deepcopy(load_json(base_path))

    if "--ask" in argv:
        return do_ask(questions, base, argv)

    if "--from-docx" in argv:
        return do_from_docx(base, argv, out)

    if "--diff" in argv:
        return do_diff(argv)

    if "--sample" in argv:
        sample_out = os.getcwd()
        if "--outdir" in argv and argv.index("--outdir") + 1 < len(argv):
            sample_out = os.path.abspath(argv[argv.index("--outdir") + 1])
        return do_sample(base, argv, sample_out)

    if "--list" in argv:
        do_list(questions, base, core_only="--core" in argv)
        return 0

    if "--show" in argv:
        import format_report
        tgt = argv[argv.index("--show") + 1] if (
            argv.index("--show") + 1 < len(argv) and
            not argv[argv.index("--show") + 1].startswith("-")) else None
        print(format_report.report(tgt or out or base_path, None))
        return 0

    # --set 路径=值（可重复），直接改 profile 里的一个键
    sets = []
    i = 0
    while "--set" in argv[i:]:
        j = argv.index("--set", i)
        sets.append(argv[j + 1])
        i = j + 2

    if "--answers" in argv:
        ap = argv[argv.index("--answers") + 1]
        answers = load_json(ap)
    elif sets:
        answers = {}          # 只改几个键，不进交互
    else:
        answers = do_interactive(questions, base, out)

    profile = apply_answers(base, answers, questions)

    for item in sets:
        if "=" not in item:
            raise SystemExit("--set 需要 路径=值 的形式，收到：%s" % item)
        k, v = item.split("=", 1)
        v = v.strip()
        try:                      # 先当 JSON 解，数字/布尔/数组/对象都支持
            v = json.loads(v)
        except ValueError:
            low = v.lower()
            if low in ("true", "false"):
                v = (low == "true")
            elif low == "null":
                v = None
            else:                 # 其余按字符串（如 宋体、F6F6F6、1-2）
                pass
        set_path(profile, k.strip(), v)
        print("  · %s = %r" % (k.strip(), v))

    def clean(o):                     # 去掉值为 None 的键，档案更干净
        if isinstance(o, dict):
            return {k: clean(v) for k, v in o.items() if v is not None}
        if isinstance(o, list):
            return [clean(x) for x in o]
        return o
    profile = clean(profile)
    profile["name"] = answers.get("name") or base.get("name") or "自定义"

    if not out:
        out = os.path.join(PROFILE_DIR, "custom.json")
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    with open(out, "w", encoding="utf-8") as fh:
        json.dump(profile, fh, ensure_ascii=False, indent=2)
    print("\n✓ 已写入 %s" % out)

    # 把改完之后的格式念一遍，用户能立刻看出改了什么
    try:
        import format_report
        print()
        print(format_report.report(out, None))
    except Exception as e:            # 报告失败不影响写盘
        print("（格式清单生成失败：%s）" % e, file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
