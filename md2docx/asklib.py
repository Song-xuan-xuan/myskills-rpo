#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
asklib.py —— 「交互式定格式」的底层工具

配合 configure.py 使用，四件事：

1. scan_md(md)          扫 md，判断这份文档到底用得上哪些配置项
                        （没有图就不问图注，没有参考文献就不问 CSL 风格）
2. build_ask(...)       把问题翻译成 ask_user_question 能直接吃的 JSON
3. from_docx(docx)      从一份现成的 Word 模板反推 profile（读样式表）
4. diff_profiles(a,b)   对比两个档案，说明改了什么
"""
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WQ = lambda t: "{%s}%s" % (W, t)

# ---------------------------------------------------------------- 分组 / 轮次

# 每轮最多问 3~4 个，符合「先骨架、再正文、再图表、最后页面」的讨论节奏
ROUND_BY_GROUP = {
    "起点": 1, "标题编号": 1, "封面": 1,
    "正文": 2, "标题": 2,
    "图表": 3, "表格": 3,
    "页面": 4, "页眉页脚": 4, "页码": 4, "目录": 4,
    "摘要": 5, "脚注": 5, "列表": 5, "参考文献": 5, "文档": 5, "其它": 5,
    "附录": 5,
}

ROUND_TITLES = {
    1: "先定骨架：从哪套起步、标题怎么编号、要不要封面",
    2: "正文与标题的字号行距",
    3: "图、表、公式的编号与外观",
    4: "页面、页眉页脚、目录、页码",
    5: "补充：摘要、脚注、列表、参考文献",
}

# 这份 md 里没有对应内容，就不必问
REQUIRES = {
    "appendix_enable": "sections", "appendix_label": "sections",
    "logo_path": "always", "logo_width": "always",
    "fig_caption": "figures", "figures_toc": "figures", "figures_toc_title": "figures",
    "tbl_caption": "tables", "tables_toc": "tables", "tables_toc_title": "tables",
    "table_style": "tables", "table_header_shading": "tables",
    "table_repeat_header": "tables", "table_cant_split": "tables",
    "eq_caption": "equations", "ref_style": "any_ref",
    "caption_bold": "any_caption", "caption_font": "any_caption",
    "caption_size": "any_caption",
    "sec_ref_style": "sections", "header_use_chapter": "sections",
    "csl_style": "citations",
    "abstract_enable": "abstract", "abstract_title": "abstract",
    "keywords_title": "abstract",
    "footnote_pos": "footnotes", "footnote_num": "footnotes",
    "footnote_restart": "footnotes",
    "list_ordered": "lists", "list_bullet": "lists",
    "code_font": "code", "code_size": "code", "code_shading": "code",
    "drop_cap": "body_text",
    "grid_lines": "always", "margin_top": "always", "margin_bottom": "always",
    "margin_left": "always", "margin_right": "always", "paper": "always",
}


def scan_md(md_path):
    """扫一份 md，返回它用到了哪些元素（用来筛问题）"""
    feats = set(["always"])
    if not md_path or not os.path.exists(md_path):
        return feats | {"figures", "tables", "equations", "sections", "body_text",
                        "any_caption", "any_ref", "lists", "citations",
                        "abstract", "footnotes", "code"}
    text = open(md_path, encoding="utf-8").read()
    head = text.split("---")[1] if text.startswith("---") else ""

    if re.search(r"!\[[^\]]*\]\([^)]+\)", text) or "![@" in text:
        feats.add("figures")
    if re.search(r"^\s*\|.*\|\s*$", text, re.M) and re.search(r"^\s*\|[\s:|-]+\|\s*$", text, re.M):
        feats.add("tables")
    if "$$" in text or re.search(r"::: \{#eq:", text):
        feats.add("equations")
    if re.search(r"^#{1,6} ", text, re.M):
        feats.add("sections")
    if re.search(r"\{\{(?:fig|tbl|eq|sec):", text):
        feats.add("any_ref")
    if "figures" in feats or "tables" in feats or "equations" in feats:
        feats.add("any_caption")
    if re.search(r"^\s*([-*+]|\d+\.)\s+", text, re.M):
        feats.add("lists")
    if "bibliography:" in head or re.search(r"(^|[^\w])@[A-Za-z]", text):
        feats.add("citations")
    if re.search(r"^abstract:", head, re.M):
        feats.add("abstract")
    if re.search(r"\[\^[^\]]+\]", text):
        feats.add("footnotes")
    if "```" in text:
        feats.add("code")
    if re.search(r"^[^#\s>|*\-+]", text, re.M):
        feats.add("body_text")
    return feats


def question_applies(q, feats):
    need = REQUIRES.get(q["id"])
    if need and need not in feats:
        return False
    return True


def questions_for_round(questions, base, round_no, feats):
    out = []
    for q in questions:
        if ROUND_BY_GROUP.get(q.get("group"), 9) != round_no:
            continue
        if not question_applies(q, feats):
            continue
        out.append(q)
    return out


# ---------------------------------------------------------------- 结构化提问

# 内部取值 -> 人话（不能把 technical/official 这种内部标识暴露给用户）
VALUE_LABELS = {
    # 档案名
    "technical": "技术报告",
    "official": "公文",
    "academic": "学位论文",
    # 常见枚举值
    "pageBottom": "页底", "beneathText": "正文下方",
    "decimal": "阿拉伯数字", "chineseCounting": "中文数字",
    "upperRoman": "大写罗马数字", "lowerRoman": "小写罗马数字",
    "continuous": "全文连续", "eachPage": "每页重排", "eachSect": "每节重排",
    "three_line": "三线表", "full_grid": "全框线", "outline": "只有外框",
    "upperRoman ": "大写罗马数字",
}

# 中文选项 -> 内部值（用户看到中文，落库时换成配置值）
CHOICE_VALUES = {
    "body_indent": {"顶格（0 字符）": 0, "1 字符": 1, "2 字符": 2},
    "grid_lines": {"不限制": "", "每页 22 行": "22", "每页 28 行": "28"},
    "table_header_shading": {"无底色": "", "浅灰 F2F2F2": "F2F2F2",
                             "浅蓝 D9E2F3": "D9E2F3"},
    "code_shading": {"浅灰 F6F6F6": "F6F6F6", "无底色": "",
                     "浅蓝 F2F7FB": "F2F7FB"},
}

# 这些问题的「推荐值」就是档案里的当前值，不用另外标
NO_RECOMMEND = {"base"}


def _configure_module():
    """按需加载 configure（它 import 了本模块，直接顶部 import 会循环）"""
    import importlib.util
    path = os.path.join(HERE, "configure.py")
    spec = importlib.util.spec_from_file_location("md2docx_configure", path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def choice_is_current(qid, choice, base, questions):
    """这个选项选下去，是不是等于档案现在的样子

    比起维护第二张「选项→配置值」表，直接把选项应用一遍再比对更可靠。
    """
    import copy
    try:
        cf = _configure_module()
        probe = copy.deepcopy(base)
        cf.apply_answers(probe, {qid: choice}, questions)
        return not diff_profiles(base, probe)
    except Exception:
        return False


def label_of(value):
    """把内部值翻译成人话；翻不出来就原样返回"""
    if value is None or value == "":
        return ""
    if isinstance(value, bool):
        return "要" if value else "不要"
    s = str(value)
    return VALUE_LABELS.get(s, s)


def _current_value(base, q):
    if q.get("derived"):
        return None          # 需要换算的项（纸张/行距/编号样式）不显示原始值
    path = q.get("path")
    if not path:
        return None
    cur = base
    for part in path.split("."):
        if not isinstance(cur, dict) or part not in cur:
            return None
        cur = cur[part]
    return cur


def build_ask(questions, base, round_no, feats, limit=4):
    """生成 ask_user_question 直接能用的 JSON

    原则（来自实测反馈）：
      - 只问「有明确选项」的项：布尔和枚举。数值/自由文本不进问答题，
        否则会给出「按当前值 / 不设置」这种没法答的选项。
      - 选项用中文，不暴露 technical / pageBottom 这类内部值。
      - 推荐项排第一，标签带「（推荐）」。
    """
    picked = []
    for q in questions_for_round(questions, base, round_no, feats):
        if q.get("skip_in_rounds"):
            continue
        if not q.get("choices") and q.get("type") != "bool":
            continue
        picked.append(q)

    if not picked:
        return None
    picked = picked[:limit]

    out = []
    for q in picked:
        cur = _current_value(base, q)
        default = q.get("default")
        opts = []
        if q.get("choices"):
            for c in q["choices"]:
                is_cur = (str(cur) == str(c)) if cur is not None else False
                if not is_cur and not q.get("derived") and q["id"] != "base":
                    is_cur = choice_is_current(q["id"], c, base, questions)
                is_def = (c == default)
                name = c
                if q["id"] == "base":
                    name = VALUE_LABELS.get(c, c)
                label = name
                if is_def:
                    label = "%s（推荐）" % name
                elif is_cur:
                    label = "%s（当前）" % name
                if q["id"] == "base":
                    desc = "%s：按「%s」这套内置格式起步" % (
                        "推荐" if is_def else ("当前" if is_cur else "可选"), name)
                else:
                    desc = "选「%s」" % name
                    if is_def:
                        desc = "推荐：%s" % name
                    elif is_cur:
                        desc = "当前设置：%s" % name
                opts.append({"label": label, "description": desc})
            # 推荐项排第一
            opts.sort(key=lambda o: 0 if "（推荐）" in o["label"] else
                      (1 if "（当前）" in o["label"] else 2))
        else:  # bool
            want = bool(cur if cur is not None else default)
            yes = {"label": "要%s" % ("（推荐）" if want else ""),
                   "description": "启用" if want else "可以打开"}
            no = {"label": "不要%s" % ("" if want else "（推荐）"),
                  "description": "保持关闭" if not want else "可以关掉"}
            opts = [yes, no] if want else [no, yes]

        item = {
            "id": q["id"],
            "header": (q.get("group") or q["id"])[:12],
            "question": q["ask"],
            "options": opts[:4],
        }
        if q["id"] in CHOICE_VALUES:
            item["value_map"] = CHOICE_VALUES[q["id"]]
        out.append(item)

    return {
        "round": round_no,
        "title": ROUND_TITLES.get(round_no, ""),
        "current": {q["id"]: label_of(_current_value(base, q)) for q in picked},
        "questions": out,
    }


def answer_to_value(qid, answer, questions, base):
    """把用户选中的中文选项翻回配置值"""
    for q in questions:
        if q["id"] != qid:
            continue
        if qid in CHOICE_VALUES:
            return CHOICE_VALUES[qid].get(answer, answer)
        if qid == "base":
            for k, v in VALUE_LABELS.items():
                if answer.startswith(k) or answer.startswith(v):
                    return k
            return answer
        return answer
    return answer


# ---------------------------------------------------------------- 从 Word 模板反推

CN_SIZE = {42: "初号", 36: "小初", 26: "一号", 24: "小一", 22: "二号",
           18: "小二", 16: "三号", 15: "小三", 14: "四号", 12: "小四",
           10.5: "五号", 9: "小五", 7.5: "六号", 6.5: "小六", 5.5: "七号"}


def _style_props(el):
    """从 <w:style> 里抽出我们关心的属性"""
    out = {}
    rPr = el.find(WQ("rPr"))
    if rPr is not None:
        rf = rPr.find(WQ("rFonts"))
        if rf is not None:
            ea = rf.get(WQ("eastAsia"))
            asc = rf.get(WQ("ascii"))
            if ea and ea not in ("Times New Roman", "Calibri", "Arial"):
                out["east_asia_font"] = ea
            if asc and asc not in ("Times New Roman", "Calibri", "Arial"):
                out["ascii_font"] = asc
        sz = rPr.find(WQ("sz"))
        if sz is not None and sz.get(WQ("val")):
            pt = int(sz.get(WQ("val"))) / 2.0
            out["size_pt"] = CN_SIZE.get(pt, pt)
        b = rPr.find(WQ("b"))
        if b is not None:
            out["bold"] = b.get(WQ("val")) not in ("0", "false")
        c = rPr.find(WQ("color"))
        if c is not None:
            val = c.get(WQ("val"))
            if val and val.upper() not in ("000000", "AUTO"):
                out["color"] = val
    pPr = el.find(WQ("pPr"))
    if pPr is not None:
        jc = pPr.find(WQ("jc"))
        if jc is not None:
            m = {"left": "left", "center": "center", "right": "right",
                 "both": "justify"}
            if jc.get(WQ("val")) in m:
                out["align"] = m[jc.get(WQ("val"))]
        ind = pPr.find(WQ("ind"))
        if ind is not None:
            if ind.get(WQ("firstLineChars")) is not None:
                out["first_line_indent_chars"] = int(ind.get(WQ("firstLineChars"))) / 100.0
            elif ind.get(WQ("firstLine")) is not None:
                out["first_line_indent_chars"] = round(
                    int(ind.get(WQ("firstLine"))) / 20.0 / (out.get("size_pt") or 12)
                    if isinstance(out.get("size_pt"), (int, float)) else 0, 2)
            if ind.get(WQ("leftChars")) is not None:
                out["left_indent_chars"] = int(ind.get(WQ("leftChars"))) / 100.0
        sp = pPr.find(WQ("spacing"))
        if sp is not None:
            line = sp.get(WQ("line"))
            if line:
                if sp.get(WQ("lineRule")) == "exact":
                    out["line_spacing_pt"] = int(line) / 20.0
                else:
                    out["line_spacing"] = round(int(line) / 240.0, 2)
            for k, key in (("before", "space_before_pt"), ("after", "space_after_pt")):
                if sp.get(WQ(k)) is not None:
                    out[key] = int(sp.get(WQ(k))) / 20.0
        if pPr.find(WQ("keepNext")) is not None:
            out["keep_next"] = True
    return out


def from_docx(docx_path):
    """读一份 Word 文档/模板，反推出 profile 片段"""
    z = zipfile.ZipFile(docx_path)
    frag = {}
    styles = {}
    if "word/styles.xml" in z.namelist():
        root = ET.fromstring(z.read("word/styles.xml"))
        for st in root.findall(WQ("style")):
            sid = st.get(WQ("styleId"))
            if st.get(WQ("type")) != "paragraph":
                continue
            props = _style_props(st)
            if props:
                styles[sid] = props
    if styles:
        frag["styles"] = styles

    # 页面
    if "word/document.xml" in z.namelist():
        doc = ET.fromstring(z.read("word/document.xml"))
        sectPr = doc.find(".//" + WQ("sectPr"))
        if sectPr is not None:
            page = {}
            pgSz = sectPr.find(WQ("pgSz"))
            if pgSz is not None:
                w = int(pgSz.get(WQ("w"), 0)) / 567.0
                h = int(pgSz.get(WQ("h"), 0)) / 567.0
                if w and h:
                    page["width_cm"] = round(w, 2)
                    page["height_cm"] = round(h, 2)
            pgMar = sectPr.find(WQ("pgMar"))
            if pgMar is not None:
                m = {}
                for k in ("top", "bottom", "left", "right"):
                    v = pgMar.get(WQ(k))
                    if v is not None:
                        m[k] = round(int(v) / 567.0, 2)
                if m:
                    page["margin_cm"] = m
            if page:
                frag["page"] = page

    # 页眉页脚文字
    for name in z.namelist():
        if re.match(r"word/header\d*\.xml$", name):
            root = ET.fromstring(z.read(name))
            text = "".join(t.text or "" for t in root.iter(WQ("t"))).strip()
            if text:
                frag.setdefault("header", {})["text"] = text
                break
    for name in sorted(z.namelist()):
        if not re.match(r"word/footer\d*\.xml$", name):
            continue
        root = ET.fromstring(z.read(name))
        # 按出现顺序把页码域换成占位符，普通文字原样保留。
        # 注意：域在 XML 里是 begin / instrText / separate / 缓存值 / end 五段，
        # 缓存值那一段要跳过，否则会拼出「第 {PAGE}1 页」这种怪东西。
        parts = []
        in_field = False
        for r in root.iter(WQ("r")):
            fld = r.find(WQ("fldChar"))
            if fld is not None:
                t = fld.get(WQ("fldCharType"))
                if t == "begin":
                    in_field = True
                elif t == "end":
                    in_field = False
                continue
            instr = r.find(WQ("instrText"))
            if instr is not None and instr.text:
                code = instr.text.strip().upper()
                if "NUMPAGES" in code or "SECTIONPAGES" in code:
                    parts.append("{NUMPAGES}")
                elif "PAGE" in code:
                    parts.append("{PAGE}")
                continue
            if in_field:          # 域的缓存值，跳过
                continue
            for t in r.findall(WQ("t")):
                if t.text:
                    parts.append(t.text)
        fmt = "".join(parts).strip()
        if fmt:
            frag.setdefault("footer", {})["format"] = fmt
            break
    return frag


# ---------------------------------------------------------------- 档案对比

def flatten(d, prefix=""):
    out = {}
    for k, v in (d or {}).items():
        if isinstance(v, dict):
            out.update(flatten(v, prefix + k + "."))
        else:
            out[prefix + k] = v
    return out


def diff_profiles(a, b):
    fa, fb = flatten(a), flatten(b)
    lines = []
    for k in sorted(set(fa) | set(fb)):
        if k.startswith("_") or k.endswith(".name"):
            continue
        va, vb = fa.get(k), fb.get(k)
        if va != vb:
            lines.append("%-42s %r -> %r" % (k, va, vb))
    return lines
