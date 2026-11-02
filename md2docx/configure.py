#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
configure.py —— 生成 / 修改 / 查看格式档案（profile）

    python3 configure.py --list                 # 全部问题（72 条）
    python3 configure.py --list --core          # 只列核心问题
    python3 configure.py                        # 终端里逐条问（回车用默认值）
    python3 configure.py --ask --md 报告.md --round 1   # 给 agent 的结构化提问 JSON
    python3 configure.py --answers a.json --out 客户A.json
    python3 configure.py --from-docx 单位模板.docx      # 从 Word 模板反推
    python3 configure.py --sample technical     # 出一页样张
    python3 configure.py --diff a.json b.json   # 两个档案差在哪
    python3 configure.py --report technical      # 把档案翻译成中文格式清单
    python3 configure.py --set styles.Caption.bold=true --base x.json --out x.json
"""
import copy
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from _common import (HERE, PROFILE_DIR, DEFAULT_PROFILE, die,
                     strip_json_comments, resolve_profile, load_profile)

QUESTIONS = os.path.join(HERE, "questions.json")


def _load_json(path):
    return load_profile(path)


def _resolve_path(arg):
    """档案名/路径 -> 绝对路径（找不到就报错）"""
    name, path = resolve_profile(arg)
    if not path:
        die("找不到 profile：%s" % arg)
    return path


# ============================================================
# ① 读 md：先看这份文档用得上哪些问题
# ============================================================

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

# ============================================================
# ② 把档案翻译成人话
# ============================================================

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
sys.path.insert(0, HERE)





def build_template(profile_path, base_docx=None):
    tmp = tempfile.mkdtemp(prefix="fmt-report.")
    out = os.path.join(tmp, "tpl.docx")
    cmd = [sys.executable, os.path.join(HERE, "build_template.py"),
           "-p", os.path.abspath(profile_path), "-o", out]
    if base_docx:
        cmd += ["--base", base_docx]
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        raise SystemExit("模板编译失败：\n" + (r.stderr or r.stdout))
    return out


class Styles:
    """解析 styles.xml，按 basedOn 继承链算出每个样式的最终效果"""

    def __init__(self, docx_path):
        import zipfile
        z = zipfile.ZipFile(docx_path)
        self.root = ET.fromstring(z.read("word/styles.xml"))
        self.styles = {}
        for st in self.root.findall(W + "style"):
            sid = st.get(W + "styleId")
            based = st.find(W + "basedOn")
            self.styles[sid] = {
                "node": st,
                "basedOn": based.get(W + "val") if based is not None else None,
            }
        dd = self.root.find(W + "docDefaults")
        self.defaults = {}
        if dd is not None:
            rpr = dd.find(W + "rPrDefault/" + W + "rPr")
            ppr = dd.find(W + "pPrDefault/" + W + "pPr")
            if rpr is not None:
                self.defaults.update(self._rpr(rpr))
            if ppr is not None:
                self.defaults.update(self._ppr(ppr))

    @staticmethod
    def _rpr(rpr):
        d = {}
        f = rpr.find(W + "rFonts")
        if f is not None:
            if f.get(W + "eastAsia"):
                d["font_ea"] = f.get(W + "eastAsia")
            if f.get(W + "ascii"):
                d["font_ascii"] = f.get(W + "ascii")
        sz = rpr.find(W + "sz")
        if sz is not None and sz.get(W + "val"):
            d["size"] = int(sz.get(W + "val")) / 2.0
        b = rpr.find(W + "b")
        if b is not None:
            d["bold"] = b.get(W + "val") not in ("0", "false")
        i = rpr.find(W + "i")
        if i is not None:
            d["italic"] = i.get(W + "val") not in ("0", "false")
        c = rpr.find(W + "color")
        if c is not None and c.get(W + "val"):
            d["color"] = c.get(W + "val")
        return d

    @staticmethod
    def _ppr(ppr):
        d = {}
        jc = ppr.find(W + "jc")
        if jc is not None:
            d["align"] = jc.get(W + "val")
        ind = ppr.find(W + "ind")
        if ind is not None:
            for src, dst in (("firstLineChars", "indent_first"),
                             ("leftChars", "indent_left"),
                             ("rightChars", "indent_right"),
                             ("hangingChars", "indent_hanging")):
                if ind.get(W + src) is not None:
                    d[dst] = int(ind.get(W + src)) / 100.0
            if ind.get(W + "firstLine") is not None and "indent_first" not in d:
                d["indent_first"] = int(ind.get(W + "firstLine")) / 20.0
        sp = ppr.find(W + "spacing")
        if sp is not None:
            if sp.get(W + "line"):
                rule = sp.get(W + "lineRule", "auto")
                val = int(sp.get(W + "line"))
                d["line"] = val / 240.0 if rule == "auto" else val / 20.0
                d["line_rule"] = rule
            if sp.get(W + "before") is not None:
                d["before"] = int(sp.get(W + "before")) / 20.0
            if sp.get(W + "after") is not None:
                d["after"] = int(sp.get(W + "after")) / 20.0
        if ppr.find(W + "keepNext") is not None:
            d["keep_next"] = True
        if ppr.find(W + "pageBreakBefore") is not None:
            d["page_break"] = True
        shd = ppr.find(W + "shd")
        if shd is not None and shd.get(W + "fill") not in (None, "auto"):
            d["shading"] = shd.get(W + "fill")
        return d

    def effective(self, sid):
        chain, cur = [], sid
        while cur and cur in self.styles and cur not in chain:
            chain.append(cur)
            cur = self.styles[cur]["basedOn"]
        out = dict(self.defaults)
        for s in reversed(chain):          # 从最根往当前合并
            node = self.styles[s]["node"]
            rpr = node.find(W + "rPr")
            if rpr is not None:
                out.update(self._rpr(rpr))
            ppr = node.find(W + "pPr")
            if ppr is not None:
                out.update(self._ppr(ppr))
        return out


# ---------------- 描述成中文 ----------------

CN_SIZE = {42: "初号", 36: "小初", 26: "一号", 24: "小一", 22: "二号",
           18: "小二", 16: "三号", 15: "小三", 14: "四号", 12: "小四",
           10.5: "五号", 9: "小五", 7.5: "六号", 6.5: "小六", 5.5: "七号"}

ALIGN = {"left": "左对齐", "center": "居中", "right": "右对齐", "both": "两端对齐"}


def size_text(v):
    if v is None:
        return ""
    cn = CN_SIZE.get(round(v, 1))
    return "%s(%g磅)" % (cn, v) if cn else "%g磅" % v


def line_text(eff):
    if "line" not in eff:
        return ""
    if eff.get("line_rule") == "auto":
        return "行距%g倍" % eff["line"]
    return "固定行距%g磅" % eff["line"]


def indent_text(eff):
    bits = []
    if eff.get("indent_first") is not None:
        bits.append("首行缩进%g字符" % eff["indent_first"] if eff["indent_first"]
                    else "顶格")
    if eff.get("indent_left"):
        bits.append("左缩进%g字符" % eff["indent_left"])
    if eff.get("indent_right"):
        bits.append("右缩进%g字符" % eff["indent_right"])
    return " ".join(bits)


def para_bits(eff):
    bits = []
    if eff.get("font_ea"):
        bits.append(eff["font_ea"])
    if eff.get("font_ascii") and eff.get("font_ascii") != eff.get("font_ea"):
        bits.append(eff["font_ascii"] + "(西文)")
    if eff.get("size"):
        bits.append(size_text(eff["size"]))
    if eff.get("bold"):
        bits.append("加粗")
    if eff.get("align"):
        bits.append(ALIGN.get(eff["align"], eff["align"]))
    t = indent_text(eff)
    if t:
        bits.append(t)
    t = line_text(eff)
    if t:
        bits.append(t)
    if eff.get("before") or eff.get("after"):
        bits.append("段前%g 段后%g" % (eff.get("before", 0), eff.get("after", 0)))
    if eff.get("shading"):
        bits.append("底色#" + eff["shading"])
    if eff.get("keep_next"):
        bits.append("与下段同页")
    if eff.get("page_break"):
        bits.append("段前分页")
    return " · ".join(bits)


SAMPLE = {"decimal": "1", "chineseCounting": "一",
          "chineseCountingThousand": "一", "upperRoman": "I",
          "lowerRoman": "i", "upperLetter": "A", "lowerLetter": "a"}


def num_text(hn, depth=3):
    """把 lvlText 模板渲染成样例：%1.%2 -> 1.1"""
    if not hn.get("enabled"):
        return "不自动编号"
    levels = hn.get("levels") or []
    samples = []
    out = []
    for lc in levels[:depth]:
        lc = lc or {}
        nf = lc.get("num_fmt", "decimal")
        cur = SAMPLE.get(nf, "1")
        samples.append(cur)
        f = lc.get("format", "") or ".".join("%%%d" % (i + 1) for i in range(len(samples)))
        for i in range(len(samples)):
            f = f.replace("%%%d" % (i + 1), samples[i])
        out.append(f)
    return " / ".join(out)


def report(profile_path, base_docx=None):
    profile_path = _resolve_path(profile_path)
    cfg = load_profile(profile_path)
    tpl = build_template(profile_path, base_docx)
    st = Styles(tpl)

    out = []
    A = out.append
    A("格式清单（profile: %s）" % os.path.basename(profile_path).replace(".json", ""))
    A("=" * 58)

    # 页面
    p = cfg.get("page") or {}
    m = p.get("margin_cm") or {}
    A("页面　　%s %.0f×%.0fmm · 边距 上%g 下%g 左%g 右%g cm" % (
        "A4" if abs(p.get("width_cm", 21) - 21) < 0.2 else "自定义",
        p.get("width_cm", 21) * 10, p.get("height_cm", 29.7) * 10,
        m.get("top", 2.54), m.get("bottom", 2.54),
        m.get("left", 3.17), m.get("right", 2.54)))
    if p.get("first_page_different"):
        A("　　　　首页不同（封面不显示页眉页脚）")

    # 正文
    body = st.effective("BodyText")
    fp = st.effective("FirstParagraph")
    A("正文　　%s" % para_bits(body))
    if para_bits(fp) != para_bits(body):
        A("　　　　（标题后首段：%s）" % para_bits(fp))

    # 标题
    hn = cfg.get("heading_numbering") or {}
    ntext = num_text(hn, 4)
    for i in range(1, 7):
        sid = "Heading%d" % i
        if sid not in st.styles:
            continue
        eff = st.effective(sid)
        extra = ""
        if hn.get("enabled") and i <= len(hn.get("levels") or []):
            extra = " · 编号「%s」" % ntext.split(" / ")[i - 1]
        A("标题%d　　%s%s" % (i, para_bits(eff), extra))

    # 题注
    caps = cfg.get("captions") or {}
    refs = cfg.get("refs") or {}
    cap = st.effective("Caption")
    A("图题注　%s · 编号「%s」 · 正文引用「%s」" % (
        para_bits(cap), caps.get("fig", "图 {n}").replace("{n}", "1"),
        refs.get("fig", "{n}").replace("{n}", "1")))
    A("表题注　%s · 编号「%s」 · 正文引用「%s」" % (
        para_bits(st.effective("TableCaption")),
        caps.get("tbl", "表 {n}").replace("{n}", "1"),
        refs.get("tbl", "{n}").replace("{n}", "1")))
    A("公式　　编号「%s」 · 正文引用「%s」 · 居中，编号靠右" % (
        caps.get("eq", "({n})").replace("{n}", "1"),
        refs.get("eq", "({n})").replace("{n}", "1")))

    # 表格
    t = cfg.get("table") or {}
    preset = {"three_line": "三线表", "full_grid": "全框线",
              "outline": "只有外框", "none": "无框线"}.get(t.get("preset"), "三线表")
    th = t.get("header") or {}
    bits = [preset]
    if th.get("bold", True):
        bits.append("表头加粗")
    if th.get("repeat", True):
        bits.append("跨页重复表头")
    if th.get("shading"):
        bits.append("表头底色#" + th["shading"])
    if t.get("full_width", True):
        bits.append("满栏")
    if t.get("cant_split"):
        bits.append("行不跨页")
    if t.get("col_widths"):
        bits.append("自定义列宽 %d 张表" % len(t["col_widths"]))
    A("交叉引用 %s" % " · ".join(
        "%s→%s" % (k, v.replace("{n}", "1").replace("{t}", "标题"))
        for k, v in sorted((cfg.get("refs") or {}).items())))
    bw = t.get("borders") or {}
    thick = (bw.get("top") or {}).get("sz", 12)
    bits.append("外线%g磅 表头线%g磅" % (thick / 8.0,
                                        ((bw.get("header_bottom") or {}).get("sz", 6)) / 8.0))
    A("表格　　%s · 单元格 %s" % (
        " · ".join(bits),
        {"center": "垂直居中", "top": "顶端对齐",
         "bottom": "底端对齐"}.get((t.get("cell") or {}).get("valign", "center"), "垂直居中")))
    A("表内容　%s" % para_bits(st.effective("TableText")
                              if "TableText" in st.styles else st.effective("Compact")))
    if "Compact" in st.styles:
        A("列表项　%s" % para_bits(st.effective("Compact")))

    # 代码 / 引用
    A("代码块　%s" % para_bits(st.effective("SourceCode")))
    if "Hyperlink" in st.styles:
        hl = st.effective("Hyperlink")
        A("超链接　颜色#%s%s" % (hl.get("color", "自动"),
                                " · 下划线" if hl.get("underline") else ""))
    A("引用块　%s" % para_bits(st.effective("BlockText")))

    # 页眉页脚
    h = cfg.get("header") or {}
    f = cfg.get("footer") or {}
    hdr_bits = []
    if h.get("use_chapter"):
        hdr_bits.append("显示当前章节标题")
    if h.get("even_text"):
        hdr_bits.append("偶数页「%s」" % h["even_text"])

    def _size(v):
        if v is None:
            return ""
        if isinstance(v, (int, float)):
            return size_text(float(v))
        for k, pt in CN_SIZE.items():
            if str(v) == k:
                return size_text(pt)
        try:
            return size_text(float(v))
        except (TypeError, ValueError):
            return str(v)

    A("页眉　　%s%s%s" % (h.get("text") or "（无）",
                          (" · " + _size(h.get("size_pt"))) if h.get("size_pt") else "",
                          (" · " + " / ".join(hdr_bits)) if hdr_bits else ""))
    A("页脚　　%s%s" % (f.get("format") or "（无）",
                        (" · " + _size(f.get("size_pt"))) if f.get("size_pt") else ""))

    # 附录
    ap = cfg.get("appendix") or {}
    if ap.get("enabled"):
        A("附录　　一级「%s A」· 二级「A.1」· 起点锚点 {%s} · 结束 {%s}" % (
            ap.get("label", "附录"), ap.get("marker", "sec:app"),
            ap.get("end_marker", "sec:endappendix")))
    if (cfg.get("cover") or {}).get("logo"):
        A("封面 logo %s（宽 %gcm）" % ((cfg["cover"])["logo"],
                                     (cfg["cover"]).get("logo_width_cm", 4)))

    # 摘要 / 脚注 / 列表 / 文档
    abst = cfg.get("abstract") or {}
    A("摘要　　%s · 标题「%s」" % (
        "开" if abst.get("enabled", True) else "关",
        abst.get("title", "摘　要")))
    notes = cfg.get("notes") or {}
    fn = notes.get("footnote") or {}
    if fn:
        A("脚注　　位置 %s · 编号 %s · %s重排" % (
            {"pageBottom": "页底", "beneathText": "正文下方"}.get(
                fn.get("pos"), fn.get("pos")),
            {"decimal": "1,2,3", "lowerRoman": "i,ii,iii",
             "upperRoman": "I,II,III", "chineseCounting": "一,二,三"}.get(
                fn.get("num_fmt"), fn.get("num_fmt")),
            {"continuous": "全文连续", "eachPage": "每页",
             "eachSect": "每节"}.get(fn.get("num_restart"), fn.get("num_restart"))))
    lists = cfg.get("lists") or {}
    if lists:
        o = lists.get("ordered") or {}
        A("列表　　有序「%s」 · 无序「%s」" % (
            o.get("format", "%1."), (lists.get("bullet") or {}).get("char", "•")))
    doc_cfg = cfg.get("document") or {}
    if doc_cfg:
        A("文档　　修订模式%s · 保护%s · 制表位%gcm" % (
            "开" if doc_cfg.get("track_changes") else "关",
            doc_cfg.get("protect") or "无",
            doc_cfg.get("default_tab_cm", 1.27)))
    if (cfg.get("drop_cap") or {}).get("enabled"):
        A("　　　　首字下沉 %s 行" % (cfg["drop_cap"].get("lines", 3)))
    grid = (cfg.get("page") or {}).get("doc_grid") or {}
    if grid:
        A("　　　　文档网格 每页%s行" % grid.get("lines_per_page", "-"))

    # 封面
    cov = cfg.get("cover") or {}
    if cov.get("enabled"):
        A("封面　　开 · 标题距顶%gcm · 页码%s · 页眉%s" % (
            cov.get("top_space_cm", 6.5),
            "显示" if cov.get("page_number") else "不显示",
            "显示" if cov.get("header") else "不显示"))
        if cov.get("above"):
            A("　　　　标题上方：%s" % " / ".join(cov["above"]))
        if cov.get("below"):
            A("　　　　标题下方：%s" % " / ".join(cov["below"]))
    else:
        A("封面　　关")

    # 参考文献
    cit = cfg.get("citations") or {}
    if cit.get("csl"):
        csl_name = str(cit["csl"]).replace("chinese-gb7714-", "国标 ")
        A("参考文献 %s（需 md 头部写 bibliography: refs.bib）" % csl_name)
    if cfg.get("landscape"):
        A("横排表　%s" % "、".join(cfg["landscape"]))

    # 目录 / 页码
    toc = cfg.get("toc") or {}
    fm = cfg.get("front_matter") or {}
    bp = cfg.get("body_paging") or {}
    A("目录　　%s · %s级 · 目录页码 %s · 目录后分页%s" % (
        "开" if toc.get("auto", True) else "关",
        toc.get("levels", "1-3"),
        {"upperRoman": "罗马 I II III", "lowerRoman": "小写罗马 i ii iii",
         "decimal": "阿拉伯 1 2 3"}.get(fm.get("page_num_fmt", "upperRoman"),
                                       fm.get("page_num_fmt", "")),
        "是" if toc.get("page_break_after", True) else "否"))
    if toc.get("tabs", True):
        A("　　　　目录条目：层级缩进 + 右对齐点线制表位")
    A("　　　　图目录 %s · 表目录 %s" % (
        "开" if (fm.get("figures") or {}).get("enabled") else "关",
        "开" if (fm.get("tables") or {}).get("enabled") else "关"))
    A("正文页码 %s · 从 %s 开始" % (
        {"decimal": "阿拉伯 1 2 3", "upperRoman": "罗马 I II III"}.get(
            bp.get("page_num_fmt", "decimal"), bp.get("page_num_fmt", "")),
        bp.get("page_num_start", 1)))

    A("=" * 58)
    return "\n".join(out)

# ============================================================
# ③ 问答 / 反推 / 样张 / 对比 / 改键
# ============================================================

QUESTIONS = os.path.join(HERE, "questions.json")
PROFILE_DIR = os.path.join(HERE, "profiles")




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
    feats = scan_md(md)
    rounds = [int(argv[argv.index("--round") + 1])] if "--round" in argv else [1, 2, 3, 4, 5]
    limit = 4
    if "--limit" in argv:
        limit = int(argv[argv.index("--limit") + 1])
    payload = []
    for r in rounds:
        item = build_ask(questions, base, r, feats, limit)
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
    frag = from_docx(src)

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
    print("\n核对一遍再转：python3 configure.py --report %s" % out)
    return 0


def do_diff(argv):
    args = [a for a in argv if not a.startswith("--")]
    if len(args) < 2:
        raise SystemExit("用法：--diff 档案A 档案B")
    a = _load_json(_resolve_path(args[0]))
    b = _load_json(_resolve_path(args[1]))
    lines = diff_profiles(a, b)
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
    prof = _resolve_path(name) if name and not name.startswith("-") else \
        os.path.join(PROFILE_DIR, "technical.json")

    # 样张只关心「各部分长什么样」，把封面/目录/摘要关掉，压到一两页
    sample_cfg = copy.deepcopy(_load_json(prof))
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
    print("  打开确认格式，或 python3 md2docx.py preview \"%s\" 导 PDF 看" % out)
    return 0


def main(argv):
    argv = list(argv)
    if not argv or "-h" in argv or "--help" in argv:
        print(__doc__.strip())
        return 0
    questions = _load_json(QUESTIONS)

    base_name = None
    if "--base" in argv:
        base_name = argv[argv.index("--base") + 1]
    out = None
    if "--out" in argv:
        out = argv[argv.index("--out") + 1]

    base_path = _resolve_path(base_name) if base_name else \
        os.path.join(PROFILE_DIR, DEFAULT_PROFILE + ".json")
    base = copy.deepcopy(_load_json(base_path))

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

    if "--show" in argv or "--report" in argv:
        flag = "--show" if "--show" in argv else "--report"
        i = argv.index(flag)
        tgt = argv[i + 1] if (i + 1 < len(argv) and
                              not argv[i + 1].startswith("-")) else None
        print(report(tgt or out or base_path, None))
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
        answers = _load_json(ap)
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
        print()
        print(report(out, None))
    except Exception as e:            # 报告失败不影响写盘
        print("（格式清单生成失败：%s）" % e, file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]) or 0)
