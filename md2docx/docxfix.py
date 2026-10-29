#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
docxfix.py —— pandoc 产出的 docx 后处理：把静态文字升级成 Word 原生"活"结构

处理 6 件事：
  1. 图 / 表 / 公式题注编号  -> SEQ 域（增删图表后 F9 自动重排）
  2. @fig:x / @tbl:x / @eq:x -> REF 域（交叉引用跟着编号自动变）
  3. {{TOC}} 占位            -> TOC 域（Word 里可更新目录）
  4. 标题 1/2/3              -> Word 多级列表（格式由 profile 决定）
  5. 表格线框 / 表头 / 单元格对齐（由 profile 的 table 段决定）
  6. 图片不溢出、斜体清零、打开时自动刷新域

必须在 pandoc 之后运行，因为 pandoc 会重写 numbering.xml。

用法：
    python3 docxfix.py in.docx out.docx
    python3 docxfix.py in.docx out.docx --profile profiles/official.json
"""
import copy, json, os, re, sys
from docx.shared import Cm

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_PROFILE = os.path.join(HERE, "profiles", "technical.json")


def wq(tag):
    return "{%s}%s" % (W, tag)


RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps",
             "strike", "dstrike", "outline", "shadow", "emboss", "imprint",
             "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing",
             "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
             "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang",
             "eastAsianLayout", "specVanish", "oMath"]

TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
               "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd",
               "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook",
               "tblCaption", "tblDescription"]

PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
             "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
             "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
             "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
             "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
             "suppressOverlap", "jc", "textDirection", "textAlignment",
             "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
             "sectPr", "pPrChange"]

TCPR_ORDER = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd",
              "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark"]

# 段落样式 -> 题注类型（pandoc 固定用这两个样式名）
CAPTION_STYLES = {"ImageCaption": "fig", "TableCaption": "tbl"}


# ---------------- 基础 ----------------

def el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn(k), str(v))
    return e


def trun(s=None):
    r = el("w:r")
    if s is None:
        r.append(el("w:tab"))
    elif not isinstance(s, str):
        r.append(s)
    else:
        t = el("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = s
        r.append(t)
    return r


def field(instr, cached):
    out = []
    r = el("w:r"); r.append(el("w:fldChar", **{"w:fldCharType": "begin"})); out.append(r)
    r = el("w:r"); i = el("w:instrText"); i.set(qn("xml:space"), "preserve")
    i.text = instr; r.append(i); out.append(r)
    r = el("w:r"); r.append(el("w:fldChar", **{"w:fldCharType": "separate"})); out.append(r)
    out.append(trun(str(cached)))
    r = el("w:r"); r.append(el("w:fldChar", **{"w:fldCharType": "end"})); out.append(r)
    return out


def reorder(elm, order):
    idx = {t: i for i, t in enumerate(order)}
    kids = list(elm)
    kids.sort(key=lambda k: idx.get(k.tag.replace(wq(""), ""), 999))
    for k in kids:
        elm.append(k)


def style_id(p):
    ps = p.find(qn("w:pPr") + "/" + qn("w:pStyle"))
    return ps.get(qn("w:val")) if ps is not None else None


def is_math(p):
    return p.tag == qn("w:p") and p.find(M + "oMathPara") is not None


def para_text(p):
    return "".join(t.text or "" for t in p.iter(qn("w:t"))).strip()


def clear_runs(p):
    for r in list(p.findall(qn("w:r"))):
        p.remove(r)
    for x in list(p.findall(qn("w:hyperlink"))):
        p.remove(x)


def text_width_twips(doc):
    s = doc.sections[0]
    return int((s.page_width - s.left_margin - s.right_margin) / 635)


def split_template(tpl):
    """把 "图 {n}" 拆成 ("图 ", "")；把 "(1)" 这种拆成 ("(", ")")"""
    if "{n}" in tpl:
        a, b = tpl.split("{n}", 1)
        return a, b
    return tpl, ""


def split_ref_template(tpl):
    """交叉引用模板：支持 {n}（编号）、{t}（标题文字）、{p}（页码）

    "图 {n}"    -> [("lit", "图 "), ("num", ""), ("lit", "")]
    "第 {n} 节" -> [("lit", "第 "), ("num", ""), ("lit", " 节")]
    "见第 {p} 页" -> [("lit", "见第 "), ("page", ""), ("lit", " 页")]
    """
    out, pos = [], 0
    for m in re.finditer(r"\{([ntp])\}", tpl):
        if m.start() > pos:
            out.append(("lit", tpl[pos:m.start()]))
        out.append(({"n": "num", "t": "txt", "p": "page"}[m.group(1)], ""))
        pos = m.end()
    if pos < len(tpl):
        out.append(("lit", tpl[pos:]))
    return out or [("num", "")]


# ---------------- 1+2. 题注编号 & 交叉引用 ----------------

def set_math_par_format(p, width):
    pPr = p.find(qn("w:pPr"))
    if pPr is None:
        pPr = el("w:pPr"); p.insert(0, pPr)
    for tag in ("tabs", "ind", "jc"):
        old = pPr.find(qn("w:" + tag))
        if old is not None:
            pPr.remove(old)
    tabs = el("w:tabs")
    tabs.append(el("w:tab", **{"w:val": "center", "w:pos": int(width / 2)}))
    tabs.append(el("w:tab", **{"w:val": "right", "w:pos": int(width)}))
    pPr.append(tabs)
    pPr.append(el("w:ind", **{"w:firstLine": "0", "w:firstLineChars": "0"}))
    pPr.append(el("w:jc", **{"w:val": "left"}))
    reorder(pPr, PPR_ORDER)


def number_block(kind, num, chapter, seq_name, bms):
    """生成「编号」这一串 run；chapter=True 时前面用 STYLEREF 取章号 -> 2-1"""
    out = []
    for name, bid, _ in bms:
        out.append(el("w:bookmarkStart", **{"w:id": bid, "w:name": name}))
    if chapter:
        out.extend(field(" STYLEREF 1 \\s ", "1"))
        out.append(trun("-"))
    instr = " SEQ %s \\* ARABIC " % seq_name
    if chapter:
        instr = " SEQ %s \\* ARABIC \\s 1 " % seq_name
    out.extend(field(instr, num))
    for name, bid, _ in bms:
        out.append(el("w:bookmarkEnd", **{"w:id": bid}))
    return out


def upgrade(doc, cfg):
    captions = cfg.get("captions") or {}
    sep = cfg.get("caption_sep", "  ")
    chapter = bool(captions.get("chapter", False))
    seq_name = {"fig": "图", "tbl": "表", "eq": "公式"}

    body = doc.element.body
    width = text_width_twips(doc)
    stack, drop = [], set()
    pending = []          # 还没被认领的书签
    counters, bmnum = {"fig": 0, "tbl": 0, "eq": 0}, {}

    for node in list(body):
        if node.tag == qn("w:bookmarkStart"):
            # pandoc 的结构是：bmStart -> [图片/公式段落] -> 题注段落 -> ... -> bmEnd
            # 标题的书签也在 bmStart 里，但它属于标题自己，不能被题注抢走。
            pending.append((node.get(qn("w:name")), node.get(qn("w:id")), node))
            continue
        if node.tag == qn("w:bookmarkEnd"):
            continue

        kind = None
        if node.tag == qn("w:p"):
            sid = style_id(node)
            if sid in CAPTION_STYLES:
                kind = CAPTION_STYLES[sid]
            elif is_math(node) and captions.get("eq"):
                kind = "eq"
            elif sid and sid.startswith("Heading"):
                # 标题吃掉了它前面的那个书签（书签留在原地，不动）
                if pending:
                    pending.pop()
                continue
        if kind is None:
            continue

        # 题注只认「紧邻它之前的那个书签」
        moved = pending[-1:] if pending else []
        if moved:
            pending.pop()
        ids = {s[1] for s in moved}
        for s in moved:
            if s[2].getparent() is not None:
                s[2].getparent().remove(s[2])

        counters[kind] += 1
        num = counters[kind]
        before, after = split_template(captions.get(kind) or "")

        if kind == "eq":
            omp = node.find(M + "oMathPara")
            omath = omp.find(M + "oMath")
            omp.remove(omath)
            node.remove(omp)
            clear_runs(node)
            set_math_par_format(node, width)
            node.append(trun())
            node.append(omath)
            node.append(trun())
            if before:
                node.append(trun(before))
            node.extend(number_block(kind, num, chapter, seq_name[kind], moved))
            if after:
                node.append(trun(after))
        else:
            text = para_text(node)
            clear_runs(node)
            if before:
                node.append(trun(before))
            node.extend(number_block(kind, num, chapter, seq_name[kind], moved))
            if after:
                node.append(trun(after))
            if text:
                node.append(trun(sep + text))

        for name, _bid, _ in moved:
            if name:
                bmnum[name] = num
        for e in body.findall(".//" + qn("w:bookmarkEnd")):
            if e.get(qn("w:id")) in ids and e.getparent() is not node:
                drop.add(e)

    for e in drop:
        if e.getparent() is not None:
            e.getparent().remove(e)
    return bmnum


CN_NUM = "零一二三四五六七八九"


def _cn_number(n):
    """1 -> 一, 11 -> 十一, 105 -> 一百零五（编号用足够了）"""
    if n < 10:
        return CN_NUM[n]
    if n < 20:
        return "十" + (CN_NUM[n % 10] if n % 10 else "")
    if n < 100:
        return CN_NUM[n // 10] + "十" + (CN_NUM[n % 10] if n % 10 else "")
    return CN_NUM[n // 100] + "百" + (CN_NUM[(n // 10) % 10] if (n // 10) % 10 else "零") \
        + (CN_NUM[n % 10] if n % 10 else "")


ROMAN = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"),
         (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"),
         (5, "V"), (4, "IV"), (1, "I")]


def _roman(n):
    out = ""
    for v, s in ROMAN:
        while n >= v:
            out += s
            n -= v
    return out


def _fmt_num(n, num_fmt):
    if num_fmt == "chineseCounting" or num_fmt == "chineseCountingThousand":
        return _cn_number(n)
    if num_fmt == "upperRoman":
        return _roman(n)
    if num_fmt == "lowerRoman":
        return _roman(n).lower()
    if num_fmt == "upperLetter":
        return chr(ord("A") + (n - 1) % 26) if n <= 26 else str(n)
    if num_fmt == "lowerLetter":
        return chr(ord("a") + (n - 1) % 26) if n <= 26 else str(n)
    return str(n)



def appendix_started(p, bms, cfg, text=None):
    """判断这个标题是不是附录的起点（或之后的附录标题）"""
    ap = cfg.get("appendix") or {}
    if not ap.get("enabled"):
        return False
    marker = ap.get("marker") or "sec:appendix"
    if any(b == marker for b in bms):
        return True
    return False


def build_appendix_numbering(doc, cfg):
    """建一套「附录 A / A.1」的多级列表，返回 numId 和绑定状态

    Word 的做法：附录用独立的多级列表，一级是字母（A、B），
    二级是 A.1。标题段落上挂段落级 numPr 指向它，覆盖样式里的编号。
    """
    ap = cfg.get("appendix") or {}
    if not ap.get("enabled"):
        return None
    try:
        num_root = doc.part.numbering_part.element
    except Exception:
        return None

    label = ap.get("label") or "附录"
    levels_cfg = ap.get("levels") or [
        {"format": "%s %%1" % label, "num_fmt": "upperLetter"},
        {"format": "%1.%2", "num_fmt": "decimal"},
        {"format": "%1.%2.%3", "num_fmt": "decimal"},
    ]
    levels = len(levels_cfg)

    used_abs = [int(a.get(wq("abstractNumId")))
                for a in num_root.findall(wq("abstractNum"))]
    used_num = [int(n.get(wq("numId"))) for n in num_root.findall(wq("num"))]
    abs_id = (max(used_abs) + 1) if used_abs else 0
    num_id = (max(used_num) + 1) if used_num else 1

    hn = cfg.get("heading_numbering") or {}
    ea = hn.get("east_asia_font") or "宋体"
    asc = hn.get("ascii_font") or "Times New Roman"

    absEl = el("w:abstractNum", **{"w:abstractNumId": abs_id})
    absEl.append(el("w:multiLevelType", **{"w:val": "multilevel"}))
    for lvl in range(levels):
        lc = levels_cfg[lvl] or {}
        l = el("w:lvl", **{"w:ilvl": lvl})
        l.append(el("w:start", **{"w:val": str(lc.get("start", 1))}))
        l.append(el("w:numFmt", **{"w:val": lc.get("num_fmt", "upperLetter" if lvl == 0 else "decimal")}))
        l.append(el("w:pStyle", **{"w:val": "Heading%d" % (lvl + 1)}))
        txt = lc.get("format") or (
            "%s %%1" % label if lvl == 0 else
            ".".join("%%%d" % (i + 1) for i in range(lvl + 1)))
        # 上级是字母、下级引用它时，要 isLgl 才会渲染成 A.1 而不是 A.一
        if lvl > 0:
            l.append(el("w:isLgl"))
        l.append(el("w:suff", **{"w:val": lc.get("suffix", "space")}))
        l.append(el("w:lvlText", **{"w:val": txt}))
        l.append(el("w:lvlJc", **{"w:val": "left"}))
        pPr = el("w:pPr")
        ind = el("w:ind", **{"w:left": "0", "w:leftChars": "0",
                             "w:firstLine": "0", "w:firstLineChars": "0"})
        pPr.append(ind); l.append(pPr)
        rPr = el("w:rPr")
        rPr.append(el("w:rFonts", **{"w:ascii": asc, "w:hAnsi": asc,
                                     "w:eastAsia": ea, "w:hint": "eastAsia"}))
        rPr.append(el("w:i", **{"w:val": "0"}))
        rPr.append(el("w:iCs", **{"w:val": "0"}))
        if lc.get("bold", hn.get("bold")):
            rPr.append(el("w:b"))
        l.append(rPr)
        absEl.append(l)

    first_num = num_root.find(wq("num"))
    if first_num is not None:
        first_num.addprevious(absEl)
    else:
        num_root.append(absEl)
    numEl = el("w:num", **{"w:numId": num_id})
    numEl.append(el("w:abstractNumId", **{"w:val": abs_id}))
    num_root.append(numEl)
    return num_id


def apply_appendix(doc, cfg, num_id):
    """从附录起点开始，把标题段落的编号挂到附录列表上"""
    if not num_id:
        return []
    ap = cfg.get("appendix") or {}
    label = ap.get("label") or "附录"
    cover_label = ap.get("prefix_text", False)
    touched = []
    inside = False
    pending = []
    for node in list(doc.element.body):
        if node.tag == qn("w:bookmarkStart"):
            nm = node.get(qn("w:name"))
            if nm:
                pending.append(nm)
            continue
        if node.tag != qn("w:p"):
            continue
        sid = style_id(node)
        if not sid or not sid.startswith("Heading"):
            continue
        bms = pending + [b.get(qn("w:name")) for b in node.iter(qn("w:bookmarkStart"))
                         if b.get(qn("w:name"))]
        end_marker = ap.get("end_marker") or "sec:endappendix"
        if any(b == end_marker for b in bms):
            inside = False
        if appendix_started(node, bms, cfg):
            inside = True
        pending = []
        if not inside:
            continue
        try:
            lvl = int(sid[7:]) - 1
        except ValueError:
            continue
        pPr = node.find(qn("w:pPr"))
        if pPr is None:
            pPr = el("w:pPr"); node.insert(0, pPr)
        old = pPr.find(qn("w:numPr"))
        if old is not None:
            pPr.remove(old)
        numPr = el("w:numPr")
        numPr.append(el("w:ilvl", **{"w:val": lvl}))
        numPr.append(el("w:numId", **{"w:val": num_id}))
        pPr.append(numPr)
        reorder(pPr, PPR_ORDER)
        # 一级附录标题前面补「附录」二字（可选）
        if lvl == 0 and cover_label and ap.get("prefix_text", True):
            text = para_text(node)
            if text and not text.startswith(label):
                for r in node.findall(qn("w:r")):
                    t = r.find(qn("w:t"))
                    if t is not None and t.text:
                        t.text = "%s %s" % (label, t.text)
                        t.set(qn("xml:space"), "preserve")
                        break
        touched.append(node)
    return touched


def collect_headings(doc, cfg):
    """扫出所有标题：返回 {书签名: (编号文字, 标题文字)}

    用来给 {{sec:x}} 交叉引用一个正确的缓存值（Word 里 F9 刷新前也能看）。
    """
    hn = cfg.get("heading_numbering") or {}
    levels = hn.get("levels") or []
    if not hn.get("enabled"):
        levels = []
    ap = cfg.get("appendix") or {}
    label = ap.get("label") or "附录"
    ap_levels = ap.get("levels") or ([
        {"format": "%s %%1" % label, "num_fmt": "upperLetter"},
        {"format": "%1.%2", "num_fmt": "decimal"},
        {"format": "%1.%2.%3", "num_fmt": "decimal"},
    ] if ap.get("enabled") else [])
    counters = [0] * max(len(levels), 1)
    ap_counters = [0] * max(len(ap_levels), 1)
    in_appendix = False
    out = {}
    body = doc.element.body
    pending = []          # pandoc 把标题的书签放在段落外面（作为兄弟节点）
    for node in body:
        if node.tag == qn("w:bookmarkStart"):
            nm = node.get(qn("w:name"))
            if nm:
                pending.append(nm)
            continue
        if node.tag != qn("w:p"):
            continue
        sid = style_id(node)
        if not sid or not sid.startswith("Heading"):
            continue
        try:
            lvl = int(sid[7:]) - 1
        except ValueError:
            continue
        text = para_text(node)
        bms = list(pending)
        for b in node.iter(qn("w:bookmarkStart")):
            nm = b.get(qn("w:name"))
            if nm and nm not in bms:
                bms.append(nm)
        # 结束标记：从这里开始回到正常章节编号
        end_marker = (cfg.get("appendix") or {}).get("end_marker") or "sec:endappendix"
        if any(b == end_marker for b in bms):
            in_appendix = False
        if appendix_started(node, bms, cfg):
            in_appendix = True
        if not bms:
            continue
        if in_appendix and ap_levels:
            if lvl < len(ap_levels):
                ap_counters[lvl] += 1
                for i in range(lvl + 1, len(ap_counters)):
                    ap_counters[i] = 0
                lc = ap_levels[lvl] or {}
                fmt = lc.get("format") or (
                    "%s %%1" % label if lvl == 0 else
                    ".".join("%%%d" % (i + 1) for i in range(lvl + 1)))
                label = fmt
                for i in range(lvl + 1):
                    # 每一级用自己的数字格式：一级字母 A，二级 A.1
                    li = ap_levels[i] or {}
                    lf = li.get("num_fmt", "upperLetter" if i == 0 else "decimal")
                    label = label.replace("%%%d" % (i + 1),
                                          _fmt_num(ap_counters[i], lf))
            else:
                label = ""
        elif lvl < len(levels):
            counters[lvl] += 1
            for i in range(lvl + 1, len(counters)):
                counters[i] = 0
            fmt = (levels[lvl] or {}).get("format") or ".".join(
                "%%%d" % (i + 1) for i in range(lvl + 1))
            label = fmt
            for i in range(lvl + 1):
                li = levels[i] or {}
                label = label.replace("%%%d" % (i + 1),
                                      _fmt_num(counters[i], li.get("num_fmt", "decimal")))
        else:
            label = ""
        for nm in bms:
            out[nm] = (label, text)
        pending = []
    return out


def refs(doc, bmnum, cfg, headings=None):
    """把 {{fig:x}} / {{tbl:x}} / {{eq:x}} / {{sec:x}}（以及旧的 @fig:x 写法）
    换成 Word 的 REF 域。

    为什么用 {{ }}：pandoc 的 --citeproc 会把任何 @key 当成文献引用吃掉，
    {{fig:x}} 不会被它碰，参考文献和交叉引用才能共存。
    """
    headings = headings or {}
    pat = re.compile(r"\{\{(fig|tbl|eq|sec):([\w\-.]+)\}\}"
                     r"|@(fig|tbl|eq|sec):([\w\-.]+)")
    tpls = {"fig": "{n}", "tbl": "{n}", "eq": "({n})", "sec": "{n}"}
    tpls.update(cfg.get("refs") or {})
    # 先收集所有可用锚点：指向不存在的锚点时不能静默变成「?」，
    # 否则 F9 之后会变成 Word 的「错误！未找到引用源」，而且用户很难发现。
    valid = set(bmnum) | set(headings)
    missing = []
    for para in doc.paragraphs:
        p = para._p
        for r in list(p.findall(qn("w:r"))):
            t = r.find(qn("w:t"))
            if t is None or not t.text or not pat.search(t.text):
                continue
            rPr = r.find(qn("w:rPr"))
            idx = list(p).index(r)
            p.remove(r)
            segs, pos = [], 0
            for m in pat.finditer(t.text):
                if m.start() > pos:
                    segs.append(trun(t.text[pos:m.start()]))
                if m.group(1):
                    kind, name = m.group(1), m.group(2)
                else:
                    kind, name = m.group(3), m.group(4)
                key = "%s:%s" % (kind, name)
                original = m.group(0)
                if key not in valid:
                    # 保留原文，让用户一眼看到写错了哪个锚点
                    if original not in missing:
                        missing.append(original)
                    segs.append(trun(original))
                    pos = m.end()
                    continue
                for ph, txt in split_ref_template(tpls.get(kind) or "{n}"):
                    if ph == "lit":
                        if txt:
                            segs.append(trun(txt))
                    elif ph == "txt":          # {t} = 标题文字
                        num, htext = headings.get(key, ("?", ""))
                        segs.extend(field(" REF %s \\h " % key, htext or "?"))
                    elif ph == "page":         # {p} = 页码
                        segs.extend(field(" REF %s \\p \\h " % key, "1"))
                    else:                      # {n} = 编号
                        if kind == "sec":
                            num, htext = headings.get(key, ("?", ""))
                            segs.extend(field(" REF %s \\w \\h " % key, num))
                        else:
                            segs.extend(field(" REF %s \\h " % key,
                                              bmnum.get(key, "?")))
                pos = m.end()
            if pos < len(t.text):
                segs.append(trun(t.text[pos:]))
            for i, nr in enumerate(segs):
                if rPr is not None and nr.tag == qn("w:r"):
                    nr.insert(0, rPr)
                p.insert(idx + i, nr)
    if missing:
        print("! 以下交叉引用找不到对应锚点，已原样保留（请检查 md 里的锚点名）：",
              file=sys.stderr)
        for x in missing:
            print("    %s" % x, file=sys.stderr)



# w:sectPr 子元素的合法顺序（Word 严格按这个校验，顺序错了会弹「修复」）
SECTPR_ORDER = ["headerReference", "footerReference", "footnotePr", "endnotePr",
                "type", "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType",
                "pgNumType", "cols", "formProt", "vAlign", "noEndnote",
                "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid",
                "printerSettings", "sectPrChange"]


def reorder_sections(doc):
    """把所有 sectPr 的子元素按 schema 顺序排好（headerReference 必须最前）"""
    n = 0
    for sectPr in doc.element.body.iter(wq("sectPr")):
        kids = list(sectPr)
        idx = {t: i for i, t in enumerate(SECTPR_ORDER)}
        # 同名的多个（如多个 headerReference）保持原有相对顺序
        order = sorted(range(len(kids)),
                       key=lambda i: (idx.get(kids[i].tag.split("}")[-1], 999), i))
        if order != list(range(len(kids))):
            n += 1
        for i in order:
            sectPr.append(kids[i])
    return n



# ---------------- 结构自检 ----------------
# Word 对 OOXML 的 schema 比 python-docx / LibreOffice 严格：
# 同一个父元素下重复出现本应唯一的子元素，Word 会弹「修复」框、拒绝打开。
# 下面这些父元素里的子元素都必须是唯一的。

UNIQUE_CHILDREN = ["pPr", "rPr", "sectPr", "tblPr", "trPr", "tcPr",
                   "tblBorders", "pBdr", "tblCellMar", "docGrid", "pgBorders"]


def roots_for_check(doc):
    out = [doc.element.body]
    for getter in ("styles", "settings"):
        try:
            out.append(getattr(doc, getter).element)
        except Exception:
            pass
    return out


def validate(doc, quiet=False):
    """检查重复子元素，返回问题列表"""
    problems = []
    body = doc.element.body
    roots = [body]
    try:
        roots.append(doc.styles.element)
        roots.append(doc.settings.element)
    except Exception:
        pass
    # 这些子元素本来就允许出现多次（按 w:type 区分）
    MULTI = {"headerReference", "footerReference", "activeWritingStyle",
             "attachedSchema", "smartTagType"}
    for root in roots:
        for parent_tag in UNIQUE_CHILDREN:
            for parent in root.iter(wq(parent_tag)):
                seen = {}
                for child in list(parent):
                    name = child.tag.split("}")[-1]
                    if name in MULTI:
                        # 同一 type 重复才算问题
                        t = child.get(wq("type")) or child.get(wq("val")) or ""
                        key = "%s[%s]" % (name, t)
                    else:
                        key = name
                    seen[key] = seen.get(key, 0) + 1
                for name, n in seen.items():
                    if n > 1:
                        problems.append("%s 下有 %d 个 %s" % (parent_tag, n, name))
    # 书签必须成对（Word 遇到孤立 bookmarkStart/End 会报损坏）
    for root in roots_for_check(doc):
        starts, ends = {}, {}
        for b in root.iter(wq("bookmarkStart")):
            bid = b.get(wq("id"))
            starts[bid] = starts.get(bid, 0) + 1
        for b in root.iter(wq("bookmarkEnd")):
            bid = b.get(wq("id"))
            ends[bid] = ends.get(bid, 0) + 1
        for bid in set(starts) | set(ends):
            if starts.get(bid, 0) != ends.get(bid, 0):
                problems.append("书签 id=%s 不成对（start %d / end %d）"
                                % (bid, starts.get(bid, 0), ends.get(bid, 0)))

    # 段落引用的样式必须存在
    try:
        have = {s.get(wq("styleId")) for s in doc.styles.element.findall(wq("style"))}
        for p in doc.element.body.iter(wq("p")):
            ps = p.find(wq("pPr") + "/" + wq("pStyle"))
            if ps is not None:
                sid = ps.get(wq("val"))
                if sid and sid not in have:
                    problems.append("引用了不存在的样式：%s" % sid)
    except Exception:
        pass

    # 跨部件引用：Word 会因为引用了不存在的部件而报「文件已损坏」，
    # 而 LibreOffice / python-docx 都能正常打开，所以必须自己查。
    try:
        docx_path = None
        parts = None
        # 从文档里拿包内所有部件名
        names = set()
        try:
            names = {str(p.partname) for p in doc.part.package.iter_parts()}
        except Exception:
            names = set()
        body = doc.element.body

        # endnote / footnote 分隔符引用
        for tag, part in (("w:endnote", "/word/endnotes.xml"),
                          ("w:footnote", "/word/footnotes.xml")):
            refs_used = [e.get(wq("id")) for e in body.iter(wq(tag.split(":")[1]))
                         if e.getparent() is not None
                         and e.getparent().tag == wq(tag.split(":")[1] + "Pr")]
            if refs_used and part not in names:
                problems.append("sectPr 引用了 %s id=%s，但包里没有 %s"
                                % (tag.split(":")[1], refs_used[:3], part))

        # 正文的脚注/尾注引用必须在对应部件里
        for kind, part, idtag in (("footnoteReference", "/word/footnotes.xml", "footnote"),
                                  ("endnoteReference", "/word/endnotes.xml", "endnote")):
            used = [e.get(wq("id")) for e in body.iter(wq(kind))]
            if not used:
                continue
            if part not in names:
                problems.append("正文有 %s，但包里没有 %s" % (kind, part))
                continue
            try:
                target = None
                for p in doc.part.package.iter_parts():
                    if str(p.partname) == part:
                        target = p
                        break
                have = {e.get(wq("id")) for e in target.element.iter(wq(idtag))}
                miss = {u for u in used if u not in have}
                if miss:
                    problems.append("%s 指向不存在的 %s id=%s"
                                    % (kind, idtag, sorted(miss)[:3]))
            except Exception:
                pass

        # 段落/样式引用的 numId 必须存在
        try:
            num_part = doc.part.numbering_part
            defined_num = {n.get(wq("numId")) for n in num_part.element.findall(wq("num"))}
            defined_abs = {a.get(wq("abstractNumId"))
                           for a in num_part.element.findall(wq("abstractNum"))}
            used_num = set()
            for root in roots_for_check(doc):
                for e in root.iter(wq("numId")):
                    v = e.get(wq("val"))
                    if v and v != "0":
                        used_num.add(v)
            miss = used_num - defined_num
            if miss:
                problems.append("引用了不存在的 numId %s" % sorted(miss)[:4])
            used_abs = {e.get(wq("val")) for e in num_part.element.iter(wq("abstractNumId"))}
            miss = used_abs - defined_abs
            if miss:
                problems.append("引用了不存在的 abstractNumId %s" % sorted(miss)[:4])
        except Exception:
            pass

        # [Content_Types].xml 里声明了但包里没有的部件
        try:
            ct = None
            for p in doc.part.package.iter_parts():
                if str(p.partname) == "/[Content_Types].xml":
                    ct = p
                    break
            if ct is not None:
                declared = {str(e.get("PartName")).lstrip("/")
                            for e in ct.element.iter() if e.get("PartName")}
                have = {str(p.partname).lstrip("/") for p in doc.part.package.iter_parts()}
                ghosts = declared - have
                if ghosts:
                    problems.append("Content_Types 声明了不存在的部件 %s"
                                    % sorted(ghosts)[:4])
        except Exception:
            pass
    except Exception:
        pass

    # 各容器的子元素顺序
    for _root in roots_for_check(doc):
      for tag, order in CONTAINER_ORDER.items():
        for parent in _root.iter(wq(tag)):
            names = [c.tag.split("}")[-1] for c in list(parent)]
            idxs = [order.index(n) for n in names if n in order]
            if idxs != sorted(idxs):
                problems.append("%s 子元素顺序不对：%s" % (tag, " ".join(names)))
    # sectPr 子元素顺序（header/footerReference 是重复出现的 choice 组，
    # 允许交错，所以比较时把它们合并成一组）
    for sectPr in doc.element.body.iter(wq("sectPr")):
        names = [c.tag.split("}")[-1] for c in list(sectPr)]
        merged = []
        for n in names:
            if n in ("headerReference", "footerReference"):
                if not merged or merged[-1] != "HDRFTR":
                    merged.append("HDRFTR")
            else:
                merged.append(n)
        rank = {"HDRFTR": 0}
        rank.update({n: i + 1 for i, n in enumerate(SECTPR_ORDER)
                     if n not in ("headerReference", "footerReference")})
        idxs = [rank[n] for n in merged if n in rank]
        if idxs != sorted(idxs):
            problems.append("sectPr 子元素顺序不对：%s" % " ".join(names))
    # 只报唯一的问题（同一类只报一次）
    uniq = []
    for p in problems:
        if p not in uniq:
            uniq.append(p)
    if uniq and not quiet:
        print("  ! 结构问题（Word 可能提示修复）：", file=sys.stderr)
        for p in uniq:
            print("    - %s" % p, file=sys.stderr)
    return uniq


TRPR_ORDER = ["cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter",
              "cantSplit", "trHeight", "tblHeader", "tblCellSpacing", "jc",
              "hidden", "ins", "del", "trPrChange"]

TCPR_ORDER = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders",
              "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign",
              "hideMark", "headers", "cellIns", "cellDel", "cellMerge",
              "tcPrChange"]

STYLE_ORDER = ["name", "aliases", "basedOn", "next", "link", "autoRedefine",
               "hidden", "uiPriority", "semiHidden", "unhideWhenUsed", "qFormat",
               "locked", "personal", "personalCompose", "personalReply", "rsid",
               "pPr", "rPr", "tblPr", "trPr", "tcPr", "tblStylePr"]



def normalize_structure(doc):
    """把所有容器的子元素按 OOXML schema 顺序排好

    Word 对这个顺序极其严格：顺序错了不一定报错，但经常直接弹
    「文件已损坏，是否修复」，而 LibreOffice / python-docx 都照常打开。
    这是本项目遇到过的最大坑，所以每次输出前统一过一遍。
    """
    fixed = 0
    roots = [doc.element.body]
    try:
        roots.append(doc.styles.element)
    except Exception:
        pass
    try:
        roots.append(doc.settings.element)
    except Exception:
        pass
    for root in roots:
        for tag, order in CONTAINER_ORDER.items():
            for parent in root.iter(wq(tag)):
                kids = list(parent)
                idx = {t: i for i, t in enumerate(order)}
                want = sorted(range(len(kids)),
                              key=lambda i: (idx.get(kids[i].tag.split("}")[-1], 999), i))
                if want != list(range(len(kids))):
                    fixed += 1
                    for i in want:
                        parent.append(kids[i])
    return fixed

SETTINGS_ORDER = [
    "writeProtection", "view", "zoom", "removePersonalInformation",
    "removeDateAndTime", "doNotDisplayPageBoundaries", "displayBackgroundShape",
    "printPostScriptOverText", "printFractionalCharacterWidth", "printFormsData",
    "embedTrueTypeFonts", "embedSystemFonts", "saveSubsetFonts", "saveFormsData",
    "mirrorMargins", "alignBordersAndEdges", "bordersDoNotSurroundHeader",
    "bordersDoNotSurroundFooter", "gutterAtTop", "hideSpellingErrors",
    "hideGrammaticalErrors", "activeWritingStyle", "proofState", "formsDesign",
    "attachedTemplate", "linkStyles", "stylePaneFormatFilter",
    "stylePaneSortMethod", "documentType", "mailMerge", "revisionView",
    "trackChanges", "doNotTrackMoves", "doNotTrackFormatting",
    "documentProtection", "autoFormatOverride", "styleLockTheme",
    "styleLockQFSet", "defaultTabStop", "autoHyphenation",
    "consecutiveHyphenLimit", "hyphenationZone", "doNotHyphenateCaps",
    "showEnvelope", "summaryLength", "clickAndTypeStyle", "defaultTableStyle",
    "evenAndOddHeaders", "bookFoldRevPrinting", "bookFoldPrinting",
    "bookFoldPrintingSheets", "drawingGridHorizontalSpacing",
    "drawingGridVerticalSpacing", "displayHorizontalDrawingGridEvery",
    "displayVerticalDrawingGridEvery", "doNotUseMarginsForDrawingGridOrigin",
    "drawingGridHorizontalOrigin", "drawingGridVerticalOrigin",
    "doNotShadeFormData", "noPunctuationKerning", "characterSpacingControl",
    "printTwoOnOne", "strictFirstAndLastChars", "noLineBreaksAfter",
    "noLineBreaksBefore", "savePreviewPicture", "doNotValidateAgainstSchema",
    "saveInvalidXml", "ignoreMixedContent", "alwaysShowPlaceholderText",
    "doNotDemarcateInvalidXml", "saveXmlDataOnly", "useXSLTWhenSaving",
    "saveThroughXslt", "showXMLTags", "alwaysMergeEmptyNamespace",
    "updateFields", "hdrShapeDefaults", "footnotePr", "endnotePr", "compat",
    "docVars", "rsids", "mathPr", "uiCompat97To2003", "attachedSchema",
    "themeFontLang", "clrSchemeMapping", "doNotIncludeSubdocsInStats",
    "doNotAutoCompressPictures", "forceUpgrade", "captions",
    "readModeInkLockDown", "smartTagType", "schemaLibrary", "shapeDefaults",
    "doNotEmbedSmartTags", "decimalSymbol", "listSeparator"]

# 容器 -> 合法子元素顺序
CONTAINER_ORDER = {
    "pPr": PPR_ORDER, "rPr": RPR_ORDER, "tblPr": TBLPR_ORDER,
    "trPr": TRPR_ORDER, "tcPr": TCPR_ORDER, "sectPr": SECTPR_ORDER,
    "settings": SETTINGS_ORDER, "style": STYLE_ORDER,
}



def insert_in_order(parent, elem, order):
    """把 elem 按 schema 顺序插到 parent 里（Word 对顺序极严，错了就报损坏）"""
    name = elem.tag.split("}")[-1]
    idx = order.index(name) if name in order else len(order)
    for i, child in enumerate(list(parent)):
        cname = child.tag.split("}")[-1]
        cidx = order.index(cname) if cname in order else len(order)
        if cidx > idx:
            parent.insert(i, elem)
            return elem
    parent.append(elem)
    return elem


def reorder_settings(doc):
    st = doc.settings.element
    reorder(st, SETTINGS_ORDER)



# ---------------- 参考文献自动编号 ----------------
# pandoc 的 citeproc 把 [1] [2] 当普通文字写进文档，删掉一条后按 F9 不会重排。
# 这里把书目编号换成 SEQ 域、正文角标换成 REF 域，就和图表编号一样是「活的」。

BIB_SEQ = "文献"
BIB_BM = "_Ref_bib_"


def _leading_bracket_number(p):
    """找出段首的 [n]（可能跨 run），返回 (n, runs_used, tail_in_last_run)"""
    runs = p.findall(qn("w:r"))
    acc, used = "", []
    for r in runs:
        t = r.find(qn("w:t"))
        acc += (t.text or "") if t is not None else ""
        used.append(r)
        m = re.match(r"^\[\s*(\d+)\s*\]", acc)
        if m:
            return m.group(1), used, acc[m.end():]
        if not re.match(r"^\[\s*\d*$", acc):
            return None, [], ""
        if len(used) > 4:
            break
    return None, [], ""


def auto_number_bibliography(doc, cfg):
    """书目编号 -> SEQ 域；正文角标 -> REF 域。返回 (编号数, 角标数)"""
    cit = cfg.get("citations") or {}
    if cit.get("number_bibliography") is False:
        return 0, 0
    body = doc.element.body

    # ① 收集书目条目：pandoc 给每条在段落前放一个 ref-<key> 书签
    entries, pending = [], []
    for node in list(body):
        if node.tag == qn("w:bookmarkStart"):
            nm = node.get(qn("w:name")) or ""
            if nm.startswith("ref-"):
                pending.append(nm)
            continue
        if node.tag == qn("w:p") and style_id(node) == "Bibliography":
            entries.append((node, pending[0] if pending else None))
        pending = []
    if not entries:
        return 0, 0

    # ② 把每条的 [n] 换成 [<SEQ>]，并在编号上挂一个自己的书签给 REF 用
    targets = {}
    numbered = 0
    next_id = 30000
    for para, key in entries:
        num, used, tail = _leading_bracket_number(para)
        if num is None:
            continue
        first = used[0]
        t = first.find(qn("w:t"))
        if t is None:
            continue
        # 第一个 run 只留 "["
        t.text = "["
        t.set(qn("xml:space"), "preserve")
        # 后续被吃掉的 run 里，最后一个可能有残留文字（一般没有）
        for r in used[1:]:
            rt = r.find(qn("w:t"))
            if rt is not None and rt.text and rt.text.strip():
                rt.text = rt.text
            if r.getparent() is not None:
                r.getparent().remove(r)
        # 在 "[" 之后插入：书签 + SEQ 域 + "]" + 残留文字
        after = [el("w:bookmarkEnd", **{"w:id": next_id})]
        block = []
        bmname = BIB_BM + (key[4:] if key else "n%d" % (numbered + 1))
        block.append(el("w:bookmarkStart", **{"w:id": next_id, "w:name": bmname}))
        block.extend(field(" SEQ %s \\* ARABIC " % BIB_SEQ, num))
        block.append(el("w:bookmarkEnd", **{"w:id": next_id}))
        block.append(trun("]"))
        if tail:
            block.append(trun(tail))
        anchor = first
        for b in block:
            anchor.addnext(b)
            anchor = b
        if key:
            targets[key] = bmname
        numbered += 1
        next_id += 1

    # ③ 正文角标：[1] 里的 "1" 换成指向书目的 REF 域
    replaced = 0
    for h in body.iter(qn("w:hyperlink")):
        anchor = h.get(qn("w:anchor")) or ""
        if not anchor.startswith("ref-") or anchor not in targets:
            continue
        runs = h.findall(qn("w:r"))
        if not runs:
            continue
        # 缓存值用原来的编号，这样在 Word 里按 F9 之前也能正确显示
        cached = "".join(t.text or "" for t in h.iter(qn("w:t"))).strip()
        rPr = runs[0].find(qn("w:rPr"))
        for r in runs:
            h.remove(r)
        for fr in field(" REF %s \\h " % targets[anchor], cached or "?"):
            if rPr is not None and fr.tag == qn("w:r"):
                fr.insert(0, copy.deepcopy(rPr))
            h.append(fr)
        replaced += 1
    return numbered, replaced


# ---------------- 列表编号 / 文档级设置 ----------------
# 注意：pandoc 会重写 numbering.xml 和 settings.xml，
# 所以这两类配置只能在后处理阶段做，写进 reference.docx 会被丢掉。


def apply_list_numbering(doc, cfg):
    """列表编号：有序列表的编号样式、无序列表的符号、缩进

    pandoc 把列表编号写在 numbering.xml 里，这里按配置改写。
    """
    lists = cfg.get("lists") or {}
    if not lists:
        return
    try:
        num_root = doc.part.numbering_part.element
    except Exception:
        return
    ordered = lists.get("ordered") or {}
    bullet = lists.get("bullet") or {}
    for absEl in num_root.findall(wq("abstractNum")):
        for lvl in absEl.findall(wq("lvl")):
            ilvl = int(lvl.get(wq("ilvl")) or 0)
            fmt = lvl.find(wq("numFmt"))
            kind = fmt.get(wq("val")) if fmt is not None else ""
            txt = lvl.find(wq("lvlText"))
            if txt is None:
                continue
            if kind == "bullet":
                if bullet.get("char"):
                    txt.set(wq("val"), bullet["char"])
                rf = lvl.find(wq("rPr") + "/" + wq("rFonts"))
                if rf is not None:
                    if bullet.get("font"):
                        for a in ("ascii", "hAnsi"):
                            rf.set(wq(a), bullet["font"])
                        rf.set(wq("hint"), "default")
            elif kind in ("decimal", "lowerLetter", "upperLetter",
                          "lowerRoman", "upperRoman", "chineseCounting"):
                if ordered.get("num_fmt"):
                    fmt.set(wq("val"), ordered["num_fmt"])
                if ordered.get("format"):
                    txt.set(wq("val"), ordered["format"].replace(
                        "%1", "%%%d" % (ilvl + 1)))
                if ordered.get("indent_chars"):
                    pPr = lvl.find(wq("pPr"))
                    if pPr is None:
                        pPr = el("w:pPr"); lvl.append(pPr)
                    ind = pPr.find(wq("ind"))
                    if ind is None:
                        ind = el("w:ind"); pPr.append(ind)
                    ch = float(ordered["indent_chars"])
                    ind.set(wq("left"), str(int(round(ch * 120))))
                    ind.set(wq("leftChars"), str(int(ch * 100)))
                    ind.set(wq("hanging"), str(int(round(ch * 84))))
                    ind.set(wq("hangingChars"), str(int(ch * 35)))


def apply_document_settings(doc, cfg):
    """文档级开关：修订模式、文档保护、默认制表位、视图缩放

    这些也属于 pandoc 会重写的 settings.xml，只能后处理。
    """
    st = doc.settings.element
    doc_cfg = cfg.get("document") or {}

    def ensure(tag, **attrs):
        e = st.find(wq(tag))
        if e is None:
            e = insert_in_order(st, el("w:" + tag), SETTINGS_ORDER)
        for k, v in attrs.items():
            e.set(wq(k), str(v))
        return e

    if doc_cfg.get("track_changes"):
        ensure("trackChanges")
    prot = doc_cfg.get("protect")
    if prot:
        old = st.find(wq("documentProtection"))
        if old is not None:
            st.remove(old)
        e = el("w:documentProtection")
        e.set(wq("edit"), prot if isinstance(prot, str) else "readOnly")
        e.set(wq("enforcement"), "1")
        insert_in_order(st, e, SETTINGS_ORDER)
    # 兼容性模式：pandoc 默认模板没有 compat 块，Word 会按「兼容性模式」打开。
    # 写 compatibilityMode=15 才是 Word 2013+ 的正常模式。
    if doc_cfg.get("compatibility_mode", 15):
        compat = st.find(wq("compat"))
        if compat is None:
            compat = insert_in_order(st, el("w:compat"), SETTINGS_ORDER)
        cm = doc_cfg.get("compatibility_mode", 15)
        found = [c for c in compat.findall(wq("compatSetting"))
                 if c.get(wq("name")) == "compatibilityMode"]
        if not found:
            e = el("w:compatSetting")
            e.set(wq("name"), "compatibilityMode")
            e.set(wq("uri"), "http://schemas.microsoft.com/office/word")
            e.set(wq("val"), str(int(cm)))
            compat.append(e)
        else:
            found[0].set(wq("val"), str(int(cm)))

    tab = doc_cfg.get("default_tab_cm")
    if tab:
        ensure("defaultTabStop", val=str(int(round(float(tab) * 567))))
    zoom = doc_cfg.get("zoom")
    if zoom:
        ensure("zoom", percent=str(int(zoom)))


# ---------------- 3. 封面 / 目录 / 分节 / 页码 ----------------

PLACEHOLDERS = ("{{TOC}}", "{{TOF}}", "{{TOT}}", "[[TOC]]", "[TOC]", "{{toc}}")


def empty_spacer(cm):
    """按厘米生成一段空白（用于封面留白），行高固定"""
    p = el("w:p")
    pPr = el("w:pPr")
    pPr.append(el("w:spacing", **{"w:before": "0", "w:after": "0",
                                  "w:line": str(int(cm * 567)),
                                  "w:lineRule": "exact"}))
    rPr = el("w:rPr"); rPr.append(el("w:sz", **{"w:val": "2"}))
    pPr.append(rPr)
    p.append(pPr)
    r = el("w:r"); rr = el("w:rPr"); rr.append(el("w:sz", **{"w:val": "2"}))
    r.append(rr); p.append(r)
    return p



def make_image_para(doc, path, width_cm=None, align="center", space_before=0,
                    space_after=12):
    """生成一个只放图片的段落（封面 logo 用）"""
    if not path or not os.path.exists(path):
        return None
    p = el("w:p")
    pPr = el("w:pPr")
    if align:
        pPr.append(el("w:jc", **{"w:val": align}))
    if space_before or space_after:
        pPr.append(el("w:spacing", **{"w:before": str(int(space_before * 20)),
                                      "w:after": str(int(space_after * 20))}))
    reorder(pPr, PPR_ORDER)
    p.append(pPr)
    run = el("w:r")
    # 用 python-docx 的 add_picture 生成 drawing，再搬到这个段落里
    tmp = doc.add_paragraph()
    try:
        r = tmp.add_run()
        if width_cm:
            r.add_picture(path, width=Cm(float(width_cm)))
        else:
            r.add_picture(path)
        for child in list(r._element):
            run.append(copy.deepcopy(child))
    except Exception:
        tmp._element.getparent().remove(tmp._element)
        return None
    tmp._element.getparent().remove(tmp._element)
    p.append(run)
    return p


def make_para(text, style=None, align=None):
    p = el("w:p")
    if style or align:
        pPr = el("w:pPr")
        if style:
            pPr.append(el("w:pStyle", **{"w:val": style}))
        if align:
            pPr.append(el("w:jc", **{"w:val": align}))
        reorder(pPr, PPR_ORDER)
        p.append(pPr)
    if text:
        p.append(trun(text))
    return p


def style_exists(doc, sid):
    return any(s.get(wq("styleId")) == sid
               for s in doc.styles.element.findall(wq("style")))


def set_para_style(p, sid):
    pPr = p.find(qn("w:pPr"))
    if pPr is None:
        pPr = el("w:pPr"); p.insert(0, pPr)
    old = pPr.find(qn("w:pStyle"))
    if old is not None:
        pPr.remove(old)
    pPr.insert(0, el("w:pStyle", **{"w:val": sid}))


def break_para(base_sect):
    """分节符：把 sectPr 放进上一节最后一个段落的 pPr

    注意：必须清掉从末尾节抄来的 header/footer 引用，
    否则每一节都会共用同一个页眉页脚部件，无法分别设页码。
    """
    p = el("w:p")
    pPr = el("w:pPr")
    sect = copy.deepcopy(base_sect)
    for tag in ("w:headerReference", "w:footerReference", "w:titlePg"):
        for e in sect.findall(qn(tag)):
            sect.remove(e)
    pPr.append(sect)
    p.append(pPr)
    return p


def toc_field_para(levels):
    p = el("w:p")
    for r in field(' TOC \\o "%s" \\h \\z \\u ' % levels, "右键“更新域”生成目录"):
        p.append(r)
    return p


def index_field_para(label):
    p = el("w:p")
    for r in field(' TOC \\h \\z \\c "%s" ' % label, "右键“更新域”生成索引"):
        p.append(r)
    return p


def set_pgnum(sectPr, fmt=None, start=None):
    old = sectPr.find(qn("w:pgNumType"))
    if old is not None:
        sectPr.remove(old)
    if fmt is None and start is None:
        return
    e = el("w:pgNumType")
    if fmt:
        e.set(qn("w:fmt"), fmt)
    if start is not None:
        e.set(qn("w:start"), str(start))
    anchor = sectPr.find(qn("w:pgMar"))
    if anchor is not None:
        anchor.addnext(e)
    else:
        sectPr.append(e)


def _style_run(r, cfg):
    rPr = r._element.get_or_add_rPr()
    ea, asc = cfg.get("east_asia_font"), cfg.get("ascii_font")
    if ea or asc:
        rf = rPr.find(wq("rFonts"))
        if rf is None:
            rf = el("w:rFonts"); rPr.insert(0, rf)
        for a in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
            if rf.get(wq(a)) is not None:
                del rf.attrib[wq(a)]
        if asc:
            rf.set(wq("ascii"), asc); rf.set(wq("hAnsi"), asc)
        if ea:
            rf.set(wq("eastAsia"), ea); rf.set(wq("hint"), "eastAsia")
    if cfg.get("size_pt") is not None:
        half = str(int(round(float(cfg["size_pt"]) * 2))) if isinstance(
            cfg["size_pt"], (int, float)) else None
        if half:
            for tag in ("sz", "szCs"):
                e = rPr.find(wq(tag))
                if e is None:
                    e = el("w:" + tag); rPr.append(e)
                e.set(wq("val"), half)
    for tag in ("i", "iCs"):
        e = rPr.find(wq(tag))
        if e is None:
            e = el("w:" + tag); rPr.append(e)
        e.set(wq("val"), "0")


def build_header(sec, cfg, enabled=True, even=False):
    h = cfg.get("header") or {}
    hdr = sec.even_page_header if even else sec.header
    hdr.is_linked_to_previous = False
    for p in list(hdr.paragraphs):
        p._element.getparent().remove(p._element)
    para = hdr.add_paragraph()
    if not enabled:
        return
    text = h.get("even_text") if even else h.get("text")
    if not text:
        return
    if even:
        para.add_run(text)
    elif h.get("use_chapter"):
        # 奇数页：固定文字 + 当前一级标题（STYLEREF 域，会跟着章节变）
        if text:
            para.add_run(text + " · ")
        for r in field(" STYLEREF 1 \\* MERGEFORMAT ", "章节标题"):
            para._p.append(r)
    else:
        para.add_run(text)
    if h.get("align"):
        para.alignment = {"left": 0, "center": 1, "right": 2}[h["align"]]
    for r in para.runs:
        _style_run(r, h)
    if h.get("bottom_border"):
        pPr = para._p.get_or_add_pPr()
        pbdr = el("w:pBdr")
        b = el("w:bottom", **{"w:val": "single", "w:sz": "6",
                              "w:space": "1", "w:color": "auto"})
        pbdr.append(b); pPr.append(pbdr)
        reorder(pPr, PPR_ORDER)


CACHED_BY_FMT = {"upperRoman": "I", "lowerRoman": "i", "decimal": "1",
                 "chineseCounting": "一", "chineseCountingThousand": "一",
                 "upperLetter": "A", "lowerLetter": "a"}


def build_footer(sec, cfg, enabled=True):
    f = cfg.get("footer") or {}
    ftr = sec.footer
    ftr.is_linked_to_previous = False
    for p in list(ftr.paragraphs):
        p._element.getparent().remove(p._element)
    para = ftr.add_paragraph()
    if not enabled or not f.get("page_number", True):
        return
    if f.get("align"):
        para.alignment = {"left": 0, "center": 1, "right": 2}[f["align"]]
    fmt = f.get("format") or "第 {PAGE} 页 共 {NUMPAGES} 页"
    pg = sec._sectPr.find(qn("w:pgNumType"))
    page_fmt = pg.get(qn("w:fmt")) if pg is not None else None
    cached = CACHED_BY_FMT.get(page_fmt or "decimal", "1")
    for piece in re.split(r"(\{PAGE\}|\{NUMPAGES\}|\{SECTIONPAGES\})", fmt):
        if piece == "{PAGE}":
            for r in field(" PAGE ", cached):
                para._p.append(r)
        elif piece == "{NUMPAGES}":
            for r in field(" NUMPAGES ", "1"):
                para._p.append(r)
        elif piece == "{SECTIONPAGES}":
            for r in field(" SECTIONPAGES ", "1"):
                para._p.append(r)
        elif piece:
            para.add_run(piece)
    for r in para.runs:
        _style_run(r, f)


def apply_page_structure(doc, cfg):
    """封面 + 前置部分 + 分节 + 页码，全部由 profile 控制"""
    body = doc.element.body
    base_sect = body.find(qn("w:sectPr"))
    if base_sect is None:
        return []
    cov = cfg.get("cover") or {}
    fm = cfg.get("front_matter") or {}
    toc = cfg.get("toc") or {}

    # 找到 pandoc 生成的标题块（Title / Subtitle / Author / Date）
    title_sids = ("Title", "Subtitle", "Author", "Date")
    title_map = {"Title": "CoverTitle", "Subtitle": "CoverSubtitle",
                 "Author": "CoverAuthor", "Date": "CoverDate"}
    title_paras = []
    for child in list(body):
        if child.tag != qn("w:p"):
            break
        sid = style_id(child)
        if sid in title_sids:
            title_paras.append(child)
        elif not para_text(child):
            continue
        else:
            break

    # 清掉 md 里的占位符段落（目录位置改由配置决定）
    for p in list(body.iter(qn("w:p"))):
        if para_text(p) in PLACEHOLDERS:
            p.getparent().remove(p)

    roles = []
    cursor = 0

    # ---------- 封面 ----------
    if cov.get("enabled"):
        nodes = []
        if cov.get("top_space_cm"):
            nodes.append(empty_spacer(float(cov["top_space_cm"])))
        # logo 放在标题上方（可选）
        logo = cov.get("logo")
        if logo:
            lp = make_image_para(doc, logo, cov.get("logo_width_cm"),
                                 cov.get("logo_align", "center"),
                                 space_after=cov.get("logo_space_after_pt", 18))
            if lp is not None:
                nodes.append(lp)
        for line in (cov.get("above") or []):
            nodes.append(make_para(line, cov.get("line_style") or "CoverLine"))
        for p in title_paras:
            if p.getparent() is not None:
                p.getparent().remove(p)
            target = title_map.get(style_id(p))
            if cov.get("restyle", True) and target and style_exists(doc, target):
                set_para_style(p, target)
            nodes.append(p)
        for line in (cov.get("below") or []):
            nodes.append(make_para(line, cov.get("line_style") or "CoverLine"))
        logo_bottom = cov.get("logo_bottom")
        if logo_bottom:
            lp = make_image_para(doc, logo_bottom, cov.get("logo_bottom_width_cm"),
                                 cov.get("logo_align", "center"),
                                 space_before=18, space_after=0)
            if lp is not None:
                nodes.append(lp)
        if cov.get("bottom_space_cm"):
            nodes.append(empty_spacer(float(cov["bottom_space_cm"])))
        nodes.append(break_para(base_sect))
        for n in nodes:
            body.insert(cursor, n); cursor += 1
        roles.append("cover")

    # ---------- 摘要在前置部分（pandoc 已经把 abstract 生成为 Abstract/AbstractTitle 段）----------
    abst = cfg.get("abstract") or {}
    meta = cfg.get("_meta") or {}
    abstract_nodes = []
    if abst.get("enabled", True):
        title_p = None
        abs_ps = []
        for p in list(body.iter(qn("w:p"))):
            sid = style_id(p)
            if sid == "AbstractTitle":
                title_p = p
            elif sid == "Abstract":
                abs_ps.append(p)
        if title_p is not None or abs_ps:
            if title_p is not None and title_p.getparent() is not None:
                title_p.getparent().remove(title_p)
                clear_runs(title_p)
                title_p.append(trun(abst.get("title", "摘　要")))
                set_para_style(title_p, abst.get("title_style") or "TOCHeading")
                abstract_nodes.append(title_p)
            for p in abs_ps:
                if p.getparent() is not None:
                    p.getparent().remove(p)
                if abst.get("style"):
                    set_para_style(p, abst["style"])
                abstract_nodes.append(p)
            kws = meta.get("keywords")
            if kws:
                if isinstance(kws, str):
                    kws = [k.strip().strip("[]") for k in
                           kws.replace("，", ",").replace("；", ",")
                              .replace("[", "").replace("]", "").split(",")
                           if k.strip()]
                else:
                    kws = [str(k).strip().strip("[]") for k in kws if str(k).strip()]
                kw_text = (abst.get("keywords_title", "关键词") + "："
                           + "；".join(str(k) for k in kws))
                kp = make_para(kw_text, abst.get("keywords_style")
                               or abst.get("style") or "Abstract")
                abstract_nodes.append(kp)
            if abstract_nodes:
                brk = el("w:p"); r = el("w:r")
                r.append(el("w:br", **{"w:type": "page"}))
                brk.append(r)
                abstract_nodes.append(brk)

    # ---------- 前置部分：摘要 / 目录 / 图目录 / 表目录 ----------
    fm_nodes = list(abstract_nodes)
    if toc.get("auto", True):
        if fm.get("toc_title"):
            fm_nodes.append(make_para(fm["toc_title"], "TOCHeading"))
        fm_nodes.append(toc_field_para(toc.get("levels", "1-3")))
    figs = fm.get("figures") or {}
    if figs.get("enabled"):
        fm_nodes.append(make_para(figs.get("title") or "图目录", "TOCHeading"))
        fm_nodes.append(index_field_para(figs.get("label") or "图"))
    tbls = fm.get("tables") or {}
    if tbls.get("enabled"):
        fm_nodes.append(make_para(tbls.get("title") or "表目录", "TOCHeading"))
        fm_nodes.append(index_field_para(tbls.get("label") or "表"))

    if fm_nodes:
        fm_nodes.append(break_para(base_sect))
        for n in fm_nodes:
            body.insert(cursor, n); cursor += 1
        roles.append("front")

    roles.append("body")

    # ---------- 分节配置 ----------
    secs = list(doc.sections)
    if len(secs) != len(roles):
        print("  ! 分节数不符（%d vs %d），跳过分节配置" % (len(secs), len(roles)),
              file=sys.stderr)
        return roles

    page = cfg.get("page") or {}
    body_pg = cfg.get("body_paging") or {}
    for sec, role in zip(secs, roles):
        sectPr = sec._sectPr
        if role == "cover":
            set_pgnum(sectPr,
                      cov.get("page_num_fmt", "decimal") if cov.get("page_number") else None,
                      cov.get("page_num_start") if cov.get("page_number") else None)
            build_header(sec, cfg, enabled=bool(cov.get("header")))
            build_footer(sec, cfg, enabled=bool(cov.get("page_number")))
        elif role == "front":
            set_pgnum(sectPr, fm.get("page_num_fmt", "upperRoman"),
                      fm.get("page_num_start", 1))
            build_header(sec, cfg, enabled=bool(fm.get("header")))
            build_footer(sec, cfg, enabled=bool(fm.get("page_number", True)))
        else:
            set_pgnum(sectPr, body_pg.get("page_num_fmt", "decimal"),
                      body_pg.get("page_num_start", 1))
            build_header(sec, cfg, enabled=True)
            # 奇偶页不同
            if (cfg.get("header") or {}).get("even_text"):
                build_header(sec, cfg, enabled=True, even=True)
                st = doc.settings.element
                if st.find(qn("w:evenAndOddHeaders")) is None:
                    insert_in_order(st, el("w:evenAndOddHeaders"), SETTINGS_ORDER)
            build_footer(sec, cfg, enabled=bool((cfg.get("footer") or {}).get("page_number", True)))
            first_diff = bool(page.get("first_page_different")) and not cov.get("enabled")
            sec.different_first_page_header_footer = first_diff
            if not first_diff:      # 清掉残留的「首页」页眉页脚引用
                for tag in ("w:headerReference", "w:footerReference"):
                    for e in sectPr.findall(qn(tag)):
                        if e.get(qn("w:type")) == "first":
                            sectPr.remove(e)
    return roles


# ---------------- 4. 标题自动编号 ----------------

def heading_numbering(doc, cfg):
    hn = cfg.get("heading_numbering") or {}
    if not hn.get("enabled"):
        return
    levels_cfg = hn.get("levels")
    if isinstance(levels_cfg, int):
        levels_cfg = [{}] * levels_cfg
    if not levels_cfg:
        fmts = hn.get("format") or []
        levels_cfg = [{"format": f} for f in fmts]
    if not levels_cfg:
        levels_cfg = [{"format": "%1"}, {"format": "%1.%2"}, {"format": "%1.%2.%3"}]
    levels = len(levels_cfg)

    numbering_part = doc.part.numbering_part
    num_root = numbering_part.element
    root = doc.styles.element

    used_abs = [int(a.get(wq("abstractNumId"))) for a in num_root.findall(wq("abstractNum"))]
    used_num = [int(n.get(wq("numId"))) for n in num_root.findall(wq("num"))]
    abs_id = (max(used_abs) + 1) if used_abs else 0
    num_id = (max(used_num) + 1) if used_num else 1

    ea = hn.get("east_asia_font") or "宋体"
    asc = hn.get("ascii_font") or "Times New Roman"

    absEl = el("w:abstractNum", **{"w:abstractNumId": abs_id})
    absEl.append(el("w:multiLevelType", **{"w:val": "multilevel"}))
    fmts = [(lc.get("num_fmt", "decimal") if isinstance(lc, dict) else "decimal")
            for lc in levels_cfg]
    for lvl in range(levels):
        lc = levels_cfg[lvl] or {}
        l = el("w:lvl", **{"w:ilvl": lvl})
        l.append(el("w:start", **{"w:val": str(lc.get("start", 1))}))
        l.append(el("w:numFmt", **{"w:val": lc.get("num_fmt", "decimal")}))
        l.append(el("w:pStyle", **{"w:val": "Heading%d" % (lvl + 1)}))
        txt = lc.get("format") or ".".join("%%%d" % (i + 1) for i in range(lvl + 1))
        # 本级引用了上级序号，而上级不是十进制时，Word 会渲染成「一.1」。
        # isLgl 强制所有被引用的序号按阿拉伯数字渲染 -> 「1.1」。
        refs_earlier = [int(m) for m in re.findall(r"%(\d)", txt) if 1 <= int(m) <= lvl]
        if any(fmts[m - 1] != "decimal" for m in refs_earlier):
            l.append(el("w:isLgl"))
        # 序号以标点结尾时后面不跟空格：「一、标题」而不是「一、 标题」
        if "suffix" in lc:
            suff = lc["suffix"]
        elif txt and txt[-1] in "、。．.）)】》":
            suff = "nothing"
        else:
            suff = "space"
        l.append(el("w:suff", **{"w:val": suff}))
        l.append(el("w:lvlText", **{"w:val": txt}))
        l.append(el("w:lvlJc", **{"w:val": "left"}))
        pPr = el("w:pPr")
        ind_chars = lc.get("indent_chars", 0)
        half = int(ind_chars * 100)
        pPr.append(el("w:ind", **{"w:left": str(int(ind_chars * 240)),
                                  "w:leftChars": str(half),
                                  "w:firstLine": "0", "w:firstLineChars": "0"}))
        l.append(pPr)
        rPr = el("w:rPr")
        rPr.append(el("w:rFonts", **{"w:ascii": asc, "w:hAnsi": asc,
                                     "w:eastAsia": ea, "w:hint": "eastAsia"}))
        rPr.append(el("w:i", **{"w:val": "0"}))
        rPr.append(el("w:iCs", **{"w:val": "0"}))
        if lc.get("bold", hn.get("bold")):
            rPr.append(el("w:b"))
        l.append(rPr)
        absEl.append(l)

    first_num = num_root.find(wq("num"))
    if first_num is not None:
        first_num.addprevious(absEl)
    else:
        num_root.append(absEl)
    numEl = el("w:num", **{"w:numId": num_id})
    numEl.append(el("w:abstractNumId", **{"w:val": abs_id}))
    num_root.append(numEl)

    for lvl in range(levels):
        st = next((s for s in root.findall(wq("style"))
                   if s.get(wq("styleId")) == "Heading%d" % (lvl + 1)), None)
        if st is None:
            continue
        pPr = st.find(wq("pPr"))
        if pPr is None:
            pPr = el("w:pPr"); st.append(pPr)
        old = pPr.find(wq("numPr"))
        if old is not None:
            pPr.remove(old)
        numPr = el("w:numPr")
        numPr.append(el("w:ilvl", **{"w:val": lvl}))
        numPr.append(el("w:numId", **{"w:val": num_id}))
        pPr.append(numPr)
        reorder(pPr, PPR_ORDER)


# ---------------- 5. 表格线框 / 表头 / 对齐 ----------------

PRESETS = {
    "three_line": {"top": {"sz": 12}, "bottom": {"sz": 12},
                   "header_bottom": {"sz": 6},
                   "left": {"val": "none"}, "right": {"val": "none"},
                   "insideH": {"val": "none"}, "insideV": {"val": "none"}},
    "full_grid": {"top": {"sz": 6}, "bottom": {"sz": 6},
                  "left": {"sz": 6}, "right": {"sz": 6},
                  "insideH": {"sz": 6}, "insideV": {"sz": 6},
                  "header_bottom": {"sz": 6}},
    "outline": {"top": {"sz": 8}, "bottom": {"sz": 8},
                "left": {"sz": 8}, "right": {"sz": 8},
                "insideH": {"val": "none"}, "insideV": {"val": "none"}},
    "none": {"top": {"val": "none"}, "bottom": {"val": "none"},
             "left": {"val": "none"}, "right": {"val": "none"},
             "insideH": {"val": "none"}, "insideV": {"val": "none"}},
}


def build_border(parent_tag, spec, default_color="auto"):
    e = el(parent_tag)
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        if side not in spec:
            continue
        s = spec[side] or {}
        val = s.get("val", "single")
        b = el("w:" + side)
        b.set(qn("w:val"), val)
        b.set(qn("w:sz"), str(s.get("sz", 6)))
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), s.get("color", default_color))
        e.append(b)
    return e


def table_anchors(body):
    """把每张表和它的题注锚点对应起来（pandoc 把表注放在表前面）"""
    out = []
    pending = []
    for el in body:
        if el.tag == qn("w:p"):
            sid = style_id(el)
            if sid == "TableCaption":
                pending = [b.get(qn("w:name")) for b in el.iter(qn("w:bookmarkStart"))
                           if b.get(qn("w:name"))]
            else:
                pending = []
        elif el.tag == qn("w:tbl"):
            out.append((el, pending))
            pending = []
    return out


def autofit_widths(tbl, total_twips, min_ratio=0.08):
    """按每列文字长度分配列宽：先按内容比例，再保证每列不小于 min_ratio"""
    rows = tbl.findall(qn("w:tr"))
    if not rows:
        return None
    ncol = 0
    for tr in rows:
        ncol = max(ncol, len(tr.findall(qn("w:tc"))))
    if not ncol:
        return None
    weight = []
    for ci in range(ncol):
        longest = 1
        for tr in rows:
            cells = tr.findall(qn("w:tc"))
            if ci < len(cells):
                t = "".join(x.text or "" for x in cells[ci].iter(qn("w:t")))
                longest = max(longest, len(t.strip()))
        weight.append(longest)
    # 用「长度开方 + 基数」压制极端值：否则一个长文本列会吃掉 70% 版面，
    # 序号/名称这种短列会被压到只剩几个字宽。
    weight = [(max(w, 2) ** 0.5) + 3.0 for w in weight]
    total = float(sum(weight)) or 1.0
    ratios = [max(min_ratio, w / total) for w in weight]
    s = sum(ratios)
    return [r / s for r in ratios]


def set_col_widths(tbl, ratios, total_twips):
    """按比例设置列宽（需要固定表格布局）"""
    ratios = [float(x) for x in ratios]
    total = sum(ratios) or 1.0
    widths = [int(total_twips * r / total) for r in ratios]
    # 消掉四舍五入误差
    widths[-1] += total_twips - sum(widths)
    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        cols = grid.findall(qn("w:gridCol"))
        for i, gc in enumerate(cols):
            if i < len(widths):
                gc.set(qn("w:w"), str(widths[i]))
    for tr in tbl.findall(qn("w:tr")):
        cells = tr.findall(qn("w:tc"))
        for i, tcEl in enumerate(cells):
            if i >= len(widths):
                break
            tcPr = tcEl.find(qn("w:tcPr"))
            if tcPr is None:
                tcPr = el("w:tcPr"); tcEl.insert(0, tcPr)
            old = tcPr.find(qn("w:tcW"))
            if old is not None:
                tcPr.remove(old)
            tcW = el("w:tcW", **{"w:w": str(widths[i]), "w:type": "dxa"})
            tcPr.insert(0, tcW)
            # 跨列单元格：宽度按跨了几列累加
            span = tcPr.find(qn("w:gridSpan"))
            if span is not None:
                n = int(span.get(qn("w:val")) or 1)
                if i + n <= len(widths):
                    tcW.set(qn("w:w"), str(sum(widths[i:i + n])))


def apply_tables(doc, cfg):
    tc = cfg.get("table") or {}
    preset = tc.get("preset", "three_line")
    spec = dict(PRESETS.get(preset, PRESETS["three_line"]))
    for k, v in (tc.get("borders") or {}).items():
        spec[k] = v
    header_cfg = {"bold": True, "repeat": True, "valign": "center",
                  "shading": None, "align": None}
    header_cfg.update(tc.get("header") or {})
    cell_valign = (tc.get("cell") or {}).get("valign", "center")
    indent = (tc.get("cell") or {}).get("margin_twips")
    full_width = tc.get("full_width", True)
    cant_split = tc.get("cant_split", False)
    col_widths = tc.get("col_widths") or {}
    per_table = tc.get("per_table") or {}
    width_twips = int((doc.sections[0].page_width - doc.sections[0].left_margin
                       - doc.sections[0].right_margin) / 635)

    for tbl, anchors in table_anchors(doc.element.body):
        # 单表覆盖
        over = None
        for a in anchors:
            if a in per_table:
                over = per_table[a]
                break
        t_spec = spec
        t_header = header_cfg
        t_full = full_width
        t_cant = cant_split
        t_widths = None
        if over:
            t_spec = dict(spec)
            for k, v in (over.get("borders") or {}).items():
                t_spec[k] = v
            if over.get("preset"):
                t_spec = dict(PRESETS.get(over["preset"], spec))
            t_header = dict(header_cfg)
            t_header.update(over.get("header") or {})
            t_full = over.get("full_width", full_width)
            t_cant = over.get("cant_split", cant_split)
            t_widths = over.get("col_widths")
        if t_widths is None:
            for a in anchors:
                if a in col_widths:
                    t_widths = col_widths[a]
                    break
        if t_widths == "auto" or (t_widths is None and tc.get("col_widths_mode") == "auto"):
            t_widths = autofit_widths(tbl, width_twips)

        tblPr = tbl.find(qn("w:tblPr"))
        if tblPr is None:
            tblPr = el("w:tblPr"); tbl.insert(0, tblPr)

        if t_full:
            tw = tblPr.find(qn("w:tblW"))
            if tw is None:
                tw = el("w:tblW"); tblPr.append(tw)
            tw.set(qn("w:type"), "pct"); tw.set(qn("w:w"), "5000")

        old = tblPr.find(qn("w:tblBorders"))
        if old is not None:
            tblPr.remove(old)
        tblPr.append(build_border("w:tblBorders", t_spec))

        if indent is not None:
            mar = el("w:tblCellMar")
            for side in ("left", "right"):
                m = el("w:" + side)
                m.set(qn("w:w"), str(indent)); m.set(qn("w:type"), "dxa")
                mar.append(m)
            old = tblPr.find(qn("w:tblCellMar"))
            if old is not None:
                tblPr.remove(old)
            tblPr.append(mar)
        reorder(tblPr, TBLPR_ORDER)

        if t_widths:
            lay = tblPr.find(qn("w:tblLayout"))
            if lay is None:
                lay = el("w:tblLayout"); tblPr.append(lay)
            lay.set(qn("w:type"), "fixed")
            reorder(tblPr, TBLPR_ORDER)
            set_col_widths(tbl, t_widths, width_twips)

        rows = tbl.findall(qn("w:tr"))
        for ri, tr in enumerate(rows):
            is_head = (ri == 0)
            trPr = tr.find(qn("w:trPr"))
            if trPr is None:
                trPr = el("w:trPr"); tr.insert(0, trPr)
            if is_head and t_header.get("repeat"):
                if trPr.find(qn("w:tblHeader")) is None:
                    trPr.append(el("w:tblHeader"))
            if t_cant and trPr.find(qn("w:cantSplit")) is None:
                trPr.append(el("w:cantSplit"))
            for tcEl in tr.findall(qn("w:tc")):
                # 单元格里的段落改用 TableText 样式（与列表用的 Compact 分开），
                # 这样才能「列表跟正文同号、表格小一号」
                if tc.get("use_table_style", True):
                    for cp in tcEl.findall(qn("w:p")):
                        cpPr = cp.find(qn("w:pPr"))
                        if cpPr is None:
                            cpPr = el("w:pPr"); cp.insert(0, cpPr)
                        old_st = cpPr.find(qn("w:pStyle"))
                        if old_st is not None:
                            cpPr.remove(old_st)
                        cpPr.insert(0, el("w:pStyle", **{"w:val": "TableText"}))
                        reorder(cpPr, PPR_ORDER)
                tcPr = tcEl.find(qn("w:tcPr"))
                if tcPr is None:
                    tcPr = el("w:tcPr"); tcEl.insert(0, tcPr)
                old = tcPr.find(qn("w:vAlign"))
                if old is not None:
                    tcPr.remove(old)
                va = t_header.get("valign") if is_head else cell_valign
                if va:
                    tcPr.append(el("w:vAlign", **{"w:val": va}))
                if is_head:
                    if t_header.get("shading"):
                        old = tcPr.find(qn("w:shd"))
                        if old is not None:
                            tcPr.remove(old)
                        tcPr.append(el("w:shd", **{"w:val": "clear",
                                                   "w:color": "auto",
                                                   "w:fill": str(header_cfg["shading"]).lstrip("#")}))
                    if "header_bottom" in spec:
                        old = tcPr.find(qn("w:tcBorders"))
                        if old is not None:
                            tcPr.remove(old)
                        tcPr.append(build_border("w:tcBorders",
                                                 {"bottom": spec["header_bottom"]}))
                reorder(tcPr, TCPR_ORDER)
                for p in tcEl.findall(qn("w:p")):
                    pPr = p.find(qn("w:pPr"))
                    if pPr is None:
                        pPr = el("w:pPr"); p.insert(0, pPr)
                    old = pPr.find(qn("w:jc"))
                    if old is not None:
                        pPr.remove(old)
                    if is_head and header_cfg.get("align"):
                        pPr.append(el("w:jc", **{"w:val": header_cfg["align"]}))
                    reorder(pPr, PPR_ORDER)
                    if is_head and header_cfg.get("bold") is not None:
                        for r in p.findall(qn("w:r")):
                            rPr = r.find(qn("w:rPr"))
                            if rPr is None:
                                rPr = el("w:rPr"); r.insert(0, rPr)
                            b = rPr.find(qn("w:b"))
                            if b is None:
                                b = el("w:b"); rPr.append(b)
                            if header_cfg["bold"]:
                                b.attrib.pop(qn("w:val"), None)
                                b.set(qn("w:val"), "1")
                            else:
                                b.set(qn("w:val"), "0")


# ---------------- 6. 图片 / 斜体 ----------------


def landscape_tables(doc, cfg):
    """把指定的表格（按题注锚点名）单独放到横向页里

    profile 里写： "landscape": ["tbl:wide"]
    """
    targets = cfg.get("landscape") or []
    if not targets:
        return 0
    body = doc.element.body
    base_sect = body.find(qn("w:sectPr"))
    if base_sect is None:
        return 0
    sec = doc.sections[-1]
    width, height = sec.page_width, sec.page_height
    made = 0
    for anchor in targets:
        key = anchor if ":" in anchor else "tbl:" + anchor
        # 找到题注段落
        cap = None
        for p in body.iter(qn("w:p")):
            for b in p.iter(qn("w:bookmarkStart")):
                if b.get(qn("w:name")) == key:
                    cap = p
                    break
            if cap is not None:
                break
        if cap is None:
            continue
        # 找那张表：pandoc 把表注放在表「前面」，所以先往后找
        tbl = None
        sib = cap.getnext()
        while sib is not None:
            if sib.tag == qn("w:tbl"):
                tbl = sib
                break
            sib = sib.getnext()
        if tbl is None:                 # 兼容表注在表后面的写法
            sib = cap.getprevious()
            while sib is not None:
                if sib.tag == qn("w:tbl"):
                    tbl = sib
                    break
                sib = sib.getprevious()
        if tbl is None:
            continue

        def make_sect(landscape):
            s = copy.deepcopy(base_sect)
            for tag in ("w:headerReference", "w:footerReference", "w:titlePg"):
                for e in s.findall(qn(tag)):
                    s.remove(e)
            pg = s.find(qn("w:pgSz"))
            if pg is None:
                pg = el("w:pgSz"); s.append(pg)
            if landscape:
                pg.set(qn("w:w"), str(int(height / 635)))
                pg.set(qn("w:h"), str(int(width / 635)))
                pg.set(qn("w:orient"), "landscape")
            else:
                pg.set(qn("w:w"), str(int(width / 635)))
                pg.set(qn("w:h"), str(int(height / 635)))
                if pg.get(qn("w:orient")):
                    del pg.attrib[qn("w:orient")]
            return s

        # 表前：结束上一节（纵向）。
        # 注意要插在「表注段落」之前而不是表之前，
        # 否则表注和表被分节符隔开，keepNext 失效，表注会单独留在一页。
        before = el("w:p")
        pPr = el("w:pPr"); pPr.append(make_sect(False)); before.append(pPr)
        cap.addprevious(before)
        # 表后：结束横向节
        after = el("w:p")
        pPr = el("w:pPr"); pPr.append(make_sect(True)); after.append(pPr)
        tbl.addnext(after)
        made += 1
    return made


def apply_drop_cap(doc, cfg):
    """首字下沉：把正文第一段的头 N 个字拆成独立的框"""
    dc = cfg.get("drop_cap") or {}
    if not dc.get("enabled"):
        return False
    lines = int(dc.get("lines", 3))
    chars = int(dc.get("chars", 1))
    body = doc.element.body
    # 找正文第一段（跳过封面/目录这类节内的段落）
    target = None
    for p in body.iter(qn("w:p")):
        sid = style_id(p)
        if sid in ("FirstParagraph", "BodyText") and para_text(p):
            target = p
            break
    if target is None:
        return False
    runs = target.findall(qn("w:r"))
    if not runs:
        return False
    text = "".join(t.text or "" for r in runs for t in r.iter(qn("w:t")))
    if len(text) <= chars:
        return False
    # 把前 chars 个字拆出来
    head_text = text[:chars]
    rest_text = text[chars:]
    target_text = "".join(t.text or "" for t in target.iter(qn("w:t")))
    idx = target_text.find(head_text)
    # 重建：先一个框段，再剩下的
    cap = el("w:p")
    pPr = el("w:pPr")
    src_pPr = target.find(qn("w:pPr"))
    if src_pPr is not None:
        st = src_pPr.find(qn("w:pStyle"))
        if st is not None:
            pPr.append(el("w:pStyle", **{"w:val": st.get(qn("w:val"))}))
    frame = el("w:framePr", **{
        "w:dropCap": "drop", "w:lines": str(lines), "w:wrap": "around",
        "w:vAnchor": "text", "w:hAnchor": "text"})
    pPr.append(frame)
    cap.append(pPr)
    cap.append(trun(head_text))
    # 剩下的字回填到原段
    for r in runs:
        target.remove(r)
    target.append(trun(rest_text))
    target.addprevious(cap)
    return True


def apply_drop_cap_wrapper(doc, cfg):
    return apply_drop_cap(doc, cfg)


def clamp_images(doc):
    limit = int((doc.sections[0].page_width - doc.sections[0].left_margin
                 - doc.sections[0].right_margin))
    WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
    A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    n = 0
    for ext in doc.element.body.iter(WP + "extent"):
        cx, cy = int(ext.get("cx")), int(ext.get("cy"))
        if cx > limit > 0:
            ratio = limit / float(cx)
            ext.set("cx", str(limit)); ext.set("cy", str(int(cy * ratio)))
            n += 1
            for pic in doc.element.body.iter(A + "ext"):
                if int(pic.get("cx") or 0) == cx:
                    pic.set("cx", str(limit)); pic.set("cy", str(int(cy * ratio)))
    return n


def kill_italics(doc):
    roots = [doc.styles.element, doc.element.body]
    for s in doc.sections:
        for hf in (s.header, s.footer, s.first_page_header, s.first_page_footer):
            try:
                if hf.is_linked_to_previous:
                    continue
                roots.append(hf._element)
            except Exception:
                pass
    for root in roots:
        for e in root.iter():
            if e.tag in (qn("w:i"), qn("w:iCs")):
                e.set(qn("w:val"), "0")


# ---------------- 配置加载 ----------------

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


def load_profile(path=None):
    path = path or DEFAULT_PROFILE
    if not os.path.isabs(path):
        cand = os.path.join(HERE, path)
        if os.path.exists(cand):
            path = cand
        elif os.path.exists(os.path.join(HERE, "profiles", path + ".json")):
            path = os.path.join(HERE, "profiles", path + ".json")
    if not os.path.exists(path):
        print("找不到 profile: %s" % path, file=sys.stderr)
        return {}
    return json.loads(strip_json_comments(open(path, encoding="utf-8").read()))


# ---------------- main ----------------

def main():
    argv = sys.argv[1:]
    meta = {}
    if "--meta" in argv:
        i = argv.index("--meta")
        try:
            meta = json.loads(argv[i + 1])
        except (IndexError, ValueError):
            meta = {}
        del argv[i:i + 2]
    if "--info" in argv:
        i = argv.index("--info")
        print(load_profile(argv[i + 1]).get("name", ""))
        return 0
    profile = None
    if "--profile" in argv:
        i = argv.index("--profile")
        profile = argv[i + 1]
        del argv[i:i + 2]
    flags = [a for a in argv if a.startswith("--")]
    args = [a for a in argv if not a.startswith("--")]
    if len(args) < 2:
        print(__doc__)
        return 1

    cfg = load_profile(profile)
    doc = Document(args[0])

    apply_list_numbering(doc, cfg)
    apply_document_settings(doc, cfg)

    if meta:
        cfg["_meta"] = meta

    bmnum = upgrade(doc, cfg)
    headings = collect_headings(doc, cfg)
    refs(doc, bmnum, cfg, headings)
    n_bib, n_cite = auto_number_bibliography(doc, cfg)
    if n_bib:
        print("  · 参考文献编号已改为 SEQ 域：%d 条书目 / %d 处角标"
              % (n_bib, n_cite), file=sys.stderr)
        if n_cite == 0:
            print("    正文角标没换成 REF 域（需要在 md 头部设 "
                  "link-citations: true 才能定位条目）", file=sys.stderr)
    # 先做分节（封面/目录/正文），再做横排页：
    # 反过来的话，横排插进来的节会让分节数对不上，页码和页眉就都不生成了
    apply_page_structure(doc, cfg)
    landscape_tables(doc, cfg)
    apply_drop_cap(doc, cfg)
    if "--no-heading-number" not in flags:
        heading_numbering(doc, cfg)
        ap_num_id = build_appendix_numbering(doc, cfg)
        apply_appendix(doc, cfg, ap_num_id)
    apply_tables(doc, cfg)
    clamp_images(doc)
    kill_italics(doc)

    settings = doc.settings.element
    if settings.find(qn("w:updateFields")) is None:
        insert_in_order(settings, el("w:updateFields", **{"w:val": "true"}),
                        SETTINGS_ORDER)

    # 最后一步：把所有容器的子元素按 OOXML schema 顺序排好。
    # Word 对顺序极严，顺序错了会弹「文件已损坏，是否修复」，
    # 而 LibreOffice / python-docx 都能正常打开（所以必须自己检查）。
    n_fix = normalize_structure(doc)
    if n_fix:
        print("  · 规范化了 %d 处元素顺序" % n_fix)
    validate(doc)

    doc.save(args[1])
    print("✓ %s" % args[1])
    return 0


if __name__ == "__main__":
    sys.exit(main())
