#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_template.py —— 从 style.json 生成 pandoc 用的 reference.docx

用法：
    python3 build_template.py                      # 用同目录 style.json
    python3 build_template.py my.json -o out.docx  # 指定配置和输出

所有样式（字体/字号/粗细/斜体/对齐/缩进/行距/段前后距）、页边距、
页眉页脚、标题自动编号、三线表样式，全部由 style.json 控制。
改完重新跑一次即可，不需要打开 Word。
"""
import json, os, re, shutil, subprocess, sys, copy, tempfile

from docx import Document
from docx.shared import Cm, Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

HERE = os.path.dirname(os.path.abspath(__file__))

# 中文字号 -> 磅
CN_SIZES = {
    "初号": 42, "小初": 36, "一号": 26, "小一": 24, "二号": 22, "小二": 18,
    "三号": 16, "小三": 15, "四号": 14, "小四": 12, "五号": 10.5, "小五": 9,
    "六号": 7.5, "小六": 6.5, "七号": 5.5, "八号": 5,
}

# 永远不允许斜体的样式（用户明确要求：全文不要斜体）
NEVER_ITALIC = True

# 除 config 明确指定外，这些样式必须显式清零首行缩进
# （因为它们都继承自 BodyText 或 Normal，不清零会被带着缩进）
NO_INDENT_DEFAULT = [
    "Normal", "Caption", "TableCaption", "ImageCaption", "Compact",
    "BlockText", "SourceCode", "Definition", "DefinitionTerm", "Title",
    "Subtitle", "Author", "Date", "Abstract", "AbstractTitle", "FootnoteText",
    "FootnoteBlockText", "TOCHeading", "Bibliography", "TableCaption", "Figure",
    "CaptionedFigure", "TOC1", "TOC2", "TOC3", "TOC4", "TOC5", "TOC6",
    "TOC7", "TOC8", "TOC9", "TableofFigures",
] + [f"Heading{i}" for i in range(1, 10)]

# 编译期的上下文：apply_ppr 里写制表位需要知道版心宽度
CTX = {}

SECTPR_ORDER = ["headerReference", "footerReference", "footnotePr", "endnotePr",
                "type", "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType",
                "pgNumType", "cols", "formProt", "vAlign", "noEndnote",
                "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid",
                "printerSettings", "sectPrChange"]

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

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WQ = lambda t: "{%s}%s" % (W, t)


# ---------- 基础工具 ----------

def strip_json_comments(text):
    """去掉 JSON 里的 // 和 /* */ 注释（认识字符串，不会误伤 http://）"""
    out, i, n = [], 0, len(text)
    in_str = False
    while i < n:
        c = text[i]
        if in_str:
            out.append(c)
            if c == "\\" and i + 1 < n:
                out.append(text[i + 1]); i += 2; continue
            if c == '"':
                in_str = False
            i += 1
            continue
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
            i += 2
            continue
        out.append(c); i += 1
    return "".join(out)


def size_to_pt(v):
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if s in CN_SIZES:
        return CN_SIZES[s]
    m = re.match(r"^([\d.]+)\s*(pt|磅)?$", s)
    if m:
        return float(m.group(1))
    raise ValueError("看不懂的字号: %r" % v)


def twips(pt_or_cm, unit="pt"):
    if unit == "pt":
        return int(round(float(pt_or_cm) * 20))
    return int(round(float(pt_or_cm) * 567))  # 1cm = 567 twips


def find_style(root, style_id):
    for s in root.findall(WQ("style")):
        if s.get(WQ("styleId")) == style_id:
            return s
    return None


def get_or_make(parent, tag, order=None, after=None):
    e = parent.find(WQ(tag))
    if e is not None:
        return e
    e = OxmlElement("w:" + tag)
    if after is not None:
        idx = list(parent).index(after) + 1
        parent.insert(idx, e)
    else:
        parent.append(e)
    return e


# w:style 子元素合法顺序
STYLE_ORDER = ["name", "aliases", "basedOn", "next", "link", "autoRedefine",
               "hidden", "uiPriority", "semiHidden", "unhideWhenUsed", "qFormat",
               "locked", "personal", "personalCompose", "personalReply", "rsid",
               "pPr", "rPr", "tblPr", "trPr", "tcPr", "tblStylePr"]
# w:pPr 子元素合法顺序（只列我们用的）
PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
             "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
             "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
             "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
             "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
             "suppressOverlap", "jc", "textDirection", "textAlignment",
             "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
             "sectPr", "pPrChange"]
# w:tblPr 子元素合法顺序
TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
               "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd",
               "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook",
               "tblCaption", "tblDescription"]
# w:rPr 子元素合法顺序
RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps",
             "strike", "dstrike", "outline", "shadow", "emboss", "imprint",
             "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing",
             "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
             "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang",
             "eastAsianLayout", "specVanish", "oMath"]


def reorder(el, order):
    """按 schema 顺序重排子元素"""
    idx = {t: i for i, t in enumerate(order)}
    kids = list(el)
    kids.sort(key=lambda k: idx.get(k.tag.replace(WQ(""), ""), 999))
    for k in kids:
        el.append(k)


def set_child(parent, tag, order, **attrs):
    e = get_or_make(parent, tag, after=None)
    if e.getparent() is None:
        parent.append(e)
    for k, v in attrs.items():
        e.set(WQ(k), str(v))
    reorder(parent, order)
    return e


# ---------- 样式写入 ----------

def insert_in_order(parent, elem, order):
    """把 elem 按 schema 顺序插到 parent 里（同名元素排在其后）"""
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


def ensure_style(root, style_id, pretty_name, kind="paragraph"):
    """reference.docx 里缺的样式（如 SourceCode，pandoc 会自己注入）就地造一个"""
    st = find_style(root, style_id)
    if st is not None:
        return st
    st = OxmlElement("w:style")
    st.set(WQ("type"), kind)
    st.set(WQ("customStyle"), "1")
    st.set(WQ("styleId"), style_id)
    name = OxmlElement("w:name"); name.set(WQ("val"), pretty_name)
    st.append(name)
    base = OxmlElement("w:basedOn"); base.set(WQ("val"), "Normal")
    st.append(base)
    root.append(st)
    return st


def apply_rpr(style_el, p):
    rPr = style_el.find(WQ("rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        pPr = style_el.find(WQ("pPr"))
        if pPr is not None:
            pPr.addnext(rPr)
        else:
            # 插到 rPr 该在的位置：pPr 后 / tblPr 前
            tbl = style_el.find(WQ("tblPr"))
            if tbl is not None:
                tbl.addprevious(rPr)
            else:
                style_el.append(rPr)

    ea = p.get("east_asia_font")
    ascii_ = p.get("ascii_font")
    if ea or ascii_:
        rf = rPr.find(WQ("rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts")
            rPr.insert(0, rf)
        # 主题字体优先级高于显式字体，必须删掉
        for a in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
            if rf.get(WQ(a)) is not None:
                del rf.attrib[WQ(a)]
        if ascii_:
            rf.set(WQ("ascii"), ascii_)
            rf.set(WQ("hAnsi"), ascii_)
        if ea:
            rf.set(WQ("eastAsia"), ea)
            rf.set(WQ("cs"), ascii_ or ea)
            rf.set(WQ("hint"), "eastAsia")

    if p.get("size_pt") is not None:
        half = str(int(round(size_to_pt(p["size_pt"]) * 2)))
        for tag in ("sz", "szCs"):
            e = rPr.find(WQ(tag))
            if e is None:
                e = OxmlElement("w:" + tag)
                rPr.append(e)
            e.set(WQ("val"), half)

    if p.get("color"):
        c = set_child(rPr, "color", RPR_ORDER, val=str(p["color"]).lstrip("#"))
        # 显式设色就要把主题色属性去掉，否则 Word 仍按主题色渲染
        for a in ("themeColor", "themeShade", "themeTint"):
            if c.get(WQ(a)) is not None:
                del c.attrib[WQ(a)]

    # 粗体
    if p.get("bold") is not None:
        for tag in ("b", "bCs"):
            e = rPr.find(WQ(tag))
            if e is None:
                e = OxmlElement("w:" + tag)
                rPr.append(e)
            if p["bold"]:
                e.attrib.pop(WQ("val"), None)
            else:
                e.set(WQ("val"), "0")

    # 斜体：一律显式关闭
    if NEVER_ITALIC or p.get("italic") is False:
        for tag in ("i", "iCs"):
            e = rPr.find(WQ(tag))
            if e is None:
                e = OxmlElement("w:" + tag)
                rPr.append(e)
            e.set(WQ("val"), "0")
        # 清掉默认打开形式
        for tag in ("i", "iCs"):
            e = rPr.find(WQ(tag))
            if e is not None and e.get(WQ("val")) is None:
                e.set(WQ("val"), "0")
    reorder(rPr, RPR_ORDER)


def apply_ppr(style_el, p):
    pPr = style_el.find(WQ("pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        name = style_el.find(WQ("name"))
        if name is not None:
            name.addnext(pPr)
        else:
            style_el.insert(0, pPr)

    def ensure(tag):
        e = pPr.find(WQ(tag))
        if e is None:
            e = OxmlElement("w:" + tag)
            pPr.append(e)
        return e

    # 对齐
    if p.get("align"):
        ensure("jc").set(WQ("val"), {"left": "left", "center": "center",
                                     "right": "right", "justify": "both"}[p["align"]])

    # 行距 / 段前后
    need_sp = any(k in p for k in ("line_spacing", "line_spacing_pt",
                                   "space_before_pt", "space_after_pt"))
    if need_sp:
        sp = ensure("spacing")
        if p.get("line_spacing_pt") is not None:      # 固定行距，如公文 28.8 磅
            sp.set(WQ("line"), str(int(round(float(p["line_spacing_pt"]) * 20))))
            sp.set(WQ("lineRule"), "exact")
        elif p.get("line_spacing") is not None:       # 倍数行距
            sp.set(WQ("line"), str(int(round(float(p["line_spacing"]) * 240))))
            sp.set(WQ("lineRule"), "auto")
        if p.get("space_before_pt") is not None:
            sp.set(WQ("before"), str(twips(p["space_before_pt"])))
            sp.set(WQ("beforeAutospacing"), "0")
        if p.get("space_after_pt") is not None:
            sp.set(WQ("after"), str(twips(p["space_after_pt"])))
            sp.set(WQ("afterAutospacing"), "0")

    # 缩进
    need_ind = any(k in p for k in ("first_line_indent_chars", "left_indent_chars",
                                    "right_indent_chars", "hanging_indent_chars"))
    if need_ind:
        ind = ensure("ind")
        base_pt = size_to_pt(p.get("size_pt", 12))
        if "first_line_indent_chars" in p:
            ch = float(p["first_line_indent_chars"])
            ind.set(WQ("firstLineChars"), str(int(ch * 100)))
            ind.set(WQ("firstLine"), str(int(round(ch * base_pt * 20))))
        if "left_indent_chars" in p:
            ch = float(p["left_indent_chars"])
            ind.set(WQ("leftChars"), str(int(ch * 100)))
            ind.set(WQ("left"), str(int(round(ch * base_pt * 20))))
        if "right_indent_chars" in p:
            ch = float(p["right_indent_chars"])
            ind.set(WQ("rightChars"), str(int(ch * 100)))
            ind.set(WQ("right"), str(int(round(ch * base_pt * 20))))
        if "hanging_indent_chars" in p:
            ch = float(p["hanging_indent_chars"])
            ind.set(WQ("hangingChars"), str(int(ch * 100)))
            ind.set(WQ("hanging"), str(int(round(ch * base_pt * 20))))

    # 制表位：目录条目要靠它把页码右对齐到版心右边界并加前导点线
    tabs_cfg = p.get("tabs")
    if tabs_cfg:
        tabs = ensure("tabs")
        for old in tabs.findall(WQ("tab")):
            tabs.remove(old)
        for t in tabs_cfg:
            val = t.get("val", "right")
            # pos 可以给 twips 数字，也可以写 "text" 表示版心右边界
            if str(t.get("pos")) == "text":
                pos = CTX.get("text_width_twips", 8306)
            else:
                pos = int(t.get("pos", 0))
            if t.get("indent_chars_pos"):
                pos = int(float(t["indent_chars_pos"]) *
                          size_to_pt(p.get("size_pt", 12)) * 20)
            e = OxmlElement("w:tab")
            e.set(WQ("val"), val)
            if t.get("leader"):
                e.set(WQ("leader"), t["leader"])
            e.set(WQ("pos"), str(pos))
            tabs.append(e)

    if p.get("shading"):
        shd = ensure("shd")
        shd.set(WQ("val"), "clear")
        shd.set(WQ("color"), "auto")
        shd.set(WQ("fill"), str(p["shading"]).lstrip("#"))

    border = p.get("left_border_sz")
    if border:
        pbdr = ensure("pBdr")
        for old in pbdr.findall(WQ("left")):      # 同一边只允许一个
            pbdr.remove(old)
        left = OxmlElement("w:left")
        left.set(WQ("val"), "single"); left.set(WQ("sz"), str(int(border)))
        left.set(WQ("space"), "8"); left.set(WQ("color"), "C8C8C8")
        pbdr.append(left)
        reorder(pbdr, ["top", "left", "bottom", "right", "between", "bar"])

    if p.get("keep_next"):
        ensure("keepNext")
    if p.get("keep_lines"):
        ensure("keepLines")
    if p.get("page_break_before"):
        ensure("pageBreakBefore")

    # 中文排版开关
    if p.get("widow_control"):
        ensure("widowControl")
    if p.get("kinsoku"):
        ensure("kinsoku")
    if p.get("overflow_punct"):
        ensure("overflowPunct")
    if p.get("auto_space_de"):
        ensure("autoSpaceDE")
    if p.get("auto_space_dn"):
        ensure("autoSpaceDN")
    if p.get("word_wrap") is False:
        ensure("wordWrap").set(WQ("val"), "0")

    reorder(pPr, PPR_ORDER)


def zero_indent(style_el):
    pPr = style_el.find(WQ("pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        name = style_el.find(WQ("name"))
        (name.addnext(pPr) if name is not None else style_el.insert(0, pPr))
    ind = pPr.find(WQ("ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        pPr.append(ind)
    ind.set(WQ("firstLine"), "0")
    ind.set(WQ("firstLineChars"), "0")
    reorder(pPr, PPR_ORDER)


# ---------- 各功能模块 ----------

def detach_heading_based(root):
    """TOCHeading 等样式若继承 Heading1，会把继承来的编号带出来（目录标题被编号成 1），
    这里把它们改挂到 Normal 上，并显式取消编号。"""
    for st in root.findall(WQ("style")):
        sid = st.get(WQ("styleId")) or ""
        if not sid.startswith("TOC"):
            continue
        based = st.find(WQ("basedOn"))
        if based is not None:
            based.set(WQ("val"), "Normal")
        pPr = st.find(WQ("pPr"))
        if pPr is None:
            pPr = OxmlElement("w:pPr")
            st.append(pPr)
        old = pPr.find(WQ("numPr"))
        if old is not None:
            pPr.remove(old)
        numPr = OxmlElement("w:numPr")
        ilvl = OxmlElement("w:ilvl"); ilvl.set(WQ("val"), "0"); numPr.append(ilvl)
        numId = OxmlElement("w:numId"); numId.set(WQ("val"), "0"); numPr.append(numId)
        pPr.insert(0, numPr)
        reorder(pPr, PPR_ORDER)


def neutralize_theme_colors(root, keep=("Hyperlink",)):
    """pandoc 默认模板里的 Heading1-9 带主题色（accent1 蓝），正式报告不该是蓝的。

    这里把所有带 w:themeColor 的颜色改成纯黑；keep 里的样式（如超链接）保留原样。
    用户在 profile 里显式写了 color 的样式不受影响（下面会跳过）。
    """
    n = 0
    for st in root.findall(WQ("style")):
        sid = st.get(WQ("styleId")) or ""
        if sid in keep:
            continue
        rPr = st.find(WQ("rPr"))
        if rPr is None:
            continue
        c = rPr.find(WQ("color"))
        if c is None:
            continue
        has_theme = any(c.get(WQ(a)) is not None
                        for a in ("themeColor", "themeShade", "themeTint"))
        if not has_theme:
            continue
        for a in ("themeColor", "themeShade", "themeTint"):
            if c.get(WQ(a)) is not None:
                del c.attrib[WQ(a)]
        c.set(WQ("val"), "000000")
        n += 1
    return n


def kill_all_italics(root):
    """全局清除斜体：所有样式 + docDefaults + 字符样式"""
    for el in root.iter():
        if el.tag in (WQ("i"), WQ("iCs")):
            el.set(WQ("val"), "0")
    dd = root.find(WQ("docDefaults"))
    if dd is not None:
        rpr = dd.find(WQ("rPrDefault") + "/" + WQ("rPr"))
        if rpr is not None:
            for tag in ("i", "iCs"):
                if rpr.find(WQ(tag)) is None:
                    e = OxmlElement("w:" + tag)
                    e.set(WQ("val"), "0")
                    rpr.append(e)
            reorder(rpr, RPR_ORDER)


def set_doc_defaults(root, cfg):
    dd = root.find(WQ("docDefaults"))
    if dd is None:
        return
    rpr = dd.find(WQ("rPrDefault") + "/" + WQ("rPr"))
    if rpr is None:
        return
    ea = cfg.get("east_asia_font")
    asc = cfg.get("ascii_font")
    if ea or asc:
        rf = rpr.find(WQ("rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts"); rpr.insert(0, rf)
        for a in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
            if rf.get(WQ(a)) is not None:
                del rf.attrib[WQ(a)]
        if asc:
            rf.set(WQ("ascii"), asc); rf.set(WQ("hAnsi"), asc)
        if ea:
            rf.set(WQ("eastAsia"), ea); rf.set(WQ("cs"), asc or ea)
            rf.set(WQ("hint"), "eastAsia")
    if cfg.get("size_pt") is not None:
        half = str(int(round(size_to_pt(cfg["size_pt"]) * 2)))
        for tag in ("sz", "szCs"):
            e = rpr.find(WQ(tag))
            if e is None:
                e = OxmlElement("w:" + tag); rpr.append(e)
            e.set(WQ("val"), half)
    reorder(rpr, RPR_ORDER)


def apply_page(doc, cfg):
    s = doc.sections[0]
    if cfg.get("width_cm"):
        s.page_width = Cm(cfg["width_cm"])
    if cfg.get("height_cm"):
        s.page_height = Cm(cfg["height_cm"])
    m = cfg.get("margin_cm", {})
    for k in ("top", "bottom", "left", "right"):
        if k in m:
            setattr(s, k + "_margin", Cm(m[k]))
    if "gutter_cm" in cfg:
        s.gutter = Cm(cfg["gutter_cm"])
    if "header_distance_cm" in cfg:
        s.header_distance = Cm(cfg["header_distance_cm"])
    if "footer_distance_cm" in cfg:
        s.footer_distance = Cm(cfg["footer_distance_cm"])
    if cfg.get("first_page_different"):
        s.different_first_page_header_footer = True
    if cfg.get("landscape"):
        s.orientation = 1        # WD_ORIENT.LANDSCAPE
        s.page_width, s.page_height = s.page_height, s.page_width
    # 文档网格（中文排版：每页行数 / 每行字数）
    grid = cfg.get("doc_grid") or {}
    if grid:
        sectPr = s._sectPr
        old = sectPr.find(WQ("docGrid"))
        if old is not None:
            sectPr.remove(old)
        g = OxmlElement("w:docGrid")
        gtype = grid.get("type", "lines")
        g.set(WQ("type"), gtype)
        if gtype == "lines":
            pitch = grid.get("line_pitch")
            if pitch is None and grid.get("lines_per_page"):
                # 行距 = 版心高度 / 每页行数
                text_h = s.page_height - s.top_margin - s.bottom_margin
                pitch = int(round(float(text_h) / 635.0 / int(grid["lines_per_page"])))
            g.set(WQ("linePitch"), str(int(pitch or 312)))
        else:
            space = grid.get("char_space")
            if space is None and grid.get("chars_per_line"):
                text_w = s.page_width - s.left_margin - s.right_margin
                # 每行字数 -> 字符间距（负值=压缩）
                per = int(round(float(text_w) / 635.0 / int(grid["chars_per_line"])))
                size = float(cfg.get("doc_default", {}).get("size_pt", 12)) * 20
                space = int(round((per - size)))
            g.set(WQ("charSpace"), str(int(space or 0)))
        sectPr.append(g)
    return s


def style_runs(par, cfg, underline_border=False):
    for r in par.runs:
        rPr = r._element.get_or_add_rPr()
        if cfg.get("ascii_font") or cfg.get("east_asia_font"):
            rf = rPr.find(WQ("rFonts"))
            if rf is None:
                rf = OxmlElement("w:rFonts"); rPr.insert(0, rf)
            for a in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
                if rf.get(WQ(a)) is not None:
                    del rf.attrib[WQ(a)]
            if cfg.get("ascii_font"):
                rf.set(WQ("ascii"), cfg["ascii_font"]); rf.set(WQ("hAnsi"), cfg["ascii_font"])
            if cfg.get("east_asia_font"):
                rf.set(WQ("eastAsia"), cfg["east_asia_font"]); rf.set(WQ("hint"), "eastAsia")
        if cfg.get("size_pt") is not None:
            half = str(int(round(size_to_pt(cfg["size_pt"]) * 2)))
            for tag in ("sz", "szCs"):
                e = rPr.find(WQ(tag))
                if e is None:
                    e = OxmlElement("w:" + tag); rPr.append(e)
                e.set(WQ("val"), half)
        for tag in ("i", "iCs"):
            e = rPr.find(WQ(tag))
            if e is None:
                e = OxmlElement("w:" + tag); rPr.append(e)
            e.set(WQ("val"), "0")
        reorder(rPr, RPR_ORDER)
    if underline_border:
        pPr = par._p.get_or_add_pPr()
        pbdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(WQ("val"), "single"); bottom.set(WQ("sz"), "6")
        bottom.set(WQ("space"), "1"); bottom.set(WQ("color"), "auto")
        pbdr.append(bottom)
        pPr.append(pbdr)
        reorder(pPr, PPR_ORDER)


def field_runs(par, instr, cached=""):
    r1 = par.add_run(); fc = OxmlElement("w:fldChar"); fc.set(WQ("fldCharType"), "begin"); r1._element.append(fc)
    r2 = par.add_run(); it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = instr; r2._element.append(it)
    r3 = par.add_run(); fc = OxmlElement("w:fldChar"); fc.set(WQ("fldCharType"), "separate"); r3._element.append(fc)
    r4 = par.add_run(cached)
    r5 = par.add_run(); fc = OxmlElement("w:fldChar"); fc.set(WQ("fldCharType"), "end"); r5._element.append(fc)


def apply_header_footer(doc, section, cfg):
    h = cfg.get("header") or {}
    f = cfg.get("footer") or {}
    align2jc = {"left": "left", "center": "center", "right": "right"}

    if h.get("text"):
        par = section.header.paragraphs[0]
        par.text = ""
        par.add_run(h["text"])
        if h.get("align"):
            par.alignment = {"left": 0, "center": 1, "right": 2}[h["align"]]
        style_runs(par, h, underline_border=h.get("bottom_border", False))

    if f.get("page_number") or f.get("text"):
        par = section.footer.paragraphs[0]
        par.text = ""
        if f.get("align"):
            par.alignment = {"left": 0, "center": 1, "right": 2}[f["align"]]
        fmt = f.get("format") or "第 {PAGE} 页 共 {NUMPAGES} 页"
        if f.get("text") and "{PAGE}" not in fmt:
            par.add_run(f["text"])
            fmt = ""
        # 按 {PAGE}/{NUMPAGES} 拆段
        parts = re.split(r"(\{PAGE\}|\{NUMPAGES\}|\{SECTIONPAGES\})", fmt)
        for ptxt in parts:
            if ptxt == "{PAGE}":
                field_runs(par, " PAGE ", "1")
            elif ptxt == "{NUMPAGES}":
                field_runs(par, " NUMPAGES ", "1")
            elif ptxt == "{SECTIONPAGES}":
                field_runs(par, " SECTIONPAGES ", "1")
            elif ptxt:
                par.add_run(ptxt)
        style_runs(par, f)

    if (cfg.get("page") or {}).get("first_page_different"):
        # 首页（封面）页眉页脚留空
        section.first_page_header.paragraphs[0].text = ""
        section.first_page_footer.paragraphs[0].text = ""



def apply_footnote_endnote(doc, cfg):
    """脚注/尾注：位置、编号格式、起始号、重编号规则、字号样式"""
    notes = cfg.get("notes") or {}
    fn, en = notes.get("footnote") or {}, notes.get("endnote") or {}
    sects = list(doc.sections)
    if not sects:
        return
    for sec in sects:
        sectPr = sec._sectPr
        for tag, c in (("w:footnotePr", fn), ("w:endnotePr", en)):
            if not c:
                continue
            old = sectPr.find(WQ(tag.split(":")[1]))
            if old is not None:
                sectPr.remove(old)
            pr = OxmlElement("w:" + tag.split(":")[1])
            if c.get("pos"):
                pr.append(OxmlElement("w:pos"))
                pr[-1].set(WQ("val"), c["pos"])
            if c.get("num_fmt"):
                e = OxmlElement("w:numFmt"); e.set(WQ("val"), c["num_fmt"])
                pr.append(e)
            if c.get("num_start"):
                e = OxmlElement("w:numStart"); e.set(WQ("val"), str(c["num_start"]))
                pr.append(e)
            if c.get("num_restart"):
                e = OxmlElement("w:numRestart")
                e.set(WQ("val"), c["num_restart"])
                pr.append(e)
            # 分隔线 / 续行分隔线：只在对应部件真的存在时才写。
            # 包里没有 word/endnotes.xml 却写 <w:endnote w:id="-1"/>，
            # Word 会判定引用悬空 -> 弹「文件可能已经损坏」。
            part = "/word/footnotes.xml" if "footnotePr" in tag else "/word/endnotes.xml"
            has_part = any(str(p.partname) == part
                           for p in doc.part.package.iter_parts())
            if has_part:
                for fid in c.get("separator_ids", [-1, 0]):
                    e = OxmlElement("w:" + tag.split(":")[1].rstrip("Pr"))
                    e.set(WQ("id"), str(fid))
                    pr.append(e)
            insert_in_order(sectPr, pr, SECTPR_ORDER)



def apply_table_details(doc, cfg):
    """表格：单元格边距、垂直对齐、行不跨页"""
    t = cfg.get("table") or {}
    root = doc.styles.element
    # 单元格边距（写进表格样式）
    margins = t.get("cell_margin_cm")
    if margins is not None:
        st = find_style(root, "Table")
        if st is not None:
            tblPr = st.find(WQ("tblPr"))
            if tblPr is None:
                tblPr = OxmlElement("w:tblPr"); st.append(tblPr)
            old = tblPr.find(WQ("tblCellMar"))
            if old is not None:
                tblPr.remove(old)
            mar = OxmlElement("w:tblCellMar")
            for side, key in (("top", "top"), ("left", "left"),
                              ("bottom", "bottom"), ("right", "right")):
                cm = margins if isinstance(margins, (int, float)) else margins.get(key, 0.1)
                e = OxmlElement("w:" + side)
                e.set(WQ("w"), str(int(round(float(cm) * 567))))
                e.set(WQ("type"), "dxa")
                mar.append(e)
            tblPr.append(mar)
            reorder(tblPr, TBLPR_ORDER)


def apply_typography(doc, cfg):
    """字符级排版：字间距、字符缩放、最小字号、文档网格、分栏、页面边框"""
    root = doc.styles.element
    # 字符级（放在 doc_default 里作为全局默认）
    dd = cfg.get("doc_default") or {}
    ddef = root.find(WQ("docDefaults"))
    if ddef is not None:
        rpr = ddef.find(WQ("rPrDefault") + "/" + WQ("rPr"))
        if rpr is not None:
            if dd.get("char_spacing_pt") is not None:
                e = rpr.find(WQ("spacing"))
                if e is None:
                    e = OxmlElement("w:spacing"); rpr.append(e)
                e.set(WQ("val"), str(int(round(float(dd["char_spacing_pt"]) * 20))))
            if dd.get("char_scale") is not None:
                e = rpr.find(WQ("w"))
                if e is None:
                    e = OxmlElement("w:w"); rpr.append(e)
                e.set(WQ("val"), str(int(dd["char_scale"])))
            if dd.get("kern_pt") is not None:
                e = rpr.find(WQ("kern"))
                if e is None:
                    e = OxmlElement("w:kern"); rpr.append(e)
                e.set(WQ("val"), str(int(round(float(dd["kern_pt"]) * 2))))
            reorder(rpr, RPR_ORDER)

    for sec in doc.sections:
        sectPr = sec._sectPr
        # 分栏
        cols = cfg.get("columns")
        if cols:
            old = sectPr.find(WQ("cols"))
            if old is not None:
                sectPr.remove(old)
            c = OxmlElement("w:cols")
            c.set(WQ("num"), str(int(cols.get("num", 2))))
            c.set(WQ("space"), str(int(round(float(cols.get("space_cm", 0.75)) * 567))))
            if cols.get("sep"):
                c.set(WQ("sep"), "1")
            if cols.get("equal_width") is False:
                c.set(WQ("equalWidth"), "0")
            sectPr.append(c)
        # 页面边框
        pb = cfg.get("page_border")
        if pb:
            old = sectPr.find(WQ("pgBorders"))
            if old is not None:
                sectPr.remove(old)
            b = OxmlElement("w:pgBorders")
            b.set(WQ("offsetFrom"), "page" if pb.get("from_page") else "text")
            for side in ("top", "left", "bottom", "right"):
                e = OxmlElement("w:" + side)
                e.set(WQ("val"), pb.get("style", "single"))
                e.set(WQ("sz"), str(int(pb.get("sz", 6))))
                e.set(WQ("space"), str(int(pb.get("space", 24))))
                e.set(WQ("color"), str(pb.get("color", "auto")).lstrip("#"))
                b.append(e)
            sectPr.append(b)



def apply_hyperlink_style(root, cfg):
    """超链接样式：颜色 / 下划线（默认 pandoc 沿用主题色）"""
    h = (cfg.get("styles") or {}).get("Hyperlink")
    if not h:
        return
    st = find_style(root, "Hyperlink")
    if st is None:
        return
    rPr = st.find(WQ("rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr"); st.append(rPr)
    if h.get("color"):
        e = rPr.find(WQ("color"))
        if e is None:
            e = OxmlElement("w:color"); rPr.append(e)
        e.set(WQ("val"), str(h["color"]).lstrip("#"))
        for a in ("themeColor", "themeShade", "themeTint"):
            if e.get(WQ(a)) is not None:
                del e.attrib[WQ(a)]
    if h.get("underline") is not None:
        e = rPr.find(WQ("u"))
        if e is None:
            e = OxmlElement("w:u"); rPr.append(e)
        e.set(WQ("val"), h["underline"] if h["underline"] else "none")
    reorder(rPr, RPR_ORDER)


def clone_table_style(root, src_id, new_id, name):
    """pandoc 给表格挂的是 FigureTable，reference.docx 里没有，克隆一份 Table"""
    if find_style(root, new_id) is not None:
        return
    src = find_style(root, src_id)
    if src is None:
        return
    st = copy.deepcopy(src)
    st.set(WQ("styleId"), new_id)
    st.attrib.pop(WQ("default"), None)
    nm = st.find(WQ("name"))
    if nm is not None:
        nm.set(WQ("val"), name)
    st.attrib.pop(WQ("customStyle"), None)
    root.append(st)


def apply_table_style(doc, cfg):
    """三线表：上下框线粗，表头下框线细，无竖线"""
    if not cfg.get("three_line", True):
        return
    root = doc.styles.element
    st = find_style(root, "Table")
    if st is None:
        return
    tblPr = st.find(WQ("tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr"); st.append(tblPr)
    old = tblPr.find(WQ("tblBorders"))
    if old is not None:
        tblPr.remove(old)
    borders = OxmlElement("w:tblBorders")
    outer = int(cfg.get("outer_border_sz", 12))     # 1/8 pt 单位, 12 = 1.5pt
    inner = int(cfg.get("inner_border_sz", 6))      # 6 = 0.75pt
    spec = {"top": outer, "bottom": outer, "left": 0, "right": 0,
            "insideH": 0, "insideV": 0}
    for side, sz in spec.items():
        e = OxmlElement("w:" + side)
        if sz:
            e.set(WQ("val"), "single"); e.set(WQ("sz"), str(sz))
            e.set(WQ("space"), "0"); e.set(WQ("color"), "auto")
        else:
            e.set(WQ("val"), "none"); e.set(WQ("sz"), "0")
            e.set(WQ("space"), "0"); e.set(WQ("color"), "auto")
        borders.append(e)
    tblPr.append(borders)
    reorder(tblPr, TBLPR_ORDER)

    # 表头行加下框线
    old_cf = st.find(WQ("tblStylePr"))
    if old_cf is not None:
        st.remove(old_cf)
    cf = OxmlElement("w:tblStylePr"); cf.set(WQ("type"), "firstRow")
    rPr = OxmlElement("w:rPr")
    rPr.append(OxmlElement("w:b"))
    for tag in ("i", "iCs"):
        e = OxmlElement("w:" + tag); e.set(WQ("val"), "0"); rPr.append(e)
    cf.append(rPr)
    tcPr = OxmlElement("w:tcPr")
    tb = OxmlElement("w:tcBorders")
    bottom = OxmlElement("w:bottom")
    bottom.set(WQ("val"), "single"); bottom.set(WQ("sz"), str(inner))
    bottom.set(WQ("space"), "0"); bottom.set(WQ("color"), "auto")
    tb.append(bottom); tcPr.append(tb); cf.append(tcPr)
    st.append(cf)


# ---------- 主流程 ----------

def build(config_path, out_path, base_docx=None):
    with open(config_path, encoding="utf-8") as fh:
        raw = fh.read()
    cfg = json.loads(strip_json_comments(raw))

    tmp_base = None
    if base_docx:
        base = os.path.abspath(base_docx)
        if not os.path.exists(base):
            raise SystemExit("找不到模板文件：%s" % base)
        print("· 以用户模板为底：%s" % base)
    else:
        # pandoc 的默认 reference.docx 取到临时目录，避免污染 skill 目录、也避免并发写冲突
        tmp_base = os.path.join(tempfile.mkdtemp(prefix="md2docx-base."), "reference.docx")
        subprocess.run(["pandoc", "-o", tmp_base, "--print-default-data-file",
                        "reference.docx"], check=True)
        base = tmp_base

    doc = Document(base)
    root = doc.styles.element

    # 版心宽度（twips）：目录页码右对齐要用
    try:
        pg = cfg.get("page") or {}
        m = pg.get("margin_cm") or {}
        w_cm = float(pg.get("width_cm", 21.0))
        left = float(m.get("left", 3.17))
        right = float(m.get("right", 2.54))
        CTX["text_width_twips"] = int(round((w_cm - left - right) / 2.54 * 1440))
    except Exception:
        CTX["text_width_twips"] = 8306

    set_doc_defaults(root, cfg.get("doc_default", {}))

    # 1) 具名样式
    for sid, props in (cfg.get("styles") or {}).items():
        st = find_style(root, sid)
        if st is None:
            pretty = (cfg.get("pretty_names") or {}).get(sid, sid)
            st = ensure_style(root, sid, pretty)
            print("  · 新建样式: %s (%s)" % (sid, pretty))
        apply_rpr(st, props)
        apply_ppr(st, props)

    # 1.5) 目录条目样式：Word 更新目录域时会用 TOC1..TOC9，样式不存在就沿用默认字体
    for i in range(1, 10):
        sid = "TOC%d" % i
        if find_style(root, sid) is None:
            st = ensure_style(root, sid, "toc %d" % i)
            nxt = OxmlElement("w:next"); nxt.set(WQ("val"), "Normal")
            st.append(nxt)
    if find_style(root, "TableofFigures") is None:
        ensure_style(root, "TableofFigures", "table of figures")
    # 表格单元格独立样式：pandoc 把单元格和紧凑列表都用 Compact，
    # 想「列表跟正文一样大、表格小一号」就做不到，所以拆出来。
    if find_style(root, "TableText") is None:
        ensure_style(root, "TableText", "Table Text")

    # 目录条目外观：没有 width 制表位时，Word 更新目录后页码会落到默认制表位
    # （1.27cm），三级条目全部顶格、没有点线前导，看起来是一团乱。
    # 这里给 TOC1..9 补上「层级缩进 + 右对齐点线制表位」。
    toc_cfg = cfg.get("toc") or {}
    if toc_cfg.get("tabs", True):
        for i in range(1, 10):
            sid = "TOC%d" % i
            st = find_style(root, sid)
            if st is None:
                continue
            user = (cfg.get("styles") or {}).get(sid) or {}
            # 用户显式写了 tabs 就完全听用户的；否则只补缺的那部分
            props = {}
            if "tabs" not in user:
                props["tabs"] = [{"val": "right", "pos": "text", "leader": "dot"}]
            if "left_indent_chars" not in user:
                props["left_indent_chars"] = (i - 1) * 2
            if props:
                apply_ppr(st, props)
        st = find_style(root, "TableofFigures")
        if st is not None:
            user = (cfg.get("styles") or {}).get("TableofFigures") or {}
            if "tabs" not in user:
                apply_ppr(st, {"tabs": [{"val": "right", "pos": "text",
                                         "leader": "dot"}]})

    # 2) 没配置的样式：确保不带首行缩进（否则会被 BodyText 传染）
    configured = set((cfg.get("styles") or {}).keys())
    for sid in NO_INDENT_DEFAULT:
        if sid in configured:
            continue
        st = find_style(root, sid)
        if st is not None:
            zero_indent(st)

    detach_heading_based(root)

    # 2.5) 主题色 -> 黑色（正式报告标题不该是 pandoc 默认的蓝）
    n_theme = neutralize_theme_colors(root)
    if n_theme:
        print("  · %d 个样式的主题色已改为黑色" % n_theme)

    # 3) 斜体全球清除
    kill_all_italics(root)
    for sid in (cfg.get("styles") or {}):
        st = find_style(root, sid)
        if st is None:
            continue
        rPr = st.find(WQ("rPr"))
        if rPr is None:
            rPr = OxmlElement("w:rPr"); st.append(rPr)
        for tag in ("i", "iCs"):
            e = rPr.find(WQ(tag))
            if e is None:
                e = OxmlElement("w:" + tag); rPr.append(e)
            e.set(WQ("val"), "0")
        reorder(rPr, RPR_ORDER)

    # 3.5) 修正所有样式的子元素顺序（schema 要求，否则 Word 可能报修复）
    for st in root.findall(WQ("style")):
        reorder(st, STYLE_ORDER)

    # 4) 页面、页眉页脚
    section = apply_page(doc, cfg.get("page") or {})
    apply_header_footer(doc, section, cfg)

    # 5) 表格：给 pandoc 用的 FigureTable 克隆一个基础样式
    #    具体的线框/表头/对齐由 docxfix.py 按 profile 直接写到每张表上
    clone_table_style(root, "Table", "FigureTable", "Figure Table")
    apply_table_details(doc, cfg.get("table") or {})

    # 6) 脚注尾注 / 列表编号 / 字符排版 / 文档开关 / 超链接样式
    apply_footnote_endnote(doc, cfg)
    apply_typography(doc, cfg)
    apply_hyperlink_style(root, cfg)

    # 所有容器的子元素按 schema 排序，否则 Word 会认为文件损坏
    TRPR_ORDER = ["cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore",
                  "wAfter", "cantSplit", "trHeight", "tblHeader",
                  "tblCellSpacing", "jc", "hidden", "ins", "del", "trPrChange"]
    TCPR_ORDER = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders",
                  "shd", "noWrap", "tcMar", "textDirection", "tcFitText",
                  "vAlign", "hideMark", "headers", "cellIns", "cellDel",
                  "cellMerge", "tcPrChange"]
    for tag, order in (("pPr", PPR_ORDER), ("rPr", RPR_ORDER),
                       ("tblPr", TBLPR_ORDER), ("trPr", TRPR_ORDER),
                       ("tcPr", TCPR_ORDER), ("sectPr", SECTPR_ORDER),
                       ("style", STYLE_ORDER)):
        for parent in doc.element.body.iter(WQ(tag)):
            reorder(parent, order)
        for parent in root.iter(WQ(tag)):
            reorder(parent, order)

    doc.save(out_path)
    if tmp_base:
        shutil.rmtree(os.path.dirname(tmp_base), ignore_errors=True)
    print("✓ 生成 %s" % out_path)


if __name__ == "__main__":
    args = sys.argv[1:]
    cfgp = os.path.join(HERE, "profiles", "technical.json")
    outp = os.path.join(tempfile.mkdtemp(prefix="md2docx-tpl."), "technical.docx")
    base_docx = None
    if "-p" in args:
        cfgp = args[args.index("-p") + 1]
    elif args and not args[0].startswith("-"):
        cfgp = args.pop(0)
    if "-o" in args:
        outp = args[args.index("-o") + 1]
    if "--base" in args:
        base_docx = args[args.index("--base") + 1]
    os.makedirs(os.path.dirname(os.path.abspath(outp)), exist_ok=True)
    sys.exit(build(cfgp, outp, base_docx))
