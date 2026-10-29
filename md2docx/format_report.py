#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
format_report.py —— 把一份 profile「翻译成人话」，输出当前生效的格式清单

    python3 format_report.py profiles/technical.json
    python3 format_report.py --profile official

清单内容来自「实际编译出来的模板 docx」，所以是真正生效的效果，
不是 JSON 里写了什么就显示什么（样式之间有继承关系）。

用来让 agent 回答「现在正文是什么字体、缩进多少、行距多少」这类问题，
或者让用户确认格式。
"""
import json
import os
import re
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
sys.path.insert(0, HERE)


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


def resolve_profile(path):
    """把名字/相对路径解析成实际文件路径"""
    if os.path.isabs(path) and os.path.exists(path):
        return path
    for cand in (path,
                 os.path.join(HERE, "profiles", path),
                 os.path.join(HERE, "profiles", path + ".json"),
                 os.path.join(os.getcwd(), path),
                 os.path.join(os.getcwd(), path + ".json")):
        if os.path.exists(cand):
            return os.path.abspath(cand)
    raise SystemExit("找不到 profile：%s" % path)


def load_profile(path):
    path = resolve_profile(path)
    return json.loads(strip_json_comments(open(path, encoding="utf-8").read()))


# ---------------- 模板编译 + 样式解析 ----------------

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
    profile_path = resolve_profile(profile_path)
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


def main():
    argv = sys.argv[1:]
    path = None
    base = None
    if "--profile" in argv:
        path = argv[argv.index("--profile") + 1]
    if "--base" in argv:
        base = argv[argv.index("--base") + 1]
    if not path:
        path = argv[0] if argv else "technical"
    print(report(path, base))
    return 0


if __name__ == "__main__":
    sys.exit(main())
