#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
preview.py —— 把生成的 docx 转成 PDF 预览

    python3 preview.py 报告.docx                    # 导出整份 PDF 并打开
    python3 preview.py 报告.docx --pages 1,2,5      # 只导这几页（拼一张竖排长图）
    python3 preview.py 报告.docx --sheet            # 关键页（首/目录/正文首/末）拼联系表
    python3 preview.py 报告.docx --sheet --cols 3   # 自定义列数
    python3 preview.py 报告.docx --no-open          # 只生成不打开
    python3 preview.py 报告.docx --out x.pdf        # 指定输出路径

**看效果优先用 --sheet**：一张图就能看完封面、目录、正文首末页，
不用逐页截图。

为什么不会弹「授予访问权限」：
    Word 是沙箱应用，直接让它打开 ~/Documents 下的文件会弹授权框。
    把文件复制到 Word 自己的沙箱容器
    （~/Library/Containers/com.microsoft.Word/Data/Documents/）再打开就不会弹。

怎么避免「此文件被锁定，锁者：另一用户」：
    Word 是单实例，同名文件被两个会话打开就会互相锁。
    所以：① 每次用「时间+进程号+随机」的唯一文件名；
         ② 开跑前清掉自己上次留下的临时文件和 ~$ 锁文件；
         ③ 转完立刻用 Cmd+W 关掉刚打开的文档（这版 Word 的
            AppleScript 没有 close 命令）；
         ④ 已经有比 docx 新的 PDF 就直接复用，根本不去碰 Word。
    还嫌脏就运行：python3 preview.py --cleanup
"""
import uuid
import os
import shutil
import subprocess
import sys
import tempfile
import time

HERE = os.path.dirname(os.path.abspath(__file__))
WORD_CONTAINER = os.path.expanduser(
    "~/Library/Containers/com.microsoft.Word/Data/Documents")
WORD_WORK = os.path.join(WORD_CONTAINER, ".md2docx_preview")

# Word 只有一个实例：两个会话同时驱动会互相干扰（open/active document 打架）。
# 用文件锁排队，避免「找不到活动文档」「转出来是别的文件」这类怪现象。
LOCK_PATH = os.path.join(tempfile.gettempdir(), "md2docx-word.lock")


def acquire_lock(timeout=240):
    try:
        import fcntl
    except ImportError:
        return None
    fh = open(LOCK_PATH, "w")
    waited = 0
    while True:
        try:
            fcntl.flock(fh, fcntl.LOCK_EX | fcntl.LOCK_NB)
            return fh
        except BlockingIOError:
            if waited == 0:
                print("另一个会话正在用 Word，排队等待…", file=sys.stderr)
            if waited >= timeout:
                fh.close()
                raise SystemExit("等了 %d 秒还没轮上，请稍后再试" % timeout)
            time.sleep(2)
            waited += 2


def release_lock(fh):
    if fh is None:
        return
    try:
        import fcntl
        fcntl.flock(fh, fcntl.LOCK_UN)
    except Exception:
        pass
    try:
        fh.close()
    except Exception:
        pass

CONVERT = '''
on run argv
    set src to item 1 of argv
    set dst to item 2 of argv
    set wantName to item 3 of argv
    with timeout of 900 seconds
        tell application "Microsoft Word"
            activate
            -- Word 只有一个实例，别人也在用时 open 可能被吞掉。
            -- 所以：反复尝试打开，并在文档列表里按文件名确认真的打开了。
            set myDoc to missing value
            repeat 6 times
                try
                    open file name src
                end try
                delay 2
                repeat with d in (get documents)
                    try
                        if (name of d) is wantName then
                            set myDoc to contents of d
                            exit repeat
                        end if
                    end try
                end repeat
                if myDoc is not missing value then exit repeat
            end repeat
            if myDoc is missing value then
                error "打不开 " & wantName & "（Word 正被其它任务占用？）"
            end if
            try
                update fields of text object of myDoc
            end try
            delay 1
            save as myDoc file name dst file format format PDF
        end tell
    end timeout
    return "ok"
end run
'''

COUNT_DOCS = '''
tell application "Microsoft Word"
    try
        return (count of documents) as string
    on error
        return "0"
    end try
end tell
'''

QUIT_WORD = '''
tell application "Microsoft Word"
    try
        quit saving no
    end try
end tell
'''


def die(msg):
    print(msg, file=sys.stderr)
    sys.exit(1)


def run_osa(script, *args):
    with tempfile.NamedTemporaryFile("w", suffix=".applescript",
                                     delete=False, encoding="utf-8") as fh:
        fh.write(script)
        path = fh.name
    try:
        return subprocess.run(["osascript", path] + list(args),
                              capture_output=True, text=True)
    finally:
        try:
            os.remove(path)
        except OSError:
            pass


def dismiss_dialogs():
    """关掉 Word 上的模态弹窗

    Word 有弹窗时（例如「文件可能已经损坏」「此文件被锁定」），
    所有 AppleScript 的 open 都会被静默忽略，表现为「打不开但也不报错」。
    """
    script = '''
tell application "System Events"
    tell process "Microsoft Word"
        set n to 0
        repeat 5 times
            set did to false
            try
                repeat with w in windows
                    try
                        if (count of sheets of w) > 0 then
                            set b to button 1 of sheet 1 of w
                            click b
                            set did to true
                            delay 0.5
                        end if
                    end try
                end repeat
            end try
            if did then set n to n + 1
            if not did then exit repeat
        end repeat
        return n as string
    end tell
end tell
'''
    r = run_osa(script)
    return (r.stdout or "").strip()



def word_state():
    """返回 Word 的进程状态：是否运行、PID、PPID、窗口数、文档数

    为什么要看 PPID：如果 Word 是被脚本拉起来的，脚本退出后 Word 的父进程
    变成 1（孤儿化），这种实例常常「进程活着、能答属性查询，但没有窗口，
    open 被静默丢弃」——看起来像僵尸，实际上是 GUI 会话没起来。
    """
    st = {"running": False, "pid": None, "ppid": None,
          "windows": None, "docs": None}
    r = subprocess.run(["pgrep", "-x", "Microsoft Word"], capture_output=True, text=True)
    pid = (r.stdout or "").strip().splitlines()
    if not pid:
        return st
    st["running"] = True
    st["pid"] = pid[0].strip()
    try:
        ps = subprocess.run(["ps", "-o", "ppid=", "-p", st["pid"]],
                            capture_output=True, text=True).stdout.strip()
        st["ppid"] = int(ps) if ps.isdigit() else None
    except Exception:
        pass
    try:
        st["docs"] = int(run_osa(COUNT_DOCS).stdout.strip() or 0)
    except ValueError:
        pass
    win = run_osa('''
tell application "System Events"
    tell process "Microsoft Word"
        try
            return (count of windows) as string
        on error
            return "0"
        end try
    end tell
end tell
''')
    try:
        st["windows"] = int((win.stdout or "0").strip() or 0)
    except ValueError:
        pass
    return st


def explain_word_state(st):
    """把 Word 状态翻译成「现在怎么办」"""
    if not st["running"]:
        print("Word：未运行。preview.py 不会替你冷启动它（冷启动容易卡在引导页）。")
        return "not-running"
    pid, ppid = st["pid"], st["ppid"]
    win, docs = st["windows"], st["docs"]
    print("Word：运行中（PID %s，父进程 %s），窗口 %s 个，文档 %s 个"
          % (pid, ppid, win, docs))
    if win == 0 and not docs:
        print()
        print("  诊断：Word 进程在，但没有窗口 —— 它被「孤儿化」了。")
        if ppid == 1:
            print("        父进程是 1，说明当初是脚本/后台拉起来的，脚本退出后它就没再建界面。")
        print("        这种实例能回答属性查询，但 open 会被静默丢弃（表现为 -2700）。")
        print()
        print("  处理（任选）：")
        print("    终端：kill -9 %s" % pid)
        print("    或：活动监视器 → 搜索 Microsoft Word → 选中 → 点 × → 强制退出")
        print("    然后从 Dock / 应用程序文件夹手动启动 Word，等它出现正常窗口。")
        return "orphaned"
    if win == 0:
        print("  诊断：Word 有文档没有窗口，可能卡在弹窗上。关掉弹窗再试。")
        return "no-window"
    return "ok"

def diagnose():
    """转换失败时给出可执行的诊断，而不是只丢一个错误码"""
    st = word_state()
    print("诊断：", file=sys.stderr)
    kind = explain_word_state(st)
    if kind == "ok":
        print("  → Word 状态正常但打不开文件，可能是另一个任务正在用它。",
              file=sys.stderr)
        print("    稍等重试，或直接双击 docx 自己看效果。", file=sys.stderr)
    elif kind == "not-running":
        print("  → 先手动启动一次 Word，确认没有弹窗，再重试 --word。",
              file=sys.stderr)


def reset_word():
    """Word 有时会进入异常状态（文档列表里出现 missing value），杀掉重启即可

    典型诱因：上次预览的临时文件被外部删除 / Word 卡在弹框上。
    """
    # 不能用 SIGKILL：macOS 会弹「执行了强制退出」错误报告，用户以为文件坏了。
    # 温和地让 Word 退出即可；退不掉就让用户自己处理。
    subprocess.run(["osascript", "-e",
                    'tell application "Microsoft Word" to quit saving no'],
                   capture_output=True, text=True)
    time.sleep(2)



# ---------------------------------------------------------------- 关键页 / 联系表

def pick_key_pages(pdf, total):
    """挑出最值得看的那几页：封面、目录、正文首页、最后一页"""
    pages = [1]
    try:
        out = subprocess.run(["pdftotext", "-layout", pdf, "-"],
                             capture_output=True, text=True).stdout
        texts = out.split("\f")
    except Exception:
        texts = []
    # 目录页：含「目 录」或大量点线
    for i, t in enumerate(texts[:8], start=1):
        if i == 1:
            continue
        if "目\u3000录" in t or "目录" in t or t.count("..") > 3:
            pages.append(i)
            break
    # 正文第一页：目录之后第一个有正文章节号的行
    start = pages[-1] + 1 if len(pages) > 1 else 2
    for i in range(start, min(total, start + 6) + 1):
        pages.append(i)
        break
    if total not in pages:
        pages.append(total)
    return sorted(set(p for p in pages if 1 <= p <= total))


def render_pages(pdf, pages, dpi=80):
    """把指定页渲染成 PNG，返回 [(页码, 路径)]"""
    out = []
    tmpdir = tempfile.mkdtemp(prefix="md2docx-pages.")
    for p in pages:
        base = os.path.join(tmpdir, "p%03d" % p)
        subprocess.run(["pdftoppm", "-png", "-r", str(dpi), "-f", str(p),
                        "-l", str(p), pdf, base], capture_output=True)
        for f in os.listdir(tmpdir):
            if f.startswith("p%03d" % p) and f.endswith(".png"):
                out.append((p, os.path.join(tmpdir, f)))
    return out


def make_sheet(images, out_path, cols=2, label=True):
    """把若干页拼成一张联系表（一张图看完，省得逐页截图）"""
    try:
        from PIL import Image, ImageDraw
    except ImportError:
        print("没有 Pillow，改成分页导出（pip install pillow 可拼图）",
              file=sys.stderr)
        return None
    if not images:
        return None
    thumbs = []
    for num, f in images:
        try:
            im = Image.open(f).convert("RGB")
        except Exception:
            continue
        w = 900
        h = int(im.height * w / im.width)
        thumbs.append((num, im.resize((w, h), Image.LANCZOS)))
    if not thumbs:
        return None
    cols = max(1, min(int(cols), len(thumbs)))
    rows = (len(thumbs) + cols - 1) // cols
    pad, bar = 24, 34
    tw = max(t.width for _, t in thumbs)
    th = max(t.height for _, t in thumbs)
    W = cols * tw + (cols + 1) * pad
    H = rows * (th + bar) + (rows + 1) * pad
    sheet = Image.new("RGB", (W, H), (238, 238, 238))
    draw = ImageDraw.Draw(sheet)
    for i, (num, im) in enumerate(thumbs):
        r, c = divmod(i, cols)
        x = pad + c * (tw + pad)
        y = pad + r * (th + bar + pad)
        sheet.paste(im, (x, y))
        draw.rectangle([x - 2, y - 2, x + im.width + 1, y + im.height + 1],
                       outline=(150, 150, 150))
        if label:
            # 用 ASCII 页码：PIL 内置字体没有中文字形，写「第 N 页」会变方框
            draw.text((x + 6, y + im.height + 8), "page %d" % num,
                      fill=(40, 40, 40))
    sheet.save(out_path)
    return out_path



def purge_leftovers(max_age_sec=0):
    """删掉本工具在 Word 容器里留下的临时文件

    这些东西留着会被 Word 认为「别人还在用」，下次打开就弹
    「此文件被锁定，锁者：另一用户」。max_age_sec=0 表示全部清掉。
    """
    n = 0
    if not os.path.isdir(WORD_WORK):
        return n
    now = time.time()
    for f in list(os.listdir(WORD_WORK)):
        if not (f.startswith("md2docx_") or f.startswith("~$md2docx_")):
            continue
        p = os.path.join(WORD_WORK, f)
        try:
            if max_age_sec and now - os.path.getmtime(p) < max_age_sec:
                continue
            os.remove(p)
            n += 1
        except OSError:
            pass
    return n


def close_our_document(basename):
    """转换完把刚打开的那个文档关掉

    Word 这版的 AppleScript 不支持 close（会报 -1708），但可以用
    Cmd+W 关掉最前面的窗口；关闭前先确认最前面那个确实是我们的文件。
    """
    script = '''
on run argv
    set wantName to item 1 of argv
    tell application "Microsoft Word"
        try
            if (name of active document) is not wantName then return "skip"
        on error
            return "skip"
        end try
        activate
    end tell
    delay 0.3
    tell application "System Events"
        tell process "Microsoft Word"
            keystroke "w" using command down
        end tell
    end tell
    return "closed"
end run
'''
    r = run_osa(script, basename)
    time.sleep(0.6)
    return (r.stdout or "").strip() == "closed"


def with_word(docx, _retry=True):
    """用 Word 渲染。文件放进 Word 自己的容器，避免弹授权框"""
    if not os.path.isdir(os.path.expanduser(
            "~/Library/Containers/com.microsoft.Word")):
        return None, None
    # 绝不「冷启动」Word：新启动的 Word 会弹引导页/登录页，
    # AppleScript 在那种状态下拿不到文档（-2700、document = missing value），
    # 而且会把 Word 留在一个卡死的状态里。
    if subprocess.run(["pgrep", "-x", "Microsoft Word"],
                      capture_output=True).returncode != 0:
        print("Word 当前没有运行。脚本不会替你启动它——冷启动的 Word 会卡在引导页。",
              file=sys.stderr)
        print("请先手动打开一次 Microsoft Word，确认没弹窗后关闭窗口（不必退出），再重试。",
              file=sys.stderr)
        return None, 0

    os.makedirs(WORD_WORK, exist_ok=True)

    # 转换前 Word 里开着几个文档（用于决定事后能不能退出 Word）
    try:
        n0 = int(run_osa(COUNT_DOCS).stdout.strip() or 0)
    except ValueError:
        n0 = 0

    # 每次用唯一文件名：Word 是单实例，多会话共用同一个名字会互相覆盖/读错文件
    # 名字必须全局唯一：Word 是单实例，两个会话用同名文件时，
    # 后打开的那个会看到「此文件被锁定，锁者：另一用户」。
    h = "%s_%d_%s" % (time.strftime("%H%M%S"), os.getpid(), uuid.uuid4().hex[:6])
    src = os.path.join(WORD_WORK, "md2docx_%s.docx" % h)
    dst = os.path.join(WORD_WORK, "md2docx_%s.pdf" % h)
    purge_leftovers()
    shutil.copy(docx, src)
    # 清掉可能残留的 Word owner 文件（~$xxx.docx）。
    # 不清的话 Word 会认为「此文件被锁定，锁者：另一用户」，直接拒绝打开。
    for junk in (os.path.join(WORD_WORK, "~$" + os.path.basename(src)),):
        if os.path.exists(junk):
            try:
                os.remove(junk)
            except OSError:
                pass
    if os.path.exists(dst):
        os.remove(dst)

    dismiss_dialogs()
    r = run_osa(CONVERT, src, dst, os.path.basename(src))
    if r.returncode != 0 or not os.path.exists(dst):
        msg = (r.stderr or r.stdout).strip()
        for junk in (src, os.path.join(WORD_WORK, "~$" + os.path.basename(src))):
            try:
                os.remove(junk)
            except OSError:
                pass
        if _retry:
            print("Word 状态异常，重启后重试…", file=sys.stderr)
            reset_word()
            return with_word(docx, _retry=False)
        print("Word 转换失败：%s" % msg, file=sys.stderr)
        diagnose()
        return None, n0

    # 不把文档留在 Word 里：留一个就会让下一次打开同名/同路径文件弹「被锁定」
    close_our_document(os.path.basename(src))
    return dst, n0


def main():
    args = sys.argv[1:]
    flags = [a for a in args if a.startswith("--")]
    pos = [a for a in args if not a.startswith("--")]

    if "--doctor" in flags:
        print("Word 预览环境检查")
        print("-" * 46)
        st = word_state()
        explain_word_state(st)
        left = []
        if os.path.isdir(WORD_WORK):
            left = [f for f in os.listdir(WORD_WORK)
                    if f.startswith(("md2docx_", "~$md2docx_"))]
        print("本工具的临时文件：%d 个%s"
              % (len(left), "（可清理：preview.py --cleanup）" if left else ""))
        print("-" * 46)
        print("提醒：格式验收优先用 verify.py，它不需要渲染器。")
        return 0

    if "--cleanup" in flags:
        n = purge_leftovers()
        print("已清理 %d 个本工具留在 Word 容器里的临时文件" % n)
        print("如果 Word 里还有 md2docx_ 开头的文档，请手动关掉。")
        return 0

    if not pos or "-h" in flags or "--help" in flags:
        print(__doc__)
        return 1
    docx = os.path.abspath(pos[0])
    if not os.path.isfile(docx):
        die("找不到文件：%s" % docx)

    out = None
    if "--out" in args:          # 注意要在完整参数里找
        i = args.index("--out")
        if i + 1 < len(args):
            out = args[i + 1]

    # 已经有一份比 docx 新的 PDF 就直接用，不必再驱动 Word
    sibling = os.path.splitext(docx)[0] + ".pdf"
    reuse = None
    if (reuse is None and os.path.exists(sibling)
            and os.path.getmtime(sibling) >= os.path.getmtime(docx)
            and "--force" not in flags):
        reuse = sibling

    if reuse:
        pdf, n0 = reuse, 0
        print("复用已有 PDF：%s（不启动 Word）" % reuse)
    elif "--word" in flags:
        lock = acquire_lock()
        try:
            pdf, n0 = with_word(docx)
        finally:
            release_lock(lock)
    else:
        # 默认不自动驱动 Word：
        # Word 在 macOS 上只有一个实例，脚本驱动它很容易弹
        # 「此文件被锁定，锁者：另一用户」，而且这版 Word 的
        # AppleScript 不能可靠关闭文档，残留会越积越多。
        print("没有可用的 PDF，且默认不自动启动 Word。", file=sys.stderr)
        print("", file=sys.stderr)
        print("想要 PDF / 联系表，任选一种：", file=sys.stderr)
        print("  ① 自己双击打开 %s → 文件 → 导出为 PDF，"
              "再运行本命令（会自动复用那份 PDF）" % os.path.basename(docx),
              file=sys.stderr)
        print("  ② 明确允许驱动 Word：python3 preview.py \"%s\" --word"
              % os.path.basename(docx), file=sys.stderr)
        print("", file=sys.stderr)
        print("如果只是想确认格式对不对，先跑 verify.py，它不需要任何渲染器：",
              file=sys.stderr)
        print("  python3 verify.py \"%s\"" % os.path.basename(docx), file=sys.stderr)
        return 3
    if not pdf or not os.path.exists(pdf):
        die("转换失败")

    if out:
        final = os.path.abspath(out)
    else:
        final = os.path.join(os.path.dirname(docx),
                             os.path.splitext(os.path.basename(docx))[0] + ".pdf")
    if os.path.abspath(pdf) != os.path.abspath(final):
        shutil.copy(pdf, final)
    # 只删「Word 容器里的临时 PDF」。
    # 绝不能用 final（那可能是用户自己导出的 PDF，删了就丢数据）。
    if os.path.abspath(pdf).startswith(os.path.abspath(WORD_WORK) + os.sep):
        try:
            os.remove(pdf)
        except OSError:
            pass
    # 顺手清掉本工具留下的旧临时文档（只删自己命名的，不碰别人的文件）
    try:
        now = time.time()
        for f in os.listdir(WORD_WORK):
            if not f.startswith("md2docx_"):
                continue
            p = os.path.join(WORD_WORK, f)
            if os.path.isfile(p) and now - os.path.getmtime(p) > 3600:
                os.remove(p)
    except OSError:
        pass

    # ---------- 关键页 / 联系表：只给一张图，避免逐页截图 ----------
    pages_arg = None
    if "--pages" in args:
        i = args.index("--pages")
        if i + 1 < len(args):
            pages_arg = args[i + 1]
    want_sheet = "--sheet" in flags or pages_arg is not None
    if want_sheet:
        cols = 2
        if "--cols" in args:
            i = args.index("--cols")
            if i + 1 < len(args):
                cols = int(args[i + 1])
        try:
            total = int(subprocess.run(["pdfinfo", final], capture_output=True,
                                       text=True).stdout.split("Pages:")[1].split()[0])
        except (IndexError, ValueError, FileNotFoundError, TypeError):
            total = 1
        if pages_arg:
            pages = [int(x) for x in re.split(r"[,\s]+", pages_arg) if x.strip().isdigit()]
        else:
            pages = pick_key_pages(final, total)
        imgs = render_pages(final, pages)
        sheet_path = os.path.splitext(final)[0] + "_联系表.png"
        got = make_sheet(imgs, sheet_path, cols=cols)
        if got:
            print("联系表：%s（第 %s 页）" % (got, ",".join(str(p) for p, _ in imgs)))
            if "--no-open" not in flags:
                subprocess.run(["open", got])
            return 0
        print("拼图失败（需要 Pillow：pip install pillow），已导出整份 PDF：%s" % final,
              file=sys.stderr)

    print("预览：%s" % final)
    if n0:
        print("提示：Word 里留着 1 个临时文档窗口，自行关掉即可（脚本已尝试自动关）。")
    if "--no-open" not in flags:
        subprocess.run(["open", final])
    return 0


if __name__ == "__main__":
    sys.exit(main())
