#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
_common.py —— 所有命令共用的几件小事

以前这些小函数在 4~5 个文件里各写了一份（strip_json_comments / resolve_profile /
die …），改一处要改五遍。现在集中在这里。

本文件不承载任何业务逻辑，只有：
  · 路径常量        HERE / PROFILE_DIR / DEFAULT_PROFILE
  · 退出与命令执行   die / run
  · 档案解析        strip_json_comments / resolve_profile / load_profile
"""
import json
import os
import re
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
PROFILE_DIR = os.path.join(HERE, "profiles")
DEFAULT_PROFILE = "technical"


def die(msg, code=1):
    print(msg, file=sys.stderr)
    sys.exit(code)


def run(cmd, **kw):
    """跑子进程并收回输出（统一用 text 模式）"""
    return subprocess.run(cmd, capture_output=True, text=True, **kw)


def strip_json_comments(text):
    """去掉 JSON 里的 // 和 /* */ 注释（认识字符串，不会误伤 http://）"""
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


def profile_candidates(arg):
    """--profile 既能给名字，也能给路径；按优先级列出候选"""
    return [arg,
            os.path.join(PROFILE_DIR, arg),
            os.path.join(PROFILE_DIR, arg + ".json"),
            os.path.join(HERE, arg),
            os.path.join(HERE, arg + ".json"),
            os.path.join(os.getcwd(), arg),
            os.path.join(os.getcwd(), arg + ".json")]


def resolve_profile(arg, default=True):
    """返回 (名字, 绝对路径)。arg 为空时给默认档案。"""
    if not arg:
        if not default:
            return None, None
        arg = DEFAULT_PROFILE
    for c in profile_candidates(arg):
        if os.path.isfile(c):
            return os.path.basename(c).replace(".json", ""), os.path.abspath(c)
    return None, None


def load_profile(path):
    """读档案（自动跳过注释）"""
    with open(path, encoding="utf-8") as fh:
        return json.loads(strip_json_comments(fh.read()))
