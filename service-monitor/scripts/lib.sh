#!/usr/bin/env bash
# lib.sh —— 监控循环的公共骨架。被 <项目>/.monitor/<服务名>.sh 用 source 引入。
#
# 设计要点：**脆弱逻辑只写一份**。
# 进程锁、心跳、超时、CHECK 报错检测、存活探针、截断、重试，全部在这里。
# 每个服务的脚本只负责回答三个问题：叫什么、怎么判断好了、推什么。
#
# 用法（在 .monitor/<服务名>.sh 里）：
#   SVC='backup-db'
#   CHECK='test -f /backup/$(date +%F).sql.gz'
#   BODY='du -h /backup/$(date +%F).sql.gz | cut -f1'
#   source "$HOME/.notify/lib.sh"
#   monitor_loop
#
# 可用参数（都有默认值，只设需要的）：
#   INTERVAL=30        多久检查一次（秒）
#   STABLE=0           判断通过后再等几秒复检一次，防「文件写到一半」
#   HEARTBEAT=1800     每多久打一行「还在等」，0=关
#   TIMEOUT=0          最长等多久（秒），0=无限
#   WATCH_PID=""       目标进程 PID；它消失而条件没满足 → 立即告警
#   WATCH_PATTERN=""   目标进程命令行特征（pgrep -f）；同上
#   BODY_MAX=300       推送正文最多多少【字符】，0=不截断
#   TITLE_DONE=""      推送标题，默认 "<服务名> 完成了"

# ═══════════════════════════════════════════════════════════════════════
# 跨平台工具：macOS 和 Linux 的 stat 参数完全不同，包一层
# ═══════════════════════════════════════════════════════════════════════
_mtime() {   # 文件最后修改时间（epoch 秒）
    case "$(uname -s)" in
        Darwin) stat -f %m "$1" 2>/dev/null || echo 0 ;;
        *)      stat -c %Y "$1" 2>/dev/null || echo 0 ;;
    esac
}

# 按【字符】截断。不能用 bash 的 ${var:0:N}：
# 它在 C locale 下按字节切，会把一个汉字切成两半，推过去是乱码（实测出现过）。
_cut_chars() {   # $1=文本 $2=最大字符数
    local text="$1" n="$2"
    if command -v python3 >/dev/null 2>&1; then
        printf '%s' "$text" | N="$n" python3 -c '
import os, sys
n = int(os.environ["N"]); s = sys.stdin.read().rstrip("\n")
print(s if len(s) <= n else s[:n] + "…（已截断，完整内容见日志）")'
    elif command -v perl >/dev/null 2>&1; then
        printf '%s' "$text" | perl -CSD -ne 'chomp; print length($_)>'"$n"' ? substr($_,0,'"$n"')."…（已截断）\n" : "$_\n"'
    else
        printf '%s' "$text" | head -c "$n"; printf '…（已截断）\n'
    fi
}

_mb() {  # 文件大小（MB，整数）
    local s; s=$(wc -c < "$1" 2>/dev/null | tr -d ' ')
    echo $(( ${s:-0} / 1048576 ))
}

# ═══════════════════════════════════════════════════════════════════════
# 主循环
# ═══════════════════════════════════════════════════════════════════════
monitor_loop() {
    : "${SVC:?lib.sh: 必须设置 SVC}"
    : "${CHECK:?lib.sh: 必须设置 CHECK}"
    : "${BODY:=}"
    : "${INTERVAL:=30}" "${STABLE:=0}" "${HEARTBEAT:=1800}" "${TIMEOUT:=0}"
    : "${WATCH_PID:=}" "${WATCH_PATTERN:=}" "${BODY_MAX:=300}" "${TITLE_DONE:=}"

    local ROOT="$HOME/.notify"
    local SRV_DIR="$ROOT/services" LOG_DIR="$ROOT/logs" ST_DIR="$ROOT/state"
    mkdir -p "$SRV_DIR" "$LOG_DIR" "$ST_DIR"

    local NOTIFY="$ROOT/notify"
    [ -x "$NOTIFY" ] || { echo "lib.sh: 找不到 $NOTIFY，先跑 ctl init" >&2; return 2; }

    export NOTIFY_SVC="$SVC"

    say() { echo "[$(date '+%F %H:%M:%S')] $*"; }

    # ── ① 启动前先把 CHECK 的语法验一遍 ──
    # 否则一个拼错的 CHECK 会安静地空转到超时，12 小时后才发现。
    if ! bash -n -c "$CHECK" 2>"$ST_DIR/.syntax.err"; then
        say "❌ CHECK 有语法错误，不启动："
        sed 's/^/    /' "$ST_DIR/.syntax.err"
        "$NOTIFY" warn "监控启动失败：判断命令语法错误" "$(head -3 "$ST_DIR/.syntax.err")"
        rm -f "$ST_DIR/.syntax.err"
        return 2
    fi

    # ── ② 进程锁：同一服务只允许一个实例 ──
    # 否则误开两次会推两遍通知，甚至互相覆盖状态。
    local LOCK="$ST_DIR/$SVC.lock"
    if ! mkdir "$LOCK" 2>/dev/null; then
        local old; old=$(cat "$LOCK/pid" 2>/dev/null)
        if [ -n "$old" ] && kill -0 "$old" 2>/dev/null; then
            say "❌ 该服务已在运行（pid $old），本次退出"
            return 3
        fi
        say "发现残留锁（进程 $old 已不在），清理后继续"
        rm -rf "$LOCK"; mkdir "$LOCK" 2>/dev/null || { say "❌ 无法建立锁"; return 3; }
    fi
    echo $$ > "$LOCK/pid"
    # shellcheck disable=SC2064
    trap "rm -rf '$LOCK'" EXIT INT TERM

    # ── ③ 日志轮转：单文件超过 5MB 就换一份，最多留 2 份 ──
    local LOGF="$LOG_DIR/$SVC.log"
    if [ -f "$LOGF" ] && [ "$(_mb "$LOGF")" -ge 5 ]; then
        mv -f "$LOGF.1" "$LOGF.2" 2>/dev/null
        mv -f "$LOGF"   "$LOGF.1" 2>/dev/null
    fi

    # ── ④ 头部信息 ──
    say "──────────────────────────────────────────────"
    say "服务     : $SVC"
    say "判断命令 : $CHECK"
    [ -n "$BODY" ] && say "摘要命令 : $BODY"
    say "间隔 ${INTERVAL}s | 稳定 ${STABLE}s | 心跳 ${HEARTBEAT}s | 超时 ${TIMEOUT:-无限}s"
    [ -n "$WATCH_PID" ]     && say "存活探针 : PID $WATCH_PID"
    [ -n "$WATCH_PATTERN" ] && say "存活探针 : 命令含 '$WATCH_PATTERN'"
    say "──────────────────────────────────────────────"

    "$NOTIFY" start "开始监控" "判断条件：$CHECK"

    local START NOW EL LAST_BEAT RC SEEN=0
    START=$(date +%s); LAST_BEAT=$START

    while :; do
        # ── ⑤ 跑判断命令，并区分「没满足」和「根本跑不起来」 ──
        bash -c "$CHECK" >/dev/null 2>"$ST_DIR/.check.err"
        RC=$?
        if [ "$RC" -eq 0 ]; then
            if [ "$STABLE" -gt 0 ]; then
                say "判断通过，等 ${STABLE}s 复检…"
                sleep "$STABLE"
                if ! bash -c "$CHECK" >/dev/null 2>&1; then
                    say "复检没通过（可能还在写），回到等待"; sleep "$INTERVAL"; continue
                fi
            fi
            say "✅ 判断通过，条件已满足"; break
        fi

        # 126/127 = 命令无法执行 / 命令不存在。这不是「还没好」，是「永远好不了」。
        if [ "$RC" -eq 126 ] || [ "$RC" -eq 127 ]; then
            say "❌ CHECK 无法执行（退出码 $RC）："
            sed 's/^/    /' "$ST_DIR/.check.err" 2>/dev/null | head -5
            "$NOTIFY" warn "监控失败：判断命令无法执行" \
                "退出码 $RC。$(head -2 "$ST_DIR/.check.err" 2>/dev/null | tr '\n' ' ')"
            return 4
        fi

        NOW=$(date +%s); EL=$(( NOW - START ))

        # ── ⑥ 超时出口 ──
        if [ "$TIMEOUT" -gt 0 ] && [ "$EL" -ge "$TIMEOUT" ]; then
            say "⏱ 超过 ${TIMEOUT}s 仍未满足"
            "$NOTIFY" warn "超时未完成" "等了 $(( EL / 60 )) 分钟仍未满足条件，请人工检查"
            return 1
        fi

        # ── ⑦ 存活探针：目标进程死了就立刻报，别傻等 ──
        # 进程因 OOM 或报错挂掉时，产物永远不会出现，只等超时可能等一整天。
        #
        # ⚠ 必须先用 SEEN 标记「到底见过它活着没」：
        #   监控可能在任务启动之前就拉起来了，此时进程本来就不在。
        #   不加这个判断，一启动就会误报「异常终止」（实测踩到）。
        if [ -n "$WATCH_PID" ] || [ -n "$WATCH_PATTERN" ]; then
            local alive=1
            if [ -n "$WATCH_PID" ]; then
                kill -0 "$WATCH_PID" 2>/dev/null || alive=0
            else
                pgrep -f "$WATCH_PATTERN" >/dev/null 2>&1 || alive=0
            fi

            if [ "$alive" -eq 1 ]; then
                SEEN=1
            elif [ "$SEEN" -eq 1 ]; then
                # 见过它活着 → 现在不见了，这才是真的异常退出
                say "⚠ 目标进程已消失，等 ${STABLE:-5}s 让它把最后的数据写完…"
                sleep "${STABLE:-5}"
                if ! bash -c "$CHECK" >/dev/null 2>&1; then
                    say "❌ 进程已退出但条件仍未满足 → 判定为异常终止"
                    "$NOTIFY" warn "异常终止" "目标进程已退出，但完成条件未满足，可能崩溃或被中断"
                    return 5
                fi
                say "条件已满足，按正常完成处理"; break
            fi
            # SEEN=0 且进程不在 = 任务还没开始，继续等，不报警
        fi

        # ── ⑧ 心跳：没有它，「还在等」和「脚本死了」外观完全一样 ──
        if [ "$HEARTBEAT" -gt 0 ] && [ $(( NOW - LAST_BEAT )) -ge "$HEARTBEAT" ]; then
            say "💓 还在等，已等 $(( EL / 60 )) 分钟"; LAST_BEAT=$NOW
        fi

        sleep "$INTERVAL"
    done

    # ── ⑨ 取摘要 + 推送 ──
    local SUMMARY
    if [ -n "$BODY" ]; then
        say "取摘要：$BODY"
        SUMMARY=$(bash -c "$BODY" 2>&1 | tail -1)
        [ -z "${SUMMARY// /}" ] && SUMMARY="（取摘要为空，详情见日志）"
    else
        SUMMARY="已完成，耗时 $(( ($(date +%s) - START) / 60 )) 分钟"
    fi

    if [ "${BODY_MAX:-0}" -gt 0 ] && [ "${#SUMMARY}" -gt "$BODY_MAX" ]; then
        SUMMARY=$(_cut_chars "$SUMMARY" "$BODY_MAX")
    fi

    say "──────────────────────────────────────────────"
    say "  $SUMMARY"
    say "──────────────────────────────────────────────"
    "$NOTIFY" done "${TITLE_DONE:-完成了}" "$SUMMARY"
    say "监控结束"
    rm -f "$ST_DIR/.check.err" "$ST_DIR/.syntax.err"
    return 0
}
