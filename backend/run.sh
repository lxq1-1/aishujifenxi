#!/usr/bin/env bash
# =============================================================================
#  AI 数据处理工作台 · 后端启动脚本 (兼容 macOS / Linux)
#
#  设计目标:
#    1. 在前后台都能稳定启动,不被 set -e + & 的副作用吞掉
#    2. 显式把 PID 写到 backend.pid,日志写到 logs/backend.log,方便排查
#    3. 端口被占用时给清晰提示,而不是静默退出
#    4. 不依赖 setsid / timeout 等 macOS 默认没装的工具
#
#  用法:
#    ./run.sh              # 前台启动 (推荐,日志直接看到)
#    ./run.sh start        # 后台启动
#    ./run.sh stop         # 后台停止
#    ./run.sh status       # 查看后台状态
#    ./run.sh tail         # 跟踪后台日志
# =============================================================================

set -u  # 不用 -e,因为我们希望每一步都明确处理失败
cd "$(dirname "$0")"

PYTHON_BIN="python3"
VENV_DIR=".venv"
LOG_DIR="logs"
LOG_FILE="$LOG_DIR/backend.log"
PID_FILE="backend.pid"
PORT="${PORT:-8000}"
HOST="${HOST:-0.0.0.0}"

mkdir -p "$LOG_DIR"

# ---------- 准备虚拟环境 ----------
if [ ! -d "$VENV_DIR" ]; then
    echo "[run.sh] 创建虚拟环境 $VENV_DIR ..."
    "$PYTHON_BIN" -m venv "$VENV_DIR" || { echo "[run.sh] venv 创建失败"; exit 1; }
fi
# shellcheck disable=SC1091
source "$VENV_DIR/bin/activate"

# ---------- 准备依赖 (失败继续尝试启动,避免一次 pip 抖动就把人卡死) ----------
# 策略:
#   1. 先快速检查关键包是否已就绪 (uvicorn/fastapi 都在则直接跳过)
#   2. 缺失时按 清华 -> 阿里 -> 官方 的顺序自动选源
#   3. 加 --prefer-binary 与 --timeout 30,降低镜像源被 reset 的概率
#   4. 仍失败: 给出明确提示,但不阻断 uvicorn 启动
echo "[run.sh] 检查依赖 ..."
if python -c "import uvicorn, fastapi" 2>/dev/null; then
    echo "[run.sh] 关键依赖已就绪,跳过 pip install"
else
    MIRRORS=(
        "https://pypi.tuna.tsinghua.edu.cn/simple"
        "https://mirrors.aliyun.com/pypi/simple"
        "https://pypi.org/simple"
    )
    ok=0
    for m in "${MIRRORS[@]}"; do
        echo "[run.sh] 尝试源: $m"
        if pip install -q --prefer-binary --timeout 30 -i "$m" -r requirements.txt; then
            ok=1
            break
        else
            echo "[run.sh]   该源失败,尝试下一个"
        fi
    done
    if [ "$ok" -eq 0 ]; then
        echo "[run.sh] 警告: 全部镜像源都失败,继续尝试启动 (若运行报错再回头修)"
    fi
fi

# ---------- 帮助函数 ----------
is_running() {
    if [ -f "$PID_FILE" ]; then
        local pid
        pid=$(cat "$PID_FILE" 2>/dev/null || echo "")
        if [ -n "$pid" ] && kill -0 "$pid" 2>/dev/null; then
            echo "$pid"
            return 0
        fi
    fi
    return 1
}

port_in_use() {
    # macOS 自带 lsof; 若 lsof 也不在则降级用 python 检测
    if command -v lsof >/dev/null 2>&1; then
        lsof -ti tcp:"$PORT" 2>/dev/null | head -n1
    else
        "$PYTHON_BIN" - <<PY 2>/dev/null
import socket
s = socket.socket()
try:
    s.bind(("0.0.0.0", $PORT))
    print("")
except OSError:
    print("$PORT")
finally:
    s.close()
PY
    fi
}

# ---------- 子命令 ----------
cmd="${1:-foreground}"

case "$cmd" in
    stop)
        if pid=$(is_running); then
            echo "[run.sh] 停止后端 PID=$pid"
            kill "$pid" 2>/dev/null || true
            sleep 1
            if kill -0 "$pid" 2>/dev/null; then
                echo "[run.sh] 仍在运行,强制 kill -9"
                kill -9 "$pid" 2>/dev/null || true
            fi
            rm -f "$PID_FILE"
        else
            echo "[run.sh] 未发现后台进程 (PID 文件不存在或进程已退出)"
        fi
        # 兜底: 杀掉所有占用端口的 uvicorn
        if pids=$(port_in_use) && [ -n "$pids" ]; then
            echo "[run.sh] 端口 $PORT 还被 PID=$pids 占用,一并清理"
            kill -9 $pids 2>/dev/null || true
        fi
        ;;

    status)
        if pid=$(is_running); then
            echo "[run.sh] 运行中 · PID=$pid · 端口=$PORT · 日志=$LOG_FILE"
        else
            echo "[run.sh] 未运行"
        fi
        if occ=$(port_in_use) && [ -n "$occ" ]; then
            echo "[run.sh] 注意: 端口 $PORT 被 PID=$occ 占用"
        fi
        ;;

    tail)
        if [ -f "$LOG_FILE" ]; then
            tail -n 200 -f "$LOG_FILE"
        else
            echo "[run.sh] 日志文件不存在: $LOG_FILE"
        fi
        ;;

    start)
        if pid=$(is_running); then
            echo "[run.sh] 已经在运行 · PID=$pid"
            exit 0
        fi
        if occ=$(port_in_use) && [ -n "$occ" ]; then
            echo "[run.sh] 端口 $PORT 被 PID=$occ 占用,先停止或换个端口"
            exit 1
        fi
        echo "[run.sh] 后台启动 · 端口=$PORT"
        # macOS 友好: 用 nohup + 完全重定向 (替代 setsid)
        nohup python -m uvicorn app.main:app --host "$HOST" --port "$PORT" \
            > "$LOG_FILE" 2>&1 < /dev/null &
        echo $! > "$PID_FILE"
        disown 2>/dev/null || true
        sleep 2
        if pid=$(is_running); then
            echo "[run.sh] 已启动 · PID=$pid"
            echo "[run.sh] 日志: $LOG_FILE  (tail -f 查看)"
        else
            echo "[run.sh] 启动失败,最近日志:"
            tail -n 40 "$LOG_FILE" 2>/dev/null || echo " (日志未生成)"
            exit 1
        fi
        ;;

    foreground|"")
        if occ=$(port_in_use) && [ -n "$occ" ]; then
            echo "[run.sh] 端口 $PORT 被 PID=$occ 占用"
            echo "          想后台跑请先:  ./run.sh stop   或   PORT=8001 ./run.sh"
            exit 1
        fi
        echo "[run.sh] 前台启动 · http://$HOST:$PORT  (Ctrl+C 退出)"
        exec python -m uvicorn app.main:app --host "$HOST" --port "$PORT"
        ;;

    *)
        echo "用法: $0 [foreground|start|stop|status|tail]"
        exit 2
        ;;
esac
