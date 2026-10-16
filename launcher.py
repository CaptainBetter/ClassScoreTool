# -*- coding: utf-8 -*-
"""一键启动脚本（由 run.bat 调用，也可以直接 `python launcher.py` 运行）。

把"建虚拟环境 -> 装依赖 -> 启动 Streamlit"这套流程放在 Python 里做，
原因是：cmd.exe 解析 .bat 文件时用的是系统 ANSI 代码页，.bat 里写中文
很容易被拆成乱码命令（"Add Python to PATH" 被当成命令执行就是这么来的）。
Python 处理和输出中文则没有这个问题。

依赖默认走清华镜像（国内快很多），可在 run.bat 里改 CS_PIP_INDEX。
"""

from __future__ import annotations

import hashlib
import os
import subprocess
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
VENV_DIR = BASE_DIR / ".venv"
VENV_PY = VENV_DIR / "Scripts" / "python.exe"
REQUIREMENTS = BASE_DIR / "requirements.txt"
APP = BASE_DIR / "app.py"
STAMP = VENV_DIR / ".deps-ok"           # 记录上次装依赖时 requirements.txt 的内容指纹
MIN_PY = (3, 9)
NEW_PY = (3, 14)                        # 太新的 Python 可能还没有某些依赖的安装包

# 下载源：优先 run.bat 里的 CS_PIP_INDEX，其次系统的 PIP_INDEX_URL，都没有就用清华镜像。
# 想改回官方源：在 run.bat 里写 set "CS_PIP_INDEX=official"
MIRROR = "https://pypi.tuna.tsinghua.edu.cn/simple"
OFFICIAL = "https://pypi.org/simple"
OFFICIAL_ALIASES = {"official", "off", "none", "pypi", "default"}


def setup_console() -> None:
    """让 Windows 命令行能正常显示中文。"""
    if os.name == "nt":
        os.system("chcp 65001 >nul 2>nul")
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def say(text: str = "") -> None:
    print(text, flush=True)


def pick_index() -> str:
    """决定用哪个下载源；返回空字符串表示用 pip 默认的官方 PyPI。"""
    raw = os.environ.get("CS_PIP_INDEX")
    if raw is None:
        raw = os.environ.get("PIP_INDEX_URL")
    if raw is None:
        return MIRROR
    raw = raw.strip()
    if raw.lower() in OFFICIAL_ALIASES:
        return ""
    return raw or MIRROR


def requirements_digest() -> str:
    try:
        return hashlib.sha256(REQUIREMENTS.read_bytes()).hexdigest()
    except OSError:
        return ""


def run(cmd: list[str], extra_env: dict[str, str] | None = None) -> int:
    """执行命令，实时把输出显示在当前窗口。

    extra_env 会叠加在当前环境变量上，只影响这个子进程。
    """
    env = None
    if extra_env:
        env = dict(os.environ)
        env.update(extra_env)
    try:
        return subprocess.call(cmd, cwd=str(BASE_DIR), env=env)
    except FileNotFoundError:
        say(f"  找不到命令：{cmd[0]}")
        return 1


def pip_install(index: str) -> int:
    """用指定源安装依赖。

    加了 --only-binary=:all:：只接受现成的安装包（wheel），不允许 pip 去编源码。
    Windows 上一般没有 C++ 编译器，一旦它偷偷开始编源码，就会"卡很久然后失败"，
    这正是很多人卡在"安装依赖"这一步的原因。
    """
    cmd = [
        str(VENV_PY), "-m", "pip", "install",
        "--disable-pip-version-check",
        "--only-binary=:all:",
        "--prefer-binary",
        "--timeout", "30",
        "--retries", "2",
        "-r", str(REQUIREMENTS),
    ]
    if index:
        cmd += ["-i", index]
    return run(cmd)


def main() -> int:
    setup_console()
    say("=" * 52)
    say("   班级二课分统计工具 - 启动中")
    say("=" * 52)
    say()

    # ---- 1. 检查 Python 版本 ----
    if sys.version_info[:2] < MIN_PY:
        say(f"  [X] 当前 Python {sys.version_info.major}.{sys.version_info.minor} 太旧了，")
        say(f"      请安装 Python {MIN_PY[0]}.{MIN_PY[1]} 或更高版本。")
        return 1
    say(f"  Python 版本：{sys.version.split()[0]}")
    say(f"  解释器路径：{sys.executable}")
    if sys.version_info[:2] >= NEW_PY:
        say("  [!] 你的 Python 比较新。如果下面提示某个包没有现成的安装包，")
        say("      就装一个 Python 3.12，再删掉本文件夹里的 .venv 后重新双击 run.bat。")

    if not APP.exists():
        say(f"  [X] 找不到 app.py，请确认它和本脚本在同一个文件夹：{BASE_DIR}")
        return 1
    if not REQUIREMENTS.exists():
        say("  [X] 找不到 requirements.txt，请确认它和本脚本在同一个文件夹。")
        return 1

    # ---- 2. 准备虚拟环境 ----
    if not VENV_PY.exists():
        say()
        say("  [1/3] 第一次运行，正在创建虚拟环境 .venv（约 10 秒）...")
        code = run([sys.executable, "-m", "venv", str(VENV_DIR)])
        if code != 0 or not VENV_PY.exists():
            say()
            say("  [X] 创建虚拟环境失败。常见原因：")
            say("      1) 杀毒软件拦截 -> 关掉后重试；")
            say("      2) 项目放在需要管理员权限的目录（如 C:\\Program Files）-> 换个目录。")
            say("      手动做法：在本文件夹打开命令行，执行 python -m venv .venv")
            return 1
    else:
        say("  虚拟环境已存在：.venv")

    # ---- 3. 安装依赖（requirements.txt 没变就跳过，省时间）----
    digest = requirements_digest()
    if STAMP.exists() and STAMP.read_text(encoding="utf-8").strip() == digest:
        say("  [2/3] 依赖已安装，跳过。")
    else:
        index = pick_index()
        say()
        say("  [2/3] 正在安装依赖（第一次需要几分钟，要联网）...")
        say(f"        下载源：{index or 'PyPI 官方源（较慢）'}")
        say("        提示：pyarrow 这类大包会安静地下一两分钟，屏幕上暂时没动静是正常的。")
        say()
        code = pip_install(index)
        if code != 0 and index:
            say()
            say("  [!] 镜像源没成功，换成 PyPI 官方源再试一次 ...")
            code = pip_install(OFFICIAL)
        if code != 0:
            say()
            say("  [X] 依赖安装失败。看完上面的报错，对号入座：")
            say("      · no matching distribution / 没有现成的安装包")
            say("        -> 这个包还没有你当前 Python 版本的安装包。")
            say("           装个 Python 3.12，删掉 .venv 再双击 run.bat。")
            say("      · Read timed out / 连接超时")
            say("        -> 换网络或用手机热点重试；也可以手动执行：")
            say(f"           {VENV_PY} -m pip install -r requirements.txt -i {MIRROR}")
            say("      · 想换回官方源：把 run.bat 里的 CS_PIP_INDEX 改成 official")
            return 1
        try:
            STAMP.write_text(digest, encoding="utf-8")
        except OSError:
            pass

    # ---- 4. 启动网页 ----
    say()
    say("  [3/3] 正在启动网页 ... 浏览器会自动打开 http://localhost:8501")
    say("        已关闭 Streamlit 的邮箱询问与使用统计（不会再问你要 Email）")
    say("        想停止程序：在本窗口按 Ctrl + C，或直接关掉窗口。")
    say()
    # 这里三重保险，确保首次运行不会再弹出 "Welcome to Streamlit! Email:"：
    #   1) 项目里的 .streamlit/config.toml 已经设了 showEmailPrompt=false
    #   2) 命令行参数（最直接）
    #   3) 环境变量（run.bat 里也设了，老版本 Streamlit 读这个）
    # 注意：直接按回车（空邮箱）会被判定为无效，Streamlit 会以退出码 -1
    # 结束进程，这就是"窗口一闪而过、网页没起来"的原因。
    streamlit_env = {
        "STREAMLIT_SERVER_SHOW_EMAIL_PROMPT": "false",
        "STREAMLIT_BROWSER_GATHER_USAGE_STATS": "false",
    }
    streamlit_flags = [
        "--server.showEmailPrompt=false",
        "--browser.gatherUsageStats=false",
    ]
    cmd = [str(VENV_PY), "-m", "streamlit", "run", str(APP)]
    cmd += streamlit_flags + sys.argv[1:]
    return run(cmd, extra_env=streamlit_env)


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        sys.exit(0)
