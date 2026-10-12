# -*- coding: utf-8 -*-
"""
ITO AI 工具集门户 —— 一键刷新（Python 编排层）

背景（2026-10-12 事故复盘）：
  1) 旧入口「刷新工具集门户.bat」被写入时用了 LF 行尾（CR=0/LF=76）。
     cmd.exe 解析 .bat 依赖 CRLF，LF 版本会在 if(...) 块处解析崩溃
     （实测：exit=255，报「此时不应有 else」），窗口一闪而过 —— 即"闪退"。
     实测对照：同一份字节改成 CRLF 后，同一套 4 步流程正常跑完。
     实测结论：3 行以内的极简 stub（无括号块）在 LF/CRLF 下都能跑。
  2) 「_bak_刷新工具集门户_20260908.bat」指向基础 python
     (binaries\\python\\versions\\3.13.12\\python.exe)，该解释器的 site-packages
     里没有 openpyxl → 必然 ModuleNotFoundError。
  3) 正常路径使用的 venv 位于 WorkBuddy 托管目录，平台/App 更新会整体替换
     其中的 vendor python，可能出现"包突然不在"的瞬时态 → 本脚本做依赖自愈。

因此本脚本承担原来 bat 里的全部编排逻辑，令入口 bat 只剩 stub；
编排层是 .py，行尾不敏感，且每次运行都会自检并修复入口 bat 的行尾。
"""
import os
import sys
import shutil
import subprocess

HERE = os.path.dirname(os.path.abspath(__file__))
BAT = os.path.join(HERE, "刷新工具集门户.bat")
GEN_SCRIPT = os.path.join(HERE, "build_portal.py")
INDEX = os.path.join(HERE, "index.html")

PORTABLE_GIT = (r"C:\Users\QwQ\.workbuddy\binaries\PortableGit"
                r"\versions\1.2.0\mingw64\bin\git.exe")
VENV_PY = (r"C:\Users\QwQ\.workbuddy\binaries\python\envs\default"
           r"\Scripts\python.exe")

COMMIT_MSG = "auto: update toolset portal"
NO_PUSH = "--no-push" in sys.argv   # 只生成 + 提交，不推送（用于验证或手动把控上线时机）


# ---------------------------------------------------------------- 输出工具
def say(msg):
    print(msg, flush=True)


def head(title):
    say("")
    say("=" * 32)
    say("  " + title)
    say("=" * 32)


# ------------------------------------------------- ① 入口 bat 行尾自检/自修复
def check_and_fix_bat_lineending():
    """若入口 bat 被写成 LF-only，cmd 解析会崩。这里检测并就地修回 CRLF。"""
    if not os.path.isfile(BAT):
        return
    raw = open(BAT, "rb").read()
    cr, lf = raw.count(b"\r"), raw.count(b"\n")
    if lf and cr == 0:
        say("[FIX ] 入口 bat 是 LF 行尾（会导致双击闪退），正在修回 CRLF ...")
        if not os.path.isfile(BAT + ".lf.bak"):
            shutil.copy2(BAT, BAT + ".lf.bak")
        with open(BAT, "wb") as f:
            f.write(raw.replace(b"\n", b"\r\n"))
        say("[OK  ] 已修复：%s" % BAT)
    elif cr and cr != lf:
        say("[WARN] 入口 bat 行尾混杂（CR=%d LF=%d），建议用记事本另存一次。" % (cr, lf))


# --------------------------------------------------------- ② 依赖自愈 openpyxl
def ensure_openpyxl():
    try:
        import openpyxl  # noqa: F401
        return True
    except ImportError:
        pass
    say("[INFO] 当前解释器缺少 openpyxl：%s" % sys.executable)
    say("[INFO] 正在自动安装（适配 WorkBuddy 托管环境被平台重装的情况）...")
    code = subprocess.run([sys.executable, "-m", "pip", "install",
                           "--disable-pip-version-check", "openpyxl"],
                          cwd=HERE).returncode
    if code != 0:
        say("[ERROR] 自动安装失败。请手动执行：")
        say('        "%s" -m pip install openpyxl' % sys.executable)
        return False
    try:
        import openpyxl  # noqa: F401
        say("[OK  ] openpyxl 已就绪。")
        return True
    except ImportError:
        say("[ERROR] 安装后仍无法导入 openpyxl。")
        return False


# ------------------------------------------------------------------- ③ git
def git_exe():
    if os.path.isfile(PORTABLE_GIT):
        return PORTABLE_GIT
    found = shutil.which("git")
    if found:
        say("[INFO] 便携版 git 不存在，改用 PATH 上的 git：%s" % found)
        return found
    return None


def git(args, cwd=HERE, **kw):
    exe = git_exe()
    if not exe:
        say("[ERROR] 未找到可用的 git。")
        return 1
    say("    $ git " + " ".join(args))
    env = dict(os.environ, GIT_TERMINAL_PROMPT="0")
    try:
        return subprocess.run([exe] + args, cwd=cwd, env=env, **kw).returncode
    except OSError as exc:
        say("[ERROR] git 调用失败：%s" % exc)
        return 1


def current_branch():
    exe = git_exe()
    if not exe:
        return "main"
    try:
        out = subprocess.run([exe, "rev-parse", "--abbrev-ref", "HEAD"],
                             cwd=HERE, capture_output=True, text=True,
                             env=dict(os.environ, GIT_TERMINAL_PROMPT="0"))
        name = (out.stdout or "").strip()
        return name or "main"
    except OSError:
        return "main"


# -------------------------------------------------------------------- 主流程
def main():
    head("ITO AI 工具集门户刷新")

    say("[0/4] 环境自检 ...")
    check_and_fix_bat_lineending()
    if not os.path.isfile(GEN_SCRIPT):
        say("[ERROR] 未找到生成脚本：%s" % GEN_SCRIPT)
        return 1
    if not ensure_openpyxl():
        return 1

    say("[1/4] 同步最新仓库（git pull --ff-only）...")
    if git(["pull", "--ff-only"]) != 0:
        say("[WARN] 拉取失败（可能离线或本地有冲突改动），将使用本地现有脚本继续。")
    else:
        say("[OK  ] 已同步到最新脚本与模板。")

    say("[2/4] 从 Excel 生成门户 ...")
    if subprocess.run([sys.executable, GEN_SCRIPT], cwd=HERE).returncode != 0:
        say("[ERROR] 生成门户失败（多半是 xlsx 结构变动或缺依赖）。")
        return 1

    say("[3/4] 提交到 Git ...")
    git(["add", "-A"])
    code = git(["commit", "-m", COMMIT_MSG])
    if code == 1:
        say("[INFO] 没有变化需要提交，直接结束。")
        say("")
        say("门户文件：%s" % INDEX)
        return 0
    if code != 0:
        say("[ERROR] Git 提交失败。")
        return code

    say("[4/4] 推送到 GitHub ...")
    if NO_PUSH:
        say("[SKIP] 已指定 --no-push：本地已生成并提交，未推送。")
        say("       需要上线时双击入口 bat（或运行不带 --no-push 的本脚本）。")
        say("门户文件：%s" % INDEX)
        return 0
    if git(["push", "origin", current_branch()]) != 0:
        say("[ERROR] Git 推送失败（检查网络 / 凭证）。")
        return 1
    say("")
    say("GitHub 同步完成。GitHub Pages 约 1-2 分钟后重构建，请硬刷新（Ctrl+Shift+R）查看。")
    say("门户文件：%s" % INDEX)
    return 0


if __name__ == "__main__":
    try:
        rc = main()
    except Exception:
        import traceback
        say("")
        say("[FATAL] 未捕获异常：")
        traceback.print_exc()
        rc = 1
    # 不在此处 pause：双击时的暂停由入口 bat 末尾的 pause 负责
    sys.exit(rc)
