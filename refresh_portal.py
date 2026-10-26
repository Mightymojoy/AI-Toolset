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

# 路径一律从当前用户 HOME 推导，禁止写死 C:\Users\<用户名>。
# PortableGit 的版本目录名（1.2.0 -> 1.3.0 -> ...）由 WorkBuddy 升级决定，
# 更不能写死 —— 具体探测逻辑见下方 git_exe()。
HOME = os.path.expanduser("~")
PGROOT = os.path.join(HOME, ".workbuddy", "binaries", "PortableGit", "versions")
VENV_PY = os.path.join(HOME, ".workbuddy", "binaries", "python", "envs",
                       "default", "Scripts", "python.exe")

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
# 2026-10-26 事故复盘：
#   PortableGit 装在 ...\binaries\PortableGit\versions\<版本号>\mingw64\bin\git.exe，
#   WorkBuddy 升级会删掉旧版本号目录（1.2.0 被删、1.3.0 新建）。
#   旧代码把 1.2.0 写死在 PORTABLE_GIT 里，升级后 isfile() 恒为 False，
#   回退 shutil.which("git") 也失效（本机未装独立 Git，WorkBuddy 自带的 git
#   只在 WorkBuddy 进程注入的 PATH 里，不是持久化环境变量），
#   于是所有 git 命令报「未找到可用的 git」：门户生成成功却推不上去，线上仍是旧版。
# 结论：git 定位必须运行时探测，且探测结果要能和「确实没有改动」区分开。
_GIT_PROBE = {"done": False, "exe": None}


def _git_candidates():
    """按优先级产出 (路径, 来源说明)。任何一级 isfile 命中即可用。"""
    # ① versions\current —— WorkBuddy 自己维护的指针文件，升级时会同步更新
    cur_file = os.path.join(PGROOT, "current")
    if os.path.isfile(cur_file):
        try:
            with open(cur_file, "r", encoding="utf-8", errors="ignore") as f:
                ver = f.read().strip()
            if ver:
                yield (os.path.join(PGROOT, ver, "mingw64", "bin", "git.exe"),
                       "versions\\current -> %s" % ver)
        except OSError:
            pass

    # ② 扫描 versions\ 各版本目录，取版本号最大的可用者
    #    （注意：不能纯字符串排序，1.10.0 会排到 1.9.0 前面）
    try:
        dirs = [d for d in os.listdir(PGROOT)
                if d != ".locks" and os.path.isdir(os.path.join(PGROOT, d))]
    except OSError:
        dirs = []

    def _verkey(name):
        return [(int(s) if s.isdigit() else -1) for s in name.split(".")]

    for d in sorted(dirs, key=_verkey, reverse=True):
        yield (os.path.join(PGROOT, d, "mingw64", "bin", "git.exe"),
               "versions\\%s（扫描）" % d)

    # ③ 系统 PATH
    found = shutil.which("git")
    if found:
        yield (found, "系统 PATH")

    # ④ WorkBuddy 安装目录的 vendor PortableGit（路径不含版本号，最稳但最靠后）
    localapp = os.environ.get("LOCALAPPDATA") or os.path.join(
        HOME, "AppData", "Local")
    yield (os.path.join(localapp, "Programs", "WorkBuddy", "resources", "vendor",
                        "PortableGit", "mingw64", "bin", "git.exe"),
           "WorkBuddy vendor")


def git_exe():
    """返回可用的 git.exe；找不到返回 None。只探测一次并缓存。"""
    if _GIT_PROBE["done"]:
        return _GIT_PROBE["exe"]

    _GIT_PROBE["done"] = True
    tried = []
    for path, how in _git_candidates():
        if path and os.path.isfile(path):
            _GIT_PROBE["exe"] = path
            say("[OK  ] git 定位成功（%s）" % how)
            say("        %s" % path)
            return path
        tried.append((path, how))

    say("[ERROR] 未找到可用的 git。已尝试以下位置：")
    for path, how in tried:
        say("        [%s] %s" % (how, path))
    say("        修复建议：确认 WorkBuddy 已安装 PortableGit，或把 git 加入系统 PATH。")
    return None


def _git_env(exe):
    """给 git 子进程准备环境变量。

    关键：只找到 git.exe 还不够。git 的很多操作（pull / push / commit 触发
    hook 等）会**再 spawn 一个 git 子进程**，它靠 PATH 查找。用户双击 bat 时
    系统 PATH 里没有 WorkBuddy 的 git，于是出现
        error: cannot spawn git: No such file or directory
    —— 表现为 pull 直接失败、commit "报错却成功"、push 必然失败。
    所以必须把 git 自己的 bin 目录前置注入 PATH。
    """
    env = dict(os.environ, GIT_TERMINAL_PROMPT="0")
    bindir = os.path.dirname(exe)                       # ...\mingw64\bin
    root = os.path.dirname(os.path.dirname(bindir))     # ...\<版本目录>
    parts = [bindir]
    usrbin = os.path.join(root, "usr", "bin")           # 内含 sh.exe 等辅助程序
    if os.path.isdir(usrbin):
        parts.append(usrbin)
    old = env.get("PATH", "")
    env["PATH"] = os.pathsep.join(parts + ([old] if old else []))
    return env


def git(args, cwd=HERE, **kw):
    exe = git_exe()
    if not exe:
        # 具体原因已由 git_exe() 完整输出，这里不再重复刷屏
        return 1
    say("    $ git " + " ".join(args))
    try:
        return subprocess.run([exe] + args, cwd=cwd, env=_git_env(exe),
                              **kw).returncode
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
                             env=_git_env(exe))
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

    # 提前探测 git：窗口一开头就显示实际使用的 git 路径；
    # 同时让后续能区分「git 不可用」与「确实没有改动」两种完全不同的结局。
    has_git = git_exe() is not None
    if not has_git:
        say("[WARN] 没有 git：本次仍会生成门户，但无法提交 / 推送。")

    say("[1/4] 同步最新仓库（git pull --ff-only）...")
    if not has_git:
        say("[SKIP] git 不可用，跳过拉取（不影响生成）。")
    elif git(["pull", "--ff-only"]) != 0:
        say("[WARN] 拉取失败（可能离线或本地有冲突改动），将使用本地现有脚本继续。")
    else:
        say("[OK  ] 已同步到最新脚本与模板。")

    say("[2/4] 从 Excel 生成门户 ...")
    if subprocess.run([sys.executable, GEN_SCRIPT], cwd=HERE).returncode != 0:
        say("[ERROR] 生成门户失败（多半是 xlsx 结构变动或缺依赖）。")
        return 1

    say("[3/4] 提交到 Git ...")
    if not has_git:
        # 关键修复：旧代码在这里会因为 git 缺失而让 commit 返回 1，
        # 被下面的 code==1 分支误判成「没有变化需要提交」，
        # 于是把一个失败说成"正常结束"。现在显式区分这两种情况。
        say("[ERROR] git 不可用 —— 本次未提交、也未推送。")
        say("        门户文件已生成（本地可打开），但线上 GitHub Pages 仍是旧版。")
        say("        修好 git 后重跑本脚本即可补上，或手动提交：%s" % INDEX)
        return 1
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
