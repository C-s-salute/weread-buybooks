# -*- coding: utf-8 -*-
"""统一配置：工作目录 / 登录态 profile / 浏览器通道 / Node 可执行文件。

业务脚本一律从这里取路径，禁止硬编码绝对路径或用户名。

环境变量
--------
WEREAD_WORK     数据工作目录（默认：当前目录）
WEREAD_PROFILE  持久化浏览器 profile 目录（默认：~/.weread-buybooks/profile；
                给相对路径时按 WEREAD_WORK 解析，可把登录态一并放进数据目录做便携模式）
WEREAD_BROWSER  浏览器通道 msedge / chrome / chromium（默认 msedge，找不到则回退 chromium）
WEREAD_NODE     node 可执行文件（默认取 PATH 中的 node）
"""
import os
import shutil
import sys
from pathlib import Path

HOME = Path.home() / '.weread-buybooks'
WORK = Path(os.environ.get('WEREAD_WORK') or os.getcwd()).expanduser().resolve()


def _abs(p, base=None):
    """相对路径按 WORK 解析成绝对路径：整套数据（含登录态）可以只靠相对路径搬走。"""
    p = Path(p).expanduser()
    if p.is_absolute():
        return p.resolve()
    return ((base or WORK) / p).resolve()


PROFILE_DIR = str(_abs(os.environ.get('WEREAD_PROFILE') or (HOME / 'profile')))

# 业务脚本统一用相对路径读写数据文件。设置了 WEREAD_WORK 时先切到该目录，
# 这样 run_pipeline 调用、以及直接 `python scripts/xxx.py` 两种跑法都能定位数据。
if os.environ.get('WEREAD_WORK'):
    try:
        os.makedirs(WORK, exist_ok=True)
        os.chdir(WORK)
    except OSError as e:
        print('无法进入 WEREAD_WORK=%s：%s' % (WORK, e))
        sys.exit(2)

# 各通道在 Windows / macOS / Linux 上的常见安装位置
_CHANNEL_HINTS = {
    'msedge': [
        r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe',
        r'C:\Program Files\Microsoft\Edge\Application\msedge.exe',
        '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge',
    ],
    'chrome': [
        r'C:\Program Files\Google\Chrome\Application\chrome.exe',
        r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
        '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
    ],
}


def channel():
    """返回 playwright 可用的 channel 名称；None 表示用内置 chromium。"""
    want = (os.environ.get('WEREAD_BROWSER') or 'msedge').strip().lower()
    if want in ('chromium', 'bundled', 'none', ''):
        return None
    for p in _CHANNEL_HINTS.get(want, []):
        if os.path.exists(p):
            return want
    # 未装该通道：优先退回另一个已装通道，再退回内置 chromium
    for alt in ('msedge', 'chrome'):
        if alt == want:
            continue
        if any(os.path.exists(p) for p in _CHANNEL_HINTS[alt]):
            return alt
    return None


def launch_kwargs(headless=True, viewport=None):
    """持久化上下文的公共启动参数。"""
    kw = {
        'headless': headless,
        'args': ['--no-first-run', '--no-default-browser-check',
                 '--disable-blink-features=AutomationControlled',
                 '--disable-gpu'],
        'viewport': viewport or {'width': 1280, 'height': 900},
    }
    ch = channel()
    if ch:
        kw['channel'] = ch
    return kw


def node_bin():
    """node 可执行文件；找不到时返回裸名交由 PATH 解析。"""
    env = (os.environ.get('WEREAD_NODE') or '').strip()
    if env and os.path.exists(env):
        return env
    for name in ('node', 'node.exe'):
        hit = shutil.which(name)
        if hit:
            return hit
    return 'node'


def ensure_profile():
    """确保 profile 目录存在（先扫码登录过才会有 cookie）。"""
    os.makedirs(PROFILE_DIR, exist_ok=True)
    return PROFILE_DIR


def has_login(ctx):
    """检查持久化上下文里是否还有微信读书登录态。"""
    try:
        names = {c.get('name') for c in ctx.cookies()}
    except Exception:
        return False
    return 'wr_vid' in names


def info():
    return {
        'work': str(WORK),
        'profile': PROFILE_DIR,
        'browser': channel() or 'chromium(bundled)',
        'node': node_bin(),
        'python': sys.version.split()[0],
    }


if __name__ == '__main__':
    for k, v in info().items():
        print('%-8s %s' % (k, v))
