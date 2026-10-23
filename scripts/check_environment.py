# -*- coding: utf-8 -*-
"""环境自检：一条命令看清「能不能跑」，不能跑时直接给修复命令。

用法：
    python scripts/check_environment.py

输出三态之一（供"环境就绪路由"使用）：
    ready       依赖齐备，直接进入抓取与报告
    partial     核心依赖齐备、可选依赖缺失；受影响功能见该行提示
    needs_setup 有阻塞项，必须先修；退出码 1
另有 unavailable 不在本脚本判定——登录态失效/被限流属于运行时故障，
由 fetch 阶段在遇到 wr_vid 缺失或空响应时报告。

为什么需要它：本技能的依赖里有 matplotlib / numpy，两者都有 C 扩展，
必须命中官方预编译 wheel。Python 版本太低或用了 32 位解释器时，pip 会
退回去编译源码（缺 MSVC 就报错），或根本找不到匹配的 wheel ——
报错信息往往只说"找不到版本"，看不出真实原因。这里把判断说清楚。
"""
import importlib.metadata as md
import os
import platform
import shutil
import sys

# 依赖下限：(包名, 最低版本, 用途)
DEPS = [
    ('openpyxl', '3.1', 'Excel 工作簿'),
    ('requests', '2.31', '抓公开书单页'),
    ('playwright', '1.40', '浏览器自动化抓取'),
    ('matplotlib', '3.8', '互动证据图表'),
    ('numpy', '1.26', 'matplotlib 的依赖 + 数值计算'),
]

MIN_PY = (3, 12)          # numpy 2.5 要求 >= 3.12；matplotlib 3.11 要求 >= 3.11
REC_PY = (3, 13)
OK, WARN, BAD = '  [OK]  ', '  [警告] ', '  [阻塞] '


def ver_tuple(s):
    out = []
    for part in str(s).split('.')[:3]:
        num = ''
        for ch in part:
            if ch.isdigit():
                num += ch
            else:
                break
        out.append(int(num) if num else 0)
    while len(out) < 3:
        out.append(0)
    return tuple(out)


def check_python():
    v = sys.version_info
    bits = platform.architecture()[0]
    is64 = bits == '64bit' and sys.maxsize > 2 ** 32
    print('Python %d.%d.%d  (%s, %s)' % (v.major, v.minor, v.micro, bits, platform.platform()))
    bad = []
    if not is64:
        bad.append('Python 是 32 位（%s）。matplotlib / numpy 的官方 wheel 只出 64 位，'
                   '32 位下 pip 只能尝试源码编译，基本必然失败。\n'
                   '          → 卸掉换装 64 位 Python：https://www.python.org/downloads/windows/ '
                   '（安装时勾选 "Install for all users"，别选 32-bit 安装包）' % bits)
    if (v.major, v.minor) < MIN_PY:
        bad.append('Python %d.%d 低于最低要求 %d.%d。最新 numpy 已要求 Python >= 3.12，'
                   'matplotlib 要求 >= 3.11；版本不够时 pip 只能装老版本，或者报"找不到匹配的版本"。\n'
                   '          → 装 Python %d.%d（64 位）后重建虚拟环境：\n'
                   '            py -%d.%d -m venv .venv && .venv\\Scripts\\activate\n'
                   '            python -m pip install -U pip\n'
                   '            pip install -r requirements.txt'
                   % (v.major, v.minor, MIN_PY[0], MIN_PY[1],
                      REC_PY[0], REC_PY[1], REC_PY[0], REC_PY[1]))
    elif (v.major, v.minor) > REC_PY:
        print(WARN + 'Python %d.%d 比推荐版本 %d.%d 新，若某个 wheel 还没跟上，可回退到 %d.%d'
              % (v.major, v.minor, REC_PY[0], REC_PY[1], REC_PY[0], REC_PY[1]))
    else:
        print(OK + 'Python 版本与位数符合要求')
    return bad, is64


def check_pip():
    """pip 太旧会解析不到新 wheel（老 pip 不认识新的 wheel tag）。"""
    try:
        pv = md.version('pip')
    except Exception:
        print(WARN + 'pip 版本未知，建议先执行：python -m pip install -U pip')
        return []
    if ver_tuple(pv) < (22, 0):
        print(WARN + 'pip %s 偏旧，可能解析不到新 wheel' % pv)
        return ['pip %s 太旧：python -m pip install -U pip' % pv]
    print(OK + 'pip %s' % pv)
    return []


def check_deps():
    bad, warn = [], []
    for name, need, use in DEPS:
        try:
            got = md.version(name)
        except Exception:
            bad.append('缺 %s（%s）：pip install "%s>=%s"' % (name, use, name, need))
            print(BAD + '%-12s 未安装    （%s）' % (name, use))
            continue
        if ver_tuple(got) < ver_tuple(need):
            warn.append('%s %s 低于建议版本 %s：pip install -U "%s>=%s"'
                        % (name, got, need, name, need))
            print(WARN + '%-12s %-9s 低于建议 %s' % (name, got, need))
        else:
            print(OK + '%-12s %-9s' % (name, got))
    return bad, warn


def check_browser():
    """playwright 需要一个可用通道；本机有 Edge/Chrome 就不必再下 chromium。"""
    try:
        import wconfig
        ch = wconfig.channel()
    except Exception as e:
        print(WARN + '无法探测浏览器通道：%s' % e)
        return []
    if ch:
        print(OK + '浏览器通道：%s' % ch)
        return []
    print(WARN + '没找到 Edge / Chrome / Playwright chromium')
    return ['缺浏览器：本机装 Edge 或 Chrome 即可（推荐），或执行 python -m playwright install chromium']


def check_node():
    try:
        import wconfig
        node = wconfig.node_bin()
    except Exception:
        node = shutil.which('node')
    if node and os.path.exists(node):
        print(OK + 'node：%s' % node)
        return []
    print(WARN + 'node 未找到（只有抓「公开书单」需要，不影响书架分析与购书方案）')
    return []


def main():
    print('=' * 62)
    print('weread-buybooks 环境自检')
    print('=' * 62)
    blockers, warnings = [], []

    b, _ = check_python()
    blockers += b
    print('-' * 62)
    blockers += check_pip()
    print('-' * 62)
    b, w = check_deps()
    blockers += b
    warnings += w
    print('-' * 62)
    warnings += check_browser()
    warnings += check_node()
    print('=' * 62)

    if blockers:
        print('总体状态：needs_setup —— 还跑不起来，按下面的命令修：\n')
        for x in blockers:
            print('  - ' + x)
        if warnings:
            print('\n顺带处理：')
            for x in warnings:
                print('  - ' + x)
        return 1
    if warnings:
        print('总体状态：partial —— 核心依赖齐了，可以直接跑；下面是可选优化：\n')
        for x in warnings:
            print('  - ' + x)
        return 0
    print('总体状态：ready —— 环境完全就绪，执行：python scripts/login.py 扫码后 run_pipeline.py')
    return 0


if __name__ == '__main__':
    sys.exit(main())
