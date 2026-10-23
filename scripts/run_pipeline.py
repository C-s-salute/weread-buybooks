# -*- coding: utf-8 -*-
"""一条命令跑完「抓取 → 评分 → Excel → Markdown → HTML → 图表」。

用法：
    python run_pipeline.py                       # 全流程，数据落在 ./weread_data
    python run_pipeline.py --work ./my_data      # 指定数据目录
    python run_pipeline.py --only score,excel,html,charts   # 只重算报告（已有数据）
    python run_pipeline.py --only fetch          # 只刷数据
    python run_pipeline.py --only booklist       # 只刷公开书单相关数据

阶段（按依赖排序）：
    fetch     账户资产 / 书架 / 书目元数据 / 阅读进度      → _wallet3, _web_shelf2, _bookmeta_v5, _progress2
    booklist  公开书单 + 书单内未加架书的详情与互动        → _booklists, _bookdetail_extra, _bookmeta_v5_extra
    score     价值评分                                    → _valuescore5.json
    excel     工作簿（顺带生成 Markdown 口径说明）        → 微信读书_权益分类与购书方案.xlsx
    html      购书方案 HTML 报告                          → 微信读书_购书方案.html
    charts    互动证据说明图                              → fig_v5_interaction.png
"""
import argparse
import os
import subprocess
import sys
import time

SCRIPTS = os.path.dirname(os.path.abspath(__file__))
STAGE_ORDER = ['fetch', 'booklist', 'score', 'excel', 'html', 'charts']

STAGE_CMD = {
    'fetch': [['fetch_all.py']],
    'booklist': [['fetch_booklist.py'], ['fetch_bookdetail.py'], ['fetch_extra.py']],
    'score': [['score_value.py']],
    'excel': [['build_excel.py']],
    'html': [['build_html.py']],
    'charts': [['make_charts.py']],
}

REQUIRED = {
    'score': ['_web_shelf2.json', '_bookmeta_v5.json'],
    'excel': ['_web_shelf2.json', '_bookmeta_v5.json', '_valuescore5.json'],
    'html': ['_web_shelf2.json', '_bookmeta_v5.json', '_valuescore5.json'],
    'charts': ['_valuescore5.json'],
}

ARTIFACTS = [
    '微信读书_权益分类与购书方案.xlsx',
    '微信读书_购书方案.html',
    '微信读书_方法与口径说明.md',
    'fig_v5_interaction.png',
]


def env_for(work):
    env = dict(os.environ)
    env['WEREAD_WORK'] = work
    env['PYTHONPATH'] = SCRIPTS + (os.pathsep + env['PYTHONPATH'] if env.get('PYTHONPATH') else '')
    env['PYTHONIOENCODING'] = 'utf-8'
    return env


def check(work, stage):
    """只检查「本阶段的输入」是否就绪（前序阶段刚生成的产物不算缺失）。"""
    missing = [f for f in REQUIRED.get(stage, [])
               if not os.path.exists(os.path.join(work, f))]
    if missing:
        print('缺少数据文件，无法执行 %s 阶段：' % stage)
        for m in missing:
            print('  -', m)
        print('请先跑前序阶段（或 --only fetch / --only booklist）')
        return False
    return True


def run(work, stage):
    env = env_for(work)
    for cmd in STAGE_CMD[stage]:
        script = os.path.join(SCRIPTS, cmd[0])
        t0 = time.time()
        print('\n=== [%s] %s ===' % (stage, cmd[0]), flush=True)
        r = subprocess.run([sys.executable, script] + cmd[1:], cwd=work, env=env)
        if r.returncode != 0:
            print('!! %s 失败（exit=%d），已中止' % (cmd[0], r.returncode))
            return r.returncode
        print('--- %s 完成，用时 %.1f 秒' % (cmd[0], time.time() - t0), flush=True)
    return 0


def main():
    ap = argparse.ArgumentParser(description='微信读书购书方案全流程')
    ap.add_argument('--work', default='./weread_data', help='数据工作目录（默认 ./weread_data）')
    ap.add_argument('--only', default='', help='只跑指定阶段，逗号分隔：' + ','.join(STAGE_ORDER))
    ap.add_argument('--skip', default='', help='跳过指定阶段，逗号分隔')
    args = ap.parse_args()

    work = os.path.abspath(args.work)
    os.makedirs(work, exist_ok=True)

    stages = STAGE_ORDER[:]
    if args.only:
        want = [s.strip() for s in args.only.split(',') if s.strip()]
        bad = [s for s in want if s not in STAGE_ORDER]
        if bad:
            ap.error('未知阶段：%s（可选：%s）' % (','.join(bad), ','.join(STAGE_ORDER)))
        stages = [s for s in STAGE_ORDER if s in want]
    if args.skip:
        skip = {s.strip() for s in args.skip.split(',') if s.strip()}
        stages = [s for s in stages if s not in skip]

    print('工作目录:', work)
    print('执行阶段:', ' → '.join(stages))

    t0 = time.time()
    for st in stages:
        if not check(work, st):
            return 2
        rc = run(work, st)
        if rc:
            return rc

    print('\n=== 全部完成，总用时 %.1f 分钟 ===' % ((time.time() - t0) / 60))
    for a in ARTIFACTS:
        p = os.path.join(work, a)
        if os.path.exists(p):
            print('  %-42s %8.1f KB' % (a, os.path.getsize(p) / 1024.0))
    return 0


if __name__ == '__main__':
    sys.exit(main())
