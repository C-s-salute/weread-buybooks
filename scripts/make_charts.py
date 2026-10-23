# -*- coding: utf-8 -*-
"""v5 互动证据说明图：左=fac_e 映射曲线，右=各类型因互动带来的平均分差。"""
import json
import collections
import os

try:
    import wconfig   # noqa: F401  统一工作目录（设置了 WEREAD_WORK 时切 cwd）
except ImportError:
    pass

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.ticker import FuncFormatter

# 中文字体：按平台依次尝试，命中即注册；全都没有时退到系统默认（会提示可能显示方块）
_CJK_FONTS = [
    ('C:/Windows/Fonts/msyh.ttc', 'Microsoft YaHei'),                    # Windows
    ('C:/Windows/Fonts/msyhbd.ttc', 'Microsoft YaHei'),
    ('/System/Library/Fonts/PingFang.ttc', 'PingFang SC'),               # macOS
    ('/System/Library/Fonts/STHeiti Medium.ttc', 'Heiti SC'),
    ('/Library/Fonts/Arial Unicode.ttf', 'Arial Unicode MS'),
    ('/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc', 'Noto Sans CJK SC'),   # Linux
    ('/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc', 'WenQuanYi Zen Hei'),
    ('/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf', 'Droid Sans Fallback'),
]
_found = []
for _path, _name in _CJK_FONTS:
    if os.path.exists(_path):
        try:
            font_manager.fontManager.addfont(_path)
            _found.append(_name)
        except Exception:
            pass
if not _found:
    print('警告：未找到中文字体，图表中文可能显示为方块。可装 Noto Sans CJK 或用 WEREAD_FONT 指定字体文件。')
_env_font = os.environ.get('WEREAD_FONT')
if _env_font and os.path.exists(_env_font):
    try:
        font_manager.fontManager.addfont(_env_font)
        _found.insert(0, font_manager.FontProperties(fname=_env_font).get_name())
    except Exception:
        pass
plt.rcParams['font.sans-serif'] = _found + ['Microsoft YaHei', 'PingFang SC',
                                            'Noto Sans CJK SC', 'WenQuanYi Zen Hei',
                                            'DejaVu Sans', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False

RED, GREEN, BLUE, GREY = '#CC0000', '#006600', '#2F5496', '#7F7F7F'

V = json.load(open('_valuescore5.json', encoding='utf-8'))
P = V['params']
S = V['scores']

FLOOR, CEIL, AT = P['FLOOR'], P['CEIL'], P['AT_MEAN']
KE, WE, EF, EC = P['K_E'], P['W_E'], P['E_FLOOR'], P['E_CEIL']
N_HILL = P['N_HILL']


def hill(x, k, n=N_HILL):
    if x <= 0:
        return 0.0
    return x ** n / (k ** n + x ** n)


def fac_of(x):
    ce = FLOOR + (CEIL - FLOOR) * hill(x, KE)
    return min(max((ce / AT) ** WE, EF), EC)


fig = plt.figure(figsize=(20, 9.6), dpi=150)
gs = fig.add_gridspec(1, 2, width_ratios=[1.02, 1], wspace=0.25,
                      left=0.06, right=0.975, top=0.80, bottom=0.245)

# ---------------- 左：映射曲线 ----------------
ax = fig.add_subplot(gs[0, 0])
xs = [0.03 + i * 0.004 for i in range(1, 800)]
ys = [fac_of(x) for x in xs]
ax.plot(xs, ys, color=BLUE, lw=3.4, zorder=3)
ax.axhline(1.0, color=GREY, ls='--', lw=1.4)
ax.axvline(1.0, color=GREY, ls='--', lw=1.4)
ax.set_xlim(0, 3.2)
ax.set_ylim(0.865, 1.075)
ax.set_xlabel('x_e =（划线密度^0.55 × 想法密度^0.45）÷ 本类中位数', fontsize=20, labelpad=10)
ax.set_ylabel('互动修正 fac_e', fontsize=21, labelpad=12)
ax.set_title('（a）互动因子是温和的补充证据：最多 +6%，最多 −12%',
             fontsize=25, pad=22, fontweight='bold', color='#1F3864')
ax.tick_params(labelsize=17)
ax.yaxis.set_major_formatter(FuncFormatter(lambda v, p: '%.2f' % v))
ax.grid(ls=':', color='#D9D9D9')
ax.set_axisbelow(True)
for s in ('top', 'right'):
    ax.spines[s].set_visible(False)

case = next((v for v in S.values() if '棋经十三篇' in (v.get('title') or '') and v.get('fac_e')), None)
marks = []
if case:
    marks.append((case['x_e'] or 0.7, case['fac_e'],
                  '《棋经十三篇》\nx_e=%.2f → fac_e=%.2f%s'
                  % (case['x_e'], case['fac_e'],
                     '（几乎不变）' if abs(case['fac_e'] - 1) < 0.005
                     else '（%+.0f%%）' % ((case['fac_e'] - 1) * 100)),
                  (1.12, 0.955), 'left'))
marks.append((2.4, fac_of(2.4), '互动明显强 → 封顶 %.2f' % EC, (1.78, 1.068), 'left'))
for xv, yv, note, xy, ha in marks:
    ax.scatter([xv], [yv], s=180, color=RED, zorder=6, edgecolor='white', lw=1.8)
    ax.annotate(note, (xv, yv), xytext=xy, fontsize=17, color='#262626', ha=ha, va='center',
                bbox=dict(boxstyle='round,pad=0.38', fc='white', ec='#BFBFBF', alpha=0.97),
                arrowprops=dict(arrowstyle='-', color='#808080', lw=1.2,
                                shrinkA=2, shrinkB=6))

ax.text(0.985, 0.04, 'Hill 曲线 K=%.2f、指数 n=%.2f、整体强度 W=%.2f'
        % (KE, N_HILL, WE), transform=ax.transAxes, ha='right', fontsize=15, color=GREY)

# ---------------- 右：按类型的平均分差 ----------------
ax2 = fig.add_subplot(gs[0, 1])
by = collections.defaultdict(list)
for v in S.values():
    if v.get('fac_e') is not None and v.get('v4_score'):
        by[v['c1']].append(v['v5' if False else 'score'] - v['v4_score'])
rows = [(k, sum(v) / len(v), len(v)) for k, v in by.items() if len(v) >= 10]
rows.sort(key=lambda r: r[1])

names = [r[0] for r in rows]
vals = [r[1] for r in rows]
cols = [GREEN if v >= 0 else RED for v in vals]
ax2.barh(range(len(names)), vals, color=cols, height=0.62, alpha=0.9, zorder=3)
ax2.set_yticks(range(len(names)))
ax2.set_yticklabels(names, fontsize=17)
ax2.axvline(0, color='#404040', lw=1.2)
ax2.set_xlabel('该类型平均分差 = v5 − v4（分）', fontsize=20, labelpad=10)
ax2.set_title('（b）各类型的价值分因互动证据升降多少', fontsize=25, pad=22,
              fontweight='bold', color='#1F3864')
ax2.tick_params(axis='x', labelsize=17)
lo, hi = min(vals + [0]), max(vals + [0])
ax2.set_xlim(lo - 1.2, hi + 1.6)
ax2.grid(axis='x', ls=':', color='#D9D9D9')
ax2.set_axisbelow(True)
for s in ('top', 'right'):
    ax2.spines[s].set_visible(False)
for i, v in enumerate(vals):
    ax2.text(v + (0.12 if v >= 0 else -0.12), i, '%+.1f' % v, va='center',
             ha='left' if v >= 0 else 'right', fontsize=16, color='#262626')

fig.suptitle('价值评分 v5 的第四层：把「划线 / 想法」按类别做成互动证据',
             fontsize=27, y=0.945, fontweight='bold', color='#1F3864')
fig.text(0.5, 0.088,
         'x_mark = 最热一条划线的人数 ÷ 阅读人数 ÷ 本类中位数    ·    x_idea = 想法数 ÷ 阅读人数 ÷ 本类中位数',
         ha='center', fontsize=16, color='#404040')
fig.text(0.5, 0.032,
         '用「人均」而非绝对量：绝对量与阅读人数的 log 相关为 0.91（划线）/ 0.94（想法），直接用等于重复奖励热度',
         ha='center', fontsize=15, color='#7F7F7F')
fig.savefig('fig_v5_interaction.png', facecolor='white')
print('fig_v5_interaction.png')
