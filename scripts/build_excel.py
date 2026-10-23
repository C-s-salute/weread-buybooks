# -*- coding: utf-8 -*-
"""微信读书书架权益分类与付费价格统计 v3

数据源：_web_shelf.json（/web/shelf/sync，登录态，含 paid / payingStatus / price / centPrice）
口径已用「阅读器页付费墙文案」实地验证：
  paid=1                      -> 永久可读（已购买/已拥有）
  paid=0 & payingStatus=2     -> 付费会员卡可读（页面显示"会员卡 19元/月 · 开通后即可阅读"）
  paid=0 & payingStatus=1     -> 体验卡可读（页面显示"你正在使用体验卡阅读付费部分"）
  paid=0 & payingStatus=0     -> 免费
输出：微信读书_权益分类与付费统计.xlsx
"""
import json
import os
import re
from datetime import datetime
from collections import Counter, defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:
    import wconfig   # noqa: F401  统一工作目录（设置了 WEREAD_WORK 时切 cwd）
except ImportError:
    pass

SHELF = json.load(open('_web_shelf2.json', encoding='utf-8'))
BOOKS = SHELF['books']
ARCHIVES = SHELF.get('archive') or []
try:
    WALLET = json.load(open('_wallet3.json', encoding='utf-8'))
except Exception:
    WALLET = None
BY_ID = {b['bookId']: b for b in BOOKS}


def _load_app_ref():
    """可选的 App 侧四类数量（用户自己 App 截屏的数字），用于「本表 vs App」对照。

    文件：_app_ref.json，形如
        {"付费会员卡可读": 0, "体验卡可读": 0, "永久可读": 0, "免费": 0}
    没有这个文件就跳过对照，只呈现本表统计——不要把任何人的 App 数字写死在代码里。
    """
    try:
        d = json.load(open('_app_ref.json', encoding='utf-8'))
    except Exception:
        return {}
    out = {}
    for k in ('付费会员卡可读', '体验卡可读', '永久可读', '免费'):
        v = d.get(k)
        if isinstance(v, (int, float)):
            out[k] = int(v)
    return out if len(out) == 4 else {}


APP_REF = _load_app_ref()

HEAD_FILL = PatternFill('solid', fgColor='2F5496')
HEAD_FONT = Font(bold=True, color='FFFFFF', size=10, name='微软雅黑')
TITLE_FONT = Font(bold=True, size=15, color='1F3864', name='微软雅黑')
SUB_FONT = Font(bold=True, size=10, color='1F3864', name='微软雅黑')
BODY_FONT = Font(size=10, name='微软雅黑')
THIN = Side(style='thin', color='BFBFBF')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
WRAP = Alignment(vertical='top', wrap_text=True)
MONEY = '#,##0.00'

CAT_COLOR = {
    '永久可读': 'E2EFDA',        # 绿：已拥有
    '付费会员卡可读': 'FCE4E4',  # 红：需付费会员卡
    '体验卡可读': 'FFF2CC',      # 黄：体验卡可读
    '免费': 'EDEDED',            # 灰：免费
}
CAT_ORDER = ['付费会员卡可读', '体验卡可读', '永久可读', '免费']


def cat_of(b):
    if str(b.get('paid')) == '1':
        return '永久可读'
    ps = str(b.get('payingStatus'))
    return {'0': '免费', '2': '付费会员卡可读', '1': '体验卡可读'}.get(ps, '其他(ps=%s)' % ps)


def price_of(b):
    """真实售价（元）：centPrice 单位为分，price=-1 等异常时回退 centPrice。"""
    c = b.get('centPrice')
    try:
        c = float(c)
        if c and c > 0:
            return round(c / 100.0, 2)
    except Exception:
        pass
    try:
        p = float(b.get('price'))
        if p and p > 0:
            return round(p, 2)
    except Exception:
        pass
    return 0.0


for b in BOOKS:
    b['_cat'] = cat_of(b)
    b['_price'] = price_of(b)

# 分类聚合
agg = defaultdict(lambda: {'n': 0, 'money': 0.0})
for b in BOOKS:
    agg[b['_cat']]['n'] += 1
    agg[b['_cat']]['money'] += b['_price']

TOTAL_N = len(BOOKS)
TOTAL_MONEY = round(sum(b['_price'] for b in BOOKS), 2)
NEED = [b for b in BOOKS if b['_cat'] in ('付费会员卡可读', '体验卡可读')]
NEED_MONEY = round(sum(b['_price'] for b in NEED), 2)
NEED_SOLDOUT = sum(1 for b in NEED if b.get('soldout') == 1)

# ---------- 赠币（积分）可直购 ----------
COIN = 0.0
COIN_EXPIRE = ''
if WALLET:
    _wa = WALLET.get('android') or {}
    COIN = round(float(_wa.get('giftBalance') or (WALLET.get('ios') or {}).get('peerGiftBalance') or 0), 2)
    _mc = WALLET.get('memberCard') or {}
    if _mc.get('expiredTime'):
        COIN_EXPIRE = datetime.fromtimestamp(_mc['expiredTime']).strftime('%Y-%m-%d %H:%M:%S')

# 标价 ≤ 赠币余额 => 可用积分一次性直购
COIN_OK = [b for b in NEED if 0 < b['_price'] <= COIN + 1e-9] if COIN > 0 else []
COIN_OK_MONEY = round(sum(b['_price'] for b in COIN_OK), 2)
COIN_OK_SOLDOUT = sum(1 for b in COIN_OK if b.get('soldout') == 1)
# 已下架的买不到，方案计算时剔除
COIN_BUY = [b for b in COIN_OK if b.get('soldout') != 1]
COIN_BUY_MONEY = round(sum(b['_price'] for b in COIN_BUY), 2)


def best_bundle(books, budget):
    """0/1 背包精确解：在预算内买到"本数最多"的组合；本数相同时取花费最少者。"""
    B = int(round(budget * 100))
    cost = [int(round(b['_price'] * 100)) for b in books]
    n = len(cost)
    if not n or B <= 0:
        return []
    f = [[0] * (B + 1) for _ in range(n + 1)]
    for i in range(1, n + 1):
        ci, fi, fp = cost[i - 1], f[i], f[i - 1]
        for w in range(B + 1):
            best = fp[w]
            if ci <= w:
                v = fp[w - ci] + 1
                if v > best:
                    best = v
            fi[w] = best
    best_c, best_w = 0, 0
    for w in range(B + 1):
        if f[n][w] > best_c:
            best_c, best_w = f[n][w], w
    picked, w = [], best_w
    for i in range(n, 0, -1):
        if f[i][w] > f[i - 1][w]:
            picked.append(books[i - 1])
            w -= cost[i - 1]
    return picked


COIN_PLAN = sorted(best_bundle(COIN_BUY, COIN), key=lambda x: x['_price'])
COIN_PLAN_SPENT = round(sum(b['_price'] for b in COIN_PLAN), 2)
COIN_PLAN_LEFT = round(COIN - COIN_PLAN_SPENT, 2)
COIN_SHEET = '积分购买方案'

# 书架分组归属（注意：shelf/sync 的 archive[] = 书架分组，非 App「书单」）
_ARCH_ROWS = sum(len(_a.get('bookIds') or []) for _a in ARCHIVES)
BOOK_ARCH = defaultdict(list)
for _a in ARCHIVES:
    for _bid in (_a.get('bookIds') or []):
        BOOK_ARCH[str(_bid)].append(_a.get('name') or '未命名书单')

# App 书单（/misc/booklist 分享页）：bookId -> 书单名列表
BOOK_BL = defaultdict(list)
try:
    for _r in json.load(open('_booklists.json', encoding='utf-8')).values():
        for _b in _r['books']:
            BOOK_BL[str(_b['bookId'])].append(_r.get('name') or '未命名书单')
except Exception:
    pass

# ---------- 评分 / 出版信息 / 下架风险（v5 刷新结果优先，旧文件补缺失字段）----------
def _load_meta():
    try:
        old = json.load(open('_bookmeta.json', encoding='utf-8'))
    except Exception:
        old = {}
    try:
        new = json.load(open('_bookmeta_v5.json', encoding='utf-8'))
    except Exception:
        new = {}
    m = {}
    for k, v in old.items():
        r = dict(v)
        nv = new.get(k) or {}
        for kk, vv in nv.items():
            if vv is not None:
                r[kk] = vv
        m[k] = r
    for k, nv in new.items():
        if k not in m:
            m[k] = dict(nv)
    return m


BOOKMETA = _load_meta()

# 数据快照时间（每次运行都重新刷新底层，这里如实展示抓取时间）
def _snap_time(fn):
    try:
        return datetime.fromtimestamp(os.path.getmtime(fn)).strftime('%Y-%m-%d %H:%M')
    except Exception:
        return '—'


SNAP = {
    '账户资产': _snap_time('_wallet3.json'),
    '书架': _snap_time('_web_shelf2.json'),
    '书目/评分/阅读人数': _snap_time('_bookmeta_v5.json'),
    '阅读进度': _snap_time('_progress2.json'),
    '价值评分': _snap_time('_valuescore5.json'),
    '公开书单': _snap_time('_booklists.json'),
}
SNAP_NOW = datetime.now().strftime('%Y-%m-%d %H:%M')


def meta_of(b):
    return BOOKMETA.get(str(b.get('bookId'))) or {}


def rating_of(b):
    v = meta_of(b).get('rating')
    return float(v) if isinstance(v, (int, float)) else None


def yseg(y):
    if not y or not str(y).isdigit():
        return '未知'
    v = int(y)
    if v <= 2012:
        return '≤2012'
    if v <= 2015:
        return '2013-2015'
    if v <= 2018:
        return '2016-2018'
    if v <= 2021:
        return '2019-2021'
    return '2022+'


def cseg(n):
    if n < 100:
        return '＜100人'
    if n < 1000:
        return '100-1k'
    if n < 10000:
        return '1k-1w'
    if n < 100000:
        return '1w-10w'
    return '≥10w'


YSEGS = ['≤2012', '2013-2015', '2016-2018', '2019-2021', '2022+', '未知']
CSEGS = ['＜100人', '100-1k', '1k-1w', '1w-10w', '≥10w']


def _laplace(sel, alls):
    """带 Laplace 平滑的下架概率"""
    return (sum(1 for x in alls if x['_sold']) + 0.5) / (len(alls) + 1) if alls else 0.0


# ---------- 价值评分 = v4 三层 + 互动证据（划线/想法）；v4/v3/v2 保留作对照 ----------
try:
    _VS5 = json.load(open('_valuescore5.json', encoding='utf-8'))
    VSCORE = _VS5.get('scores') or {}
    VSPARAM = _VS5.get('params') or {}
    VTYPE = _VS5.get('typestat') or {}
    VBASE = _VS5.get('typebase') or {}
    VEBASE = _VS5.get('etypebase') or {}
    VGLOBAL = (VSPARAM.get('global') or {})
except Exception:
    _VS5, VSCORE, VSPARAM, VTYPE, VBASE, VEBASE, VGLOBAL = {}, {}, {}, {}, {}, {}, {}

try:
    _VS4 = json.load(open('_valuescore4.json', encoding='utf-8'))
    V4SCORE = _VS4.get('scores') or {}
    VTYPE4 = _VS4.get('typestat') or {}
except Exception:
    V4SCORE, VTYPE4 = {}, {}

try:
    _VS3 = json.load(open('_valuescore3.json', encoding='utf-8'))
    V3SCORE = _VS3.get('scores') or {}
    VTYPE3 = _VS3.get('typestat') or {}
    VSPARAM3 = _VS3.get('params') or {}
except Exception:
    V3SCORE, VTYPE3, VSPARAM3 = {}, {}, {}

try:
    _VS2 = json.load(open('_valuescore2.json', encoding='utf-8'))
    VS2SCORE = _VS2.get('scores') or {}
    VSPARAM2 = _VS2.get('params') or {}
except Exception:
    VS2SCORE, VSPARAM2 = {}, {'FLOOR': 0.2, 'CEIL': 1.0, 'AT_MEAN': 0.8, 'K_REL': 0.7,
                              'N_HILL': 3.0801, 'r_mean': 0.0301, 'r_med': 0.0255,
                              'r_min': 0.003, 'r_max': 0.173, 'n': 0}

# 互动数据（旧 _reviewstats 仅作兜底；主口径已并入 _valuescore5）
try:
    REVIEWSTATS = json.load(open('_reviewstats.json', encoding='utf-8'))
except Exception:
    REVIEWSTATS = {}


def wilson_lcb(p, n, z=None):
    import math
    z = z or VSPARAM.get('Z_LCB', 1.645)
    if not n or n <= 0:
        return 0.0
    p = min(max(p, 0.0), 1.0)
    z2 = z * z
    den = 1.0 + z2 / n
    centre = p + z2 / (2.0 * n)
    half = z * math.sqrt(max(p * (1 - p) / n + z2 / (4.0 * n * n), 0.0))
    return max(0.0, min(1.0, (centre - half) / den))


def vs_of(b):
    return VSCORE.get(str(b.get('bookId'))) or {}


def split_type(c):
    """category 形如「经济理财-理财」→ (一级, 二级)"""
    c = (c or '').strip()
    if not c:
        return '未分类', '未分类'
    if '-' in c:
        a, b2 = c.split('-', 1)
        return a, c
    return c, c


for b in BOOKS:
    md = meta_of(b)
    b['_sold'] = b.get('soldout') == 1
    b['_rating'] = rating_of(b)
    b['_rcnt'] = int(md.get('ratingCount') or 0)
    b['_pub'] = (md.get('publishTime') or '')[:4]
    b['_pubseg'] = yseg(b['_pub'])
    b['_cseg'] = cseg(b['_rcnt'])
    b['_publisher'] = md.get('publisher') or ''
    _v = vs_of(b)
    b['_vscore'] = _v.get('score') or 0.0
    b['_conf'] = _v.get('conf') or 0.0
    b['_vlcb'] = _v.get('lcb') or 0.0
    b['_confstat'] = _v.get('conf_stat')
    b['_relr'] = _v.get('rel_r')
    b['_heat'] = _v.get('heat')
    b['_ratio'] = _v.get('ratio')
    b['_type1'] = _v.get('c1') or split_type(md.get('category'))[0]
    b['_type2'] = _v.get('c2') or split_type(md.get('category'))[1]
    b['_xr'] = _v.get('x_r')
    b['_xd'] = _v.get('x_d')
    b['_conf_r'] = _v.get('conf_r')
    b['_clsrel'] = _v.get('cls_rel')
    b['_clspct'] = _v.get('cls_pct')
    b['_rbar'] = _v.get('rbar')
    b['_rdbar'] = _v.get('rdbar')
    b['_fac_e'] = _v.get('fac_e')
    b['_conf_e'] = _v.get('conf_e')
    b['_xmark'] = _v.get('x_mark')
    b['_xidea'] = _v.get('x_idea')
    b['_dmark'] = _v.get('d_mark')
    b['_didea'] = _v.get('d_idea')
    b['_xe'] = _v.get('x_e')
    b['_nidea'] = _v.get('idea')
    b['_nmark'] = _v.get('mark')
    _v4 = V4SCORE.get(str(b.get('bookId'))) or {}
    b['_v4score'] = _v4.get('score')          # 可能为 None（该书在 v4 跑批之后才进书架）
    b['_has4'] = b['_v4score'] is not None
    _v3 = V3SCORE.get(str(b.get('bookId'))) or {}
    b['_v3score'] = _v3.get('score') or 0.0
    b['_v3conf'] = _v3.get('conf') or 0.0
    _v2 = VS2SCORE.get(str(b.get('bookId'))) or {}
    b['_v2score'] = _v2.get('score') or 0.0
    b['_v2conf'] = _v2.get('conf') or 0.0
    b['_vdelta'] = (round(b['_vscore'] - b['_v4score'], 1)
                    if (b['_vscore'] and b['_has4']) else 0.0)
    b['_d4'] = b['_vdelta']
    b['_d3'] = round(b['_vscore'] - b['_v3score'], 1) if b['_vscore'] else 0.0
    b['_d2'] = round(b['_vscore'] - b['_v2score'], 1) if b['_vscore'] else 0.0
    b['_neff'] = 0.0
    b['_readcnt'] = int(md.get('readingCount') or 0)
    _rs = REVIEWSTATS.get(str(b.get('bookId'))) or {}
    b['_idea'] = b.get('_nidea') if b.get('_nidea') is not None else _rs.get('idea')
    b['_mark'] = b.get('_nmark') if b.get('_nmark') is not None else None
    b['_review'] = int(md.get('reviewCount') or 0)
    b['_disc'] = round((b['_rating'] or 0) - (b['_vscore'] or 0), 1)

# 各信号的分档下架率（数据驱动，非拍脑袋）
P_CAT = {c: _laplace(None, [b for b in BOOKS if b['_cat'] == c]) for c in set(b['_cat'] for b in BOOKS)}
P_PUB = {k: _laplace(None, [b for b in BOOKS if b['_pubseg'] == k]) for k in YSEGS}
P_CNT = {k: _laplace(None, [b for b in BOOKS if b['_cseg'] == k]) for k in CSEGS}

# 下架风险分 = 0.5×权益类型 + 0.3×出版年 + 0.2×冷门度（各为该档实测下架率）
for b in BOOKS:
    b['_risk'] = round(100 * (0.5 * P_CAT[b['_cat']] + 0.3 * P_PUB[b['_pubseg']] + 0.2 * P_CNT[b['_cseg']]), 1)


# ---------- 阅读进度（/web/shelf/sync 的 bookProgress + /web/book/getProgress 的 finishTime）----------
try:
    PROG2 = json.load(open('_progress2.json', encoding='utf-8'))
except Exception:
    PROG2 = {}
PROG_SHELF = {str(x.get('bookId')): x for x in (SHELF.get('bookProgress') or [])}


def _pick(*vals):
    for v in vals:
        if v is not None:
            return v
    return None


def progress_rec(bid):
    bid = str(bid)
    a = PROG_SHELF.get(bid) or {}
    c = PROG2.get(bid) or {}
    return {
        'progress': _pick(c.get('progress'), a.get('progress')),
        'rtime': _pick(c.get('readingTime'), a.get('readingTime')),
        'finish': c.get('finishTime'),
        'start': c.get('startReadingTime'),
        'recent': _pick(c.get('updateTime'), a.get('updateTime')),
        'chapterIdx': _pick(c.get('chapterIdx'), a.get('chapterIdx')),
    }


def pstat_of(p):
    if p is None:
        return '无记录'
    if p >= 100:
        return '已读完'
    if p >= 80:
        return '快读完'
    if p >= 30:
        return '读了一半'
    if p >= 1:
        return '刚开始'
    return '未读'


PSTAT_ORDER = ['已读完', '快读完', '读了一半', '刚开始', '未读', '无记录']
PSTAT_COLOR = {'已读完': 'C6EFCE', '快读完': 'FFEB9C', '读了一半': 'FCE4D6',
               '刚开始': 'DDEBF7', '未读': 'FFFFFF', '无记录': 'F2F2F2'}


def ts(v):
    """时间戳 -> 'YYYY-MM-DD'；0/None -> ''"""
    try:
        v = int(v)
        if v <= 0:
            return ''
        return datetime.fromtimestamp(v).strftime('%Y-%m-%d')
    except Exception:
        return ''


def hours(sec):
    try:
        sec = int(sec)
    except Exception:
        return None
    if sec <= 0:
        return 0.0
    return round(sec / 3600.0, 1)


for b in BOOKS:
    _r = progress_rec(b['bookId'])
    b['_prog'] = int(_r['progress']) if isinstance(_r['progress'], (int, float)) else None
    b['_rtime'] = _r['rtime'] if isinstance(_r['rtime'], (int, float)) else 0
    b['_rhours'] = hours(b['_rtime'])
    b['_finish'] = ts(_r['finish'])
    b['_start'] = ts(_r['start'])
    b['_recent'] = ts(_r['recent'])
    b['_pstat'] = pstat_of(b['_prog'])

READ_DONE = [b for b in BOOKS if b['_pstat'] == '已读完']
NEED_DONE = [b for b in NEED if b['_pstat'] == '已读完']
NEED_READING = [b for b in NEED if b['_prog'] is not None and 0 < b['_prog'] < 100]
NEED_FAST = [b for b in NEED if b['_pstat'] == '快读完']
TOTAL_HOURS = round(sum(b['_rtime'] for b in BOOKS) / 3600.0, 1)


# ==================== 书架 + 书单 合并去重（「热书排名」「分类基准与类型评分」共用） ====================
def _hill(x, k, n):
    if not x or x <= 0:
        return 0.0
    xn = x ** n
    return xn / (k ** n + xn)


def _med(a):
    a = sorted([x for x in a if x is not None])
    if not a:
        return None
    n = len(a)
    return a[n // 2] if n % 2 else (a[n // 2 - 1] + a[n // 2]) / 2.0


_VP = VSPARAM or {}
_FL = _VP.get('FLOOR', 0.20)
_CL = _VP.get('CEIL', 1.00)
_AT = _VP.get('AT_MEAN', 0.80)
_KR = _VP.get('K_REL', 0.70)
_NH = _VP.get('N_HILL', 3.0801)
_WR = _VP.get('W_R', 0.5)
_RFL = _VP.get('REL_FLOOR', 0.65)
_RCE = _VP.get('REL_CEIL', 1.12)
_HB = _VP.get('HEAT_BASE', 0.95)
_HR = _VP.get('HEAT_RANGE', 0.10)
_KE = _VP.get('K_E', 0.50)
_WE = _VP.get('W_E', 0.40)
_WMK = _VP.get('W_MARK', 0.55)
_WID = _VP.get('W_IDEA', 0.45)
_KP = _VP.get('K_PSEUDO', 200.0)
_EF = _VP.get('E_FLOOR', 0.88)
_EC = _VP.get('E_CEIL', 1.06)
_XCL = (0.10, 4.0)
# 封顶曲线参数与 score_value.py 同源（从 json 的 params 读，保证两处口径一致）
_CAPRAW = _VP.get('CAP_RAW', 91.0)
_CAPG = _VP.get('CAP_GAMMA', 0.5)
_EG = {'mark': _med([v.get('mark') for v in (VEBASE or {}).values()]),
       'idea': _med([v.get('idea') for v in (VEBASE or {}).values()])}


def score_like_shelf(rec, c1, c2):
    """给「不在书架」的书用完全相同的公式打分（类内基准取书架算出的 typebase）。

    缺阅读人数时参与率/受众面两层退回 1.0（不奖不罚），只由 conf_stat × 互动决定。
    """
    raw = float(rec.get('rating') or 0)
    rc = float(rec.get('ratingCount') or 0)
    rd = float(rec.get('readingCount') or 0)
    idea = rec.get('idea')
    mark = rec.get('markBest') if rec.get('markBest') else None
    out = {'score': 0.0, 'conf': 0.0, 'fac_e': None, 'x_e': None, 'ratio': None,
           'rel_eff': None, 'heat': None, 'conf_stat': None, 'hasRd': rd > 0}
    if raw <= 0 or rc <= 0:
        return out
    p = raw / 100.0
    cs = min(wilson_lcb(p, rc) / p, 1.0) if p > 0 else 0.0
    base = ((VBASE.get('c2') or {}).get(c2) or (VBASE.get('c1') or {}).get(c1) or {})
    mean_r = base.get('mean_r')
    geo_rd = base.get('geo_rd')
    rel, heat = 1.0, 1.0
    if rd > 0 and mean_r and geo_rd:
        ratio = rc / rd
        rel = (_FL + (_CL - _FL) * _hill(ratio / mean_r, _KR, _NH)) / _AT
        heat = _HB + _HR * _hill(rd / geo_rd, _KR, _NH)
        out['ratio'] = round(ratio, 6)

    eb = (VEBASE or {}).get(c1) or _EG

    def smooth(cnt, prior):
        if cnt is None or rd <= 0 or not prior:
            return None
        return (float(cnt) + _KP * prior) / (rd + _KP)

    dmk = smooth(mark, eb.get('mark') or _EG.get('mark'))
    did = smooth(idea, eb.get('idea') or _EG.get('idea'))
    xmk = (min(max(dmk / eb['mark'], _XCL[0]), _XCL[1])
           if (dmk is not None and eb.get('mark')) else None)
    xid = (min(max(did / eb['idea'], _XCL[0]), _XCL[1])
           if (did is not None and eb.get('idea')) else None)
    if xmk is not None and xid is not None:
        xe = (xmk ** _WMK) * (xid ** _WID)
    else:
        xe = xid if xid is not None else xmk
    if xe is not None:
        ce = _FL + (_CL - _FL) * _hill(xe, _KE, _NH)
        fac = min(max((ce / _AT) ** _WE, _EF), _EC)
        out['x_e'] = round(xe, 4)
        out['fac_e'] = round(fac, 4)
    else:
        fac = 1.0
    rel_eff = min(max(rel ** _WR, _RFL), _RCE)
    cap = min(1.0, (raw / _CAPRAW) ** _CAPG) if _CAPRAW > 0 else 1.0
    conf = min(cs * rel_eff * fac * heat, cap)
    out.update({'score': round(raw * conf, 1), 'conf': round(conf, 4),
                'rel_eff': round(rel_eff, 3), 'heat': round(heat, 3),
                'conf_stat': round(cs, 3)})
    return out


def build_union():
    """书架全部书 + 公开书单里未加书架的书，按 bookId 合并去重。

    返回 list[dict]（同一结构，便于两张 sheet 一起排序）：
      bid / inShelf / title / author / type1 / type2 / price / rating / rc / rd /
      idea / mark / vscore / conf / fac_e / cat / sold / prog / pstat / risk / bl(所属书单)
    """
    try:
        BL = json.load(open('_booklists.json', encoding='utf-8'))
    except Exception:
        BL = {}
    try:
        DET = json.load(open('_bookdetail_extra.json', encoding='utf-8'))
    except Exception:
        DET = {}
    try:
        EXM = json.load(open('_bookmeta_v5_extra.json', encoding='utf-8'))
    except Exception:
        EXM = {}

    bln = {}
    for rec in BL.values():
        nm = rec.get('name') or '-'
        for x in rec.get('books') or []:
            bid = str(x['bookId'])
            lst = bln.setdefault(bid, [])
            if nm not in lst:
                lst.append(nm)

    rows, shelf_ids = [], set()
    for b in BOOKS:
        bid = str(b.get('bookId'))
        shelf_ids.add(bid)
        rows.append({
            'bid': bid, 'inShelf': True, 'b': b,
            'title': b.get('title') or '', 'author': b.get('author') or '',
            'type1': b['_type1'], 'type2': b['_type2'],
            'price': b['_price'], 'rating': b['_rating'], 'rc': b['_rcnt'],
            'rd': b['_readcnt'], 'idea': b.get('_idea'), 'mark': b.get('_mark'),
            'ratio': b.get('_ratio'),
            'vscore': b['_vscore'] or 0.0, 'conf': b.get('_conf') or 0.0,
            'fac_e': b.get('_fac_e'), 'cat': b['_cat'], 'sold': b['_sold'],
            'prog': b['_prog'], 'pstat': b['_pstat'], 'risk': b['_risk'],
            'bl': bln.get(bid, []),
        })

    for bid, names in bln.items():
        if bid in shelf_ids:
            continue
        m = dict(DET.get(bid) or {})
        m.update({k: v for k, v in (EXM.get(bid) or {}).items() if v is not None})
        catg = m.get('category') or ''
        c1, c2 = split_type(catg)
        sc = score_like_shelf(m, c1, c2)
        price = m.get('price')
        if not price:
            price = round((m.get('centPrice') or 0) / 100.0, 2)
        rows.append({
            'bid': bid, 'inShelf': False, 'b': None,
            'title': m.get('title') or '', 'author': m.get('author') or '',
            'type1': c1, 'type2': c2, 'price': round(float(price or 0), 2),
            'rating': float(m.get('rating') or 0), 'rc': int(m.get('ratingCount') or 0),
            'rd': int(m.get('readingCount') or 0),
            'idea': m.get('idea'), 'mark': m.get('markBest'),
            'ratio': sc.get('ratio'),
            'vscore': sc['score'], 'conf': sc['conf'], 'fac_e': sc['fac_e'],
            'cat': '未加书架', 'sold': 1 if m.get('soldout') in (1, True) else 0,
            'prog': None, 'pstat': '未加书架', 'risk': None, 'bl': names,
        })
    return rows


UNION = build_union()


def union_decorate(rows):
    """补两个派生量：热度分（跨书架+书单可比）与类内相对分。"""
    import math

    def lg(x):
        return math.log10(1 + max(0, x or 0))

    dims = [('rd', 0.40), ('rc', 0.30), ('mark', 0.20), ('idea', 0.10)]
    vals = {}
    for key, _w in dims:
        vals[key] = sorted(lg(r.get(key) or 0) for r in rows if (r.get(key) or 0) > 0)

    def pctile(sorted_vals, v):
        if not sorted_vals:
            return None
        lo = sum(1 for x in sorted_vals if x < v)
        return 100.0 * lo / len(sorted_vals)

    for r in rows:
        acc = wsum = 0.0
        used = 0
        for key, w in dims:
            if (r.get(key) or 0) <= 0:
                continue
            p = pctile(vals[key], lg(r.get(key)))
            if p is None:
                continue
            acc += w * p
            wsum += w
            used += 1
        r['heat'] = round(acc / wsum, 1) if wsum else 0.0
        r['heatDim'] = used

    by_t1 = {}
    for r in rows:
        if r['vscore']:
            by_t1.setdefault(r['type1'], []).append(r)
    for t1, sub in by_t1.items():
        scs = sorted(x['vscore'] for x in sub)
        mean = sum(scs) / len(scs)
        for r in sub:
            r['clsN'] = len(sub)
            r['clsMean'] = round(mean, 1)
            r['clsRel'] = round(r['vscore'] - mean, 1)
            r['clsPct'] = round(100.0 * sum(1 for x in scs if x < r['vscore']) / len(scs), 1)
    return rows


union_decorate(UNION)
UNION_N = len(UNION)
UNION_SHELF = sum(1 for r in UNION if r['inShelf'])
UNION_OUT = UNION_N - UNION_SHELF
UNION_BL_ONLY = sum(1 for r in UNION if not r['inShelf'] and r['bl'])
UNION_RATED = sum(1 for r in UNION if r['vscore'])


# ==================== 一次性最大量购书（每本只买一次） ====================
def knap_lex(items, budget, secondary='min_spend'):
    """0/1 背包精确解，每本书最多选一次（bookId 唯一，不重复购）。

    第一目标：本数最多（"最大量"的硬约束）
    第二目标（本数相同时才比较）：'min_spend' 花费最少 / 'max_spend' 花满预算
                                  'risk' 下架风险合计最大 / 'rating' 评分合计最大
    做法：按"恰好花掉 w 分"做 DP，状态 =(本数, 第二目标值)，字典序取优。
    返回 (选中的书列表, 总花费(元), 第二目标值)。
    """
    B = int(round(budget * 100))
    cand = [b for b in items if 0 < int(round(b['_price'] * 100)) <= B]
    n = len(cand)
    if not n or B <= 0:
        return [], 0.0, 0.0

    cost = [int(round(b['_price'] * 100)) for b in cand]
    if secondary == 'risk':
        sec = [int(round(b['_risk'] * 10)) for b in cand]
    elif secondary == 'rating':
        # 用价值评分（置信下界），避免"人少高分"的书靠虚高评分挤进方案
        sec = [int(round((b.get('_vscore') or 0) * 10)) for b in cand]
    else:
        sec = [0] * n

    NEG = (-1, -1)
    dp = [NEG] * (B + 1)
    dp[0] = (0, 0)
    choice = [bytearray(B + 1) for _ in range(n)]

    for i in range(n):
        ci, si = cost[i], sec[i]
        for w in range(B, ci - 1, -1):
            prev = dp[w - ci]
            if prev[0] < 0:
                continue
            nc = prev[0] + 1
            if secondary in ('risk', 'rating'):
                ns = prev[1] + si
            elif secondary == 'max_spend':
                ns = w
            else:
                ns = -w
            cur = dp[w]
            if nc > cur[0] or (nc == cur[0] and ns > cur[1]):
                dp[w] = (nc, ns)
                choice[i][w] = 1

    best_w, best = 0, NEG
    for w in range(B + 1):
        s = dp[w]
        if s[0] < 0:
            continue
        if s[0] > best[0] or (s[0] == best[0] and s[1] > best[1]):
            best, best_w = s, w

    picked, w = [], best_w
    for i in range(n - 1, -1, -1):
        if choice[i][w]:
            picked.append(cand[i])
            w -= cost[i]
    return picked, round(best_w / 100.0, 2), best[1]


def plan_stat(pk):
    sp = round(sum(b['_price'] for b in pk), 2)
    return dict(n=len(pk), spent=sp, left=round(COIN - sp, 2),
                risk=round(sum(b['_risk'] for b in pk) / len(pk), 1) if pk else 0,
                rating=round(sum((b['_rating'] or 0) for b in pk) / len(pk), 1) if pk else 0,
                wrating=round(sum((b['_vscore'] or 0) for b in pk) / len(pk), 1) if pk else 0)


def make_group(pool):
    """对给定候选池算 4 种"本数最多"的组合 + 上限证明数据"""
    g = {}
    for key, sec in (('min', 'min_spend'), ('rate', 'rating'),
                     ('full', 'max_spend'), ('risk', 'risk')):
        pk, sp, _v = knap_lex(pool, COIN, sec)
        if key == 'rate':
            pk = sorted(pk, key=lambda b: (-(b['_rating'] or 0), b['_price']))
        elif key == 'risk':
            pk = sorted(pk, key=lambda b: (-b['_risk'], b['_price']))
        else:
            pk = sorted(pk, key=lambda b: (-(b['_rating'] or 0), b['_price']))
        g[key] = {'plan': pk, 'stat': plan_stat(pk), 'spend': sp}

    ordered = sorted(pool, key=lambda b: (b['_price'], b.get('title') or ''))
    mx = max(v['stat']['n'] for v in g.values())
    s = 0
    n = 0
    nxt_p, nxt_t = None, ''
    for b in ordered:
        c = int(round(b['_price'] * 100))
        if s + c <= int(round(COIN * 100)):
            s += c
            n += 1
        elif nxt_p is None:
            nxt_p, nxt_t = b['_price'], b.get('title') or ''
    if nxt_p is None and OUT_POOL:
        # 池内书全被买下（或池为空）：拿"单本超预算"里最便宜的一本当兜底，
        # 保证 next_price 永远是个数，"下一本要多少钱"这句才讲得通。
        nxt_p, nxt_t = OUT_POOL[0]['_price'], OUT_POOL[0].get('title') or ''
    g['n'] = mx
    g['greedy_n'] = n
    g['cheap_sum'] = round(s / 100.0, 2)
    g['next_price'] = nxt_p
    g['next_title'] = nxt_t
    g['next_sum'] = round(s / 100.0 + nxt_p, 2) if nxt_p else None
    g['empty_pool'] = not pool
    assert mx == n, '背包解与贪心上限不一致：%d vs %d' % (mx, n)
    return g


# 候选池一：待付费（未拥有）+ 未下架 + 单本标价 ≤ 余额
COIN_POOL = [b for b in BOOKS
             if b['_cat'] in ('付费会员卡可读', '体验卡可读')
             and not b['_sold'] and 0 < b['_price'] <= COIN + 1e-9]
COIN_POOL.sort(key=lambda b: (b['_price'], b.get('title') or ''))

# 候选池二：在上面基础上排除"我已经读完（进度 100%）"的书
COIN_POOL_X = [b for b in COIN_POOL if b['_pstat'] != '已读完']

# 候选池外的书：单本标价就超过余额，背包永远不会选它们。
# 保留下来只为回答"再买下一本要多少钱"——余额极低、池子里一本书都没有时，
# 全靠它给 next_price 兜底，否则 next_price 为 None 会让后续文案格式化炸掉。
OUT_POOL = [b for b in BOOKS
            if b['_cat'] in ('付费会员卡可读', '体验卡可读')
            and not b['_sold'] and b['_price'] > COIN + 1e-9]
OUT_POOL.sort(key=lambda b: (b['_price'], b.get('title') or ''))

# ===== 零钱备选池 =====
# 余额不足时也要能给出「买得起什么」，否则报告只剩一句"买不了"。
# 口径：未下架、有标价、有评分、标价 ≤ LOW_CAP（默认 2 元），按价值评分降序。
# LOW_CAP 可用环境变量 WEREAD_LOW_PRICE 覆盖（例如 1 / 3 / 5）。
try:
    LOW_CAP = float(os.environ.get('WEREAD_LOW_PRICE') or 2.0)
except ValueError:
    LOW_CAP = 2.0
if LOW_CAP <= 0:
    LOW_CAP = 2.0

LOW_POOL = [b for b in BOOKS
            if b['_cat'] in ('付费会员卡可读', '体验卡可读')
            and not b['_sold'] and 0 < b['_price'] <= LOW_CAP + 1e-9 and b['_vscore']]
LOW_POOL.sort(key=lambda b: (-(b['_vscore'] or 0), b['_price'], b.get('title') or ''))
# 当前余额能买得起的子集（用于"其中 X 本现在就能买"）
LOW_AFFORD = [b for b in LOW_POOL if b['_price'] <= COIN + 1e-9]
LOW_CHEAPEST = min(LOW_POOL, key=lambda b: b['_price']) if LOW_POOL else None

# 全库采样清单（技能自带 data/lowprice_seed.json，由 fetch_lowprice.py 生成）。
# 用途：用户书架/书单里一本 ≤N 元的书都没有时，这份清单兜底，让报告仍能给出可买的书。
# 注意它与你书架无关——是"全库哪些书这么便宜且评分够高"，所以要单独标注来源。
LOW_SEED_PATH = os.path.normpath(
    os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'data', 'lowprice_seed.json'))


def _load_low_seed():
    try:
        with open(LOW_SEED_PATH, encoding='utf-8') as f:
            d = json.load(f)
    except Exception:
        return [], {}
    meta = d.get('meta') or {}
    books = d.get('books') or []
    # 种子清单的 cap 与当前 LOW_CAP 不一致时，按当前阈值再筛一遍（比如种子是 ≤5 元、现在想看 ≤2 元）
    seed_cap = meta.get('cap')
    if isinstance(seed_cap, (int, float)) and abs(seed_cap - LOW_CAP) > 1e-9:
        books = [b for b in books if float(b.get('price') or 0) <= LOW_CAP + 1e-9]
        meta = dict(meta, filteredFrom=seed_cap)
    books = [b for b in books if float(b.get('price') or 0) > 0]
    # 排序：先「你没加书架的」→「在书架但没买断」→「已拥有」，同组内按评分降序。
    # 不这么排的话，高分的往往正好是书架里已有的，所谓"备选"就成了重复推荐。
    def _own_rank(b):
        o = BY_ID.get(b.get('bookId'))
        if not o:
            return 0
        return 2 if o['_cat'] == '永久可读' else 1
    books.sort(key=lambda b: (_own_rank(b), -(b.get('rating') or 0),
                              float(b.get('price') or 0), b.get('title') or ''))
    return books, meta


LOW_SEED, LOW_SEED_META = _load_low_seed()
# 主表来源：书架里有就用书架的（能直接买），没有就回落到全库采样清单
LOW_MAIN = LOW_POOL if LOW_POOL else LOW_SEED
LOW_MAIN_SRC = 'shelf' if LOW_POOL else ('seed' if LOW_SEED else 'none')


APPROX_SRC = ('inferred', 'manual')


def type_disp(b):
    """分类展示：平台未给分类、靠推断或人工指定的类目加 ≈ 前缀。"""
    t = (b.get('_type1') or '').strip() or '未分类'
    return ('≈' + t) if (b.get('_type_src') in APPROX_SRC and t != '未分类') else t


SEED_INFERRED = sum(1 for b in LOW_SEED if b.get('_type_src') in APPROX_SRC)

# ==================== 零钱模式 ====================
# 触发条件：余额 ≤ LOW_TRIGGER（默认 2 元，环境变量 WEREAD_LOW_TRIGGER 可调）。
# 为什么需要这个模式：余额只有一两元时，基于书架的四套方案往往一本都排不出来
# （书架里未必有 ≤2 元的书），报告就只剩一句"买不了"。零钱模式改用**全库 2 元清单**
# 当候选池，直接给出「这点钱具体买哪几本」的 3 个方案 + 攒够 20 书币后的参考。
try:
    LOW_TRIGGER = float(os.environ.get('WEREAD_LOW_TRIGGER') or 2.0)
except ValueError:
    LOW_TRIGGER = 2.0
if LOW_TRIGGER < 0:
    LOW_TRIGGER = 2.0
LOW_MODE = COIN <= LOW_TRIGGER + 1e-9


import lowplan
import report_low


LOW_POOL_BAG = lowplan.seed_as_pool(LOW_SEED)

# 单个方案最多买几本。必须设上限：清单里有大量 0.01 元的「雪球漫画投资课/私募专刊」小册子，
# 纯按「本数最多」做背包会刷出三四十本 —— 数学最优但业务上没意义，用户要的是"买哪几本"。
try:
    LOW_MAX_BOOKS = max(1, int(os.environ.get('WEREAD_LOW_MAXBOOKS') or 3))
except ValueError:
    LOW_MAX_BOOKS = 3

LOW_PLANS, LOW_LATER = (lowplan.make_plans(LOW_POOL_BAG, COIN, LOW_MAX_BOOKS)
                        if LOW_MODE else ([], {}))

# 两张零钱 sheet 的排版在 report_low，这里把数据与样式注入过去。
report_low.configure(
    COIN=COIN, LOW_CAP=LOW_CAP, LOW_TRIGGER=LOW_TRIGGER, LOW_MAX_BOOKS=LOW_MAX_BOOKS,
    LOW_POOL=LOW_POOL, LOW_SEED=LOW_SEED, LOW_POOL_BAG=LOW_POOL_BAG,
    LOW_AFFORD=LOW_AFFORD, LOW_CHEAPEST=LOW_CHEAPEST, LOW_SEED_META=LOW_SEED_META,
    LOW_PLANS=LOW_PLANS, LOW_LATER=LOW_LATER, SEED_INFERRED=SEED_INFERRED,
    BY_ID=BY_ID, type_disp=type_disp,
    BORDER=BORDER, HEAD_FILL=HEAD_FILL, HEAD_FONT=HEAD_FONT, TITLE_FONT=TITLE_FONT,
    CENTER=CENTER, MONEY=MONEY)

POOL_IDS = [b['bookId'] for b in COIN_POOL]
DUP_ID = len(POOL_IDS) - len(set(POOL_IDS))
_tc = {}
for b in COIN_POOL:
    _tc[b.get('title') or ''] = _tc.get(b.get('title') or '', 0) + 1
DUP_TITLE = {k: v for k, v in _tc.items() if v > 1}

G_ALL = make_group(COIN_POOL)       # 组一：不排除已读完
G_X = make_group(COIN_POOL_X)       # 组二：排除已读完（推荐）
MAX_N = G_X['n']
MAIN = G_X['rate']                  # 推荐方案：排除已读完 + 本数最多 + 评分最优
MAIN_PLAN, MAIN_STAT = MAIN['plan'], MAIN['stat']

# 门檻扫描（说明"一加门槛本数就掉"）
PLAN_A_SCAN = []
for th in (0, 10, 15, 20, 25, 30, 35, 40):
    g = [b for b in COIN_POOL_X if b['_risk'] >= th]
    pk, sp, _v = knap_lex(g, COIN, 'risk')
    st = plan_stat(pk)
    PLAN_A_SCAN.append((th, len(g), st['n'], st['spent'], st['left'], st['risk'], st['rating']))

PLAN_B_SCAN = []
for th in (0, 75, 78, 80, 82, 84, 85, 87, 88, 90):
    g = [b for b in COIN_POOL_X if (b['_rating'] or 0) >= th]
    pk, sp, _v = knap_lex(g, COIN, 'rating')
    st = plan_stat(pk)
    PLAN_B_SCAN.append((th, len(g), st['n'], st['spent'], st['left'], st['rating'], st['risk']))

# ==================== 四个购买方案（合并进「积分购买方案」一个 sheet） ====================
PLAN_SHEET = COIN_SHEET          # 合并后：方案与清单在同一个 sheet


def knap_max(items, budget, wfun):
    """0/1 背包：预算内最大化 Σ wfun(b)（不限制本数）。返回 (选中, 花费)。"""
    B = int(round(budget * 100))
    cand = [b for b in items if 0 < int(round(b['_price'] * 100)) <= B]
    if not cand or B <= 0:
        return [], 0.0
    cost = [int(round(b['_price'] * 100)) for b in cand]
    wt = [max(0.0, float(wfun(b))) for b in cand]
    NEG = -1e18
    dp = [NEG] * (B + 1)
    dp[0] = 0.0
    choice = [bytearray(B + 1) for _ in range(len(cand))]
    for i, (ci, wi) in enumerate(zip(cost, wt)):
        for w in range(B, ci - 1, -1):
            if dp[w - ci] <= NEG / 2:
                continue
            v = dp[w - ci] + wi
            if v > dp[w] + 1e-12:
                dp[w] = v
                choice[i][w] = 1
    best_w = max(range(B + 1), key=lambda w: dp[w])
    picked, w = [], best_w
    for i in range(len(cand) - 1, -1, -1):
        if choice[i][w]:
            picked.append(cand[i])
            w -= cost[i]
    return picked, round(best_w / 100.0, 2)


def greedy_ratio(items, budget):
    """按 价值分 ÷ 标价（每元价值）从高到低，钱够就买 —— 性价比优先的贪心解。"""
    B = int(round(budget * 100))
    left = B
    picked = []
    ranked = sorted([b for b in items if 0 < int(round(b['_price'] * 100)) <= B],
                    key=lambda b: (-((b['_vscore'] or 0) / b['_price']), b['_price']))
    for b in ranked:
        c = int(round(b['_price'] * 100))
        if c <= left:
            picked.append(b)
            left -= c
    return picked, round((B - left) / 100.0, 2)


def plan4(pk):
    """方案统计：本数 / 花费 / 剩余 / 总价值分 / 平均价值分 / 每元价值 / 平均下架风险"""
    sp = round(sum(b['_price'] for b in pk), 2)
    tv = round(sum((b['_vscore'] or 0) for b in pk), 1)
    return dict(n=len(pk), spent=sp, left=round(COIN - sp, 2), tv=tv,
                av=round(tv / len(pk), 1) if pk else 0.0,
                per=round(tv / sp, 1) if sp else 0.0,
                risk=round(sum(b['_risk'] for b in pk) / len(pk), 1) if pk else 0.0)


def cap_note():
    """「余额实际能买 N 本」这句的说明。

    余额极低时 next_price 可能为 None（池子空了 / 连最便宜的都买不起），
    绝不能把 None 塞进 %.2f，否则整个 Excel 生成会 TypeError 中断。
    """
    left = round(COIN - G_X['cheap_sum'], 2)
    nx = G_X.get('next_price')
    if PS_CNT['n'] <= 0 and not nx:
        return '余额 %.2f 元，可买池为空，当前一本书都买不了' % COIN
    if PS_CNT['n'] <= 0:
        return ('余额 %.2f 元低于待付费书里最便宜的一本《%s》%.2f 元，当前一本都买不了；'
                '完成本周阅读任务拿到 20 书币后可以再买'
                % (COIN, G_X.get('next_title') or '（未知书名）', nx))
    if not nx:
        return ('单本标价 ≤ %.2f 元的 %d 本已全部买下，余额还剩 %.2f 元，可买池里没有更便宜的书了'
                % (COIN, PS_CNT['n'], left))
    return ('0/1 背包上限（每本只买一次）：第 %d 本最便宜也要 %.2f 元 > 剩余 %.2f 元'
            % (PS_CNT['n'] + 1, nx, left))


def cnt_plan_note():
    """方案②「本数最多」逐本表的抬头备注（空方案时不能吹"本数上限"）。"""
    if PS_CNT['n']:
        return ('② 本数最多：%d 本 / 花 %.2f 元 / 剩 %.2f 元 / 平均价值分 %.1f'
                % (PS_CNT['n'], PS_CNT['spent'], PS_CNT['left'], PS_CNT['av']))
    return '② 本数最多：0 本 / 余额 %.2f 元不足任何一本 / 剩 %.2f 元' % (COIN, PS_CNT['left'])


def cnt_plan_foot():
    """方案②逐本表脚注（同上，按是否有书切换口径）。"""
    nx = G_X.get('next_price')
    if PS_CNT['n'] and nx:
        return ('※ %d 本已是本数上限（第 %d 便宜的也要 %.2f 元 > 剩余 %.2f 元）。'
                '缺点：为了凑本数会掺进几本低分书，每元价值只有 %.1f 分/元。'
                % (PS_CNT['n'], PS_CNT['n'] + 1, nx, PS_CNT['left'], PS_CNT['per']))
    if PS_CNT['n']:
        return ('※ 可买池内 %d 本已全部买下（余额还剩 %.2f 元，没有更便宜的书）。'
                % (PS_CNT['n'], PS_CNT['left']))
    nx_txt = ('待付费书里最便宜的一本《%s》要 %.2f 元' % (G_X.get('next_title') or '（未知书名）', nx)) if nx else '可买池为空'
    return ('※ 余额 %.2f 元买不了任何一本（%s）。本方案为空是余额约束的结果，不是没有候选书；'
            '攒够书币后重跑本脚本即可。' % (COIN, nx_txt))


# ① 性价比最高：价值分 ÷ 标价 降序装满预算
# 只拿「有评分」的书做性价比方案：0 分的书每元价值也是 0，塞进去只是占预算
RATIO_PLAN, RATIO_SPENT = greedy_ratio([b for b in COIN_POOL_X if b['_vscore']], COIN)
RATIO_PLAN = sorted(RATIO_PLAN, key=lambda b: (-((b['_vscore'] or 0) / b['_price']), b['_price']))
PS_RATIO = plan4(RATIO_PLAN)

# ② 本数最多：0/1 背包本数优先 + 同本数价值分最高（MAIN_PLAN，已有）
MAIN_PLAN = sorted(MAIN_PLAN, key=lambda b: (-(b['_vscore'] or 0), b['_price']))
PS_CNT = plan4(MAIN_PLAN)


# ③ 绝对价值最高：一次只买 1 本，把积分尽量花在一本书上
# 分类口径与「分类基准与类型评分」第一节完全一致：
#   取 typestat.c1 中「有评分 ≥ 10 本」且非「未分类」的一级分类（该节共 15 类，按类均分排名）
TYPE1_STD = [k for k, v in (VTYPE.get('c1') or {}).items()
             if (v.get('nRated') or 0) >= 10 and k != '未分类']
TYPE1_STD.sort(key=lambda k: -(VTYPE['c1'][k].get('meanScore') or 0))
TYPE1_RANK = {k: i + 1 for i, k in enumerate(TYPE1_STD)}
TYPE1_MEAN = {k: (VTYPE['c1'][k].get('meanScore') or 0.0) for k in TYPE1_STD}


def _pctile_rank(vals, v):
    """池内百分位（0~100）：把参与率、阅读人数换成与评分同量纲的分数。"""
    if not vals:
        return 0.0
    return 100.0 * sum(1 for x in vals if x < v) / len(vals)


def power_score(pool):
    """给池内每本书算「综合价值 V」（价格不进 V）。

    V = 0.45×价值分 + 0.25×平台原评分
        + 0.15×参与率池内百分位 + 0.15×阅读人数池内百分位
    """
    pool = [b for b in pool if b['_vscore']]
    pr = [b['_rating'] or 0 for b in pool]
    pp = [b['_ratio'] or 0 for b in pool]
    pd = [b['_readcnt'] or 0 for b in pool]
    for b in pool:
        v = (0.45 * b['_vscore'] + 0.25 * (b['_rating'] or 0)
             + 0.15 * _pctile_rank(pp, b['_ratio'] or 0)
             + 0.15 * _pctile_rank(pd, b['_readcnt'] or 0))
        b['_V'] = round(v, 1)
        b['_pw'] = v / b['_price'] if b['_price'] else 0.0
    return pool


def power_top(pool, topn=2, types=None):
    """每个一级分类内按「综合价值 V」取前 topn 本 → {一级类: [候选书…]}

    · types = 只在这些一级分类里取（默认用 TYPE1_STD，与「分类基准与类型评分」第一节同一口径）
    · 单本购买不追求性价比，所以按 V 排序（不是 V÷价格）
    """
    groups = {}
    for b in power_score(pool):
        if types is not None and b['_type1'] not in types:
            continue
        groups.setdefault(b['_type1'], []).append(b)
    return {t1: sorted(rows, key=lambda b: (-b['_V'], b['_price']))[:topn]
            for t1, rows in groups.items()}


def single_pick(cands, budget):
    """一次只买一本：在候选里挑「剩余最少 + 价值最高」的那一本。

    · 候选 = 每个一级大类综合价值 V 前 3，且单本标价 ≤ 余额
    · 排序 = 单本绝对价值产出 = 综合价值 V × 标价
      只买一本时没有「数量」可摊薄，两个目标只能合并：
      标价越高 → 剩余越少（把积分花光）；V 越高 → 这本书本身越值。乘积越大越符合"绝对价值最高"。
    返回 (选中的那一本, 全部候选（已按上述规则降序）)
    """
    rows = []
    for t1 in sorted(cands):
        for b in cands[t1]:
            if not b['_price'] or b['_price'] > budget + 1e-9:
                continue
            b['_sv'] = round(b['_V'] * b['_price'], 1)
            b['_fill'] = b['_price'] / budget
            rows.append(b)
    rows.sort(key=lambda b: (-b['_sv'], b['_price']))
    return (rows[0] if rows else None), rows


POW_CANDS = power_top(COIN_POOL_X, 2, TYPE1_STD)
VAL_PICK, VAL_CANDS = single_pick(POW_CANDS, COIN)
VAL_PLAN = [VAL_PICK] if VAL_PICK else []
PS_VAL = plan4(VAL_PLAN)
PS_VAL['nTypePool'] = len(POW_CANDS)                       # 实际有候选的一级分类数
PS_VAL['nTypeStd'] = len(TYPE1_STD)                        # 基准表口径的一级分类总数
PS_VAL['nTypeCand'] = len({b['_type1'] for b in VAL_CANDS})  # 候选实际覆盖的大类数
PS_VAL['nCand'] = len(VAL_CANDS)                           # 候选本数（V 前 3 且 ≤ 余额）
PS_VAL['nDrop'] = sum(len(v) for v in POW_CANDS.values()) - len(VAL_CANDS)  # 因超预算被剔除的候选
PS_VAL['V'] = VAL_PICK['_V'] if VAL_PICK else 0.0
PS_VAL['sv'] = VAL_PICK['_sv'] if VAL_PICK else 0.0
PS_VAL['fill'] = VAL_PICK['_fill'] if VAL_PICK else 0.0
PS_VAL['rank'] = ([x['bookId'] for x in POW_CANDS.get(VAL_PICK['_type1'], [])].index(VAL_PICK['bookId']) + 1
                  if VAL_PICK else 0)
# 两个单目标参考项，用来看折中取舍
PS_VAL['topV'] = max(VAL_CANDS, key=lambda b: (b['_V'], b['_price'])) if VAL_CANDS else None
PS_VAL['topPrice'] = max(VAL_CANDS, key=lambda b: (b['_price'], b['_V'])) if VAL_CANDS else None

# ④ 抢下架：只看价值分 ≥ 70 的高价值书，最大化 Σ(下架风险 × 价值分)
RISK_POOL = [b for b in COIN_POOL_X if (b['_vscore'] or 0) >= 70]
RISK_PLAN, RISK_SPENT = knap_max(RISK_POOL, COIN,
                                 lambda b: (b['_risk'] or 0) * (b['_vscore'] or 0))
RISK_PLAN = sorted(RISK_PLAN, key=lambda b: (-b['_risk'], -(b['_vscore'] or 0)))
PS_RISK = plan4(RISK_PLAN)

OVERLAP_RR = len({b['bookId'] for b in RATIO_PLAN} & {b['bookId'] for b in RISK_PLAN})
OVERLAP_RC = len({b['bookId'] for b in RATIO_PLAN} & {b['bookId'] for b in MAIN_PLAN})
OVERLAP_RV = len({b['bookId'] for b in RATIO_PLAN} & {b['bookId'] for b in VAL_PLAN})
RISK_AVG = round(sum(b['_risk'] for b in COIN_POOL_X) / len(COIN_POOL_X), 1) if COIN_POOL_X else 0


def write_header(ws, headers, widths, start_row=1):
    for i, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(row=start_row, column=i, value=h)
        c.fill = HEAD_FILL
        c.font = HEAD_FONT
        c.alignment = CENTER
        c.border = BORDER
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[start_row].height = 26
    ws.freeze_panes = ws.cell(row=start_row + 1, column=1)


def put_row(ws, r, values, money_cols=(), cat=None, height=None):
    for i, v in enumerate(values, 1):
        c = ws.cell(row=r, column=i, value=v)
        c.font = BODY_FONT
        c.border = BORDER
        if i in money_cols:
            c.number_format = MONEY
            c.alignment = Alignment(horizontal='right', vertical='center')
        else:
            c.alignment = WRAP if i <= 2 else CENTER
    if cat and cat in CAT_COLOR:
        fill = PatternFill('solid', fgColor=CAT_COLOR[cat])
        for i in range(1, len(values) + 1):
            ws.cell(row=r, column=i).fill = fill
    if height:
        ws.row_dimensions[r].height = height


def book_rows(books, with_cat=True, sort_key=None):
    key = sort_key or (lambda b: (-b['_price'], str(b.get('title'))))
    rows = sorted(books, key=key)
    out = []
    for i, b in enumerate(rows, 1):
        rec = [i, b.get('title') or '', b.get('author') or '']
        if with_cat:
            rec.append(b['_cat'])
        rec += [
            b['_price'],
            ('%d%%' % b['_prog']) if b['_prog'] is not None else '—',
            b['_pstat'],
            b['_rhours'],
            b['_recent'],
            '已下架' if b.get('soldout') == 1 else '',
            b.get('category') or '未分类',
            b.get('payType'),
            b.get('bookId'),
        ]
        out.append((rec, b['_cat']))
    return out


def build_overview(wb, wallet=None):
    ws = wb.create_sheet('概览')
    ws.sheet_view.showGridLines = False
    for i, w in enumerate([26, 19, 15, 17, 20, 30], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    user = (wallet or {}).get('user') or {}
    ws['A1'] = '微信读书·账户资产与阅读权益统计'
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:F1')
    ws.row_dimensions[1].height = 34
    _acct = user.get('name')
    if not _acct:
        try:
            _acct = (json.load(open('_wallet3.json.bak', encoding='utf-8')).get('user') or {}).get('name')
        except Exception:
            _acct = None
    ws['A2'] = ('账号：%s    数据源：/web/shelf/sync + /web/pay/balance（登录态）    本次运行：%s'
                % (_acct or '未登录', SNAP_NOW))
    ws['A2'].font = Font(size=9, color='808080', name='微软雅黑')
    ws.merge_cells('A2:F2')
    ws['A3'] = ('底层数据抓取时间（每次运行都重新抓，不复用旧文件）：'
                + '　｜　'.join('%s %s' % (k, v) for k, v in SNAP.items()))
    ws['A3'].font = Font(size=9, color='2F5496', name='微软雅黑')
    ws.merge_cells('A3:F3')

    r = 4

    # ---------- 一、账户资产 ----------
    ws.cell(row=r, column=1, value='一、我的账户资产（书币 / 赠币 / 体验卡）').font = SUB_FONT
    r += 1
    for col, txt in ((1, '项目'), (2, '数值'), (4, '说明')):
        c = ws.cell(row=r, column=col, value=txt)
        c.fill = HEAD_FILL
        c.font = HEAD_FONT
        c.alignment = CENTER
        c.border = BORDER
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=6)
    ws.row_dimensions[r].height = 24
    r += 1

    def wallet_row(label, value, note, color='1F3864', size=12):
        nonlocal r
        c = ws.cell(row=r, column=1, value=label)
        c.font = Font(bold=True, size=10, name='微软雅黑')
        c.alignment = Alignment(vertical='center')
        c.border = BORDER
        v = ws.cell(row=r, column=2, value=value)
        v.font = Font(bold=True, size=size, color=color, name='微软雅黑')
        v.alignment = Alignment(horizontal='center', vertical='center')
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        n = ws.cell(row=r, column=4, value=note)
        n.font = Font(size=9, color='595959', name='微软雅黑')
        n.alignment = Alignment(vertical='center', wrap_text=True)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=6)
        ws.row_dimensions[r].height = 26
        r += 1

    if wallet:
        wa = wallet.get('android') or {}
        wi = wallet.get('ios') or {}
        mc = wallet.get('memberCard') or {}
        coin_total = float(wa.get('balance') or wi.get('peerBalance') or 0)
        coin_gift = float(wa.get('giftBalance') or wi.get('peerGiftBalance') or 0)
        coin_paid = round(coin_total - coin_gift, 2)
        gift_pct = ('%.0f%%' % (100 * coin_gift / coin_total)) if coin_total else '-'
        if mc.get('expiredTime'):
            exp_dt = datetime.fromtimestamp(mc['expiredTime']).strftime('%Y-%m-%d %H:%M:%S')
        else:
            exp_dt = '-'
        hours = round(float((wa.get('welfare') or {}).get('expiredTime') or 0) / 3600.0, 1)
        avail = [b for b in NEED if 0 < b['_price'] <= coin_gift + 1e-9 and b.get('soldout') != 1]
        left, cnt = COIN_PLAN_LEFT, len(COIN_PLAN)

        wallet_row('书币总余额', '%.2f 元' % coin_total,
                   '可用于购买电子书（微信读书「书币」）')
        wallet_row('　其中：赠币', '%.2f 元' % coin_gift,
                   '占总余额 %s；赠币有有效期，优先消耗' % gift_pct, color='C00000')
        wallet_row('　充值书币', '%.2f 元' % coin_paid,
                   '无有效期', color='404040')
        wallet_row('最近到期的权益', '体验卡 ' + exp_dt[:16],
                   '剩余约 %.1f 小时；到期后 %d 本待付费书将无法阅读（赠币不受影响）'
                   % (hours, len(NEED)),
                   color='C00000', size=11)
        wallet_row('赠币可覆盖的书', '%d 本' % len(avail),
                   '待付费 %d 本中，标价 ≤ %.2f 元的共 %d 本（含 %d 本已下架），剔除下架后可买 %d 本'
                   % (len(NEED), coin_gift, len(COIN_OK), COIN_OK_SOLDOUT, len(avail)))
        wallet_row('按本数最多买', '%d 本' % MAX_N,
                   '本数上限（0/1 背包）。推荐组合：排除已读完 + 评分最优 = %d 本 / %.2f 元 / 剩 %.2f 元；'
                   '若要不排除已读完的最省组合：%d 本 / %.2f 元 / 剩 %.2f 元'
                   % (MAIN_STAT['n'], MAIN_STAT['spent'], MAIN_STAT['left'],
                      cnt, round(COIN_PLAN_SPENT, 2), round(left, 2)))
        wallet_row('积分购买方案', '见「%s」' % COIN_SHEET,
                   '单本标价 ≤ 赠币余额的书逐本列出（%d 本），并给出四个购买方案：'
                   '① 性价比最高 / ② 本数最多 / ③ 绝对价值最高（只买 1 本）/ ④ 抢下架优先'
                   % len(COIN_OK),
                   color='1F3864', size=11)
        wallet_row('赠币逐笔到期', '需查 App',
                   '网页端 /web/pay/balance 只返回赠币总额，无明细接口；'
                   '逐笔到期请在微信读书 App「我的 → 书币」查看', color='808080', size=11)
    else:
        wallet_row('书币 / 赠币余额', '未取到', '未找到 _wallet3.json，请先运行 fetch_all.py')

    # ---------- 二、四类阅读权限 ----------
    r += 1
    ws.cell(row=r, column=1, value='二、四类阅读权限'
            + ('（与你 App 显示的数字对照）' if APP_REF else '')).font = SUB_FONT
    r += 1
    heads = (['权益类型', '本表数量', 'App显示', '差异', '标价合计(元)', '说明'] if APP_REF
             else ['权益类型', '本表数量', '标价合计(元)', '说明'])
    for i, h in enumerate(heads, 1):
        c = ws.cell(row=r, column=i, value=h)
        c.fill = HEAD_FILL
        c.font = HEAD_FONT
        c.alignment = CENTER
        c.border = BORDER
    r += 1
    note = {
        '付费会员卡可读': '需开通付费会员卡才能读全文',
        '体验卡可读': '目前靠体验卡，体验卡过期后需会员卡',
        '永久可读': '已购买/已拥有，永久可读',
        '免费': '导入书/公众号文章，无标价',
    }
    for cat in CAT_ORDER:
        n = agg[cat]['n']
        m = round(agg[cat]['money'], 2)
        if APP_REF:
            a = APP_REF[cat]
            d = n - a
            vals = [cat, n, a, ('%+d' % d) if d else '一致', m, note[cat]]
            money_cols = (5,)
        else:
            d = 0
            vals = [cat, n, m, note[cat]]
            money_cols = (3,)
        put_row(ws, r, vals, money_cols=money_cols, cat=cat)
        ws.cell(row=r, column=2).alignment = CENTER
        if APP_REF:
            ws.cell(row=r, column=3).alignment = CENTER
            dc = ws.cell(row=r, column=4)
            dc.alignment = CENTER
            dc.font = Font(size=10, name='微软雅黑', bold=(d != 0),
                           color='006100' if d == 0 else 'C00000')
        r += 1
    if APP_REF:
        put_row(ws, r, ['合计', TOTAL_N, sum(APP_REF.values()),
                        '%+d' % (TOTAL_N - sum(APP_REF.values())), TOTAL_MONEY, ''])
        ncol = 6
    else:
        put_row(ws, r, ['合计', TOTAL_N, TOTAL_MONEY, ''], money_cols=(3,))
        ncol = 4
    for i in range(1, ncol + 1):
        ws.cell(row=r, column=i).font = Font(bold=True, size=10, name='微软雅黑')

    # ---------- 三、需要花钱的书 ----------
    r += 2
    ws.cell(row=r, column=1, value='三、需要花钱的书（核心结论）').font = SUB_FONT
    r += 1
    cards = [
        ('需要付费/需会员才能读的书', '%d 本' % len(NEED), '会员卡可读 %d 本 + 体验卡可读 %d 本'
         % (agg['付费会员卡可读']['n'], agg['体验卡可读']['n'])),
        ('这些书的标价合计', '%.2f 元' % NEED_MONEY, '按单本买断价累加，仅作价值参考'),
        ('其中已下架（买不到）', '%d 本' % NEED_SOLDOUT, '仍占标价 %.2f 元'
         % round(sum(b['_price'] for b in NEED if b.get('soldout') == 1), 2)),
        ('已拥有（永久可读）', '%d 本' % agg['永久可读']['n'], '标价合计 %.2f 元，无需再付钱'
         % round(agg['永久可读']['money'], 2)),
        ('付费会员卡订阅价', '19 元/月', '需另行微信支付：书币/赠币不能抵扣，也不能用来开卡（详见 方法与口径说明.md）'),
        ('已读完但没买断', '%d 本' % len(NEED_DONE),
         '待付费书里进度 100% 的书：靠会员卡/体验卡读完，买断只是留永久副本'),
        ('在读（未读完）的书', '%d 本' % len(NEED_READING),
         '进度 1~99%%：其中快读完（≥80%%）%d 本，见「待付费书单」的阅读进度列' % len(NEED_FAST)),
        ('累计阅读时长', '%.0f 小时' % TOTAL_HOURS,
         '书架 %d 本的 readingTime 汇总（仅网页/App 同步到的记录）' % len(BOOKS)),
    ]
    for k, v, d in cards:
        ws.cell(row=r, column=1, value=k).font = Font(bold=True, size=10, name='微软雅黑')
        ws.cell(row=r, column=1).alignment = Alignment(vertical='center')
        c = ws.cell(row=r, column=2, value=v)
        c.font = Font(bold=True, size=12, color='C00000', name='微软雅黑')
        c.alignment = Alignment(horizontal='center', vertical='center')
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        d0 = ws.cell(row=r, column=4, value=d)
        d0.font = Font(size=9, color='808080', name='微软雅黑')
        d0.alignment = Alignment(vertical='center')
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=6)
        ws.row_dimensions[r].height = 24
        r += 1

    # ---------- 三之二、四个购买方案（合并后统一放在「积分购买方案」）----------
    r += 1
    ws.cell(row=r, column=1,
            value='三之二、%.2f 元的四个买法（明细逐本见「%s」sheet 第三~六节）'
                  % (COIN, PLAN_SHEET)).font = SUB_FONT
    r += 1
    for i, h in enumerate(['方案', '本数', '花费(元)', '剩余(元)', '平均价值评分', '每元价值'], 1):
        c = ws.cell(row=r, column=i, value=h)
        c.fill = HEAD_FILL
        c.font = HEAD_FONT
        c.alignment = CENTER
        c.border = BORDER
    r += 1
    for nm, st in (('① 性价比最高（推荐）', PS_RATIO),
                   ('② 本数最多', PS_CNT),
                   ('③ 绝对价值最高（只买 1 本）', PS_VAL),
                   ('④ 抢下架优先', PS_RISK)):
        _per = '—' if st is PS_VAL else st['per']   # ③ 只有 1 本，每元价值与多本方案不可比
        put_row(ws, r, [nm, st['n'], st['spent'], st['left'], st['av'], _per],
                money_cols=(3, 4))
        for i in (2, 5, 6):
            ws.cell(row=r, column=i).alignment = Alignment(horizontal='center', vertical='center')
            if i != 6 or st is not PS_VAL:
                ws.cell(row=r, column=i).number_format = '0.0'
        r += 1
    nc = ws.cell(row=r, column=1,
                 value='※ 余额 %.2f 元是书币，只能按标价买断电子书：不能提现，也不能用来支付 19 元/月的会员卡。'
                       '「单本 ≤ %.2f 元的书 %d 本」≠「%d 本都买得起」，余额实际只够买 %d 本（每本只买一次）。'
                       '③ 只有 1 本，平均价值分是单本数值、每元价值列留空（与①②④的多本合计不可比）。'
                       '四个方案都不含「已读完但未买断（未获永久授权）」的书。'
                       % (COIN, COIN, len(COIN_OK), len(COIN_OK), PS_CNT['n']))
    nc.font = Font(size=9, color='C00000', name='微软雅黑')
    nc.alignment = Alignment(vertical='center', wrap_text=True)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    ws.row_dimensions[r].height = 46

    # ---------- 四、按分类目聚合 ----------
    r += 1
    ws.cell(row=r, column=1, value='四、待付费书单·按分类目聚合（标价前 12）').font = SUB_FONT
    r += 1
    for i, h in enumerate(['分类目', '本数', '标价合计(元)', '均值(元)'], 1):
        c = ws.cell(row=r, column=i, value=h)
        c.fill = HEAD_FILL
        c.font = HEAD_FONT
        c.alignment = CENTER
        c.border = BORDER
    r += 1
    bycat = defaultdict(lambda: {'n': 0, 'm': 0.0})
    for b in NEED:
        k = b.get('category') or '未分类'
        bycat[k]['n'] += 1
        bycat[k]['m'] += b['_price']
    for k, d in sorted(bycat.items(), key=lambda x: -x[1]['m'])[:12]:
        put_row(ws, r, [k, d['n'], round(d['m'], 2), round(d['m'] / d['n'], 2)], money_cols=(3, 4))
        r += 1

    # ---------- 五、口径与校验 ----------
    r += 1
    ws.cell(row=r, column=1, value='五、口径与校验').font = SUB_FONT
    r += 1
    lines = [
        '账户资产：登录态调用 POST /web/pay/balance（body {"pf":"android_html"}），返回 balance / giftBalance / peerBalance / peerGiftBalance；',
        '   体验卡到期取 GET /web/pay/memberCardSummary?pf=ios 的 expiredTime 字段。',
        '赠币到期：该接口只返回赠币总额，不返回逐笔有效期；网页端全部 pay 接口仅 balance 与 memberCardSummary 两个，无明细接口，',
        '   故"哪笔赠币最先到期"只能在微信读书 App「我的 → 书币」查看。表中"最近到期"列出的是体验卡到期时间。',
        '权益分类：登录态调用 /web/shelf/sync，逐本读取 paid 与 payingStatus 两个字段。',
        '字段校验（可在浏览器里打开任意付费书复验）：付费墙文案与 paid / payingStatus 一一对应——',
        '   paid=0 且为会员卡档 → 显示「会员卡 19元/月 · 开通后即可阅读」',
        '   paid=0 且为体验卡档 → 显示「你正在使用体验卡阅读付费部分」',
        '   paid=1 → 直接进正文，无付费墙',
        '已拥有校验：paid=1 的书一律计入永久可读，无需再付费。',
        '价格口径：标价（售价）= centPrice/100，单位为元；非"你还要付的钱"。',
    ]
    if APP_REF:
        lines += [
            '体验卡可读本表 %d 本、App 显示 %d 本，差 %d 本：书架接口另返回「文章收藏」(mpbook) 等非电子书条目。'
            % (agg['体验卡可读']['n'], APP_REF['体验卡可读'],
               agg['体验卡可读']['n'] - APP_REF['体验卡可读']),
            '   App 四类合计 %d，本表覆盖 %d 本电子书，差额即在此。'
            % (sum(APP_REF.values()), TOTAL_N),
        ]
    for t in lines:
        ws.cell(row=r, column=1, value=t).font = Font(size=9, name='微软雅黑',
                                                    color='C00000' if '差 ' in t else '404040')
        ws.cell(row=r, column=1).alignment = Alignment(vertical='center')
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        r += 1


def build_plan_sheet(wb):
    """积分购买方案（合并原「积分可购」+「购书方案」两个 sheet）：
    余额约束 + 四个购买方案（性价比最高 / 抢下架 / 绝对价值最高 / 本数最多）+ 全部可买清单。
    """
    ws = wb.create_sheet(COIN_SHEET)
    ws.sheet_view.showGridLines = False
    widths = [5, 42, 20, 13, 11, 10, 10, 10, 10, 10, 10, 10, 10, 20, 20, 9, 15]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    N_OK = len(COIN_OK)
    N_BUY = len(COIN_BUY)
    N_SOLD = COIN_OK_SOLDOUT
    _wa = (WALLET or {}).get('android') or {}
    BAL = round(float(_wa.get('balance') or (WALLET or {}).get('ios', {}).get('peerBalance') or COIN), 2)
    PAID = round(BAL - COIN, 2)

    ws['A1'] = '积分购买方案：%.2f 元书币能买什么、按什么顺序买' % COIN
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:L1')
    ws.row_dimensions[1].height = 34
    ws['A2'] = ('余额 %.2f 元（赠币 %.2f ＋ 充值书币 %.2f，%s 刷新）　·　'
                '候选池：待付费 %d 本里，单本标价 ≤ %.2f 元的 %d 本（其中已下架 %d 本、真实可买 %d 本）　·　'
                '0/1 背包精确解，每本只买一次'
                % (BAL, COIN, PAID, SNAP['账户资产'],
                   len(NEED), COIN, N_OK, N_SOLD, N_BUY))
    ws['A2'].font = Font(size=9, color='808080', name='微软雅黑')
    ws.merge_cells('A2:L2')
    ws['A3'] = ('※ 先纠正一处容易误读的表述：下面说的「%d 本」是「单本标价 ≤ %.2f 元、任意一本单独买都付得起」的书，'
                '不是「%.2f 元能把 %d 本全买下来」。这 %d 本如果全买要 %.2f 元，是你余额的 %.1f 倍；'
                '余额实际只够买其中 %d 本（见第二节方案②）。'
                % (N_OK, COIN, COIN, N_OK, N_OK, COIN_OK_MONEY,
                   COIN_OK_MONEY / COIN if COIN else 0, PS_CNT['n']))
    ws['A3'].font = Font(size=9, color='C00000', name='微软雅黑')
    ws['A3'].alignment = Alignment(vertical='center', wrap_text=True)
    ws.merge_cells('A3:L3')
    ws.row_dimensions[3].height = 32
    ws['A4'] = ('※ 余额是「书币/赠币」，只能在微信读书里按标价买断电子书：不能提现、不能当现金用，'
                '也不能用来支付 19 元/月的付费会员卡订阅（会员卡只能另外用微信支付）。'
                '所以下面全部方案都是「拿这 %.2f 元买断书」，不涉及开卡。' % COIN)
    ws['A4'].font = Font(size=9, color='2F5496', name='微软雅黑')
    ws['A4'].alignment = Alignment(vertical='center', wrap_text=True)
    ws.merge_cells('A4:L4')
    ws.row_dimensions[4].height = 28

    r = 6

    def sec(t):
        nonlocal r
        ws.cell(row=r, column=1, value=t).font = SUB_FONT
        r += 1

    def table(heads, rows_, money_cols=(), fill=HEAD_FILL, widths_=None):
        nonlocal r
        for i, h in enumerate(heads, 1):
            c = ws.cell(row=r, column=i, value=h)
            c.fill = fill
            c.font = HEAD_FONT
            c.alignment = CENTER
            c.border = BORDER
        ws.row_dimensions[r].height = 26
        r += 1
        for row in rows_:
            put_row(ws, r, list(row), money_cols=money_cols)
            r += 1

    def note(t, color='808080', size=9, cols=17, height=26):
        nonlocal r
        c = ws.cell(row=r, column=1, value=t)
        c.font = Font(size=size, color=color, name='微软雅黑')
        c.alignment = Alignment(vertical='center', wrap_text=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=cols)
        ws.row_dimensions[r].height = height
        r += 1

    # ---------- 一、余额与范围 ----------
    def mt(spec, head=False, height=None):
        """按合并单元格写一行：spec = [(value, 起始列, 结束列, 数字格式)]"""
        nonlocal r
        for val, c1, c2, fmt in spec:
            cell = ws.cell(row=r, column=c1, value=val)
            if c2 > c1:
                ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
            for c in range(c1, c2 + 1):
                cc = ws.cell(row=r, column=c)
                cc.border = BORDER
                if head:
                    cc.fill = HEAD_FILL
            if head:
                cell.font = HEAD_FONT
                cell.alignment = CENTER
            elif isinstance(val, (int, float)):
                cell.font = BODY_FONT
                cell.alignment = Alignment(horizontal='center', vertical='center')
            else:
                cell.font = BODY_FONT
                cell.alignment = Alignment(vertical='center', wrap_text=True)
            if fmt:
                cell.number_format = fmt
        if height:
            ws.row_dimensions[r].height = height
        r += 1

    sec('一、余额能买什么（先看约束，再看方案）')
    mt([('项目', 1, 2, None), ('数值', 3, 3, None), ('说明', 4, 17, None)], head=True, height=24)
    rows_ = [
        ('可用余额', '%.2f 元' % COIN,
         '全部为书币（赠币＋充值币）；只能按标价买断电子书，不能提现、不能付会员卡'),
        ('单本买得起的书', '%d 本' % N_OK,
         '待付费 %d 本里单本标价 ≤ %.2f 元的书 —— 任选其一都付得起，但不是 %d 本全买得起'
         % (len(NEED), COIN, N_OK)),
        ('这 %d 本全买要' % N_OK, '%.2f 元' % COIN_OK_MONEY,
         '是余额的 %.1f 倍；均价 %.2f 元/本，余额只能覆盖总价的 %.1f%%'
         % (COIN_OK_MONEY / COIN if COIN else 0,
            COIN_OK_MONEY / N_OK if N_OK else 0,
            100.0 * COIN / COIN_OK_MONEY if COIN_OK_MONEY else 0)),
        ('余额实际能买', '%d 本' % PS_CNT['n'], cap_note()),
        ('其中已下架（买不到）', '%d 本' % N_SOLD,
         '标价虽在余额内，但已下架、无法购买也无法阅读；下面所有方案都已剔除'),
        ('能不能开会员卡', '不能',
         '余额不是现金，无法支付 19 元/月会员卡；会员卡需另外用微信支付，与这份方案无关'),
    ]
    for _lab, _val, _note in rows_:
        mt([(_lab, 1, 2, None), (_val, 3, 3, None), (_note, 4, 17, None)], height=26)
    note('※ 重要口径：下面四个方案，全部都不包含「你已经读完（进度 100%%）、但还没买断（= 未获永久授权）」'
         '的 %d 本书 —— 它们靠体验卡/会员卡读完了，再买只是拿个永久副本，没有阅读增量。'
         % (len(COIN_POOL) - len(COIN_POOL_X)), color='C00000', size=10, height=30)
    r += 1

    # ---------- 二、四个方案 ----------
    sec('二、四个购买方案对比（都是拿这 %.2f 元买断，选一个执行即可，不要叠加）' % COIN)
    mt([('方案', 1, 2, None), ('怎么选（目标）', 3, 6, None), ('本数', 7, 7, None),
        ('花费(元)', 8, 8, None), ('剩余(元)', 9, 9, None), ('总价值分', 10, 10, None),
        ('平均价值分', 11, 11, None), ('每元价值', 12, 12, None), ('平均下架风险', 13, 13, None)],
       head=True, height=26)
    plans_ = [
        ('① 性价比最高（推荐）',
         '按「价值评分 ÷ 标价」从高到低装，每一元买到的价值最多', PS_RATIO),
        ('② 本数最多',
         '0/1 背包本数优先，同本数下价值分合计最高 —— 书架厚度最大化', PS_CNT),
        ('③ 绝对价值最高',
         '只买 1 本：候选=基准表口径的 %d 个一级分类 × 每类 V 前 2 名（%d 个候选位，实到 %d 本）；'
         '挑一本花满积分且本身最值的'
         % (len(TYPE1_STD), len(TYPE1_STD) * 2, PS_VAL['nCand']),
         PS_VAL),
        ('④ 抢下架优先',
         '只看价值分 ≥ 70 的书，最大化「下架风险 × 价值分」——先买可能没了的好书', PS_RISK),
    ]
    for _nm, _desc, _st in plans_:
        mt([(_nm, 1, 2, None), (_desc, 3, 6, None), (_st['n'], 7, 7, '0'),
            (_st['spent'], 8, 8, MONEY), (_st['left'], 9, 9, MONEY), (_st['tv'], 10, 10, '0.0'),
            (_st['av'], 11, 11, '0.0'), (_st['per'], 12, 12, '0.0'), (_st['risk'], 13, 13, '0.0')],
           height=30)
    note('※ 四个方案共同的候选池：待付费 %d 本 → 剔除已下架 %d 本 → 再剔除「已读完但未买断（未获永久授权）」'
         ' %d 本 → 剩 %d 本可买。所以四个方案里都不会出现你已经读完、只差授权的书。'
         % (len(NEED), N_SOLD, len(COIN_POOL) - len(COIN_POOL_X), len(COIN_POOL_X)),
         color='1F3864', size=10, height=26)
    note('※ 推荐 ①：同样花 %.2f 元，它的「每元价值」%.1f 分/元最高（② 只有 %.1f 分/元）。'
         '①与④重合 %d 本 —— 性价比最高的那批书，本身也多是高价值+高风险的书，两个方案可以互相印证；'
         '①与②重合 %d/%d 本，说明在这批书里"每元价值最高"与"能买最多本"几乎是同一件事，差别主要在排序。'
         % (COIN, PS_RATIO['per'], PS_CNT['per'], OVERLAP_RR,
            OVERLAP_RC, PS_CNT['n']), color='C00000', size=10, height=34)
    note('※ ③ 与其余三个方案的量纲不同：它只有 1 本，所以「总价/每元价值」都是单本的数，'
         '不要拿去和①②④的总和比。看③只需看三个数：花满度（把积分花掉多少）、综合价值 V、以及它的候选排名。',
         color='C00000', size=10, height=26)
    note('※ 「每元价值」= 总价值分 ÷ 总花费。它天然偏爱低价书（分母小），'
         '所以①的书目会偏「薄」一些；想「只买一本、把积分一次花光」的买法，用③（③的本数栏就是 1）。'
         '「平均下架风险」是该方案书目的风险均值，候选池基准 %.1f 分。' % RISK_AVG,
         color='595959', height=28)
    r += 1

    # ---------- 三、各方案逐本 ----------
    PH = ['#', '书名', '作者', '权益类型', '价值分', '标价(元)', '累计花费',
          '剩余(元)', '每元价值', '下架风险', '阅读进度']

    def plan_table(pk, st, title, fill='C00000', extra_note=None):
        nonlocal r
        ws.cell(row=r, column=1, value=title).font = Font(bold=True, size=10, color='1F3864', name='微软雅黑')
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=11)
        r += 1
        for i, h in enumerate(PH, 1):
            c = ws.cell(row=r, column=i, value=h)
            c.fill = PatternFill('solid', fgColor=fill)
            c.font = HEAD_FONT
            c.alignment = CENTER
            c.border = BORDER
        ws.row_dimensions[r].height = 26
        r += 1
        spent = 0.0
        for i, b in enumerate(pk, 1):
            spent = round(spent + b['_price'], 2)
            prog = '—' if b['_prog'] is None else ('%d%%' % b['_prog'])
            put_row(ws, r, [i, b.get('title') or '', b.get('author') or '', b['_cat'],
                            b['_vscore'] if b['_vscore'] else '—', b['_price'], spent,
                            round(COIN - spent, 2),
                            round(b['_vscore'] / b['_price'], 1) if (b['_vscore'] and b['_price']) else '—',
                            b['_risk'], prog],
                    money_cols=(6, 7, 8), cat=b['_cat'])
            for c_, fmt in ((5, '0.0'), (9, '0.0'), (10, '0.0')):
                ws.cell(row=r, column=c_).number_format = fmt
                ws.cell(row=r, column=c_).alignment = Alignment(horizontal='center', vertical='center')
            ws.cell(row=r, column=11).alignment = Alignment(horizontal='center', vertical='center')
            r += 1
        put_row(ws, r, ['合计', '%d 本' % st['n'], '', '', st['tv'], st['spent'],
                        st['spent'], st['left'], st['per'], st['risk'], ''],
                money_cols=(6, 7, 8))
        for c_ in range(1, 12):
            ws.cell(row=r, column=c_).font = Font(bold=True, size=10, name='微软雅黑')
        for c_, fmt in ((5, '0.0'), (9, '0.0'), (10, '0.0')):
            ws.cell(row=r, column=c_).number_format = fmt
        r += 1
        if extra_note:
            note(extra_note, color='595959', height=24)
        r += 1

    sec('三、方案①「性价比最高」逐本（价值分 ÷ 标价 降序装满 %.2f 元）' % COIN)
    if RATIO_PLAN:
        _lo = min(b['_price'] for b in RATIO_PLAN)
        _nlo = sum(1 for b in RATIO_PLAN if abs(b['_price'] - _lo) < 1e-9)
        _foot_ratio = ('※ 这一份是「每一元买到的价值最多」的解：书目普遍是低价 + 高价值分的书。'
                       '其中 %d 本标价仅 %.2f 元，把「每元价值」拉到 %.0f（分母极小），'
                       '它们合计只花 %.2f 元，不影响其余书的选择。'
                       '若你更在意"每个大类都覆盖到"，看第五节方案③。'
                       % (_nlo, _lo, PS_RATIO['per'], round(_nlo * _lo, 2)))
    else:
        _foot_ratio = ('※ 余额 %.2f 元买不到任何一本，本方案为空（不是没有候选书，是预算不够）。'
                       % COIN)
    plan_table(RATIO_PLAN, PS_RATIO,
               '① 性价比最高：%d 本 / 花 %.2f 元 / 剩 %.2f 元 / 每元价值 %.1f 分'
               % (PS_RATIO['n'], PS_RATIO['spent'], PS_RATIO['left'], PS_RATIO['per']),
               'C00000', _foot_ratio)

    sec('四、方案②「本数最多」逐本（本数优先，同本数下价值分合计最高）')
    plan_table(MAIN_PLAN, PS_CNT, cnt_plan_note(), '375623', cnt_plan_foot())

    # ---- 方案③：只买一本，一张不同的表 ----
    _pv = VAL_PICK
    _vtop, _ptop = PS_VAL['topV'], PS_VAL['topPrice']

    # 15 个标准分类 × 每类前 2 名 = 30 个候选位；逐类诊断空缺
    _pooln = {t: sum(1 for x in COIN_POOL_X if x['_type1'] == t and x['_vscore'])
              for t in TYPE1_STD}
    _shelf = {t: [x for x in BOOKS if x['_type1'] == t] for t in TYPE1_STD}
    _GLOB = {x['bookId']: i + 1 for i, x in enumerate(VAL_CANDS)}
    _slot = len(TYPE1_STD) * 2
    _empty = [t for t in TYPE1_STD if not POW_CANDS.get(t)]

    sec('五、方案③「绝对价值最高」逐本（只买 1 本；候选 = 「分类基准与类型评分」同一套 15 个一级分类，'
        '每类综合价值 V 前 2 名 = %d 个候选位）' % _slot)
    PH3 = ['#', '一级分类', '书名', '作者', '标价(元)', '剩余(元)', '花满度',
           '综合价值V', '单本价值产出\n(V×标价)', '价值分', '平台评分',
           '类内V排名', '该类平均\n价值分', '类排名\n(基准表)']
    if _pv:
        ws.cell(row=r, column=1,
                value='③ 绝对价值最高（一次只买 1 本）：《%s》 %.2f 元 / 剩 %.2f 元（花满 %.1f%%）/ '
                      '综合价值 V %.1f 分 / 价值分 %.1f'
                      % (_pv['title'], _pv['_price'], PS_VAL['left'], round(100 * PS_VAL['fill'], 1),
                         PS_VAL['V'], _pv['_vscore'])
                ).font = Font(bold=True, size=10, color='1F3864', name='微软雅黑')
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=14)
        r += 1
        for i, h in enumerate(PH3, 1):
            c = ws.cell(row=r, column=i, value=h)
            c.fill = PatternFill('solid', fgColor='1F3864')
            c.font = HEAD_FONT
            c.alignment = CENTER
            c.border = BORDER
        ws.row_dimensions[r].height = 32
        r += 1
        put_row(ws, r, [1, _pv['_type1'], _pv.get('title') or '', _pv.get('author') or '',
                        _pv['_price'], PS_VAL['left'], round(100 * PS_VAL['fill'], 1), _pv['_V'],
                        _pv['_sv'], _pv['_vscore'], _pv['_rating'],
                        '%d / %d' % (PS_VAL['rank'], len(POW_CANDS.get(_pv['_type1'], []))),
                        round(TYPE1_MEAN.get(_pv['_type1']) or 0, 1),
                        TYPE1_RANK.get(_pv['_type1'], '')],
                money_cols=(5, 6))
        for c_, fmt in ((5, MONEY), (6, MONEY), (7, '0.0"%"'), (8, '0.0'), (9, '0.0'),
                        (10, '0.0'), (11, '0.0'), (13, '0.0')):
            ws.cell(row=r, column=c_).number_format = fmt
            ws.cell(row=r, column=c_).alignment = Alignment(horizontal='center', vertical='center')
        for c_ in (12, 14):
            ws.cell(row=r, column=c_).alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[r].height = 28
        r += 1
    else:
        # 余额连一本候选都买不起：不摆空表，直接说明原因（否则 _pv[...] 会 TypeError）
        _why = (('待付费书里最便宜的一本《%s》要 %.2f 元'
                 % (G_X.get('next_title') or '（未知书名）', G_X['next_price']))
                if G_X.get('next_price') else '可买池为空')
        c = ws.cell(row=r, column=1,
                    value='③ 绝对价值最高：本方案无书可选 —— 余额 %.2f 元低于全部候选标价（%s）。'
                          '攒够书币后重跑本脚本即可。' % (COIN, _why))
        c.font = Font(bold=True, size=10, color='C00000', name='微软雅黑')
        c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=14)
        ws.row_dimensions[r].height = 30
        r += 1

    # 候选清单：按 15 个标准分类逐类分组（每类 2 个候选位，缺位也留行）
    sec('　　候选清单：%d 个一级分类 × 每类 V 前 2 名 = %d 个候选位，'
        '实到 %d 本（%d 类有书 / %d 类候选位空缺）；本方案只在这 %d 本里挑'
        % (len(TYPE1_STD), _slot, PS_VAL['nCand'],
           len(TYPE1_STD) - len(_empty), len(_empty), PS_VAL['nCand']))
    PH3b = ['#', '一级分类', '书名', '作者', '标价(元)', '剩余(元)', '花满度',
            '综合价值V', '单本价值产出\n(V×标价)', '价值分', '平台评分',
            '类内V排名', '全局\n产出排名', '入选']
    for i, h in enumerate(PH3b, 1):
        c = ws.cell(row=r, column=i, value=h)
        c.fill = HEAD_FILL
        c.font = HEAD_FONT
        c.alignment = CENTER
        c.border = BORDER
    ws.row_dimensions[r].height = 32
    r += 1

    _seq = 0
    for _t in TYPE1_STD:
        _cand = POW_CANDS.get(_t, [])
        # 分组小标题
        _gh = ('类排名 %d ｜ %s ｜ 类均分 %.1f ｜ 可买池中有评分 %d 本 ｜ 本类候选 %d / 2 位'
               % (TYPE1_RANK.get(_t, 0), _t, TYPE1_MEAN.get(_t) or 0.0,
                  _pooln.get(_t, 0), len(_cand)))
        c = ws.cell(row=r, column=1, value=_gh)
        c.font = Font(bold=True, size=9, color='1F3864', name='微软雅黑')
        c.fill = PatternFill('solid', fgColor='DCE6F1')
        c.alignment = Alignment(horizontal='left', vertical='center')
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=14)
        for _c in range(1, 15):
            ws.cell(row=r, column=_c).border = BORDER
        ws.row_dimensions[r].height = 20
        r += 1

        if not _cand:
            _nb = len(_shelf.get(_t, []))
            _free = sum(1 for x in _shelf.get(_t, []) if not x['_price'])
            _sold = sum(1 for x in _shelf.get(_t, []) if x['_sold'])
            if _nb and _sold == _nb and _free == _nb:
                _why = '书架 %d 本全部已下架、且标价 0（永久可读/免费）' % _nb
            else:
                _why = ('书架 %d 本中 %d 本已下架、%d 本标价 0（免费/永久可读）'
                        % (_nb, _sold, _free))
            _cn = ('　　（该类候选位空缺：%s，待付费可买池里有评分的 0 本 → 本方案不选该类）'
                   % _why)
            c = ws.cell(row=r, column=1, value=_cn)
            c.font = Font(size=9, color='C00000', name='微软雅黑')
            c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=14)
            for _c in range(1, 15):
                ws.cell(row=r, column=_c).border = BORDER
            ws.row_dimensions[r].height = 22
            r += 1
            continue

        for _k, _b in enumerate(_cand, 1):
            _seq += 1
            hit = _b['bookId'] == _pv['bookId']
            put_row(ws, r, [_seq, _t, _b.get('title') or '', _b.get('author') or '',
                            _b['_price'], round(COIN - _b['_price'], 2),
                            round(100 * _b['_fill'], 1), _b['_V'], _b['_sv'],
                            _b['_vscore'], _b['_rating'], '%d / %d' % (_k, len(_cand)),
                            '%d / %d' % (_GLOB.get(_b['bookId'], 0), len(VAL_CANDS)),
                            '★ 本方案选它' if hit else ''],
                    money_cols=(5, 6))
            for c_, fmt in ((5, MONEY), (6, MONEY), (7, '0.0"%"'), (8, '0.0'), (9, '0.0'),
                            (10, '0.0'), (11, '0.0')):
                ws.cell(row=r, column=c_).number_format = fmt
                ws.cell(row=r, column=c_).alignment = Alignment(horizontal='center',
                                                                vertical='center')
            for c_ in (12, 13, 14):
                ws.cell(row=r, column=c_).alignment = Alignment(horizontal='center',
                                                                vertical='center')
            if hit:
                for c_ in range(1, 15):
                    ws.cell(row=r, column=c_).fill = PatternFill('solid', fgColor='FFF2CC')
                    ws.cell(row=r, column=c_).font = Font(bold=True, size=9, name='微软雅黑')
                ws.cell(row=r, column=14).font = Font(bold=True, size=9, color='C00000',
                                                      name='微软雅黑')
            r += 1
    r += 1

    note('※ ③ 的口径（只买一本，与①②④是不同的一路）：', color='C00000', size=10, height=22)
    note('   ① 候选结构：%d 个一级分类 × 每类 V 前 2 名 = %d 个候选位；'
         '其中 %d 类在待付费可买池里有可买的书，%d 类（%s）空缺 → 实到 %d 本。'
         '表里按 15 类逐类列出，空缺的类也留了一行写明原因，因此 28 是「本数」，不是「分类数」。'
         % (len(TYPE1_STD), _slot, len(TYPE1_STD) - len(_empty), len(_empty),
            '、'.join(_empty) or '无', PS_VAL['nCand']),
         color='595959', height=38)
    note('   ② 分类口径与「分类基准与类型评分」第一节完全一致：只取「有评分 ≥ 10 本、且非未分类」的一级分类，'
         '共 %d 类，类排名 = 该节按类均分从高到低的位次（可与该节 # 列逐行对照）。'
         '艺术 / 生活百科 / 漫画 / 童书等样本 < 10 本的类不进基准表，也不进本方案候选。'
         % len(TYPE1_STD),
         color='595959', height=32)
    note('   ③ 综合价值 V = 0.45×价值分 + 0.25×平台原评分 '
         '+ 0.15×参与率池内百分位 + 0.15×阅读人数池内百分位（价格不进 V）；'
         '类内 V 排名 = 该书在本类 2 个候选里的位次，全局产出排名 = 在 %d 本候选中按 V×标价 的位次。'
         % PS_VAL['nCand'],
         color='595959', height=30)
    note('   ④ 只买 1 本时的「剩余最少 + 价值最高」用乘积合并：单本价值产出 = V × 标价。'
         '标价越高 → 余额剩得越少（把积分一次花光）；V 越高 → 这本书本身越值。'
         '两者相乘，既不会为了省积分去挑一本几块钱的小书，也不会为了贵而买一本平庸的书。'
         '本方案选中全局产出排名第 1 的那本。',
         color='595959', height=32)
    if _vtop and _ptop and _pv:
        note('   ⑤ 两个单目标参考（用来看折中取舍）：'
             '只看 V 最高 → 《%s》V %.1f、价 %.2f、剩 %.2f；'
             '只看剩余最少 → 《%s》价 %.2f、V %.1f、剩 %.2f。'
             '本方案选的是两者的折中：《%s》价 %.2f、V %.1f、剩 %.2f（花满 %.1f%%）。'
             % (_vtop['title'], _vtop['_V'], _vtop['_price'], round(COIN - _vtop['_price'], 2),
                _ptop['title'], _ptop['_price'], _ptop['_V'], round(COIN - _ptop['_price'], 2),
                _pv['title'], _pv['_price'], _pv['_V'], PS_VAL['left'], round(100 * PS_VAL['fill'], 1)),
             color='595959', height=32)
    else:
        note('   ⑤ 本方案无书可选：余额 %.2f 元低于全部候选标价，'
             '单目标参考（V 最高 / 剩余最少）同样是空的 —— 不是算法出错，是预算不够。' % COIN,
             color='C00000', height=26)
    note('   ⑥ 与②④的区别：② 追本数、会掺进低价小书；④ 追下架风险、可能与 ③ 完全不重合；'
         '③ 只有 1 本，目的就是"把积分一次花在一本最值的书上"，适合你只想单次买一本、不想摊薄的情形。',
         color='595959', height=26)
    r += 1

    sec('六、方案④「抢下架」逐本（价值分 ≥ 70 且下架风险高，优先买可能没了的书）')
    plan_table(RISK_PLAN, PS_RISK,
               '④ 抢下架：%d 本 / 花 %.2f 元 / 剩 %.2f 元 / 平均下架风险 %.1f 分（候选池基准 %.1f）'
               % (PS_RISK['n'], PS_RISK['spent'], PS_RISK['left'], PS_RISK['risk'], RISK_AVG),
               '833C00',
               '※ 下架风险 = 100×(0.5×权益类型档实测下架率 + 0.3×出版年档 + 0.2×评分人数档)，'
               '均为书架实测值。这批书都是价值分 ≥ 70 的，属于"好书且可能先没"，最适合优先囤。')
    r += 1

    # ---------- 七、全部可买清单 ----------
    heads = ['#', '书名', '作者', '权益类型', '标价(元)', '每元价值', '价值分', '评分',
             '评分人数', '出版年', '下架风险', '阅读进度', '阅读状态', '分类', '所属书单',
             '状态', 'bookId']
    sec('七、全部可买清单：单本标价 ≤ %.2f 元的 %d 本（可买 %d 本在前，已下架 %d 本灰底垫底）'
        % (COIN, N_OK, N_BUY, N_SOLD))
    head_row = r
    for i, h in enumerate(heads, 1):
        c = ws.cell(row=r, column=i, value=h)
        c.fill = HEAD_FILL
        c.font = HEAD_FONT
        c.alignment = CENTER
        c.border = BORDER
    ws.row_dimensions[r].height = 28
    r += 1

    order = sorted(COIN_OK, key=lambda b: (b.get('soldout') == 1,
                                           -(b.get('_vscore') or 0),
                                           -((b.get('_vscore') or 0) / max(b['_price'], 0.01)),
                                           str(b.get('title'))))
    for i, b in enumerate(order, 1):
        _bln = BOOK_BL.get(str(b.get('bookId')), [])
        arch = '、'.join(_bln) if _bln else ('、'.join(BOOK_ARCH.get(str(b.get('bookId')), [])) or '')
        soldout = b.get('soldout') == 1
        vals = [i, b.get('title') or '', b.get('author') or '', b['_cat'], b['_price'],
                round((b['_vscore'] or 0) / b['_price'], 1) if b['_price'] else '',
                b['_vscore'] if b['_vscore'] else '—',
                b['_rating'] if b['_rating'] is not None else '',
                b['_rcnt'] or '', b['_pub'] or '', b['_risk'],
                ('%d%%' % b['_prog']) if b['_prog'] is not None else '—',
                b['_pstat'], b.get('category') or '未分类', arch,
                '已下架' if soldout else '可购买', b.get('bookId')]
        put_row(ws, r, vals, money_cols=(5,), cat=b['_cat'])
        for c_, fmt in ((6, '0.0'), (7, '0.0'), (8, '0.0'), (9, '#,##0'), (11, '0.0')):
            _cc = ws.cell(row=r, column=c_)
            _cc.number_format = fmt
            _cc.alignment = Alignment(horizontal='center', vertical='center')
        for c_ in (12, 13, 16):
            ws.cell(row=r, column=c_).alignment = Alignment(horizontal='center', vertical='center')
        pg = ws.cell(row=r, column=13)
        if pg.value in PSTAT_COLOR:
            pg.fill = PatternFill('solid', fgColor=PSTAT_COLOR[pg.value])
        st = ws.cell(row=r, column=16)
        if soldout:
            grey = PatternFill('solid', fgColor='F2F2F2')
            for j in range(1, len(vals) + 1):
                cc = ws.cell(row=r, column=j)
                cc.fill = grey
                cc.font = Font(size=10, name='微软雅黑', color='A6A6A6')
            st.font = Font(size=10, name='微软雅黑', color='C00000')
        else:
            st.font = Font(size=10, name='微软雅黑', color='006100')
        r += 1

    ws.auto_filter.ref = 'A%d:Q%d' % (head_row, r - 1)
    r += 1

    # ---------- 八、口径 ----------
    sec('八、口径与校验')
    for t in [
        '① 收录条件：待付费（会员卡可读 + 体验卡可读）且单本标价 ≤ %.2f 元、未下架；共 %d 本（含已下架 %d 本则 %d 本）。'
        % (COIN, N_BUY, N_SOLD, N_OK),
        '② 每本只买一次：候选池按 bookId 去重（%d 个 id 全部唯一，同名不同版本 %d 组）。'
        % (len(COIN_POOL), len(DUP_TITLE)),
        '③ 标价 = 详情页 centPrice ÷ 100（买断价）；若有活动折扣/赠币券，以 App 下单页为准。',
        '④ 阅读进度来自 /web/shelf/sync 的 bookProgress[].progress + /web/book/getProgress 的 finishTime；'
        '进度 100% = 已读完，与"是否已买断"无关（靠体验卡也能读完），故四个方案都排除了这批书。',
        '⑤ 价值评分 = 平台评分 × 统计下界 × 类内参与率 × 互动证据 × 受众面微调（只下调不上调，详见「价值评分」）。',
        '⑥ 下架风险分由书架实测下架率计算，仅表示"该分类/年份/冷门度下的历史下架比例"，不是平台公告。',
        '⑦ 余额是书币，只能在站内买断电子书：不能提现、不能购买会员卡、不能转赠；赠币有有效期，优先花。',
        '⑧ 已读完但未买断的 %d 本、标价超余额的 %d 本，分别在「已读完未买断」「待付费书单」sheet 单独列出。'
        % (len(NEED_DONE), len(NEED) - N_OK),
    ]:
        note(t, color='404040', height=24)
    return ws


def build_list_sheet(wb, name, books, with_cat=True, sort_key=None):
    ws = wb.create_sheet(name)
    headers = ['#', '书名', '作者']
    widths = [6, 46, 24]
    if with_cat:
        headers.append('权益类型')
        widths.append(14)
    headers += ['标价(元)', '阅读进度', '阅读状态', '阅读时长(h)', '最近阅读', '状态', '分类', 'payType', 'bookId']
    widths += [11, 11, 11, 13, 12, 9, 22, 12, 16]
    write_header(ws, headers, widths)
    money_col = headers.index('标价(元)') + 1
    iprog = headers.index('阅读进度') + 1
    istat = headers.index('阅读状态') + 1
    r = 2
    for rec, cat in book_rows(books, with_cat, sort_key):
        put_row(ws, r, rec, money_cols=(money_col,), cat=cat)
        ws.cell(row=r, column=iprog).alignment = Alignment(horizontal='center', vertical='center')
        st = ws.cell(row=r, column=istat)
        st.alignment = Alignment(horizontal='center', vertical='center')
        if st.value in PSTAT_COLOR:
            st.fill = PatternFill('solid', fgColor=PSTAT_COLOR[st.value])
        r += 1
    ws.auto_filter.ref = 'A1:%s%d' % (get_column_letter(len(headers)), r - 1)


def build_archive_sheet(wb):
    """书架分组（/web/shelf/sync 的 archive[]）—— 注意与 App 的「书单」不是同一个功能。"""
    ws = wb.create_sheet('书架分组明细(%d)' % len(ARCHIVES))
    ws.sheet_view.showGridLines = False
    widths = [20, 44, 22, 14, 11, 9, 16]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws['A1'] = '书架分组明细：%d 个分组 / %d 条书籍归属' % (len(ARCHIVES), _ARCH_ROWS)
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:G1')
    ws.row_dimensions[1].height = 30
    ws['A2'] = ('数据来源：登录态 GET https://weread.qq.com/web/shelf/sync → archive[]（每个分组含 archiveId/name/bookIds）'
                '    统计时点：%s' % datetime.now().strftime('%Y-%m-%d %H:%M'))
    ws['A2'].font = Font(size=9, color='595959', name='微软雅黑')
    ws.merge_cells('A2:G2')
    ws['A3'] = ('⚠ 这是「书架分组」，不是 App 里的「书单」。App 的「书单」（我的→书单）是独立功能，'
                '接口在 i.weread.qq.com，网页端既无入口也无接口，本工作簿拿不到。')
    ws['A3'].font = Font(size=9, bold=True, color='C00000', name='微软雅黑')
    ws.merge_cells('A3:G3')
    ws.row_dimensions[3].height = 26

    # ---------- 一、分组汇总 ----------
    r = 5
    ws.cell(row=r, column=1, value='一、分组汇总（按分组统计权益与下架）').font = SUB_FONT
    r += 1
    heads = ['分组', '本数', '已永久可读', '待付费', '已下架', '已读完未买断', '标价合计(元)']
    for i, h in enumerate(heads, 1):
        c = ws.cell(row=r, column=i, value=h)
        c.fill = HEAD_FILL
        c.font = HEAD_FONT
        c.alignment = CENTER
        c.border = BORDER
    r += 1
    _tot = [0, 0, 0, 0, 0, 0.0]
    for a in sorted(ARCHIVES, key=lambda x: -len(x.get('bookIds') or [])):
        bs = [BY_ID[str(i)] for i in (a.get('bookIds') or []) if str(i) in BY_ID]
        n = len(bs)
        p_own = sum(1 for b in bs if b['_cat'] == '永久可读')
        p_need = sum(1 for b in bs if b['_cat'] in ('会员卡可读', '体验卡可读'))
        p_sold = sum(1 for b in bs if b['_sold'])
        p_done = sum(1 for b in bs if b.get('_pstat') == '已读完')
        money = round(sum(b['_price'] for b in bs), 2)
        _tot[0] += n
        _tot[1] += p_own
        _tot[2] += p_need
        _tot[3] += p_sold
        _tot[4] += p_done
        _tot[5] += money
        put_row(ws, r, [a.get('name') or '未命名', n, p_own, p_need, p_sold, p_done, money],
                money_cols=(7,))
        for i in (2, 3, 4, 5, 6):
            ws.cell(row=r, column=i).alignment = CENTER
        r += 1
    put_row(ws, r, ['合计', _tot[0], _tot[1], _tot[2], _tot[3], _tot[4], round(_tot[5], 2)],
            money_cols=(7,))
    for i in (1, 2, 3, 4, 5, 6, 7):
        c = ws.cell(row=r, column=i)
        c.font = Font(bold=True, size=10, name='微软雅黑')
        c.alignment = CENTER if i > 1 else Alignment(horizontal='center', vertical='center')
        c.border = BORDER
        c.fill = PatternFill('solid', fgColor='D9E2F3')
    r += 1
    ws.cell(row=r, column=1,
            value='※ 分组里的书 %d%% 已永久可读、%d 本已下架，只有 %d 本是待付费——'
                  '所以这一页对「买什么书」参考价值有限，买书请看「%s」sheet。'
                  % (round(100.0 * _tot[1] / max(1, _tot[0])), _tot[3], _tot[2], PLAN_SHEET)
            ).font = Font(size=9, color='595959', name='微软雅黑')
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    ws.row_dimensions[r].height = 24

    # ---------- 二、逐本明细 ----------
    r += 2
    ws.cell(row=r, column=1, value='二、逐本明细（%d 条归属记录，同一本书可出现在多个分组）' % _ARCH_ROWS).font = SUB_FONT
    r += 1
    heads = ['分组', '书名', '作者', '权益类型', '标价(元)', '阅读状态', 'bookId']
    for i, h in enumerate(heads, 1):
        c = ws.cell(row=r, column=i, value=h)
        c.fill = HEAD_FILL
        c.font = HEAD_FONT
        c.alignment = CENTER
        c.border = BORDER
    r += 1
    head_row = r
    for a in ARCHIVES:
        nm = a.get('name') or '未命名书单'
        for bid in (a.get('bookIds') or []):
            b = BY_ID.get(str(bid))
            if not b:
                put_row(ws, r, [nm, '（书架中未找到）', '', '', 0, '', str(bid)])
                r += 1
                continue
            put_row(ws, r, [nm, b.get('title') or '', b.get('author') or '', b['_cat'],
                            b['_price'], b.get('_pstat') or '', b.get('bookId')],
                    money_cols=(5,), cat=b['_cat'])
            st = ws.cell(row=r, column=6)
            st.alignment = CENTER
            if st.value in PSTAT_COLOR:
                st.fill = PatternFill('solid', fgColor=PSTAT_COLOR[st.value])
            r += 1
    ws.auto_filter.ref = 'A%d:G%d' % (head_row, r - 1)


# ==================== 价值评分 v2（评分 × 参与率置信度） ====================
_OLD_PARAM = dict(RD_CAP=200000.0, R0=0.025, C_FLOOR=0.30)


def score_ext(rec):
    """不在书架的书：拿不到阅读人数 → 参与率未知，置信度按均值 0.8 处理（不奖不罚）。
    返回 (价值评分, 置信度, 评分人数)。"""
    raw = float(rec.get('newRating') or 0) / 10.0
    rc = float(rec.get('newRatingCount') or rec.get('ratingCount') or 0)
    if raw <= 0 or rc <= 0:
        return 0.0, 0.0, 0
    p = min(max(raw / 100.0, 0.0), 1.0)
    cf = min(wilson_lcb(p, rc) / p, 1.0) if p > 0 else VSPARAM.get('AT_MEAN', 0.8)
    return round(raw * cf, 1), round(cf, 4), int(rc)


def _old_parts(rc, rd):
    """复算旧模型（置信下界）的 C / B / n_eff，仅用于在新 sheet 里解释参数。"""
    import math
    b = min(1.0, math.log10(1 + rd) / math.log10(1 + _OLD_PARAM['RD_CAP'])) if rd > 0 else 0.75
    c = max(_OLD_PARAM['C_FLOOR'], min(1.0, (rc / rd) / _OLD_PARAM['R0'])) if rd > 0 else 1.0
    return c, b, rc * c * b


# -*- coding: utf-8 -*-
"""v9 新增/重写的工作表函数（由 build_price_excel_v9.py 引用）。"""


def _pct(x, d=1):
    return ('%.' + str(d) + 'f%%') % (x * 100) if x is not None else '—'


# ==================== 价值评分 ====================
def build_value_sheet_v4(wb):
    ws = wb.create_sheet('价值评分')
    ws.sheet_view.showGridLines = False
    widths = [5, 38, 20, 11, 10, 10, 11, 10, 10, 10, 9, 9, 9, 10, 10, 10, 10]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    P = VSPARAM or {}
    rated = [b for b in BOOKS if (b['_rating'] or 0) > 0 and b.get('_conf')]
    rated.sort(key=lambda b: -b['_vscore'])
    n = len(rated)
    avg_raw = sum(b['_rating'] for b in rated) / n if n else 0
    avg_v = sum(b['_vscore'] for b in rated) / n if n else 0
    avg_c = sum(b['_conf'] for b in rated) / n if n else 0
    n_zf = sum(1 for b in rated if b['_conf'] >= 0.999)

    ws['A1'] = '价值评分 = 统计置信下界 × 类内参与率 × 互动证据（划线/想法）× 受众面微调'
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:Q1')
    ws.row_dimensions[1].height = 30
    ws['A2'] = (f'有评分 {n} 本：平台原评分均值 {avg_raw:.1f} → 价值评分均值 {avg_v:.1f}'
                f'（折扣率 {100 * (1 - avg_v / avg_raw) if avg_raw else 0:.1f}%）'
                f'；平均置信度 {avg_c:.3f}，其中 {n_zf} 本无需下调。'
                f'全书架基准：参与率 r̄ = {VGLOBAL.get("mean_r", 0) * 100:.3f}%、'
                f'阅读人数几何均值 {VGLOBAL.get("geo_rd", 0):,.0f} 人')
    ws['A2'].font = Font(size=9, color='595959', name='微软雅黑')
    ws.merge_cells('A2:Q2')
    ws['A3'] = ('四层结构，每层只管一件事：① 样本量够不够（统计下界）② 参与率在同类里高不高 '
                '③ 读者是否真的在划线/写想法（互动证据）④ 受众面大小（只 ±5%，属商业价值而非可信度）。'
                '用「人均密度」而不是绝对量，因为绝对量与阅读人数的 log 相关高达 0.91~0.94，直接用等于把热度再奖励一次。')
    ws['A3'].font = Font(size=9, color='C00000', name='微软雅黑')
    ws.merge_cells('A3:Q3')
    ws.row_dimensions[3].height = 46

    def sub(r, text):
        ws.cell(row=r, column=1, value=text).font = SUB_FONT

    r = 5

    # ---------- 一、个案拆解 ----------
    _case = next((b for b in BOOKS if '棋经十三篇' in (b.get('title') or '')), None)
    sub(r, '一、逐项拆解示例：%s（%s 分是怎么算出来的）'
        % ((_case.get('title') if _case else '某本书'),
           ('%.1f' % _case['_vscore']) if _case else '—'))
    r += 1
    heads = ['#', '因子', '数值', '与同类比', '对分数的影响', '说明']
    for i, h in enumerate(heads, 1):
        cc = ws.cell(row=r, column=i, value=h)
        cc.fill = HEAD_FILL
        cc.font = HEAD_FONT
        cc.alignment = CENTER
        cc.border = BORDER
    ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=12)
    r += 1

    if _case:
        cv = vs_of(_case)
        _f = cv.get('fac_e') or 1.0
        if _f > 1.005:
            _fac_note = 'fac_e 高于类中位，最多加 6%%（这本 +%.0f%%）' % (100 * (_f - 1))
        elif _f < 0.995:
            _fac_note = 'fac_e 低于类中位，最多扣 12%%（这本 −%.0f%%）' % (100 * (1 - _f))
        else:
            _fac_note = 'fac_e 与类中位持平，不奖不罚'
        rows = [
            ('① 平台原评分', f'{_case["_rating"]:.1f} 分', '—', '起算点', '推荐值（newRating）÷10'),
            ('② 评分人数 / 阅读人数', f'{cv.get("rc"):,} 人 / {cv.get("rd"):,} 人',
             f'参与率 {cv.get("ratio", 0) * 100:.2f}%（同类 {cv.get("rbar", 0) * 100:.2f}%）',
             f'x_r = {cv.get("x_r"):.2f} → 参与度置信 {cv.get("conf_r"):.3f}',
             '参与率是同类均值的 0.74 倍；按平方根强度计入（rel_r = 0.797 → 0.892），只是在同类里偏"冷"，不算差'),
            ('③ 统计下界（Wilson 95%）', f'{cv.get("lcb", 0) * 100:.1f} 分',
             f'比原评分低 {_case["_rating"] - cv.get("lcb", 0) * 100:.1f} 分',
             f'conf_stat = {cv.get("conf_stat"):.3f}',
             '72 人评分的标准误约 ±4.2 分；人越多这一项越接近 1（30 万人评分 ≈ 0.999）'),
            ('④ 阅读人数（受众面）', f'{cv.get("rd"):,} 人',
             f'同类几何均值 {cv.get("rdbar"):,.0f} 人的 {cv.get("x_d"):.2f} 倍',
             f'heat = {cv.get("heat"):.3f}',
             '受众少只影响 −5%，不会像"人少=不可信"那样被压垮'),
            ('⑤ 互动证据（划线 / 想法）',
             ('想法 %d 条 / 最热划线 %d 人' % (int(cv.get('idea') or 0), int(cv.get('mark') or 0)))
             if cv.get('idea') is not None else '无互动数据',
             f'x_idea {cv.get("x_idea") or 0:.2f}、x_mark {cv.get("x_mark") or 0:.2f} → x_e {cv.get("x_e") or 0:.2f}',
             f'fac_e = {cv.get("fac_e") or 1:.3f}',
             '同类里它的想法密度是中位的 %.0f%%；%s。平台不公开全站划线总数，这里用 Top10 热门划线代理'
             % (100 * (cv.get('x_idea') or 1), _fac_note)),
            ('⑥ 合成置信度', f'{cv.get("conf"):.3f}', '—',
             f'{_case["_rating"]:.1f} × {cv.get("conf"):.3f} = **{cv.get("score"):.1f} 分**',
             '四层相乘后受封顶线限制（9.1 分以下按 √(评分÷9.1) 收紧），所以价值评分永远是平台评分的"保守下界"'),
            ('⑦ 类内位置', '比同类均值 %+.1f 分' % (_case.get('_clsrel') or 0),
             f'百分位 {_case.get("_clspct") or 0:.0f}%',
             '在同类 %d 本里排中游' % (cv.get('cls_n') or 0),
             '它的参与率确实低于同类均值，属于"书好但读者沉默"，不是"分数虚高"'),
        ]
        for i, (a, b_, c, d, e) in enumerate(rows, 1):
            put_row(ws, r, [i, a, b_, c, d, e])
            for j in (3, 4, 5):
                ws.cell(row=r, column=j).alignment = CENTER
            ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=12)
            ws.cell(row=r, column=6).alignment = Alignment(vertical='center', wrap_text=True)
            ws.row_dimensions[r].height = 30
            r += 1
    r += 1

    # ---------- 二、公式 ----------
    sub(r, '二、公式与参数（照着这张表就能复算）')
    r += 1
    heads = ['#', '项', '公式', '参数', '说明']
    for i, h in enumerate(heads, 1):
        cc = ws.cell(row=r, column=i, value=h)
        cc.fill = HEAD_FILL
        cc.font = HEAD_FONT
        cc.alignment = CENTER
        cc.border = BORDER
    ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=12)
    r += 1
    rows = [
        ['1', '统计可信度', 'conf_stat = WilsonLCB(raw/100, rc, z=1.645) ÷ (raw/100)',
         'z = 1.645（单侧 95%）', '唯一由样本量决定的折扣，自动按 1/√n 校准；rc=3→0.53、72→0.90、1000→0.97、30 万→1.00'],
        ['2', '类内参与率', 'x_r = (rc/rd) ÷ 本类 r̄；rel_r = [0.2+0.8·Hill(x_r)] ÷ 0.8',
         f'K={P.get("K_REL")}、n={P.get("N_HILL")}', 'Hill 曲线与锚点保留（x=1→0.8、0.7→0.6）；只以 rel_r^0.5 进入，因为在类内单用太陡'],
        ['3', '受众面', 'heat = 0.95 + 0.10·Hill(x_d)，x_d = rd ÷ 本类阅读人数几何均值',
         f'基准 {P.get("HEAT_BASE")}、振幅 ±{P.get("HEAT_RANGE") / 2 * 100:.0f}%', '人多最多 +5%，人少最多 −5%；受众面属于商业价值，不该当可信度'],
        ['4', '互动证据', 'x_e = x_mark^0.55 × x_idea^0.45；fac_e = clip{[0.2+0.8·Hill(x_e,K=0.50)]÷0.8}^0.40',
         f'K_E={P.get("K_E")}、W_E={P.get("W_E")}、下限 {P.get("E_FLOOR")}、上限 {P.get("E_CEIL")}',
         'x_mark = 最热划线人数/阅读人数 ÷ 本类中位数；x_idea = 想法数/阅读人数 ÷ 本类中位数；划线权重 0.55 > 评论 0.45'],
        ['5', '合成', 'conf = min(conf_stat × rel_r^0.5 × fac_e × heat, cap)',
         f'W_r = {P.get("W_R")}', '封顶线随分数上升：cap = min(1, √(评分÷9.1))，只有 9.1 分以上才可能 conf=1，以下按分数收紧'],
        ['6', '封顶线', 'cap = min(1.0, (原评分 ÷ 91)^0.5)；价值分 = 原评分 × conf',
         f'CAP_RAW={P.get("CAP_RAW")}、γ={P.get("CAP_GAMMA")}',
         '四层乘数上限叠乘 1.12×1.06×1.05=1.246，若封顶取常数 1 会让 36.9% 的书「价值分=平台评分」、模型退化；'
         '改成随分数收紧后，打满的只剩 29 本且全部 ≥9.1 分，8.5 分的书价值分约 82'],
    ]
    for i, (a, b_, c, d, e) in enumerate(rows, 1):
        put_row(ws, r, [a, b_, c, d, e])
        ws.cell(row=r, column=3).alignment = Alignment(vertical='center', wrap_text=True)
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=12)
        ws.cell(row=r, column=5).alignment = Alignment(vertical='center', wrap_text=True)
        ws.row_dimensions[r].height = 34
        r += 1
    r += 1

    # ---------- 三、全量 ----------
    sub(r, '三、全量明细（%d 本有评分的书，按 价值评分从高到低）' % n)
    r += 1
    heads = ['#', '书名', '二级分类', '标价(元)', '原评分', '评分人数', '阅读人数', '参与率',
             '统计下界', 'conf_stat', 'x_r', 'rel_r', 'heat', '互动 fac_e', '价值分',
             '类内相对', '类内百分位']
    head_row = r
    for i, h in enumerate(heads, 1):
        cc = ws.cell(row=r, column=i, value=h)
        cc.fill = HEAD_FILL
        cc.font = HEAD_FONT
        cc.alignment = CENTER
        cc.border = BORDER
    ws.row_dimensions[r].height = 30
    r += 1
    for i, b in enumerate(rated, 1):
        put_row(ws, r, [i, b.get('title') or '', b['_type2'], b['_price'], b['_rating'],
                        b['_rcnt'], b['_readcnt'],
                        _pct(b['_ratio'], 2) if b['_ratio'] else '—',
                        round((b['_vlcb'] or 0) * 100, 1),
                        round(b['_confstat'], 3) if b['_confstat'] else '—',
                        round(b['_xr'], 2) if b['_xr'] else '—',
                        round(b['_relr'], 3) if b['_relr'] else '—',
                        round(b['_heat'], 3) if b['_heat'] else '—',
                        round(b['_fac_e'], 3) if b.get('_fac_e') else '—',
                        b['_vscore'], b.get('_clsrel'), b.get('_clspct')],
                money_cols=(4,))
        for j in range(5, 18):
            ws.cell(row=r, column=j).alignment = Alignment(horizontal='center', vertical='center')
        r += 1
    ws.auto_filter.ref = 'A%d:Q%d' % (head_row, r - 1)

    r += 1
    for t in ['※ 三层「虚高」要分开看：',
              '   ① 单本可信度看「conf_stat」列：<0.7 说明这本书的分主要靠少数人投票；',
              '   ② 相对同类的位置看「类内相对」与「类内百分位」：同类型内比较，避开类型红利；',
              '   ③ 整类是否普遍虚高看「原评分 − 类均分」（见「分类基准与类型评分」的折扣率）。',
              '※ 为什么不能拿阅读人数当折扣：一本 3000 人读、72 人评的书，与一本 30 万人读、72 人评的书，'
              '评分的统计可信度是同一个数量级——差异在受众面，不在证据强度。',
              '※ fac_e 区间 0.88~1.06：互动最多加 6%、最多扣 12%，不允许"读者沉默"把一本好书压垮。']:
        cc = ws.cell(row=r, column=1, value=t)
        cc.font = Font(size=9, color='595959', name='微软雅黑')
        cc.alignment = Alignment(vertical='center', wrap_text=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=15)
        ws.row_dimensions[r].height = 22
        r += 1


# ==================== 分类基准与类型评分 ====================
# ==================== 分类基准与类型评分（按一级分类整理全部书） ====================
def build_type_sheet(wb):
    ws = wb.create_sheet('分类基准与类型评分')
    ws.sheet_view.showGridLines = False
    widths = [6, 40, 18, 20, 9, 12, 7, 9, 10, 10, 10, 9, 9, 9, 10, 12]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    c1 = VTYPE.get('c1') or {}
    g = VGLOBAL or {}

    ws['A1'] = '分类基准与类型评分：一级分类基准 + 书架∪书单全部书按类内价值分排序'
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:P1')
    ws.row_dimensions[1].height = 30
    ws['A2'] = (f'合并去重：{UNION_N} 本 ＝ 书架 {UNION_SHELF} 本 ＋ 公开书单里尚未加架 {UNION_OUT} 本；'
                f'其中有评分 {UNION_RATED} 本。类内基准仍用书架全量算出：'
                f'参与率 r̄ = {g.get("mean_r", 0) * 100:.3f}%、阅读人数几何均值 {g.get("geo_rd", 0):,.0f} 人'
                f'（算术 {g.get("mean_rd", 0):,.0f}）、平均原评分 {g.get("mean_raw", 0):.1f}；n = {g.get("n", 0)} 本')
    ws['A2'].font = Font(size=9, color='595959', name='微软雅黑')
    ws.merge_cells('A2:P2')
    ws.row_dimensions[2].height = 28
    ws['A3'] = ('读法：①「相对全局」两列 >1 表示该类整体比全书架更热/参与率更高，也就是用全书架统一均值时它会吃到类型红利；'
                '②「折扣率」= 1 − 平均价值评分 ÷ 平均原评分，越大说明该类评分被削得越狠；'
                '③ 第三节的「类内相对」= 该书价值分 − 本类均值，排序只看类内，跨类比较请回到「价值评分」。')
    ws['A3'].font = Font(size=9, color='C00000', name='微软雅黑')
    ws.merge_cells('A3:P3')
    ws.row_dimensions[3].height = 32

    def _avg_by(field):
        acc, cnt = {}, {}
        for b in BOOKS:
            val = b.get(field)
            if val is None or not b.get('_vscore'):
                continue
            k = b['_type1']
            acc[k] = acc.get(k, 0.0) + val
            cnt[k] = cnt.get(k, 0) + 1
        return {k: acc[k] / cnt[k] for k in acc}

    CS1, FE1 = _avg_by('_confstat'), _avg_by('_fac_e')
    U_CNT, U_RATED = {}, {}
    for r in UNION:
        U_CNT[r['type1']] = U_CNT.get(r['type1'], 0) + 1
        if r['vscore']:
            U_RATED[r['type1']] = U_RATED.get(r['type1'], 0) + 1

    r = 5
    ws.cell(row=r, column=1,
            value='一、一级分类基准（按 平均价值评分从高到低；基准=书架全量，样本=书架∪书单）').font = SUB_FONT
    r += 1
    heads = ['#', '一级分类', '合并样本', '书架子集', '有评分', '类内 r̄', '类内阅读\n人数(几何)',
             '相对全局\nr̄', '相对全局\n阅读人数', '平均原评分', '平均\n价值分', '折扣率',
             '统计折扣\n(均值)', '互动因子\n(均值)', '类内书单书']
    for i, h in enumerate(heads, 1):
        cc = ws.cell(row=r, column=i, value=h)
        cc.fill = HEAD_FILL
        cc.font = HEAD_FONT
        cc.alignment = CENTER
        cc.border = BORDER
    ws.row_dimensions[r].height = 32
    r += 1
    items = [(k, v) for k, v in c1.items() if (v.get('nRated') or 0) >= 10 and k != '未分类']
    items.sort(key=lambda kv: -kv[1]['meanScore'])
    for i, (k, v) in enumerate(items, 1):
        cs = CS1.get(k)
        fe = FE1.get(k)
        put_row(ws, r, [i, k, U_CNT.get(k, v.get('n', 0)), v.get('n', 0), U_RATED.get(k, 0),
                        _pct(v['rbar'], 3), int(v['rdbar']),
                        round(v['rel_r'], 2), round(v['rel_rd'], 2),
                        v['meanRaw'], v['meanScore'],
                        _pct(1 - v['meanScore'] / v['meanRaw'], 1) if v['meanRaw'] else '—',
                        round(cs, 3) if cs is not None else '',
                        round(fe, 3) if fe else '',
                        sum(1 for x in UNION if x['type1'] == k and not x['inShelf'])])
        for i2 in range(3, 16):
            ws.cell(row=r, column=i2).alignment = Alignment(horizontal='center', vertical='center')
        r += 1
    r += 1

    # ---------- 结论速读（不含版本对比） ----------
    ws.cell(row=r, column=1, value='二、结论速读').font = SUB_FONT
    r += 1
    xs = [(k, v) for k, v in c1.items() if (v.get('nRated') or 0) >= 10 and k != '未分类']
    hot = sorted(xs, key=lambda kv: -kv[1]['meanScore'])[:4]
    low = sorted(xs, key=lambda kv: kv[1]['meanScore'])[:4]
    infl = sorted(xs, key=lambda kv: kv[1]['inflate'], reverse=True)[:4]
    lines = [
        '① 类内价值评分最高的类型（已按同类基准相对化）：' +
        '、'.join('%s %.1f' % (k, v['meanScore']) for k, v in hot),
        '② 类内价值评分最低的类型：' +
        '、'.join('%s %.1f' % (k, v['meanScore']) for k, v in low),
        '③ 折扣率最高的类型（平台评分与"统计下界"差距最大，最该警惕整体虚高）：' +
        '、'.join('%s %s' % (k, _pct(1 - v['meanScore'] / v['meanRaw'], 1)) for k, v in infl),
        '④ 小说 / 传记到底吃什么红利（这些类型的参与率与阅读人数是全书架的数倍）：',
    ]
    for k in ('精品小说', '人物传记', '文学', '历史', '经济理财', '教育学习'):
        v = c1.get(k)
        if v:
            lines.append('     %s：参与率 %s（%.2f 倍全局）、阅读人数几何 %s 人（%.2f 倍）、折扣 %s、互动因子 %.3f'
                         % (k, _pct(v['rbar'], 3), v['rel_r'], format(int(v['rdbar']), ','), v['rel_rd'],
                            _pct(1 - v['meanScore'] / v['meanRaw'], 1), v.get('meanFacE') or 1.0))
    lines.append('⑤ 同一本书在不同类型里的分数不可直接比：类型基准不同，一个 70 分的古籍可能比 75 分的畅销小说更"值得"，'
                 '所以第三节按类内排序，把同类的书放在一起看。')
    for t in lines:
        cc = ws.cell(row=r, column=1, value=t)
        cc.font = Font(size=9.5, name='微软雅黑', color='333333')
        cc.alignment = Alignment(vertical='center', wrap_text=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=16)
        ws.row_dimensions[r].height = 20
        r += 1
    r += 1

    # ---------- 三、按一级分类逐本（书架 ∪ 书单） ----------
    ws.cell(row=r, column=1,
            value='三、按一级分类逐本：书架∪书单合并去重 %d 本，类内按价值分从高到低' % UNION_N).font = SUB_FONT
    r += 1
    ws.cell(row=r, column=1,
            value='（每类先给基准与统计，再列书；「在书架」=否 的书来自公开书单，价值分用同一套公式现算，'
                  '基准取该类书架基准。无评分的书排在类末。）').font = Font(size=9, color='808080', name='微软雅黑')
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=16)
    r += 1

    heads3 = ['#', '书名', '作者', '二级分类', '在书架', '权益/状态', '标价(元)', '原评分',
              '评分人数', '阅读人数', '参与率', '价值分', '类内相对', '类内百分位', '所属书单', 'bookId']
    by_t1 = {}
    for x in UNION:
        by_t1.setdefault(x['type1'], []).append(x)

    order_t1 = [k for k, _v in items] + [k for k in by_t1 if k not in [x for x, _ in items]]
    for t1 in order_t1:
        rows = by_t1.get(t1) or []
        if not rows:
            continue
        rows = sorted(rows, key=lambda x: (-(x['vscore'] or -1), -(x.get('clsRel') or -999)))
        v = c1.get(t1) or {}
        rated_n = sum(1 for x in rows if x['vscore'])
        head = ('%s　合并 %d 本（书架 %d / 书单未加架 %d，有评分 %d）'
                % (t1, len(rows), sum(1 for x in rows if x['inShelf']),
                   sum(1 for x in rows if not x['inShelf']), rated_n))
        cc = ws.cell(row=r, column=1, value=head)
        cc.font = Font(bold=True, size=11, color='FFFFFF', name='微软雅黑')
        cc.fill = PatternFill('solid', fgColor='1F3864')
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=16)
        ws.row_dimensions[r].height = 24
        r += 1
        if v:
            cc = ws.cell(row=r, column=1,
                         value='   基准：类内 r̄ %s（全局的 %.2f 倍）、阅读人数几何 %s 人（%.2f 倍）、'
                               '平均原评分 %.1f → 平均价值分 %.1f（折扣 %s）、互动因子均值 %.3f'
                               % (_pct(v['rbar'], 3), v['rel_r'], format(int(v['rdbar']), ','), v['rel_rd'],
                                  v['meanRaw'], v['meanScore'],
                                  _pct(1 - v['meanScore'] / v['meanRaw'], 1), v.get('meanFacE') or 1.0))
            cc.font = Font(size=9, color='595959', name='微软雅黑')
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=16)
            r += 1
        for i, h in enumerate(heads3, 1):
            c = ws.cell(row=r, column=i, value=h)
            c.fill = HEAD_FILL
            c.font = HEAD_FONT
            c.alignment = CENTER
            c.border = BORDER
        r += 1
        for i, x in enumerate(rows, 1):
            b = x['b']
            cat = b['_cat'] if b else ('书单·未加架')
            put_row(ws, r, [i, x['title'], x['author'], x['type2'],
                            '是' if x['inShelf'] else '否', cat, x['price'] or '—',
                            round(x['rating'], 1) if x['rating'] else '—',
                            x['rc'] or '—', x['rd'] or '—',
                            _pct(x.get('ratio') or 0, 2) if x.get('ratio') else '—',
                            x['vscore'] if x['vscore'] else '无评分',
                            x.get('clsRel') if x['vscore'] else '—',
                            x.get('clsPct') if x['vscore'] else '—',
                            '、'.join(x['bl']) if x['bl'] else '', x['bid']],
                    money_cols=(7,))
            for i2 in (5, 6, 7, 8, 9, 10, 11, 12, 13, 14):
                ws.cell(row=r, column=i2).alignment = Alignment(horizontal='center', vertical='center')
            if not x['inShelf']:
                ws.cell(row=r, column=5).fill = PatternFill('solid', fgColor='FFF2CC')
            if x['sold']:
                ws.cell(row=r, column=6).fill = PatternFill('solid', fgColor='F2F2F2')
            r += 1
        r += 1

    r += 1
    for t in ['※ 类内基准的分档处理（避免小类被一两本书带偏）：',
              '   · 一级类有评分书 ≥10 本 → 用自己的均值；<10 本 → 沿用全书架基准（不作类型调整）。',
              '   · 二级类 ≥8 本 → 用自己均值并向一级类轻度收缩（权重 n÷(n+6)，n=20 时保留 77%）；<8 本 → 继承一级类。',
              '※ 阅读人数用几何均值（= 平均数量级），不用算术均值：小说类算术 52.0 万 vs 几何 13.5 万，'
              '算术均值被个位数畅销书拉到 3.7 倍，用它当基准会让该类 9 成书被判"低于平均"。',
              '※ 「类内相对」在本 sheet 里按「书架∪书单」重新算（含书单里未加架的书）；'
              '「价值评分」sheet 的类内相对只按书架算，两者基数不同，看趋势即可，不要逐本对数。']:
        cc = ws.cell(row=r, column=1, value=t)
        cc.font = Font(size=9, color='595959', name='微软雅黑')
        cc.alignment = Alignment(vertical='center', wrap_text=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=16)
        ws.row_dimensions[r].height = 22
        r += 1



# ==================== 书架清理候选 ====================
def _prog_state(p):
    if p is None:
        return '无记录'
    if p >= 100:
        return '已读完'
    if p >= 50:
        return '读过大半'
    if p >= 10:
        return '读了一部分'
    if p >= 1:
        return '刚翻开'
    return '从未打开'


def build_cleanup_sheet(wb):
    """书架清理候选：只列「价值评分 < 40」的书（阈值以下才出现在本 sheet）。"""
    TH = 40.0
    judged = []
    for b in BOOKS:
        v = b['_vscore'] or 0
        if not v or v >= TH:
            continue
        st = _prog_state(b['_prog'])
        owned = b['_cat'] == '永久可读'
        if st in ('已读完', '读过大半'):
            act, why = '暂缓（已投入时间）', '已%s，价值分再低也已经读过了' % st
        elif st == '读了一部分':
            act, why = '暂缓（在读）', '正在读，进度 %d%%' % (b['_prog'] or 0)
        elif owned:
            act, why = '可清理（已买断）', '价值分 %.1f（原评分 %.1f，同类%+.1f），%s；已买断，删了权益仍在'
            why = why % (v, b['_rating'] or 0, b.get('_clsrel') or 0, st)
        else:
            act, why = '优先清理（未买断）', '价值分 %.1f（原评分 %.1f，同类%+.1f），%s；未买断，删了零损失'
            why = why % (v, b['_rating'] or 0, b.get('_clsrel') or 0, st)
        judged.append((act, why, st, b))

    order = {'优先清理（未买断）': 0, '可清理（已买断）': 1, '暂缓（在读）': 2, '暂缓（已投入时间）': 3}
    judged.sort(key=lambda t: (order[t[0]], t[3]['_vscore'] or 0, -(t[3]['_price'] or 0)))
    cnt = Counter(t[0] for t in judged)
    n_go = cnt['优先清理（未买断）'] + cnt['可清理（已买断）']

    ws = wb.create_sheet('书架清理候选(%d)' % len(judged))
    ws.sheet_view.showGridLines = False
    widths = [5, 36, 20, 10, 8, 9, 8, 9, 10, 9, 9, 8, 46]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws['A1'] = ('书架清理候选：价值评分 < %.0f 分的 %d 本（全书架 %d 本，有评分 %d 本）'
                % (TH, len(judged), len(BOOKS),
                   sum(1 for b in BOOKS if b['_rating'])))
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:M1')
    ws.row_dimensions[1].height = 30
    ws['A2'] = ('筛选规则：只看 价值评分 < %.0f 分（不再按 55 / 62 分档，也不再把"折扣大"当理由）。'
                '再按权益与进度分层：① 未买断 → 优先清理（删了没有任何损失）；'
                '② 已买断 → 可清理（权益还在，随时能从"已购"加回）；③ 在读 / 读过 → 暂缓（列出来只是让你看到它确实低分）。' % TH)
    ws['A2'].font = Font(size=9, color='595959', name='微软雅黑')
    ws.merge_cells('A2:M2')
    ws.row_dimensions[2].height = 28
    ws['A3'] = ('⚠ 本表只是筛选结果，我不会、也无法替你删除书架上的书。删书需要在微信读书 App 里手动操作。'
                '已买断（永久可读）的书移出书架不会丢失购买权益，随时可以从"已购"重新加回；'
                '未买断的书移出书架等于放弃这本书的收藏，但本来也没花钱。')
    ws['A3'].font = Font(size=9, color='C00000', name='微软雅黑')
    ws.merge_cells('A3:M3')
    ws.row_dimensions[3].height = 40

    ws['A5'] = ('· 优先清理（未买断）%d 本　· 可清理（已买断）%d 本　· 暂缓 %d 本　→ 建议动手 %d 本\n'
                '· 这 %d 本里未买断的标价合计 %.2f 元（本来也不花你的钱），已买断的标价合计 %.2f 元（已付出的沉没成本）'
                % (cnt['优先清理（未买断）'], cnt['可清理（已买断）'],
                   cnt['暂缓（在读）'] + cnt['暂缓（已投入时间）'], n_go,
                   len(judged),
                   sum(b['_price'] or 0 for a, _w, _s, b in judged if a.startswith('优先清理')),
                   sum(b['_price'] or 0 for a, _w, _s, b in judged if a.startswith('可清理'))))
    ws['A5'].alignment = Alignment(vertical='center', wrap_text=True)
    ws['A5'].font = Font(size=9.5, name='微软雅黑', color='1F3864')
    ws.merge_cells('A5:M7')

    r = 9
    heads = ['#', '书名', '分类', '权益', '标价(元)', '原评分', '价值分', '类内相对',
             '阅读进度', '阅读状态', '评分人数', '点评数', '建议与依据']
    for i, h in enumerate(heads, 1):
        cc = ws.cell(row=r, column=i, value=h)
        cc.fill = HEAD_FILL
        cc.font = HEAD_FONT
        cc.alignment = CENTER
        cc.border = BORDER
    ws.row_dimensions[r].height = 28
    r += 1
    head_row = r

    fills = {'优先清理（未买断）': 'FCE4E4', '可清理（已买断）': 'FFF2CC',
             '暂缓（在读）': 'EDEDED', '暂缓（已投入时间）': 'EDEDED'}
    for i, (act, why, st, b) in enumerate(judged, 1):
        put_row(ws, r, [i, b.get('title') or '', b.get('_type2') or '', b['_cat'], b['_price'],
                        b['_rating'] or '—', b['_vscore'] or '—', b.get('_clsrel'),
                        ('%d%%' % b['_prog']) if b['_prog'] is not None else '—', st,
                        b['_rcnt'] or '—', b['_review'] or 0,
                        '%s ｜ %s' % (act, why)],
                money_cols=(5,))
        for j in range(4, 13):
            ws.cell(row=r, column=j).alignment = CENTER
        ws.cell(row=r, column=13).alignment = Alignment(vertical='center', wrap_text=True)
        f = fills.get(act)
        if f:
            ws.cell(row=r, column=1).fill = PatternFill('solid', fgColor=f)
            ws.cell(row=r, column=2).fill = PatternFill('solid', fgColor=f)
        r += 1
    ws.auto_filter.ref = 'A%d:M%d' % (head_row, r - 1)

    r += 1
    for t in ['※ 为什么阈值直接定在 40 分：价值评分全书架均值约 66 分、中位 68 分，<40 分已经是"同类里明显偏低 + 证据不足"的那一档，'
              '再往上放（55 分）会混进不少其实还行的书，清单就不够干脆。',
              '※ 书的价值是"对读者的价值"：一本 38 分但你已读过的书，对你价值不低，所以已读/在读的不建议动。',
              '※ 清理前建议：先清「未买断 + 从未打开」这 %d 本，删了不损失任何权益；'
              '已买断的先别急，微信读书书架没有容量上限，清理只影响你自己找书的速度。'
              % cnt['优先清理（未买断）'],
              '※ 若某本书你记得"其实想读"，在 App 里搜索书名重新加入即可（已买断的不会再收费）。']:
        cc = ws.cell(row=r, column=1, value=t)
        cc.font = Font(size=9, color='595959', name='微软雅黑')
        cc.alignment = Alignment(vertical='center', wrap_text=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=13)
        ws.row_dimensions[r].height = 22
        r += 1


def build_hot_sheet(wb):
    """热书排名：书架 ∪ 书单 合并去重后按热度取前 150。"""
    TOPN = 150
    ws = wb.create_sheet('热书排名(Top%d)' % TOPN)
    ws.sheet_view.showGridLines = False
    widths = [5, 40, 18, 13, 9, 11, 11, 11, 11, 10, 9, 8, 18, 11]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ranked = sorted([x for x in UNION if x['heat'] and x['heatDim'] >= 1],
                    key=lambda x: (-x['heat'], -(x['rd'] or 0)))
    top = ranked[:TOPN]

    ws['A1'] = '热书排名：书架 ∪ 书单合并去重后的热度 Top %d' % TOPN
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:N1')
    ws.row_dimensions[1].height = 30
    ws['A2'] = ('热度分 = 100 ×［0.40×阅读人数百分位 + 0.30×评分人数百分位 + 0.20×最热划线人次百分位 '
                '+ 0.10×想法数百分位］，四项都先取 log10(1+x)。'
                '缺项的按可用项权重归一化，所以「数据维度」列要一起看。')
    ws['A2'].font = Font(size=9, color='C00000', name='微软雅黑')
    ws.merge_cells('A2:N2')
    ws.row_dimensions[2].height = 30
    ws['A3'] = ('样本：书架 %d 本 ＋ 公开书单里未加架 %d 本 ＝ 合并去重 %d 本；'
                '本表取热度前 %d。在架书 4 项全覆盖（数据维度 4）；书单里未加架的书已用登录态补抓'
                '阅读人数 / 想法数 / 热门划线，绝大多数也是 4 项，个别平台没给热门划线的会是 3 项。'
                % (UNION_SHELF, UNION_OUT, UNION_N, len(top)))
    ws['A3'].font = Font(size=9, color='595959', name='微软雅黑')
    ws.merge_cells('A3:N3')
    ws.row_dimensions[3].height = 30

    r = 5
    heads = ['#', '书名', '作者', '一级分类', '热度分', '数据维度', '阅读人数', '评分人数',
             '最热划线人次', '想法数', '价值分', '在书架', '所属书单', '标价(元)']
    for i, h in enumerate(heads, 1):
        c = ws.cell(row=r, column=i, value=h)
        c.fill = HEAD_FILL
        c.font = HEAD_FONT
        c.alignment = CENTER
        c.border = BORDER
    ws.row_dimensions[r].height = 30
    head_row = r
    r += 1
    for i, x in enumerate(top, 1):
        put_row(ws, r, [i, x['title'], x['author'], x['type1'], x['heat'],
                        '%d 项' % x['heatDim'], x['rd'] or 0, x['rc'] or 0,
                        int(x['mark']) if x['mark'] else '—',
                        int(x['idea']) if x['idea'] is not None else '—',
                        x['vscore'] if x['vscore'] else '—',
                        '是' if x['inShelf'] else '否',
                        '、'.join(x['bl']) if x['bl'] else '', x['price'] or '—'],
                money_cols=(14,))
        for i2 in (5, 6, 7, 8, 9, 10, 11, 12, 14):
            ws.cell(row=r, column=i2).alignment = Alignment(horizontal='center', vertical='center')
        if not x['inShelf']:
            ws.cell(row=r, column=12).fill = PatternFill('solid', fgColor='FFF2CC')
        r += 1
    ws.auto_filter.ref = 'A%d:N%d' % (head_row, r - 1)

    r += 1
    hot_cnt = sum(1 for x in top if not x['inShelf'])
    for t in ['※ 这张表回答的是「多少人在真的读它」：读的人越多、打分的人越多、划同一个句子的人越多、'
              '写想法的人越多，热度分越高。它和价值评分是两件事——',
              '   · 热度高 + 价值分高 = 又叫好又叫座，值得优先买；',
              '   · 热度高 + 价值分低 = 畅销型（营销/时点效应），不必因为热就买；',
              '   · 热度低 + 价值分高 = 小众好书，适合按需买，不用赶。',
              '※ 前 %d 本里有 %d 本来自你的公开书单（尚未加架），标黄表示"不在书架"：'
              '如果它们正好在你想买的书单里，可以优先考虑。' % (len(top), hot_cnt),
              '※ 关于「多个朋友都看过」：微信读书的"朋友在读"依赖微信好友关系，只存在于 App 端，'
              '网页版没有任何 friend / social 接口（/web/book/friend* 全 404），因此拿不到"好友"这一维，'
              '本表用的是全体读者的公开痕迹。']:
        cc = ws.cell(row=r, column=1, value=t)
        cc.font = Font(size=9, color='595959', name='微软雅黑')
        cc.alignment = Alignment(vertical='center', wrap_text=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=14)
        ws.row_dimensions[r].height = 24
        r += 1



# ==================== 书单（公开书单，多单汇总） ====================
def _bl_rows(rec, EXTRA):
    """把书单记录展开成行：在架的直接用书架数据，不在架的用补抓的详情。"""
    rows = []
    for b in rec['books']:
        bid = b['bookId']
        sh = BY_ID.get(bid)
        if sh:
            rows.append({'inShelf': True, 'bid': bid, 'b': sh, 'order': b.get('order'),
                         'title': sh.get('title') or b['title'],
                         'author': sh.get('author') or ''})
        else:
            e = EXTRA.get(bid) or {}
            sc, cf, rc = score_ext(e)
            rows.append({'inShelf': False, 'bid': bid, 'b': None, 'order': b.get('order'),
                         'title': e.get('title') or b['title'],
                         'author': e.get('author') or b.get('author') or '',
                         'price': round((e.get('centPrice') or 0) / 100.0, 2),
                         'rating': round((e.get('newRating') or 0) / 10.0, 1),
                         'rc': rc, 'conf': cf, 'vscore': sc,
                         'sold': 1 if e.get('soldout') in (1, True) else 0})
    return rows


def build_booklist_sheets(wb):
    try:
        BL = json.load(open('_booklists.json', encoding='utf-8'))
    except Exception:
        return
    try:
        EXTRA = json.load(open('_bookdetail_extra.json', encoding='utf-8'))
    except Exception:
        EXTRA = {}

    recs = sorted(BL.values(), key=lambda x: -(x.get('bookCount') or len(x['books'])))
    TITLE_OF = {}
    for rec in recs:
        for b in rec['books']:
            TITLE_OF.setdefault(b['bookId'], b.get('title') or '')
    pack = {rec['booklistId']: _bl_rows(rec, EXTRA) for rec in recs}

    def stat(rec):
        rows = pack[rec['booklistId']]
        ins = [x for x in rows if x['inShelf']]
        own = [x for x in ins if x['b']['_cat'] == '永久可读']
        mc = [x for x in ins if x['b']['_cat'] == '付费会员卡可读']
        tc = [x for x in ins if x['b']['_cat'] == '体验卡可读']
        sold = [x for x in rows if (x['b']['_sold'] if x['inShelf'] else x['sold'])]
        need = [x for x in ins if x['b']['_cat'] in ('付费会员卡可读', '体验卡可读') and not x['b']['_sold']]
        pool = [x['b'] for x in need if 0 < x['b']['_price'] <= COIN]
        pk, sp, _v = knap_lex(pool, COIN, 'rating')
        vs = [x['b']['_vscore'] for x in need if x['b']['_vscore']]
        return dict(n=len(rows), ins=len(ins), out=len(rows) - len(ins), own=len(own),
                    mc=len(mc), tc=len(tc), sold=len(sold), need=len(need),
                    money=round(sum(x['b']['_price'] for x in need), 2),
                    buy_n=len(pk), buy_money=round(sp, 2),
                    avg_v=round(sum(vs) / len(vs), 1) if vs else 0)

    # ---------------- 书单总览 ----------------
    ws = wb.create_sheet('书单总览(%d)' % len(recs))
    ws.sheet_view.showGridLines = False
    widths = [5, 16, 8, 10, 11, 12, 12, 11, 9, 14, 15, 12, 34]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    tot = sum(len(pack[r['booklistId']]) for r in recs)
    uniq = set()
    for r in recs:
        uniq.update(x['bid'] for x in pack[r['booklistId']])
    ws['A1'] = '书单汇总：%d 个书单 / %d 条归属 / %d 本去重' % (len(recs), tot, len(uniq))
    ws['A1'].font = TITLE_FONT
    ws.merge_cells('A1:M1')
    ws.row_dimensions[1].height = 30
    ws['A2'] = ('书单页 https://weread.qq.com/misc/booklist/<id> 服务端渲染，免登录即可取全部书目；'
                '不在书架的书用登录态 /web/book/info 补价格与推荐值。'
                '　本次刷新（%s）：%d 个书单 / %d 条归属 / %d 本，书单之间无重复收录。'
                % (recs[0].get('fetchedAt') if recs else '', len(recs), tot, len(uniq)))
    ws['A2'].font = Font(size=9, color='595959', name='微软雅黑')
    ws.merge_cells('A2:M2')

    r = 4
    ws.cell(row=r, column=1, value='一、各书单统计（按本数降序）').font = SUB_FONT
    r += 1
    heads = ['#', '书单', '本数', '在书架', '未加书架', '已永久可读', '会员卡可读', '体验卡可读',
             '已下架', '在架待付费标价合计(元)', '余额内可买(本/元)', '待付费书平均价值评分', '书单 ID']
    for i, h in enumerate(heads, 1):
        c = ws.cell(row=r, column=i, value=h)
        c.fill = HEAD_FILL
        c.font = HEAD_FONT
        c.alignment = CENTER
        c.border = BORDER
    ws.row_dimensions[r].height = 30
    r += 1
    head_row = r
    for i, rec in enumerate(recs, 1):
        st = stat(rec)
        put_row(ws, r, [i, rec.get('name') or '-', st['n'], st['ins'], st['out'], st['own'],
                        st['mc'], st['tc'], st['sold'], st['money'],
                        '%d 本 / %.2f' % (st['buy_n'], st['buy_money']), st['avg_v'],
                        rec['booklistId']], money_cols=(10,))
        for i2 in (3, 4, 5, 6, 7, 8, 9, 12):
            ws.cell(row=r, column=i2).alignment = CENTER
        r += 1
    put_row(ws, r, ['合计', '%d 个书单' % len(recs), tot, '', '', '', '', '', '', '', '', '', ''])
    for i2 in (3,):
        ws.cell(row=r, column=i2).alignment = CENTER
    r += 2
    ws.cell(row=r, column=1,
            value='※ 会员卡可读 = 需要付费会员卡才能读全文（会员卡 19 元/月须另行微信支付，书币/赠币不能抵扣）；'
                  '未加书架 = 书单里有、书架里没有，需要单独买断或先加书架再靠体验卡读。'
            ).font = Font(size=9, color='595959', name='微软雅黑')
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=13)
    r += 2

    # 二、全部书单合并后最值得先买
    pend = []
    for rec in recs:
        for x in pack[rec['booklistId']]:
            if x['inShelf']:
                b = x['b']
                if (b['_cat'] in ('付费会员卡可读', '体验卡可读') and not b['_sold']
                        and 0 < b['_price'] <= COIN and b['_pstat'] != '已读完'):
                    pend.append((b, rec.get('name')))
    seen2, cand = set(), []
    for b, nm in pend:
        if b['bookId'] in seen2:
            continue
        seen2.add(b['bookId'])
        cand.append((b, nm))
    cand.sort(key=lambda t: -(t[0]['_vscore'] or 0))
    ws.cell(row=r, column=1,
            value='二、全部书单合并后「余额内最值得先买」Top 20（按价值评分；已去重、已剔下架、已剔单价超余额、已剔你读完的）'
            ).font = SUB_FONT
    r += 1
    for i, h in enumerate(['#', '书名', '所属书单', '权益类型', '标价(元)', '评分', '置信度',
                           '价值评分', '每元价值', '阅读进度'], 1):
        c = ws.cell(row=r, column=i, value=h)
        c.fill = HEAD_FILL
        c.font = HEAD_FONT
        c.alignment = CENTER
        c.border = BORDER
    r += 1
    for i, (b, nm) in enumerate(cand[:20], 1):
        per = round(b['_vscore'] / b['_price'], 1) if b['_price'] else 0
        prog = '-' if b['_prog'] is None else '%d%%' % b['_prog']
        put_row(ws, r, [i, b.get('title') or '', nm, b['_cat'], b['_price'], b['_rating'],
                        b['_conf'], b['_vscore'], per, prog], money_cols=(5,), cat=b['_cat'])
        for i2 in (6, 7, 8, 9, 10):
            ws.cell(row=r, column=i2).alignment = CENTER
        r += 1
    r += 2

    # 四、链接清单 + 为何不能枚举
    ws.cell(row=r, column=1, value='三、已收录的书单链接（补录：把新链接加进 booklist_urls.txt 重跑即可）').font = SUB_FONT
    r += 1
    for i, h in enumerate(['#', '书单', '本数', '链接'], 1):
        c = ws.cell(row=r, column=i, value=h)
        c.fill = HEAD_FILL
        c.font = HEAD_FONT
        c.alignment = CENTER
        c.border = BORDER
    r += 1
    for i, rec in enumerate(recs, 1):
        put_row(ws, r, [i, rec.get('name') or '-', len(pack[rec['booklistId']]), rec.get('url') or ''])
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=13)
        ws.cell(row=r, column=2).alignment = CENTER
        ws.cell(row=r, column=3).alignment = CENTER
        r += 1
    r += 1
    for t in ['※ 书单 ID 形如 00000000_XXXXXXXXXX，后 10 位是服务端生成的乱序串，不是自增序号：',
              '  · 相邻 ID（…oulc / …oula / …ould）逐个试过，全部返回「书单不存在」；',
              '  · 按 base62 解码后与创建时间（updateTime）完全不对应，说明它是哈希/随机而非时间戳；',
              '  · App 侧接口 i.weread.qq.com 用网页登录态调用一律 -2012 登录超时，网关（/api/agent/gateway）的 18 个接口里也没有书单列表。',
              '  结论：剩下的书单无法由已知链接推测，只能你在 App 里分享一次、把链接发我。']:
        cc_ = ws.cell(row=r, column=1, value=t)
        cc_.font = Font(size=9, color='595959', name='微软雅黑')
        cc_.alignment = Alignment(vertical='center', wrap_text=True)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=13)
        r += 1

    # ---------------- 书单明细 ----------------
    ws2 = wb.create_sheet('书单明细(%d)' % tot)
    ws2.sheet_view.showGridLines = False
    w2 = [5, 14, 5, 40, 18, 9, 14, 9, 9, 8, 10, 9, 10, 10, 8, 10, 12]
    for i, w in enumerate(w2, 1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    ws2['A1'] = '书单逐本明细：%d 条 / %d 本（每个书单内无重复，书单之间也无重复）' % (tot, len(uniq))
    ws2['A1'].font = TITLE_FONT
    ws2.merge_cells('A1:Q1')
    ws2.row_dimensions[1].height = 30
    ws2['A2'] = '按书单、书单内原始顺序排列；「在书架」= 书架里已有，「权益类型」只对在架书有效。'
    ws2['A2'].font = Font(size=9, color='595959', name='微软雅黑')
    ws2.merge_cells('A2:Q2')
    heads2 = ['书单', '书单内序号', '书名', '作者', '在书架', '权益类型', '阅读进度', '标价(元)',
              '评分', '参与率', '置信度', '价值评分', '下架', 'bookId', '阅读状态', '阅读时长(h)', '最近阅读']
    r2 = 4
    for i, h in enumerate(heads2, 1):
        c = ws2.cell(row=r2, column=i, value=h)
        c.fill = HEAD_FILL
        c.font = HEAD_FONT
        c.alignment = CENTER
        c.border = BORDER
    ws2.row_dimensions[r2].height = 26
    ws2.freeze_panes = ws2.cell(row=r2 + 1, column=1)
    r2 += 1
    for rec in recs:
        for x in pack[rec['booklistId']]:
            if x['inShelf']:
                b = x['b']
                prog = '-' if b['_prog'] is None else '%d%%' % b['_prog']
                put_row(ws2, r2, [rec.get('name'), x.get('order') or '', x['title'],
                                  b.get('author') or '', '是', b['_cat'], prog, b['_price'],
                                  b['_rating'] or '', f'{(b["_ratio"] or 0) * 100:.2f}%' if b['_ratio'] else '—',
                                  b['_conf'], b['_vscore'], '已下架' if b['_sold'] else '',
                                  x['bid'], b['_pstat'],
                                  b.get('_rhours') or '', b.get('_recent') or ''], money_cols=(8,), cat=b['_cat'])
                st = ws2.cell(row=r2, column=7)
                st.alignment = CENTER
                st2 = ws2.cell(row=r2, column=15)
                st2.alignment = CENTER
            else:
                e = EXTRA.get(x['bid']) or {}
                put_row(ws2, r2, [rec.get('name'), x.get('order') or '', x['title'], x.get('author') or '', '否',
                                  '未加书架', '', x.get('price') or 0,
                                  x.get('rating') or '', '—', x.get('conf') or '',
                                  round(x.get('vscore') or 0, 1) if x.get('vscore') else '',
                                  '已下架' if x.get('sold') else '', x['bid'], '', '', ''],
                        money_cols=(8,))
            r2 += 1
    ws2.auto_filter.ref = 'A%d:Q%d' % (4, r2 - 1)


def main():
    wb = Workbook()
    wb.remove(wb.active)

    build_overview(wb, WALLET)
    build_plan_sheet(wb)
    if LOW_MODE:
        report_low.build_lowplan_sheet(wb)
    build_list_sheet(wb, '待付费书单(%d)' % len(NEED), NEED, with_cat=True)
    build_list_sheet(wb, '永久可读(%d)' % agg['永久可读']['n'],
                     [b for b in BOOKS if b['_cat'] == '永久可读'], with_cat=False)
    build_list_sheet(wb, '免费(%d)' % agg['免费']['n'],
                     [b for b in BOOKS if b['_cat'] == '免费'], with_cat=False)
    build_list_sheet(wb, '已读完未买断(%d)' % len(NEED_DONE), NEED_DONE, with_cat=True,
                     sort_key=lambda b: -b['_price'])
    build_list_sheet(wb, '在读(%d)' % len(NEED_READING), NEED_READING, with_cat=True,
                     sort_key=lambda b: (-(b['_prog'] or 0), -b['_price']))
    build_value_sheet_v4(wb)
    build_type_sheet(wb)
    build_hot_sheet(wb)
    build_cleanup_sheet(wb)
    report_low.build_lowprice_sheet(wb)
    build_booklist_sheets(wb)
    build_archive_sheet(wb)
    build_list_sheet(wb, '全量书籍(%d)' % len(BOOKS), BOOKS, with_cat=True)

    cands = ['微信读书_权益分类与购书方案.xlsx',
             '微信读书_权益分类与购书方案_new.xlsx',
             '微信读书_权益分类与付费统计.xlsx']
    out, err = None, None
    for c in cands:
        try:
            wb.save(c)
            out = c
            break
        except PermissionError as e:
            err = e
            continue
    if not out:
        raise SystemExit('所有目标文件都被占用，请关闭 Excel/预览后重试：%s' % err)
    print('saved:', out)
    print('四类:', {c: agg[c]['n'] for c in CAT_ORDER})
    print('需付费: %d 本, 标价合计 %.2f 元' % (len(NEED), NEED_MONEY))
    print('全量: %d 本, 标价合计 %.2f 元' % (TOTAL_N, TOTAL_MONEY))
    print('sheets:', wb.sheetnames)

    from docs_md import build_docs_md
    build_docs_md(globals())


if __name__ == '__main__':
    main()
