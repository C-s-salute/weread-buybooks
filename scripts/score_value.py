# -*- coding: utf-8 -*-
"""微信读书「价值评分 v5」= v4 三层 + 新增「互动证据」层（划线 / 想法）。

v5 相对 v4 的唯一增量：
  ④ 互动证据 fac_e —— 读者愿不愿意动笔，是比"投一票"更硬的质量证据
     · 划线密度 d_mark = 热门划线 Top10 累计人次 ÷ 阅读人数
     · 想法密度 d_idea = 想法数 ÷ 阅读人数
       （两者跨 3 个数量级且与阅读人数 log 相关 0.91/0.94，必须换成"人均"才不重复奖励热度）
     · 各自按类内几何均值相对化 → x_mark / x_idea
     · 合成 x_e = x_mark^0.6 × x_idea^0.4   （划线权重高于评论，用户指定）
     · conf_e = FLOOR + (CEIL-FLOOR)·Hill(x_e)，同一条 Hill 曲线（K=0.7 / n=3.0801）
     · fac_e  = clip( (conf_e/AT_MEAN)^W_E , E_FLOOR, E_CEIL )，W_E=0.70 > W_R=0.50

  score = 原评分 × min( conf_stat × rel_r^0.5 × fac_e × heat , cap )
  cap   = min(1.0, (原评分 ÷ 91)^0.5)   # 只有 9.1 分以上才允许置信度打满

  为什么 cap 不是常数 1：四个乘数的上限叠乘 = 1.12×1.06×1.05 = 1.246，足以把
  任何 conf_stat 推到 1，结果 36.9% 的书「价值分 = 原评分」，模型在它们身上
  完全退化。改成随原始分上升的封顶后，conf=1 只出现在 9.1 分以上的书里。
"""
import json
import math
import collections
import os

try:
    import wconfig   # noqa: F401  统一工作目录（设置了 WEREAD_WORK 时切 cwd）
except ImportError:
    pass

META_NEW = '_bookmeta_v5.json'
META_OLD = '_bookmeta.json'
OUT = '_valuescore5.json'

# ---- 与 v4 完全一致的曲线参数 ----
FLOOR = 0.20
CEIL = 1.00
AT_MEAN = 0.80
K_REL = 0.70
Z_LCB = 1.645
W_R = 0.5
REL_FLOOR = 0.65
REL_CEIL = 1.12
HEAT_BASE = 0.95
HEAT_RANGE = 0.10
K_SHRINK = 6.0
MIN_N = 8
MIN_N1 = 10

# ---- 封顶曲线：置信度允许打满的原始分门槛 ----
CAP_RAW = 91.0      # 0-100 尺度，即 9.1 分；低于它的书 conf 必须明显小于 1
CAP_GAMMA = 0.5     # cap = (raw / CAP_RAW) ** CAP_GAMMA；0.5 让低分区温和、8~9 分仍看得见折扣

# ---- v5 新增：互动证据 ----
W_E = 0.40          # 互动因子整体强度
W_MARK = 0.55       # 划线内部权重（略高于评论；数据只覆盖 Top10 热门划线，故不给更高）
W_IDEA = 0.45       # 评论内部权重
K_E = 0.50          # 互动曲线的半饱和点（低于参与率的 0.70：互动是弱证据，只在明显异常时才压分）
E_FLOOR = 0.88      # 互动修正下限（再差也只扣 12%）
E_CEIL = 1.06       # 互动修正上限（最多加 6%）
K_PSEUDO = 200.0    # 密度平滑的伪读者数（防止 50 人读 5 条想法就判高分）
X_CLIP = (0.10, 4.0)  # 相对度封顶，防单本异常值


def solve_n(k, at_mean, floor, ceil):
    target = (at_mean - floor) / (ceil - floor)
    return math.log(1.0 / target - 1.0) / math.log(k)


N_HILL = round(solve_n(K_REL, AT_MEAN, FLOOR, CEIL), 4)


def hill(x, k=K_REL, n=None):
    n = n or N_HILL
    if x is None or x <= 0:
        return 0.0
    xn = x ** n
    return xn / (k ** n + xn)


def conf_of(x):
    return FLOOR + (CEIL - FLOOR) * hill(x)


def wilson_lcb(p, n, z=Z_LCB):
    if n <= 0:
        return 0.0
    p = min(max(p, 0.0), 1.0)
    z2 = z * z
    den = 1.0 + z2 / n
    centre = p + z2 / (2.0 * n)
    half = z * math.sqrt(max(p * (1 - p) / n + z2 / (4.0 * n * n), 0.0))
    return max(0.0, min(1.0, (centre - half) / den))


def num(v):
    try:
        f = float(v)
        return f if f == f else None
    except Exception:
        return None


def split_cat(c):
    c = (c or '').strip()
    if not c:
        return '未分类', '未分类'
    if '-' in c:
        a, b = c.split('-', 1)
        return a, c
    return c, c


def geo(a):
    a = [x for x in a if x and x > 0]
    return math.exp(sum(math.log(x) for x in a) / len(a)) if a else 0.0


def pct_rank(sorted_vals, v):
    if not sorted_vals:
        return None
    lo, hi = 0, len(sorted_vals)
    while lo < hi:
        mid = (lo + hi) // 2
        if sorted_vals[mid] < v:
            lo = mid + 1
        else:
            hi = mid
    return round(100.0 * lo / len(sorted_vals), 1)


def load_meta():
    """以刷新后的 _bookmeta_v5.json 为主，旧文件（若有）只补缺失字段（防刷新中断掉书）。"""
    new = {}
    if os.path.exists(META_NEW):
        new = json.load(open(META_NEW, encoding='utf-8'))
    if not new:
        raise SystemExit('缺少 %s：请先跑 fetch 阶段（python run_pipeline.py --only fetch）' % META_NEW)
    try:
        old = json.load(open(META_OLD, encoding='utf-8'))
    except Exception:
        old = {}
    merged = {}
    for bid, ov in old.items():
        nv = new.get(bid) or {}
        rec = dict(ov)
        for k, v in nv.items():
            if v is not None:
                rec[k] = v
        rec['_stale'] = bid not in new
        merged[bid] = rec
    for bid, nv in new.items():
        if bid not in merged:
            merged[bid] = dict(nv, _stale=False)
    return merged, len(new)


def main():
    meta, n_fresh = load_meta()

    books = {}
    for bid, v in meta.items():
        raw = num(v.get('rating')) or 0.0
        rc = num(v.get('ratingCount')) or 0.0
        rd = num(v.get('readingCount')) or 0.0
        idea = num(v.get('idea'))
        mark = num(v.get('markBest')) if num(v.get('markBest')) else None
        c1, c2 = split_cat(v.get('category'))
        books[str(bid)] = {
            'bid': str(bid), 'title': v.get('title') or '', 'c1': c1, 'c2': c2,
            'raw': raw, 'rc': int(rc), 'rd': int(rd),
            'idea': idea, 'mark': mark,
            'ratio': (rc / rd) if rc > 0 and rd > 0 else None,
            'stale': bool(v.get('_stale')),
        }

    valid = [b for b in books.values() if b['raw'] > 0 and b['rc'] > 0 and b['rd'] > 0]
    for b in valid:
        b['lnrd'] = math.log(b['rd'])

    def agg(rows):
        n = len(rows)
        return {'n': n,
                'mean_r': sum(x['ratio'] for x in rows) / n,
                'geo_rd': math.exp(sum(x['lnrd'] for x in rows) / n),
                'mean_rd': sum(x['rd'] for x in rows) / n,
                'mean_raw': sum(x['raw'] for x in rows) / n}

    G = agg(valid)
    c1_rows = collections.defaultdict(list)
    c2_rows = collections.defaultdict(list)
    for b in valid:
        c1_rows[b['c1']].append(b)
        c2_rows[b['c2']].append(b)

    c1_base = {}
    for c1, rows in c1_rows.items():
        a = agg(rows)
        if a['n'] >= MIN_N1:
            c1_base[c1] = dict(a, inherit=False, raw_mean_r=a['mean_r'], raw_geo_rd=a['geo_rd'])
        else:
            c1_base[c1] = {'n': a['n'], 'mean_r': G['mean_r'], 'geo_rd': G['geo_rd'],
                           'mean_rd': G['mean_rd'], 'mean_raw': a['mean_raw'],
                           'inherit': True, 'raw_mean_r': a['mean_r'], 'raw_geo_rd': a['geo_rd']}

    c2_base = {}
    for c2, rows in c2_rows.items():
        a = agg(rows)
        c1 = c2.split('-', 1)[0] if '-' in c2 else c2
        p = c1_base.get(c1) or c1_base.get(c2) or G
        if a['n'] < MIN_N:
            c2_base[c2] = {'n': a['n'], 'mean_r': p['mean_r'], 'geo_rd': p['geo_rd'],
                           'mean_rd': p['mean_rd'], 'mean_raw': a['mean_raw'],
                           'inherit': True, 'w': 0.0, 'parent': c1,
                           'raw_mean_r': a['mean_r'], 'raw_geo_rd': a['geo_rd']}
        else:
            w = a['n'] / (a['n'] + K_SHRINK)
            c2_base[c2] = {
                'n': a['n'],
                'mean_r': w * a['mean_r'] + (1 - w) * p['mean_r'],
                'geo_rd': math.exp(w * math.log(a['geo_rd']) + (1 - w) * math.log(p['geo_rd'])),
                'mean_rd': w * a['mean_rd'] + (1 - w) * p['mean_rd'],
                'mean_raw': a['mean_raw'], 'inherit': False, 'w': w, 'parent': c1,
                'raw_mean_r': a['mean_r'], 'raw_geo_rd': a['geo_rd']}

    # ---------- 互动密度基准（按一级类；密度是"率"，直接取类内均值）----------
    def dens(b, key):
        """人均密度 = 计数 ÷ 阅读人数"""
        n = b.get(key)
        if n is None or b['rd'] <= 0:
            return None
        return float(n) / b['rd']

    def med(a):
        a = sorted(a)
        n = len(a)
        if not n:
            return None
        return a[n // 2] if n % 2 else (a[n // 2 - 1] + a[n // 2]) / 2.0

    e_base = {}
    for c1, rows in c1_rows.items():
        dm = [dens(b, 'mark') for b in rows if dens(b, 'mark') is not None]
        di = [dens(b, 'idea') for b in rows if dens(b, 'idea') is not None]
        e_base[c1] = {
            'n': len(rows), 'nMark': len(dm), 'nIdea': len(di),
            # 率类指标用中位数当基准：markBest/rd 有平台截断与长尾，均值会被几本拉高 3 倍
            'mark': med(dm), 'idea': med(di),
            'mark_mean': (sum(dm) / len(dm)) if dm else None,
            'idea_mean': (sum(di) / len(di)) if di else None,
        }
    # 全局兜底基准（类样本不足时用）
    e_G = {'mark': med([x['mark'] for x in e_base.values() if x['mark']]) or None,
           'idea': med([x['idea'] for x in e_base.values() if x['idea']]) or None}

    def smooth(cnt, rd, prior):
        if cnt is None or rd <= 0:
            return None
        return (float(cnt) + K_PSEUDO * prior) / (rd + K_PSEUDO)

    # ---------- 逐本打分 ----------
    out = {}
    for bid, b in books.items():
        base = c2_base.get(b['c2']) or c2_base.get(b['c1']) or G
        eb = e_base.get(b['c1']) or e_G
        rec = {'title': b['title'], 'raw': round(b['raw'], 1) if b['raw'] else 0.0,
               'rc': b['rc'], 'rd': b['rd'], 'idea': b['idea'], 'mark': b['mark'],
               'ratio': round(b['ratio'], 6) if b['ratio'] else None,
               'c1': b['c1'], 'c2': b['c2'], 'stale': b['stale'],
               'rbar': round(base.get('mean_r') or 0, 6),
               'rdbar': round(base.get('geo_rd') or 0, 1),
               'x_r': None, 'x_d': None, 'conf_r': None, 'conf_stat': None, 'lcb': None,
               'rel_r': None, 'heat': None, 'conf': None, 'score': 0.0, 'cap': None,
               'd_mark': None, 'd_idea': None, 'x_mark': None, 'x_idea': None,
               'x_e': None, 'conf_e': None, 'fac_e': None,
               'ratioMissing': b['ratio'] is None}
        if b['raw'] > 0 and b['rc'] > 0:
            p = b['raw'] / 100.0
            lcb = wilson_lcb(p, b['rc'])
            cs = (lcb / p) if p > 0 else 0.0

            if b['ratio'] and base.get('mean_r'):
                xr = b['ratio'] / base['mean_r']
                xd = b['rd'] / base['geo_rd'] if base.get('geo_rd') else 1.0
                cr = conf_of(xr)
                rel = cr / AT_MEAN
                heat = HEAT_BASE + HEAT_RANGE * hill(xd)
                rec.update({'x_r': round(xr, 4), 'x_d': round(xd, 4),
                            'conf_r': round(cr, 4), 'rel_r': round(rel, 4),
                            'heat': round(heat, 4)})
            else:
                rel, heat = 1.0, 1.0
                rec.update({'rel_r': 1.0, 'heat': 1.0})

            # --- 互动证据 ---
            dmk = smooth(b['mark'], b['rd'], eb.get('mark') or e_G.get('mark'))
            did = smooth(b['idea'], b['rd'], eb.get('idea') or e_G.get('idea'))
            xmk = xid = None
            if dmk is not None and eb.get('mark'):
                xmk = min(max(dmk / eb['mark'], X_CLIP[0]), X_CLIP[1])
            if did is not None and eb.get('idea'):
                xid = min(max(did / eb['idea'], X_CLIP[0]), X_CLIP[1])
            # 划线数据缺失（平台未公开热门划线）时，只按想法数评，不当作零互动
            if xmk is not None and xid is not None:
                xe = (xmk ** W_MARK) * (xid ** W_IDEA)
            elif xid is not None:
                xe = xid
            elif xmk is not None:
                xe = xmk
            else:
                xe = None
            if xe is not None:
                ce = FLOOR + (CEIL - FLOOR) * hill(xe, k=K_E)
                fac = min(max((ce / AT_MEAN) ** W_E, E_FLOOR), E_CEIL)
                rec.update({'d_mark': round(dmk, 5) if dmk else None,
                            'd_idea': round(did, 5) if did else None,
                            'x_mark': round(xmk, 4) if xmk else None,
                            'x_idea': round(xid, 4) if xid else None,
                            'x_e': round(xe, 4), 'conf_e': round(ce, 4),
                            'fac_e': round(fac, 4), 'markMissing': xmk is None})
            else:
                fac = 1.0
                rec['fac_e'] = 1.0

            rel_eff = min(max(rel ** W_R, REL_FLOOR), REL_CEIL)
            rec['rel_eff'] = round(rel_eff, 4)
            # 封顶随原始分上升：9.1 分以上才允许乘数到 1，以下按 √(raw/91) 收紧
            cap = min(1.0, (b['raw'] / CAP_RAW) ** CAP_GAMMA) if CAP_RAW > 0 else 1.0
            rec['cap'] = round(cap, 4)
            conf = min(cs * rel_eff * fac * heat, cap)
            rec.update({'lcb': round(lcb, 4), 'conf_stat': round(cs, 4),
                        'conf': round(conf, 4),
                        'score': round(b['raw'] * conf, 1),
                        'gap_lcb': round(b['raw'] - 100 * lcb, 1)})
        out[bid] = rec

    # ---------- 类内相对 ----------
    by_c1 = collections.defaultdict(list)
    for r in out.values():
        if r['conf']:
            by_c1[r['c1']].append(r)
    for c1, rows in by_c1.items():
        sc = sorted(x['score'] for x in rows)
        mean_sc = sum(sc) / len(sc)
        for r in rows:
            r['cls_n'] = len(rows)
            r['cls_mean'] = round(mean_sc, 1)
            r['cls_rel'] = round(r['score'] - mean_sc, 1)
            r['cls_pct'] = pct_rank(sc, r['score'])

    # ---------- 类型汇总 ----------
    tstat = {}
    for key in ('c1', 'c2'):
        names = sorted({(b['c1'] if key == 'c1' else b['c2']) for b in books.values()})
        for name in names:
            rows = [b for b in books.values() if (b['c1'] if key == 'c1' else b['c2']) == name]
            vrows = [out[b['bid']] for b in rows if out[b['bid']]['conf']]
            if not vrows:
                continue
            base = (c1_base if key == 'c1' else c2_base).get(name) or {}
            eb = e_base.get(name) or e_G
            scores = [x['score'] for x in vrows]
            raws = [x['raw'] for x in vrows]
            confs = [x['conf'] for x in vrows]
            facs = [x['fac_e'] for x in vrows if x.get('fac_e')]
            tstat.setdefault(key, {})[name] = {
                'n': len(rows), 'nRated': len(vrows),
                'rbar': round(base.get('mean_r') or 0, 6),
                'rdbar': round(base.get('geo_rd') or 0, 1),
                'rdbar_arith': round(base.get('mean_rd') or 0, 1),
                'rel_r': round((base.get('mean_r') or 0) / G['mean_r'], 3),
                'rel_rd': round((base.get('geo_rd') or 0) / G['geo_rd'], 3),
                'd_mark': round(eb.get('mark') or 0, 6),
                'd_idea': round(eb.get('idea') or 0, 6),
                'nMark': eb.get('nMark') or 0,
                'meanRaw': round(sum(raws) / len(raws), 1),
                'meanScore': round(sum(scores) / len(scores), 1),
                'meanConf': round(sum(confs) / len(confs), 3),
                'meanFacE': round(sum(facs) / len(facs), 3) if facs else None,
                'nLow': sum(1 for c in confs if c < 0.4),
                'shrunk': round(1 - (base.get('w') if base.get('w') is not None else 0), 3),
            }
            tstat[key][name]['inflate'] = round(sum(raws) / len(raws) - sum(scores) / len(scores), 1)

    # ---------- 版本对照 ----------
    prev = {}
    for tag, fn in (('v2', '_valuescore2.json'), ('v3', '_valuescore3.json'),
                    ('v4', '_valuescore4.json')):
        try:
            prev[tag] = json.load(open(fn, encoding='utf-8')).get('scores') or {}
        except Exception:
            prev[tag] = {}
    for bid, rec in out.items():
        for tag in ('v2', 'v3', 'v4'):
            rec[tag + '_score'] = (prev[tag].get(bid) or {}).get('score') or 0.0
        rec['d4'] = round(rec['score'] - rec['v4_score'], 1) if rec['score'] else 0.0
        rec['d3'] = round(rec['score'] - rec['v3_score'], 1) if rec['score'] else 0.0

    json.dump({'params': {
        'FLOOR': FLOOR, 'CEIL': CEIL, 'AT_MEAN': AT_MEAN, 'K_REL': K_REL,
        'N_HILL': N_HILL, 'Z_LCB': Z_LCB, 'W_R': W_R,
        'REL_FLOOR': REL_FLOOR, 'REL_CEIL': REL_CEIL,
        'HEAT_BASE': HEAT_BASE, 'HEAT_RANGE': HEAT_RANGE,
        'W_E': W_E, 'W_MARK': W_MARK, 'W_IDEA': W_IDEA, 'K_E': K_E,
        'E_FLOOR': E_FLOOR, 'E_CEIL': E_CEIL, 'K_PSEUDO': K_PSEUDO,
        'K_SHRINK': K_SHRINK, 'CAP_RAW': CAP_RAW, 'CAP_GAMMA': CAP_GAMMA,
        'n_fresh': n_fresh,
        'global': {'n': G['n'], 'mean_r': round(G['mean_r'], 6),
                   'geo_rd': round(G['geo_rd'], 1), 'mean_rd': round(G['mean_rd'], 1),
                   'mean_raw': round(G['mean_raw'], 1)}},
        'etypebase': e_base,
        'typebase': {'c1': c1_base, 'c2': c2_base},
        'typestat': tstat,
        'scores': out}, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)

    # ---------- 控制台 ----------
    allsc = sorted(r['score'] for r in out.values() if r['conf'])
    print('刷新覆盖 %d 本；有效评分 %d 本' % (n_fresh, len(allsc)))
    print('v5 均值=%.1f 中位=%.1f p10=%.1f max=%.1f | 原评分均值=%.1f → 折扣 %.1f%%'
          % (sum(allsc) / len(allsc), allsc[len(allsc) // 2], allsc[int(len(allsc) * .1)],
             allsc[-1], G['mean_raw'], 100 * (1 - (sum(allsc) / len(allsc)) / G['mean_raw'])))
    facs = [r['fac_e'] for r in out.values() if r.get('fac_e')]
    print('fac_e 均值=%.3f  <0.9 的 %d 本  >1.0 的 %d 本'
          % (sum(facs) / len(facs), sum(1 for f in facs if f < 0.9), sum(1 for f in facs if f > 1.0)))
    sc_rows = [r for r in out.values() if r['conf']]
    full = [r for r in sc_rows if r['conf'] >= 0.9999]
    print('封顶 CAP_RAW=%.1f（%.1f 分）γ=%.2f | conf 打满 %d 本（%.1f%%），最低原始分 %.1f'
          % (CAP_RAW, CAP_RAW / 10.0, CAP_GAMMA, len(full),
             100.0 * len(full) / max(1, len(sc_rows)), min((r['raw'] for r in full), default=0)))

    print('\n=== 棋经十三篇 ===')
    r = out.get('26062765')
    if r:
        print('  raw=%.1f rc=%d rd=%d | v3=%.1f v4=%.1f v5=%.1f | d_mark=%.4f d_idea=%.4f '
              'x_e=%.2f fac_e=%.3f conf=%.3f 类内%s(%.0f%%)'
              % (r['raw'], r['rc'], r['rd'], r['v3_score'], r['v4_score'], r['score'],
                 r['d_mark'] or 0, r['d_idea'] or 0, r['x_e'] or 0, r['fac_e'] or 0,
                 r['conf'], '%+.1f' % r['cls_rel'], r['cls_pct']))

    print('\n=== 互动因子加分最多 Top10 ===')
    for r in sorted([x for x in out.values() if x['conf']], key=lambda x: -x['d4'])[:10]:
        print('  %+5.1f  v4=%5.1f→v5=%5.1f  x_mark=%5.2f x_idea=%5.2f fac=%.2f  %s'
              % (r['d4'], r['v4_score'], r['score'], r['x_mark'] or 0, r['x_idea'] or 0,
                 r['fac_e'] or 0, r['title'][:24]))

    print('\n=== 互动因子扣分最多 Top10 ===')
    for r in sorted([x for x in out.values() if x['conf']], key=lambda x: x['d4'])[:10]:
        print('  %+5.1f  v4=%5.1f→v5=%5.1f  x_mark=%5.2f x_idea=%5.2f fac=%.2f  %s'
              % (r['d4'], r['v4_score'], r['score'], r['x_mark'] or 0, r['x_idea'] or 0,
                 r['fac_e'] or 0, r['title'][:24]))

    print('\n=== 一级分类（v5 均值）===')
    print('%-10s %5s %7s %7s %7s %7s' % ('类型', 'n', '原评分', 'v5', 'fac_e', 'd_mark'))
    for k, v in sorted(tstat['c1'].items(), key=lambda kv: -kv[1]['meanScore']):
        if v['nRated'] >= 10:
            print('%-10s %5d %7.1f %7.1f %7.3f %7.4f'
                  % (k, v['nRated'], v['meanRaw'], v['meanScore'], v['meanFacE'] or 0, v['d_mark'] or 0))
    print('\n输出 %s' % OUT)


if __name__ == '__main__':
    main()
