# -*- coding: utf-8 -*-
"""零钱模式的选书算法：候选池 → 三个买法 + 一个「再攒 20 书币」参考。

只做计算，不碰工作簿也不写文案；渲染在 report_low，编排在 build_excel。
所有函数显式接收参数，模块内不读环境变量、不依赖全局状态，可以单独验证。
"""
import re


def seed_as_pool(seed):
    """把全库种子转成背包可用的结构。

    种子只有微信评分，没有四层价值分（_vscore 需要阅读人数/划线等数据，
    全库书大多抓不到）。这里统一用微信评分当价值函数，并在文案里标注口径差异。
    """
    out = []
    for b in seed:
        p = float(b.get('price') or 0)
        r = float(b.get('rating') or 0)
        if p <= 0 or r <= 0:
            continue
        out.append({
            '_price': p, '_vscore': r, '_rating': r, '_risk': 0.0,
            'bookId': b.get('bookId'), 'title': b.get('title') or '',
            'author': b.get('author') or '', 'publisher': b.get('publisher') or '',
            'deepLink': b.get('deepLink') or '', '_type1': b.get('_type1') or '',
            '_type_src': b.get('_type_src') or '',
            '_rcnt': b.get('rcnt') or 0, '_kw': b.get('kw') or [],
        })
    return out


def stat(picks, budget):
    """一个方案的花费与评分合计。"""
    sp = round(sum(x['_price'] for x in picks), 2)
    tv = round(sum(x['_rating'] for x in picks), 1)
    return dict(n=len(picks), spent=sp, left=round(budget - sp, 2), tv=tv,
                av=round(tv / len(picks), 1) if picks else 0.0,
                per=round(tv / sp, 1) if sp else 0.0)


def series_key(title):
    """把「雪球漫画投资课第五期｜别再迷信牛基」这类标题折成同一个系列 key。

    取去掉期号/标点后的前 8 个字 —— 同一系列的不同期落到同一个 key，只保留评分最高的一期。
    """
    s = re.sub(r'[（(【\[].*?[)）】\]]', '', title or '')
    s = re.sub(r'[0-9]+', '', s)
    s = re.sub(r'[·•\-—|｜:：，,。、\s（）()《》"\']', '', s)
    return s[:8]


def dedupe_series(pool):
    """同一系列只留评分最高（同分则更便宜）的一期。"""
    best = {}
    for b in pool:
        k = series_key(b['title'])
        cur = best.get(k)
        if cur is None or (b['_rating'], -b['_price']) > (cur['_rating'], -cur['_price']):
            best[k] = b
    return list(best.values())


def pick_top(pool, budget, maxn):
    """按评分降序贪心装入，最多 maxn 本。比「本数优先背包」更符合低预算下的真实决策。"""
    left, out = budget, []
    for b in sorted(pool, key=lambda x: (-x['_rating'], x['_price'])):
        if len(out) >= maxn:
            break
        if b['_price'] <= left + 1e-9:
            out.append(b)
            left -= b['_price']
    return out


def make_plans(pool, budget, maxn):
    """零钱模式的 3 个方案 + 一个「再攒 20 书币」参考。

    ① 精选 N 本：同系列去重后按评分降序取，最多 maxn 本
    ② 单本最优：只买 1 本，评分最高
    ③ 花光预算：评分 ≥75 的书里挑最贵的 1 本（钱不剩，但也不买低分书）
    ④ 再攒 20 书币：按 余额+20 重算 ①，看「等一周能多买到什么」
    """
    b = budget
    pool = dedupe_series(pool)
    if not pool or b <= 0:
        return [], {}
    out = []
    pk = pick_top(pool, b, maxn)
    out.append(dict(key='best_n', no='①', name='精选 %d 本' % maxn, plan=pk,
                    stat=stat(pk, b),
                    desc='同一系列只留评分最高的一期，再按微信评分从高到低挑，最多 %d 本 —— '
                         '低预算下"买几本真正想读的"比"刷本数"更有意义' % maxn))
    one = max((x for x in pool if x['_price'] <= b + 1e-9),
              key=lambda x: (x['_rating'], -x['_price']), default=None)
    pk2 = [one] if one else []
    out.append(dict(key='best_one', no='②', name='只买 1 本', plan=pk2, stat=stat(pk2, b),
                    desc='在买得起的范围里挑评分最高的 1 本 —— 适合只想认真读一本'))
    good = [x for x in pool if x['_price'] <= b + 1e-9 and x['_rating'] >= 75]
    fill = max(good, key=lambda x: (x['_price'], x['_rating']), default=None)
    if fill is None:
        fill = max((x for x in pool if x['_price'] <= b + 1e-9),
                   key=lambda x: (x['_price'], x['_rating']), default=None)
    pk3 = [fill] if fill else []
    out.append(dict(key='fill', no='③', name='花光预算', plan=pk3, stat=stat(pk3, b),
                    desc='在评分 ≥75 的书里挑最贵的一本，让余额尽量不剩 —— '
                         '零钱留着也没用，但要避免为了花光而买低分书'))
    later = {}
    pk4 = pick_top(pool, b + 20, maxn)
    _t0 = (len(pk), round(sum(x['_rating'] for x in pk), 1))
    _t4 = (len(pk4), round(sum(x['_rating'] for x in pk4), 1))
    if _t4 > _t0:
        later = dict(plan=pk4, stat=stat(pk4, b + 20),
                     addN=max(0, len(pk4) - len(pk)),
                     gain=round(_t4[1] - _t0[1], 1))
    return out, later
