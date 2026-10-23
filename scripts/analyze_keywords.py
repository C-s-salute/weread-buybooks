# -*- coding: utf-8 -*-
"""分析 lowprice_seed.json：找关键词的「零贡献」与书名的「公约数」。

两个用途：
  1. **体检现有词表**：哪些词一本都没捞到（白花请求），哪些词贡献最大。
  2. **挖书名公约数**：对已命中的书名做 n-gram + 贪心集合覆盖，
     找出「用一个片段就能囊括多本书」的词（例如『法典』能覆盖『民法典·总则/合同/婚姻家庭』，
     『雪球』能覆盖 49 本雪球系专刊）。

用法：
  python analyze_keywords.py                       # 看完整报告
  python analyze_keywords.py --min-cover 3         # 只列出能覆盖 ≥3 本的书名片段
  python analyze_keywords.py --apply               # 顺便把零贡献词从输出里标出来（不改文件）
"""
import argparse
import collections
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_SEED = os.path.join(HERE, '..', 'data', 'lowprice_seed.json')

# 片段里出现这些字就丢弃（虚词、数量词、单字高频词），避免切出「的」「一个」这类噪声
BAD_CHARS = set('的了是在和与及之一二三四五六七八九十零个本全上下中大小我你他它们这那有无不就都也又还很最更'
                '些每各该其此并或但而且因所以为对从到把被让使将已经过着地得')
BAD_PREFIX = ('的', '了', '和', '与', '是', '在', '为', '对', '把', '被')


def clean_title(s):
    """去掉括注与标点，只留主体，避免「（大全集）」这类尾巴干扰片段统计。"""
    s = re.sub(r'[（(【\[].*?[)）】\]]', '', s or '')
    return re.sub(r'[·•：:，,。、！？!?\-\s《》"\']', '', s)


def load_keywords():
    try:
        sys.path.insert(0, HERE)
        from fetch_lowprice import GROUPS
        return {w for _n, _p, words in GROUPS for w in words}
    except Exception:
        return set()


def contribution(books, keywords):
    cnt = collections.Counter()
    for b in books:
        for k in b.get('kw') or []:
            cnt[k] += 1
    rows = sorted(cnt.items(), key=lambda x: -x[1])
    zero = sorted(w for w in keywords if cnt.get(w, 0) == 0)
    return cnt, rows, zero


def ngram_candidates(titles):
    cand = collections.defaultdict(set)
    for i, s in enumerate(titles):
        for n in range(2, 7):
            for j in range(len(s) - n + 1):
                g = s[j:j + n]
                if re.search(r'\d', g):
                    continue
                if any(ch in BAD_CHARS for ch in g):
                    continue
                if g.startswith(BAD_PREFIX):
                    continue
                cand[g].add(i)
    return cand


def greedy_cover(cand, total, min_cover, keywords, max_pick=40):
    """贪心集合覆盖：每次取「能覆盖最多未覆盖书」的片段。返回 [(片段, 新增覆盖, 累计覆盖)]。"""
    uncovered = set(range(total))
    out = []
    cand = {g: v for g, v in cand.items() if len(v) >= min_cover}
    while len(out) < max_pick:
        best, best_gain = None, 0
        for g, idxs in cand.items():
            gain = len(idxs & uncovered)
            if gain > best_gain or (gain == best_gain and best and len(g) < len(best)):
                best, best_gain = g, gain
        if best_gain < min_cover:
            break
        hit = len(cand[best] & uncovered)
        out.append((best, hit, len(cand[best])))
        uncovered -= cand[best]
    return out, uncovered


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--seed', default=DEFAULT_SEED)
    ap.add_argument('--min-cover', type=int, default=2, help='片段至少覆盖几本才纳入（默认 2）')
    ap.add_argument('--top', type=int, default=40, help='公约数最多列几条')
    a = ap.parse_args()

    d = json.load(open(a.seed, encoding='utf-8'))
    books, meta = d.get('books') or [], d.get('meta') or {}
    if not books:
        print('种子为空：%s' % a.seed)
        return 1

    kw = load_keywords()
    used = meta.get('keywordsUsed') or []
    if used:
        # 优先用种子记录的「本次实际跑过的词表」，这样零贡献判断才准
        scope, scope_note = set(used), '本次抓取实际用过的 %d 个词' % len(used)
    else:
        scope, scope_note = kw, '当前 GROUPS 的 %d 个词（种子未记录词表快照，判断可能不准）' % len(kw)
    cnt, rows, zero = contribution(books, scope)

    print('=' * 64)
    print('种子：%s' % os.path.abspath(a.seed))
    print('     %d 本 · cap %.2f 元 · 评分 ≥%s · 生成于 %s'
          % (len(books), meta.get('cap', 0), meta.get('minRating'), meta.get('generatedAt')))
    print('=' * 64)

    print('\n【一】关键词贡献 Top20')
    for k, n in rows[:20]:
        print('  %-14s %3d 本' % (k, n))

    if scope:
        print('\n【二】零贡献词：%d / %d（占 %.0f%%）｜口径：%s'
              % (len(zero), len(scope), 100.0 * len(zero) / len(scope), scope_note))
        for i in range(0, len(zero), 8):
            print('  ' + '、'.join(zero[i:i + 8]))
        print('  → 这些词在本次最多 %s 页的翻页深度内一本都没捞到，可以删掉省请求。'
              % meta.get('pagesByGroup', {}).get('默认', 'N'))

    titles = [b.get('title') or '' for b in books]
    cleaned = [clean_title(t) for t in titles]
    cand = ngram_candidates(cleaned)
    picked, uncovered = greedy_cover(cand, len(books), a.min_cover, kw, a.top)

    print('\n【三】书名公约数（覆盖 ≥%d 本的片段，贪心挑选）' % a.min_cover)
    print('  %-16s %-6s %-6s %s' % ('片段', '新增覆盖', '总覆盖', '已在词表?'))
    for g, hit, tot in picked:
        print('  %-16s %-8d %-8d %s' % (g, hit, tot, '✅' if g in kw else '← 建议新增'))

    print('\n  未被任何片段覆盖的书：%d 本（这些只能靠独立关键词捞）' % len(uncovered))
    for i in sorted(uncovered)[:10]:
        print('    · %s' % titles[i][:34])

    print('\n提示：公约数只是「已有书的聚合」——加进词表是为了用更少的词捞到更多同类书，')
    print('     不等于这些片段一定能搜到新书。加之前建议先用 --keywords 单独小跑验证。')
    return 0


if __name__ == '__main__':
    sys.exit(main())
