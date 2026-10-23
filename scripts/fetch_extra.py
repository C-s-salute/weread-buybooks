# -*- coding: utf-8 -*-
"""补抓「书单里但不在书架」的书：阅读人数 / 想法数 / 热门划线人次。

为什么需要：书架上的书已有 _bookmeta_v5.json（含 readingCount / idea / markBest），
但书单里没加书架的书，_bookdetail_extra.json 只补了价格与评分，
缺阅读人数 → 无法按同一套 v5 公式给它们打分，也无法和书架书一起排热度。

抓取内容（每本 3~4 次请求）：
  1) GET /web/book/info?bookId=        → deepLink / encodeId（取分享页 v 参数）
  2) GET /book-detail?type=1&v=<enc>   → readingStat.readingCount（累计阅读人数，SSR）
  3) GET /web/review/list listType=2   → totalCount（想法数）
  4) GET /web/book/bestbookmarks       → items[].totalCount（热门划线人次 markBest / markTop）

产出：_bookmeta_v5_extra.json  {bookId: {..., readingCount, idea, markBest, markTop}}
可断点续跑：_refresh_extra_state.json 记录已完成 id。
用法：python fetch_extra.py [LIMIT=50 环境变量限量测试]
"""
import json
import os
import re
import sys
import time

from playwright.sync_api import sync_playwright

import wconfig

BASE = 'https://weread.qq.com'
UDD = wconfig.PROFILE_DIR
OUT = '_bookmeta_v5_extra.json'
STATE = '_refresh_extra_state.json'
BL = '_booklists.json'
DET = '_bookdetail_extra.json'


def jload(p, default=None):
    try:
        return json.load(open(p, encoding='utf-8'))
    except Exception:
        return default if default is not None else {}


def jdump(obj, p):
    tmp = p + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(obj, f, ensure_ascii=False)
    os.replace(tmp, p)


def main():
    bl = jload(BL)
    det = jload(DET)
    ids = []
    seen = set()
    for rec in bl.values():
        for b in rec.get('books') or []:
            bid = str(b['bookId'])
            if bid in seen:
                continue
            seen.add(bid)
            ids.append(bid)
    try:
        shelf = set(str(b['bookId']) for b in (jload('_web_shelf2.json').get('books') or []))
    except Exception:
        shelf = set()
    targets = [b for b in ids if b not in shelf]
    lim = int(os.environ.get('LIMIT', '0') or 0)
    if lim:
        targets = targets[:lim]
    print('书单去重 %d 本，其中不在书架 %d 本待抓' % (len(ids), len(targets)), flush=True)

    out = jload(OUT)
    st = jload(STATE)
    done = set(st.get('done') or [])
    fails = list(st.get('fails') or [])

    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            UDD, **wconfig.launch_kwargs(headless=True))
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.goto(BASE + '/', wait_until='domcontentloaded', timeout=60000)
        page.wait_for_timeout(1000)
        if 'wr_vid' not in {c['name'] for c in ctx.cookies()}:
            print('LOGIN_LOST：登录态失效，请先运行 python scripts/login.py 扫码登录')
            return 2
        print('登录态正常', flush=True)

        t0 = time.time()
        n_new = 0
        for i, bid in enumerate(targets, 1):
            if bid in done and bid in out:
                continue
            rec = dict(out.get(bid) or {})
            try:
                info = ctx.request.get(BASE + '/web/book/info?bookId=' + bid,
                                       timeout=20000).json()
                enc = (info.get('deepLink') or '').split('v=')[-1] or info.get('encodeId') or ''
                rec.update({
                    'ok': True,
                    'title': info.get('title'),
                    'author': info.get('author'),
                    'category': info.get('category'),
                    'price': (info.get('price') if (info.get('price') or 0) > 0
                              else (info.get('centPrice') or 0) / 100.0),
                    'centPrice': info.get('centPrice'),
                    'soldout': info.get('soldout'),
                    'payType': info.get('payType'),
                    'publishTime': info.get('publishTime'),
                    'publisher': info.get('publisher'),
                    'isbn': info.get('isbn'),
                    'wordCount': info.get('totalWords') or 0,
                    'newRating': info.get('newRating'),
                    'rating': round((info.get('newRating') or 0) / 10.0, 1),
                    'ratingCount': info.get('newRatingCount') or info.get('ratingCount'),
                    'ratingCountOld': info.get('ratingCount'),
                    'encodeId': enc,
                })
                # 累计阅读人数（分享页 SSR）
                rd = None
                if enc:
                    for _ in range(2):
                        try:
                            html = ctx.request.get(BASE + '/book-detail?type=1&v=' + enc,
                                                   timeout=25000).text()
                            m = re.search(r'readingStat:\{readingCount:(\d+)\}', html)
                            if not m:
                                m = re.search(r'"readingCount":(\d+)', html)
                            if m:
                                rd = int(m.group(1))
                            if rd:
                                break
                        except Exception:
                            time.sleep(0.4)
                if rd:
                    rec['readingCount'] = rd
                rec['readingCount_at'] = int(time.time())

                try:
                    j2 = ctx.request.get(BASE + '/web/review/list',
                                         params={'bookId': bid, 'listType': '2',
                                                 'maxIdx': '0', 'count': '1'},
                                         timeout=15000).json()
                    rec['idea'] = j2.get('totalCount')
                except Exception:
                    pass
                try:
                    j3 = ctx.request.get(BASE + '/web/book/bestbookmarks',
                                         params={'bookId': bid}, timeout=20000).json()
                    bb = (j3 or {}).get('bestBookMarks') or {}
                    items = bb.get('items') or []
                    rec['markTop'] = sum(int(x.get('totalCount') or 0) for x in items)
                    rec['markBest'] = max([int(x.get('totalCount') or 0) for x in items] or [0])
                    rec['nBest'] = len(items)
                except Exception:
                    pass
                out[bid] = rec
                done.add(bid)
                n_new += 1
            except Exception as e:
                fails.append({'bookId': bid, 'err': str(e)[:120]})
                print('  FAIL %s %s' % (bid, str(e)[:80]), flush=True)

            if i % 20 == 0:
                jdump(out, OUT)
                jdump({'done': sorted(done), 'fails': fails, 'updated': int(time.time())}, STATE)
                el = time.time() - t0
                rate = n_new / el if el else 0
                eta = (len(targets) - i) / rate / 60 if rate else 0
                print('  %d/%d 新增 %d 失败 %d %.2f本/秒 剩余约 %.1f 分钟'
                      % (i, len(targets), n_new, len(fails), rate, eta), flush=True)

        jdump(out, OUT)
        jdump({'done': sorted(done), 'fails': fails, 'updated': int(time.time())}, STATE)
        print('完成：%d 本，失败 %d，用时 %.1f 分钟'
              % (len(out), len(fails), (time.time() - t0) / 60), flush=True)
        ctx.close()
    return 0


if __name__ == '__main__':
    sys.exit(main())
