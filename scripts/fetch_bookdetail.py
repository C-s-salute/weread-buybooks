# -*- coding: utf-8 -*-
"""为书单里「不在书架」的书补抓详情（价格 / 推荐值 / 分类），存 _bookdetail_extra.json。

数据源：登录态 GET /web/book/info?bookId=（含 centPrice/price/soldout/newRating/ratingCount/
        newRatingDetail/category/publishTime/totalWords）。
"""
import json
import sys
import time

from playwright.sync_api import sync_playwright

import wconfig

OUT = '_bookdetail_extra.json'
KEEP = ('title', 'author', 'price', 'centPrice', 'originalPrice', 'soldout', 'newRating',
        'ratingCount', 'newRatingCount', 'newRatingDetail', 'category', 'publishTime',
        'totalWords', 'payType', 'isBuy', 'canBuy', 'publisher', 'isbn', 'type')


def main():
    try:
        BL = json.load(open('_booklists.json', encoding='utf-8'))
    except Exception:
        print('缺少 _booklists.json')
        return 1
    shelf = json.load(open('_web_shelf2.json', encoding='utf-8'))
    in_shelf = {str(b['bookId']) for b in shelf['books']}
    store = {}
    try:
        store = json.load(open(OUT, encoding='utf-8'))
    except Exception:
        pass
    pending = []
    for rec in BL.values():
        for b in rec['books']:
            bid = b['bookId']
            if bid in in_shelf or bid in store:
                continue
            pending.append(bid)
    print('待抓 %d 本' % len(pending))
    if not pending:
        return 0
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            wconfig.PROFILE_DIR, **wconfig.launch_kwargs(headless=True))
        pg = ctx.pages[0] if ctx.pages else ctx.new_page()
        pg.goto('https://weread.qq.com/', wait_until='domcontentloaded', timeout=60000)
        pg.wait_for_timeout(800)
        for i, bid in enumerate(pending, 1):
            try:
                r = ctx.request.get('https://weread.qq.com/web/book/info?bookId=%s' % bid,
                                    timeout=20000)
                j = r.json()
                store[bid] = {k: j.get(k) for k in KEEP if k in j}
                cp = j.get('centPrice') or 0
                print('[%2d/%d] %-4s %-30s %s 元  推荐 %s (%s 人)' %
                      (i, len(pending), bid, (j.get('title') or '')[:28],
                       round(cp / 100, 2) if cp else '-', j.get('newRating'), j.get('ratingCount')))
            except Exception as e:
                print('[%2d/%d] %s FAIL %s' % (i, len(pending), bid, str(e)[:60]))
            time.sleep(0.4)
        ctx.close()
    json.dump(store, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print('saved %s  n=%d' % (OUT, len(store)))
    return 0


if __name__ == '__main__':
    sys.exit(main())
