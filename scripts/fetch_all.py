# -*- coding: utf-8 -*-
"""微信读书底层数据全量刷新（可断点续跑）。

刷新内容与来源：
  1) 账户资产        POST /web/pay/balance + /web/pay/memberCardSummary   → _wallet3.json
  2) 书架全量        GET  /web/shelf/sync                                 → _web_shelf2.json
  3) 书目元数据      GET  /web/book/info?bookId=                          → 价格/权益/评分/推荐值/分类
  4) 累计阅读人数    GET  /book-detail?type=1&v=<encodeId>  (SSR)         → readingStat.readingCount
  5) 想法数          GET  /web/review/list listType=2 totalCount          → idea
  6) 热门划线人次    GET  /web/book/bestbookmarks  (items[].totalCount)   → mark
  7) 阅读进度        GET  /web/book/getProgress?bookId=                   → _progress2.json

产出：
  _bookmeta_v5.json     {bookId: {...meta..., readingCount, idea, mark, markTop, nBest}}
  _refresh_meta.json    刷新时间戳 / 条数 / 失败数
用法：
  python fetch_all.py            # 全量刷新
  python fetch_all.py --meta     # 只刷元数据+互动（不动钱包/书架）
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
META_OUT = '_bookmeta_v5.json'
PROG_STATE = '_refresh_state.json'
LOG = '_refresh_meta.json'
SHELF_OUT = '_web_shelf2.json'
WALLET_OUT = '_wallet3.json'
PROG_OUT = '_progress2.json'


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


def cat(b):
    if str(b.get('paid')) == '1':
        return 'P'
    return {'0': 'F', '2': 'M', '1': 'T'}.get(str(b.get('payingStatus')), 'X')


def main():
    only_meta = '--meta' in sys.argv
    t_start = time.time()

    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            UDD, **wconfig.launch_kwargs(headless=True))
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        page.goto(BASE + '/', wait_until='domcontentloaded', timeout=60000)
        page.wait_for_timeout(1200)
        if 'wr_vid' not in {c['name'] for c in ctx.cookies()}:
            print('LOGIN_LOST：登录态失效，请先运行 python scripts/login.py 扫码登录')
            return 2
        print('[1/5] 登录态正常', flush=True)

        # ---------- 1. 钱包 ----------
        if not only_meta:
            wallet = {}
            H = {'Content-Type': 'application/json;charset=UTF-8'}
            for tag, pf in (('android', 'android_html'), ('ios', 'ios_html')):
                try:
                    r = ctx.request.post(BASE + '/web/pay/balance',
                                         data=json.dumps({'pf': pf}), headers=H, timeout=20000)
                    wallet[tag] = r.json()
                except Exception as e:
                    print('  wallet %s ERR %s' % (tag, e))
            try:
                wallet['memberCard'] = ctx.request.get(
                    BASE + '/web/pay/memberCardSummary?pf=ios', timeout=20000).json()
            except Exception:
                pass
            wallet['fetched_at'] = int(time.time())
            jdump(wallet, WALLET_OUT)
            print('[2/5] 钱包已刷新: %s' % json.dumps(wallet.get('android'), ensure_ascii=False)[:160], flush=True)

            # ---------- 2. 书架 ----------
            r = ctx.request.get(BASE + '/web/shelf/sync', timeout=60000)
            shelf = r.json()
            jdump(shelf, SHELF_OUT)
            print('[3/5] 书架已刷新: books=%d archive=%d progress=%d'
                  % (len(shelf.get('books') or []), len(shelf.get('archive') or []),
                     len(shelf.get('bookProgress') or [])), flush=True)
        else:
            shelf = jload(SHELF_OUT)
            print('[2-3/5] 跳过钱包/书架（--meta）', flush=True)

        books = shelf.get('books') or []
        ids = [str(b['bookId']) for b in books]
        lim = int(os.environ.get('LIMIT', '0') or 0)
        if lim:
            ids = ids[:lim]
            print('LIMIT=%d 仅测试前 %d 本' % (lim, len(ids)))
        print('共 %d 本待刷新' % len(ids), flush=True)

        # ---------- 3. 逐本元数据 + 阅读人数 + 想法 + 热门划线 ----------
        meta = jload(META_OUT)
        state = jload(PROG_STATE)
        if state.get('run_id') != os.environ.get('RUN_ID', 'v5'):
            pass
        done = set(state.get('done') or [])
        fails = list(state.get('fails') or [])

        n_new = 0
        t0 = time.time()
        for i, bid in enumerate(ids, 1):
            if bid in done and bid in meta:
                continue
            rec = meta.get(bid) or {}
            try:
                info = ctx.request.get(BASE + '/web/book/info?bookId=' + bid, timeout=20000).json()
                rec.update({
                    'ok': True, 'title': info.get('title'), 'author': info.get('author'),
                    'payType': info.get('payType'),
                    'price': (info.get('price') if (info.get('price') or 0) > 0
                              else (info.get('centPrice') or 0) / 100.0),
                    'centPrice': info.get('centPrice'), 'originalPrice': info.get('originalPrice'),
                    'soldout': info.get('soldout'), 'free': info.get('free'),
                    'mcardDiscount': info.get('mcardDiscount'),
                    'type': info.get('type'), 'category': info.get('category'),
                    'publishTime': info.get('publishTime'), 'publisher': info.get('publisher'),
                    'isbn': info.get('isbn'), 'wordCount': info.get('totalWords') or 0,
                    'newRating': info.get('newRating'),
                    'rating': round((info.get('newRating') or 0) / 10.0, 1),
                    'ratingCount': info.get('newRatingCount'),
                    'ratingCountOld': info.get('ratingCount'),
                    'ratingDetail': info.get('newRatingDetail'),
                    'encodeId': (info.get('deepLink') or '').split('v=')[-1] or info.get('encodeId'),
                })
                # 累计阅读人数（分享页 SSR）
                enc = rec['encodeId']
                rd = None
                if enc:
                    for _ in range(2):
                        try:
                            html = ctx.request.get(
                                BASE + '/book-detail?type=1&v=' + enc, timeout=25000).text()
                            m = re.search(r'readingStat:\{readingCount:(\d+)\}', html)
                            if m:
                                rd = int(m.group(1))
                            if rd is None:
                                m = re.search(r'"readingCount":(\d+)', html)
                                rd = int(m.group(1)) if m else None
                            if rd:
                                break
                        except Exception:
                            time.sleep(0.4)
                if rd:
                    rec['readingCount'] = rd
                rec['readingCount_at'] = int(time.time())

                # 想法数（评论总量）
                try:
                    j2 = ctx.request.get(BASE + '/web/review/list',
                                         params={'bookId': bid, 'listType': '2',
                                                 'maxIdx': '0', 'count': '1'}, timeout=15000).json()
                    rec['idea'] = j2.get('totalCount')
                except Exception:
                    pass
                # 热门划线人次（Top10 累计 + 条数）
                try:
                    j3 = ctx.request.get(BASE + '/web/book/bestbookmarks',
                                         params={'bookId': bid}, timeout=20000).json()
                    bb = (j3 or {}).get('bestBookMarks') or {}
                    items = bb.get('items') or []
                    rec['markTop'] = sum(int(x.get('totalCount') or 0) for x in items)
                    rec['markBest'] = max([int(x.get('totalCount') or 0) for x in items] or [0])
                    rec['nBest'] = len(items)
                    rec['nBestChapter'] = bb.get('totalCount')
                except Exception:
                    pass
                meta[bid] = rec
                done.add(bid)
                n_new += 1
            except Exception as e:
                fails.append({'bookId': bid, 'err': str(e)[:120]})
                print('  FAIL %s %s' % (bid, str(e)[:80]), flush=True)

            if i % 25 == 0:
                jdump(meta, META_OUT)
                jdump({'done': sorted(done), 'fails': fails, 'updated': int(time.time())}, PROG_STATE)
                el = time.time() - t0
                rate = n_new / el if el else 0
                eta = (len(ids) - i) / rate / 60 if rate else 0
                print('  %d/%d  新增 %d  失败 %d  %.2f本/秒  预计剩余 %.1f 分钟'
                      % (i, len(ids), n_new, len(fails), rate, eta), flush=True)

        jdump(meta, META_OUT)
        jdump({'done': sorted(done), 'fails': fails, 'updated': int(time.time())}, PROG_STATE)
        print('[4/5] 元数据完成: %d 本，失败 %d，用时 %.1f 分钟'
              % (len(meta), len(fails), (time.time() - t0) / 60), flush=True)

        # ---------- 4. 阅读进度 ----------
        if not only_meta:
            targets = [b for b in books if cat(b) in ('M', 'T')]
            prog = {}
            t1 = time.time()
            for i, b in enumerate(targets, 1):
                bid = str(b['bookId'])
                try:
                    j = ctx.request.get(BASE + '/web/book/getProgress',
                                        params={'bookId': bid}, timeout=15000).json()
                    bk = j.get('book') or {}
                    prog[bid] = {'bookId': bid, 'progress': bk.get('progress'),
                                 'chapterIdx': bk.get('chapterIdx'),
                                 'readingTime': bk.get('readingTime'),
                                 'updateTime': bk.get('updateTime'),
                                 'finishTime': bk.get('finishTime'),
                                 'startReadingTime': bk.get('startReadingTime')}
                except Exception:
                    pass
                if i % 50 == 0:
                    print('  progress %d/%d' % (i, len(targets)), flush=True)
            jdump(prog, PROG_OUT)
            print('[5/5] 阅读进度完成: %d 本，用时 %.1f 秒' % (len(prog), time.time() - t1), flush=True)

        ctx.close()

    jdump({'finished_at': int(time.time()), 'elapsed_min': round((time.time() - t_start) / 60, 1),
           'books': len(ids), 'meta': len(meta), 'fails': len(fails),
           'wallet': jload(WALLET_OUT).get('android') if not only_meta else None},
          LOG)
    print('DONE 总用时 %.1f 分钟' % ((time.time() - t_start) / 60))
    return 0


if __name__ == '__main__':
    sys.exit(main())
