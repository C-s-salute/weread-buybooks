# -*- coding: utf-8 -*-
"""抓取微信读书公开书单（/misc/booklist/<booklistId>）并解析为 _booklists.json。

用法：
  1) 把书单分享链接逐行写进 booklist_urls.txt
  2) python fetch_booklist.py
无需登录：公开书单页服务端渲染，__NUXT__ 里带 booklist + bookEntities + bookIds。
"""
import json
import os
import re
import subprocess
import sys
from datetime import datetime

import requests

import wconfig

URLS = 'booklist_urls.txt'
OUT = '_booklists.json'
TMP = '_bl_tmp'
HOST = 'https://weread.qq.com'
ALLOWED_HOSTS = ('weread.qq.com',)
UA = ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
      '(KHTML, like Gecko) Chrome/120.0 Safari/537.36')


def extract_nuxt(html_path):
    nuxt_js = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'nuxt_extract.js')
    r = subprocess.run([wconfig.node_bin(), nuxt_js, html_path],
                       capture_output=True, text=True, encoding='utf-8')
    if r.returncode != 0:
        raise RuntimeError('nuxt eval fail: %s' % (r.stderr or '')[:120])
    return json.loads(r.stdout)


def parse_one(url):
    from urllib.parse import urlparse
    host = (urlparse(url.strip()).hostname or '').lower()
    if host not in ALLOWED_HOSTS:
        raise ValueError('只接受 %s 的分享链接，当前来自: %s'
                         % ('/'.join(ALLOWED_HOSTS), host or '(空)'))
    m = re.search(r'/misc/booklist/([0-9A-Za-z_]+)', url)
    if not m:
        raise ValueError('无法从链接中解析 booklistId: %s' % url)
    bid = m.group(1)
    os.makedirs(TMP, exist_ok=True)
    hp = os.path.join(TMP, bid + '.html')
    r = requests.get('%s/misc/booklist/%s' % (HOST, bid),
                     headers={'User-Agent': UA}, timeout=30)
    r.raise_for_status()
    open(hp, 'w', encoding='utf-8').write(r.text)
    nuxt = extract_nuxt(hp)
    data = (nuxt.get('data') or [{}])[0]
    bl = data.get('booklist') or {}
    if not bl:
        raise RuntimeError('书单不存在或已删除: %s' % bid)
    ents = data.get('bookEntities') or {}
    ids = [str(x) for x in (data.get('bookIds') or [])]
    books = []
    for i, bkid in enumerate(ids, 1):
        e = ents.get(bkid) or {}
        books.append({
            'order': i,
            'bookId': bkid,
            'title': e.get('title') or '',
            'author': e.get('author') or '',
            'rating': e.get('ratingString') or '',
            'ratingTitle': e.get('ratingTitle') or '',
            'soldout': 1 if e.get('isSoldOut') in (1, True) else 0,
            'inEntity': bool(e),
        })
    return bid, {
        'booklistId': bid,
        'name': bl.get('name') or '',
        'description': bl.get('description') or '',
        'authorName': (bl.get('author') or {}).get('name') or '',
        'authorVid': (bl.get('author') or {}).get('userVid'),
        'bookCount': bl.get('bookCount') or len(ids),
        'url': url.strip(),
        'fetchedAt': datetime.now().strftime('%Y-%m-%d %H:%M'),
        'bookIds': ids,
        'books': books,
    }


def main():
    if not os.path.exists(URLS):
        print('缺少 %s，请把书单链接逐行写入' % URLS)
        return 1
    urls = [l.strip() for l in open(URLS, encoding='utf-8') if l.strip()
            and not l.strip().startswith('#')]
    store = {}
    if os.path.exists(OUT):
        store = json.load(open(OUT, encoding='utf-8'))
    for u in urls:
        try:
            bid, rec = parse_one(u)
            store[bid] = rec
            print('[ok] %-26s %2d 本  %s' % (rec['name'], len(rec['books']), bid))
        except Exception as e:
            print('[fail] %s -> %s' % (u, str(e)[:100]))
    json.dump(store, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print('saved %s, 书单数 %d' % (OUT, len(store)))
    return 0


if __name__ == '__main__':
    sys.exit(main())
