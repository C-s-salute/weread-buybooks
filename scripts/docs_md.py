# -*- coding: utf-8 -*-
"""把「可获取信息与可做功能」「互动证据」「口径说明」三块内容写成 Markdown。

这三块是"说明性文本 + 支撑数据"，放在 Excel 里占 sheet 且不好读，所以从工作簿里移出来。
用法（在 build_excel.py 的 main 里）：
    from docs_md import build_docs_md
    build_docs_md(globals())
"""
import json

try:
    import wconfig   # noqa: F401  统一工作目录（设置了 WEREAD_WORK 时切 cwd）
except ImportError:
    pass


def _tbl(heads, rows, aligns=None):
    out = ['| ' + ' | '.join(str(h) for h in heads) + ' |',
           '|' + '|'.join(['---'] * len(heads)) + '|']
    for r in rows:
        cells = []
        for v in r:
            s = '' if v is None else str(v)
            cells.append(s.replace('|', '／').replace('\n', ' '))
        out.append('| ' + ' | '.join(cells) + ' |')
    return '\n'.join(out)


def _fmt(v, d=1, dash='—'):
    if v is None:
        return dash
    try:
        return ('%.' + str(d) + 'f') % float(v)
    except Exception:
        return str(v)


def build_docs_md(G, path='微信读书_方法与口径说明.md'):
    BOOKS = G['BOOKS']
    agg = G['agg']
    COIN = G['COIN']
    SNAP_NOW = G['SNAP_NOW']
    VSPARAM = G.get('VSPARAM') or {}
    VGLOBAL = G.get('VGLOBAL') or {}
    VTYPE = G.get('VTYPE') or {}
    UNION = G.get('UNION') or []
    PS_RATIO, PS_CNT = G['PS_RATIO'], G['PS_CNT']
    PS_VAL, PS_RISK = G['PS_VAL'], G['PS_RISK']
    VAL_PICK = G['VAL_PICK']
    TYPE1_STD = G['TYPE1_STD']
    POW_CANDS = G['POW_CANDS']
    VAL_CANDS = G['VAL_CANDS']
    COIN_OK, COIN_OK_MONEY = G['COIN_OK'], G['COIN_OK_MONEY']
    COIN_POOL, COIN_POOL_X = G['COIN_POOL'], G['COIN_POOL_X']
    LOW_POOL = G.get('LOW_POOL') or []
    LOW_AFFORD = G.get('LOW_AFFORD') or []
    LOW_CAP = G.get('LOW_CAP') or 2.0
    LOW_SEED = G.get('LOW_SEED') or []
    LOW_SEED_META = G.get('LOW_SEED_META') or {}
    NEED = G['NEED']
    MAX_N, G_X = G['MAX_N'], G['G_X']
    try:
        BLX = json.load(open('_booklists.json', encoding='utf-8'))
    except Exception:
        BLX = {}

    rated = [b for b in BOOKS if (b['_rating'] or 0) > 0 and b.get('_conf')]
    n_rated = len(rated)
    have = [b for b in rated if b.get('_fac_e') is not None]
    shelf_n = sum(1 for x in UNION if x['inShelf'])
    out_n = len(UNION) - shelf_n
    rated_u = sum(1 for x in UNION if x['vscore'])
    bl_tot = sum(len(r.get('books') or []) for r in BLX.values()) if BLX else 0
    bl_uniq = len({str(b.get('bookId')) for r in BLX.values() for b in (r.get('books') or [])}) if BLX else 0

    L = []
    A = L.append
    A('# 微信读书书架分析：方法与口径说明')
    A('')
    A('> 这份文档承接原工作簿里「可获取信息与可做功能」「互动证据」「口径说明」三个 sheet 的全部内容。')
    A('> 工作簿只保留数据表（方案、清单、评分、排名），方法与口径统一放这里，便于随参数一起存档。')
    A('')
    A('- 数据快照：%s' % SNAP_NOW)
    A('- 书架：%d 本　·　公开书单：%d 个 / %d 条归属 / %d 本去重'
      % (len(BOOKS), len(BLX), bl_tot, bl_uniq))
    A('- 合并去重（书架 ∪ 书单）：%d 本（书架 %d ＋ 书单里未加架 %d），其中有评分 %d 本'
      % (len(UNION), shelf_n, out_n, rated_u))
    A('- 书币余额：%.2f 元（赠币＋充值币，只能在站内买断电子书）' % COIN)
    A('')

    # ============================================================ 一
    A('## 一、可获取信息与可做功能')
    A('')
    A('全部为登录态接口实测。✅＝已接入工作簿　🟡＝可接但未接　❌＝网页端拿不到。')
    A('')
    A('### 1.1 书籍维度：每本书能拿到的字段')
    A('')
    A(_tbl(['字段', '来源接口 / 字段名', '覆盖率', '说明', '状态'], [
        ['权益类型', '/web/shelf/sync → paid + payingStatus', '%d/%d' % (len(BOOKS), len(BOOKS)),
         '永久可读 / 会员卡可读 / 体验卡可读 / 免费，已实测校准', '✅'],
        ['买断价', '详情页 NUXT centPrice ÷ 100', '%d/%d' % (len(BOOKS), len(BOOKS)),
         '会员卡可读、体验卡可读、已拥有都有标价', '✅'],
        ['评分（推荐值）', 'bookInfo.newRating ÷ 10', '%d/%d' % (n_rated, len(BOOKS)),
         '区间 %.1f~%.1f，中位 %.1f' % (min(b['_rating'] for b in rated),
                                     max(b['_rating'] for b in rated),
                                     sorted(b['_rating'] for b in rated)[n_rated // 2]), '✅'],
        ['评分人数', 'newRatingCount', '%d/%d' % (n_rated, len(BOOKS)), '用于 Wilson 下界与下架风险', '✅'],
        ['累计阅读人数', '/book-detail?type=1&v=<encodeId> 分享页 SSR readingStat.readingCount',
         '%d/%d' % (sum(1 for b in BOOKS if b['_readcnt']), len(BOOKS)),
         '注意：/web/book/readInfo 给的是「当前在读」，不是累计阅读人数', '✅'],
        ['想法数', '/web/review/list listType=2 → totalCount', '%d/%d' % (sum(1 for b in BOOKS if b['_idea'] is not None), len(BOOKS)),
         '读者在划线下写的短评数；listType=3 是星评人数，不是划线数', '✅'],
        ['热门划线人次', '/web/book/bestbookmarks → items[].totalCount', '%d/%d' % (sum(1 for b in BOOKS if b.get('_mark')), len(BOOKS)),
         '平台只公开 Top10 热门划线，不公开全站划线总数，故取最热一条的人数（markBest）', '✅'],
        ['出版时间 / 出版社', 'bookInfo.publishTime / publisher', '99% / 95%', 'ISBN 网页端多为 null', '✅'],
        ['分类', 'shelf.books[].category', '%d/%d' % (len(BOOKS), len(BOOKS)), '如「经济理财-理财」，拆成一级/二级', '✅'],
        ['是否下架', 'shelf.books[].soldout', '%d/%d' % (len(BOOKS), len(BOOKS)), '下架后不可读不可买', '✅'],
        ['是否已买断', 'shelf.books[].paid', '%d/%d' % (len(BOOKS), len(BOOKS)), '布尔值，无成交价、无购买时间', '✅'],
        ['阅读进度 %', 'shelf.bookProgress[].progress', '—', '0-100，100＝读完（含靠会员卡读完的）', '✅'],
        ['阅读时长', 'bookProgress[].readingTime（秒）', '—', '可汇总成总阅读时长', '✅'],
        ['开始读 / 读完时间', '/web/book/getProgress → startReadingTime / finishTime', '—', '需逐本调一次', '✅'],
        ['书架分组归属', 'shelf/sync → archive[]', '%d 组' % len(G['ARCHIVES']), '这是分组，不是 App 书单', '✅'],
        ['纸书 SKU / 讲书 / 有声', 'paperBook / hasLecture / lectureBooks', '少数', '本账号暂无讲书', '✅'],
        ['章节目录 / 章节字数', 'POST /web/book/chapterInfos', '单本可用', '批量传多 id 返回空，需逐本', '🟡'],
        ['我的划线 / 想法', 'bookmarklist、/web/review/list?mine=1', '单本可用', '需逐本调', '🟡'],
        ['简介', '详情页 bookInfo.intro', '~97%', '太长，未进表', '🟡'],
        ['历史订单 / 成交价', '/web/pay/order/list 等均 404', '—', '拿不到当时实际付了多少钱', '❌'],
        ['好友评论 / 好友在读', '/web/book/friend* 全 404', '—', '好友关系只在 App 端，需 App 鉴权或 wrk- 密钥', '❌'],
    ]))
    A('')
    A('### 1.2 账户维度：钱与权益')
    A('')
    A(_tbl(['字段', '来源接口', '当前值', '说明', '状态'], [
        ['书币总余额', 'POST /web/pay/balance {"pf":"android_html"}', '%.2f 元' % COIN, '总余额＝赠币＋充值', '✅'],
        ['赠币余额', '同上 giftBalance', '%.2f 元' % COIN, '有有效期，先扣赠币', '✅'],
        ['体验卡到期', '/web/pay/memberCardSummary → expiredTime', (G['COIN_EXPIRE'] or '—')[:19],
         'remainTime 秒级倒计时', '✅'],
        ['会员卡状态', '同上 expired / isPaying / payingUsedDay', '见接口', '曾开过会员卡，故多本书读到 100%', '✅'],
        ['赠币逐笔到期', '—', '—', '网页端无接口，只能 App「我的→书币」看', '❌'],
        ['优惠券列表', '/web/pay/coupon(s) 均 404', '—', '只能在下单页看到可用券', '❌'],
    ]))
    A('')
    A('### 1.3 账号维度：阅读统计')
    A('')
    A(_tbl(['字段', '来源', '当前值', '说明', '状态'], [
        ['累计阅读时长', 'bookProgress[].readingTime 汇总', '%.0f 小时' % G['TOTAL_HOURS'],
         '仅统计同步到云端的记录', '✅'],
        ['已读完本数', 'progress == 100', '%d 本（全库）' % len(G['READ_DONE']), '含已买断的', '✅'],
        ['在读本数', '0 < progress < 100', '%d 本（待付费书里）' % len(G['NEED_READING']), '可做「续读清单」', '✅'],
        ['书架总数', 'shelf.bookCount / pureBookCount', '%d / %d' % (len(BOOKS), len(BOOKS) + 1),
         '与 App 差 1~2 本＝文章收藏 + 有声/专辑', '✅'],
        ['年度/周度阅读报告', '/web/readdata/detail 等 4 个路径均 404', '—', '网页端未开放，App 内有', '❌'],
    ]))
    A('')
    A('### 1.4 基于这些数据还能做什么')
    A('')
    A(_tbl(['功能', '怎么做', '用到的数据', '价值', '难度'], [
        ['续读清单 Top N', '按 progress 40~99% 排序，附剩余章节估算', '进度、章节、字数', '快速找回「读一半」的书', '低'],
        ['已读完不必买', 'progress=100 且 paid=0', '进度 + paid', '避免重复花钱（已做）', '已做'],
        ['买断 vs 会员卡', '待付费标价合计 vs 19 元×月数（会员卡须另行微信支付，书币不能抵扣）', '标价、进度',
         '把「买断多少本」与「另外开卡」分开算', '低'],
        ['下架预警', '每天跑一次 soldout，变动即提醒', 'soldout', '体验卡书下架率最高，预警有用', '中'],
        ['阅读画像', '按分类统计读完率/时长/评分偏好', '分类、进度、时长', '看清自己真正在读什么', '低'],
        ['笔记/划线资产导出', '逐本拉 bookmarklist + review，导出 Markdown', '划线、想法', '沉淀读书笔记', '中'],
        ['价格变动监控', '每天存一份 centPrice，做价差表', '标价快照', '抓促销窗口再买', '中'],
    ]))
    A('')

    # ============================================================ 二
    A('## 二、互动证据（划线 / 想法）')
    A('')
    A('### 2.1 为什么用「人均密度」而不是绝对数量')
    A('')
    for t in [
        '- 绝对量与阅读人数的 log 相关系数：想法数 0.935、热门划线人次 0.914、评分人数 0.913。'
        '「划线多」绝大部分只是因为「读的人多」，直接用绝对量等于把热度再奖励一次。',
        '- 口径改为「每读者互动率」：d_mark ＝ 最热一条划线的人数 ÷ 阅读人数；d_idea ＝ 想法数 ÷ 阅读人数。',
        '- 密度对小样本不稳（50 人读、5 条想法＝10%%），故向类均值做伪计数平滑：'
        'd ＝ (n ＋ %d × 本类均值) ÷ (阅读人数 ＋ %d)，相当于 200 个「虚拟读者」。' % (int(VSPARAM.get('K_PSEUDO', 200)), int(VSPARAM.get('K_PSEUDO', 200))),
        '- 再按一级分类的类内中位数相对化：x_mark ＝ d_mark ÷ 本类中位，x_idea 同理（消除类型差异）。',
        '- 合成：x_e ＝ x_mark^0.55 × x_idea^0.45（划线权重高于评论）；'
        'fac_e ＝ clip｛[(0.2＋0.8·Hill(x_e, K=0.50)) ÷ 0.8]^0.40, 0.88, 1.06｝——最多加 6%、最多扣 12%。',
        '- 若某本书 bestbookmarks 返回空（平台没给热门划线），fac_e 只由想法数决定，**不当作零互动惩罚**。',
        '- `listType=3` 的 totalCount 不是划线数，而是「星评人数」（与 ratingCountOld 完全相等），旧版曾误用，已纠正。',
    ]:
        A(t)
    A('')

    n = len(have) or 1
    bands = [('0.88（触底：互动明显偏低）', lambda v: v <= 0.885),
             ('0.885–0.93', lambda v: 0.885 < v <= 0.93),
             ('0.93–0.97', lambda v: 0.93 < v <= 0.97),
             ('0.97–1.00（接近类中位）', lambda v: 0.97 < v <= 1.0),
             ('>1.00（互动高于类中位，加分）', lambda v: v > 1.0)]
    A('### 2.2 互动修正 fac_e 的分布（%d 本有评分、可算互动因子）' % len(have))
    A('')
    rows = []
    for label, fn in bands:
        sub = [b for b in have if fn(b['_fac_e'])]
        if not sub:
            continue
        rows.append([label, len(sub), '%.1f%%' % (100.0 * len(sub) / n),
                     _fmt(sum(b['_vscore'] for b in sub) / len(sub))])
    A(_tbl(['区间', '本数', '占比', '平均价值分'], rows))
    A('')

    def board(title, items):
        A('### ' + title)
        A('')
        A(_tbl(['#', '书名', '二级分类', '阅读人数', '想法数', '最热划线', 'x_e', 'fac_e', '价值分'],
               [[i, b.get('title') or '', b.get('_type2') or '', b['_readcnt'],
                 int(b['_idea']) if b.get('_idea') is not None else '—',
                 int(b['_mark']) if b.get('_mark') is not None else '（平台未公开）',
                 _fmt(b.get('_xe'), 2), _fmt(b.get('_fac_e'), 3), _fmt(b['_vscore'])]
                for i, b in enumerate(items, 1)]))
        A('')

    pool = [b for b in have if b.get('_fac_e')]
    board('2.3 互动最强（fac_e 最高，同分按 x_e 排）Top 20',
          sorted(pool, key=lambda b: (-b['_fac_e'], -(b.get('_xe') or 0)))[:20])
    board('2.4 互动最弱（fac_e 触底 0.88）Top 20',
          [b for b in sorted(pool, key=lambda b: (b['_fac_e'], b.get('_xe') or 0))
           if b['_fac_e'] < 0.9][:20])

    A('### 2.5 各类型的互动密度（一级分类，有评分书 ≥10 本）')
    A('')
    rows = []
    for name, v in (VTYPE.get('c1') or {}).items():
        if (v.get('nRated') or 0) < 10:
            continue
        rows.append([name, v['nRated'], _fmt(v.get('d_mark'), 4), _fmt(v.get('d_idea'), 4),
                     _fmt(v.get('meanFacE'), 3), _fmt(v.get('meanScore'))])
    rows.sort(key=lambda x: -(x[4] if isinstance(x[4], float) else 0))
    A(_tbl(['一级分类', '有评分', '类内 d_mark', '类内 d_idea', '平均 fac_e', '平均价值分'], rows))
    A('')

    # ============================================================ 三
    A('## 三、口径说明')
    A('')
    A('### 3.1 「书单」与「书架分组」是两个不同功能')
    A('')
    A('- 「书架分组明细」来自 `/web/shelf/sync` 的 `archive[]`，共 %d 组 / %d 条归属（%s）。'
      % (len(G['ARCHIVES']), G['_ARCH_ROWS'], '、'.join(a.get('name') or '-' for a in G['ARCHIVES'])))
    A('- App 里的「书单」（我的 → 书单，可公开分享的那种）是独立功能，不在书架数据里；'
      '其接口在 i.weread.qq.com，网页登录态调用一律返回 -2012「登录超时」。')
    A('- 因此书单只能靠分享链接抓：`https://weread.qq.com/misc/booklist/<id>` 服务端渲染，免登录即可取全部书目。')
    A('')
    A('### 3.2 四类权益判定规则')
    A('')
    A('```')
    A('paid = 1                     → 永久可读（已购买 / 已拥有 / 获赠）')
    A('paid = 0 且 payingStatus = 2 → 付费会员卡可读')
    A('paid = 0 且 payingStatus = 1 → 体验卡可读')
    A('paid = 0 且 payingStatus = 0 → 免费（导入书 / 公众号文章，无标价）')
    A('```')
    A('')
    A('本表得到 %d / %d / %d / %d，其中三档与 App 显示完全吻合。付费墙文案也逐一印证：'
      '会员卡档显示「会员卡 19 元/月 · 开通后即可阅读」，体验卡档显示「你正在使用体验卡阅读付费部分」。'
      % (agg['付费会员卡可读']['n'], agg['体验卡可读']['n'], agg['永久可读']['n'], agg['免费']['n']))
    A('')
    A('### 3.3 与 App 差几本的原因')
    A('')
    A('书架接口另返回「文章收藏」(mpbook) 1 条，且 pureBookCount 比 bookCount 多 1；'
      'App 侧还会把有声书/专辑计入。网页接口不返回这些条目，所以本表覆盖 %d 本电子书。' % len(BOOKS))
    A('')
    A('### 3.4 价格口径')
    A('')
    A('- 标价(元) ＝ centPrice ÷ 100，是这本书的买断价，来源为书籍详情页。')
    A('- 标价 ≠ 你还需支付的钱：永久可读的书同样有标价，但你已经拥有。')
    A('- 会员卡/体验卡可读的书，理论上靠订阅即可读，不必逐本买断；标价合计只表示「全部买断要多少钱」。')
    A('')
    A('### 3.5 账户资产')
    A('')
    A('- 书币余额来自 `POST /web/pay/balance`；`pf=ios_html` 时数据出现在 peerBalance / peerGiftBalance，两侧一致。')
    A('- 体验卡到期时间：`/web/pay/memberCardSummary?pf=ios` 的 expiredTime，本次为 %s。' % (G['COIN_EXPIRE'] or '未取到'))
    A('- 余额是书币/赠币，**不是现金**：不能提现、不能转赠、不能支付 19 元/月的付费会员卡（开卡须另行微信支付）。')
    A('- 赠币逐笔到期时间网页端无接口，只能到 App「我的 → 书币」查看。')
    A('')
    A('### 3.6 「积分购买方案」sheet 的口径')
    A('')
    A('- 收录条件：待付费书单（会员卡可读＋体验卡可读，共 %d 本）中，单本标价 ≤ %.2f 元的书 → %d 本。'
      % (len(NEED), COIN, len(COIN_OK)))
    A('- **「%d 本」的含义**：任选其中一本都付得起，不等于余额能把 %d 本全买下；'
      '这 %d 本全买要 %.2f 元，是余额的 %.1f 倍。余额实际只够买 %d 本。'
      % (len(COIN_OK), len(COIN_OK), len(COIN_OK), COIN_OK_MONEY,
         COIN_OK_MONEY / COIN if COIN else 0, PS_CNT['n']))
    A('- 每本只买一次：候选池按 bookId 去重（0/1 背包）。已下架的书虽标价够但买不到，全部剔除。')
    A('- **四个方案全部不包含**「已读完（进度 100%%）、但还没买断（= 未获永久授权）」的 %d 本：'
      '它们靠体验卡/会员卡读完了，再买只是拿个永久副本，没有阅读增量。'
      % (len(COIN_POOL) - len(COIN_POOL_X)))
    A('')
    A('四个方案：')
    A('')
    A(_tbl(['方案', '目标', '算法', '本数', '花费', '剩余', '平均价值分'], [
        ['① 性价比最高', '每一元买到的价值最多', '按 价值分÷标价 降序贪心装满预算', PS_RATIO['n'],
         '%.2f' % PS_RATIO['spent'], '%.2f' % PS_RATIO['left'], _fmt(PS_RATIO['av'])],
        ['② 本数最多', '书架厚度最大化', '0/1 背包本数优先，同本数下价值分合计最高', PS_CNT['n'],
         '%.2f' % PS_CNT['spent'], '%.2f' % PS_CNT['left'], _fmt(PS_CNT['av'])],
        ['③ 绝对价值最高', '只买 1 本，把积分一次花在一本最值的书上',
         '基准表口径的 %d 个一级分类 × 每类 V 前 2 = %d 个候选位（实到 %d 本）；排序= V×标价'
         % (PS_VAL['nTypeStd'], PS_VAL['nTypeStd'] * 2, PS_VAL['nCand']),
         PS_VAL['n'], '%.2f' % PS_VAL['spent'], '%.2f' % PS_VAL['left'], _fmt(PS_VAL['av'])],
        ['④ 抢下架', '先买可能没了的书', '价值分≥70 的书上做 0/1 背包，max Σ(下架风险×价值分)', PS_RISK['n'],
         '%.2f' % PS_RISK['spent'], '%.2f' % PS_RISK['left'], _fmt(PS_RISK['av'])],
    ]))
    A('')
    A('③ 的实现（只买一本，与其余三个方案逻辑完全不同）：')
    A('')
    A('- 综合价值 `V = 0.45×价值分 + 0.25×平台原评分 + 0.15×参与率池内百分位 + 0.15×阅读人数池内百分位`，价格不进 V。')
    A('- 分类口径与「分类基准与类型评分」第一节完全一致：只取「有评分 ≥ 10 本、且非未分类」的一级分类，'
      '共 %d 类，类排名=该节按类均分从高到低的位次（可与该节 # 列逐行对照）。'
      '艺术 / 生活百科 / 漫画 / 童书等样本 < 10 本的类不进基准表，也不进本方案候选。'
      % PS_VAL['nTypeStd'])
    _empty = [t for t in TYPE1_STD if not POW_CANDS.get(t)]
    A('- 候选结构：**%d 类 × 每类 V 前 2 名 = %d 个候选位**；其中 %d 类在待付费可买池里有可买的书、'
      '%d 类（%s）候选位空缺 → **实到 %d 本**（单本标价均 ≤ %.2f 元，无需剔除）。'
      '注意 %d 是**本数**不是分类数：%s的书全部已下架、且标价 0（免费/永久可读），不在待付费可买池。'
      % (PS_VAL['nTypeStd'], PS_VAL['nTypeStd'] * 2, PS_VAL['nTypePool'], len(_empty),
         '、'.join(_empty) or '无', PS_VAL['nCand'], COIN, PS_VAL['nCand'],
         '、'.join(_empty) if _empty else ''))
    A('- 表格按 15 类**逐类分组**列出（每组 2 行，空缺的类留一行写原因），组内按 V 降序，'
      '并给「类内 V 排名」与「全局产出排名」两列。')
    if VAL_PICK:
        A('- 排序：**单本价值产出 = V × 标价**。只买一本时"剩余最少"与"价值最高"只能合并——'
          '标价越高余额剩得越少，V 越高这本书本身越值。选中的是 **《%s》%.2f 元、剩 %.2f 元（花满 %.1f%%）、V %.1f**。'
          % (VAL_PICK['title'], VAL_PICK['_price'], PS_VAL['left'], 100 * PS_VAL['fill'], PS_VAL['V']))
        A('- 两个单目标参考：只看 V 最高 → 《%s》V %.1f、价 %.2f；只看剩余最少 → 《%s》价 %.2f、V %.1f。'
          % (PS_VAL['topV']['title'], PS_VAL['topV']['_V'], PS_VAL['topV']['_price'],
             PS_VAL['topPrice']['title'], PS_VAL['topPrice']['_price'], PS_VAL['topPrice']['_V']))
    else:
        A('- **本方案无书可选**：余额 %.2f 元低于全部候选标价，"V × 标价"的排序结果为空。'
          '这是预算约束的结果，不是算法异常；攒够书币后重跑即可。' % COIN)
    A('')
    A('### 3.7 评分与下架风险')
    A('')
    A('- 评分来源：详情页 bookInfo.newRating ÷ 10；评分人数取 newRatingCount。'
      '%d 本中 %d 本有评分（%.1f%%），区间 %.1f~%.1f，中位 %.1f。'
      % (len(BOOKS), n_rated, 100.0 * n_rated / len(BOOKS),
         min(b['_rating'] for b in rated), max(b['_rating'] for b in rated),
         sorted(b['_rating'] for b in rated)[n_rated // 2]))
    A('- 下架风险分 ＝ 100 ×（0.5×权益类型档下架率 ＋ 0.3×出版年档下架率 ＋ 0.2×评分人数档下架率），'
      '三档下架率均为书架实测值（含 Laplace 平滑）。')
    A('- 下架高度集中在「体验卡可读」：%d 本中 %d 本已下架（%.1f%%）；'
      '「付费会员卡可读」%d 本 %d 本（%.1f%%）；「永久可读」%d 本 %d 本（%.1f%%）。'
      % (len([b for b in BOOKS if b['_cat'] == '体验卡可读']),
         sum(1 for b in BOOKS if b['_cat'] == '体验卡可读' and b['_sold']),
         100.0 * sum(1 for b in BOOKS if b['_cat'] == '体验卡可读' and b['_sold'])
         / max(1, len([b for b in BOOKS if b['_cat'] == '体验卡可读'])),
         agg['付费会员卡可读']['n'],
         sum(1 for b in BOOKS if b['_cat'] == '付费会员卡可读' and b['_sold']),
         100.0 * sum(1 for b in BOOKS if b['_cat'] == '付费会员卡可读' and b['_sold'])
         / max(1, agg['付费会员卡可读']['n']),
         agg['永久可读']['n'],
         sum(1 for b in BOOKS if b['_cat'] == '永久可读' and b['_sold']),
         100.0 * sum(1 for b in BOOKS if b['_cat'] == '永久可读' and b['_sold'])
         / max(1, agg['永久可读']['n'])))
    A('- 评分本身未纳入风险模型：已下架书评分中位 73.4、未下架 79.5，仅弱相关。')
    A('- 阅读进度：`/web/shelf/sync` 的 bookProgress[].progress（0-100）＋ `/web/book/getProgress` 的 finishTime；'
      '进度 100% 不等于已购买。')
    A('')
    A('### 3.8 价值评分模型')
    A('')
    A('```')
    A('conf_stat = WilsonLCB(raw/100, rc, z=1.645) ÷ (raw/100)          # ① 样本量')
    A('rel_r     = [0.20 + 0.80·Hill(x_r)] ÷ 0.80,  x_r = 本类参与率相对度   # ② 类内参与率')
    A('heat      = 0.95 + 0.10·Hill(x_d),           x_d = 阅读人数 ÷ 本类几何均值   # ③ 受众面')
    A('fac_e     = clip{[(0.2+0.8·Hill(x_e, K=0.50))÷0.8]^0.40, 0.88, 1.06}   # ④ 互动证据')
    A('cap       = min(1.0, (raw ÷ 91)^0.5)                    # 封顶线：9.1 分以上才允许打满')
    A('conf      = min(conf_stat × rel_r^0.5 × fac_e × heat, cap)')
    A('价值分     = 平台原评分 × conf        # 只下调不上调，永远 ≤ 平台评分')
    A('```')
    A('')
    A('- 封顶线为什么随分数走：四个乘数的上限叠乘 = 1.12×1.06×1.05 = 1.246，足以把任何 conf_stat 推到 1。'
      '若 cap 写成常数 1，只要样本够大、类内够热，任何分数的书都会被打成「价值分 = 平台评分」，'
      '模型在它们身上完全失去信息量（改造前实测 36.9%% 的书如此，最低出现在 4.85 分）。'
      '现在 cap=√(raw÷91)：conf 打满只剩 %d 本（%.1f%%），全部在 9.1 分以上；8.5 分的书价值分约 82，仍有可见折扣。'
      % (len([b for b in BOOKS if b['_vscore'] and (b.get('_conf') or 0) >= 0.9999]),
         100.0 * len([b for b in BOOKS if b['_vscore'] and (b.get('_conf') or 0) >= 0.9999])
         / max(1, len([b for b in BOOKS if b['_vscore']]))))
    A('- 参数：K_REL=%s、N_HILL=%s、W_R=%s、REL_FLOOR=%s、REL_CEIL=%s、HEAT_BASE=%s、HEAT_RANGE=%s、'
      'K_E=%s、W_E=%s、W_MARK=%s、W_IDEA=%s、E_FLOOR=%s、E_CEIL=%s、K_PSEUDO=%s、'
      'CAP_RAW=%s、CAP_GAMMA=%s。参数集中在 '
      '`score_value.py` 顶部，改完重跑 `score_value.py` 即可。'
      % (VSPARAM.get('K_REL'), VSPARAM.get('N_HILL'), VSPARAM.get('W_R'),
         VSPARAM.get('REL_FLOOR'), VSPARAM.get('REL_CEIL'), VSPARAM.get('HEAT_BASE'),
         VSPARAM.get('HEAT_RANGE'), VSPARAM.get('K_E'), VSPARAM.get('W_E'),
         VSPARAM.get('W_MARK'), VSPARAM.get('W_IDEA'), VSPARAM.get('E_FLOOR'),
         VSPARAM.get('E_CEIL'), VSPARAM.get('K_PSEUDO'),
         VSPARAM.get('CAP_RAW'), VSPARAM.get('CAP_GAMMA')))
    A('- 类内基准的分档处理：一级类有评分书 ≥10 本用自己的均值，否则沿用全书架基准；'
      '二级类 ≥8 本时向一级类轻度收缩（权重 n÷(n+6)），<8 本直接继承一级类。')
    A('- 阅读人数基准用几何均值（平均数量级），不用算术均值：小说类算术 52.0 万 vs 几何 13.5 万，'
      '算术均值被个位数畅销书拉到 3.7 倍。')
    A('- 局限：① Wilson 只修正抽样误差，修不了「读完的人才来打分」的自选择偏差；'
      '② 平台不公开全站划线总数，只能用 Top10 热门划线代理；'
      '③ 互动密度与「书好不好」不是单调关系——小众经典读者不爱发言，所以 fac_e 最低只到 0.88。')
    A('- 为什么不能拿阅读人数当可信度折扣：一本 3000 人读、72 人评的书，与 30 万人读、72 人评的书，'
      '评分的统计可信度是同一个数量级，差异在受众面而非证据强度。')
    A('')
    A('### 3.9 书单（/misc/booklist）怎么拿到的')
    A('')
    A('- 公开书单可分享，链接形如 `https://weread.qq.com/misc/booklist/<id>`，页面服务端渲染（__NUXT__），'
      '**不需要登录**即可拿到书单名、作者、按顺序的 bookIds 与每本推荐值。')
    A('- 逐本价格/推荐值人数用登录态 `/web/book/info` 补；在书架的书直接用书架数据。')
    A('- 已接入 `booklist_urls.txt` 里的全部书单：%d 个 / %d 条归属 / %d 本去重（书单之间无重复收录）。'
      % (len(BLX), bl_tot, bl_uniq))
    A('- 书单 ID 后 10 位是服务端乱序串（相邻 ID 实测 404、base62 解码与创建时间不对应、App 接口需 App 鉴权），'
      '无法由已知链接推测其余书单，只能逐个分享。')
    A('')
    A('### 3.10 合并去重与两张排名表的口径')
    A('')
    A('- 合并去重：书架 %d 本 ＋ 书单里未加架 %d 本 ＝ %d 本（按 bookId 去重，书单之间无重复）。'
      % (shelf_n, out_n, len(UNION)))
    A('- 书单里未加架的书，原先只有价格与评分；本次用登录态补抓了累计阅读人数、想法数、热门划线人次，'
      '因此可以用与书架完全相同的 公式打分（类内基准取该类书架基准）。')
    A('- 「热书排名」热度分 ＝ 100 ×［0.40×阅读人数百分位 ＋ 0.30×评分人数百分位 ＋ '
      '0.20×最热划线人次百分位 ＋ 0.10×想法数百分位］，四项先取 log10(1+x)，缺项按可用项权重归一化。')
    A('- 「分类基准与类型评分」第三节按一级分类列出合并去重后的全部书，按类内价值分降序；'
      '类内相对 ＝ 该书价值分 − 本类均值（基数含书单书，与「价值评分」sheet 的「类内相对」不完全一致，看趋势即可）。')
    A('')
    A('### 3.11 「零钱备选(≤%.0f 元)」sheet 的口径' % LOW_CAP)
    A('')
    A('- 存在意义：余额低于最便宜的书时，四套方案的本数全是 0，'
      '这张表回答"这点钱还能买什么、攒到多少能买第一本"，避免报告只剩一句"买不了"。')
    A('- 分两个来源，**不要混为一谈**：')
    A('  - **① 你的书架/书单**：待付费（会员卡可读 / 体验卡可读）＋ 未下架 ＋ 有标价 ＋ 有价值评分 '
      '＋ 标价 ≤ %.2f 元，按价值评分降序。这是你自己的书，能直接下单。' % LOW_CAP)
    A('  - **② 微信读书全库**：`data/lowprice_seed.json`，由 `fetch_lowprice.py` 用关键词矩阵'
      '调 `/web/search/name_books` 采样得到（接口返回里直接带 price / newRating，不必逐本查价）。'
      '与你书架无关，只在书架这一档为空/不够挑时兜底。')
    if LOW_POOL:
        _vs = [b['_vscore'] for b in LOW_POOL]
        A('- 本次 ① 区：**%d 本**，价值分 %.1f ~ %.1f（中位 %.1f）；'
          '余额 %.2f 元能买其中 **%d 本**。'
          % (len(LOW_POOL), min(_vs), max(_vs), sorted(_vs)[len(_vs) // 2], COIN, len(LOW_AFFORD)))
    else:
        A('- 本次 ① 区：**0 本**（你的书架/书单里没有 ≤ %.2f 元、未下架、有评分的书），'
          '此时 ② 区自动成为主内容。' % LOW_CAP)
    if LOW_SEED:
        _mr = LOW_SEED_META.get('minRating')
        _pg = LOW_SEED_META.get('pagesByGroup') or {}
        _pg_txt = '、'.join('%s %s 页' % (k, v) for k, v in _pg.items()) if _pg else '-'
        A('- 本次 ② 区：**%d 本**（评分 ≥ %s）。采样口径：%d 个关键词，'
          '**本轮 %s 次请求 / %s 分钟**（含补分类），累计 %s 次请求、扫描 %s 条，生成于 %s%s。'
          % (len(LOW_SEED), ('%g' % _mr) if isinstance(_mr, (int, float)) else '70',
             LOW_SEED_META.get('keywordCount') or 0,
             LOW_SEED_META.get('lastRunRequests') or LOW_SEED_META.get('requests') or '-',
             LOW_SEED_META.get('lastRunMinutes') or LOW_SEED_META.get('elapsedMin') or '-',
             LOW_SEED_META.get('requests') or '-',
             LOW_SEED_META.get('scanned') or '-',
             LOW_SEED_META.get('generatedAt') or '未知',
             ('（分 %s 轮抓取后合并去重）' % LOW_SEED_META['rounds'])
             if (LOW_SEED_META.get('rounds') or 1) > 1 else ''))
        A('- 分组页数（热门类目翻更深）：%s。' % _pg_txt)
        A('- **分类纠偏**：平台 `category` 有明显不当的（如《道德经》→「历史」、《论语》→「教育学习」、'
          '《红楼梦》→「历史」），按 `fetch_lowprice.py` 的 `CLASS_FIX` 规则改写为「国学古籍 / 古典文学 / '
          '法律法规 / 期刊杂志」，共纠正 **%s 本**；其余沿用平台分类。'
          '原始平台分类保留在种子的 `platformType1` 字段，可逐本核对。'
          % (LOW_SEED_META.get('typeFixed') or 0))
        A('- **分类推断**：平台对一部分低价书不返回分类（本次 %s 本），再查详情页也拿不到。'
          '这些书按书名 → 作者 → 出版社的顺序归入最接近的一级分类，报告里分类名前的 **≈** 就是这个意思；'
          '本次推断 %s 本，剩余 %s 本标「未分类」。推断规则在 `fetch_lowprice.py` 的 '
          '`TYPE_RULES2` / `AUTHOR_RULES` / `PUB_RULES`，可用 `--reclassify` 只重算分类、不重抓。'
          % (LOW_SEED_META.get('typeEmptyBefore') or 0,
             LOW_SEED_META.get('typeInferred') or 0,
             len([b for b in LOW_SEED if not (b.get('_type1') or '').strip()])))
        A('- **② 区是采样，不是全库穷举**：微信读书没有「按价格浏览全库」的接口，'
          '能否命中某本低价书取决于关键词是否覆盖到它。加词或加深页数后重跑 `fetch_lowprice.py` 即可扩容。')
    else:
        A('- 本次 ② 区：**未生成**。跑一次 `python scripts/fetch_lowprice.py` 生成 `data/lowprice_seed.json`。')
    A('- 与四方案的关系：四方案是"给定余额求最优组合"，本表是"给定低价档位求高价值书"，'
      '两者不互相替代。余额够时本表也会照常列出（低价好书本来就值得关注）。')
    A('')
    A('### 3.12 数据局限')
    A('')
    A('- 网页接口不返回「是否已购买」的成交价与购买时间，只返回 paid 布尔标记。')
    A('- 有声书/专辑、文章收藏不在 books 列表中，未纳入统计。')
    A('- 权益状态是账号级、随时间变化（体验卡 %s 到期），本表为 %s 快照。'
      % ((G['COIN_EXPIRE'] or '未知')[:16], SNAP_NOW))
    A('- 平台不公开全站划线总数，划线维度只能用「Top10 最热一条的人次」代理，读者沉默的书会被低估一点。')
    A('')

    with open(path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(L))
    print('docs:', path, len(L), 'lines')
    return path
